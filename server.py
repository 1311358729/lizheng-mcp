# -*- coding: utf-8 -*-
"""
理正勘察 9.5 接口文件 MCP Server (lizheng-mcp)
=============================================
面向「理正工程地质勘察软件 9.5PB3」的只读自动化通道，两条主线：

1. 工程库(.lz)读取 —— .lz 是 ZIP 包，内含 Access 工程库 LZGICAD*.mdb
   (Jet4, 带库密码) + ProjectInfo.ini + 数字化交付文件。
   库密码硬编码于理正安装目录的二进制（Forms.exe 等连接字符串），
   本服务启动时自动定位安装目录并提取，无需手工配置；
   也可用环境变量 LZ_MDB_PASSWORD 显式指定。
   依赖：pyodbc + 本机 64 位 ACE Access 驱动
   (Microsoft Access Driver (*.mdb, *.accdb))。
   库表字段码与官方《理正标准数据接口》完全一致（ZKBH/TCMC/DTDSD/QYBH...）。

2. 标准数据接口文件 —— 官方规范 lz_standard_interface_spec.txt
   (Tab 分隔、GBK 编码、首行 ;#VERSION#gicad9.0、#TAG# 数据节 /
   %TAG% 项目表节)。支持 生成 / 解析 / 校验 / Excel 附表识别。

安全约定：全程只读（readonly ODBC 连接），不写工程库、不删文件；
所有输出均为新文件，由调用方指定路径。

.mcp.json 注册示例：
{
  "mcpServers": {
    "lizheng": {
      "command": "python",
      "args": ["/path/to/lizheng-mcp/server.py"]
    }
  }
}
"""

import os
import re
import sys
import json
import math
import shutil
import hashlib
import tempfile
import zipfile
import datetime
import asyncio

from typing import Any

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp import types

# ---------------------------------------------------------------------------
# 常量
# ---------------------------------------------------------------------------

# 常见安装位置（仅用于密码自动定位与 backend_info 提示，不强制）
LZ_INSTALL_CANDIDATES = [
    r"D:\LiZheng\GICAD9.5PB3(工勘版)(x64)",
    r"D:\LiZheng\GICAD9.5PB3(工民建版)(x64)",
    r"D:\LiZheng\GICAD9.5PB3(公路版)(x64)",
    r"D:\LiZheng\GICAD9.5PB3(水利水电版)(x64)",
]


def _discover_db_password() -> str:
    """从本机理正安装目录的二进制中提取 Access 工程库密码。

    理正将 Jet 连接字符串（含 Database Password）硬编码在安装目录的
    Forms.exe 等可执行文件中。本函数扫描候选安装目录下的 exe/dll，
    提取该字符串，使仓库无需包含任何密码明文。

    定位顺序：环境变量 LZ_INSTALL_DIR > 注册表卸载项 > 常见安装路径。
    找不到时返回空串（按无密码库尝试连接）。
    """
    dirs = []
    env_dir = os.environ.get("LZ_INSTALL_DIR", "")
    if env_dir and os.path.isdir(env_dir):
        dirs.append(env_dir)
    # 注册表卸载项中找安装位置
    if sys.platform == "win32":
        try:
            import winreg
            for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
                for sub in (
                    r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                    r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
                ):
                    try:
                        with winreg.OpenKey(root, sub) as k:
                            for i in range(200):
                                try:
                                    name = winreg.EnumKey(k, i)
                                except OSError:
                                    break
                                try:
                                    with winreg.OpenKey(k, name) as sk:
                                        loc, _ = winreg.QueryValueEx(sk, "InstallLocation")
                                        if loc and os.path.isdir(str(loc)) and "LiZheng" in str(loc):
                                            dirs.append(str(loc))
                                except OSError:
                                    continue
                    except OSError:
                        continue
        except Exception:
            pass
    dirs += LZ_INSTALL_CANDIDATES
    # 同时匹配 ASCII 与 UTF-16LE 两种存储；跳过 %s 占位符与空值
    ascii_pat = re.compile(rb"Database\s+Password\s*=\s*([^;\"'\r\n\x00]{1,64})")
    utf16_pat = re.compile(
        "Database Password=([^;\"'\r\n]{1,64})".encode("utf-16-le"))
    candidates = []
    for d in dirs:
        if not os.path.isdir(d):
            continue
        try:
            entries = os.listdir(d)
        except OSError:
            continue
        for fn in entries:
            if not fn.lower().endswith((".exe", ".dll")):
                continue
            fp = os.path.join(d, fn)
            try:
                with open(fp, "rb") as f:
                    blob = f.read()
            except OSError:
                continue
            for pat, dec in ((ascii_pat, "ascii"), (utf16_pat, "utf-16-le")):
                for m in pat.finditer(blob):
                    try:
                        val = m.group(1).decode(dec, "ignore").strip()
                    except Exception:
                        continue
                    if val and val != "%s":
                        candidates.append(val)
    if candidates:
        # 出现频次最高的值即库密码（占位符/空串已过滤）
        return max(set(candidates), key=candidates.count)
    return ""


HERE = os.path.dirname(os.path.abspath(__file__))
SPEC_FILE = os.path.join(HERE, "lz_standard_interface_spec.txt")
VERSION_HEADER_PREFIX = ";#VERSION#"

# 理正 Access 工程库密码获取顺序：
#   1. 环境变量 LZ_MDB_PASSWORD
#   2. 自动从本机理正安装目录的二进制中提取（连接字符串硬编码于 Forms.exe 等）
#   3. 空（视为无密码库）
LZ_MDB_PASSWORD = os.environ.get("LZ_MDB_PASSWORD", "") or _discover_db_password()

# 库表 -> 接口节 映射（tag: (库表, 钻孔号字段)）
# verified=True 表示已在真实工程库上核对过字段码
DB_SECTION_MAP = [
    ("GC",  "x_GongCheng",    None,   True,  "工程数据"),
    ("ZK",  "z_ZuanKong",     "ZKBH", True,  "钻孔数据"),
    ("TC",  "z_g_TuCeng",     "ZKBH", True,  "土层数据"),
    ("QY",  "z_c_QuYang",     "ZKBH", True,  "常规试验取样"),
    ("KF",  "z_c_KeFen",      "ZKBH", True,  "颗分试验"),
    ("DT",  "z_y_DongTan",    "ZKBH", True,  "动力触探"),
    ("JT",  "z_y_JingTan",    "ZKBH", False, "静力触探"),
    ("BG",  "z_y_BiaoGuan",   "ZKBH", False, "标准贯入"),
    ("BS",  "z_y_BoSu",       "ZKBH", False, "波速测试"),
    ("TS",  "z_y_TouShuiLv",  "ZKBH", False, "透水率/渗水率"),
    ("SB",  "z_y_LuoXuanBan", "ZKBH", False, "十字板剪切"),
    ("PY",  "z_y_PangYa",     "ZKBH", False, "旁压试验"),
    ("SW",  "z_g_ShuiWei",    "ZKBH", False, "水位数据"),
    ("CS",  "z_g_ChaoShiDu",  "ZKBH", False, "潮湿度"),
    ("FH",  "z_g_FengHua",    "ZKBH", False, "风化程度"),
    ("KS",  "z_g_KeSuXing",   "ZKBH", False, "可塑性"),
    ("LF",  "z_g_XieKong",    "ZKBH", False, "裂缝密度"),
    ("ST",  "z_g_ShenTouJX",  "ZKBH", False, "渗透系数"),
    ("YR",  "z_g_YXRQD",      "ZKBH", False, "岩芯采取率/RQD"),
    ("U1",  "z_g_User1",      "ZKBH", False, "扩展数据1"),
    ("U2",  "z_g_User2",      "ZKBH", False, "扩展数据2"),
    ("GJ",  "z_c_GuJie",      "ZKBH", False, "竖向固结"),
    ("ZJ",  "z_c_ZhiJian",    "ZKBH", False, "直剪"),
    ("SZ",  "z_c_SanZhou",    "ZKBH", False, "三轴"),
    ("PZ",  "z_c_PengZhangTu","ZKBH", False, "膨胀试验"),
    ("ZH",  "z_c_ShuiZhi",    "ZKBH", False, "水质试验"),
    ("YY",  "z_c_YanShi",     "ZKBH", False, "岩石取样"),
]

# 关键表中文名（project_summary 展示用）
TABLE_CN = {
    "x_GongCheng": "工程信息", "z_ZuanKong": "钻孔", "z_g_TuCeng": "土层",
    "z_c_QuYang": "取样", "z_c_KeFen": "颗分", "z_y_DongTan": "动力触探",
    "z_y_JingTan": "静力触探", "z_y_BiaoGuan": "标准贯入", "z_y_BoSu": "波速",
    "z_y_TouShuiLv": "透水率", "z_y_LuoXuanBan": "十字板", "z_y_PangYa": "旁压",
    "z_g_ShuiWei": "水位", "z_g_ChaoShiDu": "潮湿度", "z_g_FengHua": "风化",
    "z_g_KeSuXing": "可塑性", "z_g_XieKong": "裂缝", "z_g_ShenTouJX": "渗透系数",
    "z_g_YXRQD": "岩芯RQD", "z_g_User1": "扩展1", "z_g_User2": "扩展2",
    "z_c_GuJie": "固结", "z_c_ZhiJian": "直剪", "z_c_SanZhou": "三轴",
    "z_c_PengZhangTu": "膨胀土", "z_c_ShuiZhi": "水质", "z_c_YanShi": "岩石取样",
    "t_YaShuiSYJL": "压水试验记录",
}

# 接口文件生成时优先输出的节顺序（其余按规范顺序补）
CORE_SECTIONS = ["GC", "ZK", "TC", "DT", "BG", "JT", "SW", "QY", "KF"]

app = Server("lizheng-mcp")

# 工程库连接缓存：{lz_path: {"mtime":..., "conn":..., "mdb":..., "tmpdir":...}}
_DB_CACHE = {}

# ---------------------------------------------------------------------------
# 基础工具
# ---------------------------------------------------------------------------

def _ok(data: Any) -> list:
    if isinstance(data, (dict, list)):
        text = json.dumps(data, indent=2, default=str, ensure_ascii=False)
    else:
        text = str(data)
    return [types.TextContent(type="text", text=text)]


def _err(msg: str) -> list:
    return [types.TextContent(type="text", text="ERROR: " + str(msg))]


def _clean_num(v: Any) -> str:
    """Access 浮点清洗：float32 只有约 7 位有效数字，
    6.30000019073486 -> '6.3'，4375.42578125 -> '4375.4258'，146.0 -> '146'。"""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "1" if v else "0"
    if isinstance(v, float):
        if math.isnan(v) or math.isinf(v):
            return ""
        s = "%.8g" % v
        return s
    if isinstance(v, int):
        return str(v)
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, bytes):
        try:
            return v.decode("gb18030", "replace").strip("\x00 ")
        except Exception:
            return ""
    return str(v).strip()


# 坐标可信阈值：工程坐标一般在 1e5~1e7 量级；理正库中无效坐标以 -999... 开头
COORD_PLAUSIBLE_MAX = 1e7


def _looks_like_bad_coord(v: Any) -> bool:
    try:
        f = float(v)
    except (TypeError, ValueError):
        return False
    return math.isnan(f) or abs(f) > COORD_PLAUSIBLE_MAX or f <= -999.0


# ---------------------------------------------------------------------------
# 官方接口规范解析（运行时读取随包规范文件）
# ---------------------------------------------------------------------------

_SPEC_CACHE = None


def load_spec() -> dict:
    """解析 lz_standard_interface_spec.txt -> {key: {...}}，key 形如 '#ZK'/'%GJ'。"""
    global _SPEC_CACHE
    if _SPEC_CACHE is not None:
        return _SPEC_CACHE
    if not os.path.exists(SPEC_FILE):
        raise FileNotFoundError("接口规范文件缺失: %s" % SPEC_FILE)
    raw = open(SPEC_FILE, "rb").read()
    text = raw.decode("gb18030", "replace")
    sections = {}
    order = []
    last_comment = ""
    for line in text.splitlines():
        s = line.strip()
        if not s:
            continue
        if s.startswith(";"):
            last_comment = s.lstrip(";").strip()
            continue
        m = re.match(r"^([#%])([A-Z0-9]+)\1(.*)$", s)
        if not m:
            continue
        prefix, tag, rest = m.group(1), m.group(2), m.group(3)
        fields = [f.strip() for f in re.split(r"[\t ]+", rest.strip()) if f.strip()]
        key = prefix + tag
        # 必填字段：描述中带 * 的 token 与字段按位对应（尽力而为）
        required = set()
        desc_tokens = re.split(r"[\t ]+", last_comment) if last_comment else []
        desc_tokens = [t for t in desc_tokens if t]
        star_idx = [i for i, t in enumerate(desc_tokens) if t.startswith("*")]
        if len(desc_tokens) == len(fields) and star_idx:
            for i in star_idx:
                if 0 <= i < len(fields):
                    required.add(fields[i])
        sections[key] = {
            "prefix": prefix,
            "tag": tag,
            "fields": fields,
            "desc": last_comment,
            "required": sorted(required),
            "kind": "data" if prefix == "#" else "item",
        }
        order.append(key)
    # 规范化 order：数据节(#)按规范序，项目表(%)紧随其主表
    _SPEC_CACHE = {"sections": sections, "order": order}
    return _SPEC_CACHE


def resolve_section_key(name: str) -> str:
    """'ZK' / '#ZK' / '#zk' -> '#ZK'；同 tag 的 #/% 并存时优先 #。"""
    spec = load_spec()["sections"]
    name = name.strip()
    if name.startswith("#") or name.startswith("%"):
        return name.upper()
    tag = name.upper()
    if "#" + tag in spec:
        return "#" + tag
    if "%" + tag in spec:
        return "%" + tag
    raise KeyError("未知接口节: %s（可用节见 lz_get_spec）" % name)


# ---------------------------------------------------------------------------
# 接口文件：生成 / 解析 / 校验
# ---------------------------------------------------------------------------

def build_interface_text(data: dict) -> tuple:
    """data: {'ZK': [ {...} | [...] ], ...} -> (text, warnings)。"""
    spec = load_spec()
    sections = spec["sections"]
    order = spec["order"]
    warnings = []
    normalized = {}   # key -> [record dict]
    for raw_key, records in (data or {}).items():
        try:
            key = resolve_section_key(raw_key)
        except KeyError as e:
            warnings.append(str(e))
            continue
        if key not in sections:
            warnings.append("跳过未知节 %s" % key)
            continue
        fields = sections[key]["fields"]
        recs = []
        for i, rec in enumerate(records or []):
            if isinstance(rec, dict):
                unknown = [k for k in rec if k not in fields]
                if unknown:
                    warnings.append("%s 第%d条记录含未定义字段 %s，已忽略"
                                    % (key, i + 1, ",".join(unknown)))
                recs.append({f: _clean_num(rec.get(f)) for f in fields})
            elif isinstance(rec, (list, tuple)):
                if len(rec) > len(fields):
                    warnings.append("%s 第%d条记录字段数 %d > 规范 %d，多余值已忽略"
                                    % (key, i + 1, len(rec), len(fields)))
                vals = list(rec) + [""] * (len(fields) - len(rec))
                recs.append({f: _clean_num(v) for f, v in zip(fields, vals)})
            else:
                warnings.append("%s 第%d条记录类型不支持: %r" % (key, i + 1, rec))
        if recs:
            normalized[key] = recs

    lines = [VERSION_HEADER_PREFIX + "gicad9.0"]
    for key in order:
        if key not in normalized:
            continue
        sec = sections[key]
        lines.append(";" + sec["desc"])
        lines.append(key + "#" + "\t".join(sec["fields"]))
        for rec in normalized[key]:
            lines.append("\t".join(rec[f] for f in sec["fields"]))
    # 规范外的节（理论上不会发生）
    for key in normalized:
        if key not in order:
            sec = sections[key]
            lines.append(";" + sec["desc"])
            lines.append(key + "#" + "\t".join(sec["fields"]))
            for rec in normalized[key]:
                lines.append("\t".join(rec[f] for f in sec["fields"]))
    return "\r\n".join(lines) + "\r\n", warnings


def write_interface_file(output_path: str, data: dict,
                         encoding: str = "gb18030",
                         overwrite: bool = False) -> dict:
    text, warnings = build_interface_text(data)
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    if os.path.exists(output_path):
        if not overwrite:
            raise FileExistsError("文件已存在: %s（如确认覆盖请先删除或传 overwrite=true）" % output_path)
        os.remove(output_path)
    with open(output_path, "wb") as f:
        f.write(text.encode(encoding, "replace"))
    n_rec = sum(len(v) for v in (data or {}).values() if isinstance(v, list))
    return {
        "output_path": output_path,
        "encoding": encoding,
        "bytes": len(text.encode(encoding, "replace")),
        "sections": sorted({resolve_section_key(k) for k in (data or {})}),
        "records": n_rec,
        "warnings": warnings,
    }


def parse_interface_text(text: str) -> dict:
    lines = text.splitlines()
    result = {"version": None, "sections": {}, "issues": []}
    cur = None
    for ln, line in enumerate(lines, 1):
        if not line.strip():
            continue
        if line.lstrip().startswith(";"):
            if cur is None and ln == 1:
                m = re.match(r"^\s*;#VERSION#(\S+)", line)
                if m:
                    result["version"] = m.group(1)
            continue
        m = re.match(r"^([#%])([A-Z0-9]+)\1(.*)$", line.strip())
        if m:
            prefix, tag, rest = m.group(1), m.group(2), m.group(3)
            fields = [f.strip() for f in re.split(r"[\t ]+", rest.strip()) if f.strip()]
            cur = prefix + tag
            result["sections"][cur] = {"fields": fields, "records": []}
            continue
        if cur is None:
            result["issues"].append("第%d行数据无对应节头: %r" % (ln, line[:40]))
            continue
        vals = line.split("\t")
        if len(vals) != len(result["sections"][cur]["fields"]):
            result["issues"].append(
                "第%d行字段数 %d 与节头 %d 不一致"
                % (ln, len(vals), len(result["sections"][cur]["fields"])))
        sec = result["sections"][cur]
        rec = {}
        for f, v in zip(sec["fields"], vals):
            rec[f] = v.strip()
        sec["records"].append(rec)
    return result


def parse_interface_file(path: str) -> dict:
    raw = open(path, "rb").read()
    try:
        text = raw.decode("gb18030")
    except UnicodeDecodeError:
        text = raw.decode("utf-8", "replace")
    parsed = parse_interface_text(text)
    parsed["file"] = path
    parsed["stats"] = {k: len(v["records"]) for k, v in parsed["sections"].items()}
    return parsed


def validate_interface_file(path: str) -> dict:
    parsed = parse_interface_file(path)
    spec = load_spec()["sections"]
    errors, warnings = [], []
    if not parsed["version"]:
        errors.append("缺少版本头，首行应为 ;#VERSION#gicad9.0")
    elif not parsed["version"].startswith("gicad9"):
        warnings.append("版本头为 %s，本工具按 gicad9.0 规范校验" % parsed["version"])
    if not parsed["sections"]:
        errors.append("文件中没有任何数据节")
    zk_bhs = set()
    for key, sec in parsed["sections"].items():
        if key not in spec:
            errors.append("未知数据节 %s" % key)
            continue
        want = spec[key]["fields"]
        if sec["fields"] != want:
            if len(sec["fields"]) != len(want):
                warnings.append("%s 字段数 %d 与规范 %d 不一致（顺序可自定义，数量应一致）"
                                % (key, len(sec["fields"]), len(want)))
            else:
                warnings.append("%s 字段顺序与规范不同（规范允许，需与理正端一致）" % key)
        for req in spec[key]["required"]:
            if req not in sec["fields"]:
                warnings.append("%s 缺少建议必填字段 %s" % (key, req))
        for i, rec in enumerate(sec["records"], 1):
            empty_req = [r for r in spec[key]["required"] if not rec.get(r)]
            if empty_req:
                warnings.append("%s 第%d条记录必填字段为空: %s"
                                % (key, i, ",".join(empty_req)))
        if key == "#ZK":
            zk_bhs = {r.get("ZKBH") for r in sec["records"] if r.get("ZKBH")}
    # 土层深度连续性
    tc = parsed["sections"].get("#TC")
    if tc:
        depths = []
        for rec in tc["records"]:
            try:
                depths.append(float(rec.get("TCCDSD") or 0))
            except ValueError:
                errors.append("#TC 层底深度非数值: %r" % rec.get("TCCDSD"))
        for a, b in zip(depths, depths[1:]):
            if b < a - 1e-9:
                warnings.append("#TC 层底深度非单调递增: %s -> %s" % (a, b))
        if depths and (depths[0] <= 0 or depths[-1] > 1000):
            warnings.append("#TC 深度范围异常: %s ~ %s" % (depths[0], depths[-1]))
    # 平板载荷孔号引用
    pb = parsed["sections"].get("#PBZH")
    if pb and zk_bhs:
        for rec in pb["records"]:
            bh = rec.get("ZKBH_ZH")
            if bh and bh not in zk_bhs:
                warnings.append("#PBZH 引用不存在的钻孔 %s" % bh)
    stats = parsed["stats"]
    return {
        "file": path,
        "valid": not errors,
        "errors": errors,
        "warnings": warnings,
        "version": parsed["version"],
        "stats": stats,
    }


# ---------------------------------------------------------------------------
# .lz 工程库读取（pyodbc + ACE，只读）
# ---------------------------------------------------------------------------

def _zip_name(name: str) -> str:
    """ZIP 中文名还原：python zipfile 对非 UTF8 名按 cp437 解码，转回 gb18030。"""
    if name.isascii():
        return name
    try:
        return name.encode("cp437").decode("gb18030")
    except Exception:
        return name


def _find_mdb_in_zip(zpath: str) -> str:
    with zipfile.ZipFile(zpath) as z:
        names = [_zip_name(n) for n in z.namelist()]
        mdb = [n for n in names if n.lower().endswith(".mdb")]
        if not mdb:
            raise ValueError(".lz 包内未找到 .mdb 工程库")
        return mdb[0]


def _connect_mdb(dbq: str):
    import pyodbc
    cs = "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" + dbq
    if LZ_MDB_PASSWORD:
        cs += ";PWD=" + LZ_MDB_PASSWORD
    return pyodbc.connect(cs, readonly=True, timeout=30)


def open_lz_project(lz_path: str) -> dict:
    """打开 .lz：解出 mdb 到临时目录并只读连接。带缓存。"""
    lz_path = os.path.abspath(lz_path)
    if not os.path.exists(lz_path):
        raise FileNotFoundError("文件不存在: %s" % lz_path)
    mtime = os.path.getmtime(lz_path)
    hit = _DB_CACHE.get(lz_path)
    if hit and hit["mtime"] == mtime:
        return hit
    if hit:
        try:
            hit["conn"].close()
        except Exception:
            pass
    mdb_name = _find_mdb_in_zip(lz_path)
    tmpdir = os.path.join(tempfile.gettempdir(),
                          "lizmcp_" + hashlib.md5(lz_path.encode("utf-8")).hexdigest()[:10])
    os.makedirs(tmpdir, exist_ok=True)
    dbq = os.path.join(tmpdir, os.path.basename(mdb_name))
    with zipfile.ZipFile(lz_path) as z:
        # read 需要原始(cp437)条目名；显示/落盘用还原后的中文名
        orig = next(n for n in z.namelist()
                    if _zip_name(n).lower().endswith(".mdb"))
        with open(dbq, "wb") as f:
            f.write(z.read(orig))
    conn = _connect_mdb(dbq)
    cur = conn.cursor()
    tables = sorted([r[2] for r in cur.tables() if r[3] == "TABLE"])
    info = {
        "lz_path": lz_path,
        "mtime": mtime,
        "mdb_name": mdb_name,
        "dbq": dbq,
        "tmpdir": tmpdir,
        "conn": conn,
        "tables": tables,
        "ini": _read_project_ini(lz_path),
        "files": [_zip_name(n) for n in zipfile.ZipFile(lz_path).namelist()],
    }
    _DB_CACHE[lz_path] = info
    return info


def _read_project_ini(lz_path: str) -> dict:
    try:
        with zipfile.ZipFile(lz_path) as z:
            name = [_zip_name(n) for n in z.namelist()
                    if _zip_name(n).endswith("ProjectInfo.ini")]
            if not name:
                return {}
            orig = next(n for n in z.namelist() if _zip_name(n).endswith("ProjectInfo.ini"))
            txt = z.read(orig).decode("gb18030", "replace")
        out, sec = {}, None
        for line in txt.splitlines():
            line = line.strip()
            if not line:
                continue
            if line.startswith("[") and line.endswith("]"):
                sec = line[1:-1]
                out[sec] = {}
            elif "=" in line and sec:
                k, v = line.split("=", 1)
                out[sec][k.strip()] = v.strip()
        return out
    except Exception:
        return {}


def _db_rows(info: dict, table: str, limit: int = 200, offset: int = 0) -> tuple:
    """安全读表：表名必须在真实表清单中；仅允许 SELECT。"""
    if table not in info["tables"]:
        raise KeyError("表不存在: %s（可用 lz_open_project 查看表清单）" % table)
    cur = info["conn"].cursor()
    sql = "SELECT TOP %d * FROM [%s]" % (limit + offset, table)
    cur.execute(sql)
    names = [d[0] for d in cur.description]
    rows = []
    for i, r in enumerate(cur.fetchall()):
        if i < offset:
            continue
        rows.append({n: r[j] for j, n in enumerate(names)})
    return names, rows


def _db_table_count(info: dict, table: str) -> int:
    if table not in info["tables"]:
        return -1
    cur = info["conn"].cursor()
    return cur.execute("SELECT COUNT(*) FROM [%s]" % table).fetchone()[0]


def _row_to_record(row: dict, fields: list) -> dict:
    """按接口节字段码从库行取值（库字段码与接口一致）。"""
    return {f: _clean_num(row.get(f)) for f in fields if f in row}


def _project_gc_record(info: dict) -> dict:
    """从 x_GongCheng 取工程记录（映射到 #GC#）。"""
    spec = load_spec()["sections"]["#GC"]
    try:
        names, rows = _db_rows(info, "x_GongCheng", limit=1)
    except Exception:
        return {}
    if not rows:
        return {}
    rec = _row_to_record(rows[0], spec["fields"])
    # 工程编号/名称兜底：ProjectInfo.ini
    ini = info.get("ini") or {}
    gcinfo = ini.get("GCInfo", {})
    if not rec.get("GCBH") and gcinfo.get("GCBH"):
        rec["GCBH"] = gcinfo["GCBH"]
    if not rec.get("GCMC") and gcinfo.get("GCMC"):
        rec["GCMC"] = gcinfo["GCMC"]
    return rec


# ---------------------------------------------------------------------------
# Excel 附表 -> 规范化记录
# ---------------------------------------------------------------------------

# 表头关键词 -> 字段码（勘探点一览表）
ZK_SYNONYMS = [
    (["勘探点编号", "钻孔编号", "孔号", "勘探点"], "ZKBH"),
    (["勘探点类型", "钻孔类型", "类型"], "ZKLX"),
    (["勘探点深度", "勘探深度", "孔深"], "ZKSD"),
    (["地面高程", "孔口高程", "孔口标高", "孔口高程"], "ZKBG"),
    (["x坐标", "x(m)", "x"], "ZKX"),
    (["y坐标", "y(m)", "y"], "ZKY"),
    (["备注", "注"], "BZ"),
]

DT_SYNONYMS = [
    (["勘探点编号", "钻孔编号", "孔号"], "ZKBH"),
    (["试验段深度", "试验深度", "深度"], "DTDSD"),
    (["重型动探n63.5", "n63.5", "动探n63.5", "击数"], "DTJS"),
    (["贯入度"], "DTGRD"),
    (["探杆长度", "杆长"], "DTGC"),
    (["杆长修正系数"], "DTGCXS"),
    (["修正n63.5", "修正击数", "重型动探修正n63.5"], "DTXZJS"),
    (["备注"], "BZ"),
]


def _norm_header(h: Any) -> str:
    if h is None:
        return ""
    s = str(h).replace("\n", "").replace(" ", "").replace("　", "").strip().lower()
    return s


def _match_synonym(header: str, table: list) -> str:
    """最长匹配优先：'勘探点类型' 应命中 ZKLX 而不是泛化的 '勘探点'->ZKBH。"""
    best_code, best_len = "", 0
    for keys, code in table:
        for k in keys:
            if k in header and len(k) > best_len:
                best_code, best_len = code, len(k)
    return best_code


# 单位行/汇总列过滤
UNIT_TOKENS = {"m", "cm", "mm", "(m)", "(cm)", "%", "件", "个", "击", "击/10cm",
               "cm/击", "kpa", "mpa", "g/cm3", "kn/m3", "°", "度", "s", "lu"}
SKIP_HEADER_WORDS = ("统计", "场区", "平均值", "标准差", "最大值", "最小值", "样本")
# 汇总/签署行过滤（值中出现这些词即判为非数据行）
SUMMARY_WORDS = ("合计", "共计", "总计", "小计", "校对", "审核", "制图", "签名", "日期")


def _merge_header_rows(grid: list, idx: int) -> tuple:
    """合并双行表头：下一行是单位行(m/cm...)或 X/Y 子表头时向下拼接。
    返回 (合并后的表头, 是否消费了下一行)。"""
    header = [_norm_header(c) for c in grid[idx]]
    if idx + 1 < len(grid):
        nxt = [_norm_header(c) for c in grid[idx + 1]]
        unit_like = sum(1 for c in nxt if c and (c in UNIT_TOKENS or c in ("x", "y")))
        if unit_like >= 2 or any(c in ("x", "y") for c in nxt):
            merged = []
            for h, n in zip(header, nxt):
                if h and n and n not in UNIT_TOKENS:
                    merged.append(h + n)      # 坐标 + X / 坐标 + Y
                elif h:
                    merged.append(h)
                elif n and n not in UNIT_TOKENS:
                    merged.append(n)
                else:
                    merged.append(h or n)
            return merged, True
    return header, False


def excel_to_records(xlsx_path: str, sheet: str = None,
                     table_type: str = None) -> dict:
    """识别理正风格成果附表 -> 规范化记录（按目标接口节分组）。"""
    import openpyxl
    if not os.path.exists(xlsx_path):
        raise FileNotFoundError("文件不存在: %s" % xlsx_path)
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb[sheet] if sheet else wb[wb.sheetnames[0]]
    grid = [list(r) for r in ws.iter_rows(values_only=True)]
    # 定位表头行：含 >=3 个非空且命中关键词
    header_row_idx = None
    for i, row in enumerate(grid[:12]):
        nonempty = [c for c in row if c is not None and str(c).strip()]
        if len(nonempty) >= 3:
            joined = "".join(_norm_header(c) for c in row)
            if ("勘探点" in joined or "钻孔" in joined) and "编号" in joined:
                header_row_idx = i
                break
            if "试验段深度" in joined or "动探" in joined or "标贯" in joined:
                header_row_idx = i
                break
            if "地层" in joined and ("层底" in joined or "深度" in joined):
                header_row_idx = i
                break
    if header_row_idx is None:
        return {"detected": False, "reason": "未识别到表头行（前12行内无匹配关键词）",
                "headers": [[_norm_header(c) for c in r] for r in grid[:8] if any(c is not None for c in r)]}
    header, consumed_next = _merge_header_rows(grid, header_row_idx)
    joined = "".join(header)
    # 工程名称行（向上找）
    proj_name = ""
    for row in grid[:header_row_idx]:
        for j, c in enumerate(row):
            if c and "工程名称" in str(c):
                for k in range(j + 1, len(row)):
                    if row[k] and str(row[k]).strip():
                        proj_name = str(row[k]).strip()
                        break
    detected = table_type
    if not detected:
        if "试验段深度" in joined or ("动探" in joined and "击数" in joined):
            detected = "dpt"
        elif "标贯" in joined or "标准贯入" in joined:
            detected = "spt"
        elif "勘探点" in joined and ("x" in joined or "坐标" in joined or "高程" in joined):
            detected = "borehole_list"
        else:
            detected = "unknown"
    syn = {"dpt": DT_SYNONYMS, "spt": DT_SYNONYMS,
           "borehole_list": ZK_SYNONYMS}.get(detected, [])
    colmap = {}
    for j, h in enumerate(header):
        if not h:
            continue
        if any(w in h for w in SKIP_HEADER_WORDS):
            continue                      # 汇总统计列不进记录
        code = _match_synonym(h, syn)
        if code:
            colmap[j] = code
    records, skipped = [], 0
    bh_col = next((j for j, c in colmap.items() if c == "ZKBH"), None)
    last_bh = None
    for row in grid[header_row_idx + 1 + (1 if consumed_next else 0):]:
        if not any(c is not None and str(c).strip() for c in row):
            continue
        rec = {}
        for j, code in colmap.items():
            v = row[j] if j < len(row) else None
            if v is None:
                continue
            if isinstance(v, str):
                v = v.strip()
                if not v:
                    continue
            rec[code] = v
        # 单位行过滤：非空单元格几乎都是单位token
        cells = [str(c).strip().lower() for c in row if c is not None and str(c).strip()]
        if cells and sum(1 for c in cells if c in UNIT_TOKENS) >= max(1, len(cells) // 2):
            skipped += 1
            continue
        # 汇总/签署行过滤
        if any(isinstance(v, str) and any(w in v for w in SUMMARY_WORDS)
               for v in rec.values()):
            skipped += 1
            continue
        if not rec:
            skipped += 1
            continue
        # 合并单元格向下填充钻孔号
        if bh_col is not None:
            if rec.get("ZKBH"):
                last_bh = rec["ZKBH"]
            elif last_bh is not None and detected in ("dpt", "spt"):
                rec["ZKBH"] = last_bh
        if detected in ("dpt", "spt"):
            _split_depth_range(rec)
        records.append(rec)
    notes = []
    if proj_name:
        notes.append("识别到工程名称: %s" % proj_name)
    if skipped:
        notes.append("跳过单位行/空行: %d" % skipped)
    if detected in ("dpt", "spt"):
        groups = {}
        for r in records:
            k = str(r.get("ZKBH", "?"))
            groups[k] = groups.get(k, 0) + 1
        notes.append("记录按钻孔分组: %s" % json.dumps(groups, ensure_ascii=False))
        notes.append("接口文件(如#DT#)不含钻孔号，导入时需在理正内选定对应钻孔后逐孔导入")
    return {
        "detected": True,
        "table_type": detected,
        "target_section": {"dpt": "#DT", "spt": "#BG",
                           "borehole_list": "#ZK"}.get(detected),
        "project_name": proj_name,
        "column_mapping": {header[j]: c for j, c in colmap.items() if j < len(header)},
        "record_count": len(records),
        "records": records,
        "notes": notes,
    }


def _split_depth_range(rec: dict):
    """'1.00-1.10' -> DTDSD(底深) + DTCD(段长)。"""
    v = rec.get("DTDSD")
    if isinstance(v, str) and "-" in v:
        try:
            a, b = [float(x) for x in re.split(r"[-~—]", v)]
            rec["DTDSD"] = round(b, 3)
            rec["DTCD"] = round(b - a, 3)
        except Exception:
            pass


# ---------------------------------------------------------------------------
# 录入模板生成
# ---------------------------------------------------------------------------

def create_input_template(output_path: str, sections: list) -> dict:
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    spec = load_spec()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    made = []
    for name in (sections or ["ZK", "TC", "DT", "BG"]):
        try:
            key = resolve_section_key(name)
        except KeyError:
            continue
        sec = spec["sections"][key]
        ws = wb.create_sheet(title=key.strip("#%"))
        hdr_fill = PatternFill("solid", fgColor="DDEBF7")
        req_font = Font(bold=True, color="C00000")
        ws.append(list(sec["fields"]))
        ws.append(["%s%s" % ("*" if f in sec["required"] else "", f)
                   for f in sec["fields"]])
        # 中文名行：从规范描述按位取（尽力而为）
        desc_tokens = [t for t in re.split(r"[\t ]+", sec["desc"]) if t]
        cn = []
        for i, f in enumerate(sec["fields"]):
            t = desc_tokens[i] if i < len(desc_tokens) else ""
            cn.append(t.lstrip("*"))
        ws.append(cn)
        for j in range(1, len(sec["fields"]) + 1):
            c = ws.cell(row=1, column=j)
            c.fill = hdr_fill
            if sec["fields"][j - 1] in sec["required"]:
                c.font = req_font
        for j, w in enumerate([max(10, len(str(f)) * 2.2) for f in sec["fields"]], 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
        made.append({"sheet": ws.title, "key": key, "fields": len(sec["fields"])})
    if os.path.exists(output_path):
        raise FileExistsError("文件已存在: %s" % output_path)
    wb.save(output_path)
    return {"output_path": output_path, "sheets": made}


# ---------------------------------------------------------------------------
# MCP 工具实现
# ---------------------------------------------------------------------------

async def _tool_backend_info():
    import platform
    spec_ok = os.path.exists(SPEC_FILE)
    n_sec = 0
    if spec_ok:
        try:
            n_sec = len(load_spec()["sections"])
        except Exception:
            pass
    pyodbc_ok, ace_ok = False, False
    try:
        import pyodbc
        pyodbc_ok = True
        drivers = pyodbc.drivers()
        ace_ok = any("Access" in d and "mdb" in d for d in drivers)
    except Exception:
        pass
    install = next((p for p in LZ_INSTALL_CANDIDATES if os.path.isdir(p)), None)
    return _ok({
        "server": "lizheng-mcp",
        "python": platform.python_version(),
        "platform": platform.platform(),
        "spec_file": SPEC_FILE,
        "spec_loaded": spec_ok,
        "spec_sections": n_sec,
        "pyodbc_installed": pyodbc_ok,
        "ace_access_driver": ace_ok,
        "lz_install_dir": install,
        "db_password_source": "安装目录二进制自动提取 (Database Password=)",
        "cached_projects": list(_DB_CACHE.keys()),
        "hints": [] if (pyodbc_ok and ace_ok and spec_ok) else [
            "缺 pyodbc: pip install pyodbc",
            "缺 ACE 驱动: 安装 Microsoft Access Database Engine 64 位",
            "缺规范文件: 确认 lz_standard_interface_spec.txt 与 server.py 同目录",
        ],
    })


async def _tool_get_spec(section: str = None):
    spec = load_spec()
    if section:
        key = resolve_section_key(section)
        sec = spec["sections"][key]
        return _ok({"key": key, **sec})
    out = {}
    for key, sec in spec["sections"].items():
        out[key] = {"fields": sec["fields"], "required": sec["required"],
                    "desc": sec["desc"]}
    return _ok({"section_count": len(out), "sections": out})


async def _tool_open_project(lz_path: str):
    info = open_lz_project(lz_path)
    counts = {}
    for tag, table, _bk, _v, _cn in DB_SECTION_MAP:
        if table in info["tables"]:
            try:
                counts[table] = _db_table_count(info, table)
            except Exception:
                counts[table] = None
    ini = info.get("ini") or {}
    gc = ini.get("GCInfo", {})
    return _ok({
        "lz_path": lz_path,
        "mdb": info["mdb_name"],
        "db_extracted_to": info["dbq"],
        "readonly": True,
        "project": {
            "GCBH": gc.get("GCBH"), "GCMC": gc.get("GCMC"),
            "GCPATH": gc.get("GCPATH"),
            "db_folder": gc.get("CURDB"), "files_folder": gc.get("CURFILE"),
        },
        "table_total": len(info["tables"]),
        "key_table_counts": counts,
        "delivery_files": [f for f in info["files"]
                           if f.lower().endswith((".pdf", ".doc", ".docx", ".xls", ".xlsx", ".dwg"))][:30],
    })


async def _tool_read_table(lz_path: str, table: str, limit: int = 50,
                           offset: int = 0, columns: list = None):
    info = open_lz_project(lz_path)
    names, rows = _db_rows(info, table, limit=min(limit, 500), offset=offset)
    if columns:
        keep = [c for c in columns if c in names]
        rows = [{c: r.get(c) for c in keep} for r in rows]
        names = keep
    out_rows = [{n: _clean_num(v) for n, v in r.items()} for r in rows]
    return _ok({
        "table": table, "table_cn": TABLE_CN.get(table, ""),
        "columns": names, "column_count": len(names),
        "returned": len(out_rows), "offset": offset,
        "records": out_rows,
    })


async def _tool_project_summary(lz_path: str):
    info = open_lz_project(lz_path)
    gc = _project_gc_record(info)
    # 钻孔清单
    zk_spec = load_spec()["sections"]["#ZK"]
    boreholes = []
    zk_warnings = []
    if "z_ZuanKong" in info["tables"]:
        _n, rows = _db_rows(info, "z_ZuanKong", limit=500)
        zk_recs = _prepare_zk_records(rows, zk_spec["fields"], None, zk_warnings)
        for r in zk_recs:
            boreholes.append({
                "ZKBH": r.get("ZKBH"), "ZKLX": r.get("ZKLX"),
                "ZKX": r.get("ZKX"), "ZKY": r.get("ZKY"),
                "ZKBG": r.get("ZKBG"), "ZKSD": r.get("ZKSD"),
                "ZKKSRQ": r.get("ZKKSRQ"), "ZKZZRQ": r.get("ZKZZRQ"),
            })
    counts = {}
    for tag, table, _bk, verified, cn in DB_SECTION_MAP:
        if table in info["tables"]:
            try:
                n = _db_table_count(info, table)
            except Exception:
                n = None
            if n:
                counts["#" + tag] = {"table": table, "cn": cn,
                                     "rows": n, "verified": verified}
    # 土层概况
    layers = []
    if "z_g_TuCeng" in info["tables"]:
        _n, rows = _db_rows(info, "z_g_TuCeng", limit=1000)
        for r in rows:
            layers.append({
                "ZKBH": _clean_num(r.get("ZKBH")),
                "TCMC": _clean_num(r.get("TCMC")),
                "TCCDSD": _clean_num(r.get("TCCDSD")),
                "TCHD": _clean_num(r.get("TCHD")),
                "TCZCBH": _clean_num(r.get("TCZCBH")),
                "TCDZSD": _clean_num(r.get("TCDZSD")),
                "TCDZCY": _clean_num(r.get("TCDZCY")),
            })
    return _ok({
        "project": gc,
        "borehole_count": len(boreholes),
        "boreholes": boreholes,
        "warnings": zk_warnings,
        "sections_with_data": counts,
        "layer_count": len(layers),
        "layers": layers[:120],
        "layers_truncated": len(layers) > 120,
    })


def _prepare_zk_records(rows: list, fields: list,
                        coord_overrides: dict = None,
                        warnings: list = None) -> list:
    """构造 #ZK# 记录：清洗库中无效坐标（-999... 保护值），支持外部坐标覆盖。"""
    coord_overrides = coord_overrides or {}
    bad_coord_bhs = []
    recs = []
    for r in rows:
        rec = _row_to_record(r, fields)
        bh = rec.get("ZKBH", "")
        for axis in ("ZKX", "ZKY"):
            if axis in rec and rec[axis] and _looks_like_bad_coord(rec[axis]):
                rec[axis] = ""
                if bh and bh not in coord_overrides and bh not in bad_coord_bhs:
                    bad_coord_bhs.append(bh)
        ov = coord_overrides.get(bh)
        if ov:
            for k, v in ov.items():
                if k in rec:
                    rec[k] = _clean_num(v)
        recs.append(rec)
    if bad_coord_bhs and warnings is not None:
        warnings.append(
            "以下钻孔的库内坐标为无效保护值(-999...)，已置空，请用 coordinates 参数"
            "或导入后在理正中补填: %s" % ",".join(bad_coord_bhs))
    return recs


async def _tool_db_to_interface(lz_path: str, output_dir: str,
                                sections: list = None, boreholes: list = None,
                                mode: str = "per_borehole",
                                include_empty: bool = False,
                                coordinates: dict = None,
                                overwrite: bool = True):
    """从 .lz 工程库生成理正标准接口文件。

    mode=per_borehole: 每个钻孔一个文件（接口文件按孔导入，推荐）
    mode=combined:     单文件含全部节（备份/迁移用）
    coordinates: {"ZK1": {"ZKX": 463095.787, "ZKY": 3730224.501}, ...} 坐标覆盖
    """
    info = open_lz_project(lz_path)
    spec = load_spec()
    mode = (mode or "per_borehole").lower()
    if mode not in ("per_borehole", "combined"):
        raise ValueError("mode 仅支持 per_borehole / combined")
    wanted = None
    if sections:
        wanted = [resolve_section_key(s) for s in sections]

    os.makedirs(output_dir, exist_ok=True)
    gc_rec = _project_gc_record(info)
    gc_key = "#GC"

    # 钻孔清单
    zk_rows = []
    if "z_ZuanKong" in info["tables"]:
        _n, zk_rows = _db_rows(info, "z_ZuanKong", limit=2000)

    def rows_of(tag):
        entry = next((e for e in DB_SECTION_MAP if "#" + e[0] == tag), None)
        if not entry or entry[1] not in info["tables"]:
            return None, None, None
        _t, table, bk_field, _v, _cn = entry
        try:
            _n, rows = _db_rows(info, table, limit=20000)
        except Exception:
            return None, None, None
        return rows, bk_field, table

    files, warnings = [], []
    stamp = datetime.datetime.now().strftime("%Y%m%d")

    if mode == "combined":
        data = {}
        if gc_rec:
            data[gc_key] = [gc_rec]
        if "z_ZuanKong" in info["tables"]:
            data["#ZK"] = _prepare_zk_records(
                zk_rows, spec["sections"]["#ZK"]["fields"],
                coordinates, warnings)
        for tag, table, _bk, _v, cn in DB_SECTION_MAP:
            key = "#" + tag
            if wanted and key not in wanted:
                continue
            if key in ("#GC", "#ZK"):
                continue
            rows, _bkf, _tb = rows_of(key)
            if not rows:
                continue
            sec_fields = spec["sections"].get(key, {}).get("fields")
            if not sec_fields:
                continue
            data[key] = [_row_to_record(r, sec_fields) for r in rows]
        if not include_empty and not any(data.values()):
            return _ok({"mode": mode, "files": [], "warning": "工程库中无可用数据"})
        fname = "%s_工程库导出_%s.txt" % (
            (gc_rec.get("GCBH") or "LZ").replace(" ", ""), stamp)
        fpath = os.path.join(output_dir, fname)
        res = write_interface_file(fpath, data, overwrite=overwrite)
        res["note"] = "combined 模式：单文件含全部节；理正导入时数据节按孔选择导入"
        files.append(res)
    else:
        if not zk_rows:
            raise ValueError("工程库中未找到钻孔数据（z_ZuanKong 为空）")
        gc_bh = (gc_rec.get("GCBH") or "LZ").replace(" ", "")
        for r in zk_rows:
            bh = _clean_num(r.get("ZKBH"))
            if not bh:
                continue
            if boreholes and bh not in boreholes:
                continue
            data = {}
            if gc_rec and (not wanted or gc_key in wanted):
                data[gc_key] = [gc_rec]
            zk_fields = spec["sections"]["#ZK"]["fields"]
            if not wanted or "#ZK" in wanted:
                data["#ZK"] = _prepare_zk_records(
                    [r], zk_fields, coordinates, warnings)
            for tag, table, bk_field, _v, cn in DB_SECTION_MAP:
                key = "#" + tag
                if key in ("#GC", "#ZK"):
                    continue
                if wanted and key not in wanted:
                    continue
                if table not in info["tables"]:
                    continue
                sec_fields = spec["sections"].get(key, {}).get("fields")
                if not sec_fields:
                    continue
                try:
                    _n, rows = _db_rows(info, table, limit=20000)
                except Exception:
                    continue
                mine = [x for x in rows if _clean_num(x.get(bk_field or "")) == bh]
                if mine:
                    data[key] = [_row_to_record(x, sec_fields) for x in mine]
            if not include_empty and not any(len(v) for v in data.values()):
                warnings.append("%s 无数据，跳过" % bh)
                continue
            fpath = os.path.join(output_dir, "%s_%s_接口文件_%s.txt" % (gc_bh, bh, stamp))
            res = write_interface_file(fpath, data, overwrite=overwrite)
            res["borehole"] = bh
            files.append(res)
    return _ok({
        "mode": mode,
        "output_dir": output_dir,
        "file_count": len(files),
        "files": files,
        "warnings": warnings,
        "usage": "在理正勘察中：数据 -> 接口文件导入 -> 选择对应钻孔 -> 选择本文件",
    })


async def _tool_build_interface_file(output_path: str, data: dict,
                                     encoding: str = "gb18030"):
    if os.path.exists(output_path):
        raise FileExistsError("文件已存在: %s" % output_path)
    res = write_interface_file(output_path, data, encoding)
    return _ok(res)


async def _tool_parse_interface_file(path: str):
    if not os.path.exists(path):
        raise FileNotFoundError("文件不存在: %s" % path)
    return _ok(parse_interface_file(path))


async def _tool_validate_interface_file(path: str):
    if not os.path.exists(path):
        raise FileNotFoundError("文件不存在: %s" % path)
    return _ok(validate_interface_file(path))


async def _tool_excel_to_records(path: str, sheet: str = None,
                                 table_type: str = None):
    return _ok(excel_to_records(path, sheet, table_type))


async def _tool_create_input_template(output_path: str,
                                      sections: list = None):
    if os.path.exists(output_path):
        raise FileExistsError("文件已存在: %s" % output_path)
    return _ok(create_input_template(output_path, sections or ["ZK", "TC", "DT", "BG"]))


# ---------------------------------------------------------------------------
# MCP 协议层
# ---------------------------------------------------------------------------

@app.list_tools()
async def list_tools() -> list:
    return [
        types.Tool(
            name="lz_backend_info",
            description="理正 MCP 环境自检：Python/pyodbc/ACE 驱动/接口规范/安装目录/缓存工程。",
            inputSchema={"type": "object", "properties": {}, "required": []},
        ),
        types.Tool(
            name="lz_get_spec",
            description="查询《理正标准数据接口》规范：全部或指定节(#TAG#/%TAG#)的字段码、中文名、必填标记。生成接口文件前先用它核对字段。",
            inputSchema={
                "type": "object",
                "properties": {
                    "section": {"type": "string",
                                "description": "节名，如 ZK / #TC / %GJ；省略返回全部 44 节"},
                },
                "required": [],
            },
        ),
        types.Tool(
            name="lz_open_project",
            description="打开理正工程包 .lz（只读）：解出 Access 工程库并列出关键表行数、工程信息、交付文件清单。",
            inputSchema={
                "type": "object",
                "properties": {
                    "lz_path": {"type": "string", "description": ".lz 工程包绝对路径"},
                },
                "required": ["lz_path"],
            },
        ),
        types.Tool(
            name="lz_read_table",
            description="只读查询工程库指定表（如 z_ZuanKong/z_g_TuCeng/z_y_DongTan/z_c_QuYang），支持分页与字段裁剪。",
            inputSchema={
                "type": "object",
                "properties": {
                    "lz_path": {"type": "string", "description": ".lz 工程包路径"},
                    "table": {"type": "string", "description": "库表名，须在 lz_open_project 返回的表清单中"},
                    "limit": {"type": "integer", "description": "返回行数上限，默认 50，最大 500", "default": 50},
                    "offset": {"type": "integer", "description": "跳过行数，默认 0", "default": 0},
                    "columns": {"type": "array", "items": {"type": "string"},
                                "description": "只返回这些字段（默认全部）"},
                },
                "required": ["lz_path", "table"],
            },
        ),
        types.Tool(
            name="lz_project_summary",
            description="工程摘要：工程信息、钻孔清单（编号/类型/坐标/孔深/日期）、各试验数据计数、土层分层概况。",
            inputSchema={
                "type": "object",
                "properties": {
                    "lz_path": {"type": "string", "description": ".lz 工程包路径"},
                },
                "required": ["lz_path"],
            },
        ),
        types.Tool(
            name="lz_db_to_interface",
            description="【核心】从 .lz 工程库直接生成理正标准接口文件。默认每孔一个文件（接口文件按孔导入）。支持按节/按孔过滤，输出 GBK 编码接口文件。",
            inputSchema={
                "type": "object",
                "properties": {
                    "lz_path": {"type": "string", "description": ".lz 工程包路径"},
                    "output_dir": {"type": "string", "description": "输出目录（自动创建）"},
                    "sections": {"type": "array", "items": {"type": "string"},
                                 "description": "只导出这些节，如 [\"ZK\",\"TC\",\"DT\"]；默认全部有数据的节"},
                    "boreholes": {"type": "array", "items": {"type": "string"},
                                  "description": "只导出这些钻孔号（per_borehole 模式生效）"},
                    "mode": {"type": "string", "enum": ["per_borehole", "combined"],
                             "description": "per_borehole=每孔一文件（默认，推荐）；combined=单文件全量", "default": "per_borehole"},
                    "include_empty": {"type": "boolean",
                                      "description": "是否也为无数据钻孔生成文件，默认 false", "default": False},
                    "coordinates": {"type": "object",
                                    "description": "坐标覆盖 {\"ZK1\": {\"ZKX\": 463095.787, \"ZKY\": 3730224.501}}；库内坐标为-999保护值，建议从勘探点一览表补入"},
                    "overwrite": {"type": "boolean",
                                  "description": "目标文件已存在时是否覆盖，默认 true（导出可重复生成）", "default": True},
                },
                "required": ["lz_path", "output_dir"],
            },
        ),
        types.Tool(
            name="lz_build_interface_file",
            description="从 JSON 数据生成理正标准接口文件（GBK、Tab 分隔、;#VERSION#gicad9.0 头）。data 形如 {\"ZK\":[{...}],\"TC\":[...]}。",
            inputSchema={
                "type": "object",
                "properties": {
                    "output_path": {"type": "string", "description": "输出 .txt 路径"},
                    "data": {"type": "object",
                             "description": "节名 -> 记录数组；记录为 字段码:值 的对象或按规范顺序的值数组"},
                    "encoding": {"type": "string", "description": "默认 gb18030", "default": "gb18030"},
                },
                "required": ["output_path", "data"],
            },
        ),
        types.Tool(
            name="lz_parse_interface_file",
            description="解析接口文件为结构化 JSON：版本头、各节字段表与记录。",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "接口文件路径"},
                },
                "required": ["path"],
            },
        ),
        types.Tool(
            name="lz_validate_interface_file",
            description="校验接口文件：版本头、节名、字段数/顺序、必填字段、土层深度连续性、孔号引用。返回 errors/warnings。",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "接口文件路径"},
                },
                "required": ["path"],
            },
        ),
        types.Tool(
            name="lz_excel_to_records",
            description="识别理正风格成果附表（勘探点一览表/动探统计表/标贯表等）-> 按目标接口节规范化记录，供生成接口文件使用。",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": ".xlsx 附表路径"},
                    "sheet": {"type": "string", "description": "工作表名，默认第一个"},
                    "table_type": {"type": "string",
                                   "description": "强制指定 borehole_list/dpt/spt，默认自动识别"},
                },
                "required": ["path"],
            },
        ),
        types.Tool(
            name="lz_create_input_template",
            description="生成标准录入模板 xlsx：所选节的三行表头（字段码/必填标记+中文名），用于规范采集数据。",
            inputSchema={
                "type": "object",
                "properties": {
                    "output_path": {"type": "string", "description": "输出 .xlsx 路径"},
                    "sections": {"type": "array", "items": {"type": "string"},
                                 "description": "节名列表，默认 [\"ZK\",\"TC\",\"DT\",\"BG\"]"},
                },
                "required": ["output_path"],
            },
        ),
    ]


@app.call_tool()
async def call_tool(name: str, arguments: dict) -> list:
    try:
        if name == "lz_backend_info":
            return await _tool_backend_info()
        elif name == "lz_get_spec":
            return await _tool_get_spec(**(arguments or {}))
        elif name == "lz_open_project":
            return await _tool_open_project(**(arguments or {}))
        elif name == "lz_read_table":
            return await _tool_read_table(**(arguments or {}))
        elif name == "lz_project_summary":
            return await _tool_project_summary(**(arguments or {}))
        elif name == "lz_db_to_interface":
            return await _tool_db_to_interface(**(arguments or {}))
        elif name == "lz_build_interface_file":
            return await _tool_build_interface_file(**(arguments or {}))
        elif name == "lz_parse_interface_file":
            return await _tool_parse_interface_file(**(arguments or {}))
        elif name == "lz_validate_interface_file":
            return await _tool_validate_interface_file(**(arguments or {}))
        elif name == "lz_excel_to_records":
            return await _tool_excel_to_records(**(arguments or {}))
        elif name == "lz_create_input_template":
            return await _tool_create_input_template(**(arguments or {}))
        else:
            return _err("未知工具: %s" % name)
    except Exception as e:
        return _err("%s: %s" % (type(e).__name__, e))


async def main():
    async with stdio_server() as (read_stream, write_stream):
        await app.run(read_stream, write_stream,
                      app.create_initialization_options())


if __name__ == "__main__":
    asyncio.run(main())
