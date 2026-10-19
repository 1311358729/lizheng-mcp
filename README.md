# lizheng-mcp — 理正勘察 9.5 接口文件 MCP Server

面向**理正工程地质勘察软件 9.5PB3** 的只读自动化通道。把理正的工程数据
（.lz 工程库）和官方《标准数据接口》文件接进 AI 工作流：

- **读工程库**：打开 .lz → 查询工程摘要 / 任意表 / 钻孔清单
- **导接口文件**：从 .lz 一键批量生成理正标准数据接口文件（默认每孔一个文件），可直接在理正里导入
- **接口文件工具链**：规范速查 / 生成 / 解析 / 校验（版本头、字段数、必填项、深度连续性）
- **Excel 附表识别**：自动识别《勘探点一览表》《动力触探试验成果统计表》等成果附表 → 规范化记录

全程**只读**：不写工程库、不删文件，所有输出都是调用方指定路径的新文件。

---

## 这个 MCP 现在能用来干嘛

| 场景 | 典型说法 | 用到 |
|---|---|---|
| 新项目建库 | "打开这个 .lz，看看工程信息和钻孔清单" | `lz_open_project` / `lz_project_summary` |
| 批量导接口文件 | "从工程库导出全部钻孔的接口文件，坐标用附表1 的" | `lz_db_to_interface` |
| 接口文件体检 | "校验这个接口文件能不能导入" | `lz_validate_interface_file` |
| 附表转数据 | "把这份勘探点一览表识别成钻孔记录" | `lz_excel_to_records` |
| 查规范 | "#DT 动力触探节有哪些字段，哪些必填" | `lz_get_spec` |
| 库里查数 | "读一下 z_g_TuCeng 表前 50 行" | `lz_read_table` |
| 手工拼接口文件 | "按规范生成一个 #ZK 节文件" | `lz_build_interface_file` / `lz_write...` |

一条典型工作流：

```
.lz 工程库 ──lz_db_to_interface──▶ 每孔一个接口文件(GBK/Tab)
                                      │
附表1(真实坐标) ──lz_excel_to_records─┘（坐标覆盖）
                                      ▼
                          理正软件内一键导入（按孔）
```

> 背景：接口文件是理正官方的批量数据通道（钻孔、土层、标贯、动探、水位、
> 波速、岩芯、取样、固结、颗分、直剪、旁压、水土分析……共 45 个节）。
> 以前要 Excel VBA 拼格式，现在说一句话就行。

---

## 环境要求

- **Windows**（依赖注册表定位与 Access ODBC 驱动）
- Python 3.10+（64 位）
- 本机已安装 **64 位 ACE Access 驱动**（随 Office 或 Access 数据库引擎安装）：
  驱动名 `Microsoft Access Driver (*.mdb, *.accdb)`
- 理正勘察 9.5PB3 已安装（用于**自动提取工程库密码**，见下）

```bash
pip install -r requirements.txt
```

## 工程库密码（无需配置）

.lz 内的 Access 工程库带库密码。理正把连接字符串硬编码在安装目录的二进制里，
本服务启动时**自动定位安装目录并提取**（环境变量 `LZ_INSTALL_DIR` > 注册表
卸载项 > 常见安装路径 `D:\LiZheng\...`），仓库中不含任何密码明文。
如需显式指定：`LZ_MDB_PASSWORD` 环境变量。

## 注册到 MCP 客户端

```json
{
  "mcpServers": {
    "lizheng": {
      "command": "python",
      "args": ["C:/path/to/lizheng-mcp/server.py"],
      "env": {}
    }
  }
}
```

## 自测

```bash
# 不依赖工程库的部分（backend_info / get_spec）
python test_stdio.py

# 完整链路（传一个真实 .lz）
python test_stdio.py "C:/path/to/工程.lz"
```

---

## 工具列表（11 个）

| 工具 | 说明 |
|---|---|
| `lz_backend_info` | 环境体检：pyodbc / ACE 驱动 / 规范节数 / 密码来源 |
| `lz_get_spec` | 接口规范速查：按节名（如 ZK/TC/DT）返回字段、必填项 |
| `lz_open_project` | 打开 .lz：解出 mdb、只读连接、返回表清单 |
| `lz_project_summary` | 工程摘要：工程信息、钻孔清单、各节数据量 |
| `lz_read_table` | 只读查询任意库表（分页） |
| `lz_db_to_interface` | **核心**：.lz → 理正标准接口文件（每孔一文件 / 合并 / 坐标覆盖） |
| `lz_build_interface_file` | 按规范把记录组装成接口文件文本 |
| `lz_parse_interface_file` | 解析接口文件 → 结构化 JSON |
| `lz_validate_interface_file` | 校验接口文件（错误/警告分级） |
| `lz_excel_to_records` | 识别成果附表 → 规范化记录（自动表头识别、合并单元格填充） |
| `lz_create_input_template` | 生成 Excel 录入模板 |

## 文件说明

```
lizheng-mcp/
├── server.py                        # MCP 服务端（全部逻辑）
├── test_stdio.py                    # 协议级握手测试
├── lz_standard_interface_spec.txt   # 官方《理正标准数据接口》规范原文（GBK）
├── samples/
│   └── ZK1_接口文件_sample.txt       # 导出的接口文件样例（GBK/Tab）
├── requirements.txt
├── LICENSE
└── README.md
```

## .lz 文件格式（逆向结论）

`.lz` 是 ZIP 包，内部：

| 条目 | 内容 |
|---|---|
| `LZGICAD1.mdb` | Access 工程库（Jet4，361 张表，带库密码）。表名前缀：`x_` 工程信息、`z_` 钻孔数据（`z_g_` 通用地质 / `z_y_` 原位试验 / `z_c_` 室内试验）、`t_` 统计表、`g_`/`k_` 字典 |
| `ProjectInfo.ini` | 工程编号、名称、目录指针 |
| `*.pdf` | 数字化交付的剖面图等 |

库表字段码与接口规范完全一致（`ZKBH`/`TCMC`/`DTDSD`/`QYBH`…），
这是 `lz_db_to_interface` 能直接映射的原因。

## 已知坑与处理

1. **库内钻孔坐标是保护值**（`-999...`）：导出时自动置空，用
   `coordinates` 参数从《勘探点一览表》覆盖真实坐标。
2. **`TCMS` 土层描述是理正压缩编码**（`LZPS&F...`）：原样导出，理正可解回。
3. **接口文件不带钻孔号**：按孔逐文件导入，与理正的导入交互一致。
4. 规范分节定界符 `#TAG#`（数据节）/ `%TAG%`（项目表节），收尾也用 `#`/`%`。

## 免责声明

本项目为个人开发的只读辅助工具，与北京理正软件股份有限公司无关。
"理正" 商标归其所有者所有。请在拥有合法授权的情况下使用理正软件，
并遵守相关许可协议。使用本项目产生的一切后果由使用者自行承担。

## License

MIT
