# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

DownloadPrivateBond 是一个私募债募集说明书自动化下载工具。工作流程：
1. 从名单截图 OCR 识别公司名称和债券简称
2. 使用 Excel + WIND 插件刷新债券字段（债券代码、债券全称、发行方式）
3. 登录 Ratingdog 下载对应债券的募集说明书 PDF

**仅支持 Windows**：WIND 刷新依赖 Excel COM (`pywin32`) 和本地安装的 WIND 金融终端插件。

## 项目结构

```
DownloadPrivateBond/
  data/                         # 运行时：OCR 识别结果 bond_list.json
  reference/                    # 样例文件（截图、Excel、PDF）
  output/                       # 运行时：生成的 Excel 和日志
  Download/                     # 运行时：下载的 PDF
  skills/download-private-bond/
    scripts/                    # 核心脚本
    assets/                     # Excel 公式模板
    references/                 # 流程文档
    agents/                     # Skill 配置
  tests/                        # 单元测试
```

## 常用命令

### 运行测试
```powershell
# 运行全部测试
python -m unittest discover -s tests -p "test_*.py" -v

# 运行单个测试文件
python -m unittest tests.test_prepare_bond_excel -v
python -m unittest tests.test_download_ratingdog_announcements -v

# 语法检查
python -m py_compile skills\download-private-bond\scripts\prepare_bond_excel.py skills\download-private-bond\scripts\download_ratingdog_announcements.py
```

### 生成 WIND Excel
```powershell
# 完整流程（刷新 WIND 公式）
python skills\download-private-bond\scripts\prepare_bond_excel.py

# 跳过 WIND 刷新（仅用于开发测试）
python skills\download-private-bond\scripts\prepare_bond_excel.py --skip-wind

# 指定输入/输出
python skills\download-private-bond\scripts\prepare_bond_excel.py --input data\bond_list.json --output-dir output
```

### 下载 Ratingdog PDF
```powershell
# 自动使用 output 目录下最新的 Excel
python skills\download-private-bond\scripts\download_ratingdog_announcements.py

# 指定 Excel 文件
python skills\download-private-bond\scripts\download_ratingdog_announcements.py --excel "output\信评需求私募债_YYYYMMDD_HHMMSS.xlsx"

# 指定下载目录和账号
python skills\download-private-bond\scripts\download_ratingdog_announcements.py --download-dir Download --username "手机号" --password "密码"
```

## 配置账号

```powershell
# 复制示例配置
Copy-Item "skills\download-private-bond\.env.example" "skills\download-private-bond\.env"
```

编辑 `.env` 文件：
```
RATINGDOG_USERNAME=手机号
RATINGDOG_PASSWORD=密码
```

## 核心脚本架构

### prepare_bond_excel.py
- `load_ocr_records()` - 从 JSON 加载 OCR 识别结果
- `fill_template_with_excel()` - 使用 Excel COM 填充模板并保留公式
- `refresh_wind_values()` - 触发 WIND 公式计算并固化值
- `remove_public_and_sort_rows()` - 过滤公募债并按发行人排序
- 输出：`output\信评需求私募债_YYYYMMDD_HHMMSS.xlsx`

### download_ratingdog_announcements.py
- `login()` - 登录 Ratingdog 并跳转到租户公告页
- `search_tenant_announcements()` - 按年份和关键词搜索公告
- `find_and_download()` - 匹配标题含"募集说明书"的附件并下载
- `process_rows()` - 遍历债券列表，处理每一只债券
- 输出：`Download\*.pdf` 和 `output\download_log_YYYYMMDD_HHMMSS.txt`

## 关键实现细节

### 路径解析
两个脚本都使用 `find_project_root()` 通过检测 `data` 和 `reference` 目录来定位项目根目录，支持从任意位置运行。

### WIND 刷新机制
- 使用 `win32com.client.Dispatch("Excel.Application")` 打开 Excel
- 调用 `CalculateFullRebuild()` 触发全量重算
- 轮询检测 C2:E 列值，等待非空且非 "Fetching..."
- 通过 `Range.Value = Range.Value` 固化公式为值
- 超时默认 120 秒

### Ratingdog 下载流程
- 从债券全称正则提取年份 `(20\d{2})年`
- 日期范围设为 `YYYY-01-01` 至 `YYYY-12-31`
- 搜索关键词 = 债券全称 + "募集说明书"
- 标题匹配时统一全角括号、半角括号、空格、引号
- 下载后验证无 `.crdownload` 文件才继续

### 环境依赖
- `openpyxl` - Excel 文件读写
- `selenium` - 浏览器自动化
- `pywin32` - Windows 专属，Excel COM 操作
- Chrome + ChromeDriver
- Microsoft Excel + WIND 金融终端插件

## 输入输出约定

### bond_list.json 格式
```json
[
  {
    "company_name": "发行人公司名称",
    "bond_short_name": "债券简称"
  }
]
```

### Excel 列定义
- A 列：公司名称
- B 列：债券简称
- C 列：债券代码（WIND 公式）
- D 列：债券全称（WIND 公式）
- E 列：发行方式（WIND 公式）

### 日志格式
制表符分隔：`公司名称\t债券简称\t债券全称\t错误原因`

## 开发注意事项

- `.env` 文件已被 `.gitignore` 排除，不要提交真实账号
- `data/*`、`output/*`、`Download/*` 也被排除，只保留 `.gitkeep`
- 修改脚本后必须运行测试和语法检查
- 真实联调时先用单行 Excel 测试，确认下载到的是募集说明书 PDF
