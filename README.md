# DownloadPrivateBond

从名单截图识别私募债清单，使用 WIND/Excel 刷新债券字段，并在 Ratingdog 下载对应债券的募集说明书 PDF。

## 适用范围

完整流程目前面向 Windows：

- Excel + WIND 插件用于刷新模板中的 WIND 公式。
- `pywin32` / Excel COM 用于驱动 Excel 刷新和固化公式值。
- Selenium + Chrome 用于登录 Ratingdog 并下载公告附件。

macOS 暂不支持完整 WIND 刷新流程。已有刷新后 Excel 的情况下，可以单独尝试 Ratingdog 下载脚本，但没有作为正式支持路径验证。

## 目录结构

```text
DownloadPrivateBond/
  data/                         # agent/OCR 写入 bond_list.json
  reference/                    # 样例截图、样例 Excel、样例募集说明书
  output/                       # 生成的 Excel 和日志
  Download/                     # 下载的 PDF
  skills/download-private-bond/ # Codex skill 与核心脚本
  tests/                        # 单元测试
```

## 环境准备

建议使用 Python 3.10+。

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

还需要本机已安装并可正常使用：

- Microsoft Excel
- WIND 金融终端 / WIND Excel 插件
- Google Chrome
- 与 Chrome 版本匹配的 ChromeDriver，或可由 Selenium 自动管理的 ChromeDriver

## 配置账号

复制示例配置：

```powershell
Copy-Item "skills\download-private-bond\.env.example" "skills\download-private-bond\.env"
```

填写：

```text
RATINGDOG_USERNAME=手机号
RATINGDOG_PASSWORD=密码
```

真实 `.env` 已被 `.gitignore` 排除，不要提交到 GitHub。

## 使用流程

### 1. 准备 `bond_list.json`

触发 `download-private-bond` skill 时，优先把名单截图作为附件传给 agent。agent 识别后写入：

```text
data\bond_list.json
```

格式：

```json
[
  {
    "company_name": "发行人公司名称",
    "bond_short_name": "债券简称"
  }
]
```

如果没有传图片，skill 会去 `reference` 目录找当天最新图片；仍找不到时会提醒补图。

### 2. 生成并刷新 Excel

```powershell
python skills\download-private-bond\scripts\prepare_bond_excel.py
```

输出文件位于：

```text
output\信评需求私募债_YYYYMMDD_HHMMSS.xlsx
```

脚本会：

- 复制 skill 内置模板 `skills\download-private-bond\assets\公式模板.xlsx`
- 填入公司名称和债券简称
- 刷新 WIND 公式
- 删除发行方式为“公募”的行
- 按发行人名称升序排序

### 3. 下载募集说明书

默认自动读取 `output` 目录下最新的 `信评需求私募债_*.xlsx`：

```powershell
python skills\download-private-bond\scripts\download_ratingdog_announcements.py
```

也可以指定 Excel：

```powershell
python skills\download-private-bond\scripts\download_ratingdog_announcements.py --excel "output\信评需求私募债_YYYYMMDD_HHMMSS.xlsx"
```

下载结果默认位于：

```text
Download\
```

日志默认位于：

```text
output\download_log_YYYYMMDD_HHMMSS.txt
```

## 样例文件

`reference` 中保留了三类样例：

- `名单.png`：名单截图样例
- `样例_信评需求私募债.xlsx`：WIND 刷新后的 Excel 样例
- `样例_安徽省铁路发展基金股份有限公司2025年募集说明书.pdf`：募集说明书 PDF 样例

## 验证

```powershell
python -m unittest discover -s tests -p "test_*.py" -v
python -m py_compile skills\download-private-bond\scripts\prepare_bond_excel.py skills\download-private-bond\scripts\download_ratingdog_announcements.py
```

## 注意事项

- 不要提交 `skills\download-private-bond\.env`。
- 不要提交批量生成的 Excel、日志、PDF 下载结果。
- 如果 WIND 返回 `0` 或债券全称缺少年份，下载脚本会跳过并写入日志。
- 下载目标必须是标题同时匹配债券全称且包含“募集说明书”的公告附件。
