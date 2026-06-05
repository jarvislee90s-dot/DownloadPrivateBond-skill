---
name: download-private-bond
description: 私募债募集说明书下载工作流。用于根据名单截图提取公司名称和债券简称，生成 data/bond_list.json，套用内置 WIND 公式模板刷新债券字段，过滤公募债并排序，然后登录 Ratingdog 主体评级页面按发行人和债券全称下载“募集说明书”PDF。触发场景包括：处理 DownloadPrivateBond 项目、从 名单.png 生成私募债 Excel、刷新 WIND 私募债清单、下载 Ratingdog 相关公告附件、批量下载私募债募集说明书。
---

# 私募债募集说明书下载

## 快速判断

在 DownloadPrivateBond 项目内执行本 skill。遇到以下任务时使用：

- 从名单截图识别“公司名称、债券简称”。
- 用 `prepare_bond_excel.py` 生成并刷新 WIND Excel。
- 用 `download_ratingdog_announcements.py` 登录 Ratingdog 下载募集说明书。
- 排查日期筛选、私募勾选、附件下载、日志输出等问题。

详细网页交互规则见 [workflow.md](references/workflow.md)。

## 仓库约定

- 固定资产位于 `skills\download-private-bond`。
- 项目根部 `data`、`reference`、`output`、`Download` 是运行时目录。
- `reference` 仅保留样例截图、样例 Excel、样例募集说明书。
- 真实账号密码放在 `skills\download-private-bond\.env`，但该文件不提交；提交 `.env.example`。
- 完整流程默认仅支持 Windows，因为 WIND 刷新依赖 Excel COM / `pywin32`。

## 输入图片

- 如果用户在触发 skill 时提供图片，优先使用该图片识别名单。
- 如果没有提供图片，在项目 `reference` 目录查找文件名或修改日期包含今天日期的最新图片。
- 如果仍没有找到图片，暂停并提醒用户提供名单截图。
- 识别完成后，由 agent 写入 `data\bond_list.json`，不要把该动态文件放入 skill 目录。

## 工作流

1. 准备截图识别结果。
   - 将 agent/OCR 识别的记录写入 `data\bond_list.json`。
   - JSON 每条记录包含 `company_name` 和 `bond_short_name`。
   - `bond_list.json` 只是轻量输入；下载阶段需要 WIND 刷出的 Excel，不直接用它下载。

2. 生成 WIND Excel。
   - 运行 `python skills\download-private-bond\scripts\prepare_bond_excel.py`。
   - 脚本会复制 skill 内置的 `assets\公式模板.xlsx`，填入 A/B 列，刷新 C/D/E 列 WIND 公式，删除 `发行方式=公募` 的行，并按发行人名称升序排序。
   - 输出在 `output\信评需求私募债_YYYYMMDD_HHMMSS.xlsx`。

3. 下载 Ratingdog 公告。
   - 优先使用 `skills\download-private-bond\.env` 中的 `RATINGDOG_USERNAME` 和 `RATINGDOG_PASSWORD`。
   - 也可用环境变量覆盖：

```powershell
$env:RATINGDOG_USERNAME="手机号"
$env:RATINGDOG_PASSWORD="密码"
python skills\download-private-bond\scripts\download_ratingdog_announcements.py
```

   - 未传 `--excel` 时，脚本自动使用项目 `output` 目录下最新的 `信评需求私募债_*.xlsx`。
   - 需要指定文件时再传：

```powershell
python skills\download-private-bond\scripts\download_ratingdog_announcements.py --excel "output\信评需求私募债_YYYYMMDD_HHMMSS.xlsx"
```

4. 检查输出。
   - PDF 默认下载到 `Download`。
   - 日志默认写入 `output\download_log_YYYYMMDD_HHMMSS.txt`。
   - 日志中记录未匹配、无附件、下载超时、债券全称无年份等情况。

## 关键约束

- 下载目标必须是标题中同时匹配债券全称且包含“募集说明书”的公告附件。
- 先按债券全称中的年份筛 `YYYY-01-01` 到 `YYYY-12-31`，再勾选“只看私募”，然后搜索。
- 如果年度私募范围内没有匹配，再撤销日期和私募筛选，在全部公告里搜索一次。
- 相邻两行发行人相同，保持同一个发行人详情页，逐只债券搜索；发行人变化后关闭详情标签并回到“主体评级”。
- 不读取或展示 `.env`、密钥、凭据文件内容。

## 验证

修改脚本后至少运行：

```powershell
python -m unittest discover -s tests -p "test_*.py" -v
python -m py_compile skills\download-private-bond\scripts\prepare_bond_excel.py skills\download-private-bond\scripts\download_ratingdog_announcements.py
```

真实联调时，先用单行 Excel 烟测，确认下载到的是“募集说明书”PDF，再运行全量任务。
