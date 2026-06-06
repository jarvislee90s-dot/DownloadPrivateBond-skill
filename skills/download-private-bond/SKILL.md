---
name: download-private-bond
description: 用于处理 DownloadPrivateBond 私募债名单截图或粘贴的债券简称，准备 WIND 债券数据，并从 Ratingdog 下载私募债募集说明书 PDF。
---

# 私募债募集说明书下载

> 运行前请务必：
> 1. **提供最新的名单截图**作为输入,或**手动粘贴债券简称**（不要依赖 `data\bond_list.json` 中的旧数据）
> 2. **Download文件夹中的PDF不受影响**，会永久保留
> 3. **完整流程结束后**，可手动运行清理脚本删除临时文件（见第5步）
> 4. 如需重新运行，直接提供新截图即可

## 工作流

### 1. 准备截图识别结果

触发 skill 时提供输入数据，支持两种方式：

**方式一：提供截图（自动识别）**
- 将名单截图作为附件上传
- agent 自动识别截图中的"公司名称"和"债券简称"
- 识别完成后写入 `data\bond_list.json`
- 支持两种格式：
  - **完整格式**：`{ "company_name": "发行人名称", "bond_short_name": "债券简称" }`
  - **简写格式**：`{ "bond_short_name": "债券简称" }`（公司名称通过WIND公式自动填充）

**方式二：粘贴文字（跳过识别）**
- 如果无法提供截图，直接将债券简称以文字形式粘贴发送
- agent 跳过图像识别，直接将文字录入
- 公司名称通过 WIND 公式自动填充

**示例（粘贴文字）：**
```
26杭开F1
26嘉州01
26船投K1
```
- agent 会跳过图像识别，直接将文字录入为 `bond_short_name`

**生成的 JSON 格式：**
```json
[
  { "bond_short_name": "26杭开F1" },
  { "bond_short_name": "26嘉州01" },
  { "bond_short_name": "26船投K1" }
]
```

**识别常见问题：**
- 债券简称和发行人名称可能用**空格、下划线或制表符**分割
- 识别完成后请检查 `data\bond_list.json` 内容是否正确

**⚠️ 重要：每次新运行的标准**
- **必须以截图或粘贴的文字为准**，不能使用之前残留的 JSON
- 运行前请确认 `data\bond_list.json` 已更新为当前输入
- 旧日志和 Excel 仅供参考，不作为当前运行依据
- 如需重新运行，请重新提供截图或粘贴文字

**说明：**
- 支持两种格式：完整格式（含公司名称）或简写格式（仅债券简称）
- `bond_list.json` 只是轻量输入；下载阶段默认使用 WIND 刷新后导出的同名 JSON
- 公司名称可以通过 WIND 公式自动填充，所以只提供债券简称也可以

### 2. 生成 WIND Excel

```powershell
python skills\download-private-bond\scripts\prepare_bond_excel.py
```

**处理流程：**
- 复制内置模板 `assets\公式模板.xlsx`
- 填入债券简称到 A 列
- 刷新 WIND 公式（需本地 Excel + WIND 插件）
  - A列：债券简称（输入值，必填）
  - B列：公司名称（WIND公式自动查询）
  - C列：债券代码（WIND公式自动查询）
  - D列：债券全称（WIND公式自动查询）
  - E列：发行方式（WIND公式自动查询）
- 刷新后立刻从 Excel 缓存值导出同名 JSON：`output\信评需求私募债_YYYYMMDD_HHMMSS.json`
  - JSON 保留未过滤、未排序的 WIND 原始结果，便于排查 Excel COM 或公式缓存问题
  - 如果 Excel COM 报错但缓存值完整，脚本会打印警告并继续写 JSON 与最终 Excel
  - 如果 WIND 返回错误码（如 `-214682...`）或缓存值不完整，脚本会停止，避免把错误值固化
- 删除发行方式为"公募"的行，按发行人排序
- 输出：`output\信评需求私募债_YYYYMMDD_HHMMSS.xlsx`

**跳过 WIND 刷新（仅测试）：**
```powershell
python skills\download-private-bond\scripts\prepare_bond_excel.py --skip-wind
```

### 3. 下载 Ratingdog 公告

```powershell
python skills\download-private-bond\scripts\download_ratingdog_announcements.py
```

**处理逻辑：**
- 自动读取 `output` 目录下最新的 `信评需求私募债_*.json`
- 下载脚本从 JSON 读取 `bond_full_name`、`bond_short_name`、`company_name` 等字段，不再打开 Excel
- 如传入旧参数 `--excel xxx.xlsx`，脚本只会自动改用同名 `xxx.json`，不会读取 xlsx 文件
- 读取 JSON 后会排除发行方式为"公募"的债券
- 第一轮处理时**自动排除债券简称含"PPN"的债券**（公开渠道无法获取）
- 按年份筛选、搜索、下载每只债券的募集说明书

**环境变量（优先于 .env）：**
```powershell
$env:RATINGDOG_USERNAME="手机号"
$env:RATINGDOG_PASSWORD="密码"
```

**输出位置：**
- PDF 下载到 `Download\`
- 日志写入 `output\download_log_YYYYMMDD_HHMMSS.txt`

### 4. 第二轮搜索（可选）

当第一轮存在下载失败的债券时，进行第二轮搜索：

**情况A：浏览器未关闭（复用会话）**
```powershell
# 直接在现有浏览器中继续搜索失败债券
python skills\download-private-bond\scripts\download_ratingdog_announcements.py
```

**情况B：浏览器已关闭（创建重试 JSON）**

1. **查看第一轮日志，找到失败的债券**：
   ```powershell
   # 找到最新的日志文件
   Get-ChildItem output\download_log_*.txt | Sort-Object LastWriteTime -Descending | Select-Object -First 1
   ```

2. **从日志中提取【下载失败的名单】**，创建重试 JSON

3. **运行重试**：
   ```powershell
   python skills\download-private-bond\scripts\download_ratingdog_announcements.py --json "output\信评需求私募债_重试.json"
   ```

**第二轮日志特点：**
- 日志标题会标注 **【第二轮】**
- 会自动查找并汇总第一轮的数据
- 生成**汇总报告**：`output\下载汇总报告_YYYYMMDD_HHMMSS.txt`
- 汇总报告包含：
  - 第一轮统计（总数、成功、PPN、失败）
  - 第二轮统计（重试数、成功、仍失败）
  - 最终汇总（总计、总成功、总失败、成功率）

> **说明：** PPN债券已在第一轮自动排除，不会进入第二轮。

### 5. 清理临时文件（可选）

完整流程结束后，清理临时文件（保留Download文件夹和日志）：

```powershell
python skills\download-private-bond\scripts\cleanup_residual_files.py
```

**说明：**
- 此步骤**在第二轮搜索完成后**执行
- 清理 `data\bond_list.json`、临时输入 JSON、历史 Excel/JSON 文件
- **保留** `Download\` 文件夹（已下载的PDF）
- **保留** `output\download_log_*.txt` 日志（用于问题排查）
- **保留** `.gitkeep` 文件

**模拟运行（查看将要删除的文件）：**
```powershell
python skills\download-private-bond\scripts\cleanup_residual_files.py --dry-run
```

> **注意：** 清理后如需重新运行，请重新提供名单截图。

## 关键约束

- **仅支持 Windows**：WIND 刷新依赖 `pywin32` 和本地 Excel
- **下载目标**：标题必须同时匹配债券全称且包含"募集说明书"
- **日期筛选**：按债券全称中的年份设置 `YYYY-01-01` 至 `YYYY-12-31`
- **搜索关键词**：债券全称 + "募集说明书"（自动去掉 `(品种X)` 后缀）
- **账号安全**：`.env` 已被 `.gitignore` 排除，不要提交真实账号

## 依赖检查

运行前确认已安装：
- Python 3.10+
- `openpyxl`、`selenium`、`pywin32` (Windows)
- Microsoft Excel + WIND 金融终端插件
- Google Chrome + 匹配版本的 ChromeDriver
