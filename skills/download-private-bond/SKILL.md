---
name: download-private-bond
description: 私募债募集说明书下载工作流。根据名单截图提取债券信息，生成 WIND Excel，登录 Ratingdog 按发行人和债券全称下载募集说明书 PDF。触发场景：处理 DownloadPrivateBond 项目、从名单.png 生成私募债清单、刷新 WIND 数据、下载 Ratingdog 公告附件。
---

# 私募债募集说明书下载

> **⚠️ 重要提示：每次运行必须以截图中的债券清单为准**
> 
> 运行前请务必：
> 1. **提供最新的名单截图**作为输入（不要依赖 `data\bond_list.json` 中的旧数据）
> 2. **Download文件夹中的PDF不受影响**，会永久保留
> 3. **完整流程结束后**，可手动运行清理脚本删除临时文件（见第5步）
> 4. 如需重新运行，直接提供新截图即可

## 工作流

### 1. 准备债券清单

**方式一：提供截图（自动识别）**
- 将名单截图作为附件传给 agent
- agent 自动识别并写入 `data\bond_list.json`

**方式二：提供文字列表（跳过识别）**
- 直接将债券简称列表以文字形式发送给 agent
- **格式示例：**
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

**说明：**
- 支持两种格式：完整格式（含公司名称）或简写格式（仅债券简称）
- 公司名称可通过WIND公式自动填充，所以简写格式即可
- `bond_list.json` 是轻量输入，下载阶段需要WIND刷新的Excel

### 2. 生成 WIND Excel

```powershell
python skills\download-private-bond\scripts\prepare_bond_excel.py
```

**处理流程：**
- 复制内置模板 `assets\公式模板.xlsx`
- 填入债券简称（A列公司名称可选填，留空则由WIND公式自动填充）
- 刷新 WIND 公式（需本地 Excel + WIND 插件）
  - A列：公司名称（WIND公式自动查询）
  - C列：债券代码（WIND公式自动查询）
  - D列：债券全称（WIND公式自动查询）
  - E列：发行方式（WIND公式自动查询）
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
- 自动读取 `output` 目录下最新的 `信评需求私募债_*.xlsx`
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

**日志格式（分类统计）：**
```
================================================================================
下载完成汇总
================================================================================
总债券数: 25
完成下载: 20
排除PPN: 2
下载失败: 3

【总名单】
  - 26杭开F1: 杭州临平开发投资集团有限公司
  - 26嘉州01: 乐山国有资产投资运营(集团)有限公司
  ...

【完成下载的名单】
  ✓ 26杭开F1: 杭州临平开发投资集团有限公司
  ✓ 26嘉州01: 乐山国有资产投资运营(集团)有限公司
  ...

【排除的PPN名单】
  ⊘ 26西湖投资PPN001: 杭州西湖投资集团有限公司 (PPN债券无法下载)
  ⊘ 26余杭交通PPN003: 杭州余杭交通投资集团有限公司 (PPN债券无法下载)

【下载失败的名单】
  ✗ 26湖交02: 湖州市交通投资集团有限公司
  ✗ 26经控02: 嘉兴经济技术开发区国有资本投资控股集团有限公司
  ✗ 26金华01: 金华市城市发展集团有限公司

================================================================================

[详细失败记录]
制表符分隔：公司名称  债券简称  债券全称  失败原因
```

### 4. 第二轮搜索（可选）

当第一轮存在下载失败的债券时，进行第二轮搜索：

**情况A：浏览器未关闭（复用会话）**
```powershell
# 直接在现有浏览器中继续搜索失败债券
python skills\download-private-bond\scripts\download_ratingdog_announcements.py
```

**情况B：浏览器已关闭（创建重试Excel）**

1. **查看第一轮日志，找到失败的债券**：
   ```powershell
   # 找到最新的日志文件
   Get-ChildItem output\download_log_*.txt | Sort-Object LastWriteTime -Descending | Select-Object -First 1
   ```

2. **从日志中提取【下载失败的名单】**，创建重试Excel

3. **运行重试**：
   ```powershell
   python skills\download-private-bond\scripts\download_ratingdog_announcements.py --excel "output\信评需求私募债_重试.xlsx"
   ```

> **说明：** PPN债券已在第一轮自动排除，不会进入第二轮。

### 5. 清理临时文件（可选）

完整流程结束后，清理临时文件（保留Download文件夹和日志）：

```powershell
python skills\download-private-bond\scripts\cleanup_residual_files.py
```

**说明：**
- 此步骤**在第二轮搜索完成后**执行
- 清理 `data\bond_list.json` 和历史Excel文件
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
