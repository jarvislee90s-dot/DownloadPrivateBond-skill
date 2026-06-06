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

### 1. 准备截图识别结果

如果用户在触发 skill 时提供图片，优先使用该图片识别名单。

- 识别完成后，将记录写入 `data\bond_list.json`
- 支持两种格式：
  - **完整格式**：`{ "company_name": "发行人名称", "bond_short_name": "债券简称" }`
  - **简写格式**：`{ "bond_short_name": "债券简称" }`（公司名称通过WIND公式自动填充）
- `bond_list.json` 是轻量输入；下载阶段需要WIND刷新的Excel

**⚠️ 重要：每次新运行的标准**

- **必须以截图中的债券清单为准**，不能使用之前残留的JSON或Excel
- 运行前请确认 `data\bond_list.json` 已更新为当前截图的识别结果
- 旧日志和Excel仅供参考，不作为当前运行依据
- 如需重新运行，请重新提供截图或确认JSON内容已更新

**示例（简写格式）：**
```json
[
  { "bond_short_name": "26杭开F1" },
  { "bond_short_name": "26嘉州01" }
]
```

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

- 自动读取 `output` 目录下最新的 `信评需求私募债_*.xlsx`
- 也可指定文件：`--excel "output\信评需求私募债_YYYYMMDD_HHMMSS.xlsx"`

**环境变量（优先于 .env）：**
```powershell
$env:RATINGDOG_USERNAME="手机号"
$env:RATINGDOG_PASSWORD="密码"
```

**输出位置：**
- PDF 下载到 `Download\`
- 日志写入 `output\download_log_YYYYMMDD_HHMMSS.txt`
- 日志格式：制表符分隔的公司名、债券简称、债券全称、错误原因

### 4. 第二轮搜索（可选）

当第一轮存在失败债券时，进行第二轮搜索：

**情况A：浏览器未关闭（复用会话）**
```powershell
# 直接在现有浏览器中继续搜索失败债券
python skills\download-private-bond\scripts\download_ratingdog_announcements.py
```

**情况B：浏览器已关闭（创建重试Excel）**

1. **复制原Excel**：
   ```powershell
   Copy-Item "output\信评需求私募债_YYYYMMDD_HHMMSS.xlsx" "output\信评需求私募债_重试.xlsx"
   ```

2. **手动编辑重试Excel**：
   - 打开 `output\信评需求私募债_重试.xlsx`
   - **删除已成功下载的行**（保留失败的债券）
   - **删除债券简称含"PPN"的行**（公开渠道无法获取募集说明书）
   - 保存文件

3. **运行重试**：
   ```powershell
   python skills\download-private-bond\scripts\download_ratingdog_announcements.py --excel "output\信评需求私募债_重试.xlsx"
   ```

**筛选规则：**
- 从第一轮日志中查看失败记录
- **排除债券简称包含"PPN"的债券**（定向债务融资工具无法下载）
- 仅保留非PPN的失败债券进入第二轮

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
