# 国企公文格式（GB/T 9704）WorkBuddy Skill

一个用于生成 / 排版**党政机关、国企公文**的 WorkBuddy 技能。复制你提供的「标准版模板」docx，
仅替换正文内容，从而 **100% 继承模板的页面设置与页脚页码**，严格符合 GB/T 9704 格式规范。

> 本 skill 是「标准工作流」的固化：以模板为唯一权威来源，遇到不确定之处应先向用户确认，
> 不得自行臆测；所有输出为可直接打开的 Word 文档。

---

## ✨ 功能特性

- **倒梯形标题**：方正小标宋简体，二号，居中，多行时首行最长、逐行递减。
- **日期副标题**：楷体，居中；日期后**自动空一行**（与标准版一致）。
- **主送单位 / 正文**：仿宋_GB2312，三号，正文两端对齐、首行缩进 2 字符。
- **三级标题体系**：
  - 一级（一、二、三…）黑体加粗
  - 二级（（一）（二）…）楷体
  - 三级（（1）（2）…）仿宋
- **小标题不分页**：所有标题设 `keepNext`，避免小标题与正文被分页拆开。
- **版心基准**：每页 22 行 × 26 字（可视情况增减 1–2 行并保持版心高度）。
- **页脚页码**：`— 页码 —`，宋体，右对齐，PAGE 域自动更新。
- **自带验证器**：生成后对照规范逐项核对，输出 ✅/❌ 清单。

---

## 📁 目录结构

```
国企公文格式/
├── SKILL.md                      # 工作流 + 核心纪律 + 完整格式规范
├── README.md                     # 本文件
├── LICENSE                       # MIT 开源协议
├── .gitignore
├── scripts/
│   ├── gen_gongwen.py            # 生成器：复制模板 → 替换正文
│   └── verify_gongwen.py         # 验证器：对照 GB/T 9704 自动核对
└── references/
    └── format_spec.md            # 完整格式规范镜像（可单独查阅）
```

---

## 📦 安装

### 方式零：用自带安装脚本（最简单，无需 git）

本包已内置自安装脚本：`install.ps1`（Windows）/ `install.sh`（macOS / Linux）。
给它一个**可直接下载的 zip 地址**（GitHub Release、自有服务器等均可，不一定要公开开源），
它会自动解压到 skills 目录，装完即可用。

```powershell
# Windows：从任意直链安装
.\install.ps1 -Url "https://你的直链/国企公文格式-skill.zip"
# 或从本地 zip 安装
.\install.ps1 -Path "C:\下载\国企公文格式-skill.zip"
```

```bash
# macOS / Linux
./install.sh "https://你的直链/国企公文格式-skill.zip"
```

> 有了这个安装包 + 一个直链，下次只需对 AI 说：
> 「从 <链接> 下载国企公文格式-skill.zip，用里面的 install 脚本装到你的 skills 文件夹」，
> AI 即可一键装好该 skill。

### 方式一：Git 克隆（推荐，便于更新）

```powershell
git clone <本仓库地址> "$env:USERPROFILE\.workbuddy\skills\国企公文格式"
```

### 方式二：下载 ZIP 后解压

1. 下载仓库的 `main.zip`；
2. 解压后将其内容放入：
   - Windows：`%USERPROFILE%\.workbuddy\skills\国企公文格式\`
   - macOS/Linux：`~/.workbuddy/skills/国企公文格式/`

### 方式三：一键安装（PowerShell，从 GitHub Releases / main 分支）

```powershell
$url = "https://github.com/<用户名>/<仓库名>/archive/refs/heads/main.zip"
$tmp = New-Item -ItemType Directory -Path "$env:TEMP\gww_install" -Force
Invoke-WebRequest $url -OutFile "$tmp\skill.zip"
Expand-Archive "$tmp\skill.zip" -DestinationPath "$tmp\ext"
$dest = "$env:USERPROFILE\.workbuddy\skills\国企公文格式"
New-Item -ItemType Directory -Path $dest -Force | Out-Null
Copy-Item "$tmp\ext\*-main\*" $dest -Recurse -Force
Write-Host "安装完成 -> $dest"
```

---

## 🔧 依赖

- **Python 3.8+**
- **lxml**：`pip install lxml`

> 若使用 WorkBuddy 内置托管 Python，请在该环境的 venv 中安装 lxml。

---

## 🚀 使用方法

### 1. 准备内容 JSON

内容用二维数组 `[[kind, text], ...]` 表示，`kind ∈ {title, date, zhusong, body, h1, h2, h3}`：

| kind     | 含义                         | 示例                          |
|----------|------------------------------|-------------------------------|
| `title`  | 标题（可多行，每行一项）     | `"关于近期体育事业发展情况"`  |
| `date`   | 日期副标题                   | `"(2026 年 08 月 16 日)"`     |
| `zhusong`| 主送单位                     | `"主送单位：市体育局"`        |
| `body`   | 正文段落                     | `"今年以来，……"`             |
| `h1`     | 一级标题（一、二、三）       | `"一、总体运行情况"`          |
| `h2`     | 二级标题（（一））           | `"（一）市场规模"`            |
| `h3`     | 三级标题（（1））            | `"（1）区域分布"`             |

> 日期后**无需**在 JSON 里写空项，生成器会自动插入空行。

### 2. 生成公文

```powershell
python gen_gongwen.py -t 标准版模板.docx -c 内容.json -o 输出.docx
```

### 3. 验证格式

```powershell
python verify_gongwen.py 输出.docx
```

---

## ⚠️ 注意事项

1. **标准版模板需自备**：生成器通过 `-t` 复制你的「公文格式标准版.docx」，
   本仓库**不含**该模板（通常含企业/单位专属版式）。请将其放在可访问路径并传入 `-t`。
2. **中文字体依赖**：标题/正文使用了 方正小标宋简体、楷体_GB2312、仿宋_GB2312、黑体。
   若目标机器未安装这些字体，Word 会做字体替换，但**版式（字号、对齐、缩进、行距、页码）仍严格生效**。
3. **跨平台**：脚本仅依赖标准库 + lxml，Windows / macOS / Linux 均可运行。

---

## 📄 开源协议

[MIT License](./LICENSE) — 可自由使用、修改、分发。
