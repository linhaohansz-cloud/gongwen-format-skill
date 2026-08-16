# 国企公文格式 skill —— Windows 自安装脚本
# 用法（三选一）：
#   1) 从直链下载并安装：
#      .\install.ps1 -Url "https://你的直链/国企公文格式-skill.zip"
#   2) 从本地 zip 安装：
#      .\install.ps1 -Path "C:\下载\国企公文格式-skill.zip"
#   3) 脚本与 zip 同目录时直接运行：
#      .\install.ps1
param(
    [string]$Url,
    [string]$Path
)

$ErrorActionPreference = 'Stop'
$dest = Join-Path $env:USERPROFILE ".workbuddy\skills\国企公文格式"
$tmp  = New-Item -ItemType Directory -Force -Path (Join-Path $env:TEMP "gww_install") | Select-Object -ExpandProperty FullName
$zip  = Join-Path $tmp "skill.zip"

if ($Url) {
    Write-Host "正在下载: $Url"
    Invoke-WebRequest -Uri $Url -OutFile $zip
} elseif ($Path) {
    Copy-Item $Path $zip
} else {
    # 默认：脚本同目录下的 zip
    $zip = Join-Path $PSScriptRoot "国企公文格式-skill.zip"
}

if (-not (Test-Path $zip)) {
    Write-Error "未找到安装包: $zip`n请通过 -Url 或 -Path 指定，或把 zip 放在脚本同目录。"
    exit 1
}

Expand-Archive -Path $zip -DestinationPath $tmp -Force
$skillFile = Get-ChildItem -Path $tmp -Recurse -Filter SKILL.md | Select-Object -First 1
$srcDir = if ($skillFile) { $skillFile.DirectoryName } else { $tmp }

New-Item -ItemType Directory -Force -Path $dest | Out-Null
Copy-Item -Path (Join-Path $srcDir '*') -Destination $dest -Recurse -Force

Write-Host ""
Write-Host "✅ 已安装到: $dest"
Write-Host "下次直接对 AI 说：用「国企公文格式」skill 即可。"
