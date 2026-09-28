# 把本目录推送到 GitHub
# 用法：
#   1) 先在 GitHub 网页端创建空仓库（不要勾选 README / .gitignore / License）
#   2) 在仓库目录内执行：
#      powershell -ExecutionPolicy Bypass -File scripts\push-to-github.ps1 -RepoUrl https://github.com/Planeer-starry/rhine-lab-ppt
param(
    [Parameter(Mandatory = $true)]
    [string]$RepoUrl,
    [string]$UserName = 'Planeer-starry',
    [string]$UserEmail = '',
    [string]$Branch = 'main',
    [string]$CommitMessage = 'feat: RHINE LAB PPT style skill (design system + renderer + master templates)'
)

$ErrorActionPreference = 'Stop'

# 仓库根：-File 方式下 $PSScriptRoot 指向 scripts；否则向上查找含 SKILL.md 的目录
function Resolve-RepoRoot {
    param([string]$ScriptRoot)
    if ($ScriptRoot) {
        $parent = Split-Path -Parent $ScriptRoot
        if ($parent -and (Test-Path -LiteralPath (Join-Path $parent 'SKILL.md'))) {
            return (Resolve-Path -LiteralPath $parent).Path
        }
    }
    $dir = (Get-Location).Path
    for ($i = 0; $i -lt 6; $i++) {
        if (Test-Path -LiteralPath (Join-Path $dir 'SKILL.md')) { return $dir }
        $up = Split-Path -Parent $dir
        if (-not $up -or $up -eq $dir) { break }
        $dir = $up
    }
    return ''
}

$RepoRoot = Resolve-RepoRoot -ScriptRoot $PSScriptRoot
if (-not $RepoRoot) {
    throw '找不到仓库根（需要含 SKILL.md 的目录）。请在仓库目录内运行，或改用仓库根的一键脚本。'
}

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    throw '未找到 git。请先安装 Git for Windows: https://git-scm.com/download/win，然后重开 PowerShell 窗口。'
}

if ($RepoUrl -notlike '*.git') { $RepoUrl = "$RepoUrl.git" }
if ($RepoUrl -match '<') { throw '请把 -RepoUrl 换成你真实的仓库地址。' }

Write-Host "仓库根：$RepoRoot"
Push-Location $RepoRoot
try {
    if (-not (Test-Path -LiteralPath '.git')) {
        Write-Host '初始化本地仓库……'
        & git init -b $Branch 2>$null
        if ($LASTEXITCODE -ne 0) {
            & git init | Out-Null
            & git symbolic-ref HEAD "refs/heads/$Branch"
        }
        & git config user.name  $UserName
        if ($UserEmail) { & git config user.email $UserEmail } else { & git config user.email "$UserName@users.noreply.github.com" }
    } else {
        Write-Host '已存在 .git，沿用当前历史。'
    }

    & git add -A
    $staged = & git diff --cached --name-only
    if ($staged) {
        & git commit -m $CommitMessage | Out-Host
    } else {
        Write-Host '没有需要提交的改动。'
    }

    & git branch -M $Branch

    $existing = & git remote
    if ($existing -contains 'origin') { & git remote set-url origin $RepoUrl } else { & git remote add origin $RepoUrl }

    Write-Host "推送到 $RepoUrl ……"
    & git push -u origin $Branch | Out-Host
    if ($LASTEXITCODE -ne 0) { throw '推送失败。请检查仓库是否已创建、以及 Personal Access Token 是否具备 Contents 读写权限。' }

    Write-Host ''
    Write-Host ('完成。仓库地址：' + ($RepoUrl -replace '\.git$',''))
}
finally {
    Pop-Location
}
