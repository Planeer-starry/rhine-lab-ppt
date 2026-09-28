# 把 rhine-lab-ppt 安装为 DSH 技能
# 用法：
#   powershell -ExecutionPolicy Bypass -File scripts\install-skill.ps1 [-Scope user|project] [-ProjectRoot <路径>]
# 也可以用 Invoke-Expression 直接执行本文件内容（会自动向上查找仓库根）。
param(
    [ValidateSet('user','project')]
    [string]$Scope = 'user',
    [string]$ProjectRoot = '',
    [string]$RepoRoot = ''
)

$ErrorActionPreference = 'Stop'
$SkillName = 'rhine-lab-ppt'

# 确定仓库根：显式参数 > $PSScriptRoot 的父目录 > 向上查找含 SKILL.md 的目录
function Resolve-RepoRoot {
    param([string]$Explicit, [string]$ScriptRoot)
    if ($Explicit -and (Test-Path -LiteralPath (Join-Path $Explicit 'SKILL.md'))) {
        return (Resolve-Path -LiteralPath $Explicit).Path
    }
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

$root = Resolve-RepoRoot -Explicit $RepoRoot -ScriptRoot $PSScriptRoot
if (-not $root) {
    throw '找不到仓库根（需要含 SKILL.md 的目录）。请用 -RepoRoot 显式指定，或在仓库目录内运行。'
}
Write-Host "仓库根：$root"

switch ($Scope) {
    'user' {
        $dshHome = $env:DSH_HOME
        if (-not $dshHome) { $dshHome = Join-Path $env:APPDATA 'dsh-desktop\harness' }
        $target = Join-Path $dshHome "skills\$SkillName"
    }
    'project' {
        if (-not $ProjectRoot) { $ProjectRoot = (Get-Location).Path }
        if (-not (Test-Path -LiteralPath $ProjectRoot)) { throw "项目根不存在：$ProjectRoot" }
        $target = Join-Path $ProjectRoot ".dsh\skills\$SkillName"
    }
}

Write-Host "安装到：$target"

if (Test-Path -LiteralPath $target) {
    Write-Host '目标已存在，先移除旧版本……'
    Remove-Item -LiteralPath $target -Recurse -Force
}
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $target) | Out-Null
New-Item -ItemType Directory -Force -Path $target | Out-Null

# 只复制技能运行所需的内容，跳过仓库元数据
$include = @('SKILL.md', 'references', 'scripts', 'templates', 'docs')
foreach ($item in $include) {
    $src = Join-Path $root $item
    if (Test-Path -LiteralPath $src) {
        Copy-Item -LiteralPath $src -Destination $target -Recurse -Force
    }
}

Write-Host ''
Write-Host "已安装。技能清单会在无需重启的情况下刷新，之后 Agent 即可按名称加载：$SkillName"
Write-Host '验证：目录中应存在 SKILL.md，且 frontmatter 含 name 与 description。'
Get-ChildItem -LiteralPath $target | Select-Object Mode, Name | Format-Table -AutoSize
