# DRIVER Session Start Hook (PowerShell)
# Injects the using-driver skill at the start of every session

# Derive plugin root from this script's location if not set
if (-not $env:CLAUDE_PLUGIN_ROOT) {
    $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $env:CLAUDE_PLUGIN_ROOT = Split-Path -Parent $ScriptDir
}

$SkillPath = Join-Path $env:CLAUDE_PLUGIN_ROOT "skills\using-driver\SKILL.md"

if (Test-Path $SkillPath) {
    Write-Output "<EXTREMELY-IMPORTANT>"
    Get-Content $SkillPath -Raw
    Write-Output "</EXTREMELY-IMPORTANT>"
}
