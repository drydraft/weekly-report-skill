<#
.SYNOPSIS
    Cross-platform installer for weekly-report-skill (Windows PowerShell).
.DESCRIPTION
    Installs the single-source weekly-report skill core and, where required,
    the target platform's external plugin/bundle wrapper without duplicating source files.
.PARAMETER Target
    One or more targets: pi, codex, codex-plugin, dsh, dsh-bundle,
    antigravity, antigravity-skill, claude, claude-plugin, opencode, all.
.PARAMETER HomeDir
    Override user home directory (useful for isolated staging or verification).
.PARAMETER DryRun
    Print planned actions without writing files.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet(
        'pi', 'codex', 'codex-plugin', 'dsh', 'dsh-bundle',
        'antigravity', 'antigravity-skill', 'claude', 'claude-plugin',
        'opencode', 'all'
    )]
    [string[]]$Target,

    [string]$HomeDir = $env:USERPROFILE,

    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path

$CoreDirs = @('agents', 'evals', 'references', 'tests', 'workflows')
$CoreFiles = @('LICENSE', 'README.md', 'SKILL.md')
$CoreScripts = @('generate_report.py', 'validate_payload.py')

function Copy-SkillCore {
    param(
        [Parameter(Mandatory = $true)][string]$Destination
    )

    if ($DryRun) {
        Write-Host "[DryRun] Install skill core -> $Destination"
        return
    }

    New-Item -ItemType Directory -Force -Path $Destination | Out-Null

    foreach ($file in $CoreFiles) {
        $src = Join-Path $RepoRoot $file
        if (Test-Path $src) {
            Copy-Item -Path $src -Destination (Join-Path $Destination $file) -Force
        }
    }

    foreach ($dir in $CoreDirs) {
        $srcDir = Join-Path $RepoRoot $dir
        $dstDir = Join-Path $Destination $dir
        if (Test-Path $srcDir) {
            if (Test-Path $dstDir) { Remove-Item -Recurse -Force $dstDir }
            New-Item -ItemType Directory -Force -Path $dstDir | Out-Null
            Get-ChildItem -Path $srcDir -Recurse -File |
                Where-Object { $_.FullName -notlike '*\__pycache__\*' -and $_.Extension -ne '.pyc' } |
                ForEach-Object {
                    $rel = $_.FullName.Substring($srcDir.Length + 1)
                    $outPath = Join-Path $dstDir $rel
                    New-Item -ItemType Directory -Force -Path (Split-Path $outPath -Parent) | Out-Null
                    Copy-Item -Path $_.FullName -Destination $outPath -Force
                }
        }
    }

    $dstScripts = Join-Path $Destination 'scripts'
    New-Item -ItemType Directory -Force -Path $dstScripts | Out-Null
    foreach ($scriptName in $CoreScripts) {
        $srcScript = Join-Path $RepoRoot (Join-Path 'scripts' $scriptName)
        if (Test-Path $srcScript) {
            Copy-Item -Path $srcScript -Destination (Join-Path $dstScripts $scriptName) -Force
        }
    }

    # Copy default template while preserving any user-supplied custom company templates.
    $dstTemplates = Join-Path $Destination 'templates'
    New-Item -ItemType Directory -Force -Path $dstTemplates | Out-Null
    $defaultTpl = Join-Path $RepoRoot 'templates\工作周报模板.xlsx'
    if (Test-Path $defaultTpl) {
        Copy-Item -Path $defaultTpl -Destination (Join-Path $dstTemplates '工作周报模板.xlsx') -Force
    }
}

function Copy-Wrapper {
    param(
        [Parameter(Mandatory = $true)][string]$PlatformSubdir,
        [Parameter(Mandatory = $true)][string]$PluginRoot
    )

    $wrapperSrc = Join-Path $RepoRoot (Join-Path 'platforms' $PlatformSubdir)
    if (-not (Test-Path $wrapperSrc)) {
        throw "Platform wrapper directory not found: $wrapperSrc"
    }

    if ($DryRun) {
        Write-Host "[DryRun] Overlay wrapper ($PlatformSubdir) -> $PluginRoot"
        return
    }

    New-Item -ItemType Directory -Force -Path $PluginRoot | Out-Null
    Get-ChildItem -Path $wrapperSrc -Recurse -File | ForEach-Object {
        $rel = $_.FullName.Substring($wrapperSrc.Length + 1)
        $outPath = Join-Path $PluginRoot $rel
        New-Item -ItemType Directory -Force -Path (Split-Path $outPath -Parent) | Out-Null
        Copy-Item -Path $_.FullName -Destination $outPath -Force
    }
}

$ExpandedTargets = New-Object System.Collections.Generic.List[string]
foreach ($item in $Target) {
    if ($item -eq 'all') {
        foreach ($t in @('pi', 'codex', 'dsh', 'antigravity', 'claude')) {
            if (-not $ExpandedTargets.Contains($t)) { $ExpandedTargets.Add($t) }
        }
    } else {
        if (-not $ExpandedTargets.Contains($item)) { $ExpandedTargets.Add($item) }
    }
}

foreach ($t in $ExpandedTargets) {
    switch ($t) {
        'pi' {
            $dest = Join-Path $HomeDir '.pi\agent\skills\weekly-report'
            Copy-SkillCore -Destination $dest
            Write-Host "Installed [pi] -> $dest (run /reload in Pi)"
        }
        'codex' {
            $dest = Join-Path $HomeDir '.codex\skills\weekly-report'
            Copy-SkillCore -Destination $dest
            Write-Host "Installed [codex] -> $dest (includes agents/openai.yaml)"
        }
        'codex-plugin' {
            $root = Join-Path $HomeDir '.codex\plugins\weekly-report'
            Copy-Wrapper -PlatformSubdir 'codex' -PluginRoot $root
            Copy-SkillCore -Destination (Join-Path $root 'skills\weekly-report')
            Write-Host "Installed [codex-plugin] -> $root"
        }
        'dsh' {
            $dest = Join-Path $HomeDir '.dsh\skills\weekly-report'
            Copy-SkillCore -Destination $dest
            Write-Host "Installed [dsh] -> $dest"
        }
        'dsh-bundle' {
            $root = Join-Path $HomeDir '.dsh\local-plugins\dsh-weekly-report'
            Copy-Wrapper -PlatformSubdir 'dsh' -PluginRoot $root
            Copy-SkillCore -Destination (Join-Path $root 'skills\weekly-report')
            Write-Host "Installed [dsh-bundle] -> $root (register in DSH profile bundles)"
        }
        'antigravity' {
            $root = Join-Path $HomeDir '.gemini\config\plugins\weekly-report'
            Copy-Wrapper -PlatformSubdir 'antigravity' -PluginRoot $root
            Copy-SkillCore -Destination (Join-Path $root 'skills\weekly-report')
            Write-Host "Installed [antigravity plugin] -> $root (run: agy plugin validate `"$root`" --strict)"
        }
        'antigravity-skill' {
            $dest = Join-Path $HomeDir '.gemini\config\skills\weekly-report'
            Copy-SkillCore -Destination $dest
            Write-Host "Installed [antigravity standalone skill] -> $dest"
        }
        'claude' {
            $dest = Join-Path $HomeDir '.claude\skills\weekly-report'
            Copy-SkillCore -Destination $dest
            Write-Host "Installed [claude] -> $dest"
        }
        'claude-plugin' {
            $root = Join-Path $HomeDir '.claude\plugins\weekly-report'
            Copy-Wrapper -PlatformSubdir 'claude-code' -PluginRoot $root
            Copy-SkillCore -Destination (Join-Path $root 'skills\weekly-report')
            Write-Host "Installed [claude-plugin] -> $root"
        }
        'opencode' {
            $dest = Join-Path $HomeDir '.config\opencode\superpowers\skills\weekly-report'
            Copy-SkillCore -Destination $dest
            Write-Host "Installed [opencode] -> $dest"
        }
    }
}
