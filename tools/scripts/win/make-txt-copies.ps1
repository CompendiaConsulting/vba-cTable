#requires -Version 5.1
<#
Run from the repository root:
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\scripts\win\make-txt-copies.ps1
Uses the bundled Windows PowerShell; no modules or packages required.
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

try {
    $repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../..'))
    $output = Join-Path $repoRoot 'txt-copies'
    $reparsePoint = [IO.FileAttributes]::ReparsePoint

    # Walk explicitly so links/junctions and excluded folders are never followed.
    function Get-VbaSources([string] $Directory) {
        foreach ($item in Get-ChildItem -LiteralPath $Directory -Force) {
            if (($item.Attributes -band $reparsePoint) -ne 0) { continue }
            if ($item.PSIsContainer) {
                if ($item.Name -eq '.git' -or $item.FullName -eq $output) { continue }
                Get-VbaSources $item.FullName
            }
            elseif ($item.Extension -ieq '.bas' -or $item.Extension -ieq '.cls') {
                $item
            }
        }
    }

    # Refuse links within an old output tree before recursive removal.
    function Assert-PlainDirectory([string] $Directory) {
        foreach ($item in Get-ChildItem -LiteralPath $Directory -Force) {
            if (($item.Attributes -band $reparsePoint) -ne 0) {
                throw "Remove the link/junction from txt-copies before rerunning: $($item.FullName)"
            }
            if ($item.PSIsContainer) { Assert-PlainDirectory $item.FullName }
        }
    }

    $existing = Get-Item -LiteralPath $output -Force -ErrorAction SilentlyContinue
    if ($null -ne $existing) {
        if (-not $existing.PSIsContainer -or ($existing.Attributes -band $reparsePoint) -ne 0) {
            throw 'txt-copies must be a real directory, not a link or file.'
        }
        Assert-PlainDirectory $output
    }

    $sources = @(Get-VbaSources $repoRoot)
    if ($null -ne $existing) { Remove-Item -LiteralPath $output -Recurse -Force }
    [void][IO.Directory]::CreateDirectory($output)

    $count = 0
    foreach ($source in $sources) {
        $relative = $source.FullName.Substring($repoRoot.Length).TrimStart([char[]]'\/')
        $stem = $relative.Substring(0, $relative.Length - $source.Extension.Length)
        $suffix = $source.Extension.Substring(1).ToLowerInvariant()
        $destination = Join-Path $output ($stem + '_' + $suffix + '.txt')
        [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($destination))
        # Byte-for-byte copy; false prevents silent destination collisions.
        [IO.File]::Copy($source.FullName, $destination, $false)
        $count++
    }
    Write-Host "Created $count text copies in $output"
}
catch {
    [Console]::Error.WriteLine("Error: " + $_.Exception.Message)
    exit 1
}
