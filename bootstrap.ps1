# Claude Cowork setup kit - one-line installer.
#
#   irm https://raw.githubusercontent.com/tiger3c/claude-cowork-setup/main/bootstrap.ps1 | iex
#
# Downloads this kit from GitHub to %USERPROFILE%\ClaudeCoworkSetup and runs setup.ps1.
# (Kept ASCII-only on purpose: irm | iex may mis-decode non-ASCII text.)
& {
    $ErrorActionPreference = 'Stop'
    $ProgressPreference = 'SilentlyContinue'
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor 3072

    $repo = 'tiger3c/claude-cowork-setup'
    $dest = Join-Path $env:USERPROFILE 'ClaudeCoworkSetup'
    $work = Join-Path $env:TEMP ('ccs-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Force -Path $work | Out-Null
    try {
        Write-Host "Downloading setup kit from github.com/$repo ..."
        $zip = Join-Path $work 'kit.zip'
        Invoke-WebRequest -UseBasicParsing -Uri "https://github.com/$repo/archive/refs/heads/main.zip" -OutFile $zip
        Expand-Archive -LiteralPath $zip -DestinationPath $work -Force
        $src = Get-ChildItem -LiteralPath $work -Directory | Select-Object -First 1
        if (-not $src -or -not (Test-Path (Join-Path $src.FullName 'setup.ps1'))) { throw 'setup.ps1 not found in the downloaded kit.' }
        # Merge into the kit folder (keeps installers already placed in files\)
        $null = & robocopy.exe $src.FullName $dest /E /NFL /NDL /NJH /NJS /NP
        if ($LASTEXITCODE -ge 8) { throw "Copy failed (robocopy exit code $LASTEXITCODE)." }
    } finally {
        Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
    }
    Write-Host "Kit saved to $dest"
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $dest 'setup.ps1')
}
