#Requires -Version 5.1
<#
.SYNOPSIS
    Claude Cowork 원클릭 설치 도우미 (Windows)

.DESCRIPTION
    1) PC 사양 / 가상화 점검
    2) 인터넷 연결 확인
    3) Virtual Machine Platform 켜기 (Cowork 필수)
    4) Claude Desktop(.msix) 설치
         - files\ 폴더나 다운로드 폴더에 설치 파일이 있으면 그걸 쓰고 (오프라인판 우선)
         - 없으면 Anthropic 공식 주소에서 받아서 서명/게시자를 확인한 뒤 설치
    5) Claude Code 설치 (가상화가 안 되는 PC에서는 Cowork 대신 쓰는 대안)
    6) 재시작이 필요하면 재시작 후 자동으로 다시 떠서 마무리

    여러 번 실행해도 안전합니다. 이미 끝난 단계는 건너뜁니다.

.PARAMETER CheckOnly
    아무것도 바꾸지 않고 점검만 합니다. 관리자 권한이 필요 없습니다.
.PARAMETER SkipClaudeCode
    Claude Code 설치를 묻지 않고 건너뜁니다.
.PARAMETER ClaudeCodeOnly
    Claude Code만 설치합니다 (관리자 권한 불필요).
.PARAMETER Resumed
    재시작 후 자동 실행될 때 내부적으로 쓰는 옵션입니다.

.NOTES
    Windows PowerShell 5.1 기준. 종료 코드: 0 = 정상, 2 = 문제 있음.
#>
[CmdletBinding()]
param(
    [switch]$CheckOnly,
    [switch]$SkipClaudeCode,
    [switch]$ClaudeCodeOnly,
    [switch]$Resumed,
    [long]$RegisteredAt = 0,   # 내부용: 자동 재개를 등록한 시각 (UTC ticks)
    [switch]$Elevated          # 내부용: 관리자 권한 요청 후 다시 실행된 것
)

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

# =============================================================== 상수 / 상태
$PackageName = 'Claude'
$PublisherId = 'pzs8sxrjxfjjc'     # Anthropic 서명 게시자 해시 (패키지 패밀리 이름 뒷부분)
$PFN         = "${PackageName}_${PublisherId}"
$MinBuild    = 19041               # Windows 10 2004
$RunOnceKey  = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\RunOnce'
$RunOnceName = 'ClaudeCoworkSetup'
$CpuArch     = if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64' -or $env:PROCESSOR_ARCHITEW6432 -eq 'ARM64') { 'arm64' } else { 'x64' }
$Urls = @{
    Msix          = "https://claude.ai/api/desktop/win32/$CpuArch/msix/latest/redirect"
    Offline       = "https://claude.ai/api/desktop/win32/$CpuArch/offline/latest/redirect"
    Readiness     = "https://claude.ai/api/desktop/win32/$CpuArch/cowork-readiness-check/latest/redirect"
    CodeInstaller = 'https://claude.ai/install.ps1'
    CodeLatest    = 'https://downloads.claude.ai/claude-code-releases/latest'
}
$KitRoot       = $PSScriptRoot
$FilesDir      = Join-Path $KitRoot 'files'
$LogDir        = Join-Path $KitRoot 'logs'
$ClaudeCodeExe = Join-Path $env:USERPROFILE '.local\bin\claude.exe'

$Report = [ordered]@{
    Virtualization = 'Unknown'   # OK / Off / Unknown
    NeedRestart    = $false
    Fatal          = $false
    Online         = $false
    Desktop        = '-'
    DesktopOffline = $false
    ClaudeCode     = '-'
    Readiness      = '-'
}
$script:LogFile        = $null
$script:IsAdmin        = $false
$script:SameUser       = $true
$script:LatestMsix     = $null
$script:DeclinedCode   = $false
$script:DownloadedMsix = $null
$script:IsVm              = $false
$script:HypervisorPresent = $false
$script:HypervisorMissing = $false
$script:VirtError         = $null
$script:RestartFailed     = $false

# =============================================================== 출력 / 입력
function Write-Ok   ([string]$m) { Write-Host "  [OK]   $m" -ForegroundColor Green }
function Write-Warn ([string]$m) { Write-Host "  [주의] $m" -ForegroundColor Yellow }
function Write-Bad  ([string]$m) { Write-Host "  [실패] $m" -ForegroundColor Red }
function Write-Info ([string]$m) { Write-Host "         $m" }
function Write-Step ([string]$m) { Write-Host ''; Write-Host "== $m" -ForegroundColor Cyan }

function Read-YesNo([string]$Question, [bool]$Default = $true) {
    if ($CheckOnly) { return $false }
    $hint = if ($Default) { '[Y/n]' } else { '[y/N]' }
    $a = Read-Host "  $Question $hint"
    if ([string]::IsNullOrWhiteSpace($a)) { return $Default }
    # 한글 자판 상태에서 y/n 을 누르면 'ㅛ'/'ㅜ' 가 입력되므로 같이 받아줌
    return ($a.Trim() -match '^(y|yes|ㅛ|예|네|응)$')
}

function Stop-Log { try { Stop-Transcript | Out-Null } catch {} }

function Exit-Setup([int]$Code) {
    Stop-Log
    if ($script:LogFile) { Write-Host ''; Write-Host "  로그 파일: $($script:LogFile)" -ForegroundColor DarkGray }
    if (-not $CheckOnly) { [void](Read-Host "`n  Enter 키를 누르면 창이 닫힙니다") }
    exit $Code
}

function Start-Log {
    foreach ($d in @($LogDir, (Join-Path $env:TEMP 'ClaudeCoworkSetup'))) {
        try {
            New-Item -ItemType Directory -Force -Path $d | Out-Null
            $f = Join-Path $d ('setup-{0:yyyyMMdd-HHmmss}.log' -f (Get-Date))
            Start-Transcript -Path $f -Force | Out-Null
            $script:LogFile = $f
            return
        } catch {}
    }
}

# 네이티브 명령(bcdedit, dism, curl) 은 stderr 를 오류로 취급하지 않도록 실행
function Invoke-Native([string]$Exe, [string[]]$Arguments, [switch]$Show) {
    $old = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        if ($Show) { & $Exe @Arguments | Out-Host; return }
        $out = & $Exe @Arguments
        return ,@($out)
    } finally { $ErrorActionPreference = $old }
}

# =============================================================== 권한 / 환경
function Test-IsAdmin {
    $p = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-PassThroughArgs {
    $a = @()
    if ($SkipClaudeCode)   { $a += '-SkipClaudeCode' }
    if ($Resumed)          { $a += '-Resumed' }
    if ($RegisteredAt -gt 0) { $a += @('-RegisteredAt', "$RegisteredAt") }
    return $a
}

function Request-Elevation {
    Write-Host ''
    Write-Host '  관리자 권한이 필요합니다. 잠시 후 뜨는 창에서 [예]를 눌러주세요.' -ForegroundColor Yellow
    $argLine = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Elevated " + ((Get-PassThroughArgs) -join ' ')
    try {
        Start-Process -FilePath powershell.exe -Verb RunAs -ArgumentList $argLine | Out-Null
        exit 0
    } catch {
        Write-Host '  관리자 권한을 받지 못했습니다. 다시 실행해서 [예]를 눌러주세요.' -ForegroundColor Red
        [void](Read-Host '  Enter 키를 누르면 창이 닫힙니다')
        exit 2
    }
}

function Test-SameUser {
    # UAC 에서 다른 관리자 계정 암호를 넣은 경우: 사용자별 설치/PATH/자동 재개를 그 계정에 하면 안 됨
    try {
        $console = (Get-CimInstance Win32_ComputerSystem).UserName
        if (-not $console) { return $true }
        return ($console -ieq [Security.Principal.WindowsIdentity]::GetCurrent().Name)
    } catch { return $true }
}

function Get-KitDriveType {
    try {
        if ($KitRoot -like '\\*') { return 'Network' }
        return [string]([IO.DriveInfo]::new((Split-Path $KitRoot -Qualifier))).DriveType
    } catch { return 'Unknown' }
}

function Initialize-Network {
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch {}
    # 회사 프록시(로그인 필요한 프록시) 대응
    try { [Net.WebRequest]::DefaultWebProxy.Credentials = [Net.CredentialCache]::DefaultNetworkCredentials } catch {}
}

# =============================================================== 다운로드
function Resolve-Download([string]$Url) {
    # 공식 'latest' 주소가 가리키는 실제 파일 주소를 (파일은 받지 않고) 알아냄
    $req = [Net.HttpWebRequest]::Create($Url)
    $req.Method = 'GET'
    $req.AllowAutoRedirect = $false
    $req.Timeout = 30000
    $resp = $null
    try { $resp = $req.GetResponse() }
    catch {
        $ex = $_.Exception
        while ($ex -and -not ($ex -is [Net.WebException])) { $ex = $ex.InnerException }
        if ($ex -and $ex.Response) { $resp = $ex.Response } else { throw }
    }
    try { $code = [int]$resp.StatusCode; $loc = $resp.Headers['Location'] } finally { $resp.Close() }
    if ($code -lt 300 -or $code -ge 400 -or -not $loc) { throw "다운로드 주소를 받지 못했습니다 (HTTP $code)" }
    if ($loc -notmatch '^https://') { throw "https 가 아닌 주소라서 거부합니다: $loc" }
    $ver = $null
    if ($loc -match '/(\d+(?:\.\d+){1,3})/[^/]+$') { try { $ver = [version]$Matches[1] } catch {} }
    return [pscustomobject]@{
        Url      = $loc
        Version  = $ver
        FileName = [IO.Path]::GetFileName(([uri]$loc).AbsolutePath)
    }
}

function Save-Download([string]$Url, [string]$Dest) {
    $part = "$Dest.partial"
    $curl = Join-Path $env:windir 'System32\curl.exe'
    if (Test-Path $curl) {
        # 진행률 표시 + 이어받기 + 재시도
        [void](Invoke-Native $curl @('-L', '--fail', '--retry', '3', '--retry-delay', '5', '-C', '-', '-o', $part, $Url))
        if ($LASTEXITCODE -eq 0 -and (Test-Path $part)) { Move-Item -Force $part $Dest; return }
        Write-Warn "curl 로 받기 실패 (코드 $LASTEXITCODE). 다른 방법으로 다시 시도합니다."
    }
    Remove-Item $part -Force -ErrorAction SilentlyContinue
    try {
        Import-Module BitsTransfer -ErrorAction Stop
        $ProgressPreference = 'Continue'
        Start-BitsTransfer -Source $Url -Destination $part -DisplayName 'Claude 설치 파일' -ErrorAction Stop
        Move-Item -Force $part $Dest
        return
    } catch {
        Write-Warn "BITS 로 받기 실패: $($_.Exception.Message)"
    }
    $ProgressPreference = 'SilentlyContinue'   # PS 5.1 의 Invoke-WebRequest 는 진행률 표시가 켜져 있으면 매우 느림
    Remove-Item $part -Force -ErrorAction SilentlyContinue
    Write-Info '마지막 방법으로 받는 중... (진행률이 표시되지 않습니다)'
    Invoke-WebRequest -UseBasicParsing -Uri $Url -OutFile $part
    Move-Item -Force $part $Dest
}

function Get-DownloadDir {
    foreach ($d in @($FilesDir, (Join-Path $env:TEMP 'ClaudeCoworkSetup'))) {
        try {
            New-Item -ItemType Directory -Force -Path $d | Out-Null
            $t = Join-Path $d '.write-test'
            [IO.File]::WriteAllText($t, 'x')
            Remove-Item $t -Force
            return $d
        } catch {}
    }
    throw '설치 파일을 내려받을 폴더를 만들 수 없습니다.'
}

function Get-DownloadsFolder {
    try {
        $p = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
        if ($p) { return $p }
    } catch {}
    return (Join-Path $env:USERPROFILE 'Downloads')
}

function Get-SearchDirs {
    return @($FilesDir, (Get-DownloadsFolder)) | Select-Object -Unique | Where-Object { $_ -and (Test-Path -LiteralPath $_) }
}

# =============================================================== MSIX 확인
function Get-PublisherIdFromName([string]$Publisher) {
    # Windows 패키지 패밀리 이름의 게시자 해시: SHA-256(UTF-16LE) 앞 8바이트 -> base32(13자)
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $hash = $sha.ComputeHash([Text.Encoding]::Unicode.GetBytes($Publisher)) } finally { $sha.Dispose() }
    $bits = (-join ($hash[0..7] | ForEach-Object { [Convert]::ToString($_, 2).PadLeft(8, '0') })) + '0'
    $alphabet = '0123456789abcdefghjkmnpqrstvwxyz'
    return (-join (0..12 | ForEach-Object { $alphabet[[Convert]::ToInt32($bits.Substring($_ * 5, 5), 2)] }))
}

function Get-MsixInfo([string]$Path) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $zip.GetEntry('AppxManifest.xml')
        if (-not $entry) { return $null }
        $reader = New-Object IO.StreamReader($entry.Open())
        try { [xml]$xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
    } finally { $zip.Dispose() }
    $id   = $xml.Package.Identity
    $size = (Get-Item -LiteralPath $Path).Length
    return [pscustomobject]@{
        Path      = $Path
        Name      = [string]$id.Name
        Publisher = [string]$id.Publisher
        Version   = [version]$id.Version
        Arch      = [string]$id.ProcessorArchitecture
        SizeGB    = [math]::Round($size / 1GB, 2)
        IsOffline = ($size -ge 1.5GB)   # 일반판 약 0.3GB, 오프라인판은 수 GB
    }
}

function Test-IsClaudeMsix($Info) {
    return ($Info -and $Info.Name -eq $PackageName -and (Get-PublisherIdFromName $Info.Publisher) -eq $PublisherId)
}

function Find-ClaudeMsix {
    $found = @()
    foreach ($d in Get-SearchDirs) {
        $filter = if ($d -eq $FilesDir) { '*.msix' } else { 'Claude*.msix' }
        foreach ($f in @(Get-ChildItem -LiteralPath $d -Filter $filter -File -ErrorAction SilentlyContinue)) {
            try { $i = Get-MsixInfo $f.FullName }
            catch { Write-Warn "$($f.Name): 열 수 없는 파일이라 건너뜁니다 (다운로드가 덜 됐거나 손상)."; continue }
            if (-not (Test-IsClaudeMsix $i)) { continue }
            if ($i.Arch -and $i.Arch -ne 'neutral' -and $i.Arch -ne $CpuArch) {
                Write-Warn "$($f.Name): $($i.Arch) 용 파일입니다. 이 PC는 $CpuArch 라서 건너뜁니다."
                continue
            }
            $found += $i
        }
    }
    if ($found.Count -eq 0) { return $null }
    return ($found | Sort-Object -Property @{ Expression = { $_.Version }; Descending = $true }, @{ Expression = { $_.SizeGB }; Descending = $true } | Select-Object -First 1)
}

function Find-ReadinessCheck {
    foreach ($d in Get-SearchDirs) {
        $cands = @(Get-ChildItem -LiteralPath $d -Filter '*.exe' -File -ErrorAction SilentlyContinue |
                   Where-Object { $d -eq $FilesDir -or $_.Name -match 'readiness|cowork' } |
                   Sort-Object LastWriteTime -Descending)
        foreach ($f in $cands) {
            $sig = Get-AuthenticodeSignature -FilePath $f.FullName
            if ($sig.Status -eq 'Valid' -and $sig.SignerCertificate.Subject -match '(^|,\s*)O="?Anthropic, PBC') { return $f.FullName }
        }
    }
    return $null
}

function Get-InstalledClaude {
    $p = @(Get-AppxPackage -Name $PackageName -ErrorAction SilentlyContinue | Where-Object { $_.PackageFamilyName -eq $PFN })
    if ($p.Count -eq 0) { return $null }
    return ($p | Sort-Object -Property @{ Expression = { [version]$_.Version }; Descending = $true } | Select-Object -First 1)
}

function Show-InstallHint([string]$Message) {
    $hints = [ordered]@{
        '0x80073D28'                         = '관리자 권한이 필요합니다. setup.bat 으로 다시 실행하고 [예]를 누르세요.'
        '0x80073D02'                         = 'Claude 앱이 실행 중입니다. 작업 표시줄 오른쪽 아래 Claude 아이콘에서 완전히 종료한 뒤 다시 실행하세요.'
        '0x80073CF6|0x80073CFB|0x80073CF3'   = '기존 Claude 설치와 충돌합니다. [설정 > 앱 > 설치된 앱]에서 Claude를 제거한 뒤 다시 실행하세요.'
        '0x800B0109|0x800B010A|0x800B0101|0x80096004' = '인증서 확인에 실패했습니다. PC 날짜/시간이 맞는지 확인하고 Windows 업데이트 후 다시 실행하세요.'
        '0x80070070|0x80070027'              = '디스크 공간이 부족합니다. C: 드라이브를 정리한 뒤 다시 실행하세요.'
        '0x80073CFF'                         = '이 PC는 정책상 앱 직접 설치(사이드로드)가 막혀 있습니다. 회사 PC라면 IT 담당자에게 문의하세요.'
    }
    foreach ($k in $hints.Keys) { if ($Message -match $k) { Write-Info "-> $($hints[$k])"; return } }
    Write-Info '-> README.md 의 [문제 해결]에서 오류 코드를 찾아보거나, 로그 파일을 Claude에게 보여주세요.'
}

# =============================================================== 1. PC 점검
function Test-System {
    Write-Step '1. PC 점검'
    $os  = Get-CimInstance Win32_OperatingSystem
    $cs  = Get-CimInstance Win32_ComputerSystem
    $cpu = @(Get-CimInstance Win32_Processor)[0]

    $build = [int]$os.BuildNumber
    if ($build -ge $MinBuild) { Write-Ok "$($os.Caption) (빌드 $build)" }
    else {
        Write-Bad "$($os.Caption) (빌드 $build): Windows 10 2004(빌드 $MinBuild) 이상이 필요합니다. Windows 업데이트부터 하세요."
        $Report.Fatal = $true
    }
    if (-not [Environment]::Is64BitOperatingSystem) { Write-Bad '32비트 Windows는 지원하지 않습니다.'; $Report.Fatal = $true }
    if (@(98, 99, 100, 101) -contains [int]$os.OperatingSystemSKU) {
        Write-Warn 'Home 에디션입니다. Cowork이 되는 PC도, 안 되는 PC도 있다는 보고가 있습니다. 안 되면 Claude Code를 쓰세요.'
    }

    Write-Ok ("CPU: {0} ({1})" -f "$($cpu.Name)".Trim(), $CpuArch)

    $ramGB = [math]::Round($cs.TotalPhysicalMemory / 1GB)
    if ($ramGB -ge 8) { Write-Ok "메모리 ${ramGB}GB" } else { Write-Warn "메모리 ${ramGB}GB: 8GB 이상 권장 (느리거나 Cowork이 안 될 수 있음)" }

    $drives = @($env:SystemDrive, (Split-Path $env:LOCALAPPDATA -Qualifier)) | ForEach-Object { $_.TrimEnd(':').ToUpper() } | Select-Object -Unique
    foreach ($dl in $drives) {
        $freeGB = [math]::Floor((Get-PSDrive -Name $dl).Free / 1GB)
        if ($freeGB -ge 20) { Write-Ok "${dl}: 드라이브 여유 공간 ${freeGB}GB" }
        else { Write-Warn "${dl}: 드라이브 여유 공간 ${freeGB}GB: 20GB 이상 권장" }
    }

    $isVm = ("$($cs.Model)" -match 'Virtual|VMware|VirtualBox|KVM|QEMU|HVM domU|Parallels|Cloud PC') -or
            ("$($cs.Manufacturer)" -match 'QEMU|Xen|VMware|innotek|Parallels|Amazon EC2|Google')
    $script:IsVm              = $isVm
    $script:HypervisorPresent = [bool]$cs.HypervisorPresent
    if ($isVm) {
        Write-Warn "가상 PC(VM/VDI/클라우드 PC) 안에서 실행 중입니다: $($cs.Manufacturer) $($cs.Model)"
        Write-Info '호스트에서 "중첩 가상화"를 허용해야 Cowork이 됩니다. 일단 계속 진행합니다.'
        $Report.Virtualization = 'Unknown'
    } elseif ($cs.HypervisorPresent) {
        Write-Ok '가상화 사용 가능 (하이퍼바이저 동작 중)'
        $Report.Virtualization = 'OK'
    } elseif ($cpu.VirtualizationFirmwareEnabled -eq $true) {
        Write-Ok 'BIOS 가상화(Intel VT-x / AMD-V) 켜짐'
        $Report.Virtualization = 'OK'
    } elseif ($cpu.VirtualizationFirmwareEnabled -eq $false) {
        Write-Bad 'BIOS/UEFI 에서 가상화가 꺼져 있습니다. 이대로는 Cowork을 쓸 수 없습니다 (Claude Code는 가능).'
        Write-Info 'README.md 의 [BIOS 가상화 켜기]를 따라 켜면 Cowork도 쓸 수 있습니다.'
        $Report.Virtualization = 'Off'
    } else {
        Write-Warn '가상화 상태를 확인하지 못했습니다. 일단 계속 진행합니다.'
        $Report.Virtualization = 'Unknown'
    }
}

# =============================================================== 2. 인터넷
function Test-Network {
    Write-Step '2. 인터넷 연결'
    try {
        $script:LatestMsix = Resolve-Download $Urls.Msix
        Write-Ok "claude.ai 연결됨 (최신 Claude Desktop $($script:LatestMsix.Version))"
        $Report.Online = $true
    } catch {
        Write-Warn "claude.ai 에 연결하지 못했습니다: $($_.Exception.Message)"
        Write-Info '로그인과 다운로드에 인터넷이 필요합니다. 회사망이면 프록시/방화벽을 확인하세요.'
        return
    }
    try {
        [void](Invoke-WebRequest -UseBasicParsing -Uri $Urls.CodeLatest -TimeoutSec 20)
        Write-Ok 'downloads.claude.ai 연결됨'
    } catch {
        Write-Warn 'downloads.claude.ai 에 연결하지 못했습니다. Cowork 실행 환경과 Claude Code 다운로드가 막힐 수 있습니다.'
    }
}

# =============================================================== 3. 가상화 설정
function Get-VmpState {
    try { return [string](Get-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform -ErrorAction Stop).State }
    catch {
        # 관리자 권한이 없을 때 (점검 모드)
        $f = Get-CimInstance Win32_OptionalFeature -Filter "Name='VirtualMachinePlatform'" -ErrorAction SilentlyContinue
        if (-not $f) { return 'Absent' }
        switch ([int]$f.InstallState) { 1 { return 'Enabled' } 2 { return 'Disabled' } default { return 'Absent' } }
    }
}

function Enable-Vmp {
    try {
        Enable-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform -All -NoRestart -ErrorAction Stop | Out-Null
    } catch {
        Write-Info "PowerShell 방식이 실패해서 DISM 으로 다시 시도합니다: $($_.Exception.Message)"
        [void](Invoke-Native 'dism.exe' @('/online', '/enable-feature', '/featurename:VirtualMachinePlatform', '/all', '/norestart', '/quiet'))
        if (@(0, 3010) -notcontains $LASTEXITCODE) { throw "Virtual Machine Platform 을 켜지 못했습니다 (DISM 코드 $LASTEXITCODE)" }
    }
}

function Set-Virtualization {
    Write-Step '3. Cowork용 가상화 설정'
    if ($Report.Virtualization -eq 'Off') { Write-Info 'BIOS 가상화가 꺼져 있어서 건너뜁니다.'; return }

    $vmp = Get-VmpState
    switch ($vmp) {
        'Enabled'       { Write-Ok 'Virtual Machine Platform 켜짐' }
        'EnablePending' { Write-Warn '켜기가 예약되어 있습니다. 재시작하면 적용됩니다.'; $Report.NeedRestart = $true }
        default {
            if ($CheckOnly) { Write-Warn "Virtual Machine Platform 꺼짐 ($vmp): 설치할 때 자동으로 켭니다." }
            else {
                Write-Info 'Virtual Machine Platform 켜는 중... (1~2분)'
                Enable-Vmp
                Write-Ok 'Virtual Machine Platform 켰습니다 (재시작 후 적용)'
                $Report.NeedRestart = $true
            }
        }
    }

    if ($vmp -eq 'Enabled') {
        $svc = @(Get-Service -Name vmcompute, hns -ErrorAction SilentlyContinue)
        if ($svc.Count -ge 2) { Write-Ok '가상화 서비스(vmcompute, hns) 확인' }
        else { Write-Warn '가상화 서비스(vmcompute, hns)가 안 보입니다. Cowork이 안 되면 PC를 [다시 시작] 해보세요.' }
        if (-not $script:HypervisorPresent -and -not $script:IsVm) {
            Write-Warn '가상화 기능은 켜져 있지만 하이퍼바이저가 아직 동작하지 않습니다.'
            Write-Info '[다시 시작]을 안 했다면 재시작하세요. 이미 했다면 아래 하이퍼바이저 설정과 BIOS 가상화를 확인합니다.'
            $script:HypervisorMissing = $true
        }
    }

    if (-not $script:IsAdmin) { return }   # bcdedit 은 관리자 전용
    $bcd = (Invoke-Native 'bcdedit.exe' @('/enum', '{current}')) -join "`n"
    if ($bcd -match '(?im)^\s*hypervisorlaunchtype\s+Off\b') {
        Write-Warn '하이퍼바이저가 꺼져 있습니다 (hypervisorlaunchtype = Off).'
        Write-Info 'VMware/VirtualBox/안드로이드 에뮬레이터/일부 게임 보안 프로그램이 꺼둔 경우가 많습니다.'
        Write-Info 'Cowork을 쓰려면 켜야 합니다. (오래된 VMware/VirtualBox는 영향을 받을 수 있음)'
        if (Read-YesNo '하이퍼바이저를 켤까요?' $true) {
            [void](Invoke-Native 'bcdedit.exe' @('/set', '{current}', 'hypervisorlaunchtype', 'auto'))
            if ($LASTEXITCODE -eq 0) { Write-Ok '하이퍼바이저 켬 (재시작 후 적용)'; $Report.NeedRestart = $true }
            else { Write-Bad "bcdedit 실패 (코드 $LASTEXITCODE)" }
        } else {
            $Report.Virtualization = 'Off'
        }
    } elseif ($bcd -and $LASTEXITCODE -eq 0) {
        Write-Ok '하이퍼바이저 부팅 설정 정상'
        # 설정은 정상인데 하이퍼바이저가 안 돌고 있으면 재시작이 필요한 상태
        if ($script:HypervisorMissing) { $Report.NeedRestart = $true }
    } else {
        Write-Warn '하이퍼바이저 부팅 설정을 확인하지 못했습니다 (bcdedit 실패). 계속 진행합니다.'
    }
}

# =============================================================== 예전 .exe 버전
function Remove-LegacyClaude {
    $dir = Join-Path $env:LOCALAPPDATA 'AnthropicClaude'
    $u   = Join-Path $dir 'Update.exe'
    if (-not (Test-Path $u)) { return }
    # 삭제 후에도 Update.exe 와 .dead 표시가 남으므로, 실제 앱 폴더(app-*)가 있을 때만 예전 버전으로 판단
    if ((Test-Path (Join-Path $dir '.dead')) -or -not (Get-ChildItem -LiteralPath $dir -Directory -Filter 'app-*' -ErrorAction SilentlyContinue)) { return }
    Write-Step '예전 버전 감지'
    Write-Warn '예전 방식(.exe)으로 설치된 Claude가 있습니다. 이 버전에는 Cowork이 없고, 새 버전과 따로 남습니다.'
    Write-Info '(대화 기록은 Claude 계정에 저장되어 있어서 지워지지 않습니다)'
    if (Read-YesNo '예전 버전을 삭제할까요? 먼저 Claude 앱을 닫아주세요.' $true) {
        Invoke-Native $u @('--uninstall', '-s') -Show
        if ($LASTEXITCODE -eq 0) { Write-Ok '예전 버전 삭제 완료' }
        else { Write-Warn "예전 버전 삭제가 끝나지 않았습니다 (코드 $LASTEXITCODE). [설정 > 앱]에서 직접 제거해도 됩니다." }
    }
}

# =============================================================== 4. Claude Desktop
function Install-ClaudeDesktop {
    Write-Step '4. Claude Desktop 설치'
    $installed = Get-InstalledClaude
    if ($installed) {
        Write-Ok "이미 설치되어 있습니다 (버전 $($installed.Version)). 업데이트는 앱이 알아서 합니다."
        $Report.Desktop = "설치됨 ($($installed.Version))"
        return
    }

    $msix = Find-ClaudeMsix
    if ($msix) {
        $kind = if ($msix.IsOffline) { '오프라인판' } else { '일반판' }
        Write-Ok ("설치 파일 찾음: {0} (버전 {1}, {2}GB, {3})" -f (Split-Path $msix.Path -Leaf), $msix.Version, $msix.SizeGB, $kind)
    }

    if ($CheckOnly) {
        if (-not $msix) {
            if ($Report.Online) { Write-Info "설치 파일이 없습니다. 설치할 때 공식 주소에서 자동으로 받습니다 (약 300MB, 버전 $($script:LatestMsix.Version))." }
            else { Write-Bad '설치 파일도 없고 인터넷도 안 됩니다. README.md 의 [방법 C: USB]를 보세요.' }
        }
        $Report.Desktop = '미설치'
        return
    }

    if (-not $msix) {
        if (-not $Report.Online) { throw '설치 파일이 없고 인터넷에도 연결되지 않습니다. README.md 의 [방법 C: USB]를 보세요.' }
        $r = if ($script:LatestMsix) { $script:LatestMsix } else { Resolve-Download $Urls.Msix }
        $dest = Join-Path (Get-DownloadDir) $r.FileName
        Write-Info "공식 주소에서 Claude Desktop $($r.Version) 내려받는 중... (약 300MB)"
        Save-Download $r.Url $dest
        $script:DownloadedMsix = $dest
        $msix = Get-MsixInfo $dest
        if (-not (Test-IsClaudeMsix $msix)) { throw '받은 파일이 Claude 설치 파일이 아닙니다. 잠시 후 다시 실행하세요.' }
        Write-Ok '다운로드 완료'
    }

    Write-Info '서명 확인 중... (파일이 크면 1~2분 걸립니다)'
    $sig = Get-AuthenticodeSignature -FilePath $msix.Path
    switch ([string]$sig.Status) {
        'Valid'                  { Write-Ok 'Anthropic 공식 서명 확인' }
        'NotSupportedFileFormat' { Write-Warn '이 Windows에서는 미리 서명을 확인할 수 없습니다. 설치할 때 Windows가 확인합니다.' }
        default {
            Write-Bad "설치 파일의 서명이 올바르지 않습니다 ($($sig.Status))."
            Show-InstallHint '0x800B0109'
            $Report.Desktop = '설치 실패 (서명)'
            return
        }
    }

    Write-Info '설치 중... (몇 분 걸릴 수 있습니다)'
    $provErr = $null
    $userErr = $null
    try { Add-AppxProvisionedPackage -Online -PackagePath $msix.Path -SkipLicense | Out-Null }   # 모든 사용자 + Cowork 서비스 등록
    catch { $provErr = $_.Exception.Message }
    if ($script:SameUser) {
        try { Add-AppxPackage -Path $msix.Path -ForceApplicationShutdown }                       # 지금 계정에 바로 등록
        catch { $userErr = $_.Exception.Message }
    }

    $installed = Get-InstalledClaude
    $allErr = (@($provErr, $userErr) | Where-Object { $_ }) -join ' / '
    if ($installed -or $allErr -match '0x80073D06') {
        $v = if ($installed) { $installed.Version } else { '최신' }
        Write-Ok "Claude Desktop 설치 완료 (버전 $v)"
        if ($provErr -and $provErr -notmatch '0x80073D06') {
            Write-Warn "모든 사용자용 등록은 실패했습니다. Cowork이 안 되면 setup.bat 을 다시 실행하세요: $provErr"
        }
        $Report.Desktop = "설치 완료 ($v)"
        $Report.DesktopOffline = $msix.IsOffline
        if ($script:DownloadedMsix) { Remove-Item -LiteralPath $script:DownloadedMsix -Force -ErrorAction SilentlyContinue }
    } elseif (-not $provErr) {
        # 모든 사용자용 설치는 됐지만 지금 계정에는 아직 등록되지 않음 -> 다시 로그인하면 등록됨
        Write-Ok '모든 사용자용으로 설치했습니다.'
        if ($userErr) { Write-Warn "지금 계정 등록은 실패했습니다: $userErr" }
        Write-Warn '시작 메뉴에 Claude가 안 보이면 로그아웃 후 다시 로그인하세요.'
        $Report.Desktop = '설치 완료 (다시 로그인 필요)'
    } else {
        Write-Bad "설치 오류: $allErr"
        Show-InstallHint $allErr
        $Report.Desktop = '설치 실패'
    }
}

# =============================================================== 5. Claude Code
function Send-EnvironmentChanged {
    try {
        if (-not ('ClaudeSetup.NativeMethods' -as [type])) {
            Add-Type -Namespace ClaudeSetup -Name NativeMethods -MemberDefinition @'
[DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
public static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg, UIntPtr wParam, string lParam, uint fuFlags, uint uTimeout, out UIntPtr lpdwResult);
'@
        }
        $r = [UIntPtr]::Zero
        [void][ClaudeSetup.NativeMethods]::SendMessageTimeout([IntPtr]0xffff, 0x1A, [UIntPtr]::Zero, 'Environment', 2, 5000, [ref]$r)
    } catch {}
}

function Add-UserPath {
    $bin = Split-Path $ClaudeCodeExe
    $key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('Environment', $true)
    try {
        $raw = [string]$key.GetValue('Path', '', [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
        $parts = @($raw -split ';' | Where-Object { $_ })
        foreach ($p in $parts) {
            if ([Environment]::ExpandEnvironmentVariables($p).TrimEnd('\') -ieq $bin.TrimEnd('\')) { return }
        }
        if ($CheckOnly) { Write-Warn "PATH 에 $bin 이 없습니다 (설치할 때 자동으로 추가)."; return }
        $key.SetValue('Path', (($parts + $bin) -join ';'), [Microsoft.Win32.RegistryValueKind]::ExpandString)
    } finally { $key.Close() }
    Send-EnvironmentChanged
    Write-Ok "PATH 에 $bin 추가 (새로 여는 터미널부터 claude 명령 사용 가능)"
}

function Install-ClaudeCode([bool]$Recommended, [bool]$NoPrompt = $false) {
    Write-Step '5. Claude Code (터미널용 Claude)'
    if (Test-Path $ClaudeCodeExe) {
        Write-Ok "이미 설치되어 있습니다: $ClaudeCodeExe"
        Add-UserPath
        $Report.ClaudeCode = '설치됨'
        return
    }
    if ($CheckOnly) { Write-Info '설치되어 있지 않습니다.'; $Report.ClaudeCode = '미설치'; return }
    if ($SkipClaudeCode) { Write-Info '건너뜁니다.'; $Report.ClaudeCode = '건너뜀'; return }
    if (-not $script:SameUser) {
        Write-Warn '관리자 계정이 지금 로그인한 계정과 달라서 여기서는 설치하지 않습니다.'
        Write-Info '대신 install-claude-code.bat 을 더블클릭하세요 (관리자 권한 불필요).'
        $Report.ClaudeCode = 'install-claude-code.bat 으로 설치 필요'
        return
    }
    if (-not $Report.Online) { Write-Warn '인터넷이 안 돼서 건너뜁니다.'; $Report.ClaudeCode = '건너뜀 (인터넷 없음)'; return }

    if ($Recommended) {
        Write-Info '이 PC는 Cowork을 쓸 수 없어서, 대신 Claude Code 설치를 권장합니다.'
        Write-Info 'Claude Code는 가상화 없이 터미널에서 파일 읽기/수정/정리, 명령 실행을 할 수 있습니다.'
    }
    $q = if ($Recommended) { 'Claude Code를 설치할까요?' } else { 'Claude Code(claude 명령)도 같이 설치할까요?' }
    if (-not $NoPrompt -and -not (Read-YesNo $q $true)) { $Report.ClaudeCode = '건너뜀'; $script:DeclinedCode = $true; return }

    $tmp = Join-Path $env:TEMP ('claude-code-install-{0}.ps1' -f [guid]::NewGuid().ToString('N'))
    try {
        Invoke-WebRequest -UseBasicParsing -Uri $Urls.CodeInstaller -OutFile $tmp
        if ((Get-Content -LiteralPath $tmp -Raw) -notmatch 'downloads\.claude\.ai') { throw '공식 설치 스크립트를 받지 못했습니다.' }
        Write-Info '공식 설치 스크립트 실행 중...'
        $quoted = $tmp -replace "'", "''"   # 사용자 이름에 ' 가 있어도 안전하게
        $cmd = "[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor 3072; & '$quoted'; exit `$LASTEXITCODE"
        Invoke-Native 'powershell.exe' @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', $cmd) -Show
    } catch {
        Write-Bad "Claude Code 설치 실패: $($_.Exception.Message)"
    } finally {
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    }

    if (Test-Path $ClaudeCodeExe) {
        Write-Ok 'Claude Code 설치 완료'
        Add-UserPath
        $Report.ClaudeCode = '설치 완료'
    } else {
        Write-Bad 'Claude Code가 설치되지 않았습니다. install-claude-code.bat 으로 다시 시도하세요.'
        $Report.ClaudeCode = '설치 실패'
    }
}

# =============================================================== 6. 마무리
function Invoke-ReadinessCheck {
    Write-Step '6. 공식 Cowork 점검 프로그램 (선택)'
    if ($Report.Virtualization -eq 'Off') { Write-Info 'Cowork을 쓸 수 없는 PC라서 건너뜁니다.'; return }
    $exe = Find-ReadinessCheck
    if (-not $exe) { Write-Info '점검 프로그램이 없어서 건너뜁니다 (없어도 됩니다).'; $Report.Readiness = '건너뜀'; return }
    if ($CheckOnly) { Write-Ok "점검 프로그램 있음: $(Split-Path $exe -Leaf)"; return }
    Write-Info "'This computer is ready for Cowork' 가 나오면 정상입니다. 점검 창을 닫으면 계속됩니다."
    try { Start-Process -FilePath $exe -Wait -NoNewWindow; $Report.Readiness = '실행함' }
    catch { Write-Warn "점검 프로그램 실행 실패: $($_.Exception.Message)" }
}

function Start-ClaudeApp {
    try {
        $app = Get-StartApps | Where-Object { $_.AppID -like "$PFN!*" } | Select-Object -First 1
        if ($app) {
            # explorer 를 거쳐 실행하면 관리자 권한이 아닌 일반 권한으로 열림
            Start-Process -FilePath explorer.exe -ArgumentList "shell:AppsFolder\$($app.AppID)"
            return $true
        }
    } catch {}
    return $false
}

function Test-RebootedSinceRegistration {
    # RunOnce 는 "재시작"이 아니라 "다음 로그인" 때 실행됨 -> 실제로 재부팅했는지 부팅 시각으로 확인
    # (빠른 시작이 켜진 상태의 [시스템 종료]는 부팅 시각이 안 바뀌므로 '재시작 안 함'으로 판단됨)
    if ($RegisteredAt -le 0) { return $true }
    try { return ((Get-CimInstance Win32_OperatingSystem).LastBootUpTime.ToUniversalTime().Ticks -gt $RegisteredAt) }
    catch { return $true }
}

function Register-Resume {
    if (-not $script:SameUser) { return $false }
    $cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Resumed -RegisteredAt $([DateTime]::UtcNow.Ticks)"
    if ($script:DeclinedCode -or $SkipClaudeCode) { $cmd += ' -SkipClaudeCode' }
    try {
        if (-not (Test-Path $RunOnceKey)) { New-Item -Path $RunOnceKey -Force | Out-Null }
        Set-ItemProperty -Path $RunOnceKey -Name $RunOnceName -Value $cmd
        return $true
    } catch { return $false }
}

function Get-CoworkStatus {
    if ($Report.Fatal)                            { return '불가 (위 [실패] 항목 확인)' }
    if ($script:VirtError)                        { return '불가 (가상화 설정 실패) -> Claude Code 사용' }
    if ($script:RestartFailed)                    { return '불가 (재시작 후에도 가상화 미적용) -> Claude Code 사용' }
    if ($Report.Virtualization -eq 'Off')         { return '불가 (가상화 꺼짐) -> Claude Code 사용' }
    if ($Report.Desktop -like '*실패*')           { return '불가 (Claude Desktop 설치 실패)' }
    if ($CheckOnly)                               { return '점검만 함 (설치는 setup.bat)' }
    if ($Report.NeedRestart)                      { return '재시작 후 사용 가능' }
    if ($Report.Virtualization -eq 'Unknown')     { return '사용 가능할 수 있음 (앱에서 확인)' }
    return '사용 가능'
}

function Show-Summary([bool]$AutoResume = $false) {
    Write-Step '결과'
    Write-Host ("  Cowork          : {0}" -f (Get-CoworkStatus))
    Write-Host ("  Claude Desktop  : {0}" -f $Report.Desktop)
    Write-Host ("  Claude Code     : {0}" -f $Report.ClaudeCode)
    if ($CheckOnly) { return }
    Write-Host ''
    Write-Host '  다음 할 일' -ForegroundColor Cyan
    if ($script:RestartFailed) {
        Write-Host '   1) README.md 의 [BIOS 가상화 켜기]로 BIOS 설정 확인'
        Write-Host '   2) 그동안은 Claude Code(터미널) 또는 Claude 앱 일반 채팅을 쓰세요'
    } elseif ($Report.NeedRestart) {
        Write-Host '   1) PC를 [다시 시작] (시스템 종료 X)'
        if ($AutoResume) { Write-Host '   2) 로그인하면 이 설치 창이 자동으로 다시 떠서 마무리합니다 ([예] 클릭)' }
        else             { Write-Host '   2) 다시 시작한 뒤 setup.bat 을 한 번 더 실행하세요' }
    } elseif ($Report.Desktop -notlike '*실패*' -and $Report.Desktop -ne '-') {
        Write-Host '   1) Claude 앱에서 로그인 (이메일 또는 Google)'
        Write-Host '   2) Cowork 탭 선택 -> 작업할 폴더를 고르고 할 일을 입력'
        if (-not $Report.DesktopOffline -and $Report.Desktop -like '설치 완료*') {
            Write-Host '      (Cowork 첫 실행 때 실행 환경을 내려받느라 몇 분 걸릴 수 있습니다)'
        }
    }
    if (@('설치됨', '설치 완료') -contains $Report.ClaudeCode) {
        Write-Host '   *) Claude Code: 새 터미널(PowerShell)을 열고 claude 입력 -> 로그인'
    }
}

# =============================================================== 메인
function Invoke-Main {
    if (-not $PSCommandPath) { Write-Host 'setup.bat 으로 실행해 주세요.'; return }

    # 32비트 PowerShell 에서 실행된 경우 64비트로 다시 실행
    if ($env:PROCESSOR_ARCHITEW6432 -and -not [Environment]::Is64BitProcess) {
        $ps64 = Join-Path $env:windir 'sysnative\WindowsPowerShell\v1.0\powershell.exe'
        $passArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath) + (Get-PassThroughArgs)
        if ($CheckOnly)      { $passArgs += '-CheckOnly' }
        if ($ClaudeCodeOnly) { $passArgs += '-ClaudeCodeOnly' }
        & $ps64 @passArgs
        exit $LASTEXITCODE
    }

    Initialize-Network
    $script:IsAdmin = Test-IsAdmin

    # Claude Code만 설치 (관리자 권한 불필요)
    if ($ClaudeCodeOnly) {
        Start-Log
        $Report.Online = $true
        Install-ClaudeCode $false $true
        Exit-Setup ($(if ($Report.ClaudeCode -like '*실패*') { 2 } else { 0 }))
    }

    if (-not $script:IsAdmin -and -not $CheckOnly) {
        if ($Elevated) {
            # 권한 요청을 했는데도 관리자 권한이 없음 (UAC 꺼진 일반 계정 등) -> 무한 반복 방지
            Write-Host '  관리자 권한을 받지 못했습니다. 관리자 계정으로 Windows에 로그인한 뒤 다시 실행하세요.' -ForegroundColor Red
            [void](Read-Host '  Enter 키를 누르면 창이 닫힙니다')
            exit 2
        }
        if ((Get-KitDriveType) -eq 'Network') {
            Write-Host '  네트워크 드라이브에서는 관리자 권한으로 실행이 안 됩니다.' -ForegroundColor Red
            Write-Host '  이 폴더를 바탕화면 등 내 PC에 복사한 뒤 다시 실행하세요.' -ForegroundColor Red
            [void](Read-Host '  Enter 키를 누르면 창이 닫힙니다')
            exit 2
        }
        Request-Elevation
    }

    Start-Log
    $script:SameUser = Test-SameUser
    Write-Host ''
    Write-Host '  ============================================' -ForegroundColor Cyan
    Write-Host '    Claude Cowork 설치 도우미 (Windows)' -ForegroundColor Cyan
    if ($CheckOnly) { Write-Host '    점검 모드: 아무것도 바꾸지 않습니다' -ForegroundColor Cyan }
    if ($Resumed)   { Write-Host '    재시작 후 마무리 단계' -ForegroundColor Cyan }
    Write-Host '  ============================================' -ForegroundColor Cyan
    if (-not $script:SameUser) {
        Write-Warn '지금 로그인한 계정과 다른 관리자 계정으로 실행 중입니다. 일부 단계(Claude Code, 자동 재개)는 건너뜁니다.'
    }

    Test-System
    Test-Network
    if ($Report.Fatal) { Show-Summary; Exit-Setup 2 }

    # 단계별로 오류를 가둬서, 한 단계가 실패해도 나머지(특히 Claude Code 대안과 재시작 안내)는 계속 진행
    try { Set-Virtualization }
    catch {
        Write-Bad "가상화 설정 실패: $($_.Exception.Message)"
        $script:VirtError = $_.Exception.Message
        $Report.Virtualization = 'Off'
    }
    if (-not $CheckOnly) {
        try { Remove-LegacyClaude } catch { Write-Warn "예전 버전 확인 중 오류 (무시하고 계속): $($_.Exception.Message)" }
    }
    try { Install-ClaudeDesktop }
    catch {
        Write-Bad "Claude Desktop 설치 실패: $($_.Exception.Message)"
        Show-InstallHint $_.Exception.Message
        $Report.Desktop = '설치 실패'
    }
    try { Install-ClaudeCode ($Report.Virtualization -eq 'Off' -or $Report.Desktop -like '*실패*') }
    catch { Write-Bad "Claude Code 설치 실패: $($_.Exception.Message)"; $Report.ClaudeCode = '설치 실패' }

    $bad = $Report.Desktop -like '*실패*' -or $Report.ClaudeCode -like '*실패*' -or [bool]$script:VirtError
    $exitCode = $(if ($bad) { 2 } else { 0 })

    if ($Report.NeedRestart -and -not $CheckOnly) {
        if ($Resumed -and (Test-RebootedSinceRegistration)) {
            Write-Warn '재시작했는데도 가상화 설정이 아직 적용되지 않았습니다.'
            Write-Info 'BIOS 가상화가 꺼져 있거나 이 PC(예: Home 에디션)에서 지원되지 않는 경우입니다.'
            $script:RestartFailed = $true
            Show-Summary $false
            Exit-Setup 2
        }
        if ($Resumed) { Write-Info '아직 PC를 [다시 시작]하지 않은 것 같습니다 (로그아웃/로그인 또는 빠른 시작 종료만 함).' }
        $auto = Register-Resume
        Show-Summary $auto
        if ((Get-KitDriveType) -eq 'Removable') { Write-Warn 'USB에서 실행 중입니다. 재시작할 때 USB를 뽑지 마세요.' }
        Write-Host ''
        if (Read-YesNo '지금 다시 시작할까요? (열려 있는 작업은 먼저 저장하세요)' $false) {
            Stop-Log
            Restart-Computer -Force
            exit $exitCode
        }
        Exit-Setup $exitCode
    }

    try { Invoke-ReadinessCheck } catch { Write-Warn "점검 프로그램 오류 (무시하고 계속): $($_.Exception.Message)" }
    if (-not $CheckOnly -and $Report.Desktop -notlike '*실패*' -and $Report.Desktop -ne '-') {
        if (Start-ClaudeApp) { Write-Host ''; Write-Ok 'Claude 앱을 열었습니다.' }
        else { Write-Host ''; Write-Info '시작 메뉴에서 Claude를 직접 실행하세요.' }
    }
    Show-Summary
    Exit-Setup $exitCode
}

# 테스트용으로 dot-source(. .\setup.ps1) 하면 함수만 불러오고 실행하지 않음
if ($MyInvocation.InvocationName -ne '.') {
    try { Invoke-Main }
    catch {
        Write-Host ''
        Write-Host "  [오류] 예상하지 못한 문제가 생겼습니다: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "         (setup.ps1 $($_.InvocationInfo.ScriptLineNumber)번째 줄)" -ForegroundColor DarkGray
        Write-Host '         README.md 의 [문제 해결]을 보거나, 아래 로그 파일을 Claude에게 보여주세요.'
        Exit-Setup 2
    }
}
