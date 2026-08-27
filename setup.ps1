<#
  setup.ps1 - Cai "Tro ly trinh van ban" tren may Windows. Chay 1 lan la xong.
  - Khong can quyen admin.
  - Chay lai lan nua = cap nhat len ban moi nhat.

  Cach nhanh nhat (mo PowerShell, dan 1 dong roi Enter):
    iex (iwr -useb 'https://raw.githubusercontent.com/hungvumoh/trinh-van-ban-voffice/main/setup.ps1').Content

  Hoac double-click file install.bat.
#>

$ErrorActionPreference = 'Stop'
$ProgressPreference     = 'SilentlyContinue'   # tai nhanh hon nhieu

$Repo       = 'hungvumoh/trinh-van-ban-voffice'
$Branch     = 'main'
$InstallDir = Join-Path $env:LOCALAPPDATA 'TroLyTrinhVanBan'
$PyMin      = [Version]'3.9'
$PyDownload = 'https://www.python.org/ftp/python/3.12.10/python-3.12.10-amd64.exe'
$PyUserExe  = "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe"

function Say($m) { Write-Host $m -ForegroundColor Cyan }
function Ok ($m) { Write-Host $m -ForegroundColor Green }

try {
    # ---------- 1. Python ----------
    Say "`n[1/4] Kiem tra Python..."
    $py = $null
    foreach ($c in @($PyUserExe,
                     "$env:LOCALAPPDATA\Programs\Python\Python311\python.exe",
                     "$env:LOCALAPPDATA\Programs\Python\Python310\python.exe")) {
        if (Test-Path $c) { $py = $c; break }
    }
    if (-not $py) {
        foreach ($name in @('py', 'python')) {
            $g = Get-Command $name -ErrorAction SilentlyContinue
            if ($g -and $g.Source -notlike '*WindowsApps*') {
                try {
                    $v = & $g.Source -c "import sys;print('%d.%d'%sys.version_info[:2])" 2>$null
                    if ($v -and [Version]$v -ge $PyMin) { $py = $g.Source; break }
                } catch { }
            }
        }
    }
    if (-not $py) {
        Say "  Chua co Python phu hop - tai bo cai chinh thuc tu python.org..."
        $tmp = Join-Path $env:TEMP 'python-setup.exe'
        Invoke-WebRequest $PyDownload -OutFile $tmp -UseBasicParsing
        Say "  Dang cai (am tham, khong can admin)..."
        Start-Process $tmp -Wait -ArgumentList @(
            '/quiet', 'InstallAllUsers=0', 'PrependPath=1', 'Include_tcltk=1',
            'Include_pip=1', 'Include_launcher=1', 'Include_test=0')
        Remove-Item $tmp -ErrorAction SilentlyContinue
        $py = $PyUserExe
        if (-not (Test-Path $py)) {
            throw "Cai Python xong nhung khong thay $py. Cai tay python.org 3.12 (tick Add to PATH) roi chay lai."
        }
    }
    Ok "  Python: $py"

    # ---------- 2. Tai ma nguon ----------
    Say "`n[2/4] Tai chuong trinh -> $InstallDir"
    $zip   = Join-Path $env:TEMP 'tvb.zip'
    $stage = Join-Path $env:TEMP 'tvb-stage'
    Invoke-WebRequest "https://github.com/$Repo/archive/refs/heads/$Branch.zip" -OutFile $zip -UseBasicParsing
    if (Test-Path $stage) { Remove-Item $stage -Recurse -Force }
    Expand-Archive $zip $stage -Force
    $src = (Get-ChildItem $stage -Directory | Select-Object -First 1).FullName

    New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
    # giu lai noi_nhan.json nguoi dung da sua tay (neu co) -> .bak truoc khi ghi de
    $nn = Join-Path $InstallDir 'noi_nhan.json'
    if (Test-Path $nn) { Copy-Item $nn "$nn.bak" -Force }
    Copy-Item (Join-Path $src '*') $InstallDir -Recurse -Force
    Remove-Item $zip -ErrorAction SilentlyContinue
    Remove-Item $stage -Recurse -Force -ErrorAction SilentlyContinue
    Ok "  Da tai xong."

    # ---------- 3. Thu vien ----------
    Say "`n[3/4] Cai thu vien Python (co the mat vai phut)..."
    & $py -m pip install --upgrade pip --quiet --disable-pip-version-check
    & $py -m pip install --upgrade --quiet --disable-pip-version-check -r (Join-Path $InstallDir 'requirements.txt')
    if ($LASTEXITCODE -ne 0) { throw "pip install that bai (ma loi $LASTEXITCODE). Kiem tra ket noi mang / proxy." }
    Ok "  Xong thu vien."

    # ---------- 4. Loi tat ----------
    Say "`n[4/4] Tao loi tat..."
    $pyDir = Split-Path $py
    $pyw = Join-Path $pyDir 'pythonw.exe'
    if (-not (Test-Path $pyw)) { $pyw = Join-Path $pyDir 'pyw.exe' }
    if (-not (Test-Path $pyw)) { $pyw = $py }
    $target = Join-Path $InstallDir 'trinh_van_ban.py'
    $sh = New-Object -ComObject WScript.Shell
    foreach ($dir in @([Environment]::GetFolderPath('Desktop'),
                       (Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs'))) {
        $lnk = $sh.CreateShortcut((Join-Path $dir 'Tro ly trinh van ban.lnk'))
        $lnk.TargetPath       = $pyw
        $lnk.Arguments        = "`"$target`""
        $lnk.WorkingDirectory = $InstallDir
        $lnk.IconLocation     = "$pyw,0"
        $lnk.Save()
    }
    Ok "  Da tao loi tat (Desktop + Start Menu)."

    Ok "`n=== HOAN TAT ==="
    Write-Host "Mo 'Tro ly trinh van ban' tu Desktop."
    Write-Host "Lan dau: dang nhap + tick 'Nho dang nhap'."
    Write-Host "Muon tu chay khi bat may: trong chuong trinh tick 'Khoi dong cung Windows'."
    Write-Host "Cap nhat sau nay: chay lai lenh cai dat nay."

    Start-Process $pyw -ArgumentList "`"$target`"" -WorkingDirectory $InstallDir
}
catch {
    Write-Host "`nLOI: $_" -ForegroundColor Red
    Write-Host "Chup man hinh nay gui nguoi cap phan mem." -ForegroundColor Yellow
    exit 1
}
