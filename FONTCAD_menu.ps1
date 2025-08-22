<# FONTCAD_menu.ps1
 - 1) Install ALL  2) Install AUTOCAD (SHX)
 - Tu nap script cai chinh tu RAW GitHub
#>

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Config
$RawInstallUrl = "https://raw.githubusercontent.com/SurveyRoyal/Survey/main/Install_Fonts_FromZip.ps1"
$ZipUrl        = "https://github.com/SurveyRoyal/Survey/releases/download/CAIDATFONT/FONTCAD.zip"
$DestShx       = "C:\FONTCAD\SHX"

function Ensure-InstallFn {
  if (-not (Get-Command Install-Fonts_FromZip -ErrorAction SilentlyContinue)) {
    Write-Host "`n-> Loading installer script from GitHub..." -ForegroundColor Yellow
    irm $RawInstallUrl | iex
  }
}

function Add-SupportPath($path) {
  try {
    Write-Host "-> Add $path to AutoCAD Support Path..." -ForegroundColor Yellow
    $keys = Get-ChildItem "HKCU:\Software\Autodesk\AutoCAD" -Recurse -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match "\\Profiles\\[^\\]+\\General$" }
    $t=0;$m=0
    foreach ($k in $keys) {
      $cur = (Get-ItemProperty -Path $k.PSPath -Name "ACAD" -ErrorAction SilentlyContinue).ACAD
      if ($cur) {
        if ($cur -notmatch [regex]::Escape($path)) {
          Set-ItemProperty -Path $k.PSPath -Name "ACAD" -Value "$path;$cur"
          $t++
        } else { $m++ }
      }
    }
    Write-Host "   Updated: $t, Exists: $m" -ForegroundColor Green
  } catch { Write-Host "   !! Cannot update Support Path: $($_.Exception.Message)" -ForegroundColor Red }
}

function Install-ALL {
  Ensure-InstallFn
  Write-Host "`n=== Install ALL (SHX + Windows fonts + CTB) ===" -ForegroundColor Cyan
  Write-Host "ZIP: $ZipUrl" -ForegroundColor DarkGray
  Install-Fonts_FromZip -Zip $ZipUrl -DoShx -DoTtf -DoPlot -OnlyNew -DestShx $DestShx
  Add-SupportPath $DestShx
  Write-Host ">>> DONE. Log: $DestShx\font_install.log`n" -ForegroundColor Green
}

function Install-AUTOCAD {
  Ensure-InstallFn
  Write-Host "`n=== Install AUTOCAD (only SHX) ===" -ForegroundColor Cyan
  Write-Host "ZIP: $ZipUrl" -ForegroundColor DarkGray
  Install-Fonts_FromZip -Zip $ZipUrl -DoShx -OnlyNew -DestShx $DestShx
  Add-SupportPath $DestShx
  Write-Host ">>> DONE. Log: $DestShx\font_install.log`n" -ForegroundColor Green
}

function Show-Menu {
  Clear-Host
  Write-Host "==============================="
  Write-Host "        MENU CAI DAT FONTCAD   "
  Write-Host "==============================="
  Write-Host "1. Install ALL (SHX + Windows fonts + CTB)"
  Write-Host "2. Install AUTOCAD (chi SHX)"
  Write-Host "3. Mo file log"
  Write-Host "4. Thoat"
  Write-Host "==============================="
}

do {
  Show-Menu
  $c = Read-Host "Nhap lua chon (1-4)"
  switch ($c) {
    "1" { Install-ALL;    Pause }
    "2" { Install-AUTOCAD;Pause }
    "3" { ii "$DestShx\font_install.log" -ErrorAction SilentlyContinue; Pause }
    "4" { break }
    Default { Write-Host "Lua chon khong hop le!" -ForegroundColor Red; Start-Sleep 1 }
  }
} while ($true)
