<# FONTCAD_menu.ps1
 - Menu 1) Install ALL  2) Install AUTOCAD (chỉ SHX)
 - Mặc định SHX -> C:\FONTCAD\SHX
 - ZIP fonts của bạn: Survey/CAIDATFONT/FONTCAD.zip
 - Cần chạy PowerShell "Run as Administrator"
#>

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ---- cấu hình ----
$RawInstallUrl = "https://raw.githubusercontent.com/SurveyRoyal/Survey/main/Install_Fonts_FromZip.ps1"
$ZipUrl        = "https://github.com/SurveyRoyal/Survey/releases/download/CAIDATFONT/FONTCAD.zip"
$DestShx       = "C:\FONTCAD\SHX"

function Ensure-InstallFn {
  if (-not (Get-Command Install-Fonts_FromZip -ErrorAction SilentlyContinue)) {
    Write-Host "-> Dang nap script cai dat tu GitHub..." -f Yellow
    irm $RawInstallUrl | iex
  }
}

function Add-SupportPath($path) {
  try {
    Write-Host "-> Them $path vao AutoCAD Support Path (cac profile HKCU)..." -f Yellow
    $keys = Get-ChildItem "HKCU:\Software\Autodesk\AutoCAD" -Recurse -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match "\\Profiles\\[^\\]+\\General$" }
    $t=0;$m=0
    foreach ($k in $keys) {
      $cur = (Get-ItemProperty -Path $k.PSPath -Name "ACAD" -ErrorAction SilentlyContinue).ACAD
      if ($cur) {
        if ($cur -notmatch [regex]::Escape($path)) {
          $new = "$path;$cur"
          Set-ItemProperty -Path $k.PSPath -Name "ACAD" -Value $new
          $t++
        } else { $m++ }
      }
    }
    Write-Host "   -> Updated: $t, Da co san: $m" -f Green
  } catch {
    Write-Host "   !! Khong the cap nhat Support Path: $($_.Exception.Message)" -f Red
  }
}

function Install-ALL {
  Ensure-InstallFn
  Write-Host "`n=== Cai ALL (SHX + Windows fonts + CTB) ===" -f Cyan
  Install-Fonts_FromZip -ZipUrl $ZipUrl -DoShx -DoTtf -DoPlot -OnlyNew -DestShx $DestShx
  Add-SupportPath $DestShx
  Write-Host ">>> Hoan tat. Log: $DestShx\font_install.log`n" -f Green
}

function Install-AUTOCAD {
  Ensure-InstallFn
  Write-Host "`n=== Cai chi cho AutoCAD (SHX) ===" -f Cyan
  Install-Fonts_FromZip -ZipUrl $ZipUrl -DoShx -OnlyNew -DestShx $DestShx
  Add-SupportPath $DestShx
  Write-Host ">>> Hoan tat. Log: $DestShx\font_install.log`n" -f Green
}

function Show-Menu {
  Clear-Host
  Write-Host "==============================="
  Write-Host "     MENU CAI DAT FONTCAD     "
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
    Default { Write-Host "Lua chon khong hop le!" -f Red; Start-Sleep 1 }
  }
} while ($true)
