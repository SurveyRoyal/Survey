<# Install_Fonts_FromZip.ps1
   - Tải 1 file ZIP (bên trong có thể có: SHX, Font window, CTB, ...)
   - SHX (*.shx, *.shp, *.pfb, *.pfa)  -> COPY PHẲNG vào C:\FONTCAD\SHX (hoặc -DestShx)
   - Windows fonts (*.ttf, *.otf, *.ttc) -> C:\Windows\Fonts (hoặc -DestTtf)
   - Plot styles (*.ctb, *.stb) -> thư mục “Plot Styles” (auto dò, hoặc -DestPlot)
   - Chỉ copy file mới/mới hơn khi dùng -OnlyNew (khuyên dùng)
   - Ghi log tại C:\FONTCAD\SHX\font_install.log
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)] [string]$ZipUrl,
  [string]$DestShx = "C:\FONTCAD\SHX",
  [string]$DestTtf = "$env:WINDIR\Fonts",
  [string]$DestPlot,
  [switch]$DoShx,
  [switch]$DoTtf,
  [switch]$DoPlot,
  [switch]$OnlyNew
)

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ---------- helpers ----------
function Write-Log([string]$msg) {
  $logDir  = $DestShx
  $logFile = Join-Path $logDir "font_install.log"
  if (!(Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
  $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
  Add-Content -Path $logFile -Value "$ts  $msg"
}

function Download-File($url,$dst) {
  try { Start-BitsTransfer -Source $url -Destination $dst -ErrorAction Stop }
  catch { Invoke-WebRequest -Uri $url -OutFile $dst -UseBasicParsing -ErrorAction Stop }
}

function Should-Copy($src,$dst) {
  if (-not (Test-Path $dst)) { return $true }
  $s = Get-Item $src; $d = Get-Item $dst
  return (($s.LastWriteTimeUtc -gt $d.LastWriteTimeUtc) -or ($s.Length -ne $d.Length))
}

function Find-PlotStylesDir {
  if ($DestPlot -and (Test-Path $DestPlot)) { return $DestPlot }
  $root = Join-Path $env:APPDATA "Autodesk"
  if (Test-Path $root) {
    $hit = Get-ChildItem -Path $root -Directory -Recurse -ErrorAction SilentlyContinue |
           Where-Object { $_.Name -ieq "Plot Styles" } |
           Select-Object -First 1
    if ($hit) { return $hit.FullName }
  }
  $fallback = Join-Path $env:USERPROFILE "Documents\AutoCAD Plot Styles"
  if (!(Test-Path $fallback)) { New-Item -ItemType Directory -Path $fallback | Out-Null }
  return $fallback
}

# ---------- main ----------
Write-Log "=== Start from ZIP: $ZipUrl ==="
$tmp = Join-Path $env:TEMP ("FONTCAD_" + [guid]::NewGuid())
$zip = "$tmp.zip"
New-Item -ItemType Directory -Path $tmp | Out-Null

try {
  Write-Log "Downloading ZIP..."
  Download-File $ZipUrl $zip
  if ((Get-Item $zip).Length -lt 1024) { throw "ZIP too small/invalid." }

  Write-Log "Expanding ZIP..."
  Expand-Archive -Path $zip -DestinationPath $tmp -Force

  if ($DoShx) {
    if (!(Test-Path $DestShx)) { New-Item -ItemType Directory -Path $DestShx | Out-Null }
    $shxExt = @("*.shx","*.shp","*.pfb","*.pfa")
    $copied=0; $skipped=0
    foreach ($pat in $shxExt) {
      Get-ChildItem -Path $tmp -Recurse -File -Include $pat -ErrorAction SilentlyContinue | ForEach-Object {
        $dst = Join-Path $DestShx $_.Name         # FLATTEN (Support Path không đệ quy)
        if ($OnlyNew) {
          if (Should-Copy $_.FullName $dst) { Copy-Item $_.FullName -Destination $dst -Force; $copied++ } else { $skipped++ }
        } else {
          Copy-Item $_.FullName -Destination $dst -Force; $copied++
        }
      }
    }
    Write-Log "SHX: Copied=$copied Skipped=$skipped Dest=$DestShx"
  }

  if ($DoTtf) {
    if (!(Test-Path $DestTtf)) { New-Item -ItemType Directory -Path $DestTtf | Out-Null }
    $ttfExt = @("*.ttf","*.otf","*.ttc")
    $copied=0; $skipped=0
    foreach ($pat in $ttfExt) {
      Get-ChildItem -Path $tmp -Recurse -File -Include $pat -ErrorAction SilentlyContinue | ForEach-Object {
        $dst = Join-Path $DestTtf $_.Name
        if ($OnlyNew) {
          if (Should-Copy $_.FullName $dst) { Copy-Item $_.FullName -Destination $dst -Force; $copied++ } else { $skipped++ }
        } else {
          Copy-Item $_.FullName -Destination $dst -Force; $copied++
        }
      }
    }
    Write-Log "TTF: Copied=$copied Skipped=$skipped Dest=$DestTtf"
  }

  if ($DoPlot) {
    $plotDir = Find-PlotStylesDir
    $copied=0; $skipped=0
    foreach ($pat in @("*.ctb","*.stb")) {
      Get-ChildItem -Path $tmp -Recurse -File -Include $pat -ErrorAction SilentlyContinue | ForEach-Object {
        $dst = Join-Path $plotDir $_.Name
        if ($OnlyNew) {
          if (Should-Copy $_.FullName $dst) { Copy-Item $_.FullName -Destination $dst -Force; $copied++ } else { $skipped++ }
        } else {
          Copy-Item $_.FullName -Destination $dst -Force; $copied++
        }
      }
    }
    Write-Log "PLOT: Copied=$copied Skipped=$skipped Dest=$plotDir"
  }

  Write-Log "=== Done ==="
  Write-Output "DONE"
}
catch {
  Write-Log "ERROR: $($_.Exception.Message)"; throw
}
finally {
  Remove-Item $zip -Force -ErrorAction SilentlyContinue
  Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
