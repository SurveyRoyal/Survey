<# Install_Fonts_FromZip.ps1 (robust)
   - Param -Zip: http/https URL hoac duong dan .zip local
   - SHX (.shx/.shp/.pfb/.pfa)  -> C:\FONTCAD\SHX (flatten)
   - TTF/OTF/TTC                -> %WINDIR%\Fonts
   - CTB/STB                    -> Plot Styles (auto find)
   - Only copy new/newer khi -OnlyNew
   - Log: C:\FONTCAD\SHX\font_install.log
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)][Alias('ZipUrl','ZipPath')] [string]$Zip,
  [string]$DestShx = "C:\FONTCAD\SHX",
  [string]$DestTtf = "$env:WINDIR\Fonts",
  [string]$DestPlot,
  [switch]$DoShx, [switch]$DoTtf, [switch]$DoPlot, [switch]$OnlyNew
)

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# --- input sanitize & classify ---
$Zip = ($Zip -replace '[\u200B-\u200D\uFEFF]', '').Trim()
if ([string]::IsNullOrWhiteSpace($Zip)) {
  throw "Zip/Url is empty. Pass -Zip <URL or local .zip>"
}
$IsUrl = $false
try {
  $u = [Uri]$Zip
  $IsUrl = $u.IsAbsoluteUri -and ($u.Scheme -in @('http','https'))
} catch { $IsUrl = $false }

# ---------- helpers ----------
function Write-Log([string]$msg) {
  $logDir  = $DestShx
  $logFile = Join-Path $logDir "font_install.log"
  if (!(Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
  $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
  Add-Content -Path $logFile -Value "$ts  $msg"
  Write-Host "  $msg"
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

function Download-File($url,$dst) {
  $url = ($url -replace '[\u200B-\u200D\uFEFF]', '').Trim()
  if (-not [Uri]::IsWellFormedUriString($url,[UriKind]::Absolute)) {
    throw "Invalid URL: '$url'"
  }
  $headers = @{ 'User-Agent'='Mozilla/5.0'; 'Accept'='*/*' }
  try {
    Write-Host "-> Downloading (Invoke-WebRequest)..."
    Invoke-WebRequest -Uri $url -OutFile $dst -UseBasicParsing -Headers $headers -MaximumRedirection 10 -ErrorAction Stop
  } catch {
    Write-Host "-> IWR failed, trying BITS..."
    Start-BitsTransfer -Source $url -Destination $dst -ErrorAction Stop
  }
}

# ---------- main ----------
Write-Log "=== Start Install from: $Zip ==="

$tmpRoot = Join-Path $env:TEMP ("FONTCAD_" + [guid]::NewGuid())
New-Item -ItemType Directory -Path $tmpRoot | Out-Null
$zipPath = $null

try {
  if ($IsUrl) {
    $zipPath = Join-Path $tmpRoot "FONTCAD.zip"
    Write-Log "Downloading ZIP to: $zipPath"
    Download-File -url $Zip -dst $zipPath
    $sizeMB = [Math]::Round((Get-Item $zipPath).Length / 1MB,2)
    Write-Log "Downloaded: ${sizeMB} MB"
    if ((Get-Item $zipPath).Length -lt 1024) { throw "ZIP too small/invalid." }
  }
  elseif (Test-Path -LiteralPath $Zip) {
    $zipPath = (Resolve-Path -LiteralPath $Zip).Path
    Write-Log "Using local ZIP: $zipPath"
  }
  else {
    throw "Not a valid URL or existing local path: '$Zip'"
  }

  $extract = Join-Path $tmpRoot "unzipped"
  Write-Log "Expanding ZIP..."
  Expand-Archive -Path $zipPath -DestinationPath $extract -Force

  if ($DoShx) {
    if (!(Test-Path $DestShx)) { New-Item -ItemType Directory -Path $DestShx | Out-Null }
    $copied=0; $skipped=0
    foreach ($pat in @("*.shx","*.shp","*.pfb","*.pfa")) {
      Get-ChildItem -Path $extract -Recurse -File -Include $pat -ErrorAction SilentlyContinue | ForEach-Object {
        $dst = Join-Path $DestShx $_.Name
        if ($OnlyNew) {
          if (Should-Copy $_.FullName $dst) { Copy-Item $_.FullName -Destination $dst -Force; $copied++ } else { $skipped++ }
        } else { Copy-Item $_.FullName -Destination $dst -Force; $copied++ }
      }
    }
    Write-Log "SHX -> $DestShx | Copied=$copied Skipped=$skipped"
  }

  if ($DoTtf) {
    if (!(Test-Path $DestTtf)) { New-Item -ItemType Directory -Path $DestTtf | Out-Null }
    $copied=0; $skipped=0
    foreach ($pat in @("*.ttf","*.otf","*.ttc")) {
      Get-ChildItem -Path $extract -Recurse -File -Include $pat -ErrorAction SilentlyContinue | ForEach-Object {
        $dst = Join-Path $DestTtf $_.Name
        if ($OnlyNew) {
          if (Should-Copy $_.FullName $dst) { Copy-Item $_.FullName -Destination $dst -Force; $copied++ } else { $skipped++ }
        } else { Copy-Item $_.FullName -Destination $dst -Force; $copied++ }
      }
    }
    Write-Log "TTF/OTF -> $DestTtf | Copied=$copied Skipped=$skipped"
  }

  if ($DoPlot) {
    $plotDir = Find-PlotStylesDir
    $copied=0; $skipped=0
    foreach ($pat in @("*.ctb","*.stb")) {
      Get-ChildItem -Path $extract -Recurse -File -Include $pat -ErrorAction SilentlyContinue | ForEach-Object {
        $dst = Join-Path $plotDir $_.Name
        if ($OnlyNew) {
          if (Should-Copy $_.FullName $dst) { Copy-Item $_.FullName -Destination $dst -Force; $copied++ } else { $skipped++ }
        } else { Copy-Item $_.FullName -Destination $dst -Force; $copied++ }
      }
    }
    Write-Log "Plot Styles -> $plotDir | Copied=$copied Skipped=$skipped"
  }

  Write-Log "=== Done ==="
  Write-Output "DONE"
}
catch {
  Write-Log "ERROR: $($_.Exception.Message)"
  throw
}
finally {
  Remove-Item $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue
}

