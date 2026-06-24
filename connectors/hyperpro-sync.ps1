#requires -version 5
<#
  Siir x HYPER PRO auto-sync
  --------------------------
  Reads HYPER PRO's Firebird PRODUIT table and pushes to the local Siir server:
    * the catalogue at the NORMAL price (PV1_TTC), and
    * the CURRENTLY-ACTIVE promotions (PROMO=1 AND today within D1..D2 AND a real
      discount PP1<PV1) as old/new price pairs.
  The server replaces the prior synced/imported promotions each run, so the promo
  page always mirrors the POS (no stale promos pile up). Read-only on HYPER PRO.

  Deploy on a store PC (PowerShell):
    mkdir C:\ProgramData\SiirSync 2>$null
    curl.exe -L -o C:\ProgramData\SiirSync\hyperpro-sync.ps1 `
      https://raw.githubusercontent.com/nudiiir/siir-releases/main/connectors/hyperpro-sync.ps1
    powershell -ExecutionPolicy Bypass -File C:\ProgramData\SiirSync\hyperpro-sync.ps1 -ApiKey sk_xxx
    schtasks /create /tn "SiirSync" /sc minute /mo 15 /ru SYSTEM /f ^
      /tr "powershell -NoProfile -ExecutionPolicy Bypass -File C:\ProgramData\SiirSync\hyperpro-sync.ps1 -ApiKey sk_xxx"
#>

param(
  [Parameter(Mandatory = $true)] [string] $ApiKey,
  [string] $DbPath  = 'C:\HYPER PRO\raya.FDB',
  [string] $SiirUrl = 'http://localhost:3000/api/stock',
  [string] $FbUser  = 'SYSDBA',
  [string] $FbPass  = 'masterkey',
  [string] $Charset = 'WIN1252'   # if names look garbled, try UTF8 or NONE
)

$ErrorActionPreference = 'Stop'
$work = Join-Path $env:ProgramData 'SiirSync'
New-Item -ItemType Directory -Force $work | Out-Null
$sqlFile  = Join-Path $work 'query.sql'
$itemsTxt = Join-Path $work 'items.txt'
$promoTxt = Join-Path $work 'promos.txt'
$payload  = Join-Path $work 'payload.json'
$log      = Join-Path $work 'sync.log'
function Log($m) { "[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $m | Add-Content $log }
function J($s) { if (-not $s) { return '""' }; '"' + $s.Replace('\', '\\').Replace('"', '\"').Replace("`r", ' ').Replace("`n", ' ').Replace("`t", ' ') + '"' }

try {
  $isql = @(
    (Get-ChildItem 'C:\Program Files\Firebird', 'C:\Program Files (x86)\Firebird' -Recurse -Filter isql.exe -ErrorAction SilentlyContinue |
       Select-Object -First 1 -ExpandProperty FullName),
    'C:\HYPER PRO\isql.exe'
  ) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
  if (-not $isql) { throw 'isql.exe not found (part of Firebird). Set its full path in the script.' }

  foreach ($f in @($itemsTxt, $promoTxt)) { if (Test-Path $f) { Remove-Item $f -Force } }

  # US = ASCII unit separator (31): a delimiter that never appears in product text.
  @"
SET HEADING OFF;
SET LIST OFF;
OUTPUT '$itemsTxt';
SELECT CODE_BARRE || ASCII_CHAR(31) || COALESCE(PRODUIT,'') || ASCII_CHAR(31) ||
  CAST(CAST(COALESCE(PV1_TTC,0) AS NUMERIC(18,2)) AS VARCHAR(20))
FROM PRODUIT WHERE CODE_BARRE IS NOT NULL AND CODE_BARRE <> '';
OUTPUT;
OUTPUT '$promoTxt';
SELECT CODE_BARRE || ASCII_CHAR(31) || COALESCE(PRODUIT,'') || ASCII_CHAR(31) ||
  CAST(CAST(PV1_TTC AS NUMERIC(18,2)) AS VARCHAR(20)) || ASCII_CHAR(31) ||
  CAST(CAST(PP1_TTC AS NUMERIC(18,2)) AS VARCHAR(20))
FROM PRODUIT
WHERE COALESCE(PROMO,0)=1 AND COALESCE(PP1_TTC,0)>0 AND COALESCE(PV1_TTC,0)>0
  AND PP1_TTC < PV1_TTC
  AND (D1 IS NULL OR D1 <= CURRENT_DATE) AND (D2 IS NULL OR D2 >= CURRENT_DATE)
  AND CODE_BARRE IS NOT NULL AND CODE_BARRE <> '';
OUTPUT;
"@ | Set-Content -Encoding ASCII $sqlFile

  & $isql -q -u $FbUser -p $FbPass -ch $Charset -i $sqlFile "localhost:$DbPath" > $null 2>&1
  if (-not (Test-Path $itemsTxt)) { throw 'Query produced no output - check DB path / Firebird user+password.' }

  $US = [char]31
  $itemLines  = Get-Content -Encoding Default $itemsTxt | Where-Object { $_.Trim() -ne '' }
  $promoLines = if (Test-Path $promoTxt) { Get-Content -Encoding Default $promoTxt | Where-Object { $_.Trim() -ne '' } } else { @() }

  $sb = New-Object System.Text.StringBuilder
  [void]$sb.Append('{"items":[')
  $first = $true
  foreach ($line in $itemLines) {
    $f = $line.Split($US); if ($f.Count -lt 3) { continue }
    if (-not $first) { [void]$sb.Append(',') }; $first = $false
    [void]$sb.Append('{"barcode":').Append((J $f[0].Trim())).Append(',"name":').Append((J $f[1].Trim())).Append(',"price":').Append($f[2].Trim()).Append('}')
  }
  [void]$sb.Append('],"promos":[')
  $first = $true
  foreach ($line in $promoLines) {
    $f = $line.Split($US); if ($f.Count -lt 4) { continue }
    if (-not $first) { [void]$sb.Append(',') }; $first = $false
    [void]$sb.Append('{"barcode":').Append((J $f[0].Trim())).Append(',"title":').Append((J $f[1].Trim())).Append(',"oldPrice":').Append($f[2].Trim()).Append(',"newPrice":').Append($f[3].Trim()).Append('}')
  }
  [void]$sb.Append(']}')
  [IO.File]::WriteAllText($payload, $sb.ToString(), (New-Object System.Text.UTF8Encoding($false)))

  $resp = & curl.exe -s -X POST $SiirUrl -H "x-api-key: $ApiKey" -H "Content-Type: application/json; charset=utf-8" --data-binary "@$payload"
  Log "synced $($itemLines.Count) products, $($promoLines.Count) active promos -> $resp"
  Write-Output "OK: $($itemLines.Count) products, $($promoLines.Count) active promos -> $resp"
}
catch {
  Log "ERROR: $($_.Exception.Message)"
  Write-Output "ERROR: $($_.Exception.Message)"
  exit 1
}
