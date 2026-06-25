#requires -version 5
<#
  Siir x HYPER PRO auto-sync
  --------------------------
  Reads HYPER PRO's Firebird DB and pushes to the local Siir server:
    * the catalogue at the NORMAL price (PV1_TTC) under EVERY barcode a product
      has - primary CODE_BARRE, synonym barcodes (CODEBARRE table), and the
      carton barcode CB_COLIS (priced PV1 x NBR_CB) - so no scan says "not found".
    * the CURRENTLY-ACTIVE promotions (PROMO=1 AND today within D1..D2 AND a real
      discount), including the POS quantity threshold QTE_PROMO (minQty), under
      the primary AND synonym barcodes.
  The server replaces the prior synced promotions each run. Read-only on HYPER PRO.

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

# Active-promo WHERE clause, reused for the primary + synonym arms (p = PRODUIT alias).
$promoWhere = "COALESCE({0}.PROMO,0)=1 AND COALESCE({0}.PP1_TTC,0)>0 AND COALESCE({0}.PV1_TTC,0)>0 AND {0}.PP1_TTC < {0}.PV1_TTC AND ({0}.D1 IS NULL OR {0}.D1 <= CURRENT_DATE) AND ({0}.D2 IS NULL OR {0}.D2 >= CURRENT_DATE)"

try {
  $isql = @(
    (Get-ChildItem 'C:\Program Files\Firebird', 'C:\Program Files (x86)\Firebird' -Recurse -Filter isql.exe -ErrorAction SilentlyContinue |
       Select-Object -First 1 -ExpandProperty FullName),
    'C:\HYPER PRO\isql.exe'
  ) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
  if (-not $isql) { throw 'isql.exe not found (part of Firebird). Set its full path in the script.' }

  foreach ($f in @($itemsTxt, $promoTxt)) { if (Test-Path $f) { Remove-Item $f -Force } }

  $U = 'ASCII_CHAR(31)'   # unit-separator delimiter (never appears in product text)
  $pv  = 'CAST(CAST(COALESCE({0}.PV1_TTC,0) AS NUMERIC(18,2)) AS VARCHAR(20))'
  @"
SET HEADING OFF;
SET LIST OFF;
OUTPUT '$itemsTxt';
SELECT p.CODE_BARRE || $U || COALESCE(p.PRODUIT,'') || $U || $($pv -f 'p')
  FROM PRODUIT p WHERE p.CODE_BARRE IS NOT NULL AND p.CODE_BARRE <> ''
UNION ALL
SELECT cb.CODE_BARRE_SYN || $U || COALESCE(p.PRODUIT,'') || $U || $($pv -f 'p')
  FROM CODEBARRE cb JOIN PRODUIT p ON p.CODE_BARRE = cb.CODE_BARRE
  WHERE cb.CODE_BARRE_SYN IS NOT NULL AND cb.CODE_BARRE_SYN <> ''
UNION ALL
SELECT p.CB_COLIS || $U || COALESCE(p.PRODUIT,'') || $U ||
  CAST(CAST(COALESCE(p.PV1_TTC,0)*COALESCE(p.NBR_CB,1) AS NUMERIC(18,2)) AS VARCHAR(20))
  FROM PRODUIT p WHERE p.CB_COLIS IS NOT NULL AND p.CB_COLIS <> '';
OUTPUT;
OUTPUT '$promoTxt';
SELECT p.CODE_BARRE || $U || COALESCE(p.PRODUIT,'') || $U || $($pv -f 'p') || $U ||
  CAST(CAST(p.PP1_TTC AS NUMERIC(18,2)) AS VARCHAR(20)) || $U || CAST(COALESCE(p.QTE_PROMO,1) AS VARCHAR(10))
  FROM PRODUIT p WHERE p.CODE_BARRE IS NOT NULL AND p.CODE_BARRE <> '' AND $($promoWhere -f 'p')
UNION ALL
SELECT cb.CODE_BARRE_SYN || $U || COALESCE(p.PRODUIT,'') || $U || $($pv -f 'p') || $U ||
  CAST(CAST(p.PP1_TTC AS NUMERIC(18,2)) AS VARCHAR(20)) || $U || CAST(COALESCE(p.QTE_PROMO,1) AS VARCHAR(10))
  FROM CODEBARRE cb JOIN PRODUIT p ON p.CODE_BARRE = cb.CODE_BARRE
  WHERE cb.CODE_BARRE_SYN IS NOT NULL AND cb.CODE_BARRE_SYN <> '' AND $($promoWhere -f 'p');
OUTPUT;
"@ | Set-Content -Encoding ASCII $sqlFile

  & $isql -q -u $FbUser -p $FbPass -ch $Charset -i $sqlFile "localhost:$DbPath" > $null 2>&1
  if (-not (Test-Path $itemsTxt)) { throw 'Query produced no output - check DB path / Firebird user+password.' }

  $US = [char]31
  $itemLines  = Get-Content -Encoding Default $itemsTxt | Where-Object { $_.Trim() -ne '' }
  $promoLines = if (Test-Path $promoTxt) { Get-Content -Encoding Default $promoTxt | Where-Object { $_.Trim() -ne '' } } else { @() }

  $sb = New-Object System.Text.StringBuilder
  [void]$sb.Append('{"items":[')
  $seen = New-Object System.Collections.Generic.HashSet[string]   # dedup: primary wins
  $nItems = 0; $first = $true
  foreach ($line in $itemLines) {
    $f = $line.Split($US); if ($f.Count -lt 3) { continue }
    $bc = $f[0].Trim(); if ($bc -eq '' -or -not $seen.Add($bc)) { continue }
    if (-not $first) { [void]$sb.Append(',') }; $first = $false; $nItems++
    [void]$sb.Append('{"barcode":').Append((J $bc)).Append(',"name":').Append((J $f[1].Trim())).Append(',"price":').Append($f[2].Trim()).Append('}')
  }
  [void]$sb.Append('],"promos":[')
  $pseen = New-Object System.Collections.Generic.HashSet[string]
  $nPromos = 0; $first = $true
  foreach ($line in $promoLines) {
    $f = $line.Split($US); if ($f.Count -lt 5) { continue }
    $bc = $f[0].Trim(); if ($bc -eq '' -or -not $pseen.Add($bc)) { continue }
    if (-not $first) { [void]$sb.Append(',') }; $first = $false; $nPromos++
    [void]$sb.Append('{"barcode":').Append((J $bc)).Append(',"title":').Append((J $f[1].Trim())).Append(',"oldPrice":').Append($f[2].Trim()).Append(',"newPrice":').Append($f[3].Trim()).Append(',"minQty":').Append($f[4].Trim()).Append('}')
  }
  [void]$sb.Append(']}')
  [IO.File]::WriteAllText($payload, $sb.ToString(), (New-Object System.Text.UTF8Encoding($false)))

  $resp = & curl.exe -s -X POST $SiirUrl -H "x-api-key: $ApiKey" -H "Content-Type: application/json; charset=utf-8" --data-binary "@$payload"
  Log "synced $nItems barcodes, $nPromos active promos -> $resp"
  Write-Output "OK: $nItems barcodes, $nPromos active promos -> $resp"
}
catch {
  Log "ERROR: $($_.Exception.Message)"
  Write-Output "ERROR: $($_.Exception.Message)"
  exit 1
}
