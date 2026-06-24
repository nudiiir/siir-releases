#requires -version 5
<#
  Siir x HYPER PRO auto-sync
  --------------------------
  Reads HYPER PRO's Firebird PRODUIT table and pushes the live catalogue to the
  local Siir server (/api/stock). The price sent is the EFFECTIVE price, exactly
  as HYPER PRO shows it:
     promo price (PP1_TTC) ONLY when PROMO=1 AND today is within D1..D2,
     otherwise the normal sale price (PV1_TTC).
  Read-only on HYPER PRO's database - it never writes to it.

  Deploy on a store PC (PowerShell):
    mkdir C:\ProgramData\SiirSync 2>$null
    curl.exe -L -o C:\ProgramData\SiirSync\hyperpro-sync.ps1 `
      https://raw.githubusercontent.com/nudiiir/siir-releases/main/connectors/hyperpro-sync.ps1

    # test once (ApiKey = Siir panel -> "amr raf3 al-catalogue"):
    powershell -ExecutionPolicy Bypass -File C:\ProgramData\SiirSync\hyperpro-sync.ps1 -ApiKey sk_xxx

    # then schedule every 15 min, running as SYSTEM (works when logged out):
    schtasks /create /tn "SiirSync" /sc minute /mo 15 /ru SYSTEM /f ^
      /tr "powershell -NoProfile -ExecutionPolicy Bypass -File C:\ProgramData\SiirSync\hyperpro-sync.ps1 -ApiKey sk_xxx"
#>

param(
  [Parameter(Mandatory = $true)] [string] $ApiKey,
  [string] $DbPath  = 'C:\HYPER PRO\raya.FDB',     # the .FDB in HYPER PRO's status bar
  [string] $SiirUrl = 'http://localhost:3000/api/stock',
  [string] $FbUser  = 'SYSDBA',
  [string] $FbPass  = 'masterkey',                 # Firebird default; change if the store set one
  [string] $Charset = 'WIN1252'                    # if names look garbled, try UTF8 or NONE
)

$ErrorActionPreference = 'Stop'
$work = Join-Path $env:ProgramData 'SiirSync'
New-Item -ItemType Directory -Force $work | Out-Null
$sqlFile = Join-Path $work 'query.sql'
$rows    = Join-Path $work 'rows.txt'
$payload = Join-Path $work 'scanner_prices.txt'
$log     = Join-Path $work 'sync.log'
function Log($m) { "[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $m | Add-Content $log }

try {
  $isql = @(
    (Get-ChildItem 'C:\Program Files\Firebird','C:\Program Files (x86)\Firebird' -Recurse -Filter isql.exe -ErrorAction SilentlyContinue |
       Select-Object -First 1 -ExpandProperty FullName),
    'C:\HYPER PRO\isql.exe'
  ) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
  if (-not $isql) { throw 'isql.exe not found (part of Firebird). Set its full path in the script.' }

  if (Test-Path $rows) { Remove-Item $rows -Force }

  @"
SET HEADING OFF;
SET LIST OFF;
OUTPUT '$rows';
SELECT CODE_BARRE || '|' || COALESCE(PRODUIT,'') || '|' ||
  CAST(CAST(COALESCE(
    CASE WHEN COALESCE(PROMO,0)=1 AND COALESCE(PP1_TTC,0)>0
              AND (D1 IS NULL OR D1 <= CURRENT_DATE)
              AND (D2 IS NULL OR D2 >= CURRENT_DATE)
         THEN PP1_TTC ELSE PV1_TTC END, 0) AS NUMERIC(18,2)) AS VARCHAR(20))
FROM PRODUIT
WHERE CODE_BARRE IS NOT NULL AND CODE_BARRE <> '';
OUTPUT;
"@ | Set-Content -Encoding ASCII $sqlFile

  & $isql -q -u $FbUser -p $FbPass -ch $Charset -i $sqlFile "localhost:$DbPath" > $null 2>&1
  if (-not (Test-Path $rows)) { throw 'Query produced no output - check DB path / Firebird user+password.' }

  $lines = Get-Content -Encoding Default $rows | Where-Object { $_.Trim() -ne '' }
  $outList = New-Object System.Collections.Generic.List[string]
  $outList.Add('CODE_BARRE|PRODUIT|PRIX')
  if ($lines) { $outList.AddRange([string[]]$lines) }
  [IO.File]::WriteAllLines($payload, $outList, (New-Object Text.UTF8Encoding($false)))

  $resp = & curl.exe -s -X POST $SiirUrl -H "x-api-key: $ApiKey" -H "Content-Type: text/plain; charset=utf-8" --data-binary "@$payload"
  Log "synced $($lines.Count) products -> $resp"
  Write-Output "OK: synced $($lines.Count) products -> $resp"
}
catch {
  Log "ERROR: $($_.Exception.Message)"
  Write-Output "ERROR: $($_.Exception.Message)"
  exit 1
}
