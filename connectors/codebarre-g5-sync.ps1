#requires -version 5
<#
  Siir x CodeBarre G5 (SoftMANAGING) auto-sync
  --------------------------------------------
  Reads CodeBarre G5's SQL Server database [BarreCodeBASE] and pushes to the
  local Siir server in ONE call (POST /api/stock):
    * the catalogue at the current sale price  (Produit JOIN Stock.PrixV)
    * the CURRENTLY-ACTIVE flash promotions     (VenteFlash: PrixFlash, while
      today is within D1..D2 AND the current time within its daily hours;
      old price = Stock.PrixV, only real discounts newPrice < oldPrice)
  The server replaces the prior synced/imported promotions each run, so stale
  ones never pile up. Read-only on the POS database.

  Deploy on a store PC (PowerShell, as admin):
    mkdir C:\ProgramData\SiirSync 2>$null
    curl.exe -L -o C:\ProgramData\SiirSync\codebarre-g5-sync.ps1 `
      https://raw.githubusercontent.com/nudiiir/siir-releases/main/connectors/codebarre-g5-sync.ps1
    # test one sync:
    powershell -ExecutionPolicy Bypass -File C:\ProgramData\SiirSync\codebarre-g5-sync.ps1 -ApiKey sk_xxx -Once
    # then run every 15 min in the background:
    schtasks /create /tn "SiirSync" /sc minute /mo 15 /ru SYSTEM /f `
      /tr "powershell -NoProfile -ExecutionPolicy Bypass -File C:\ProgramData\SiirSync\codebarre-g5-sync.ps1 -ApiKey sk_xxx -Once"

  If SQL needs a login (not Windows auth):  add  -SqlUser sa -SqlPass yourpass
  If the instance isn't auto-detected:       add  -Server ".\SQLEXPRESS"
#>

param(
  [Parameter(Mandatory = $true)] [string] $ApiKey,
  [string] $Server   = '',                              # SQL instance; '' = auto-detect
  [string] $Database = 'BarreCodeBASE',
  [string] $SiirUrl  = 'http://localhost:3000/api/stock',
  [string] $SqlUser  = '',                              # '' = Windows (integrated) auth
  [string] $SqlPass  = '',
  [switch] $Once,                                       # one sync then exit (for Task Scheduler)
  [int]    $IntervalSeconds = 300                       # loop mode interval
)

$ErrorActionPreference = 'Stop'
$work = Join-Path $env:ProgramData 'SiirSync'
New-Item -ItemType Directory -Force $work | Out-Null
$log = Join-Path $work 'sync.log'
function Log($m, $c = 'Gray') {
  $line = "[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $m
  Write-Host $line -ForegroundColor $c
  try { Add-Content -Path $log -Value $line } catch {}
}

function Get-CandidateServers {
  if ($Server) { return @($Server) }
  $list = New-Object System.Collections.Generic.List[string]
  Get-Service -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -eq 'MSSQLSERVER' -or $_.Name -like 'MSSQL$*' } |
    ForEach-Object {
      if ($_.Name -eq 'MSSQLSERVER') { $list.Add('.') }
      else { $list.Add('.\' + $_.Name.Substring(6)) }   # MSSQL$SQLEXPRESS -> .\SQLEXPRESS
    }
  foreach ($s in '.', '.\SQLEXPRESS', 'localhost', '(local)', 'localhost\SQLEXPRESS') {
    if (-not $list.Contains($s)) { $list.Add($s) }
  }
  return $list
}

function New-Conn {
  $auth = if ($SqlUser) { "User ID=$SqlUser;Password=$SqlPass" } else { 'Integrated Security=SSPI' }
  foreach ($srv in (Get-CandidateServers)) {
    $cs = "Server=$srv;Database=$Database;$auth;TrustServerCertificate=True;Connect Timeout=5"
    try {
      $c = New-Object System.Data.SqlClient.SqlConnection $cs
      $c.Open()
      Log "connected to SQL Server: $srv" 'Green'
      return $c
    } catch {
      Log "$srv : $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'DarkGray'
    }
  }
  throw "Could not connect to SQL Server (database=$Database). Pass -Server '.\SQLEXPRESS' or -SqlUser/-SqlPass."
}

function Invoke-Rows($conn, $sql) {
  $cmd = $conn.CreateCommand(); $cmd.CommandText = $sql; $cmd.CommandTimeout = 120
  $dt = New-Object System.Data.DataTable
  (New-Object System.Data.SqlClient.SqlDataAdapter $cmd).Fill($dt) | Out-Null
  return $dt
}

$PRODUCT_SQL = @'
SELECT p.Code AS Code, p.Designation AS Designation, s.PrixV AS PrixV
FROM dbo.Produit p
INNER JOIN dbo.Stock s ON s.Code = p.Code
WHERE s.PrixV IS NOT NULL AND s.PrixV > 0
'@

$PROMO_SQL = @'
SELECT vf.Code AS barcode, s.PrixV AS oldPrice, vf.PrixFlash AS newPrice
FROM dbo.VenteFlash vf
INNER JOIN dbo.Stock s ON s.Code = vf.Code
WHERE vf.PrixFlash > 0 AND s.PrixV > vf.PrixFlash
  AND CAST(GETDATE() AS date) BETWEEN vf.DateDebut AND vf.DateFin
  AND CAST(GETDATE() AS time) BETWEEN ISNULL(vf.HeureDebut, CAST('00:00:00' AS time))
                                  AND ISNULL(vf.HeureFin,   CAST('23:59:59' AS time))
'@

# PowerShell 5.1 unwraps single-element arrays in ConvertTo-Json : force a JSON array.
function To-JsonArray($items) {
  $a = @($items)
  if ($a.Count -eq 0) { return '[]' }
  if ($a.Count -eq 1) { return '[' + ($a[0] | ConvertTo-Json -Depth 5 -Compress) + ']' }
  return ($a | ConvertTo-Json -Depth 5 -Compress)
}

function Sync-Once {
  $conn = New-Conn
  try {
    $prod  = Invoke-Rows $conn $PRODUCT_SQL
    $promo = Invoke-Rows $conn $PROMO_SQL
  } finally { $conn.Close() }

  $products = New-Object System.Collections.Generic.List[object]
  foreach ($r in $prod.Rows) {
    $products.Add([ordered]@{ Code = [string]$r.Code; Designation = [string]$r.Designation; PrixV = [double]$r.PrixV })
  }
  $promos = New-Object System.Collections.Generic.List[object]
  foreach ($r in $promo.Rows) {
    $promos.Add([ordered]@{ barcode = [string]$r.barcode; oldPrice = [double]$r.oldPrice; newPrice = [double]$r.newPrice })
  }

  if ($products.Count -eq 0) { Log "no products with a price found - nothing pushed (check the DB)." 'Yellow'; return }

  $json  = '{"products":' + (To-JsonArray $products) + ',"promos":' + (To-JsonArray $promos) + '}'
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
  $resp  = Invoke-RestMethod -Uri $SiirUrl -Method Post -Headers @{ 'x-api-key' = $ApiKey } -ContentType 'application/json; charset=utf-8' -Body $bytes -TimeoutSec 120
  Log ("OK - {0} products, {1} promos pushed" -f $resp.productCount, $resp.promoCount) 'Cyan'
}

if ($Once) {
  Sync-Once
} else {
  Log "Siir sync running - every $IntervalSeconds s. Ctrl+C to stop." 'Yellow'
  while ($true) {
    try { Sync-Once } catch { Log "error: $($_.Exception.Message)" 'Red' }
    Start-Sleep -Seconds $IntervalSeconds
  }
}
