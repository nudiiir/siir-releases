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
  $cmd = $conn.CreateCommand(); $cmd.CommandText = $sql; $cmd.CommandTimeout = 600
  $dt = New-Object System.Data.DataTable
  (New-Object System.Data.SqlClient.SqlDataAdapter $cmd).Fill($dt) | Out-Null
  return $dt
}

# Price = the store's proven logic: most-recent of StockDET (price history, by
# DateAch) and Ventes (last UNIT sale, by DateVente), with a MultiCodeBare
# sibling-barcode fallback. COLLATE DATABASE_DEFAULT on every Code comparison
# (this POS mixes Arabic/French column collations).
$PRODUCT_SQL = @'
SET NOCOUNT ON;
SELECT Code, PrixV, DateAch INTO #LatestStock FROM (
  SELECT LTRIM(RTRIM(Code)) COLLATE DATABASE_DEFAULT AS Code, PrixV, DateAch,
         ROW_NUMBER() OVER (PARTITION BY LTRIM(RTRIM(Code)) ORDER BY DateAch DESC) AS rn
  FROM dbo.StockDET WHERE PrixV > 0
) x WHERE rn = 1;

SELECT Code, PrixV, DateVente INTO #LatestVente FROM (
  SELECT LTRIM(RTRIM(v2.Code)) COLLATE DATABASE_DEFAULT AS Code, v2.PrixV, v2.DateVente, v2.NumVente,
         ROW_NUMBER() OVER (PARTITION BY LTRIM(RTRIM(v2.Code)) ORDER BY v2.DateVente DESC, v2.NumVente DESC) AS rn
  FROM dbo.Ventes v2
  WHERE v2.PrixV > 0 AND v2.Qte = 1
    AND NOT (
      LTRIM(RTRIM(v2.Code)) COLLATE DATABASE_DEFAULT IN (SELECT LTRIM(RTRIM(CodeSTK)) COLLATE DATABASE_DEFAULT FROM dbo.DeuxCodeEmbalage)
      AND v2.PrixV <> ISNULL((SELECT TOP 1 s.PrixV FROM dbo.StockDET s WHERE LTRIM(RTRIM(s.Code)) COLLATE DATABASE_DEFAULT = LTRIM(RTRIM(v2.Code)) COLLATE DATABASE_DEFAULT ORDER BY s.DateAch DESC), v2.PrixV)
    )
) v WHERE rn = 1;

SELECT
  LTRIM(RTRIM(p.Code)) AS Code,
  LTRIM(RTRIM(p.Designation)) AS Designation,
  CAST(CASE
    WHEN sd.DateAch IS NOT NULL AND vt.DateVente IS NOT NULL THEN CASE WHEN vt.DateVente >= sd.DateAch THEN vt.PrixV ELSE sd.PrixV END
    WHEN sd.DateAch IS NOT NULL THEN sd.PrixV
    WHEN vt.DateVente IS NOT NULL THEN vt.PrixV
    ELSE sib.PrixV
  END AS real) AS PrixV
FROM dbo.Produit p
LEFT JOIN #LatestStock sd ON sd.Code = LTRIM(RTRIM(p.Code)) COLLATE DATABASE_DEFAULT
LEFT JOIN #LatestVente vt ON vt.Code = LTRIM(RTRIM(p.Code)) COLLATE DATABASE_DEFAULT
OUTER APPLY (
  SELECT TOP 1 ls.PrixV FROM #LatestStock ls
  WHERE ls.Code IN (
    SELECT LTRIM(RTRIM(mb2.MultiCode)) COLLATE DATABASE_DEFAULT FROM dbo.MultiCodeBare mb2
    WHERE LTRIM(RTRIM(mb2.Pere)) COLLATE DATABASE_DEFAULT = (
      SELECT TOP 1 LTRIM(RTRIM(mb1.Pere)) COLLATE DATABASE_DEFAULT FROM dbo.MultiCodeBare mb1
      WHERE LTRIM(RTRIM(mb1.MultiCode)) COLLATE DATABASE_DEFAULT = LTRIM(RTRIM(p.Code)) COLLATE DATABASE_DEFAULT
    ) AND mb2.MultiCode IS NOT NULL AND mb2.MultiCode <> ''
  ) ORDER BY ls.DateAch DESC
) sib
WHERE CASE
    WHEN sd.DateAch IS NOT NULL AND vt.DateVente IS NOT NULL THEN CASE WHEN vt.DateVente >= sd.DateAch THEN vt.PrixV ELSE sd.PrixV END
    WHEN sd.DateAch IS NOT NULL THEN sd.PrixV
    WHEN vt.DateVente IS NOT NULL THEN vt.PrixV
    ELSE sib.PrixV
  END > 0;

DROP TABLE #LatestStock; DROP TABLE #LatestVente;
'@

$PROMO_SQL = @'
SELECT vf.Code AS barcode, s.PrixV AS oldPrice, vf.PrixFlash AS newPrice
FROM dbo.VenteFlash vf
INNER JOIN dbo.Stock s ON s.Code COLLATE DATABASE_DEFAULT = vf.Code COLLATE DATABASE_DEFAULT
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
    if ($null -eq $r) { continue }
    $code = [Convert]::ToString($r.Item('Code')).Trim()
    if (-not $code) { continue }
    $products.Add([ordered]@{
      Code        = $code
      Designation = [Convert]::ToString($r.Item('Designation')).Trim()
      PrixV       = [Convert]::ToDouble($r.Item('PrixV'))
    })
  }
  $promos = New-Object System.Collections.Generic.List[object]
  foreach ($r in $promo.Rows) {
    if ($null -eq $r) { continue }
    $bc = [Convert]::ToString($r.Item('barcode')).Trim()
    if (-not $bc) { continue }
    $promos.Add([ordered]@{
      barcode  = $bc
      oldPrice = [Convert]::ToDouble($r.Item('oldPrice'))
      newPrice = [Convert]::ToDouble($r.Item('newPrice'))
    })
  }

  if ($products.Count -eq 0) { Log "no products with a price found - nothing pushed (check the DB)." 'Yellow'; return }

  $json  = '{"products":' + (To-JsonArray $products) + ',"promos":' + (To-JsonArray $promos) + '}'
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
  $resp  = Invoke-RestMethod -Uri $SiirUrl -Method Post -Headers @{ 'x-api-key' = $ApiKey } -ContentType 'application/json; charset=utf-8' -Body $bytes -TimeoutSec 120
  Log ("OK - {0} products, {1} promos pushed" -f $resp.productCount, $resp.promoCount) 'Cyan'
}

function Report-Error($e) {
  Log ("error [line {0}]: {1} :: {2}" -f $e.InvocationInfo.ScriptLineNumber, $e.InvocationInfo.Line.Trim(), $e.Exception.Message) 'Red'
}

if ($Once) {
  try { Sync-Once } catch { Report-Error $_; exit 1 }
} else {
  Log "Siir sync running - every $IntervalSeconds s. Ctrl+C to stop." 'Yellow'
  while ($true) {
    try { Sync-Once } catch { Report-Error $_ }
    Start-Sleep -Seconds $IntervalSeconds
  }
}
