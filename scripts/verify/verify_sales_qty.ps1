Add-Type -AssemblyName System.Data
$conn = New-Object System.Data.SqlClient.SqlConnection('Server=(localdb)\mssqllocaldb;Database=ocdata;Integrated Security=True')
$conn.Open(); $cmd = $conn.CreateCommand()

# ItemId -> POS (from prior session)
$posMap = @{27='101';28='102';29='201';30='301';31='401';32='501';33='502'}
$expected = @{
    '2026-06-30' = @{101=25;102=10;201=30;301=8;401=15;501=40;502=12}
    '2026-07-01' = @{101=28;102=12;201=35;301=10;401=18;501=45;502=15}
    '2026-07-02' = @{101=22;102=8;201=28;301=12;401=20;501=38;502=10}
    '2026-07-03' = @{101=30;102=15;201=40;301=9;401=16;501=50;502=18}
    '2026-07-04' = @{101=35;102=18;201=45;301=14;401=22;501=55;502=20}
}

$cmd.CommandText = @"
SELECT CONVERT(varchar(10), ss.SalesDate, 23) AS D, ps.Product, ps.Qty, ps.GrossSalesTotal
FROM oc.ProductSale ps
JOIN oc.SalesSource ss ON ss.SalesSourceId = ps.SalesSource
WHERE ss.SalesDate >= '2026-06-29' AND ss.SalesDate <= '2026-07-05'
ORDER BY ss.SalesDate, ps.Product
"@
$r = $cmd.ExecuteReader()
$byDate = @{}
while ($r.Read()) {
    $d = $r[0]; $pid = [int]$r[1]; $qty = [double]$r[2]
    $pos = $posMap[$pid]
    if (-not $byDate.ContainsKey($d)) { $byDate[$d] = @{} }
    $byDate[$d][$pos] = $qty
}
$r.Close()

Write-Output "=== SALES QTY CHECK ==="
foreach ($d in ($expected.Keys | Sort-Object)) {
    $dayOk = $true
    Write-Output "--- $d ---"
    if (-not $byDate.ContainsKey($d)) {
        Write-Output "MISSING_DAY"
        continue
    }
    foreach ($pos in ($expected[$d].Keys | Sort-Object)) {
        $exp = $expected[$d][$pos]
        $got = $byDate[$d][$pos]
        if ($null -eq $got) {
            Write-Output "POS $pos MISSING"
            $dayOk = $false
        } elseif ([Math]::Abs($got - $exp) -gt 0.01) {
            Write-Output "POS $pos BAD got=$got exp=$exp"
            $dayOk = $false
        } else {
            Write-Output "POS $pos OK $got"
        }
    }
    if ($dayOk) { Write-Output "DAY_OK" }
}

$conn.Close()
