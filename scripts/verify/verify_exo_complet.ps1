Add-Type -AssemblyName System.Data
$out = "E:\OC DOCS\verify_exo_complet.txt"
$lines = New-Object System.Collections.Generic.List[string]
$conn = New-Object System.Data.SqlClient.SqlConnection('Server=(localdb)\mssqllocaldb;Database=ocdata;Integrated Security=True;Connect Timeout=15')

function Add-RS($cmd, $title) {
    $lines.Add("")
    $lines.Add("=== $title ===")
    $r = $cmd.ExecuteReader()
    $lines.Add((1..$r.FieldCount | ForEach-Object { $r.GetName($_ - 1) }) -join '|')
    while ($r.Read()) {
        $vals = 1..$r.FieldCount | ForEach-Object {
            $v = $r.GetValue($_ - 1)
            if ($v -is [DBNull]) { '' } elseif ($v -is [DateTime]) { $v.ToString('yyyy-MM-dd') } else { $v }
        }
        $lines.Add(($vals -join '|'))
    }
    $r.Close()
}

$posMap = @{27='101';28='102';29='201';30='301';31='401';32='501';33='502'}
$expectedSales = @{
    '2026-06-30' = @{101=25;102=10;201=30;301=8;401=15;501=40;502=12}
    '2026-07-01' = @{101=28;102=12;201=35;301=10;401=18;501=45;502=15}
    '2026-07-02' = @{101=22;102=8;201=28;301=12;401=20;501=38;502=10}
    '2026-07-03' = @{101=30;102=15;201=40;301=9;401=16;501=50;502=18}
    '2026-07-04' = @{101=35;102=18;201=45;301=14;401=22;501=55;502=20}
}
$expectedInv = @{
    'BP-2026-0142' = @{Date='2026-06-30';TTC=5.28}
    'DC-2026-0318' = @{Date='2026-06-30';TTC=191.40}
    'DC-2026-0320' = @{Date='2026-07-02';TTC=180.95}
    'BN-2026-0088' = @{Date='2026-07-03';TTC=52.80}
    'EH-2026-0205' = @{Date='2026-07-04';TTC=88.00}
}

try {
    $conn.Open()
    $cmd = $conn.CreateCommand()

    $cmd.CommandText = 'SELECT InventoryId, Finalized FROM oc.Inventory ORDER BY InventoryId'
    Add-RS $cmd 'INVENTORIES'

    $cmd.CommandText = @"
SELECT inv.InvoiceNumber, CONVERT(varchar(10), inv.InvoiceDate, 23) AS InvDate,
    s.Name AS Supplier, CAST(inv.Total AS decimal(10,2)) AS Total
FROM oc.Invoice inv
LEFT JOIN oc.Supplier s ON s.SupplierId = inv.Supplier
ORDER BY inv.InvoiceId
"@
    Add-RS $cmd 'INVOICES'

    $cmd.CommandText = @"
SELECT tt.Descrip, COUNT(tp.Item) AS Products
FROM oc.TillTape tt
LEFT JOIN oc.TillTapeProduct tp ON tp.TillTape = tt.TillTapeId
GROUP BY tt.Descrip
"@
    Add-RS $cmd 'TILLTAPE'

    $cmd.CommandText = @"
SELECT ss.SalesSourceId, CONVERT(varchar(10), ss.SalesDate, 23) AS SalesDate,
    CASE WHEN tts.SalesSource IS NOT NULL THEN 'TillTape' ELSE 'Import' END AS Src,
    SUM(ps.Qty) AS TotalQty, CAST(SUM(ps.GrossSalesTotal) AS decimal(10,2)) AS Gross
FROM oc.SalesSource ss
LEFT JOIN oc.TillTapeSale tts ON tts.SalesSource = ss.SalesSourceId
LEFT JOIN oc.ProductSale ps ON ps.SalesSource = ss.SalesSourceId
WHERE ss.SalesDate >= '2026-06-29' AND ss.SalesDate <= '2026-07-05'
GROUP BY ss.SalesSourceId, ss.SalesDate, tts.SalesSource
ORDER BY ss.SalesDate, ss.SalesSourceId
"@
    Add-RS $cmd 'SALES_BY_DAY'

    $cmd.CommandText = @"
SELECT CONVERT(varchar(10), ss.SalesDate, 23) AS D, ps.Product AS ItemId, ps.Qty, ps.GrossSalesTotal
FROM oc.ProductSale ps
JOIN oc.SalesSource ss ON ss.SalesSourceId = ps.SalesSource
WHERE ss.SalesDate >= '2026-06-29' AND ss.SalesDate <= '2026-07-05'
ORDER BY ss.SalesDate, ps.Product
"@
    Add-RS $cmd 'SALES_LINES'

    $cmd.CommandText = @"
SELECT CONVERT(varchar(10), SalesDate, 23) AS D, GrossSales, NetSales
FROM oc.DailySales
WHERE SalesDate >= '2026-06-29' AND SalesDate <= '2026-07-05'
ORDER BY SalesDate
"@
    try { Add-RS $cmd 'DAILY_SALES' } catch { $lines.Add('DAILY_SALES|none') }

    $cmd.CommandText = "SELECT COUNT(*) FROM oc.Waste"
    try { $lines.Add("WASTE_COUNT|$($cmd.ExecuteScalar())") } catch { $lines.Add('WASTE_COUNT|table_missing') }

    $cmd.CommandText = @"
SELECT i.Descrip, q.QtyOnHand, ru.Uom
FROM oc.ItemQtyOnHand q
JOIN oc.Item i ON i.ItemId = q.Item
LEFT JOIN oc.Uom ru ON ru.UomId = i.RecipeUom
WHERE i.Descrip IN (N'Pain burger', N'Cola 355 ml', N'Bouf', N'Cola') OR i.Descrip LIKE N'%Cola%'
ORDER BY i.Descrip
"@
    Add-RS $cmd 'QOH'

    $conn.Close()

    # Parse sales lines for qty check
    $byDate = @{}
    $inSales = $false
    foreach ($line in $lines) {
        if ($line -eq '=== SALES_LINES ===') { $inSales = $true; continue }
        if ($line -like '=== *' -and $line -ne '=== SALES_LINES ===') { $inSales = $false }
        if (-not $inSales -or $line -like 'D|*' -or $line -eq '') { continue }
        $p = $line -split '\|'
        if ($p.Count -lt 4) { continue }
        $d = $p[0]; $itemId = [int]$p[1]; $qty = [double]$p[2]
        $pos = $posMap[$itemId]
        if (-not $byDate.ContainsKey($d)) { $byDate[$d] = @{} }
        if ($pos) { $byDate[$d][$pos] = $qty }
    }

    $lines.Add("")
    $lines.Add("=== VERDICT ===")

    # Inventories
    $invText = ($lines -join "`n")
    if ($invText -match 'INVENTORIES' -and $invText -match 'True|1') {
        $lines.Add('OPENING_INVENTORY|OK (Finalize)')
    } else {
        $lines.Add('OPENING_INVENTORY|CHECK')
    }

    foreach ($num in $expectedInv.Keys) {
        if ($invText -match [regex]::Escape($num)) {
            $lines.Add("INVOICE_FOUND|$num")
        } else {
            $lines.Add("INVOICE_MISSING|$num")
        }
    }

    if ($invText -match 'Tout') { $lines.Add('TILLTAPE_LIST|OK') } else { $lines.Add('TILLTAPE_LIST|MISSING') }

    foreach ($d in ($expectedSales.Keys | Sort-Object)) {
        if (-not $byDate.ContainsKey($d)) {
            $lines.Add("SALES_MISSING_DAY|$d")
            continue
        }
        $dayOk = $true
        foreach ($pos in $expectedSales[$d].Keys) {
            $exp = $expectedSales[$d][$pos]
            $got = $byDate[$d][$pos]
            if ($null -eq $got -or [Math]::Abs($got - $exp) -gt 0.01) {
                $lines.Add("SALES_BAD|$d|POS$pos|got=$got|exp=$exp")
                $dayOk = $false
            }
        }
        if ($dayOk) { $lines.Add("SALES_OK|$d") }
    }

    if ($invText -match 'DAILY_SALES' -and $invText -notmatch 'DAILY_SALES\|none') {
        $n = ([regex]::Matches($invText, '2026-0[67]')).Count
        $lines.Add("DAILY_SALES|partial_or_done")
    } else {
        $lines.Add('DAILY_SALES|NOT_YET')
    }

    if ($invText -match 'WASTE_COUNT\|0') { $lines.Add('WASTE|NOT_YET') }
    elseif ($invText -match 'WASTE_COUNT\|[1-9]') { $lines.Add('WASTE|OK') }
    else { $lines.Add('WASTE|NOT_YET') }

    $lines | Set-Content $out -Encoding UTF8
    Write-Output "OK -> $out"
} catch {
    "ERR: $($_.Exception.Message)" | Set-Content $out -Encoding UTF8
    Write-Output "ERR $($_.Exception.Message)"
    exit 1
}
