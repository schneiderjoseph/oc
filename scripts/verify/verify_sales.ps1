Add-Type -AssemblyName System.Data
$out = "E:\OC DOCS\verify_sales_db.txt"
$lines = New-Object System.Collections.Generic.List[string]
$conn = New-Object System.Data.SqlClient.SqlConnection('Server=(localdb)\mssqllocaldb;Database=ocdata;Integrated Security=True;Connect Timeout=10')

function Add-RS($cmd, $title) {
    $lines.Add($title)
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

$expected = @{
    'Burger classique' = 25
    'Burger bacon' = 10
    'Frites moyennes' = 30
    'Salade César' = 8
    'Bol chili' = 15
    'Cola 355 ml' = 40
    'Cola' = 40
    'Eau 500 ml' = 12
    'Eau' = 12
}

try {
    $conn.Open()
    $cmd = $conn.CreateCommand()

    $cmd.CommandText = @"
SELECT tt.TillTapeId, tt.Descrip, COUNT(tp.Product) AS ProductCount
FROM oc.TillTape tt
LEFT JOIN oc.TillTapeProduct tp ON tp.TillTape = tt.TillTapeId
GROUP BY tt.TillTapeId, tt.Descrip
"@
    Add-RS $cmd 'TILLTAPE_LISTS'

    $cmd.CommandText = @"
SELECT ts.TillTapeSaleId, ts.SalesDate, tt.Descrip AS TillTapeList, ts.TotalQty, ts.TotalSales
FROM oc.TillTapeSale ts
LEFT JOIN oc.TillTape tt ON tt.TillTapeId = ts.TillTape
ORDER BY ts.SalesDate
"@
    Add-RS $cmd 'TILLTAPE_SALES'

    $cmd.CommandText = @"
SELECT ss.SalesSourceId, ss.SalesDate, ss.UsageSource
FROM oc.SalesSource ss
WHERE ss.SalesDate >= '2026-06-29' AND ss.SalesDate <= '2026-07-05'
ORDER BY ss.SalesDate
"@
    Add-RS $cmd 'SALES_SOURCES_WEEK'

    $cmd.CommandText = @"
SELECT ss.SalesDate, p.Descrip AS Product, p.PosId, rsi.QtySold, rsi.GrossSales, rsi.Price
FROM oc.RetailSale rs
JOIN oc.SalesSource ss ON ss.SalesSourceId = rs.SalesSource
JOIN oc.RetailSaleItem rsi ON rsi.RetailSale = rs.RetailSaleId
JOIN oc.Product p ON p.ProductId = rsi.Product
WHERE ss.SalesDate >= '2026-06-29' AND ss.SalesDate <= '2026-07-05'
ORDER BY ss.SalesDate, p.Descrip
"@
    Add-RS $cmd 'RETAIL_SALE_LINES'

    $cmd.CommandText = @"
SELECT ss.SalesDate, si.PluNumber, si.Descrip, si.QtySold, si.GrossSales, si.Status
FROM oc.SalesItem si
JOIN oc.SalesSource ss ON ss.SalesSourceId = si.SalesSource
WHERE ss.SalesDate >= '2026-06-29' AND ss.SalesDate <= '2026-07-05'
ORDER BY ss.SalesDate, si.PluNumber
"@
    Add-RS $cmd 'SALES_ITEMS'

    $cmd.CommandText = @"
SELECT i.Descrip, q.QtyOnHand, ru.Uom
FROM oc.ItemQtyOnHand q
JOIN oc.Item i ON i.ItemId = q.Item
LEFT JOIN oc.Uom ru ON ru.UomId = i.RecipeUom
WHERE i.Descrip IN (N'Pain burger', N'Bœuf haché 80/20', N'Bouf', N'Pommes de terre', N'Cola 355 ml')
ORDER BY i.Descrip
"@
    Add-RS $cmd 'QOH_AFTER_LUNDI_SALES'

    $conn.Close()

    $text = $lines -join "`n"
    $lines.Add('VERDICT')
    if ($text -match '2026-06-30') { $lines.Add('OK|Sales found for 2026-06-30') } else { $lines.Add('MISSING|No sales on 2026-06-30') }

    $lines | Set-Content $out -Encoding UTF8
    Write-Output "OK -> $out"
} catch {
    "ERR: $($_.Exception.Message)" | Set-Content $out -Encoding UTF8
    Write-Output "ERR $($_.Exception.Message)"
    exit 1
}
