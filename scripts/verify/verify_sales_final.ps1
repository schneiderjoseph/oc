Add-Type -AssemblyName System.Data
$conn = New-Object System.Data.SqlClient.SqlConnection('Server=(localdb)\mssqllocaldb;Database=ocdata;Integrated Security=True')
$conn.Open(); $cmd = $conn.CreateCommand()

$cmd.CommandText = @"
SELECT tp.Idx, tp.Item, i.Descrip, i.Type
FROM oc.TillTapeProduct tp
JOIN oc.Item i ON i.ItemId = tp.Item
JOIN oc.TillTape tt ON tt.TillTapeId = tp.TillTape
WHERE tt.Descrip LIKE N'%Tout%'
ORDER BY tp.Idx
"@
$r = $cmd.ExecuteReader(); Write-Output '---TillTape products---'; while($r.Read()){ Write-Output "idx=$($r[0]) item=$($r[1]) | $($r[2]) type=$($r[3])" }; $r.Close()

$cmd.CommandText = @"
SELECT ps.Product, ps.Qty, ps.GrossSalesTotal
FROM oc.ProductSale ps WHERE ps.SalesSource = 0 ORDER BY ps.Product
"@
$r = $cmd.ExecuteReader(); Write-Output '---Sales---'; while($r.Read()){ Write-Output "productId=$($r[0]) qty=$($r[1]) gross=$($r[2])" }; $r.Close()

$cmd.CommandText = "SELECT SUM(Qty) AS TotalQty, SUM(GrossSalesTotal) AS TotalGross FROM oc.ProductSale WHERE SalesSource=0"
$r = $cmd.ExecuteReader(); if($r.Read()){ Write-Output "TOTAL qty=$($r[0]) gross=$($r[1])" }; $r.Close()

$cmd.CommandText = @"
SELECT ss.SalesDate, ss.UsageSource FROM oc.SalesSource ss WHERE ss.SalesSourceId=0
"@
$r = $cmd.ExecuteReader(); if($r.Read()){ Write-Output "DATE=$($r[0]) usageSource=$($r[1])" }; $r.Close()

$conn.Close()
