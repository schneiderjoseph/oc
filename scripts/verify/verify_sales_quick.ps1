Add-Type -AssemblyName System.Data
$conn = New-Object System.Data.SqlClient.SqlConnection('Server=(localdb)\mssqllocaldb;Database=ocdata;Integrated Security=True')
$conn.Open(); $cmd = $conn.CreateCommand()

$cmd.CommandText = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA='oc' AND TABLE_NAME LIKE '%Product%' ORDER BY TABLE_NAME"
$r = $cmd.ExecuteReader(); while($r.Read()){ Write-Output $r[0] }; $r.Close()

$cmd.CommandText = @"
SELECT tt.Descrip, COUNT(tp.Item) AS Cnt
FROM oc.TillTape tt LEFT JOIN oc.TillTapeProduct tp ON tp.TillTape = tt.TillTapeId
GROUP BY tt.Descrip
"@
$r = $cmd.ExecuteReader(); Write-Output '---TILLTAPE---'; while($r.Read()){ Write-Output "$($r[0]) | products=$($r[1])" }; $r.Close()

$cmd.CommandText = @"
SELECT ss.SalesDate, ss.SalesSourceId, ss.UsageSource, tts.TillTape
FROM oc.SalesSource ss
LEFT JOIN oc.TillTapeSale tts ON tts.SalesSource = ss.SalesSourceId
ORDER BY ss.SalesDate
"@
$r = $cmd.ExecuteReader(); Write-Output '---SOURCES---'; while($r.Read()){ Write-Output "$($r[0]) | id=$($r[1]) | usage=$($r[2]) | tilltape=$($r[3])" }; $r.Close()

$cmd.CommandText = @"
SELECT ss.SalesDate, si.PluNumber, si.Descrip, si.QtySold, si.GrossSales, si.Status
FROM oc.SalesItem si JOIN oc.SalesSource ss ON ss.SalesSourceId = si.SalesSource
ORDER BY ss.SalesDate, si.PluNumber
"@
$r = $cmd.ExecuteReader(); Write-Output '---SALESITEMS---'; while($r.Read()){ Write-Output "$($r[0]) | $($r[1]) | $($r[2]) | qty=$($r[3]) | gross=$($r[4]) | $($r[5])" }; $r.Close()

$cmd.CommandText = @"
SELECT rs.SalesDate, rs.UsageSource, rsi.Qty, i.Descrip
FROM oc.RetailSale rs
JOIN oc.RetailSaleItem rsi ON rsi.RetailSale = rs.RetailSaleId
JOIN oc.Item i ON i.ItemId = rsi.Item
ORDER BY rs.SalesDate, i.Descrip
"@
$r = $cmd.ExecuteReader(); Write-Output '---RETAILSALE---'; while($r.Read()){ Write-Output "$($r[0]) | $($r[1]) | $($r[2]) | $($r[3])" }; $r.Close()

$conn.Close()
