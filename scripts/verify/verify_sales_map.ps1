Add-Type -AssemblyName System.Data
$conn = New-Object System.Data.SqlClient.SqlConnection('Server=(localdb)\mssqllocaldb;Database=ocdata;Integrated Security=True')
$conn.Open(); $cmd = $conn.CreateCommand()

$cmd.CommandText = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA='oc' AND TABLE_NAME LIKE '%Product%' ORDER BY TABLE_NAME"
$r = $cmd.ExecuteReader(); while($r.Read()){ Write-Output $r[0] }; $r.Close()

$cmd.CommandText = @"
SELECT ps.Product, ps.Qty, ps.GrossSalesTotal, ps.TotalCost
FROM oc.ProductSale ps WHERE ps.SalesSource = 0 ORDER BY ps.Product
"@
$r = $cmd.ExecuteReader(); Write-Output '---ProductSale source 0---'; while($r.Read()){ Write-Output "prod=$($r[0]) qty=$($r[1]) gross=$($r[2]) cost=$($r[3])" }; $r.Close()

# Find product table - maybe MenuRecipe or similar
$cmd.CommandText = @"
SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA='oc'
AND COLUMN_NAME IN ('PosId','POSId','PluNumber')
"@
$r = $cmd.ExecuteReader(); Write-Output '---tables with POS---'; while($r.Read()){ Write-Output $r[0] }; $r.Close()

$cmd.CommandText = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA='oc' AND TABLE_NAME='MenuRecipe' ORDER BY ORDINAL_POSITION"
$r = $cmd.ExecuteReader(); Write-Output '---MenuRecipe cols---'; while($r.Read()){ Write-Output $r[0] }; $r.Close()

$cmd.CommandText = @"
SELECT mr.MenuRecipeId, mr.Descrip, mr.PosId
FROM oc.MenuRecipe mr ORDER BY mr.PosId
"@
try {
    $r = $cmd.ExecuteReader(); Write-Output '---MenuRecipe---'; while($r.Read()){ Write-Output "$($r[2]) | id=$($r[0]) | $($r[1])" }; $r.Close()
} catch {}

$cmd.CommandText = @"
SELECT i.Descrip, q.QtyOnHand FROM oc.ItemQtyOnHand q
JOIN oc.Item i ON i.ItemId = q.Item
WHERE i.Descrip IN (N'Pain burger', N'Bouf', N'Pommes de terre', N'Cola 355 ml')
"@
$r = $cmd.ExecuteReader(); Write-Output '---QOH---'; while($r.Read()){ Write-Output "$($r[0])=$($r[1])" }; $r.Close()

$conn.Close()
