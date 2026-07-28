Add-Type -AssemblyName System.Data
$conn = New-Object System.Data.SqlClient.SqlConnection('Server=(localdb)\mssqllocaldb;Database=ocdata;Integrated Security=True')
$conn.Open(); $cmd = $conn.CreateCommand()

$cmd.CommandText = @"
SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES t
WHERE t.TABLE_SCHEMA='oc'
AND EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS c WHERE c.TABLE_SCHEMA='oc' AND c.TABLE_NAME=t.TABLE_NAME AND c.COLUMN_NAME='PosId')
"@
$r = $cmd.ExecuteReader(); Write-Output '---PosId tables---'; while($r.Read()){ Write-Output $r[0] }; $r.Close()

$cmd.CommandText = @"
SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES t
WHERE t.TABLE_SCHEMA='oc'
AND EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS c WHERE c.TABLE_SCHEMA='oc' AND c.TABLE_NAME=t.TABLE_NAME AND c.COLUMN_NAME='Descrip')
AND TABLE_NAME LIKE '%Menu%'
"@
$r = $cmd.ExecuteReader(); Write-Output '---Menu tables---'; while($r.Read()){ Write-Output $r[0] }; $r.Close()

# Product might be in a view - try ProductPrice
$cmd.CommandText = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA='oc' AND TABLE_NAME='ProductPrice'"
$r = $cmd.ExecuteReader(); Write-Output '---ProductPrice---'; while($r.Read()){ Write-Output $r[0] }; $r.Close()

$cmd.CommandText = @"
SELECT pp.Product, pp.Price FROM oc.ProductPrice pp ORDER BY pp.Product
"@
$r = $cmd.ExecuteReader(); Write-Output '---prices---'; while($r.Read()){ Write-Output "$($r[0])=$($r[1])" }; $r.Close()

$conn.Close()
