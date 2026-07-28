Add-Type -AssemblyName System.Data
$conn = New-Object System.Data.SqlClient.SqlConnection('Server=(localdb)\mssqllocaldb;Database=ocdata;Integrated Security=True')
$conn.Open(); $cmd = $conn.CreateCommand()

foreach ($t in @('ProductSale','Recipe','MenuProduct')) {
    Write-Output "=== $t cols ==="
    $cmd.CommandText = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA='oc' AND TABLE_NAME='$t' ORDER BY ORDINAL_POSITION"
    $r = $cmd.ExecuteReader(); while($r.Read()){ Write-Output $r[0] }; $r.Close()
}

$cmd.CommandText = "SELECT * FROM oc.ProductSale"
$r = $cmd.ExecuteReader()
Write-Output '---ProductSale---'
$cols = 1..$r.FieldCount | ForEach-Object { $r.GetName($_-1) }
Write-Output ($cols -join '|')
while($r.Read()){
    $vals = 1..$r.FieldCount | ForEach-Object { $v=$r.GetValue($_-1); if($v -is [DBNull]){''}elseif($v -is [DateTime]){$v.ToString('yyyy-MM-dd')}else{$v} }
    Write-Output ($vals -join '|')
}
$r.Close()

$cmd.CommandText = @"
SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA='oc'
AND (TABLE_NAME LIKE '%Recipe%' OR TABLE_NAME LIKE '%Menu%' OR TABLE_NAME LIKE '%Plu%')
ORDER BY TABLE_NAME
"@
$r = $cmd.ExecuteReader(); Write-Output '---RECIPE TABLES---'; while($r.Read()){ Write-Output $r[0] }; $r.Close()

$cmd.CommandText = "SELECT COUNT(*) FROM oc.Recipe WHERE Type='M'"
Write-Output "Menu recipes=$($cmd.ExecuteScalar())"

$cmd.CommandText = @"
SELECT TOP 20 r.RecipeId, r.Descrip, r.PosId, r.Type
FROM oc.Recipe r WHERE r.Type = 'M' ORDER BY r.PosId
"@
$r = $cmd.ExecuteReader(); Write-Output '---PRODUCTS---'; while($r.Read()){ Write-Output "$($r[2]) | $($r[1]) | type=$($r[3])" }; $r.Close()

$conn.Close()
