Add-Type -AssemblyName System.Data
$conn = New-Object System.Data.SqlClient.SqlConnection('Server=(localdb)\mssqllocaldb;Database=ocdata;Integrated Security=True')
$conn.Open()
$cmd = $conn.CreateCommand()
$tables = @('TillTape','TillTapeProduct','TillTapeSale','RetailSale','RetailSaleItem','SalesSource','Product')
foreach ($t in $tables) {
    Write-Output "=== oc.$t ==="
    $cmd.CommandText = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA='oc' AND TABLE_NAME='$t' ORDER BY ORDINAL_POSITION"
    $r = $cmd.ExecuteReader()
    while ($r.Read()) { Write-Output $r[0] }
    $r.Close()
}
$conn.Close()
