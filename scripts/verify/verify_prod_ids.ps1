Add-Type -AssemblyName System.Data
$conn = New-Object System.Data.SqlClient.SqlConnection('Server=(localdb)\mssqllocaldb;Database=ocdata;Integrated Security=True')
$conn.Open(); $cmd = $conn.CreateCommand()
$cmd.CommandText = @"
SELECT i.ItemId, i.Descrip, i.Type, i.PosId
FROM oc.Item i WHERE i.ItemId BETWEEN 27 AND 33 ORDER BY i.ItemId
"@
$r = $cmd.ExecuteReader()
while($r.Read()){ Write-Output "$($r[0]) | POS $($r[3]) | $($r[1]) | type=$($r[2])" }
$r.Close()
$conn.Close()
