Add-Type -AssemblyName System.Data
$conn = New-Object System.Data.SqlClient.SqlConnection('Server=(localdb)\mssqllocaldb;Database=ocdata;Integrated Security=True')
$conn.Open()
$cmd = $conn.CreateCommand()
$cmd.CommandText = @"
SELECT i.Descrip, cs.CaseDescrip, cs.PurchasePrice, cs.SplitQty, su.Uom AS SplitUom, pu.Uom AS PurchaseUom
FROM oc.CaseSize cs
JOIN oc.Item i ON i.ItemId = cs.Item
LEFT JOIN oc.Uom su ON su.UomId = cs.SplitUom
LEFT JOIN oc.Uom pu ON pu.UomId = cs.PurchaseUom
WHERE i.ItemId IN (2, 3, 5, 12, 16)
ORDER BY i.Descrip
"@
$r = $cmd.ExecuteReader()
Write-Output '=== CASE SIZES ==='
while ($r.Read()) {
    Write-Output "$($r[0]) | split=$($r[3]) $($r[4]) | purch_uom=$($r[5]) | price=$($r[2])"
}
$r.Close()
$conn.Close()
