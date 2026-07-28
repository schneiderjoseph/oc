Add-Type -AssemblyName System.Data
$conn = New-Object System.Data.SqlClient.SqlConnection('Server=(localdb)\mssqllocaldb;Database=ocdata;Integrated Security=True')
$conn.Open()
$cmd = $conn.CreateCommand()
$cmd.CommandText = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA='oc' AND TABLE_NAME LIKE '%Case%' ORDER BY TABLE_NAME"
$r = $cmd.ExecuteReader()
Write-Output '=== CASE TABLES ==='
while ($r.Read()) { Write-Output $r[0] }
$r.Close()

$cmd.CommandText = @"
SELECT i.ItemId, i.Descrip, ic.PurchaseQty, ic.CaseQty, q.QtyOnHand, ru.Uom
FROM oc.InventoryCount ic
JOIN oc.Item i ON i.ItemId = ic.Item
LEFT JOIN oc.ItemQtyOnHand q ON q.Item = ic.Item
LEFT JOIN oc.Uom ru ON ru.UomId = i.RecipeUom
WHERE ic.Countsheet = (SELECT MAX(InventoryId) FROM oc.Inventory)
  AND i.ItemId IN (2, 3, 5, 12)
ORDER BY i.ItemId
"@
$r = $cmd.ExecuteReader()
Write-Output '=== OPENING ITEMS 2,3,5,12 ==='
while ($r.Read()) {
    Write-Output "$($r[1]) | open_cs=$($r[2]) | case_col=$($r[3]) | qoh=$($r[4]) $($r[5])"
}
$r.Close()
$conn.Close()
