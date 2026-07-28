Add-Type -AssemblyName System.Data
$conn = New-Object System.Data.SqlClient.SqlConnection('Server=(localdb)\mssqllocaldb;Database=ocdata;Integrated Security=True')
$conn.Open()
$cmd = $conn.CreateCommand()

$queries = @{
    'BOEUF_ITEMS' = "SELECT ItemId, Descrip FROM oc.Item WHERE Descrip LIKE N'%B%uf%' ORDER BY Descrip"
    'DC0318_ITEMIDS' = @"
SELECT inv.InvoiceNumber, i.ItemId, i.Descrip, ii.Qty, q.QtyOnHand, ru.Uom
FROM oc.InvoiceItem ii
JOIN oc.Invoice inv ON inv.InvoiceId = ii.Invoice
JOIN oc.Item i ON i.ItemId = ii.Item
LEFT JOIN oc.ItemQtyOnHand q ON q.Item = i.ItemId
LEFT JOIN oc.Uom ru ON ru.UomId = i.RecipeUom
WHERE inv.InvoiceNumber = 'DC-2026-0318'
"@
    'COLA_QOH' = @"
SELECT i.Descrip, q.QtyOnHand, ic.PurchaseQty AS OpeningCs
FROM oc.Item i
LEFT JOIN oc.ItemQtyOnHand q ON q.Item = i.ItemId
LEFT JOIN oc.InventoryCount ic ON ic.Item = i.ItemId AND ic.Countsheet = (SELECT MAX(InventoryId) FROM oc.Inventory)
WHERE i.Descrip = N'Cola 355 ml'
"@
    'CASE_COLA' = @"
SELECT i.Descrip, cs.Qty, u.Uom
FROM oc.ItemCaseSize cs
JOIN oc.Item i ON i.ItemId = cs.Item
JOIN oc.Uom u ON u.UomId = cs.Uom
WHERE i.Descrip IN (N'Cola 355 ml', N'Pain burger', N'Bœuf haché 80/20', N'Bœuf', N'Pommes de terre')
"@
}

foreach ($k in $queries.Keys) {
    Write-Output "=== $k ==="
    $cmd.CommandText = $queries[$k]
    $r = $cmd.ExecuteReader()
    $cols = 1..$r.FieldCount | ForEach-Object { $r.GetName($_ - 1) }
    Write-Output ($cols -join '|')
    while ($r.Read()) {
        $vals = 1..$r.FieldCount | ForEach-Object {
            $v = $r.GetValue($_ - 1)
            if ($v -is [DBNull]) { '' } else { $v }
        }
        Write-Output ($vals -join '|')
    }
    $r.Close()
}
$conn.Close()
