Add-Type -AssemblyName System.Data
$out = "E:\OC DOCS\verify_inventory_db.txt"
$conn = New-Object System.Data.SqlClient.SqlConnection('Server=(localdb)\mssqllocaldb;Database=ocdata;Integrated Security=True;Connect Timeout=10')
$lines = New-Object System.Collections.Generic.List[string]
function Add-RS($cmd, $title) {
    $lines.Add($title)
    $r = $cmd.ExecuteReader()
    $lines.Add((1..$r.FieldCount | ForEach-Object { $r.GetName($_ - 1) }) -join '|')
    while ($r.Read()) {
        $vals = 1..$r.FieldCount | ForEach-Object {
            $v = $r.GetValue($_ - 1)
            if ($v -is [DBNull]) { '' } else { $v }
        }
        $lines.Add(($vals -join '|'))
    }
    $r.Close()
}
try {
    $conn.Open(); $cmd = $conn.CreateCommand()
    $cmd.CommandText = 'SELECT MAX(InventoryId) FROM oc.Inventory'
    $invId = [int]$cmd.ExecuteScalar()

    $cmd.CommandText = @"
SELECT i.Descrip, ic.PurchaseQty, ic.CaseQty, ic.PakQty,
    s.QtyOnHand AS SummaryQty, q.QtyOnHand AS LiveQty, ru.Uom
FROM oc.InventoryCount ic
JOIN oc.Item i ON i.ItemId = ic.Item
LEFT JOIN oc.InventorySummary s ON s.Inventory = ic.Countsheet AND s.Item = ic.Item
LEFT JOIN oc.ItemQtyOnHand q ON q.Item = ic.Item
LEFT JOIN oc.Uom ru ON ru.UomId = i.RecipeUom
WHERE ic.Countsheet = $invId AND i.Type = 'P'
ORDER BY i.Descrip
"@
    Add-RS $cmd 'PREPS_COUNT_SUMMARY_LIVE'

    $cmd.CommandText = @"
SELECT i.Descrip, s.QtyOnHand, s.TotalValue, s.WasCounted
FROM oc.InventorySummary s
JOIN oc.Item i ON i.ItemId = s.Item
WHERE s.Inventory = $invId AND i.Type = 'P'
"@
    Add-RS $cmd 'PREPS_IN_SUMMARY_ONLY'

    $conn.Close()
    $lines | Set-Content $out -Encoding UTF8
    Write-Output "OK"
} catch { Write-Output "ERR $($_.Exception.Message)"; exit 1 }
