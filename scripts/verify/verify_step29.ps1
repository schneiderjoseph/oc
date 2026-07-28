Add-Type -AssemblyName System.Data
$out = "E:\OC DOCS\verify_step29.txt"
$lines = New-Object System.Collections.Generic.List[string]
$conn = New-Object System.Data.SqlClient.SqlConnection('Server=(localdb)\mssqllocaldb;Database=ocdata;Integrated Security=True;Connect Timeout=10')

function Add-RS($cmd, $title) {
    $lines.Add($title)
    $r = $cmd.ExecuteReader()
    $lines.Add((1..$r.FieldCount | ForEach-Object { $r.GetName($_ - 1) }) -join '|')
    while ($r.Read()) {
        $vals = 1..$r.FieldCount | ForEach-Object {
            $v = $r.GetValue($_ - 1)
            if ($v -is [DBNull]) { '' } elseif ($v -is [DateTime]) { $v.ToString('dd/MM/yyyy') } else { $v }
        }
        $lines.Add(($vals -join '|'))
    }
    $r.Close()
}

$expectedInvoices = @(
    @{ Num = 'BP-2026-0142'; Date = '30/06/2026'; TTC = 5.28 }
    @{ Num = 'DC-2026-0318'; Date = '30/06/2026'; TTC = 191.40 }
    @{ Num = 'DC-2026-0320'; Date = '02/07/2026'; TTC = 180.95 }
    @{ Num = 'BN-2026-0088'; Date = '03/07/2026'; TTC = 52.80 }
    @{ Num = 'EH-2026-0205'; Date = '04/07/2026'; TTC = 88.00 }
)

try {
    $conn.Open()
    $cmd = $conn.CreateCommand()

    $cmd.CommandText = @"
SELECT inv.InvoiceNumber, CONVERT(varchar(10), inv.InvoiceDate, 103) AS InvDate,
    s.Name AS Supplier, inv.Total, inv.ItemTotal, inv.AdjustmentTotal, inv.ExpenseTotal
FROM oc.Invoice inv
LEFT JOIN oc.Supplier s ON s.SupplierId = inv.Supplier
ORDER BY inv.InvoiceId
"@
    Add-RS $cmd 'INVOICES'

    $cmd.CommandText = @"
SELECT inv.InvoiceNumber, i.Descrip, ii.Qty, u.Uom, ii.UnitCost, ii.LineTotal
FROM oc.InvoiceItem ii
JOIN oc.Invoice inv ON inv.InvoiceId = ii.Invoice
JOIN oc.Item i ON i.ItemId = ii.Item
LEFT JOIN oc.Uom u ON u.UomId = ii.Uom
ORDER BY inv.InvoiceId, ii.Idx
"@
    Add-RS $cmd 'LINES'

    $cmd.CommandText = @"
SELECT i.Descrip, q.QtyOnHand, ru.Uom
FROM oc.ItemQtyOnHand q
JOIN oc.Item i ON i.ItemId = q.Item
LEFT JOIN oc.Uom ru ON ru.UomId = i.RecipeUom
WHERE i.Type = 'I'
ORDER BY i.Descrip
"@
    Add-RS $cmd 'ALL_ITEM_QOH'

    $cmd.CommandText = 'SELECT MAX(InventoryId) FROM oc.Inventory'
    $invId = [int]$cmd.ExecuteScalar()
    $cmd.CommandText = @"
SELECT i.Descrip, ic.PurchaseQty, ic.CaseQty, q.QtyOnHand AS LiveQty, ru.Uom
FROM oc.InventoryCount ic
JOIN oc.Item i ON i.ItemId = ic.Item
LEFT JOIN oc.ItemQtyOnHand q ON q.Item = ic.Item
LEFT JOIN oc.Uom ru ON ru.UomId = i.RecipeUom
WHERE ic.Countsheet = $invId AND i.Descrip IN (N'Pain burger', N'Pommes de terre', N'Bœuf haché 80/20', N'Laitue romaine', N'Oignons')
ORDER BY i.Descrip
"@
    Add-RS $cmd 'OPENING_VS_LIVE'

    $conn.Close()

    $text = $lines -join "`n"
    $lines.Add('CHECKLIST')
    foreach ($e in $expectedInvoices) {
        if ($text -match [regex]::Escape($e.Num)) {
            $lines.Add("FOUND|$($e.Num)")
        } else {
            $lines.Add("MISSING|$($e.Num)")
        }
    }

    if ($text -match 'Pain burger\|48') { $lines.Add('OK|Pain QOH 48 ea (24+24)') }
    elseif ($text -match 'Pain burger\|24') { $lines.Add('WARN|Pain QOH 24 — facture lundi pas comptée?') }
    else { $lines.Add('CHECK|Pain QOH — voir ALL_ITEM_QOH') }

    $lines | Set-Content $out -Encoding UTF8
    Write-Output "OK -> $out"
} catch {
    "ERR: $($_.Exception.Message)" | Set-Content $out -Encoding UTF8
    Write-Output "ERR $($_.Exception.Message)"
    exit 1
}
