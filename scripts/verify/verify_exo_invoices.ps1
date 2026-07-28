Add-Type -AssemblyName System.Data
$out = "E:\OC DOCS\verify_invoices_db.txt"
$conn = New-Object System.Data.SqlClient.SqlConnection('Server=(localdb)\mssqllocaldb;Database=ocdata;Integrated Security=True;Connect Timeout=10')
$lines = New-Object System.Collections.Generic.List[string]

function Add-RS($cmd, $title) {
    $lines.Add($title)
    $r = $cmd.ExecuteReader()
    $lines.Add((1..$r.FieldCount | ForEach-Object { $r.GetName($_ - 1) }) -join '|')
    while ($r.Read()) {
        $vals = 1..$r.FieldCount | ForEach-Object {
            $v = $r.GetValue($_ - 1)
            if ($v -is [DBNull]) { '' } elseif ($v -is [DateTime]) { $v.ToString('yyyy-MM-dd') } else { $v }
        }
        $lines.Add(($vals -join '|'))
    }
    $r.Close()
}

# Expected from exercice (first 2 invoices)
$expected = @{
    'BP-2026-0142' = @{
        Supplier = 'Boulangerie'
        Lines = @(@('Pain burger', 2, 2.40, 4.80))
        HT = 4.80; Tax = 0.48; TTC = 5.28
    }
    'DC-2026-0318' = @{
        Supplier = 'Distrib'
        Lines = @(
            @('Bœuf', 2, 42.00, 84.00),
            @('Pommes', 2, 22.00, 44.00),
            @('Laitue', 1, 28.00, 28.00),
            @('Oignons', 1, 18.00, 18.00)
        )
        HT = 174.00; Tax = 17.40; TTC = 191.40
    }
}

try {
    $conn.Open()
    $cmd = $conn.CreateCommand()

    $cmd.CommandText = @"
SELECT TOP 5 inv.InvoiceId, inv.InvoiceNumber, inv.InvoiceDate, inv.Total, inv.ItemTotal,
    inv.ExpenseTotal, inv.AdjustmentTotal, s.Name AS Supplier
FROM oc.Invoice inv
LEFT JOIN oc.Supplier s ON s.SupplierId = inv.Supplier
ORDER BY inv.InvoiceId
"@
    Add-RS $cmd 'INVOICE_HEADERS'

    $cmd.CommandText = @"
SELECT inv.InvoiceNumber, s.Name AS Supplier, i.Descrip AS Item,
    ii.Qty, u.Uom, ii.UnitCost, ii.LineTotal, ii.SplitCase, ii.CostOverridden
FROM oc.InvoiceItem ii
JOIN oc.Invoice inv ON inv.InvoiceId = ii.Invoice
JOIN oc.Supplier s ON s.SupplierId = inv.Supplier
JOIN oc.Item i ON i.ItemId = ii.Item
LEFT JOIN oc.Uom u ON u.UomId = ii.Uom
ORDER BY inv.InvoiceId, ii.Idx
"@
    Add-RS $cmd 'INVOICE_LINES'

    $cmd.CommandText = @"
SELECT inv.InvoiceNumber, t.Code AS TaxCode, it.Amount
FROM oc.InvoiceTax it
JOIN oc.Invoice inv ON inv.InvoiceId = it.Invoice
LEFT JOIN oc.Tax t ON t.TaxId = it.Tax
ORDER BY inv.InvoiceId
"@
    Add-RS $cmd 'INVOICE_TAXES'

    # Qty on hand after purchases - key items
    $cmd.CommandText = @"
SELECT i.Descrip, q.QtyOnHand, ru.Uom
FROM oc.ItemQtyOnHand q
JOIN oc.Item i ON i.ItemId = q.Item
LEFT JOIN oc.Uom ru ON ru.UomId = i.RecipeUom
WHERE i.Descrip LIKE N'%Pain%' OR i.Descrip LIKE N'%Bœuf%' OR i.Descrip LIKE N'%Pommes%'
   OR i.Descrip LIKE N'%Laitue%' OR i.Descrip LIKE N'%Oignon%'
ORDER BY i.Descrip
"@
    Add-RS $cmd 'QOH_AFTER_INVOICES'

    $conn.Close()

    # Simple verdict
    $lines.Add('VERDICT')
    $content = $lines -join "`n"
    if ($content -match 'BP-2026-0142') { $lines.Add('FOUND|BP-2026-0142') } else { $lines.Add('MISSING|BP-2026-0142') }
    if ($content -match 'DC-2026-0318') { $lines.Add('FOUND|DC-2026-0318') } else { $lines.Add('MISSING|DC-2026-0318|check invoice number entered') }

    $lines | Set-Content $out -Encoding UTF8
    Write-Output "OK -> $out"
} catch {
    "ERR: $($_.Exception.Message)" | Set-Content $out -Encoding UTF8
    Write-Output "ERR $($_.Exception.Message)"
    exit 1
}
