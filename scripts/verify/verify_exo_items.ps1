Add-Type -AssemblyName System.Data
$out = "E:\OC DOCS\verify_items_db.txt"
$conn = New-Object System.Data.SqlClient.SqlConnection('Server=(localdb)\mssqllocaldb;Database=ocdata;Integrated Security=True;Connect Timeout=10')
$lines = New-Object System.Collections.Generic.List[string]
try {
    $conn.Open()
    $cmd = $conn.CreateCommand()

    $cmd.CommandText = "SELECT Type, COUNT(*) AS Cnt FROM oc.Item GROUP BY Type ORDER BY Type"
    $r = $cmd.ExecuteReader()
    $lines.Add('ITEM_TYPES')
    while ($r.Read()) { $lines.Add("$($r.GetValue(0))| $($r.GetValue(1))") }
    $r.Close()

    $cmd.CommandText = @"
SELECT i.ItemId, i.Type, i.Descrip,
    (SELECT COUNT(*) FROM oc.CaseSize cs WHERE cs.Item = i.ItemId) AS CaseSizeCount,
    (SELECT COUNT(*) FROM oc.CaseSize cs
        JOIN oc.CaseSizeCost csc ON csc.CaseSize = cs.CaseSizeId AND csc.IsDeleted = 0
        WHERE cs.Item = i.ItemId) AS ActiveCaseCosts
FROM oc.Item i
WHERE EXISTS (SELECT 1 FROM oc.CaseSize cs WHERE cs.Item = i.ItemId)
ORDER BY i.Descrip
"@
    $r = $cmd.ExecuteReader()
    $lines.Add('ITEMS_WITH_CASESIZE')
    $hdr = for ($i=0; $i -lt $r.FieldCount; $i++) { $r.GetName($i) }
    $lines.Add(($hdr -join '|'))
    while ($r.Read()) {
        $vals = for ($i=0; $i -lt $r.FieldCount; $i++) { $r.GetValue($i) }
        $lines.Add(($vals -join '|'))
    }
    $r.Close()

    $cmd.CommandText = @"
SELECT COUNT(DISTINCT i.ItemId) AS ItemCount, COUNT(cs.CaseSizeId) AS CaseSizeCount
FROM oc.Item i
JOIN oc.CaseSize cs ON cs.Item = i.ItemId
"@
    $r = $cmd.ExecuteReader()
    $lines.Add('TOTALS')
    if ($r.Read()) { $lines.Add("items|$($r.GetValue(0))|case_sizes|$($r.GetValue(1))") }
    $r.Close()

    $conn.Close()
    $lines | Set-Content -Path $out -Encoding UTF8
    Write-Output "OK -> $out"
} catch {
    "ERR: $($_.Exception.Message)" | Set-Content -Path $out -Encoding UTF8
    Write-Output "ERR $($_.Exception.Message)"
    exit 1
}
