Add-Type -AssemblyName System.Data
$out = "E:\OC DOCS\verify_preps_db.txt"
$conn = New-Object System.Data.SqlClient.SqlConnection('Server=(localdb)\mssqllocaldb;Database=ocdata;Integrated Security=True;Connect Timeout=10')
$lines = New-Object System.Collections.Generic.List[string]
try {
    $conn.Open()
    $cmd = $conn.CreateCommand()

    $cmd.CommandText = "SELECT TypeId, TypeName FROM oc.TypeMap ORDER BY TypeId"
    $r = $cmd.ExecuteReader()
    $lines.Add('TYPEMAP')
    while ($r.Read()) { $lines.Add("$($r.GetValue(0))|$($r.GetValue(1))") }
    $r.Close()

    $cmd.CommandText = @"
SELECT i.ItemId, i.Type, i.Descrip,
    ru.Uom AS RecipeUom,
    bu.Uom AS BatchUom,
    py.CountQty, cu.Uom AS CountUom,
    py.RecipeQty,
    pd.ShowOnPrepSheets,
    ps.Description AS PrepStation,
    pd.ShelfLifeMinutes,
    id.ActualizeUsage,
    (SELECT COUNT(*) FROM oc.Ingredient ing WHERE ing.Recipe = i.ItemId) AS IngredCount
FROM oc.Item i
JOIN oc.PrepYield py ON py.Item = i.ItemId
JOIN oc.Uom ru ON ru.UomId = i.RecipeUom
JOIN oc.Uom bu ON bu.UomId = py.BatchUom
JOIN oc.Uom cu ON cu.UomId = py.CountUom
LEFT JOIN oc.PrepDetail pd ON pd.Item = i.ItemId
LEFT JOIN oc.PrepStation ps ON ps.PrepStationId = pd.PrepStation
LEFT JOIN oc.ItemDetail id ON id.Item = i.ItemId
ORDER BY i.Descrip
"@
    $r = $cmd.ExecuteReader()
    $lines.Add('PREPS')
    $hdr = for ($i=0; $i -lt $r.FieldCount; $i++) { $r.GetName($i) }
    $lines.Add(($hdr -join '|'))
    while ($r.Read()) {
        $vals = for ($i=0; $i -lt $r.FieldCount; $i++) {
            $v = $r.GetValue($i)
            if ($v -is [DBNull]) { '' } else { $v }
        }
        $lines.Add(($vals -join '|'))
    }
    $r.Close()

    $cmd.CommandText = @"
SELECT r.Descrip AS Prep, r.ItemId AS PrepId,
    ing.Idx,
    i.Descrip AS Ingredient,
    i.Type AS IngType,
    ing.Qty,
    u.Uom AS Uom
FROM oc.Ingredient ing
JOIN oc.Item r ON ing.Recipe = r.ItemId
JOIN oc.Item i ON ing.Item = i.ItemId
JOIN oc.Uom u ON u.UomId = ing.Uom
WHERE EXISTS (SELECT 1 FROM oc.PrepYield py WHERE py.Item = r.ItemId)
ORDER BY r.Descrip, ing.Idx
"@
    $r = $cmd.ExecuteReader()
    $lines.Add('INGREDIENTS')
    $hdr = for ($i=0; $i -lt $r.FieldCount; $i++) { $r.GetName($i) }
    $lines.Add(($hdr -join '|'))
    while ($r.Read()) {
        $vals = for ($i=0; $i -lt $r.FieldCount; $i++) {
            $v = $r.GetValue($i)
            if ($v -is [DBNull]) { '' } else { $v }
        }
        $lines.Add(($vals -join '|'))
    }
    $r.Close()

    $cmd.CommandText = "SELECT COUNT(*) FROM oc.Item i JOIN oc.PrepYield py ON py.Item = i.ItemId"
    $prepCount = $cmd.ExecuteScalar()
    $lines.Add("PREP_COUNT|$prepCount")

    $cmd.CommandText = @"
SELECT i.ItemId, i.Descrip, id.ActualizeUsage, id.TrackInventory
FROM oc.Item i
LEFT JOIN oc.ItemDetail id ON id.Item = i.ItemId
WHERE i.ItemId IN (22, 23, 24, 25, 26)
ORDER BY i.ItemId
"@
    $r = $cmd.ExecuteReader()
    $lines.Add('PREP_FLAGS')
    while ($r.Read()) {
        $lines.Add("$($r.GetValue(0))|$($r.GetValue(1))|$($r.GetValue(2))|$($r.GetValue(3))")
    }
    $r.Close()

    $conn.Close()
    $lines | Set-Content -Path $out -Encoding UTF8
    Write-Output "OK -> $out ($prepCount preps)"
} catch {
    "ERR: $($_.Exception.Message)" | Set-Content -Path $out -Encoding UTF8
    Write-Output "ERR $($_.Exception.Message)"
    exit 1
}
