Add-Type -AssemblyName System.Data
$out = "E:\OC DOCS\verify_sales_db.txt"
$lines = New-Object System.Collections.Generic.List[string]
$conn = New-Object System.Data.SqlClient.SqlConnection('Server=(localdb)\mssqllocaldb;Database=ocdata;Integrated Security=True;Connect Timeout=10')

function Add-RS($cmd, $title) {
    $lines.Add($title)
    $r = $cmd.ExecuteReader()
    $lines.Add((1..$r.FieldCount | ForEach-Object { $r.GetName($_ - 1) }) -join '|')
    while ($r.Read()) {
        $vals = 1..$r.FieldCount | ForEach-Object {
            $v = $r.GetValue($_ - 1)
            if ($v -is [DBNull]) { '' } elseif ($v -is [DateTime]) { $v.ToString('yyyy-MM-dd HH:mm') } else { $v }
        }
        $lines.Add(($vals -join '|'))
    }
    $r.Close()
}

try {
    $conn.Open()
    $cmd = $conn.CreateCommand()

    # Tables ventes
    $cmd.CommandText = @"
SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES
WHERE TABLE_SCHEMA = 'oc' AND (TABLE_NAME LIKE '%Sales%' OR TABLE_NAME LIKE '%Till%' OR TABLE_NAME LIKE '%Retail%')
ORDER BY TABLE_NAME
"@
    Add-RS $cmd 'SALES_TABLES'

    $cmd.CommandText = @"
SELECT TABLE_NAME, COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_SCHEMA = 'oc' AND TABLE_NAME LIKE '%Sales%'
ORDER BY TABLE_NAME, ORDINAL_POSITION
"@
    Add-RS $cmd 'SALES_COLUMNS'

    $conn.Close()
    $lines | Set-Content $out -Encoding UTF8
    Write-Output "OK -> $out"
} catch {
    "ERR: $($_.Exception.Message)" | Set-Content $out -Encoding UTF8
    Write-Output "ERR $($_.Exception.Message)"
    exit 1
}
