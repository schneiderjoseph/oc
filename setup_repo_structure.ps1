# Réorganise le dossier en structure repo (idempotent)
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

function Ensure-Dir($p) { if (-not (Test-Path $p)) { New-Item -ItemType Directory -Path $p | Out-Null } }

$dirs = @(
    "scripts\audit", "scripts\build", "scripts\verify", "scripts\tools",
    "docs\formation", "docs\la-reserve", "docs\schema",
    "transcriptions\clean", "transcriptions\raw", "livres", "output"
)
foreach ($d in $dirs) { Ensure-Dir (Join-Path $root $d) }

# Audit
@("oc_db_full_audit.py", "build_lareserve_audit_package.py", "requirements-oc-audit.txt") | ForEach-Object {
    $src = Join-Path $root $_
    if (Test-Path $src) { Move-Item $src (Join-Path $root "scripts\audit\$_") -Force }
}

# Build
@(
    "build_exercice.py", "build_parcours_lineaire.py", "rebuild_final.py",
    "exercice_locale.py", "generate_ventes_csv.py",
    "build_corrige.py", "build_corrigé.py", "calc_corrigé.py"
) | ForEach-Object {
    $src = Join-Path $root $_
    if (Test-Path $src) { Move-Item $src (Join-Path $root "scripts\build\$_") -Force }
}

# Verify
Get-ChildItem $root -File -Filter "verify_*" | ForEach-Object {
    Move-Item $_.FullName (Join-Path $root "scripts\verify\$($_.Name)") -Force
}
@("compare_exo_items.py", "verify_step29.py") | ForEach-Object {
    $src = Join-Path $root $_
    if (Test-Path $src) { Move-Item $src (Join-Path $root "scripts\verify\$_") -Force }
}

# Tools
Get-ChildItem $root -File -Filter "*.py" | ForEach-Object {
    Move-Item $_.FullName (Join-Path $root "scripts\tools\$($_.Name)") -Force
}

# Formation docs
@(
    "00_PLAYLIST_ORDRE.md", "video_book_mapping.json", "video_book_mapping.txt",
    "book_structure.txt", "book_structure_out.txt", "semantic_coverage_summary.txt",
    "semantic_coverage_report.json", "comparison_report_summary.txt", "comparison_report.json"
) | ForEach-Object {
    $src = Join-Path $root $_
    if (Test-Path $src) { Move-Item $src (Join-Path $root "docs\formation\$_") -Force }
}

# La Réserve
@("Mail_LaReserve_Optimum_Control.txt") | ForEach-Object {
    $src = Join-Path $root $_
    if (Test-Path $src) { Move-Item $src (Join-Path $root "docs\la-reserve\$_") -Force }
}

# Schema
@("_schema_dump.txt") | ForEach-Object {
    $src = Join-Path $root $_
    if (Test-Path $src) {
        $dst = Join-Path $root "docs\schema\oc_schema_tables.txt"
        if (Test-Path $dst) { Remove-Item $dst -Force }
        Move-Item $src $dst -Force
    }
}

# Transcriptions clean
Get-ChildItem $root -File -Filter "*_transcription_clean.txt" | ForEach-Object {
    Move-Item $_.FullName (Join-Path $root "transcriptions\clean\$($_.Name)") -Force
}

# Transcriptions raw (dossier OC)
$oc = Join-Path $root "OC"
if (Test-Path $oc) {
    Get-ChildItem $oc -File | ForEach-Object {
        Move-Item $_.FullName (Join-Path $root "transcriptions\raw\$($_.Name)") -Force
    }
    Remove-Item $oc -Force -ErrorAction SilentlyContinue
}

Write-Output "Structure repo OK : $root"
