<#
.SYNOPSIS
    Обновление региональных ипотечных таблиц по новой выгрузке — версия
    для Windows без установки чего-либо (PowerShell встроен в Windows,
    работает через COM-автоматизацию уже установленного Excel).

.DESCRIPTION
    Логика 1:1 повторяет update_regional_tables.py (там же и проверена
    самопроверкой + синтетическим тестом) — см. комментарии в том файле
    для деталей структуры листов. Здесь та же логика реализована через
    Excel COM, чтобы не требовалось ставить Python/openpyxl.

    Берёт свежескачанную "Статистические ряды_регионы.xlsx" и переносит
    новые месяцы/годы в мастер-файл "Статистические ряды_регионы (с
    таблицами).xlsx" (листы кол-во/объем/ДДУ), не трогая форматирование —
    вставка колонок делается через настоящий Excel (.Insert), поэтому
    формулы (A/B/D на листе ДДУ) он сам корректно пересчитывает адреса,
    и формат соседней колонки копируется автоматически.

    По умолчанию — dry-run (только отчёт в консоль). Пишет в файл только
    с флагом -Apply.

.PARAMETER RawPath
    Путь к новой выгрузке "Статистические ряды_регионы.xlsx".

.PARAMETER MasterPath
    Путь к мастер-файлу. По умолчанию — файл в той же папке, где лежит
    сам скрипт.

.PARAMETER Apply
    Записать изменения. Без этого флага — только отчёт, файл не трогается.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File update_regional_tables.ps1 -RawPath "C:\Downloads\Статистические ряды_регионы.xlsx"

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File update_regional_tables.ps1 -RawPath "C:\Downloads\...xlsx" -Apply
#>
param(
    [Parameter(Mandatory = $true)][string]$RawPath,
    [string]$MasterPath = (Join-Path $PSScriptRoot "Статистические ряды_регионы (с таблицами).xlsx"),
    [switch]$Apply
)

$ErrorActionPreference = "Stop"

$HEADER_ROW = 4
$FIRST_DATA_ROW = 5
$EXCLUDED_PROGRAMS = @("Все программы", "Ипотека в отдельных регионах")
$DDU_LABEL_REGEX = '^Покупка по ДДУ,\s*(шт\.|млн руб\.)$'

function Get-ColumnBlock($ws, $col, $firstRow, $lastRow) {
    # Один bulk-читает столбец в 2D-массив (COM-диапазоны в PowerShell
    # 1-индексные), возвращает обычный PS-массив object[] (индекс с 0).
    if ($lastRow -lt $firstRow) { return @() }
    $range = $ws.Range($ws.Cells($firstRow, $col), $ws.Cells($lastRow, $col))
    $raw = $range.Value2
    $n = $lastRow - $firstRow + 1
    if ($n -eq 1) { return @($raw) }
    $out = New-Object object[] $n
    for ($i = 0; $i -lt $n; $i++) { $out[$i] = $raw[$i + 1, 1] }
    return $out
}

function Get-UsedLastRow($ws) {
    $u = $ws.UsedRange
    return $u.Row + $u.Rows.Count - 1
}

function ColKey($val) {
    if ($val -is [double] -or $val -is [int]) {
        $d = [datetime]::FromOADate([double]$val)
        return "date:$($d.Year)-$($d.Month)"
    }
    return "text:$val"
}

function FmtKey($val) {
    if ($val -is [double] -or $val -is [int]) {
        $d = [datetime]::FromOADate([double]$val)
        return $d.ToString("yyyy-MM")
    }
    return "$val"
}

function Get-Rows {
    <#
    Универсальный сбор строк с листа: заголовки регионов/программ
    (пустая ячейка в data-колонке) плюс отфильтрованные строки данных.

    -DeriveDimFromLabel: тип строки (Dim) берётся из самой метки через
    regex (лист ДДУ), а не из отдельной колонки-размерности.
    #>
    param(
        $ws, [int]$labelCol, [int]$dimCol, [int]$dataStartCol,
        [string]$targetDim = $null, [string[]]$excludeSet = $null,
        [switch]$DeriveDimFromLabel, [string]$labelRegex = $null
    )
    $lastRow = Get-UsedLastRow $ws
    $labels = Get-ColumnBlock $ws $labelCol $FIRST_DATA_ROW $lastRow
    $datas = Get-ColumnBlock $ws $dataStartCol $FIRST_DATA_ROW $lastRow
    $dims = $null
    if (-not $DeriveDimFromLabel) { $dims = Get-ColumnBlock $ws $dimCol $FIRST_DATA_ROW $lastRow }

    $rows = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $labels.Length; $i++) {
        $label = $labels[$i]
        if ($null -eq $label) { continue }
        $hasData = $null -ne $datas[$i]
        if (-not $hasData) {
            $rows.Add([PSCustomObject]@{ Row = $FIRST_DATA_ROW + $i; Label = $label; Dim = $null })
            continue
        }
        if ($DeriveDimFromLabel) {
            $trimmed = $label.Trim()
            if ($trimmed -notmatch $labelRegex) { continue }
            $rows.Add([PSCustomObject]@{ Row = $FIRST_DATA_ROW + $i; Label = $label; Dim = $trimmed })
        } else {
            $dim = $dims[$i]
            if ($null -eq $dim) { continue }
            if ($excludeSet -and ($excludeSet -contains $label.Trim())) { continue }
            if ($targetDim -and ($dim -ne $targetDim)) { continue }
            $rows.Add([PSCustomObject]@{ Row = $FIRST_DATA_ROW + $i; Label = $label; Dim = $dim })
        }
    }
    return $rows
}

function Check-RowAlignment {
    param([string]$Name, $MasterRows, $RawRows)
    $n = [Math]::Min($MasterRows.Count, $RawRows.Count)
    for ($i = 0; $i -lt $n; $i++) {
        $m = $MasterRows[$i]; $r = $RawRows[$i]
        if ($m.Label -ne $r.Label -or $m.Dim -ne $r.Dim) {
            throw "[$Name] Строки не совпадают на позиции $i (мастер row=$($m.Row)): " +
                  "в мастере ('$($m.Label)','$($m.Dim)'), в новой выгрузке ('$($r.Label)','$($r.Dim)'). " +
                  "Похоже на новый/пропавший регион или программу - разберись руками."
        }
    }
    if ($RawRows.Count -gt $MasterRows.Count) {
        $extra = $RawRows.Count - $MasterRows.Count
        Write-Host "[$Name] В новой выгрузке $extra новых строк в конце (новый регион/программа?), не добавлены автоматически."
    } elseif ($RawRows.Count -lt $MasterRows.Count) {
        Write-Host "[$Name] ВНИМАНИЕ: в новой выгрузке меньше строк, чем в мастере - проверь, тот ли файл дали."
    }
    return $RawRows[0..($n - 1)]
}

function Sync-Sheet {
    param(
        [string]$Name, $RawWb, $MasterWb,
        [string]$RawSheet, [int]$RawLabelCol, [int]$RawDimCol, [int]$RawDataStartCol,
        [string]$MasterSheet, [int]$MasterLabelCol, [int]$MasterDimCol, [int]$MasterDataStartCol,
        [string]$TargetDim, [string[]]$ExcludeSet,
        [bool]$IsDdu, [bool]$Apply, $RevisionsOut, [int]$AutoFilterStartCol
    )
    $rawWs = $RawWb.Sheets.Item($RawSheet)
    $masterWs = $MasterWb.Sheets.Item($MasterSheet)

    if ($IsDdu) {
        $rawRows = Get-Rows -ws $rawWs -labelCol $RawLabelCol -dimCol 0 -dataStartCol $RawDataStartCol -DeriveDimFromLabel -labelRegex $DDU_LABEL_REGEX
        $masterSig = Get-Rows -ws $masterWs -labelCol $MasterLabelCol -dimCol 0 -dataStartCol $MasterDataStartCol -DeriveDimFromLabel -labelRegex $DDU_LABEL_REGEX
    } else {
        $rawRows = Get-Rows -ws $rawWs -labelCol $RawLabelCol -dimCol $RawDimCol -dataStartCol $RawDataStartCol -targetDim $TargetDim -excludeSet $ExcludeSet
        $masterSig = Get-Rows -ws $masterWs -labelCol $MasterLabelCol -dimCol $MasterDimCol -dataStartCol $MasterDataStartCol
    }

    $alignedRawRows = Check-RowAlignment -Name $Name -MasterRows $masterSig -RawRows $rawRows

    $masterLastRow = Get-UsedLastRow $masterWs
    $masterLastCol = $masterWs.UsedRange.Column + $masterWs.UsedRange.Columns.Count - 1
    $masterKeys = New-Object System.Collections.Generic.List[string]
    for ($c = $MasterDataStartCol; $c -le $masterLastCol; $c++) {
        $v = $masterWs.Cells($HEADER_ROW, $c).Value2
        if ($null -eq $v) { continue }
        $masterKeys.Add((ColKey $v))
    }

    $rawLastCol = $rawWs.UsedRange.Column + $rawWs.UsedRange.Columns.Count - 1
    $newColsAdded = New-Object System.Collections.Generic.List[object]
    $pointer = 0

    for ($rc = $RawDataStartCol; $rc -le $rawLastCol; $rc++) {
        $rv = $rawWs.Cells($HEADER_ROW, $rc).Value2
        if ($null -eq $rv) { continue }
        $key = ColKey $rv

        if ($pointer -lt $masterKeys.Count -and $masterKeys[$pointer] -eq $key) {
            $masterCol = $MasterDataStartCol + $pointer
            # сверка на пересмотр значений
            $rawColVals = Get-ColumnBlock $rawWs $rc $FIRST_DATA_ROW (Get-UsedLastRow $rawWs)
            $masterColVals = Get-ColumnBlock $masterWs $masterCol $FIRST_DATA_ROW $masterLastRow
            for ($i = 0; $i -lt $alignedRawRows.Count; $i++) {
                $rawRow = $alignedRawRows[$i].Row
                $masterRow = $masterSig[$i].Row
                $rawVal = $rawColVals[$rawRow - $FIRST_DATA_ROW]
                $masterVal = $masterColVals[$masterRow - $FIRST_DATA_ROW]
                if ($null -eq $rawVal -and $null -eq $masterVal) { continue }
                $differs = $false
                if (($rawVal -is [double] -or $rawVal -is [int]) -and ($masterVal -is [double] -or $masterVal -is [int])) {
                    $differs = [Math]::Abs([double]$rawVal - [double]$masterVal) -gt [Math]::Max(1e-6, [Math]::Abs([double]$masterVal) * 1e-9)
                } else {
                    $differs = $rawVal -ne $masterVal
                }
                if ($differs) {
                    $RevisionsOut.Add([PSCustomObject]@{
                        Sheet = $Name; Col = (FmtKey $rv); Label = $alignedRawRows[$i].Label
                        Old = $masterVal; New = $rawVal; Row = $masterRow
                    })
                }
            }
            $pointer++
            continue
        }

        if ($masterKeys.Contains($key)) {
            Write-Host "[$Name] Колонка $(FmtKey $rv) стоит не по порядку относительно мастер-файла - пропускаю, разберись руками."
            continue
        }

        # новая колонка - вставляем перед текущей позицией pointer
        $insertCol = $MasterDataStartCol + $pointer
        $newColsAdded.Add([PSCustomObject]@{ Col = $insertCol; Value = $rv })
        if ($Apply) {
            $masterWs.Columns($insertCol).Insert([System.Reflection.Missing]::Value, 0) | Out-Null
            $masterWs.Cells($HEADER_ROW, $insertCol).Value2 = $rv
            $rawColVals = Get-ColumnBlock $rawWs $rc $FIRST_DATA_ROW (Get-UsedLastRow $rawWs)
            for ($i = 0; $i -lt $alignedRawRows.Count; $i++) {
                $rawRow = $alignedRawRows[$i].Row
                $masterRow = $masterSig[$i].Row
                $val = $rawColVals[$rawRow - $FIRST_DATA_ROW]
                if ($null -ne $val) { $masterWs.Cells($masterRow, $insertCol).Value2 = $val }
            }
            $masterLastCol++
        }
        $masterKeys.Insert($pointer, $key)
        $pointer++
    }

    if ($Apply -and $newColsAdded.Count -gt 0) {
        $masterWs.AutoFilterMode = $false
        $newLastCol = $MasterDataStartCol + $masterKeys.Count - 1
        $lastRowNow = Get-UsedLastRow $masterWs
        $masterWs.Range($masterWs.Cells($HEADER_ROW, $AutoFilterStartCol), $masterWs.Cells($lastRowNow, $newLastCol)).AutoFilter($null) | Out-Null
    }

    return $newColsAdded
}

# ---- main ----

Write-Host "Читаю новую выгрузку: $RawPath"
Write-Host "Читаю мастер-файл: $MasterPath"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
$excel.Calculation = -4135  # xlCalculationManual, чтобы не тормозило на каждой правке

try {
    $rawWb = $excel.Workbooks.Open($RawPath, [Type]::Missing, $true)  # ReadOnly
    $masterWb = $excel.Workbooks.Open($MasterPath)

    $revisions = New-Object System.Collections.Generic.List[object]

    $newKolVo = Sync-Sheet -Name "кол-во" -RawWb $rawWb -MasterWb $masterWb `
        -RawSheet "01_02_01" -RawLabelCol 1 -RawDimCol 2 -RawDataStartCol 3 `
        -MasterSheet "кол-во" -MasterLabelCol 1 -MasterDimCol 3 -MasterDataStartCol 4 `
        -TargetDim "Всего, шт." -ExcludeSet $EXCLUDED_PROGRAMS `
        -IsDdu $false -Apply $Apply.IsPresent -RevisionsOut $revisions -AutoFilterStartCol 1

    $newObyem = Sync-Sheet -Name "объем" -RawWb $rawWb -MasterWb $masterWb `
        -RawSheet "01_02_01" -RawLabelCol 1 -RawDimCol 2 -RawDataStartCol 3 `
        -MasterSheet "объем" -MasterLabelCol 1 -MasterDimCol 3 -MasterDataStartCol 4 `
        -TargetDim "Всего, млн руб." -ExcludeSet $EXCLUDED_PROGRAMS `
        -IsDdu $false -Apply $Apply.IsPresent -RevisionsOut $revisions -AutoFilterStartCol 1

    $newDdu = Sync-Sheet -Name "ДДУ" -RawWb $rawWb -MasterWb $masterWb `
        -RawSheet "01_02_03" -RawLabelCol 1 -RawDimCol 0 -RawDataStartCol 2 `
        -MasterSheet "ДДУ" -MasterLabelCol 3 -MasterDimCol 0 -MasterDataStartCol 5 `
        -TargetDim $null -ExcludeSet $null `
        -IsDdu $true -Apply $Apply.IsPresent -RevisionsOut $revisions -AutoFilterStartCol 3

    Write-Host "`n=== ОТЧЁТ ==="
    foreach ($pair in @(@("кол-во", $newKolVo), @("объем", $newObyem), @("ДДУ", $newDdu))) {
        $name = $pair[0]; $cols = $pair[1]
        if ($cols.Count -gt 0) {
            $keys = ($cols | ForEach-Object { FmtKey $_.Value }) -join ", "
            Write-Host "[$name] Новые колонки ($($cols.Count)): $keys"
        } else {
            Write-Host "[$name] Новых колонок нет."
        }
    }

    if ($revisions.Count -gt 0) {
        Write-Host "`nПересмотр значений задним числом ($($revisions.Count) ячеек, первые 20):"
        $revisions | Select-Object -First 20 | ForEach-Object {
            Write-Host "  [$($_.Sheet)] $($_.Col) / строка $($_.Row) ('$($_.Label)'): было $($_.Old) -> стало $($_.New)"
        }
    } else {
        Write-Host "`nПересмотров значений в существующих колонках не найдено."
    }

    if ($Apply) {
        $masterWb.Save()
        Write-Host "`nСохранено: $MasterPath"
        Write-Host "Дальше: посмотри диф в Excel, потом закоммить и запушь руками."
    } else {
        Write-Host "`nЭто был dry-run (без -Apply), файл не менялся."
    }

    $rawWb.Close($false)
    $masterWb.Close($Apply.IsPresent)
} finally {
    $excel.Calculation = -4105  # xlCalculationAutomatic
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
