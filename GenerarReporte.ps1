Add-Type -AssemblyName System.Drawing

function OleColor([int]$r,[int]$g,[int]$b) {
    return [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb($r,$g,$b))
}
$cDarkBlue  = OleColor 31  73  125
$cWhite     = OleColor 255 255 255
$cLightBlue = OleColor 235 241 250
$cGreen     = OleColor   0 128   0
$cOrange    = OleColor 191 144   0
$cRed       = OleColor 192   0   0
$cGray      = OleColor  89  89  89
$cYellow    = OleColor 255 235 156
$cRedBg     = OleColor 255 199 206

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false

$srcPath = "C:\Users\bchevasco\OneDrive - Articulos Promocionales SA\Escritorio\Asistencia\Asistencia.xlsx"
$srcWb = $xl.Workbooks.Open($srcPath)
$srcWs = $srcWb.Sheets.Item(1)
$totalRows = $srcWs.UsedRange.Rows.Count
$data = @()
for ($r = 3; $r -le $totalRows; $r++) {
    $data += [PSCustomObject]@{
        Apellidos = $srcWs.Cells.Item($r,1).Text
        Nombre    = $srcWs.Cells.Item($r,2).Text
        Grupo     = $srcWs.Cells.Item($r,4).Text
        Fecha     = $srcWs.Cells.Item($r,5).Text
        Turno     = $srcWs.Cells.Item($r,7).Text
        Entro     = $srcWs.Cells.Item($r,8).Text
    }
}
$srcWb.Close($false)

$wd = @('Lun','Mar','Mie','Jue','Vie')
$dataWD = $data | Where-Object {
    $f = $_.Fecha
    $prefix = if ($f -match '^(\w+)\s') { $matches[1] } else { '' }
    $wd -contains $prefix -or $prefix -eq 'Mi'
}
$dataWD = $data | Where-Object {
    $f = $_.Fecha
    ($f -match '^Lun\s' -or $f -match '^Mar\s' -or $f -match '^Mi' -or $f -match '^Jue\s' -or $f -match '^Vie\s')
}

$personas = $dataWD | ForEach-Object { "$($_.Apellidos)|$($_.Nombre)|$($_.Grupo)" } | Sort-Object -Unique

$analisis = @()
foreach ($p in $personas) {
    $pts = $p -split '\|'
    $apell = $pts[0]; $nom = $pts[1]; $grp = $pts[2]
    $rows = $dataWD | Where-Object { $_.Apellidos -eq $apell -and $_.Nombre -eq $nom }
    $nFeriado  = [int](@($rows | Where-Object { $_.Fecha -match '\(F\)' }).Count)
    $nDescanso = [int](@($rows | Where-Object { $_.Turno -eq 'Descanso' -and $_.Fecha -notmatch '\(F\)' }).Count)
    $rowsLab   = @($rows | Where-Object { $_.Turno -ne 'Descanso' -and $_.Fecha -notmatch '\(F\)' })
    $nPresente = [int](@($rowsLab | Where-Object { $_.Entro -ne '' }).Count)
    $nAusente  = [int](@($rowsLab | Where-Object { $_.Entro -eq '' }).Count)
    $nLab      = [int]$rowsLab.Count
    $turnoVal  = if ($rowsLab.Count -gt 0) { $rowsLab[0].Turno } else { '' }
    $ausentes  = @($rowsLab | Where-Object { $_.Entro -eq '' } | Select-Object Fecha, Turno)
    $analisis += [PSCustomObject]@{
        Apellidos = $apell; Nombre = $nom; Grupo = $grp; Turno = $turnoVal
        nLab = $nLab; nFeriado = $nFeriado; nDescanso = $nDescanso
        nPresente = $nPresente; nAusente = $nAusente
        Ausentes = $ausentes
    }
}

$wb = $xl.Workbooks.Add()

# ==========================================
# SHEET 1: RESUMEN GENERAL
# ==========================================
$ws1 = $wb.Sheets.Item(1)
$ws1.Name = "Resumen"

$ws1.Range("A1:I1").Merge()
$ws1.Cells.Item(1,1).Value2 = "REPORTE DE ASISTENCIA - LUNES A VIERNES"
$ws1.Cells.Item(1,1).Font.Bold = $true
$ws1.Cells.Item(1,1).Font.Size = 14
$ws1.Cells.Item(1,1).Font.Name = "Arial"
$ws1.Cells.Item(1,1).HorizontalAlignment = -4108
$ws1.Cells.Item(1,1).VerticalAlignment = -4108
$ws1.Cells.Item(1,1).Interior.Color = $cDarkBlue
$ws1.Cells.Item(1,1).Font.Color = $cWhite
$ws1.Rows.Item(1).RowHeight = 30

$ws1.Range("A2:I2").Merge()
$ws1.Cells.Item(2,1).Value2 = "Periodo: 01/05/2026 - 29/05/2026   |   Generado: 14/05/2026   |   * Fechas desde 15/05 son dias futuros (aun no ocurridos)"
$ws1.Cells.Item(2,1).Font.Italic = $true
$ws1.Cells.Item(2,1).Font.Size = 9
$ws1.Cells.Item(2,1).Font.Name = "Arial"
$ws1.Cells.Item(2,1).Font.Color = $cGray
$ws1.Cells.Item(2,1).HorizontalAlignment = -4108

$headers = @("Apellidos y Nombre","Grupo","Turno","Dias Lab. L-V","Feriados","Desc. en Semana","Dias Presentes","Dias Ausentes","% Asistencia")
for ($c = 1; $c -le $headers.Count; $c++) {
    $cell = $ws1.Cells.Item(4,$c)
    $cell.Value2 = $headers[$c-1]
    $cell.Font.Bold = $true; $cell.Font.Name = "Arial"; $cell.Font.Size = 10
    $cell.Font.Color = $cWhite; $cell.Interior.Color = $cDarkBlue
    $cell.HorizontalAlignment = -4108; $cell.VerticalAlignment = -4108
    $cell.WrapText = $true; $cell.Borders.LineStyle = 1
}
$ws1.Rows.Item(4).RowHeight = 32

$row = 5; $ci = 0
foreach ($a in $analisis) {
    $bgColor = if ($ci % 2 -eq 0) { $cWhite } else { $cLightBlue }; $ci++

    $ws1.Cells.Item($row,1).Value2 = "$($a.Apellidos), $($a.Nombre)"
    $ws1.Cells.Item($row,2).Value2 = $a.Grupo
    $ws1.Cells.Item($row,3).Value2 = $a.Turno
    $ws1.Cells.Item($row,4).Formula = "=$($a.nLab)"
    $ws1.Cells.Item($row,5).Formula = "=$($a.nFeriado)"
    $ws1.Cells.Item($row,6).Formula = "=$($a.nDescanso)"
    $ws1.Cells.Item($row,7).Formula = "=$($a.nPresente)"
    $ws1.Cells.Item($row,8).Formula = "=$($a.nAusente)"

    for ($c = 1; $c -le 8; $c++) {
        $ws1.Cells.Item($row,$c).Font.Name = "Arial"; $ws1.Cells.Item($row,$c).Font.Size = 10
        $ws1.Cells.Item($row,$c).Interior.Color = $bgColor; $ws1.Cells.Item($row,$c).Borders.LineStyle = 1
        if ($c -ge 4) { $ws1.Cells.Item($row,$c).HorizontalAlignment = -4108 }
    }

    $pct = if ($a.nLab -gt 0) { $a.nPresente / $a.nLab } else { 0 }
    $pCell = $ws1.Cells.Item($row,9)
    if ($a.nLab -gt 0) {
        $pCell.Formula = "=G$row/D$row"
        $pCell.NumberFormat = "0.0%"
        $pCell.Font.Bold = $true
        if ($pct -ge 0.9) { $pCell.Font.Color = $cGreen }
        elseif ($pct -ge 0.7) { $pCell.Font.Color = $cOrange }
        else { $pCell.Font.Color = $cRed }
    } else { $pCell.Value2 = "N/A" }
    $pCell.Font.Name = "Arial"; $pCell.Font.Size = 10
    $pCell.Interior.Color = $bgColor; $pCell.HorizontalAlignment = -4108; $pCell.Borders.LineStyle = 1

    if ($a.nAusente -ge 5) { $ws1.Cells.Item($row,8).Interior.Color = $cRedBg }
    $row++
}

$lastDataRow = $row - 1
$ws1.Range("A$row`:I$row").Interior.Color = $cDarkBlue
$ws1.Cells.Item($row,1).Value2 = "TOTALES"
$ws1.Cells.Item($row,1).Font.Bold = $true; $ws1.Cells.Item($row,1).Font.Name = "Arial"
$ws1.Cells.Item($row,1).Font.Color = $cWhite; $ws1.Cells.Item($row,1).Borders.LineStyle = 1
$colL = @('','A','B','C','D','E','F','G','H','I')
for ($c = 4; $c -le 8; $c++) {
    $l = $colL[$c]
    $ws1.Cells.Item($row,$c).Formula = "=SUM(${l}5:${l}$lastDataRow)"
    $ws1.Cells.Item($row,$c).Font.Bold = $true; $ws1.Cells.Item($row,$c).Font.Name = "Arial"
    $ws1.Cells.Item($row,$c).Font.Color = $cWhite; $ws1.Cells.Item($row,$c).HorizontalAlignment = -4108
    $ws1.Cells.Item($row,$c).Borders.LineStyle = 1
}
foreach ($c in @(1,2,3,9)) {
    $ws1.Cells.Item($row,$c).Interior.Color = $cDarkBlue
    $ws1.Cells.Item($row,$c).Borders.LineStyle = 1
}

$cw = @(35,18,24,14,10,16,16,15,14)
for ($c = 1; $c -le $cw.Count; $c++) { $ws1.Columns.Item($c).ColumnWidth = $cw[$c-1] }
$ws1.Activate()
$xl.ActiveWindow.SplitRow = 4; $xl.ActiveWindow.FreezePanes = $true

# ==========================================
# SHEET 2: DETALLE AUSENCIAS
# ==========================================
$ws2 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $ws1)
$ws2.Name = "Detalle Ausencias"

$ws2.Range("A1:E1").Merge()
$ws2.Cells.Item(1,1).Value2 = "DETALLE DE AUSENCIAS POR PERSONA - LUNES A VIERNES"
$ws2.Cells.Item(1,1).Font.Bold = $true; $ws2.Cells.Item(1,1).Font.Size = 13; $ws2.Cells.Item(1,1).Font.Name = "Arial"
$ws2.Cells.Item(1,1).HorizontalAlignment = -4108; $ws2.Cells.Item(1,1).VerticalAlignment = -4108
$ws2.Cells.Item(1,1).Interior.Color = $cDarkBlue; $ws2.Cells.Item(1,1).Font.Color = $cWhite
$ws2.Rows.Item(1).RowHeight = 28

$ws2.Range("A2:E2").Merge()
$ws2.Cells.Item(2,1).Value2 = "Fondo amarillo = fecha futura (a partir del 15/05/2026).  Fondo rojo = ausencia ocurrida."
$ws2.Cells.Item(2,1).Font.Italic = $true; $ws2.Cells.Item(2,1).Font.Size = 9; $ws2.Cells.Item(2,1).Font.Name = "Arial"
$ws2.Cells.Item(2,1).Font.Color = $cGray

$row2 = 4
foreach ($a in $analisis) {
    $ws2.Range("A$row2`:E$row2").Merge()
    $ws2.Cells.Item($row2,1).Value2 = "$($a.Apellidos), $($a.Nombre)  |  $($a.Grupo)  |  Turno: $($a.Turno)  |  Ausencias: $($a.nAusente) de $($a.nLab) dias laborales"
    $ws2.Cells.Item($row2,1).Font.Bold = $true; $ws2.Cells.Item($row2,1).Font.Name = "Arial"; $ws2.Cells.Item($row2,1).Font.Size = 10
    $ws2.Cells.Item($row2,1).Interior.Color = OleColor 68 114 196; $ws2.Cells.Item($row2,1).Font.Color = $cWhite
    $ws2.Cells.Item($row2,1).HorizontalAlignment = -4131; $ws2.Cells.Item($row2,1).Borders.LineStyle = 1
    $row2++

    if ($a.nAusente -eq 0) {
        $ws2.Cells.Item($row2,1).Value2 = "     Sin ausencias en el periodo."
        $ws2.Cells.Item($row2,1).Font.Name = "Arial"; $ws2.Cells.Item($row2,1).Font.Size = 9
        $ws2.Cells.Item($row2,1).Font.Color = $cGreen; $ws2.Cells.Item($row2,1).Font.Italic = $true
        $row2++
    } else {
        $sh = @("Fecha","Dia","Turno Asignado","Estado","Obs.")
        for ($c = 1; $c -le 5; $c++) {
            $ws2.Cells.Item($row2,$c).Value2 = $sh[$c-1]
            $ws2.Cells.Item($row2,$c).Font.Bold = $true; $ws2.Cells.Item($row2,$c).Font.Name = "Arial"; $ws2.Cells.Item($row2,$c).Font.Size = 9
            $ws2.Cells.Item($row2,$c).Interior.Color = OleColor 149 179 215; $ws2.Cells.Item($row2,$c).Font.Color = OleColor 31 73 125
            $ws2.Cells.Item($row2,$c).HorizontalAlignment = -4108; $ws2.Cells.Item($row2,$c).Borders.LineStyle = 1
        }
        $row2++

        foreach ($aus in $a.Ausentes) {
            $fechaStr = $aus.Fecha
            $dayMap = @{ 'Lun'='Lunes'; 'Mar'='Martes'; 'Jue'='Jueves'; 'Vie'='Viernes' }
            $dayFull = 'Miercoles'
            foreach ($k in $dayMap.Keys) { if ($fechaStr -match "^$k") { $dayFull = $dayMap[$k]; break } }
            $dayNum = 0
            if ($fechaStr -match '(\d{2})-05-2026') { $dayNum = [int]$matches[1] }
            $isFuture = $dayNum -ge 15
            $bgAus = if ($isFuture) { $cYellow } else { $cRedBg }
            $estado = if ($isFuture) { "Pendiente (futuro)" } else { "AUSENTE" }
            $obs = if ($isFuture) { "Dia pendiente" } else { "Sin registro de entrada" }

            $vals = @($fechaStr, $dayFull, $aus.Turno, $estado, $obs)
            for ($c = 1; $c -le 5; $c++) {
                $ws2.Cells.Item($row2,$c).Value2 = $vals[$c-1]
                $ws2.Cells.Item($row2,$c).Font.Name = "Arial"; $ws2.Cells.Item($row2,$c).Font.Size = 9
                $ws2.Cells.Item($row2,$c).Interior.Color = $bgAus; $ws2.Cells.Item($row2,$c).Borders.LineStyle = 1
                if ($c -eq 4 -and -not $isFuture) {
                    $ws2.Cells.Item($row2,$c).Font.Bold = $true
                    $ws2.Cells.Item($row2,$c).Font.Color = OleColor 192 0 0
                }
            }
            $row2++
        }
    }
    $row2++
}

$cw2 = @(28, 12, 26, 22, 28)
for ($c = 1; $c -le $cw2.Count; $c++) { $ws2.Columns.Item($c).ColumnWidth = $cw2[$c-1] }

# ==========================================
# SHEET 3: ASISTENCIA REAL HASTA HOY
# ==========================================
$ws3 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $ws2)
$ws3.Name = "Asistencia Actual"

$ws3.Range("A1:G1").Merge()
$ws3.Cells.Item(1,1).Value2 = "ASISTENCIA REAL AL 14/05/2026 (solo dias ocurridos)"
$ws3.Cells.Item(1,1).Font.Bold = $true; $ws3.Cells.Item(1,1).Font.Size = 13; $ws3.Cells.Item(1,1).Font.Name = "Arial"
$ws3.Cells.Item(1,1).HorizontalAlignment = -4108; $ws3.Cells.Item(1,1).VerticalAlignment = -4108
$ws3.Cells.Item(1,1).Interior.Color = OleColor 0 112 192; $ws3.Cells.Item(1,1).Font.Color = $cWhite
$ws3.Rows.Item(1).RowHeight = 28

$ws3.Range("A2:G2").Merge()
$ws3.Cells.Item(2,1).Value2 = "Dias laborales computados: 05/05 (Lun) al 14/05 (Mie) = 8 dias.  Feriado 01/05 excluido."
$ws3.Cells.Item(2,1).Font.Italic = $true; $ws3.Cells.Item(2,1).Font.Size = 9; $ws3.Cells.Item(2,1).Font.Name = "Arial"
$ws3.Cells.Item(2,1).Font.Color = $cGray

$headers3 = @("Apellidos y Nombre","Grupo","Dias Laborales","Presentes","Ausentes","% Asistencia","Dias Ausente (detalle)")
for ($c = 1; $c -le $headers3.Count; $c++) {
    $cell = $ws3.Cells.Item(4,$c)
    $cell.Value2 = $headers3[$c-1]
    $cell.Font.Bold = $true; $cell.Font.Name = "Arial"; $cell.Font.Size = 10
    $cell.Font.Color = $cWhite; $cell.Interior.Color = OleColor 0 112 192
    $cell.HorizontalAlignment = -4108; $cell.VerticalAlignment = -4108; $cell.WrapText = $true; $cell.Borders.LineStyle = 1
}
$ws3.Rows.Item(4).RowHeight = 32

$actualDates = @('05-05-2026','06-05-2026','07-05-2026','08-05-2026','11-05-2026','12-05-2026','13-05-2026','14-05-2026')

$row3 = 5; $ci3 = 0
foreach ($a in $analisis) {
    $rows = $dataWD | Where-Object { $_.Apellidos -eq $a.Apellidos -and $_.Nombre -eq $a.Nombre }

    $rowsAct = @($rows | Where-Object {
        $f = $_.Fecha
        $matched = $false
        foreach ($d in $actualDates) { if ($f -match $d) { $matched = $true; break } }
        $matched -and $_.Turno -ne 'Descanso' -and $_.Fecha -notmatch '\(F\)'
    })
    $nLabAct  = [int]$rowsAct.Count
    $nPresAct = [int](@($rowsAct | Where-Object { $_.Entro -ne '' }).Count)
    $nAusAct  = [int](@($rowsAct | Where-Object { $_.Entro -eq '' }).Count)
    $ausDates = ($rowsAct | Where-Object { $_.Entro -eq '' } | ForEach-Object { $_.Fecha }) -join ", "

    $bgColor3 = if ($ci3 % 2 -eq 0) { $cWhite } else { $cLightBlue }; $ci3++

    $ws3.Cells.Item($row3,1).Value2 = "$($a.Apellidos), $($a.Nombre)"
    $ws3.Cells.Item($row3,2).Value2 = $a.Grupo
    $ws3.Cells.Item($row3,3).Formula = "=$nLabAct"
    $ws3.Cells.Item($row3,4).Formula = "=$nPresAct"
    $ws3.Cells.Item($row3,5).Formula = "=$nAusAct"

    for ($c = 1; $c -le 5; $c++) {
        $ws3.Cells.Item($row3,$c).Font.Name = "Arial"; $ws3.Cells.Item($row3,$c).Font.Size = 10
        $ws3.Cells.Item($row3,$c).Interior.Color = $bgColor3; $ws3.Cells.Item($row3,$c).Borders.LineStyle = 1
        if ($c -ge 3) { $ws3.Cells.Item($row3,$c).HorizontalAlignment = -4108 }
    }

    $pctCell3 = $ws3.Cells.Item($row3,6)
    if ($nLabAct -gt 0) {
        $pct3 = $nPresAct / $nLabAct
        $pctCell3.Formula = "=D$row3/C$row3"
        $pctCell3.NumberFormat = "0.0%"; $pctCell3.Font.Bold = $true
        if ($pct3 -ge 0.9) { $pctCell3.Font.Color = $cGreen }
        elseif ($pct3 -ge 0.7) { $pctCell3.Font.Color = $cOrange }
        else { $pctCell3.Font.Color = $cRed }
    } else { $pctCell3.Value2 = "N/A" }
    $pctCell3.Font.Name = "Arial"; $pctCell3.Font.Size = 10
    $pctCell3.Interior.Color = $bgColor3; $pctCell3.HorizontalAlignment = -4108; $pctCell3.Borders.LineStyle = 1

    $detalle = if ($ausDates -ne '') { $ausDates } else { 'Sin ausencias' }
    $ws3.Cells.Item($row3,7).Value2 = $detalle
    $ws3.Cells.Item($row3,7).Font.Name = "Arial"; $ws3.Cells.Item($row3,7).Font.Size = 9
    $ws3.Cells.Item($row3,7).Interior.Color = $bgColor3; $ws3.Cells.Item($row3,7).Borders.LineStyle = 1
    if ($nAusAct -gt 0) { $ws3.Cells.Item($row3,7).Font.Color = $cRed }

    if ($nAusAct -ge 2) { $ws3.Cells.Item($row3,5).Interior.Color = $cRedBg }
    $row3++
}

$cw3 = @(35,18,14,12,12,14,60)
for ($c = 1; $c -le $cw3.Count; $c++) { $ws3.Columns.Item($c).ColumnWidth = $cw3[$c-1] }
$ws3.Activate()
$xl.ActiveWindow.SplitRow = 4; $xl.ActiveWindow.FreezePanes = $true

# Reorder: Asistencia Actual first, then Resumen, then Detalle
$ws3.Move($wb.Sheets.Item(1))

$outPath = "C:\Users\bchevasco\OneDrive - Articulos Promocionales SA\Escritorio\Asistencia\Reporte_Asistencia.xlsx"
$wb.SaveAs($outPath, 51)
$wb.Close($false)
$xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null

Write-Host "LISTO: $outPath"
