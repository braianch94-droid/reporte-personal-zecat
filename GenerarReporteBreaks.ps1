Add-Type -AssemblyName System.Drawing
function OleColor([int]$r,[int]$g,[int]$b) { return [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb($r,$g,$b)) }
function ParseTime([string]$t) {
    if ($t -eq '') { return $null }
    $p = $t -split ':'; return [int]$p[0]*60 + [int]$p[1]
}
function MinToStr([int]$min) {
    $h = [math]::Floor($min/60); $m = $min % 60
    return "$($h.ToString('00')):$($m.ToString('00'))"
}

$cDarkBlue  = OleColor 31  73  125
$cBlue2     = OleColor 0  112  192
$cWhite     = OleColor 255 255 255
$cLightBlue = OleColor 235 241 250
$cGreen     = OleColor   0 128   0
$cGreenBg   = OleColor 198 239 206
$cOrange    = OleColor 191 144   0
$cOrangeBg  = OleColor 255 235 156
$cRed       = OleColor 192   0   0
$cRedBg     = OleColor 255 199 206
$cGray      = OleColor  89  89  89
$cGrayBg    = OleColor 242 242 242
$cPurple    = OleColor 112  48 160
$cPurpleBg  = OleColor 228 208 244

# ---- READ SOURCE ----
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$srcWb = $xl.Workbooks.Open("C:\Users\bchevasco\OneDrive - Articulos Promocionales SA\Escritorio\Asistencia\Asistencia.xlsx")
$srcWs = $srcWb.Sheets.Item(1)
$totalRows = $srcWs.UsedRange.Rows.Count
$data = @()
for ($r = 3; $r -le $totalRows; $r++) {
    $data += [PSCustomObject]@{
        Apellidos=$srcWs.Cells.Item($r,1).Text; Nombre=$srcWs.Cells.Item($r,2).Text
        Grupo=$srcWs.Cells.Item($r,4).Text; Fecha=$srcWs.Cells.Item($r,5).Text
        Turno=$srcWs.Cells.Item($r,7).Text; E1=$srcWs.Cells.Item($r,8).Text
        S1=$srcWs.Cells.Item($r,12).Text; E2=$srcWs.Cells.Item($r,16).Text
        S2=$srcWs.Cells.Item($r,20).Text; HT=$srcWs.Cells.Item($r,27).Text
    }
}
$srcWb.Close($false)

$dataWD = $data | Where-Object {
    $f = $_.Fecha
    ($f -match '^Lun\s' -or $f -match '^Mar\s' -or $f -match '^Mi' -or $f -match '^Jue\s' -or $f -match '^Vie\s')
}

# ---- ANALYZE BREAKS ----
$DESCANSO_MIN = 15   # minutos descanso esperado
$ALMUERZO_MIN = 45   # minutos almuerzo esperado
$TOTAL_MIN    = 60   # total esperado

$breakRows = @()
$personas = $dataWD | ForEach-Object { "$($_.Apellidos)|$($_.Nombre)|$($_.Grupo)" } | Sort-Object -Unique

$summary = @()
foreach ($p in $personas) {
    $pts = $p -split '\|'
    $apell = $pts[0]; $nom = $pts[1]; $grp = $pts[2]

    $presentes = @($dataWD | Where-Object {
        $_.Apellidos -eq $apell -and $_.Nombre -eq $nom -and
        $_.Turno -ne 'Descanso' -and $_.Fecha -notmatch '\(F\)' -and $_.E1 -ne ''
    })

    $conBreak = @($presentes | Where-Object { $_.S1 -ne '' -and $_.E2 -ne '' })
    $sinBreak = @($presentes | Where-Object { $_.S1 -eq '' -or $_.E2 -eq '' })

    $dayBreaks = @()
    foreach ($row in $conBreak) {
        $s1m = ParseTime $row.S1
        $e2m = ParseTime $row.E2
        if ($s1m -ne $null -and $e2m -ne $null) {
            $bMin = $e2m - $s1m
            if ($bMin -lt 0) { $bMin += 1440 }
            $diff = $bMin - $TOTAL_MIN
            $tipo = if ($bMin -le 20) { "Descanso" } elseif ($bMin -le 65) { "Almuerzo" } else { "Combinado" }
            $estado = if ($bMin -lt 10) { "SIN BREAK" }
                      elseif ($tipo -eq "Descanso" -and $bMin -le 20) { "OK" }
                      elseif ($tipo -eq "Almuerzo" -and $bMin -ge 40 -and $bMin -le 65) { "OK" }
                      elseif ($tipo -eq "Combinado" -and $bMin -le 75) { "LARGO" }
                      elseif ($bMin -gt 75) { "MUY LARGO" }
                      else { "CORTO" }
            $dayBreaks += [PSCustomObject]@{
                Apellidos=$apell; Nombre=$nom; Grupo=$grp; Fecha=$row.Fecha; Turno=$row.Turno
                S1=$row.S1; E2=$row.E2; BreakMin=$bMin; BreakStr=(MinToStr $bMin)
                TipoBreak=$tipo; Estado=$estado; Diferencia=$diff
            }
            $breakRows += $dayBreaks[-1]
        }
    }

    $nOK    = [int](@($dayBreaks | Where-Object { $_.Estado -eq 'OK' }).Count)
    $nLargo = [int](@($dayBreaks | Where-Object { $_.Estado -like '*LARGO*' }).Count)
    $nCorto = [int](@($dayBreaks | Where-Object { $_.Estado -like '*CORTO*' }).Count)
    $nSin   = [int](@($dayBreaks | Where-Object { $_.Estado -eq 'SIN BREAK' }).Count)
    $avgMin = if ($dayBreaks.Count -gt 0) { [int][math]::Round(($dayBreaks | Measure-Object -Property BreakMin -Average).Average) } else { 0 }
    $maxMin = if ($dayBreaks.Count -gt 0) { [int]($dayBreaks | Measure-Object -Property BreakMin -Maximum).Maximum } else { 0 }
    $minMin = if ($dayBreaks.Count -gt 0) { [int]($dayBreaks | Measure-Object -Property BreakMin -Minimum).Minimum } else { 0 }

    $summary += [PSCustomObject]@{
        Apellidos=$apell; Nombre=$nom; Grupo=$grp
        Presentes=$presentes.Count; ConBreak=$conBreak.Count; SinBreakData=$sinBreak.Count
        OK=$nOK; Largo=$nLargo; Corto=$nCorto; SinBreak=$nSin
        AvgMin=$avgMin; AvgStr=(MinToStr $avgMin)
        MaxMin=$maxMin; MaxStr=(MinToStr $maxMin)
        MinMin=$minMin; MinStr=(MinToStr $minMin)
    }
}

# ---- CREATE WORKBOOK ----
$wb = $xl.Workbooks.Add()

# ==========================================
# HOJA 1: RESUMEN
# ==========================================
$ws1 = $wb.Sheets.Item(1)
$ws1.Name = "Resumen Descansos"

$ws1.Range("A1:J1").Merge()
$ws1.Cells.Item(1,1).Value2 = "REPORTE DE DESCANSOS Y ALMUERZO - LUNES A VIERNES"
$ws1.Cells.Item(1,1).Font.Bold = $true; $ws1.Cells.Item(1,1).Font.Size = 14; $ws1.Cells.Item(1,1).Font.Name = "Arial"
$ws1.Cells.Item(1,1).HorizontalAlignment = -4108; $ws1.Cells.Item(1,1).VerticalAlignment = -4108
$ws1.Cells.Item(1,1).Interior.Color = $cDarkBlue; $ws1.Cells.Item(1,1).Font.Color = $cWhite
$ws1.Rows.Item(1).RowHeight = 30

$ws1.Range("A2:J2").Merge()
$ws1.Cells.Item(2,1).Value2 = "Descanso esperado: 15 min  |  Almuerzo esperado: 45 min  |  Total: 60 min  |  Periodo: 01/05 - 14/05/2026 (dias ocurridos)"
$ws1.Cells.Item(2,1).Font.Italic = $true; $ws1.Cells.Item(2,1).Font.Size = 9; $ws1.Cells.Item(2,1).Font.Name = "Arial"
$ws1.Cells.Item(2,1).Font.Color = $cGray

# Legend row 3
$ws1.Range("A3:B3").Merge()
$ws1.Cells.Item(3,1).Value2 = "VERDE = dentro del tiempo"
$ws1.Cells.Item(3,1).Interior.Color = $cGreenBg; $ws1.Cells.Item(3,1).Font.Name = "Arial"; $ws1.Cells.Item(3,1).Font.Size = 9; $ws1.Cells.Item(3,1).Borders.LineStyle = 1
$ws1.Range("C3:D3").Merge()
$ws1.Cells.Item(3,3).Value2 = "NARANJA = mayor al esperado"
$ws1.Cells.Item(3,3).Interior.Color = $cOrangeBg; $ws1.Cells.Item(3,3).Font.Name = "Arial"; $ws1.Cells.Item(3,3).Font.Size = 9; $ws1.Cells.Item(3,3).Borders.LineStyle = 1
$ws1.Range("E3:F3").Merge()
$ws1.Cells.Item(3,5).Value2 = "ROJO = corto o sin break"
$ws1.Cells.Item(3,5).Interior.Color = $cRedBg; $ws1.Cells.Item(3,5).Font.Name = "Arial"; $ws1.Cells.Item(3,5).Font.Size = 9; $ws1.Cells.Item(3,5).Borders.LineStyle = 1
$ws1.Range("G3:H3").Merge()
$ws1.Cells.Item(3,7).Value2 = "GRIS = sin datos en sistema"
$ws1.Cells.Item(3,7).Interior.Color = $cGrayBg; $ws1.Cells.Item(3,7).Font.Name = "Arial"; $ws1.Cells.Item(3,7).Font.Size = 9; $ws1.Cells.Item(3,7).Borders.LineStyle = 1

# Headers
$hdrs = @("Apellidos y Nombre","Grupo","Dias Presentes","Dias con Datos","Dias sin Datos Sistema","Breaks OK","Breaks Largos","Breaks Cortos","Promedio Break","Minimo","Maximo")
for ($c = 1; $c -le $hdrs.Count; $c++) {
    $cell = $ws1.Cells.Item(5,$c)
    $cell.Value2 = $hdrs[$c-1]
    $cell.Font.Bold = $true; $cell.Font.Name = "Arial"; $cell.Font.Size = 10
    $cell.Font.Color = $cWhite; $cell.Interior.Color = $cDarkBlue
    $cell.HorizontalAlignment = -4108; $cell.VerticalAlignment = -4108; $cell.WrapText = $true; $cell.Borders.LineStyle = 1
}
$ws1.Rows.Item(5).RowHeight = 36

$row = 6; $ci = 0
foreach ($s in $summary) {
    $bgBase = if ($ci % 2 -eq 0) { $cWhite } else { $cLightBlue }; $ci++

    $ws1.Cells.Item($row,1).Value2 = "$($s.Apellidos), $($s.Nombre)"
    $ws1.Cells.Item($row,2).Value2 = $s.Grupo
    $ws1.Cells.Item($row,3).Formula = "=$($s.Presentes)"
    $ws1.Cells.Item($row,4).Formula = "=$($s.ConBreak)"
    $ws1.Cells.Item($row,5).Formula = "=$($s.SinBreakData)"
    $ws1.Cells.Item($row,6).Formula = "=$($s.OK)"
    $ws1.Cells.Item($row,7).Formula = "=$($s.Largo)"
    $ws1.Cells.Item($row,8).Formula = "=$($s.Corto)"
    $ws1.Cells.Item($row,9).Value2 = if ($s.ConBreak -gt 0) { $s.AvgStr } else { "N/D" }
    $ws1.Cells.Item($row,10).Value2 = if ($s.ConBreak -gt 0) { $s.MinStr } else { "N/D" }
    $ws1.Cells.Item($row,11).Value2 = if ($s.ConBreak -gt 0) { $s.MaxStr } else { "N/D" }

    for ($c = 1; $c -le 11; $c++) {
        $ws1.Cells.Item($row,$c).Font.Name = "Arial"; $ws1.Cells.Item($row,$c).Font.Size = 10
        $ws1.Cells.Item($row,$c).Borders.LineStyle = 1
        if ($c -ge 3) { $ws1.Cells.Item($row,$c).HorizontalAlignment = -4108 }
    }

    # Color cells based on status
    if ($s.SinBreakData -eq $s.Presentes) {
        for ($c = 4; $c -le 11; $c++) { $ws1.Cells.Item($row,$c).Interior.Color = $cGrayBg }
        $ws1.Cells.Item($row,5).Font.Color = $cGray; $ws1.Cells.Item($row,5).Font.Bold = $true
        foreach ($c in @(1,2,3)) { $ws1.Cells.Item($row,$c).Interior.Color = $bgBase }
    } else {
        for ($c = 1; $c -le 11; $c++) { $ws1.Cells.Item($row,$c).Interior.Color = $bgBase }
        if ($s.OK -gt 0) { $ws1.Cells.Item($row,6).Interior.Color = $cGreenBg; $ws1.Cells.Item($row,6).Font.Color = $cGreen; $ws1.Cells.Item($row,6).Font.Bold = $true }
        if ($s.Largo -gt 0) { $ws1.Cells.Item($row,7).Interior.Color = $cOrangeBg; $ws1.Cells.Item($row,7).Font.Color = $cOrange; $ws1.Cells.Item($row,7).Font.Bold = $true }
        if ($s.Corto -gt 0) { $ws1.Cells.Item($row,8).Interior.Color = $cRedBg; $ws1.Cells.Item($row,8).Font.Color = $cRed; $ws1.Cells.Item($row,8).Font.Bold = $true }
        # Color avg cell
        if ($s.ConBreak -gt 0) {
            if ($s.AvgMin -ge 10 -and $s.AvgMin -le 65) { $ws1.Cells.Item($row,9).Interior.Color = $cGreenBg; $ws1.Cells.Item($row,9).Font.Color = $cGreen; $ws1.Cells.Item($row,9).Font.Bold = $true }
            elseif ($s.AvgMin -gt 65) { $ws1.Cells.Item($row,9).Interior.Color = $cOrangeBg; $ws1.Cells.Item($row,9).Font.Color = $cOrange; $ws1.Cells.Item($row,9).Font.Bold = $true }
            else { $ws1.Cells.Item($row,9).Interior.Color = $cRedBg; $ws1.Cells.Item($row,9).Font.Color = $cRed; $ws1.Cells.Item($row,9).Font.Bold = $true }
        }
    }
    $row++
}

$cw1 = @(32,18,14,14,20,12,14,14,16,12,12)
for ($c = 1; $c -le $cw1.Count; $c++) { $ws1.Columns.Item($c).ColumnWidth = $cw1[$c-1] }
$ws1.Rows.Item(1).RowHeight = 30; $ws1.Rows.Item(5).RowHeight = 36

# ==========================================
# HOJA 2: DETALLE DIA A DIA
# ==========================================
$ws2 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $ws1)
$ws2.Name = "Detalle por Dia"

$ws2.Range("A1:I1").Merge()
$ws2.Cells.Item(1,1).Value2 = "DETALLE DE BREAKS POR DIA - LUNES A VIERNES"
$ws2.Cells.Item(1,1).Font.Bold = $true; $ws2.Cells.Item(1,1).Font.Size = 13; $ws2.Cells.Item(1,1).Font.Name = "Arial"
$ws2.Cells.Item(1,1).HorizontalAlignment = -4108; $ws2.Cells.Item(1,1).VerticalAlignment = -4108
$ws2.Cells.Item(1,1).Interior.Color = $cDarkBlue; $ws2.Cells.Item(1,1).Font.Color = $cWhite
$ws2.Rows.Item(1).RowHeight = 28

$ws2.Range("A2:I2").Merge()
$ws2.Cells.Item(2,1).Value2 = "Break=Descanso(15m)+Almuerzo(45m)=60m total.  El sistema registra UN break por dia (salida S1 y retorno E2)."
$ws2.Cells.Item(2,1).Font.Italic = $true; $ws2.Cells.Item(2,1).Font.Size = 9; $ws2.Cells.Item(2,1).Font.Name = "Arial"
$ws2.Cells.Item(2,1).Font.Color = $cGray

$hdrs2 = @("Apellidos y Nombre","Grupo","Fecha","Turno","Hora Salida","Hora Retorno","Duracion","Tipo Break","Estado","Diferencia vs esperado")
for ($c = 1; $c -le $hdrs2.Count; $c++) {
    $cell = $ws2.Cells.Item(4,$c)
    $cell.Value2 = $hdrs2[$c-1]
    $cell.Font.Bold = $true; $cell.Font.Name = "Arial"; $cell.Font.Size = 10
    $cell.Font.Color = $cWhite; $cell.Interior.Color = $cDarkBlue
    $cell.HorizontalAlignment = -4108; $cell.VerticalAlignment = -4108; $cell.WrapText = $true; $cell.Borders.LineStyle = 1
}
$ws2.Rows.Item(4).RowHeight = 32

$row2 = 5; $ci2 = 0
$sortedBreaks = $breakRows | Sort-Object Apellidos, Nombre, Fecha
foreach ($b in $sortedBreaks) {
    $bgBase2 = if ($ci2 % 2 -eq 0) { $cWhite } else { $cLightBlue }; $ci2++

    $diffStr = if ($b.Diferencia -gt 0) { "+$($b.Diferencia) min" } elseif ($b.Diferencia -lt 0) { "$($b.Diferencia) min" } else { "Exacto" }

    $bgRow = switch ($b.Estado) {
        'OK'        { $cGreenBg }
        'LARGO'     { $cOrangeBg }
        'MUY LARGO' { $cOrangeBg }
        'CORTO'     { $cRedBg }
        'SIN BREAK' { $cRedBg }
        default     { $bgBase2 }
    }
    $fgRow = switch ($b.Estado) {
        'OK'        { $cGreen }
        'LARGO'     { $cOrange }
        'MUY LARGO' { OleColor 180 0 0 }
        'CORTO'     { $cRed }
        'SIN BREAK' { $cRed }
        default     { OleColor 0 0 0 }
    }

    $vals = @("$($b.Apellidos), $($b.Nombre)", $b.Grupo, $b.Fecha, $b.Turno, $b.S1, $b.E2, $b.BreakStr, $b.TipoBreak, $b.Estado, $diffStr)
    for ($c = 1; $c -le 10; $c++) {
        $cell = $ws2.Cells.Item($row2,$c)
        $cell.Value2 = $vals[$c-1]
        $cell.Font.Name = "Arial"; $cell.Font.Size = 9; $cell.Borders.LineStyle = 1
        $cell.Interior.Color = $bgBase2
        if ($c -ge 5) { $cell.HorizontalAlignment = -4108 }
    }
    # Color only the key columns
    foreach ($c in @(7,8,9,10)) {
        $ws2.Cells.Item($row2,$c).Interior.Color = $bgRow
        $ws2.Cells.Item($row2,$c).Font.Color = $fgRow
        $ws2.Cells.Item($row2,$c).Font.Bold = if ($b.Estado -ne 'OK') { $true } else { $false }
    }
    $row2++
}

$cw2 = @(30,16,20,24,14,14,12,14,14,22)
for ($c = 1; $c -le $cw2.Count; $c++) { $ws2.Columns.Item($c).ColumnWidth = $cw2[$c-1] }

$ws2.Activate()
$xl.ActiveWindow.SplitRow = 4; $xl.ActiveWindow.FreezePanes = $true

# ==========================================
# HOJA 3: RESUMEN POR PERSONA (detallado)
# ==========================================
$ws3 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $ws2)
$ws3.Name = "Por Persona"

$ws3.Range("A1:G1").Merge()
$ws3.Cells.Item(1,1).Value2 = "ANALISIS DE BREAKS POR PERSONA"
$ws3.Cells.Item(1,1).Font.Bold = $true; $ws3.Cells.Item(1,1).Font.Size = 13; $ws3.Cells.Item(1,1).Font.Name = "Arial"
$ws3.Cells.Item(1,1).HorizontalAlignment = -4108; $ws3.Cells.Item(1,1).VerticalAlignment = -4108
$ws3.Cells.Item(1,1).Interior.Color = $cBlue2; $ws3.Cells.Item(1,1).Font.Color = $cWhite
$ws3.Rows.Item(1).RowHeight = 28

$ws3.Range("A2:G2").Merge()
$ws3.Cells.Item(2,1).Value2 = "Descanso esperado: 15 min  |  Almuerzo esperado: 45 min  |  COTUGNO y KAPP solo registran el descanso de 15 min (sin almuerzo en sistema)"
$ws3.Cells.Item(2,1).Font.Italic = $true; $ws3.Cells.Item(2,1).Font.Size = 9; $ws3.Cells.Item(2,1).Font.Name = "Arial"
$ws3.Cells.Item(2,1).Font.Color = $cGray

$row3 = 4
foreach ($s in $summary) {
    $personBreaks = $breakRows | Where-Object { $_.Apellidos -eq $s.Apellidos -and $_.Nombre -eq $s.Nombre }

    # Person header
    $ws3.Range("A$row3`:G$row3").Merge()
    $label = "$($s.Apellidos), $($s.Nombre)  |  $($s.Grupo)"
    if ($s.SinBreakData -eq $s.Presentes) { $label += "  |  SIN DATOS DE BREAK EN SISTEMA" }
    else { $label += "  |  Promedio: $($s.AvgStr)  |  Min: $($s.MinStr)  |  Max: $($s.MaxStr)" }
    $ws3.Cells.Item($row3,1).Value2 = $label
    $ws3.Cells.Item($row3,1).Font.Bold = $true; $ws3.Cells.Item($row3,1).Font.Name = "Arial"; $ws3.Cells.Item($row3,1).Font.Size = 10
    $hdrColor = if ($s.SinBreakData -eq $s.Presentes) { OleColor 128 128 128 } else { OleColor 68 114 196 }
    $ws3.Cells.Item($row3,1).Interior.Color = $hdrColor; $ws3.Cells.Item($row3,1).Font.Color = $cWhite
    $ws3.Cells.Item($row3,1).HorizontalAlignment = -4131; $ws3.Cells.Item($row3,1).Borders.LineStyle = 1
    $row3++

    if ($s.SinBreakData -eq $s.Presentes) {
        $ws3.Cells.Item($row3,1).Value2 = "     El sistema no registra datos de salida/retorno de break para esta persona."
        $ws3.Cells.Item($row3,1).Font.Name = "Arial"; $ws3.Cells.Item($row3,1).Font.Size = 9
        $ws3.Cells.Item($row3,1).Font.Color = $cGray; $ws3.Cells.Item($row3,1).Font.Italic = $true
        $row3 += 2; continue
    }

    $sh3 = @("Fecha","Salida Break","Retorno","Duracion","Tipo","Estado","Diferencia")
    for ($c = 1; $c -le 7; $c++) {
        $ws3.Cells.Item($row3,$c).Value2 = $sh3[$c-1]
        $ws3.Cells.Item($row3,$c).Font.Bold = $true; $ws3.Cells.Item($row3,$c).Font.Name = "Arial"; $ws3.Cells.Item($row3,$c).Font.Size = 9
        $ws3.Cells.Item($row3,$c).Interior.Color = OleColor 149 179 215; $ws3.Cells.Item($row3,$c).Font.Color = OleColor 31 73 125
        $ws3.Cells.Item($row3,$c).HorizontalAlignment = -4108; $ws3.Cells.Item($row3,$c).Borders.LineStyle = 1
    }
    $row3++

    foreach ($b in $personBreaks) {
        $diffStr = if ($b.Diferencia -gt 0) { "+$($b.Diferencia) min" } elseif ($b.Diferencia -lt 0) { "$($b.Diferencia) min" } else { "Exacto" }
        $bgRow3 = switch ($b.Estado) {
            'OK'        { $cGreenBg }
            'LARGO'     { $cOrangeBg }
            'MUY LARGO' { $cOrangeBg }
            default     { $cRedBg }
        }
        $fgRow3 = switch ($b.Estado) {
            'OK'        { $cGreen }
            'LARGO'     { $cOrange }
            'MUY LARGO' { OleColor 180 0 0 }
            default     { $cRed }
        }
        $vals3 = @($b.Fecha, $b.S1, $b.E2, $b.BreakStr, $b.TipoBreak, $b.Estado, $diffStr)
        for ($c = 1; $c -le 7; $c++) {
            $ws3.Cells.Item($row3,$c).Value2 = $vals3[$c-1]
            $ws3.Cells.Item($row3,$c).Font.Name = "Arial"; $ws3.Cells.Item($row3,$c).Font.Size = 9
            $ws3.Cells.Item($row3,$c).Borders.LineStyle = 1
            if ($c -le 3) { $ws3.Cells.Item($row3,$c).Interior.Color = $cWhite }
            else {
                $ws3.Cells.Item($row3,$c).Interior.Color = $bgRow3
                $ws3.Cells.Item($row3,$c).Font.Color = $fgRow3
                if ($b.Estado -ne 'OK') { $ws3.Cells.Item($row3,$c).Font.Bold = $true }
            }
        }
        $row3++
    }
    $row3++
}

$cw3 = @(22, 14, 14, 14, 14, 14, 18)
for ($c = 1; $c -le $cw3.Count; $c++) { $ws3.Columns.Item($c).ColumnWidth = $cw3[$c-1] }

# Order sheets
$ws3.Move($wb.Sheets.Item(1))

# ---- SAVE ----
$outPath = "C:\Users\bchevasco\OneDrive - Articulos Promocionales SA\Escritorio\Asistencia\Reporte_Descansos.xlsx"
$wb.SaveAs($outPath, 51)
$wb.Close($false)
$xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
Write-Host "LISTO: $outPath"
