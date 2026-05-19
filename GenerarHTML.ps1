Add-Type -AssemblyName System.Drawing
function ParseTime([string]$t) {
    if ($t -eq '') { return $null }
    $p = $t -split ':'; return [int]$p[0]*60 + [int]$p[1]
}
function MinToStr([int]$min) {
    $h = [math]::Floor($min/60); $m = $min % 60
    return "$($h.ToString('00')):$($m.ToString('00'))"
}
function ToJson([string]$s) { return $s -replace "'","\\'" -replace '"','\"' }
function ParseTurnoEnd([string]$turno) {
    $times = [regex]::Matches($turno, '\d{1,2}:\d{2}')
    if ($times.Count -ge 2) { return ParseTime $times[$times.Count - 1].Value }
    return $null
}
function GetFechaDate([string]$fecha) {
    $m = [regex]::Match($fecha, '\d{2}-\d{2}-\d{4}')
    if ($m.Success) {
        try { return [datetime]::ParseExact($m.Value, 'dd-MM-yyyy', $null) } catch { return $null }
    }
    return $null
}

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
        Atraso1=$srcWs.Cells.Item($r,10).Text
        S1=$srcWs.Cells.Item($r,12).Text; Adelanto1=$srcWs.Cells.Item($r,14).Text
        E2=$srcWs.Cells.Item($r,16).Text; Atraso2=$srcWs.Cells.Item($r,18).Text
        S2=$srcWs.Cells.Item($r,20).Text; Adelanto2=$srcWs.Cells.Item($r,22).Text
        HT=$srcWs.Cells.Item($r,27).Text
    }
}
$srcWb.Close($false); $xl.Quit()
$data = @($data | Where-Object { $_.Apellidos -notlike '*AIRALA*' })

# Solo tomar datos a partir del 08/04/2026
$startDate = [datetime]::ParseExact('08-04-2026', 'dd-MM-yyyy', $null)
$data = @($data | Where-Object {
    $fd = GetFechaDate $_.Fecha
    $fd -eq $null -or $fd -ge $startDate
})

$dataWD = $data | Where-Object {
    $f = $_.Fecha
    ($f -match '^Lun\s' -or $f -match '^Mar\s' -or $f -match '^Mi' -or $f -match '^Jue\s' -or $f -match '^Vie\s')
}

$personas = $dataWD | ForEach-Object { "$($_.Apellidos)|$($_.Nombre)|$($_.Grupo)" } | Sort-Object -Unique

# ---- FERIADOS CONOCIDOS (agregar fechas en formato dd-MM-yyyy) ----
$feriadosFechas = @('02-04-2026', '03-04-2026')
$feriadosPattern = ($feriadosFechas | ForEach-Object { [regex]::Escape($_) }) -join '|'

# ---- JUSTIFICACIONES / VIAJES (no cuentan como ausencia real) ----
$justificaciones = @(
    [PSCustomObject]@{
        Apellidos  = 'DEMARCHIS'
        FechaDesde = [datetime]::ParseExact('08-04-2026', 'dd-MM-yyyy', $null)
        FechaHasta = [datetime]::ParseExact('10-04-2026', 'dd-MM-yyyy', $null)
        Motivo     = 'Viaje a Chile'
        Icono      = '&#9992;'
    }
)

$allPersonas = @()
foreach ($p in $personas) {
    $pts = $p -split '\|'
    $apell = $pts[0]; $nom = $pts[1]; $grp = $pts[2]
    $rows = $dataWD | Where-Object { $_.Apellidos -eq $apell -and $_.Nombre -eq $nom }

    # Asistencia
    $nFeriado   = [int](@($rows | Where-Object { $_.Fecha -match '\(F\)' -or ($feriadosPattern -ne '' -and $_.Fecha -match $feriadosPattern) }).Count)
    $nDescanso  = [int](@($rows | Where-Object { $_.Turno -eq 'Descanso' -and $_.Fecha -notmatch '\(F\)' -and ($feriadosPattern -eq '' -or $_.Fecha -notmatch $feriadosPattern) }).Count)
    $rowsLab    = @($rows | Where-Object { $_.Turno -ne 'Descanso' -and $_.Fecha -notmatch '\(F\)' -and ($feriadosPattern -eq '' -or $_.Fecha -notmatch $feriadosPattern) })
    $nLab       = [int]$rowsLab.Count

    # Justificaciones para esta persona (viajes, etc)
    $justifPersona = @($justificaciones | Where-Object { $apell -like "*$($_.Apellidos)*" })

    # Calcular ausDetails con flag Justificado (antes de nPresente para poder usarlo)
    $ausDetails = @($rowsLab | Where-Object { $_.E1 -eq '' -and $_.S1 -eq '' -and $_.E2 -eq '' -and $_.S2 -eq '' } | ForEach-Object {
        $rowF = $_.Fecha; $rowT = $_.Turno
        $isFut    = $rowF -match '\b(1[5-9]|[2-9]\d)-05-2026\b'
        $fd       = GetFechaDate $rowF
        $jMatch   = if ($fd -ne $null -and $justifPersona.Count -gt 0) {
            $justifPersona | Where-Object { $fd -ge $_.FechaDesde -and $fd -le $_.FechaHasta } | Select-Object -First 1
        } else { $null }
        $isJustif = $jMatch -ne $null
        [PSCustomObject]@{ Fecha=$rowF; Turno=$rowT; Futuro=$isFut; Justificado=$isJustif; Motivo=if($isJustif){$jMatch.Motivo}else{''} }
    })

    # Presente = cualquier fichada O dia justificado
    $nJustif   = [int](@($ausDetails | Where-Object { $_.Justificado }).Count)
    $nPresente = [int](@($rowsLab | Where-Object { $_.E1 -ne '' -or $_.S1 -ne '' -or $_.E2 -ne '' -or $_.S2 -ne '' }).Count) + $nJustif
    $nAusente  = [int](@($rowsLab | Where-Object { $_.E1 -eq '' -and $_.S1 -eq '' -and $_.E2 -eq '' -and $_.S2 -eq '' }).Count) - $nJustif
    $pct       = if ($nLab -gt 0) { [math]::Round($nPresente/$nLab*100,1) } else { 0 }

    # Breaks
    $rowsConBreak = @($rowsLab | Where-Object { $_.E1 -ne '' -and $_.S1 -ne '' -and $_.E2 -ne '' })
    $breakDays = @()
    foreach ($row in $rowsConBreak) {
        $s1m = ParseTime $row.S1; $e2m = ParseTime $row.E2
        if ($s1m -ne $null -and $e2m -ne $null) {
            $bMin = $e2m - $s1m; if ($bMin -lt 0) { $bMin += 1440 }
            $tipo = if ($bMin -le 22) { "Descanso" } elseif ($bMin -le 65) { "Almuerzo" } else { "Combinado" }
            $estado = if ($bMin -lt 10) { "sin-break" }
                      elseif ($tipo -eq "Descanso" -and $bMin -le 22) { "ok" }
                      elseif ($bMin -ge 40 -and $bMin -le 65) { "ok" }
                      elseif ($bMin -gt 65 -and $bMin -le 90) { "largo" }
                      elseif ($bMin -gt 90) { "muy-largo" }
                      else { "corto" }
            $diff = $bMin - 60
            $diffStr = if ($diff -gt 0) { "+${diff}m" } elseif ($diff -lt 0) { "${diff}m" } else { "exacto" }
            $breakDays += [PSCustomObject]@{
                Fecha=$row.Fecha; S1=$row.S1; E2=$row.E2
                BreakMin=$bMin; BreakStr=(MinToStr $bMin)
                Tipo=$tipo; Estado=$estado; Diff=$diff; DiffStr=$diffStr
            }
        }
    }
    $nConBreak = [int]$rowsConBreak.Count
    $nSinBreakData = [int](@($rowsLab | Where-Object { $_.E1 -ne '' -and ($_.S1 -eq '' -or $_.E2 -eq '') }).Count)
    $avgBreak = if ($breakDays.Count -gt 0) { [int][math]::Round(($breakDays | Measure-Object -Property BreakMin -Average).Average) } else { 0 }
    $maxBreak = if ($breakDays.Count -gt 0) { [int]($breakDays | Measure-Object -Property BreakMin -Maximum).Maximum } else { 0 }
    $minBreak = if ($breakDays.Count -gt 0) { [int]($breakDays | Measure-Object -Property BreakMin -Minimum).Minimum } else { 0 }
    $nBreakOK    = [int](@($breakDays | Where-Object { $_.Estado -eq 'ok' }).Count)
    $nBreakLargo = [int](@($breakDays | Where-Object { $_.Estado -like '*largo*' }).Count)
    $nBreakCorto = [int](@($breakDays | Where-Object { $_.Estado -like '*corto*' -or $_.Estado -eq 'sin-break' }).Count)

    # ---- NOVEDADES ----
    $rowsPresReal = @($rowsLab | Where-Object { ($_.E1 -ne '' -or $_.S1 -ne '' -or $_.E2 -ne '' -or $_.S2 -ne '') -and $_.Fecha -notmatch '\b(1[5-9]|[2-9]\d)-05-2026\b' })

    # Ausencias reales (no futuras, sin fichada, sin justificacion)
    $ausReales = @($ausDetails | Where-Object { -not $_.Futuro -and -not $_.Justificado } |
        ForEach-Object { [PSCustomObject]@{ Fecha=$_.Fecha; Turno=$_.Turno } })

    # Atrasos entrada
    $atrasos = @($rowsPresReal | Where-Object { $_.Atraso1 -ne '00:00' -and $_.Atraso1 -ne '' } | ForEach-Object {
        $min = ParseTime $_.Atraso1
        [PSCustomObject]@{ Fecha=$_.Fecha; Turno=$_.Turno; Entro=$_.E1; AtrasoStr=$_.Atraso1; AtrasoMin=$min }
    })
    $totalAtrasoMin = if ($atrasos.Count -gt 0) { [int]($atrasos | Measure-Object -Property AtrasoMin -Sum).Sum } else { 0 }

    # Demora retorno break (Atraso2)
    $demoraBreak = @($rowsPresReal | Where-Object { $_.Atraso2 -ne '00:00' -and $_.Atraso2 -ne '' } | ForEach-Object {
        $min = ParseTime $_.Atraso2
        [PSCustomObject]@{ Fecha=$_.Fecha; Retorno=$_.E2; DemoraStr=$_.Atraso2; DemoraMin=$min }
    })

    # Break excedido (real, ya calculado en breakDays, filtrar solo dias reales y largo)
    $breakExcedido = @($breakDays | Where-Object {
        ($_.Estado -like '*largo*') -and $_.Fecha -notmatch '\b(1[5-9]|[2-9]\d)-05-2026\b'
    })

    # Salidas anticipadas (los viernes todos salen 1h antes, descontar 60 min)
    $adelantos = @($rowsPresReal | ForEach-Object {
        $a1m = ParseTime $_.Adelanto1; $a2m = ParseTime $_.Adelanto2
        $best = if ($a2m -ne $null -and $a2m -gt 0) { $a2m } elseif ($a1m -ne $null -and $a1m -gt 0) { $a1m } else { 0 }
        if ($_.Fecha -match '^Vie\s') { $best = [math]::Max(0, $best - 60) }
        if ($best -gt 0) {
            [PSCustomObject]@{ Fecha=$_.Fecha; Turno=$_.Turno; AdelantoStr=(MinToStr $best); AdelantoMin=$best }
        }
    } | Where-Object { $_ -ne $null })
    $totalAdelantoMin = if ($adelantos.Count -gt 0) { [int]($adelantos | Measure-Object -Property AdelantoMin -Sum).Sum } else { 0 }

    $nNovedades = $ausReales.Count + $atrasos.Count + $demoraBreak.Count + $breakExcedido.Count + $adelantos.Count

    # ---- TOTAL INCONSISTENCIAS (para tabla resumen) ----
    $demoraBreakTotalMin  = if ($demoraBreak.Count  -gt 0) { [int]($demoraBreak  | Measure-Object -Property DemoraMin  -Sum).Sum } else { 0 }
    $breakExcedidoExcessMin = if ($breakExcedido.Count -gt 0) { [int](($breakExcedido | ForEach-Object { [math]::Max(0, $_.BreakMin - 60) }) | Measure-Object -Sum).Sum } else { 0 }
    $totalInconsistenciaMin = $totalAtrasoMin + $totalAdelantoMin + $demoraBreakTotalMin + $breakExcedidoExcessMin

    # ---- HORAS EXTRAS ----
    $rowsSabOT = @($data | Where-Object {
        $_.Apellidos -eq $apell -and $_.Nombre -eq $nom -and
        $_.Fecha -match '^S' -and
        ($_.E1 -ne '' -or $_.S1 -ne '' -or $_.E2 -ne '' -or $_.S2 -ne '') -and
        $_.Fecha -notmatch '\b(1[5-9]|[2-9]\d)-05-2026\b'
    })

    $htDays = @()

    # Lun-Vie: horas trabajadas (entrada->salida) vs esperadas (9h Lun-Jue, 8h Vie) -> 50%
    foreach ($row in $rowsPresReal) {
        $entMin  = ParseTime $row.E1
        $exitMin = if ($row.S2 -ne '') { ParseTime $row.S2 } elseif ($row.S1 -ne '') { ParseTime $row.S1 } else { $null }
        if ($entMin -eq $null -or $exitMin -eq $null) { continue }
        $workedMin   = $exitMin - $entMin
        if ($workedMin -lt 0) { $workedMin += 1440 }
        $expectedMin = if ($row.Fecha -match '^Vie\s') { 480 } else { 540 }
        $otMin = [math]::Max(0, $workedMin - $expectedMin)
        if ($otMin -ge 30) {
            $salidaStr = if ($row.S2 -ne '') { $row.S2 } else { $row.S1 }
            $htDays += [PSCustomObject]@{
                Fecha=$row.Fecha; Turno=$row.Turno
                Entrada=$row.E1; Salida=$salidaStr
                TurnoFin=(MinToStr ($entMin + $expectedMin))
                OTMin=$otMin; OT50Min=$otMin; OT100Min=0; EsDia='semana'
            }
        }
    }

    # Sabado: todos son HE -> hasta 13:00 = 50%, despues = 100%
    foreach ($row in $rowsSabOT) {
        $entMin  = ParseTime $row.E1
        $exitMin = if ($row.S2 -ne '') { ParseTime $row.S2 } elseif ($row.S1 -ne '') { ParseTime $row.S1 } else { $null }
        if ($entMin -eq $null -or $exitMin -eq $null) { continue }
        $limite  = 780  # 13:00 hs
        $ot50    = [math]::Max(0, [math]::Min($exitMin, $limite) - $entMin)
        $ot100   = [math]::Max(0, $exitMin - $limite)
        $salidaSab = if ($row.S2 -ne '') { $row.S2 } else { $row.S1 }
        $htDays += [PSCustomObject]@{
            Fecha=$row.Fecha; Turno=$row.Turno
            Entrada=$row.E1; Salida=$salidaSab
            TurnoFin='--'
            OTMin=($ot50+$ot100); OT50Min=$ot50; OT100Min=$ot100; EsDia='sabado'
        }
    }

    $totalOT50  = if ($htDays.Count -gt 0) { [int]($htDays | Measure-Object -Property OT50Min  -Sum).Sum } else { 0 }
    $totalOT100 = if ($htDays.Count -gt 0) { [int]($htDays | Measure-Object -Property OT100Min -Sum).Sum } else { 0 }
    $totalOTMin = $totalOT50 + $totalOT100

    $allPersonas += [PSCustomObject]@{
        Apellidos=$apell; Nombre=$nom; Grupo=$grp
        nLab=$nLab; nPresente=$nPresente; nAusente=$nAusente; Pct=$pct
        nFeriado=$nFeriado; nDescanso=$nDescanso
        AusDetails=$ausDetails
        nConBreak=$nConBreak; nSinBreakData=$nSinBreakData
        AvgBreak=$avgBreak; MaxBreak=$maxBreak; MinBreak=$minBreak
        nBreakOK=$nBreakOK; nBreakLargo=$nBreakLargo; nBreakCorto=$nBreakCorto
        BreakDays=$breakDays
        AusReales=$ausReales; Atrasos=$atrasos; TotalAtrasoMin=$totalAtrasoMin
        DemoraBreak=$demoraBreak; BreakExcedido=$breakExcedido
        Adelantos=$adelantos; TotalAdelantoMin=$totalAdelantoMin
        nNovedades=$nNovedades
        JustifPersona=$justifPersona
        HTDays=$htDays; TotalOT50=$totalOT50; TotalOT100=$totalOT100; TotalOTMin=$totalOTMin
        TotalInconsistenciaMin=$totalInconsistenciaMin
    }
}

$grupos = $allPersonas | ForEach-Object { $_.Grupo } | Sort-Object -Unique
$generadoEn = Get-Date -Format "dd/MM/yyyy HH:mm"

# ---- BUILD HTML ----
$html = @'
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Reporte de Personal - Articulos Promocionales SA</title>
<style>
  :root {
    --accent: #e8383d;
    --accent-dark: #c02428;
    --bg-page: #111214;
    --bg-card: #1a1c1f;
    --bg-inner: #212326;
    --bg-hover: #252729;
    --border: #2e3035;
    --border-light: #383b40;
    --text-primary: #f0f0f0;
    --text-secondary: #9a9da3;
    --text-muted: #5a5d63;
    --green: #22c55e;
    --green-bg: rgba(34,197,94,.15);
    --orange: #f59e0b;
    --orange-bg: rgba(245,158,11,.15);
    --red: #ef4444;
    --red-bg: rgba(239,68,68,.15);
    --purple: #a855f7;
    --purple-bg: rgba(168,85,247,.15);
    --white: #ffffff;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', Arial, sans-serif; font-size: 13px; background: var(--bg-page); color: var(--text-primary); }

  header { background: #0d0e10; color: white; padding: 0 24px; display:flex; align-items:center; gap:16px; height:58px; border-bottom: 1px solid var(--border); box-shadow: 0 2px 12px rgba(0,0,0,.5); }
  .header-logo { width:34px; height:34px; border-radius:50%; object-fit:cover; flex-shrink:0; }
  .header-logo-placeholder { width:34px; height:34px; border-radius:50%; background:var(--accent); display:flex; align-items:center; justify-content:center; font-weight:900; font-size:16px; color:white; flex-shrink:0; }
  .header-title { font-size:16px; font-weight:700; letter-spacing:.3px; flex:1; }
  .header-sub { font-size:11px; color:var(--text-muted); white-space:nowrap; }
  .btn-refresh { display:flex; align-items:center; gap:6px; padding:7px 16px; background:var(--accent); color:white; border:none; border-radius:6px; font-size:12px; font-weight:700; cursor:pointer; transition:background .15s; white-space:nowrap; }
  .btn-refresh:hover { background:var(--accent-dark); }
  .refresh-modal { display:none; position:fixed; inset:0; background:rgba(0,0,0,.7); z-index:1000; align-items:center; justify-content:center; }
  .refresh-modal.open { display:flex; }
  .refresh-box { background:#1a1c1f; border:1px solid var(--border-light); border-radius:12px; padding:28px 32px; max-width:420px; width:90%; }
  .refresh-box h3 { font-size:15px; font-weight:700; margin-bottom:16px; color:var(--text-primary); }
  .refresh-step { display:flex; gap:12px; align-items:flex-start; margin-bottom:12px; font-size:13px; color:var(--text-secondary); }
  .refresh-num { background:var(--accent); color:white; border-radius:50%; width:22px; height:22px; display:flex; align-items:center; justify-content:center; font-size:11px; font-weight:700; flex-shrink:0; }
  .refresh-box code { background:var(--bg-inner); border:1px solid var(--border-light); border-radius:4px; padding:2px 7px; font-family:monospace; font-size:12px; color:var(--orange); }
  .btn-reload { margin-top:20px; width:100%; padding:10px; background:var(--accent); color:white; border:none; border-radius:6px; font-size:13px; font-weight:700; cursor:pointer; transition:background .15s; }
  .btn-reload:hover { background:var(--accent-dark); }
  .btn-close-modal { position:absolute; top:12px; right:16px; background:none; border:none; color:var(--text-muted); font-size:18px; cursor:pointer; }

  .main-tabs { display:flex; background: #0d0e10; border-bottom: 1px solid var(--border); padding: 0 24px; gap:2px; }
  .main-tab { padding: 12px 20px; cursor:pointer; font-size:12px; font-weight:600; color:var(--text-secondary); border-bottom: 2px solid transparent; margin-bottom:-1px; transition: all .2s; letter-spacing:.3px; text-transform:uppercase; }
  .main-tab:hover { color: var(--text-primary); }
  .main-tab.active { color: var(--accent); border-bottom-color: var(--accent); }

  .toolbar { background: #0d0e10; padding: 10px 24px; display:flex; align-items:center; gap:8px; border-bottom:1px solid var(--border); flex-wrap:wrap; }
  .toolbar label { font-weight:600; color:var(--text-muted); font-size:11px; text-transform:uppercase; letter-spacing:.5px; }
  .group-btn { padding: 5px 14px; border:1px solid var(--border-light); border-radius:20px; background:transparent; color:var(--text-secondary); cursor:pointer; font-size:11px; font-weight:600; transition: all .15s; }
  .group-btn:hover { border-color: var(--accent); color: var(--accent); }
  .group-btn.active { background: var(--accent); border-color: var(--accent); color:white; }
  .expand-all-btn { margin-left:auto; padding:5px 14px; border:1px solid var(--border-light); border-radius:6px; background:transparent; cursor:pointer; font-size:11px; color:var(--text-secondary); transition:all .15s; }
  .expand-all-btn:hover { background:var(--bg-inner); color:var(--text-primary); }

  .content { padding: 20px 24px; }
  .tab-panel { display:none; }
  .tab-panel.active { display:block; }

  .group-section { margin-bottom: 28px; }
  .group-header { display:flex; align-items:center; gap:10px; margin-bottom:12px; }
  .group-title { font-size:13px; font-weight:700; color:var(--text-primary); text-transform:uppercase; letter-spacing:1px; border-left:3px solid var(--accent); padding-left:10px; }
  .group-badge { background:var(--bg-inner); color:var(--text-secondary); border:1px solid var(--border-light); border-radius:12px; padding:2px 10px; font-size:11px; font-weight:700; }

  .person-card { background:var(--bg-card); border-radius:8px; box-shadow:0 1px 4px rgba(0,0,0,.4); margin-bottom:8px; overflow:hidden; border:1px solid var(--border); transition:border-color .15s; }
  .person-card:hover { border-color: var(--border-light); }
  .person-header { display:flex; align-items:center; gap:12px; padding:12px 16px; cursor:pointer; user-select:none; transition: background .15s; }
  .person-header:hover { background: var(--bg-hover); }
  .chevron { transition: transform .25s; color:var(--text-muted); font-size:10px; }
  .person-card.open .chevron { transform: rotate(90deg); }
  .person-name { font-weight:700; font-size:13px; flex:1; color:var(--text-primary); }
  .person-grupo { font-size:11px; color:var(--text-muted); }

  .pill { display:inline-block; border-radius:20px; padding:2px 10px; font-size:10px; font-weight:700; letter-spacing:.3px; }
  .pill.green  { background:var(--green-bg); color:var(--green); border:1px solid rgba(34,197,94,.3); }
  .pill.orange { background:var(--orange-bg); color:var(--orange); border:1px solid rgba(245,158,11,.3); }
  .pill.red    { background:var(--red-bg); color:var(--red); border:1px solid rgba(239,68,68,.3); }
  .pill.gray   { background:var(--bg-inner); color:var(--text-secondary); border:1px solid var(--border-light); }
  .pill.purple { background:var(--purple-bg); color:var(--purple); border:1px solid rgba(168,85,247,.3); }

  .person-body { display:none; padding:0 16px 16px; border-top:1px solid var(--border); }
  .person-card.open .person-body { display:block; }

  .two-col { display:grid; grid-template-columns:1fr 1fr; gap:14px; margin-top:14px; }
  @media(max-width:700px){ .two-col { grid-template-columns:1fr; } }

  .section-box { background:var(--bg-inner); border:1px solid var(--border); border-radius:8px; padding:12px; }
  .section-box h4 { font-size:10px; font-weight:700; color:var(--text-muted); text-transform:uppercase; letter-spacing:.8px; margin-bottom:10px; padding-bottom:6px; border-bottom:1px solid var(--border); }

  .stat-row { display:flex; justify-content:space-between; flex-wrap:wrap; gap:8px; margin-bottom:10px; }
  .stat { text-align:center; }
  .stat .val { font-size:22px; font-weight:700; color:var(--text-primary); }
  .stat .lbl { font-size:10px; color:var(--text-muted); text-transform:uppercase; letter-spacing:.4px; }

  table { width:100%; border-collapse:collapse; font-size:12px; }
  th { background:#0d0e10; color:var(--text-secondary); padding:6px 8px; text-align:left; font-size:10px; text-transform:uppercase; letter-spacing:.5px; border-bottom:1px solid var(--border); }
  td { padding:5px 8px; border-bottom:1px solid var(--border); color:var(--text-primary); }
  tr:last-child td { border-bottom:none; }
  tr:hover td { background:rgba(255,255,255,.02); }
  tr.future td { color:var(--orange); font-style:italic; }
  tr.future td:first-child::before { content:"* "; }
  .td-ok    { background:var(--green-bg); color:var(--green); font-weight:700; text-align:center; }
  .td-largo { background:var(--orange-bg); color:var(--orange); font-weight:700; text-align:center; }
  .td-corto, .td-sin-break { background:var(--red-bg); color:var(--red); font-weight:700; text-align:center; }
  .td-muy-largo { background:rgba(239,68,68,.25); color:#ff6b6b; font-weight:700; text-align:center; }
  .no-data { color:var(--text-muted); font-style:italic; font-size:12px; padding:8px 0; }
  .note { font-size:11px; color:var(--text-muted); font-style:italic; margin-top:8px; }
  .progress-bar { height:6px; border-radius:3px; background:var(--border); margin-top:6px; overflow:hidden; }
  .progress-fill { height:100%; border-radius:3px; }
  .summary-chips { display:flex; flex-wrap:wrap; gap:6px; margin-top:8px; }
</style>
</head>
<body>
<header>
  <div class="header-logo-placeholder">Z</div>
  <div class="header-title">Reporte de Personal &mdash; Articulos Promocionales SA</div>
  <div class="header-sub">Periodo: 01/05/2026 &ndash; 29/05/2026 &nbsp;&#183;&nbsp; Generado: $generadoEn</div>
  <button class="btn-refresh" onclick="document.getElementById('refreshModal').classList.add('open')">&#8635; Actualizar</button>
</header>

<div class="refresh-modal" id="refreshModal" onclick="if(event.target===this)this.classList.remove('open')">
  <div class="refresh-box" style="position:relative">
    <button class="btn-close-modal" onclick="document.getElementById('refreshModal').classList.remove('open')">&#10005;</button>
    <h3>&#8635; Como actualizar el reporte</h3>
    <div class="refresh-step"><div class="refresh-num">1</div><div>Guardar los cambios en <code>Asistencia.xlsx</code> y cerrar Excel</div></div>
    <div class="refresh-step"><div class="refresh-num">2</div><div>Hacer doble clic en <code>Actualizar.bat</code> en la carpeta Asistencia</div></div>
    <div class="refresh-step"><div class="refresh-num">3</div><div>El reporte se va a abrir automaticamente con los datos actualizados</div></div>
    <div style="margin-top:16px;padding:10px 14px;background:rgba(245,158,11,.1);border-radius:6px;font-size:11px;color:var(--orange)">
      &#9432; Si ya ejecutaste <code>Actualizar.bat</code>, presiona el boton para recargar esta pagina:
    </div>
    <button class="btn-reload" onclick="location.reload(true)">&#8635; Recargar pagina ahora</button>
  </div>
</div>

<div class="main-tabs">
  <div class="main-tab active" onclick="switchTab('resumen')">&#128203; Resumen</div>
  <div class="main-tab" onclick="switchTab('asistencia')">&#128197; Asistencia</div>
  <div class="main-tab" onclick="switchTab('descansos')">&#9200; Descansos y Almuerzo</div>
  <div class="main-tab" onclick="switchTab('novedades')">&#9888; Novedades e Inconsistencias</div>
  <div class="main-tab" onclick="switchTab('horasextras')">&#9201; Horas Extras</div>
</div>

<div class="toolbar">
  <label>Filtrar grupo:</label>
  <button class="group-btn active" onclick="filterGroup('all',this)">Todos</button>
'@

foreach ($g in $grupos) {
    $n = ($allPersonas | Where-Object { $_.Grupo -eq $g }).Count
    $html += "  <button class=`"group-btn`" onclick=`"filterGroup('$g',this)`">$g ($n)</button>`n"
}

$html += @'
  <button class="expand-all-btn" onclick="toggleAll()">Expandir / Colapsar todo</button>
</div>

<div class="content">
'@

# ==========================================
# TAB RESUMEN
# ==========================================
$html += '<div id="tab-resumen" class="tab-panel active">'
$html += @'
<style>
  .resumen-table { width:100%; border-collapse:collapse; font-size:13px; }
  .resumen-table th { background:#0d0e10; color:var(--text-secondary); padding:10px 14px; text-align:left; font-size:10px; text-transform:uppercase; letter-spacing:.6px; border-bottom:2px solid var(--border); }
  .resumen-table th.tc { text-align:center; }
  .resumen-table td { padding:9px 14px; border-bottom:1px solid var(--border); color:var(--text-primary); vertical-align:middle; }
  .resumen-table tr:last-child td { border-bottom:none; }
  .resumen-table .tc { text-align:center; }
  .resumen-table .td-name { font-weight:700; }
  .resumen-table .td-grupo { color:var(--text-muted); font-size:11px; }
  .resumen-group-row td { background:#0d0e10; color:var(--text-secondary); font-size:10px; font-weight:700; text-transform:uppercase; letter-spacing:1px; padding:7px 14px; border-left:3px solid var(--accent); }
  .check-ok  { color:var(--green);  font-size:16px; font-weight:700; }
  .check-bad { color:var(--red);    font-size:16px; font-weight:700; }
  .num-circle { display:inline-flex; align-items:center; justify-content:center; width:22px; height:22px; border-radius:50%; background:var(--bg-inner); color:var(--text-muted); font-size:11px; font-weight:700; border:1px solid var(--border-light); }
  .incons-ok  { color:var(--green); font-weight:700; }
  .incons-bad { color:var(--red);   font-weight:700; }
  .rsm-row { cursor:pointer; user-select:none; }
  .rsm-row:hover td { background:rgba(255,255,255,.05) !important; }
  .rsm-chevron { display:inline-block; transition:transform .22s; margin-right:7px; font-size:10px; color:var(--text-muted); }
  .rsm-row.open .rsm-chevron { transform:rotate(90deg); }
  .resumen-detail td { background:var(--bg-inner) !important; padding:14px 16px !important; border-bottom:2px solid var(--border) !important; }
  .rsm-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(190px,1fr)); gap:10px; }
  .rsm-block { background:var(--bg-card); border:1px solid var(--border); border-radius:6px; padding:10px 12px; }
  .rsm-block h6 { font-size:10px; font-weight:700; text-transform:uppercase; letter-spacing:.5px; margin-bottom:7px; padding-bottom:5px; border-bottom:1px solid var(--border); }
  .rsm-item { display:flex; justify-content:space-between; align-items:center; font-size:12px; padding:3px 0; color:var(--text-secondary); border-bottom:1px solid var(--border); }
  .rsm-item:last-child { border-bottom:none; }
  .rsm-item .rsm-fecha { color:var(--text-primary); }
  .rsm-item .rsm-val { font-weight:700; }
</style>
'@

$html += "<div style='background:rgba(232,56,61,.08);border:1px solid rgba(232,56,61,.25);border-radius:8px;padding:10px 16px;margin-bottom:16px;font-size:12px;color:var(--text-secondary);'>"
$html += "&#128203;&nbsp; <strong>Resumen del periodo 01/05 &ndash; 14/05/2026</strong> &nbsp;&#183;&nbsp; Las inconsistencias incluyen atrasos de entrada, salidas anticipadas, demoras en retorno de break y breaks excedidos. Umbral de alerta: <strong style='color:var(--accent)'>15 minutos</strong>.</div>"

$html += "<div class='section-box' style='padding:0;overflow:hidden;margin-bottom:0'>"
$html += "<table class='resumen-table'>"
$html += "<thead><tr>"
$html += "<th style='width:40px'>#</th>"
$html += "<th>Operario</th>"
$html += "<th>Grupo</th>"
$html += "<th class='tc'>Dias Lab.</th>"
$html += "<th class='tc'>Presentes</th>"
$html += "<th class='tc'>100% Asist.</th>"
$html += "<th class='tc'>Inconsistencias</th>"
$html += "<th class='tc'>Estado</th>"
$html += "</tr></thead><tbody>"

$rowNum = 0
foreach ($g in $grupos) {
    $personasGrupo = $allPersonas | Where-Object { $_.Grupo -eq $g }
    $html += "<tr class='resumen-group-row' data-group='$g'><td colspan='8'>$g &mdash; $($personasGrupo.Count) persona$(if($personasGrupo.Count -gt 1){'s'})</td></tr>"
    foreach ($per in $personasGrupo) {
        $rowNum++
        $rsmId        = "$($per.Apellidos)_$($per.Nombre)" -replace '\s','_' -replace '[^a-zA-Z0-9_]','X'
        $realAusRsm   = [int]($per.AusReales.Count)
        $asistOK      = $realAusRsm -eq 0
        $inconsMin    = [int]$per.TotalInconsistenciaMin
        $inconsStr    = MinToStr $inconsMin
        $estadoOK     = $asistOK -and $inconsMin -le 15
        $asistSymbol  = if ($asistOK)  { "<span class='check-ok'>&#10003;</span>"  } else { "<span class='check-bad'>&#10007;</span>" }
        $estadoSymbol = if ($estadoOK) { "<span class='check-ok'>&#10003;</span>"  } else { "<span class='check-bad'>&#10007;</span>" }
        $inconsClass  = if ($estadoOK) { "incons-ok" } else { "incons-bad" }
        $presColor    = if ($per.nAusente -eq 0) { "color:var(--green)" } elseif ($realAusRsm -gt 0) { "color:var(--red)" } else { "color:var(--orange)" }

        # Fila principal (clickeable)
        $html += "<tr id='rsm_$rsmId' class='rsm-row' data-group='$g' onclick='toggleRsmDetail(`"$rsmId`")'>"
        $html += "<td class='tc'><span class='num-circle'>$rowNum</span></td>"
        $html += "<td class='td-name'><span class='rsm-chevron'>&#9654;</span>$($per.Apellidos), $($per.Nombre)</td>"
        $html += "<td class='td-grupo'>$($per.Grupo)</td>"
        $html += "<td class='tc'>$($per.nLab)</td>"
        $html += "<td class='tc' style='$presColor'>$($per.nPresente)</td>"
        $html += "<td class='tc'>$asistSymbol</td>"
        $html += "<td class='tc $inconsClass'>$inconsStr</td>"
        $html += "<td class='tc'>$estadoSymbol</td>"
        $html += "</tr>"

        # Fila detalle (colapsada por defecto)
        $html += "<tr id='detail_$rsmId' class='resumen-detail' data-group='$g' style='display:none'>"
        $html += "<td colspan='8'><div class='rsm-grid'>"

        if ($per.Atrasos.Count -gt 0) {
            $html += "<div class='rsm-block'>"
            $html += "<h6 style='color:var(--orange)'>&#128336; Atrasos de entrada &mdash; $(MinToStr $per.TotalAtrasoMin) total</h6>"
            foreach ($a in $per.Atrasos) {
                $html += "<div class='rsm-item'><span class='rsm-fecha'>$($a.Fecha)</span><span class='rsm-val' style='color:var(--orange)'>$($a.AtrasoStr)</span></div>"
            }
            $html += "</div>"
        }

        if ($per.Adelantos.Count -gt 0) {
            $html += "<div class='rsm-block'>"
            $html += "<h6 style='color:var(--purple)'>&#9194; Salidas anticipadas &mdash; $(MinToStr $per.TotalAdelantoMin) total</h6>"
            foreach ($a in $per.Adelantos) {
                $html += "<div class='rsm-item'><span class='rsm-fecha'>$($a.Fecha)</span><span class='rsm-val' style='color:var(--purple)'>$($a.AdelantoStr)</span></div>"
            }
            $html += "</div>"
        }

        if ($per.DemoraBreak.Count -gt 0) {
            $html += "<div class='rsm-block'>"
            $html += "<h6 style='color:var(--orange)'>&#9203; Demoras en retorno de break</h6>"
            foreach ($d in $per.DemoraBreak) {
                $html += "<div class='rsm-item'><span class='rsm-fecha'>$($d.Fecha)</span><span class='rsm-val' style='color:var(--orange)'>$($d.DemoraStr)</span></div>"
            }
            $html += "</div>"
        }

        if ($per.BreakExcedido.Count -gt 0) {
            $html += "<div class='rsm-block'>"
            $html += "<h6 style='color:var(--orange)'>&#9200; Breaks excedidos (&gt;65 min)</h6>"
            foreach ($b in $per.BreakExcedido) {
                $excesoBE = $b.BreakMin - 60
                $html += "<div class='rsm-item'><span class='rsm-fecha'>$($b.Fecha)</span><span class='rsm-val' style='color:var(--orange)'>$($b.BreakStr) (+${excesoBE}m)</span></div>"
            }
            $html += "</div>"
        }

        if ($per.AusReales.Count -gt 0) {
            $html += "<div class='rsm-block'>"
            $html += "<h6 style='color:var(--red)'>&#128683; Ausencias ($($per.AusReales.Count))</h6>"
            foreach ($a in $per.AusReales) {
                $html += "<div class='rsm-item'><span class='rsm-fecha'>$($a.Fecha)</span><span class='rsm-val' style='color:var(--red)'>Ausente</span></div>"
            }
            $html += "</div>"
        }

        if ($inconsMin -eq 0 -and $per.AusReales.Count -eq 0) {
            $html += "<div style='color:var(--green);font-size:12px;font-style:italic;padding:4px 0'>&#10003; Sin inconsistencias ni ausencias en el periodo.</div>"
        }

        $html += "</div></td></tr>"
    }
}

$html += "</tbody></table></div>"
$html += '</div>'  # tab-resumen

# ---- TAB ASISTENCIA ----
$html += '<div id="tab-asistencia" class="tab-panel">'

foreach ($g in $grupos) {
    $personasGrupo = $allPersonas | Where-Object { $_.Grupo -eq $g }
    $html += "<div class='group-section' data-group='$g'>"
    $html += "<div class='group-header'><div class='group-title'>$g</div><div class='group-badge'>$($personasGrupo.Count) personas</div></div>"

    foreach ($per in $personasGrupo) {
        $pctColor = if ($per.Pct -ge 90) { "green" } elseif ($per.Pct -ge 70) { "orange" } else { "red" }
        $pctFill = if ($per.Pct -ge 90) { "var(--green)" } elseif ($per.Pct -ge 70) { "var(--orange)" } else { "var(--red)" }
        $cardId = "$($per.Apellidos)_$($per.Nombre)" -replace '\s','_' -replace '[^a-zA-Z0-9_]','X'

        $html += "<div class='person-card' id='a_$cardId' data-group='$g'>"
        $html += "<div class='person-header' onclick='toggle(`"a_$cardId`")'>"
        $html += "<span class='chevron'>&#9654;</span>"
        $html += "<span class='person-name'>$($per.Apellidos), $($per.Nombre)</span>"
        $html += "<span class='person-grupo'>$($per.Grupo)</span>"
        $html += "<span class='pill $pctColor'>$($per.Pct)% asistencia</span>"
        $html += "<span class='pill gray'>$($per.nPresente)/$($per.nLab) dias</span>"
        if ($per.nAusente -gt 0) {
            $futCount = [int]($per.AusDetails | Where-Object { $_.Futuro }).Count
            $realCount = $per.nAusente - $futCount
            if ($realCount -gt 0) { $html += "<span class='pill red'>$realCount ausencias reales</span>" }
            if ($futCount -gt 0) { $html += "<span class='pill orange'>$futCount pendientes</span>" }
        }
        $html += "</div>"  # person-header

        $html += "<div class='person-body'>"
        $html += "<div class='two-col'>"

        # LEFT: stats
        $html += "<div class='section-box'>"
        $html += "<h4>&#128200; Resumen del Periodo</h4>"
        $html += "<div class='stat-row'>"
        $html += "<div class='stat'><div class='val' style='color:var(--blue-dark)'>$($per.nLab)</div><div class='lbl'>Dias Laborales</div></div>"
        $html += "<div class='stat'><div class='val' style='color:var(--green)'>$($per.nPresente)</div><div class='lbl'>Presentes</div></div>"
        $realAus = [int](($per.AusDetails | Where-Object { -not $_.Futuro }).Count)
        $futAus  = [int](($per.AusDetails | Where-Object { $_.Futuro }).Count)
        $html += "<div class='stat'><div class='val' style='color:var(--red)'>$realAus</div><div class='lbl'>Ausentes Reales</div></div>"
        $html += "<div class='stat'><div class='val' style='color:var(--orange)'>$futAus</div><div class='lbl'>Dias Futuros</div></div>"
        $html += "</div>"
        $html += "<div style='margin-top:6px;font-size:11px;color:var(--gray)'>$($per.Pct)% asistencia</div>"
        $html += "<div class='progress-bar'><div class='progress-fill' style='width:$($per.Pct)%;background:$pctFill'></div></div>"
        if ($per.nFeriado -gt 0) { $html += "<div class='note'>Feriados en el periodo: $($per.nFeriado)</div>" }
        if ($per.nDescanso -gt 0) { $html += "<div class='note'>Descansos en semana (turno rotativo): $($per.nDescanso)</div>" }
        $html += "</div>"  # section-box

        # RIGHT: ausencias
        $html += "<div class='section-box'>"
        $html += "<h4>&#128680; Dias Ausentes</h4>"
        if ($per.nAusente -eq 0) {
            $html += "<div class='no-data' style='color:var(--green)'>Sin ausencias en el periodo</div>"
        } else {
            $html += "<table><tr><th>Fecha</th><th>Turno</th><th>Estado</th></tr>"
            foreach ($a in $per.AusDetails) {
                if ($a.Justificado) {
                    $html += "<tr><td>$($a.Fecha)</td><td>$($a.Turno)</td><td style='color:var(--green);font-weight:700'>&#9992; $($a.Motivo)</td></tr>"
                } elseif ($a.Futuro) {
                    $html += "<tr class='future'><td>$($a.Fecha)</td><td>$($a.Turno)</td><td>Pendiente</td></tr>"
                } else {
                    $html += "<tr><td>$($a.Fecha)</td><td>$($a.Turno)</td><td style='color:var(--red);font-weight:700'>AUSENTE</td></tr>"
                }
            }
            $html += "</table>"
        }
        $html += "</div>"  # section-box
        $html += "</div></div></div>"  # two-col, person-body, person-card
    }
    $html += "</div>"  # group-section
}

$html += '</div>'  # tab-asistencia

# ---- TAB DESCANSOS ----
$html += '<div id="tab-descansos" class="tab-panel">'
$html += "<div style='background:rgba(245,158,11,.1);border:1px solid rgba(245,158,11,.3);border-radius:8px;padding:10px 16px;margin-bottom:16px;font-size:12px;color:var(--orange);'>"
$html += "&#9432;&nbsp; <strong>Descanso esperado: 15 min</strong> &nbsp;|&nbsp; <strong>Almuerzo esperado: 45 min</strong> &nbsp;|&nbsp; Total: 60 min. "
$html += "El sistema registra UN evento de break por dia (salida S1 y retorno E2). Si el break es &lt;22 min = solo descanso; 40-65 min = almuerzo; &gt;65 min = se paso de tiempo.</div>"

foreach ($g in $grupos) {
    $personasGrupo = $allPersonas | Where-Object { $_.Grupo -eq $g }
    $html += "<div class='group-section' data-group='$g'>"
    $html += "<div class='group-header'><div class='group-title'>$g</div><div class='group-badge'>$($personasGrupo.Count) personas</div></div>"

    foreach ($per in $personasGrupo) {
        $cardId2 = "b_$("$($per.Apellidos)_$($per.Nombre)" -replace '\s','_' -replace '[^a-zA-Z0-9_]','X')"
        $hasData = $per.nConBreak -gt 0

        $breakStatus = if (-not $hasData) { "gray" }
                       elseif ($per.nBreakLargo -gt 0 -or $per.nBreakCorto -gt 0) { "orange" }
                       else { "green" }

        $html += "<div class='person-card' id='$cardId2' data-group='$g'>"
        $html += "<div class='person-header' onclick='toggle(`"$cardId2`")'>"
        $html += "<span class='chevron'>&#9654;</span>"
        $html += "<span class='person-name'>$($per.Apellidos), $($per.Nombre)</span>"
        $html += "<span class='person-grupo'>$($per.Grupo)</span>"

        if (-not $hasData) {
            $html += "<span class='pill gray'>Sin datos de break</span>"
        } else {
            $html += "<span class='pill $breakStatus'>Prom: $(MinToStr $per.AvgBreak)</span>"
            if ($per.nBreakOK -gt 0) { $html += "<span class='pill green'>$($per.nBreakOK) OK</span>" }
            if ($per.nBreakLargo -gt 0) { $html += "<span class='pill orange'>$($per.nBreakLargo) largos</span>" }
            if ($per.nBreakCorto -gt 0) { $html += "<span class='pill red'>$($per.nBreakCorto) cortos</span>" }
            if ($per.nSinBreakData -gt 0) { $html += "<span class='pill gray'>$($per.nSinBreakData) sin dato</span>" }
        }
        $html += "</div>"  # person-header

        $html += "<div class='person-body'>"
        $html += "<div class='two-col'>"

        # LEFT: break stats
        $html += "<div class='section-box'>"
        $html += "<h4>&#9200; Resumen de Breaks</h4>"
        if (-not $hasData) {
            $html += "<div class='no-data'>El sistema no registra salida/retorno de break para esta persona.</div>"
            $html += "<div class='note'>Asistio $($per.nPresente) dias. Los dias trabajados no tienen datos de break disponibles en el reporte fuente.</div>"
        } else {
            $html += "<div class='stat-row'>"
            $html += "<div class='stat'><div class='val'>$(MinToStr $per.AvgBreak)</div><div class='lbl'>Promedio</div></div>"
            $html += "<div class='stat'><div class='val'>$(MinToStr $per.MinBreak)</div><div class='lbl'>Minimo</div></div>"
            $html += "<div class='stat'><div class='val'>$(MinToStr $per.MaxBreak)</div><div class='lbl'>Maximo</div></div>"
            $html += "<div class='stat'><div class='val'>$($per.nConBreak)</div><div class='lbl'>Dias con dato</div></div>"
            $html += "</div>"
            $html += "<div class='summary-chips'>"
            if ($per.nBreakOK -gt 0) { $html += "<span class='pill green'>$($per.nBreakOK) dentro del tiempo</span>" }
            if ($per.nBreakLargo -gt 0) { $html += "<span class='pill orange'>$($per.nBreakLargo) se paso del tiempo</span>" }
            if ($per.nBreakCorto -gt 0) { $html += "<span class='pill red'>$($per.nBreakCorto) break corto</span>" }
            if ($per.nSinBreakData -gt 0) { $html += "<span class='pill gray'>$($per.nSinBreakData) dias sin dato de break</span>" }
            $html += "</div>"

            # Classify what type of break they mainly take
            $nDescansoTipo = [int](@($per.BreakDays | Where-Object { $_.Tipo -eq 'Descanso' }).Count)
            $nAlmuerzoTipo = [int](@($per.BreakDays | Where-Object { $_.Tipo -eq 'Almuerzo' }).Count)
            if ($nDescansoTipo -gt 0 -and $nAlmuerzoTipo -eq 0) {
                $html += "<div class='note' style='color:var(--orange)'>&#9888; Solo se registra el descanso de 15 min. El almuerzo de 45 min no aparece en el sistema.</div>"
            } elseif ($nDescansoTipo -gt 0 -and $nAlmuerzoTipo -gt 0) {
                $html += "<div class='note'>Algunos dias solo descanso ($nDescansoTipo), otros con almuerzo ($nAlmuerzoTipo).</div>"
            }
        }
        $html += "</div>"  # section-box

        # RIGHT: break detail table
        $html += "<div class='section-box'>"
        $html += "<h4>&#128203; Detalle por Dia</h4>"
        if ($per.BreakDays.Count -eq 0) {
            $html += "<div class='no-data'>Sin registros de break disponibles.</div>"
        } else {
            $html += "<table><tr><th>Fecha</th><th>Salida</th><th>Retorno</th><th>Duracion</th><th>Tipo</th><th>Estado</th><th>Dif.</th></tr>"
            foreach ($b in $per.BreakDays) {
                $tdCls = "td-$($b.Estado)"
                $diffColor = if ($b.Diff -gt 0) { "color:var(--orange)" } elseif ($b.Diff -lt 0) { "color:var(--red)" } else { "color:var(--green)" }
                $html += "<tr><td>$($b.Fecha)</td><td>$($b.S1)</td><td>$($b.E2)</td>"
                $html += "<td class='$tdCls'>$($b.BreakStr)</td>"
                $html += "<td>$($b.Tipo)</td>"
                $html += "<td class='$tdCls'>$($b.Estado)</td>"
                $html += "<td style='$diffColor;font-weight:600;text-align:center'>$($b.DiffStr)</td></tr>"
            }
            $html += "</table>"
        }
        $html += "</div>"  # section-box
        $html += "</div></div></div>"  # two-col, person-body, person-card
    }
    $html += "</div>"  # group-section
}

$html += '</div>'  # tab-descansos

# ==========================================
# TAB NOVEDADES
# ==========================================
$html += '<div id="tab-novedades" class="tab-panel">'
$html += @'
<style>
  .nov-note { background:rgba(245,158,11,.1);border-left:3px solid var(--orange);padding:10px 14px;border-radius:6px;margin-bottom:16px;font-size:12px;color:var(--orange); }
  .nov-persona { background:var(--bg-card);border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,.4);margin-bottom:8px;overflow:hidden;border:1px solid var(--border);transition:border-color .15s; }
  .nov-persona:hover { border-color:var(--border-light); }
  .nov-persona.limpio .nov-header { border-left:3px solid var(--green); }
  .nov-persona.con-nov .nov-header { border-left:3px solid var(--accent); }
  .nov-header { padding:12px 16px;cursor:pointer;display:flex;align-items:center;gap:12px;user-select:none;background:var(--bg-card);transition:background .15s; }
  .nov-header:hover { background:var(--bg-hover); }
  .nov-body { display:none; padding:0 16px 16px; border-top:1px solid var(--border); }
  .nov-persona.open .nov-body { display:block; }
  .nov-persona.open .chevron { transform:rotate(90deg); }
  .nov-name { font-weight:700;font-size:13px;flex:1;color:var(--text-primary); }
  .nov-grupo { font-size:11px;color:var(--text-muted); }
  .nov-grid { display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:14px;margin-top:14px; }
  .nov-block { background:var(--bg-inner);border:1px solid var(--border);border-radius:8px;padding:12px; }
  .nov-block h5 { font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px;padding-bottom:6px;border-bottom:1px solid var(--border); }
  .nov-block h5.red-h  { color:var(--red); }
  .nov-block h5.ora-h  { color:var(--orange); }
  .nov-block h5.pur-h  { color:var(--purple); }
  .nov-block h5.blu-h  { color:var(--text-secondary); }
  .nov-total { font-size:18px;font-weight:700;margin-bottom:4px; }
  .nov-sub  { font-size:10px;color:var(--text-muted); }
  .nov-table { width:100%;border-collapse:collapse;font-size:12px;margin-top:6px; }
  .nov-table th { background:#0d0e10;color:var(--text-secondary);padding:5px 8px;text-align:left;font-size:10px;text-transform:uppercase;letter-spacing:.5px;border-bottom:1px solid var(--border); }
  .nov-table td { padding:5px 8px;border-bottom:1px solid var(--border);color:var(--text-primary); }
  .nov-table tr:last-child td { border-bottom:none; }
  .nov-table tr:hover td { background:rgba(255,255,255,.02); }
  .highlight-red { background:var(--red-bg);color:var(--red);font-weight:700;padding:2px 6px;border-radius:4px; }
  .highlight-ora { background:var(--orange-bg);color:var(--orange);font-weight:700;padding:2px 6px;border-radius:4px; }
  .highlight-pur { background:var(--purple-bg);color:var(--purple);font-weight:700;padding:2px 6px;border-radius:4px; }
  .no-nov { color:var(--green);font-size:12px;font-style:italic;margin-top:6px; }
</style>
'@

$html += "<div class='nov-note'>&#9432; Solo se muestran <strong>dias ocurridos hasta el 14/05/2026</strong>. Dias futuros no se consideran como novedad.</div>"

foreach ($g in $grupos) {
    $personasGrupo = $allPersonas | Where-Object { $_.Grupo -eq $g }
    $html += "<div class='group-section' data-group='$g'>"
    $html += "<div class='group-header'><div class='group-title'>$g</div><div class='group-badge'>$($personasGrupo.Count) personas</div></div>"

    foreach ($per in $personasGrupo) {
        $novClass = if ($per.nNovedades -gt 0) { "con-nov" } else { "limpio" }
        $novId = "n_$("$($per.Apellidos)_$($per.Nombre)" -replace '\s','_' -replace '[^a-zA-Z0-9_]','X')"

        $html += "<div class='nov-persona $novClass' id='$novId' data-group='$g'>"
        $html += "<div class='nov-header' onclick='toggle(`"$novId`")'>"
        $html += "<span class='chevron' style='font-size:11px;color:var(--gray)'>&#9654;</span>"
        $html += "<span class='nov-name'>$($per.Apellidos), $($per.Nombre)</span>"
        $html += "<span class='nov-grupo'>$($per.Grupo)</span>"

        if ($per.nNovedades -eq 0) {
            $html += "<span class='pill green'>Sin novedades</span>"
        } else {
            if ($per.AusReales.Count -gt 0)    { $html += "<span class='pill red'>$($per.AusReales.Count) ausencia$(if($per.AusReales.Count-gt 1){'s'})</span>" }
            if ($per.Atrasos.Count -gt 0)       { $html += "<span class='pill orange'>$($per.Atrasos.Count) atraso$(if($per.Atrasos.Count-gt 1){'s'}) ($(MinToStr $per.TotalAtrasoMin))</span>" }
            if ($per.Adelantos.Count -gt 0)     { $html += "<span class='pill purple'>$($per.Adelantos.Count) salida$(if($per.Adelantos.Count-gt 1){'s'}) anticipada$(if($per.Adelantos.Count-gt 1){'s'})</span>" }
            if ($per.BreakExcedido.Count -gt 0) { $html += "<span class='pill orange'>$($per.BreakExcedido.Count) break excedido</span>" }
            if ($per.DemoraBreak.Count -gt 0)   { $html += "<span class='pill orange'>$($per.DemoraBreak.Count) demora en retorno</span>" }
        }
        $html += "</div>"  # nov-header

        $html += "<div class='nov-body'><div class='nov-grid'>"

        # JUSTIFICACIONES / VIAJES
        if ($per.JustifPersona.Count -gt 0) {
            $html += "<div class='nov-block'>"
            $html += "<h5 style='color:var(--green)'>&#9992; Justificaciones ($($per.JustifPersona.Count))</h5>"
            foreach ($j in $per.JustifPersona) {
                $diasJustif = @($per.AusDetails | Where-Object { $_.Justificado -and $_.Motivo -eq $j.Motivo })
                $html += "<div class='nov-total' style='color:var(--green)'>$($j.Motivo)</div>"
                $html += "<div class='nov-sub'>$($j.FechaDesde.ToString('dd/MM')) al $($j.FechaHasta.ToString('dd/MM/yyyy'))</div>"
                if ($diasJustif.Count -gt 0) {
                    $html += "<table class='nov-table' style='margin-top:8px'><tr><th>Fecha</th><th>Turno</th><th>Motivo</th></tr>"
                    foreach ($d in $diasJustif) {
                        $html += "<tr><td>$($d.Fecha)</td><td>$($d.Turno)</td><td style='color:var(--green)'>&#9992; $($d.Motivo)</td></tr>"
                    }
                    $html += "</table>"
                } else {
                    $html += "<div style='font-size:11px;color:var(--text-muted);margin-top:4px'>No hay dias sin fichada en ese rango.</div>"
                }
            }
            $html += "</div>"
        }

        # AUSENCIAS
        $html += "<div class='nov-block'>"
        $html += "<h5 class='red-h'>&#128683; Ausencias ($($per.AusReales.Count))</h5>"
        if ($per.AusReales.Count -eq 0) {
            $html += "<div class='no-nov'>Sin ausencias en el periodo</div>"
        } else {
            $html += "<div class='nov-total' style='color:var(--red)'>$($per.AusReales.Count)</div><div class='nov-sub'>dias sin presentarse</div>"
            $html += "<table class='nov-table'><tr><th>Fecha</th><th>Turno</th></tr>"
            foreach ($a in $per.AusReales) { $html += "<tr><td>$($a.Fecha)</td><td>$($a.Turno)</td></tr>" }
            $html += "</table>"
        }
        $html += "</div>"

        # ATRASOS
        $html += "<div class='nov-block'>"
        $html += "<h5 class='ora-h'>&#128336; Atrasos en entrada ($($per.Atrasos.Count) dias)</h5>"
        if ($per.Atrasos.Count -eq 0) {
            $html += "<div class='no-nov'>Sin atrasos registrados</div>"
        } else {
            $html += "<div class='nov-total' style='color:var(--orange)'>$(MinToStr $per.TotalAtrasoMin)</div>"
            $html += "<div class='nov-sub'>total acumulado en $($per.Atrasos.Count) dia$(if($per.Atrasos.Count-gt 1){'s'})</div>"
            $html += "<table class='nov-table'><tr><th>Fecha</th><th>Turno</th><th>Entro</th><th>Demora</th></tr>"
            foreach ($a in $per.Atrasos) {
                $cls = if ($a.AtrasoMin -ge 60) { "highlight-red" } elseif ($a.AtrasoMin -ge 15) { "highlight-ora" } else { "" }
                $html += "<tr><td>$($a.Fecha)</td><td>$($a.Turno)</td><td>$($a.Entro)</td><td><span class='$cls'>$($a.AtrasoStr)</span></td></tr>"
            }
            $html += "</table>"
        }
        $html += "</div>"

        # SALIDAS ANTICIPADAS
        $html += "<div class='nov-block'>"
        $html += "<h5 class='pur-h'>&#9194; Salidas anticipadas ($($per.Adelantos.Count) dias)</h5>"
        if ($per.Adelantos.Count -eq 0) {
            $html += "<div class='no-nov'>Sin salidas anticipadas</div>"
        } else {
            $html += "<div class='nov-total' style='color:var(--purple)'>$(MinToStr $per.TotalAdelantoMin)</div>"
            $html += "<div class='nov-sub'>tiempo anticipado en $($per.Adelantos.Count) dia$(if($per.Adelantos.Count-gt 1){'s'})</div>"
            $html += "<table class='nov-table'><tr><th>Fecha</th><th>Turno</th><th>Se fue antes</th></tr>"
            foreach ($a in $per.Adelantos) {
                $cls = if ($a.AdelantoMin -ge 60) { "highlight-red" } elseif ($a.AdelantoMin -ge 30) { "highlight-ora" } else { "highlight-pur" }
                $html += "<tr><td>$($a.Fecha)</td><td>$($a.Turno)</td><td><span class='$cls'>$($a.AdelantoStr)</span></td></tr>"
            }
            $html += "</table>"
        }
        $html += "</div>"

        # BREAKS IRREGULARES
        $nBreakIrreg = $per.BreakExcedido.Count + $per.DemoraBreak.Count
        $html += "<div class='nov-block'>"
        $html += "<h5 class='ora-h'>&#9203; Irregularidades en break ($nBreakIrreg)</h5>"
        if ($nBreakIrreg -eq 0) {
            $html += "<div class='no-nov'>Sin irregularidades en breaks reales</div>"
        } else {
            if ($per.BreakExcedido.Count -gt 0) {
                $html += "<div style='font-size:11px;font-weight:700;color:var(--orange);margin-bottom:4px'>Break largo (mas de 65 min):</div>"
                $html += "<table class='nov-table'><tr><th>Fecha</th><th>Salida</th><th>Retorno</th><th>Duracion</th><th>Exceso</th></tr>"
                foreach ($b in $per.BreakExcedido) {
                    $exceso = $b.BreakMin - 60
                    $html += "<tr><td>$($b.Fecha)</td><td>$($b.S1)</td><td>$($b.E2)</td><td><span class='highlight-ora'>$($b.BreakStr)</span></td><td><span class='highlight-ora'>+${exceso}m</span></td></tr>"
                }
                $html += "</table>"
            }
            if ($per.DemoraBreak.Count -gt 0) {
                $html += "<div style='font-size:11px;font-weight:700;color:var(--orange);margin-top:8px;margin-bottom:4px'>Demora en retorno del break:</div>"
                $html += "<table class='nov-table'><tr><th>Fecha</th><th>Retorno</th><th>Demora</th></tr>"
                foreach ($d in $per.DemoraBreak) {
                    $html += "<tr><td>$($d.Fecha)</td><td>$($d.Retorno)</td><td><span class='highlight-ora'>$($d.DemoraStr)</span></td></tr>"
                }
                $html += "</table>"
            }
        }
        $html += "</div>"

        $html += "</div></div></div>"  # nov-grid, nov-body, nov-persona
    }
    $html += "</div>"  # group-section
}

$html += '</div>'  # tab-novedades

# ==========================================
# TAB HORAS EXTRAS
# ==========================================
$html += '<div id="tab-horasextras" class="tab-panel">'
$html += "<div style='background:rgba(245,158,11,.1);border:1px solid rgba(245,158,11,.3);border-radius:8px;padding:10px 16px;margin-bottom:16px;font-size:12px;color:var(--orange);'>"
$html += "&#9432;&nbsp; <strong>Lun-Vie:</strong> horas que exceden el fin de turno = 50%. &nbsp;|&nbsp; <strong>Sabado hasta 13:00:</strong> 50% &nbsp;|&nbsp; <strong>Sabado despues 13:00:</strong> 100%. &nbsp;|&nbsp; Los viernes el turno termina 1 hora antes.</div>"

foreach ($g in $grupos) {
    $personasGrupo = $allPersonas | Where-Object { $_.Grupo -eq $g }
    $html += "<div class='group-section' data-group='$g'>"
    $html += "<div class='group-header'><div class='group-title'>$g</div><div class='group-badge'>$($personasGrupo.Count) personas</div></div>"

    foreach ($per in $personasGrupo) {
        $htId    = "h_$("$($per.Apellidos)_$($per.Nombre)" -replace '\s','_' -replace '[^a-zA-Z0-9_]','X')"
        $hasHT   = $per.TotalOTMin -gt 0
        $htClass = if ($hasHT) { "con-nov" } else { "limpio" }

        $html += "<div class='nov-persona $htClass' id='$htId' data-group='$g'>"
        $html += "<div class='nov-header' onclick='toggle(`"$htId`")'>"
        $html += "<span class='chevron' style='font-size:11px;color:var(--text-muted)'>&#9654;</span>"
        $html += "<span class='nov-name'>$($per.Apellidos), $($per.Nombre)</span>"
        $html += "<span class='nov-grupo'>$($per.Grupo)</span>"
        if (-not $hasHT) {
            $html += "<span class='pill green'>Sin horas extras</span>"
        } else {
            $html += "<span class='pill orange'>Total: $(MinToStr $per.TotalOTMin)</span>"
            if ($per.TotalOT50  -gt 0) { $html += "<span class='pill gray'>50%: $(MinToStr $per.TotalOT50)</span>" }
            if ($per.TotalOT100 -gt 0) { $html += "<span class='pill red'>100%: $(MinToStr $per.TotalOT100)</span>" }
        }
        $html += "</div>"  # nov-header

        $html += "<div class='nov-body'>"
        if (-not $hasHT) {
            $html += "<div class='no-data' style='padding:12px 0'>Sin horas extras registradas en el periodo.</div>"
        } else {
            $html += "<div class='two-col' style='margin-top:14px'>"

            # LEFT: resumen
            $html += "<div class='section-box'>"
            $html += "<h4>&#128336; Resumen Horas Extras</h4>"
            $html += "<div class='stat-row'>"
            $html += "<div class='stat'><div class='val' style='color:var(--orange)'>$(MinToStr $per.TotalOTMin)</div><div class='lbl'>Total HE</div></div>"
            $html += "<div class='stat'><div class='val'>$(MinToStr $per.TotalOT50)</div><div class='lbl'>Al 50%</div></div>"
            $html += "<div class='stat'><div class='val' style='color:var(--red)'>$(MinToStr $per.TotalOT100)</div><div class='lbl'>Al 100%</div></div>"
            $html += "<div class='stat'><div class='val'>$($per.HTDays.Count)</div><div class='lbl'>Dias con HE</div></div>"
            $html += "</div>"
            $nSemana = [int](@($per.HTDays | Where-Object { $_.EsDia -eq 'semana' }).Count)
            $nSabado = [int](@($per.HTDays | Where-Object { $_.EsDia -eq 'sabado' }).Count)
            $html += "<div class='summary-chips'>"
            if ($nSemana -gt 0) { $html += "<span class='pill gray'>$nSemana dia$(if($nSemana -gt 1){'s'}) Lun-Vie</span>" }
            if ($nSabado -gt 0) { $html += "<span class='pill orange'>$nSabado sabado$(if($nSabado -gt 1){'s'})</span>" }
            $html += "</div></div>"  # summary-chips, section-box

            # RIGHT: detalle por dia
            $html += "<div class='section-box'>"
            $html += "<h4>&#128203; Detalle por Dia</h4>"
            $html += "<table><tr><th>Fecha</th><th>Turno</th><th>Entrada</th><th>Salida</th><th>Fin turno</th><th>HE</th><th>50%</th><th>100%</th></tr>"
            foreach ($d in $per.HTDays) {
                $rowStyle = if ($d.EsDia -eq 'sabado') { " style='border-left:2px solid var(--orange)'" } else { "" }
                $html += "<tr$rowStyle>"
                $html += "<td>$($d.Fecha)</td><td>$($d.Turno)</td><td>$($d.Entrada)</td><td>$($d.Salida)</td>"
                $html += "<td>$(if($d.EsDia -eq 'sabado'){'<span style=''color:var(--orange);font-weight:700''>Sabado</span>'}else{$d.TurnoFin})</td>"
                $html += "<td><span class='highlight-ora'>$(MinToStr $d.OTMin)</span></td>"
                $html += "<td>$(if($d.OT50Min -gt 0){"<span class='highlight-ora'>$(MinToStr $d.OT50Min)</span>"}else{'<span style=''color:var(--text-muted)''>-</span>'})</td>"
                $html += "<td>$(if($d.OT100Min -gt 0){"<span class='highlight-red'>$(MinToStr $d.OT100Min)</span>"}else{'<span style=''color:var(--text-muted)''>-</span>'})</td>"
                $html += "</tr>"
            }
            $html += "</table>"
            $html += "</div>"  # section-box

            $html += "</div>"  # two-col
        }
        $html += "</div></div>"  # nov-body, nov-persona
    }
    $html += "</div>"  # group-section
}

$html += '</div>'  # tab-horasextras
$html += '</div>'  # content

$html += @'
<script>
function toggle(id) {
    var card = document.getElementById(id);
    card.classList.toggle('open');
}
function toggleRsmDetail(id) {
    var detail = document.getElementById('detail_' + id);
    var row    = document.getElementById('rsm_'    + id);
    var isOpen = detail.style.display !== 'none';
    detail.style.display = isOpen ? 'none' : '';
    row.classList.toggle('open', !isOpen);
}
function switchTab(tab) {
    document.querySelectorAll('.tab-panel').forEach(function(el){ el.classList.remove('active'); });
    document.querySelectorAll('.main-tab').forEach(function(el){ el.classList.remove('active'); });
    document.getElementById('tab-' + tab).classList.add('active');
    event.target.classList.add('active');
    currentTab = tab;
}
var currentTab = 'resumen';
var allExpanded = false;
function filterGroup(group, btn) {
    document.querySelectorAll('.group-btn').forEach(function(b){ b.classList.remove('active'); });
    btn.classList.add('active');
    document.querySelectorAll('.group-section').forEach(function(s){
        s.style.display = (group === 'all' || s.dataset.group === group) ? '' : 'none';
    });
    document.querySelectorAll('tr[data-group]').forEach(function(r){
        r.style.display = (group === 'all' || r.dataset.group === group) ? '' : 'none';
    });
}
function toggleAll() {
    allExpanded = !allExpanded;
    var map = { asistencia:'a_', descansos:'b_', novedades:'n_', horasextras:'h_' };
    var prefix = map[currentTab] || 'a_';
    var sel = currentTab === 'asistencia' || currentTab === 'descansos'
        ? '.person-card[id^="' + prefix + '"]'
        : '.nov-persona[id^="' + prefix + '"]';
    document.querySelectorAll(sel).forEach(function(c){
        if (c.style.display !== 'none') {
            if (allExpanded) c.classList.add('open'); else c.classList.remove('open');
        }
    });
}
</script>
</body>
</html>
'@

$outPath = "C:\Users\bchevasco\OneDrive - Articulos Promocionales SA\Escritorio\Asistencia\Reporte_Personal.html"
[System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
Write-Host "LISTO: $outPath"

