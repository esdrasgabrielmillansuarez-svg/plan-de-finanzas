$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $excel.Workbooks.Add()

while ($workbook.Worksheets.Count -gt 1) {
    $workbook.Worksheets.Item(2).Delete()
}

# Config (oculta)
$sheetConfig = $workbook.Worksheets.Item(1)
$sheetConfig.Name = "Config"
$sheetConfig.Range("A1").Value = "LISTA DE CATEGORIAS DE INGRESOS"
$sheetConfig.Range("A1").Font.Bold = $true
$sheetConfig.Range("A2").Value = "Salario"
$sheetConfig.Range("A3").Value = "Freelance"
$sheetConfig.Range("A4").Value = "Inversiones"
$sheetConfig.Range("A5").Value = "Ventas"
$sheetConfig.Range("A6").Value = "Otros"

$sheetConfig.Range("C1").Value = "CATEGORIAS DE GASTOS"
$sheetConfig.Range("C1").Font.Bold = $true
$sheetConfig.Range("C2").Value = "Vivienda"
$sheetConfig.Range("C3").Value = "Alimentación"
$sheetConfig.Range("C4").Value = "Transporte"
$sheetConfig.Range("C5").Value = "Servicios"
$sheetConfig.Range("C6").Value = "Entretenimiento"
$sheetConfig.Range("C7").Value = "Deudas"
$sheetConfig.Range("C8").Value = "Ahorro"
$sheetConfig.Range("C9").Value = "Salud"
$sheetConfig.Range("C10").Value = "Educación"
$sheetConfig.Range("C11").Value = "Otros"

$sheetConfig.Range("E1").Value = "METODOS PAGO/CUENTAS"
$sheetConfig.Range("E1").Font.Bold = $true
$sheetConfig.Range("E2").Value = "Efectivo"
$sheetConfig.Range("E3").Value = "Tarjeta de Débito"
$sheetConfig.Range("E4").Value = "Tarjeta de Crédito"
$sheetConfig.Range("E5").Value = "Cuenta Bancaria"
$sheetConfig.Range("E6").Value = "Otros"
$sheetConfig.Columns("A:H").EntireColumn.AutoFit()

# Diario
$sheetDiario = $workbook.Worksheets.Add()
$sheetDiario.Name = "Control Diario"
$sheetDiario.Move($workbook.Worksheets.Item("Config"))

# AREA ESPECIFICA PARA AÑADIR GASTO/INGRESO
$sheetDiario.Range("A1").Value = "ÁREA RÁPIDA DE REGISTRO"
$sheetDiario.Range("A1").Font.Bold = $true
$sheetDiario.Range("A1").Font.Size = 14
$sheetDiario.Range("A1").Font.Color = 0xFFFFFF
$sheetDiario.Range("A1:J1").Merge()
$sheetDiario.Range("A1").Interior.Color = 0x0F3D56
$sheetDiario.Range("A1").HorizontalAlignment = -4108
$sheetDiario.Rows.Item(1).RowHeight = 32

$sheetDiario.Range("A3").Value = "Fecha"
$sheetDiario.Range("A4").NumberFormat = "dd/mm/yyyy"
$sheetDiario.Range("B3").Value = "Tipo"
$sheetDiario.Range("C3").Value = "Categoría"
$sheetDiario.Range("D3").Value = "Subcategoría"
$sheetDiario.Range("E3").Value = "Descripción / Concepto"
$sheetDiario.Range("F3").Value = "Monto"
$sheetDiario.Range("G3").Value = "Cuenta / Método"
$sheetDiario.Range("H3").Value = "Estado"
$sheetDiario.Range("I3").Value = "AÑADIR"
$sheetDiario.Range("A3:I3").Interior.Color = 0x1F4E79
$sheetDiario.Range("A3:I3").Font.Bold = $true
$sheetDiario.Range("A3:I3").Font.Color = 0xFFFFFF
$sheetDiario.Range("A3:I3").HorizontalAlignment = -4108
$sheetDiario.Range("A4").Formula = "=TODAY()"
$sheetDiario.Range("F4").NumberFormat = "`$#,##0.00"

# Validaciones área rápida
$validationTipoRap = $sheetDiario.Range("B4").Validation
$validationTipoRap.Delete()
$validationTipoRap.Add(3, 1, 1, "Ingreso,Gasto,Transferencia")
$validationCatRap = $sheetDiario.Range("C4").Validation
$validationCatRap.Delete()
$validationCatRap.Add(3, 1, 1, "=Config!`$A`$2:`$A`$6,Config!`$C`$2:`$C`$11")
$validationCuentaRap = $sheetDiario.Range("G4").Validation
$validationCuentaRap.Delete()
$validationCuentaRap.Add(3, 1, 1, "=Config!`$E`$2:`$E`$6")
$validationEstadoRap = $sheetDiario.Range("H4").Validation
$validationEstadoRap.Delete()
$validationEstadoRap.Add(3, 1, 1, "Pagado,Pendiente")

$sheetDiario.Range("I4").Value = "➕ AGREGAR"
$sheetDiario.Range("I4").Font.Bold = $true
$sheetDiario.Range("I4").Interior.Color = 0x70AD47
$sheetDiario.Range("I4").Font.Color = 0xFFFFFF
$sheetDiario.Range("I4").HorizontalAlignment = -4108
$sheetDiario.Range("A4:I4").Borders.LineStyle = 1
$sheetDiario.Range("A4:I4").Borders.Color = 0xCCCCCC

# Historial
$sheetDiario.Range("A7").Value = "HISTORIAL DE MOVIMIENTOS"
$sheetDiario.Range("A7").Font.Bold = $true
$sheetDiario.Range("A7").Font.Size = 12
$sheetDiario.Range("A7:J7").Merge()
$sheetDiario.Range("A7").Interior.Color = 0xF2F2F2
$sheetDiario.Range("A7").HorizontalAlignment = -4108

$headersDiario = @("Fecha", "Mes", "Tipo", "Categoría", "Subcategoría", "Descripción / Concepto", "Monto", "Cuenta / Método", "Estado")
for ($i=0; $i -lt $headersDiario.Count; $i++) {
    $cell = $sheetDiario.Cells.Item(8, $i+1)
    $cell.Value = $headersDiario[$i]
    $cell.Font.Bold = $true
    $cell.Font.Color = 0xFFFFFF
    $cell.Interior.Color = 0x0F3D56
}
$rangeDiario = $sheetDiario.Range("A8").CurrentRegion
$listObj = $sheetDiario.ListObjects.Add(1, $rangeDiario, $null, 1)
$listObj.Name = "TablaDiario"
$listObj.TableStyle = "TableStyleMedium9"
$sheetDiario.Range("B9").Formula = "=SIERROR(TEXTO([@Fecha],""mmmm""),"""")"
$sheetDiario.Range("B9:B2000").FillDown()
$sheetDiario.Columns("A:J").EntireColumn.AutoFit()

# Presupuesto_vs_Real oculta
$sheetPres = $workbook.Worksheets.Add()
$sheetPres.Name = "Presupuesto_vs_Real"
$sheetPres.Move($workbook.Worksheets.Item("Control Diario"))

# Control_Deudas_Ahorro oculta
$sheetDeuda = $workbook.Worksheets.Add()
$sheetDeuda.Name = "Control_Deudas_Ahorro"
$sheetDeuda.Move($workbook.Worksheets.Item("Presupuesto_vs_Real"))

# Dashboard
$sheetDash = $workbook.Worksheets.Add()
$sheetDash.Name = "Reporte Mensual"
$sheetDash.Move($workbook.Worksheets.Item("Control_Deudas_Ahorro"))

$sheetDash.Range("A1").Value = "REPORTE MENSUAL - DASHBOARD FINANCIERO"
$sheetDash.Range("A1").Font.Bold = $true
$sheetDash.Range("A1").Font.Size = 22
$sheetDash.Range("A1").Font.Color = 0xFFFFFF
$sheetDash.Range("A1:M1").Merge()
$sheetDash.Range("A1").Interior.Color = 0x0F3D56
$sheetDash.Range("A1").HorizontalAlignment = -4108
$sheetDash.Rows.Item(1).RowHeight = 40

$kpiLabels = @("TOTAL INGRESOS", "TOTAL GASTOS", "BALANCE NETO", "TASA DE AHORRO", "RATIO ENDEUDAMIENTO")
$kpiFormulas = @("=SUMAR.SI('Control Diario'[Tipo],""Ingreso"",'Control Diario'[Monto])","=SUMAR.SI('Control Diario'[Tipo],""Gasto"",'Control Diario'[Monto])","=B4-B5","=SIERROR((B4-B5)/B4,0)","=SIERROR(SUM('Control_Deudas_Ahorro'!D5:D12)/SUMAR.SI('Control Diario'[Tipo],""Ingreso"",'Control Diario'[Monto]),0)")
$colorsKPI = @(0x1F77B4, 0xD62728, 0x2CA02C, 0x9467BD, 0xFF7F0E)
for ($i=0; $i -lt $kpiLabels.Count; $i++) {
    $colK = 1 + ($i * 3)
    $c1 = [char](64 + $colK)
    $c2 = [char](64 + $colK + 1)
    $c3 = [char](64 + $colK + 2)
    $sheetDash.Range($c1 + "3:" + $c3 + "3").Merge()
    $sheetDash.Range($c1 + "3").Value = $kpiLabels[$i]
    $sheetDash.Range($c1 + "3").Font.Bold = $true
    $sheetDash.Range($c1 + "3").Font.Color = 0xFFFFFF
    $sheetDash.Range($c1 + "3").Interior.Color = $colorsKPI[$i]
    $sheetDash.Range($c1 + "3").HorizontalAlignment = -4108
    $sheetDash.Range($c1 + "4:" + $c3 + "4").Merge()
    $sheetDash.Range($c1 + "4").Formula = $kpiFormulas[$i]
    $sheetDash.Range($c1 + "4").Font.Bold = $true
    $sheetDash.Range($c1 + "4").Font.Size = 16
    $sheetDash.Range($c1 + "4").HorizontalAlignment = -4108
    $sheetDash.Range($c1 + "4").Interior.Color = 0xF8F9FA
    if ($i -ge 3) {
        $sheetDash.Range($c1 + "4").NumberFormat = "0.00%"
    } else {
        $sheetDash.Range($c1 + "4").NumberFormat = "`$#,##0.00"
    }
    $sheetDash.Range($c1 + "3:" + $c3 + "4").Borders.LineStyle = 1
    $sheetDash.Range($c1 + "3:" + $c3 + "4").Borders.Color = 0xCCCCCC
}
$sheetDash.Rows.Item(3).RowHeight = 24
$sheetDash.Rows.Item(4).RowHeight = 40

$meses = @("Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre")
$sheetDash.Range("A7").Value = "INGRESOS VS GASTOS MES A MES"
$sheetDash.Range("A7").Font.Bold = $true
$sheetDash.Range("A7").Font.Size = 12
$sheetDash.Range("A7:M7").Merge()
$sheetDash.Range("A8").Value = "Mes"
$sheetDash.Range("B8").Value = "Ingresos"
$sheetDash.Range("C8").Value = "Gastos"
$sheetDash.Range("D8").Value = "Saldo Neto"
$sheetDash.Range("A8:D8").Interior.Color = 0x0F3D56
$sheetDash.Range("A8:D8").Font.Bold = $true
$sheetDash.Range("A8:D8").Font.Color = 0xFFFFFF
$sheetDash.Range("A8:D8").HorizontalAlignment = -4108
for ($i=0; $i -lt $meses.Count; $i++) {
    $row = 9 + $i
    $sheetDash.Cells.Item($row, 1).Value = $meses[$i]
    $sheetDash.Cells.Item($row, 2).Formula = "=SUMAR.SI.CONJUNTO('Control Diario'[Monto],'Control Diario'[Tipo],""Ingreso"",'Control Diario'[Mes],A$row)"
    $sheetDash.Cells.Item($row, 3).Formula = "=SUMAR.SI.CONJUNTO('Control Diario'[Monto],'Control Diario'[Tipo],""Gasto"",'Control Diario'[Mes],A$row)"
    $sheetDash.Cells.Item($row, 4).Formula = "=B$row-C$row"
    $sheetDash.Range("B$row:D$row").NumberFormat = "`$#,##0.00"
    $sheetDash.Range("A$row:D$row").Borders.LineStyle = 1
    $sheetDash.Range("A$row:D$row").Borders.Color = 0xE0E0E0
    if ($i % 2 -eq 1) {
        $sheetDash.Range("A$row:D$row").Interior.Color = 0xF8F9FA
    }
}
$sheetDash.Range("A9:D20").Borders.LineStyle = 1
$sheetDash.Range("A9:D20").Borders.Color = 0xCCCCCC

$sheetDash.Range("F7").Value = "DISTRIBUCIÓN GASTOS POR CATEGORÍA"
$sheetDash.Range("F7").Font.Bold = $true
$sheetDash.Range("F7").Font.Size = 12
$sheetDash.Range("F7:M7").Merge()
$sheetDash.Range("F8").Value = "Categoría"
$sheetDash.Range("G8").Value = "Total"
$sheetDash.Range("H8").Value = "%"
$sheetDash.Range("F8:H8").Interior.Color = 0x0F3D56
$sheetDash.Range("F8:H8").Font.Bold = $true
$sheetDash.Range("F8:H8").Font.Color = 0xFFFFFF
$sheetDash.Range("F8:H8").HorizontalAlignment = -4108
$catsConfigList = @("Vivienda","Alimentación","Transporte","Servicios","Entretenimiento","Deudas","Ahorro","Salud","Educación","Otros")
for ($i=0; $i -lt $catsConfigList.Count; $i++) {
    $row = 9 + $i
    $sheetDash.Cells.Item($row, 6).Value = $catsConfigList[$i]
    $sheetDash.Cells.Item($row, 7).Formula = "=SUMAR.SI.CONJUNTO('Control Diario'[Monto],'Control Diario'[Tipo],""Gasto"",'Control Diario'[Categoría],F$row)"
    $sheetDash.Cells.Item($row, 8).Formula = "=SIERROR(G$row/$G$21,0)"
    $sheetDash.Range("G$row").NumberFormat = "`$#,##0.00"
    $sheetDash.Range("H$row").NumberFormat = "0.00%"
    $sheetDash.Range("F$row:H$row").Borders.LineStyle = 1
    $sheetDash.Range("F$row:H$row").Borders.Color = 0xE0E0E0
    if ($i % 2 -eq 1) {
        $sheetDash.Range("F$row:H$row").Interior.Color = 0xF8F9FA
    }
}
$sheetDash.Cells.Item(21, 6).Value = "TOTAL"
$sheetDash.Cells.Item(21, 6).Font.Bold = $true
$sheetDash.Cells.Item(21, 7).Formula = "=SUM(G9:G20)"
$sheetDash.Cells.Item(21, 7).NumberFormat = "`$#,##0.00"
$sheetDash.Cells.Item(21, 7).Font.Bold = $true
$sheetDash.Cells.Item(21, 8).Formula = "=SUM(H9:H20)"
$sheetDash.Cells.Item(21, 8).NumberFormat = "0.00%"
$sheetDash.Cells.Item(21, 8).Font.Bold = $true
$sheetDash.Range("F21:H21").Interior.Color = 0xD9E1F2
$sheetDash.Range("F21:H21").Font.Bold = $true
$sheetDash.Range("F8:H21").Borders.LineStyle = 1
$sheetDash.Range("F8:H21").Borders.Color = 0xCCCCCC

$chart1 = $sheetDash.Shapes.AddChart2(227, 51)
$chart1.Chart.SetSourceData($sheetDash.Range("A8:D20"))
$chart1.Left = 550
$chart1.Top = 120
$chart1.Width = 380
$chart1.Height = 240
$chart1.Chart.ChartTitle.Text = "Ingresos vs Gastos"

$chart2 = $sheetDash.Shapes.AddChart2(251, 5)
$chart2.Chart.SetSourceData($sheetDash.Range("F8:H21"), 2)
$chart2.Left = 550
$chart2.Top = 380
$chart2.Width = 380
$chart2.Height = 240
$chart2.Chart.ChartTitle.Text = "Distribución por Categoría"

$sheetDash.Columns("A:M").EntireColumn.AutoFit()

# Ocultar hojas
$sheetConfig.Visible = 2
$sheetPres.Visible = 2
$sheetDeuda.Visible = 2

$sheetDiario.Activate()

$filePath = "C:\Users\Suarezmillan\Documents\proyectos\juego\plan de finanzas.xlsx"
$workbook.SaveAs($filePath)
$workbook.Close($false)
$excel.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

Write-Host "OK"
