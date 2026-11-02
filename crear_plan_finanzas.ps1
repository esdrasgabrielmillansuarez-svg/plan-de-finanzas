$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $excel.Workbooks.Add()

while ($workbook.Worksheets.Count -gt 1) {
    $workbook.Worksheets.Item(2).Delete()
}

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

$sheetConfig.Range("E1").Value = "METODOS DE PAGO/CUENTAS"
$sheetConfig.Range("E1").Font.Bold = $true
$sheetConfig.Range("E2").Value = "Efectivo"
$sheetConfig.Range("E3").Value = "Tarjeta de Débito"
$sheetConfig.Range("E4").Value = "Tarjeta de Crédito"
$sheetConfig.Range("E5").Value = "Cuenta Bancaria"
$sheetConfig.Range("E6").Value = "Otros"

$sheetConfig.Range("G1").Value = "PRESUPUESTO MENSUAL ASIGNADO"
$sheetConfig.Range("G1").Font.Bold = $true
$sheetConfig.Range("G2").Value = "Categoría"
$sheetConfig.Range("H2").Value = "Presupuesto"
$sheetConfig.Range("G3").Value = "Vivienda"
$sheetConfig.Range("H3").Value = 0
$sheetConfig.Range("G4").Value = "Alimentación"
$sheetConfig.Range("H4").Value = 0
$sheetConfig.Range("G5").Value = "Transporte"
$sheetConfig.Range("H5").Value = 0
$sheetConfig.Range("G6").Value = "Servicios"
$sheetConfig.Range("H6").Value = 0
$sheetConfig.Range("G7").Value = "Entretenimiento"
$sheetConfig.Range("H7").Value = 0
$sheetConfig.Range("G8").Value = "Deudas"
$sheetConfig.Range("H8").Value = 0
$sheetConfig.Range("G9").Value = "Ahorro"
$sheetConfig.Range("H9").Value = 0
$sheetConfig.Range("G10").Value = "Salud"
$sheetConfig.Range("H10").Value = 0
$sheetConfig.Range("G11").Value = "Educación"
$sheetConfig.Range("H11").Value = 0
$sheetConfig.Range("G12").Value = "Otros"
$sheetConfig.Range("H12").Value = 0
$sheetConfig.Columns("A:H").EntireColumn.AutoFit()

$sheetDiario = $workbook.Worksheets.Add()
$sheetDiario.Name = "Diario"
$sheetDiario.Move($workbook.Worksheets.Item("Config"))

$headersDiario = @("Fecha", "Mes", "Tipo", "Categoría", "Subcategoría", "Descripción / Concepto", "Monto", "Cuenta / Método", "Estado")
for ($i=0; $i -lt $headersDiario.Count; $i++) {
    $cell = $sheetDiario.Cells.Item(1, $i+1)
    $cell.Value = $headersDiario[$i]
    $cell.Font.Bold = $true
    $cell.Interior.Color = 0xD9D9D9
}

$rangeDiario = $sheetDiario.Range("A1").CurrentRegion
$listObj = $sheetDiario.ListObjects.Add(1, $rangeDiario, $null, 1)
$listObj.Name = "TablaDiario"
$listObj.TableStyle = "TableStyleMedium2"

$sheetDiario.Range("B2").Formula = "=SIERROR(TEXTO([@Fecha],""mmmm""),"""")"
$sheetDiario.Range("B2:B1000").FillDown()

$validationTipo = $sheetDiario.Range("C2:C1000").Validation
$validationTipo.Delete()
$validationTipo.Add(3, 1, 1, "Ingreso,Gasto,Transferencia")

$validationCat = $sheetDiario.Range("D2:D1000").Validation
$validationCat.Delete()
$validationCat.Add(3, 1, 1, "=Config!`$A`$2:`$A`$6,Config!`$C`$2:`$C`$11")

$validationCuenta = $sheetDiario.Range("H2:H1000").Validation
$validationCuenta.Delete()
$validationCuenta.Add(3, 1, 1, "=Config!`$E`$2:`$E`$6")

$validationEstado = $sheetDiario.Range("I2:I1000").Validation
$validationEstado.Delete()
$validationEstado.Add(3, 1, 1, "Pagado,Pendiente")

$sheetDiario.Columns("A:I").EntireColumn.AutoFit()

$sheetPres = $workbook.Worksheets.Add()
$sheetPres.Name = "Presupuesto_vs_Real"
$sheetPres.Move($workbook.Worksheets.Item("Diario"))

$sheetPres.Range("A1").Value = "PRESUPUESTO VS REAL - COMPARATIVO MENSUAL"
$sheetPres.Range("A1").Font.Bold = $true
$sheetPres.Range("A1").Font.Size = 14
$sheetPres.Range("A2").Value = "Categoría"
$meses = @("Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre")
$col = 2
foreach ($mes in $meses) {
    $sheetPres.Cells.Item(2, $col).Value = "$mes - Presupuesto"
    $sheetPres.Cells.Item(2, $col+1).Value = "$mes - Real"
    $sheetPres.Cells.Item(2, $col+2).Value = "$mes - Desviación"
    $sheetPres.Cells.Item(2, $col+3).Value = "$mes - % Cumplimiento"
    $col += 4
}

$cats = @("Vivienda","Alimentación","Transporte","Servicios","Entretenimiento","Deudas","Ahorro","Salud","Educación","Otros")
for ($i=0; $i -lt $cats.Count; $i++) {
    $rowIdx = 3 + $i
    $sheetPres.Cells.Item($rowIdx, 1).Value = $cats[$i]
    $sheetPres.Cells.Item($rowIdx, 2).Formula = "=SIERROR(BUSCARV(A$rowIdx,Config!`$G`$3:`$H`$12,2,FALSE),0)"
}

$sheetPres.Cells.Item(3, 3).Formula = "=SIERROR(SUMAR.SI.CONJUNTO(Diario[Monto],Diario[Tipo],""Gasto"",Diario[Mes],""Enero"",Diario[Categoría],A3),0)"
$sheetPres.Cells.Item(3, 4).Formula = "=B3-C3"
$sheetPres.Cells.Item(3, 5).Formula = "=SIERROR(C3/B3,0)"
$sheetPres.Range("E3").NumberFormat = "0.00%"

$sheetPres.Range("C3:E12").FillDown()
$sheetPres.Range("B3:B12").FillDown()

$sheetPres.Columns.EntireColumn.AutoFit()

$sheetDeuda = $workbook.Worksheets.Add()
$sheetDeuda.Name = "Control_Deudas_Ahorro"
$sheetDeuda.Move($workbook.Worksheets.Item("Presupuesto_vs_Real"))

$sheetDeuda.Range("A1").Value = "CONTROL DE DEUDAS, ACTIVOS Y AHORRO"
$sheetDeuda.Range("A1").Font.Bold = $true
$sheetDeuda.Range("A1").Font.Size = 14
$sheetDeuda.Range("A3").Value = "SECCIÓN A: DEUDAS / TARJETAS DE CRÉDITO"
$sheetDeuda.Range("A3").Font.Bold = $true

$deudaHeaders = @("Concepto/Entidad", "Saldo Total", "Tasa de Interés (%)", "Pago Mínimo", "Fecha de Corte", "Fecha de Pago", "Saldo Pendiente", "Estado")
for ($i=0; $i -lt $deudaHeaders.Count; $i++) {
    $cell = $sheetDeuda.Cells.Item(4, $i+1)
    $cell.Value = $deudaHeaders[$i]
    $cell.Font.Bold = $true
    $cell.Interior.Color = 0xD9D9D9
}
$sheetDeuda.Range("G5").Formula = "=B5"

$sheetDeuda.Range("A12").Value = "SECCIÓN B: FONDOS Y AHORRO"
$sheetDeuda.Range("A12").Font.Bold = $true
$ahorroHeaders = @("Meta/Fondo", "Objetivo", "Saldo Acumulado", "Fecha Objetivo", "Falta por Alcanzar", "% Alcanzado")
for ($i=0; $i -lt $ahorroHeaders.Count; $i++) {
    $cell = $sheetDeuda.Cells.Item(13, $i+1)
    $cell.Value = $ahorroHeaders[$i]
    $cell.Font.Bold = $true
    $cell.Interior.Color = 0xD9D9D9
}
$sheetDeuda.Range("E14").Formula = "=B14-C14"
$sheetDeuda.Range("F14").Formula = "=SIERROR(C14/B14,0)"
$sheetDeuda.Range("F14").NumberFormat = "0.00%"

$sheetDeuda.Columns.EntireColumn.AutoFit()

$sheetDash = $workbook.Worksheets.Add()
$sheetDash.Name = "Dashboard"
$sheetDash.Move($workbook.Worksheets.Item("Control_Deudas_Ahorro"))

$sheetDash.Range("A1").Value = "DASHBOARD - PANEL DE CONTROL FINANCIERO"
$sheetDash.Range("A1").Font.Bold = $true
$sheetDash.Range("A1").Font.Size = 16
$sheetDash.Range("A3").Value = "INDICADORES CLAVE (KPI's)"
$sheetDash.Range("A3").Font.Bold = $true

$kpis = @("TOTAL INGRESOS ANUALES", "TOTAL GASTOS ANUALES", "BALANCE NETO / AHORRO TOTAL", "TASA DE AHORRO PROMEDIO", "RATIO DE ENDEUDAMIENTO")
$kpiFormulas = @("=SUMAR.SI(Diario[Tipo],""Ingreso"",Diario[Monto])","=SUMAR.SI(Diario[Tipo],""Gasto"",Diario[Monto])","=B4-B5","=SIERROR((B4-B5)/B4,0)","=SIERROR(SUM(Control_Deudas_Ahorro!D5:D10)/SUMAR.SI(Diario[Tipo],""Ingreso"",Diario[Monto]),0)")
for ($i=0; $i -lt $kpis.Count; $i++) {
    $row = 4 + $i
    $sheetDash.Cells.Item($row, 1).Value = $kpis[$i]
    $sheetDash.Cells.Item($row, 1).Interior.Color = 0xF2F2F2
    $sheetDash.Cells.Item($row, 1).Font.Bold = $true
    $sheetDash.Cells.Item($row, 2).Formula = $kpiFormulas[$i]
    if ($i -eq 3 -or $i -eq 4) {
        $sheetDash.Cells.Item($row, 2).NumberFormat = "0.00%"
    } else {
        $sheetDash.Cells.Item($row, 2).NumberFormat = "`$#,##0.00"
    }
}

$sheetDash.Range("A15").Value = "RESUMEN MENSUAL INGRESOS VS GASTOS"
$sheetDash.Range("A15").Font.Bold = $true
$sheetDash.Range("A16").Value = "Mes"
$sheetDash.Range("B16").Value = "Ingresos"
$sheetDash.Range("C16").Value = "Gastos"
$sheetDash.Range("D16").Value = "Saldo Neto"
for ($i=0; $i -lt $meses.Count; $i++) {
    $row = 17 + $i
    $sheetDash.Cells.Item($row, 1).Value = $meses[$i]
    $sheetDash.Cells.Item($row, 2).Formula = "=SUMAR.SI.CONJUNTO(Diario[Monto],Diario[Tipo],""Ingreso"",Diario[Mes],A$row)"
    $sheetDash.Cells.Item($row, 3).Formula = "=SUMAR.SI.CONJUNTO(Diario[Monto],Diario[Tipo],""Gasto"",Diario[Mes],A$row)"
    $sheetDash.Cells.Item($row, 4).Formula = "=B$row-C$row"
    $sheetDash.Range("B$row:D$row").NumberFormat = "`$#,##0.00"
}

$sheetDash.Range("A4:B8").Borders.LineStyle = 1
$sheetDash.Range("A16:D28").Borders.LineStyle = 1
$sheetDash.Columns.EntireColumn.AutoFit()

$filePath = "C:\Users\Suarezmillan\Documents\proyectos\juego\plan de finanzas.xlsx"
$workbook.SaveAs($filePath)
$workbook.Close($false)
$excel.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

Write-Host "OK"
