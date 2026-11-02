$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $excel.Workbooks.Add()

while ($workbook.Worksheets.Count -gt 1) {
    $workbook.Worksheets.Item(2).Delete()
}

# Config
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
$catsConfig = @("Vivienda","Alimentación","Transporte","Servicios","Entretenimiento","Deudas","Ahorro","Salud","Educación","Otros")
for ($i=0; $i -lt $catsConfig.Count; $i++) {
    $sheetConfig.Cells.Item(3+$i, 7).Value = $catsConfig[$i]
    $sheetConfig.Cells.Item(3+$i, 8).Value = 0
}
$sheetConfig.Columns("A:H").EntireColumn.AutoFit()

# Diario
$sheetDiario = $workbook.Worksheets.Add()
$sheetDiario.Name = "Diario"
$sheetDiario.Move($workbook.Worksheets.Item("Config"))
$headersDiario = @("Fecha", "Mes", "Tipo", "Categoría", "Subcategoría", "Descripción / Concepto", "Monto", "Cuenta / Método", "Estado")
for ($i=0; $i -lt $headersDiario.Count; $i++) {
    $cell = $sheetDiario.Cells.Item(1, $i+1)
    $cell.Value = $headersDiario[$i]
    $cell.Font.Bold = $true
    $cell.Font.Color = 0xFFFFFF
    $cell.Interior.Color = 0x1F4E79
}
$rangeDiario = $sheetDiario.Range("A1").CurrentRegion
$listObj = $sheetDiario.ListObjects.Add(1, $rangeDiario, $null, 1)
$listObj.Name = "TablaDiario"
$listObj.TableStyle = "TableStyleDark2"
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

# Presupuesto_vs_Real
$sheetPres = $workbook.Worksheets.Add()
$sheetPres.Name = "Presupuesto_vs_Real"
$sheetPres.Move($workbook.Worksheets.Item("Diario"))
$sheetPres.Range("A1").Value = "PRESUPUESTO VS REAL - COMPARATIVO MENSUAL"
$sheetPres.Range("A1").Font.Bold = $true
$sheetPres.Range("A1").Font.Size = 14
$sheetPres.Range("A1").Font.Color = 0xFFFFFF
$sheetPres.Range("A1:AZ1").Merge()
$sheetPres.Range("A1").Interior.Color = 0x1F4E79
$sheetPres.Range("A1").HorizontalAlignment = -4108

$meses = @("Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre")
$sheetPres.Range("A3").Value = "Categoría"
$col = 2
foreach ($mes in $meses) {
    $sheetPres.Cells.Item(3, $col).Value = "$mes - Presupuesto"
    $sheetPres.Cells.Item(3, $col+1).Value = "$mes - Real"
    $sheetPres.Cells.Item(3, $col+2).Value = "$mes - Desviación"
    $sheetPres.Cells.Item(3, $col+3).Value = "% Cumplimiento"
    $col += 4
}
$sheetPres.Range("A3:AZ3").Interior.Color = 0xD9E1F2
$sheetPres.Range("A3:AZ3").Font.Bold = $true
$sheetPres.Range("A3:AZ3").HorizontalAlignment = -4108

for ($i=0; $i -lt $catsConfig.Count; $i++) {
    $rowIdx = 4 + $i
    $sheetPres.Cells.Item($rowIdx, 1).Value = $catsConfig[$i]
    $sheetPres.Cells.Item($rowIdx, 1).Font.Bold = $true
    $sheetPres.Cells.Item($rowIdx, 2).Formula = "=SIERROR(BUSCARV(A$rowIdx,Config!`$G`$3:`$H`$12,2,FALSE),0)"
    $sheetPres.Cells.Item($rowIdx, 2).NumberFormat = "`$#,##0.00"
}
$sheetPres.Range("B4:B13").FillDown()
$sheetPres.Columns.EntireColumn.AutoFit()

# Control_Deudas_Ahorro
$sheetDeuda = $workbook.Worksheets.Add()
$sheetDeuda.Name = "Control_Deudas_Ahorro"
$sheetDeuda.Move($workbook.Worksheets.Item("Presupuesto_vs_Real"))
$sheetDeuda.Range("A1").Value = "CONTROL DE DEUDAS, ACTIVOS Y AHORRO"
$sheetDeuda.Range("A1").Font.Bold = $true
$sheetDeuda.Range("A1").Font.Size = 14
$sheetDeuda.Range("A1").Font.Color = 0xFFFFFF
$sheetDeuda.Range("A1:H1").Merge()
$sheetDeuda.Range("A1").Interior.Color = 0x1F4E79
$sheetDeuda.Range("A1").HorizontalAlignment = -4108

$sheetDeuda.Range("A3").Value = "SECCIÓN A: DEUDAS / TARJETAS DE CRÉDITO"
$sheetDeuda.Range("A3").Font.Bold = $true
$sheetDeuda.Range("A3").Font.Color = 0xFFFFFF
$sheetDeuda.Range("A3:H3").Merge()
$sheetDeuda.Range("A3").Interior.Color = 0xC00000
$deudaHeaders = @("Concepto/Entidad", "Saldo Total", "Tasa de Interés (%)", "Pago Mínimo", "Fecha de Corte", "Fecha de Pago", "Saldo Pendiente", "Estado")
for ($i=0; $i -lt $deudaHeaders.Count; $i++) {
    $cell = $sheetDeuda.Cells.Item(4, $i+1)
    $cell.Value = $deudaHeaders[$i]
    $cell.Font.Bold = $true
    $cell.Font.Color = 0xFFFFFF
    $cell.Interior.Color = 0x404040
}
$sheetDeuda.Range("G5").Formula = "=B5"
$sheetDeuda.Range("G5:G12").FillDown()

$sheetDeuda.Range("A14").Value = "SECCIÓN B: FONDOS Y AHORRO"
$sheetDeuda.Range("A14").Font.Bold = $true
$sheetDeuda.Range("A14").Font.Color = 0xFFFFFF
$sheetDeuda.Range("A14:F14").Merge()
$sheetDeuda.Range("A14").Interior.Color = 0x548235
$ahorroHeaders = @("Meta/Fondo", "Objetivo", "Saldo Acumulado", "Fecha Objetivo", "Falta por Alcanzar", "% Alcanzado")
for ($i=0; $i -lt $ahorroHeaders.Count; $i++) {
    $cell = $sheetDeuda.Cells.Item(15, $i+1)
    $cell.Value = $ahorroHeaders[$i]
    $cell.Font.Bold = $true
    $cell.Font.Color = 0xFFFFFF
    $cell.Interior.Color = 0x404040
}
$sheetDeuda.Range("E16").Formula = "=B16-C16"
$sheetDeuda.Range("F16").Formula = "=SIERROR(C16/B16,0)"
$sheetDeuda.Range("F16").NumberFormat = "0.00%"
$sheetDeuda.Range("E16:F22").FillDown()
$sheetDeuda.Columns.EntireColumn.AutoFit()

# Dashboard
$sheetDash = $workbook.Worksheets.Add()
$sheetDash.Name = "Dashboard"
$sheetDash.Move($workbook.Worksheets.Item("Control_Deudas_Ahorro"))

$sheetDash.Range("A1").Value = "FINANZAS | DASHBOARD INTERACTIVO"
$sheetDash.Range("A1").Font.Bold = $true
$sheetDash.Range("A1").Font.Size = 20
$sheetDash.Range("A1").Font.Color = 0xFFFFFF
$sheetDash.Range("A1:N1").Merge()
$sheetDash.Range("A1").Interior.Color = 0x0F3D56
$sheetDash.Range("A1").HorizontalAlignment = -4108
$sheetDash.Rows.Item(1).RowHeight = 36

$sheetDash.Range("A3").Value = "INDICADORES CLAVE"
$sheetDash.Range("A3").Font.Bold = $true
$sheetDash.Range("A3").Font.Size = 12
$sheetDash.Range("A3:N3").Merge()

$kpiLabels = @("TOTAL INGRESOS", "TOTAL GASTOS", "BALANCE NETO", "TASA DE AHORRO", "RATIO ENDEUDAMIENTO")
$kpiFormulas = @("=SUMAR.SI(Diario[Tipo],""Ingreso"",Diario[Monto])","=SUMAR.SI(Diario[Tipo],""Gasto"",Diario[Monto])","=B5-B6","=SIERROR((B5-B6)/B5,0)","=SIERROR(SUM(Control_Deudas_Ahorro!D5:D12)/SUMAR.SI(Diario[Tipo],""Ingreso"",Diario[Monto]),0)")
$colorsKPI = @(0x1F4E79, 0xC00000, 0x548235, 0x7030A0, 0xED7D31)
for ($i=0; $i -lt $kpiLabels.Count; $i++) {
    $colK = 1 + ($i * 3)
    $c1 = [char](64 + $colK)
    $c2 = [char](64 + $colK + 1)
    $c3 = [char](64 + $colK + 2)
    $sheetDash.Range($c1 + "5:" + $c3 + "5").Merge()
    $sheetDash.Range($c1 + "5").Value = $kpiLabels[$i]
    $sheetDash.Range($c1 + "5").Font.Bold = $true
    $sheetDash.Range($c1 + "5").Font.Color = 0xFFFFFF
    $sheetDash.Range($c1 + "5").Interior.Color = $colorsKPI[$i]
    $sheetDash.Range($c1 + "5").HorizontalAlignment = -4108
    $sheetDash.Range($c1 + "6:" + $c3 + "6").Merge()
    $sheetDash.Range($c1 + "6").Formula = $kpiFormulas[$i]
    $sheetDash.Range($c1 + "6").Font.Bold = $true
    $sheetDash.Range($c1 + "6").Font.Size = 14
    $sheetDash.Range($c1 + "6").HorizontalAlignment = -4108
    $sheetDash.Range($c1 + "6").Interior.Color = 0xF8F9FA
    if ($i -ge 3) {
        $sheetDash.Range($c1 + "6").NumberFormat = "0.00%"
    } else {
        $sheetDash.Range($c1 + "6").NumberFormat = "`$#,##0.00"
    }
    $sheetDash.Range($c1 + "5:" + $c3 + "6").Borders.LineStyle = 1
    $sheetDash.Range($c1 + "5:" + $c3 + "6").Borders.Color = 0xCCCCCC
}
$sheetDash.Rows.Item(5).RowHeight = 22
$sheetDash.Rows.Item(6).RowHeight = 36

$sheetDash.Range("A10").Value = "RESUMEN MENSUAL INGRESOS VS GASTOS"
$sheetDash.Range("A10").Font.Bold = $true
$sheetDash.Range("A10").Font.Size = 12
$sheetDash.Range("A10:N10").Merge()
$sheetDash.Range("A11").Value = "Mes"
$sheetDash.Range("B11").Value = "Ingresos"
$sheetDash.Range("C11").Value = "Gastos"
$sheetDash.Range("D11").Value = "Saldo Neto"
$sheetDash.Range("A11:D11").Interior.Color = 0x1F4E79
$sheetDash.Range("A11:D11").Font.Bold = $true
$sheetDash.Range("A11:D11").Font.Color = 0xFFFFFF
$sheetDash.Range("A11:D11").HorizontalAlignment = -4108
for ($i=0; $i -lt $meses.Count; $i++) {
    $row = 12 + $i
    $sheetDash.Cells.Item($row, 1).Value = $meses[$i]
    $sheetDash.Cells.Item($row, 2).Formula = "=SUMAR.SI.CONJUNTO(Diario[Monto],Diario[Tipo],""Ingreso"",Diario[Mes],A$row)"
    $sheetDash.Cells.Item($row, 3).Formula = "=SUMAR.SI.CONJUNTO(Diario[Monto],Diario[Tipo],""Gasto"",Diario[Mes],A$row)"
    $sheetDash.Cells.Item($row, 4).Formula = "=B$row-C$row"
    $sheetDash.Range("B$row:D$row").NumberFormat = "`$#,##0.00"
    $sheetDash.Range("A$row:D$row").Borders.LineStyle = 1
    $sheetDash.Range("A$row:D$row").Borders.Color = 0xE0E0E0
    if ($i % 2 -eq 1) {
        $sheetDash.Range("A$row:D$row").Interior.Color = 0xF8F9FA
    }
}
$sheetDash.Range("A12:D23").Borders.LineStyle = 1
$sheetDash.Range("A12:D23").Borders.Color = 0xCCCCCC

$sheetDash.Range("F10").Value = "DISTRIBUCIÓN POR CATEGORÍA (GASTOS)"
$sheetDash.Range("F10").Font.Bold = $true
$sheetDash.Range("F10").Font.Size = 12
$sheetDash.Range("F10:J10").Merge()
$sheetDash.Range("F11").Value = "Categoría"
$sheetDash.Range("G11").Value = "Total Gastado"
$sheetDash.Range("H11").Value = "% del Total"
$sheetDash.Range("F11:H11").Interior.Color = 0x1F4E79
$sheetDash.Range("F11:H11").Font.Bold = $true
$sheetDash.Range("F11:H11").Font.Color = 0xFFFFFF
$sheetDash.Range("F11:H11").HorizontalAlignment = -4108
for ($i=0; $i -lt $catsConfig.Count; $i++) {
    $row = 12 + $i
    $sheetDash.Cells.Item($row, 6).Value = $catsConfig[$i]
    $sheetDash.Cells.Item($row, 7).Formula = "=SUMAR.SI.CONJUNTO(Diario[Monto],Diario[Tipo],""Gasto"",Diario[Categoría],F$row)"
    $sheetDash.Cells.Item($row, 8).Formula = "=SIERROR(G$row/$G$24,0)"
    $sheetDash.Range("G$row").NumberFormat = "`$#,##0.00"
    $sheetDash.Range("H$row").NumberFormat = "0.00%"
    $sheetDash.Range("F$row:H$row").Borders.LineStyle = 1
    $sheetDash.Range("F$row:H$row").Borders.Color = 0xE0E0E0
    if ($i % 2 -eq 1) {
        $sheetDash.Range("F$row:H$row").Interior.Color = 0xF8F9FA
    }
}
$sheetDash.Cells.Item(24, 6).Value = "TOTAL"
$sheetDash.Cells.Item(24, 6).Font.Bold = $true
$sheetDash.Cells.Item(24, 7).Formula = "=SUM(G12:G23)"
$sheetDash.Cells.Item(24, 7).NumberFormat = "`$#,##0.00"
$sheetDash.Cells.Item(24, 7).Font.Bold = $true
$sheetDash.Cells.Item(24, 8).Formula = "=SUM(H12:H23)"
$sheetDash.Cells.Item(24, 8).NumberFormat = "0.00%"
$sheetDash.Cells.Item(24, 8).Font.Bold = $true
$sheetDash.Range("F24:H24").Interior.Color = 0xD9E1F2
$sheetDash.Range("F24:H24").Font.Bold = $true
$sheetDash.Range("F11:H24").Borders.LineStyle = 1
$sheetDash.Range("F11:H24").Borders.Color = 0xCCCCCC

$sheetDash.Range("J10").Value = "ACCESO RÁPIDO"
$sheetDash.Range("J10").Font.Bold = $true
$sheetDash.Range("J10").Font.Size = 12
$sheetDash.Range("J10:N10").Merge()
$sheetDash.Range("J12").Value = "📝 Ir a Diario"
$sheetDash.Range("J12").Font.Bold = $true
$sheetDash.Range("J12").Interior.Color = 0xE8F5E9
$sheetDash.Range("J12:N12").Borders.LineStyle = 1
$sheetDash.Range("J13").Value = "📊 Ir a Presupuesto vs Real"
$sheetDash.Range("J13").Font.Bold = $true
$sheetDash.Range("J13").Interior.Color = 0xE3F2FD
$sheetDash.Range("J13:N13").Borders.LineStyle = 1
$sheetDash.Range("J14").Value = "💳 Ir a Control Deudas/Ahorro"
$sheetDash.Range("J14").Font.Bold = $true
$sheetDash.Range("J14").Interior.Color = 0xFFF3E0
$sheetDash.Range("J14:N14").Borders.LineStyle = 1
$sheetDash.Range("J15").Value = "⚙️ Ir a Configuración"
$sheetDash.Range("J15").Font.Bold = $true
$sheetDash.Range("J15").Interior.Color = 0xF3E5F5
$sheetDash.Range("J15:N15").Borders.LineStyle = 1

$sheetDash.Columns("A:N").EntireColumn.AutoFit()
$sheetDash.Activate()

$filePath = "C:\Users\Suarezmillan\Documents\proyectos\juego\plan de finanzas.xlsx"
$workbook.SaveAs($filePath)
$workbook.Close($false)
$excel.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

Write-Host "OK"
