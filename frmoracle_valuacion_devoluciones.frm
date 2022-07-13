VERSION 5.00
Begin VB.Form frmoracle_valuacion_devoluciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de valuación de devoluciones"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Dep. Mex."
      Height          =   315
      Left            =   2175
      TabIndex        =   10
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmd_general 
      Caption         =   "Alm. Gral"
      Height          =   315
      Left            =   1200
      TabIndex        =   9
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmd_calidad 
      Caption         =   "Calidad"
      Height          =   315
      Left            =   345
      TabIndex        =   8
      Top             =   0
      Width           =   855
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   840
      Left            =   60
      TabIndex        =   2
      Top             =   435
      Width           =   4245
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   315
         Width           =   1140
      End
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   2865
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   300
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   375
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2325
         TabIndex        =   5
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3945
      Picture         =   "frmoracle_valuacion_devoluciones.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   0
      Picture         =   "frmoracle_valuacion_devoluciones.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   30
      TabIndex        =   7
      Top             =   270
      Width           =   4275
   End
End
Attribute VB_Name = "frmoracle_valuacion_devoluciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_calidad_Click()
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim año As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_consecutivo As Integer
   Dim var_afectacion_movimiento As String
   Dim var_vistas As String
   Dim iFila As Long, iCol As Integer, i As Integer
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            cnn.CommandTimeout = 720
            cnn.BeginTrans
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open "select max(inte_tem_consecutivo) as numero from tb_oracle_valuacion_devoluciones", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!numero), 0, rs!numero)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into tb_oracle_valuacion_devoluciones (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            var_dia = CStr(Day(CDate(txt_inicio)))
            var_mes = CStr(Month(CDate(txt_inicio)))
            var_año = CStr(Year(CDate(txt_inicio)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            
            var_fecha_inicio = var_dia + "-" + var_mes + "-" + var_año
            VAR_FECHA_INICIO_TABLA = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            var_fecha_fin_1 = CDate(txt_fin) + 1
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = var_dia + "-" + var_mes + "-" + var_año
            
            
            
            var_fecha_fin_1 = CDate(txt_fin)
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            VAR_FECHA_FIN_TABLA = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
             
            rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            If var_unidad_organizacional = 93 Then
               var_cadena = " select distinct OL.line_id, agente, nombre_agente, cliente, nombre_cliente, 'SIDDC_'||CAST(NUMERO AS VARCHAR(50)) AS FOLI0, FECHA_FIN, ORDER_NUMBER, ORDERED_ITEM, ORDER_QUANTITY_UOM, pricing_quantity, UNIT_SELLING_PRICE, ol.attribute11, pricing_quantity *  unit_selling_price as importe, xv.description, REFERENCIA, NUMERO_REFERENCIA  from xxvia_tb_devoluciones_clientes x, OE_ORDER_HEADERS_ALL  OE, OE_ORDER_LINES_ALL OL, xxvia_system_items_b xv where trunc(to_date(x.fecha_fin,'DD/MM/YYYY'))  >= to_date('" + Me.txt_inicio + "','DD-MM-YYYY')  AND trunc(to_date(x.fecha_fin,'DD/MM/YYYY')) < TO_DATE('" + Me.txt_fin + "','DD-MM-YYYY') + 1 AND ESTATUS = 'I' AND MOVIMIENTO = 'DC' AND OE.orig_sys_document_ref = 'SIDDC_'||CAST(NUMERO AS VARCHAR(50)) AND OE.HEADER_ID = OL.HEADER_ID and ol.inventory_item_id = xv.inventory_item_id and oe.ship_from_org_id = xv.organization_id AND xv.organization_id = '" + var_unidad_organizacional + "' and almacen = 'CDI_ALMCAL' UNION "
               var_cadena = var_cadena + " select distinct OL.line_id, agente, nombre_agente, cliente, nombre_cliente, 'SIDDCS_'||CAST(NUMERO AS VARCHAR(50)) AS FOLI0, FECHA_FIN, ORDER_NUMBER, ORDERED_ITEM, ORDER_QUANTITY_UOM, pricing_quantity, UNIT_SELLING_PRICE, ol.attribute11, pricing_quantity *  unit_selling_price as importe, xv.description, REFERENCIA, NUMERO_REFERENCIA "
               var_cadena = var_cadena + " from xxvia_tb_devoluciones_clientes x, OE_ORDER_HEADERS_ALL  OE, OE_ORDER_LINES_ALL OL, xxvia_system_items_b xv where trunc(to_date(x.fecha_fin,'DD/MM/YYYY'))  >= to_date('" + Me.txt_inicio + "','DD-MM-YYYY')  AND trunc(to_date(x.fecha_fin,'DD/MM/YYYY')) < TO_DATE('" + Me.txt_fin + "','DD-MM-YYYY') + 1 AND ESTATUS = 'I' AND MOVIMIENTO = 'DCS' AND OE.orig_sys_document_ref = 'SIDDCS_'||CAST(NUMERO AS VARCHAR(50)) AND OE.HEADER_ID = OL.HEADER_ID and ol.inventory_item_id = xv.inventory_item_id and oe.ship_from_org_id = xv.organization_id AND xv.organization_id = '" + var_unidad_organizacional + "' and almacen = 'CDI_ALMCAL' ORDER BY FECHA_FIN DESC"
            Else
               var_cadena = " select distinct OL.line_id, agente, nombre_agente, cliente, nombre_cliente, 'SIDDC_'||CAST(NUMERO AS VARCHAR(50)) AS FOLI0, FECHA_FIN, ORDER_NUMBER, ORDERED_ITEM, ORDER_QUANTITY_UOM, pricing_quantity, UNIT_SELLING_PRICE, ol.attribute11, pricing_quantity *  unit_selling_price as importe, xv.description, REFERENCIA, NUMERO_REFERENCIA "
               var_cadena = var_cadena + " from xxvia_tb_devoluciones_clientes x, OE_ORDER_HEADERS_ALL  OE, OE_ORDER_LINES_ALL OL, xxvia_system_items_b xv where trunc(to_date(x.fecha_fin,'DD/MM/YYYY'))  >= to_date('" + Me.txt_inicio + "','DD-MM-YYYY')  AND trunc(to_date(x.fecha_fin,'DD/MM/YYYY')) < TO_DATE('" + Me.txt_fin + "','DD-MM-YYYY') + 1 AND ESTATUS = 'I' AND MOVIMIENTO = 'DC' AND OE.orig_sys_document_ref = 'SIDDC_'||CAST(NUMERO AS VARCHAR(50)) AND OE.HEADER_ID = OL.HEADER_ID and ol.inventory_item_id = xv.inventory_item_id and oe.ship_from_org_id = xv.organization_id AND xv.organization_id = '" + var_unidad_organizacional + "' and almacen in ('TEXCALIDAD','CDISTEX_PT') ORDER BY FECHA_FIN DESC"
            End If
            rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               While Not rs.EOF
                     var_fecha_fin_1 = CDate(rs!fecha_fin)
                     var_dia = CStr(Day(var_fecha_fin_1))
                     var_mes = CStr(Month(var_fecha_fin_1))
                     var_año = CStr(Year(var_fecha_fin_1))
                     If Len(Trim(var_dia)) = 1 Then
                        var_dia = "0" + var_dia
                     End If
                     If Len(Trim(var_mes)) = 1 Then
                        var_mes = "0" + var_mes
                     End If
                     VAR_FECHA_MOVIMIENTO = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                     
                     
                     var_cadena = "INSERT INTO TB_ORACLE_VALUACION_DEVOLUCIONES (INTE_TEM_CONSECUTIVO, FECHA_INICIO, FECHA_FIN, AGENTE, NOMBRE_AGENTE, CLIENTE, NOMBRE_CLIENTE, CODIGO, DESCRIPCION, CANTIDAD, PRECIO, FACTURA, IMPORTE, FECHA, FOLIO_SID, PEDIDO, REFERENCIA, NUMERO_REFERENCIA)"
                     var_cadena = var_cadena + " VALUES (" + CStr(var_consecutivo) + ", " + VAR_FECHA_INICIO_TABLA + "," + VAR_FECHA_FIN_TABLA + ",'" + CStr(IIf(IsNull(rs!Agente), 0, rs!Agente)) + "','" + CStr(rs!NOMBRE_AGENTE) + "','" + CStr(rs!Cliente) + "','" + rs!nombre_cliente + "','" + CStr(IIf(IsNull(rs!ORDERED_ITEM), 0, rs!ORDERED_ITEM)) + "','" + CStr(rs!Description) + "'," + CStr(rs!PRICING_QUANTITY) + "," + CStr(IIf(IsNull(rs!unit_selling_price), 0, rs!unit_selling_price)) + ",'" + CStr(IIf(IsNull(rs!attribute11), "", rs!attribute11)) + "'," + CStr(IIf(IsNull(rs!Importe), 0, rs!Importe)) + "," + VAR_FECHA_MOVIMIENTO + ", '" + rs!FOLI0 + "'," + CStr(rs!order_number) + ",'" + IIf(IsNull(rs!Referencia), "", rs!Referencia) + "','" + CStr(IIf(IsNull(rs!numero_referencia), "", rs!numero_referencia)) + "')"
                     rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               rsaux.Open "DELETE FROM TB_ORACLE_VALUACION_DEVOLUCIONES WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND FECHA_INICIO IS NULL", cnn, adOpenDynamic, adLockOptimistic
               
               rsaux.Open "select distinct PEDIDO, cliente, referencia FROM  TB_ORACLE_VALUACION_DEVOLUCIONES where INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
               
               While Not rsaux.EOF
                     rsaux2.Open "select * from ra_customer_trx_all WHERE  ct_reference = '" + CStr(rsaux!pedido) + "' and bill_to_site_use_id = " + CStr(rsaux!Cliente) + "  and attribute7 = '" + IIf(IsNull(rsaux!Referencia), "", rsaux!Referencia) + "' and interface_header_attribute10 = '" + var_unidad_organizacional + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     'rsaux2.Open "select trx_number from ra_customer_trx_all WHERE  ct_reference = '" + CStr(rsaux!pedido) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     VAR_NOTAS_CREDITO = ""
                     While Not rsaux2.EOF
                           If VAR_NOTAS_CREDITO = "" Then
                              VAR_NOTAS_CREDITO = CStr(rsaux2!trx_number)
                           Else
                              VAR_NOTAS_CREDITO = VAR_NOTAS_CREDITO + ", " + CStr(rsaux2!trx_number)
                           End If
                           rsaux2.MoveNext
                     Wend
                     rsaux2.Close
                     rsaux2.Open "UPDATE TB_ORACLE_VALUACION_DEVOLUCIONES SET NOTA_CREDITO = '" + VAR_NOTAS_CREDITO + "' WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO = " + CStr(rsaux!pedido), cnn, adOpenDynamic, adLockOptimistic
                     rsaux.MoveNext
               Wend
               rsaux.Close
               
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_valuacion_devoluciones_detalle.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_VALUACION_DEVOLUCIONES_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Valuación de devoluciones a detalle"
               frmvistasprevias.Show 1
               Set reporte = Nothing
    
               var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  x = 0
                  If x = 1 Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_oracle_valuacion_devoluciones_detalle_excel.rpt")
                     reporte.RecordSelectionFormula = "{VW_ORACLE_VALUACION_DEVOLUCIONES_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     reporte.ExportOptions.FormatType = crEFTExcel80
                     reporte.ExportOptions.DestinationType = crEDTDiskFile
                     archivo = "c:\reportessid\valuacion_devoluciones_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     reporte.ExportOptions.DiskFileName = archivo
                     reporte.Export False
                     Set reporte = Nothing
                  Else
                     Set oexcel = CreateObject("Excel.Application")
                     Set owbook = oexcel.Workbooks.Add
                     Set osheet = owbook.Worksheets(1)
                     osheet.Name = "VALUACION DE DEVOLUCIONES"
                     Screen.MousePointer = vbHourglass
                     iFila = 1
                     ifila2 = 1
                     icol2 = 1
                     iCol = 1
                     var_cadena = " SELECT FECHA_INICIO,FECHA_FIN, AGENTE, NOMBRE_AGENTE, CLIENTE, NOMBRE_CLIENTE, CODIGO, DESCRIPCION, CANTIDAD, PRECIO, FACTURA, IMPORTE FECHA, FOLIO_SID, PEDIDO, NOTA_CREDITO, REFERENCIA, NUMERO_REFERENCIA FROM VW_ORACLE_VALUACION_DEVOLUCIONES_DETALLE WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo)
                     rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     For i = 0 To rsaux10.Fields.Count - 1
                         osheet.Cells(iFila, i + 1) = rsaux10.Fields(i).Name
                     Next
                     iFila = iFila + 1
                     With osheet
                         ' carga los registros del recordset
                         .Cells(iFila, iCol).CopyFromRecordset rsaux10
                         'oExcel.Columns(1).Select
                         'oExcel.Selection.NumberFormat = "#,##0.00"
                         'oExcel.Columns(1).Select
                         'oExcel.Selection.Font.Color = vbRed
                         .Columns.AutoFit ' ajusta el ancho de las columnas
                     End With
                     owbook.SaveAs "c:\reportessid\valuacion_devoluciones_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     oexcel.Visible = True
                     Set oexcel = Nothing
                     Screen.MousePointer = vbDefault
                     rsaux10.Close
                  
                  
                  End If
                  MsgBox "Se a terminado de guardar el archivo " + archivo
               End If
            
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_valuacion_devoluciones_concentrado.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_VALUACION_DEVOLUCIONES_CONCENTRADO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Pedidos cargados"
               frmvistasprevias.Show 1
               Set reporte = Nothing
    
               var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  x = 0
                  If x = 1 Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_oracle_valuacion_devoluciones_concentrado_excel.rpt")
                     reporte.RecordSelectionFormula = "{VW_ORACLE_VALUACION_DEVOLUCIONES_CONCENTRADO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     reporte.ExportOptions.FormatType = crEFTExcel80
                     reporte.ExportOptions.DestinationType = crEDTDiskFile
                     archivo = "c:\reportessid\valuacion_devoluciones_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     reporte.ExportOptions.DiskFileName = archivo
                     reporte.Export False
                     Set reporte = Nothing
                     MsgBox "Se a terminado de guardar el archivo " + archivo
                  Else
                     Set oexcel = CreateObject("Excel.Application")
                     Set owbook = oexcel.Workbooks.Add
                     Set osheet = owbook.Worksheets(1)
                     osheet.Name = "VALUACION DE DEVOLUCIONES"
                     Screen.MousePointer = vbHourglass
                     iFila = 1
                     ifila2 = 1
                     icol2 = 1
                     iCol = 1
                     var_cadena = "SELECT FECHA_INICIO, FECHA_FIN, AGENTE, NOMBRE_AGENTE, CLIENTE, NOMBRE_CLIENTE, IMPORTE FECHA, FOLIO_SID, PEDIDO, CANTIDAD, NOTA_CREDITO, REFERENCIA, NUMERO_REFERENCIA FROM VW_ORACLE_VALUACION_DEVOLUCIONES_CONCENTRADO WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo)
                     rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     For i = 0 To rsaux10.Fields.Count - 1
                         osheet.Cells(iFila, i + 1) = rsaux10.Fields(i).Name
                     Next
                     iFila = iFila + 1
                     With osheet
                         ' carga los registros del recordset
                         .Cells(iFila, iCol).CopyFromRecordset rsaux10
                         'oExcel.Columns(1).Select
                         'oExcel.Selection.NumberFormat = "#,##0.00"
                         'oExcel.Columns(1).Select
                         'oExcel.Selection.Font.Color = vbRed
                         .Columns.AutoFit ' ajusta el ancho de las columnas
                     End With
                     archivo = "c:\reportessid\valuacion_devoluciones_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     owbook.SaveAs archivo
                     oexcel.Visible = True
                     Set oexcel = Nothing
                     Screen.MousePointer = vbDefault
                     rsaux10.Close
                     MsgBox "Se a terminado de guardar el archivo " + archivo
                  End If
               End If
            
            Else
               MsgBox "No existen pedidos para el periodo indicado", vbOKOnly, "ATENCION"
            End If
            rs.Close
            rs.Open "delete from TB_oracle_valuacion_Devoluciones WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "La fecha de inicio debe de ser menor o igual a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de Inicio incorrecta", vbOKOnly, "ATENCION"
   End If

End Sub

Private Sub cmd_general_Click()
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim año As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_consecutivo As Integer
   Dim var_afectacion_movimiento As String
   Dim var_vistas As String
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            cnn.CommandTimeout = 720
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) as numero from tb_oracle_valuacion_devoluciones", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!numero), 0, rs!numero)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into tb_oracle_valuacion_devoluciones (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            var_dia = CStr(Day(CDate(txt_inicio)))
            var_mes = CStr(Month(CDate(txt_inicio)))
            var_año = CStr(Year(CDate(txt_inicio)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            
            var_fecha_inicio = var_dia + "-" + var_mes + "-" + var_año
            VAR_FECHA_INICIO_TABLA = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            var_fecha_fin_1 = CDate(txt_fin) + 1
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = var_dia + "-" + var_mes + "-" + var_año
            
            
            
            var_fecha_fin_1 = CDate(txt_fin)
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            VAR_FECHA_FIN_TABLA = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
             
            rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
              
            var_cadena = " select distinct OL.line_id, agente, nombre_agente, cliente, nombre_cliente, 'SIDDC_'||CAST(NUMERO AS VARCHAR(50)) AS FOLI0, FECHA_FIN, ORDER_NUMBER, ORDERED_ITEM, ORDER_QUANTITY_UOM, pricing_quantity, UNIT_SELLING_PRICE, ol.attribute11, pricing_quantity *  unit_selling_price as importe, xv.description, referencia from xxvia_tb_devoluciones_clientes x, OE_ORDER_HEADERS_ALL  OE, OE_ORDER_LINES_ALL OL, xxvia_system_items_b xv where trunc(to_date(x.fecha_fin,'DD/MM/YYYY'))  >= to_date('" + Me.txt_inicio + "','DD-MM-YYYY')  AND trunc(to_date(x.fecha_fin,'DD/MM/YYYY')) < TO_DATE('" + Me.txt_fin + "','DD-MM-YYYY') + 1 AND ESTATUS = 'I' AND MOVIMIENTO = 'DC' AND OE.orig_sys_document_ref = 'SIDDC_'||CAST(NUMERO AS VARCHAR(50)) AND OE.HEADER_ID = OL.HEADER_ID and ol.inventory_item_id = xv.inventory_item_id and oe.ship_from_org_id = xv.organization_id AND xv.organization_id = '" + var_unidad_organizacional + "' and almacen = 'CDI_ALMPT' ORDER BY FECHA_FIN DESC"
            
            rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               While Not rs.EOF
                     var_fecha_fin_1 = CDate(rs!fecha_fin)
                     var_dia = CStr(Day(var_fecha_fin_1))
                     var_mes = CStr(Month(var_fecha_fin_1))
                     var_año = CStr(Year(var_fecha_fin_1))
                     If Len(Trim(var_dia)) = 1 Then
                        var_dia = "0" + var_dia
                     End If
                     If Len(Trim(var_mes)) = 1 Then
                        var_mes = "0" + var_mes
                     End If
                     VAR_FECHA_MOVIMIENTO = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                     
                     
                     var_cadena = "INSERT INTO TB_ORACLE_VALUACION_DEVOLUCIONES (INTE_TEM_CONSECUTIVO, FECHA_INICIO, FECHA_FIN, AGENTE, NOMBRE_AGENTE, CLIENTE, NOMBRE_CLIENTE, CODIGO, DESCRIPCION, CANTIDAD, PRECIO, FACTURA, IMPORTE, FECHA, FOLIO_SID, PEDIDO, referencia)"
                     var_cadena = var_cadena + " VALUES (" + CStr(var_consecutivo) + ", " + VAR_FECHA_INICIO_TABLA + "," + VAR_FECHA_FIN_TABLA + ",'" + CStr(IIf(IsNull(rs!Agente), 0, rs!Agente)) + "','" + CStr(rs!NOMBRE_AGENTE) + "','" + CStr(rs!Cliente) + "','" + rs!nombre_cliente + "','" + CStr(IIf(IsNull(rs!ORDERED_ITEM), 0, rs!ORDERED_ITEM)) + "','" + CStr(rs!Description) + "'," + CStr(rs!PRICING_QUANTITY) + "," + CStr(rs!unit_selling_price) + ",'" + CStr(IIf(IsNull(rs!attribute11), "", rs!attribute11)) + "'," + CStr(rs!Importe) + "," + VAR_FECHA_MOVIMIENTO + ", '" + rs!FOLI0 + "'," + CStr(rs!order_number) + ",'" + IIf(IsNull(rs!Referencia), "", rs!Referencia) + "')"
                     rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               rsaux.Open "DELETE FROM TB_ORACLE_VALUACION_DEVOLUCIONES WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND FECHA_INICIO IS NULL", cnn, adOpenDynamic, adLockOptimistic
               
               rsaux.Open "select distinct PEDIDO, cliente, referencia FROM  TB_ORACLE_VALUACION_DEVOLUCIONES where INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
               
               While Not rsaux.EOF
                     rsaux2.Open "select * from ra_customer_trx_all WHERE  ct_reference = '" + CStr(rsaux!pedido) + "' and bill_to_site_use_id = " + CStr(rsaux!Cliente) + "  and attribute7 = '" + IIf(IsNull(rsaux!Referencia), "", rsaux!Referencia) + "' and interface_header_attribute10 = '" + var_unidad_organizacional + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     VAR_NOTAS_CREDITO = ""
                     While Not rsaux2.EOF
                           If VAR_NOTAS_CREDITO = "" Then
                              VAR_NOTAS_CREDITO = CStr(rsaux2!trx_number)
                           Else
                              VAR_NOTAS_CREDITO = VAR_NOTAS_CREDITO + ", " + CStr(rsaux2!trx_number)
                           End If
                           rsaux2.MoveNext
                     Wend
                     rsaux2.Close
                     rsaux2.Open "UPDATE TB_ORACLE_VALUACION_DEVOLUCIONES SET NOTA_CREDITO = '" + VAR_NOTAS_CREDITO + "' WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO = " + CStr(rsaux!pedido), cnn, adOpenDynamic, adLockOptimistic
                     rsaux.MoveNext
               Wend
               rsaux.Close
               
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_valuacion_devoluciones_detalle.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_VALUACION_DEVOLUCIONES_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Valuación de devoluciones a detalle"
               frmvistasprevias.Show 1
               Set reporte = Nothing
    
               var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  x = 0
                  If x = 1 Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_oracle_valuacion_devoluciones_detalle_excel.rpt")
                     reporte.RecordSelectionFormula = "{VW_ORACLE_VALUACION_DEVOLUCIONES_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     reporte.ExportOptions.FormatType = crEFTExcel80
                     reporte.ExportOptions.DestinationType = crEDTDiskFile
                     archivo = "c:\reportessid\valuacion_devoluciones_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     reporte.ExportOptions.DiskFileName = archivo
                     reporte.Export False
                     Set reporte = Nothing
                  Else
                     Set oexcel = CreateObject("Excel.Application")
                     Set owbook = oexcel.Workbooks.Add
                     Set osheet = owbook.Worksheets(1)
                     osheet.Name = "VALUACION DE DEVOLUCIONES"
                     Screen.MousePointer = vbHourglass
                     iFila = 1
                     ifila2 = 1
                     icol2 = 1
                     iCol = 1
                     var_cadena = " SELECT FECHA_INICIO,FECHA_FIN, AGENTE, NOMBRE_AGENTE, CLIENTE, NOMBRE_CLIENTE, CODIGO, DESCRIPCION, CANTIDAD, PRECIO, FACTURA, IMPORTE FECHA, FOLIO_SID, PEDIDO, NOTA_CREDITO, REFERENCIA FROM VW_ORACLE_VALUACION_DEVOLUCIONES_DETALLE WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo)
                     rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     For i = 0 To rsaux10.Fields.Count - 1
                         osheet.Cells(iFila, i + 1) = rsaux10.Fields(i).Name
                     Next
                     iFila = iFila + 1
                     With osheet
                         ' carga los registros del recordset
                         .Cells(iFila, iCol).CopyFromRecordset rsaux10
                         'oExcel.Columns(1).Select
                         'oExcel.Selection.NumberFormat = "#,##0.00"
                         'oExcel.Columns(1).Select
                         'oExcel.Selection.Font.Color = vbRed
                         .Columns.AutoFit ' ajusta el ancho de las columnas
                     End With
                     owbook.SaveAs "c:\reportessid\valuacion_devoluciones_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     oexcel.Visible = True
                     Set oexcel = Nothing
                     Screen.MousePointer = vbDefault
                     rsaux10.Close
                  End If
                  MsgBox "Se a terminado de guardar el archivo " + archivo
               End If
            
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_valuacion_devoluciones_concentrado.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_VALUACION_DEVOLUCIONES_CONCENTRADO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Pedidos cargados"
               frmvistasprevias.Show 1
               Set reporte = Nothing
    
               var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  x = 0
                  If x = 1 Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_oracle_valuacion_devoluciones_concentrado_excel.rpt")
                     reporte.RecordSelectionFormula = "{VW_ORACLE_VALUACION_DEVOLUCIONES_CONCENTRADO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     reporte.ExportOptions.FormatType = crEFTExcel80
                     reporte.ExportOptions.DestinationType = crEDTDiskFile
                     archivo = "c:\reportessid\valuacion_devoluciones_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     reporte.ExportOptions.DiskFileName = archivo
                     reporte.Export False
                     Set reporte = Nothing
                     MsgBox "Se a terminado de guardar el archivo " + archivo
                  Else
                     Set oexcel = CreateObject("Excel.Application")
                     Set owbook = oexcel.Workbooks.Add
                     Set osheet = owbook.Worksheets(1)
                     osheet.Name = "VALUACION DE DEVOLUCIONES"
                     Screen.MousePointer = vbHourglass
                     iFila = 1
                     ifila2 = 1
                     icol2 = 1
                     iCol = 1
                     var_cadena = "SELECT FECHA_INICIO, FECHA_FIN, AGENTE, NOMBRE_AGENTE, CLIENTE, NOMBRE_CLIENTE, IMPORTE FECHA, FOLIO_SID, PEDIDO, CANTIDAD, NOTA_CREDITO, REFERENCIA FROM VW_ORACLE_VALUACION_DEVOLUCIONES_CONCENTRADO WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo)
                     rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     For i = 0 To rsaux10.Fields.Count - 1
                         osheet.Cells(iFila, i + 1) = rsaux10.Fields(i).Name
                     Next
                     iFila = iFila + 1
                     With osheet
                         ' carga los registros del recordset
                         .Cells(iFila, iCol).CopyFromRecordset rsaux10
                         'oExcel.Columns(1).Select
                         'oExcel.Selection.NumberFormat = "#,##0.00"
                         'oExcel.Columns(1).Select
                         'oExcel.Selection.Font.Color = vbRed
                         .Columns.AutoFit ' ajusta el ancho de las columnas
                     End With
                     archivo = "c:\reportessid\valuacion_devoluciones_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     owbook.SaveAs archivo
                     oexcel.Visible = True
                     Set oexcel = Nothing
                     Screen.MousePointer = vbDefault
                     rsaux10.Close
                     MsgBox "Se a terminado de guardar el archivo " + archivo
                  End If
               End If
            
            Else
               MsgBox "No existen pedidos para el periodo indicado", vbOKOnly, "ATENCION"
            End If
            rs.Close
            rs.Open "delete from TB_oracle_valuacion_Devoluciones WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "La fecha de inicio debe de ser menor o igual a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de Inicio incorrecta", vbOKOnly, "ATENCION"
   End If

End Sub

Private Sub cmd_imprimir_Click()
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim año As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_consecutivo As Integer
   Dim var_afectacion_movimiento As String
   Dim var_vistas As String
   Dim iFila As Long, iCol As Integer, i As Integer
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            cnn.CommandTimeout = 720
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) as numero from tb_oracle_valuacion_devoluciones", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!numero), 0, rs!numero)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into tb_oracle_valuacion_devoluciones (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            var_dia = CStr(Day(CDate(txt_inicio)))
            var_mes = CStr(Month(CDate(txt_inicio)))
            var_año = CStr(Year(CDate(txt_inicio)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            
            var_fecha_inicio = var_dia + "-" + var_mes + "-" + var_año
            VAR_FECHA_INICIO_TABLA = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            var_fecha_fin_1 = CDate(txt_fin) + 1
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = var_dia + "-" + var_mes + "-" + var_año
            
            
            
            var_fecha_fin_1 = CDate(txt_fin)
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            VAR_FECHA_FIN_TABLA = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
             
            rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
              
            var_cadena = " select distinct OL.line_id, agente, nombre_agente, cliente, nombre_cliente, 'SIDDC_'||CAST(NUMERO AS VARCHAR(50)) AS FOLI0, FECHA_FIN, ORDER_NUMBER, ORDERED_ITEM, ORDER_QUANTITY_UOM, pricing_quantity, UNIT_SELLING_PRICE, ol.attribute11, pricing_quantity *  unit_selling_price as importe, xv.description, x.referencia, x.numero_referencia from xxvia_tb_devoluciones_clientes x, OE_ORDER_HEADERS_ALL  OE, OE_ORDER_LINES_ALL OL, xxvia_system_items_b xv where trunc(to_date(x.fecha_fin,'DD/MM/YYYY'))  >= to_date('" + Me.txt_inicio + "','DD-MM-YYYY')  AND trunc(to_date(x.fecha_fin,'DD/MM/YYYY')) < TO_DATE('" + Me.txt_fin + "','DD-MM-YYYY') + 1 AND ESTATUS = 'I' AND MOVIMIENTO = 'DC' AND OE.orig_sys_document_ref = 'SIDDC_'||CAST(NUMERO AS VARCHAR(50)) AND OE.HEADER_ID = OL.HEADER_ID and ol.inventory_item_id = xv.inventory_item_id and oe.ship_from_org_id = xv.organization_id AND xv.organization_id = '" + var_unidad_organizacional + "' union "
            var_cadena = var_cadena + " select distinct OL.line_id, agente, nombre_agente, cliente, nombre_cliente, 'SIDDCS_'||CAST(NUMERO AS VARCHAR(50)) AS FOLI0, FECHA_FIN, ORDER_NUMBER, ORDERED_ITEM, ORDER_QUANTITY_UOM, pricing_quantity, UNIT_SELLING_PRICE, ol.attribute11, pricing_quantity *  unit_selling_price as importe, xv.description, X.REFERENCIA, X.NUMERO_REFERENCIA from xxvia_tb_devoluciones_clientes x, OE_ORDER_HEADERS_ALL  OE, OE_ORDER_LINES_ALL OL, xxvia_system_items_b xv where trunc(to_date(x.fecha_fin,'DD/MM/YYYY'))  >= to_date('" + Me.txt_inicio + "','DD-MM-YYYY')  AND trunc(to_date(x.fecha_fin,'DD/MM/YYYY')) < TO_DATE('" + Me.txt_fin + "','DD-MM-YYYY') + 1 AND ESTATUS = 'I' AND MOVIMIENTO = 'DCS' AND OE.orig_sys_document_ref = 'SIDDCS_'||CAST(NUMERO AS VARCHAR(50)) AND OE.HEADER_ID = OL.HEADER_ID and ol.inventory_item_id = xv.inventory_item_id and oe.ship_from_org_id = xv.organization_id AND xv.organization_id = '" + var_unidad_organizacional + "' ORDER BY FECHA_FIN "
            rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               While Not rs.EOF
                     var_fecha_fin_1 = CDate(rs!fecha_fin)
                     var_dia = CStr(Day(var_fecha_fin_1))
                     var_mes = CStr(Month(var_fecha_fin_1))
                     var_año = CStr(Year(var_fecha_fin_1))
                     If Len(Trim(var_dia)) = 1 Then
                        var_dia = "0" + var_dia
                     End If
                     If Len(Trim(var_mes)) = 1 Then
                        var_mes = "0" + var_mes
                     End If
                     VAR_FECHA_MOVIMIENTO = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                     
                     
                     var_cadena = "INSERT INTO TB_ORACLE_VALUACION_DEVOLUCIONES (INTE_TEM_CONSECUTIVO, FECHA_INICIO, FECHA_FIN, AGENTE, NOMBRE_AGENTE, CLIENTE, NOMBRE_CLIENTE, CODIGO, DESCRIPCION, CANTIDAD, PRECIO, FACTURA, IMPORTE, FECHA, FOLIO_SID, PEDIDO, REFERENCIA, NUMERO_REFERENCIA)"
                     var_cadena = var_cadena + " VALUES (" + CStr(var_consecutivo) + ", " + VAR_FECHA_INICIO_TABLA + "," + VAR_FECHA_FIN_TABLA + ",'" + CStr(IIf(IsNull(rs!Agente), 0, rs!Agente)) + "','" + CStr(rs!NOMBRE_AGENTE) + "','" + CStr(rs!Cliente) + "','" + rs!nombre_cliente + "','" + CStr(IIf(IsNull(rs!ORDERED_ITEM), 0, rs!ORDERED_ITEM)) + "','" + CStr(rs!Description) + "'," + CStr(rs!PRICING_QUANTITY) + "," + CStr(IIf(IsNull(rs!unit_selling_price), 0, rs!unit_selling_price)) + ",'" + CStr(IIf(IsNull(rs!attribute11), "", rs!attribute11)) + "'," + CStr(IIf(IsNull(rs!Importe), 0, rs!Importe)) + "," + VAR_FECHA_MOVIMIENTO + ", '" + rs!FOLI0 + "'," + CStr(rs!order_number) + ",'" + IIf(IsNull(rs!Referencia), "", rs!Referencia) + "','" + CStr(IIf(IsNull(rs!numero_referencia), "", rs!numero_referencia)) + "')"
                     
                     rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               rsaux.Open "DELETE FROM TB_ORACLE_VALUACION_DEVOLUCIONES WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND FECHA_INICIO IS NULL", cnn, adOpenDynamic, adLockOptimistic
               
               rsaux.Open "select distinct PEDIDO, cliente, referencia FROM  TB_ORACLE_VALUACION_DEVOLUCIONES where INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
               
               While Not rsaux.EOF
                     rsaux2.Open "select * from ra_customer_trx_all WHERE  ct_reference = '" + CStr(rsaux!pedido) + "' and bill_to_site_use_id = " + CStr(rsaux!Cliente), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     VAR_NOTAS_CREDITO = ""
                     While Not rsaux2.EOF
                           If VAR_NOTAS_CREDITO = "" Then
                              VAR_NOTAS_CREDITO = CStr(rsaux2!trx_number)
                           Else
                              VAR_NOTAS_CREDITO = VAR_NOTAS_CREDITO + ", " + CStr(rsaux2!trx_number)
                           End If
                           rsaux2.MoveNext
                     Wend
                     rsaux2.Close
                     rsaux2.Open "UPDATE TB_ORACLE_VALUACION_DEVOLUCIONES SET NOTA_CREDITO = '" + VAR_NOTAS_CREDITO + "' WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO = " + CStr(rsaux!pedido), cnn, adOpenDynamic, adLockOptimistic
                     rsaux.MoveNext
               Wend
               rsaux.Close
               
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_valuacion_devoluciones_detalle.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_VALUACION_DEVOLUCIONES_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Valuación de devoluciones a detalle"
               frmvistasprevias.Show 1
               Set reporte = Nothing
    
               var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  x = 1
                  If x = 0 Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_oracle_valuacion_devoluciones_detalle.rpt")
                     reporte.RecordSelectionFormula = "{VW_ORACLE_VALUACION_DEVOLUCIONES_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     reporte.ExportOptions.FormatType = crEFTExcel80
                     reporte.ExportOptions.DestinationType = crEDTDiskFile
                     archivo = "c:\reportessid\valuacion_devoluciones_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     reporte.ExportOptions.DiskFileName = archivo
                     reporte.Export False
                     Set reporte = Nothing
                     MsgBox "Se a terminado de guardar el archivo " + archivo
                  Else
                     Set oexcel = CreateObject("Excel.Application")
                     Set owbook = oexcel.Workbooks.Add
                     Set osheet = owbook.Worksheets(1)
                     osheet.Name = "VALUACION DE DEVOLUCIONES"
                     Screen.MousePointer = vbHourglass
                     iFila = 1
                     ifila2 = 1
                     icol2 = 1
                     iCol = 1
                     var_cadena = " SELECT FECHA_INICIO, FECHA_FIN, AGENTE, NOMBRE_AGENTE, CLIENTE, NOMBRE_CLIENTE, CODIGO, DESCRIPCION, CANTIDAD, PRECIO, FACTURA, IMPORTE, FECHA, FOLIO_SID, PEDIDO, NOTA_CREDITO, REFERENCIA, NUMERO_REFERENCIA FROM VW_ORACLE_VALUACION_DEVOLUCIONES_DETALLE WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo)
                     rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     For i = 0 To rsaux10.Fields.Count - 1
                         osheet.Cells(iFila, i + 1) = rsaux10.Fields(i).Name
                     Next
                     iFila = iFila + 1
                     With osheet
                         ' carga los registros del recordset
                         .Cells(iFila, iCol).CopyFromRecordset rsaux10
                         'oExcel.Columns(1).Select
                         'oExcel.Selection.NumberFormat = "#,##0.00"
                         'oExcel.Columns(1).Select
                         'oExcel.Selection.Font.Color = vbRed
                         .Columns.AutoFit ' ajusta el ancho de las columnas
                     End With
                     archivo = "c:\reportessid\valuacion_devoluciones_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     owbook.SaveAs archivo
                     oexcel.Visible = True
                     Set oexcel = Nothing
                     Screen.MousePointer = vbDefault
                     rsaux10.Close
                     MsgBox "Se a terminado de guardar el archivo " + archivo
                  End If
               End If
            
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_valuacion_devoluciones_concentrado.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_VALUACION_DEVOLUCIONES_CONCENTRADO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Pedidos cargados"
               frmvistasprevias.Show 1
               Set reporte = Nothing
    
               var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  x = 1
                  If x = 0 Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_oracle_valuacion_devoluciones_concentrado.rpt")
                     reporte.RecordSelectionFormula = "{VW_ORACLE_VALUACION_DEVOLUCIONES_CONCENTRADO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     reporte.ExportOptions.FormatType = crEFTExcel80
                     reporte.ExportOptions.DestinationType = crEDTDiskFile
                     archivo = "c:\reportessid\valuacion_devoluciones_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     reporte.ExportOptions.DiskFileName = archivo
                     reporte.Export False
                     Set reporte = Nothing
                     MsgBox "Se a terminado de guardar el archivo " + archivo
                  Else
                     Set oexcel = CreateObject("Excel.Application")
                     Set owbook = oexcel.Workbooks.Add
                     Set osheet = owbook.Worksheets(1)
                     osheet.Name = "VALUACION DE DEVOLUCIONES"
                     Screen.MousePointer = vbHourglass
                     iFila = 1
                     ifila2 = 1
                     icol2 = 1
                     iCol = 1
                     var_cadena = "SELECT FECHA_INICIO, FECHA_FIN, AGENTE, NOMBRE_AGENTE, CLIENTE, NOMBRE_CLIENTE, IMPORTE, FECHA, FOLIO_SID, PEDIDO, CANTIDAD, NOTA_CREDITO, REFERENCIA, NUMERO_REFERENCIA FROM VW_ORACLE_VALUACION_DEVOLUCIONES_CONCENTRADO WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo)
                     rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     For i = 0 To rsaux10.Fields.Count - 1
                         osheet.Cells(iFila, i + 1) = rsaux10.Fields(i).Name
                     Next
                     iFila = iFila + 1
                     With osheet
                         ' carga los registros del recordset
                         .Cells(iFila, iCol).CopyFromRecordset rsaux10
                         'oExcel.Columns(1).Select
                         'oExcel.Selection.NumberFormat = "#,##0.00"
                         'oExcel.Columns(1).Select
                         'oExcel.Selection.Font.Color = vbRed
                         .Columns.AutoFit ' ajusta el ancho de las columnas
                     End With
                     archivo = "c:\reportessid\valuacion_devoluciones_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     owbook.SaveAs archivo
                     oexcel.Visible = True
                     Set oexcel = Nothing
                     Screen.MousePointer = vbDefault
                     rsaux10.Close
                     MsgBox "Se a terminado de guardar el archivo " + archivo
                  End If
               End If
            
            Else
               MsgBox "No existen pedidos para el periodo indicado", vbOKOnly, "ATENCION"
            End If
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open "delete from TB_oracle_valuacion_Devoluciones WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "La fecha de inicio debe de ser menor o igual a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de Inicio incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim año As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_consecutivo As Integer
   Dim var_afectacion_movimiento As String
   Dim var_vistas As String
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            cnn.CommandTimeout = 720
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) as numero from tb_oracle_valuacion_devoluciones", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!numero), 0, rs!numero)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into tb_oracle_valuacion_devoluciones (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            var_dia = CStr(Day(CDate(txt_inicio)))
            var_mes = CStr(Month(CDate(txt_inicio)))
            var_año = CStr(Year(CDate(txt_inicio)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            
            var_fecha_inicio = var_dia + "-" + var_mes + "-" + var_año
            VAR_FECHA_INICIO_TABLA = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            var_fecha_fin_1 = CDate(txt_fin) + 1
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = var_dia + "-" + var_mes + "-" + var_año
            
            
            
            var_fecha_fin_1 = CDate(txt_fin)
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            VAR_FECHA_FIN_TABLA = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
             
            rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
              
            var_cadena = " select distinct OL.line_id, agente, nombre_agente, cliente, nombre_cliente, 'SIDDC_'||CAST(NUMERO AS VARCHAR(50)) AS FOLI0, FECHA_FIN, ORDER_NUMBER, ORDERED_ITEM, ORDER_QUANTITY_UOM, pricing_quantity, UNIT_SELLING_PRICE, ol.attribute11, pricing_quantity *  unit_selling_price as importe, xv.description from xxvia_tb_devoluciones_clientes x, OE_ORDER_HEADERS_ALL  OE, OE_ORDER_LINES_ALL OL, xxvia_system_items_b xv where trunc(to_date(x.fecha_fin,'DD/MM/YYYY'))  >= to_date('" + Me.txt_inicio + "','DD-MM-YYYY')  AND trunc(to_date(x.fecha_fin,'DD/MM/YYYY')) < TO_DATE('" + Me.txt_fin + "','DD-MM-YYYY') + 1 AND ESTATUS = 'I' AND MOVIMIENTO = 'DC' AND OE.orig_sys_document_ref = 'SIDDC_'||CAST(NUMERO AS VARCHAR(50)) AND OE.HEADER_ID = OL.HEADER_ID and ol.inventory_item_id = xv.inventory_item_id and oe.ship_from_org_id = xv.organization_id AND xv.organization_id = '" + var_unidad_organizacional + "' and almacen = 'RECEPMEX' ORDER BY FECHA_FIN DESC"
            
            rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               While Not rs.EOF
                     var_fecha_fin_1 = CDate(rs!fecha_fin)
                     var_dia = CStr(Day(var_fecha_fin_1))
                     var_mes = CStr(Month(var_fecha_fin_1))
                     var_año = CStr(Year(var_fecha_fin_1))
                     If Len(Trim(var_dia)) = 1 Then
                        var_dia = "0" + var_dia
                     End If
                     If Len(Trim(var_mes)) = 1 Then
                        var_mes = "0" + var_mes
                     End If
                     VAR_FECHA_MOVIMIENTO = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                     
                     
                     var_cadena = "INSERT INTO TB_ORACLE_VALUACION_DEVOLUCIONES (INTE_TEM_CONSECUTIVO, FECHA_INICIO, FECHA_FIN, AGENTE, NOMBRE_AGENTE, CLIENTE, NOMBRE_CLIENTE, CODIGO, DESCRIPCION, CANTIDAD, PRECIO, FACTURA, IMPORTE, FECHA, FOLIO_SID, PEDIDO)"
                     var_cadena = var_cadena + " VALUES (" + CStr(var_consecutivo) + ", " + VAR_FECHA_INICIO_TABLA + "," + VAR_FECHA_FIN_TABLA + ",'" + CStr(IIf(IsNull(rs!Agente), 0, rs!Agente)) + "','" + CStr(rs!NOMBRE_AGENTE) + "','" + CStr(rs!Cliente) + "','" + rs!nombre_cliente + "','" + CStr(IIf(IsNull(rs!ORDERED_ITEM), 0, rs!ORDERED_ITEM)) + "','" + CStr(rs!Description) + "'," + CStr(rs!PRICING_QUANTITY) + "," + CStr(rs!unit_selling_price) + ",'" + CStr(IIf(IsNull(rs!attribute11), "", rs!attribute11)) + "'," + CStr(rs!Importe) + "," + VAR_FECHA_MOVIMIENTO + ", '" + rs!FOLI0 + "'," + CStr(rs!order_number) + ")"
                     rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               rsaux.Open "DELETE FROM TB_ORACLE_VALUACION_DEVOLUCIONES WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND FECHA_INICIO IS NULL", cnn, adOpenDynamic, adLockOptimistic
               
               rsaux.Open "select distinct PEDIDO, cliente, referencia FROM  TB_ORACLE_VALUACION_DEVOLUCIONES where INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
               
               While Not rsaux.EOF
                     rsaux2.Open "select * from ra_customer_trx_all WHERE  ct_reference = '" + CStr(rsaux!pedido) + "' and bill_to_site_use_id = " + CStr(rsaux!Cliente), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     VAR_NOTAS_CREDITO = ""
                     While Not rsaux2.EOF
                           If VAR_NOTAS_CREDITO = "" Then
                              VAR_NOTAS_CREDITO = CStr(rsaux2!trx_number)
                           Else
                              VAR_NOTAS_CREDITO = VAR_NOTAS_CREDITO + ", " + CStr(rsaux2!trx_number)
                           End If
                           rsaux2.MoveNext
                     Wend
                     rsaux2.Close
                     rsaux2.Open "UPDATE TB_ORACLE_VALUACION_DEVOLUCIONES SET NOTA_CREDITO = '" + VAR_NOTAS_CREDITO + "' WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO = " + CStr(rsaux!pedido), cnn, adOpenDynamic, adLockOptimistic
                     rsaux.MoveNext
               Wend
               rsaux.Close
               
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_valuacion_devoluciones_detalle.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_VALUACION_DEVOLUCIONES_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Valuación de devoluciones a detalle"
               frmvistasprevias.Show 1
               Set reporte = Nothing
    
               var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  x = 0
                  If x = 1 Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_oracle_valuacion_devoluciones_detalle_excel.rpt")
                     reporte.RecordSelectionFormula = "{VW_ORACLE_VALUACION_DEVOLUCIONES_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     reporte.ExportOptions.FormatType = crEFTExcel80
                     reporte.ExportOptions.DestinationType = crEDTDiskFile
                     archivo = "c:\reportessid\valuacion_devoluciones_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     reporte.ExportOptions.DiskFileName = archivo
                     reporte.Export False
                     Set reporte = Nothing
                  Else
                     Set oexcel = CreateObject("Excel.Application")
                     Set owbook = oexcel.Workbooks.Add
                     Set osheet = owbook.Worksheets(1)
                     osheet.Name = "VALUACION DE DEVOLUCIONES"
                     Screen.MousePointer = vbHourglass
                     iFila = 1
                     ifila2 = 1
                     icol2 = 1
                     iCol = 1
                     var_cadena = " SELECT FECHA_INICIO,FECHA_FIN, AGENTE, NOMBRE_AGENTE, CLIENTE, NOMBRE_CLIENTE, CODIGO, DESCRIPCION, CANTIDAD, PRECIO, FACTURA, IMPORTE FECHA, FOLIO_SID, PEDIDO, NOTA_CREDITO, REFERENCIA FROM VW_ORACLE_VALUACION_DEVOLUCIONES_DETALLE WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo)
                     rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     For i = 0 To rsaux10.Fields.Count - 1
                         osheet.Cells(iFila, i + 1) = rsaux10.Fields(i).Name
                     Next
                     iFila = iFila + 1
                     With osheet
                         ' carga los registros del recordset
                         .Cells(iFila, iCol).CopyFromRecordset rsaux10
                         'oExcel.Columns(1).Select
                         'oExcel.Selection.NumberFormat = "#,##0.00"
                         'oExcel.Columns(1).Select
                         'oExcel.Selection.Font.Color = vbRed
                         .Columns.AutoFit ' ajusta el ancho de las columnas
                     End With
                     owbook.SaveAs "c:\reportessid\valuacion_devoluciones_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     oexcel.Visible = True
                     Set oexcel = Nothing
                     Screen.MousePointer = vbDefault
                     rsaux10.Close
                  End If
                  MsgBox "Se a terminado de guardar el archivo " + archivo
               End If
            
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_valuacion_devoluciones_concentrado.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_VALUACION_DEVOLUCIONES_CONCENTRADO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Pedidos cargados"
               frmvistasprevias.Show 1
               Set reporte = Nothing
    
               var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  x = 0
                  If x = 1 Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_oracle_valuacion_devoluciones_concentrado_excel.rpt")
                     reporte.RecordSelectionFormula = "{VW_ORACLE_VALUACION_DEVOLUCIONES_CONCENTRADO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     reporte.ExportOptions.FormatType = crEFTExcel80
                     reporte.ExportOptions.DestinationType = crEDTDiskFile
                     archivo = "c:\reportessid\valuacion_devoluciones_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     reporte.ExportOptions.DiskFileName = archivo
                     reporte.Export False
                     Set reporte = Nothing
                     MsgBox "Se a terminado de guardar el archivo " + archivo
                  Else
                     Set oexcel = CreateObject("Excel.Application")
                     Set owbook = oexcel.Workbooks.Add
                     Set osheet = owbook.Worksheets(1)
                     osheet.Name = "VALUACION DE DEVOLUCIONES"
                     Screen.MousePointer = vbHourglass
                     iFila = 1
                     ifila2 = 1
                     icol2 = 1
                     iCol = 1
                     var_cadena = "SELECT FECHA_INICIO, FECHA_FIN, AGENTE, NOMBRE_AGENTE, CLIENTE, NOMBRE_CLIENTE, IMPORTE FECHA, FOLIO_SID, PEDIDO, CANTIDAD, NOTA_CREDITO, REFERENCIA FROM VW_ORACLE_VALUACION_DEVOLUCIONES_CONCENTRADO WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo)
                     rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     For i = 0 To rsaux10.Fields.Count - 1
                         osheet.Cells(iFila, i + 1) = rsaux10.Fields(i).Name
                     Next
                     iFila = iFila + 1
                     With osheet
                         ' carga los registros del recordset
                         .Cells(iFila, iCol).CopyFromRecordset rsaux10
                         'oExcel.Columns(1).Select
                         'oExcel.Selection.NumberFormat = "#,##0.00"
                         'oExcel.Columns(1).Select
                         'oExcel.Selection.Font.Color = vbRed
                         .Columns.AutoFit ' ajusta el ancho de las columnas
                     End With
                     archivo = "c:\reportessid\valuacion_devoluciones_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     owbook.SaveAs archivo
                     oexcel.Visible = True
                     Set oexcel = Nothing
                     Screen.MousePointer = vbDefault
                     rsaux10.Close
                     MsgBox "Se a terminado de guardar el archivo " + archivo
                  End If
               End If
            
            Else
               MsgBox "No existen pedidos para el periodo indicado", vbOKOnly, "ATENCION"
            End If
            rs.Close
            rs.Open "delete from TB_oracle_valuacion_Devoluciones WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "La fecha de inicio debe de ser menor o igual a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de Inicio incorrecta", vbOKOnly, "ATENCION"
   End If

End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 3000
   Left = 3500
   txt_inicio = Date
   txt_fin = Date
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fin) Then
         frmcalendario.mes = CDate(Me.txt_fin)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      txt_fin = var_fecha_general
   End If
End Sub

Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_inicio) Then
         frmcalendario.mes = CDate(Me.txt_inicio)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      txt_inicio = var_fecha_general
   End If
End Sub

