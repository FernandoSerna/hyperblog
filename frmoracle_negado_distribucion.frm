VERSION 5.00
Begin VB.Form frmoracle_negado_distribucion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de negado de distribución"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1330
      Picture         =   "frmoracle_negado_distribucion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Negado de distribución concentrado"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1020
      Picture         =   "frmoracle_negado_distribucion.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Negado de distribución VXT"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_concentrado 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   690
      Picture         =   "frmoracle_negado_distribucion.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Negado de distribución concentrado"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_resumen 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   360
      Picture         =   "frmoracle_negado_distribucion.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Negado de distribución resumen"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   30
      Picture         =   "frmoracle_negado_distribucion.frx":0408
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Negado de distribución"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6150
      Picture         =   "frmoracle_negado_distribucion.frx":050A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   840
      Left            =   75
      TabIndex        =   0
      Top             =   450
      Width           =   6405
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   4065
         TabIndex        =   2
         Top             =   300
         Width           =   2160
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   780
         TabIndex        =   1
         Top             =   315
         Width           =   2160
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   3720
         TabIndex        =   4
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   375
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   45
      TabIndex        =   7
      Top             =   285
      Width           =   6450
   End
End
Attribute VB_Name = "frmoracle_negado_distribucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter


Private Sub cmd_concentrado_Click()
   If IsDate(Me.txt_inicio) Then
      If IsDate(Me.txt_fin) Then
         If CDate(Me.txt_inicio) <= CDate(Me.txt_fin) Then
            cnn.BeginTrans
            rs.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_ORACLE_NEGADO_DISTRIBUCION", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
            Else
               var_consecutivo = 0
            End If
            rs.Close
            var_consecutivo = var_consecutivo + 1
            rs.Open "INSERT INTO TB_TEMP_ORACLE_NEGADO_DISTRIBUCION (INTE_TEM_CONSECUTIVO) VALUES ('" + CStr(var_consecutivo) + "')"
            cnn.CommitTrans
            var_dia_s = CStr(Day(CDate(txt_inicio)))
            var_mes_s = CStr(Month(CDate(Me.txt_inicio)))
            var_año_s = CStr(Year(CDate(Me.txt_inicio)))
            var_hora_s = CStr(Hour(CDate(Me.txt_inicio)))
            var_minuto_s = CStr(Minute(CDate(Me.txt_inicio)))
            var_segundo_s = CStr(Second(CDate(Me.txt_inicio)))
            If Len(var_dia_s) = 1 Then
               var_dia_s = "0" + var_dia_s
            End If
            If Len(var_mes_s) = 1 Then
               var_mes_s = "0" + var_mes_s
            End If
            If Len(var_año_s) = 2 Then
               var_año_s = "20" + var_dia_s
            End If
            If Len(var_hora_s) < 2 Then
               var_hora_s = "0" + var_hora_s
            End If
            If Len(var_minuto_s) < 2 Then
               var_minuto_s = "0" + var_minuto_s
            End If
            If Len(var_segundo_s) < 2 Then
               var_segundo_s = "0" + var_segundo_s
            End If
            
            var_fecha_inicio = var_dia_s + "-" + var_mes_s + "-" + var_año_s + " " + var_hora_s + ":" + var_minuto_s + ":" + var_segundo_s
            'var_fecha_inicio_sql = "{d '" + var_año_s + "-" + var_mes_s + "-" + var_dia_s + "'}"
            var_fecha_inicio_sql = " CONVERT(datetime, '" + var_año_s + "-" + var_mes_s + "-" + var_dia_s + "T" + var_hora_s + ":" + var_minuto_s + ":" + var_segundo_s + "',126)"
            var_dia_s = CStr(Day(CDate(Me.txt_fin)))
            var_mes_s = CStr(Month(CDate(Me.txt_fin)))
            var_año_s = CStr(Year(CDate(Me.txt_fin)))
            var_hora_s = CStr(Hour(CDate(Me.txt_fin)))
            var_minuto_s = CStr(Minute(CDate(Me.txt_fin)))
            var_segundo_s = CStr(Second(CDate(Me.txt_fin)))
            
            If Len(var_dia_s) = 1 Then
               var_dia_s = "0" + var_dia_s
            End If
            If Len(var_mes_s) = 1 Then
               var_mes_s = "0" + var_mes_s
            End If
            If Len(var_año_s) = 2 Then
               var_año_s = "20" + var_dia_s
            End If
            If Len(var_hora_s) < 2 Then
               var_hora_s = "0" + var_hora_s
            End If
            If Len(var_minuto_s) < 2 Then
               var_minuto_s = "0" + var_minuto_s
            End If
            If Len(var_segundo_s) < 2 Then
               var_segundo_s = "0" + var_segundo_s
            End If
            
            var_fecha_inicio_o = var_dia_s + "-" + var_mes_s + "-" + var_año_s + " " + var_hora_s + ":" + var_minuto_s + ":" + var_segundo_s
            
            'var_fecha_fin_sql = "{d '" + var_año_s + "-" + var_mes_s + "-" + var_dia_s + "'}"
            var_fecha_fin_sql = " CONVERT(datetime, '" + var_año_s + "-" + var_mes_s + "-" + var_dia_s + "T" + var_hora_s + ":" + var_minuto_s + ":" + var_segundo_s + "',126)"
            
            var_dia_s = CStr(Day(CDate(Me.txt_fin)))
            var_mes_s = CStr(Month(CDate(Me.txt_fin)))
            var_año_s = CStr(Year(CDate(Me.txt_fin)))
            var_hora_s = CStr(Hour(CDate(Me.txt_fin)))
            var_minuto_s = CStr(Minute(CDate(Me.txt_fin)))
            var_segundo_s = CStr(Second(CDate(Me.txt_fin)))
            If Len(var_dia_s) = 1 Then
               var_dia_s = "0" + var_dia_s
            End If
            If Len(var_mes_s) = 1 Then
               var_mes_s = "0" + var_mes_s
            End If
            If Len(var_año_s) = 2 Then
               var_año_s = "20" + var_dia_s
            End If
            If Len(var_hora_s) < 2 Then
               var_hora_s = "0" + var_hora_s
            End If
            If Len(var_minuto_s) < 2 Then
               var_minuto_s = "0" + var_minuto_s
            End If
            If Len(var_segundo_s) < 2 Then
               var_segundo_s = "0" + var_segundo_s
            End If
            
            var_fecha_fin = var_dia_s + "-" + var_mes_s + "-" + var_año_s + " " + var_hora_s + ":" + var_minuto_s + ":" + var_segundo_s
            
            rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            rs.Open "alter session set nls_date_format='DD-MM-YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            'rs.Open "select A.*, B.DESCRIPTION, C.LINEA from xxvia_tb_negado_distribucion A, xxvia_system_items_b b, xxvia_vw_categorias_item_b C where fecha_NEGADO > to_date('" + var_fecha_inicio + "','DD-MM-YYYY hh24:mi:ss') AND FECHA_NEGADO < TO_DATE('" + var_fecha_fin + "','DD-MM-YYYY hh24:mi:ss') and a.organization_id = b.organization_id and a.inventory_item_id = b.inventory_item_id and a.organization_id = C.organization_id and a.inventory_item_id = C.item_id and a.cantidad > 0 and a.organization_id = " + CStr(var_unidad_organizacional)
            rs.Open "select A.*, B.DESCRIPTION, b.attribute2, b.attribute3, b.attribute4, b.attribute5, b.attribute6, b.attribute7, C.LINEA from xxvia_tb_negado_distribucion A, xxvia_system_items_b b, xxvia_vw_categorias_item_b C where fecha_NEGADO > to_date('" + var_fecha_inicio + "','DD-MM-YYYY hh24:mi:ss') AND FECHA_NEGADO < TO_DATE('" + var_fecha_fin + "','DD-MM-YYYY hh24:mi:ss') and a.organization_id = b.organization_id and a.inventory_item_id = b.inventory_item_id and a.organization_id = C.organization_id and a.inventory_item_id = C.item_id and a.cantidad > 0 and a.organization_id = " + CStr(var_unidad_organizacional)
            If Not rs.EOF Then
               While Not rs.EOF
                     var_cadena = "insert into TB_TEMP_ORACLE_NEGADO_DISTRIBUCION (inte_tem_consecutivo, fecha_inicio, fecha_fin, pedido, tipo_pedido, ruta, cliente, codigo, cantidad_surtir, cantidad_surtida, negado_distribucion, causa_negado, descripcion, LINEA, ubicacion1, ubicacion2, ubicacion3, ubicacion4, ubicacion5, ubicacion6, existencia, disponible) "
                     var_causa_negado = IIf(IsNull(rs!nombre_causa_negado), "", rs!nombre_causa_negado)
                     var_ubicacion1 = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                     var_ubicacion2 = IIf(IsNull(rs!attribute3), "", rs!attribute3)
                     var_ubicacion3 = IIf(IsNull(rs!attribute4), "", rs!attribute4)
                     var_ubicacion4 = IIf(IsNull(rs!attribute5), "", rs!attribute5)
                     var_ubicacion5 = IIf(IsNull(rs!attribute6), "", rs!attribute6)
                     var_ubicacion6 = IIf(IsNull(rs!attribute7), "", rs!attribute7)
                     If rs!cantidad > 0 And var_causa_negado = "" Then
                        var_causa_negado = "NO LOCALIZADO"
                     End If
                     
                     strconsulta = "select * from Xxvia_vw_existencias_inv where organization_id = ? and subinventory_code = ? and segment1 = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(var_unidad_organizacional))
                          .Parameters.Append parametro
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, "CDI_ALMPT")
                          .Parameters.Append parametro
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rs!SEGMENT1)
                          .Parameters.Append parametro
                     End With
                     Set rsaux9 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     If Not rsaux9.EOF Then
                        var_existencias = IIf(IsNull(rsaux9!CANTMANO), 0, rsaux9!CANTMANO)
                        var_disponible = IIf(IsNull(rsaux9!Disponible), 0, rsaux9!Disponible)
                     Else
                        var_existencias = 0
                        var_disponible = 0
                     End If
                     var_cadena = var_cadena + " values (" + CStr(var_consecutivo) + "," + var_fecha_inicio_sql + "," + var_fecha_fin_sql + "," + CStr(rs!source_header_number) + ",'','','','" + rs!SEGMENT1 + "'," + CStr(rs!CANTIDAD_PEDIDA) + "," + CStr(rs!CANTIDAD_SURTIDA) + "," + CStr(rs!cantidad) + ",'" + var_causa_negado + "','" + rs!Description + "', '" + rs!Linea + "','" + var_ubicacion1 + "','" + var_ubicacion2 + "','" + var_ubicacion3 + "', '" + var_ubicacion4 + "','" + var_ubicacion5 + "','" + var_ubicacion6 + "'," + CStr(var_existencias) + "," + CStr(var_disponible) + ")"
                     rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     rsaux9.Close
                     rs.MoveNext
               Wend
               rs.Close
               'rs.Open "select distinct codigo from TB_TEMP_ORACLE_NEGADO_DISTRIBUCION where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and codigo is not null", cnn, adOpenDynamic, adLockOptimistic
               'While Not rs.EOF
               '      strconsulta = "select * from xxvia_system_items_b where segment1 = ? and organization_id = ?"
               '      With comandoORA
               '           .ActiveConnection = cnnoracle_4
               '           .CommandType = adCmdText
               '           .CommandText = strconsulta
               '           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rs!codigo)
               '           .Parameters.Append parametro
               '           Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
               '           .Parameters.Append parametro
               '      End With
               '      Set rsaux9 = comandoORA.execute
               '      Set comandoORA = Nothing
               '      Set parametro = Nothing
               '      rsaux1.Open "update TB_TEMP_ORACLE_NEGADO_DISTRIBUCION set descripcion = '" + rsaux9!Description + "' where codigo = '" + rs!codigo + "' and inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
               '      rsaux9.Close
               '      rs.MoveNext
               'Wend
               'rs.Close
               x = 1
               If x = 0 Then
                  rsaux10.Open "select distinct pedido from TB_TEMP_ORACLE_NEGADO_DISTRIBUCION where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " and pedido is not null", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux10.EOF
                        var_orden = rsaux10!pedido
                        strconsulta = "SELECT HEADER_ID, source_document_id, SHIP_TO_ORG_ID, A.NAME AS RUTA FROM OE_ORDER_HEADERS_ALL OHA, XXVIA_VENDEDORES A WHERE ORDER_NUMBER  = ? AND OHA.SALESREP_ID = A.SALESREP_ID"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_orden)
                             .Parameters.Append parametro
                        End With
                        Set rsaux7 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                  
                        If Not rsaux7.EOF Then
                           VAR_HEADER_ID = IIf(IsNull(rsaux7!header_id), 0, rsaux7!header_id)
                           var_requisicion = IIf(IsNull(rsaux7!source_document_id), "", rsaux7!source_document_id)
                           var_establecimiento = IIf(IsNull(rsaux7!ship_to_org_id), "0", rsaux7!ship_to_org_id)
                           var_nombre_agente_str = rsaux7!ruta
                        Else
                           VAR_HEADER_ID = 0
                        End If
                        rsaux7.Close
                  
                        var_cadena = " SELECT a.source_header_type_name, HCAS.CUST_ACCOUNT_ID, HCAS.CUST_ACCT_SITE_ID as customer_id, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, oha.attribute8, oha.attribute9 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND to_number(source_header_number) = " + CStr(var_orden)
                        var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID   AND A.SOURCE_HEADER_ID = " + CStr(VAR_HEADER_ID)
                        If rs.State = 1 Then
                           rs.Close
                        End If
                        rs.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                 
                        If rs!source_header_type_name = "VIA_PEDIDO_INTERNO" Or rs!source_header_type_name = "TEX_PEDIDO_INTERNO" Then
                           If var_pedido_tienda = 0 Then
                              txt_cliente = IIf(IsNull(rs!customer_name), "", rs!customer_name)
                              rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(var_requisicion) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux2.EOF Then
                                 txt_entrega = Replace(IIf(IsNull(rsaux2!Description), "", rsaux2!Description), "'", " ")
                              End If
                              rsaux2.Close
                           Else
                              txt_cliente = IIf(IsNull(rs!ATTRIBUTE9), "", rs!ATTRIBUTE9)
                           End If
                        Else
                           txt_cliente = IIf(IsNull(rs!customer_name), "", rs!customer_name)
                        End If
                        If txt_cliente = "VIANNEY TEXTIL HOGAR SA DE CV" Then
                           var_nombre_agente_str = "VIANNEY TEXTIL HOGAR SA DE CV"
                           txt_cliente = txt_entrega
                        End If
                        rsaux11.Open "update TB_TEMP_ORACLE_NEGADO_DISTRIBUCION set tipo_pedido = '" + rs!source_header_type_name + "', RUTA = '" + var_nombre_agente_str + "', CLIENTE = '" + txt_cliente + "' where pedido = " + CStr(var_orden), cnn, adOpenDynamic, adLockOptimistic
                        rs.Close
                        rsaux10.MoveNext
                  Wend
                  rsaux10.Close
               End If
               rsaux10.Open "DELETE FROM TB_TEMP_ORACLE_NEGADO_DISTRIBUCION WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO IS NULL", cnn, adOpenDynamic, adLockOptimistic
               x = 0
               If x = 1 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_negado_distribucion_concentrado.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_NEGADO_DISTRIBUCION_CONCENTRADO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\negado_distribucion_concentrado_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + ".xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               Else
                  Set oexcel = CreateObject("Excel.Application")
                  Set owbook = oexcel.Workbooks.Add
                  Set osheet = owbook.Worksheets(1)
                  osheet.Name = "NEGADO DISTRIBUCION"
                  Screen.MousePointer = vbHourglass
                  iFila = 1
                  ifila2 = 1
                  icol2 = 1
                  iCol = 1
                  var_cadena = "SELECT fecha_inicio, FECHA_FIN, CODIGO, DESCRIPCION, LINEA,  expr1 as CANTIDAD_PEDIDA, Expr2 AS CANTIDAD_SURTIDA, EXPR3 CANTIDAD_NEGADA, CAUSA_NEGADO, UBICACION1, UBICACION2, UBICACION3, UBICACION4, UBICACION5, UBICACION6, EXISTENCIA, DISPONIBLE   FROM VW_ORACLE_NEGADO_DISTRIBUCION_CONCENTRADO WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo)
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
                  archivo = "c:\reportessid\negado_distribucion_concentrado_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + ".xls"
                  owbook.SaveAs archivo
                  oexcel.Visible = True
                  Set oexcel = Nothing
                  Screen.MousePointer = vbDefault
                  rsaux10.Close
               
               
               End If
               MsgBox "Se a terminado de guardar el archivo " + archivo
            Else
               MsgBox "No existe negado para el periodo seleccionado", vbOKOnly, "ATENCION"
            End If
            If rs.State = 1 Then
               rs.Close
            End If
         Else
            MsgBox "La fecha inicial debe de ser menor o igual a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha inicial incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
   If IsDate(Me.txt_inicio) Then
      If IsDate(Me.txt_fin) Then
         If CDate(Me.txt_inicio) <= CDate(Me.txt_fin) Then
            cnn.BeginTrans
            rs.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_ORACLE_NEGADO_DISTRIBUCION", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
            Else
               var_consecutivo = 0
            End If
            rs.Close
            var_consecutivo = var_consecutivo + 1
            rs.Open "INSERT INTO TB_TEMP_ORACLE_NEGADO_DISTRIBUCION (INTE_TEM_CONSECUTIVO) VALUES ('" + CStr(var_consecutivo) + "')"
            cnn.CommitTrans
            var_dia_s = CStr(Day(CDate(txt_inicio)))
            var_mes_s = CStr(Month(CDate(Me.txt_inicio)))
            var_año_s = CStr(Year(CDate(Me.txt_inicio)))
            var_hora_s = CStr(Hour(CDate(Me.txt_inicio)))
            var_minuto_s = CStr(Minute(CDate(Me.txt_inicio)))
            var_segundo_s = CStr(Second(CDate(Me.txt_inicio)))
            If Len(var_dia_s) = 1 Then
               var_dia_s = "0" + var_dia_s
            End If
            If Len(var_mes_s) = 1 Then
               var_mes_s = "0" + var_mes_s
            End If
            If Len(var_año_s) = 2 Then
               var_año_s = "20" + var_dia_s
            End If
            If Len(var_hora_s) < 2 Then
               var_hora_s = "0" + var_hora_s
            End If
            If Len(var_minuto_s) < 2 Then
               var_minuto_s = "0" + var_minuto_s
            End If
            If Len(var_segundo_s) < 2 Then
               var_segundo_s = "0" + var_segundo_s
            End If
            
            var_fecha_inicio = var_dia_s + "-" + var_mes_s + "-" + var_año_s + " " + var_hora_s + ":" + var_minuto_s + ":" + var_segundo_s
            'var_fecha_inicio_sql = "{d '" + var_año_s + "-" + var_mes_s + "-" + var_dia_s + "'}"
            var_fecha_inicio_sql = " CONVERT(datetime, '" + var_año_s + "-" + var_mes_s + "-" + var_dia_s + "T" + var_hora_s + ":" + var_minuto_s + ":" + var_segundo_s + "',126)"
            var_dia_s = CStr(Day(CDate(Me.txt_fin)))
            var_mes_s = CStr(Month(CDate(Me.txt_fin)))
            var_año_s = CStr(Year(CDate(Me.txt_fin)))
            var_hora_s = CStr(Hour(CDate(Me.txt_fin)))
            var_minuto_s = CStr(Minute(CDate(Me.txt_fin)))
            var_segundo_s = CStr(Second(CDate(Me.txt_fin)))
            
            If Len(var_dia_s) = 1 Then
               var_dia_s = "0" + var_dia_s
            End If
            If Len(var_mes_s) = 1 Then
               var_mes_s = "0" + var_mes_s
            End If
            If Len(var_año_s) = 2 Then
               var_año_s = "20" + var_dia_s
            End If
            If Len(var_hora_s) < 2 Then
               var_hora_s = "0" + var_hora_s
            End If
            If Len(var_minuto_s) < 2 Then
               var_minuto_s = "0" + var_minuto_s
            End If
            If Len(var_segundo_s) < 2 Then
               var_segundo_s = "0" + var_segundo_s
            End If
            
            var_fecha_inicio_o = var_dia_s + "-" + var_mes_s + "-" + var_año_s + " " + var_hora_s + ":" + var_minuto_s + ":" + var_segundo_s
            
            'var_fecha_fin_sql = "{d '" + var_año_s + "-" + var_mes_s + "-" + var_dia_s + "'}"
            var_fecha_fin_sql = " CONVERT(datetime, '" + var_año_s + "-" + var_mes_s + "-" + var_dia_s + "T" + var_hora_s + ":" + var_minuto_s + ":" + var_segundo_s + "',126)"
            
            var_dia_s = CStr(Day(CDate(Me.txt_fin)))
            var_mes_s = CStr(Month(CDate(Me.txt_fin)))
            var_año_s = CStr(Year(CDate(Me.txt_fin)))
            var_hora_s = CStr(Hour(CDate(Me.txt_fin)))
            var_minuto_s = CStr(Minute(CDate(Me.txt_fin)))
            var_segundo_s = CStr(Second(CDate(Me.txt_fin)))
            If Len(var_dia_s) = 1 Then
               var_dia_s = "0" + var_dia_s
            End If
            If Len(var_mes_s) = 1 Then
               var_mes_s = "0" + var_mes_s
            End If
            If Len(var_año_s) = 2 Then
               var_año_s = "20" + var_dia_s
            End If
            If Len(var_hora_s) < 2 Then
               var_hora_s = "0" + var_hora_s
            End If
            If Len(var_minuto_s) < 2 Then
               var_minuto_s = "0" + var_minuto_s
            End If
            If Len(var_segundo_s) < 2 Then
               var_segundo_s = "0" + var_segundo_s
            End If
            
            var_fecha_fin = var_dia_s + "-" + var_mes_s + "-" + var_año_s + " " + var_hora_s + ":" + var_minuto_s + ":" + var_segundo_s
            
            rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            rs.Open "alter session set nls_date_format='DD-MM-YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            rs.Open "select A.*, B.DESCRIPTION, C.LINEA from xxvia_tb_negado_distribucion A, xxvia_system_items_b b, xxvia_vw_categorias_item_b C where fecha_NEGADO > to_date('" + var_fecha_inicio + "','DD-MM-YYYY hh24:mi:ss') AND FECHA_NEGADO < TO_DATE('" + var_fecha_fin + "','DD-MM-YYYY hh24:mi:ss') and a.organization_id = b.organization_id and a.inventory_item_id = b.inventory_item_id and a.organization_id = C.organization_id and a.inventory_item_id = C.item_id and a.organization_id = " + CStr(var_unidad_organizacional)
            If Not rs.EOF Then
               While Not rs.EOF
                     var_cadena = "insert into TB_TEMP_ORACLE_NEGADO_DISTRIBUCION (inte_tem_consecutivo, fecha_inicio, fecha_fin, pedido, tipo_pedido, ruta, cliente, codigo, cantidad_surtir, cantidad_surtida, negado_distribucion, causa_negado, descripcion, LINEA) "
                     var_causa_negado = IIf(IsNull(rs!nombre_causa_negado), "", rs!nombre_causa_negado)
                     If rs!cantidad > 0 And var_causa_negado = "" Then
                        var_causa_negado = "NO LOCALIZADO"
                     End If
                     var_cadena = var_cadena + " values (" + CStr(var_consecutivo) + "," + var_fecha_inicio_sql + "," + var_fecha_fin_sql + "," + CStr(rs!source_header_number) + ",'','','','" + rs!SEGMENT1 + "'," + CStr(rs!CANTIDAD_PEDIDA) + "," + CStr(rs!CANTIDAD_SURTIDA) + "," + CStr(rs!cantidad) + ",'" + var_causa_negado + "','" + rs!Description + "', '" + rs!Linea + "')"
                     rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               rs.Close
               'rs.Open "select distinct codigo from TB_TEMP_ORACLE_NEGADO_DISTRIBUCION where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and codigo is not null", cnn, adOpenDynamic, adLockOptimistic
               'While Not rs.EOF
               '      strconsulta = "select * from xxvia_system_items_b where segment1 = ? and organization_id = ?"
               '      With comandoORA
               '           .ActiveConnection = cnnoracle_4
               '           .CommandType = adCmdText
               '           .CommandText = strconsulta
               '           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rs!codigo)
               '           .Parameters.Append parametro
               '           Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
               '           .Parameters.Append parametro
               '      End With
               '      Set rsaux9 = comandoORA.execute
               '      Set comandoORA = Nothing
               '      Set parametro = Nothing
               '      rsaux1.Open "update TB_TEMP_ORACLE_NEGADO_DISTRIBUCION set descripcion = '" + rsaux9!Description + "' where codigo = '" + rs!codigo + "' and inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
               '      rsaux9.Close
               '      rs.MoveNext
               'Wend
               'rs.Close
               rsaux10.Open "select distinct pedido from TB_TEMP_ORACLE_NEGADO_DISTRIBUCION where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " and pedido is not null", cnn, adOpenDynamic, adLockOptimistic
               While Not rsaux10.EOF
                     var_orden = rsaux10!pedido
                     strconsulta = "SELECT HEADER_ID, source_document_id, SHIP_TO_ORG_ID, A.NAME AS RUTA FROM OE_ORDER_HEADERS_ALL OHA, XXVIA_VENDEDORES A WHERE ORDER_NUMBER  = ? AND OHA.SALESREP_ID = A.SALESREP_ID and ship_From_org_id = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_orden)
                          .Parameters.Append parametro
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_unidad_organizacional)
                          .Parameters.Append parametro
                     End With
                     Set rsaux7 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
               
                     If Not rsaux7.EOF Then
                        VAR_HEADER_ID = IIf(IsNull(rsaux7!header_id), 0, rsaux7!header_id)
                        var_requisicion = IIf(IsNull(rsaux7!source_document_id), "", rsaux7!source_document_id)
                        var_establecimiento = IIf(IsNull(rsaux7!ship_to_org_id), "0", rsaux7!ship_to_org_id)
                        var_nombre_agente_str = rsaux7!ruta
                     Else
                        VAR_HEADER_ID = 0
                     End If
                     rsaux7.Close
               
                     var_cadena = " SELECT a.source_header_type_name, HCAS.CUST_ACCOUNT_ID, HCAS.CUST_ACCT_SITE_ID as customer_id, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, oha.attribute8, oha.attribute9 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND to_number(source_header_number) = " + CStr(var_orden)
                     var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID   AND A.SOURCE_HEADER_ID = " + CStr(VAR_HEADER_ID)
                     If rs.State = 1 Then
                        rs.Close
                     End If
                     rs.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                     If rs!source_header_type_name = "VIA_PEDIDO_INTERNO" Or rs!source_header_type_name = "TEX_PEDIDO_INTERNO" Then
                        If var_pedido_tienda = 0 Then
                           txt_cliente = IIf(IsNull(rs!customer_name), "", rs!customer_name)
                           rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(var_requisicion) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If Not rsaux2.EOF Then
                              txt_entrega = Replace(IIf(IsNull(rsaux2!Description), "", rsaux2!Description), "'", " ")
                           End If
                           rsaux2.Close
                        Else
                           txt_cliente = IIf(IsNull(rs!ATTRIBUTE9), "", rs!ATTRIBUTE9)
                        End If
                     Else
                        txt_cliente = IIf(IsNull(rs!customer_name), "", rs!customer_name)
                     End If
                     If txt_cliente = "VIANNEY TEXTIL HOGAR SA DE CV" Then
                        var_nombre_agente_str = "VIANNEY TEXTIL HOGAR SA DE CV"
                        txt_cliente = txt_entrega
                     End If
                     rsaux11.Open "update TB_TEMP_ORACLE_NEGADO_DISTRIBUCION set tipo_pedido = '" + rs!source_header_type_name + "', RUTA = '" + var_nombre_agente_str + "', CLIENTE = '" + txt_cliente + "' where pedido = " + CStr(var_orden), cnn, adOpenDynamic, adLockOptimistic
                     End If
                     rs.Close
                     rsaux10.MoveNext
               Wend
               rsaux10.Close
               x = 0
               If x = 1 Then
                  rsaux10.Open "DELETE FROM TB_TEMP_ORACLE_NEGADO_DISTRIBUCION WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO IS NULL", cnn, adOpenDynamic, adLockOptimistic
                  Set reporte = appl.OpenReport(App.Path + "\rep_oracle_negado_distribucion.rpt")
                  reporte.RecordSelectionFormula = "{VW_ORACLE_NEGADO_DISTRIBUCION.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  reporte.ExportOptions.FormatType = crEFTExcel80
                  reporte.ExportOptions.DestinationType = crEDTDiskFile
                  archivo = "c:\reportessid\negado_distribucion_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + ".xls"
                  reporte.ExportOptions.DiskFileName = archivo
                  reporte.Export False
                  Set reporte = Nothing
               Else
                  Set oexcel = CreateObject("Excel.Application")
                  Set owbook = oexcel.Workbooks.Add
                  Set osheet = owbook.Worksheets(1)
                  osheet.Name = "NEGADO DISTRIBUCION"
                  Screen.MousePointer = vbHourglass
                  iFila = 1
                  ifila2 = 1
                  icol2 = 1
                  iCol = 1
                  var_cadena = "select FECHA_INICIO, FECHA_FIN, PEDIDO, TIPO_PEDIDO, RUTA, CLIENTE, CODIGO, DESCRIPCION, LINEA, CANTIDAD_SURTIR, CANTIDAD_SURTIDA, NEGADO_DISTRIBUCION, CAUSA_NEGADO from VW_ORACLE_NEGADO_DISTRIBUCION WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + "  AND CODIGO IS NOT NULL"
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
                  archivo = "c:\reportessid\negado_distribucion_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + ".xls"
                  owbook.SaveAs archivo
                  oexcel.Visible = True
                  Set oexcel = Nothing
                  Screen.MousePointer = vbDefault
                  rsaux10.Close
               End If
               MsgBox "Se a terminado de guardar el archivo " + archivo
            Else
               MsgBox "No existe negado para el periodo seleccionado", vbOKOnly, "ATENCION"
            End If
            If rs.State = 1 Then
               rs.Close
            End If
         Else
            MsgBox "La fecha inicial debe de ser menor o igual a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha inicial incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_resumen_Click()
   If IsDate(Me.txt_inicio) Then
      If IsDate(Me.txt_fin) Then
         If CDate(Me.txt_inicio) <= CDate(Me.txt_fin) Then
            cnn.BeginTrans
            rs.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_ORACLE_NEGADO_DISTRIBUCION", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
            Else
               var_consecutivo = 0
            End If
            rs.Close
            var_consecutivo = var_consecutivo + 1
            rs.Open "INSERT INTO TB_TEMP_ORACLE_NEGADO_DISTRIBUCION (INTE_TEM_CONSECUTIVO) VALUES ('" + CStr(var_consecutivo) + "')"
            cnn.CommitTrans
            var_dia_s = CStr(Day(CDate(txt_inicio)))
            var_mes_s = CStr(Month(CDate(Me.txt_inicio)))
            var_año_s = CStr(Year(CDate(Me.txt_inicio)))
            var_hora_s = CStr(Hour(CDate(Me.txt_inicio)))
            var_minuto_s = CStr(Minute(CDate(Me.txt_inicio)))
            var_segundo_s = CStr(Second(CDate(Me.txt_inicio)))
            If Len(var_dia_s) = 1 Then
               var_dia_s = "0" + var_dia_s
            End If
            If Len(var_mes_s) = 1 Then
               var_mes_s = "0" + var_mes_s
            End If
            If Len(var_año_s) = 2 Then
               var_año_s = "20" + var_dia_s
            End If
            If Len(var_hora_s) < 2 Then
               var_hora_s = "0" + var_hora_s
            End If
            If Len(var_minuto_s) < 2 Then
               var_minuto_s = "0" + var_minuto_s
            End If
            If Len(var_segundo_s) < 2 Then
               var_segundo_s = "0" + var_segundo_s
            End If
            
            var_fecha_inicio = var_dia_s + "-" + var_mes_s + "-" + var_año_s + " " + var_hora_s + ":" + var_minuto_s + ":" + var_segundo_s
            'var_fecha_inicio_sql = "{d '" + var_año_s + "-" + var_mes_s + "-" + var_dia_s + "'}"
            var_fecha_inicio_sql = " CONVERT(datetime, '" + var_año_s + "-" + var_mes_s + "-" + var_dia_s + "T" + var_hora_s + ":" + var_minuto_s + ":" + var_segundo_s + "',126)"
            var_dia_s = CStr(Day(CDate(Me.txt_fin)))
            var_mes_s = CStr(Month(CDate(Me.txt_fin)))
            var_año_s = CStr(Year(CDate(Me.txt_fin)))
            var_hora_s = CStr(Hour(CDate(Me.txt_fin)))
            var_minuto_s = CStr(Minute(CDate(Me.txt_fin)))
            var_segundo_s = CStr(Second(CDate(Me.txt_fin)))
            
            If Len(var_dia_s) = 1 Then
               var_dia_s = "0" + var_dia_s
            End If
            If Len(var_mes_s) = 1 Then
               var_mes_s = "0" + var_mes_s
            End If
            If Len(var_año_s) = 2 Then
               var_año_s = "20" + var_dia_s
            End If
            If Len(var_hora_s) < 2 Then
               var_hora_s = "0" + var_hora_s
            End If
            If Len(var_minuto_s) < 2 Then
               var_minuto_s = "0" + var_minuto_s
            End If
            If Len(var_segundo_s) < 2 Then
               var_segundo_s = "0" + var_segundo_s
            End If
            
            var_fecha_inicio_o = var_dia_s + "-" + var_mes_s + "-" + var_año_s + " " + var_hora_s + ":" + var_minuto_s + ":" + var_segundo_s
            
            'var_fecha_fin_sql = "{d '" + var_año_s + "-" + var_mes_s + "-" + var_dia_s + "'}"
            var_fecha_fin_sql = " CONVERT(datetime, '" + var_año_s + "-" + var_mes_s + "-" + var_dia_s + "T" + var_hora_s + ":" + var_minuto_s + ":" + var_segundo_s + "',126)"
            
            var_dia_s = CStr(Day(CDate(Me.txt_fin)))
            var_mes_s = CStr(Month(CDate(Me.txt_fin)))
            var_año_s = CStr(Year(CDate(Me.txt_fin)))
            var_hora_s = CStr(Hour(CDate(Me.txt_fin)))
            var_minuto_s = CStr(Minute(CDate(Me.txt_fin)))
            var_segundo_s = CStr(Second(CDate(Me.txt_fin)))
            If Len(var_dia_s) = 1 Then
               var_dia_s = "0" + var_dia_s
            End If
            If Len(var_mes_s) = 1 Then
               var_mes_s = "0" + var_mes_s
            End If
            If Len(var_año_s) = 2 Then
               var_año_s = "20" + var_dia_s
            End If
            If Len(var_hora_s) < 2 Then
               var_hora_s = "0" + var_hora_s
            End If
            If Len(var_minuto_s) < 2 Then
               var_minuto_s = "0" + var_minuto_s
            End If
            If Len(var_segundo_s) < 2 Then
               var_segundo_s = "0" + var_segundo_s
            End If
            
            var_fecha_fin = var_dia_s + "-" + var_mes_s + "-" + var_año_s + " " + var_hora_s + ":" + var_minuto_s + ":" + var_segundo_s
            
            rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            rs.Open "alter session set nls_date_format='DD-MM-YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            'rs.Open "select A.*, B.DESCRIPTION, C.LINEA from xxvia_tb_negado_distribucion A, xxvia_system_items_b b, xxvia_vw_categorias_item_b C where fecha_NEGADO > to_date('" + var_fecha_inicio + "','DD-MM-YYYY hh24:mi:ss') AND FECHA_NEGADO < TO_DATE('" + var_fecha_fin + "','DD-MM-YYYY hh24:mi:ss') and a.organization_id = b.organization_id and a.inventory_item_id = b.inventory_item_id and a.organization_id = C.organization_id and a.inventory_item_id = C.item_id and a.organization_id = " + CStr(var_unidad_organizacional)
            rs.Open "select A.*, B.DESCRIPTION, C.LINEA from xxvia_tb_negado_distribucion A, xxvia_system_items_b b, xxvia_vw_categorias_item_b C where fecha_NEGADO > to_date('" + var_fecha_inicio + "','DD-MM-YYYY hh24:mi:ss') AND FECHA_NEGADO < TO_DATE('" + var_fecha_fin + "','DD-MM-YYYY hh24:mi:ss') and a.organization_id = b.organization_id and a.inventory_item_id = b.inventory_item_id and a.organization_id = C.organization_id and a.inventory_item_id = C.item_id and a.cantidad > 0 and a.organization_id = " + CStr(var_unidad_organizacional)
            If Not rs.EOF Then
               While Not rs.EOF
                     var_cadena = "insert into TB_TEMP_ORACLE_NEGADO_DISTRIBUCION (inte_tem_consecutivo, fecha_inicio, fecha_fin, pedido, tipo_pedido, ruta, cliente, codigo, cantidad_surtir, cantidad_surtida, negado_distribucion, causa_negado, descripcion, LINEA) "
                     var_causa_negado = IIf(IsNull(rs!nombre_causa_negado), "", rs!nombre_causa_negado)
                     If rs!cantidad > 0 And var_causa_negado = "" Then
                        var_causa_negado = "NO LOCALIZADO"
                     End If
                     var_cadena = var_cadena + " values (" + CStr(var_consecutivo) + "," + var_fecha_inicio_sql + "," + var_fecha_fin_sql + "," + CStr(rs!source_header_number) + ",'','','','" + rs!SEGMENT1 + "'," + CStr(rs!CANTIDAD_PEDIDA) + "," + CStr(rs!CANTIDAD_SURTIDA) + "," + CStr(rs!cantidad) + ",'" + var_causa_negado + "','" + rs!Description + "', '" + rs!Linea + "')"
                     rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               rs.Close
               'rs.Open "select distinct codigo from TB_TEMP_ORACLE_NEGADO_DISTRIBUCION where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and codigo is not null", cnn, adOpenDynamic, adLockOptimistic
               'While Not rs.EOF
               '      strconsulta = "select * from xxvia_system_items_b where segment1 = ? and organization_id = ?"
               '      With comandoORA
               '           .ActiveConnection = cnnoracle_4
               '           .CommandType = adCmdText
               '           .CommandText = strconsulta
               '           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rs!codigo)
               '           .Parameters.Append parametro
               '           Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
               '           .Parameters.Append parametro
               '      End With
               '      Set rsaux9 = comandoORA.execute
               '      Set comandoORA = Nothing
               '      Set parametro = Nothing
               '      rsaux1.Open "update TB_TEMP_ORACLE_NEGADO_DISTRIBUCION set descripcion = '" + rsaux9!Description + "' where codigo = '" + rs!codigo + "' and inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
               '      rsaux9.Close
               '      rs.MoveNext
               'Wend
               'rs.Close
               rsaux10.Open "select distinct pedido from TB_TEMP_ORACLE_NEGADO_DISTRIBUCION where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " and pedido is not null", cnn, adOpenDynamic, adLockOptimistic
               While Not rsaux10.EOF
                     var_orden = rsaux10!pedido
                     strconsulta = "SELECT HEADER_ID, source_document_id, SHIP_TO_ORG_ID, A.NAME AS RUTA FROM OE_ORDER_HEADERS_ALL OHA, XXVIA_VENDEDORES A WHERE ORDER_NUMBER  = ? AND OHA.SALESREP_ID = A.SALESREP_ID"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_orden)
                          .Parameters.Append parametro
                     End With
                     Set rsaux7 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
               
                     If Not rsaux7.EOF Then
                        VAR_HEADER_ID = IIf(IsNull(rsaux7!header_id), 0, rsaux7!header_id)
                        var_requisicion = IIf(IsNull(rsaux7!source_document_id), "", rsaux7!source_document_id)
                        var_establecimiento = IIf(IsNull(rsaux7!ship_to_org_id), "0", rsaux7!ship_to_org_id)
                        var_nombre_agente_str = rsaux7!ruta
                     Else
                        VAR_HEADER_ID = 0
                     End If
                     rsaux7.Close
               
                     var_cadena = " SELECT a.source_header_type_name, HCAS.CUST_ACCOUNT_ID, HCAS.CUST_ACCT_SITE_ID as customer_id, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, oha.attribute8, oha.attribute9 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND to_number(source_header_number) = " + CStr(var_orden)
                     var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID   AND A.SOURCE_HEADER_ID = " + CStr(VAR_HEADER_ID)
                     If rs.State = 1 Then
                        rs.Close
                     End If
                     rs.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
              
                     If rs!source_header_type_name = "VIA_PEDIDO_INTERNO" Or rs!source_header_type_name = "TEX_PEDIDO_INTERNO" Then
                        If var_pedido_tienda = 0 Then
                           txt_cliente = IIf(IsNull(rs!customer_name), "", rs!customer_name)
                           rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(var_requisicion) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If Not rsaux2.EOF Then
                              txt_entrega = Replace(IIf(IsNull(rsaux2!Description), "", rsaux2!Description), "'", " ")
                           End If
                           rsaux2.Close
                        Else
                           txt_cliente = IIf(IsNull(rs!ATTRIBUTE9), "", rs!ATTRIBUTE9)
                        End If
                     Else
                        txt_cliente = IIf(IsNull(rs!customer_name), "", rs!customer_name)
                     End If
                     If txt_cliente = "VIANNEY TEXTIL HOGAR SA DE CV" Then
                        var_nombre_agente_str = "VIANNEY TEXTIL HOGAR SA DE CV"
                        txt_cliente = txt_entrega
                     End If
                     rsaux11.Open "update TB_TEMP_ORACLE_NEGADO_DISTRIBUCION set tipo_pedido = '" + rs!source_header_type_name + "', RUTA = '" + var_nombre_agente_str + "', CLIENTE = '" + txt_cliente + "' where pedido = " + CStr(var_orden), cnn, adOpenDynamic, adLockOptimistic
                     rs.Close
                     rsaux10.MoveNext
               Wend
               rsaux10.Close
               rsaux10.Open "DELETE FROM TB_TEMP_ORACLE_NEGADO_DISTRIBUCION WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO IS NULL", cnn, adOpenDynamic, adLockOptimistic
               x = 0
               If x = 1 Then
                  Set reporte = appl.OpenReport(App.Path + "\rep_oracle_negado_distribucion.rpt")
                  reporte.RecordSelectionFormula = "{VW_ORACLE_NEGADO_DISTRIBUCION.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and  {VW_ORACLE_NEGADO_DISTRIBUCION.NEGADO_DISTRIBUCION} > 0"
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  reporte.ExportOptions.FormatType = crEFTExcel80
                  reporte.ExportOptions.DestinationType = crEDTDiskFile
                  archivo = "c:\reportessid\negado_distribucion_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + ".xls"
                  reporte.ExportOptions.DiskFileName = archivo
                  reporte.Export False
                  Set reporte = Nothing
               Else
                  Set oexcel = CreateObject("Excel.Application")
                  Set owbook = oexcel.Workbooks.Add
                  Set osheet = owbook.Worksheets(1)
                  osheet.Name = "NEGADO DISTRIBUCION"
                  Screen.MousePointer = vbHourglass
                  iFila = 1
                  ifila2 = 1
                  icol2 = 1
                  iCol = 1
                  var_cadena = "select FECHA_INICIO, FECHA_FIN, PEDIDO, TIPO_PEDIDO, RUTA, CLIENTE, CODIGO, DESCRIPCION, LINEA, CANTIDAD_SURTIR, CANTIDAD_SURTIDA, NEGADO_DISTRIBUCION, CAUSA_NEGADO from VW_ORACLE_NEGADO_DISTRIBUCION WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + "  AND CODIGO IS NOT NULL AND NEGADO_DISTRIBUCION > 0"
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
                  archivo = "c:\reportessid\negado_distribucion_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + ".xls"
                  owbook.SaveAs archivo
                  oexcel.Visible = True
                  Set oexcel = Nothing
                  Screen.MousePointer = vbDefault
                  rsaux10.Close
                  'reporte.ExportOptions.DiskFileName = archivo
               End If
               MsgBox "Se a terminado de guardar el archivo " + archivo
            Else
               MsgBox "No existe negado para el periodo seleccionado", vbOKOnly, "ATENCION"
            End If
            If rs.State = 1 Then
               rs.Close
            End If
         Else
            MsgBox "La fecha inicial debe de ser menor o igual a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha inicial incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   If IsDate(Me.txt_inicio) Then
      If IsDate(Me.txt_fin) Then
         If CDate(Me.txt_inicio) <= CDate(Me.txt_fin) Then
            cnn.BeginTrans
            rs.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_ORACLE_NEGADO_DISTRIBUCION", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
            Else
               var_consecutivo = 0
            End If
            rs.Close
            var_consecutivo = var_consecutivo + 1
            rs.Open "INSERT INTO TB_TEMP_ORACLE_NEGADO_DISTRIBUCION (INTE_TEM_CONSECUTIVO) VALUES ('" + CStr(var_consecutivo) + "')"
            cnn.CommitTrans
            var_dia_s = CStr(Day(CDate(txt_inicio)))
            var_mes_s = CStr(Month(CDate(Me.txt_inicio)))
            var_año_s = CStr(Year(CDate(Me.txt_inicio)))
            var_hora_s = CStr(Hour(CDate(Me.txt_inicio)))
            var_minuto_s = CStr(Minute(CDate(Me.txt_inicio)))
            var_segundo_s = CStr(Second(CDate(Me.txt_inicio)))
            If Len(var_dia_s) = 1 Then
               var_dia_s = "0" + var_dia_s
            End If
            If Len(var_mes_s) = 1 Then
               var_mes_s = "0" + var_mes_s
            End If
            If Len(var_año_s) = 2 Then
               var_año_s = "20" + var_dia_s
            End If
            If Len(var_hora_s) < 2 Then
               var_hora_s = "0" + var_hora_s
            End If
            If Len(var_minuto_s) < 2 Then
               var_minuto_s = "0" + var_minuto_s
            End If
            If Len(var_segundo_s) < 2 Then
               var_segundo_s = "0" + var_segundo_s
            End If
            
            var_fecha_inicio = var_dia_s + "-" + var_mes_s + "-" + var_año_s + " " + var_hora_s + ":" + var_minuto_s + ":" + var_segundo_s
            'var_fecha_inicio_sql = "{d '" + var_año_s + "-" + var_mes_s + "-" + var_dia_s + "'}"
            var_fecha_inicio_sql = " CONVERT(datetime, '" + var_año_s + "-" + var_mes_s + "-" + var_dia_s + "T" + var_hora_s + ":" + var_minuto_s + ":" + var_segundo_s + "',126)"
            var_dia_s = CStr(Day(CDate(Me.txt_fin)))
            var_mes_s = CStr(Month(CDate(Me.txt_fin)))
            var_año_s = CStr(Year(CDate(Me.txt_fin)))
            var_hora_s = CStr(Hour(CDate(Me.txt_fin)))
            var_minuto_s = CStr(Minute(CDate(Me.txt_fin)))
            var_segundo_s = CStr(Second(CDate(Me.txt_fin)))
            
            If Len(var_dia_s) = 1 Then
               var_dia_s = "0" + var_dia_s
            End If
            If Len(var_mes_s) = 1 Then
               var_mes_s = "0" + var_mes_s
            End If
            If Len(var_año_s) = 2 Then
               var_año_s = "20" + var_dia_s
            End If
            If Len(var_hora_s) < 2 Then
               var_hora_s = "0" + var_hora_s
            End If
            If Len(var_minuto_s) < 2 Then
               var_minuto_s = "0" + var_minuto_s
            End If
            If Len(var_segundo_s) < 2 Then
               var_segundo_s = "0" + var_segundo_s
            End If
            
            var_fecha_inicio_o = var_dia_s + "-" + var_mes_s + "-" + var_año_s + " " + var_hora_s + ":" + var_minuto_s + ":" + var_segundo_s
            
            'var_fecha_fin_sql = "{d '" + var_año_s + "-" + var_mes_s + "-" + var_dia_s + "'}"
            var_fecha_fin_sql = " CONVERT(datetime, '" + var_año_s + "-" + var_mes_s + "-" + var_dia_s + "T" + var_hora_s + ":" + var_minuto_s + ":" + var_segundo_s + "',126)"
            
            var_dia_s = CStr(Day(CDate(Me.txt_fin)))
            var_mes_s = CStr(Month(CDate(Me.txt_fin)))
            var_año_s = CStr(Year(CDate(Me.txt_fin)))
            var_hora_s = CStr(Hour(CDate(Me.txt_fin)))
            var_minuto_s = CStr(Minute(CDate(Me.txt_fin)))
            var_segundo_s = CStr(Second(CDate(Me.txt_fin)))
            If Len(var_dia_s) = 1 Then
               var_dia_s = "0" + var_dia_s
            End If
            If Len(var_mes_s) = 1 Then
               var_mes_s = "0" + var_mes_s
            End If
            If Len(var_año_s) = 2 Then
               var_año_s = "20" + var_dia_s
            End If
            If Len(var_hora_s) < 2 Then
               var_hora_s = "0" + var_hora_s
            End If
            If Len(var_minuto_s) < 2 Then
               var_minuto_s = "0" + var_minuto_s
            End If
            If Len(var_segundo_s) < 2 Then
               var_segundo_s = "0" + var_segundo_s
            End If
            
            var_fecha_fin = var_dia_s + "-" + var_mes_s + "-" + var_año_s + " " + var_hora_s + ":" + var_minuto_s + ":" + var_segundo_s
            
            rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            rs.Open "alter session set nls_date_format='DD-MM-YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
           'rs.Open "select A.*, B.DESCRIPTION, C.LINEA from xxvia_tb_negado_distribucion A, xxvia_system_items_b b, xxvia_vw_categorias_item_b C where fecha_NEGADO > to_date('" + var_fecha_inicio + "','DD-MM-YYYY hh24:mi:ss') AND FECHA_NEGADO < TO_DATE('" + var_fecha_fin + "','DD-MM-YYYY hh24:mi:ss') and a.organization_id = b.organization_id and a.inventory_item_id = b.inventory_item_id and a.organization_id = C.organization_id and a.inventory_item_id = C.item_id and a.cantidad > 0 and a.organization_id = " + CStr(var_unidad_organizacional)
            rs.Open "select A.*, B.DESCRIPTION, C.LINEA from xxvia_tb_negado_distribucion A, xxvia_system_items_b b, xxvia_vw_categorias_item_b C, oe_order_headers_all d where fecha_NEGADO > to_date('" + var_fecha_inicio + "','DD-MM-YYYY  hh24:mi:ss') AND FECHA_NEGADO < TO_DATE('" + var_fecha_fin + "','DD-MM-YYYY hh24:mi:ss') and a.organization_id = b.organization_id and a.inventory_item_id = b.inventory_item_id and a.organization_id = C.organization_id and a.inventory_item_id = C.item_id and a.organization_id = 93 and d.order_number = to_char(a.source_header_number) and order_type_id in (1106,1161,1049, 1556, 1421)", cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               While Not rs.EOF
                     var_cadena = "insert into TB_TEMP_ORACLE_NEGADO_DISTRIBUCION (inte_tem_consecutivo, fecha_inicio, fecha_fin, pedido, tipo_pedido, ruta, cliente, codigo, cantidad_surtir, cantidad_surtida, negado_distribucion, causa_negado, descripcion, LINEA) "
                     var_causa_negado = IIf(IsNull(rs!nombre_causa_negado), "", rs!nombre_causa_negado)
                     If rs!cantidad > 0 And var_causa_negado = "" Then
                        var_causa_negado = "NO LOCALIZADO"
                     End If
                     var_cadena = var_cadena + " values (" + CStr(var_consecutivo) + "," + var_fecha_inicio_sql + "," + var_fecha_fin_sql + "," + CStr(rs!source_header_number) + ",'','','','" + rs!SEGMENT1 + "'," + CStr(rs!CANTIDAD_PEDIDA) + "," + CStr(rs!CANTIDAD_SURTIDA) + "," + CStr(rs!cantidad) + ",'" + var_causa_negado + "','" + rs!Description + "', '" + rs!Linea + "')"
                     rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               rs.Close
               rsaux10.Open "select distinct pedido from TB_TEMP_ORACLE_NEGADO_DISTRIBUCION where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " and pedido is not null", cnn, adOpenDynamic, adLockOptimistic
               While Not rsaux10.EOF
                     var_orden = rsaux10!pedido
                     strconsulta = "SELECT HEADER_ID, source_document_id, SHIP_TO_ORG_ID, A.NAME AS RUTA FROM OE_ORDER_HEADERS_ALL OHA, XXVIA_VENDEDORES A WHERE ORDER_NUMBER  = ? AND OHA.SALESREP_ID = A.SALESREP_ID"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_orden)
                          .Parameters.Append parametro
                     End With
                     Set rsaux7 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
               
                     If Not rsaux7.EOF Then
                        VAR_HEADER_ID = IIf(IsNull(rsaux7!header_id), 0, rsaux7!header_id)
                        var_requisicion = IIf(IsNull(rsaux7!source_document_id), "", rsaux7!source_document_id)
                        var_establecimiento = IIf(IsNull(rsaux7!ship_to_org_id), "0", rsaux7!ship_to_org_id)
                        var_nombre_agente_str = rsaux7!ruta
                     Else
                        VAR_HEADER_ID = 0
                     End If
                     rsaux7.Close
               
                     var_cadena = " SELECT a.source_header_type_name, HCAS.CUST_ACCOUNT_ID, HCAS.CUST_ACCT_SITE_ID as customer_id, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, oha.attribute8, oha.attribute9 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND to_number(source_header_number) = " + CStr(var_orden)
                     var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID   AND A.SOURCE_HEADER_ID = " + CStr(VAR_HEADER_ID)
                     If rs.State = 1 Then
                        rs.Close
                     End If
                     rs.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
              
                     If rs!source_header_type_name = "VIA_PEDIDO_INTERNO" Or rs!source_header_type_name = "TEX_PEDIDO_INTERNO" Then
                        If var_pedido_tienda = 0 Then
                           txt_cliente = IIf(IsNull(rs!customer_name), "", rs!customer_name)
                           rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(var_requisicion) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If Not rsaux2.EOF Then
                              txt_entrega = Replace(IIf(IsNull(rsaux2!Description), "", rsaux2!Description), "'", " ")
                           End If
                           rsaux2.Close
                        Else
                           txt_cliente = IIf(IsNull(rs!ATTRIBUTE9), "", rs!ATTRIBUTE9)
                        End If
                     Else
                        txt_cliente = IIf(IsNull(rs!customer_name), "", rs!customer_name)
                     End If
                     If txt_cliente = "VIANNEY TEXTIL HOGAR SA DE CV" Then
                        var_nombre_agente_str = "VIANNEY TEXTIL HOGAR SA DE CV"
                        txt_cliente = txt_entrega
                     End If
                     rsaux11.Open "update TB_TEMP_ORACLE_NEGADO_DISTRIBUCION set tipo_pedido = '" + rs!source_header_type_name + "', RUTA = '" + var_nombre_agente_str + "', CLIENTE = '" + txt_cliente + "' where pedido = " + CStr(var_orden), cnn, adOpenDynamic, adLockOptimistic
                     rs.Close
                     rsaux10.MoveNext
               Wend
               rsaux10.Close
               x = 0
               If x = 1 Then
                  rsaux10.Open "DELETE FROM TB_TEMP_ORACLE_NEGADO_DISTRIBUCION WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO IS NULL", cnn, adOpenDynamic, adLockOptimistic
                  Set reporte = appl.OpenReport(App.Path + "\rep_oracle_negado_distribucion.rpt")
                  reporte.RecordSelectionFormula = "{VW_ORACLE_NEGADO_DISTRIBUCION.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  reporte.ExportOptions.FormatType = crEFTExcel80
                  reporte.ExportOptions.DestinationType = crEDTDiskFile
                  archivo = "c:\reportessid\negado_distribucion_VXT_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + ".xls"
                  reporte.ExportOptions.DiskFileName = archivo
                  reporte.Export False
                  Set reporte = Nothing
                  MsgBox "Se a terminado de guardar el archivo " + archivo
               Else
                  Set oexcel = CreateObject("Excel.Application")
                  Set owbook = oexcel.Workbooks.Add
                  Set osheet = owbook.Worksheets(1)
                  osheet.Name = "NEGADO DISTRIBUCION"
                  Screen.MousePointer = vbHourglass
                  iFila = 1
                  ifila2 = 1
                  icol2 = 1
                  iCol = 1
                  var_cadena = "select FECHA_INICIO, FECHA_FIN, PEDIDO, TIPO_PEDIDO, RUTA, CLIENTE, CODIGO, DESCRIPCION, LINEA, CANTIDAD_SURTIR, CANTIDAD_SURTIDA, NEGADO_DISTRIBUCION, CAUSA_NEGADO from VW_ORACLE_NEGADO_DISTRIBUCION WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + "  AND CODIGO IS NOT NULL"
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
                  archivo = "c:\reportessid\negado_distribucion_VXT_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + ".xls"
                  owbook.SaveAs archivo
                  oexcel.Visible = True
                  Set oexcel = Nothing
                  Screen.MousePointer = vbDefault
                  rsaux10.Close
               End If
               MsgBox "Se a terminado de guardar el archivo " + archivo
            Else
               MsgBox "No existe negado para el periodo seleccionado", vbOKOnly, "ATENCION"
            End If
            If rs.State = 1 Then
               rs.Close
            End If
         Else
            MsgBox "La fecha inicial debe de ser menor o igual a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha inicial incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Command2_Click()
   MsgBox CStr(Date)
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 3000
   Left = 3500
   txt_inicio = CDate(Round(CDbl(Date), 0)) - 0.08333333
   txt_fin = CDate(Round(CDbl(Date + 1), 0)) - 0.08333333
   If var_clave_usuario_global = "U0000000356" Then
      Me.cmd_imprimir.Enabled = False
      Me.cmd_concentrado.Enabled = False
      Me.cmd_resumen.Enabled = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fin) Then
         frmcalendario.mes = CDate(Me.txt_fin)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      'txt_fin = var_fecha_general
      txt_fin = CDate(Round(CDbl(var_fecha_general + 1), 0)) - 0.08333333
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
      txt_inicio = CDate(Round(CDbl(var_fecha_general + 1), 0)) - 0.08333333
      
      'txt_inicio = var_fecha_general
   End If
End Sub

