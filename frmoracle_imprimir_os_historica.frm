VERSION 5.00
Begin VB.Form frmoracle_imprimir_os_historica 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Orden de surtido historica"
   ClientHeight    =   810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   825
      Left            =   15
      TabIndex        =   0
      Top             =   -75
      Width           =   4605
      Begin VB.TextBox txt_ordern_surtido 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2265
         TabIndex        =   2
         Top             =   270
         Width           =   2250
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Orden de surtido:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   150
         TabIndex        =   1
         Top             =   315
         Width           =   2100
      End
   End
End
Attribute VB_Name = "frmoracle_imprimir_os_historica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter

Private Sub Form_Load()
   Top = 2800
   Left = 3600
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_ordern_surtido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_ordern_surtido) Then
         If rs.State = 1 Then
            rs.Close
         End If
         rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         strconsulta = "select * from OE_ORDER_HEADERS_ALL where ORDER_NUMBER = ? "
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_ordern_surtido)
              .Parameters.Append parametro
         End With
         Set rs = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         If Not rs.EOF Then
            strconsulta = "SELECT * FROM WSH_DELIVERABLES_V WHERE SOURCE_HEADER_NUMBER = ? AND RELEASED_STATUS = 'Y'"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_ordern_surtido))
                 .Parameters.Append parametro
            End With
            Set rsaux = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            If Not rsaux.EOF Then
               VAR_ESTATUS = 1
            Else
               VAR_ESTATUS = 0
            End If
            rsaux.Close
            strconsulta = "SELECT ORDERED_ITEM, B.DESCRIPTION, SUM(NVL(ORDERED_QUANTITY,0) + NVL(CANCELLED_QUANTITY,0)) AS CANTIDAD FROM OE_ORDER_LINES_ALL A, XXVIA_SYSTEM_ITEMS_B B WHERE HEADER_ID = ? AND A.ORDERED_ITEM= B.SEGMENT1 AND B.ORGANIZATION_ID = ? GROUP BY ORDERED_ITEM, B.DESCRIPTION"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, rs!header_id)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(var_unidad_organizacional))
                 .Parameters.Append parametro
            End With
            Set rsaux = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            If Not rsaux.EOF Then
               cnn.BeginTrans
               If rsaux1.State = 1 Then
                  rsaux1.Close
               End If
               rsaux1.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_ORACLE_ORDEN_SURTIDO_HISTORICA", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux1.EOF Then
                  var_consecutivo = IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value)
               Else
                  var_consecutivo = 0
               End If
               rsaux1.Close
               var_consecutivo = var_consecutivo + 1
               rsaux1.Open "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_HISTORICA (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
               cnn.CommitTrans
               While Not rsaux.EOF
                     rsaux1.Open "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_HISTORICA (INTE_TEM_CONSECUTIVO, PEDIDO, CODIGO, DESCRIPCION, CANTIDAD_PEDIDA, CANTIDAD_SURTIR, CANTIDAD_SURTIDA, ESTATUS) VALUES (" + CStr(var_consecutivo) + "," + Me.txt_ordern_surtido + ",'" + rsaux!ORDERED_ITEM + "','" + rsaux!Description + "'," + CStr(rsaux!cantidad) + ",0,0," + CStr(VAR_ESTATUS) + ")"
                     rsaux.MoveNext
               Wend
               
               
               
               
               
               
               
               strconsulta = "SELECT   segment1, SUM(cantidad_pedida) CANTIDAD_PEDIDA, SUM(CANTIDAD_SURTIDA) CANTIDAD_SURTIDA, SUM(CANTIDAD) AS CANTIDAD_NEGADAD From xxvia_Tb_negado_distribucion Where source_header_number = ? GROUP BY segment1"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, Me.txt_ordern_surtido)
                    .Parameters.Append parametro
               End With
               Set rsaux1 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               While Not rsaux1.EOF
                     rsaux2.Open "UPDATE TB_TEMP_ORACLE_ORDEN_SURTIDO_HISTORICA SET CANTIDAD_SURTIR = " + CStr(IIf(IsNull(rsaux1!CANTIDAD_PEDIDA), 0, rsaux1!CANTIDAD_PEDIDA)) + ", CANTIDAD_SURTIDA = 0 WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND CODIGO = '" + rsaux1!SEGMENT1 + "'", cnn, adOpenDynamic, adLockOptimistic
                     rsaux1.MoveNext
               Wend
               rsaux1.Close
               
               strconsulta = "SELECT   segment1, SUM(floa_Sal_Cantidad_leida) cantidad_surtida from xxvia_tb_salidas_Cajas Where source_header_number = ? GROUP BY segment1"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, Me.txt_ordern_surtido)
                    .Parameters.Append parametro
               End With
               Set rsaux1 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               While Not rsaux1.EOF
                     rsaux2.Open "UPDATE TB_TEMP_ORACLE_ORDEN_SURTIDO_HISTORICA SET CANTIDAD_SURTIDA = " + CStr(IIf(IsNull(rsaux1!CANTIDAD_SURTIDA), 0, rsaux1!CANTIDAD_SURTIDA)) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND CODIGO = '" + rsaux1!SEGMENT1 + "'", cnn, adOpenDynamic, adLockOptimistic
                     rsaux1.MoveNext
               Wend
               rsaux1.Close
               
               
               
               strconsulta = "SELECT DISTINCT CUST_ACCOUNT_ID, site_use_id,A.DELIVERY_ID, HL.ADDRESS1 AS CUSTOMER_NAME, a.source_header_type_name, oha.source_document_id, j.NAME AS nombre_ruta, j.salesrep_id AS clave_ruta FROM hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, XXVIA_VENDEDORES j Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID = OHA.INVOICE_TO_ORG_ID AND to_number(source_header_number) = ? AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND oha.salesrep_id = j.salesrep_id"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, Me.txt_ordern_surtido)
                    .Parameters.Append parametro
               End With
               Set rsaux1 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               If Not rsaux1.EOF Then
                  
                  strconsulta = "SELECT distinct  creation_date FROM WSH_DLVB_DLVY_V where delivery_id = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, rsaux1!delivery_id)
                       .Parameters.Append parametro
                  End With
                  Set rsaux2 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  If Not rsaux2.EOF Then
                     var_fecha_creacion = rsaux2!creation_Date
                  Else
                     
                     strconsulta = "SELECT distinct  creation_date FROM oe_order_headers_all where order_number = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_ordern_surtido)
                          .Parameters.Append parametro
                     End With
                     Set rsaux3 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     If Not rsaux3.EOF Then
                        var_fecha_creacion = rsaux3!creation_Date
                     Else
                        var_fecha_creacion = Now
                     End If
                  End If
                  rsaux2.Close
                  var_dia_s = CStr(Day(CDate(var_fecha_creacion)))
                  var_mes_s = CStr(Month(CDate(var_fecha_creacion)))
                  var_año_s = CStr(Year(CDate(var_fecha_creacion)))
                  If Len(var_dia_s) = 1 Then
                     var_dia_s = "0" + var_dia_s
                  End If
                  If Len(var_mes_s) = 1 Then
                     var_mes_s = "0" + var_mes_s
                  End If
                  If Len(var_año_s) = 2 Then
                     var_año_s = "20" + var_dia_s
                  End If
                  
                  var_fecha_os = "{d '" + var_año_s + "-" + var_mes_s + "-" + var_dia_s + "'}"
                  
                  
                  VAR_NOMBRE_PROVEEDOR = ""
                  If rsaux1!source_header_type_name = "VIA_PEDIDO_INTERNO" Or rsaux1!source_header_type_name = "TEX_PEDIDO_INTERNO" Then
                     rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = " + CStr(rsaux1!source_document_id) + " AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux2.EOF Then
                        VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                     End If
                     rsaux2.Close
                  Else
                     rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux4.EOF Then
                        VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                     End If
                     rsaux4.Close
                  End If
                  rsaux4.Open "UPDATE TB_TEMP_ORACLE_ORDEN_SURTIDO_HISTORICA SET fecha = " + var_fecha_os + ", RUTA = '" + rsaux1!nombre_ruta + "', AGENTE = '" + VAR_NOMBRE_PROVEEDOR + "', CLIENTE = '" + rsaux1!customer_name + "' WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
               End If
               rsaux1.Close
               rsaux1.Open "SELECT * FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES WHERE PEDIDO = '" + Me.txt_ordern_surtido + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux1.EOF Then
                  rsaux4.Open "UPDATE  TB_TEMP_ORACLE_ORDEN_SURTIDO_HISTORICA SET EMBARQUE = '" + CStr(IIf(IsNull(rsaux1!Embarque), "", rsaux1!Embarque)) + "'", cnn, adOpenDynamic, adLockOptimistic
                  If CStr(IIf(IsNull(rsaux1!Embarque), "", rsaux1!Embarque)) <> "" Then
                     strconsulta = "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(rsaux1!Embarque))
                          .Parameters.Append parametro
                     End With
                     Set rsaux3 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     If Not rsaux3.EOF Then
                        If IIf(IsNull(rsaux3!FECHA_INICIO), "", rsaux3!FECHA_INICIO) <> "" Then
                           var_dia = CStr(Day(rsaux3!FECHA_INICIO))
                           var_mes = CStr(Month(rsaux3!FECHA_INICIO))
                           var_año = CStr(Year(rsaux3!FECHA_INICIO))
                           If Len(var_dia) = 1 Then
                              var_dia = "0" + var_dia
                           End If
                           If Len(var_mes) = 1 Then
                              var_mes = "0" + var_mes
                           End If
                           var_fecha_str = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                           rsaux4.Open "UPDATE  TB_TEMP_ORACLE_ORDEN_SURTIDO_HISTORICA SET FECHA_EMBARQUE = " + var_fecha_str, cnn, adOpenDynamic, adLockOptimistic
                        End If
                     End If
                  End If
               End If
               rsaux1.Close
               rsaux1.Open "delete from TB_TEMP_ORACLE_ORDEN_SURTIDO_HISTORICA where inte_tem_Consecutivo = " + CStr(var_consecutivo) + " and pedido is null", cnn, adOpenDynamic, adLockOptimistic
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_historica.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO_HISTORICA.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Ordenes de surtido historica"
               frmvistasprevias.Show 1
               Set reporte = Nothing
               var_si = MsgBox("¿Desea exportar el pedido a excel?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_historica.rpt")
                  reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO_HISTORICA.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  reporte.ExportOptions.FormatType = crEFTExcel80
                  reporte.ExportOptions.DestinationType = crEDTDiskFile
                  archivo = "c:\reportessid\orden_surtido_" + Me.txt_ordern_surtido + "_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                  reporte.ExportOptions.DiskFileName = archivo
                  reporte.Export False
                  Set reporte = Nothing
                  MsgBox "Se a terminado de guardar el archivo " + archivo
               End If
               
               
               rsaux1.Open "delete from TB_TEMP_ORACLE_ORDEN_SURTIDO_HISTORICA where inte_tem_Consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            Else
               MsgBox "La orden de surtido esta vacia", vbOKOnly, "ATENCION"
            End If
            rsaux.Close
            
         Else
            MsgBox "La orden de surtido no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
      
      Else
         MsgBox "Número de orden de surtido incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
End Sub
