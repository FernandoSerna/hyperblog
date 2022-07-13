VERSION 5.00
Begin VB.Form frmoracle_reporte_pedidos_cargados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedidos cargados"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   30
      Picture         =   "frmoracle_reporte_pedidos_cargados.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3960
      Picture         =   "frmoracle_reporte_pedidos_cargados.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   840
      Left            =   75
      TabIndex        =   0
      Top             =   435
      Width           =   4245
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   2865
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   300
         Width           =   1140
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2325
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
      Top             =   270
      Width           =   4275
   End
End
Attribute VB_Name = "frmoracle_reporte_pedidos_cargados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comandoORA As New adodb.Command
Dim parametro As adodb.Parameter
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

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
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            cnn.CommandTimeout = 720
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_ORACLE_PEDIDOS_CARGADOS", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into TB_TEMP_ORACLE_PEDIDOS_CARGADOS (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
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
              
            var_cadena = "SELECT oha.source_document_id, oha.header_id, TL.NAME AS TIPO_PEDIDO, oha.flow_status_code, j.name as ruta, oha.salesrep_id, oha.order_type_id, oha.header_id, oha.ordered_date, oha.order_number,  HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  f.orig_system_reference, sum(ordered_quantity) as cantidad from oe_order_lines_all ohl, hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU, hz_cust_acct_sites_all f, JTF_RS_SALESREPS j, OE_TRANSACTION_TYPES_TL TL Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND  HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND  HCSU.site_use_code = 'BILL_TO' and f.cust_acct_site_id = HCAS.CUST_ACCT_SITE_ID and ordered_date >= to_date('" + var_fecha_inicio + "','DD-MM-YYYY') AND ORDERED_DATE < TO_DATE('" + var_fecha_fin + "','DD-MM-YYYY') "
            var_cadena = var_cadena + " AND OHA.ship_from_org_id = " + var_unidad_organizacional
            var_cadena = var_cadena + " and oha.salesrep_id = j.salesrep_id AND tl.transaction_type_id = oha.order_type_id AND TL.LANGUAGE = 'ESA' and oha.header_id = ohl.header_id group by oha.source_document_id, oha.header_id, TL.NAME , oha.flow_status_code, j.name , oha.salesrep_id, oha.order_type_id, oha.header_id, oha.ordered_date, oha.order_number,  HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1,  f.orig_system_reference order by TL.NAME, oha.flow_status_code, j.name, HL.ADDRESS1"
            rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               While Not rs.EOF
                     var_ruta = IIf(IsNull(rs!ruta), "", rs!ruta)
                     var_nombre_cliente = rs!CUSTOMER_NAME
                     If rs!tipo_pedido = "VIA_PEDIDO_INTERNO" Or rs!tipo_pedido = "TEX_PEDIDO_INTERNO" Then
                        var_ruta = "TIENDAS"
                        If rsaux2.State = 1 Then
                           rsaux2.Close
                        End If
                        rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(IIf(IsNull(rs!source_document_id), 0, rs!source_document_id)) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux2.EOF Then
                           var_nombre_cliente = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                        End If
                        rsaux2.Close
                     End If
               
               
                     var_cadena = "INSERT INTO TB_TEMP_ORACLE_PEDIDOS_CARGADOS (INTE_TEM_CONSECUTIVO, FECHA_INICIO, FECHA_FIN, SOURCE_DOCUMENT_ID, HEADER_ID, TIPO_PEDIDO, FLOW_STATUS_CODE, SALESREP_ID, ORDER_TYPE_ID, ORDERED_DATE, ORDER_NUMBER, CUST_ACCT_SITE_ID, CUSTOMER_NAME, CANTIDAD, RUTA)"
                     var_cadena = var_cadena + " VALUES (" + CStr(var_consecutivo) + ", " + VAR_FECHA_INICIO_TABLA + "," + VAR_FECHA_FIN_TABLA + "," + CStr(IIf(IsNull(rs!source_document_id), 0, rs!source_document_id)) + "," + CStr(rs!header_id) + ",'" + rs!tipo_pedido + "','" + rs!FLOW_STATUS_CODE + "'," + CStr(IIf(IsNull(rs!salesrep_id), 0, rs!salesrep_id)) + "," + CStr(rs!ORDER_TYPE_ID) + ",'" + CStr(rs!ORDERED_DATE) + "'," + CStr(rs!ORDER_NUMBER) + "," + CStr(rs!CUST_ACCT_SITE_ID) + ",'" + var_nombre_cliente + "'," + CStr(IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)) + ",'" + var_ruta + "')"
                     'MsgBox var_cadena
                     rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               rsaux2.Open "SELECT DISTINCT ORDER_NUMBER FROM TB_TEMP_ORACLE_PEDIDOS_CARGADOS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND ORDER_NUMBER IS NOT NULL", cnn, adOpenDynamic, adLockOptimistic
               While Not rsaux2.EOF
                     strconsulta = "SELECT INTE_EMB_EMBARQUE FROM XXVIA_TB_sALIDAS_CAJAS WHERE SOURCE_HEADER_NUMBER = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(rsaux2(0).Value))
                          .Parameters.Append parametro
                     End With
                     Set rsaux9 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     If Not rsaux9.EOF Then
                        rsaux3.Open "UPDATE TB_TEMP_ORACLE_PEDIDOS_CARGADOS SET EMBARQUE = " + CStr(rsaux9(0).Value) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND ORDER_NUMBER = '" + CStr(rsaux2(0).Value) + "'", cnn, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux9.Close
                     rsaux2.MoveNext
               Wend
               rsaux2.Close
               
               rsaux.Open "DELETE FROM TB_TEMP_ORACLE_PEDIDOS_CARGADOS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND FECHA_INICIO IS NULL", cnn, adOpenDynamic, adLockOptimistic
               
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_PEDIDOS_CARGADOS.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_PEDIDOS_CARGADOS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
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
                  Set reporte = appl.OpenReport(App.Path + "\rep_oracle_PEDIDOS_CARGADOS.rpt")
                  reporte.RecordSelectionFormula = "{VW_ORACLE_PEDIDOS_CARGADOS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  reporte.ExportOptions.FormatType = crEFTExcel80
                  reporte.ExportOptions.DestinationType = crEDTDiskFile
                  archivo = "c:\reportessid\pedidos_cargados_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                  reporte.ExportOptions.DiskFileName = archivo
                  reporte.Export False
                  Set reporte = Nothing
                  MsgBox "Se a terminado de guardar el archivo " + archivo
               End If
            Else
               MsgBox "No existen pedidos para el periodo indicado", vbOKOnly, "ATENCION"
            End If
            rs.Close
            rs.Open "delete from TB_TEMP_ORACLE_PEDIDOS_CARGADOS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
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
