VERSION 5.00
Begin VB.Form frmoracle_reporte_control_relacion_mayoreo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control relación mayoreo"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   840
      Left            =   90
      TabIndex        =   4
      Top             =   435
      Width           =   4245
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   2
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
      Left            =   3975
      Picture         =   "frmoracle_reporte_control_relacion_mayoreo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   45
      Picture         =   "frmoracle_reporte_control_relacion_mayoreo.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   60
      TabIndex        =   7
      Top             =   270
      Width           =   4275
   End
End
Attribute VB_Name = "frmoracle_reporte_control_relacion_mayoreo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
            rs.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_ORACLE_CONTROL_RELACION_MAYOREO", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!numero), 0, rs!numero)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into TB_TEMP_ORACLE_CONTROL_RELACION_MAYOREO (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
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
            
            var_fecha_inicio = var_dia + "/" + var_mes + "/" + var_año
            var_fecha_inicio_reporte = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
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
            var_fecha_FIN_REPORTE = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
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
            var_fecha_fin = var_dia + "/" + var_mes + "/" + var_año
            
            
            
             
            rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
              
            var_cadena = "SELECT EMBARQUE FROM XXVIA_tB_ENCABEZADO_EMBARQUES WHERE fecha_fin >= to_date('" + var_fecha_inicio + "','dd/mm/yyyy') and fecha_fin < to_date('" + var_fecha_fin + "','dd/mm/yyyy') and organizacion = " + var_unidad_organizacional
            rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_cadena_embarques = ""
               While Not rs.EOF
                     If var_cadena_embarques = "" Then
                        var_cadena_embarques = CStr(rs!Embarque)
                     Else
                        var_cadena_embarques = var_cadena_embarques + "," + CStr(rs!Embarque)
                     End If
                     rs.MoveNext
               Wend
               rsaux.Open "select a.source_document_id, E.NAME RUTA, f.NAME TIPO, embarque, fecha_fin fecha,  a.order_number pedido, sum(shipped_quantity) piezas, sum(shipped_quantity*unit_selling_price) IMPORTE from oe_order_headers_all a, oe_order_lines_all b, XXVIA_ORDENES_EMBARQUES c, xxvia_tb_encabezado_embarques d, JTF_RS_SALESREPS E, oe_transaction_types_tl f  where a.header_id = b.header_id and inte_emb_embarque in (" + var_cadena_embarques + ") and inte_emb_embarque = embarque and  c.source_header_number = a.order_number  and  shipped_quantity >0 AND A.SALESREP_ID = E.SALESREP_ID  and f.transaction_type_id = a.order_type_id and source_lang = 'ESA' and e.org_id = " + var_empresa + " group by a.source_document_id, E.NAME, f.NAME, embarque, fecha_fin,  a.order_number", cnnoracle_4, adOpenDynamic, adLockOptimistic
              'rsaux.Open "select a.source_document_id, E.NAME RUTA, SOURCE_HEADER_TYPE_NAME TIPO, embarque, fecha_fin FECHA,  order_number PEDIDO, sum(shipped_quantity) PIEZAS, sum(shipped_quantity*unit_price) IMPORTE from oe_order_headers_all a, wsh_deliverables_v b, XXVIA_ORDENES_EMBARQUES c, xxvia_tb_encabezado_embarques d, JTF_RS_SALESREPS E  where a.order_number = b.source_header_number and inte_emb_embarque in (" + var_cadena_embarques + ") and inte_emb_embarque = embarque and  c.source_header_number = order_number  and  released_status = 'C' AND A.SALESREP_ID = E.SALESREP_ID   group by source_document_id, E.NAME, SOURCE_HEADER_TYPE_NAME, embarque, fecha_fin,  order_number", cnnoracle_4, adOpenDynamic, adLockOptimistic
               While Not rsaux.EOF
                     var_fecha_embarque = rsaux!Fecha
                     var_dia = CStr(Day(var_fecha_embarque))
                     var_mes = CStr(Month(var_fecha_embarque))
                     var_año = CStr(Year(var_fecha_embarque))
                     If Len(Trim(var_dia)) = 1 Then
                        var_dia = "0" + var_dia
                     End If
                     If Len(Trim(var_mes)) = 1 Then
                        var_mes = "0" + var_mes
                     End If
                     var_fecha_embarque_s = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                     var_ruta = rsaux!ruta
                     If rsaux!tipo = "VIA_PEDIDO_INTERNO" Then
                        rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(rsaux!source_document_id) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux2.EOF Then
                              var_ruta = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                        End If
                        rsaux2.Close
                        
                     End If
                     rsaux1.Open "INSERT INTO TB_TEMP_ORACLE_CONTROL_RELACION_MAYOREO (INTE_TEM_CONSECUTIVO, EMBARQUE, FECHA, TIPO, DESTINO, PIEZAS, IMPORTE, SOURCE_DOCUMENT_ID, PEDIDO, FECHA_INICIO, FECHA_FIN) VALUES (" + CStr(var_consecutivo) + "," + CStr(rsaux!Embarque) + ", " + var_fecha_embarque_s + ",'" + rsaux!tipo + "','" + var_ruta + "'," + CStr(rsaux!PIEZAS) + "," + CStr(rsaux!Importe) + ",'" + CStr(IIf(IsNull(rsaux!source_document_id), "", rsaux!source_document_id)) + "'," + CStr(rsaux!PEDIDO) + "," + var_fecha_inicio_reporte + "," + var_fecha_FIN_REPORTE + ")", cnn, adOpenDynamic, adLockOptimistic
                     rsaux.MoveNext
               Wend
               rsaux.Close
               rsaux.Open "delete from TB_TEMP_ORACLE_CONTROL_RELACION_MAYOREO where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and embarque is null", cnn, adOpenDynamic, adLockOptimistic
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_control_relacion_mayoreo.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_CONTROL_RELACION_MAYOREO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
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
                  Set reporte = appl.OpenReport(App.Path + "\rep_oracle_control_relacion_mayoreo.rpt")
                  reporte.RecordSelectionFormula = "{VW_ORACLE_CONTROL_RELACION_MAYOREO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
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
               MsgBox "No existen embarques para el periodo indicado", vbOKOnly, "ATENCION"
            End If
            rs.Close
            rs.Open "delete from TB_TEMP_ORACLE_CONTROL_RELACION_MAYOREO WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
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
   Call activa_forma(var_activa_forma_lineas)
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

