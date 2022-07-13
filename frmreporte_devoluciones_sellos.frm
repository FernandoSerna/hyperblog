VERSION 5.00
Begin VB.Form frmreporte_devoluciones_sellos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de sello de devoluciones"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   840
      Left            =   45
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
      Left            =   3930
      Picture         =   "frmreporte_devoluciones_sellos.frx":0000
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
      Picture         =   "frmreporte_devoluciones_sellos.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   15
      TabIndex        =   7
      Top             =   270
      Width           =   4275
   End
End
Attribute VB_Name = "frmreporte_devoluciones_sellos"
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
            var_cadena = "select * from tb_devoluciones "
            rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               rsaux.Open "DELETE FROM tb_Temp_oracle_reporte_sellos WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO IS NULL", cnn, adOpenDynamic, adLockOptimistic
               While Not rs.EOF
                     var_fecha_embarque = rs!FECHA_INICIO
                     var_dia = CStr(Day(var_fecha_embarque))
                     var_mes = CStr(Month(var_fecha_embarque))
                     var_año = CStr(Year(var_fecha_embarque))
                     If Len(Trim(var_dia)) = 1 Then
                        var_dia = "0" + var_dia
                     End If
                     If Len(Trim(var_mes)) = 1 Then
                        var_mes = "0" + var_mes
                     End If
                     var_fecha_embarque_inicio_s = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                     
                     var_si_Fecha_fin = CStr(IIf(IsNull(rs!fecha_fin), "", rs!fecha_fin))
                     If var_si_Fecha_fin <> "" Then
                        var_fecha_embarque = rs!fecha_fin
                        var_dia = CStr(Day(var_fecha_embarque))
                        var_mes = CStr(Month(var_fecha_embarque))
                        var_año = CStr(Year(var_fecha_embarque))
                        If Len(Trim(var_dia)) = 1 Then
                           var_dia = "0" + var_dia
                        End If
                        If Len(Trim(var_mes)) = 1 Then
                           var_mes = "0" + var_mes
                        End If
                        var_fecha_embarque_fin_s = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                     End If
                     
                     var_fecha_embarque = CDate(Me.txt_inicio)
                     var_dia = CStr(Day(var_fecha_embarque))
                     var_mes = CStr(Month(var_fecha_embarque))
                     var_año = CStr(Year(var_fecha_embarque))
                     If Len(Trim(var_dia)) = 1 Then
                        var_dia = "0" + var_dia
                     End If
                     If Len(Trim(var_mes)) = 1 Then
                        var_mes = "0" + var_mes
                     End If
                     var_fecha_inicio_s = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                     
                     var_fecha_embarque = CDate(Me.txt_fin)
                     var_dia = CStr(Day(var_fecha_embarque))
                     var_mes = CStr(Month(var_fecha_embarque))
                     var_año = CStr(Year(var_fecha_embarque))
                     If Len(Trim(var_dia)) = 1 Then
                        var_dia = "0" + var_dia
                     End If
                     If Len(Trim(var_mes)) = 1 Then
                        var_mes = "0" + var_mes
                     End If
                     var_fecha_fin_s = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                     
                     rsaux.Open "SELECT * FROM TB_USUARIOS WHERE VCHA_USU_USUARIO_ID = '" + IIf(IsNull(rs!USUARIO), "", rs!USUARIO) + "'", cnn, adOpenDynamic, adLockOptimistic
                     VAR_USUARIO = ""
                     If Not rsaux.EOF Then
                        VAR_USUARIO = IIf(IsNull(rsaux!vcha_usu_nombre), "", rsaux!vcha_usu_nombre) + " " + IIf(IsNull(rsaux!vcha_usu_apellidos), "", rsaux!vcha_usu_apellidos)
                     End If
                     rsaux.Close
                     
                     
                     
                     var_ruta = rs!ruta
                     If rs!tipo_pedido = "VIA_PEDIDO_INTERNO" Then
                        rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(IIf(IsNull(rs!source_document_id), 0, rs!source_document_id)) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux2.EOF Then
                              var_ruta = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                        End If
                        rsaux2.Close
                     End If
                     
                     var_cadena = "INSERT INTO TB_TEMP_ORACLE_REPORTE_SELLOS (INTE_TEM_CONSECUTIVO, INICIO_REPORTE, FIN_REPORTE, RUTA, CLIENTE, EMBARQUE, CAJA, ESTATUS, PEDIDO, TIPO_CAJA, SELLO, MAQUINA, USUARIO, TIPO_PEDIDO, FECHA_INICIO, FECHA_FIN, CANTIDAD)  Values"
                     If var_si_Fecha_fin <> "" Then
                        var_cadena = var_cadena + "(" + CStr(var_consecutivo) + "," + var_fecha_inicio_s + "," + var_fecha_fin_s + ",'" + IIf(IsNull(var_ruta), "", var_ruta) + "','" + rs!Cliente + "', " + CStr(rs!Embarque) + "," + CStr(rs!Caja) + ",'" + IIf(IsNull(rs!estatus), "", rs!estatus) + "'," + CStr(rs!pedido) + ",'" + rs!tipo_caja + "','" + IIf(IsNull(rs!sello), "", rs!sello) + "','" + rs!maquina + "','" + VAR_USUARIO + "','" + rs!tipo_pedido + "'," + var_fecha_embarque_inicio_s + "," + var_fecha_embarque_fin_s + "," + CStr(rs!cantidad) + ")"
                     Else
                        var_cadena = var_cadena + "(" + CStr(var_consecutivo) + "," + var_fecha_inicio_s + "," + var_fecha_fin_s + ",'" + IIf(IsNull(var_ruta), "", var_ruta) + "','" + rs!Cliente + "', " + CStr(rs!Embarque) + "," + CStr(rs!Caja) + ",'" + IIf(IsNull(rs!estatus), "", rs!estatus) + "'," + CStr(rs!pedido) + ",'" + rs!tipo_caja + "','" + IIf(IsNull(rs!sello), "", rs!sello) + "','" + rs!maquina + "','" + VAR_USUARIO + "','" + rs!tipo_pedido + "'," + var_fecha_embarque_inicio_s + ",NULL," + CStr(rs!cantidad) + ")"
                     End If
                     
                     'MsgBox var_cadena
                     rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_reporte_sellos.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_REPORTE_SELLOS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\reporte_sellos_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            Else
               MsgBox "No existen embarques para el periodo indicado", vbOKOnly, "ATENCION"
            End If
            rs.Close
            rs.Open "delete from TB_TEMP_ORACLE_REPORTE_SELLOS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
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


Private Sub txt_sello_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
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
            rs.Open "select max(inte_tem_consecutivo) as numero from tb_Temp_oracle_reporte_sellos", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into tb_Temp_oracle_reporte_sellos (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
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
            'var_cadena = "select a.source_document_id, E.NAME RUTA, CP.NOMBRE AS CLIENTE, inte_emb_embarque embarque, inte_paq_Caja caja, char_paq_estatus estatus, source_header_number pedido, tipo_Caja, sello, maquina, usuario, tipo_pedido, fecha_inicio, fecha_fin, sum(floa_sal_Cantidad_leida) as cantidad from oe_order_headers_all a, JTF_RS_SALESREPS E, oe_transaction_types_tl f, XXVIA_VW_CLIENTE_DEL_PEDIDO CP, xxvia_Tb_salidas_cajas a, xxvia_tb_encabezado_embarques b Where CP.ORDER_NUMBER = a.ORDER_NUMBER"
            'var_cadena = var_cadena + " AND inte_emb_embarque = embarque and floa_Sal_cantidad_leida > 0 AND A.SALESREP_ID = E.SALESREP_ID and f.transaction_type_id = a.order_type_id and source_lang = 'ESA' and e.org_id = 92 and floa_Sal_cantidad_leida > 0 and sello = '" + Me.txt_sello + "'  AND CP.ORDER_NUMBER = SOURCE_HEADER_NUMBER group by a.source_document_id, E.NAME, CP.NOMBRE, inte_emb_embarque, inte_paq_Caja, char_paq_estatus, source_header_number, tipo_Caja, sello, maquina, usuario, tipo_pedido, fecha_inicio, fecha_fin order by inte_emb_embarque desc"
            
            
            
            
            
            rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               rsaux.Open "DELETE FROM tb_Temp_oracle_reporte_sellos WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO IS NULL", cnn, adOpenDynamic, adLockOptimistic
               While Not rs.EOF
                     var_fecha_embarque = rs!FECHA_INICIO
                     var_dia = CStr(Day(var_fecha_embarque))
                     var_mes = CStr(Month(var_fecha_embarque))
                     var_año = CStr(Year(var_fecha_embarque))
                     If Len(Trim(var_dia)) = 1 Then
                        var_dia = "0" + var_dia
                     End If
                     If Len(Trim(var_mes)) = 1 Then
                        var_mes = "0" + var_mes
                     End If
                     var_fecha_embarque_inicio_s = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                     
                     var_si_Fecha_fin = CStr(IIf(IsNull(rs!fecha_fin), "", rs!fecha_fin))
                     If var_si_Fecha_fin <> "" Then
                        var_fecha_embarque = rs!fecha_fin
                        var_dia = CStr(Day(var_fecha_embarque))
                        var_mes = CStr(Month(var_fecha_embarque))
                        var_año = CStr(Year(var_fecha_embarque))
                        If Len(Trim(var_dia)) = 1 Then
                           var_dia = "0" + var_dia
                        End If
                        If Len(Trim(var_mes)) = 1 Then
                           var_mes = "0" + var_mes
                        End If
                        var_fecha_embarque_fin_s = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                     End If
                     
                     var_fecha_embarque = CDate(Me.txt_inicio)
                     var_dia = CStr(Day(var_fecha_embarque))
                     var_mes = CStr(Month(var_fecha_embarque))
                     var_año = CStr(Year(var_fecha_embarque))
                     If Len(Trim(var_dia)) = 1 Then
                        var_dia = "0" + var_dia
                     End If
                     If Len(Trim(var_mes)) = 1 Then
                        var_mes = "0" + var_mes
                     End If
                     var_fecha_inicio_s = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                     
                     var_fecha_embarque = CDate(Me.txt_fin)
                     var_dia = CStr(Day(var_fecha_embarque))
                     var_mes = CStr(Month(var_fecha_embarque))
                     var_año = CStr(Year(var_fecha_embarque))
                     If Len(Trim(var_dia)) = 1 Then
                        var_dia = "0" + var_dia
                     End If
                     If Len(Trim(var_mes)) = 1 Then
                        var_mes = "0" + var_mes
                     End If
                     var_fecha_fin_s = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                     
                     rsaux.Open "SELECT * FROM TB_USUARIOS WHERE VCHA_USU_USUARIO_ID = '" + IIf(IsNull(rs!USUARIO), "", rs!USUARIO) + "'", cnn, adOpenDynamic, adLockOptimistic
                     VAR_USUARIO = ""
                     If Not rsaux.EOF Then
                        VAR_USUARIO = IIf(IsNull(rsaux!vcha_usu_nombre), "", rsaux!vcha_usu_nombre) + " " + IIf(IsNull(rsaux!vcha_usu_apellidos), "", rsaux!vcha_usu_apellidos)
                     End If
                     rsaux.Close
                     
                     var_ruta = rs!ruta
                     If rs!tipo_pedido = "VIA_PEDIDO_INTERNO" Then
                        rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(IIf(IsNull(rs!source_document_id), 0, rs!source_document_id)) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux2.EOF Then
                              var_ruta = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                        End If
                        rsaux2.Close
                     End If
                     
                     var_cadena = "INSERT INTO TB_TEMP_ORACLE_REPORTE_SELLOS (INTE_TEM_CONSECUTIVO, INICIO_REPORTE, FIN_REPORTE, RUTA, CLIENTE, EMBARQUE, CAJA, ESTATUS, PEDIDO, TIPO_CAJA, SELLO, MAQUINA, USUARIO, TIPO_PEDIDO, FECHA_INICIO, FECHA_FIN, CANTIDAD)  Values"
                     If var_si_Fecha_fin <> "" Then
                        var_cadena = var_cadena + "(" + CStr(var_consecutivo) + "," + var_fecha_inicio_s + "," + var_fecha_fin_s + ",'" + var_ruta + "','" + rs!Cliente + "', " + CStr(rs!Embarque) + "," + CStr(rs!Caja) + ",'" + IIf(IsNull(rs!estatus), "", rs!estatus) + "'," + CStr(rs!pedido) + ",'" + rs!tipo_caja + "','" + IIf(IsNull(rs!sello), "", rs!sello) + "','" + rs!maquina + "','" + VAR_USUARIO + "','" + rs!tipo_pedido + "'," + var_fecha_embarque_inicio_s + "," + var_fecha_embarque_fin_s + "," + CStr(rs!cantidad) + ")"
                     Else
                        var_cadena = var_cadena + "(" + CStr(var_consecutivo) + "," + var_fecha_inicio_s + "," + var_fecha_fin_s + ",'" + var_ruta + "','" + rs!Cliente + "', " + CStr(rs!Embarque) + "," + CStr(rs!Caja) + ",'" + IIf(IsNull(rs!estatus), "", rs!estatus) + "'," + CStr(rs!pedido) + ",'" + rs!tipo_caja + "','" + IIf(IsNull(rs!sello), "", rs!sello) + "','" + rs!maquina + "','" + VAR_USUARIO + "','" + rs!tipo_pedido + "'," + var_fecha_embarque_inicio_s + ",NULL," + CStr(rs!cantidad) + ")"
                     End If
                     
                     'MsgBox var_cadena
                     rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_reporte_sellos.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_REPORTE_SELLOS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\reporte_sellos_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            Else
               MsgBox "No existen embarques para el periodo indicado", vbOKOnly, "ATENCION"
            End If
            rs.Close
            rs.Open "delete from TB_TEMP_ORACLE_REPORTE_SELLOS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "La fecha de inicio debe de ser menor o igual a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de Inicio incorrecta", vbOKOnly, "ATENCION"
   End If
   End If
End Sub

