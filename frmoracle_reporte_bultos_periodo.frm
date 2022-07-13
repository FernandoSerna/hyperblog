VERSION 5.00
Begin VB.Form frmoracle_reporte_bultos_periodo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rerporte de bultos por embarque por periodo"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   840
      Left            =   180
      TabIndex        =   2
      Top             =   435
      Width           =   4245
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   300
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
      Left            =   4065
      Picture         =   "frmoracle_reporte_bultos_periodo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmoracle_reporte_bultos_periodo.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   150
      TabIndex        =   7
      Top             =   270
      Width           =   4275
   End
End
Attribute VB_Name = "frmoracle_reporte_bultos_periodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objConn As New ADODB.Connection
Dim objCmd As New ADODB.Command
Dim objParm As ADODB.Parameter
Dim clnt As New SoapClient30
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter


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
            
            rs.Open "SELECT * FROM TB_ORACLE_TRANSPORTES", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  strconsulta = "SELECT * FROM XXVIA_TB_TRANSPORTES WHERE CLAVE = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 20, CStr(rs!clave))
                       .Parameters.Append parametro
                  End With
                  Set rsaux9 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  If Not rsaux9.EOF Then
                     strconsulta = "UPDATE XXVIA_TB_TRANSPORTES SET NOMBRE = '" + IIf(IsNull(rs!NOMBRE), "", rs!NOMBRE) + "' WHERE CLAVE = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 20, CStr(rs!clave))
                          .Parameters.Append parametro
                     End With
                     Set rsaux10 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     
                  Else
                  
                     strconsulta = "INSERT INTO XXVIA_TB_TRANSPORTES (CLAVE, NOMBRE) VALUES (?,?)"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 20, CStr(rs!clave))
                          .Parameters.Append parametro
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rs!NOMBRE), "", rs!NOMBRE))
                          .Parameters.Append parametro
                     End With
                     Set rsaux10 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                  End If
                  rs.MoveNext
            Wend
            rs.Close
            
            
            var_fecha_fin = var_dia + "/" + var_mes + "/" + var_año
            var_cadena = "select inte_Emb_embarque EMBARQUE, b.fecha_fin FECHA_ENVIO, source_header_number PEDIDO, vehiculo, C.clave, C.nombre, direccion, inte_paq_caja CAJA_EMBARQUE, tipo_caja, sello, caja_pedido, auditada, sum(floa_sal_cantidad_leida) as cantidad, B.transporte, D.NOMBRE AS NOMBRE_UNIDAD, c.clave from xxvia_tb_Salidas_cajas a, xxvia_Tb_encabezado_Embarques b, XXVIA_VW_DIRECCIONES_PEDIDOS c, XXVIA_TB_TRANSPORTES D where a.INTE_EMB_EMBARQUE = b.EMBARQUE and fecha_fin>= to_Date('" + var_fecha_inicio + "','DD/MM/YYYY') AND FECHA_fin < TO_DATE('" + var_fecha_fin + "','DD/MM/YYYY') and floa_sal_Cantidad_leida > 0 and a.source_header_number = c.order_number(+) AND B.TRANSPORTE = D.CLAVE and nvl(C.clave,' ') <> ' ' group by  inte_Emb_embarque, b.fecha_fin, source_header_number , vehiculo, C.clave, C.nombre, direccion, inte_paq_caja, tipo_caja, sello, caja_pedido, auditada, B.transporte, D.NOMBRE  order by source_header_number, caja_pedido, B.transporte"
            'MsgBox var_cadena
            rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            
            cnn.BeginTrans
            rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) AS CONSECUTIVO FROM TB_TEMP_ORACLE_REPORTE_BULTOS_EMBARQUE_PERIODO", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rsaux!CONSECUTIVO), 0, rsaux!CONSECUTIVO)
            End If
            If rsaux1.State = 1 Then
               rsaux1.Close
            End If
            rsaux1.Open "INSERT INTO TB_TEMP_ORACLE_REPORTE_BULTOS_EMBARQUE_PERIODO (INTE_tEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo + 1) + ")", cnn, adOpenDynamic, adLockOptimistic
            
            rsaux.Close
            cnn.CommitTrans
            var_consecutivo = var_consecutivo + 1
            While Not rs.EOF
                  var_mesaje = "INSERT INTO TB_TEMP_ORACLE_REPORTE_BULTOS_EMBARQUE_PERIODO  (INTE_TEM_CONSECUTIVO, EMBARQUE, FECHA_ENVIO, PEDIDO, VEHICULO, CLAVE, NOMBRE, DIRECCION, CAJA_EMBARQUE, TIPO_CAJA, SELLO,CAJA_PEDIDO, AUDITADA, CANTIDAD, TRANSPORTE, NOMBRE_UNIDAD) VALUES (" + CStr(var_consecutivo) + ", '" + CStr(rs!Embarque) + "','" + CStr(rs!FECHA_ENVIO) + "','" + CStr(rs!pedido) + "', '" + IIf(IsNull(rs!vehiculo), "", rs!vehiculo) + "', '" + rs!clave + "', '" + rs!NOMBRE + "', '" + rs!direccion + "', " + CStr(rs!CAJA_EMBARQUE) + ",'" + IIf(IsNull(rs!tipo_caja), "", rs!tipo_caja) + "', '" + IIf(IsNull(rs!sello), "", rs!sello) + "'," + CStr(IIf(IsNull(rs!caja_pedido), "", rs!caja_pedido)) + ", '" + CStr(IIf(IsNull(rs!auditada), "", rs!auditada)) + "', " + CStr(rs!cantidad) + ",'" + rs!transporte + "','" + rs!NOMBRE_UNIDAD + "')"
                  'Me.Text1 = var_mesaje
                  rsaux.Open "INSERT INTO TB_TEMP_ORACLE_REPORTE_BULTOS_EMBARQUE_PERIODO  (INTE_TEM_CONSECUTIVO, EMBARQUE, FECHA_ENVIO, PEDIDO, VEHICULO, CLAVE, NOMBRE, DIRECCION, CAJA_EMBARQUE, TIPO_CAJA, SELLO,CAJA_PEDIDO, AUDITADA, CANTIDAD, TRANSPORTE, NOMBRE_UNIDAD) VALUES (" + CStr(var_consecutivo) + ", '" + CStr(rs!Embarque) + "','" + CStr(rs!FECHA_ENVIO) + "','" + CStr(rs!pedido) + "', '" + IIf(IsNull(rs!vehiculo), "", rs!vehiculo) + "', '" + rs!clave + "', '" + rs!NOMBRE + "', '" + Mid(rs!direccion, 1, 100) + "', " + CStr(rs!CAJA_EMBARQUE) + ",'" + IIf(IsNull(rs!tipo_caja), "", rs!tipo_caja) + "', '" + Mid(Trim(IIf(IsNull(rs!sello), "", rs!sello)), 1, 20) + "'," + CStr(IIf(IsNull(rs!caja_pedido), "0", rs!caja_pedido)) + ", '" + CStr(IIf(IsNull(rs!auditada), "", rs!auditada)) + "', " + CStr(rs!cantidad) + ",'" + rs!transporte + "','" + rs!NOMBRE_UNIDAD + "')", cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
            Wend
            rsaux.Open "SELECT DISTINCT CLAVE FROM TB_TEMP_ORACLE_REPORTE_BULTOS_EMBARQUE_PERIODO WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND CLAVE IS NOT NULL", cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux.EOF
                     strconsulta = "SELECT * FROM XXVIA_TB_CLIENTES_RUTAS_DISTR A, XXVIA_tB_RUTAS_DISTRIBUCION B WHERE A.RUTA = B.RUTA AND ESTABLECIMIENTO = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 20, CStr(rsaux!clave))
                          .Parameters.Append parametro
                     End With
                     Set rsaux10 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     If Not rsaux10.EOF Then
                        rsaux11.Open "UPDATE TB_TEMP_ORACLE_REPORTE_BULTOS_EMBARQUE_PERIODO set ruta = '" + rsaux10!nombre_ruta + "'  WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND CLAVE = '" + rsaux!clave + "'", cnn, adOpenDynamic, adLockOptimistic
                     End If
                  rsaux.MoveNext
            Wend
            rsaux.Close
            rsaux.Open "SELECT DISTINCT CLAVE FROM TB_TEMP_ORACLE_REPORTE_BULTOS_EMBARQUE_PERIODO WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND CLAVE IS NOT NULL and ruta is null", cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux.EOF
                     strconsulta = "SELECT nombre_ruta FROM XXVIA_TB_CLIENTES_RUTAS_DISTR A, XXVIA_tB_RUTAS_DISTRIBUCION B, xxvia_vw_clientes_bcp c WHERE A.RUTA = B.RUTA and a.establecimiento = to_char(c.site_use_id) AND party_site_number  = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 20, CStr(rsaux!clave))
                          .Parameters.Append parametro
                     End With
                     Set rsaux10 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     If Not rsaux10.EOF Then
                        rsaux11.Open "UPDATE TB_TEMP_ORACLE_REPORTE_BULTOS_EMBARQUE_PERIODO set ruta = '" + rsaux10!nombre_ruta + "'  WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND CLAVE = '" + rsaux!clave + "'", cnn, adOpenDynamic, adLockOptimistic
                     End If
                  rsaux.MoveNext
            Wend
            rsaux.Close
            
            rs.MoveFirst
            rs.Close
            rs.Open "select * from TB_TEMP_ORACLE_REPORTE_BULTOS_EMBARQUE_PERIODO where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " and clave is not null", cnn, adOpenDynamic, adLockOptimistic
                     
            If Not rs.EOF Then
               Set oexcel = CreateObject("Excel.Application")
               Set owbook = oexcel.Workbooks.Add
               Set osheet = owbook.Worksheets(1)
               var_cadena = "PERIODO DEL " + Replace(var_fecha_inicio, "/", "_") + " AL " + Replace(var_fecha_fin, "/", "_")
               'MsgBox var_cadena
               osheet.Name = "DEL " + Replace(var_fecha_inicio, "/", "_") + " AL " + Replace(CStr(CDate(var_fecha_fin) - 1), "/", "_")
               Screen.MousePointer = vbHourglass
               iFila = 1
               ifila2 = 1
               icol2 = 1
               iCol = 1
               'rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               For i = 0 To rs.Fields.Count - 1
                   osheet.Cells(iFila, i + 1) = rs.Fields(i).Name
               Next
               rs.MoveFirst
               var_cantidad = 0
               While Not rs.EOF
                     var_cantidad = var_cantidad + IIf(IsNull(rs!cantidad), 0, rs!cantidad)
                     rs.MoveNext
               Wend
               rs.MoveFirst
               iFila = iFila + 1
               With osheet
                  ' carga los registros del recordset
                  .Cells(iFila, iCol).CopyFromRecordset rs
                  oexcel.Columns(13).Select
                  oexcel.Selection.NumberFormat = "###,###,##0.00"
                  
                  'oExcel.Columns(1).Select
                  'oExcel.Selection.Font.Color = vbRed
                  .Columns.AutoFit ' ajusta el ancho de las columnas
                  VAR_FILAS = iFila + rs.RecordCount
                  
                  .Cells(VAR_FILAS, 10).Value = "TOTAL BULTOS:"
                  .Cells(VAR_FILAS, 10).Font.Bold = True
                  .Cells(VAR_FILAS, 10).horizontalAlignment = xlRight
                  
                  .Cells(VAR_FILAS, 11).Value = rs.RecordCount
                  .Cells(VAR_FILAS, 11).NumberFormat = "###,###,##0"
                  .Cells(VAR_FILAS, 11).Font.Bold = True
                  
                  .Cells(VAR_FILAS, 12).Value = "TOTAL PIEZAS:"
                  .Cells(VAR_FILAS, 12).Font.Bold = True
                  .Cells(VAR_FILAS, 12).horizontalAlignment = xlRight
                  
                  .Cells(VAR_FILAS, 13).Value = var_cantidad
                  .Cells(VAR_FILAS, 13).Font.Bold = True
                  .Cells(VAR_FILAS, 13).NumberFormat = "###,###,##0.00"
                  .Columns.AutoFit
               End With
               
               
               With osheet
                  .Cells(1, 1).Font.Bold = True
                  .Cells(1, 2).Font.Bold = True
                  .Cells(1, 3).Font.Bold = True
                  .Cells(1, 4).Font.Bold = True
                  .Cells(1, 5).Font.Bold = True
                  .Cells(1, 6).Font.Bold = True
                  .Cells(1, 7).Font.Bold = True
                  .Cells(1, 8).Font.Bold = True
                  .Cells(1, 9).Font.Bold = True
                  .Cells(1, 10).Font.Bold = True
                  .Cells(1, 11).Font.Bold = True
                  .Cells(1, 12).Font.Bold = True
                  .Cells(1, 13).Font.Bold = True
                  .Columns.AutoFit
               End With
               
               
               
               owbook.SaveAs "c:\reportessid\reporte_de_bultos_por_embarque_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               oexcel.Visible = True
               Set oexcel = Nothing
               Screen.MousePointer = vbDefault
               
               'rs.Close
               rsaux11.Open "delete from TB_TEMP_ORACLE_REPORTE_BULTOS_EMBARQUE_PERIODO where inte_Tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            Else
               MsgBox "No existen embarques para el periodo indicado", vbOKOnly, "ATENCION"
            End If
            If rs.State = 1 Then
               rs.Close
            End If
            If rsaux10.State = 1 Then
               rsaux10.Close
            End If
         
         
         
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
