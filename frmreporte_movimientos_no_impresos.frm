VERSION 5.00
Begin VB.Form frmreporte_movimientos_no_impresos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de movimientos no impresos"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmreporte_movimientos_no_impresos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3975
      Picture         =   "frmreporte_movimientos_no_impresos.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   840
      Left            =   90
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
      Left            =   0
      TabIndex        =   7
      Top             =   270
      Width           =   4395
   End
End
Attribute VB_Name = "frmreporte_movimientos_no_impresos"
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
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_MOVIMIENTOS_NO_IMPRESOS", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into TB_TEMP_MOVIMIENTOS_NO_IMPRESOS (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            var_fecha_fin_1 = CDate(txt_fin)
            var_dia = CStr(Day(CDate(txt_inicio)))
            var_mes = CStr(Month(CDate(txt_inicio)))
            var_año = CStr(Year(CDate(txt_inicio)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
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
            var_fecha_fin_2 = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            cnn.CommandTimeout = 3600
            
            
            
            
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
            var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            var_cadena = "insert into TB_TEMP_MOVIMIENTOS_NO_IMPRESOS (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_EMP_EMPRESA_ID, VCHA_EMP_NOMBRE, VCHA_UOR_UNIDAD_ID, VCHA_UOR_NOMBRE, VCHA_ALM_ALMACEN_ID, VCHA_ALM_NOMBRE, VCHA_MOV_MOVIMIENTO_ID, VCHA_MOV_NOMBRE, INTE_EMO_NUMERO, DTIM_EMO_FECHA, VCHA_TEM_USUARIO, VCHA_TEM_MAQUINA, VCHA_TEM_AFECTACION, FLOA_tEM_PIEZAS) "
            var_cadena = var_cadena + "SELECT TOP 100 PERCENT " + CStr(var_consecutivo) + " AS Expr1, " + var_fecha_inicio + " AS Expr2, " + var_fecha_fin + " AS Expr3, dbo.TB_EMPRESAS.VCHA_EMP_EMPRESA_ID, dbo.TB_EMPRESAS.VCHA_EMP_NOMBRE, dbo.TB_UNIDADESORGANIZACIONALES.VCHA_UOR_UNIDAD_ID, dbo.TB_UNIDADESORGANIZACIONALES.VCHA_UOR_NOMBRE, dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_MOVIMIENTOS.VCHA_MOV_NOMBRE, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO, dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA, dbo.TB_USUARIOS.VCHA_USU_NOMBRE + ' ' + dbo.TB_USUARIOS.VCHA_USU_APELLIDOS AS NOMBRE, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_AUD_MAQUINA, TB_MOVIMIENTOS.CHAR_MOV_AFECTACION, 0 FROM dbo.TB_ENCABEZADO_MOVIMIENTOS INNER JOIN dbo.TB_MOVIMIENTOS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID INNER JOIN "
            var_cadena = var_cadena + " dbo.TB_EMPRESAS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_EMPRESAS.VCHA_EMP_EMPRESA_ID INNER JOIN dbo.TB_UNIDADESORGANIZACIONALES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_UNIDADESORGANIZACIONALES.VCHA_UOR_UNIDAD_ID INNER JOIN dbo.TB_ALMACENES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID INNER JOIN dbo.TB_USUARIOS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_AUD_USUARIO = dbo.TB_USUARIOS.VCHA_USU_USUARIO_ID WHERE (dbo.TB_ENCABEZADO_MOVIMIENTOS.CHAR_EMO_ESTATUS IS NULL) OR (dbo.TB_ENCABEZADO_MOVIMIENTOS.CHAR_EMO_ESTATUS = '') AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA >= " + var_fecha_inicio + ") AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA <= " + var_fecha_fin + "+ 1 - .000001) AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') ORDER BY dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA"
            
            
            
            
            
            
            'var_cadena = var_cadena + " SELECT " + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin + ", dbo.TB_EMPRESAS.VCHA_EMP_EMPRESA_ID, dbo.TB_EMPRESAS.VCHA_EMP_NOMBRE, dbo.TB_UNIDADESORGANIZACIONALES.VCHA_UOR_UNIDAD_ID, dbo.TB_UNIDADESORGANIZACIONALES.VCHA_UOR_NOMBRE, dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_MOVIMIENTOS.VCHA_MOV_NOMBRE, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO, dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA FROM dbo.TB_ENCABEZADO_MOVIMIENTOS INNER JOIN"
            'var_cadena = var_cadena + " dbo.TB_MOVIMIENTOS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID INNER JOIN dbo.TB_EMPRESAS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_EMPRESAS.VCHA_EMP_EMPRESA_ID INNER JOIN dbo.TB_UNIDADESORGANIZACIONALES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_UNIDADESORGANIZACIONALES.VCHA_UOR_UNIDAD_ID INNER JOIN dbo.TB_ALMACENES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID WHERE (dbo.TB_ENCABEZADO_MOVIMIENTOS.CHAR_EMO_ESTATUS IS NULL) OR (dbo.TB_ENCABEZADO_MOVIMIENTOS.CHAR_EMO_ESTATUS = '') AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA >= " + var_fecha_inicio + ") AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA <= " + var_fecha_fin + "+1-.000001) and dbo.TB_ENCABEZADO_MOVIMIENTOS.vcha_emp_empresa_id = '" + var_empresa + "'"
            'var_cadena = var_cadena + " ORDER BY dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA"
            
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            
            var_cadena = "SELECT SUM(dbo.TB_TEMPORAL_SALIDAS.FLOA_SAL_CANTIDAD) AS CANTIDAD, dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.INTE_TEM_CONSECUTIVO, dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.VCHA_EMP_EMPRESA_ID, dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.VCHA_UOR_UNIDAD_ID, dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.INTE_EMO_NUMERO FROM dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS INNER JOIN dbo.TB_TEMPORAL_SALIDAS ON dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.VCHA_EMP_EMPRESA_ID = dbo.TB_TEMPORAL_SALIDAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.VCHA_UOR_UNIDAD_ID = dbo.TB_TEMPORAL_SALIDAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_TEMPORAL_SALIDAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.INTE_EMO_NUMERO = dbo.TB_TEMPORAL_SALIDAS.INTE_SAL_NUMERO GROUP BY dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.INTE_TEM_CONSECUTIVO, dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.VCHA_EMP_EMPRESA_ID, "
            var_cadena = var_cadena + " dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.VCHA_UOR_UNIDAD_ID, dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.INTE_EMO_NUMERO Having (dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + ")"
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  rsaux.Open "UPDATE TB_TEMP_MOVIMIENTOS_NO_IMPRESOS SET FLOA_tEM_PIEZAS = " + CStr(IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND VCHA_EMP_EMPRESA_ID = '" + rs!VCHA_EMP_EMPRESA_ID + "' AND VCHA_UOR_UNIDAD_ID = '" + rs!vcha_uor_unidad_id + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + rs!VCHA_MOV_MOVIMIENTO_ID + "' AND INTE_EMO_NUMERO = " + CStr(rs!INTE_EMO_NUMERO), cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
            Wend
            rs.Close
            
            
            
            var_cadena = "SELECT     dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.INTE_TEM_CONSECUTIVO, dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.VCHA_EMP_EMPRESA_ID, dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.VCHA_UOR_UNIDAD_ID, dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.INTE_EMO_NUMERO, SUM(dbo.tb_Temporal_entradas.FLOA_ENT_CANTIDAD) AS CANTIDAD FROM dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS INNER JOIN dbo.tb_Temporal_entradas ON dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.VCHA_EMP_EMPRESA_ID = dbo.tb_Temporal_entradas.VCHA_EMP_EMPRESA_ID AND dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.VCHA_UOR_UNIDAD_ID = dbo.tb_Temporal_entradas.VCHA_UOR_UNIDAD_ID AND dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.VCHA_ALM_ALMACEN_ID = dbo.tb_Temporal_entradas.VCHA_ALM_ALMACEN_ID AND dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.VCHA_MOV_MOVIMIENTO_ID = dbo.tb_Temporal_entradas.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.INTE_EMO_NUMERO = dbo.tb_Temporal_entradas.INTE_ENT_NUMERO "
            var_cadena = var_cadena + " GROUP BY dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.INTE_TEM_CONSECUTIVO, dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.VCHA_EMP_EMPRESA_ID, dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.VCHA_UOR_UNIDAD_ID, dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.INTE_EMO_NUMERO Having (dbo.TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + ")"
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  rsaux.Open "UPDATE TB_TEMP_MOVIMIENTOS_NO_IMPRESOS SET FLOA_tEM_PIEZAS = " + CStr(IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND VCHA_EMP_EMPRESA_ID = '" + rs!VCHA_EMP_EMPRESA_ID + "' AND VCHA_UOR_UNIDAD_ID = '" + rs!vcha_uor_unidad_id + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + rs!VCHA_MOV_MOVIMIENTO_ID + "' AND INTE_EMO_NUMERO = " + CStr(rs!INTE_EMO_NUMERO), cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
            Wend
            rs.Close
            'MsgBox var_consecutivo
            
            rs.Open "delete from TB_TEMP_MOVIMIENTOS_NO_IMPRESOS where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and dtim_tem_fecha_inicio is null", cnn, adOpenDynamic, adLockOptimistic
           
            Set reporte = appl.OpenReport(App.Path + "\rep_movimientos_no_impresos.rpt")
            reporte.RecordSelectionFormula = "{TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de movimientos no impresos"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            
            var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_movimientos_no_impresos.rpt")
               reporte.RecordSelectionFormula = "{TB_TEMP_MOVIMIENTOS_NO_IMPRESOS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\reporte_movimientos_no_impresos_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
            End If
            
            
            
            
            rs.Open "delete from TB_TEMP_MOVIMIENTOS_NO_IMPRESOS where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
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
    var_swpassword = False
    var_modifica_registro = False
    Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub txt_fin_GotFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = "Presione F5 para seleccionar la fecha"
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

Private Sub txt_fin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   End If
End Sub

Private Sub txt_fin_LostFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = ""
End Sub

Private Sub txt_inicio_GotFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = "Presione F5 para seleccionar la fecha"
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

Private Sub txt_inicio_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_inicio_LostFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = ""
End Sub


