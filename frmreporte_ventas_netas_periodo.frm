VERSION 5.00
Begin VB.Form frmreporte_ventas_netas_periodo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de ventas netas por periodo"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   840
      Left            =   120
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
      Left            =   4035
      Picture         =   "frmreporte_ventas_netas_periodo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frmreporte_ventas_netas_periodo.frx":063A
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
      Width           =   4395
   End
End
Attribute VB_Name = "frmreporte_ventas_netas_periodo"
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
   Dim a?o As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_consecutivo As Integer
   Dim var_afectacion_movimiento As String
   Dim var_vistas As String
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_REPORTE_VENTAS_NETAS_PERIODO", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into TB_TEMP_REPORTE_VENTAS_NETAS_PERIODO (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            var_fecha_fin_1 = CDate(txt_fin)
            var_dia = CStr(Day(CDate(txt_inicio)))
            var_mes = CStr(Month(CDate(txt_inicio)))
            var_a?o = CStr(Year(CDate(txt_inicio)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_inicio = "{d '" + var_a?o + "-" + var_mes + "-" + var_dia + "'}"
            
            var_fecha_fin_1 = CDate(Me.txt_fin)
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_a?o = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = "{d '" + var_a?o + "-" + var_mes + "-" + var_dia + "'}"
            
            cnn.CommandTimeout = 3600
            
            var_cadena = "INSERT INTO TB_TEMP_REPORTE_VENTAS_NETAS_PERIODO (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_EMP_EMPRESA_ID, VCHA_CLI_CLAVE_ID, FLOA_TEM_IMPORTE, VCHA_CAN_CANAL_VENTA_ID, VCHA_CAN_NOMBRE, VCHA_AGE_AGENTE_ID, VCHA_TIT_TITULAR_ID, VCHA_GAC_GRUPO_ACTUAL_ID, VCH_GRE_GRUPO_REAL_ID) "
            var_cadena = var_cadena + " SELECT " + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin + ",dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID, SUM(dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_IMPORTE_NETO) AS IMPORTE, dbo.TB_CANALESVENTAS.VCHA_CAN_CANAL_VENTA_ID, dbo.TB_CANALESVENTAS.VCHA_CAN_NOMBRE, dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.VW_CLIENTES.VCHA_TIT_TITULAR_ID, dbo.VW_CLIENTES.VCHA_GAC_GRUPO_ACTUAL_ID , dbo.VW_CLIENTES.VCHA_GRE_GRUPO_REAL_ID FROM dbo.VW_CLIENTES INNER JOIN dbo.TB_ENCABEZADO_CARTERA ON dbo.VW_CLIENTES.VCHA_CLI_CLAVE_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID INNER JOIN dbo.TB_CANALESVENTAS INNER JOIN dbo.TB_AGENTES ON dbo.TB_CANALESVENTAS.VCHA_CAN_CANAL_VENTA_ID = dbo.TB_AGENTES.VCHA_CAN_CANAL_VENTA_ID ON dbo.VW_CLIENTES.VCHA_AGE_AGENTE_ID = dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID WHERE (dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_TIPO_DOCUMENTO = 'FA') AND (dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS <> 'C' OR "
            var_cadena = var_cadena + " dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS IS NULL) and dtim_car_fecha >= " + var_fecha_inicio + " and dtim_car_fecha < " + var_fecha_fin + "+1 GROUP BY dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID, dbo.TB_CANALESVENTAS.VCHA_CAN_CANAL_VENTA_ID, dbo.TB_CANALESVENTAS.VCHA_CAN_NOMBRE, dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.VW_CLIENTES.VCHA_TIT_TITULAR_ID, dbo.VW_CLIENTES.VCHA_GAC_GRUPO_ACTUAL_ID , dbo.VW_CLIENTES.VCHA_GRE_GRUPO_REAL_ID"
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            
            var_cadena = "INSERT INTO TB_TEMP_REPORTE_VENTAS_NETAS_PERIODO (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_EMP_EMPRESA_ID, VCHA_CLI_CLAVE_ID, FLOA_TEM_IMPORTE, VCHA_CAN_CANAL_VENTA_ID, VCHA_CAN_NOMBRE, VCHA_AGE_AGENTE_ID, VCHA_TIT_TITULAR_ID, VCHA_GAC_GRUPO_ACTUAL_ID, VCH_GRE_GRUPO_REAL_ID) "
            var_cadena = var_cadena + " SELECT  " + CStr(var_consecutivo) + ", " + var_fecha_inicio + "," + var_fecha_fin + ", dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID, 0 - SUM(dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_IMPORTE_NETO) AS IMPORTE, dbo.TB_CANALESVENTAS.VCHA_CAN_CANAL_VENTA_ID, dbo.TB_CANALESVENTAS.VCHA_CAN_NOMBRE, dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.VW_CLIENTES.VCHA_TIT_TITULAR_ID, dbo.VW_CLIENTES.VCHA_GAC_GRUPO_ACTUAL_ID , dbo.VW_CLIENTES.VCHA_GRE_GRUPO_REAL_ID FROM dbo.VW_CLIENTES INNER JOIN dbo.TB_ENCABEZADO_CARTERA ON dbo.VW_CLIENTES.VCHA_CLI_CLAVE_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID INNER JOIN dbo.TB_CANALESVENTAS INNER JOIN dbo.TB_AGENTES ON dbo.TB_CANALESVENTAS.VCHA_CAN_CANAL_VENTA_ID = dbo.TB_AGENTES.VCHA_CAN_CANAL_VENTA_ID ON"
            var_cadena = var_cadena + " dbo.VW_CLIENTES.VCHA_AGE_AGENTE_ID = dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID WHERE (dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_TIPO_DOCUMENTO = 'NC' OR dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_TIPO_DOCUMENTO = 'CA') AND (dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS <> 'C' OR dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS IS NULL) AND DTIM_CAR_FECHA >= " + var_fecha_inicio + " AND DTIM_cAR_FECHA < " + var_fecha_fin + "+1 GROUP BY dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID, dbo.TB_CANALESVENTAS.VCHA_CAN_CANAL_VENTA_ID, dbo.TB_CANALESVENTAS.VCHA_CAN_NOMBRE, dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.VW_CLIENTES.VCHA_TIT_TITULAR_ID , dbo.VW_CLIENTES.VCHA_GAC_GRUPO_ACTUAL_ID, dbo.VW_CLIENTES.VCHA_GRE_GRUPO_REAL_ID "
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            
            rs.Open "delete from TB_TEMP_REPORTE_VENTAS_NETAS_PERIODO where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and dtim_tem_fecha_inicio is null", cnn, adOpenDynamic, adLockOptimistic
            
            Set reporte = appl.OpenReport(App.Path + "\rep_ventas_netas_periodo.rpt")
            reporte.RecordSelectionFormula = "{VW_TEMP_REPORTE_VENTAS_NETAS_PERIDO.inte_tem_consecutivo} = " + CStr(var_consecutivo) + " and {VW_TEMP_REPORTE_VENTAS_NETAS_PERIDO.vcha_emp_empresa_id} = '" + var_empresa + "'"
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\Reporte_ventas_netas_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
            MsgBox "Se a terminado de guardar el archivo " + archivo
            
            
           
            rs.Open "delete from TB_TEMP_REPORTE_VENTAS_NETAS_PERIODO where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
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
    Call activa_forma(var_activa_forma_articulos2)
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



