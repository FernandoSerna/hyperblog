VERSION 5.00
Begin VB.Form frmreporte_entradas_calidad_auditoria 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte entradas a calidad"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frmreporte_entradas_calidad_auditoria.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4005
      Picture         =   "frmreporte_entradas_calidad_auditoria.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   840
      Left            =   120
      TabIndex        =   0
      Top             =   450
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
      Left            =   30
      TabIndex        =   7
      Top             =   285
      Width           =   4395
   End
End
Attribute VB_Name = "frmreporte_entradas_calidad_auditoria"
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
            rs.Open "select max(INTE_TEM_CONSsECUTIVO) as numero from TB_TEMP_REPORTE_ENTRADAS_CALIDAD_AUDITORIA", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into TB_TEMP_REPORTE_ENTRADAS_CALIDAD_AUDITORIA (INTE_TEM_CONSsECUTIVO) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            var_fecha_fin_1 = CDate(txt_fin) + 1
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
            
            
            
            
            var_cadena = "INSERT INTO TB_TEMP_REPORTE_ENTRADAS_CALIDAD_AUDITORIA SELECT " + CStr(var_consecutivo) + ",  " + var_fecha_inicio + ", " + var_fecha_fin + ",dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_MOVIMIENTOS.VCHA_MOV_NOMBRE, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO, dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA, dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL, dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD , dbo.TB_ENTRADAS.FLOA_ENT_COSTO, dbo.TB_ENTRADAS.FLOA_ENT_PRECIO FROM  dbo.TB_CLIENTES INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_CLI_CLAVE_ID INNER JOIN  dbo.TB_AGENTES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_AGE_AGENTE_ID = dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID INNER JOIN dbo.TB_ARTICULOS INNER JOIN "
            var_cadena = var_cadena + " dbo.TB_ENTRADAS ON dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID = dbo.TB_ENTRADAS.VCHA_ART_ARTICULO_ID ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_ENTRADAS.INTE_ENT_NUMERO INNER JOIN dbo.TB_MOVIMIENTOS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID whERE (dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = 'CA' OR dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = 'CAVT') AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA >= " + var_fecha_inicio + ") AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA <= " + var_fecha_fin + ")"
            var_cadena = var_cadena + "  and tb_articulos.vcha_Art_articulo_id <> '---' and tb_entradas.vcha_alm_almacen_id = '14' ORDER BY dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA"
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
             
            var_cadena = "INSERT INTO TB_TEMP_REPORTE_ENTRADAS_CALIDAD_AUDITORIA SELECT " + CStr(var_consecutivo) + ",  " + var_fecha_inicio + ", " + var_fecha_fin + ", dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_MOVIMIENTOS.VCHA_MOV_NOMBRE, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO, dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA, dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL, dbo.TB_ENTRADAS.FLOA_ENT_CANTIDAD , dbo.TB_ENTRADAS.FLOA_ENT_COSTO, dbo.TB_ENTRADAS.FLOA_ENT_PRECIO FROM dbo.TB_ARTICULOS INNER JOIN dbo.TB_ENTRADAS ON dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID = dbo.TB_ENTRADAS.VCHA_ART_ARTICULO_ID INNER JOIN dbo.TB_ENCABEZADO_MOVIMIENTOS ON dbo.TB_ENTRADAS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENTRADAS.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID AND "
            var_cadena = var_cadena + " dbo.TB_ENTRADAS.VCHA_ALM_ALMACEN_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENTRADAS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENTRADAS.INTE_ENT_NUMERO = dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO INNER JOIN dbo.TB_MOVIMIENTOS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID INNER JOIN dbo.TB_CLIENTES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_PRO_PROVEEDOR_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID INNER JOIN dbo.TB_AGENTES ON dbo.TB_CLIENTES.VCHA_AGE_AGENTE_ID = dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID WHERE (dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = 'DT') AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA >= " + var_fecha_inicio + ") AND (dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA <= " + var_fecha_fin + ")  and tb_articulos.vcha_Art_articulo_id <> '---'  and tb_entradas.vcha_alm_almacen_id = '14' ORDER BY dbo.TB_ENCABEZADO_MOVIMIENTOS.DTIM_EMO_FECHA"
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            rs.Open "delete from TB_TEMP_REPORTE_ENTRADAS_CALIDAD_AUDITORIA where inte_tem_conssecutivo = " + CStr(var_consecutivo) + " and dtim_emo_fecha is null", cnn, adOpenDynamic, adLockOptimistic
            Set reporte = appl.OpenReport(App.Path + "\rep_entradas_calidad_auditoria.rpt")
            reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_CALIDAD_AUDITORIA.INTE_TEM_CONSsECUTIVO} = " + CStr(var_consecutivo)
            For ntablas = 1 To reporte.Database.Tables.Count
               reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\Reporte_entradas_calidad_auditoria_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
            MsgBox "Se a terminado de guardar el archivo " + archivo
            rs.Open "delete from tb_temp_reporte_entradas_Calidad_auditoria where INTE_TEM_CONSsECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
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


