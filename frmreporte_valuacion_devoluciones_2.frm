VERSION 5.00
Begin VB.Form frmreporte_valuacion_devoluciones_2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Valuacion de devoluciones"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   720
      Left            =   60
      TabIndex        =   3
      Top             =   420
      Width           =   4335
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   690
         TabIndex        =   5
         Top             =   255
         Width           =   1080
      End
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   2700
         TabIndex        =   4
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   315
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2385
         TabIndex        =   6
         Top             =   315
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   2
      Top             =   330
      Width           =   4485
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   45
      Picture         =   "frmreporte_valuacion_devoluciones_2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4035
      Picture         =   "frmreporte_valuacion_devoluciones_2.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "frmreporte_valuacion_devoluciones_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
   Dim var_consecutivo As Double
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim año As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   
   'On Error GoTo salir:
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            
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
            cnn.BeginTrans
            rs.Open "select max(isnull(INTE_TEM_CONSECUTIVO,0)) as numero from TB_TEMP_REPORTE_ENTRADAS_PRODUCCION", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rsaux.Open "INSERT INTO TB_TEMP_REPORTE_ENTRADAS_PRODUCCION (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTM_TEM_FECHA_FINAL, VCHA_AUR_MAQUINA, VCHA_AUD_USUARIO, VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO) Values ( " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + ", '" + fun_NombrePc + "','" + var_clave_usuario_global + "', '','', '', '',0)", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
             'rs.Open "INSERT INTO TB_TEMP_VALUACION_DEVOLUCIONES (INTE_TVD_CONSECUTIVO, DTIM_TVD_FECHA_INICIO, DTIM_TVD_FECHA_FIN, VCHA_EMP_EMPRESA_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA) select * from tb_encabezado_cartera where dtim_Car_fecha >= " + var_fecha_inicio + " and vcha_car_tipo_documento = 'NC' and VCHA_CAR_DOCUMENTO = 'DV' and dtim_car_fecha <= " + var_fecha_fin, cnn, adOpenDynamic, adLockOptimistic
              rsaux.Open "INSERT INTO TB_TEMP_REPORTE_ENTRADAS_PRODUCCION (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTM_TEM_FECHA_FINAL, VCHA_AUR_MAQUINA, VCHA_AUD_USUARIO, VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO) select " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + "-.00001, '" + fun_NombrePc + "','" + var_clave_usuario_global + "', VCHA_EMP_EMPRESA_ID , VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO from tb_encabezado_movimientos where vcha_mov_movimiento_id = 'DC' and dtim_emo_fecha >= " + var_fecha_inicio + " and dtim_emo_fecha <= " + var_fecha_fin + "-.00001", cnn, adOpenDynamic, adLockOptimistic

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
             
             'If Not rs.EOF Then
             '   While Not rs.EOF
             '      var_cadena = "INSERT INTO TB_TEMP_VALUACION_DEVOLUCIONES (INTE_TVD_CONSECUTIVO, DTIM_TVD_FECHA_INICIO, DTIM_TVD_FECHA_FIN, VCHA_EMP_EMPRESA_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA)"
             '      var_cadena = var_cadena + "Values ( " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + ", '" + rs!VCHA_EMP_EMPRESA_ID + "','" + rs!VCHA_CAR_TIPO_DOCUMENTO + "', '" + rs!vcha_ser_serie_id + "', " + CStr(rs!inte_car_numero) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "')"
             '      rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
             '      rs.MoveNext
             '   Wend
             'End If
             'rs.Close
         
         Set reporte = appl.OpenReport(App.Path + "\rep_devoluciones_clientes.rpt")
         reporte.RecordSelectionFormula = "{VW_devolucion_cliente_concentrado.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_devolucion_cliente_concentrado.vcha_Emp_empresa_id} = '" + var_empresa + "'"
         frmvistasprevias.cr.ReportSource = reporte
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Reporte de Valuación de Devoluciones"
         frmvistasprevias.Show 1
         Set reporte = Nothing
         var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            Set reporte = appl.OpenReport(App.Path + "\rep_devoluciones_clientes.rpt")
            reporte.RecordSelectionFormula = "{VW_devolucion_cliente_concentrado.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_devolucion_cliente_concentrado.vcha_Emp_empresa_id} = '" + var_empresa + "'"
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\rep_valuacion_devoluciones_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
            MsgBox "Se a terminado de guardar el archivo " + archivo
         End If
         
         rs.Open "delete from TB_TEMP_REPORTE_ENTRADAS_PRODUCCION where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "La fecha de inicio debe de ser mayor", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha Final Incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de Inicio Incorrecta", vbOKOnly, "ATENCION"
   End If
   Exit Sub
salir:
   MsgBox "A surgido un error al generar el reporte", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
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
   Call activa_forma(var_activa_forma_reporte_valuacion_devoluciones)
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fin) Then
         frmcalendario.mes.Value = CDate(Me.txt_fin)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_fin = var_fecha_general
   End If
End Sub

Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_inicio) Then
         frmcalendario.mes.Value = CDate(Me.txt_inicio)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_inicio = var_fecha_general
   End If
End Sub

