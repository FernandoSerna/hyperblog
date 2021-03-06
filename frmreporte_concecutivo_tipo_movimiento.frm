VERSION 5.00
Begin VB.Form frmreporte_concecutivo_tipo_documento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consecutivo por tipo de documento"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_dolares 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frmreporte_concecutivo_tipo_movimiento.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Imprimir Movimiento en dolares"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   840
      Left            =   90
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
      Left            =   3975
      Picture         =   "frmreporte_concecutivo_tipo_movimiento.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmreporte_concecutivo_tipo_movimiento.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   0
      Width           =   330
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
Attribute VB_Name = "frmreporte_concecutivo_tipo_documento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_dolares_Click()
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
            rs.Open "select max(inte_temp_consecutivo) as numero from TB_TEMP_CONSECUTIVO_CARTERA", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into TB_TEMP_CONSECUTIVO_CARTERA (inte_temp_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            var_fecha_fin_1 = CDate(txt_fin) + 1
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
            rs.Open " INSERT INTO TB_TEMP_CONSECUTIVO_CARTERA ( INTE_TEMP_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_EMP_EMPRESA_ID, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, VCHA_CLI_CLAVE_ID) select " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + "-.000001, VCHA_EMP_EMPRESA_ID, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, VCHA_CLI_CLAVE_ID from tb_encabezado_cartera where dtim_car_fecha >= " + var_fecha_inicio + " and dtim_car_fecha <= " + var_fecha_fin + "-.0000001 and char_car_afectacion = '+'", cnn, adOpenDynamic, adLockOptimistic
            
            
            var_fecha_fin_1 = CDate(txt_fin)
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_a?o = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin_2 = "{d '" + var_a?o + "-" + var_mes + "-" + var_dia + "'}"
            
            
            If var_empresa = "18" Then
               var_si = MsgBox("?Deseas el reporte dividido por agente?", vbYesNo, "ATENCION")
            Else
               var_si = 6
            End If
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_cartera_consecutivo_CARGOS_2_DOLARES.rpt")
               reporte.RecordSelectionFormula = "{VW_CARTERA_CONSECUTIVO_CARGOS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_CARTERA_CONSECUTIVO_CARGOS.INTE_TEMP_CONSECUTIVO} = " + CStr(var_consecutivo)
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Reporte Consecutivo de Movimientos por Agente (cargos)"
               frmvistasprevias.Show 1
               Set reporte = Nothing
               var_si = MsgBox("?Desea importar el reporte?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  Set reporte = appl.OpenReport(App.Path + "\rep_cartera_consecutivo_CARGOS_2_DOLARES.rpt")
                  reporte.RecordSelectionFormula = "{VW_CARTERA_CONSECUTIVO_CARGOS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_CARTERA_CONSECUTIVO_CARGOS.INTE_TEMP_CONSECUTIVO} = " + CStr(var_consecutivo)
                  For ntablas = 1 To reporte.Database.Tables.Count
                     reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  reporte.ExportOptions.FormatType = crEFTExcel80
                  reporte.ExportOptions.DestinationType = crEDTDiskFile
                  archivo = "c:\reportessid\reporte_consecutivo_movimientos_cargos" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                  reporte.ExportOptions.DiskFileName = archivo
                  reporte.Export False
                  Set reporte = Nothing
                  MsgBox "Se a terminado de guardar el archivo " + archivo
               End If
            Else
               Set reporte = appl.OpenReport(App.Path + "\rep_cartera_consecutivo_CARGOS_2_no_dividido_DOLARES.rpt")
               reporte.RecordSelectionFormula = "{VW_CARTERA_CONSECUTIVO_CARGOS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_CARTERA_CONSECUTIVO_CARGOS.INTE_TEMP_CONSECUTIVO} = " + CStr(var_consecutivo)
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Reporte Consecutivo de Movimientos por Agente (cargos)"
               frmvistasprevias.Show 1
               Set reporte = Nothing
               var_si = MsgBox("?Desea importar el reporte?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  Set reporte = appl.OpenReport(App.Path + "\rep_cartera_consecutivo_CARGOS_2_no_dividido_DOLARES.rpt")
                  reporte.RecordSelectionFormula = "{VW_CARTERA_CONSECUTIVO_CARGOS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_CARTERA_CONSECUTIVO_CARGOS.INTE_TEMP_CONSECUTIVO} = " + CStr(var_consecutivo)
                  For ntablas = 1 To reporte.Database.Tables.Count
                     reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  reporte.ExportOptions.FormatType = crEFTExcel80
                  reporte.ExportOptions.DestinationType = crEDTDiskFile
                  archivo = "c:\reportessid\reporte_consecutivo_movimientos_cargos" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                  reporte.ExportOptions.DiskFileName = archivo
                  reporte.Export False
                  Set reporte = Nothing
                  MsgBox "Se a terminado de guardar el archivo " + archivo
               End If
            End If
            rs.Open "delete from TB_TEMP_CONSECUTIVO_CARTERA where INTE_TEMP_CONSECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            
            rs.Open " INSERT INTO [TB_TEMP_CONSECUTIVO_CARTERA] ( [INTE_TEMP_CONSECUTIVO], [DTIM_TEM_FECHA_INICIO], [DTIM_TEM_FECHA_FIN], [VCHA_EMP_EMPRESA_ID], [VCHA_CAR_DOCUMENTO], [VCHA_SER_SERIE_ID], [INTE_CAR_NUMERO], [VCHA_CLI_CLAVE_ID]) select " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + "-.000001, VCHA_EMP_EMPRESA_ID, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, VCHA_CLI_CLAVE_ID from tb_encabezado_cartera where dtim_car_fecha >= " + var_fecha_inicio + " and dtim_car_fecha <= " + var_fecha_fin + "-.0000001 and char_car_afectacion = '-'", cnn, adOpenDynamic, adLockOptimistic
            Set reporte = appl.OpenReport(App.Path + "\rep_cartera_consecutivo_abonos_2_DOLARES.rpt")
            reporte.RecordSelectionFormula = "{VW_CONSECUTIVO_CARTERA_ABONOS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_CONSECUTIVO_CARTERA_ABONOS.INTE_TEMP_CONSECUTIVO} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte Consecutivo de Movimientos por Agente (abonos)"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            var_si = MsgBox("?Desea importar el reporte?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_cartera_consecutivo_abonos_2_DOLARES.rpt")
               reporte.RecordSelectionFormula = "{VW_CONSECUTIVO_CARTERA_ABONOS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_CONSECUTIVO_CARTERA_ABONOS.INTE_TEMP_CONSECUTIVO} = " + CStr(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\reporte_consecutivo_movimientos_abonos" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
            rs.Open "delete from TB_TEMP_CONSECUTIVO_CARTERA where INTE_TEMp_COnsECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
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
            var_fecha_fin_1 = CDate(txt_fin) + 1
            var_dia = CStr(Day(CDate(txt_inicio)))
            var_mes = CStr(Month(CDate(txt_inicio)))
            var_a?o = CStr(Year(CDate(txt_inicio)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_inicio = var_a?o + "," + var_mes + "," + var_dia
            
            
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_a?o = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = var_a?o + "," + var_mes + "," + var_dia
            
            
            
            var_fecha_fin_1 = CDate(txt_fin)
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_a?o = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin_2 = "{d '" + var_a?o + "-" + var_mes + "-" + var_dia + "'}"
            
            Set reporte = appl.OpenReport(App.Path + "\rep_oracle_consecutivo_tipo_movimientos_cargos.rpt")
            reporte.RecordSelectionFormula = "{XXVIA_JV_CONSECUTIVO_DE_CARGOS.TRX_DATE} >= CDate (" + var_fecha_inicio + ") AND {XXVIA_JV_CONSECUTIVO_DE_CARGOS.TRX_DATE} < cdate(" + var_fecha_fin + ")"
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo "TVIA_TEST", "TVIA_TEST", "apps", "apps"
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte Consecutivo de Movimientos por Agente (cargos)"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            var_si = MsgBox("?Desea importar el reporte?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_consecutivo_tipo_movimientos_cargos.rpt")
               reporte.RecordSelectionFormula = "{XXVIA_JV_CONSECUTIVO_DE_CARGOS.TRX_DATE} >= CDate (" + var_fecha_inicio + ") AND {XXVIA_JV_CONSECUTIVO_DE_CARGOS.TRX_DATE} < cdate(" + var_fecha_fin + ")"
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo "TVIA_TEST", "TVIA_TEST", "apps", "apps"
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\reporte_consecutivo_movimientos_cargos" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
            Set reporte = appl.OpenReport(App.Path + "\rep_oracle_consecutivo_tipo_movimientos_abonos.rpt")
            reporte.RecordSelectionFormula = "{XXVIA_JV_CONSECUTIVO_DE_ABONOS.TRX_DATE}  >= CDate (" + var_fecha_inicio + ") AND {XXVIA_JV_CONSECUTIVO_DE_ABONOS.TRX_DATE} < cdate(" + var_fecha_fin + ")"
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo "TVIA_TEST", "TVIA_TEST", "apps", "apps"
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte Consecutivo de Movimientos por Agente (abonos)"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            var_si = MsgBox("?Desea importar el reporte?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_consecutivo_tipo_movimientos_abonos.rpt")
               reporte.RecordSelectionFormula = "{XXVIA_JV_CONSECUTIVO_DE_ABONOS.TRX_DATE}   >= CDate (" + var_fecha_inicio + ") AND {XXVIA_JV_CONSECUTIVO_DE_ABONOS.TRX_DATE} < cdate(" + var_fecha_fin + ")"
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo "TVIA_TEST", "TVIA_TEST", "apps", "apps"
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\reporte_consecutivo_movimientos_abonos" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
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
   If var_unidad_organizacional = "01" Or var_unidad_organizacional = "02" Or var_unidad_organizacional = "03" Or var_unidad_organizacional = "04" Or var_unidad_organizacional = "06" Then
      If UCase(parametros(0)) = "ADMCDINDUSTRIAL" Then
         Me.cmd_dolares.Visible = False
      Else
         Me.cmd_dolares.Visible = False
      End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    var_swpassword = False
    var_modifica_registro = False
    Call activa_forma(var_activa_forma_existencias_generales)
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

