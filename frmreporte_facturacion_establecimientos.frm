VERSION 5.00
Begin VB.Form frmreporte_facturacion_establecimientos 
   Caption         =   "Form1"
   ClientHeight    =   1455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   720
      Left            =   60
      TabIndex        =   3
      Top             =   510
      Width           =   4335
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   2730
         TabIndex        =   5
         Top             =   255
         Width           =   1095
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   855
         TabIndex        =   4
         Top             =   255
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2415
         TabIndex        =   7
         Top             =   315
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   405
         TabIndex        =   6
         Top             =   315
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   30
      TabIndex        =   2
      Top             =   345
      Width           =   4440
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      Picture         =   "frmreporte_facturacion_establecimientos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4050
      Picture         =   "frmreporte_facturacion_establecimientos.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
End
Attribute VB_Name = "frmreporte_facturacion_establecimientos"
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
   txt_almacen = "-"
   If txt_almacen <> "" Then
      If IsDate(txt_inicio) Then
         If IsDate(txt_fin) Then
            If CDate(txt_inicio) <= CDate(txt_fin) Then
               cnn.BeginTrans
               rs.Open "select max(inte_TEM_consecutivo) from tb_temp_kardex", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
               Else
                  var_consecutivo = 0
               End If
               var_consecutivo = var_consecutivo + 1
               rs.Close
               rs.Open "Insert into tb_temp_kardex (INTE_tem_CONSECUTIVO) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
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
               rs.Open "EXEC SP_REPORTE_FACTURACION " + CStr(var_consecutivo) + ",'" + txt_almacen + "'," + var_fecha_inicio + ", " + var_fecha_fin, cnn, adOpenDynamic, adLockOptimistic
               rs.Open "delete from TB_TEMP_KARDEX where INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND vcha_alm_almacen_id IS NULL", cnn, adOpenDynamic, adLockOptimistic
               Set reporte = appl.OpenReport(App.Path + "\REP_REPORTE_FACTURACION_ESTAMPADOS.rpt")
               reporte.RecordSelectionFormula = "{VW_TEMP_KARDEX.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND {VW_TEMP_KARDEX.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "'"
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Reporte de movimientos"
               frmvistasprevias.Show 1
               Set reporte = Nothing
               
               var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  Set reporte = appl.OpenReport(App.Path + "\REP_REPORTE_FACTURACION_ESTAMPADOS.rpt")
                  reporte.RecordSelectionFormula = "{VW_TEMP_KARDEX.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND {VW_TEMP_KARDEX.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "'"
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  reporte.ExportOptions.FormatType = crEFTExcel80
                  reporte.ExportOptions.DestinationType = crEDTDiskFile
                  archivo = "c:\reportessid\Reporte_facturacion_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                  reporte.ExportOptions.DiskFileName = archivo
                  reporte.Export False
                  Set reporte = Nothing
                  MsgBox "Se a terminado de guardar el archivo " + archivo
               End If
               
               
               
               rs.Open "delete from TB_TEMP_KARDEX where INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            Else
               MsgBox "La fecha de inicio debe de ser mayor", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Fecha Final Incorrecta", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha de Inicio Incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Clave de almacén incorrecto", vbOKOnly, "ATENCION"
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




