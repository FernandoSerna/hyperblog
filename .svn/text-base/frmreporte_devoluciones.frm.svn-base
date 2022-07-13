VERSION 5.00
Begin VB.Form frmreporte_devoluciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Reporte Devoluciones"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3390
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2970
      Picture         =   "frmreporte_devoluciones.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   0
      Picture         =   "frmreporte_devoluciones.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Caption         =   " Fecha "
      Height          =   645
      Left            =   75
      TabIndex        =   1
      Top             =   405
      Width           =   3270
      Begin VB.TextBox txt_fecha 
         Height          =   345
         Left            =   900
         TabIndex        =   2
         Top             =   195
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   30
      TabIndex        =   0
      Top             =   330
      Width           =   3285
   End
End
Attribute VB_Name = "frmreporte_devoluciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_mes As Integer

Private Sub cmd_imprimir_Click()
   Set reporte = appl.OpenReport(App.Path + "\rep_valuacion_devoluciones_nota_credito.rpt")
   reporte.RecordSelectionFormula = "{vw_devolucion_nota_credito.dtim_emo_fecha} = cdate('" + txt_fecha + "') and {vw_devolucion_nota_credito.vcha_mov_movimiento_id} = 'CA' and {vw_devolucion_nota_credito.vcha_emp_empresa_id} = '" + var_empresa + "'"
   frmvistasprevias.cr.ReportSource = reporte
   For ntablas = 1 To reporte.Database.Tables.Count
   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
   Next ntablas
        frmvistasprevias.cr.ViewReport
   frmvistasprevias.Caption = "Reporte de valuacion de devoluciones a detalle"
   frmvistasprevias.Show 1
   Set reporte = Nothing
   var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      Set reporte = appl.OpenReport(App.Path + "\rep_valuacion_devoluciones_nota_credito.rpt")
      reporte.RecordSelectionFormula = "{vw_devolucion_nota_credito.dtim_emo_fecha} = cdate('" + txt_fecha + "')  and {vw_devolucion_nota_credito.vcha_mov_movimiento_id} = 'CA'  and {vw_devolucion_nota_credito.vcha_emp_empresa_id} = '" + var_empresa + "'"
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      reporte.ExportOptions.FormatType = crEFTExcel80
      reporte.ExportOptions.DestinationType = crEDTDiskFile
      archivo = "c:\reportessid\Reporte_devoluciones_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
      reporte.ExportOptions.DiskFileName = archivo
      reporte.Export False
      Set reporte = Nothing
      MsgBox "Se a terminado de guardar el archivo " + archivo
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 4000
   txt_fecha = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_reporte_comisiones)
End Sub

Private Sub txt_fecha_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fecha) Then
         frmcalendario.mes = CDate(Me.txt_fecha)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      txt_fecha = var_fecha_general
   End If
End Sub

Private Sub txt_fecha_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub
