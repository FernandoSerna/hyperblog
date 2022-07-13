VERSION 5.00
Begin VB.Form frmreporte_surtido 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Surtido"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   2310
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Pedido "
      Height          =   630
      Left            =   60
      TabIndex        =   4
      Top             =   495
      Width           =   2145
      Begin VB.TextBox txt_pedido 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   210
         Width           =   1860
      End
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   45
      TabIndex        =   2
      Top             =   360
      Width           =   2235
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      Picture         =   "frmreporte_surtido.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1875
      Picture         =   "frmreporte_surtido.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
End
Attribute VB_Name = "frmreporte_surtido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report


Private Sub cmd_imprimir_Click()
   Dim var_cadena As String
   var_cadena = ""
   Set reporte = appl.OpenReport(App.Path + "\rep_cantidad_Surtido.rpt")
   reporte.RecordSelectionFormula = "{VW_REPORTE_SURTIDO.inte_ped_numero} = " + txt_pedido
   frmvistasprevias.cr.ReportSource = reporte
   For ntablas = 1 To reporte.Database.Tables.Count
       reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
   Next ntablas
   frmvistasprevias.cr.ViewReport
   frmvistasprevias.Caption = "Reporte de Surtido de Pedidos"
   frmvistasprevias.Show 1
   Set reporte = Nothing
   var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      Set reporte = appl.OpenReport(App.Path + "\rep_cantidad_Surtido.rpt")
      reporte.RecordSelectionFormula = "{VW_REPORTE_SURTIDO.inte_ped_numero} = " + txt_pedido
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      reporte.ExportOptions.FormatType = crEFTExcel80
      reporte.ExportOptions.DestinationType = crEDTDiskFile
      archivo = "c:\reportessid\Surtido_pedido_" + txt_pedido + "_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
      reporte.ExportOptions.DiskFileName = archivo
      reporte.Export False
      Set reporte = Nothing
      MsgBox "Se a terminado de guardar el archivo " + archivo
  End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 5000
   txt_inicio = Date
   txt_fin = Date
'   txt_pedido.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_despliega_menu = True Then
      var_swpassword = False
      var_modifica_registro = False
   End If
   Call activa_forma(var_activa_forma_reporte_comisiones)
End Sub

Private Sub txt_agente_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_agente_LostFocus()
   If Trim(txt_agente) <> "" Then
      rs.Open "select * from tb_agentes where vcha_age_agente_id = '" + txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
      Else
         MsgBox "Clave de agente incorrecto", vbOKOnly, "ATENCION"
         txt_agente = ""
         txt_nombre = ""
      End If
      rs.Close
   Else
      txt_nombre = ""
   End If
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(txt_fin) Then
         frmcalendario.mes.Value = CDate(txt_fin)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_fin = var_fecha_general
   End If
End Sub

Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(txt_inicio) Then
         frmcalendario.mes.Value = CDate(txt_inicio)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_inicio = var_fecha_general
   End If
End Sub

Private Sub txt_inicio_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_pedido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   End If
End Sub
