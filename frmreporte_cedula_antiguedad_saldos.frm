VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmreporte_cedula_antiguedad_saldos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cedula de Saldos"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   1710
      Left            =   0
      TabIndex        =   6
      Top             =   -120
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1200
         Left            =   45
         TabIndex        =   7
         Top             =   480
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   2117
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7937
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   8
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Agente "
      Height          =   810
      Left            =   30
      TabIndex        =   5
      Top             =   555
      Width           =   5655
      Begin VB.TextBox txt_nombre_agente 
         Height          =   345
         Left            =   1365
         TabIndex        =   1
         Top             =   300
         Width           =   4170
      End
      Begin VB.TextBox txt_agente 
         Height          =   345
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   1230
      End
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   45
      TabIndex        =   4
      Top             =   360
      Width           =   5685
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      Picture         =   "frmreporte_cedula_antiguedad_saldos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5325
      Picture         =   "frmreporte_cedula_antiguedad_saldos.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
End
Attribute VB_Name = "frmreporte_cedula_antiguedad_saldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
   If var_clave_usuario_global = "U0000000047" Then
      Set reporte = appl.OpenReport(App.Path + "\rep_antiguedad_saldos_direcciones.rpt")
   Else
      Set reporte = appl.OpenReport(App.Path + "\rep_antiguedad_saldos_direcciones.rpt")
      'Set reporte = appl.OpenReport(App.Path + "\rep_antiguedad_saldos_sin_direcciones.rpt")
   End If
   reporte.RecordSelectionFormula = "{VW_ANTIGUEDAD_SALDOS_DIRECCIONES.VCHA_AGE_AGENTE_ID} = '" + txt_agente + "'"
   frmvistasprevias.cr.ReportSource = reporte
   For ntablas = 1 To reporte.Database.Tables.Count
   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
   Next ntablas
        frmvistasprevias.cr.ViewReport
   frmvistasprevias.Caption = "Reporte de Cedula de Saldos"
   frmvistasprevias.Show 1
   Set reporte = Nothing
   var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      If var_clave_usuario_global = "U0000000047" Then
         Set reporte = appl.OpenReport(App.Path + "\rep_antiguedad_saldos_direcciones.rpt")
      Else
         Set reporte = appl.OpenReport(App.Path + "\rep_antiguedad_saldos_sin_direcciones.rpt")
      End If
      reporte.RecordSelectionFormula = "{VW_ANTIGUEDAD_SALDOS_DIRECCIONES.VCHA_AGE_AGENTE_ID} = '" + txt_agente + "'"
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      reporte.ExportOptions.FormatType = crEFTExcel80
      reporte.ExportOptions.DestinationType = crEDTDiskFile
      archivo = "c:\reportessid\Reporte_cedula_saldos_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
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
   Top = 2500
   Left = 2500
   frm_lista.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_reporte_comisiones)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         txt_agente = lv_lista.selectedItem
         txt_nombre_agente = lv_lista.selectedItem.SubItems(1)
      Else
         txt_agente = ""
         txt_nombre_agente = ""
      End If
      txt_agente.SetFocus
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
    frm_lista.Visible = False
End Sub

Private Sub txt_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_AGENTES where vcha_Emp_empresa_id = '" + var_empresa + "' order by vcha_AGE_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_agente_KeyPress(KeyAscii As Integer)
     Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_agente_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_agente) <> "" Then
      txt_agente = UCase(txt_agente)
      rs.Open "select * from tb_agentes where vcha_age_agente_id = '" + txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_agente = rs!VCHA_AGE_NOMBRE
         var_agente = rs!VCHA_AGE_AGENTE_ID
         rs.Close
      Else
         rs.Close
         MsgBox "Clave de agente incorrecta", vbOKOnly, "ATENCION"
         txt_agente = ""
         txt_nombre_agente = ""
      End If
   End If
End Sub

Private Sub txt_nombre_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_AGENTES  where vcha_emp_empresa_id = '" + var_empresa + "' order by vcha_AGE_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      cmd_imprimir.SetFocus
   End If
End Sub
