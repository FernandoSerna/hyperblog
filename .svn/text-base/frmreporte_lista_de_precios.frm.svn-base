VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmreporte_lista_de_precios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de listas de precios"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   1800
      Left            =   15
      TabIndex        =   6
      Top             =   -30
      Width           =   6810
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1290
         Left            =   45
         TabIndex        =   7
         Top             =   450
         Width           =   6720
         _ExtentX        =   11853
         _ExtentY        =   2275
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
            Object.Width           =   10583
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   8
         Top             =   135
         Width           =   6720
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6390
      Picture         =   "frmreporte_lista_de_precios.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmreporte_lista_de_precios.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   75
      TabIndex        =   5
      Top             =   270
      Width           =   6645
   End
   Begin VB.Frame Frame1 
      Caption         =   " Lista de precios "
      Height          =   1095
      Left            =   105
      TabIndex        =   4
      Top             =   540
      Width           =   6600
      Begin VB.TextBox txt_nombre 
         Height          =   350
         Left            =   960
         TabIndex        =   1
         Top             =   435
         Width           =   5520
      End
      Begin VB.TextBox txt_clave 
         Height          =   350
         Left            =   150
         TabIndex        =   0
         Top             =   435
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmreporte_lista_de_precios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
   If Me.txt_clave <> "" Then
      Set reporte = appl.OpenReport(App.Path + "\REP_lista_precios.rpt")
      reporte.RecordSelectionFormula = "{VW_LISTA_PRECIOS.VCHA_LIS_LISTA_ID} = '" + Me.txt_clave + "'"
      frmvistasprevias.cr.ReportSource = reporte
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      frmvistasprevias.cr.ViewReport
      frmvistasprevias.Caption = "Reporte de Lista de Precios"
      frmvistasprevias.Show 1
      Set reporte = Nothing
      var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         Set reporte = appl.OpenReport(App.Path + "\REP_lista_precios.rpt")
         reporte.RecordSelectionFormula = "{VW_LISTA_PRECIOS.VCHA_LIS_LISTA_ID} = '" + Me.txt_clave + "'"
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         reporte.ExportOptions.FormatType = crEFTExcel80
         reporte.ExportOptions.DestinationType = crEDTDiskFile
         archivo = "c:\reportessid\Reporte_lista_precios_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
         reporte.ExportOptions.DiskFileName = archivo
         reporte.Export False
         Set reporte = Nothing
         MsgBox "Se a terminado de guardar el archivo " + archivo
      End If
   Else
      MsgBox "No se a selecionado una lista de precios", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Form_Load()
   Me.frm_lista.Visible = False
   Top = 3000
   Left = 3000
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frm_lista.Visible = False
   End If
   If KeyAscii = 13 Then
      Me.txt_clave = Me.lv_lista.selectedItem
      Me.txt_nombre = Me.lv_lista.selectedItem.SubItems(1)
      Me.lv_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub txt_clave_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_listadeprecios order by vcha_lis_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_LIS_LISTA_iD)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_lIS_NOMBRE), "", rs!VCHA_lIS_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "LISTA DE PRECIOS"
      var_tipo_lista = 21
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 5400
      Else
         lv_lista.ColumnHeaders(2).Width = 5600
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_clave_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
    If KeyAscii = 27 Then
       Unload Me
    End If
End Sub

Private Sub txt_clave_LostFocus()
   If Me.txt_clave <> "" Then
      rs.Open "select * from TB_LISTADEPRECIOS where vcha_lis_lista_id = '" + Me.txt_clave + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre = IIf(IsNull(rs!VCHA_lIS_NOMBRE), "", rs!VCHA_lIS_NOMBRE)
      Else
         MsgBox "Clave de lista incorrecta", vbOKOnly, "ATENCION"
         Me.txt_clave = ""
         Me.txt_nombre = ""
      End If
      rs.Close
   Else
      Me.txt_nombre = ""
   End If
End Sub

Private Sub txt_nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   Else
      If KeyAscii = 27 Then
         Unload Me
      Else
         KeyAscii = 0
      End If
   End If
End Sub
