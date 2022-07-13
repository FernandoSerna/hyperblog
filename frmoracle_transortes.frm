VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_transortes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transportes"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frm_lista 
      Height          =   3000
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7845
      Begin VB.TextBox txt_transporte 
         Height          =   405
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   7575
      End
      Begin MSComctlLib.ListView lv_lista 
         Height          =   2190
         Left            =   45
         TabIndex        =   0
         Top             =   720
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   3863
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Volumen"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "frmoracle_transortes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
   If Me.lv_lista.ListItems.Count > 0 Then
      var_transporte_global = Me.lv_lista.selectedItem
   End If
End Sub

Private Sub Form_Load()
      var_ventana = 2
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_oracle_transportes order by nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!CLAVE)
            list_item.SubItems(1) = IIf(IsNull(rs!nombre), "", rs!nombre)
            list_item.SubItems(2) = IIf(IsNull(rs!VOLUMEN), "", rs!VOLUMEN)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TRANSPORTES"
      VAR_TIPO_LISTA = 100
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      'lv_lista.SetFocus
End Sub

Private Sub lv_lista_GotFocus()
   If Me.lv_lista.ListItems.Count > 0 Then
      var_transporte_global = Me.lv_lista.selectedItem
   End If
End Sub

Private Sub lv_lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
   If Me.lv_lista.ListItems.Count > 0 Then
      var_transporte_global = Me.lv_lista.selectedItem
   End If

End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If Me.lv_lista.ListItems.Count > 0 Then
      var_transporte_global = Me.lv_lista.selectedItem
      Unload Me
   End If
End Sub

Private Sub txt_transporte_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 40 Then
      If Me.lv_lista.ListItems.Count > 0 Then
         Me.lv_lista.SetFocus
      End If
   End If
End Sub

Private Sub txt_transporte_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       If Trim(Me.txt_transporte) <> "" Then
          Me.lv_lista.ListItems.Clear
          rs.Open "select * from tb_oracle_transportes where nombre like '%" + Me.txt_transporte + "%'", cnn, adOpenDynamic, adLockOptimistic
          While Not rs.EOF
                Set list_item = lv_lista.ListItems.Add(, , rs!CLAVE)
                list_item.SubItems(1) = IIf(IsNull(rs!nombre), "", rs!nombre)
                list_item.SubItems(2) = IIf(IsNull(rs!VOLUMEN), "", rs!VOLUMEN)
                rs.MoveNext
          Wend
          rs.Close
       Else
          Me.lv_lista.ListItems.Clear
          rs.Open "select * from tb_oracle_transportes order by nombre", cnn, adOpenDynamic, adLockOptimistic
          While Not rs.EOF
                Set list_item = lv_lista.ListItems.Add(, , rs!CLAVE)
                list_item.SubItems(1) = IIf(IsNull(rs!nombre), "", rs!nombre)
                list_item.SubItems(2) = IIf(IsNull(rs!VOLUMEN), "", rs!VOLUMEN)
                rs.MoveNext
          Wend
          rs.Close
       End If
   End If
   
End Sub
