VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmlista 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   8655
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm_lista 
      Height          =   2595
      Left            =   45
      TabIndex        =   0
      Top             =   -45
      Width           =   8565
      Begin MSComctlLib.ListView lv_lista 
         Height          =   2430
         Left            =   45
         TabIndex        =   1
         Top             =   120
         Width           =   8460
         _ExtentX        =   14923
         _ExtentY        =   4286
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
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   12347
         EndProperty
      End
   End
End
Attribute VB_Name = "frmlista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
   frm_lista.Visible = False
End Sub

Private Sub Form_Load()
   If var_lista_transportes = 1 Then
      rs.Open "select * from xxvia_Tb_Transportes where exportaciones = 1", cnnoracle_4, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!clave)
            list_item.SubItems(1) = IIf(IsNull(rs!nombre), "", rs!nombre)
            rs.MoveNext
      Wend
      rs.Close
   End If
   If lv_lista.ListItems.Count > 0 Then
      lv_lista.SetFocus
   End If
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         var_clave_lista_global = lv_lista.selectedItem
         var_nombre_lista_global = lv_lista.selectedItem.SubItems(1)
         frmlista.Visible = False
      End If
   End If
   If KeyAscii = 27 Then
      var_clave_lista_global = ""
      var_nombre_lista_global = ""
      frmlista.Visible = False
   End If
End Sub
