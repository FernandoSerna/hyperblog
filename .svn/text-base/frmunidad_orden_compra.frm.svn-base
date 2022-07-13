VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmunidad_orden_compra 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Organización"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   1
         Top             =   480
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   3228
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
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Unidad"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         Caption         =   " Organización"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   2
         Top             =   120
         Width           =   5610
      End
   End
End
Attribute VB_Name = "frmunidad_orden_compra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Dim list_item As ListItem
   
   rs.Open "select * from TB_UNIDADES_ORDEN_COMPRA order by VCHA_ORC_NOMBRE", cnn, adOpenDynamic, adLockOptimistic
   numero_items_tallas = 0
   While Not rs.EOF
      Set list_item = lv_lista.ListItems.Add(, , rs(0).Value)
      list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
      rs.MoveNext:
      numero_items_tallas = numero_items_tallas + 1
    Wend
    rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
    var_unidad_OC = lv_lista.selectedItem
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 27 Then
      Unload Me
   End If
End Sub
