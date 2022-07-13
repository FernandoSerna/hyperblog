VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmservidores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Servidores"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lv_servidores 
      Height          =   3165
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   5583
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Clave"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Servidor"
         Object.Width           =   9525
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nombre"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Base de Datos"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmservidores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
        var_servidor = ""
        rs.Open "SELECT * FROM tb_Servidores ORDER BY VCHA_SER_NOMBRE", cnn_distribucion, adOpenDynamic, adLockOptimistic
        lv_servidores.ListItems.Clear
        While Not rs.EOF
              Set list_item = lv_servidores.ListItems.Add(, , rs!VCHA_sER_SERVIDOR_ID)
              list_item.SubItems(1) = IIf(IsNull(rs!VCHA_SER_NOMBRE), "", rs!VCHA_SER_NOMBRE)
              list_item.SubItems(2) = IIf(IsNull(rs!vcha_Ser_Servidor), "", rs!vcha_Ser_Servidor)
              list_item.SubItems(3) = IIf(IsNull(rs!vcha_Ser_base_datos), "", rs!vcha_Ser_base_datos)
              rs.MoveNext
        Wend
        rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_servidor = "" Then
      var_servidor = lv_servidores.selectedItem
   End If
End Sub

Private Sub lv_servidores_DblClick()
   var_servidor = lv_servidores.selectedItem
   Unload Me
End Sub

Private Sub lv_servidores_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      var_servidor = lv_servidores.selectedItem
      Unload Me
   End If
End Sub
