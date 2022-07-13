VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmlistado_almacenes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Almacenes"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6750
   Begin MSComctlLib.ListView lv_almacenes 
      Height          =   3345
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   5900
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Clave"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "                               Almacén"
         Object.Width           =   11201
      EndProperty
   End
End
Attribute VB_Name = "frmlistado_almacenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Top = 2000
   Left = 2500
   Dim list_item As ListItem
   If var_empresa = "28" Then
      rs.Open "select * from tb_almacenes where vcha_Emp_Empresa_id = '" + var_empresa + "' order by vcha_alm_nombre", cnn, adOpenDynamic, adLockOptimistic
   Else
      rs.Open "select * from tb_almacenes order by vcha_alm_nombre", cnn, adOpenDynamic, adLockOptimistic
   End If
   While Not rs.EOF
      Set list_item = lv_almacenes.ListItems.Add(, , rs!VCHA_ALM_ALMACEN_ID)
      list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
      rs.MoveNext:
    Wend
    rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_listado_almacenes)
End Sub

Private Sub lv_almacenes_DblClick()
   frmubicaciones_almacen.txt_almacen = lv_almacenes.selectedItem
   frmubicaciones_almacen.txt_nombre_almacen = lv_almacenes.selectedItem.SubItems(1)
End Sub

Private Sub lv_almacenes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      frmubicaciones_almacen.txt_almacen = lv_almacenes.selectedItem
      frmubicaciones_almacen.txt_nombre_almacen = lv_almacenes.selectedItem.SubItems(1)
   End If
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub
