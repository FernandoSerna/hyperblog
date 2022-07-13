VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcatalogos_canales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de catálogo por canal de venta"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5955
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   5475
      Left            =   150
      TabIndex        =   15
      Top             =   1740
      Width           =   5655
      Begin MSComctlLib.ListView lv_catalogos 
         Height          =   5250
         Left            =   45
         TabIndex        =   16
         Top             =   150
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   9260
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave Canal de Venta"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Canal de Venta"
            Object.Width           =   4851
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Clave catálogo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Catálogo"
            Object.Width           =   4762
         EndProperty
      End
   End
   Begin VB.TextBox txt_articulo 
      Height          =   285
      Left            =   6270
      TabIndex        =   14
      Top             =   1995
      Width           =   825
   End
   Begin VB.Frame Frame2 
      Caption         =   " Catálogo "
      Height          =   615
      Left            =   150
      TabIndex        =   11
      Top             =   1080
      Width           =   5655
      Begin VB.ComboBox cmb_catalogos 
         Height          =   315
         Left            =   1260
         TabIndex        =   13
         Top             =   210
         Width           =   4320
      End
      Begin VB.TextBox txt_catalogo 
         Height          =   315
         Left            =   210
         MaxLength       =   50
         TabIndex        =   12
         Top             =   210
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmcatalogos_canales.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      Picture         =   "frmcatalogos_canales.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Guardar Alt + G"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   810
      Picture         =   "frmcatalogos_canales.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      Picture         =   "frmcatalogos_canales.frx":02D6
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1470
      Picture         =   "frmcatalogos_canales.frx":03D8
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5460
      Picture         =   "frmcatalogos_canales.frx":04DA
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2895
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   3
      Top             =   30
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   " Canal de Venta "
      Height          =   615
      Left            =   150
      TabIndex        =   0
      Top             =   420
      Width           =   5655
      Begin VB.TextBox txt_canal_venta 
         Height          =   315
         Left            =   225
         MaxLength       =   50
         TabIndex        =   2
         Top             =   210
         Width           =   1005
      End
      Begin VB.ComboBox cmb_canales_venta 
         Height          =   315
         Left            =   1260
         TabIndex        =   1
         Top             =   210
         Width           =   4320
      End
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   10
      Top             =   285
      Width           =   5655
   End
End
Attribute VB_Name = "frmcatalogos_canales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub pro_llena_listview1()
Dim list_item As ListItem
   numero_items_rutas = 0
   rs.Open "select * from vw_catalogos_canales_venta where vcha_art_articulo_id = '" + txt_Articulo + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = lv_catalogos.ListItems.Add(, , rs!vcha_can_canal_venta_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_can_nombre), "", rs!vcha_can_nombre)
            list_item.SubItems(2) = IIf(IsNull(rs!vcha_Cat_catalogo_id), "", rs!vcha_Cat_catalogo_id)
            list_item.SubItems(3) = IIf(IsNull(rs!vcha_cat_nombre), "", rs!vcha_cat_nombre)
            rs.MoveNext:
            numero_items_rutas = numero_items_rutas + 1
       Wend
       txt_canal_venta = lv_catalogos.selectedItem
       cmb_canales_venta = lv_catalogos.selectedItem.SubItems(1)
       txt_catalogo = lv_catalogos.selectedItem.SubItems(2)
       cmb_catalogos = lv_catalogos.selectedItem.SubItems(3)
   End If
   rs.Close
End Sub

Private Sub cmb_canales_venta_Click()
 txt_canal_venta = Obtener_llave(cnn, rs, "TB_CANALESVENTAS", "VCHA_CAN_NOMBRE", cmb_canales_venta, 0, "T")
End Sub

Private Sub cmb_canales_venta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_catalogo.SetFocus
   End If
End Sub

Private Sub cmb_catalogos_Click()
   txt_catalogo = Obtener_llave(cnn, rs, "TB_CATALOGOS", "VCHA_CAT_NOMBRE", cmb_catalogos, 0, "T")
End Sub

Private Sub cmd_eliminar_Click()
   Dim si As Integer
   si = MsgBox("¿Deseas eliminar el registro?", vbYesNo, "ATENCION")
   If si = 6 Then
      rs.Open "delete from tb_catalogos_canal_Venta where vcha_cat_Catalogo_id = '" + txt_catalogo + "' and vcha_can_canal_Venta_id = '" + txt_canal_venta + "' and vcha_Art_articulo_id = '" + txt_Articulo + "'", cnn, adOpenDynamic, adLockOptimistic
      lv_catalogos.ListItems.Remove (lv_catalogos.selectedItem.Index)
   End If
End Sub

Private Sub cmd_guardar_Click()
   If Trim(txt_catalogo) <> "" And Trim(txt_canal_venta) <> "" Then
      rs.Open "select * from TB_CATALOGOS_CANAL_VENTA where vcha_cat_catalogo_id = '" + txt_catalogo + "' and vcha_can_canal_Venta_id = '" + txt_canal_venta + "' and vcha_art_articulo_id = '" + txt_Articulo + "'", cnn, adOpenDynamic, adLockOptimistic
      If rs.EOF Then
         rsaux.Open "insert into TB_CATALOGOS_CANAL_VENTA (vcha_cat_catalogo_id, vcha_can_canal_venta_id, vcha_art_articulo_id) values ('" + txt_catalogo + "', '" + txt_canal_venta + "', '" + txt_Articulo + "')"
         Dim list_item As ListItem
         Set list_item = lv_catalogos.ListItems.Add(, , txt_canal_venta)
         list_item.SubItems(1) = cmb_canales_venta
         list_item.SubItems(2) = txt_catalogo
         list_item.SubItems(3) = cmb_catalogos
         list_item.EnsureVisible
         list_item.Selected = True
      Else
         MsgBox "El catálogo ya habia sido asignado al canal de venta con anterioridad", vbOKOnly, "ATENCION"
      End If
      rs.Close
      txt_canal_venta.Enabled = False
      cmb_canales_venta.Enabled = False
      txt_catalogo.Enabled = False
      cmb_catalogos.Enabled = False
   End If
End Sub

Private Sub cmd_nuevo_Click()
   txt_canal_venta = ""
   cmb_canales_venta = ""
   txt_catalogo = ""
   cmb_catalogos = ""
   txt_canal_venta.Enabled = True
   cmb_canales_venta.Enabled = True
   txt_catalogo.Enabled = True
   cmb_catalogos.Enabled = True
   txt_canal_venta.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub


Private Sub Form_Load()
   var_cadena_seguridad = ""
   txt_canal_venta.Enabled = False
   cmb_canales_venta.Enabled = False
   txt_catalogo.Enabled = False
   cmb_catalogos.Enabled = False
   txt_Articulo = frmarticulos2.txt_codigo
   rs.Open "select * from tb_canalesventas order by vcha_Can_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
   Call RecsetToCombo(cmb_canales_venta.hwnd, rs, 1)
   rs.Close
   rs.Open "select * from tb_catalogos order by vcha_cat_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
   Call RecsetToCombo(cmb_catalogos.hwnd, rs, 1)
   rs.Close
   Call pro_llena_listview1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_catalogos_canales)
End Sub

Private Sub lv_catalogos_ItemClick(ByVal Item As MSComctlLib.ListItem)
   txt_canal_venta = lv_catalogos.selectedItem
   cmb_canales_venta = lv_catalogos.selectedItem.SubItems(1)
   txt_catalogo = lv_catalogos.selectedItem.SubItems(2)
   cmb_catalogos = lv_catalogos.selectedItem.SubItems(3)
End Sub

Private Sub txt_canal_venta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      cmb_canales_venta.SetFocus
   End If
End Sub

Private Sub txt_canal_venta_LostFocus()
   If Trim(txt_canal_venta) <> "" Then
      rs.Open "select * from tb_canalesventas where vcha_can_Canal_Venta_id = '" + txt_canal_venta + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         cmb_canales_venta = rs!vcha_can_nombre
      Else
         MsgBox "Clave de canal de venta incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   End If
End Sub

Private Sub txt_catalogo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      cmb_catalogos.SetFocus
   End If
End Sub

Private Sub txt_catalogo_LostFocus()
   If Trim(txt_catalogo) <> "" Then
      rs.Open "select * from tb_catalogos where vcha_cat_catalogo_id = '" + txt_catalogo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         cmb_catalogos = rs!vcha_cat_nombre
      Else
         MsgBox "Clave de catálogo incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   End If
End Sub
