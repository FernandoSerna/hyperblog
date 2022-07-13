VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmaplicacion_descuentos_canal_venta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aplicación de pagos a catálogos segun el canal de venta."
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   765
      TabIndex        =   18
      Top             =   390
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   30
         TabIndex        =   19
         Top             =   495
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
         TabIndex        =   20
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5325
      Left            =   120
      TabIndex        =   16
      Top             =   1875
      Width           =   6555
      Begin MSComctlLib.ListView lv_descuentos 
         Height          =   5100
         Left            =   45
         TabIndex        =   17
         Top             =   165
         Width           =   6420
         _ExtentX        =   11324
         _ExtentY        =   8996
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Catálogo"
            Object.Width           =   6967
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descuento 1"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descuento 2"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Clave Catalogo"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6300
      Picture         =   "frmaplicacion_descuentos_canal_venta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1410
      Picture         =   "frmaplicacion_descuentos_canal_venta.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1080
      Picture         =   "frmaplicacion_descuentos_canal_venta.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   765
      Picture         =   "frmaplicacion_descuentos_canal_venta.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   435
      Picture         =   "frmaplicacion_descuentos_canal_venta.frx":0910
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Guardar Alt + G"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmaplicacion_descuentos_canal_venta.frx":0A12
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Caption         =   " Aplica Descuentos"
      Height          =   1440
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   6540
      Begin VB.CheckBox chk_descuento_2 
         Caption         =   "Aplica descuento 2"
         Height          =   195
         Left            =   3570
         TabIndex        =   13
         Top             =   1065
         Width           =   1725
      End
      Begin VB.TextBox txt_catalogo 
         Height          =   315
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   9
         Top             =   615
         Width           =   1035
      End
      Begin VB.TextBox txt_nombre_catalogo 
         Height          =   315
         Left            =   2370
         MaxLength       =   50
         TabIndex        =   11
         Top             =   615
         Width           =   4095
      End
      Begin VB.CheckBox chk_descuento_1 
         Caption         =   "Aplica descuento 1"
         Height          =   195
         Left            =   1320
         TabIndex        =   12
         Top             =   1065
         Width           =   1725
      End
      Begin VB.TextBox txt_nombre_canal_venta 
         Height          =   315
         Left            =   2370
         MaxLength       =   50
         TabIndex        =   8
         Top             =   270
         Width           =   4095
      End
      Begin VB.TextBox txt_canal_venta 
         Height          =   315
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   7
         Top             =   270
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "Catálogo:"
         Height          =   285
         Left            =   90
         TabIndex        =   15
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Canal de Venta:"
         Height          =   285
         Left            =   75
         TabIndex        =   10
         Top             =   285
         Width           =   1275
      End
   End
   Begin VB.Frame Frame5 
      Height          =   120
      Left            =   90
      TabIndex        =   14
      Top             =   285
      Width           =   6555
   End
End
Attribute VB_Name = "frmaplicacion_descuentos_canal_venta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tipo_lista As Integer
Dim var_nuevo As Boolean
Dim var_cambios As Boolean

Private Sub cmd_deshacer_Click()
   var_nuevo = False
End Sub

Private Sub cmd_eliminar_Click()
   var_si = MsgBox("¿Desea eliminar el registro?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      rs.Open "DELETE FROM TB_APLICACION_DESCUENTOS_CANAL_VENTA_CATALOGOS where VCHA_CAN_CANAL_VENTA_ID = '" + Me.txt_canal_venta + "' and VCHA_CAT_CATALOGO_ID = '" + Me.txt_catalogo + "'", cnn, adOpenDynamic, adLockOptimistic
      lv_descuentos.ListItems.Remove (lv_descuentos.selectedItem.Index)
      lv_descuentos.SetFocus
   End If
End Sub

Private Sub cmd_guardar_Click()
   If var_nuevo = True Then
      rs.Open "select * from TB_APLICACION_DESCUENTOS_CANAL_VENTA_CATALOGOS where VCHA_CAN_CANAL_VENTA_ID = '" + Me.txt_canal_venta + "' and VCHA_CAT_CATALOGO_ID = '" + Me.txt_catalogo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         rsaux.Open "update TB_APLICACION_DESCUENTOS_CANAL_VENTA_CATALOGOS set INTE_APL_APLICA_DESCUENTO_1 = " + CStr(chk_descuento_1.Value) + ", INTE_APL_APLICA_DESCUENTO_2 = " + CStr(Me.chk_descuento_2) + " where VCHA_CAN_CANAL_VENTA_ID = '" + Me.txt_canal_venta + "' and VCHA_CAT_CATALOGO_ID = '" + Me.txt_catalogo + "'"
         lv_descuentos.selectedItem.SubItems(1) = chk_descuento_1
         lv_descuentos.selectedItem.SubItems(2) = chk_descuento_2
      Else
         rsaux.Open "insert TB_APLICACION_DESCUENTOS_CANAL_VENTA_CATALOGOS (VCHA_CAN_CANAL_VENTA_ID, VCHA_CAT_CATALOGO_ID, INTE_APL_APLICA_DESCUENTO_1, INTE_APL_APLICA_DESCUENTO_2) VALUES ('" + Me.txt_canal_venta + "', '" + Me.txt_catalogo + "', " + CStr(Me.chk_descuento_1) + ", " + CStr(Me.chk_descuento_2) + ")", cnn, adOpenDynamic, adLockOptimistic
         Dim list_item As ListItem
         Set list_item = lv_descuentos.ListItems.Add(, , txt_nombre_catalogo)
         list_item.SubItems(1) = chk_descuento_1
         list_item.SubItems(2) = chk_descuento_2
         list_item.SubItems(3) = txt_catalogo
    End If
      rs.Close
   Else
      rsaux.Open "update TB_APLICACION_DESCUENTOS_CANAL_VENTA_CATALOGOS set INTE_APL_APLICA_DESCUENTO_1 = " + CStr(chk_descuento_1.Value) + ", INTE_APL_APLICA_DESCUENTO_2 = " + CStr(Me.chk_descuento_2) + " where VCHA_CAN_CANAL_VENTA_ID = '" + Me.txt_canal_venta + "' and VCHA_CAT_CATALOGO_ID = '" + Me.txt_catalogo + "'"
      lv_descuentos.selectedItem.SubItems(1) = chk_descuento_1
      lv_descuentos.selectedItem.SubItems(2) = chk_descuento_2
   End If
   var_nuevo = False
End Sub

Private Sub cmd_nuevo_Click()
   Call pro_limpiatextos(Me)
   Me.txt_canal_venta.SetFocus
   var_nuevo = True
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 2500
   frm_lista.Visible = False
   var_nuevo = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_aplicacion_descuentos_canal_venta)
End Sub

Private Sub lv_descuentos_GotFocus()
   If lv_descuentos.ListItems.Count > 0 Then
      Me.txt_nombre_catalogo = lv_descuentos.selectedItem
      Me.chk_descuento_1 = lv_descuentos.selectedItem.SubItems(1)
      Me.chk_descuento_2 = lv_descuentos.selectedItem.SubItems(2)
      Me.txt_catalogo = lv_descuentos.selectedItem.SubItems(3)
   Else
      Me.txt_nombre_catalogo = ""
      Me.chk_descuento_1 = 0
      Me.chk_descuento_2 = 0
      Me.txt_catalogo = ""
   End If
End Sub

Private Sub lv_descuentos_ItemClick(ByVal Item As MSComctlLib.ListItem)
   If lv_descuentos.ListItems.Count > 0 Then
      Me.txt_nombre_catalogo = lv_descuentos.selectedItem
      Me.chk_descuento_1 = lv_descuentos.selectedItem.SubItems(1)
      Me.chk_descuento_2 = lv_descuentos.selectedItem.SubItems(2)
      Me.txt_catalogo = lv_descuentos.selectedItem.SubItems(3)
   End If
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_tipo_lista = 1 Then
         Me.txt_canal_venta = lv_lista.selectedItem
         Me.txt_nombre_canal_venta = lv_lista.selectedItem.SubItems(1)
         Me.txt_canal_venta.SetFocus
      End If
      If var_tipo_lista = 2 Then
         Me.txt_catalogo = lv_lista.selectedItem
         Me.txt_nombre_catalogo = lv_lista.selectedItem.SubItems(1)
         Me.txt_catalogo.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub txt_canal_venta_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_CANALESVENTAS order by VCHA_CAN_NOMBRE", cnn, adOpenDynamic, adLockOptimistic
      
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_can_canal_venta_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_can_nombre), "", rs!vcha_can_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CANALES DE VENTA"
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

Private Sub txt_canal_venta_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_canal_venta_LostFocus()
   If Trim(txt_canal_venta) <> "" Then
      rs.Open "select * from tb_canalesventas where vcha_can_canal_Venta_id = '" + txt_canal_venta + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_canal_venta = IIf(IsNull(rs!vcha_can_nombre), "", rs!vcha_can_nombre)
         lv_descuentos.ListItems.Clear
         Dim list_item As ListItem
         rsaux.Open "select * from vw_APLICACION_DESCUENTOS_CANAL_VENTA_CATALOGOS where vcha_can_canal_Venta_id = '" + Me.txt_canal_venta + "' order by vcha_cat_nombre", cnn, adOpenDynamic, adLockOptimistic
         numero_items_clases = 0
         While Not rsaux.EOF
               Set list_item = lv_descuentos.ListItems.Add(, , rsaux!vcha_cat_nombre)
               list_item.SubItems(1) = IIf(IsNull(rsaux!INTE_APL_APLICA_DESCUENTO_1), 0, rsaux!INTE_APL_APLICA_DESCUENTO_1)
               list_item.SubItems(2) = IIf(IsNull(rsaux!INTE_APL_APLICA_DESCUENTO_2), 0, rsaux!INTE_APL_APLICA_DESCUENTO_2)
               list_item.SubItems(3) = IIf(IsNull(rsaux!vcha_Cat_catalogo_id), "", rsaux!vcha_Cat_catalogo_id)
               rsaux.MoveNext:
               numero_items_clases = numero_items_clases + 1
         Wend
         rsaux.Close
      Else
         txt_canal_venta = ""
         txt_nombre_canal_venta = ""
         lv_descuentos.ListItems.Clear
         MsgBox "Clave de canal de venta incorrecto", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      lv_descuentos.ListItems.Clear
      txt_nombre_canal_venta = ""
      txt_catalogo = ""
      txt_nombre_catalogo = ""
      chk_descuento_1 = 0
      chk_descuento_2 = 0
   End If
End Sub

Private Sub txt_catalogo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_CATALOGOS order by VCHA_CAT_NOMBRE", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_Cat_catalogo_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_cat_nombre), "", rs!vcha_cat_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CATALOGOS"
      var_tipo_lista = 2
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

Private Sub txt_catalogo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_catalogo_LostFocus()
   If Trim(txt_catalogo) <> "" Then
      rs.Open "SELECT * FROM TB_CATALOGOS WHERE VCHA_CAT_CATALOGO_ID = '" + txt_catalogo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_catalogo = IIf(IsNull(rs!vcha_cat_nombre), "", rs!vcha_cat_nombre)
      Else
         txt_nombre_catalogo = ""
         txt_catalogo = ""
         MsgBox "Clave de catálogo incorrecto", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      txt_nombre_catalogo = ""
   End If
End Sub

Private Sub txt_nombre_canal_venta_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_CANALESVENTAS order by VCHA_CAN_NOMBRE", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_can_canal_venta_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_can_nombre), "", rs!vcha_can_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CANALES DE VENTA"
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

Private Sub txt_nombre_canal_venta_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_catalogo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_CATALOGOS order by VCHA_CAT_NOMBRE", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_Cat_catalogo_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_cat_nombre), "", rs!vcha_cat_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CATALOGOS"
      var_tipo_lista = 2
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

Private Sub txt_nombre_catalogo_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
       KeyAscii = 0
    End If
    Call pro_enfoque(KeyAscii)
End Sub
