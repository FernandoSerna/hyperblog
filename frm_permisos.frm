VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmpermisos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Permisos a Movimientos"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9915
   Icon            =   "frm_permisos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   9915
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   2535
      TabIndex        =   19
      Top             =   1845
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   20
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
         TabIndex        =   21
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frm_permisos.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   435
      Picture         =   "frm_permisos.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Guardar Alt + G"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   765
      Picture         =   "frm_permisos.frx":0ACE
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1095
      Picture         =   "frm_permisos.frx":0BA0
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1425
      Picture         =   "frm_permisos.frx":0CA2
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9285
      Picture         =   "frm_permisos.frx":0DA4
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   45
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   1350
      Left            =   135
      TabIndex        =   14
      Top             =   450
      Width           =   9615
      Begin VB.TextBox txt_nombre_almacen_2 
         Height          =   315
         Left            =   3315
         TabIndex        =   12
         Top             =   930
         Width           =   5400
      End
      Begin VB.TextBox txt_nombre_almacen_1 
         Height          =   315
         Left            =   3315
         TabIndex        =   10
         Top             =   570
         Width           =   5400
      End
      Begin VB.TextBox txt_nombre_movimiento 
         Height          =   315
         Left            =   3315
         TabIndex        =   8
         Top             =   210
         Width           =   5400
      End
      Begin VB.TextBox txt_movimiento 
         Height          =   315
         Left            =   2115
         TabIndex        =   7
         Top             =   210
         Width           =   1170
      End
      Begin VB.TextBox txt_almacen_1 
         Height          =   315
         Left            =   2115
         TabIndex        =   9
         Top             =   570
         Width           =   1170
      End
      Begin VB.TextBox txt_almacen_2 
         Height          =   315
         Left            =   2115
         TabIndex        =   11
         Top             =   930
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Almacen 2:"
         Height          =   195
         Left            =   1020
         TabIndex        =   18
         Top             =   990
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Almacen 1:"
         Height          =   195
         Left            =   1020
         TabIndex        =   17
         Top             =   630
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Movimiento:"
         Height          =   195
         Index           =   0
         Left            =   1020
         TabIndex        =   16
         Top             =   270
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3390
      Left            =   135
      TabIndex        =   0
      Top             =   1755
      Width           =   9615
      Begin MSComctlLib.ListView lv_permisos 
         Height          =   3195
         Left            =   60
         TabIndex        =   13
         Top             =   135
         Width           =   9480
         _ExtentX        =   16722
         _ExtentY        =   5636
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Movimiento"
            Object.Width           =   5997
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Almacen 1"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Almacen 2"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Usuario"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Clave Movmiento"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "cve almacen 1"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "cve almacen2"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "tipo movimiento"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   5220
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   12
      ImageHeight     =   12
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_permisos.frx":13DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_permisos.frx":1CB8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   4665
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_permisos.frx":2592
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_permisos.frx":2E6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_permisos.frx":3746
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_permisos.frx":3CE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_permisos.frx":45BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_permisos.frx":4E98
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_permisos.frx":5772
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_permisos.frx":5884
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_permisos.frx":5996
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_permisos.frx":5AA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_permisos.frx":5BBA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   90
      TabIndex        =   15
      Top             =   315
      Width           =   9675
   End
End
Attribute VB_Name = "frmpermisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim numero_items_permisos As Integer
Dim VAR_TIPO_LISTA As Integer




Private Sub cmd_deshacer_Click()
   Call pro_textos
End Sub

Private Sub cmd_eliminar_Click()
   Call pro_elimina_permisos
   rs.Open "select * from vw_permisos_movimientos where vcha_usu_usuario_id = '" + var_usuario_permiso + "'", cnn, adOpenDynamic, adLockOptimistic
   If rs.BOF Then
      cmd_guardar.Enabled = False
      cmd_deshacer.Enabled = False
      cmd_eliminar.Enabled = False
   Else
      cmd_guardar.Enabled = True
      cmd_deshacer.Enabled = True
      cmd_eliminar.Enabled = True
   End If
   rs.Close
End Sub

Private Sub cmd_guardar_Click()
   Call pro_guardar_permisos
   rs.Open "select * from vw_permisos_movimientos where vcha_usu_usuario_id = '" + var_usuario_permiso + "'", cnn, adOpenDynamic, adLockOptimistic
   If rs.BOF Then
      cmd_guardar.Enabled = False
      cmd_deshacer.Enabled = False
      cmd_eliminar.Enabled = False
   Else
      cmd_guardar.Enabled = True
      cmd_deshacer.Enabled = True
      cmd_eliminar.Enabled = True
   End If
   rs.Close
End Sub

Private Sub cmd_imprimir_Click()
   If vector_valida_passwords(var_indice_menu) = "*" Then
      frmpasswords.Show
   Else
      Call gPrintListView(lv_permisos, "LISTADO DE permisos")
   End If
End Sub

Private Sub cmd_nuevo_Click()
   Call pro_limpiatextos(Me)
   Me.txt_almacen_1.Enabled = True
   Me.txt_almacen_2.Enabled = True
   Me.txt_movimiento.Enabled = True
   Me.txt_nombre_almacen_1.Enabled = True
   Me.txt_nombre_almacen_2.Enabled = True
   Me.txt_nombre_movimiento.Enabled = True
   txt_movimiento.SetFocus: var_modifica_registro = False
   cmd_guardar.Enabled = True
   cmd_deshacer.Enabled = True
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 78 Then
      cmd_nuevo_Click
   End If
   If Shift = 4 And KeyCode = 71 Then
      cmd_guardar_Click
   End If
   If Shift = 4 And KeyCode = 68 Then
      cmd_deshacer_Click
   End If
   If Shift = 4 And KeyCode = 69 Then
      cmd_eliminar_Click
   End If
   If Shift = 4 And KeyCode = 73 Then
      cmd_imprimir_Click
   End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      If frm_lista.Visible = False Then
         Unload Me
      Else
         frm_lista.Visible = False
      End If
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 1000
   Left = 1100
   frm_lista.Visible = False
   var_modifica_registro = True
   lv_permisos.SmallIcons = ImageList
   Call pro_encabezadosView(Me, lv_permisos, False)
   Call pro_llena_listview1
   pro_textos
   rs.Open "select * from vw_permisos_movimientos where vcha_usu_usuario_id = '" + var_usuario_permiso + "'", cnn, adOpenDynamic, adLockOptimistic
   If rs.BOF Then
      cmd_guardar.Enabled = False
      cmd_deshacer.Enabled = False
      cmd_eliminar.Enabled = False
   Else
      cmd_guardar.Enabled = True
      cmd_deshacer.Enabled = True
      cmd_eliminar.Enabled = True
   End If
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'If var_despliega_menu = True Then
   '   var_swpassword = False
   '   var_modifica_registro = False
   'End If
   Call activa_forma(var_activa_forma_permisos)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If VAR_TIPO_LISTA = 1 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_movimiento = lv_lista.selectedItem
            txt_nombre_movimiento = lv_lista.selectedItem.SubItems(1)
         Else
            txt_movimiento = ""
            txt_nombre_movimiento = ""
         End If
        
         txt_movimiento.SetFocus
      End If
      If VAR_TIPO_LISTA = 2 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_almacen_1 = lv_lista.selectedItem
            txt_nombre_almacen_1 = lv_lista.selectedItem.SubItems(1)
         Else
            txt_almacen_1 = ""
            txt_nombre_almacen_1 = ""
         End If
         txt_almacen_1.SetFocus
      End If
      If VAR_TIPO_LISTA = 3 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_almacen_2 = lv_lista.selectedItem
            txt_nombre_almacen_2 = lv_lista.selectedItem.SubItems(1)
         Else
            txt_almacen_2 = ""
            txt_nombre_almacen_2 = ""
         End If
         txt_almacen_2.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      If VAR_TIPO_LISTA = 1 Then
         txt_movimiento.SetFocus
      End If
      If VAR_TIPO_LISTA = 2 Then
         txt_almacen_1.SetFocus
      End If
      If VAR_TIPO_LISTA = 3 Then
         txt_almacen_2.SetFocus
      End If
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub lv_permisos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_permisos, ColumnHeader)
End Sub

Private Sub lv_permisos_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Set lv_permisos.selectedItem = Item
   pro_textos
   var_modifica_registro = True
   txt_movimiento.Enabled = False
End Sub



Sub pro_guardar_permisos()
Dim ok As Boolean
Set TB_PERMISOS_MOVIMIENTOS = New TB_PERMISOS_MOVIMIENTOS
    ok = True
    If txt_movimiento <> "" And txt_almacen_1 <> "" Then
        If var_hubo_cambios Then
            ok = TB_PERMISOS_MOVIMIENTOS.Anadir(var_usuario_permiso, txt_movimiento, txt_almacen_1, txt_almacen_2)
            If ok Then
                bitacora = True
                pro_actualiza_ListView
                txt_movimiento.Enabled = False
                MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
                txt_registros = lv_permisos.ListItems.Count
                var_modifica_registro = True
            Else
                MsgBox "No se puede grabar registro: " + TB_PERMISOS_MOVIMIENTOS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
        End If
    End If
Set TB_PERMISOS_MOVIMIENTOS = Nothing: var_hubo_cambios = False
End Sub

Sub pro_elimina_permisos()
   Dim var_llave_usuarios As String
   Set TB_PERMISOS_MOVIMIENTOS = New TB_PERMISOS_MOVIMIENTOS
   On Error GoTo salir:
   ok = True
   If txt_movimiento <> "" And txt_almacen_1 <> "" And var_modifica_registro = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_PERMISOS_MOVIMIENTOS.Eliminar(var_usuario_permiso, txt_movimiento, txt_almacen_1, txt_almacen_2)
      Else
         GoTo salir:
      End If
      If ok Then
         numero_items_permisos = numero_items_permisos - 1
         MsgBox "Se Elimino Correctamente el Registro", vbInformation
         lv_permisos.ListItems.Remove (lv_permisos.selectedItem.Index)
         Call pro_limpiatextos(Me)
         txt_registros = lv_permisos.ListItems.Count
         lv_permisos.selectedItem.Selected = True
         pro_textos
      Else
         MsgBox "No se puede eliminar registro: " + TB_PERMISOS_MOVIMIENTOS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
   Set TB_PERMISOS_MOVIMIENTOS = Nothing
End Sub


Sub pro_llena_listview1()
   Dim list_item As ListItem
   
   rs.Open "select * from vw_permisos_movimientos where vcha_usu_usuario_id = '" + var_usuario_permiso + "'", cnn, adOpenDynamic, adLockOptimistic
   numero_items_permisos = 0
   While Not rs.EOF
      rsaux.Open "select secondary_inventory_name VCHA_ALM_ALMACEN_ID, description VCHA_ALM_NOMBRE from mtl_secondary_inventories where secondary_inventory_name = '" + IIf(IsNull(rs!vcha_per_almacen_1), "", rs!vcha_per_almacen_1) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rsaux.EOF Then
         var_nombre_almacen = IIf(IsNull(rsaux!vcha_alm_nombre), "", rsaux!vcha_alm_nombre)
      End If
      rsaux.Close
      Set list_item = lv_permisos.ListItems.Add(, , rs!vcha_mov_nombre)
      list_item.SubItems(1) = var_nombre_almacen
      list_item.SubItems(2) = IIf(IsNull(rs!vcha_alm_almacen2), "", rs!vcha_alm_almacen2)
      list_item.SubItems(3) = var_usuario_permiso
      list_item.SubItems(4) = IIf(IsNull(rs!VCHA_MOV_MOVIMIENTO_ID), "", rs!VCHA_MOV_MOVIMIENTO_ID)
      list_item.SubItems(5) = IIf(IsNull(rs!vcha_per_almacen_1), "", rs!vcha_per_almacen_1)
      list_item.SubItems(6) = IIf(IsNull(rs!vcha_per_almacen_2), "", rs!vcha_per_almacen_2)
      list_item.SubItems(7) = IIf(IsNull(rs!CHAR_MOV_AFECTACION), "", rs!CHAR_MOV_AFECTACION)
      rs.MoveNext:
      numero_items_permisos = numero_items_permisos + 1
    Wend
    rs.Close
End Sub


Sub pro_textos()
On Error GoTo err0:
   var_n = lv_permisos.ListItems.Count
   If var_n > 0 Then
      txt_movimiento = lv_permisos.selectedItem.SubItems(4)
      txt_almacen_1 = lv_permisos.selectedItem.SubItems(5)
      txt_almacen_2 = lv_permisos.selectedItem.SubItems(6)
      txt_nombre_movimiento = lv_permisos.selectedItem
      txt_nombre_almacen_1 = lv_permisos.selectedItem.SubItems(1)
      txt_nombre_almacen_2 = lv_permisos.selectedItem.SubItems(2)
   End If
   Me.txt_almacen_1.Enabled = False
   Me.txt_almacen_2.Enabled = False
   Me.txt_movimiento.Enabled = False
   Me.txt_nombre_almacen_1.Enabled = False
   Me.txt_nombre_almacen_2.Enabled = False
   Me.txt_nombre_movimiento.Enabled = False
   var_numero_renglones = lv_permisos.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_permisos.ColumnHeaders(2).Width = 2750
   Else
      lv_permisos.ColumnHeaders(2).Width = 3000.18
   End If
   var_hubo_cambios = False
err0:
End Sub

Private Sub pro_actualiza_ListView()
   lv_permisos.ListItems.Clear
   Call pro_llena_listview1
End Sub



Private Sub txt_almacen_1_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_almacen_1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select secondary_inventory_name VCHA_ALM_ALMACEN_ID, description VCHA_ALM_NOMBRE from mtl_secondary_inventories", cnnoracle_4, adOpenDynamic, adLockOptimistic
      'rs.Open "select * from VW_MOVIMIENTOS_almacenes WHERE VCHA_MOV_MOVIMIENTO_ID = '" + txt_movimiento + "' order by vcha_alm_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ALM_ALMACEN_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_alm_nombre), "", rs!vcha_alm_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ALMACENES"
      VAR_TIPO_LISTA = 2
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

Private Sub txt_almacen_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_almacen_1_LostFocus()
   If Trim(txt_almacen_1) <> "" Then
       rs.Open "select secondary_inventory_name VCHA_ALM_ALMACEN_ID, description VCHA_ALM_NOMBRE from mtl_secondary_inventories WHERE secondary_inventory_name = '" + Me.txt_almacen_1 + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      'rs.Open "SELECT * FROM VW_MOVIMIENTOS_ALMACENES WHERE VCHA_ALM_ALMACEN_ID = '" + txt_almacen_1 + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + txt_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_almacen_1 = IIf(IsNull(rs!vcha_alm_nombre), "", rs!vcha_alm_nombre)
      Else
         MsgBox "Clave de almacen incorrecta", vbOKOnly, "ATENCION"
         txt_nombre_almacen_1 = ""
         txt_almacen_1 = ""
      End If
      rs.Close
   Else
      txt_nombre_almacen = ""
   End If
End Sub

Private Sub txt_almacen_2_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_almacen_2_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from VW_MOVIMIENTOS_almacenes WHERE VCHA_MOV_MOVIMIENTO_ID = '" + txt_movimiento + "' order by vcha_alm_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ALM_ALMACEN_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_alm_nombre), "", rs!vcha_alm_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ALMACENES"
      VAR_TIPO_LISTA = 3
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

Private Sub txt_almacen_2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_almacen_2_LostFocus()
   If Trim(txt_almacen_2) <> "" Then
      rs.Open "SELECT * FROM VW_MOVIMIENTOS_ALMACENES WHERE VCHA_ALM_ALMACEN_ID = '" + txt_almacen_2 + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + txt_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_almacen_2 = IIf(IsNull(rs!vcha_alm_nombre), "", rs!vcha_alm_nombre)
      Else
         MsgBox "Clave de almacen incorrecta", vbOKOnly, "ATENCION"
         txt_nombre_almacen_2 = ""
         txt_almacen_2 = ""
      End If
      rs.Close
   Else
      txt_nombre_almacen_2 = ""
   End If
End Sub

Private Sub txt_movimiento_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_movimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_movimientos order by vcha_mov_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_MOV_MOVIMIENTO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_mov_nombre), "", rs!vcha_mov_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "MOVIMIENTOS"
      VAR_TIPO_LISTA = 1
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

Private Sub txt_movimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_movimiento_LostFocus()
   If Trim(txt_movimiento) <> "" Then
      rs.Open "select * from tb_movimientos where vcha_mov_movimiento_id = '" + txt_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_movimiento = IIf(IsNull(rs!vcha_mov_nombre), "", rs!vcha_mov_nombre)
      Else
         MsgBox "Clave de Movimiento incorrecto", vbOKOnly, "ATENCION"
         txt_movimiento = ""
         txt_nombre_movimiento = ""
      End If
      rs.Close
   End If
End Sub

Private Sub txt_nombre_almacen_1_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_almacen_1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from VW_MOVIMIENTOS_almacenes WHERE VCHA_MOV_MOVIMIENTO_ID = '" + txt_movimiento + "' order by vcha_alm_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ALM_ALMACEN_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_alm_nombre), "", rs!vcha_alm_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ALMACENES"
      VAR_TIPO_LISTA = 2
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

Private Sub txt_nombre_almacen_1_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_almacen_2_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_almacen_2_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from VW_MOVIMIENTOS_almacenes WHERE VCHA_MOV_MOVIMIENTO_ID = '" + txt_movimiento + "' order by vcha_alm_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ALM_ALMACEN_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_alm_nombre), "", rs!vcha_alm_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ALMACENES"
      VAR_TIPO_LISTA = 3
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

Private Sub txt_nombre_almacen_2_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub

Private Sub txt_nombre_movimiento_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_movimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_movimientos order by vcha_mov_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_MOV_MOVIMIENTO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_mov_nombre), "", rs!vcha_mov_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "MOVIMIENTOS"
      VAR_TIPO_LISTA = 1
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

Private Sub txt_nombre_movimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub
