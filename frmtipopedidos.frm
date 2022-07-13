VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmtipopedidos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Pedidos"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   Icon            =   "frmtipopedidos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5955
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   135
      TabIndex        =   28
      Top             =   600
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   29
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
         TabIndex        =   30
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmtipopedidos.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      Picture         =   "frmtipopedidos.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Guardar Alt + G"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   810
      Picture         =   "frmtipopedidos.frx":0ACE
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      Picture         =   "frmtipopedidos.frx":0BA0
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1470
      Picture         =   "frmtipopedidos.frx":0CA2
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
      Picture         =   "frmtipopedidos.frx":0DA4
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   23
      Top             =   285
      Width           =   5655
   End
   Begin VB.Frame Frame3 
      Height          =   3825
      Left            =   150
      TabIndex        =   21
      Top             =   3405
      Width           =   5655
      Begin MSComctlLib.ListView lv_tipopedidos 
         Height          =   3600
         Left            =   45
         TabIndex        =   22
         Top             =   135
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   6350
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   6879
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "resurtible"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "autorizacion"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "tipo cliente"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "carga"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "dias"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "movimiento"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "iva"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ImageList icono_encabezado 
         Left            =   0
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
               Picture         =   "frmtipopedidos.frx":13DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmtipopedidos.frx":1CB8
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Tipo de Pedidos"
      Height          =   2955
      Left            =   150
      TabIndex        =   18
      Top             =   420
      Width           =   5655
      Begin VB.TextBox txt_nombre_tipo_cliente 
         Height          =   315
         Left            =   2445
         MaxLength       =   50
         TabIndex        =   12
         Top             =   1350
         Width           =   3060
      End
      Begin VB.TextBox txt_nombre_movimiento 
         Height          =   315
         Left            =   2445
         MaxLength       =   50
         TabIndex        =   16
         Top             =   2235
         Width           =   3060
      End
      Begin VB.TextBox txt_iva 
         Height          =   315
         Left            =   1155
         MaxLength       =   2
         TabIndex        =   17
         Top             =   2565
         Width           =   1275
      End
      Begin VB.TextBox txt_movimiento 
         Height          =   315
         Left            =   1155
         MaxLength       =   50
         TabIndex        =   15
         Top             =   2235
         Width           =   1275
      End
      Begin VB.TextBox txt_caducidad 
         Height          =   315
         Left            =   1155
         MaxLength       =   50
         TabIndex        =   14
         Top             =   1905
         Width           =   1275
      End
      Begin VB.CheckBox chk_requiere_archivo_carga 
         Caption         =   "Requiere Archivo de Carga"
         Height          =   210
         Left            =   1155
         TabIndex        =   13
         Top             =   1680
         Width           =   3030
      End
      Begin VB.TextBox txt_tipo_cliente 
         Height          =   315
         Left            =   1155
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1350
         Width           =   1275
      End
      Begin VB.CheckBox chk_requiere_autorizacion 
         Caption         =   "Requiere Autorización"
         Height          =   210
         Left            =   1155
         TabIndex        =   10
         Top             =   1140
         Width           =   2025
      End
      Begin VB.CheckBox chk_resurtible 
         Caption         =   "Resurtible"
         Height          =   210
         Left            =   1155
         TabIndex        =   9
         Top             =   915
         Width           =   1410
      End
      Begin VB.TextBox txt_tipo_pedido 
         Height          =   315
         Left            =   1155
         MaxLength       =   50
         TabIndex        =   7
         Top             =   255
         Width           =   1275
      End
      Begin VB.TextBox txt_nombre_tipo_pedido 
         Height          =   315
         Left            =   1155
         MaxLength       =   50
         TabIndex        =   8
         Top             =   585
         Width           =   4350
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "IVA:"
         Height          =   195
         Index           =   5
         Left            =   165
         TabIndex        =   27
         Top             =   2640
         Width           =   300
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Movimiento"
         Height          =   195
         Index           =   4
         Left            =   165
         TabIndex        =   26
         Top             =   2310
         Width           =   810
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Caducidad:"
         Height          =   195
         Index           =   3
         Left            =   165
         TabIndex        =   25
         Top             =   1980
         Width           =   810
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Clientes:"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   24
         Top             =   1380
         Width           =   960
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   20
         Top             =   255
         Width           =   450
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   19
         Top             =   615
         Width           =   885
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2700
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   15
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmtipopedidos.frx":2592
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtipopedidos.frx":2E6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtipopedidos.frx":3746
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtipopedidos.frx":3CE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtipopedidos.frx":45BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtipopedidos.frx":4E98
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtipopedidos.frx":5772
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtipopedidos.frx":5884
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtipopedidos.frx":5996
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtipopedidos.frx":5AA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtipopedidos.frx":5BBA
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmtipopedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim bitacora As Boolean
Dim numero_items_tipopedidos As Integer
Dim var_tipo_lista As Integer


Private Sub chk_requiere_archivo_carga_Click()
   var_hubo_cambios = True
End Sub

Private Sub chk_requiere_autorizacion_Click()
   var_hubo_cambios = True
End Sub

Private Sub chk_resurtible_Click()
   var_hubo_cambios = True
End Sub

Private Sub cmd_deshacer_Click()
   Call pro_textos
End Sub

Private Sub cmd_eliminar_Click()
   var_opcion_seguridad = 2
   var_acepta_seguridad = 1
   If var_global_permiso3 = 1 Then
      var_acepta_seguridad = 2
      If var_global_permiso4 = 1 Then
         frmpasswords2.Show 1
      Else
         frmpasswords.Show 1
      End If
   End If
   If var_acepta_seguridad = 1 Then
      Call pro_elimina_tipopedidos
      rs.Open "select * from tb_tipopedidos", cnn, adOpenDynamic, adLockOptimistic
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
   End If
End Sub

Private Sub cmd_guardar_Click()
Dim var_posible As Boolean
   var_posible = True
   If var_modifica_registro_tipopedido = False Then
      rs.Open "select * from TB_tipopedidos where char_tpe_tipo_pedido_id = '" + Me.txt_tipo_pedido + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_posible = False
      End If
      rs.Close
   End If
   If var_posible = True Then
      var_opcion_seguridad = 2
      var_acepta_seguridad = 1
      If var_global_permiso3 = 1 Then
         var_acepta_seguridad = 2
         If var_global_permiso4 = 1 Then
            frmpasswords2.Show 1
         Else
            frmpasswords.Show 1
         End If
      End If
      If var_acepta_seguridad = 1 Then
         Call pro_guardar_tipopedidos
         rs.Open "select * from tb_tipopedidos", cnn, adOpenDynamic, adLockOptimistic
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
      End If
   Else
      MsgBox "Clave de tipo de pedido ya existe", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call gPrintListView(lv_tipopedidos, "LISTADO DE tipopedidos")
        End If

End Sub

Private Sub cmd_nuevo_Click()
        Call pro_limpiatextos(Me)
        txt_tipo_pedido.Enabled = True
        txt_tipo_pedido.SetFocus: var_modifica_registro_tipopedido = False
        cmd_guardar.Enabled = True
        cmd_deshacer.Enabled = True
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_tipopedido = False Then
      var_si = MsgBox("No se han guardado los cambios, ¿Desea salir?", vbYesNo, "ATENCION")
      If var_si <> 6 Then
         GoTo salir:
      End If
   Else
      If var_hubo_cambios = True Then
         var_si = MsgBox("No se han guardado los cambios, ¿Desea salir?", vbYesNo, "ATENCION")
         If var_si <> 6 Then
            GoTo salir:
         End If
      End If
   End If
   Unload Me
   Exit Sub
salir:
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

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 0
   Left = 2900
   frm_lista.Visible = False
   
   rs.Open "select * from tb_tipopedidos", cnn, adOpenDynamic, adLockOptimistic
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
   var_modifica_registro_tipopedido = True
   lv_tipopedidos.SmallIcons = ImageList
   Call pro_encabezadosView(Me, lv_tipopedidos, False)
   Call pro_llena_listview1
   pro_textos
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_despliega_menu = True Then
      var_swpassword = False
      var_modifica_registro_tipopedido = False
   End If
   Call activa_forma(var_activa_forma_tipopedidos)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_tipo_lista = 1 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_tipo_cliente = lv_lista.selectedItem
            txt_nombre_tipo_cliente = lv_lista.selectedItem.SubItems(1)
         Else
            txt_tipo_cliente = ""
            txt_nombre_tipo_cliente = ""
         End If
         txt_tipo_cliente.SetFocus
      End If
      If var_tipo_lista = 2 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_movimiento = lv_lista.selectedItem
            txt_nombre_movimiento = lv_lista.selectedItem.SubItems(1)
         Else
            txt_movimiento = ""
            txt_nombre_movimiento = ""
         End If
         txt_movimiento.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub lv_tipopedidos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_tipopedidos, ColumnHeader)
End Sub

Private Sub lv_tipopedidos_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_tipopedidos.selectedItem = Item
        pro_textos
        var_modifica_registro_tipopedido = True
        txt_tipo_pedido.Enabled = False

End Sub



Sub pro_guardar_tipopedidos()

Dim ok As Boolean
Set TB_TIPOPEDIDOS = New TB_TIPOPEDIDOS
Set TB_BITACORA_TIPOPEDIDOS = New TB_BITACORA_TIPOPEDIDOS
    
    ok = True
    If txt_tipo_pedido <> "" And txt_nombre_tipo_pedido <> "" And txt_iva <> "" Then
        If var_hubo_cambios Then
            rs.Open "select * from tb_tipopedidos where CHAR_TPE_TIPO_PEDIDO_ID = '" + txt_tipo_pedido + "'", cnn, adOpenDynamic, adLockOptimistic
            If Trim(txt_caducidad) = "" Then
               txt_caducidad = 0
            End If
            If Trim(txt_iva) = "" Then
               txt_iva = 0
            End If
            ok = TB_TIPOPEDIDOS.Anadir(txt_tipo_pedido, txt_nombre_tipo_pedido, chk_resurtible, chk_requiere_autorizacion, txt_tipo_cliente, chk_requiere_archivo_carga, CDbl(txt_caducidad), txt_movimiento, CDbl(txt_iva))
            If ok Then
                bitacora = True
                If var_modifica_registro_tipopedido = False Then
                   var_operacion_bitacora = "I"
                   bitacora = TB_BITACORA_TIPOPEDIDOS.Anadir(txt_tipo_pedido, "VCHA_TPE_NOMBRE", "", txt_nombre_tipo_pedido, var_clave_usuario_global, fun_NombrePc, Date)
                Else
                   var_operacion_bitacora = "M"
                   If rs(0) <> txt_tipo_pedido Then
                      bitacora = TB_BITACORA_TIPOPEDIDOS.Anadir(txt_tipo_pedido, "CHAR_TPE_TIPO_PEDIDO_ID", rs(0), txt_tipo_pedido, var_clave_usuario_global, fun_NombrePc, Date)
                   End If
                   If rs(1) <> txt_nombre_tipo_pedido Then
                      bitacora = TB_BITACORA_TIPOPEDIDOS.Anadir(txt_tipo_pedido, "VCHA_TPE_NOMBRE", rs(1), txt_nombre_tipo_pedido, var_clave_usuario_global, fun_NombrePc, Date)
                   End If
                   If rs(2) <> chk_resurtible Then
                      bitacora = TB_BITACORA_TIPOPEDIDOS.Anadir(txt_tipo_pedido, "INTE_TPE_RESURTIBLE", rs(2), chk_resurtible, var_clave_usuario_global, fun_NombrePc, Date)
                   End If
                   If rs(3) <> chk_requiere_autorizacion Then
                      bitacora = TB_BITACORA_TIPOPEDIDOS.Anadir(txt_tipo_pedido, "INTE_TPE_AUTORIZACION", rs(3), chk_requiere_autorizacion, var_clave_usuario_global, fun_NombrePc, Date)
                   End If
                End If
                rs.Close
                pro_actualiza_ListView
                txt_tipo_pedido.Enabled = False
                MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
                txt_registros = lv_tipopedidos.ListItems.Count
                var_modifica_registro_tipopedido = True
            Else
                MsgBox "No se puede grabar registro: " + TB_TIPOPEDIDOS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
        End If
    Else
       MsgBox "Información incompleta", vbOKOnly, "ATENCION"
    End If
    
Set TB_TIPOPEDIDOS = Nothing: var_hubo_cambios = False

End Sub

Sub pro_elimina_tipopedidos()
On Error GoTo salir:
   Dim var_llave_usuarios As String
   Set TB_TIPOPEDIDOS = New TB_TIPOPEDIDOS
   Set TB_BITACORA_TIPOPEDIDOS = New TB_BITACORA_TIPOPEDIDOS
   ok = True
   bitacora = True
   If txt_tipo_pedido <> "" And txt_nombre_tipo_pedido <> "" And var_modifica_registro_tipopedido = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_TIPOPEDIDOS.Eliminar(txt_tipo_pedido)
      Else
         GoTo salir:
      End If
      If ok Then
        numero_items_tipopedidos = numero_items_tipopedidos - 1
        MsgBox "Se Elimino Correctamente el Registro", vbInformation
        var_operacion_bitacora = "E"
        bitacora = TB_BITACORA_TIPOPEDIDOS.Anadir(txt_tipo_pedido, "VCHA_TAL_TALLA_ID", "", txt_tipo_pedido, var_clave_usuario_global, fun_NombrePc, Date)
        lv_tipopedidos.ListItems.Remove (lv_tipopedidos.selectedItem.Index)
        Call pro_limpiatextos(Me)
        txt_registros = lv_tipopedidos.ListItems.Count
        If lv_tipopedidos.selectedItem.Selected <> False Then
           lv_tipopedidos.selectedItem.Selected = True
        End If
        pro_textos
      Else
        MsgBox "No se puede eliminar registro: " + TB_TIPOPEDIDOS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
   Set TB_TIPOPEDIDOS = Nothing
End Sub


Sub pro_llena_listview1()
   Dim list_item As ListItem
   rs.Open "select * from TB_tipopedidos", cnn, adOpenDynamic, adLockOptimistic
   numero_items_tipopedidos = 0
   While Not rs.EOF
      Set list_item = lv_tipopedidos.ListItems.Add(, , rs!char_tpe_tipo_pedido_id)
      list_item.SubItems(1) = IIf(IsNull(rs!vcha_tpe_nombre), "", rs!vcha_tpe_nombre)
      list_item.SubItems(2) = IIf(IsNull(rs!inte_tpe_resurtible), 0, rs!INTE_TPE_AUTORIZACION)
      list_item.SubItems(3) = IIf(IsNull(rs!INTE_TPE_AUTORIZACION), 0, rs!INTE_TPE_AUTORIZACION)
      list_item.SubItems(4) = IIf(IsNull(rs!vcha_tcl_tipo_cliente_id), "", rs!vcha_tcl_tipo_cliente_id)
      list_item.SubItems(5) = IIf(IsNull(rs!INTE_TPE_CARGA_ARCHIVO), 0, rs!INTE_TPE_CARGA_ARCHIVO)
      list_item.SubItems(6) = IIf(IsNull(rs!inte_tpe_dias_caducidad), 0, rs!inte_tpe_dias_caducidad)
      list_item.SubItems(7) = IIf(IsNull(rs!VCHA_MOV_MOVIMIENTO_ID), "", rs!VCHA_MOV_MOVIMIENTO_ID)
      list_item.SubItems(8) = IIf(IsNull(rs!FLOA_TPE_IVA), 0, rs!FLOA_TPE_IVA)
      rs.MoveNext:
      numero_items_tipopedidos = numero_items_tipopedidos + 1
    Wend
    rs.Close
End Sub


Sub pro_textos()
'On Error GoTo err0:
   Dim var_n As Integer
   var_n = lv_tipopedidos.ListItems.Count
   If var_n > 0 Then
        txt_tipo_pedido = lv_tipopedidos.selectedItem
        txt_nombre_tipo_pedido = lv_tipopedidos.selectedItem.SubItems(1)
        chk_resurtible = lv_tipopedidos.selectedItem.SubItems(2)
        chk_requiere_autorizacion = lv_tipopedidos.selectedItem.SubItems(3)
        txt_tipo_cliente = lv_tipopedidos.selectedItem.SubItems(4)
        chk_requiere_archivo_carga = lv_tipopedidos.selectedItem.SubItems(5)
        txt_caducidad = lv_tipopedidos.selectedItem.SubItems(6)
        txt_movimiento = lv_tipopedidos.selectedItem.SubItems(7)
        txt_iva = lv_tipopedidos.selectedItem.SubItems(8)
        rs.Open "SELECT * FROM TB_TIPOSCLIENTES WHERE VCHA_TCL_TIPO_CLIENTE_ID = '" + txt_tipo_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
           txt_nombre_tipo_cliente = IIf(IsNull(rs!VCHA_TCL_nombre), "", rs!VCHA_TCL_nombre)
        Else
           txt_nombre_tipo_cliente = ""
        End If
        rs.Close
        rs.Open "SELECT * FROM TB_MOVIMIENTOS WHERE VCHA_MOV_MOVIMIENTO_ID = '" + txt_movimiento + "' AND CHAR_MOV_DOCUMENTO = 'F'", cnn, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
           txt_nombre_movimiento = IIf(IsNull(rs!vcha_mov_nombre), "", rs!vcha_mov_nombre)
        Else
           txt_nombre_movimiento = ""
        End If
        rs.Close
   End If
   var_numero_renglones = lv_tipopedidos.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_tipopedidos.ColumnHeaders(2).Width = 3850
   Else
      lv_tipopedidos.ColumnHeaders(2).Width = 4099.71
   End If
   var_hubo_cambios = False
   var_modifica_registro_tipopedido = True
   Me.txt_tipo_pedido.Enabled = False
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem

    If var_modifica_registro_tipopedido = False Then
        Set list_item = lv_tipopedidos.ListItems.Add(, , txt_tipo_pedido)
        list_item.SubItems(1) = txt_nombre_tipo_pedido
        list_item.SubItems(2) = chk_resurtible
        list_item.SubItems(3) = chk_requiere_autorizacion
        list_item.SubItems(4) = txt_tipo_cliente
        list_item.SubItems(5) = chk_requiere_archivo_carga
        list_item.SubItems(6) = txt_caducidad
        list_item.SubItems(7) = txt_movimiento
        list_item.SubItems(8) = txt_iva
        list_item.EnsureVisible
        list_item.Selected = True
        numero_items_tipopedidos = numero_items_tipopedidos + 1
    Else
        lv_tipopedidos.ListItems.Item(lv_tipopedidos.selectedItem.Index).Checked = False
        lv_tipopedidos.ListItems.Item(lv_tipopedidos.selectedItem.Index) = txt_tipo_pedido
        lv_tipopedidos.ListItems.Item(lv_tipopedidos.selectedItem.Index).ListSubItems(1) = txt_nombre_tipo_pedido
        lv_tipopedidos.ListItems.Item(lv_tipopedidos.selectedItem.Index).ListSubItems(2) = chk_resurtible
        lv_tipopedidos.ListItems.Item(lv_tipopedidos.selectedItem.Index).ListSubItems(3) = chk_requiere_autorizacion
        lv_tipopedidos.ListItems.Item(lv_tipopedidos.selectedItem.Index).ListSubItems(4) = txt_tipo_cliente
        lv_tipopedidos.ListItems.Item(lv_tipopedidos.selectedItem.Index).ListSubItems(5) = chk_requiere_archivo_carga
        lv_tipopedidos.ListItems.Item(lv_tipopedidos.selectedItem.Index).ListSubItems(6) = txt_caducidad
        lv_tipopedidos.ListItems.Item(lv_tipopedidos.selectedItem.Index).ListSubItems(7) = txt_movimiento
        lv_tipopedidos.ListItems.Item(lv_tipopedidos.selectedItem.Index).ListSubItems(8) = txt_iva
        lv_tipopedidos.ListItems.Item(lv_tipopedidos.selectedItem.Index).Selected = True
    End If
'    lv_tipopedidos.SetFocus
End Sub






Private Sub txt_caducidad_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_iva_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_iva_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Me.cmd_guardar.Enabled = True Then
         Me.cmd_guardar.SetFocus
      End If
   End If
End Sub

Private Sub txt_movimiento_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_movimiento_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para accesar al catálogo"
End Sub

Private Sub txt_movimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_MOVIMIENTOS WHERE  CHAR_MOV_DOCUMENTO =  'F' order by vcha_MOV_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_MOV_MOVIMIENTO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_mov_nombre), "", rs!vcha_mov_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "MOVIMIENTOS"
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
   If KeyCode = 117 Then
      var_activa_forma_tipopedidos = Me.Name
      Me.Enabled = False
      frmmovimientos.Show
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
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_movimiento) <> "" Then
      rs.Open "SELECT * FROM TB_MOVIMIENTOS WHERE VCHA_MOV_MOVIMIENTO_ID = '" + txt_movimiento + "' AND CHAR_MOV_DOCUMENTO = 'F'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_movimiento = IIf(IsNull(rs!vcha_mov_nombre), "", rs!vcha_mov_nombre)
      Else
         txt_nombre_movimiento = ""
         txt_movimiento = ""
         MsgBox "Clave de movimiento incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      txt_nombre_movimiento = ""
   End If
End Sub

Private Sub txt_nombre_movimiento_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_movimiento_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para accesar al catálogo"
End Sub

Private Sub txt_nombre_movimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_MOVIMIENTOS WHERE  CHAR_MOV_DOCUMENTO =  'F' order by vcha_MOV_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_MOV_MOVIMIENTO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_mov_nombre), "", rs!vcha_mov_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "MOVIMIENTOS"
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
   If KeyCode = 117 Then
      var_activa_forma_movimientos = Me.Name
      Me.Enabled = False
      frmmovimientos.Show
   End If
End Sub

Private Sub txt_nombre_movimiento_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_movimiento_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_tipo_cliente_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_tipo_cliente_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para accesar al catálogo"
End Sub

Private Sub txt_nombre_tipo_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_TIPOSCLIENTES order by vcha_TCL_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_tcl_tipo_cliente_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_TCL_nombre), "", rs!VCHA_TCL_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TIPO CLIENTES"
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
   If KeyCode = 117 Then
      var_activa_forma_tiposclientes = Me.Name
      Me.Enabled = False
      frmtiposclientes.Show
   End If
End Sub

Private Sub txt_nombre_tipo_cliente_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_tipo_cliente_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_tipo_pedido_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_tipo_pedido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_tipo_cliente_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_tipo_cliente_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para accesar al catálogo"
End Sub

Private Sub txt_tipo_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_TIPOSCLIENTES order by vcha_TCL_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_tcl_tipo_cliente_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_TCL_nombre), "", rs!VCHA_TCL_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TIPO CLIENTES"
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
   If KeyCode = 117 Then
      var_activa_forma_tiposclientes = Me.Name
      Me.Enabled = False
      frmtiposclientes.Show
   End If
End Sub

Private Sub txt_tipo_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_tipo_cliente_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_tipo_cliente) <> "" Then
      rs.Open "SELECT * FROM TB_TIPOSCLIENTES WHERE VCHA_TCL_TIPO_CLIENTE_ID = '" + txt_tipo_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_tipo_cliente = IIf(IsNull(rs!VCHA_TCL_nombre), "", rs!VCHA_TCL_nombre)
      Else
         Me.txt_tipo_cliente = ""
         txt_nombre_tipo_cliente = ""
         MsgBox "Clave de tipo de cliente incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      txt_nombre_tipo_cliente = ""
   End If
End Sub

Private Sub txt_tipo_pedido_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_tipo_pedido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub
