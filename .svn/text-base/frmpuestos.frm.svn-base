VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmpuestos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Puestos"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   Icon            =   "frmpuestos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5955
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   120
      TabIndex        =   19
      Top             =   450
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
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5460
      Picture         =   "frmpuestos.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1470
      Picture         =   "frmpuestos.frx":0F04
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      Picture         =   "frmpuestos.frx":1006
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   810
      Picture         =   "frmpuestos.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      Picture         =   "frmpuestos.frx":11DA
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Guardar Alt + G"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmpuestos.frx":12DC
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1455
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   17
      Top             =   5805
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   "  Puestos "
      Height          =   1500
      Left            =   150
      TabIndex        =   13
      Top             =   405
      Width           =   5655
      Begin VB.TextBox txt_nombre_menu 
         Height          =   315
         Left            =   2265
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   870
         Width           =   3285
      End
      Begin VB.CheckBox chk_supervisor 
         Caption         =   "Supervisor"
         Height          =   210
         Left            =   1110
         TabIndex        =   11
         Top             =   1230
         Width           =   1245
      End
      Begin VB.TextBox txt_menu 
         Height          =   315
         Left            =   1110
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   9
         Top             =   870
         Width           =   1125
      End
      Begin VB.TextBox txt_nombre_puesto 
         Height          =   315
         Left            =   1110
         MaxLength       =   50
         TabIndex        =   8
         Top             =   525
         Width           =   4440
      End
      Begin VB.TextBox txt_puesto 
         Height          =   315
         Left            =   1110
         MaxLength       =   50
         TabIndex        =   7
         Top             =   180
         Width           =   1125
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Menu:"
         Height          =   195
         Index           =   2
         Left            =   315
         TabIndex        =   16
         Top             =   930
         Width           =   450
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   1
         Left            =   315
         TabIndex        =   15
         Top             =   585
         Width           =   600
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   14
         Top             =   240
         Width           =   450
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5310
      Left            =   150
      TabIndex        =   0
      Top             =   1890
      Width           =   5655
      Begin MSComctlLib.ListView lv_puestos 
         Height          =   5100
         Left            =   45
         TabIndex        =   12
         Top             =   150
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   8996
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   6879
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "menu"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "supervisor"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "hace referencia"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "referencia"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "folio"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   255
      Top             =   5595
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
            Picture         =   "frmpuestos.frx":13DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpuestos.frx":1CB8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3450
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpuestos.frx":2592
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpuestos.frx":2E6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpuestos.frx":3746
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpuestos.frx":3CE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpuestos.frx":45BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpuestos.frx":4E98
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpuestos.frx":5772
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpuestos.frx":5884
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpuestos.frx":5996
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpuestos.frx":5AA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpuestos.frx":5BBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpuestos.frx":5CCC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   18
      Top             =   285
      Width           =   5655
   End
End
Attribute VB_Name = "frmpuestos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim numero_items_puestos As Integer
Dim bitacora As Boolean
Private Sub chk_supervisor_Click()
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
      Call pro_elimina_puestos
      rs.Open "select * from tb_puestos", cnn, adOpenDynamic, adLockOptimistic
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
      Call pro_guardar_puestos
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open "select * from tb_puestos", cnn, adOpenDynamic, adLockOptimistic
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

Private Sub cmd_imprimir_Click()
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call gPrintListView(lv_puestos, "LISTADO DE puestos")
        End If

End Sub

Private Sub cmd_nuevo_Click()
        Call pro_limpiatextos(Me)
        txt_puesto.Enabled = True
        txt_puesto.SetFocus: var_modifica_registro_puesto = False
        cmd_guardar.Enabled = True
        cmd_deshacer.Enabled = True
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_puesto = False Then
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
   frm_lista.Visible = False
   var_cadena_seguridad = ""
   Top = 0
   Left = 2900
   var_modifica_registro_puesto = True
   lv_puestos.SmallIcons = ImageList1
   Call pro_encabezadosView(Me, lv_puestos, False)
   Call pro_llena_listview1
   pro_textos
   rs.Open "select * from tb_puestos", cnn, adOpenDynamic, adLockOptimistic
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
   If var_despliega_menu = True Then
      var_swpassword = False
      var_modifica_registro_puesto = False
   End If
   Call activa_forma(var_activa_forma_puestos)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         txt_menu = lv_lista.selectedItem
         txt_nombre_menu = lv_lista.selectedItem.SubItems(1)
      Else
         txt_menu = ""
         txt_nombre_menu = ""
      End If
      txt_menu.SetFocus
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub lv_puestos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_puestos, ColumnHeader)
End Sub

Private Sub lv_puestos_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_puestos.selectedItem = Item
        pro_textos
        var_modifica_registro_puesto = True
        txt_puesto.Enabled = False

End Sub



Sub pro_guardar_puestos()
   Dim ok As Boolean
   Set TB_PUESTOS = New TB_PUESTOS
   Set TB_BITACORA_PUESTOS = New TB_BITACORA_PUESTOS
   ok = True
   If txt_puesto <> "" And txt_nombre_puesto <> "" And txt_menu <> "" Then
      If var_hubo_cambios Then
         rs.Open "select * from tb_PUESTOS where vcha_PUE_PUESTO_id = '" + txt_puesto + "'", cnn, adOpenDynamic, adLockOptimistic
         ok = TB_PUESTOS.Anadir(txt_puesto, txt_nombre_puesto, txt_menu, chk_supervisor)
         If ok Then
            bitacora = True
            If var_modifica_registro_puesto = False Then
               var_operacion_bitacora = "I"
               bitacora = TB_BITACORA_PUESTOS.Anadir(txt_puesto, "VCHA_PUE_DESCRIPCION", var_operacion_bitacora, "", txt_nombre_puesto, var_clave_usuario_global, fun_NombrePc, Date)
            Else
               var_operacion_bitacora = "M"
               If rs(0) <> txt_puesto Then
                  bitacora = TB_BITACORA_PUESTOS.Anadir(txt_puesto, "VCHA_PUE_PUESTO_ID", var_operacion_bitacora, rs(0), txt_puesto, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(1) <> txt_nombre_puesto Then
                  bitacora = TB_BITACORA_PUESTOS.Anadir(txt_puesto, "VCHA_PUE_DESCRIPCION", var_operacion_bitacora, rs(1), txt_nombre_puesto, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(2) <> txt_menu Then
                  bitacora = TB_BITACORA_PUESTOS.Anadir(txt_puesto, "VCHA_MEN_MENU_ID", var_operacion_bitacora, rs(0), txt_puesto, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(3) <> chk_supervisor Then
                  bitacora = TB_BITACORA_PUESTOS.Anadir(txt_puesto, "VCHA_PUE_SUPERVISOR", var_operacion_bitacora, rs(1), txt_nombre_puesto, var_clave_usuario_global, fun_NombrePc, Date)
               End If
            End If
            rs.Close
            pro_actualiza_ListView
            txt_puesto.Enabled = False
            MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
            txt_registros = lv_puestos.ListItems.Count
            var_modifica_registro_puesto = True
         Else
            MsgBox "No se puede grabar registro: " + TB_PUESTOS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
         End If
      End If
   End If
   Set TB_PUESTOS = Nothing: var_hubo_cambios = False
End Sub

Sub pro_elimina_puestos()
   Dim var_llave_usuarios As String
   Set TB_PUESTOS = New TB_PUESTOS
   Set TB_BITACORA_PUESTOS = New TB_BITACORA_PUESTOS
   On Error GoTo salir:
   ok = True
   If txt_puesto <> "" And txt_nombre_puesto <> "" And var_modifica_registro_puesto = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_PUESTOS.Eliminar(txt_puesto)
      Else
         GoTo salir:
      End If
      If ok Then
         var_operacion_bitacora = "E"
         bitacora = TB_BITACORA_PUESTOS.Anadir(txt_puesto, "VCHA_PUE_DESCRIPCION", var_operacion_bitacora, txt_nombre_puesto, "", var_clave_usuario_global, fun_NombrePc, Date)
         numero_items_puestos = numero_items_puestos - 1
         MsgBox "Se Elimino Correctamente el Registro", vbInformation
         lv_puestos.ListItems.Remove (lv_puestos.selectedItem.Index)
         Call pro_limpiatextos(Me)
         txt_registros = lv_puestos.ListItems.Count
         lv_puestos.selectedItem.Selected = True
         pro_textos
      Else
         MsgBox "No se puede eliminar registro: " + TB_PUESTOS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
   Set TB_PUESTOS = Nothing
End Sub


Sub pro_llena_listview1()
   Dim list_item As ListItem
   rs.Open "select * from TB_puestos", cnn, adOpenDynamic, adLockOptimistic
   numero_items_puestos = 0
   While Not rs.EOF
      Set list_item = lv_puestos.ListItems.Add(, , rs(0).Value)
      list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
      list_item.SubItems(2) = IIf(IsNull(rs(2).Value), "", rs(2).Value)
      list_item.SubItems(3) = IIf(IsNull(rs(3).Value), "", rs(3).Value)
      rs.MoveNext:
      numero_items_puestos = numero_items_puestos + 1
    Wend
    rs.Close
End Sub


Sub pro_textos()
On Error GoTo err0:
Dim var_n As Double
var_n = lv_puestos.ListItems.Count
   If var_n > 0 Then
      txt_puesto = lv_puestos.selectedItem
      txt_nombre_puesto = lv_puestos.selectedItem.SubItems(1)
      txt_menu = lv_puestos.selectedItem.SubItems(2)
      chk_supervisor = lv_puestos.selectedItem.SubItems(3)
   End If
   rs.Open "select * from tb_menus where vcha_men_menu_id = '" + txt_menu + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      txt_nombre_menu = IIf(IsNull(rs!vcha_men_descripcion), "", rs!vcha_men_descripcion)
   Else
      txt_nombre_menu = ""
   End If
   rs.Close
   var_numero_renglones = lv_puestos.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_puestos.ColumnHeaders(2).Width = 3850
   Else
      lv_puestos.ColumnHeaders(2).Width = 4099.71
   End If
   var_hubo_cambios = False
   var_modifica_registro_puesto = True
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem

    If var_modifica_registro_puesto = False Then
        Set list_item = lv_puestos.ListItems.Add(, , txt_puesto)
        list_item.SubItems(1) = txt_nombre_puesto
        list_item.SubItems(2) = txt_menu
        list_item.SubItems(3) = chk_supervisor
        list_item.EnsureVisible
        list_item.Selected = True
        numero_items_puestos = numero_items_puestos + 1
    Else
        lv_puestos.ListItems.Item(lv_puestos.selectedItem.Index).Checked = False
        
        lv_puestos.ListItems.Item(lv_puestos.selectedItem.Index) = txt_puesto
        lv_puestos.ListItems.Item(lv_puestos.selectedItem.Index).ListSubItems(1) = txt_nombre_puesto
        lv_puestos.ListItems.Item(lv_puestos.selectedItem.Index).ListSubItems(2) = txt_menu
        lv_puestos.ListItems.Item(lv_puestos.selectedItem.Index).ListSubItems(3) = chk_supervisor
        lv_puestos.ListItems.Item(lv_puestos.selectedItem.Index).Selected = True
    End If
'    lv_puestos.SetFocus
End Sub

Private Sub txt_puestos_Change(Index As Integer)
   var_hubo_cambios = True
End Sub

Private Sub txt_puestos_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
      var_hubo_cambios = True
   End If
End Sub


Private Sub txt_menu_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_menu_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_menus order by vcha_men_descripcion", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_men_menu_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_men_descripcion), "", rs!vcha_men_descripcion)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "MENUS"
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

Private Sub txt_menu_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_menu_LostFocus()
   If Trim(txt_menu) <> "" Then
      rs.Open "select * from tb_menus where vcha_men_menu_id = '" + txt_menu + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_menu = IIf(IsNull(rs!vcha_men_descripcion), "", rs!vcha_men_descripcion)
      Else
         MsgBox "Clave de menu incorrecto", vbOKOnly, "ATENCION"
         txt_menu = ""
         txt_nombre_menu = ""
      End If
      rs.Close
   Else
      txt_nombre_menu = ""
      txt_menu = ""
   End If
End Sub

Private Sub txt_nombre_menu_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_menu_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Me.cmd_guardar.Enabled = True Then
         Me.cmd_guardar.SetFocus
      End If
   End If
End Sub

Private Sub txt_nombre_puesto_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_puesto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_puesto_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_puesto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub
