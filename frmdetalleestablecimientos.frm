VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmdetalle_establecimientos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de establecimientos"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   Icon            =   "frmdetalleestablecimientos.frx":0000
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
      Left            =   120
      TabIndex        =   14
      Top             =   2025
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   15
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
            Object.Width           =   2646
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
         TabIndex        =   16
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frmdetalleestablecimientos.frx":08CA
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
      Picture         =   "frmdetalleestablecimientos.frx":09CC
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
      Picture         =   "frmdetalleestablecimientos.frx":0ACE
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
      Picture         =   "frmdetalleestablecimientos.frx":0BA0
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
      Picture         =   "frmdetalleestablecimientos.frx":0CA2
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5415
      Picture         =   "frmdetalleestablecimientos.frx":0DA4
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   45
      Width           =   330
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2775
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   13
      Top             =   90
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   " Detalle de establecimientos "
      Height          =   810
      Left            =   150
      TabIndex        =   0
      Top             =   435
      Width           =   5655
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   315
         Left            =   2220
         TabIndex        =   8
         Top             =   300
         Width           =   3360
      End
      Begin VB.TextBox txt_cliente 
         Height          =   315
         Left            =   705
         MaxLength       =   50
         TabIndex        =   7
         Top             =   300
         Width           =   1500
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   345
         Width           =   525
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5985
      Left            =   135
      TabIndex        =   10
      Top             =   1215
      Width           =   5670
      Begin MSComctlLib.ImageList icono_encabezado 
         Left            =   5475
         Top             =   765
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
               Picture         =   "frmdetalleestablecimientos.frx":13DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdetalleestablecimientos.frx":1CB8
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lv_detalle_establecimientos 
         Height          =   5715
         Left            =   45
         TabIndex        =   12
         Top             =   195
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   10081
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave cliente"
            Object.Width           =   2559
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre Cliente"
            Object.Width           =   6879
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   45
      TabIndex        =   11
      Top             =   285
      Width           =   5685
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
            Picture         =   "frmdetalleestablecimientos.frx":2592
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdetalleestablecimientos.frx":2E6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdetalleestablecimientos.frx":3746
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdetalleestablecimientos.frx":3CE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdetalleestablecimientos.frx":45BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdetalleestablecimientos.frx":4E98
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdetalleestablecimientos.frx":5772
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdetalleestablecimientos.frx":5884
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdetalleestablecimientos.frx":5996
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdetalleestablecimientos.frx":5AA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdetalleestablecimientos.frx":5BBA
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmdetalle_establecimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
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
      Call pro_elimina_detalle_establecimientos
      rs.Open "select * from tb_detalle_establecimientos", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If rs.BOF Then
         cmd_guardar.Enabled = False
         cmd_deshacer.Enabled = False
         cmd_eliminar.Enabled = False
      Else
         cmd_guardar.Enabled = True
         cmd_deshacer.Enabled = True
         cmd_eliminar.Enabled = True
      End If
      Me.txt_nombre_cliente.Enabled = False
      rs.Close
   End If
End Sub

Private Sub cmd_guardar_Click()
Dim var_posible As Boolean
   var_posible = True
   If var_modifica_registro_detalle_establecimiento = False Then
      rs.Open "select a.VCHA_CLI_CLAVE_ID,b.VCHA_CLI_NOMBRE from TB_DETALLE_establecimientos a, TB_clientes b where b.vcha_tit_titular_id = '" + vartitular + "' and a.VCHA_ESB_ESTABLECIMIENTO_ID = '" & varestablecimiento & "' and a.VCHA_CLI_CLAVE_ID = b.VCHA_CLI_CLAVE_ID and a.vcha_cli_clave_id = '" + txt_cliente + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
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
         Call pro_guardar_detalle_establecimientos
         rs.Open "select * from tb_detalle_establecimientos", cnn_distribucion, adOpenDynamic, adLockOptimistic
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
     MsgBox "Clave de cliente ya existe", vbOKOnly, "ATENCION"
  End If
End Sub

Private Sub cmd_imprimir_Click()
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call gPrintListView(lv_detalle_establecimientos, "LISTADO DE detalle_establecimientos")
        End If

End Sub

Private Sub cmd_nuevo_Click()
   Call pro_limpiatextos(Me)
   txt_cliente.Enabled = True
   txt_nombre_cliente.Enabled = True
   txt_cliente.SetFocus: var_modifica_registro_detalle_establecimiento = False
   cmd_guardar.Enabled = True
   cmd_deshacer.Enabled = True
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_detalle_establecimiento = False Then
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
   varestablecimiento = frmestablecimientos.txt_establecimiento
   var_modifica_registro_detalle_establecimiento = True
   Call pro_encabezadosView(Me, lv_detalle_establecimientos, False)
   Call pro_llena_listview1
   pro_textos
   rs.Open "select a.vcha_cli_clave_id,b.vcha_cli_nombre from TB_DETALLE_establecimientos a, TB_clientes b where a.vcha_esb_establecimiento_id = '" & varestablecimiento & "' and a.vcha_cli_clave_id = b.vcha_cli_clave_id", cnn_distribucion, adOpenDynamic, adLockOptimistic
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
    var_swpassword = False
    var_modifica_registro_detalle_establecimiento = False
    Call activa_forma(var_activa_forma_detalle_establecimientos)
End Sub

Private Sub lv_detalle_establecimientos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_detalle_establecimientos, ColumnHeader)
End Sub

Private Sub lv_detalle_establecimientos_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_detalle_establecimientos.selectedItem = Item
        pro_textos
        var_modifica_registro_detalle_establecimiento = True
        txt_cliente.Enabled = False

End Sub



Sub pro_guardar_detalle_establecimientos()
   Dim ok As Boolean
   Dim varaceptar As Boolean
   Dim varmensaje As String
   varaceptar = True
   If Trim(Me.txt_cliente) <> "" Then
      If var_modifica_registro_detalle_establecimiento = False Then
         rs.Open "select * from TB_DETALLE_ESTABLECIMIENTOS where  VCHA_CLI_CLAVE_ID = '" & txt_cliente & "' and VCHA_ESB_ESTABLECIMIENTO_ID = '" + varestablecimiento + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            varaceptar = False
         End If
         rs.Close
      End If
      If varaceptar = True Then
         rs.Open "select * from TB_clientes where  VCHA_CLI_CLAVE_ID = '" & txt_cliente & "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
         vardetallecliente = rs(1).Value
         rs.Close
         Set TB_DETALLE_ESTABLECIMIENTOS = New TB_DETALLE_ESTABLECIMIENTOS
         ok = True
         If txt_cliente <> "" Then
            If var_hubo_cambios Then
               ok = TB_DETALLE_ESTABLECIMIENTOS.Anadir(txt_cliente, varestablecimiento)
               If ok Then
                  pro_actualiza_ListView
                  txt_cliente.Enabled = False
                  MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
                  txt_registros = lv_detalle_establecimientos.ListItems.Count
                   var_modifica_registro_detalle_establecimiento = True
               Else
                   MsgBox "No se puede grabar registro: " + TB_DETALLE_ESTABLECIMIENTOS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
               End If
           End If
       End If
    Else
       varmensaje = "Ya existe una relación entre el cliente " + cmb_detalle_establecimientos + " y el establecimiento " + frmestablecimientos.txt_establecimiento
       MsgBox varmensaje, vbOKOnly, "ATENCION"
    End If
  Else
      MsgBox "No se a seleccionado un cliente", vbOKOnly, "ATENCION"
  End If
Set TB_DETALLE_ESTABLECIMIENTOS = Nothing: var_hubo_cambios = False

End Sub

Sub pro_elimina_detalle_establecimientos()
   Dim var_llave_usuarios As String
   Set TB_DETALLE_ESTABLECIMIENTOS = New TB_DETALLE_ESTABLECIMIENTOS
   On Error GoTo salir:
   ok = True
   If txt_cliente <> "" And var_modifica_registro_detalle_establecimiento = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_DETALLE_ESTABLECIMIENTOS.Eliminar(txt_cliente, varestablecimiento)
      Else
         GoTo salir:
      End If
      If ok Then
        MsgBox "Se Elimino Correctamente el Registro", vbInformation
        lv_detalle_establecimientos.ListItems.Remove (lv_detalle_establecimientos.selectedItem.Index)
        Call pro_limpiatextos(Me)
        txt_registros = lv_detalle_establecimientos.ListItems.Count
        lv_detalle_establecimientos.selectedItem.Selected = True
        pro_textos
      Else
        MsgBox "No se puede eliminar registro: " + TB_DETALLE_ESTABLECIMIENTOS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
   Set TB_DETALLE_ESTABLECIMIENTOS = Nothing
End Sub


Sub pro_llena_listview1()
   Dim list_item As ListItem
   
   rs.Open "select a.VCHA_CLI_CLAVE_ID,b.VCHA_CLI_NOMBRE from TB_DETALLE_establecimientos a, TB_clientes b where b.vcha_tit_titular_id = '" + vartitular + "' and a.VCHA_ESB_ESTABLECIMIENTO_ID = '" & varestablecimiento & "' and a.VCHA_CLI_CLAVE_ID = b.VCHA_CLI_CLAVE_ID", cnn_distribucion, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
      Set list_item = lv_detalle_establecimientos.ListItems.Add(, , rs!vcha_cli_clave_id)
      list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
      rs.MoveNext:
    Wend
    rs.Close
End Sub


Sub pro_textos()
On Error GoTo err0:
   txt_cliente = lv_detalle_establecimientos.selectedItem
   rs.Open "SELECT * FROM TB_CLIENTES WHERE VCHA_CLI_CLAVE_ID = '" + txt_cliente + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      txt_nombre_cliente = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
   Else
      txt_nombre_cliente = ""
   End If
   rs.Close
   txt_cliente.Enabled = False
   txt_nombre_cliente.Enabled = False
   var_modifica_registro_detalle_establecimiento = True
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem

    If var_modifica_registro_detalle_establecimiento = False Then
        Set list_item = lv_detalle_establecimientos.ListItems.Add(, , txt_cliente)
        list_item.SubItems(1) = txt_nombre_cliente
        list_item.EnsureVisible
        list_item.Selected = True
    Else
        lv_detalle_establecimientos.ListItems.Item(lv_detalle_establecimientos.selectedItem.Index).Checked = False
        lv_detalle_establecimientos.ListItems.Item(lv_detalle_establecimientos.selectedItem.Index) = txt_cliente
        lv_detalle_establecimientos.ListItems.Item(lv_detalle_establecimientos.selectedItem.Index).ListSubItems(1) = txt_nombre_cliente
        lv_detalle_establecimientos.ListItems.Item(lv_detalle_establecimientos.selectedItem.Index).Selected = True
    End If
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         txt_cliente = lv_lista.selectedItem
         txt_nombre_cliente = lv_lista.selectedItem.SubItems(1)
      Else
         txt_cliente = ""
         txt_nombre_cliente = ""
      End If
      txt_cliente.SetFocus
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub


Private Sub txt_cliente_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_cliente_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_clientes where vcha_tit_titular_id = '" + vartitular + "' order by vcha_cli_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CLIENTES"
      var_tipo_lista = 2
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3770.71
      Else
         lv_lista.ColumnHeaders(2).Width = 3999.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub


Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_cliente_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_cliente) <> "" Then
      rs.Open "select * from tb_clientes where vcha_cli_clave_id = '" + txt_cliente + "' and vcha_tit_titular_id = '" + vartitular + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_cliente = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
      Else
         txt_cliente = ""
         txt_nombre_cliente = ""
         MsgBox "Clave de cliente incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      txt_nombre_cliente = ""
   End If
End Sub

Private Sub txt_nombre_cliente_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_nombre_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_clientes order by vcha_cli_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CLIENTES"
      var_tipo_lista = 2
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3770.71
      Else
         lv_lista.ColumnHeaders(2).Width = 3999.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
       KeyAscii = 0
    Else
      If Me.cmd_guardar.Enabled = True Then
         Me.cmd_guardar.SetFocus
      End If
    End If
End Sub

Private Sub txt_nombre_cliente_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub
