VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmdescuentos_catalogos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descuentos por vencimiento de catálogos"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
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
      Left            =   150
      TabIndex        =   20
      Top             =   450
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   21
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
         TabIndex        =   22
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5055
      Left            =   150
      TabIndex        =   17
      Top             =   2160
      Width           =   5655
      Begin MSComctlLib.ListView lv_descuentos_catalogos 
         Height          =   4860
         Left            =   45
         TabIndex        =   18
         Top             =   150
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   8573
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Limite Inferior"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Limite Superior"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Porcentaje"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Clave Canal"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Descuentos por vencimiento de catalogos "
      Height          =   1710
      Left            =   150
      TabIndex        =   12
      Top             =   435
      Width           =   5655
      Begin VB.TextBox txt_nombre_canal_venta 
         Height          =   315
         Left            =   2385
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   270
         Width           =   3195
      End
      Begin VB.TextBox txt_canal_venta 
         Height          =   315
         Left            =   1455
         TabIndex        =   6
         Top             =   270
         Width           =   900
      End
      Begin VB.TextBox txt_limite_inferior 
         Height          =   315
         Left            =   1455
         MaxLength       =   10
         TabIndex        =   8
         Top             =   600
         Width           =   900
      End
      Begin VB.TextBox txt_limite_superior 
         Height          =   315
         Left            =   1455
         MaxLength       =   50
         TabIndex        =   9
         Top             =   930
         Width           =   900
      End
      Begin VB.TextBox txt_porcentaje 
         Height          =   315
         Left            =   1455
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1260
         Width           =   1620
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Canal de Venta:"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   16
         Top             =   330
         Width           =   1140
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Limite Inferior:"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   15
         Top             =   660
         Width           =   975
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Limite Superior:"
         Height          =   195
         Index           =   2
         Left            =   270
         TabIndex        =   14
         Top             =   990
         Width           =   1080
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Porcentaje:"
         Height          =   195
         Index           =   4
         Left            =   270
         TabIndex        =   13
         Top             =   1320
         Width           =   810
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   -30
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   11
      Top             =   1425
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmdescuentos_catalogos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   465
      Picture         =   "frmdescuentos_catalogos.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Guardar Alt + G"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   780
      Picture         =   "frmdescuentos_catalogos.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1110
      Picture         =   "frmdescuentos_catalogos.frx":02D6
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      Picture         =   "frmdescuentos_catalogos.frx":03D8
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5460
      Picture         =   "frmdescuentos_catalogos.frx":04DA
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   -105
      Top             =   450
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
            Picture         =   "frmdescuentos_catalogos.frx":0B14
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdescuentos_catalogos.frx":13EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   -120
      Top             =   1935
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
            Picture         =   "frmdescuentos_catalogos.frx":1CC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdescuentos_catalogos.frx":25A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdescuentos_catalogos.frx":2E7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdescuentos_catalogos.frx":3418
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdescuentos_catalogos.frx":3CF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdescuentos_catalogos.frx":45CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdescuentos_catalogos.frx":4EA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdescuentos_catalogos.frx":4FBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdescuentos_catalogos.frx":50CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdescuentos_catalogos.frx":51DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdescuentos_catalogos.frx":52F0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   19
      Top             =   285
      Width           =   5655
   End
End
Attribute VB_Name = "frmdescuentos_catalogos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim numero_items_descuentos_catalogos As Integer
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
      Call pro_elimina_descuentos_catalogos
      rs.Open "select * from tb_descuentos_catalogos", cnn, adOpenDynamic, adLockOptimistic
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
   If Not IsNumeric(Me.txt_limite_inferior) Then
      txt_limite_inferior = 0
   End If
   If Not IsNumeric(Me.txt_limite_superior) Then
      txt_limite_superior = 0
   End If
   If Not IsNumeric(Me.txt_porcentaje) Then
      txt_porcentaje = 0
   End If
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
      Call pro_guardar_descuentos_catalogos
      cmd_guardar.Enabled = False
      cmd_deshacer.Enabled = False
   End If
End Sub

Private Sub cmd_imprimir_Click()
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call gPrintListView(lv_descuentos_catalogos, "LISTADO DE descuentos_catalogos")
        End If
End Sub

Private Sub cmd_nuevo_Click()
   Call pro_limpiatextos(Me)
   txt_canal_venta.Enabled = True
   txt_canal_venta.SetFocus: var_modifica_registro = False
   cmd_guardar.Enabled = True
   cmd_deshacer.Enabled = True
   Me.txt_canal_venta.Enabled = True
   Me.txt_limite_inferior.Enabled = True
   Me.txt_limite_superior.Enabled = True
   Me.txt_nombre_canal_venta.Enabled = True
   Me.txt_porcentaje.Enabled = True
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro = False Then
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
   numero_items_descuentos_catalogos = 0
   var_modifica_registro = True
   lv_descuentos_catalogos.SmallIcons = ImageList1
   Call pro_encabezadosView(Me, lv_descuentos_catalogos, False)
   cmd_guardar.Enabled = False
   cmd_deshacer.Enabled = False
   rsaux.Open "select * from tb_descuentos_catalogos", cnn, adOpenDynamic, adLockOptimistic
   If Not rsaux.EOF Then
      txt_canal_venta = rsaux!vcha_can_canal_venta_id
      Call pro_llena_listview1
      pro_textos
   End If
   rsaux.Close
   frm_lista.Visible = False
   Me.txt_canal_venta.Enabled = False
   Me.txt_limite_inferior.Enabled = False
   Me.txt_limite_superior.Enabled = False
   Me.txt_nombre_canal_venta.Enabled = False
   Me.txt_porcentaje.Enabled = False
 End Sub

Private Sub Form_Unload(Cancel As Integer)
    var_swpassword = False
    var_modifica_registro = False
    Call activa_forma(var_activa_forma_descuentos_catalogos)
End Sub

Private Sub lv_descuentos_catalogos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_descuentos_catalogos, ColumnHeader)
End Sub

Private Sub lv_descuentos_catalogos_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_descuentos_catalogos.selectedItem = Item
        pro_textos
        var_modifica_registro = True
        txt_canal_venta.Enabled = True
End Sub



Sub pro_guardar_descuentos_catalogos()
   Dim ok As Boolean
   Set TB_DESCUENTOS_CATALOGOS = New TB_DESCUENTOS_CATALOGOS
   If txt_canal_venta <> "" And txt_limite_inferior <> "" And txt_limite_superior <> "" And txt_porcentaje <> "" Then
      ok = TB_DESCUENTOS_CATALOGOS.Anadir(txt_canal_venta, txt_limite_inferior, txt_limite_superior, txt_porcentaje)
      If ok Then
         pro_actualiza_ListView
         txt_canal_venta.Enabled = False
         MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
         txt_registros = lv_descuentos_catalogos.ListItems.Count
         var_modifica_registro = True
      Else
         MsgBox "No se puede grabar registro: " + TB_DESCUENTOS_CATALOGOS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
End Sub

Sub pro_elimina_descuentos_catalogos()
   Dim var_llave_usuarios As String
   Set TB_DESCUENTOS_CATALOGOS = New TB_DESCUENTOS_CATALOGOS
   On Error GoTo salir
   ok = True
   If txt_canal_venta <> "" And txt_limite_inferior <> "" And var_modifica_registro = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_DESCUENTOS_CATALOGOS.Eliminar(txt_canal_venta, txt_limite_inferior, txt_limite_superior, txt_porcentaje)
      Else
         GoTo salir:
      End If
      If ok Then
         MsgBox "Se Elimino Correctamente el Registro", vbInformation
         lv_descuentos_catalogos.ListItems.Remove (lv_descuentos_catalogos.selectedItem.Index)
         numero_items_descuentos_catalogos = numero_items_descuentos_catalogos - 1
         Call pro_limpiatextos(Me)
         txt_registros = lv_descuentos_catalogos.ListItems.Count
         lv_descuentos_catalogos.selectedItem.Selected = True
         pro_textos
      Else
         MsgBox "No se puede grabar registro: " + TB_DESCUENTOS_CATALOGOS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
   Set TB_DESCUENTOS_CATALOGOS = Nothing
End Sub


Sub pro_llena_listview1()
   Dim list_item As ListItem
   rs.Open "select * from TB_descuentos_catalogos where vcha_can_canal_venta_id = '" + txt_canal_venta + "'", cnn, adOpenDynamic, adLockOptimistic
   lv_descuentos_catalogos.ListItems.Clear
   While Not rs.EOF
      Set list_item = lv_descuentos_catalogos.ListItems.Add(, , rs!INTE_DES_LIMITE_INFERIOR)
      list_item.SubItems(1) = IIf(IsNull(rs!INTE_DES_LIMITE_SUPERIOR), "", rs!INTE_DES_LIMITE_SUPERIOR)
      list_item.SubItems(2) = IIf(IsNull(rs!FLOA_DES_DESCUENTO), "", rs!FLOA_DES_DESCUENTO)
      list_item.SubItems(3) = IIf(IsNull(rs!vcha_can_canal_venta_id), "", rs!vcha_can_canal_venta_id)
      rs.MoveNext:
      numero_items_descuentos_catalogos = numero_items_descuentos_catalogos + 1
   Wend
   rs.Close
   If numero_items_descuentos_catalogos > 11 Then
      lv_descuentos_catalogos.ColumnHeaders(3).Width = 1300.09
   Else
      lv_descuentos_catalogos.ColumnHeaders(3).Width = 1500.09
   End If
   var_n = lv_descuentos_catalogos.ListItems.Count
   var_modifica_registro = True
End Sub


Sub pro_textos()
'On Error GoTo err0:
        txt_limite_inferior = lv_descuentos_catalogos.selectedItem
        txt_limite_superior = lv_descuentos_catalogos.selectedItem.SubItems(1)
        txt_porcentaje = lv_descuentos_catalogos.selectedItem.SubItems(2)
        txt_canal_venta = lv_descuentos_catalogos.selectedItem.SubItems(3)
        rs.Open "select * from tb_canalesventas where vcha_can_canal_venta_id = '" + txt_canal_venta + "'", cnn, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
           txt_nombre_canal_venta = IIf(IsNull(rs!vcha_can_nombre), "", rs!vcha_can_nombre)
        Else
           txt_nombre_canal_venta = ""
        End If
        rs.Close
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro = True Then
        Set list_item = lv_descuentos_catalogos.ListItems.Add(, , txt_limite_inferior)
        list_item.SubItems(1) = txt_limite_superior
        list_item.SubItems(2) = txt_porcentaje
        list_item.SubItems(3) = txt_canal_venta
        list_item.EnsureVisible
        list_item.Selected = True
       numero_items_descuentos_catalogos = numero_items_descuentos_catalogos + 1
    Else
        lv_descuentos_catalogos.ListItems.Item(lv_descuentos_catalogos.selectedItem.Index).Checked = False
        lv_descuentos_catalogos.ListItems.Item(lv_descuentos_catalogos.selectedItem.Index) = txt_limite_inferior
        lv_descuentos_catalogos.ListItems.Item(lv_descuentos_catalogos.selectedItem.Index).ListSubItems(1) = txt_limite_superior
        lv_descuentos_catalogos.ListItems.Item(lv_descuentos_catalogos.selectedItem.Index).ListSubItems(2) = txt_porcentaje
        lv_descuentos_catalogos.ListItems.Item(lv_descuentos_catalogos.selectedItem.Index).ListSubItems(3) = txt_canal_venta
        lv_descuentos_catalogos.ListItems.Item(lv_descuentos_catalogos.selectedItem.Index).Selected = True
    End If
    If numero_items_descuentos_catalogos > 11 Then
       lv_descuentos_catalogos.ColumnHeaders(3).Width = 1300
    Else
       lv_descuentos_catalogos.ColumnHeaders(3).Width = 1500.09
    End If
    lv_descuentos_catalogos.SetFocus
End Sub





Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         txt_canal_venta = lv_lista.selectedItem
         txt_nombre_canal_venta = lv_lista.selectedItem.SubItems(1)
      Else
         txt_canal_venta = ""
         txt_nombre_canal_venta = ""
      End If
      txt_canal_venta.SetFocus
      frm_lista.Visible = False
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub txt_canal_venta_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_canal_venta_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_canalesventas order by vcha_can_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_can_canal_venta_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_can_nombre), "", rs!vcha_can_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CANALES DE VENTA"
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
      frmcanalesventas.Show
   End If
End Sub

Private Sub txt_canal_venta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_canal_venta_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_canal_venta) <> "" Then
      rsaux.Open "select * from TB_CANALESVENTAS where vcha_can_canal_venta_id = '" + txt_canal_venta + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux.EOF Then
         txt_nombre_canal_venta = IIf(IsNull(rsaux!vcha_can_nombre), "", rsaux!vcha_can_nombre)
         Call pro_llena_listview1
      Else
         txt_canal_venta = ""
         txt_nombre_canal_venta = ""
      End If
      rsaux.Close
   Else
      cmb_canales_venta = ""
   End If
End Sub

Private Sub txt_limite_inferior_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_limite_inferior_LostFocus()
   If Not IsNumeric(txt_limite_inferior) Then
      MsgBox "Limite inferior incorrecto", vbOKOnly, "ATENCION"
      txt_limite_inferior = 0
   End If
End Sub

Private Sub txt_limite_superior_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_limite_superior_LostFocus()
   If Not IsNumeric(txt_limite_superior) Then
      MsgBox "Limite superior incorrecto", vbOKOnly, "ATENCION"
      txt_limite_superior = 0
   End If
End Sub

Private Sub txt_nombre_canal_venta_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_nombre_canal_venta_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_canalesventas order by vcha_can_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_can_canal_venta_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_can_nombre), "", rs!vcha_can_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CANALES DE VENTA"
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
      frmcanalesventas.Show
   End If
End Sub

Private Sub txt_nombre_canal_venta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_canal_venta_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_porcentaje_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If Me.cmd_guardar.Enabled = True Then
         Me.cmd_guardar.SetFocus
      End If
   End If
End Sub

Private Sub txt_porcentaje_LostFocus()
   If Not IsNumeric(txt_porcentaje) Then
      MsgBox "Porcentaje incorrecto", vbOKOnly, "ATENCION"
      txt_porcentaje = 0
   End If
End Sub
