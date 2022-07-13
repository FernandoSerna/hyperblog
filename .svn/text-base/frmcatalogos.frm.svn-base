VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcatalogos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogos"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   Icon            =   "frmcatalogos.frx":0000
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
      Left            =   210
      TabIndex        =   23
      Top             =   2805
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   24
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
         TabIndex        =   25
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmcatalogos.frx":08CA
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
      Picture         =   "frmcatalogos.frx":09CC
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
      Picture         =   "frmcatalogos.frx":0ACE
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
      Picture         =   "frmcatalogos.frx":0BA0
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
      Picture         =   "frmcatalogos.frx":0CA2
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
      Picture         =   "frmcatalogos.frx":0DA4
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2910
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   16
      Top             =   30
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   " Catálogos "
      Height          =   1725
      Left            =   150
      TabIndex        =   0
      Top             =   405
      Width           =   5655
      Begin VB.CheckBox chk_vigente 
         Caption         =   "Vigente"
         Height          =   315
         Left            =   2595
         TabIndex        =   27
         Top             =   1320
         Width           =   1635
      End
      Begin VB.CheckBox chk_catalogo 
         Caption         =   "Catálogo"
         Height          =   315
         Left            =   1260
         TabIndex        =   26
         Top             =   1320
         Width           =   1110
      End
      Begin VB.TextBox txt_nombre_agrupador 
         Height          =   315
         Left            =   2415
         MaxLength       =   50
         TabIndex        =   22
         Top             =   960
         Width           =   3045
      End
      Begin VB.TextBox txt_agrupador 
         Height          =   315
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   20
         Top             =   960
         Width           =   1110
      End
      Begin VB.CommandButton cmd_vigencias 
         Height          =   315
         Left            =   5100
         Picture         =   "frmcatalogos.frx":13DE
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Vigencias por canal de venta"
         Top             =   615
         Width           =   330
      End
      Begin VB.TextBox txt_catalogos 
         Height          =   315
         Index           =   1
         Left            =   1275
         MaxLength       =   50
         TabIndex        =   8
         Top             =   615
         Width           =   3795
      End
      Begin VB.TextBox txt_catalogos 
         Height          =   315
         Index           =   0
         Left            =   1275
         TabIndex        =   7
         Top             =   270
         Width           =   1170
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Agrupador:"
         Height          =   195
         Index           =   0
         Left            =   345
         TabIndex        =   21
         Top             =   1020
         Width           =   780
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Index           =   2
         Left            =   375
         TabIndex        =   10
         Top             =   675
         Width           =   885
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   1
         Left            =   375
         TabIndex        =   9
         Top             =   315
         Width           =   450
      End
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   165
      Top             =   4935
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
            Picture         =   "frmcatalogos.frx":2650
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcatalogos.frx":2F2A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   14
      Top             =   285
      Width           =   5655
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   150
      TabIndex        =   11
      Top             =   2145
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   1890
         TabIndex        =   15
         Top             =   150
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3795
         TabIndex        =   18
         Top             =   150
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ir al primero"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Un Registro Atras"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Un registro adelante"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ir al ultimo"
               ImageIndex      =   10
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Busqueda de catalogo:"
         Height          =   195
         Left            =   195
         TabIndex        =   12
         Top             =   195
         Width           =   1650
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4470
      Left            =   150
      TabIndex        =   13
      Top             =   2700
      Width           =   5655
      Begin MSComctlLib.ListView lv_catalogos 
         Height          =   4245
         Left            =   45
         TabIndex        =   17
         Top             =   150
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   7488
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
         NumItems        =   8
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
            Text            =   "vigente"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "fecha inicio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "fecha fin"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "AGRUPADOR"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Catalogo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "vigente"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   45
         Top             =   2505
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
               Picture         =   "frmcatalogos.frx":3804
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcatalogos.frx":40DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcatalogos.frx":49B8
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcatalogos.frx":4F54
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcatalogos.frx":5830
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcatalogos.frx":610A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcatalogos.frx":69E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcatalogos.frx":6AF6
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcatalogos.frx":6C08
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcatalogos.frx":6D1A
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmcatalogos.frx":6E2C
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmcatalogos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim numero_items_catalogos As Integer
Dim bitacora As Boolean






Private Sub chk_catalogo_Click()
   var_hubo_cambios = True
End Sub

Private Sub chk_vigente_Click()
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
      Call pro_elimina_catalogos
      rs.Open "select * from tb_catalogos", cnn, adOpenDynamic, adLockOptimistic
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
   Else
      MsgBox "Imposible realizar la acción solicitada", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_guardar_Click()
Dim var_posible As Boolean
   var_posible = True
   If var_modifica_registro_catalogo = False Then
      rs.Open "SELECT * FROM TB_CATALOGOS WHERE VCHA_CAT_CATALOGO_ID = '" + Me.txt_catalogos(0) + "'", cnn, adOpenDynamic, adLockOptimistic
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
         Call pro_guardar_catalogos
         rs.Open "select * from tb_catalogos", cnn, adOpenDynamic, adLockOptimistic
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
      Else
         MsgBox "Imposible realizar la acción solicitada", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Clave de catálogo ya existe", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
      If vector_valida_passwords(var_indice_menu) = "*" Then
         frmpasswords.Show
      Else
         Call gPrintListView(lv_catalogos, "LISTADO DE catalogos")
      End If

End Sub

Private Sub cmd_nuevo_Click()
   Call pro_limpiatextos(Me)
   txt_catalogos(0).Enabled = True
   txt_catalogos(0).SetFocus: var_modifica_registro_catalogo = False
   cmd_guardar.Enabled = True
   cmd_deshacer.Enabled = True
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_catalogo = False Then
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

Private Sub cmd_vigencias_Click()
   If Trim(txt_catalogos(1)) <> "" Then
      frmvigencias_catalogo_canal_venta.txt_catalogo = txt_catalogos(0)
      frmvigencias_catalogo_canal_venta.Caption = txt_catalogos(1)
      frmvigencias_catalogo_canal_venta.Show 1
   Else
      MsgBox "No se a seleccionado un catálogo", vbOKOnly, "ATENCION"
   End If
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
    var_modifica_registro_catalogo = True
    'lv_catalogos.SmallIcons = ImageList1
    'Call pro_encabezadosView(Me, lv_catalogos, False)
    Call pro_llena_listview1
    pro_textos

    'Call pro_AsignarAViewColor(lv_catalogos, Picture1, vbWhite, vbGray)
    rs.Open "select * from tb_catalogos", cnn, adOpenDynamic, adLockOptimistic
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
    Me.txt_catalogos(0).Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Top = 0
   Left = 2900
   If var_despliega_menu = True Then
      var_swpassword = False
      var_modifica_registro_catalogo = False
   End If
   Call activa_forma(var_activa_forma_catalogos)
End Sub

Private Sub lv_catalogos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_catalogos, ColumnHeader)
End Sub

Private Sub lv_catalogos_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Set lv_catalogos.selectedItem = Item
   pro_textos
   var_modifica_registro_catalogo = True
   txt_catalogos(0).Enabled = False
End Sub



Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         txt_agrupador = lv_lista.selectedItem
         txt_nombre_agrupador = lv_lista.selectedItem.SubItems(1)
      End If
      txt_agrupador.SetFocus
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_catalogos.SetFocus
      Call pro_avanzar(Me, lv_catalogos, Button)
      lv_catalogos.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_catalogos.ListItems(1).Selected = True
      lv_catalogos.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 4 Then
      numero_items_catalogos = lv_catalogos.ListItems.Count
      lv_catalogos.ListItems(numero_items_catalogos).Selected = True
      lv_catalogos.selectedItem.EnsureVisible
      pro_textos
   End If
err0:
End Sub


Sub pro_guardar_catalogos()

Dim ok As Boolean

Set TB_CATALOGOS = New TB_CATALOGOS
Set TB_BITACORA_CATALOGOS = New TB_BITACORA_CATALOGOS
    
    ok = True
    If txt_catalogos(0) <> "" And txt_catalogos(1) <> "" Then
        If var_hubo_cambios Then
            rs.Open "select * from tb_catalogos where vcha_cat_catalogo_id = '" + txt_catalogos(0).Text + "'", cnn, adOpenDynamic, adLockOptimistic
            ok = TB_CATALOGOS.Anadir(txt_catalogos(0), txt_catalogos(1), txt_agrupador)
            If ok Then
                bitacora = True
                rsaux10.Open "update tb_Catalogos set inte_cat_catalogo = " + CStr(Me.chk_catalogo.Value) + ", inte_cat_vigente = " + CStr(Me.chk_vigente.Value) + " where vcha_cat_catalogo_id = '" + Me.txt_catalogos(0) + "'", cnn, adOpenDynamic, adLockOptimistic
                If var_modifica_registro_catalogo = False Then
                   var_operacion_bitacora = "I"
                   bitacora = TB_BITACORA_CATALOGOS.Anadir(txt_catalogos(0), "VCHA_CAT_NOMBRE", var_operacion_bitacora, "", txt_catalogos(1), var_clave_usuario_global, fun_NombrePc, Date)
                Else
                   var_operacion_bitacora = "M"
                   If rs(0) <> txt_catalogos(0) Then
                      bitacora = TB_BITACORA_CATALOGOS.Anadir(txt_catalogos(0), "VCHA_CAT_CATALOGO_ID", var_operacion_bitacora, rs(0), txt_catalogos(0), var_clave_usuario_global, fun_NombrePc, Date)
                   End If
                   If rs(1) <> txt_catalogos(1) Then
                      bitacora = TB_BITACORA_CATALOGOS.Anadir(txt_catalogos(0), "VCHA_CAT_NOMBRE", var_operacion_bitacora, rs(1), txt_catalogos(1), var_clave_usuario_global, fun_NombrePc, Date)
                   End If
                End If
                rs.Close
                pro_actualiza_ListView
                txt_catalogos(0).Enabled = False
                MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
                txt_registros = lv_catalogos.ListItems.Count
                var_modifica_registro_catalogo = True
            Else
                MsgBox "No se puede grabar registro: " + TB_CATALOGOS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
        End If
    End If
    
Set TB_CATALOGOS = Nothing: var_hubo_cambios = False

End Sub

Sub pro_elimina_catalogos()
   Dim var_llave_usuarios As String
   Set TB_CATALOGOS = New TB_CATALOGOS
   Set TB_BITACORA_CATALOGOS = New TB_BITACORA_CATALOGOS
   On Error GoTo salir:
   ok = True
   If txt_catalogos(0) <> "" And txt_catalogos(1) <> "" And var_modifica_registro_catalogo = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_CATALOGOS.Eliminar(txt_catalogos(0))
      Else
         GoTo salir:
      End If
      If ok Then
         var_operacion_bitacora = "E"
         bitacora = TB_BITACORA_CATALOGOS.Anadir(txt_catalogos(0), "VCHA_CAT_NOMBRE", var_operacion_bitacora, "", txt_catalogos(1), var_clave_usuario_global, fun_NombrePc, Date)
         numero_items_catalogos = numero_items_catalogos - 1
         MsgBox "Se Elimino Correctamente el Registro", vbInformation
         lv_catalogos.ListItems.Remove (lv_catalogos.selectedItem.Index)
         Call pro_limpiatextos(Me)
         txt_registros = lv_catalogos.ListItems.Count
         lv_catalogos.selectedItem.Selected = True
         pro_textos
      Else
         MsgBox "No se puede eliminar registro: " + TB_CATALOGOS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
   Set TB_CATALOGOS = Nothing
End Sub


Sub pro_llena_listview1()
   Dim list_item As ListItem
   rs.Open "select * from TB_catalogos", cnn, adOpenDynamic, adLockOptimistic
   numero_items_catalogos = 0
   While Not rs.EOF
      Set list_item = lv_catalogos.ListItems.Add(, , rs!vcha_cat_catalogo_id)
      list_item.SubItems(1) = IIf(IsNull(rs!VCHA_cat_NOMBRE), "", rs!VCHA_cat_NOMBRE)
      list_item.SubItems(5) = IIf(IsNull(rs!vcha_agr_agrupador_catalogo_id), "", rs!vcha_agr_agrupador_catalogo_id)
      list_item.SubItems(6) = IIf(IsNull(rs!inte_cat_catalogo), 0, rs!inte_cat_catalogo)
      list_item.SubItems(7) = IIf(IsNull(rs!inte_cat_vigente), 0, rs!inte_cat_vigente)
      rs.MoveNext:
      numero_items_catalogos = numero_items_catalogos + 1
   Wend
   rs.Close
End Sub


Sub pro_textos()
On Error GoTo err0:
Dim var_n As Double
   var_n = lv_catalogos.ListItems.Count
   If var_n > 0 Then
      txt_catalogos(0) = lv_catalogos.selectedItem
      txt_catalogos(1) = lv_catalogos.selectedItem.SubItems(1)
      txt_agrupador = lv_catalogos.selectedItem.SubItems(5)
      Me.chk_catalogo = lv_catalogos.selectedItem.SubItems(6)
      Me.chk_vigente = lv_catalogos.selectedItem.SubItems(7)
   End If
   rs.Open "SELECT * FROM TB_AGRUPADORES_CATALOGOS WHERE VCHA_AGR_AGRUPADOR_CATALOGO_ID = '" + txt_agrupador + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      txt_nombre_agrupador = IIf(IsNull(rs!vcha_agr_nombre), "", rs!vcha_agr_nombre)
   Else
      txt_agrupador = ""
      txt_nombre_agrupador = ""
   End If
   rs.Close
   var_numero_renglones = lv_catalogos.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_catalogos.ColumnHeaders(2).Width = 3850
   Else
      lv_catalogos.ColumnHeaders(2).Width = 4099.9
   End If
   var_modifica_registro_catalogo = True
   var_hubo_cambios = False
   Me.txt_catalogos(0).Enabled = False
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem

    If var_modifica_registro_catalogo = False Then
        Set list_item = lv_catalogos.ListItems.Add(, , txt_catalogos(0))
        list_item.SubItems(1) = txt_catalogos(1)
        list_item.SubItems(5) = txt_agrupador
        list_item.SubItems(6) = Me.chk_catalogo.Value
        list_item.SubItems(7) = Me.chk_vigente
        list_item.EnsureVisible
        list_item.Selected = True
        numero_items_catalogos = numero_items_catalogos + 1
    Else
        lv_catalogos.ListItems.Item(lv_catalogos.selectedItem.Index).Checked = False
        lv_catalogos.ListItems.Item(lv_catalogos.selectedItem.Index) = txt_catalogos(0)
        lv_catalogos.ListItems.Item(lv_catalogos.selectedItem.Index).ListSubItems(1) = txt_catalogos(1)
        lv_catalogos.ListItems.Item(lv_catalogos.selectedItem.Index).ListSubItems(5) = txt_agrupador
        lv_catalogos.ListItems.Item(lv_catalogos.selectedItem.Index).ListSubItems(6) = Me.chk_catalogo.Value
        lv_catalogos.ListItems.Item(lv_catalogos.selectedItem.Index).ListSubItems(7) = Me.chk_vigente
        lv_catalogos.ListItems.Item(lv_catalogos.selectedItem.Index).Selected = True
    End If
'    lv_catalogos.SetFocus
End Sub

Private Sub txt_agrupador_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_agrupador_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_agrupadores_catalogos order by VCHA_AGR_NOMBRE", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_agr_agrupador_catalogo_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_agr_nombre), "", rs!vcha_agr_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGRUPADORES DE CATALOGOS"
      var_tipo_lista = 9
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

Private Sub txt_agrupador_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
 End Sub

Private Sub txt_agrupador_LostFocus()
   If Trim(txt_agrupador) <> "" Then
      rs.Open "SELECT * FROM TB_AGRUPADORES_CATALOGOS WHERE VCHA_AGR_AGRUPADOR_CATALOGO_ID = '" + txt_agrupador + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_agrupador = IIf(IsNull(rs!vcha_agr_nombre), "", rs!vcha_agr_nombre)
      Else
         txt_agrupador = ""
         txt_nombre_agrupador = ""
         MsgBox "Nombre de agrupador incorrecto", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      txt_nombre_agrupador = ""
   End If
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_catalogos, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub

Private Sub txt_catalogos_Change(Index As Integer)
   var_hubo_cambios = True
End Sub

Private Sub txt_catalogos_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Index < 1 Then
         Call pro_enfoque(KeyAscii)
      Else
         If Me.cmd_guardar.Enabled = True Then
           Me.cmd_guardar.SetFocus
         End If
     End If
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
      var_hubo_cambios = True
   End If
End Sub

Private Sub txt_nombre_agrupador_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_agrupador_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      cmd_guardar.SetFocus
   End If
   
End Sub
