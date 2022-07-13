VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmunidadesorganizacionales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plantas"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   Icon            =   "frmunidadesorganizacionales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5955
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   135
      TabIndex        =   25
      Top             =   435
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   26
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
         TabIndex        =   27
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5460
      Picture         =   "frmunidadesorganizacionales.frx":08CA
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
      Picture         =   "frmunidadesorganizacionales.frx":0F04
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
      Picture         =   "frmunidadesorganizacionales.frx":1006
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
      Picture         =   "frmunidadesorganizacionales.frx":1108
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
      Picture         =   "frmunidadesorganizacionales.frx":11DA
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
      Picture         =   "frmunidadesorganizacionales.frx":12DC
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Caption         =   " Plantas "
      Height          =   1665
      Left            =   150
      TabIndex        =   14
      Top             =   405
      Width           =   5655
      Begin VB.TextBox txt_nombre_empresa 
         Height          =   315
         Left            =   2205
         TabIndex        =   8
         Top             =   255
         Width           =   3360
      End
      Begin VB.CommandButton cmd_series 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5190
         Picture         =   "frmunidadesorganizacionales.frx":13DE
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Series"
         Top             =   1245
         Width           =   330
      End
      Begin VB.TextBox txt_empresa 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   7
         Top             =   255
         Width           =   900
      End
      Begin VB.TextBox txt_unidad 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   9
         Top             =   585
         Width           =   900
      End
      Begin VB.TextBox txt_nombre_unidad 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   10
         Text            =   " "
         Top             =   915
         Width           =   4260
      End
      Begin VB.TextBox txt_email 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1245
         Width           =   3885
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Empresa:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   18
         Top             =   315
         Width           =   660
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   17
         Top             =   645
         Width           =   450
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   16
         Top             =   975
         Width           =   600
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "E-mail:"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   15
         Top             =   1305
         Width           =   465
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5220
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   30
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
            Picture         =   "frmunidadesorganizacionales.frx":14E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmunidadesorganizacionales.frx":1DBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmunidadesorganizacionales.frx":2694
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmunidadesorganizacionales.frx":2C30
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmunidadesorganizacionales.frx":350C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmunidadesorganizacionales.frx":3DE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmunidadesorganizacionales.frx":46C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmunidadesorganizacionales.frx":47D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmunidadesorganizacionales.frx":48E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmunidadesorganizacionales.frx":49F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmunidadesorganizacionales.frx":4B08
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   23
      Top             =   285
      Width           =   5655
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   150
      TabIndex        =   19
      Top             =   2115
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   285
         Left            =   1710
         TabIndex        =   12
         Top             =   165
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3705
         TabIndex        =   20
         Top             =   150
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList"
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
         Caption         =   "Busqueda de planta:"
         Height          =   195
         Left            =   195
         TabIndex        =   21
         Top             =   195
         Width           =   1470
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4515
      Left            =   150
      TabIndex        =   22
      Top             =   2670
      Width           =   5655
      Begin MSComctlLib.ListView lv_unidadesorganizacionales 
         Height          =   4350
         Left            =   45
         TabIndex        =   13
         Top             =   120
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   7673
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Empresa"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Clave"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nombre"
            Object.Width           =   6879
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "email"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "empresa"
            Object.Width           =   2
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
               Picture         =   "frmunidadesorganizacionales.frx":4C1A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmunidadesorganizacionales.frx":54F4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmunidadesorganizacionales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim bitacora As Boolean
Dim numero_items_unidadesorganizacionales As Integer





Private Sub cmd_deshacer_Click()
      txt_empresa.Enabled = False
      cmb_unidadesorganizacionales.Enabled = False
      txt_unidad.Enabled = False
      Call pro_textos

End Sub

Private Sub cmd_eliminar_Click()
   txt_empresa.Enabled = False
   txt_unidad.Enabled = False
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
      Call pro_elimina_unidadesorganizacionales
      rs.Open "select * from tb_unidadesorganizacionales", cnn, adOpenDynamic, adLockOptimistic
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
   If var_modifica_registro_unidadorganizacional = False Then
      rs.Open "select * from TB_UNIDADESORGANIZACIONALES where vcha_uor_unidad_id = '" + Me.txt_unidad + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_posible = False
      End If
      rs.Close
   End If
   If var_posible = True Then
      txt_empresa.Enabled = False
      txt_unidad.Enabled = False
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
         Call pro_guardar_unidadesorganizacionales
         rs.Open "select * from tb_unidadesorganizacionales", cnn, adOpenDynamic, adLockOptimistic
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
      MsgBox "Clave de planta ya existe", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
      If vector_valida_passwords(var_indice_menu) = "*" Then
         frmpasswords.Show
      Else
         Call gPrintListView(lv_unidadesorganizacionales, "LISTADO DE unidadesorganizacionales")
      End If

End Sub

Private Sub cmd_nuevo_Click()
      Call pro_limpiatextos(Me)
      txt_empresa.Enabled = True
      txt_empresa.SetFocus: var_modifica_registro_unidadorganizacional = False
      cmd_guardar.Enabled = True
      cmd_deshacer.Enabled = True
      txt_empresa.Enabled = True
      txt_unidad.Enabled = True
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_unidadorganizacional = False Then
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

Private Sub cmd_series_Click()
   If Trim(txt_unidad) <> "" Then
      frmseries.txt_empresa = txt_empresa
      frmseries.txt_unidad = txt_unidad
      frmseries.Show 1
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
   rs.Open "select * from tb_unidadesorganizacionales", cnn, adOpenDynamic, adLockOptimistic
   If rs.BOF Then
      txt_empresa.Enabled = False
      txt_unidad.Enabled = False
      txt_unidad.Enabled = False
      cmd_guardar.Enabled = False
      cmd_deshacer.Enabled = False
      cmd_eliminar.Enabled = False
   Else
      cmd_guardar.Enabled = True
      cmd_deshacer.Enabled = True
      cmd_eliminar.Enabled = True
      txt_empresa.Enabled = False
      txt_unidad.Enabled = False
   End If
   rs.Close
   var_modifica_registro_unidadorganizacional = True
   lv_unidadesorganizacionales.SmallIcons = ImageList
   Call pro_llena_listview1
   pro_textos
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_despliega_menu = True Then
      var_swpassword = False
      var_modifica_registro_unidadorganizacional = False
   End If
   Call activa_forma(var_activa_forma_unidadesorganizacionales)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         txt_empresa = lv_lista.selectedItem
         txt_nombre_empresa = lv_lista.selectedItem.SubItems(1)
      Else
         txt_empresa = ""
         txt_nombre_empresa = ""
      End If
      txt_empresa.SetFocus
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub lv_unidadesorganizacionales_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_unidadesorganizacionales, ColumnHeader)
End Sub

Private Sub lv_unidadesorganizacionales_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Set lv_unidadesorganizacionales.selectedItem = Item
   pro_textos
   var_modifica_registro_unidadorganizacional = True
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_unidadesorganizacionales.SetFocus
      Call pro_avanzar(Me, lv_unidadesorganizacionales, Button)
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_unidadesorganizacionales.ListItems(1).Selected = True
      Me.lv_unidadesorganizacionales.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 4 Then
      numero_items_unidadesorganizacionales = Me.lv_unidadesorganizacionales.ListItems.Count
      lv_unidadesorganizacionales.ListItems(numero_items_unidadesorganizacionales).Selected = True
      Me.lv_unidadesorganizacionales.selectedItem.EnsureVisible
      pro_textos
   End If
err0:
End Sub


Sub pro_guardar_unidadesorganizacionales()
Dim ok As Boolean
Set TB_UNIDADESORGANIZACIONALES = New TB_UNIDADESORGANIZACIONALES
Set TB_BITACORA_UNIDADESORG_I = New TB_BITACORA_UNIDADESORG_I
   ok = True
   If txt_empresa <> "" And txt_unidad <> "" Then
      If var_hubo_cambios Then
         rs.Open "select * from tb_unidadesorganizacionales where vcha_uor_unidad_id = '" + txt_unidad + "'", cnn, adOpenDynamic, adLockOptimistic
         ok = TB_UNIDADESORGANIZACIONALES.Anadir(txt_empresa, txt_unidad, txt_nombre_unidad, txt_email)
         If ok Then
            bitacora = True
            If var_modifica_registro_unidadorganizacional = False Then
               var_operacion_bitacora = "I"
               bitacora = TB_BITACORA_UNIDADESORG_I.Anadir(txt_unidad, "VCHA_UOR_NOMBRE", "I", "", txt_unidad, var_clave_usuario_global, fun_NombrePc, Date)
            Else
               var_operacion_bitacora = "M"
               If rs(0) <> txt_empresa Then
                  bitacora = TB_BITACORA_UNIDADESORG_I.Anadir(txt_unidad, "VCHA_EMP_EMPRESA_ID", "M", rs(0).Value, txt_empresa, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(1) <> txt_unidad Then
                  bitacora = TB_BITACORA_UNIDADESORG_I.Anadir(txt_unidad, "VCHA_UOR_UNIDAD_ID", "M", rs(1).Value, txt_unidad, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(2) <> txt_nombre_unidad Then
                  bitacora = TB_BITACORA_UNIDADESORG_I.Anadir(txt_unidad, "VCHA_UOR_NOMBRE", "M", rs(2).Value, txt_nombre_unidad, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(3) <> txt_email Then
                  bitacora = TB_BITACORA_UNIDADESORG_I.Anadir(txt_unidad, "VCHA_UOR_MAIL", "M", rs(3).Value, txt_email, var_clave_usuario_global, fun_NombrePc, Date)
               End If
            End If
            rs.Close
            pro_actualiza_ListView
            txt_empresa.Enabled = False
            MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
            txt_registros = lv_unidadesorganizacionales.ListItems.Count
            var_modifica_registro_unidadorganizacional = True
         Else
            MsgBox "No se puede grabar registro: " + TB_UNIDADESORGANIZACIONALES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
         End If
      End If
   End If
   Set TB_UNIDADESORGANIZACIONALES = Nothing: var_hubo_cambios = False
End Sub

Sub pro_elimina_unidadesorganizacionales()
'On Error GoTo SALIR:
   Dim var_llave_usuarios As String
   Set TB_UNIDADESORGANIZACIONALES = New TB_UNIDADESORGANIZACIONALES
   Set TB_BITACORA_UNIDADESORG_I = New TB_BITACORA_UNIDADESORG_I
   ok = True
   bitacora = True
   If txt_empresa <> "" And txt_unidad <> "" And var_modifica_registro_unidadorganizacional = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_UNIDADESORGANIZACIONALES.Eliminar(txt_unidad)
      Else
         GoTo salir:
      End If
      If ok Then
        numero_items_unidadesorganizacionales = numero_items_unidadesorganizacionales - 1
        MsgBox "Se Elimino Correctamente el Registro", vbInformation
        var_operacion_bitacora = "E"
        bitacora = TB_BITACORA_UNIDADESORG_I.Anadir(txt_unidad, "VCHA_UOR_UNIDAD_ID", "E", "", txt_nombre_unidad, var_clave_usuario_global, fun_NombrePc, Date)
        lv_unidadesorganizacionales.ListItems.Remove (lv_unidadesorganizacionales.selectedItem.Index)
        Call pro_limpiatextos(Me)
        txt_registros = lv_unidadesorganizacionales.ListItems.Count
        If lv_unidadesorganizacionales.ListItems.Count > 0 Then
           If lv_unidadesorganizacionales.selectedItem.Selected <> False Then
              lv_unidadesorganizacionales.selectedItem.Selected = True
           End If
        End If
        pro_textos
      Else
        MsgBox "No se puede eliminar registro: " + TB_UNIDADESORGANIZACIONALES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
   Set TB_UNIDADESORGANIZACIONALES = Nothing
End Sub


Sub pro_llena_listview1()
   Dim list_item As ListItem
   lv_unidadesorganizacionales.ListItems.Clear
   rsaux2.Open "select * from TB_unidadesorganizacionales", cnn, adOpenDynamic, adLockOptimistic
   numero_items_unidadesorganizacionales = 0
   While Not rsaux2.EOF
      Set list_item = lv_unidadesorganizacionales.ListItems.Add(, , rsaux2(0).Value)
      list_item.SubItems(1) = IIf(IsNull(rsaux2(1).Value), "", rsaux2(1).Value)
      list_item.SubItems(2) = IIf(IsNull(rsaux2(2).Value), "", rsaux2(2).Value)
      list_item.SubItems(3) = IIf(IsNull(rsaux2(3).Value), "", rsaux2(3).Value)
      rsaux2.MoveNext:
      numero_items_unidadesorganizacionales = numero_items_unidadesorganizacionales + 1
    Wend
    rsaux2.Close
End Sub


Sub pro_textos()
On Error GoTo err0:
Dim var_n As Double
var_n = lv_unidadesorganizacionales.ListItems.Count
   If var_n > 0 Then
      If var_modifica_registro_unidadorganizacional = False Then
         txt_unidad = ""
         txt_nombre_unidad = ""
         txt_email = ""
      Else
         txt_empresa = lv_unidadesorganizacionales.selectedItem
         txt_unidad = lv_unidadesorganizacionales.selectedItem.SubItems(1)
         txt_nombre_unidad = lv_unidadesorganizacionales.selectedItem.SubItems(2)
         txt_email = lv_unidadesorganizacionales.selectedItem.SubItems(3)
      End If
      rs.Open "select * from tb_empresas where vcha_emp_empresa_id = '" + txt_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_empresa = IIf(IsNull(rs!vcha_emp_nombre), "", rs!vcha_emp_nombre)
      Else
         txt_nombre_empresa = ""
      End If
      rs.Close
   End If
   var_numero_renglones = lv_unidadesorganizacionales.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_unidadesorganizacionales.ColumnHeaders(3).Width = 3850
   Else
      lv_unidadesorganizacionales.ColumnHeaders(3).Width = 4099.71
   End If
   var_hubo_cambios = False
   var_modifica_registro_unidadorganizacional = True
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem

    If var_modifica_registro_unidadorganizacional = False Then
        Set list_item = lv_unidadesorganizacionales.ListItems.Add(, , txt_empresa)
        list_item.SubItems(1) = txt_unidad
        list_item.SubItems(2) = txt_nombre_unidad
        list_item.SubItems(3) = txt_email
        list_item.EnsureVisible
        list_item.Selected = True
        numero_items_unidadesorganizacionales = numero_items_unidadesorganizacionales + 1
    Else
        lv_unidadesorganizacionales.ListItems.Item(lv_unidadesorganizacionales.selectedItem.Index).Checked = False
        lv_unidadesorganizacionales.ListItems.Item(lv_unidadesorganizacionales.selectedItem.Index) = txt_empresa
        lv_unidadesorganizacionales.ListItems.Item(lv_unidadesorganizacionales.selectedItem.Index).ListSubItems(1) = txt_unidad
        lv_unidadesorganizacionales.ListItems.Item(lv_unidadesorganizacionales.selectedItem.Index).ListSubItems(2) = txt_nombre_unidad
        lv_unidadesorganizacionales.ListItems.Item(lv_unidadesorganizacionales.selectedItem.Index).ListSubItems(3) = txt_email
        lv_unidadesorganizacionales.ListItems.Item(lv_unidadesorganizacionales.selectedItem.Index).Selected = True
    End If
'    lv_unidadesorganizacionales.SetFocus
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(lv_unidadesorganizacionales, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub



Private Sub txt_email_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_email_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Me.cmd_guardar.Enabled = True Then
         Me.cmd_guardar.SetFocus
      End If
   End If
End Sub

Private Sub txt_empresa_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_empresa_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_empresas order by vcha_emp_empresa_id", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_emp_empresa_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_emp_nombre), "", rs!vcha_emp_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Empresas"
      var_tipo_lista = 1
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_empresa_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_empresa_LostFocus()
   If Trim(txt_empresa) <> "" Then
      rs.Open "select * from tb_empresas where vcha_emp_empresa_id = '" + txt_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_empresa = IIf(IsNull(rs!vcha_emp_nombre), "", rs!vcha_emp_nombre)
      Else
         MsgBox "Clave de empresa no existe", vbOKOnly, "ATENCION"
         txt_empresa = ""
         txt_nombre_empresa = ""
      End If
      rs.Close
   End If
End Sub

Private Sub txt_nombre_empresa_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_empresa_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_empresas order by vcha_emp_empresa_id", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_emp_empresa_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_emp_nombre), "", rs!vcha_emp_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Empresas"
      var_tipo_lista = 1
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_empresa_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_unidad_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_unidad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_unidad_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_unidad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub
