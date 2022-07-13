VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsubtiposusos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subtipo de Usos de Productos"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   Icon            =   "frmsubtiposusos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5955
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   150
      TabIndex        =   22
      Top             =   390
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   23
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
         TabIndex        =   24
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmsubtiposusos.frx":08CA
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
      Picture         =   "frmsubtiposusos.frx":09CC
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
      Picture         =   "frmsubtiposusos.frx":0ACE
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
      Picture         =   "frmsubtiposusos.frx":0BA0
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
      Picture         =   "frmsubtiposusos.frx":0CA2
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
      Picture         =   "frmsubtiposusos.frx":0DA4
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Caption         =   " Subtipo de uso"
      Height          =   1380
      Left            =   150
      TabIndex        =   0
      Top             =   450
      Width           =   5655
      Begin VB.TextBox txt_nombre_uso 
         Height          =   315
         Left            =   2505
         MaxLength       =   50
         TabIndex        =   8
         Top             =   255
         Width           =   3060
      End
      Begin VB.TextBox txt_nombre_subtipo_uso 
         Height          =   315
         Left            =   1365
         MaxLength       =   50
         TabIndex        =   10
         Top             =   945
         Width           =   4200
      End
      Begin VB.TextBox txt_subtipo_uso 
         Height          =   315
         Left            =   1365
         MaxLength       =   50
         TabIndex        =   9
         Top             =   600
         Width           =   1140
      End
      Begin VB.TextBox txt_uso 
         Height          =   315
         Left            =   1365
         MaxLength       =   50
         TabIndex        =   7
         Top             =   255
         Width           =   1125
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion:"
         Height          =   195
         Index           =   4
         Left            =   105
         TabIndex        =   16
         Top             =   1005
         Width           =   885
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Subtipo de uso:"
         Height          =   195
         Index           =   2
         Left            =   105
         TabIndex        =   12
         Top             =   660
         Width           =   1110
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Uso de producto:"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   11
         Top             =   315
         Width           =   1230
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   150
      TabIndex        =   13
      Top             =   1800
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   2520
         TabIndex        =   21
         Top             =   150
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   4095
         TabIndex        =   19
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
         Caption         =   "Busqueda de subtipo de uso:"
         Height          =   195
         Left            =   195
         TabIndex        =   14
         Top             =   195
         Width           =   2070
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4875
      Left            =   150
      TabIndex        =   15
      Top             =   2325
      Width           =   5655
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   17
         Top             =   285
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSComctlLib.ImageList icono_encabezado 
         Left            =   0
         Top             =   15
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
               Picture         =   "frmsubtiposusos.frx":13DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmsubtiposusos.frx":1CB8
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lv_subtiposusos 
         Height          =   4695
         Left            =   45
         TabIndex        =   18
         Top             =   135
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   8281
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
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
            Text            =   "tipoarticulo"
            Object.Width           =   0
         EndProperty
      End
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
            Picture         =   "frmsubtiposusos.frx":2592
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubtiposusos.frx":2E6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubtiposusos.frx":3746
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubtiposusos.frx":3CE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubtiposusos.frx":45BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubtiposusos.frx":4E98
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubtiposusos.frx":5772
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubtiposusos.frx":5884
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubtiposusos.frx":5996
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubtiposusos.frx":5AA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubtiposusos.frx":5BBA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   20
      Top             =   285
      Width           =   5655
   End
End
Attribute VB_Name = "frmsubtiposusos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim numero_items_subtiposusos As Integer
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
      Call pro_elimina_subtiposusos
      rs.Open "select * from tb_subtiposusos", cnn, adOpenDynamic, adLockOptimistic
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
   If Trim(txt_uso) <> "" Then
      var_posible = True
      If var_modifica_registro_uso = False Then
         rs.Open "select * from TB_subtiposusos where VCHA_SUS_SUBTIPO_USO_ID = '" + Me.txt_subtipo_uso + "'", cnn, adOpenDynamic, adLockOptimistic
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
            Call pro_guardar_subtiposusos
            rs.Open "select * from tb_subtiposusos", cnn, adOpenDynamic, adLockOptimistic
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
         MsgBox "Clave de subtipo de uso ya existe", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Se debe de indicar el tipo de uso", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call gPrintListView(lv_subtiposusos, "LISTADO DE subtiposusos")
        End If

End Sub

Private Sub cmd_nuevo_Click()
        Call pro_limpiatextos(Me)
        txt_uso.Enabled = True
        txt_uso.SetFocus: var_modifica_registro_uso = False
        txt_nombre_uso.Enabled = True
        txt_subtipo_uso.Enabled = True
        cmd_guardar.Enabled = True
        cmd_deshacer.Enabled = True
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_uso = False Then
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
   var_modifica_registro_uso = True
   lv_subtiposusos.SmallIcons = ImageList
   Call pro_encabezadosView(Me, lv_subtiposusos, False)
   Call pro_llena_listview1
   pro_textos
   rs.Open "select * from tb_subtiposusos", cnn, adOpenDynamic, adLockOptimistic
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
      var_modifica_registro_uso = False
   End If
   Call activa_forma(var_activa_forma_subtiposusos)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         txt_uso = lv_lista.selectedItem
         txt_nombre_uso = lv_lista.selectedItem.SubItems(1)
      Else
         txt_uso = ""
         txt_nombre_uso = ""
      End If
      txt_uso.SetFocus
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub lv_subtiposusos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_subtiposusos, ColumnHeader)
End Sub

Private Sub lv_subtiposusos_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_subtiposusos.selectedItem = Item
        pro_textos
        var_modifica_registro_uso = True
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_subtiposusos.SetFocus
      Call pro_avanzar(Me, lv_subtiposusos, Button)
      Me.lv_subtiposusos.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_subtiposusos.ListItems(1).Selected = True
      Me.lv_subtiposusos.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 4 Then
      numero_items_subtiposusos = Me.lv_subtiposusos.ListItems.Count
      lv_subtiposusos.ListItems(numero_items_subtiposusos).Selected = True
      Me.lv_subtiposusos.selectedItem.EnsureVisible
      pro_textos
   End If
err0:
End Sub


Sub pro_guardar_subtiposusos()
   Dim ok As Boolean
   Set TB_SUBTIPOSUSOS = New TB_SUBTIPOSUSOS
   Set TB_BITACORA_SUBTIPOSUSOS = New TB_BITACORA_SUBTIPOSUSOS
   If txt_uso <> "" And txt_subtipo_uso <> "" Then
      If var_hubo_cambios Then
         rs.Open "select * from tb_subtiposusos where vcha_sus_subtipo_uso_id = '" + txt_subtipo_uso + "'", cnn, adOpenDynamic, adLockOptimistic
         ok = TB_SUBTIPOSUSOS.Anadir(txt_uso, txt_subtipo_uso, txt_nombre_subtipo_uso)
         If ok Then
            bitacora = True
            If var_modifica_registro_uso = False Then
               var_operacion_bitacora = "I"
               bitacora = TB_BITACORA_SUBTIPOSUSOS.Anadir(txt_subtipo_uso, "VCHA_SUS_NOMBRE", var_operacion_bitacora, "", txt_nombre_subtipo_uso, var_clave_usuario_global, fun_NombrePc, Date)
            Else
               var_operacion_bitacora = "M"
               If rs(0) <> txt_uso Then
                  bitacora = TB_BITACORA_SUBTIPOSUSOS.Anadir(txt_subtipo_uso, "VCHA_LIN_LINEA_ID", var_operacion_bitacora, rs(0), txt_uso, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(1) <> txt_subtipo_uso Then
                  bitacora = TB_BITACORA_SUBTIPOSUSOS.Anadir(txt_subtipo_uso, "VCHA_SLI_SUBLINEA_ID", var_operacion_bitacora, rs(1), txt_subtipo_uso, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(2) <> txt_nombre_subtipo_uso Then
                  bitacora = TB_BITACORA_SUBTIPOSUSOS.Anadir(txt_subtipo_uso, "VCHA_SLI_NOMBRE", var_operacion_bitacora, rs(2), txt_nombre_subtipo_uso, var_clave_usuario_global, fun_NombrePc, Date)
               End If
            End If
            rs.Close
            pro_actualiza_ListView
            txt_uso.Enabled = False
            MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
            txt_registros = lv_subtiposusos.ListItems.Count
            var_modifica_registro_uso = True
            Call pro_textos
         Else
            MsgBox "No se puede grabar registro: " + TB_SUBTIPOSUSOS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
         End If
      End If
   End If
   Set TB_SUBTIPOSUSOS = Nothing: var_hubo_cambios = False
End Sub

Sub pro_elimina_subtiposusos()
   Dim var_llave_usuarios As String
   Set TB_SUBTIPOSUSOS = New TB_SUBTIPOSUSOS
   Set TB_BITACORA_SUBTIPOSUSOS = New TB_BITACORA_SUBTIPOSUSOS
   On Error GoTo salir
   ok = True
   If txt_uso <> "" And txt_subtipo_uso <> "" And var_modifica_registro_uso = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_SUBTIPOSUSOS.Eliminar(txt_subtipo_uso)
      Else
         GoTo salir:
      End If
      If ok Then
         var_operacion_bitacora = "E"
         bitacora = TB_BITACORA_SUBTIPOSUSOS.Anadir(txt_subtipo_uso, "VCHA_SUS_NOMBRE", var_operacion_bitacora, txt_nombre_subtipo_uso, "", var_clave_usuario_global, fun_NombrePc, Date)
         numero_items_subtiposusos = numero_items_subtiposusos - 1
         MsgBox "Se Elimino Correctamente el Registro", vbInformation
         lv_subtiposusos.ListItems.Remove (lv_subtiposusos.selectedItem.Index)
         Call pro_limpiatextos(Me)
         txt_registros = lv_subtiposusos.ListItems.Count
         lv_subtiposusos.selectedItem.Selected = True
         pro_textos
      Else
         MsgBox "No se puede grabar registro: " + TB_SUBTIPOSUSOS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
   Set TB_SUBTIPOSUSOS = Nothing
End Sub


Sub pro_llena_listview1()

Dim list_item As ListItem

    rs.Open "select * from TB_subtiposusos", cnn, adOpenDynamic, adLockOptimistic
    numero_items_subtiposusos = 0
     While Not rs.EOF
        Set list_item = lv_subtiposusos.ListItems.Add(, , rs(1).Value)
        list_item.SubItems(1) = IIf(IsNull(rs(2).Value), "", rs(2).Value)
        list_item.SubItems(2) = IIf(IsNull(rs(0).Value), "", rs(0).Value)
    rs.MoveNext:
    numero_items_subtiposusos = numero_items_subtiposusos + 1
    Wend
    rs.Close

End Sub


Sub pro_textos()
'On Error GoTo err0:
Dim var_n As Integer
   var_n = lv_subtiposusos.ListItems.Count
   If var_n > 0 Then
      txt_subtipo_uso = lv_subtiposusos.selectedItem
      txt_nombre_subtipo_uso = lv_subtiposusos.selectedItem.SubItems(1)
      txt_uso = lv_subtiposusos.selectedItem.SubItems(2)
      rs.Open "select * from tb_usos where vcha_uso_uso_id = '" + txt_uso + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_uso = IIf(IsNull(rs!VCHA_USO_NOMBRE), "", rs!VCHA_USO_NOMBRE)
      Else
         txt_nombre_uso = ""
      End If
      rs.Close
      txt_uso.Enabled = False
      txt_nombre_uso.Enabled = False
      txt_subtipo_uso.Enabled = False
   End If
   var_numero_renglones = lv_subtiposusos.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_subtiposusos.ColumnHeaders(2).Width = 3850
   Else
      lv_subtiposusos.ColumnHeaders(2).Width = 4099.71
   End If
   var_hubo_cambios = False
   var_modifica_registro_uso = True
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro_uso = False Then
        Set list_item = lv_subtiposusos.ListItems.Add(, , txt_subtipo_uso)
        list_item.SubItems(1) = txt_nombre_subtipo_uso
        list_item.SubItems(2) = txt_uso
        list_item.EnsureVisible
        list_item.Selected = True
        numero_items_subtiposusos = numero_items_subtiposusos + 1
    Else
        lv_subtiposusos.ListItems.Item(lv_subtiposusos.selectedItem.Index).Checked = False
        lv_subtiposusos.ListItems.Item(lv_subtiposusos.selectedItem.Index) = txt_subtipo_uso
        lv_subtiposusos.ListItems.Item(lv_subtiposusos.selectedItem.Index).ListSubItems(1) = txt_nombre_subtipo_uso
        lv_subtiposusos.ListItems.Item(lv_subtiposusos.selectedItem.Index).ListSubItems(2) = txt_uso
        lv_subtiposusos.ListItems.Item(lv_subtiposusos.selectedItem.Index).Selected = True
    End If
    lv_subtiposusos.SetFocus
End Sub



Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_subtiposusos, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub

Private Sub txt_nombre_subtipo_uso_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_subtipo_uso_KeyPress(KeyAscii As Integer)
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

Private Sub txt_nombre_uso_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_uso_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para accesar al catálogo"
End Sub

Private Sub txt_nombre_uso_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_usos order by vcha_uso_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_USO_USO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_USO_NOMBRE), "", rs!VCHA_USO_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "USOS"
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
      var_activa_forma_usos = Me.Name
      Me.Enabled = False
      frmusos.Show
   End If
End Sub

Private Sub txt_nombre_uso_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_uso_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_subtipo_uso_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_subtipo_uso_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_uso_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_uso_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para accesar al catálogo"
End Sub

Private Sub txt_uso_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_usos order by vcha_uso_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_USO_USO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_USO_NOMBRE), "", rs!VCHA_USO_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "USOS"
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
      var_activa_forma_usos = Me.Name
      Me.Enabled = False
      frmusos.Show
   End If
End Sub

Private Sub txt_uso_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_uso_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_uso) <> "" Then
      rs.Open "SELECT * FROM TB_USOS WHERE VCHA_USO_USO_ID = '" + txt_uso + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_uso = IIf(IsNull(rs!VCHA_USO_NOMBRE), "", rs!VCHA_USO_NOMBRE)
      Else
         MsgBox "Clave de uso incorrecto", vbOKOnly, "ATENCION"
         txt_uso = ""
         txt_nombre_uso = ""
      End If
      rs.Close
   Else
      txt_nombre_uso = ""
   End If
End Sub
