VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmpersonal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personal"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   5895
   Begin VB.Frame Frame3 
      Height          =   5250
      Left            =   120
      TabIndex        =   16
      Top             =   1980
      Width           =   5655
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
               Picture         =   "frmpersonal.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmpersonal.frx":08DA
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lv_personas 
         Height          =   5055
         Left            =   45
         TabIndex        =   17
         Top             =   135
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   8916
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
            Text            =   "Clave"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   6879
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Tonos "
      Height          =   1020
      Left            =   120
      TabIndex        =   13
      Top             =   420
      Width           =   5655
      Begin VB.TextBox txt_clave_persona 
         Height          =   315
         Left            =   1065
         MaxLength       =   50
         TabIndex        =   6
         Top             =   225
         Width           =   1035
      End
      Begin VB.TextBox txt_nombre_persona 
         Height          =   315
         Left            =   1065
         MaxLength       =   50
         TabIndex        =   7
         Top             =   585
         Width           =   4500
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   15
         Top             =   270
         Width           =   450
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   14
         Top             =   630
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   1860
         TabIndex        =   10
         Top             =   150
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3675
         TabIndex        =   11
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
               Object.ToolTipText     =   "Nuevo Registro"
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
         Caption         =   "Busqueda de persona:"
         Height          =   195
         Left            =   195
         TabIndex        =   12
         Top             =   195
         Width           =   1605
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2745
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   8
      Top             =   105
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5430
      Picture         =   "frmpersonal.frx":11B4
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      Picture         =   "frmpersonal.frx":17EE
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1110
      Picture         =   "frmpersonal.frx":18F0
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   780
      Picture         =   "frmpersonal.frx":19F2
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   450
      Picture         =   "frmpersonal.frx":1AC4
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Guardar Alt + G"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmpersonal.frx":1BC6
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   45
      Width           =   330
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   -30
      Top             =   15
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpersonal.frx":1CC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpersonal.frx":25A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpersonal.frx":2E7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpersonal.frx":3418
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpersonal.frx":3CF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpersonal.frx":45CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpersonal.frx":4EA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpersonal.frx":4FBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpersonal.frx":50CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpersonal.frx":51DE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   120
      TabIndex        =   18
      Top             =   300
      Width           =   5655
   End
End
Attribute VB_Name = "frmpersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim numero_items_personas As Integer

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
      Call pro_elimina_personas
      rs.Open "select * from tb_personal", cnn, adOpenDynamic, adLockOptimistic
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
   If var_modifica_registro_personas = False Then
      rs.Open "select * from tb_personal where VCHA_PER_PERSONAL_ID = '" + Me.txt_clave_persona + "'", cnn, adOpenDynamic, adLockOptimistic
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
         Call pro_guardar_personas
         rs.Open "select * from tb_personal", cnn, adOpenDynamic, adLockOptimistic
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
      MsgBox "Clave de personas ya existe", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call gPrintListView(lv_personas, "LISTADO DE personas")
        End If

End Sub

Private Sub cmd_nuevo_Click()
   Call pro_limpiatextos(Me)
   txt_clave_persona.Enabled = True
   txt_clave_persona.SetFocus: var_modifica_registro_personas = False
   cmd_guardar.Enabled = True
   cmd_deshacer.Enabled = True
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_personas = False Then
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
   var_modifica_registro_personas = True
   lv_personas.SmallIcons = ImageList1
   Call pro_encabezadosView(Me, lv_personas, False)
   Call pro_llena_listview1
   pro_textos
   rs.Open "select * from tb_personal", cnn, adOpenDynamic, adLockOptimistic
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
   Me.txt_clave_persona.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_despliega_menu = True Then
      var_swpassword = False
      var_modifica_registro_personas = False
   End If
   Call activa_forma(var_activa_forma_personas)
End Sub

Private Sub lv_personas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_personas, ColumnHeader)
End Sub

Private Sub lv_personas_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_personas.selectedItem = Item
        pro_textos
        var_modifica_registro_personas = True
        txt_clave_persona.Enabled = False

End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_personas.SetFocus
      Call pro_avanzar(Me, lv_personas, Button)
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_personas.ListItems(1).Selected = True
      lv_personas.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 4 Then
      numero_items_personas = lv_personas.ListItems.Count
      lv_personas.ListItems(numero_items_personas).Selected = True
      lv_personas.selectedItem.EnsureVisible
      pro_textos
   End If
err0:
End Sub


Sub pro_guardar_personas()
   Dim ok As Boolean
   ok = True
   If txt_clave_persona <> "" And txt_nombre_persona <> "" Then
      If var_hubo_cambios Then
         If var_modifica_registro_personas = False Then
            rs.Open "insert into tb_personal (VCHA_PER_PERSONAL_ID, VCHA_PER_NOMBRE) values ('" + txt_clave_persona + "', '" + txt_nombre_persona + "')", cnn, adOpenDynamic, adLockOptimistic
         Else
            rs.Open "update tb_personal set VCHA_PER_NOMBRE = '" + txt_nombre_persona + "' where vcha_per_personal_id = '" + txt_clave_persona + "'", cnn, adOpenDynamic, adLockOptimistic
         End If
         pro_actualiza_ListView
         txt_clave_persona.Enabled = False
         MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
         txt_registros = lv_personas.ListItems.Count
         var_modifica_registro_personas = True
      End If
   End If
   Set tb_personal = Nothing: var_hubo_cambios = False
End Sub

Sub pro_elimina_personas()
   If txt_clave_persona <> "" And txt_nombre_persona <> "" And var_modifica_registro_personas = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         rs.Open "delete from tb_personal where vcha_per_personal_id = '" + txt_clave_persona + "'", cnn, adOpenDynamic, adLockOptimistic
         var_operacion_bitacora = "E"
         numero_items_personas = numero_items_personas - 1
         MsgBox "Se Elimino Correctamente el Registro", vbInformation
         lv_personas.ListItems.Remove (lv_personas.selectedItem.Index)
          Call pro_limpiatextos(Me)
         txt_registros = lv_personas.ListItems.Count
         lv_personas.selectedItem.Selected = True
         pro_textos
      Else
         GoTo salir:
      End If
   End If
salir:
   Set tb_personal = Nothing
End Sub


Sub pro_llena_listview1()
   Dim list_item As ListItem
   rs.Open "select * from tb_personal", cnn, adOpenDynamic, adLockOptimistic
   numero_items_personas = 0
   While Not rs.EOF
      Set list_item = lv_personas.ListItems.Add(, , rs!vcha_per_personal_id)
      list_item.SubItems(1) = IIf(IsNull(rs!vcha_per_nombre), "", rs!vcha_per_nombre)
      rs.MoveNext:
      numero_items_personas = numero_items_personas + 1
    Wend
    rs.Close
End Sub


Sub pro_textos()
On Error GoTo err0:
Dim var_n As Double
   var_n = lv_personas.ListItems.Count
   If var_n > 0 Then
      txt_clave_persona = lv_personas.selectedItem
      txt_nombre_persona = lv_personas.selectedItem.SubItems(1)
   End If
   var_numero_renglones = lv_personas.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_personas.ColumnHeaders(2).Width = 3850
   Else
      lv_personas.ColumnHeaders(2).Width = 4099.71
   End If
   var_modifica_registro_personas = True
   var_hubo_cambios = False
   Me.txt_clave_persona.Enabled = False
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro_personas = False Then
        Set list_item = lv_personas.ListItems.Add(, , txt_clave_persona)
        list_item.SubItems(1) = txt_nombre_persona
        list_item.EnsureVisible
        list_item.Selected = True
        numero_items_personas = numero_items_personas + 1
    Else
        lv_personas.ListItems.Item(lv_personas.selectedItem.Index).Checked = False
        lv_personas.ListItems.Item(lv_personas.selectedItem.Index) = txt_clave_persona
        lv_personas.ListItems.Item(lv_personas.selectedItem.Index).ListSubItems(1) = txt_nombre_persona
        lv_personas.ListItems.Item(lv_personas.selectedItem.Index).Selected = True
    End If
'    lv_personas.SetFocus
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_personas, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub


Private Sub txt_clave_persona_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_clave_persona_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_persona_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_persona_KeyPress(KeyAscii As Integer)
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
