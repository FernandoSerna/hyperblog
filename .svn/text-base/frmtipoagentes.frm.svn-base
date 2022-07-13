VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmtipoagentes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo agentes"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   Icon            =   "frmtipoagentes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5955
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmtipoagentes.frx":08CA
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
      Picture         =   "frmtipoagentes.frx":09CC
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
      Picture         =   "frmtipoagentes.frx":0ACE
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
      Picture         =   "frmtipoagentes.frx":0BA0
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
      Picture         =   "frmtipoagentes.frx":0CA2
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
      Picture         =   "frmtipoagentes.frx":0DA4
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
      Left            =   1200
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   17
      Top             =   285
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   " Tipo de Agente "
      Height          =   1005
      Left            =   150
      TabIndex        =   0
      Top             =   465
      Width           =   5655
      Begin VB.TextBox txt_tipoagentes 
         Height          =   315
         Index           =   1
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   8
         Top             =   585
         Width           =   4215
      End
      Begin VB.TextBox txt_tipoagentes 
         Height          =   315
         Index           =   0
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   7
         Top             =   255
         Width           =   900
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   10
         Top             =   615
         Width           =   885
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   9
         Top             =   270
         Width           =   450
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   150
      TabIndex        =   12
      Top             =   1440
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   2475
         TabIndex        =   19
         Top             =   150
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   4050
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
         Caption         =   "Busqueda de Tipo de Agentes:"
         Height          =   195
         Left            =   165
         TabIndex        =   13
         Top             =   210
         Width           =   2205
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5265
      Left            =   150
      TabIndex        =   14
      Top             =   1950
      Width           =   5655
      Begin MSComctlLib.ListView lv_tipoagentes 
         Height          =   5040
         Left            =   45
         TabIndex        =   16
         Top             =   165
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   8890
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
            Text            =   "Descripción"
            Object.Width           =   6879
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   3255
      Top             =   -255
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
            Picture         =   "frmtipoagentes.frx":13DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtipoagentes.frx":1CB8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   15
      Top             =   285
      Width           =   5655
   End
   Begin MSComctlLib.ImageList ImageList1 
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
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtipoagentes.frx":2592
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtipoagentes.frx":2E6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtipoagentes.frx":3746
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtipoagentes.frx":3CE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtipoagentes.frx":45BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtipoagentes.frx":4E98
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtipoagentes.frx":5772
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtipoagentes.frx":5884
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtipoagentes.frx":5996
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtipoagentes.frx":5AA8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lab_paises 
      AutoSize        =   -1  'True
      Caption         =   "Busqueda de pais:"
      Height          =   195
      Index           =   3
      Left            =   150
      TabIndex        =   11
      Top             =   1620
      Width           =   1320
   End
End
Attribute VB_Name = "frmtipoagentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim numero_items_tipoagentes As Integer
Dim bitacora As Boolean




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
      Call pro_elimina_tipoagentes
      rs.Open "select * from tb_tipoagentes", cnn, adOpenDynamic, adLockOptimistic
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
   If var_modifica_registro_tipoagente = False Then
      rs.Open "select * from tb_tipoagentes where VCHA_TAG_TIPOAGENTE_ID = '" + Me.txt_tipoagentes(0) + "'", cnn, adOpenDynamic, adLockOptimistic
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
         Call pro_guardar_tipoagentes
         rs.Open "select * from tb_tipoagentes", cnn, adOpenDynamic, adLockOptimistic
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
     MsgBox "Clave de tipo de agente ya existe", vbOKOnly, "ATENCION"
  End If
End Sub

Private Sub cmd_imprimir_Click()
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call gPrintListView(lv_tipoagentes, "LISTADO DE tipoagentes")
        End If

End Sub

Private Sub cmd_nuevo_Click()
        Call pro_limpiatextos(Me)
        txt_tipoagentes(0).Enabled = True
        txt_tipoagentes(0).SetFocus: var_modifica_registro_tipoagente = False
        cmd_guardar.Enabled = True
        cmd_deshacer.Enabled = True
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_tipoagente = False Then
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
   var_modifica_registro_tipoagente = True
   lv_tipoagentes.SmallIcons = ImageList1
   Call pro_encabezadosView(Me, lv_tipoagentes, False)
   Call pro_llena_listview1
   pro_textos
   rs.Open "select * from tb_tipoagentes", cnn, adOpenDynamic, adLockOptimistic
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
   var_modifica_registro_tipoagente = False
   Call activa_forma(var_activa_forma_tipoagentes)
End Sub

Private Sub lv_tipoagentes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_tipoagentes, ColumnHeader)
End Sub

Private Sub lv_tipoagentes_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_tipoagentes.selectedItem = Item
        pro_textos
        var_modifica_registro_tipoagente = True
        txt_tipoagentes(0).Enabled = False

End Sub


Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_tipoagentes.SetFocus
      Call pro_avanzar(Me, lv_tipoagentes, Button)
      lv_tipoagentes.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_tipoagentes.ListItems(1).Selected = True
      lv_tipoarticulos.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 4 Then
      numero_items_tipoagentes = lv_tipoagentes.ListItems.Count
      lv_tipoagentes.ListItems(numero_items_tipoagentes).Selected = True
      lv_tipoagentes.selectedItem.EnsureVisible
      pro_textos
   End If
err0:
End Sub


Sub pro_guardar_tipoagentes()
   Dim ok As Boolean
   Set TB_TIPOAGENTES = New TB_TIPOAGENTES
   Set TB_BITACORA_TIPOAGENTES = New TB_BITACORA_TIPOAGENTES
   ok = True
   If txt_tipoagentes(0) <> "" And txt_tipoagentes(1) <> "" Then
      If var_hubo_cambios Then
         rs.Open "SELECT * FROM TB_TIPOAGENTES WHERE VCHA_TAG_TIPOAGENTE_ID = '" + txt_tipoagentes(0) + "'", cnn, adOpenDynamic, adLockOptimistic
         ok = TB_TIPOAGENTES.Anadir(txt_tipoagentes(0), txt_tipoagentes(1))
         If ok Then
            bitacora = True
            If var_modifica_registro_tipoagente = False Then
               var_operacion_bitacora = "I"
               bitacora = TB_BITACORA_TIPOAGENTES.Anadir(txt_tipoagentes(0), "VCHA_TAG_NOMBRE", var_operacion_bitacora, "", txt_tipoagentes(1), var_clave_usuario_global, fun_NombrePc, Date)
            Else
               var_operacion_bitacora = "M"
               If rs(0) <> txt_tipoagentes(0) Then
                  bitacora = TB_BITACORA_TIPOAGENTES.Anadir(txt_tipoagentes(0), "VCHA_TAG_TIPOAGENTE_ID", var_operacion_bitacora, rs(0), txt_tipoagentes(0), var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(1) <> txt_tipoagentes(1) Then
                  bitacora = TB_BITACORA_TIPOAGENTES.Anadir(txt_tipoagentes(0), "VCHA_TAG_NOMBRE", var_operacion_bitacora, rs(1), txt_tipoagentes(1), var_clave_usuario_global, fun_NombrePc, Date)
               End If
            End If
            rs.Close
            pro_actualiza_ListView
            txt_tipoagentes(0).Enabled = False
            MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
            txt_registros = lv_tipoagentes.ListItems.Count
            var_modifica_registro_tipoagente = True
         Else
            MsgBox "No se puede grabar registro: " + TB_TIPOAGENTES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
         End If
      End If
   End If
   Set TB_TIPOAGENTES = Nothing: var_hubo_cambios = False
End Sub

Sub pro_elimina_tipoagentes()
   Dim var_llave_usuarios As String
   Set TB_TIPOAGENTES = New TB_TIPOAGENTES
   Set TB_BITACORA_TIPOAGENTES = New TB_BITACORA_TIPOAGENTES
   On Error GoTo salir:
   ok = True
   If txt_tipoagentes(0) <> "" And txt_tipoagentes(1) <> "" And var_modifica_registro_tipoagente = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_TIPOAGENTES.Eliminar(txt_tipoagentes(0))
      Else
         GoTo salir:
      End If
      If ok Then
         var_operacion_bitacora = "E"
         bitacora = TB_BITACORA_TIPOAGENTES.Anadir(txt_tipoagentes(0), "VCHA_TAG_NOMBRE", var_operacion_bitacora, txt_tipoagentes(1), "", var_clave_usuario_global, fun_NombrePc, Date)
         numero_items_tipoagentes = numero_items_tipoagentes - 1
         MsgBox "Se Elimino Correctamente el Registro", vbInformation
         lv_tipoagentes.ListItems.Remove (lv_tipoagentes.selectedItem.Index)
         Call pro_limpiatextos(Me)
         txt_registros = lv_tipoagentes.ListItems.Count
         lv_tipoagentes.selectedItem.Selected = True
         pro_textos
      Else
         MsgBox "No se puede grabar registro: " + TB_TIPOAGENTES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
   Set TB_TIPOAGENTES = Nothing
End Sub


Sub pro_llena_listview1()
Dim list_item As ListItem
   numero_items_tipoagentes = 0
   rs.Open "select * from TB_tipoagentes", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set list_item = lv_tipoagentes.ListItems.Add(, , rs(0).Value)
         list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
         rs.MoveNext:
         numero_items_tipoagentes = numero_items_tipoagentes + 1
   Wend
   rs.Close
   txt_tipoagentes(0).Enabled = False
End Sub


Sub pro_textos()
On Error GoTo err0:
Dim var_n As Double
   var_n = lv_tipoagentes.ListItems.Count
   If var_n > 0 Then
      txt_tipoagentes(0) = lv_tipoagentes.selectedItem
      txt_tipoagentes(1) = lv_tipoagentes.selectedItem.SubItems(1)
   End If
   var_numero_renglones = lv_tipoagentes.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_tipoagentes.ColumnHeaders(2).Width = 3850
   Else
      lv_tipoagentes.ColumnHeaders(2).Width = 4099.71
   End If
   var_hubo_cambios = False
   var_modifica_registro_tipoagente = True
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro_tipoagente = False Then
        Set list_item = lv_tipoagentes.ListItems.Add(, , txt_tipoagentes(0))
        list_item.SubItems(1) = txt_tipoagentes(1)
        list_item.EnsureVisible
        list_item.Selected = True
        numero_items_tipoagentes = numero_items_tipoagentes + 1
    Else
        lv_tipoagentes.ListItems.Item(lv_tipoagentes.selectedItem.Index).Checked = False
        lv_tipoagentes.ListItems.Item(lv_tipoagentes.selectedItem.Index) = txt_tipoagentes(0)
        lv_tipoagentes.ListItems.Item(lv_tipoagentes.selectedItem.Index).ListSubItems(1) = txt_tipoagentes(1)
        lv_tipoagentes.ListItems.Item(lv_tipoagentes.selectedItem.Index).Selected = True
    End If
    lv_tipoagentes.SetFocus
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_tipoagentes, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub

Private Sub txt_tipoagentes_Change(Index As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub

Private Sub txt_tipoagentes_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Index = 0 Then
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
