VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmagrupador_catalogos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Agrupadores de Catálogos"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   5160
      Left            =   150
      TabIndex        =   18
      Top             =   2055
      Width           =   5655
      Begin MSComctlLib.ListView lv_catalogos 
         Height          =   4950
         Left            =   45
         TabIndex        =   9
         Top             =   150
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   8731
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
         NumItems        =   5
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
               Picture         =   "frmagrupador_catalogos.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmagrupador_catalogos.frx":08DA
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmagrupador_catalogos.frx":11B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmagrupador_catalogos.frx":1750
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmagrupador_catalogos.frx":202C
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmagrupador_catalogos.frx":2906
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmagrupador_catalogos.frx":31E0
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmagrupador_catalogos.frx":32F2
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmagrupador_catalogos.frx":3404
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmagrupador_catalogos.frx":3516
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmagrupador_catalogos.frx":3628
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   150
      TabIndex        =   15
      Top             =   1470
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   1935
         TabIndex        =   8
         Top             =   150
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3795
         TabIndex        =   16
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
         TabIndex        =   17
         Top             =   195
         Width           =   1650
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Agrupadores de Catálogos "
      Height          =   1035
      Left            =   150
      TabIndex        =   11
      Top             =   405
      Width           =   5655
      Begin VB.TextBox txt_agrupador 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   6
         Top             =   270
         Width           =   1170
      End
      Begin VB.TextBox txt_nombre_agrupador 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   7
         Top             =   615
         Width           =   4290
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   1
         Left            =   375
         TabIndex        =   13
         Top             =   315
         Width           =   450
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Index           =   2
         Left            =   375
         TabIndex        =   12
         Top             =   660
         Width           =   885
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2910
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   10
      Top             =   30
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5460
      Picture         =   "frmagrupador_catalogos.frx":373A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1470
      Picture         =   "frmagrupador_catalogos.frx":3D74
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      Picture         =   "frmagrupador_catalogos.frx":3E76
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   810
      Picture         =   "frmagrupador_catalogos.frx":3F78
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      Picture         =   "frmagrupador_catalogos.frx":404A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Guardar Alt + G"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmagrupador_catalogos.frx":414C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
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
            Picture         =   "frmagrupador_catalogos.frx":424E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmagrupador_catalogos.frx":4B28
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
End
Attribute VB_Name = "frmagrupador_catalogos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim numero_items_causas_devolucion As Integer







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
      Call pro_elimina_causas_devolucion
      rs.Open "select * from tb_causas_devolucion", cnn, adOpenDynamic, adLockOptimistic
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
   If var_modifica_registro_agrupador_catalogos = False Then
      rs.Open "SELECT * From tb_agrupadores_catalogos WHERE vcha_agr_agrupador_catalogo_id = '" + Me.txt_agrupador + "'", cnn, adOpenDynamic, adLockOptimistic
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
         Call pro_guardar_causas_devolucion
         rs.Open "select * from tb_causas_devolucion", cnn, adOpenDynamic, adLockOptimistic
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
      MsgBox "Clave de causa de devolución ya existe", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call gPrintListView(lv_catalogos, "LISTADO DE causas_devolucion")
        End If

End Sub

Private Sub cmd_nuevo_Click()
   Call pro_limpiatextos(Me)
   txt_agrupador.Enabled = True
   txt_agrupador.SetFocus: var_modifica_registro_agrupador_catalogos = False
   cmd_guardar.Enabled = True
   cmd_deshacer.Enabled = True
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_agrupador_catalogos = False Then
      var_si = MsgBox("No se han guardado los cambios, żDesea salir?", vbYesNo, "ATENCION")
      If var_si <> 6 Then
         GoTo salir:
      End If
   Else
      If var_hubo_cambios = True Then
         var_si = MsgBox("No se han guardado los cambios, żDesea salir?", vbYesNo, "ATENCION")
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
   numero_items_causas_devolucion = 0
   var_modifica_registro_agrupador_catalogos = True
   lv_catalogos.SmallIcons = ImageList1
   Call pro_llena_listview1
   pro_textos
   rs.Open "select * from TB_AGRUPADORES_CATALOGOS", cnn, adOpenDynamic, adLockOptimistic
   If rs.EOF Then
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
   var_modifica_registro_agrupador_catalogos = False
   Call activa_forma(var_activa_forma_agrupador_catalogos)
End Sub

Private Sub lv_catalogos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_catalogos, ColumnHeader)
End Sub

Private Sub lv_catalogos_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_catalogos.selectedItem = Item
        pro_textos
        var_modifica_registro_agrupador_catalogos = True
        txt_agrupador.Enabled = True
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo err0:
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
      numero_items_causas_devolucion = lv_catalogos.ListItems.Count
      lv_catalogos.ListItems(numero_items_causas_devolucion).Selected = True
      lv_catalogos.selectedItem.EnsureVisible
      pro_textos
   End If
err0:
End Sub


Sub pro_guardar_causas_devolucion()
Dim ok As Boolean
Set TB_AGRUPADOR_CATALOGOS = New TB_AGRUPADOR_CATALOGOS

   If txt_agrupador <> "" And txt_nombre_agrupador <> "" Then
      If var_hubo_cambios Then
         ok = TB_AGRUPADOR_CATALOGOS.Anadir(txt_agrupador, txt_nombre_agrupador)
         If ok Then
            pro_actualiza_ListView
            txt_agrupador.Enabled = False
            MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
            txt_registros = lv_catalogos.ListItems.Count
            var_modifica_registro_agrupador_catalogos = True
         Else
            MsgBox "No se puede grabar registro: " + TB_CAUSAS_DEVOLUCION.MensajeError, vbOKOnly + vbCritical, "ATENCION"
         End If
      End If
   End If
 Set TB_CAUSAS_DEVOLUCION = Nothing: var_hubo_cambios = False

End Sub

Sub pro_elimina_causas_devolucion()
   Dim var_llave_usuarios As String
   Set TB_AGRUPADOR_CATALOGOS = New TB_AGRUPADOR_CATALOGOS
   'On Error GoTo SALIR
   ok = True
   If txt_agrupador <> "" And txt_nombre_agrupador <> "" And var_modifica_registro_agrupador_catalogos = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_AGRUPADOR_CATALOGOS.Eliminar(txt_agrupador)
      Else
         GoTo salir:
      End If
      If ok Then
         MsgBox "Se Elimino Correctamente el Registro", vbInformation
         lv_catalogos.ListItems.Remove (lv_catalogos.selectedItem.Index)
         numero_items_causas_devolucion = numero_items_causas_devolucion - 1
         Call pro_limpiatextos(Me)
         txt_registros = lv_catalogos.ListItems.Count
         If lv_catalogos.ListItems.Count > 0 Then
            lv_catalogos.selectedItem.Selected = True
         End If
         pro_textos
      Else
         MsgBox "No se puede grabar registro: " + TB_CAUSAS_DEVOLUCION.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
Set TB_CAUSAS_DEVOLUCION = Nothing
End Sub


Sub pro_llena_listview1()
   numero_items_causas_devolucion = 0
   Dim list_item As ListItem
   rs.Open "select * from tb_agrupadores_catalogos", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
      Set list_item = lv_catalogos.ListItems.Add(, , rs!vcha_agr_agrupador_catalogo_id)
      list_item.SubItems(1) = IIf(IsNull(rs!vcha_agr_nombre), "", rs!vcha_agr_nombre)
      rs.MoveNext:
     numero_items_causas_devolucion = numero_items_causas_devolucion + 1
   Wend
   rs.Close
End Sub


Sub pro_textos()
On Error GoTo err0:
   Dim var_n As Double
   var_n = lv_catalogos.ListItems.Count
   If var_n > 0 Then
      txt_agrupador = lv_catalogos.selectedItem
      txt_nombre_agrupador = lv_catalogos.selectedItem.SubItems(1)
   End If
   var_numero_renglones = lv_catalogos.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_catalogos.ColumnHeaders(2).Width = 3850
   Else
      lv_catalogos.ColumnHeaders(2).Width = 4099.9
   End If
   var_modifica_registro_agrupador_catalogos = True
   var_hubo_cambios = False
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro_agrupador_catalogos = False Then
        Set list_item = lv_catalogos.ListItems.Add(, , txt_agrupador)
        list_item.SubItems(1) = txt_nombre_agrupador
        list_item.EnsureVisible
        list_item.Selected = True
        numero_items_causas_devolucion = numero_items_causas_devolucion + 1
    Else
        lv_catalogos.ListItems.Item(lv_catalogos.selectedItem.Index).Checked = False
        lv_catalogos.ListItems.Item(lv_catalogos.selectedItem.Index) = txt_agrupador
        lv_catalogos.ListItems.Item(lv_catalogos.selectedItem.Index).ListSubItems(1) = txt_nombre_agrupador
        lv_catalogos.ListItems.Item(lv_catalogos.selectedItem.Index).Selected = True
    End If
End Sub



Private Sub txt_agrupador_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_agrupador_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(lv_catalogos, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub

Private Sub txt_nombre_agrupador_Change()
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_agrupador_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Me.cmd_guardar.Enabled = True Then
         Me.cmd_guardar.SetFocus
      End If
   End If
End Sub
