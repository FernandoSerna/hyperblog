VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmestampados_textilera 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estampados textilera"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   18
      Top             =   30
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   135
      TabIndex        =   14
      Top             =   1545
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   2070
         TabIndex        =   15
         Top             =   150
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3975
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
         Caption         =   "Busqueda de estampado:"
         Height          =   195
         Left            =   195
         TabIndex        =   17
         Top             =   195
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Estampados "
      Height          =   1065
      Left            =   135
      TabIndex        =   9
      Top             =   450
      Width           =   5655
      Begin VB.TextBox txt_nombre 
         Height          =   315
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   11
         Top             =   630
         Width           =   4155
      End
      Begin VB.TextBox txt_estampados 
         Height          =   315
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   10
         Top             =   270
         Width           =   975
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Index           =   2
         Left            =   345
         TabIndex        =   13
         Top             =   675
         Width           =   885
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   1
         Left            =   345
         TabIndex        =   12
         Top             =   315
         Width           =   450
      End
   End
   Begin VB.Frame Frame3 
      Height          =   240
      Left            =   135
      TabIndex        =   7
      Top             =   2115
      Visible         =   0   'False
      Width           =   5655
      Begin MSComctlLib.ListView lv_estampados 
         Height          =   30
         Left            =   45
         TabIndex        =   8
         Top             =   165
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   53
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
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      Picture         =   "frmestampados_textilera.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   465
      Picture         =   "frmestampados_textilera.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Guardar Alt + G"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   795
      Picture         =   "frmestampados_textilera.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1125
      Picture         =   "frmestampados_textilera.frx":02D6
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1455
      Picture         =   "frmestampados_textilera.frx":03D8
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5445
      Picture         =   "frmestampados_textilera.frx":04DA
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   750
      Top             =   4830
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
            Picture         =   "frmestampados_textilera.frx":0B14
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmestampados_textilera.frx":13EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   855
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
            Picture         =   "frmestampados_textilera.frx":1CC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmestampados_textilera.frx":25A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmestampados_textilera.frx":2E7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmestampados_textilera.frx":3418
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmestampados_textilera.frx":3CF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmestampados_textilera.frx":45CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmestampados_textilera.frx":4EA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmestampados_textilera.frx":4FBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmestampados_textilera.frx":50CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmestampados_textilera.frx":51DE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   135
      TabIndex        =   6
      Top             =   285
      Width           =   5655
   End
End
Attribute VB_Name = "frmestampados_textilera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim numero_items_estampados As Integer
Dim bitacora As Boolean



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
      Call pro_elimina_estampados
      rs.Open "select * from tb_estampados", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
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
   If var_modifica_registro_estampado = False Then
      rs.Open "select * from tb_estampados where vcha_est_estampado_id = '" + Me.txt_estampados + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
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
         Call pro_guardar_estampados
         'rs.Open "select * from tb_estampados", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
         'If rs.BOF Then
         '   cmd_guardar.Enabled = False
         '   cmd_deshacer.Enabled = False
         '   cmd_eliminar.Enabled = False
         'Else
         '   cmd_guardar.Enabled = True
         '   cmd_deshacer.Enabled = True
         '   cmd_eliminar.Enabled = True
         'End If
         'rs.Close
      End If
   Else
      MsgBox "Clave de estampado ya existe", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call gPrintListView(lv_estampados, "LISTADO DE estampados")
        End If

End Sub

Private Sub cmd_nuevo_Click()
        Call pro_limpiatextos(Me)
        txt_estampados.Enabled = True
        txt_estampados.SetFocus: var_modifica_registro_estampado = False
        cmd_guardar.Enabled = True
        cmd_deshacer.Enabled = True
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_estampado = False Then
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
   If Shift = 4 And KeyCode = 69 Then
      cmd_eliminar_Click
   End If
   If Shift = 4 And KeyCode = 73 Then
      cmd_imprimir_Click
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 3000
   Left = 2900
   var_modifica_registro_estampado = True
   lv_estampados.SmallIcons = ImageList1
   'Call pro_encabezadosView(Me, lv_estampados, False)
   'Call pro_llena_listview1
   'pro_textos
'   rs.Open "select * from tb_estampados", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
'   If rs.BOF Then
'      cmd_guardar.Enabled = False
'      cmd_deshacer.Enabled = False
'      cmd_eliminar.Enabled = False
'   Else
'      cmd_guardar.Enabled = True
'      cmd_deshacer.Enabled = True
'      cmd_eliminar.Enabled = True
'   End If
'   rs.Close
   Me.txt_estampados.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_despliega_menu = True Then
      var_swpassword = False
      var_modifica_registro_estampado = False
   End If
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub lv_estampados_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_estampados, ColumnHeader)
End Sub




Sub pro_guardar_estampados()

Dim ok As Boolean

Set TB_ESTAMPADOS_TEXTILERA = New TB_ESTAMPADOS_TEXTILERA
Set TB_BITACORA_ESTAMPADOS = New TB_BITACORA_ESTAMPADOS
    ok = True
    If txt_estampados <> "" And txt_nombre <> "" Then
        If var_hubo_cambios Then
            rs.Open "Select * from tb_estampados where vcha_est_estampado_id = '" + txt_estampados + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
            ok = TB_ESTAMPADOS_TEXTILERA.Anadir(txt_estampados, txt_nombre)
            If ok Then
                bitacora = True
                If var_modifica_registro_estampado = False Then
                   var_operacion_bitacora = "I"
                   bitacora = TB_BITACORA_ESTAMPADOS.Anadir(txt_estampados, "VCHA_EST_NOMBRE", var_operacion_bitacora, "", txt_nombre, var_clave_usuario_global, fun_NombrePc, Date)
                Else
                   var_operacion_bitacora = "M"
                   If rs(0) <> txt_estampados Then
                      bitacora = TB_BITACORA_ESTAMPADOS.Anadir(txt_estampados, "VCHA_EST_ESTAMPADO_ID", var_operacion_bitacora, rs(0), txt_estampados, var_clave_usuario_global, fun_NombrePc, Date)
                   End If
                   If rs(1) <> txt_nombre Then
                      bitacora = TB_BITACORA_ESTAMPADOS.Anadir(txt_estampados, "VCHA_EST_NOMBRE", var_operacion_bitacora, rs(1), txt_nombre, var_clave_usuario_global, fun_NombrePc, Date)
                   End If
                End If
                rs.Close
                
                pro_actualiza_ListView
                txt_estampados.Enabled = False
                MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
                txt_registros = lv_estampados.ListItems.Count
                var_modifica_registro_estampado = True
            Else
                MsgBox "No se puede grabar registro: " + TB_ESTAMPADOS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
        End If
    End If
    
Set TB_ESTAMPADOS = Nothing: var_hubo_cambios = False

End Sub

Sub pro_elimina_estampados()
   Dim var_llave_usuarios As String
   Set TB_ESTAMPADOS = New TB_ESTAMPADOS
   Set TB_BITACORA_ESTAMPADOS = New TB_BITACORA_ESTAMPADOS
   'On Error GoTo SALIR:
   ok = True
   If txt_estampados <> "" And txt_nombre <> "" And var_modifica_registro_estampado = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_ESTAMPADOS.Eliminar(txt_estampados)
      Else
         GoTo salir:
      End If
      If ok Then
         var_operacion_bitacora = "E"
         bitacora = TB_BITACORA_ESTAMPADOS.Anadir(txt_estampados, "VCHA_EST_NOMBRE", var_operacion_bitacora, txt_nombre, "", var_clave_usuario_global, fun_NombrePc, Date)
         numero_items_estampados = numero_items_estampados - 1
         MsgBox "Se Elimino Correctamente el Registro", vbInformation
         Call pro_limpiatextos(Me)
         txt_registros = lv_estampados.ListItems.Count
         If lv_estampados.ListItems.Count > 0 Then
            lv_estampados.selectedItem.Selected = True
         End If
      Else
        MsgBox "No se puede eliminar registro: " + TB_ESTAMPADOS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
   Set TB_ESTAMPADOS = Nothing
End Sub


Sub pro_llena_listview1()
'   Dim list_item As ListItem
 '  rs.Open "select * from tb_estampados", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
  ' numero_items_estampados = 0
   'While Not rs.EOF
'      Set list_item = lv_estampados.ListItems.Add(, , rs(0).Value)
 '     list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
  '    rs.MoveNext:
   '   numero_items_estampados = numero_items_estampados + 1
    'Wend
    'rs.Close
End Sub



Private Sub pro_actualiza_ListView()
'Dim list_item As ListItem

 '   If var_modifica_registro_estampado = False Then
  '      Set list_item = lv_estampados.ListItems.Add(, , txt_estampados)
   '     list_item.SubItems(1) = txt_nombre
    '    list_item.EnsureVisible
     '   list_item.Selected = True
      '  numero_items_estampados = numero_items_estampados + 1
'    Else
 '       lv_estampados.ListItems.Item(lv_estampados.selectedItem.Index).Checked = False
  '      lv_estampados.ListItems.Item(lv_estampados.selectedItem.Index) = txt_estampados
   '     lv_estampados.ListItems.Item(lv_estampados.selectedItem.Index).ListSubItems(1) = txt_nombre
    '    lv_estampados.ListItems.Item(lv_estampados.selectedItem.Index).Selected = True
    'End If
End Sub

Private Sub txt_estampados_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
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



Private Sub txt_estampados_LostFocus()
   If Trim(Me.txt_estampados) <> "" Then
      rs.Open "select * from tb_estampados where vcha_est_estampado_id = '" + Me.txt_estampados + "'", cnn_etiquetas_textilera, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
         var_modifica_registro_estampado = True
      Else
         Me.txt_nombre = ""
         var_modifica_registro_estampado = False
      End If
      rs.Close
   Else
      Me.txt_nombre = ""
   End If
End Sub

Private Sub txt_nombre_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub
