VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsublineas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sublineas"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   Icon            =   "frmsublineas.frx":0000
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
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5460
      Picture         =   "frmsublineas.frx":08CA
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
      Picture         =   "frmsublineas.frx":0F04
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
      Picture         =   "frmsublineas.frx":1006
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
      Picture         =   "frmsublineas.frx":1108
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
      Picture         =   "frmsublineas.frx":11DA
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
      Picture         =   "frmsublineas.frx":12DC
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
      Left            =   2610
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   19
      Top             =   135
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   " Sublineas"
      Height          =   1395
      Left            =   150
      TabIndex        =   0
      Top             =   420
      Width           =   5655
      Begin VB.TextBox txt_nombre_linea 
         Height          =   315
         Left            =   2010
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         Top             =   240
         Width           =   3570
      End
      Begin VB.TextBox txt_nombre_sublinea 
         Height          =   315
         Left            =   780
         MaxLength       =   50
         TabIndex        =   10
         Top             =   960
         Width           =   4770
      End
      Begin VB.TextBox txt_sublinea 
         Height          =   315
         Left            =   780
         MaxLength       =   50
         TabIndex        =   9
         Top             =   600
         Width           =   1200
      End
      Begin VB.TextBox txt_linea 
         Height          =   315
         Left            =   780
         MaxLength       =   50
         TabIndex        =   7
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   13
         Top             =   1020
         Width           =   600
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   12
         Top             =   660
         Width           =   450
      End
      Begin VB.Label lab_estados 
         AutoSize        =   -1  'True
         Caption         =   "Linea:"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   11
         Top             =   300
         Width           =   435
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   150
      TabIndex        =   14
      Top             =   1845
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   1875
         TabIndex        =   18
         Top             =   150
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3825
         TabIndex        =   21
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
         Caption         =   "Busqueda de sublinea:"
         Height          =   195
         Left            =   195
         TabIndex        =   15
         Top             =   195
         Width           =   1620
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4860
      Left            =   150
      TabIndex        =   16
      Top             =   2355
      Width           =   5655
      Begin MSComctlLib.ListView lv_sublineas 
         Height          =   4650
         Left            =   45
         TabIndex        =   20
         Top             =   150
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   8202
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
            Text            =   "sublinea"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   4395
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
            Picture         =   "frmsublineas.frx":13DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsublineas.frx":1CB8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   17
      Top             =   285
      Width           =   5655
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
            Picture         =   "frmsublineas.frx":2592
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsublineas.frx":2E6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsublineas.frx":3746
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsublineas.frx":3CE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsublineas.frx":45BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsublineas.frx":4E98
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsublineas.frx":5772
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsublineas.frx":5884
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsublineas.frx":5996
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsublineas.frx":5AA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsublineas.frx":5BBA
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmsublineas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim numero_items_sublineas As Integer




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
      Call pro_elimina_sublineas
      rs.Open "select * from tb_sublineas", cnn, adOpenDynamic, adLockOptimistic
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
   If var_modifica_registro_sublinea = False Then
      rs.Open "select * from tb_sublineas where VCHA_SLI_SUBLINEA_ID = '" + txt_sublinea + "'", cnn, adOpenDynamic, adLockOptimistic
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
         Call pro_guardar_sublineas
         rs.Open "select * from tb_sublineas", cnn, adOpenDynamic, adLockOptimistic
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
     MsgBox "Clave de subliena ya existe", vbOKOnly, "ATENCION"
  End If
End Sub

Private Sub cmd_imprimir_Click()
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call gPrintListView(lv_sublineas, "LISTADO DE sublineas")
        End If

End Sub

Private Sub cmd_nuevo_Click()
   Call pro_limpiatextos(Me)
   txt_linea.Enabled = True
   txt_linea.SetFocus: var_modifica_registro_sublinea = False
   txt_sublinea.Enabled = True
   txt_nombre_sublinea.Enabled = True
   Me.txt_nombre_linea.Enabled = True
   cmd_guardar.Enabled = True
   cmd_deshacer.Enabled = True
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_sublinea = False Then
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
   var_modifica_registro_sublinea = True
   lv_sublineas.SmallIcons = ImageList
   Call pro_encabezadosView(Me, lv_sublineas, False)
   Call pro_llena_listview1
   pro_textos
   rs.Open "select * from tb_sublineas", cnn, adOpenDynamic, adLockOptimistic
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
   Me.txt_linea.Enabled = False
   Me.txt_nombre_linea.Enabled = False
   Me.txt_nombre_sublinea.Enabled = False
   Me.txt_sublinea.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_despliega_menu = True Then
      var_swpassword = False
      var_modifica_registro_sublinea = False
   End If
   Call activa_forma(var_activa_forma_sublineas)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         txt_linea = lv_lista.selectedItem
         txt_nombre_linea = lv_lista.selectedItem.SubItems(1)
      Else
         txt_linea = ""
         txt_nombre_linea = ""
      End If
      txt_linea.SetFocus
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub lv_sublineas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_sublineas, ColumnHeader)
End Sub

Private Sub lv_sublineas_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_sublineas.selectedItem = Item
        pro_textos
        var_modifica_registro_sublinea = True
        txt_linea.Enabled = False

End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_sublineas.SetFocus
      Call pro_avanzar(Me, lv_sublineas, Button)
      lv_sublineas.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_sublineas.ListItems(1).Selected = True
      lv_sublineas.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 4 Then
      numero_items_sublineas = lv_sublineas.ListItems.Count
      lv_sublineas.ListItems(numero_items_sublineas).Selected = True
      lv_sublineas.selectedItem.EnsureVisible
      pro_textos
   End If
err0:
End Sub


Sub pro_guardar_sublineas()
   Dim ok As Boolean
   Set tb_sublineas = New tb_sublineas
   Set TB_BITACORA_SUBLINEAS = New TB_BITACORA_SUBLINEAS
   ok = True
   If txt_linea <> "" And txt_sublinea <> "" And txt_nombre_sublinea <> "" Then
      If var_hubo_cambios Then
         rs.Open "select * from tb_sublineas where vcha_sli_sublinea_id = '" + txt_sublinea + "'", cnn, adOpenDynamic, adLockOptimistic
         ok = tb_sublineas.Anadir(txt_linea, txt_sublinea, txt_nombre_sublinea)
         If ok Then
            bitacora = True
            If var_modifica_registro_sublinea = False Then
               var_operacion_bitacora = "I"
               bitacora = TB_BITACORA_SUBLINEAS.Anadir(txt_sublinea, "VCHA_SLI_NOMBRE", var_operacion_bitacora, "", txt_nombre_sublinea, var_clave_usuario_global, fun_NombrePc, Date)
            Else
               var_operacion_bitacora = "M"
               If rs(0) <> txt_linea Then
                  bitacora = TB_BITACORA_SUBLINEAS.Anadir(txt_sublinea, "VCHA_LIN_LINEA_ID", var_operacion_bitacora, rs(0), txt_linea, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(1) <> txt_sublinea Then
                  bitacora = TB_BITACORA_SUBLINEAS.Anadir(txt_sublinea, "VCHA_SLI_SUBLINEA_ID", var_operacion_bitacora, rs(1), txt_sublinea, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(2) <> txt_nombre_sublinea Then
                  bitacora = TB_BITACORA_SUBLINEAS.Anadir(txt_sublinea, "VCHA_SLI_NOMBRE", var_operacion_bitacora, rs(2), txt_nombre_sublinea, var_clave_usuario_global, fun_NombrePc, Date)
               End If
            End If
            rs.Close
            pro_actualiza_ListView
            txt_linea.Enabled = False
            MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
            txt_registros = lv_sublineas.ListItems.Count
            var_modifica_registro_sublinea = True
         Else
            MsgBox "No se puede grabar registro: " + tb_sublineas.MensajeError, vbOKOnly + vbCritical, "ATENCION"
         End If
      End If
   End If
   Set tb_sublineas = Nothing: var_hubo_cambios = False
End Sub

Sub pro_elimina_sublineas()
   Dim var_llave_usuarios As String
   Set tb_sublineas = New tb_sublineas
   Set TB_BITACORA_SUBLINEAS = New TB_BITACORA_SUBLINEAS
   ok = True
   On Error GoTo salir:
   If txt_linea <> "" And txt_sublinea <> "" And txt_nombre_sublinea <> "" And var_modifica_registro_sublinea = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = tb_sublineas.Eliminar(txt_sublinea)
      Else
         GoTo salir:
      End If
      If ok Then
         var_operacion_bitacora = "E"
         bitacora = TB_BITACORA_SUBLINEAS.Anadir(txt_sublinea, "VCHA_SLI_NOMBRE", var_operacion_bitacora, txt_nombre_sublinea, "", var_clave_usuario_global, fun_NombrePc, Date)
         numero_items_sublineas = numero_items_sublineas - 1
         MsgBox "Se Elimino Correctamente el Registro", vbInformation
         lv_sublineas.ListItems.Remove (lv_sublineas.selectedItem.Index)
         Call pro_limpiatextos(Me)
         txt_registros = lv_sublineas.ListItems.Count
         lv_sublineas.selectedItem.Selected = True
         pro_textos
      Else
         MsgBox "No se puede eliminar registro: " + tb_sublineas.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
   Set tb_sublineas = Nothing
End Sub


Sub pro_llena_listview1()
   Dim list_item As ListItem
   rs.Open "select * from TB_sublineas", cnn, adOpenDynamic, adLockOptimistic
   numero_items_sublineas = 0
   While Not rs.EOF
      Set list_item = lv_sublineas.ListItems.Add(, , rs(1).Value)
      list_item.SubItems(1) = IIf(IsNull(rs(2).Value), "", rs(2).Value)
      list_item.SubItems(2) = IIf(IsNull(rs(0).Value), "", rs(0).Value)
      rs.MoveNext:
      numero_items_sublineas = numero_items_sublineas + 1
    Wend
    rs.Close
End Sub


Sub pro_textos()
'On Error GoTo err0:
Dim var_n As Double
   var_n = lv_sublineas.ListItems.Count
   If var_n > 0 Then
      txt_linea = lv_sublineas.selectedItem.SubItems(2)
      txt_sublinea = lv_sublineas.selectedItem
      txt_nombre_sublinea = lv_sublineas.selectedItem.SubItems(1)
      rs.Open "select * from tb_lineas where vcha_lin_linea_id = '" + txt_linea + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_linea = IIf(IsNull(rs!vcha_lin_nombre), "", rs!vcha_lin_nombre)
      Else
         txt_nombre_linea = ""
      End If
      rs.Close
      txt_linea.Enabled = False
      txt_nombre_linea.Enabled = False
      txt_sublinea.Enabled = False
      Me.txt_nombre_sublinea.Enabled = True
   End If
   var_numero_renglones = lv_sublineas.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_sublineas.ColumnHeaders(2).Width = 3850
   Else
      lv_sublineas.ColumnHeaders(2).Width = 4099.71
   End If
   var_hubo_cambios = False
   var_modifica_registro_sublinea = True
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem

    If var_modifica_registro_sublinea = False Then
        Set list_item = lv_sublineas.ListItems.Add(, , txt_sublinea)
        list_item.SubItems(1) = txt_nombre_sublinea
        list_item.SubItems(2) = txt_linea
        list_item.EnsureVisible
        list_item.Selected = True
        numero_items_sublineas = numero_items_sublineas + 1
    Else
        lv_sublineas.ListItems.Item(lv_sublineas.selectedItem.Index).Checked = False
        lv_sublineas.ListItems.Item(lv_sublineas.selectedItem.Index) = txt_sublinea
        lv_sublineas.ListItems.Item(lv_sublineas.selectedItem.Index).ListSubItems(1) = txt_nombre_sublinea
        lv_sublineas.ListItems.Item(lv_sublineas.selectedItem.Index).ListSubItems(2) = txt_linea
        lv_sublineas.ListItems.Item(lv_sublineas.selectedItem.Index).Selected = True
    End If
'    lv_sublineas.SetFocus
End Sub



Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_sublineas, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub

Private Sub txt_linea_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_linea_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para accesar al catálogo"
End Sub

Private Sub txt_linea_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_lineas order by vcha_lin_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_LIN_LINEA_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_lin_nombre), "", rs!vcha_lin_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "LINEAS"
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
      var_activa_forma_lineas = Me.Name
      Me.Enabled = False
      frmlineas.Show
   End If
End Sub


Private Sub txt_linea_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_linea_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_linea) <> "" Then
      rs.Open "select * from tb_lineas where vcha_lin_linea_id = '" + txt_linea + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_linea = IIf(IsNull(rs!vcha_lin_nombre), "", rs!vcha_lin_nombre)
      Else
         txt_linea = ""
         txt_nombre_linea = ""
         MsgBox "Clave de linea incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      txt_nombre_linea = ""
   End If
End Sub

Private Sub txt_nombre_linea_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_linea_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_nombre_linea_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_lineas order by vcha_lin_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_LIN_LINEA_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_lin_nombre), "", rs!vcha_lin_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "LINEAS"
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
      var_activa_forma_lineas = Me.Name
      Me.Enabled = False
      frmlineas.Show
   End If
End Sub

Private Sub txt_nombre_linea_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_linea_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_sublinea_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_sublinea_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Me.cmd_guardar.Enabled = True Then
         If Me.cmd_guardar.Enabled = True Then
            Me.cmd_guardar.SetFocus
         End If
      End If
   End If
End Sub

Private Sub txt_sublinea_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_sublinea_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub
