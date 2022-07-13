VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmdetalle_familia_agrupadores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de familias de agrupadores"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   Icon            =   "frmdetallefamiliaagrupadores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5850
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   90
      TabIndex        =   14
      Top             =   1275
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   1995
         TabIndex        =   15
         Top             =   143
         Width           =   1725
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   4125
         TabIndex        =   16
         Top             =   135
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
               ImageIndex      =   11
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Un Registro Atras"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Un registro adelante"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ir al ultimo"
               ImageIndex      =   14
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Busqueda de agrupador:"
         Height          =   195
         Left            =   195
         TabIndex        =   17
         Top             =   203
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      Picture         =   "frmdetallefamiliaagrupadores.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   465
      Picture         =   "frmdetallefamiliaagrupadores.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Guardar Alt + G"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   795
      Picture         =   "frmdetallefamiliaagrupadores.frx":0ACE
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1125
      Picture         =   "frmdetallefamiliaagrupadores.frx":0BA0
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1455
      Picture         =   "frmdetallefamiliaagrupadores.frx":0CA2
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5445
      Picture         =   "frmdetallefamiliaagrupadores.frx":0DA4
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   30
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
      TabIndex        =   6
      Top             =   45
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   " Detalle de familias de agrupadores "
      Height          =   810
      Left            =   90
      TabIndex        =   0
      Top             =   420
      Width           =   5655
      Begin VB.TextBox txt_nombre_articulo 
         Height          =   315
         Left            =   2265
         MaxLength       =   50
         TabIndex        =   13
         Top             =   360
         Width           =   3330
      End
      Begin VB.TextBox txt_articulo 
         Height          =   315
         Left            =   750
         MaxLength       =   50
         TabIndex        =   1
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   2
         Top             =   375
         Width           =   540
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5370
      Left            =   105
      TabIndex        =   3
      Top             =   1815
      Width           =   5670
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   600
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   14
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdetallefamiliaagrupadores.frx":13DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdetallefamiliaagrupadores.frx":1CB8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdetallefamiliaagrupadores.frx":2592
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdetallefamiliaagrupadores.frx":2B2E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdetallefamiliaagrupadores.frx":3408
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdetallefamiliaagrupadores.frx":3CE2
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdetallefamiliaagrupadores.frx":45BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdetallefamiliaagrupadores.frx":48D6
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdetallefamiliaagrupadores.frx":4BF0
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdetallefamiliaagrupadores.frx":518C
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdetallefamiliaagrupadores.frx":54A6
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdetallefamiliaagrupadores.frx":55B8
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdetallefamiliaagrupadores.frx":56CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdetallefamiliaagrupadores.frx":57DC
               Key             =   ""
            EndProperty
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
               Picture         =   "frmdetallefamiliaagrupadores.frx":58EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdetallefamiliaagrupadores.frx":61C8
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lv_detalle_familia_agrupadores 
         Height          =   5160
         Left            =   45
         TabIndex        =   5
         Top             =   150
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   9102
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ancho"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "alto"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "largo"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   105
      TabIndex        =   4
      Top             =   285
      Width           =   5685
   End
End
Attribute VB_Name = "frmdetalle_familia_agrupadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean





Private Sub cmd_deshacer_Click()
       Call pro_textos

End Sub

Private Sub cmd_eliminar_Click()
        var_resultado = InStr(1, var_menus, Me.Caption)
        var_inicio = var_resultado + Len(Me.Caption) + 3
        If Mid(var_menus, var_inicio, 1) = "1" Then
            Set var_forma = frmdetalle_familia_agrupadores
            var_swpassword = True
            sw_primera_validacion = False
            frmpasswords.Show 1
        Else
            If Mid(var_menus, var_inicio, 2) = "01" Then
                Set var_forma = frmdetalle_familia_agrupadores
                var_swpassword = True
                sw_primera_validacion = False
                frmpasswords2.txt_supervisor = var_supervisor
                frmpasswords2.Show 1
            Else
               Call pro_elimina_detalle_familia_agrupadores
               rs.Open "select * from tb_detalle_familia_agrupadores", cnn, adOpenDynamic, adLockOptimistic
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
        End If

End Sub

Private Sub cmd_guardar_Click()
        var_resultado = InStr(1, var_menus, Me.Caption)
        var_inicio = var_resultado + Len(Me.Caption) + 3
        If Mid(var_menus, var_inicio, 1) = "1" Then
            Set var_forma = frmdetalle_familia_agrupadores
            var_swpassword = True
            sw_primera_validacion = False
            frmpasswords.Show 1
        Else
            If Mid(var_menus, var_inicio, 2) = "01" Then
                Set var_forma = frmdetalle_familia_agrupadores
                var_swpassword = True
                sw_primera_validacion = False
                frmpasswords2.txt_supervisor = var_supervisor
                frmpasswords2.Show 1
            Else
               Call pro_guardar_detalle_familia_agrupadores
               rs.Open "select * from tb_detalle_familia_agrupadores", cnn, adOpenDynamic, adLockOptimistic
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
        End If

End Sub

Private Sub cmd_imprimir_Click()
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call gPrintListView(lv_detalle_familia_agrupadores, "LISTADO DE detalle_familia_agrupadores")
        End If

End Sub

Private Sub cmd_nuevo_Click()
        Call pro_limpiatextos(Me)
        txt_articulo.Enabled = True
        txt_articulo.SetFocus: var_modifica_registro = False
        cmd_guardar.Enabled = True
        cmd_deshacer.Enabled = True
End Sub

Private Sub cmd_salir_Click()
   Unload Me
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

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 2900
   varfamiliaagrupador = frmfamilia_agrupadores.txt_familia_agrupadores(0)
   var_modifica_registro = True
   lv_detalle_familia_agrupadores.SmallIcons = ImageList1
   Call pro_encabezadosView(Me, lv_detalle_familia_agrupadores, False)
   Call pro_llena_listview1
   pro_textos
   rs.Open "select a.VCHA_AGR_AGRUPADOR_ID,b.vcha_agr_nombre from TB_DETALLE_FAMILIA_AGRUPADORES a,TB_AGRUPADORES b where  a.VCHA_AGR_AGRUPADOR_ID = b.VCHA_AGR_AGRUPADOR_ID and  a.VCHA_FAG_FAMILIA_AGRUPADOR_ID = '" & varfamiliaagrupador & "'", cnn, adOpenDynamic, adLockOptimistic
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
    var_modifica_registro = False
    If var_activa_menu = True Then
       Frmmenu2.Enabled = True
    End If
End Sub

Private Sub lv_detalle_familia_agrupadores_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_detalle_familia_agrupadores.selectedItem = Item
        pro_textos
        var_modifica_registro = True
        txt_articulo.Enabled = False

End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        lv_detalle_familia_agrupadores.SetFocus
        Call pro_avanzar(Me, lv_detalle_familia_agrupadores, Button)
        pro_textos
    Case 2
        lv_detalle_familia_agrupadores.SetFocus
        Call pro_avanzar(Me, lv_detalle_familia_agrupadores, Button)
        pro_textos
    Case 3
        Call pro_busca_registro(lv_detalle_familia_agrupadores, txt_buscar, False)
        txt_buscar = ""
        pro_textos
    End Select
End Sub


Sub pro_guardar_detalle_familia_agrupadores()
   Dim ok As Boolean
   rs.Open "select vcha_agr_nombre from TB_AGRUPADORES where  vcha_agr_agrupador_id = '" & txt_articulo & "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      vardetalleagrupador = rs(0).Value
      Set TB_detalle_familia_agrupadores = New TB_detalle_familia_agrupadores
      ok = True
      If txt_articulo <> "" Then
         If var_hubo_cambios Then
            ok = TB_detalle_familia_agrupadores.Anadir(varfamiliaagrupador, txt_articulo)
            If ok Then
                pro_actualiza_ListView
                txt_articulo.Enabled = False
                MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
                txt_registros = lv_detalle_familia_agrupadores.ListItems.Count
                var_modifica_registro = True
            Else
                MsgBox "No se puede grabar registro: " + TB_detalle_familia_agrupadores.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
        End If
    End If
    Else
       MsgBox "El agrupador no existe", vbOKOnly, "ATENCION"
    End If
   rs.Close
    
Set TB_detalle_familia_agrupadores = Nothing: var_hubo_cambios = False

End Sub

Sub pro_elimina_detalle_familia_agrupadores()
   Dim var_llave_usuarios As String
   Set TB_detalle_familia_agrupadores = New TB_detalle_familia_agrupadores
   On Error GoTo salir:
   ok = True
   If txt_articulo <> "" And var_modifica_registro = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_detalle_familia_agrupadores.Eliminar(varfamiliaagrupador, txt_articulo)
      Else
         GoTo salir:
      End If
      If ok Then
        MsgBox "Se Elimino Correctamente el Registro", vbInformation
        lv_detalle_familia_agrupadores.ListItems.Remove (lv_detalle_familia_agrupadores.selectedItem.Index)
        Call pro_limpiatextos(Me)
        txt_registros = lv_detalle_familia_agrupadores.ListItems.Count
        lv_detalle_familia_agrupadores.selectedItem.Selected = True
        pro_textos
      Else
        MsgBox "No se puede eliminar registro: " + TB_detalle_familia_agrupadores.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
   Set TB_detalle_familia_agrupadores = Nothing
End Sub


Sub pro_llena_listview1()
   Dim list_item As ListItem
   rs.Open "select a.VCHA_AGR_AGRUPADOR_ID,b.vcha_agr_nombre from TB_DETALLE_FAMILIA_AGRUPADORES a,TB_AGRUPADORES b where  a.VCHA_AGR_AGRUPADOR_ID = b.VCHA_AGR_AGRUPADOR_ID and  a.VCHA_FAG_FAMILIA_AGRUPADOR_ID = '" & varfamiliaagrupador & "'", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
      Set list_item = lv_detalle_familia_agrupadores.ListItems.Add(, , rs(0).Value): list_item.SmallIcon = 9
      list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
      rs.MoveNext:
    Wend
    rs.Close
End Sub


Sub pro_textos()
On Error GoTo err0:
        txt_articulo = lv_detalle_familia_agrupadores.selectedItem
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem

    If var_modifica_registro = False Then
        Set list_item = lv_detalle_familia_agrupadores.ListItems.Add(, , txt_articulo): list_item.SmallIcon = 9
        list_item.SubItems(1) = vardetalleagrupador
        list_item.EnsureVisible
        list_item.Selected = True
    Else
        lv_detalle_familia_agrupadores.ListItems.Item(lv_detalle_familia_agrupadores.selectedItem.Index).Checked = False
        lv_detalle_familia_agrupadores.ListItems.Item(lv_detalle_familia_agrupadores.selectedItem.Index) = txt_articulo
        lv_detalle_familia_agrupadores.ListItems.Item(lv_detalle_familia_agrupadores.selectedItem.Index).ListSubItems(1) = vardetalleagrupador
        lv_detalle_familia_agrupadores.ListItems.Item(lv_detalle_familia_agrupadores.selectedItem.Index).Selected = True
    End If
End Sub


Private Sub txt_articulo_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_articulo_LostFocus()
   If Trim(txt_articulo) <> "" Then
      rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_articulo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_articulo = IIf(IsNull(rs!VCHA_ART_NOMBRE_ESPAÑOL), "", rs!VCHA_ART_NOMBRE_ESPAÑOL)
      Else
         txt_nombre_articulo = ""
         txt_articulo = ""
         MsgBox "Clave de articulo incorrecto", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      txt_nombre_articulo = ""
   End If
End Sub

Private Sub txt_nombre_articulo_Change()
   var_hubo_cambios = True
End Sub
