VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmprioridades 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prioridades"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   Icon            =   "frmprioridades.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5955
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmprioridades.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   465
      Picture         =   "frmprioridades.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Guardar Alt + G"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   795
      Picture         =   "frmprioridades.frx":0ACE
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
      Picture         =   "frmprioridades.frx":0BA0
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1455
      Picture         =   "frmprioridades.frx":0CA2
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
      Picture         =   "frmprioridades.frx":0DA4
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   18
      Top             =   285
      Width           =   5655
   End
   Begin VB.Frame Frame3 
      Height          =   5325
      Left            =   150
      TabIndex        =   16
      Top             =   1890
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
               Picture         =   "frmprioridades.frx":13DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmprioridades.frx":1CB8
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lv_prioridades 
         Height          =   5130
         Left            =   30
         TabIndex        =   17
         Top             =   135
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   9049
         View            =   3
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
            Text            =   "Clasificación"
            Object.Width           =   4904
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Valor"
            Object.Width           =   4762
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "alto"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "largo"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList2 
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
            NumListImages   =   10
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmprioridades.frx":2592
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmprioridades.frx":2E6C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmprioridades.frx":3746
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmprioridades.frx":3CE2
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmprioridades.frx":45BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmprioridades.frx":4E96
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmprioridades.frx":5770
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmprioridades.frx":5A8A
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmprioridades.frx":5DA4
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmprioridades.frx":6340
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList3 
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
               Picture         =   "frmprioridades.frx":665A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmprioridades.frx":6F34
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Prioridades "
      Height          =   945
      Left            =   150
      TabIndex        =   13
      Top             =   405
      Width           =   5655
      Begin VB.TextBox txt_prioridades 
         Height          =   315
         Index           =   0
         Left            =   1935
         MaxLength       =   3
         TabIndex        =   7
         Top             =   225
         Width           =   690
      End
      Begin VB.TextBox txt_prioridades 
         Height          =   315
         Index           =   1
         Left            =   1935
         TabIndex        =   8
         Top             =   555
         Width           =   1020
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clasificación:"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   15
         Top             =   240
         Width           =   930
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   14
         Top             =   585
         Width           =   405
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   150
      TabIndex        =   9
      Top             =   1350
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   1905
         TabIndex        =   10
         Top             =   150
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3690
         TabIndex        =   11
         Top             =   165
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
         Caption         =   "Busqueda de Prioridad:"
         Height          =   195
         Left            =   195
         TabIndex        =   12
         Top             =   195
         Width           =   1650
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2700
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   90
      Width           =   255
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   -75
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
            Picture         =   "frmprioridades.frx":780E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmprioridades.frx":80E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmprioridades.frx":89C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmprioridades.frx":8F5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmprioridades.frx":983A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmprioridades.frx":A114
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmprioridades.frx":A9EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmprioridades.frx":AB00
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmprioridades.frx":AC12
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmprioridades.frx":AD24
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmprioridades.frx":AE36
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmprioridades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim numero_items_prioridades As Integer
Dim bitacora As Boolean
Dim var_mismo_nombre_caja As String





Private Sub cmd_deshacer_Click()
       Call pro_textos

End Sub

Private Sub cmd_eliminar_Click()
   var_resultado = InStr(1, var_menus, Me.Caption)
   var_inicio = var_resultado + Len(Me.Caption) + 3
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
      Call pro_elimina_prioridades
      rs.Open "select * from tb_prioridades", cnn, adOpenDynamic, adLockOptimistic
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
      MsgBox "Imposible ejecutar la acción solicitada", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_guardar_Click()
   var_resultado = InStr(1, var_menus, Me.Caption)
   var_inicio = var_resultado + Len(Me.Caption) + 3
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
      Call pro_guardar_prioridades
      rs.Open "select * from tb_prioridades", cnn, adOpenDynamic, adLockOptimistic
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
      MsgBox "Imposible ejecutar la acción solicitada", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call gPrintListView(lv_prioridades, "LISTADO DE prioridades")
        End If

End Sub

Private Sub cmd_nuevo_Click()
        Call pro_limpiatextos(Me)
        txt_prioridades(0).Enabled = True
        txt_prioridades(0).SetFocus: var_modifica_registro_prioridad = False
        cmd_guardar.Enabled = True
        cmd_deshacer.Enabled = True
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_prioridad = False Then
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
   var_modifica_registro_prioridad = True
   lv_prioridades.SmallIcons = ImageList1
   Call pro_encabezadosView(Me, lv_prioridades, False)
   Call pro_llena_listview1
   pro_textos
   rs.Open "select * from tb_prioridades", cnn, adOpenDynamic, adLockOptimistic
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
      var_modifica_registro_prioridad = False
   End If
   Call activa_forma(var_activa_forma_prioridades)
End Sub

Private Sub lv_prioridades_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_prioridades.selectedItem = Item
        pro_textos
        var_modifica_registro_prioridad = True
        txt_prioridades(0).Enabled = False

End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_prioridades.SetFocus
      Call pro_avanzar(Me, lv_prioridades, Button)
      lv_prioridades.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_prioridades.ListItems(1).Selected = True
      lv_prioridades.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 4 Then
      numero_items_prioridades = lv_prioridades.ListItems.Count
      lv_prioridades.ListItems(numero_items_prioridades).Selected = True
      lv_prioridades.selectedItem.EnsureVisible
      pro_textos
   End If
err0:
End Sub


Sub pro_guardar_prioridades()

Dim ok As Boolean

Set TB_PRIORIDADES = New TB_PRIORIDADES
Set TB_BITACORA_PRIORIDADES = New TB_BITACORA_PRIORIDADES
    ok = True
    If txt_prioridades(0) <> "" And txt_prioridades(1) <> "" And txt_prioridades(2) <> "" Then
        If var_hubo_cambios Then
            rs.Open "select * from tb_prioridades where char_pri_prioridad_id = '" + txt_prioridades(0) + "'", cnn, adOpenDynamic, adLockOptimistic
            ok = TB_PRIORIDADES.Anadir(txt_prioridades(0), txt_prioridades(1), txt_prioridades(2))
            If ok Then
                bitacora = True
                If var_modifica_registro_prioridad = False Then
                   var_operacion_bitacora = "I"
                   bitacora = TB_BITACORA_PRIORIDADES.Anadir(txt_prioridades(0), "VCHA_PRI_NOMBRE", var_operacion_bitacora, "", txt_prioridades(1), var_clave_usuario_global, fun_NombrePc, Date)
                Else
                   var_operacion_bitacora = "M"
                   If rs(0) <> txt_prioridades(0) Then
                      bitacora = TB_BITACORA_PRIORIDADES.Anadir(txt_prioridades(0), "VCHA_PRI_PRIORIDAD_ID", var_operacion_bitacora, rs(0), txt_prioridades(0), var_clave_usuario_global, fun_NombrePc, Date)
                   End If
                   If rs(1) <> txt_prioridades(1) Then
                      bitacora = TB_BITACORA_PRIORIDADES.Anadir(txt_prioridades(0), "VCHA_PRI_NOMBRE", var_operacion_bitacora, rs(1), txt_prioridades(1), var_clave_usuario_global, fun_NombrePc, Date)
                   End If
                   If rs(2) <> txt_prioridades(2) Then
                      bitacora = TB_BITACORA_PRIORIDADES.Anadir(txt_prioridades(0), "VCHA_PRI_NUMERO_ORDEN", var_operacion_bitacora, rs(2), txt_prioridades(2), var_clave_usuario_global, fun_NombrePc, Date)
                   End If
                End If
                rs.Close
                pro_actualiza_ListView
                txt_prioridades(0).Enabled = False
                MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
                txt_registros = lv_prioridades.ListItems.Count
                var_modifica_registro_prioridad = True
            Else
                MsgBox "No se puede grabar registro: " + TB_PRIORIDADES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
        End If
    End If
    
Set TB_PRIORIDADES = Nothing: var_hubo_cambios = False

End Sub
 
Sub pro_elimina_prioridades()
   Dim var_llave_usuarios As String
   Set TB_PRIORIDADES = New TB_PRIORIDADES
   Set TB_BITACORA_PRIORIDADES = New TB_BITACORA_PRIORIDADES
   On Error GoTo salir:
   ok = True
   If txt_prioridades(0) <> "" And txt_prioridades(1) <> "" And var_modifica_registro_prioridad = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_PRIORIDADES.Eliminar(txt_prioridades(0))
      Else
         GoTo salir:
      End If
      If ok Then
        bitacora = True
        var_operacion_bitacora = "E"
        bitacora = TB_BITACORA_PRIORIDADES.Anadir(txt_prioridades(0), "VCHA_PRI_NOMBRE", var_operacion_bitacora, "", txt_prioridades(1), var_clave_usuario_global, fun_NombrePc, Date)
        numero_items_prioridades = numero_items_prioridades - 1
        MsgBox "Se Elimino Correctamente el Registro", vbInformation
        lv_prioridades.ListItems.Remove (lv_prioridades.selectedItem.Index)
        Call pro_limpiatextos(Me)
        txt_registros = lv_prioridades.ListItems.Count
        lv_prioridades.selectedItem.Selected = True
        pro_textos
      Else
        MsgBox "No se puede eliminar registro: " + TB_PRIORIDADES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
   Set TB_PRIORIDADES = Nothing
End Sub


Sub pro_llena_listview1()
   Dim list_item As ListItem
   rs.Open "select * from TB_prioridades", cnn, adOpenDynamic, adLockOptimistic
   numero_items_prioridades = 0
   While Not rs.EOF
      Set list_item = lv_prioridades.ListItems.Add(, , rs(0).Value)
      list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
      rs.MoveNext:
      numero_items_prioridades = numero_items_prioridades + 1
    Wend
    rs.Close
End Sub


Sub pro_textos()
On Error GoTo err0:
        txt_prioridades(0) = lv_prioridades.selectedItem
        txt_prioridades(1) = lv_prioridades.selectedItem.SubItems(1)
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem

    If var_modifica_registro_prioridad = False Then
        Set list_item = lv_prioridades.ListItems.Add(, , txt_prioridades(0))
        list_item.SubItems(1) = txt_prioridades(1)
        list_item.EnsureVisible
        list_item.Selected = True
        numero_items_prioridades = numero_items_prioridades + 1
    Else
        lv_prioridades.ListItems.Item(lv_prioridades.selectedItem.Index).Checked = False
        lv_prioridades.ListItems.Item(lv_prioridades.selectedItem.Index) = txt_prioridades(0)
        lv_prioridades.ListItems.Item(lv_prioridades.selectedItem.Index).ListSubItems(1) = txt_prioridades(1)
        lv_prioridades.ListItems.Item(lv_prioridades.selectedItem.Index).Selected = True
    End If
'    lv_prioridades.SetFocus
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_prioridades, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub

Private Sub txt_prioridades_Change(Index As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub

Private Sub txt_prioridades_GotFocus(Index As Integer)
   If Index = 1 Then
      var_mismo_nombre_caja = txt_prioridades(1)
   End If
End Sub

Private Sub txt_prioridades_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Index < 2 Then
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

Private Sub txt_prioridades_LostFocus(Index As Integer)
   If Index = 1 Then
      If var_mismo_nombre_caja <> txt_prioridades(1) Then
         rs.Open "select * from tb_prioridades where vcha_pri_nombre  = '" + txt_prioridades(1) + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            MsgBox "Ya existe una caja con el nombre " + txt_prioridades(1), vbOKOnly, "ATENCION"
            txt_prioridades(1) = var_mismo_nombre_caja
         End If
         rs.Close
      End If
   End If
End Sub

