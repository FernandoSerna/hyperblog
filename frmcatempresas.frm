VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcatempresas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cat?logo de empresas"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   Icon            =   "frmcatempresas.frx":0000
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
      Picture         =   "frmcatempresas.frx":08CA
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
      Picture         =   "frmcatempresas.frx":09CC
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
      Picture         =   "frmcatempresas.frx":0ACE
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
      Picture         =   "frmcatempresas.frx":0BA0
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
      Picture         =   "frmcatempresas.frx":0CA2
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5430
      Picture         =   "frmcatempresas.frx":0DA4
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
      Left            =   1365
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   21
      Top             =   6135
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   " Empresas "
      Height          =   2295
      Left            =   150
      TabIndex        =   0
      Top             =   465
      Width           =   5655
      Begin VB.TextBox txt_empresas 
         Height          =   315
         Index           =   5
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   23
         Top             =   1905
         Width           =   4260
      End
      Begin VB.TextBox txt_empresas 
         Height          =   315
         Index           =   4
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   17
         Top             =   1575
         Width           =   4260
      End
      Begin VB.TextBox txt_empresas 
         Height          =   315
         Index           =   3
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   16
         Top             =   1245
         Width           =   1560
      End
      Begin VB.TextBox txt_empresas 
         Height          =   315
         Index           =   2
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   9
         Top             =   915
         Width           =   4245
      End
      Begin VB.TextBox txt_empresas 
         Height          =   315
         Index           =   1
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   8
         Top             =   585
         Width           =   4245
      End
      Begin VB.TextBox txt_empresas 
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
         Caption         =   "Gerente:"
         Height          =   195
         Index           =   6
         Left            =   165
         TabIndex        =   20
         Top             =   1935
         Width           =   615
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Giro:"
         Height          =   195
         Index           =   5
         Left            =   165
         TabIndex        =   19
         Top             =   1605
         Width           =   330
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "RFC:"
         Height          =   195
         Index           =   4
         Left            =   165
         TabIndex        =   18
         Top             =   1275
         Width           =   360
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Direcci?n:"
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   12
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   11
         Top             =   615
         Width           =   600
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   10
         Top             =   255
         Width           =   450
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   150
      TabIndex        =   13
      Top             =   2760
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   1905
         TabIndex        =   26
         Top             =   150
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3705
         TabIndex        =   25
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
         Caption         =   "Busqueda de empresa:"
         Height          =   195
         Left            =   195
         TabIndex        =   14
         Top             =   195
         Width           =   1635
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3915
      Left            =   150
      TabIndex        =   15
      Top             =   3285
      Width           =   5655
      Begin MSComctlLib.ListView lv_empresas 
         Height          =   3735
         Left            =   45
         TabIndex        =   22
         Top             =   135
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   6588
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
         NumItems        =   6
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
            Text            =   "direccion"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "rfc"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "giro"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "gerente"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3660
      Top             =   1710
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
            Picture         =   "frmcatempresas.frx":13DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcatempresas.frx":1CB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcatempresas.frx":2592
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcatempresas.frx":2B2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcatempresas.frx":340A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcatempresas.frx":3CE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcatempresas.frx":45BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcatempresas.frx":46D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcatempresas.frx":47E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcatempresas.frx":48F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcatempresas.frx":4A06
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame5 
      Height          =   120
      Left            =   150
      TabIndex        =   24
      Top             =   285
      Width           =   5655
   End
End
Attribute VB_Name = "frmcatempresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean






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
      Call pro_elimina_empresas
      rs.Open "select * from tb_empresas", cnn, adOpenDynamic, adLockOptimistic
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
      MsgBox "Imposible realizar la acci?n solicitada", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_guardar_Click()
Dim var_posible As Boolean
   var_posible = True
   If var_modifica_registro_empresa = False Then
      rs.Open "select * from tb_empresas where vcha_emp_empresa_id = '" + Me.txt_empresas(0) + "'", cnn, adOpenDynamic, adLockOptimistic
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
         Call pro_guardar_empresas
         rs.Open "select * from tb_empresas", cnn, adOpenDynamic, adLockOptimistic
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
      MsgBox "Clave de empresa ya existe", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
      If vector_valida_passwords(var_indice_menu) = "*" Then
         frmpasswords.Show
      Else
         Call gPrintListView(lv_empresas, "LISTADO DE empresas")
      End If

End Sub

Private Sub cmd_nuevo_Click()
      Call pro_limpiatextos(Me)
      txt_empresas(0).Enabled = True
      txt_empresas(0).SetFocus: var_modifica_registro_empresa = False
      cmd_guardar.Enabled = True
      cmd_deshacer.Enabled = True
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_empresa = False Then
      var_si = MsgBox("No se han guardado los cambios, ?Desea salir?", vbYesNo, "ATENCION")
      If var_si <> 6 Then
         GoTo SALIR:
      End If
   Else
      If var_hubo_cambios = True Then
         var_si = MsgBox("No se han guardado los cambios, ?Desea salir?", vbYesNo, "ATENCION")
         If var_si <> 6 Then
            GoTo SALIR:
         End If
      End If
   End If
   Unload Me
   Exit Sub
SALIR:
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
   var_modifica_registro_empresa = True
   lv_empresas.SmallIcons = ImageList1
   Call pro_llena_listview1
   pro_textos
   rs.Open "select * from tb_empresas", cnn, adOpenDynamic, adLockOptimistic
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
   var_modifica_registro_empresa = False
   Call activa_forma(var_activa_forma_catempresas)
End Sub

Private Sub lv_empresas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_empresas, ColumnHeader)
End Sub

Private Sub lv_empresas_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_empresas.selectedItem = Item
        pro_textos
        var_modifica_registro_empresa = True
        txt_empresas(0).Enabled = True
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
   If Button.Index < 3 Then
      lv_empresas.SetFocus
      Call pro_avanzar(Me, lv_empresas, Button)
      pro_textos
   Else
      Call pro_busca_registro(lv_empresas, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub

Sub pro_guardar_empresas()
Dim ok As Boolean
Set TB_EMPRESAS = New TB_EMPRESAS
    If txt_empresas(0) <> "" And txt_empresas(1) <> "" Then
        If var_hubo_cambios Then
            ok = TB_EMPRESAS.Anadir(txt_empresas(0), txt_empresas(1), txt_empresas(2), txt_empresas(3), txt_empresas(4), txt_empresas(5))
            If ok Then
                pro_actualiza_ListView
                txt_empresas(0).Enabled = False
                MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
                txt_registros = lv_empresas.ListItems.Count
                var_modifica_registro_empresa = True
            Else
                MsgBox "No se puede grabar registro: " + TB_EMPRESAS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
        End If
    End If
    
Set TB_EMPRESAS = Nothing: var_hubo_cambios = False

End Sub

Sub pro_elimina_empresas()
   Dim var_llave_usuarios As String
   Set TB_EMPRESAS = New TB_EMPRESAS
   'On Error GoTo SALIR:
   ok = True
   If txt_empresas(0) <> "" And txt_empresas(1) <> "" And var_modifica_registro_empresa = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_EMPRESAS.Eliminar(txt_empresas(0))
      Else
         GoTo SALIR:
      End If
      If ok Then
         MsgBox "Se Elimino Correctamente el Registro", vbInformation
         lv_empresas.ListItems.Remove (lv_empresas.selectedItem.Index)
         Call pro_limpiatextos(Me)
         txt_registros = lv_empresas.ListItems.Count
         lv_empresas.selectedItem.Selected = True
         pro_textos
      Else
         MsgBox "No se puede grabar registro: " + TB_EMPRESAS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
SALIR:
Set TB_EMPRESAS = Nothing

End Sub


Sub pro_llena_listview1()

Dim list_item As ListItem

    rs.Open "select * from TB_empresas", cnn, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
        Set list_item = lv_empresas.ListItems.Add(, , rs(0).Value)
        list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
        list_item.SubItems(2) = IIf(IsNull(rs(2).Value), "", rs(2).Value)
        list_item.SubItems(3) = IIf(IsNull(rs(3).Value), "", rs(3).Value)
        list_item.SubItems(4) = IIf(IsNull(rs(4).Value), "", rs(4).Value)
        list_item.SubItems(5) = IIf(IsNull(rs(5).Value), "", rs(5).Value)
    rs.MoveNext:
    Wend
    rs.Close

End Sub


Sub pro_textos()
On Error GoTo err0:
Dim var_n As Double
var_n = lv_empresas.ListItems.Count
   If var_n > 0 Then
      txt_empresas(0) = lv_empresas.selectedItem
      txt_empresas(1) = lv_empresas.selectedItem.SubItems(1)
      txt_empresas(2) = lv_empresas.selectedItem.SubItems(2)
      txt_empresas(3) = lv_empresas.selectedItem.SubItems(3)
      txt_empresas(4) = lv_empresas.selectedItem.SubItems(4)
      txt_empresas(5) = lv_empresas.selectedItem.SubItems(5)
   End If
   var_numero_renglones = lv_empresas.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_empresas.ColumnHeaders(2).Width = 3850
   Else
      lv_empresas.ColumnHeaders(2).Width = 4099.9
   End If
   var_modifica_registro_empresa = True
   var_hubo_cambios = False
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro_empresa = False Then
        Set list_item = lv_empresas.ListItems.Add(, , txt_empresas(0))
        list_item.SubItems(1) = txt_empresas(1)
        list_item.SubItems(2) = txt_empresas(2)
        list_item.SubItems(3) = txt_empresas(3)
        list_item.SubItems(4) = txt_empresas(4)
        list_item.SubItems(5) = txt_empresas(5)
        list_item.EnsureVisible
        list_item.Selected = True
    Else
        lv_empresas.ListItems.Item(lv_empresas.selectedItem.Index).Checked = False
        lv_empresas.ListItems.Item(lv_empresas.selectedItem.Index) = txt_empresas(0)
        lv_empresas.ListItems.Item(lv_empresas.selectedItem.Index).ListSubItems(1) = txt_empresas(1)
        lv_empresas.ListItems.Item(lv_empresas.selectedItem.Index).ListSubItems(2) = txt_empresas(2)
        lv_empresas.ListItems.Item(lv_empresas.selectedItem.Index).ListSubItems(3) = txt_empresas(3)
        lv_empresas.ListItems.Item(lv_empresas.selectedItem.Index).ListSubItems(4) = txt_empresas(4)
        lv_empresas.ListItems.Item(lv_empresas.selectedItem.Index).ListSubItems(5) = txt_empresas(5)
        lv_empresas.ListItems.Item(lv_empresas.selectedItem.Index).Selected = True
    End If
    lv_empresas.SetFocus
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_empresas, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub

Private Sub txt_empresas_Change(Index As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub

Private Sub txt_empresas_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Index < 4 Then
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
