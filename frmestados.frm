VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmestados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo de estados"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   Icon            =   "frmestados.frx":0000
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
      Top             =   465
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1905
         Left            =   45
         TabIndex        =   23
         Top             =   420
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   3360
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
         Top             =   135
         Width           =   5610
      End
   End
   Begin VB.Frame frm_filtro 
      Height          =   1200
      Left            =   180
      TabIndex        =   25
      Top             =   450
      Width           =   5640
      Begin VB.CommandButton cmd_aceptar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   45
         Picture         =   "frmestados.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Aceptar Alt + A"
         Top             =   390
         Width           =   330
      End
      Begin VB.CommandButton cmd_cancelar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   375
         Picture         =   "frmestados.frx":0A14
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Cancelar Alt + C"
         Top             =   390
         Width           =   330
      End
      Begin VB.Frame Frame5 
         Height          =   120
         Left            =   30
         TabIndex        =   32
         Top             =   645
         Width           =   5565
      End
      Begin VB.TextBox txt_filtro_pais 
         Height          =   315
         Left            =   915
         TabIndex        =   26
         Top             =   810
         Width           =   900
      End
      Begin VB.TextBox txt_filtro_nombre_pais 
         Height          =   315
         Left            =   1830
         TabIndex        =   28
         Top             =   810
         Width           =   3675
      End
      Begin VB.Label lab_estados 
         AutoSize        =   -1  'True
         Caption         =   "Pais:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   31
         Top             =   855
         Width           =   345
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   "Seleccione un país"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   27
         Top             =   120
         Width           =   5565
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5445
      Picture         =   "frmestados.frx":0B5E
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1470
      Picture         =   "frmestados.frx":1198
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      Picture         =   "frmestados.frx":129A
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   810
      Picture         =   "frmestados.frx":139C
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      Picture         =   "frmestados.frx":146E
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Guardar Alt + G"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmestados.frx":1570
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   2895
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   13
      Top             =   60
      Visible         =   0   'False
      Width           =   210
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   285
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
            Picture         =   "frmestados.frx":1672
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmestados.frx":1F4C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   11
      Top             =   285
      Width           =   5655
   End
   Begin VB.Frame Frame1 
      Caption         =   " Estados "
      Height          =   1335
      Left            =   150
      TabIndex        =   0
      Top             =   465
      Width           =   5655
      Begin VB.TextBox txt_nombre_pais 
         Height          =   315
         Left            =   1845
         MaxLength       =   50
         TabIndex        =   2
         Top             =   255
         Width           =   3675
      End
      Begin VB.TextBox txt_nombre_estado 
         Height          =   315
         Left            =   930
         MaxLength       =   50
         TabIndex        =   4
         Top             =   915
         Width           =   4590
      End
      Begin VB.TextBox txt_estado 
         Height          =   315
         Left            =   930
         MaxLength       =   50
         TabIndex        =   3
         Top             =   585
         Width           =   900
      End
      Begin VB.TextBox txt_pais 
         Height          =   315
         Left            =   930
         MaxLength       =   50
         TabIndex        =   1
         Top             =   255
         Width           =   900
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   2
         Left            =   255
         TabIndex        =   7
         Top             =   960
         Width           =   600
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   1
         Left            =   255
         TabIndex        =   6
         Top             =   630
         Width           =   450
      End
      Begin VB.Label lab_estados 
         AutoSize        =   -1  'True
         Caption         =   "Pais:"
         Height          =   195
         Index           =   0
         Left            =   255
         TabIndex        =   5
         Top             =   300
         Width           =   345
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   150
      TabIndex        =   8
      Top             =   1785
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   1830
         TabIndex        =   12
         Top             =   150
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3825
         TabIndex        =   15
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
         Caption         =   "Busqueda de estados:"
         Height          =   195
         Left            =   195
         TabIndex        =   9
         Top             =   195
         Width           =   1590
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4875
      Left            =   150
      TabIndex        =   10
      Top             =   2325
      Width           =   5655
      Begin MSComctlLib.ListView lv_estados 
         Height          =   4695
         Left            =   45
         TabIndex        =   14
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
            Text            =   "Nombre"
            Object.Width           =   6879
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "pais"
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
            Picture         =   "frmestados.frx":2826
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmestados.frx":3100
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmestados.frx":39DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmestados.frx":3F76
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmestados.frx":4852
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmestados.frx":512C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmestados.frx":5A06
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmestados.frx":5B18
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmestados.frx":5C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmestados.frx":5D3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmestados.frx":5E4E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmestados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim numero_items_estados As Integer


Private Sub cmd_aceptar_Click()
   Me.cmd_deshacer.Enabled = True
   Me.cmd_eliminar.Enabled = True
   Me.cmd_guardar.Enabled = True
   Me.cmd_imprimir.Enabled = True
   Me.cmd_nuevo.Enabled = True
   Me.txt_buscar.Enabled = True
   Me.txt_estado.Enabled = False
   Me.txt_nombre_estado.Enabled = True
   Me.lv_estados.Enabled = True
   txt_pais = txt_filtro_pais
   txt_nombre_pais = txt_filtro_nombre_pais
   var_modifica_registro_estado = True
   lv_estados.ListItems.Clear
   Call pro_encabezadosView(Me, lv_estados, False)
   Call pro_llena_listview1
   pro_textos
   
   rs.Open "select * from tb_estados where vcha_pai_pais_id = '" + txt_filtro_pais + "'", cnn, adOpenDynamic, adLockOptimistic
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
   frm_filtro.Visible = False
   txt_pais.SetFocus
End Sub

Private Sub cmd_cancelar_Click()
   frm_filtro.Visible = False
End Sub

Private Sub cmd_deshacer_Click()
   txt_estado.Enabled = False
   txt_pais.Enabled = False
   txt_nombre_pais.Enabled = False
   Call pro_textos
End Sub

Private Sub cmd_eliminar_Click()
   txt_estado.Enabled = False
   txt_pais.Enabled = False
   txt_nombre_pais.Enabled = False
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
      Call pro_elimina_estados
      rs.Open "select * from tb_estados", cnn, adOpenDynamic, adLockOptimistic
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
   If var_modifica_registro_estado = False Then
      rs.Open "select * from tb_estados where vcha_est_estado_id = '" + txt_estado + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_posible = False
      End If
      rs.Close
   End If
   var_posible = True
   If var_posible = True Then
      txt_estado.Enabled = False
      txt_pais.Enabled = False
      txt_nombre_pais.Enabled = False
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
         Call pro_guardar_estados
         If rs.State = 1 Then
            rs.Close
         End If
         rs.Open "select * from tb_estados", cnn, adOpenDynamic, adLockOptimistic
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
      MsgBox "Clave de estado ya existe", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
   X = 1 + 1
End Sub

Private Sub cmd_nuevo_Click()
   txt_estado = ""
   txt_nombre_estado = ""
   txt_pais.Enabled = True
   txt_nombre_pais.Enabled = True
   txt_nombre_estado.Enabled = True
   txt_nombre_estado.SetFocus: var_modifica_registro_estado = False
   'txt_estado.Enabled = True
   cmd_guardar.Enabled = True
   cmd_deshacer.Enabled = True
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_estado = False Then
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
   txt_estado.Enabled = False
   Me.cmd_nuevo.Enabled = False
   Me.cmd_deshacer.Enabled = False
   Me.cmd_eliminar.Enabled = False
   Me.cmd_guardar.Enabled = False
   Me.cmd_imprimir.Enabled = False
   Me.cmd_nuevo.Enabled = False
   Me.txt_buscar.Enabled = False
   Me.txt_estado.Enabled = False
   Me.txt_nombre_estado.Enabled = False
   Me.lv_estados.Enabled = False
   var_modifica_registro_estado = True
   frm_filtro.Visible = False
'   txt_filtro_pais.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    var_swpassword = False
    var_modifica_registro_estado = False
    Call activa_forma(var_activa_forma_estados)
End Sub

Private Sub lv_estados_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_estados, ColumnHeader)
End Sub

Private Sub lv_estados_ItemClick(ByVal Item As MSComctlLib.ListItem)
   txt_estado.Enabled = False
   Set lv_estados.selectedItem = Item
   pro_textos
   var_modifica_registro_estado = True
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         txt_filtro_pais = lv_lista.selectedItem
         txt_filtro_nombre_pais = lv_lista.selectedItem.SubItems(1)
      Else
         txt_filtro_pais = ""
         txt_filtro_nombre_pais = ""
      End If
      txt_filtro_pais.SetFocus
      frm_lista.Visible = False
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_estados, txt_buscar, True)
      txt_buscar = ""
      pro_textos
   End If
End Sub

Private Sub txt_filtro_nombre_pais_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_paises order by vcha_pai_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_PAI_PAIS_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_pai_nombre), "", rs!vcha_pai_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "PAISES"
      VAR_TIPO_LISTA = 1
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
      Me.Enabled = False
      var_activa_forma_paises = Me.Name
      frmpaises.Show
   End If
End Sub

Private Sub txt_filtro_nombre_pais_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      cmd_aceptar.SetFocus
   End If
End Sub

Private Sub txt_filtro_pais_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para accesar al catálogo"
End Sub

Private Sub txt_filtro_pais_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_paises order by vcha_pai_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_PAI_PAIS_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_pai_nombre), "", rs!vcha_pai_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "PAISES"
      VAR_TIPO_LISTA = 1
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
      Me.Enabled = False
      var_activa_forma_paises = Me.Name
      frmpaises.Show
   End If
End Sub

Private Sub txt_filtro_pais_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_filtro_pais_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_filtro_pais) <> "" Then
      rs.Open "select * from tb_paises where vcha_pai_pais_id = '" + txt_filtro_pais + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_filtro_nombre_pais = IIf(IsNull(rs!vcha_pai_nombre), "", rs!vcha_pai_nombre)
      Else
         txt_filtro_pais = ""
         txt_filtro_nombre_pais = ""
         MsgBox "Clave de pais incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      txt_filtro_nombre_pais = ""
   End If
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_estados.SetFocus
      Call pro_avanzar(Me, lv_estados, Button)
      lv_estados.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_estados.ListItems(1).Selected = True
      lv_estados.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 4 Then
      numero_items_estados = lv_estados.ListItems.Count
      lv_estados.ListItems(numero_items_estados).Selected = True
      lv_estados.selectedItem.EnsureVisible
      pro_textos
   End If
err0:
End Sub


Sub pro_guardar_estados()

Dim ok As Boolean

Set TB_ESTADOS = New TB_ESTADOS
Set TB_BITACORA_ESTADOS = New TB_BITACORA_ESTADOS
    
    ok = True
    If txt_pais <> "" And txt_nombre_estado <> "" Then
        If var_hubo_cambios Then
           rs.Open "select * from tb_estados where vcha_est_estado_id = '" + txt_estado + "'", cnn, adOpenDynamic, adLockOptimistic
            var_estado_regreso = txt_estado
            ok = TB_ESTADOS.Anadir(txt_pais, txt_estado, txt_nombre_estado)
            txt_estado = var_estado_regreso
            If ok Then
               bitacora = True
               If var_modifica_registro_estado = False Then
                  var_operacion_bitacora = "I"
                  bitacora = TB_BITACORA_ESTADOS.Anadir(txt_pais, txt_estado, "VCHA_EST_NOMBRE", var_operacion_bitacora, "", txt_nombre_estado, var_clave_usuario_global, fun_NombrePc, Date)
               Else
                  var_operacion_bitacora = "M"
                  If rs(0) <> txt_pais Then
                     bitacora = TB_BITACORA_ESTADOS.Anadir(txt_pais, txt_estado, "VCHA_PAI_PAIS_ID", var_operacion_bitacora, rs(0), txt_pais, var_clave_usuario_global, fun_NombrePc, Date)
                  End If
                  If rs(1) <> txt_estado Then
                     bitacora = TB_BITACORA_ESTADOS.Anadir(txt_pais, txt_estado, "VCHA_EST_ESTADO_ID", var_operacion_bitacora, rs(1), txt_estado, var_clave_usuario_global, fun_NombrePc, Date)
                  End If
                  If rs(2) <> txt_nombre_estado Then
                     bitacora = TB_BITACORA_ESTADOS.Anadir(txt_pais, txt_estado, "VCHA_EST_NOMBRE", var_operacion_bitacora, rs(2), txt_nombre_estado, var_clave_usuario_global, fun_NombrePc, Date)
                  End If
               End If
               rs.Close
               pro_actualiza_ListView
               txt_pais.Enabled = False
               MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
               txt_registros = lv_estados.ListItems.Count
               var_modifica_registro_estado = True
           Else
               MsgBox "No se puede grabar registro: " + TB_ESTADOS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
           End If
       End If
   End If
Set TB_ESTADOS = Nothing: var_hubo_cambios = False

End Sub

Sub pro_elimina_estados()
   Dim var_llave_usuarios As String
   Set TB_ESTADOS = New TB_ESTADOS
   Set TB_BITACORA_ESTADOS = New TB_BITACORA_ESTADOS
   ok = True
   On Error GoTo salir:
   If txt_pais <> "" And txt_estado <> "" And txt_nombre_estado _
      <> "" And txt_nombre_estado <> "" And var_modifica_registro_estado = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_ESTADOS.Eliminar(txt_estado)
      Else
         GoTo salir:
      End If
      If ok Then
         var_operacion_bitacora = "E"
         bitacora = TB_BITACORA_ESTADOS.Anadir(txt_pais, txt_estado, "VCHA_EST_NOMBRE", var_operacion_bitacora, txt_nombre_estado, "", var_clave_usuario_global, fun_NombrePc, Date)
         numero_items_estados = numero_items_estados - 1
         MsgBox "Se Elimino Correctamente el Registro", vbInformation
         lv_estados.ListItems.Remove (lv_estados.selectedItem.Index)
         Call pro_limpiatextos(Me)
         txt_registros = lv_estados.ListItems.Count
         lv_estados.selectedItem.Selected = True
         pro_textos
      Else
        MsgBox "No se puede eliminar registro: " + TB_ESTADOS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
   Set TB_ESTADOS = Nothing
End Sub


Sub pro_llena_listview1()
   Dim list_item As ListItem
   rs.Open "select * from TB_estados where vcha_pai_pais_id = '" + Me.txt_pais + "'", cnn, adOpenDynamic, adLockOptimistic
   numero_items_estados = 0
   While Not rs.EOF
      Set list_item = lv_estados.ListItems.Add(, , rs(1).Value)
      list_item.SubItems(1) = IIf(IsNull(rs(2).Value), "", rs(2).Value)
      list_item.SubItems(2) = IIf(IsNull(rs(0).Value), "", rs(0).Value)
      rs.MoveNext:
      numero_items_estados = numero_items_estados + 1
    Wend
    rs.Close
End Sub


Sub pro_textos()
On Error GoTo err0:
Dim var_n As Integer
   var_n = lv_estados.ListItems.Count
   If var_n > 0 Then
      txt_pais = lv_estados.selectedItem.SubItems(2)
      txt_estado = lv_estados.selectedItem
      txt_nombre_estado = lv_estados.selectedItem.SubItems(1)
      rs.Open "select * from tb_paises where vcha_pai_pais_id = '" + txt_pais + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_pais = IIf(IsNull(rs!vcha_pai_nombre), "", rs!vcha_pai_nombre)
      Else
         txt_nombre_pais = ""
      End If
      rs.Close
   End If
   var_numero_renglones = lv_estados.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_estados.ColumnHeaders(2).Width = 3850
   Else
      lv_estados.ColumnHeaders(2).Width = 4099.71
   End If
   var_modifica_registro_estado = True
   var_hubo_cambios = False
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro_estado = False Then
        Set list_item = lv_estados.ListItems.Add(, , txt_estado)
        list_item.SubItems(1) = txt_nombre_estado
        list_item.SubItems(2) = txt_pais
        list_item.EnsureVisible
        list_item.Selected = True
        numero_items_estados = numero_items_estados + 1
    Else
        lv_estados.ListItems.Item(lv_estados.selectedItem.Index).Checked = False
        lv_estados.ListItems.Item(lv_estados.selectedItem.Index) = txt_estado
        lv_estados.ListItems.Item(lv_estados.selectedItem.Index).ListSubItems(1) = txt_nombre_estado
        lv_estados.ListItems.Item(lv_estados.selectedItem.Index).ListSubItems(2) = txt_pais
        lv_estados.ListItems.Item(lv_estados.selectedItem.Index).Selected = True
    End If
'    lv_estados.SetFocus
End Sub

Private Sub txt_estado_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub


Private Sub txt_nombre_estado_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   If KeyAscii = 13 Then
      If Me.cmd_guardar.Enabled = True Then
         Me.cmd_guardar.SetFocus
      End If
   End If
End Sub


Private Sub txt_nombre_pais_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Me.cmd_nuevo.Enabled = False
      Me.cmd_deshacer.Enabled = False
      Me.cmd_eliminar.Enabled = False
      Me.cmd_guardar.Enabled = False
      Me.cmd_imprimir.Enabled = False
      Me.cmd_nuevo.Enabled = False
      Me.txt_buscar.Enabled = False
      Me.txt_estado.Enabled = False
      Me.txt_nombre_estado.Enabled = False
      Me.lv_estados.Enabled = False
      txt_filtro_pais = ""
      txt_filtro_nombre_pais = ""
      Me.frm_filtro.Visible = True
      txt_filtro_pais.SetFocus
   End If
End Sub

Private Sub txt_nombre_pais_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub


Private Sub txt_pais_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para accesar al catálogo"
End Sub

Private Sub txt_pais_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Me.cmd_nuevo.Enabled = False
      Me.cmd_deshacer.Enabled = False
      Me.cmd_eliminar.Enabled = False
      Me.cmd_guardar.Enabled = False
      Me.cmd_imprimir.Enabled = False
      Me.cmd_nuevo.Enabled = False
      Me.txt_buscar.Enabled = False
      Me.txt_estado.Enabled = False
      Me.txt_nombre_estado.Enabled = False
      Me.lv_estados.Enabled = False
      txt_filtro_pais = ""
      txt_filtro_nombre_pais = ""
      Me.frm_filtro.Visible = True
      txt_filtro_pais.SetFocus
   End If
End Sub

Private Sub txt_pais_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_pais_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_pais) <> "" Then
      rs.Open "select * from tb_paises where vcha_pai_pais_id = '" + txt_pais + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_pais = IIf(IsNull(rs!vcha_pai_nombre), "", rs!vcha_pai_nombre)
      Else
         txt_pais = ""
         txt_nombre_pais = ""
         MsgBox "Clave de pais incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      txt_nombre_pais = ""
   End If
End Sub
