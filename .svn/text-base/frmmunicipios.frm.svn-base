VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmunicipios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Municipios"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5955
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   165
      TabIndex        =   27
      Top             =   510
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   28
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
         TabIndex        =   29
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame frm_filtro 
      Height          =   1590
      Left            =   180
      TabIndex        =   30
      Top             =   585
      Width           =   5655
      Begin VB.TextBox txt_filtro_estado 
         Height          =   315
         Left            =   915
         TabIndex        =   34
         Top             =   1185
         Width           =   900
      End
      Begin VB.TextBox txt_filtro_nombre_estado 
         Height          =   315
         Left            =   1830
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   1185
         Width           =   3675
      End
      Begin VB.TextBox txt_filtro_nombre_pais 
         Height          =   315
         Left            =   1830
         TabIndex        =   33
         Top             =   840
         Width           =   3675
      End
      Begin VB.TextBox txt_filtro_pais 
         Height          =   315
         Left            =   915
         TabIndex        =   32
         Top             =   840
         Width           =   900
      End
      Begin VB.CommandButton cmd_cancelar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   390
         Picture         =   "frmmunicipios.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Cancelar Alt + C"
         Top             =   390
         Width           =   330
      End
      Begin VB.CommandButton cmd_aceptar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   60
         Picture         =   "frmmunicipios.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Aceptar Alt + A"
         Top             =   390
         Width           =   330
      End
      Begin VB.Frame Frame5 
         Height          =   120
         Left            =   15
         TabIndex        =   31
         Top             =   645
         Width           =   5610
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   40
         Top             =   1245
         Width           =   540
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Seleccione el estado"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   38
         Top             =   120
         Width           =   5580
      End
      Begin VB.Label lab_estados 
         AutoSize        =   -1  'True
         Caption         =   "Pais:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   36
         Top             =   900
         Width           =   345
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4200
      Left            =   150
      TabIndex        =   25
      Top             =   3000
      Width           =   5655
      Begin MSComctlLib.ListView lv_municipios 
         Height          =   3975
         Left            =   45
         TabIndex        =   26
         Top             =   150
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   7011
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
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "estado"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "telefono"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   150
      TabIndex        =   21
      Top             =   2460
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   1770
         TabIndex        =   22
         Top             =   165
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3690
         TabIndex        =   23
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
         Caption         =   "Busqueda de ciudad:"
         Height          =   195
         Left            =   195
         TabIndex        =   24
         Top             =   195
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Ciudades "
      Height          =   1995
      Left            =   150
      TabIndex        =   14
      Top             =   465
      Width           =   5655
      Begin VB.TextBox txt_nombre_estado 
         Height          =   315
         Left            =   2205
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   9
         Top             =   585
         Width           =   3315
      End
      Begin VB.TextBox txt_nombre_pais 
         Height          =   315
         Left            =   2205
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   7
         Top             =   255
         Width           =   3315
      End
      Begin VB.TextBox txt_pais 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   6
         Top             =   255
         Width           =   900
      End
      Begin VB.TextBox txt_estado 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   8
         Top             =   585
         Width           =   900
      End
      Begin VB.TextBox txt_municipio 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   10
         Top             =   915
         Width           =   900
      End
      Begin VB.TextBox txt_nombre_municipio 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1245
         Width           =   4215
      End
      Begin VB.TextBox txt_clave_telefono 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   12
         Top             =   1575
         Width           =   900
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Pais:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   19
         Top             =   285
         Width           =   345
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   18
         Top             =   615
         Width           =   540
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   17
         Top             =   945
         Width           =   450
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   16
         Top             =   1275
         Width           =   600
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave teléfono:"
         Height          =   195
         Index           =   5
         Left            =   195
         TabIndex        =   15
         Top             =   1605
         Width           =   1065
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1455
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   13
      Top             =   5715
      Width           =   255
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5460
      Picture         =   "frmmunicipios.frx":0294
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
      Picture         =   "frmmunicipios.frx":08CE
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
      Picture         =   "frmmunicipios.frx":09D0
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
      Picture         =   "frmmunicipios.frx":0AD2
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
      Picture         =   "frmmunicipios.frx":0BA4
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
      Picture         =   "frmmunicipios.frx":0CA6
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   255
      Top             =   5445
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
            Picture         =   "frmmunicipios.frx":0DA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmunicipios.frx":1682
            Key             =   ""
         EndProperty
      EndProperty
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
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmunicipios.frx":1F5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmunicipios.frx":2836
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmunicipios.frx":3110
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmunicipios.frx":36AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmunicipios.frx":3F88
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmunicipios.frx":4862
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmunicipios.frx":513C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmunicipios.frx":524E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmunicipios.frx":5360
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmunicipios.frx":5472
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmunicipios.frx":5584
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
Attribute VB_Name = "frmmunicipios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim numero_items_MUNICIPIOS As Integer
Dim bitacora As Boolean
Dim var_tipo_lista As Integer




Private Sub cmd_aceptar_Click()
   Me.cmd_nuevo.Enabled = False
   Me.cmd_deshacer.Enabled = True
   Me.cmd_eliminar.Enabled = True
   Me.cmd_guardar.Enabled = True
   Me.cmd_imprimir.Enabled = True
   Me.cmd_nuevo.Enabled = True
   Me.txt_buscar.Enabled = True
   Me.txt_estado.Enabled = True
   Me.txt_nombre_estado.Enabled = True
   lv_municipios.Enabled = True
   txt_pais = txt_filtro_pais
   txt_nombre_pais = txt_filtro_nombre_pais
   txt_estado = txt_filtro_estado
   txt_nombre_estado = txt_filtro_nombre_estado
   var_modifica_registro_municipio = True
   lv_municipios.SmallIcons = ImageList1
   lv_municipios.ListItems.Clear
   Call pro_encabezadosView(Me, lv_municipios, False)
   Call pro_llena_listview1
   pro_textos
   rs.Open "select * from tb_municipios where vcha_pai_pais_id = '" + txt_pais + "' and vcha_est_estado_id = '" + txt_estado + "'", cnn, adOpenDynamic, adLockOptimistic
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
   Me.txt_pais.SetFocus
End Sub

Private Sub cmd_cancelar_Click()
   Me.cmd_deshacer.Enabled = True
   Me.cmd_eliminar.Enabled = True
   Me.cmd_guardar.Enabled = True
   Me.cmd_imprimir.Enabled = True
   Me.cmd_nuevo.Enabled = True
   Me.txt_buscar.Enabled = True
   Me.txt_estado.Enabled = True
   Me.txt_nombre_estado.Enabled = True
   lv_municipios.Enabled = True
   rs.Open "select * from tb_municipios where vcha_pai_pais_id = '" + txt_pais + "' and vcha_est_estado_id = '" + txt_estado + "'", cnn, adOpenDynamic, adLockOptimistic
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
End Sub

Private Sub cmd_deshacer_Click()
      txt_pais.Enabled = False
      txt_estado.Enabled = False
      txt_municipio.Enabled = False
      txt_nombre_pais.Enabled = False
      txt_nombre_estado.Enabled = False
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
      txt_pais.Enabled = False
      txt_estado.Enabled = False
      txt_municipio.Enabled = False
      txt_nombre_pais.Enabled = False
      txt_nombre_estado.Enabled = False
      Call pro_elimina_MUNICIPIOS
      rs.Open "select * from tb_MUNICIPIOS", cnn, adOpenDynamic, adLockOptimistic
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
   If var_modifica_registro_municipio = False Then
      rs.Open "select * from tb_municipios where vcha_mun_municipio_id = '" + txt_municipio + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_posible = False
      End If
      rs.Close
   End If
   var_posible = True
   If var_posible = True Then
      If Trim(txt_pais) = "" Or Trim(txt_estado) = "" Then
         MsgBox "Información Incompleta", vbOKOnly, "ATENCION"
      Else
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
            txt_pais.Enabled = False
            txt_estado.Enabled = False
            txt_municipio.Enabled = False
            txt_nombre_pais.Enabled = False
            txt_nombre_estado.Enabled = False
            Call pro_guardar_MUNICIPIOS
            rs.Open "select * from tb_municipios", cnn, adOpenDynamic, adLockOptimistic
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
   Else
      MsgBox "Clave de municipio ya existe", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
      If vector_valida_passwords(var_indice_menu) = "*" Then
         frmpasswords.Show
      Else
         Call gPrintListView(lv_municipios, "LISTADO DE MUNICIPIOS")
      End If

End Sub

Private Sub cmd_nuevo_Click()
   txt_pais.Enabled = True
   txt_estado.Enabled = True
   txt_nombre_pais.Enabled = True
   txt_nombre_estado.Enabled = True
   txt_pais.Enabled = True
   Me.txt_municipio = ""
   Me.txt_nombre_municipio = ""
   Me.txt_clave_telefono = ""
   txt_nombre_municipio.SetFocus: var_modifica_registro_municipio = False
   cmd_guardar.Enabled = True
   cmd_deshacer.Enabled = True
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_municipio = False Then
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




Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_municipios_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_municipios, ColumnHeader)
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_municipios, txt_buscar, True)
      txt_buscar = ""
      pro_textos
   End If
End Sub

Private Sub txt_estado_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para accesar al catálogo"
End Sub

Private Sub txt_filtro_estado_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_estados where vcha_pai_pais_id = '" + txt_filtro_pais + "' order by vcha_est_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_EST_ESTADO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ESTADOS DE " + txt_nombre_pais
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
      var_activa_forma_estados = Me.Name
      Me.Enabled = False
      frmestados.Show
   End If
End Sub

Private Sub txt_filtro_estado_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_filtro_estado_LostFocus()
   If Trim(Me.txt_filtro_estado) <> "" Then
      If Trim(txt_filtro_pais) <> "" Then
         rs.Open "select * from tb_estados where vcha_pai_pais_id = '" + Me.txt_filtro_pais + "' and vcha_est_estado_id = '" + Me.txt_filtro_estado + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_filtro_nombre_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
         Else
            Me.txt_filtro_estado = ""
            Me.txt_filtro_nombre_estado = ""
            MsgBox "Clave de estado incorrecta", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         MsgBox "No se a seleccionado ni un país", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_filtro_nombre_estado_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_estados where vcha_pai_pais_id = '" + txt_filtro_pais + "' order by vcha_est_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_EST_ESTADO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ESTADOS DE " + txt_nombre_pais
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
      var_activa_forma_estados = Me.Name
      Me.Enabled = False
      frmestados.Show
   End If
End Sub

Private Sub txt_filtro_nombre_estado_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
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
      var_tipo_lista = 1
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
      var_activa_forma_paises = Me.Name
      Me.Enabled = False
      frmpaises.Show
   End If
End Sub

Private Sub txt_filtro_nombre_pais_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
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
      var_tipo_lista = 1
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
      var_activa_forma_paises = Me.Name
      Me.Enabled = False
      frmpaises.Show
   End If
End Sub

Private Sub txt_filtro_pais_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_filtro_pais_LostFocus()
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
   End If
End Sub

Private Sub txt_nombre_estado_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para accesar al catálogo"
End Sub

Private Sub txt_nombre_estado_KeyDown(KeyCode As Integer, Shift As Integer)
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
      lv_municipios.Enabled = False
      txt_filtro_pais = ""
      txt_filtro_nombre_pais = ""
      txt_filtro_estado = ""
      txt_filtro_nombre_estado = ""
      frm_filtro.Visible = True
      Me.txt_filtro_pais.SetFocus
   End If
End Sub

Private Sub txt_nombre_estado_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
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
   Me.cmd_nuevo.Enabled = False
   Me.cmd_deshacer.Enabled = False
   Me.cmd_eliminar.Enabled = False
   Me.cmd_guardar.Enabled = False
   Me.cmd_imprimir.Enabled = False
   Me.cmd_nuevo.Enabled = False
   Me.txt_buscar.Enabled = False
   Me.txt_estado.Enabled = False
   Me.txt_nombre_estado.Enabled = False
   frm_filtro.Visible = False
   var_cadena_seguridad = ""
   var_modifica_registro_municipio = True
   Top = 0
   Left = 2900
   txt_municipio.Enabled = False
   frm_lista.Visible = False
   cmd_guardar.Enabled = False
   cmd_deshacer.Enabled = False
   cmd_eliminar.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    var_swpassword = False
    var_modifica_registro_municipio = False
    Call activa_forma(var_activa_forma_municipios)
End Sub

Private Sub lv_municipios_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Set lv_municipios.selectedItem = Item
   pro_textos
   var_modifica_registro_municipio = True
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_tipo_lista = 1 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_filtro_pais = lv_lista.selectedItem
            txt_filtro_nombre_pais = lv_lista.selectedItem.SubItems(1)
         Else
            txt_filtro_pais = ""
            txt_filtro_nombre_pais = ""
         End If
         txt_filtro_pais.SetFocus
      End If
      If var_tipo_lista = 2 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_filtro_estado = lv_lista.selectedItem
            txt_filtro_nombre_estado = lv_lista.selectedItem.SubItems(1)
         Else
            txt_filtro_estado = ""
            txt_filtro_nombre_estado = ""
         End If
         txt_filtro_estado.SetFocus
      End If
      frm_lista.Visible = False
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_municipios.SetFocus
      Call pro_avanzar(Me, lv_municipios, Button)
      lv_municipios.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_municipios.ListItems(1).Selected = True
      lv_municipios.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 4 Then
      numero_items_MUNICIPIOS = lv_municipios.ListItems.Count
      lv_municipios.ListItems(numero_items_MUNICIPIOS).Selected = True
      lv_municipios.selectedItem.EnsureVisible
      pro_textos
   End If
err0:
End Sub


Sub pro_guardar_MUNICIPIOS()

Dim ok As Boolean

Set TB_MUNICIPIOS = New TB_MUNICIPIOS
Set TB_BITACORA_MUNICIPIOS = New TB_BITACORA_MUNICIPIOS
    
    
    If txt_pais <> "" And txt_estado <> "" Then
        If var_hubo_cambios Then
            rs.Open "select * from tb_MUNICIPIOS where vcha_mun_municipio_id = '" + txt_municipio + "'", cnn, adOpenDynamic, adLockOptimistic
            var_municipio_regreso = txt_municipio
            ok = TB_MUNICIPIOS.Anadir(txt_pais, txt_estado, txt_municipio, txt_nombre_municipio, txt_clave_telefono)
            txt_municipio = var_municipio_regreso
            If ok Then
                bitacora = True
                If var_modifica_registro_municipio = False Then
                   var_operacion_bitacora = "I"
                   bitacora = TB_BITACORA_MUNICIPIOS.Anadir(txt_pais, txt_estado, txt_municipio, "VCHA_CIU_NOMBRE", var_operacion_bitacora, "", txt_nombre_municipio, var_clave_usuario_global, fun_NombrePc, Date)
                Else
                   var_operacion_bitacora = "M"
                   If rs!VCHA_PAI_PAIS_ID <> txt_pais Then
                      bitacora = TB_BITACORA_MUNICIPIOS.Anadir(txt_pais, txt_estado, txt_municipio, "VCHA_PAI_PAIS_ID", var_operacion_bitacora, rs!VCHA_PAI_PAIS_ID, txt_pais, var_clave_usuario_global, fun_NombrePc, Date)
                   End If
                   If rs!VCHA_EST_ESTADO_ID <> txt_estado Then
                      bitacora = TB_BITACORA_MUNICIPIOS.Anadir(txt_pais, txt_estado, txt_municipio, "VCHA_EST_ESTADO_ID", var_operacion_bitacora, rs!VCHA_EST_ESTADO_ID, txt_estado, var_clave_usuario_global, fun_NombrePc, Date)
                   End If
                   If rs!VCHA_MUN_MUNICIPIO_ID <> txt_municipio Then
                      bitacora = TB_BITACORA_MUNICIPIOS.Anadir(txt_pais, txt_estado, txt_municipio, "VCHA_CIU_CIUDAD_ID", var_operacion_bitacora, rs!VCHA_MUN_MUNICIPIO_ID, txt_municipio, var_clave_usuario_global, fun_NombrePc, Date)
                   End If
                   If rs!vcha_mun_nombre <> txt_nombre_municipio Then
                      bitacora = TB_BITACORA_MUNICIPIOS.Anadir(txt_pais, txt_estado, txt_municipio, "VCHA_CIU_NOMBRE_ID", var_operacion_bitacora, rs!vcha_mun_nombre, txt_nombre_municipio, var_clave_usuario_global, fun_NombrePc, Date)
                   End If
                   If rs!VCHA_MUN_CLAVE_TELEFONO <> txt_clave_telefono Then
                      bitacora = TB_BITACORA_MUNICIPIOS.Anadir(txt_pais, txt_estado, txt_municipio, "VCHA_CIU_CLAVE_TELEFONO", var_operacion_bitacora, rs!VCHA_MUN_CLAVE_TELEFONO, txt_clave_telefono, var_clave_usuario_global, fun_NombrePc, Date)
                   End If
                End If
                rs.Close
             
                pro_actualiza_ListView
                txt_pais.Enabled = False
                MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
                txt_registros = lv_municipios.ListItems.Count
                var_modifica_registro_municipio = True
            Else
                MsgBox "No se puede grabar registro: " + TB_MUNICIPIOS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
        End If
    End If
    
Set TB_MUNICIPIOS = Nothing: var_hubo_cambios = False

End Sub

Sub pro_elimina_MUNICIPIOS()
Dim var_llave_usuarios As String

Set TB_MUNICIPIOS = New TB_MUNICIPIOS
Set TB_BITACORA_MUNICIPIOS = New TB_BITACORA_MUNICIPIOS
On Error GoTo salir:
   ok = True
   If txt_pais <> "" And txt_estado <> "" And var_modifica_registro_municipio = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_MUNICIPIOS.Eliminar(txt_municipio)
      Else
         GoTo salir:
      End If
      If ok Then
         var_operacion_bitacora = "E"
         bitacora = TB_BITACORA_MUNICIPIOS.Anadir(txt_pais, txt_estado, txt_municipio, "VCHA_CIU_NOMBRE", var_operacion_bitacora, "", txt_nombre_municipio, var_clave_usuario_global, fun_NombrePc, Date)
         numero_items_MUNICIPIOS = numero_items_MUNICIPIOS - 1
         MsgBox "Se Elimino Correctamente el Registro", vbInformation
         lv_municipios.ListItems.Remove (lv_municipios.selectedItem.Index)
         Call pro_limpiatextos(Me)
         txt_registros = lv_municipios.ListItems.Count
         lv_municipios.selectedItem.Selected = True
         pro_textos
      Else
         MsgBox "No se puede grabar registro: " + TB_MUNICIPIOS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
Set TB_MUNICIPIOS = Nothing
End Sub


Sub pro_llena_listview1()

Dim list_item As ListItem
   numero_items_MUNICIPIOS = 0
    rs.Open "select * from TB_MUNICIPIOS where vcha_pai_pais_id = '" + txt_pais + "' and vcha_est_estado_id = '" + txt_estado + "'", cnn, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
        Set list_item = lv_municipios.ListItems.Add(, , rs!VCHA_MUN_MUNICIPIO_ID)
        list_item.SubItems(1) = IIf(IsNull(rs!vcha_mun_nombre), "", rs!vcha_mun_nombre)
        list_item.SubItems(2) = IIf(IsNull(rs!VCHA_PAI_PAIS_ID), "", rs!VCHA_PAI_PAIS_ID)
        list_item.SubItems(3) = IIf(IsNull(rs!VCHA_EST_ESTADO_ID), "", rs!VCHA_EST_ESTADO_ID)
        list_item.SubItems(4) = IIf(IsNull(rs!VCHA_MUN_CLAVE_TELEFONO), "", rs!VCHA_MUN_CLAVE_TELEFONO)
    rs.MoveNext:
    numero_items_MUNICIPIOS = numero_items_MUNICIPIOS + 1
    Wend
    rs.Close
End Sub


Sub pro_textos()
   'On Error GoTo err0:
   Dim var_n As Integer
   var_n = lv_municipios.ListItems.Count
   If var_n > 0 Then
      txt_municipio = lv_municipios.selectedItem
      txt_nombre_municipio = lv_municipios.selectedItem.SubItems(1)
      txt_pais = lv_municipios.selectedItem.SubItems(2)
      txt_estado = lv_municipios.selectedItem.SubItems(3)
      txt_clave_telefono = lv_municipios.selectedItem.SubItems(4)
      rs.Open "select * from tb_paises where vcha_pai_pais_id = '" + txt_pais + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_pais = IIf(IsNull(rs!vcha_pai_nombre), "", rs!vcha_pai_nombre)
      Else
         txt_nombre_pais = ""
      End If
      rs.Close
      rs.Open "select * from tb_estados where vcha_pai_pais_id = '" & txt_pais & "' and vcha_est_estado_id = '" + txt_estado + "'", cnn, adOpenDynamic, adLockBatchOptimistic
      If Not rs.EOF Then
         txt_nombre_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
      Else
         txt_nombre_estado = ""
      End If
      rs.Close
   End If
   var_numero_renglones = lv_municipios.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_municipios.ColumnHeaders(2).Width = 3850
   Else
      lv_municipios.ColumnHeaders(2).Width = 4099.71
   End If
   var_hubo_cambios = False
   var_modifica_registro_municipio = True
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro_municipio = False Then
        Set list_item = lv_municipios.ListItems.Add(, , txt_municipio)
        list_item.SubItems(1) = txt_nombre_municipio
        list_item.SubItems(2) = txt_pais
        list_item.SubItems(3) = txt_estado
        list_item.SubItems(4) = txt_clave_telefono
        list_item.EnsureVisible
        list_item.Selected = True
        numero_items_MUNICIPIOS = numero_items_MUNICIPIOS + 1
    Else
        lv_municipios.ListItems.Item(lv_municipios.selectedItem.Index).Checked = False
        lv_municipios.ListItems.Item(lv_municipios.selectedItem.Index) = txt_municipio
        lv_municipios.ListItems.Item(lv_municipios.selectedItem.Index).ListSubItems(1) = txt_nombre_municipio
        lv_municipios.ListItems.Item(lv_municipios.selectedItem.Index).ListSubItems(2) = txt_pais
        lv_municipios.ListItems.Item(lv_municipios.selectedItem.Index).ListSubItems(3) = txt_estado
        lv_municipios.ListItems.Item(lv_municipios.selectedItem.Index).ListSubItems(4) = txt_clave_telefono
        lv_municipios.ListItems.Item(lv_municipios.selectedItem.Index).Selected = True
    End If
    lv_municipios.SetFocus
End Sub

Private Sub txt_clave_telefono_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_clave_telefono_KeyPress(KeyAscii As Integer)
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

Private Sub txt_estado_Change()
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub

Private Sub txt_estado_KeyDown(KeyCode As Integer, Shift As Integer)
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
      lv_municipios.Enabled = False
      txt_filtro_pais = ""
      txt_filtro_nombre_pais = ""
      txt_filtro_estado = ""
      txt_filtro_nombre_estado = ""
      frm_filtro.Visible = True
      Me.txt_filtro_pais.SetFocus
   End If
End Sub

Private Sub txt_estado_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_estado_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_estado) <> "" Then
      rs.Open "SELECT * FROM TB_ESTADOS WHERE VCHA_EST_ESTADO_ID = '" + txt_estado + "' AND vcha_pai_pais_id = '" + txt_pais + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
      Else
         txt_estado = ""
         MsgBox "Clave de Municipio incorrecta", vbOKOnly, "ATENCION"
         txt_nombre_estado = ""
      End If
      rs.Close
   Else
      txt_nombre_estado = ""
   End If
End Sub

Private Sub txt_municipio_Change()
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub

Private Sub txt_municipio_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_estado_Change()
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub


Private Sub txt_nombre_estado_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_estados where vcha_pai_pais_id = '" + txt_pais + "' order by vcha_pai_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_EST_ESTADO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ESTADOS"
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
      frmestados.Show
   End If
End Sub

Private Sub txt_nombre_municipio_Change()
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_municipio_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_pais_Change()
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub


Private Sub txt_nombre_pais_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para accesar al catálogo"
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
      lv_municipios.Enabled = False
      txt_filtro_pais = ""
      txt_filtro_nombre_pais = ""
      txt_filtro_estado = ""
      txt_filtro_nombre_estado = ""
      frm_filtro.Visible = True
      Me.txt_filtro_pais.SetFocus
   End If
End Sub

Private Sub txt_nombre_pais_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_pais_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_pais_Change()
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
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
      lv_municipios.Enabled = False
      txt_filtro_pais = ""
      txt_filtro_nombre_pais = ""
      txt_filtro_estado = ""
      txt_filtro_nombre_estado = ""
      frm_filtro.Visible = True
      Me.txt_filtro_pais.SetFocus
   End If
End Sub

Private Sub txt_pais_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
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
         MsgBox "Clave de Pais incorrecta", vbOKOnly, "ATENCION"
         txt_nombre_pais = ""
      End If
      rs.Close
   Else
      txt_nombre_pais = ""
   End If
End Sub
