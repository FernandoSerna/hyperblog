VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcolonias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo de colonias"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   Icon            =   "frmcolonias.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5955
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   135
      TabIndex        =   31
      Top             =   3690
      Width           =   5655
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   30
         TabIndex        =   32
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
         TabIndex        =   33
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame frm_filtro 
      Height          =   1590
      Left            =   120
      TabIndex        =   53
      Top             =   3855
      Width           =   5655
      Begin VB.TextBox txt_filtro_estado 
         Height          =   315
         Left            =   915
         TabIndex        =   56
         Top             =   1185
         Width           =   900
      End
      Begin VB.TextBox txt_filtro_nombre_estado 
         Height          =   315
         Left            =   1830
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   1185
         Width           =   3675
      End
      Begin VB.TextBox txt_filtro_nombre_pais 
         Height          =   315
         Left            =   1830
         TabIndex        =   55
         Top             =   840
         Width           =   3675
      End
      Begin VB.TextBox txt_filtro_pais 
         Height          =   315
         Left            =   915
         TabIndex        =   54
         Top             =   840
         Width           =   900
      End
      Begin VB.CommandButton cmd_cancelar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   420
         Picture         =   "frmcolonias.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Cancelar Alt + C"
         Top             =   390
         Width           =   330
      End
      Begin VB.CommandButton cmd_aceptar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   75
         Picture         =   "frmcolonias.frx":0A14
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Aceptar Alt + A"
         Top             =   390
         Width           =   330
      End
      Begin VB.Frame Frame6 
         Height          =   120
         Left            =   15
         TabIndex        =   58
         Top             =   645
         Width           =   5610
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Index           =   11
         Left            =   225
         TabIndex        =   63
         Top             =   1245
         Width           =   540
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         Caption         =   " Seleccione el estado"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   62
         Top             =   120
         Width           =   5580
      End
      Begin VB.Label lab_estados 
         AutoSize        =   -1  'True
         Caption         =   "Pais:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   60
         Top             =   900
         Width           =   345
      End
   End
   Begin VB.Frame frm_filtrar_colonias 
      Height          =   2310
      Left            =   150
      TabIndex        =   36
      Top             =   3795
      Width           =   5685
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   375
         Picture         =   "frmcolonias.frx":0B5E
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Guardar Alt + G"
         Top             =   390
         Width           =   330
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   45
         Picture         =   "frmcolonias.frx":0C60
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Nuevo Alt + N"
         Top             =   390
         Width           =   330
      End
      Begin VB.Frame Frame5 
         Height          =   30
         Left            =   15
         TabIndex        =   50
         Top             =   780
         Width           =   5640
      End
      Begin VB.TextBox Text8 
         Height          =   315
         Left            =   1020
         TabIndex        =   44
         Top             =   1560
         Width           =   990
      End
      Begin VB.TextBox Text7 
         Height          =   315
         Left            =   1020
         TabIndex        =   43
         Top             =   900
         Width           =   990
      End
      Begin VB.TextBox Text6 
         Height          =   315
         Left            =   1020
         TabIndex        =   42
         Top             =   1230
         Width           =   990
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   1020
         TabIndex        =   41
         Top             =   1890
         Width           =   990
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   2025
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   1230
         Width           =   3540
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   2025
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   900
         Width           =   3540
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   2025
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   1560
         Width           =   3540
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   2025
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   1890
         Width           =   3540
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   49
         Top             =   120
         Width           =   5610
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Pais:"
         Height          =   195
         Index           =   10
         Left            =   225
         TabIndex        =   48
         Top             =   960
         Width           =   345
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Index           =   9
         Left            =   225
         TabIndex        =   47
         Top             =   1290
         Width           =   540
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Municipio:"
         Height          =   195
         Index           =   8
         Left            =   225
         TabIndex        =   46
         Top             =   1620
         Width           =   720
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad:"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   45
         Top             =   1950
         Width           =   540
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmcolonias.frx":0D62
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
      Picture         =   "frmcolonias.frx":0E64
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
      Picture         =   "frmcolonias.frx":0F66
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
      Picture         =   "frmcolonias.frx":1038
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
      Picture         =   "frmcolonias.frx":113A
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
      Picture         =   "frmcolonias.frx":123C
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
      TabIndex        =   27
      Top             =   6045
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   165
      Top             =   5775
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
            Picture         =   "frmcolonias.frx":1876
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcolonias.frx":2150
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Colonias"
      Height          =   2655
      Left            =   150
      TabIndex        =   0
      Top             =   405
      Width           =   5655
      Begin VB.TextBox txt_nombre_ciudad 
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1230
         Width           =   3540
      End
      Begin VB.TextBox txt_nombre_municipio 
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   900
         Width           =   3540
      End
      Begin VB.TextBox txt_nombre_pais 
         Height          =   315
         Left            =   2055
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   3540
      End
      Begin VB.TextBox txt_nombre_estado 
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   570
         Width           =   3540
      End
      Begin VB.TextBox txt_ciudad 
         Height          =   315
         Left            =   1035
         TabIndex        =   13
         Top             =   1230
         Width           =   990
      End
      Begin VB.TextBox txt_nombre_colonia 
         Height          =   315
         Left            =   1035
         MaxLength       =   50
         TabIndex        =   17
         Top             =   2220
         Width           =   4530
      End
      Begin VB.TextBox txt_codigo_postal 
         Height          =   315
         Left            =   1035
         MaxLength       =   50
         TabIndex        =   16
         Top             =   1890
         Width           =   990
      End
      Begin VB.TextBox txt_colonia 
         Height          =   315
         Left            =   1035
         MaxLength       =   50
         TabIndex        =   15
         Top             =   1575
         Width           =   990
      End
      Begin VB.TextBox txt_estado 
         Height          =   315
         Left            =   1035
         TabIndex        =   9
         Top             =   570
         Width           =   990
      End
      Begin VB.TextBox txt_pais 
         Height          =   315
         Left            =   1035
         TabIndex        =   7
         Top             =   240
         Width           =   990
      End
      Begin VB.TextBox txt_municipio 
         Height          =   315
         Left            =   1035
         TabIndex        =   11
         Top             =   900
         Width           =   990
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad:"
         Height          =   195
         Index           =   3
         Left            =   255
         TabIndex        =   35
         Top             =   1290
         Width           =   540
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   29
         Top             =   2280
         Width           =   600
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "C.P."
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   26
         Top             =   1950
         Width           =   300
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   25
         Top             =   1620
         Width           =   450
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Municipio:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   20
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   630
         Width           =   540
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Pais:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   300
         Width           =   345
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   150
      TabIndex        =   21
      Top             =   3015
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   1980
         TabIndex        =   34
         Top             =   165
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3795
         TabIndex        =   30
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
         Caption         =   "Busqueda de colonia:"
         Height          =   195
         Left            =   195
         TabIndex        =   22
         Top             =   195
         Width           =   1545
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3690
      Left            =   150
      TabIndex        =   23
      Top             =   3510
      Width           =   5655
      Begin MSComctlLib.ListView lv_colonias 
         Height          =   3450
         Left            =   45
         TabIndex        =   28
         Top             =   180
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   6085
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
         NumItems        =   7
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
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "estado"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Municipio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ciudad"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "cp"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   2475
      Top             =   60
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
            Picture         =   "frmcolonias.frx":2A2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcolonias.frx":3304
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcolonias.frx":3BDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcolonias.frx":417A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcolonias.frx":4A56
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcolonias.frx":5330
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcolonias.frx":5C0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcolonias.frx":5D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcolonias.frx":5E2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcolonias.frx":5F40
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   24
      Top             =   285
      Width           =   5655
   End
End
Attribute VB_Name = "frmcolonias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim numero_items_colonias As Double
Dim var_tipo_lista As Integer




Private Sub cmd_aceptar_Click()
   txt_pais = txt_filtro_pais
   txt_nombre_pais = txt_filtro_pais
   txt_estado = txt_filtro_estado
   txt_nombre_estado = txt_filtro_nombre_estado
   Me.cmd_nuevo.Enabled = True
   Me.cmd_deshacer.Enabled = True
   Me.cmd_eliminar.Enabled = True
   Me.cmd_guardar.Enabled = True
   Me.cmd_imprimir.Enabled = True
   Me.cmd_nuevo.Enabled = True
   Me.txt_buscar.Enabled = True
   Me.txt_estado.Enabled = True
   Me.txt_nombre_estado.Enabled = True
   lv_colonias.Enabled = True
   lv_colonias.ListItems.Clear
   Call pro_encabezadosView(Me, lv_colonias, False)
   Call pro_llena_listview1
   pro_textos
   rs.Open "select * from tb_COLONIAS where vcha_pai_pais_id = '" + txt_pais + "' and vcha_est_estado_id = '" + txt_estado + "'", cnn, adOpenDynamic, adLockOptimistic
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

Private Sub cmd_cancelar_Click()
   frm_filtro.Visible = False
End Sub

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
      Call pro_elimina_colonias
      rs.Open "select * from tb_colonias", cnn, adOpenDynamic, adLockOptimistic
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
   If var_modifica_registro_colonia = False Then
      rs.Open "select * from tb_colonias where vcha_col_colonia_id = '" + txt_colonia + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_posible = False
      End If
      rs.Close
   End If
   var_posible = True
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
         Call pro_guardar_colonias
         If rs.State = 1 Then
            rs.Close
         End If
         rs.Open "select * from tb_colonias", cnn, adOpenDynamic, adLockOptimistic
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
      MsgBox "Clave de colonia incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call gPrintListView(lv_colonias, "LISTADO DE colonias")
        End If

End Sub

Private Sub cmd_nuevo_Click()
   txt_nombre_pais.Enabled = True
   txt_estado.Enabled = True
   txt_nombre_estado.Enabled = True
   txt_municipio.Enabled = True
   txt_nombre_municipio.Enabled = True
   txt_ciudad.Enabled = True
   txt_nombre_ciudad.Enabled = True
   'txt_colonia.Enabled = True
   txt_pais.Enabled = True
   txt_nombre_pais.Enabled = True
   txt_ciudad = ""
   txt_nombre_ciudad = ""
   txt_municipio = ""
   txt_nombre_municipio = ""
   txt_colonia = ""
   txt_nombre_colonia = ""
                          
   txt_municipio.SetFocus: var_modifica_registro_colonia = False
   var_modifica_registro_colonia = False
   cmd_guardar.Enabled = True
   cmd_deshacer.Enabled = True
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_colonia = False Then
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



Private Sub Form_Initialize()
   txt_pais.SetFocus
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
   Me.frm_filtrar_colonias.Visible = False
   frm_filtro.Visible = False
   var_cadena_seguridad = ""
   Top = 0
   Left = 2900
   var_modifica_registro_colonia = True
   Me.cmd_nuevo.Enabled = False
   Me.cmd_deshacer.Enabled = False
   Me.cmd_eliminar.Enabled = False
   Me.cmd_guardar.Enabled = False
   Me.cmd_imprimir.Enabled = False
   Me.cmd_nuevo.Enabled = False
   Me.txt_buscar.Enabled = False
   Me.txt_estado.Enabled = False
   Me.txt_nombre_estado.Enabled = False
   lv_colonias.Enabled = False
   txt_filtro_pais = ""
   txt_filtro_nombre_pais = ""
   txt_filtro_estado = ""
   txt_filtro_nombre_estado = ""
   frm_filtro.Visible = False
   frm_lista.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    var_modifica_registro_colonia = False
    Call activa_forma(var_activa_forma_colonias)
End Sub

Private Sub lv_colonias_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_colonias, ColumnHeader)
End Sub

Private Sub lv_colonias_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Set lv_colonias.selectedItem = Item
   pro_textos
   var_modifica_registro_colonia = True
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
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
      If var_tipo_lista = 3 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_municipio = lv_lista.selectedItem
            txt_nombre_municipio = lv_lista.selectedItem.SubItems(1)
         Else
            txt_municipio = ""
            txt_nombre_municipio = ""
         End If
         txt_municipio.SetFocus
      End If
      If var_tipo_lista = 4 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_ciudad = lv_lista.selectedItem
            txt_nombre_ciudad = lv_lista.selectedItem.SubItems(1)
         Else
            txt_ciudad = ""
            txt_nombre_ciudad = ""
         End If
         txt_ciudad.SetFocus
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
      lv_colonias.SetFocus
      Call pro_avanzar(Me, lv_colonias, Button)
      lv_colonias.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_colonias.ListItems(1).Selected = True
      lv_colonias.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 4 Then
      numero_items_colonias = lv_colonias.ListItems.Count
      lv_colonias.ListItems(numero_items_colonias).Selected = True
      lv_colonias.selectedItem.EnsureVisible
      pro_textos
   End If
err0:
End Sub


Sub pro_guardar_colonias()

Dim ok As Boolean

Set TB_COLONIAS = New TB_COLONIAS
Set TB_BITACORA_COLONIAS = New TB_BITACORA_COLONIAS
    
    If txt_nombre_colonia <> "" And txt_pais <> "" And txt_estado <> "" Then
        If var_hubo_cambios Then
            var_colonia_regreso = txt_colonia
            rs.Open "select * from tb_colonias where vcha_pai_pais_id = '" + txt_pais + "' and  vcha_est_estado_id = '" + txt_estado + "' and vcha_mun_municipio_id = '" + txt_municipio + "' and vcha_col_colonia_id = '" + txt_colonia + "'", cnn, adOpenDynamic, adLockOptimistic
            ok = TB_COLONIAS.Anadir(txt_pais, txt_estado, txt_municipio, txt_ciudad, txt_colonia, txt_nombre_colonia, txt_codigo_postal)
            txt_colonia = var_colonia_regreso
            If ok Then
                bitacora = True
                If var_modifica_registro_colonia = False Then
                   var_operacion_bitacora = "I"
                   bitacora = TB_BITACORA_COLONIAS.Anadir(txt_pais, txt_colonia, txt_municipio, txt_colonia, "VCHA_COL_NOMBRE", var_operacion_bitacora, "", txt_nombre_colonia, var_clave_usuario_global, fun_NombrePc, Date)
                Else
                   var_operacion_bitacora = "M"
                   If rs!VCHA_PAI_PAIS_ID <> txt_pais Then
                      bitacora = TB_BITACORA_COLONIAS.Anadir(txt_pais, txt_estado, txt_municipio, txt_colonia, "VCHA_PAI_PAIS_ID", var_operacion_bitacora, rs!VCHA_PAI_PAIS_ID, txt_pais, var_clave_usuario_global, fun_NombrePc, Date)
                   End If
                   If rs!VCHA_EST_ESTADO_ID <> txt_estado Then
                      bitacora = TB_BITACORA_COLONIAS.Anadir(txt_pais, txt_estado, txt_municipio, txt_colonia, "VCHA_EST_ESTADO_ID", var_operacion_bitacora, rs!VCHA_EST_ESTADO_ID, txt_estado, var_clave_usuario_global, fun_NombrePc, Date)
                   End If
                   If rs!VCHA_MUN_MUNICIPIO_ID <> txt_municipio Then
                      bitacora = TB_BITACORA_COLONIAS.Anadir(txt_pais, txt_estado, txt_municipio, txt_colonia, "VCHA_MUN_CIUDAD_ID", var_operacion_bitacora, rs!VCHA_MUN_MUNICIPIO_ID, txt_municipio, var_clave_usuario_global, fun_NombrePc, Date)
                   End If
                   If rs!VCHA_COL_COLONIA_ID <> txt_colonia Then
                      bitacora = TB_BITACORA_COLONIAS.Anadir(txt_pais, txt_estado, txt_municipio, txt_colonia, "VCHA_COL_COLONIA_ID", var_operacion_bitacora, rs!VCHA_COL_COLONIA_ID, txt_colonia, var_clave_usuario_global, fun_NombrePc, Date)
                   End If
                   If rs!vcha_col_cp <> txt_codigo_postal Then
                      bitacora = TB_BITACORA_COLONIAS.Anadir(txt_pais, txt_estado, txt_municipio, txt_colonia, "VCHA_COL_CP", var_operacion_bitacora, rs!vcha_col_cp, txt_codigo_postal, var_clave_usuario_global, fun_NombrePc, Date)
                   End If
                   If rs!vcha_col_nombre <> txt_nombre_colonia Then
                      bitacora = TB_BITACORA_COLONIAS.Anadir(txt_pais, txt_estado, txt_municipio, txt_colonia, "VCHA_COL_NOMBRE", var_operacion_bitacora, rs!vcha_col_nombre, txt_nombre_colonia, var_clave_usuario_global, fun_NombrePc, Date)
                   End If
                End If
                rs.Close
                pro_actualiza_ListView
                txt_pais.Enabled = False
                MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
                txt_registros = lv_colonias.ListItems.Count
                var_modifica_registro_colonia = True
            Else
                MsgBox "No se puede grabar registro: " + TB_COLONIAS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
        End If
    End If
    
Set TB_COLONIAS = Nothing: var_hubo_cambios = False

End Sub

Sub pro_elimina_colonias()
Dim var_llave_usuarios As String
Set TB_COLONIAS = New TB_COLONIAS
Set TB_BITACORA_COLONIAS = New TB_BITACORA_COLONIAS
   ok = True
   If txt_colonia <> "" And var_modifica_registro_colonia = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_COLONIAS.Eliminar(txt_colonia)
      Else
         GoTo salir:
      End If
      If ok Then
         var_operacion_bitacora = "I"
         bitacora = TB_BITACORA_COLONIAS.Anadir(txt_pais, txt_estado, txt_ciudad, txt_colonia, "VCHA_COL_NOMBRE", var_operacion_bitacora, txt_nombre_colonia, "", var_clave_usuario_global, fun_NombrePc, Date)
         numero_items_colonias = numero_items_colonias - 1
         MsgBox "Se Elimino Correctamente el Registro", vbInformation
         lv_colonias.ListItems.Remove (lv_colonias.selectedItem.Index)
         Call pro_limpiatextos(Me)
         txt_registros = lv_colonias.ListItems.Count
         If lv_colonias.ListItems.Count > 0 Then
            lv_colonias.selectedItem.Selected = True
            pro_textos
         Else
            txt_pais = ""
            txt_nombre_pais = ""
            txt_estado = ""
            txt_nombre_estado = ""
            txt_municipio = ""
            txt_nombre_municipio = ""
            txt_ciudad = ""
            txt_nombre_ciudad = ""
            txt_colonia = ""
            txt_nombre_colonia = ""
            txt_codigo_postal = ""
         End If
      Else
         MsgBox "No se puede grabar registro: " + TB_COLONIAS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
Set TB_COLONIAS = Nothing

End Sub


Sub pro_llena_listview1()

Dim list_item As ListItem
Dim var_hubo As Boolean
    numero_items_colonias = 0
    rs.Open "select * from TB_colonias WHERE VCHA_PAI_PAIS_ID = '" + txt_pais + "' AND VCHA_EST_ESTADO_ID = '" + txt_estado + "'", cnn, adOpenDynamic, adLockOptimistic
    var_hubo = False
    While Not rs.EOF
       var_hubo = True
        Set list_item = lv_colonias.ListItems.Add(, , rs!VCHA_COL_COLONIA_ID)
        list_item.SubItems(1) = IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
        list_item.SubItems(2) = IIf(IsNull(rs!VCHA_PAI_PAIS_ID), "", rs!VCHA_PAI_PAIS_ID)
        list_item.SubItems(3) = IIf(IsNull(rs!VCHA_EST_ESTADO_ID), "", rs!VCHA_EST_ESTADO_ID)
        list_item.SubItems(4) = IIf(IsNull(rs!VCHA_MUN_MUNICIPIO_ID), "", rs!VCHA_MUN_MUNICIPIO_ID)
        list_item.SubItems(5) = IIf(IsNull(rs!VCHA_CIU_CIUDAD_ID), "", rs!VCHA_CIU_CIUDAD_ID)
        list_item.SubItems(6) = IIf(IsNull(rs!vcha_col_cp), "", rs!vcha_col_cp)
    rs.MoveNext:
    numero_items_colonias = numero_items_colonias + 1
    Wend
    rs.Close
End Sub


Sub pro_textos()
'On Error GoTo err0:
   Dim var_n As Double
   var_n = lv_colonias.ListItems.Count
   If var_n > 0 Then
      txt_colonia = lv_colonias.selectedItem
      txt_nombre_colonia = lv_colonias.selectedItem.SubItems(1)
      txt_pais = lv_colonias.selectedItem.SubItems(2)
      txt_estado = lv_colonias.selectedItem.SubItems(3)
      txt_municipio = lv_colonias.selectedItem.SubItems(4)
      txt_ciudad = lv_colonias.selectedItem.SubItems(5)
      txt_codigo_postal = lv_colonias.selectedItem.SubItems(6)
      rs.Open "SELECT * FROM TB_PAISES WHERE VCHA_PAI_PAIS_ID = '" + txt_pais + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_pais = IIf(IsNull(rs!vcha_pai_nombre), "", rs!vcha_pai_nombre)
      Else
         txt_nombre_pais = ""
      End If
      rs.Close
      rs.Open "SELECT * FROM TB_ESTADOS WHERE VCHA_EST_ESTADO_ID = '" + txt_estado + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
      Else
         txt_nombre_estado = ""
      End If
      rs.Close
      rs.Open "SELECT * FROM TB_MUNICIPIOS WHERE VCHA_MUN_MUNICIPIO_ID = '" + txt_municipio + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_municipio = IIf(IsNull(rs!vcha_mun_nombre), "", rs!vcha_mun_nombre)
      Else
         txt_nombre_municipio = ""
      End If
      rs.Close
      rs.Open "SELECT * FROM TB_CIUDADES WHERE VCHA_CIU_CIUDAD_ID = '" + txt_ciudad + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
      Else
        txt_nombre_ciudad = ""
      End If
      rs.Close
   Else
      txt_municipio = ""
      txt_nombre_municipio = ""
      txt_ciudad = ""
      txt_nombre_ciudad = ""
      txt_colonia = ""
      txt_nombre_colonia = ""
      txt_codigo_postal = ""
   End If
   var_numero_renglones = lv_colonias.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_colonias.ColumnHeaders(2).Width = 3850
   Else
      lv_colonias.ColumnHeaders(2).Width = 4099.9
   End If
   txt_municipio.Enabled = False
   txt_nombre_municipio.Enabled = False
   'txt_ciudad.Enabled = False
   'txt_nombre_ciudad.Enabled = False
   txt_colonia.Enabled = False
   var_modifica_registro_colonia = True
   var_hubo_cambios = False
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro_colonia = False Then
        Set list_item = lv_colonias.ListItems.Add(, , txt_colonia)
        list_item.SubItems(1) = txt_nombre_colonia
        list_item.SubItems(2) = txt_pais
        list_item.SubItems(3) = txt_estado
        list_item.SubItems(4) = txt_municipio
        list_item.SubItems(5) = txt_ciudad
        list_item.SubItems(6) = txt_codigo_postal
        list_item.EnsureVisible
        list_item.Selected = True
        numero_items_colonias = numero_items_colonias + 1
    Else
        lv_colonias.ListItems.Item(lv_colonias.selectedItem.Index).Checked = False
        lv_colonias.ListItems.Item(lv_colonias.selectedItem.Index) = txt_colonia
        lv_colonias.ListItems.Item(lv_colonias.selectedItem.Index).ListSubItems(1) = txt_nombre_colonia
        lv_colonias.ListItems.Item(lv_colonias.selectedItem.Index).ListSubItems(2) = txt_pais
        lv_colonias.ListItems.Item(lv_colonias.selectedItem.Index).ListSubItems(3) = txt_estado
        lv_colonias.ListItems.Item(lv_colonias.selectedItem.Index).ListSubItems(4) = txt_municipio
        lv_colonias.ListItems.Item(lv_colonias.selectedItem.Index).ListSubItems(5) = txt_ciudad
        lv_colonias.ListItems.Item(lv_colonias.selectedItem.Index).ListSubItems(6) = txt_codigo_postal
        lv_colonias.ListItems.Item(lv_colonias.selectedItem.Index).Selected = True
    End If
    lv_colonias.SetFocus
End Sub



Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_colonias, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub

Private Sub txt_ciudad_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_ciudad_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_ciudades where vcha_pai_pais_id = '" + txt_pais + "' and vcha_est_estado_id = '" + txt_estado + "' order by vcha_ciu_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_CIU_CIUDAD_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CIUDADES DE " + txt_nombre_estado
      var_tipo_lista = 4
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
      var_activa_forma_ciudades = Me.Name
      frmciudades.Show
   End If
End Sub

Private Sub txt_ciudad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_ciudad_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_ciudad) <> "" Then
      rs.Open "SELECT * FROM TB_CIUDADES WHERE VCHA_PAI_PAIS_ID = '" + txt_pais + "' AND VCHA_EST_ESTADO_ID = '" + txt_estado + "' AND VCHA_CIU_CIUDAD_ID = '" + txt_ciudad + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
      Else
         MsgBox "Clave de ciudad incorrecta", vbOKOnly, "ATENCION"
         txt_nombre_ciudad = ""
         txt_ciudad = ""
      End If
      rs.Close
   Else
      txt_nombre_ciudad = ""
   End If
End Sub

Private Sub txt_codigo_postal_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_colonia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_estado_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
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
      lv_colonias.Enabled = False
      txt_filtro_pais = ""
      txt_filtro_nombre_pais = ""
      txt_filtro_estado = ""
      txt_filtro_nombre_estado = ""
      frm_filtro.Visible = True
      txt_filtro_pais.SetFocus
   End If
End Sub

Private Sub txt_estado_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
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
      Me.Enabled = False
      var_activa_forma_estados = Me.Name
      frmestados.Show
   End If
End Sub

Private Sub txt_filtro_estado_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_filtro_estado_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_filtro_estado) <> "" Then
      rs.Open "SELECT * FROM TB_ESTADOS WHERE VCHA_PAI_PAIS_ID = '" + txt_filtro_pais + "' AND VCHA_EST_ESTADO_ID = '" + txt_filtro_estado + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_filtro_nombre_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
      Else
         MsgBox "Clave de estado incorrecta", vbOKOnly, "ATENCION"
         txt_filtro_nombre_estado = ""
         txt_filtro_estado = ""
      End If
      rs.Close
   Else
      txt_filtro_nombre_estado = ""
   End If
End Sub

Private Sub txt_filtro_nombre_estado_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_estados where vcha_pai_pais_id = '" + txt_pais + "' order by vcha_est_nombre", cnn, adOpenDynamic, adLockOptimistic
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
      Me.Enabled = False
      var_activa_forma_estados = Me.Name
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
      Me.Enabled = False
      var_activa_forma_paises = Me.Name
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
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
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
      rs.Open "SELECT * FROM TB_PAISES WHERE VCHA_PAI_PAIS_ID = '" + txt_filtro_pais + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_filtro_nombre_pais = IIf(IsNull(rs!vcha_pai_nombre), "", rs!vcha_pai_nombre)
      Else
         MsgBox "Clave de pais inocorrecta", vbOKOnly, "ATENCION"
         txt_filtro_pais = ""
         txt_filtro_nombre_pais = ""
      End If
      rs.Close
   Else
      txt_filtro_nombre_pais = ""
   End If
End Sub

Private Sub txt_municipio_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_municipio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_municipios where vcha_pai_pais_id = '" + txt_pais + "' and vcha_est_estado_id = '" + txt_estado + "' order by vcha_mun_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_MUN_MUNICIPIO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_mun_nombre), "", rs!vcha_mun_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "MUNICIPIOS DE" + txt_nombre_estado
      var_tipo_lista = 3
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
      var_activa_forma_municipios = Me.Name
      frmmunicipios.Show
   End If
End Sub

Private Sub txt_municipio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_municipio_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_municipio) <> "" Then
      rs.Open "select * from tb_municipios where vcha_mun_municipio_id = '" + txt_municipio + "' and vcha_pai_pais_id = '" + txt_pais + "' and vcha_est_estado_id = '" + txt_estado + "' AND VCHA_MUN_MUNICIPIO_ID = '" + txt_municipio + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_municipio = IIf(IsNull(rs!vcha_mun_nombre), "", rs!vcha_mun_nombre)
      Else
         MsgBox "Clave de municipio incorrecta", vbOKOnly, "ATENCION"
         txt_nombre_municipio = ""
         txt_municipio = ""
      End If
      rs.Close
   Else
      txt_nombre_municipio = ""
   End If
End Sub

Private Sub txt_nombre_ciudad_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_ciudad_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_ciudades where vcha_pai_pais_id = '" + txt_pais + "' and vcha_est_estado_id = '" + txt_estado + "' order by vcha_ciu_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_CIU_CIUDAD_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CIUDADES DE " + txt_nombre_estado
      var_tipo_lista = 4
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
      var_activa_forma_ciudades = Me.Name
      frmciudades.Show
   End If
End Sub

Private Sub txt_nombre_ciudad_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_ciudad_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_colonia_KeyPress(KeyAscii As Integer)
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

Private Sub txt_nombre_estado_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
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
      lv_colonias.Enabled = False
      txt_filtro_pais = ""
      txt_filtro_nombre_pais = ""
      txt_filtro_estado = ""
      txt_filtro_nombre_estado = ""
      frm_filtro.Visible = True
      txt_filtro_pais.SetFocus
   End If
End Sub

Private Sub txt_nombre_estado_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_estado_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_municipio_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_municipio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_municipios where vcha_pai_pais_id = '" + txt_pais + "' and vcha_est_estado_id = '" + txt_estado + "' order by vcha_mun_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_MUN_MUNICIPIO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_mun_nombre), "", rs!vcha_mun_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "MUNICIPIOS DE" + txt_nombre_estado
      var_tipo_lista = 3
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
      var_activa_forma_municipios = Me.Name
      frmmunicipios.Show
   End If
End Sub

Private Sub txt_nombre_municipio_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_municipio_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_pais_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
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
      lv_colonias.Enabled = False
      txt_filtro_pais = ""
      txt_filtro_nombre_pais = ""
      txt_filtro_estado = ""
      txt_filtro_nombre_estado = ""
      frm_filtro.Visible = True
      txt_filtro_pais.SetFocus
   End If
End Sub

Private Sub txt_nombre_pais_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_pais_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_pais_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para filtrar información"
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
      lv_colonias.Enabled = False
      txt_filtro_pais = ""
      txt_filtro_nombre_pais = ""
      txt_filtro_estado = ""
      txt_filtro_nombre_estado = ""
      frm_filtro.Visible = True
      txt_filtro_pais.SetFocus
   End If
End Sub

Private Sub txt_pais_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

