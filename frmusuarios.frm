VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmusuarios 
   Appearance      =   0  'Flat
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios"
   ClientHeight    =   7350
   ClientLeft      =   135
   ClientTop       =   1005
   ClientWidth     =   11670
   Icon            =   "frmusuarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11670
   Begin VB.Frame frm_seleccionplantas 
      Height          =   1830
      Left            =   6525
      TabIndex        =   44
      Top             =   3600
      Width           =   4965
      Begin MSComctlLib.ListView lv_seleccionplantas 
         Height          =   1410
         Left            =   30
         TabIndex        =   47
         Top             =   360
         Width           =   4890
         _ExtentX        =   8625
         _ExtentY        =   2487
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
            Text            =   "CLAVE"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Planta"
            Object.Width           =   8467
         EndProperty
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000002&
         Caption         =   " Seleccione la Planta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   30
         TabIndex        =   46
         Top             =   120
         Width           =   4890
      End
   End
   Begin VB.Frame frm_seleccionempresas 
      Height          =   1830
      Left            =   6510
      TabIndex        =   62
      Top             =   780
      Width           =   4965
      Begin MSComctlLib.ListView lv_seleccionempresas 
         Height          =   1410
         Left            =   30
         TabIndex        =   63
         Top             =   360
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   2487
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
            Text            =   "CLAVE"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Planta"
            Object.Width           =   8467
         EndProperty
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000002&
         Caption         =   " Seleccione la Empresa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   30
         TabIndex        =   64
         Top             =   120
         Width           =   4890
      End
   End
   Begin VB.Frame Frame14 
      Height          =   2970
      Left            =   6465
      TabIndex        =   57
      Top             =   -60
      Width           =   5070
      Begin VB.CommandButton cmd_eliminar_empresa 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   390
         Picture         =   "frmusuarios.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Eliminar Alt + E"
         Top             =   375
         Width           =   330
      End
      Begin VB.CommandButton cmd_nueva_empresa 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   60
         Picture         =   "frmusuarios.frx":09CC
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Nuevo Alt + N"
         Top             =   375
         Width           =   330
      End
      Begin VB.Frame Frame15 
         Height          =   2160
         Left            =   45
         TabIndex        =   58
         Top             =   750
         Width           =   4950
         Begin MSComctlLib.ListView lv_empresas 
            Height          =   2175
            Left            =   0
            TabIndex        =   59
            Top             =   -15
            Width           =   4950
            _ExtentX        =   8731
            _ExtentY        =   3836
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Usuario"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "clave"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Empresa"
               Object.Width           =   8565
            EndProperty
         End
      End
      Begin VB.Frame Frame16 
         Height          =   120
         Left            =   30
         TabIndex        =   60
         Top             =   615
         Width           =   5010
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000002&
         Caption         =   " Empresas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   30
         TabIndex        =   61
         Top             =   120
         Width           =   4995
      End
   End
   Begin VB.Frame Frame6 
      Height          =   2910
      Left            =   6450
      TabIndex        =   27
      Top             =   2880
      Width           =   5070
      Begin VB.CommandButton cmd_eliminar_planta 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   375
         Picture         =   "frmusuarios.frx":0ACE
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "Eliminar Alt + E"
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton cmd_nuevo_planta 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   45
         Picture         =   "frmusuarios.frx":0BD0
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Nuevo Alt + N"
         Top             =   360
         Width           =   330
      End
      Begin VB.Frame Frame7 
         Height          =   120
         Left            =   30
         TabIndex        =   31
         Top             =   600
         Width           =   5010
      End
      Begin VB.Frame Frame9 
         Height          =   2190
         Left            =   75
         TabIndex        =   28
         Top             =   675
         Width           =   4890
         Begin MSComctlLib.ListView lv_plantas 
            Height          =   2145
            Left            =   0
            TabIndex        =   29
            Top             =   45
            Width           =   4905
            _ExtentX        =   8652
            _ExtentY        =   3784
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Usuario"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "clave"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Planta"
               Object.Width           =   8565
            EndProperty
         End
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000002&
         Caption         =   " Plantas "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   30
         TabIndex        =   30
         Top             =   120
         Width           =   4995
      End
   End
   Begin VB.CommandButton cmd_permisos 
      Caption         =   "Permisos a Movimientos"
      Height          =   345
      Left            =   6465
      TabIndex        =   56
      Top             =   6870
      Width           =   5070
   End
   Begin VB.Frame frm_seleccionpuestos 
      Height          =   855
      Left            =   6435
      TabIndex        =   51
      Top             =   5985
      Width           =   4995
      Begin VB.ComboBox cmb_puestos 
         Height          =   315
         Left            =   60
         Sorted          =   -1  'True
         TabIndex        =   53
         Top             =   405
         Width           =   4890
      End
      Begin VB.TextBox txt_clave_puesto 
         Height          =   285
         Left            =   90
         TabIndex        =   54
         Top             =   420
         Width           =   1005
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000002&
         Caption         =   " Seleccion del puesto "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   30
         TabIndex        =   52
         Top             =   120
         Width           =   4920
      End
   End
   Begin VB.Frame frm_seleccionmodulos 
      Height          =   1575
      Left            =   1440
      TabIndex        =   45
      Top             =   4320
      Width           =   4995
      Begin MSComctlLib.ListView lv_seleccionmodulos 
         Height          =   1200
         Left            =   15
         TabIndex        =   49
         Top             =   330
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   2117
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
            Text            =   "clave"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Modulos"
            Object.Width           =   8555
         EndProperty
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000002&
         Caption         =   " Seleccione el Modulo del Sistema"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   30
         TabIndex        =   48
         Top             =   120
         Width           =   4920
      End
   End
   Begin VB.Frame Frame10 
      Height          =   1965
      Left            =   9780
      TabIndex        =   39
      Top             =   3810
      Visible         =   0   'False
      Width           =   30
      Begin VB.CommandButton cmd_eliminar_modulo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   360
         Picture         =   "frmusuarios.frx":0CD2
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Eliminar Alt + E"
         Top             =   375
         Width           =   330
      End
      Begin VB.CommandButton cmd_nuevo_modulo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   30
         Picture         =   "frmusuarios.frx":0DD4
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   "Nuevo Alt + N"
         Top             =   375
         Width           =   330
      End
      Begin VB.Frame Frame11 
         Height          =   1185
         Left            =   45
         TabIndex        =   40
         Top             =   750
         Width           =   4965
         Begin MSComctlLib.ListView lv_modulos 
            Height          =   1200
            Left            =   0
            TabIndex        =   41
            Top             =   -15
            Width           =   4965
            _ExtentX        =   8758
            _ExtentY        =   2117
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
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "USUARIO"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "PLANTA"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "CLAVE"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Modulo"
               Object.Width           =   8565
            EndProperty
         End
      End
      Begin VB.Frame Frame12 
         Height          =   120
         Left            =   30
         TabIndex        =   42
         Top             =   615
         Width           =   4995
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000002&
         Caption         =   " Modulos del Sistema"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   30
         TabIndex        =   43
         Top             =   120
         Width           =   4995
      End
   End
   Begin VB.Frame frm_password 
      Height          =   1515
      Left            =   3180
      TabIndex        =   34
      Top             =   690
      Width           =   2700
      Begin VB.TextBox txt_confirmar_password 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   945
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   14
         Top             =   1125
         Width           =   1620
      End
      Begin VB.TextBox txt_password 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   945
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   780
         Width           =   1620
      End
      Begin VB.TextBox txt_clave_usuario 
         Height          =   315
         Left            =   930
         MaxLength       =   10
         TabIndex        =   12
         Top             =   435
         Width           =   1620
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Confirmar:"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   1170
         Width           =   705
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   825
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   480
         Width           =   585
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000002&
         Caption         =   " Usuarios "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   30
         TabIndex        =   35
         Top             =   120
         Width           =   2625
      End
   End
   Begin VB.Frame Frame5 
      Height          =   7275
      Left            =   135
      TabIndex        =   0
      Top             =   -60
      Width           =   6255
      Begin VB.CommandButton cmd_salir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5835
         Picture         =   "frmusuarios.frx":0ED6
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salir"
         Top             =   375
         Width           =   330
      End
      Begin VB.CommandButton cmd_imprimir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1365
         Picture         =   "frmusuarios.frx":1510
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Imprimir Alt + I"
         Top             =   375
         Width           =   330
      End
      Begin VB.CommandButton cmd_eliminar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1035
         Picture         =   "frmusuarios.frx":1612
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Eliminar Alt + E"
         Top             =   375
         Width           =   330
      End
      Begin VB.CommandButton cmd_deshacer 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   705
         Picture         =   "frmusuarios.frx":1714
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Deshacer Alt + D"
         Top             =   375
         Width           =   330
      End
      Begin VB.CommandButton cmd_guardar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   375
         Picture         =   "frmusuarios.frx":17E6
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Guardar Alt + G"
         Top             =   375
         Width           =   330
      End
      Begin VB.CommandButton cmd_nuevo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   45
         Picture         =   "frmusuarios.frx":18E8
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Nuevo Alt + N"
         Top             =   375
         Width           =   330
      End
      Begin VB.Frame Frame1 
         Caption         =   " Usuarios "
         Height          =   1995
         Index           =   0
         Left            =   30
         TabIndex        =   17
         Top             =   750
         Width           =   6165
         Begin VB.TextBox txt_nomina 
            Height          =   315
            Left            =   1230
            MaxLength       =   50
            TabIndex        =   75
            Top             =   1920
            Visible         =   0   'False
            Width           =   1770
         End
         Begin VB.CheckBox chk_permiso_cambio_transporte 
            Caption         =   "Permiso para cambiar transporte en volumen"
            Height          =   210
            Left            =   1230
            TabIndex        =   74
            Top             =   1680
            Width           =   3960
         End
         Begin VB.TextBox txt_apellidos_usuario 
            Height          =   315
            Left            =   1230
            MaxLength       =   50
            TabIndex        =   9
            Top             =   990
            Width           =   4440
         End
         Begin VB.CheckBox chk_permiso 
            Caption         =   "Requiere permisos para movimientos"
            Height          =   210
            Left            =   1230
            TabIndex        =   10
            Top             =   1350
            Width           =   3960
         End
         Begin VB.TextBox txt_nombre_usuario 
            Height          =   315
            Left            =   1230
            MaxLength       =   50
            TabIndex        =   8
            Top             =   645
            Width           =   4440
         End
         Begin VB.TextBox txt_usuario 
            Height          =   315
            Left            =   1230
            MaxLength       =   50
            TabIndex        =   7
            Top             =   300
            Width           =   1170
         End
         Begin MSComctlLib.Toolbar Toolbar3 
            Height          =   330
            Left            =   5700
            TabIndex        =   11
            Top             =   1050
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Asignar Clave de Acceso"
                  ImageIndex      =   11
               EndProperty
            EndProperty
         End
         Begin VB.Label lab_paises 
            AutoSize        =   -1  'True
            Caption         =   "N. Nomina:"
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   76
            Top             =   1980
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Label lab_paises 
            AutoSize        =   -1  'True
            Caption         =   "Apellidos:"
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   73
            Top             =   1050
            Width           =   675
         End
         Begin VB.Label lab_paises 
            AutoSize        =   -1  'True
            Caption         =   "Nombre:"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   19
            Top             =   705
            Width           =   600
         End
         Begin VB.Label lab_paises 
            AutoSize        =   -1  'True
            Caption         =   "Clave:"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   18
            Top             =   360
            Width           =   450
         End
      End
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   2205
         TabIndex        =   16
         Top             =   2895
         Width           =   1350
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7320
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   15
         Top             =   5250
         Width           =   255
      End
      Begin VB.Frame Frame2 
         Height          =   540
         Left            =   30
         TabIndex        =   20
         Top             =   2730
         Width           =   6165
         Begin MSComctlLib.Toolbar tool_atras_siguiente 
            Height          =   330
            Left            =   4305
            TabIndex        =   26
            Top             =   180
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
            Caption         =   "Busqueda de Usuario:"
            Height          =   195
            Left            =   495
            TabIndex        =   21
            Top             =   195
            Width           =   1575
         End
      End
      Begin MSComctlLib.ImageList icono_encabezado 
         Left            =   3810
         Top             =   105
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
               Picture         =   "frmusuarios.frx":19EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmusuarios.frx":22C4
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   3195
         Top             =   120
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
               Picture         =   "frmusuarios.frx":2B9E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmusuarios.frx":3478
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmusuarios.frx":3D52
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmusuarios.frx":42EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmusuarios.frx":4BCA
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmusuarios.frx":54A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmusuarios.frx":5D7E
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmusuarios.frx":5E90
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmusuarios.frx":5FA2
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmusuarios.frx":60B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmusuarios.frx":61C6
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame4 
         Height          =   120
         Left            =   45
         TabIndex        =   24
         Top             =   615
         Width           =   6135
      End
      Begin VB.Frame Frame3 
         Height          =   4005
         Left            =   30
         TabIndex        =   22
         Top             =   3225
         Width           =   6165
         Begin MSComctlLib.ListView lv_usuarios 
            Height          =   3765
            Left            =   30
            TabIndex        =   23
            Top             =   165
            Width           =   6105
            _ExtentX        =   10769
            _ExtentY        =   6641
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
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Clave"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nombre"
               Object.Width           =   3193
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Apellidos"
               Object.Width           =   5115
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "USUARIO"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "PASSWORD"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "permiso"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Permiso cambiar transporte"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Nomina"
               Object.Width           =   0
            EndProperty
         End
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000002&
         Caption         =   " Usuarios "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   30
         TabIndex        =   25
         Top             =   120
         Width           =   6195
      End
   End
   Begin VB.Frame Frame8 
      Height          =   1065
      Left            =   6465
      TabIndex        =   32
      Top             =   5760
      Width           =   5070
      Begin VB.CommandButton cmd_eliminar_puesto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   360
         Picture         =   "frmusuarios.frx":62D8
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Eliminar Alt + E"
         Top             =   330
         Width           =   330
      End
      Begin VB.CommandButton cmd_nuevo_puesto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   30
         Picture         =   "frmusuarios.frx":63DA
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Nuevo Alt + N"
         Top             =   330
         Width           =   330
      End
      Begin VB.TextBox txt_puesto 
         Height          =   300
         Left            =   45
         TabIndex        =   55
         Top             =   705
         Width           =   4965
      End
      Begin VB.Frame Frame13 
         Height          =   120
         Left            =   15
         TabIndex        =   50
         Top             =   570
         Width           =   5025
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000002&
         Caption         =   " Puesto "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   30
         TabIndex        =   33
         Top             =   120
         Width           =   4995
      End
   End
End
Attribute VB_Name = "frmusuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim bitacora As Boolean
Dim numero_items_tallas As Integer
Dim var_unidad_organizacional As String
Dim var_bloque As String
Dim var_puesto As String



Private Sub chk_permiso_cambio_transporte_Click()
   var_hubo_cambios = True
End Sub

Private Sub chk_permiso_Click()
   If chk_permiso.Value = 1 Then
      Me.cmd_permisos.Enabled = True
   Else
      Me.cmd_permisos.Enabled = False
   End If
End Sub

Private Sub cmb_puestos_Click()
   rs.Open "select * from tb_puestos where VCHA_PUE_DESCRIPCION = '" + Me.cmb_puestos + "'", cnn, adOpenDynamic, adLockOptimistic
   var_puesto = rs(0).Value
   rs.Close
   'var_puesto = Obtener_llave(cnn, rs, "TB_puestos", "VCHA_PUE_DESCRIPCION", cmb_puestos, 0, "T")
   
   txt_clave_puesto = var_puesto
End Sub

Private Sub cmb_puestos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      frm_seleccionpuestos.Visible = False
   End If
   If KeyAscii = 13 Then
      rsaux.Open "select * from VW_RELACIONES_BLOQUES_PUESTOS where vcha_usu_usuario_id = '" + txt_usuario + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_blo_bloque_id = '" + var_bloque + "'", cnn, adOpenDynamic, adLockOptimistic
      If rsaux.EOF Then
         rs.Open "select * from VW_RELACIONES_BLOQUES_PUESTOS where vcha_usu_usuario_id = '" + txt_usuario + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_blo_bloque_id = '" + var_bloque + "' and vcha_pue_puesto_id = '" + txt_clave_puesto + "'", cnn, adOpenDynamic, adLockOptimistic
         If rs.EOF Then
            Set TB_REL_BLOQUES_PUESTOS = New TB_REL_BLOQUES_PUESTOS
            var_anadir = TB_REL_BLOQUES_PUESTOS.Anadir(txt_usuario, var_unidad_organizacional, var_bloque, var_puesto)
            frm_seleccionpuestos.Visible = False
            txt_puesto = cmb_puestos
         Else
            MsgBox "Ya existe una relacion del Bloque con este puesto", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         MsgBox "Ya existe un puesto asignado al bloque anterior", vbOKOnly, "ATENCION"
      End If
      rsaux.Close
      frm_seleccionpuestos.Visible = False
      End If
End Sub

Private Sub cmb_puestos_LostFocus()
   frm_seleccionpuestos.Visible = False
End Sub


Private Sub cmd_deshacer_Click()
   Call pro_textos
End Sub

Private Sub cmd_eliminar_Click()
   var_resultado = InStr(1, var_menus, Me.Caption)
   var_inicio = var_resultado + Len(Me.Caption) + 3
   If Mid(var_menus, var_inicio, 1) = "1" Then
      Set var_forma = frmusuarios
      var_swpassword = True
      sw_primera_validacion = False
      frmpasswords.Show
   Else
      If Mid(var_menus, var_inicio, 2) = "01" Then
         Set var_forma = frmusuarios
         var_swpassword = True
         sw_primera_validacion = False
         frmpasswords2.txt_supervisor = var_supervisor
         frmpasswords2.Show
      Else
         Call pro_elimina_usuarios
         rs.Open "select * from tb_usuarios", cnn, adOpenDynamic, adLockOptimistic
         If rs.BOF Then
         End If
         rs.Close
      End If
   End If
End Sub

Private Sub cmd_eliminar_empresa_Click()
   Set TB_REL_USUARIOS_EMPRESAS_E = New TB_REL_USUARIOS_EMPRESAS_E
   Set TB_REL_BLOQUES_PUESTOS_E = New TB_REL_BLOQUES_PUESTOS_E
   Set TB_REL_BLOQUES_UNIDADES_E = New TB_REL_BLOQUES_UNIDADES_E
   Set TB_REL_USUARIOS_UNIDADES_E = New TB_REL_USUARIOS_UNIDADES_E
   si = MsgBox("Al eliminar la empresa se eliminara toda la informaci?n relacionada a ella ?Deseas continuar?", vbOKCancel, "ATENCION")
   If si = 1 Then
      Dim i As Integer
      i = lv_plantas.ListItems.Count
      Dim j As Integer
      For j = 1 To i
          lv_plantas.ListItems.Item(j).Selected = True
          var_unidad_organizacional = lv_plantas.selectedItem.SubItems(1)
          var_anadir = TB_REL_BLOQUES_PUESTOS_E.Anadir(txt_usuario, var_unidad_organizacional, var_bloque)
          var_anadir = TB_REL_BLOQUES_UNIDADES_E.Anadir(txt_usuario, var_unidad_organizacional)
          var_anadir = TB_REL_USUARIOS_UNIDADES_E.Anadir(txt_usuario, var_unidad_organizacional)
      Next j
      var_anadir = TB_REL_USUARIOS_EMPRESAS_E.Anadir(txt_usuario, lv_empresas.selectedItem.SubItems(1))
      lv_plantas.ListItems.Clear
      lv_modulos.ListItems.Clear
      txt_puesto = ""
      pro_textos
   End If
End Sub

Private Sub cmd_eliminar_modulo_Click()
   Set TB_REL_BLOQUES_PUESTOS_E = New TB_REL_BLOQUES_PUESTOS_E
   Set TB_REL_BLOQUES_UNIDADES_E = New TB_REL_BLOQUES_UNIDADES_E
   si = MsgBox("Al eliminar este bloque se eliminara su puesto relacionado a el, ?Deseas continuar?", vbOKCancel, "ATENCION")
   If si = 1 Then
      var_anadir = TB_REL_BLOQUES_PUESTOS_E.Anadir(txt_usuario, var_unidad_organizacional, var_bloque)
      var_anadir = TB_REL_BLOQUES_UNIDADES_E.Anadir(txt_usuario, var_unidad_organizacional)
   End If
   pro_textos
End Sub

Private Sub cmd_eliminar_planta_Click()
   Set TB_REL_BLOQUES_PUESTOS_E = New TB_REL_BLOQUES_PUESTOS_E
   Set TB_REL_BLOQUES_UNIDADES_E = New TB_REL_BLOQUES_UNIDADES_E
   Set TB_REL_USUARIOS_UNIDADES_E = New TB_REL_USUARIOS_UNIDADES_E
   si = MsgBox("Al eliminar esta Planta se eliminar tambi?n sus bloques y los puestos relacionados a ella, ?Deseas continuar?", vbOKCancel, "ATENCION")
   If si = 1 Then
      var_anadir = TB_REL_BLOQUES_PUESTOS_E.Anadir(txt_usuario, var_unidad_organizacional, var_bloque)
      var_anadir = TB_REL_BLOQUES_UNIDADES_E.Anadir(txt_usuario, var_unidad_organizacional)
      var_anadir = TB_REL_USUARIOS_UNIDADES_E.Anadir(txt_usuario, var_unidad_organizacional)
   End If
   pro_textos
End Sub

Private Sub cmd_eliminar_puesto_Click()
   Set TB_REL_BLOQUES_PUESTOS_E = New TB_REL_BLOQUES_PUESTOS_E
   si = MsgBox("?Esta seguro de eliminar el puesto?", vbOKCancel, "ATENCION")
   If si = 1 Then
      var_anadir = TB_REL_BLOQUES_PUESTOS_E.Anadir(txt_usuario, var_unidad_organizacional, var_bloque)
   End If
   pro_textos
End Sub

Private Sub cmd_guardar_Click()
   Call pro_guardar_usuarios
End Sub

Private Sub cmd_imprimir_Click()
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call gPrintListView(lv_usuarios, "LISTADO DE usuarios")
        End If
End Sub

Private Sub cmd_nueva_empresa_Click()
   Dim list_item As ListItem
   contador = 0
   lv_seleccionempresas.ListItems.Clear
   rs.Open "select * from TB_EMPRESAS", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set list_item = lv_seleccionempresas.ListItems.Add(, , rs!VCHA_EMP_EMPRESA_ID)
         list_item.SubItems(1) = IIf(IsNull(rs!vcha_emp_nombre), "", rs!vcha_emp_nombre)
         contador = contador + 1
         rs.MoveNext:
   Wend
   rs.Close
   var_n = lv_seleccionempresas.ListItems.Count
   var_numero_renglones = lv_seleccionempresas.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_seleccionempresas.ColumnHeaders(2).Width = 4500
   Else
      lv_seleccionempresas.ColumnHeaders(2).Width = 4800
   End If
   frm_seleccionempresas.Visible = True
   lv_seleccionempresas.SetFocus
End Sub

Private Sub cmd_nuevo_Click()
   lv_empresas.ListItems.Clear
   Me.lv_plantas.ListItems.Clear
   Me.lv_modulos.ListItems.Clear
   Me.cmb_puestos = ""
   Call pro_limpiatextos(Me)
   txt_nombre_usuario.Enabled = True
   txt_nombre_usuario.SetFocus: var_modifica_registro_usuario = False
   Me.chk_permiso.Value = 0
   cmd_guardar.Enabled = True
   cmd_deshacer.Enabled = True
   txt_usuario.Enabled = False
End Sub

Private Sub cmd_nuevo_modulo_Click()
   Dim i As Integer
   Dim list_item As ListItem
   i = lv_plantas.ListItems.Count
   If i > 0 Then
      lv_seleccionmodulos.ListItems.Clear
      rs.Open "select * from TB_BLOQUES", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_seleccionmodulos.ListItems.Add(, , rs(0).Value)
            list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
            rs.MoveNext:
      Wend
      rs.Close
      frm_seleccionmodulos.Visible = True
      lv_seleccionmodulos.SetFocus
      var_n = lv_seleccionmodulos.ListItems.Count
      var_numero_renglones = lv_seleccionmodulos.Height / 312.5
      If var_n > var_numero_renglones Then
         lv_seleccionmodulos.ColumnHeaders(2).Width = 4550
      Else
         lv_seleccionmodulos.ColumnHeaders(2).Width = 4800
      End If
   Else
      MsgBox "No se a seleccionado ninguna planta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_planta_Click()
   Dim list_item As ListItem
   Dim i As Integer
   i = lv_empresas.ListItems.Count
   contador = 0
   If i > 0 Then
      lv_seleccionplantas.ListItems.Clear
      rs.Open "select * from TB_UNIDADESORGANIZACIONALES where vcha_emp_empresa_id = '" + lv_empresas.selectedItem.SubItems(1) + "'", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_seleccionplantas.ListItems.Add(, , rs!vcha_uor_unidad_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_uor_nombre), "", Trim(rs!vcha_uor_nombre))
            contador = contador + 1
            rs.MoveNext:
      Wend
      rs.Close
      frm_seleccionplantas.Visible = True
      lv_seleccionplantas.SetFocus
      var_n = lv_seleccionplantas.ListItems.Count
      var_numero_renglones = lv_seleccionplantas.Height / 312.5
      If var_n > var_numero_renglones Then
         lv_seleccionplantas.ColumnHeaders(2).Width = 4550
      Else
         lv_seleccionplantas.ColumnHeaders(2).Width = 4800
      End If
   Else
      MsgBox "No se a seleccionado ninguna empresa", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_puesto_Click()
   frm_seleccionpuestos.Visible = True
   cmb_puestos.SetFocus
End Sub

Private Sub cmd_permisos_Click()
   var_usuario_permiso = txt_usuario
   frmusuarios.Enabled = False
   var_activa_forma_permisos = Me.Name
   frmpermisos.Show
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 0
   Left = 0
   frm_seleccionmodulos.Visible = False
   frm_seleccionplantas.Visible = False
   frm_password.Visible = False
   frm_seleccionpuestos.Visible = False
   frm_seleccionempresas.Visible = False
   var_modifica_registro_usuario = True
   Call pro_llena_listview1
   rs.Open "select * from tb_puestos", cnn, adOpenDynamic, adLockBatchOptimistic
   Call RecsetToCombo(cmb_puestos.hwnd, rs, 1)
   rs.Close
   pro_textos
   If lv_usuarios.ListItems.Count = 0 Then
      cmd_guardar.Enabled = False
      cmd_deshacer.Enabled = False
      cmd_eliminar.Enabled = False
   Else
      cmd_guardar.Enabled = True
      cmd_deshacer.Enabled = True
      cmd_eliminar.Enabled = True
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_despliega_menu = True Then
      var_swpassword = False
      var_modifica_registro_usuario = False
   End If
   Call activa_forma(var_activa_forma_usuarios)
End Sub

Private Sub lv_empresas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_empresas, ColumnHeader)
End Sub

Private Sub lv_empresas_ItemClick(ByVal Item As MSComctlLib.ListItem)
   lv_plantas.ListItems.Clear
   lv_modulos.ListItems.Clear
   txt_puesto = ""
   rs.Open "select * from vw_relaciones_usuarios_unidades where vcha_usu_usuario_id = '" + txt_usuario + "' and vcha_emp_empresa_id = '" + lv_empresas.selectedItem.SubItems(1) + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
         Set list_item = lv_plantas.ListItems.Add(, , rs!vcha_usu_usuario_id)
         list_item.SubItems(1) = IIf(IsNull(rs!vcha_uor_unidad_id), "", rs!vcha_uor_unidad_id)
         list_item.SubItems(2) = IIf(IsNull(rs!vcha_uor_nombre), "", rs!vcha_uor_nombre)
         rs.MoveNext:
         numero_items_usuarios = numero_items_usuarios + 1
      Wend
      var_unidad_organizacional = lv_plantas.selectedItem.SubItems(1)
   Else
      var_unidad_organizacional = ""
   End If
   rs.Close
   i = lv_plantas.ListItems.Count
   If i > 0 Then
      lv_plantas.ListItems.Item(1).Selected = True
      lv_modulos.ListItems.Clear
      var_unidad_organizacional = lv_plantas.selectedItem.SubItems(1)
      rs.Open "select * from VW_RELACIONES_BLOQUES_UNIDADES where vcha_usu_usuario_id = '" + txt_usuario + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         While Not rs.EOF
            Set list_item = lv_modulos.ListItems.Add(, , rs(0).Value)
            list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
            list_item.SubItems(2) = IIf(IsNull(rs(2).Value), "", rs(2).Value)
            list_item.SubItems(3) = IIf(IsNull(rs(3).Value), "", rs(3).Value)
            rs.MoveNext:
            numero_items_usuarios = numero_items_usuarios + 1
         Wend
         var_bloque = lv_modulos.selectedItem.SubItems(2)
      Else
         var_bloque = ""
      End If
      rs.Close
      rs.Open "select * from VW_RELACIONES_BLOQUES_PUESTOS where vcha_usu_usuario_id = '" + txt_usuario + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_blo_bloque_id = '" + var_bloque + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         If IsNull(rs(4).Value) Then
            txt_puesto = ""
            var_puesto = ""
         Else
            txt_puesto = rs(4).Value
            var_puesto = rs(3).Value
         End If
      Else
         txt_puesto = ""
         var_puesto = ""
      End If
      rs.Close
   End If
End Sub

Private Sub lv_modulos_Click()
   If lv_modulos.ListItems.Count > 0 Then
      var_bloque = lv_modulos.selectedItem.SubItems(2)
      rs.Open "select * from VW_RELACIONES_BLOQUES_PUESTOS where vcha_usu_usuario_id = '" + txt_usuario + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_blo_bloque_id = '" + var_bloque + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         If IsNull(rs(4).Value) Then
            txt_puesto = ""
            var_puesto = ""
         Else
            txt_puesto = rs(4).Value
            var_puesto = rs(3).Value
         End If
      Else
         txt_puesto = ""
         var_puesto = ""
      End If
      rs.Close
   End If
End Sub

Private Sub lv_modulos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_modulos, ColumnHeader)
End Sub

Private Sub lv_modulos_ItemClick(ByVal Item As MSComctlLib.ListItem)
   var_bloque = lv_modulos.selectedItem.SubItems(2)
   rs.Open "select * from VW_RELACIONES_BLOQUES_PUESTOS where vcha_usu_usuario_id = '" + txt_usuario + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_blo_bloque_id = '" + var_bloque + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      If IsNull(rs(4).Value) Then
         txt_puesto = ""
         var_puesto = ""
      Else
         txt_puesto = rs(4).Value
         var_puesto = rs(3).Value
      End If
   Else
      txt_puesto = ""
      var_puesto = ""
   End If
   rs.Close
End Sub

Private Sub lv_plantas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_plantas, ColumnHeader)
End Sub

Private Sub lv_plantas_ItemClick(ByVal Item As MSComctlLib.ListItem)
   lv_modulos.ListItems.Clear
   var_unidad_organizacional = lv_plantas.selectedItem.SubItems(1)
   rs.Open "select * from VW_RELACIONES_BLOQUES_UNIDADES where vcha_usu_usuario_id = '" + txt_usuario + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
         Set list_item = lv_modulos.ListItems.Add(, , rs(0).Value)
         list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
         list_item.SubItems(2) = IIf(IsNull(rs(2).Value), "", rs(2).Value)
         list_item.SubItems(3) = IIf(IsNull(rs(3).Value), "", rs(3).Value)
         rs.MoveNext:
         numero_items_usuarios = numero_items_usuarios + 1
      Wend
      var_bloque = lv_modulos.selectedItem.SubItems(2)
   Else
      var_bloque = ""
   End If
   rs.Close
   rs.Open "select * from VW_RELACIONES_BLOQUES_PUESTOS where vcha_usu_usuario_id = '" + txt_usuario + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_blo_bloque_id = '" + var_bloque + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      If IsNull(rs(4).Value) Then
         txt_puesto = ""
         var_puesto = ""
      Else
         txt_puesto = rs(4).Value
         var_puesto = rs(3).Value
      End If
   Else
      txt_puesto = ""
      var_puesto = ""
   End If
   rs.Close
End Sub

Private Sub lv_seleccionempresas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_seleccionempresas, ColumnHeader)
End Sub

Private Sub lv_seleccionempresas_DblClick()
   Set TB_REL_USUARIOS_EMPRESAS_I = New TB_REL_USUARIOS_EMPRESAS_I
   var_anadir = TB_REL_USUARIOS_EMPRESAS_I.Anadir(txt_usuario, lv_seleccionempresas.selectedItem)
   frm_seleccionempresas.Visible = False
   pro_textos
End Sub

Private Sub lv_seleccionempresas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call lv_seleccionempresas_DblClick
   End If
   If KeyAscii = 27 Then
      frm_seleccionempresas.Visible = False
   End If
End Sub

Private Sub lv_seleccionmodulos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_seleccionmodulos, ColumnHeader)
End Sub

Private Sub lv_seleccionmodulos_DblClick()
   Set TB_REL_BLOQUES_UNIDADES_I = New TB_REL_BLOQUES_UNIDADES_I
   rs.Open "select * from tb_relaciones_bloques_unidades where VCHA_USU_USUARIO_ID = '" + txt_usuario + "' and VCHA_UOR_UNIDAD_ID  = '" + var_unidad_organizacional + "' AND VCHA_BLO_BLOQUE_ID = '" + lv_seleccionmodulos.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      MsgBox "Ya existe relacion entre la planta y el bloque", vbOKOnly, "ATENCION"
   Else
      var_inserta = TB_REL_BLOQUES_UNIDADES_I.Anadir(txt_usuario, var_unidad_organizacional, lv_seleccionmodulos.selectedItem)
      Set list_item = lv_modulos.ListItems.Add(, , txt_usuario)
      var_bloque = lv_seleccionmodulos.selectedItem
      list_item.SubItems(1) = var_unidad_organizacional
      list_item.SubItems(2) = lv_seleccionmodulos.selectedItem
      list_item.SubItems(3) = lv_seleccionmodulos.selectedItem.SubItems(1)
   End If
   rs.Close
   frm_seleccionmodulos.Visible = False
End Sub

Private Sub lv_seleccionmodulos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      frm_seleccionmodulos.Visible = False
   End If
   If KeyAscii = 13 Then
      Set TB_REL_BLOQUES_UNIDADES_I = New TB_REL_BLOQUES_UNIDADES_I
      rs.Open "select * from tb_relaciones_bloques_unidades where VCHA_USU_USUARIO_ID = '" + txt_usuario + "' and VCHA_UOR_UNIDAD_ID  = '" + var_unidad_organizacional + "' AND VCHA_BLO_BLOQUE_ID = '" + lv_seleccionmodulos.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         MsgBox "Ya existe relacion entre la planta y el bloque", vbOKOnly, "ATENCION"
      Else
         var_inserta = TB_REL_BLOQUES_UNIDADES_I.Anadir(txt_usuario, var_unidad_organizacional, lv_seleccionmodulos.selectedItem)
         var_bloque = lv_seleccionmodulos.selectedItem
         Set list_item = lv_modulos.ListItems.Add(, , txt_usuario)
         list_item.SubItems(1) = var_unidad_organizacional
         list_item.SubItems(2) = lv_seleccionmodulos.selectedItem
         list_item.SubItems(3) = lv_seleccionmodulos.selectedItem.SubItems(1)
      End If
      rs.Close
      frm_seleccionmodulos.Visible = False
   End If
End Sub

Private Sub lv_seleccionmodulos_LostFocus()
   frm_seleccionmodulos.Visible = False
End Sub

Private Sub lv_seleccionplantas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_seleccionplantas, ColumnHeader)
End Sub

Private Sub lv_seleccionplantas_DblClick()
   Set TB_REL_USUARIOS_UNIDADES_I = New TB_REL_USUARIOS_UNIDADES_I
   rs.Open "select * from TB_relaciones_usuarios_unidades WHERE vcha_usu_usuario_id = '" + txt_usuario + "' and VCHA_UOR_UNIDAD_ID = '" + lv_seleccionplantas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      MsgBox "Ya existe relacion entre el usuario y la planta", vbOKOnly, "ATENCION"
   Else
      var_inserta = TB_REL_USUARIOS_UNIDADES_I.Anadir(txt_usuario, lv_seleccionplantas.selectedItem)
      var_unidad_organizacional = lv_seleccionplantas.selectedItem
      Set list_item = lv_plantas.ListItems.Add(, , txt_usuario)
      list_item.SubItems(1) = lv_seleccionplantas.selectedItem
      list_item.SubItems(2) = lv_seleccionplantas.selectedItem.SubItems(1)
   End If
   rs.Close
   frm_seleccionplantas.Visible = False
End Sub

Private Sub lv_seleccionplantas_KeyPress(KeyAscii As Integer)
   Set TB_REL_USUARIOS_UNIDADES_I = New TB_REL_USUARIOS_UNIDADES_I
   If KeyAscii = 27 Then
      frm_seleccionplantas.Visible = False
   End If
   If KeyAscii = 13 Then
      rs.Open "select * from TB_RELACIONES_USUARIOS_UNIDADES WHERE vcha_usu_usuario_id = '" + txt_usuario + "' and VCHA_UOR_UNIDAD_ID = '" + lv_seleccionplantas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         MsgBox "Ya existe relacion entre el usuario y la planta", vbOKOnly, "ATENCION"
      Else
         Set TB_REL_BLOQUES_UNIDADES_I = New TB_REL_BLOQUES_UNIDADES_I
         var_inserta = TB_REL_USUARIOS_UNIDADES_I.Anadir(txt_usuario, lv_seleccionplantas.selectedItem)
         var_unidad_organizacional = lv_seleccionplantas.selectedItem
         var_inserta = TB_REL_BLOQUES_UNIDADES_I.Anadir(txt_usuario, var_unidad_organizacional, 1)
         Set list_item = lv_plantas.ListItems.Add(, , txt_usuario)
         list_item.SubItems(1) = lv_seleccionplantas.selectedItem
         list_item.SubItems(2) = lv_seleccionplantas.selectedItem.SubItems(1)
      End If
      rs.Close
      frm_seleccionplantas.Visible = False
   End If
End Sub

Private Sub lv_seleccionplantas_LostFocus()
   frm_seleccionplantas.Visible = False
End Sub

Private Sub lv_usuarios_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_usuarios, ColumnHeader)
End Sub

Private Sub lv_usuarios_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lv_empresas.ListItems.Clear
    lv_plantas.ListItems.Clear
    lv_modulos.ListItems.Clear
    txt_puesto = ""
    Set lv_usuarios.selectedItem = Item
        pro_textos
        var_modifica_registro_usuario = True
        txt_usuario.Enabled = False
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_usuarios.SetFocus
      Call pro_avanzar(Me, lv_usuarios, Button)
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_usuarios.ListItems(1).Selected = True
      lv_usuarios.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 4 Then
      numero_items_usuarios = lv_usuarios.ListItems.Count
      lv_usuarios.ListItems(numero_items_usuarios).Selected = True
      lv_usuarios.selectedItem.EnsureVisible
      pro_textos
   End If
err0:
End Sub


Sub pro_guardar_usuarios()
   Dim ok As Boolean
   Set Tb_usuarios = New Tb_usuarios
   Set TB_BITACORA_USUARIOS = New TB_BITACORA_USUARIOS
   ok = True
   If txt_nombre_usuario <> "" Then
      If var_hubo_cambios Then
         rs.Open "select * from tb_usuarios where vcha_USU_USUARIO_id = '" + txt_usuario + "'", cnn, adOpenDynamic, adLockOptimistic
         var_usuario_regreso = txt_usuario
         ok = Tb_usuarios.Anadir(txt_usuario, txt_nombre_usuario, txt_apellidos_usuario, txt_clave_usuario, txt_password, var_clave_usuario_global, chk_permiso)
         txt_usuario = var_usuario_regreso
         If ok Then
            bitacora = True
            If var_modifica_registro_usuario = False Then
               var_operacion_bitacora = "I"
               bitacora = TB_BITACORA_USUARIOS.Anadir(txt_usuario, "VCHA_USU_NOMBRE", var_operacion_bitacora, "", txt_nombre_usuario, var_clave_usuario_global, fun_NombrePc, Date)
            Else
               var_operacion_bitacora = "M"
               If rs(0) <> txt_usuario Then
                  bitacora = TB_BITACORA_USUARIOS.Anadir(txt_usuario, "VCHA_USU_USUARIO_ID", var_operacion_bitacora, rs(0), txt_usuario, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(1) <> txt_nombre_usuario Then
                  bitacora = TB_BITACORA_USUARIOS.Anadir(txt_usuario, "VCHA_USU_NOMBRE", var_operacion_bitacora, rs(1), txt_nombre_usuario, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(2) <> txt_apellidos_usuario Then
                  bitacora = TB_BITACORA_USUARIOS.Anadir(txt_usuario, "VCHA_USU_USUARIO", var_operacion_bitacora, rs(2), txt_apellidos_usuario, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(3) <> txt_clave_usuario Then
                  bitacora = TB_BITACORA_USUARIOS.Anadir(txt_usuario, "VCHA_USU_PASSWORD", var_operacion_bitacora, rs(3), txt_clave_usuario, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(4) <> chk_permisos Then
                  bitacora = TB_BITACORA_USUARIOS.Anadir(txt_usuario, "VCHA_USU_PERMISO", var_operacion_bitacora, rs(4), chk_permiso, var_clave_usuario_global, fun_NombrePc, Date)
               End If
            End If
            rs.Close
            rs.Open "update tb_usuarios set INTE_USU_PERMISO_CAMBIAR_TRANSPORTE = " + CStr(Me.chk_permiso_cambio_transporte) + ", vcha_numero_nomina = '" + Me.txt_nomina + "' where vcha_usu_usuario_id = '" + txt_usuario + "'", cnn, adOpenDynamic, adLockOptimistic
            pro_actualiza_ListView
            txt_usuario.Enabled = False
            MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
            txt_registros = lv_usuarios.ListItems.Count
            var_modifica_registro_usuario = True
         Else
            MsgBox "No se puede grabar registro: " + Tb_usuarios.MensajeError, vbOKOnly + vbCritical, "ATENCION"
         End If
      End If
   End If
   Set Tb_usuarios = Nothing: var_hubo_cambios = False
End Sub

Sub pro_elimina_usuarios()
   Dim var_llave_usuarios As String
   Dim i As Integer
   Set Tb_usuarios = New Tb_usuarios
   Set TB_BITACORA_USUARIOS = New TB_BITACORA_USUARIOS
   'On Error GoTo salir:
   ok = True
   If txt_usuario = "01" Then
      MsgBox "El administrador no puede ser eliminado", vbOKOnly, "ATENCION"
   Else
     If txt_usuario <> "" And txt_nombre_usuario <> "" And var_modifica_registro_usuario = True Then
        If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
            ok = Tb_usuarios.Eliminar(txt_usuario)
         Else
            GoTo SALIR:
         End If
         If ok Then
            var_operacion_bitacora = "E"
            bitacora = TB_BITACORA_USUARIOS.Anadir(txt_usuario, "VCHA_USU_NOMBRE", var_operacion_bitacora, txt_nombre_usuario, "", var_clave_usuario_global, fun_NombrePc, Date)
            numero_items_usuarios = numero_items_usuarios - 1
            rs.Open "delete from tb_relacion_usuarios_empresas where VCHA_USU_USUARIO_ID = '" + txt_usuario + "'", cnn, adOpenDynamic, adLockOptimistic
            rs.Open "delete from TB_RELACIONES_USUARIOS_UNIDADES where vcha_usu_usuario_id = '" + txt_usuario + "'", cnn, adOpenDynamic, adLockOptimistic
            rs.Open "delete from TB_RELACIONES_BLOQUES_PUESTOS where vcha_usu_usuario_id = '" + txt_usuario + "'", cnn, adOpenDynamic, adLockOptimistic
            rs.Open "delete from TB_RELACIONES_BLOQUES_UNIDADES where vcha_usu_usuario_id = '" + txt_usuario + "'", cnn, adOpenDynamic, adLockOptimistic
            MsgBox "Se Elimino Correctamente el Registro", vbInformation
            lv_usuarios.ListItems.Remove (lv_usuarios.selectedItem.Index)
            Call pro_limpiatextos(Me)
            txt_registros = lv_usuarios.ListItems.Count
            i = lv_usuarios.ListItems.Count
            If i > 0 Then
               lv_usuarios.selectedItem.Selected = True
            End If
            lv_empresas.ListItems.Clear
            lv_plantas.ListItems.Clear
            lv_modulos.ListItems.Clear
            txt_puesto = ""
            pro_textos
         Else
            MsgBox "No se puede eliminar registro: " + Tb_usuarios.MensajeError, vbOKOnly + vbCritical, "ATENCION"
         End If
      End If
   End If
SALIR:
   Set Tb_usuarios = Nothing
End Sub


Sub pro_llena_listview1()
   Dim list_item As ListItem
   rs.Open "select * from TB_usuarios", cnn, adOpenDynamic, adLockOptimistic
   numero_items_usuarios = 0
   While Not rs.EOF
      If rs!vcha_usu_usuario_id <> "1" Then
         Set list_item = lv_usuarios.ListItems.Add(, , rs(0).Value)
         list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
         list_item.SubItems(2) = IIf(IsNull(rs(2).Value), "", rs(2).Value)
         list_item.SubItems(3) = IIf(IsNull(rs(3).Value), "", rs(3).Value)
         list_item.SubItems(4) = IIf(IsNull(rs(4).Value), "", rs(4).Value)
         list_item.SubItems(5) = IIf(IsNull(rs(6).Value), "", rs(6).Value)
         list_item.SubItems(6) = IIf(IsNull(rs!INTE_USU_PERMISO_CAMBIAR_TRANSPORTE), 0, rs!INTE_USU_PERMISO_CAMBIAR_TRANSPORTE)
         
      End If
      rs.MoveNext:
      numero_items_usuarios = numero_items_usuarios + 1
    Wend
    rs.Close
    
End Sub


Sub pro_textos()
On Error GoTo err0:
   var_n = lv_usuarios.ListItems.Count
   If lv_usuarios.ListItems.Count > 0 Then
      txt_usuario = lv_usuarios.selectedItem
      txt_nombre_usuario = lv_usuarios.selectedItem.SubItems(1)
      txt_apellidos_usuario = lv_usuarios.selectedItem.SubItems(2)
      txt_clave_usuario = lv_usuarios.selectedItem.SubItems(3)
      txt_password = lv_usuarios.selectedItem.SubItems(4)
      txt_confirmar_password = lv_usuarios.selectedItem.SubItems(4)
      chk_permiso.Value = lv_usuarios.selectedItem.SubItems(5)
      If lv_usuarios.selectedItem.SubItems(6) = " " Then
         Me.lv_usuarios.selectedItem.SubItems(6) = "0"
      End If
      Me.chk_permiso_cambio_transporte = CDbl(lv_usuarios.selectedItem.SubItems(6))
      If chk_permiso.Value = 1 Then
         cmd_permisos.Enabled = True
      Else
         cmd_permisos.Enabled = False
      End If
      Me.txt_nomina = Me.lv_usuarios.selectedItem.SubItems(7)
      lv_empresas.ListItems.Clear
      rs.Open "select * from VW_RELACION_USUARIOS_EMPRESAS where vcha_usu_usuario_id = '" + txt_usuario + "'"
      If Not rs.EOF Then
         While Not rs.EOF
               Set list_item = lv_empresas.ListItems.Add(, , rs!vcha_usu_usuario_id)
               list_item.SubItems(1) = IIf(IsNull(rs!VCHA_EMP_EMPRESA_ID), "", rs!VCHA_EMP_EMPRESA_ID)
               list_item.SubItems(2) = IIf(IsNull(rs!vcha_emp_nombre), "", rs!vcha_emp_nombre)
               rs.MoveNext:
               numero_items_usuarios = numero_items_usuarios + 1
         Wend
      End If
      rs.Close
      Dim i As Integer
      i = lv_empresas.ListItems.Count
      If i > 0 Then
         lv_plantas.ListItems.Clear
         rs.Open "select * from vw_relaciones_usuarios_unidades where vcha_usu_usuario_id = '" + txt_usuario + "' and vcha_emp_empresa_id = '" + lv_empresas.selectedItem.SubItems(1) + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            While Not rs.EOF
                  Set list_item = lv_plantas.ListItems.Add(, , rs!vcha_usu_usuario_id)
                  list_item.SubItems(1) = IIf(IsNull(rs!vcha_uor_unidad_id), "", rs!vcha_uor_unidad_id)
                  list_item.SubItems(2) = IIf(IsNull(rs!vcha_uor_nombre), "", rs!vcha_uor_nombre)
                  rs.MoveNext:
                  numero_items_usuarios = numero_items_usuarios + 1
            Wend
            var_unidad_organizacional = lv_plantas.selectedItem.SubItems(1)
         Else
            var_unidad_organizacional = ""
         End If
         rs.Close
         i = lv_plantas.ListItems.Count
         If i > 0 Then
            lv_modulos.ListItems.Clear
            rs.Open "select * from VW_RELACIONES_BLOQUES_UNIDADES where vcha_usu_usuario_id = '" + txt_usuario + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               While Not rs.EOF
                     Set list_item = lv_modulos.ListItems.Add(, , rs(0).Value)
                     list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
                     list_item.SubItems(2) = IIf(IsNull(rs(2).Value), "", rs(2).Value)
                     list_item.SubItems(3) = IIf(IsNull(rs(3).Value), "", rs(3).Value)
                     rs.MoveNext:
                     numero_items_usuarios = numero_items_usuarios + 1
               Wend
               var_bloque = lv_modulos.selectedItem.SubItems(2)
            Else
               var_bloque = ""
            End If
            rs.Close
            rs.Open "select * from VW_RELACIONES_BLOQUES_PUESTOS where vcha_usu_usuario_id = '" + txt_usuario + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_blo_bloque_id = '" + var_bloque + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               If IsNull(rs(4).Value) Then
                  txt_puesto = ""
                  var_puesto = ""
               Else
                  txt_puesto = rs(4).Value
                  var_puesto = rs(3).Value
               End If
            Else
               txt_puesto = ""
               var_puesto = ""
            End If
            rs.Close
         End If
      End If
   End If
   If lv_usuarios.ListItems.Count = 0 Then
      cmd_guardar.Enabled = False
      cmd_deshacer.Enabled = False
      cmd_eliminar.Enabled = False
   Else
      cmd_guardar.Enabled = True
      cmd_deshacer.Enabled = True
      cmd_eliminar.Enabled = True
   End If
   txt_usuario.Enabled = False
   var_numero_renglones = lv_usuarios.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_usuarios.ColumnHeaders(3).Width = 2800
   Else
      lv_usuarios.ColumnHeaders(3).Width = 3000
   End If
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem

    If var_modifica_registro_usuario = False Then
        Set list_item = lv_usuarios.ListItems.Add(, , txt_usuario)
        list_item.SubItems(1) = txt_nombre_usuario
        list_item.SubItems(2) = txt_apellidos_usuario
        list_item.SubItems(3) = txt_clave_usuario
        list_item.SubItems(4) = txt_password
        list_item.SubItems(6) = Me.chk_permiso_cambio_transporte
        list_item.SubItems(7) = Me.txt_nomina
        list_item.EnsureVisible
        list_item.Selected = True
        numero_items_usuarios = numero_items_usuarios + 1
    Else
        lv_usuarios.ListItems.Item(lv_usuarios.selectedItem.Index).Checked = False
        lv_usuarios.ListItems.Item(lv_usuarios.selectedItem.Index) = txt_usuario
        lv_usuarios.ListItems.Item(lv_usuarios.selectedItem.Index).ListSubItems(1) = txt_nombre_usuario
        lv_usuarios.ListItems.Item(lv_usuarios.selectedItem.Index).ListSubItems(2) = txt_apellidos_usuario
        lv_usuarios.ListItems.Item(lv_usuarios.selectedItem.Index).ListSubItems(3) = txt_clave_usuario
        lv_usuarios.ListItems.Item(lv_usuarios.selectedItem.Index).ListSubItems(4) = txt_password
        lv_usuarios.ListItems.Item(lv_usuarios.selectedItem.Index).ListSubItems(6) = Me.chk_permiso_cambio_transporte
        lv_usuarios.ListItems.Item(lv_usuarios.selectedItem.Index).ListSubItems(6) = Me.txt_nomina
        lv_usuarios.ListItems.Item(lv_usuarios.selectedItem.Index).Selected = True
    End If
'    lv_usuarios.SetFocus
End Sub


Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
   frm_password.Visible = True
   txt_clave_usuario.SetFocus
End Sub

Private Sub txt_apellidos_usuario_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_apellidos_usuario_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_usuarios, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub

Private Sub txt_clave_usuario_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_clave_usuario_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 27 Then
      Me.frm_password.Visible = False
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_confirmar_password_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_confirmar_password_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Me.frm_password.Visible = False
   End If
   If KeyAscii = 27 Then
      Me.frm_password.Visible = False
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_usuario_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_usuario_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nomina_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_password_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_password_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 27 Then
      Me.frm_password.Visible = False
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_usuario_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_usuario_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub
