VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmtitulares 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Titulares"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   Icon            =   "frmtitulares.frx":0000
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
      Left            =   165
      TabIndex        =   48
      Top             =   1155
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   49
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
         Left            =   15
         TabIndex        =   50
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame frm_colonias 
      Height          =   2400
      Left            =   180
      TabIndex        =   45
      Top             =   1260
      Width           =   5685
      Begin MSComctlLib.ListView lv_colonias 
         Height          =   1830
         Left            =   45
         TabIndex        =   46
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
         NumItems        =   10
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
            Text            =   "pais"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "nombre pais"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "estado"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "nombre estado"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "municipio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "nombre municipio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "ciudad"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "nombre ciudad"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lbl_colonias 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   47
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmtitulares.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      Picture         =   "frmtitulares.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Guardar Alt + G"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   810
      Picture         =   "frmtitulares.frx":0ACE
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      Picture         =   "frmtitulares.frx":0BA0
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1470
      Picture         =   "frmtitulares.frx":0CA2
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5460
      Picture         =   "frmtitulares.frx":0DA4
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5505
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   39
      Top             =   60
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   150
      TabIndex        =   29
      Top             =   5310
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   1695
         TabIndex        =   25
         Top             =   150
         Width           =   2160
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   4080
         TabIndex        =   41
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
         Caption         =   "Busqueda de Titular:"
         Height          =   195
         Left            =   195
         TabIndex        =   30
         Top             =   195
         Width           =   1470
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Titulares "
      Height          =   4845
      Left            =   150
      TabIndex        =   26
      Top             =   465
      Width           =   5655
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5220
         Picture         =   "frmtitulares.frx":13DE
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Generar pedido "
         Top             =   4440
         Width           =   330
      End
      Begin VB.TextBox txt_fecha 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   59
         Top             =   4455
         Width           =   1350
      End
      Begin VB.CheckBox chk_descuento 
         Caption         =   "Marca descuento"
         Height          =   270
         Left            =   1290
         TabIndex        =   57
         Top             =   4200
         Width           =   1995
      End
      Begin VB.TextBox txt_unificador 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4155
         MaxLength       =   50
         TabIndex        =   55
         Top             =   3900
         Width           =   1350
      End
      Begin VB.TextBox txt_curp 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   53
         Top             =   3900
         Width           =   1965
      End
      Begin VB.TextBox txt_titular_anterior 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   51
         Top             =   3555
         Width           =   1965
      End
      Begin VB.CommandButton cmd_automatico 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5190
         Picture         =   "frmtitulares.frx":14E0
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Crear grupos automaticamente"
         Top             =   3225
         Width           =   330
      End
      Begin VB.TextBox txt_codigo_postal 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   9
         Top             =   915
         Width           =   900
      End
      Begin VB.TextBox txt_nombre_municipio 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1290
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   15
         Top             =   1905
         Width           =   4230
      End
      Begin VB.TextBox txt_nombre_colonia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1290
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1245
         Width           =   4230
      End
      Begin VB.TextBox txt_nombre_ciudad 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1290
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   13
         Top             =   1575
         Width           =   4230
      End
      Begin VB.TextBox txt_nombre_estado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1290
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   17
         Top             =   2235
         Width           =   4230
      End
      Begin VB.TextBox txt_nombre_pais 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1290
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   19
         Top             =   2565
         Width           =   4230
      End
      Begin VB.TextBox txt_nombre_grupo_real 
         Height          =   315
         Left            =   2205
         TabIndex        =   23
         Top             =   3225
         Width           =   2925
      End
      Begin VB.TextBox txt_limite_credito 
         Height          =   315
         Left            =   4005
         MaxLength       =   50
         TabIndex        =   21
         Top             =   2895
         Width           =   1515
      End
      Begin VB.TextBox txt_telefono 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   20
         Top             =   2895
         Width           =   1560
      End
      Begin VB.TextBox txt_nombre_titular 
         Height          =   315
         Left            =   2445
         MaxLength       =   50
         TabIndex        =   7
         Top             =   255
         Width           =   3060
      End
      Begin VB.TextBox txt_domicilio 
         Height          =   315
         Left            =   1290
         MaxLength       =   100
         TabIndex        =   8
         Top             =   585
         Width           =   4215
      End
      Begin VB.TextBox txt_colonia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1290
         TabIndex        =   10
         Top             =   1245
         Width           =   900
      End
      Begin VB.TextBox txt_ciudad 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1290
         TabIndex        =   12
         Top             =   1575
         Width           =   900
      End
      Begin VB.TextBox txt_estado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   2235
         Width           =   900
      End
      Begin VB.TextBox txt_pais 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2565
         Width           =   900
      End
      Begin VB.TextBox txt_titular 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1290
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   6
         Top             =   255
         Width           =   1140
      End
      Begin VB.TextBox txt_grupo_real 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   22
         Top             =   3225
         Width           =   900
      End
      Begin VB.TextBox txt_municipio 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1290
         TabIndex        =   14
         Top             =   1905
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha desc:"
         Height          =   195
         Left            =   135
         TabIndex        =   58
         Top             =   4515
         Width           =   885
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Unificador:"
         Height          =   195
         Index           =   13
         Left            =   3345
         TabIndex        =   56
         Top             =   3960
         Width           =   765
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "CURP:"
         Height          =   195
         Index           =   12
         Left            =   165
         TabIndex        =   54
         Top             =   3960
         Width           =   495
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Anterior:"
         Height          =   195
         Index           =   11
         Left            =   165
         TabIndex        =   52
         Top             =   3615
         Width           =   585
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "C.P."
         Height          =   195
         Index           =   10
         Left            =   180
         TabIndex        =   44
         Top             =   975
         Width           =   300
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Municipio:"
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   43
         Top             =   1965
         Width           =   720
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Limite Crédito:"
         Height          =   195
         Index           =   9
         Left            =   2985
         TabIndex        =   42
         Top             =   2955
         Width           =   990
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono:"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   40
         Top             =   2955
         Width           =   675
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Colonia:"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   37
         Top             =   1305
         Width           =   570
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Domicilio:"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   36
         Top             =   645
         Width           =   675
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad:"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   35
         Top             =   1635
         Width           =   540
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   34
         Top             =   2295
         Width           =   540
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Pais:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   33
         Top             =   2625
         Width           =   345
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   28
         Top             =   315
         Width           =   450
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Grupo real (F5):"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   27
         Top             =   3285
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1485
      Left            =   150
      TabIndex        =   31
      Top             =   5835
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
               Picture         =   "frmtitulares.frx":15E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmtitulares.frx":1EBC
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lv_titulares 
         Height          =   1230
         Left            =   45
         TabIndex        =   38
         Top             =   165
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   2170
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   17
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
            Text            =   "ciudad"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "colonia"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "domicilio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "telefono"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "grupo real"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Limite"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "municipio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "cp"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "anterior"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "curp"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "unificador"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Marca descuento"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Fecha descuento"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   32
      Top             =   285
      Width           =   5655
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
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtitulares.frx":2796
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtitulares.frx":3070
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtitulares.frx":394A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtitulares.frx":3EE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtitulares.frx":47C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtitulares.frx":509C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtitulares.frx":5976
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtitulares.frx":5A88
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtitulares.frx":5B9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtitulares.frx":5CAC
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmtitulares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_hubo_cambios As Boolean
Dim var_guardar_cambios As Boolean
Dim numero_items_titulares As Integer
Dim var_tipo_lista As Integer


Private Sub chk_descuento_Click()
   var_hubo_cambios = True
End Sub

Private Sub cmd_automatico_Click()
   Dim cmd As New Command
   Dim var_si As Integer
   Dim var_j As Integer
   Dim var_clave_string_ga As String
   Dim var_clave_string_gr As String
   rs.Open "select * from tb_gruposreales where vcha_gre_nombre = '" + Me.txt_nombre_titular + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
   var_j = 6
   If Not rs.EOF Then
      MsgBox "Ya existe un grupo con el nombre " + txt_nombre_titular + ", ¿desea crear otro?", vbYesNo, "ATENCION"
   End If
   rs.Close
   If var_j = 6 Then
      var_si = MsgBox("¿Deseas crear los grupos al titular?", vbYesNo, "ATENCION")
      If var_si = 6 Then
      rsaux5.Open "select vcha_emp_conexion from tb_empresas where len(vcha_emp_conexion) > 0 and vcha_emp_empresa_id = '02' ORDER BY INTE_EMP_ORDEN_CONEXION", cnn_distribucion, adOpenDynamic, adLockOptimistic
      var_clave_string_gr = ""
      var_clave_string_ga = ""
      While Not rsaux5.EOF
            var_conexion_importacion = IIf(IsNull(rsaux5(0).Value), "", rsaux5(0).Value)
            If Trim(var_conexion_importacion) <> "" Then
               If cnn_importacion.State = 1 Then
                  cnn_importacion.Close
               End If
               'MsgBox var_conexion_importacion
               cnn_importacion.Open var_conexion_importacion
               
               Set cmd.ActiveConnection = cnn_importacion
               cmd.CommandType = adCmdStoredProc
               cmd.CommandText = "SP_CREACION_GRUPOS"
               cmd("@TITULAR") = txt_titular
               cmd("@NOMBRE_TITULAR") = txt_nombre_titular
               cmd("@CLAVE_STRING_GR") = var_clave_string_gr
               cmd("@CLAVE_STRING_GA") = var_clave_string_ga
               cmd("@CLAVE_EMPRESA") = var_empresa
               'MsgBox cnn_importacion
               cmd.execute
               var_clave_string_gr = cmd("@CLAVE_STRING_GR")
               var_clave_string_ga = cmd("@CLAVE_STRING_GA")
               txt_grupo_real = cmd("@CLAVE_STRING_GR")
               txt_nombre_grupo_real = txt_nombre_titular
               Set cmd = Nothing
               rsaux.Open "select * from tb_gruposreales where vcha_gre_grupo_real_id = '" + txt_grupo_real + "'", cnn_importacion, adOpenDynamic, adLockOptimistic
               'rs.Open "update tb_gruposreales set vcha_emp_empresa_id = '" + var_empresa + "', vcha_gac_grupo_actual_id = '" + rsaux!VCHA_GAC_GRUPO_aCTUAL_ID + "' where vcha_gre_grupo_real_id = '" + txt_grupo_real + "'", cnn_importacion, adOpenDynamic, adLockOptimistic
               'rs.Open "update tb_gruposactuales set vcha_emp_empresa_id = '" + var_empresa + "' where vcha_gac_grupo_actual_id = '" + rsaux!VCHA_GAC_GRUPO_aCTUAL_ID + "'", cnn_importacion, adOpenDynamic, adLockOptimistic
               rsaux.Close
            End If
            rsaux5.MoveNext
      Wend
      rsaux5.Close
      Else
         MsgBox "Se a cancelado la creación de los grupos", vbOKOnly, "ATENCION"
      End If
   End If
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
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         rsaux5.Open "select vcha_emp_conexion from tb_empresas where len(vcha_emp_conexion) > 0", cnn_distribucion, adOpenDynamic, adLockOptimistic
         While Not rsaux5.EOF
               var_conexion_importacion = IIf(IsNull(rsaux5(0).Value), "", rsaux5(0).Value)
               If Trim(var_conexion_importacion) <> "" Then
                  If cnn_importacion.State = 1 Then
                     cnn_importacion.Close
                  End If
                  cnn_importacion.Open var_conexion_importacion
                  Call pro_elimina_titulares
               End If
               rsaux5.MoveNext
         Wend
         rsaux5.Close
      End If
      numero_items_titulares = numero_items_titulares - 1
      MsgBox "Se Elimino Correctamente el Registro", vbInformation
      lv_titulares.ListItems.Remove (lv_titulares.selectedItem.Index)
      Call pro_limpiatextos(Me)
      txt_registros = lv_titulares.ListItems.Count
      lv_titulares.selectedItem.Selected = True
      pro_textos
      
      
      rs.Open "select * from tb_titulares", cnn_distribucion, adOpenDynamic, adLockOptimistic
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
   If Not IsNumeric(Me.txt_limite_credito) Then
      Me.txt_limite_credito = 0
   End If
   If txt_grupo_real = "" Or txt_nombre_titular = "" Or txt_pais = "" Or txt_estado = "" Or txt_domicilio = "" Then
      MsgBox "Falta información", vbOKOnly, "ATENCION"
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
         var_posible = 0
         If Me.chk_descuento.Value = 1 Then
            If Not IsDate(Me.txt_fecha) Then
               var_posible = 1
            End If
         End If
         If var_posible = 1 Then
            MsgBox "La fecha de la aplicación del descuento es incorrecta", vbOKOnly, ""
            Me.chk_descuento.Value = 0
         End If
         rsaux5.Open "select vcha_emp_conexion from tb_empresas where len(vcha_emp_conexion) > 0 and vcha_emp_empresa_id = '02' ORDER BY INTE_EMP_ORDEN_CONEXION", cnn_distribucion, adOpenDynamic, adLockOptimistic
         While Not rsaux5.EOF
               var_conexion_importacion = IIf(IsNull(rsaux5(0).Value), "", rsaux5(0).Value)
               If Trim(var_conexion_importacion) <> "" Then
                  If cnn_importacion.State = 1 Then
                     cnn_importacion.Close
                  End If
                  cnn_importacion.Open var_conexion_importacion
                  Call pro_guardar_titulares
               End If
               rsaux5.MoveNext
         Wend
         rsaux5.Close
         pro_actualiza_ListView
         MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
         txt_registros = lv_titulares.ListItems.Count
         var_modifica_registro_titular = True
         var_hubo_cambios = False
         If rs.State = 1 Then
            rs.Close
         End If
         rs.Open "select * from tb_titulares", cnn_distribucion, adOpenDynamic, adLockOptimistic
         If rs.BOF Then
            var_guardar_cambios = False
            cmd_guardar.Enabled = False
            cmd_deshacer.Enabled = False
            cmd_eliminar.Enabled = False
         Else
            cmd_guardar.Enabled = True
            cmd_deshacer.Enabled = True
            cmd_eliminar.Enabled = True
            var_guardar_cambios = False
            lv_titulares.SetFocus
         End If
         rs.Close
      End If
   End If
End Sub

Private Sub cmd_imprimir_Click()
   Dim var_si As Integer
   var_si = MsgBox("Deseas imprimir solo el nombre", vbYesNo, "ATENION")
   If var_si = 6 Then
      Set reporte = appl.OpenReport(App.Path + "\rep_catalogo_titulares_solo_nombre.rpt")
      frmvistasprevias.cr.ReportSource = reporte
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      frmvistasprevias.cr.ViewReport
      frmvistasprevias.Caption = "Catálogo de Titulares"
      frmvistasprevias.Show
      
      reporte.ExportOptions.FormatType = crEFTExcel80
      reporte.ExportOptions.DestinationType = crEDTDiskFile
      reporte.ExportOptions.DiskFileName = "c:\catalogo_titulares.xls"
      reporte.Export False
      Set reporte = Nothing
   Else
      Set reporte = appl.OpenReport(App.Path + "\rep_catalogo_titulares.rpt")
      frmvistasprevias.cr.ReportSource = reporte
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      frmvistasprevias.cr.ViewReport
      frmvistasprevias.Caption = "Catálogo de Titulares"
      frmvistasprevias.Show
      Set reporte = Nothing
   End If
End Sub

Private Sub cmd_nuevo_Click()
        Call pro_limpiatextos(Me)
        txt_titular.Enabled = False
        txt_nombre_titular.SetFocus: var_modifica_registro_titular = False
        cmd_guardar.Enabled = True
        cmd_deshacer.Enabled = True
        var_guardar_cambios = True
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_titular = False Then
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

Private Sub Command2_Click()
      var_tipo_datos_adicionales = 3
      var_hubo_cambios = True
      If Trim(Me.txt_titular) <> "" Then
         rs.Open "select * from tb_titulares where vcha_tit_titular_id = '" + Me.txt_titular + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_clave_titular_global = Me.txt_titular
            var_nombre_cliente_ad = IIf(IsNull(rs!vcha_tit_nombre_2), "", rs!vcha_tit_nombre_2)
            var_paterno_cliente_ad = IIf(IsNull(rs!vcha_tit_paterno), "", rs!vcha_tit_paterno)
            var_materno_cliente_ad = IIf(IsNull(rs!vcha_tit_materno), "", rs!vcha_tit_materno)
            var_numero_cliente_ad = IIf(IsNull(rs!vcha_tit_numero), "", rs!vcha_tit_numero)
            var_clave_tel_pais_ad = IIf(IsNull(rs!vcha_tit_clave_tel_pais), "", rs!vcha_tit_clave_tel_pais)
            var_clave_tel_estado_ad = IIf(IsNull(rs!vcha_tit_clave_tel_estado), "", rs!vcha_tit_clave_tel_estado)
            var_calle_cliente_ad = IIf(IsNull(rs!vcha_tit_calle), "", rs!vcha_tit_calle)
            var_numero_interno_cliente_ad = IIf(IsNull(rs!vcha_tit_numero_interno), "", rs!vcha_tit_numero_interno)
            frmdatos_adisionales.Show 1
         Else
            MsgBox "El establecimiento no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
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
   var_cadena_seguridad = ""
   Top = 0
   Left = 2900
   frm_lista.Visible = False
   frm_colonias.Visible = False
   txt_titular.Enabled = False
   var_modifica_registro_titular = True
   Call pro_encabezadosView(Me, lv_titulares, False)
   Call pro_llena_listview1
   pro_textos
   rs.Open "select * from vw_titulares_1 where vcha_emp_empresa_id = '" + var_empresa + "' order by vcha_tit_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
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
   Dim var_guardar As Integer
   Call activa_forma(var_activa_forma_titulares)
End Sub

Private Sub lv_colonias_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_colonias, ColumnHeader)
End Sub

Private Sub lv_colonias_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_colonia = lv_colonias.selectedItem
      txt_nombre_colonia = lv_colonias.selectedItem.SubItems(1)
      txt_pais = lv_colonias.selectedItem.SubItems(2)
      txt_nombre_pais = lv_colonias.selectedItem.SubItems(3)
      txt_estado = lv_colonias.selectedItem.SubItems(4)
      txt_nombre_estado = lv_colonias.selectedItem.SubItems(5)
      txt_municipio = lv_colonias.selectedItem.SubItems(6)
      txt_nombre_municipio = lv_colonias.selectedItem.SubItems(7)
      txt_ciudad = lv_colonias.selectedItem.SubItems(8)
      txt_nombre_ciudad = lv_colonias.selectedItem.SubItems(9)
      frm_colonias.Visible = False
      Me.txt_telefono.SetFocus
   End If
   If KeyAscii = 27 Then
      frm_colonias.Visible = False
   End If
End Sub

Private Sub lv_colonias_LostFocus()
   frm_colonias.Visible = False
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_tipo_lista = 10 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_ciudad = lv_lista.selectedItem
            txt_nombre_ciudad = lv_lista.selectedItem.SubItems(1)
         Else
            txt_ciudad = ""
            txt_nombre_ciudad = ""
         End If
         txt_ciudad.SetFocus
      End If
      If var_tipo_lista = 11 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_municipio = lv_lista.selectedItem
            txt_nombre_municipio = lv_lista.selectedItem.SubItems(1)
         Else
            txt_municipio = ""
            txt_nombre_municipio = ""
         End If
         txt_municipio.SetFocus
      End If
      If var_tipo_lista = 12 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_estado = lv_lista.selectedItem
            txt_nombre_estado = lv_lista.selectedItem.SubItems(1)
         Else
            txt_estado = ""
            txt_nombre_estado = ""
         End If
         txt_estado.SetFocus
      End If
      If var_tipo_lista = 13 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_pais = lv_lista.selectedItem
            txt_nombre_pais = lv_lista.selectedItem.SubItems(1)
         Else
            txt_pais = ""
            txt_nombre_pais = ""
         End If
         txt_pais.SetFocus
      End If
      If var_tipo_lista = 14 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_colonia = lv_lista.selectedItem
            txt_nombre_colonia = lv_lista.selectedItem.SubItems(1)
         Else
            txt_colonia = ""
            txt_nombre_colonia = ""
         End If
         txt_colonia.SetFocus
      End If
      If var_tipo_lista = 15 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_grupo_real = lv_lista.selectedItem
            txt_nombre_grupo_real = lv_lista.selectedItem.SubItems(1)
         Else
            txt_grupo_real = ""
            txt_nombre_grupo_real = ""
         End If
         txt_grupo_real.SetFocus
      End If
      If var_tipo_lista = 16 Then
         If lv_lista.ListItems.Count > 0 Then
            rs.Open "select * from vw_colonias where vcha_col_cp = '" + Me.txt_codigo_postal + "' AND VCHA_PAI_PAIS_ID = '" + lv_lista.selectedItem + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
            If rs.RecordCount = 1 Then
               txt_colonia = IIf(IsNull(rs!VCHA_COL_COLONIA_ID), "", rs!VCHA_COL_COLONIA_ID)
               txt_nombre_colonia = IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
               txt_pais = IIf(IsNull(rs!VCHA_PAI_PAIS_ID), "", rs!VCHA_PAI_PAIS_ID)
               txt_nombre_pais = IIf(IsNull(rs!vcha_pai_nombre), "", rs!vcha_pai_nombre)
               txt_estado = IIf(IsNull(rs!VCHA_EST_ESTADO_ID), "", rs!VCHA_EST_ESTADO_ID)
               txt_nombre_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
               txt_municipio = IIf(IsNull(rs!VCHA_MUN_MUNICIPIO_ID), "", rs!VCHA_MUN_MUNICIPIO_ID)
               txt_nombre_municipio = IIf(IsNull(rs!vcha_mun_nombre), "", rs!vcha_mun_nombre)
               txt_ciudad = IIf(IsNull(rs!VCHA_CIU_CIUDAD_ID), "", rs!VCHA_CIU_CIUDAD_ID)
               txt_nombre_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
               rs.Close
               txt_telefono.SetFocus
               frm_lista.Visible = False
            Else
               'rs.MoveFirst
               If Not rs.EOF Then
                  lv_colonias.ListItems.Clear
                  While Not rs.EOF
                        Set list_item = lv_colonias.ListItems.Add(, , rs!VCHA_COL_COLONIA_ID)
                        list_item.SubItems(1) = IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
                        list_item.SubItems(2) = IIf(IsNull(rs!VCHA_PAI_PAIS_ID), "", rs!VCHA_PAI_PAIS_ID)
                        list_item.SubItems(3) = IIf(IsNull(rs!vcha_pai_nombre), "", rs!vcha_pai_nombre)
                        list_item.SubItems(4) = IIf(IsNull(rs!VCHA_EST_ESTADO_ID), "", rs!VCHA_EST_ESTADO_ID)
                        list_item.SubItems(5) = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                        list_item.SubItems(6) = IIf(IsNull(rs!VCHA_MUN_MUNICIPIO_ID), "", rs!VCHA_MUN_MUNICIPIO_ID)
                        list_item.SubItems(7) = IIf(IsNull(rs!vcha_mun_nombre), "", rs!vcha_mun_nombre)
                        list_item.SubItems(8) = IIf(IsNull(rs!VCHA_CIU_CIUDAD_ID), "", rs!VCHA_CIU_CIUDAD_ID)
                        list_item.SubItems(9) = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                        rs.MoveNext
                  Wend
                  lbl_colonias = "COLONIAS DEL C.P. " + txt_codigo_postal
                  var_n = lv_colonias.ListItems.Count
                  If var_n > 6 Then
                     lv_colonias.ColumnHeaders(2).Width = 4270.71
                  Else
                     lv_colonias.ColumnHeaders(2).Width = 4499.71
                  End If
                  frm_colonias.Visible = True
                  lv_colonias.SetFocus
               Else
                  MsgBox "Código postal incorrecto", vbOKOnly, "ATENCION"
               End If
               rs.Close
            End If
         Else
            txt_colonia = ""
            txt_nombre_colonia = ""
            txt_pais = ""
            txt_nombre_pais = ""
            txt_estado = ""
            txt_nombre_estado = ""
            txt_municipio = ""
            txt_nombre_municipio = ""
            txt_ciudad = ""
            txt_nombre_ciudad = ""
         End If
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

Private Sub lv_titulares_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_titulares, ColumnHeader)
End Sub

Private Sub lv_titulares_ItemClick(ByVal item As MSComctlLib.ListItem)
    Set lv_titulares.selectedItem = item
        pro_textos
        var_modifica_registro_titular = True

End Sub


Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_titulares.SetFocus
      Call pro_avanzar(Me, lv_titulares, Button)
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_titulares.ListItems(1).Selected = True
      lv_titulares.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 4 Then
      numero_items_titulares = lv_titulares.ListItems.Count
      lv_titulares.ListItems(numero_items_titulares).Selected = True
      lv_titulares.selectedItem.EnsureVisible
      pro_textos
   End If
err0:
End Sub


Sub pro_guardar_titulares()
   Dim ok As Boolean
   Set TB_TITULARES = New TB_TITULARES
   Set TB_BITACORA_TITULARES = New TB_BITACORA_TITULARES
   ok = True
   If txt_grupo_real <> "" And txt_pais <> "" And txt_estado <> "" Then
      If var_hubo_cambios Then
         'MsgBox CNN.ConnectionString
         rs.Open "select * from tb_titulares where vcha_tit_titular_id = '" + txt_titular + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
         'If Not rs.EOF Then
         var_titular_regreso = txt_titular
         ok = TB_TITULARES.Anadir(txt_grupo_real, txt_titular, txt_nombre_titular, txt_pais, txt_estado, txt_municipio, txt_ciudad, txt_colonia, txt_domicilio, txt_telefono, txt_limite_credito, txt_codigo_postal)
         If Trim(var_titular_regreso) <> "" Then
            txt_titular = var_titular_regreso
         End If
         If Trim(Me.txt_titular_anterior) = "" Then
            Me.txt_titular_anterior = Me.txt_titular
         End If
         If ok Then
            If IsDate(Me.txt_fecha) Then
               If Me.chk_descuento.Value = 0 Then
                  rsaux4.Open "update tb_titulares set vcha_emp_empresa_id = '" + var_empresa + "', vcha_tit_titular_anterior_id = '" + Me.txt_titular_anterior + "', inte_tit_marca_descuento = 0, dtim_tit_fecha_Descuento =  NULL where vcha_tit_titular_id = '" + txt_titular + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
               Else
                  var_dia = CStr(Day(CDate(Me.txt_fecha)))
                  var_mes = CStr(Month(CDate(Me.txt_fecha)))
                  var_año = CStr(Year(CDate(Me.txt_fecha)))
                  If Len(Trim(var_dia)) = 1 Then
                     var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                  
                  rsaux4.Open "update tb_titulares set vcha_emp_empresa_id = '" + var_empresa + "', vcha_tit_titular_anterior_id = '" + Me.txt_titular_anterior + "', inte_tit_marca_descuento = 1, dtim_tit_fecha_Descuento  = " + var_fecha_fin + " where vcha_tit_titular_id = '" + txt_titular + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
               End If
            Else
               If Me.chk_descuento.Value = 1 Then
                  Me.chk_descuento = 0
                  rsaux4.Open "update tb_titulares set vcha_emp_empresa_id = '" + var_empresa + "', vcha_tit_titular_anterior_id = '" + Me.txt_titular_anterior + "', inte_tit_marca_descuento = 0 where vcha_tit_titular_id = '" + txt_titular + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
               Else
                  rsaux4.Open "update tb_titulares set vcha_emp_empresa_id = '" + var_empresa + "', vcha_tit_titular_anterior_id = '" + Me.txt_titular_anterior + "', inte_tit_marca_descuento = 0 where vcha_tit_titular_id = '" + txt_titular + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
               End If
            End If
            bitacora = True
            If var_modifica_registro_titular = False Then
               var_operacion_bitacora = "I"
               bitacora = TB_BITACORA_TITULARES.Anadir(txt_titular, "VCHA_TIT_NOMBRE", var_operacion_bitacora, "", txt_nombre_titular, var_clave_usuario_global, fun_NombrePc, Date)
            Else
               var_operacion_bitacora = "M"
               If rs!vcha_gre_grupo_real_id <> txt_grupo_real Then
                  bitacora = TB_BITACORA_TITULARES.Anadir(txt_titular, "VCHA_GRE_GRUPO_REAL_ID", var_operacion_bitacora, rs!vcha_gre_grupo_real_id, txt_grupo_real, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!vcha_tit_titular_id <> txt_titular Then
                  bitacora = TB_BITACORA_TITULARES.Anadir(txt_titular, "VCHA_TIT_TITULAR_ID", var_operacion_bitacora, rs!vcha_tit_titular_id, txt_titular, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!VCHA_TIT_NOMBRE <> txt_nombre_titular Then
                  bitacora = TB_BITACORA_TITULARES.Anadir(txt_titular, "VCHA_TIT_NOMBRE", var_operacion_bitacora, rs!VCHA_TIT_NOMBRE, txt_nombre_titular, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!VCHA_PAI_PAIS_ID <> txt_pais Then
                  bitacora = TB_BITACORA_TITULARES.Anadir(txt_titular, "VCHA_PAI_PAIS_ID", var_operacion_bitacora, rs!VCHA_PAI_PAIS_ID, txt_pais, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!VCHA_EST_ESTADO_ID <> txt_estado Then
                  bitacora = TB_BITACORA_TITULARES.Anadir(txt_titular, "VCHA_EST_ESTADO_ID", var_operacion_bitacora, rs!VCHA_EST_ESTADO_ID, txt_estado, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!VCHA_CIU_CIUDAD_ID <> txt_ciudad Then
                  bitacora = TB_BITACORA_TITULARES.Anadir(txt_titular, "VCHA_CIU_CIUDAD_ID", var_operacion_bitacora, rs!VCHA_CIU_CIUDAD_ID, txt_ciudad, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!VCHA_COL_COLONIA_ID <> txt_colonia Then
                  bitacora = TB_BITACORA_TITULARES.Anadir(txt_titular, "VCHA_COL_COLONIA_ID", var_operacion_bitacora, rs!VCHA_COL_COLONIA_ID, txt_colonia, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!VCHA_TIT_DOMICILIO <> txt_domicilio Then
                  bitacora = TB_BITACORA_TITULARES.Anadir(txt_titular, "VCHA_TIT_DOMICILIO", var_operacion_bitacora, rs!VCHA_TIT_DOMICILIO, txt_domicilio, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!VCHA_TIT_TELEFONO <> txt_telefono Then
                  bitacora = TB_BITACORA_TITULARES.Anadir(txt_titular, "VCHA_TIT_TELEFONO", var_operacion_bitacora, rs!VCHA_TIT_TELEFONO, txt_telefono, var_clave_usuario_global, fun_NombrePc, Date)
               End If
            End If
            rs.Close
         Else
            MsgBox "No se puede grabar registro: " + TB_TITULARES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
         End If
         'End If
      End If
   End If
   Set TB_TITULARES = Nothing
End Sub

Sub pro_elimina_titulares()
   Dim var_llave_usuarios As String
   Set TB_TITULARES = New TB_TITULARES
   Set TB_BITACORA_TITULARES = New TB_BITACORA_TITULARES
   On Error GoTo salir:
   ok = True
   If txt_grupo_real <> "" And txt_titular <> "" And var_modifica_registro_titular = True Then
      'If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_TITULARES.Eliminar(txt_titular)
      'Else
      '   GoTo salir:
      'End If
      If ok Then
         var_operacion_bitacora = "E"
         bitacora = TB_BITACORA_TITULARES.Anadir(txt_titular, "VCHA_TIT_NOMBRE", var_operacion_bitacora, txt_nombre_titular, "", var_clave_usuario_global, fun_NombrePc, Date)
         If numero_items_titulares > 11 Then
            lv_titulares.ColumnHeaders(2).Width = 4200
         Else
            lv_titulares.ColumnHeaders(2).Width = 4499.71
         End If
      Else
         MsgBox "No se puede eliminar registro: " + TB_TITULARES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
   Set TB_TITULARES = Nothing
End Sub


Sub pro_llena_listview1()
   Dim list_item As ListItem
   rs.Open "select DISTINCT vcha_tit_titular_id,vcha_tit_nombre, vcha_pai_pais_id, vcha_est_estado_id, vcha_ciu_ciudad_id, VCHA_COL_COLONIA_ID, vcha_tit_domicilio, vcha_tit_telefono, vcha_gre_grupo_real_id, floa_tit_limite_credito, floa_tit_limite_credito, vcha_mun_municipio_id, vcha_tit_cp, vcha_tit_titular_anterior_id, vcha_tit_curp, inte_tit_unificador, inte_tit_marca_descuento, dtim_TIT_fecha_descuento  from vw_titulares_1 where VCHA_TIT_TITULAR_ID <> '' and vcha_emp_empresa_id = '" + var_empresa + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
   numero_items_titulares = 0
   While Not rs.EOF
      Set list_item = lv_titulares.ListItems.Add(, , rs!vcha_tit_titular_id)
      list_item.SubItems(1) = IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE)
      list_item.SubItems(2) = IIf(IsNull(rs!VCHA_PAI_PAIS_ID), "", rs!VCHA_PAI_PAIS_ID)
      list_item.SubItems(3) = IIf(IsNull(rs!VCHA_EST_ESTADO_ID), "", rs!VCHA_EST_ESTADO_ID)
      list_item.SubItems(4) = IIf(IsNull(rs!VCHA_CIU_CIUDAD_ID), "", rs!VCHA_CIU_CIUDAD_ID)
      list_item.SubItems(5) = IIf(IsNull(rs!VCHA_COL_COLONIA_ID), "", rs!VCHA_COL_COLONIA_ID)
      list_item.SubItems(6) = IIf(IsNull(rs!VCHA_TIT_DOMICILIO), "", rs!VCHA_TIT_DOMICILIO)
      list_item.SubItems(7) = IIf(IsNull(rs!VCHA_TIT_TELEFONO), "", rs!VCHA_TIT_TELEFONO)
      list_item.SubItems(8) = IIf(IsNull(rs!vcha_gre_grupo_real_id), "", rs!vcha_gre_grupo_real_id)
      list_item.SubItems(9) = IIf(IsNull(rs!floa_tit_limite_credito), "", rs!floa_tit_limite_credito)
      list_item.SubItems(10) = IIf(IsNull(rs!VCHA_MUN_MUNICIPIO_ID), "", rs!VCHA_MUN_MUNICIPIO_ID)
      list_item.SubItems(11) = IIf(IsNull(rs!VCHA_TIT_CP), "", rs!VCHA_TIT_CP)
      list_item.SubItems(12) = IIf(IsNull(rs!VCHA_TIT_TITULAR_ANTERIOR_ID), "", rs!VCHA_TIT_TITULAR_ANTERIOR_ID)
      list_item.SubItems(13) = IIf(IsNull(rs!vcha_tit_curp), "", rs!vcha_tit_curp)
      list_item.SubItems(14) = IIf(IsNull(rs!inte_tit_unificador), 0, rs!inte_tit_unificador)
      list_item.SubItems(15) = IIf(IsNull(rs!inte_tit_marca_Descuento), 0, rs!inte_tit_marca_Descuento)
      list_item.SubItems(16) = IIf(IsNull(rs!dtim_tit_fecha_descuento), "", rs!dtim_tit_fecha_descuento)
      rs.MoveNext:
      numero_items_titulares = numero_items_titulares + 1
   Wend
   rs.Close
End Sub


Sub pro_textos()
On Error GoTo err0:
Dim var_n As Double
var_n = lv_titulares.ListItems.Count
   If var_n > 0 Then
      txt_titular = lv_titulares.selectedItem
      txt_nombre_titular = lv_titulares.selectedItem.SubItems(1)
      txt_pais = lv_titulares.selectedItem.SubItems(2)
      txt_estado = lv_titulares.selectedItem.SubItems(3)
      txt_ciudad = lv_titulares.selectedItem.SubItems(4)
      txt_colonia = lv_titulares.selectedItem.SubItems(5)
      txt_domicilio = lv_titulares.selectedItem.SubItems(6)
      txt_telefono = lv_titulares.selectedItem.SubItems(7)
      txt_grupo_real = lv_titulares.selectedItem.SubItems(8)
      txt_limite_credito = lv_titulares.selectedItem.SubItems(9)
      txt_municipio = lv_titulares.selectedItem.SubItems(10)
      txt_codigo_postal = lv_titulares.selectedItem.SubItems(11)
      txt_titular_anterior = lv_titulares.selectedItem.SubItems(12)
      Me.txt_curp = lv_titulares.selectedItem.SubItems(13)
      Me.txt_unificador = lv_titulares.selectedItem.SubItems(14)
      Me.chk_descuento = lv_titulares.selectedItem.SubItems(15)
      Me.txt_fecha = lv_titulares.selectedItem.SubItems(16)
   End If
   rs.Open "select * from tb_paises where vcha_pai_pais_id = '" + txt_pais + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      txt_nombre_pais = IIf(IsNull(rs!vcha_pai_nombre), "", rs!vcha_pai_nombre)
   Else
      txt_nombre_pais = ""
   End If
   rs.Close
   rs.Open "select * from tb_estados where vcha_est_Estado_id = '" + txt_estado + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      txt_nombre_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
   Else
      txt_nombre_estado = ""
   End If
   rs.Close
   rs.Open "select * from tb_municipios where vcha_mun_municipio_id = '" + txt_municipio + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      txt_nombre_municipio = IIf(IsNull(rs!vcha_mun_nombre), "", rs!vcha_mun_nombre)
   Else
      txt_nombre_municipio = ""
   End If
   rs.Close
   rs.Open "select * from tb_ciudades where vcha_ciu_ciudad_id = '" + txt_ciudad + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      txt_nombre_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
   Else
      txt_nombre_ciudad = ""
   End If
   rs.Close
   rs.Open "select * from tb_colonias where vcha_col_colonia_id = '" + txt_colonia + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      txt_nombre_colonia = IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
   Else
      txt_nombre_colonia = ""
   End If
   rs.Close
   rs.Open "select * from TB_GRUPOSREALES where vcha_gre_grupo_real_id = '" + txt_grupo_real + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      txt_nombre_grupo_real = IIf(IsNull(rs!VCHA_GRE_NOMBRE), "", rs!VCHA_GRE_NOMBRE)
   Else
      txt_nombre_grupo_real = ""
   End If
   rs.Close
   
   var_numero_renglones = lv_titulares.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_titulares.ColumnHeaders(2).Width = 3850
   Else
      lv_titulares.ColumnHeaders(2).Width = 4099.71
   End If
   var_hubo_cambios = False
   var_modifica_registro_titular = True
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro_titular = False Then
        Set list_item = lv_titulares.ListItems.Add(, , txt_titular)
        list_item.SubItems(1) = txt_nombre_titular
        list_item.SubItems(2) = txt_pais
        list_item.SubItems(3) = txt_estado
        list_item.SubItems(4) = txt_ciudad
        list_item.SubItems(5) = txt_colonia
        list_item.SubItems(6) = txt_domicilio
        list_item.SubItems(7) = txt_telefono
        list_item.SubItems(8) = txt_grupo_real
        list_item.SubItems(9) = txt_limite_credito
        list_item.SubItems(10) = txt_municipio
        list_item.SubItems(11) = txt_codigo_postal
        list_item.SubItems(12) = txt_titular_anterior
        list_item.SubItems(13) = Me.txt_curp
        list_item.SubItems(14) = Me.txt_unificador
        list_item.SubItems(15) = Me.chk_descuento
        list_item.SubItems(16) = Me.txt_fecha
        list_item.EnsureVisible
        list_item.Selected = True
        numero_items_titulares = numero_items_titulares + 1
    Else
        lv_titulares.ListItems.item(lv_titulares.selectedItem.Index).Checked = False
        lv_titulares.ListItems.item(lv_titulares.selectedItem.Index) = txt_titular
        lv_titulares.ListItems.item(lv_titulares.selectedItem.Index).ListSubItems(1) = txt_nombre_titular
        lv_titulares.ListItems.item(lv_titulares.selectedItem.Index).ListSubItems(2) = txt_pais
        lv_titulares.ListItems.item(lv_titulares.selectedItem.Index).ListSubItems(3) = txt_estado
        lv_titulares.ListItems.item(lv_titulares.selectedItem.Index).ListSubItems(4) = txt_ciudad
        lv_titulares.ListItems.item(lv_titulares.selectedItem.Index).ListSubItems(5) = txt_colonia
        lv_titulares.ListItems.item(lv_titulares.selectedItem.Index).ListSubItems(6) = txt_domicilio
        lv_titulares.ListItems.item(lv_titulares.selectedItem.Index).ListSubItems(7) = txt_telefono
        lv_titulares.ListItems.item(lv_titulares.selectedItem.Index).ListSubItems(8) = txt_grupo_real
        lv_titulares.ListItems.item(lv_titulares.selectedItem.Index).ListSubItems(9) = txt_limite_credito
        lv_titulares.ListItems.item(lv_titulares.selectedItem.Index).ListSubItems(10) = txt_municipio
        lv_titulares.ListItems.item(lv_titulares.selectedItem.Index).ListSubItems(11) = txt_codigo_postal
        lv_titulares.ListItems.item(lv_titulares.selectedItem.Index).ListSubItems(12) = txt_titular_anterior
        lv_titulares.ListItems.item(lv_titulares.selectedItem.Index).ListSubItems(13) = Me.txt_curp
        lv_titulares.ListItems.item(lv_titulares.selectedItem.Index).ListSubItems(14) = Me.txt_unificador
        lv_titulares.ListItems.item(lv_titulares.selectedItem.Index).ListSubItems(15) = Me.chk_descuento
        lv_titulares.ListItems.item(lv_titulares.selectedItem.Index).ListSubItems(16) = Me.txt_fecha
        lv_titulares.ListItems.item(lv_titulares.selectedItem.Index).Selected = True
    End If
End Sub



Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_titulares, Me.txt_buscar, True)
      txt_buscar = ""
      pro_textos
   End If
End Sub

Private Sub txt_ciudad_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_ciudad_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_ciudades where vcha_pai_pais_id = '" + txt_pais + "' and vcha_est_estado_id = '" + txt_estado + "' order by vcha_ciu_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_CIU_CIUDAD_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CIUDADES DE " + txt_nombre_estado
      var_tipo_lista = 10
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 7 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
   If KeyCode = 117 Then
      frmciudades.Show
   End If
End Sub

Private Sub txt_ciudad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_codigo_postal_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_codigo_postal_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_codigo_postal_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Me.Enabled = False
      var_activa_forma_direcciones = Me.Name
      frmtitulares.Enabled = False
      frmdirecciones.Show
      If var_aceptar_direccion = True Then
         txt_pais = var_dir_pais
         txt_nombre_pais = var_dir_nombre_pais
         txt_estado = var_dir_estado
         txt_nombre_estado = var_dir_nombre_estado
         txt_municipio = var_dir_municipio
         txt_nombre_municipio = var_dir_nombre_municipio
         txt_ciudad = var_dir_ciudad
         txt_nombre_ciudad = var_dir_nombre_ciudad
         txt_colonia = var_dir_colonia
         txt_nombre_colonia = var_dir_nombre_colonia
         txt_codigo_postal = var_dir_codigo_postal
      End If
   End If
End Sub

Private Sub txt_codigo_postal_KeyPress(KeyAscii As Integer)
   Dim var_n As Integer
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   If KeyAscii = 13 Then
      If Trim(txt_codigo_postal) <> "" Then
         rs.Open "select distinct vcha_pai_pais_id from tb_colonias where vcha_col_cp = '" + txt_codigo_postal + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
         Dim var_ren As Integer
         var_ren = rs.RecordCount
         rs.Close
         If var_ren > 1 Then
            lv_lista.ListItems.Clear
            rsaux.Open "select DISTINCT VCHA_PAI_PAIS_ID, VCHA_PAI_NOMBRE from vw_colonias order by vcha_pai_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
            While Not rsaux.EOF
                  Set list_item = lv_lista.ListItems.Add(, , rsaux!VCHA_PAI_PAIS_ID)
                  list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_pai_nombre), "", rsaux!vcha_pai_nombre)
                  rsaux.MoveNext
            Wend
            rsaux.Close
            lbl_lista = "SELECCIONE EL PAIS"
            var_tipo_lista = 16
             var_n = lv_lista.ListItems.Count
            If var_n > 6 Then
               lv_lista.ColumnHeaders(2).Width = 4270.71
            Else
               lv_lista.ColumnHeaders(2).Width = 4499.71
            End If
            frm_lista.Visible = True
            lv_lista.SetFocus
         Else
            rs.Open "select * from vw_colonias where vcha_col_cp = '" + txt_codigo_postal + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               If rs.RecordCount = 1 Then
                  txt_colonia = IIf(IsNull(rs!VCHA_COL_COLONIA_ID), "", rs!VCHA_COL_COLONIA_ID)
                  txt_nombre_colonia = IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
                  txt_pais = IIf(IsNull(rs!VCHA_PAI_PAIS_ID), "", rs!VCHA_PAI_PAIS_ID)
                  txt_nombre_pais = IIf(IsNull(rs!vcha_pai_nombre), "", rs!vcha_pai_nombre)
                  txt_estado = IIf(IsNull(rs!VCHA_EST_ESTADO_ID), "", rs!VCHA_EST_ESTADO_ID)
                  txt_nombre_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                  txt_municipio = IIf(IsNull(rs!VCHA_MUN_MUNICIPIO_ID), "", rs!VCHA_MUN_MUNICIPIO_ID)
                  txt_nombre_municipio = IIf(IsNull(rs!vcha_mun_nombre), "", rs!vcha_mun_nombre)
                  txt_ciudad = IIf(IsNull(rs!VCHA_CIU_CIUDAD_ID), "", rs!VCHA_CIU_CIUDAD_ID)
                  txt_nombre_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                  txt_telefono.SetFocus
               Else
                  lv_colonias.ListItems.Clear
                  While Not rs.EOF
                        Set list_item = lv_colonias.ListItems.Add(, , rs!VCHA_COL_COLONIA_ID)
                        list_item.SubItems(1) = IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
                        list_item.SubItems(2) = IIf(IsNull(rs!VCHA_PAI_PAIS_ID), "", rs!VCHA_PAI_PAIS_ID)
                        list_item.SubItems(3) = IIf(IsNull(rs!vcha_pai_nombre), "", rs!vcha_pai_nombre)
                        list_item.SubItems(4) = IIf(IsNull(rs!VCHA_EST_ESTADO_ID), "", rs!VCHA_EST_ESTADO_ID)
                        list_item.SubItems(5) = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                        list_item.SubItems(6) = IIf(IsNull(rs!VCHA_MUN_MUNICIPIO_ID), "", rs!VCHA_MUN_MUNICIPIO_ID)
                        list_item.SubItems(7) = IIf(IsNull(rs!vcha_mun_nombre), "", rs!vcha_mun_nombre)
                        list_item.SubItems(8) = IIf(IsNull(rs!VCHA_CIU_CIUDAD_ID), "", rs!VCHA_CIU_CIUDAD_ID)
                        list_item.SubItems(9) = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                        rs.MoveNext
                  Wend
                  lbl_colonias = "COLONIAS DEL C.P. " + txt_codigo_postal
                  var_n = lv_colonias.ListItems.Count
                  If var_n > 6 Then
                     lv_colonias.ColumnHeaders(2).Width = 4270.71
                  Else
                     lv_colonias.ColumnHeaders(2).Width = 4499.71
                  End If
                  frm_colonias.Visible = True
                  lv_colonias.SetFocus
               End If
            Else
               MsgBox "Código postal incorrecto", vbOKOnly, "ATENCION"
            End If
            rs.Close
         End If
      Else
         txt_telefono.SetFocus
      End If
   End If
End Sub

Private Sub txt_codigo_postal_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_colonia_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_colonia_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_colonias where vcha_pai_pais_id = '" + txt_pais + "' and vcha_est_estado_id = '" + txt_estado + "' order by vcha_col_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_COL_COLONIA_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "COLONIAS DE " + txt_nombre_estado
      var_tipo_lista = 14
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 7 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
   If KeyCode = 117 Then
      frmciudades.Show
   End If
End Sub

Private Sub txt_colonia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_domicilio_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_domicilio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_estado_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_estado_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_estados where vcha_pai_pais_id = '" + txt_pais + "' order by vcha_est_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_EST_ESTADO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ESTADOS DE " + txt_nombre_pais
      var_tipo_lista = 12
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 7 Then
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

Private Sub txt_estado_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_fecha_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_fecha_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fecha) Then
         frmcalendario.mes.Value = CDate(Me.txt_fecha)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      Me.txt_fecha = var_fecha_general
   End If
End Sub

Private Sub txt_grupo_real_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_grupo_real_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_gruposreales  WHERE vcha_emp_empresa_id = '" + var_empresa + "' order by vcha_gre_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_gre_grupo_real_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_GRE_NOMBRE), "", rs!VCHA_GRE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "GRUPOS REALES"
      var_tipo_lista = 15
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 7 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
   If KeyCode = 117 Then
      var_activa_forma_gruposreales = Me.Name
      Me.Enabled = False
      frmgruposreales.Show
   End If
End Sub

Private Sub txt_grupo_real_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_grupo_real_LostFocus()
   If Trim(txt_grupo_real) <> "" Then
      rs.Open "SELECT * FROM TB_GRUPOSREALES WHERE VCHA_GRE_GRUPO_rEAL_ID = '" + txt_grupo_real + "'  and vcha_emp_empresa_id = '" + var_empresa + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_grupo_actual = IIf(IsNull(rs!VCHA_GRE_NOMBRE), "", rs!VCHA_GRE_NOMBRE)
      Else
         MsgBox "Clave de grupo real incorrecto", vbOKOnly, "ATENCION"
         txt_grupo_real = ""
         txt_nombre_grupo_real = ""
      End If
      rs.Close
   Else
      txt_grupo_real = ""
      txt_nombre_grupo_real = ""
   End If
End Sub

Private Sub txt_limite_credito_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_limite_credito_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_municipio_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_municipio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_municipios where vcha_pai_pais_id = '" + txt_pais + "' and vcha_est_estado_id = '" + txt_estado + "' order by vcha_mun_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_MUN_MUNICIPIO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_mun_nombre), "", rs!vcha_mun_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "MUNICIPIOS DE" + txt_nombre_estado
      var_tipo_lista = 11
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 7 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
   If KeyCode = 117 Then
      frmmunicipios.Show
   End If
End Sub

Private Sub txt_municipio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_ciudad_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_ciudad_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_colonia_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_colonia_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_estado_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_estado_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_grupo_real_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_grupo_real_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_gruposreales order by vcha_gre_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_gre_grupo_real_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_GRE_NOMBRE), "", rs!VCHA_GRE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "GRUPOS REALES"
      var_tipo_lista = 15
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 7 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
   If KeyCode = 117 Then
      var_activa_forma_gruposreales = Me.Name
      Me.Enabled = False
      frmgruposreales.Show
   End If
End Sub

Private Sub txt_nombre_grupo_real_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
   If KeyAscii = 13 Then
      If Me.cmd_guardar.Enabled = True Then
         Me.cmd_guardar.SetFocus
      End If
   End If
End Sub

Private Sub txt_nombre_municipio_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_municipio_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_pais_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_pais_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_titular_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_titular_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_pais_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_pais_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_paises order by vcha_pai_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_PAI_PAIS_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_pai_nombre), "", rs!vcha_pai_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "PAISES"
      var_tipo_lista = 13
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 7 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
   If KeyCode = 117 Then
      var_catalogo_articulos = True
      frmpaises.Show
   End If
End Sub

Private Sub txt_pais_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_telefono_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_telefono_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_titular_anterior_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_titular_anterior_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_titular_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_titular_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub
