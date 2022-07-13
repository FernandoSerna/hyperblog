VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#11.0#0"; "VBSKFREE.OCX"
Begin VB.Form frmarticulos 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Articulos"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9930
   Icon            =   "frmarticulos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   9930
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   195
      Left            =   120
      TabIndex        =   54
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmd_buscar 
      Height          =   255
      Left            =   2880
      Picture         =   "frmarticulos.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   480
      Width           =   375
   End
   Begin vbskfree.Skinner Skinner1 
      Left            =   0
      Top             =   1800
      _ExtentX        =   1270
      _ExtentY        =   1270
      MaxButton       =   0
      MinButton       =   0
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3120
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmarticulos.frx":0E54
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmarticulos.frx":0FAE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra_almacen 
      Caption         =   "Productos Almacen"
      Height          =   5415
      Left            =   3360
      TabIndex        =   20
      Top             =   720
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CheckBox chk_prorratear 
         Caption         =   "Artuiculo Prorrateable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1560
         TabIndex        =   12
         Top             =   4800
         Width           =   2655
      End
      Begin VB.TextBox txt_articulos 
         Alignment       =   1  'Right Justify
         DataField       =   "cp"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   405
         Index           =   6
         Left            =   1560
         TabIndex        =   9
         Top             =   3720
         Width           =   2055
      End
      Begin VB.ComboBox cbo_almacen 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   1920
         Width           =   2415
      End
      Begin VB.ComboBox cbo_sublinea 
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   2640
         Width           =   1095
      End
      Begin VB.ComboBox cbo_linea 
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   2280
         Width           =   1095
      End
      Begin VB.ComboBox cbo_proveedor 
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txt_articulos 
         Alignment       =   1  'Right Justify
         DataField       =   "cp"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   405
         Index           =   8
         Left            =   3720
         TabIndex        =   11
         Top             =   4200
         Width           =   2055
      End
      Begin VB.TextBox txt_articulos 
         Alignment       =   1  'Right Justify
         DataField       =   "cp"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   405
         Index           =   7
         Left            =   1560
         TabIndex        =   10
         Top             =   4200
         Width           =   2055
      End
      Begin VB.TextBox txt_articulos 
         DataField       =   "cp"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   5
         Left            =   4080
         TabIndex        =   8
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txt_articulos 
         DataField       =   "nombre"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   2
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox txt_articulos 
         BackColor       =   &H00FFFFFF&
         DataField       =   "id"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox cbo_articulo 
         Height          =   315
         Index           =   2
         Left            =   1560
         TabIndex        =   7
         Top             =   3000
         Width           =   1095
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   5520
         TabIndex        =   51
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Un Registro Atras"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   90
         Index           =   3
         Left            =   240
         Picture         =   "frmarticulos.frx":12C8
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   5715
      End
      Begin VB.Image Image1 
         Height          =   90
         Index           =   1
         Left            =   240
         Picture         =   "frmarticulos.frx":1705
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   5715
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Almacen"
         Height          =   195
         Index           =   10
         Left            =   840
         TabIndex        =   53
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Existencia"
         Height          =   195
         Index           =   9
         Left            =   4320
         TabIndex        =   52
         Top             =   3960
         Width           =   720
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cto. Promedio"
         Height          =   195
         Index           =   8
         Left            =   495
         TabIndex        =   50
         Top             =   4320
         Width           =   990
      End
      Begin VB.Label lab_fecha_articulo 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4320
         TabIndex        =   30
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Alta"
         Height          =   195
         Left            =   3120
         TabIndex        =   29
         Top             =   480
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion"
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   28
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo del Articulo"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   1320
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sublinea"
         Height          =   195
         Index           =   2
         Left            =   840
         TabIndex        =   26
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Linea"
         Height          =   195
         Index           =   1
         Left            =   1065
         TabIndex        =   25
         Top             =   2280
         Width           =   390
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   0
         Left            =   735
         TabIndex        =   24
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ultimo Costo"
         Height          =   195
         Index           =   7
         Left            =   585
         TabIndex        =   23
         Top             =   3840
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicacion"
         Height          =   195
         Index           =   4
         Left            =   3240
         TabIndex        =   22
         Top             =   3000
         Width           =   720
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad"
         Height          =   195
         Index           =   2
         Left            =   960
         TabIndex        =   21
         Top             =   3120
         Width           =   510
      End
   End
   Begin VB.Frame fra_produccion 
      Caption         =   "Articulos de Produccion"
      Height          =   4935
      Left            =   3360
      TabIndex        =   31
      Top             =   840
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox txt_articulo2 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   405
         Index           =   8
         Left            =   2640
         TabIndex        =   49
         Top             =   4320
         Width           =   2295
      End
      Begin VB.TextBox txt_articulo2 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   405
         Index           =   7
         Left            =   2640
         TabIndex        =   48
         Top             =   3840
         Width           =   2295
      End
      Begin VB.TextBox txt_articulo2 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   405
         Index           =   6
         Left            =   2640
         TabIndex        =   47
         Top             =   3360
         Width           =   2295
      End
      Begin VB.TextBox txt_articulo2 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   405
         Index           =   5
         Left            =   2640
         TabIndex        =   46
         Top             =   2880
         Width           =   2295
      End
      Begin VB.TextBox txt_articulo2 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   405
         Index           =   4
         Left            =   2640
         TabIndex        =   45
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox txt_articulo2 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   405
         Index           =   3
         Left            =   2640
         TabIndex        =   44
         Top             =   1920
         Width           =   2295
      End
      Begin VB.TextBox txt_articulo2 
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   43
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txt_articulo2 
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   42
         Top             =   960
         Width           =   4455
      End
      Begin VB.TextBox txt_articulo2 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   32
         Top             =   600
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   90
         Index           =   0
         Left            =   480
         Picture         =   "frmarticulos.frx":1B42
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   5235
      End
      Begin VB.Label lab_articulos2 
         AutoSize        =   -1  'True
         Caption         =   "Clave del Articulo"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   41
         Top             =   600
         Width           =   1230
      End
      Begin VB.Label lab_articulos2 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   40
         Top             =   960
         Width           =   840
      End
      Begin VB.Label lab_articulos2 
         AutoSize        =   -1  'True
         Caption         =   "Talla"
         Height          =   195
         Index           =   2
         Left            =   960
         TabIndex        =   39
         Top             =   1320
         Width           =   345
      End
      Begin VB.Label lab_articulos2 
         AutoSize        =   -1  'True
         Caption         =   "Materia Prima"
         Height          =   195
         Index           =   3
         Left            =   1440
         TabIndex        =   38
         Top             =   2040
         Width           =   960
      End
      Begin VB.Label lab_articulos2 
         AutoSize        =   -1  'True
         Caption         =   "Avios"
         Height          =   195
         Index           =   4
         Left            =   2040
         TabIndex        =   37
         Top             =   2520
         Width           =   390
      End
      Begin VB.Label lab_articulos2 
         AutoSize        =   -1  'True
         Caption         =   "Mano de Obra"
         Height          =   195
         Index           =   5
         Left            =   1440
         TabIndex        =   36
         Top             =   3000
         Width           =   1020
      End
      Begin VB.Label lab_articulos2 
         AutoSize        =   -1  'True
         Caption         =   "Gastos de Fabricacion"
         Height          =   195
         Index           =   6
         Left            =   840
         TabIndex        =   35
         Top             =   3480
         Width           =   1590
      End
      Begin VB.Label lab_articulos2 
         AutoSize        =   -1  'True
         Caption         =   "Precio de Lista"
         Height          =   195
         Index           =   7
         Left            =   1440
         TabIndex        =   34
         Top             =   3960
         Width           =   1050
      End
      Begin VB.Label lab_articulos2 
         AutoSize        =   -1  'True
         Caption         =   "Precio  Costo"
         Height          =   195
         Index           =   8
         Left            =   1560
         TabIndex        =   33
         Top             =   4440
         Width           =   945
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   19
      Top             =   6000
      Width           =   375
   End
   Begin VB.TextBox txt_buscar 
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   2655
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   540
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   953
      ButtonWidth     =   1191
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo"
            Key             =   "a"
            Object.ToolTipText     =   "Agregar un Registro"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Guardar"
            Object.ToolTipText     =   "Guardar Registros"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Eliminar"
            Object.ToolTipText     =   "Eliminar Registros"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Object.ToolTipText     =   "Imprimir Lista"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir de Esta Forma"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lv_articulos 
      Height          =   4935
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   8705
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ARTICULO"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "PROVEEDOR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "LINEA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "SUBLINEA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "UBICACION"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ULTIMO COSTO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "CTO PROMEDIO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "UNIDAD"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "ARTICULO_ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "EXISTENCIA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "ALMACEN_ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "PRORRATEAR"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmarticulos.frx":1F7F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmarticulos.frx":2859
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmarticulos.frx":3133
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmarticulos.frx":3A0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmarticulos.frx":42E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmarticulos.frx":4BC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmarticulos.frx":549B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmarticulos.frx":57B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmarticulos.frx":5ACF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tool_atras_siguiente 
      Height          =   330
      Left            =   8520
      TabIndex        =   16
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Un Registro Atras"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Un Registro Adelante"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   600
      Top             =   5640
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
            Picture         =   "frmarticulos.frx":63A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmarticulos.frx":6C83
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   90
      Index           =   4
      Left            =   3360
      Picture         =   "frmarticulos.frx":755D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6075
   End
   Begin VB.Label txt_registros 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   1800
      TabIndex        =   18
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Registros"
      Height          =   195
      Index           =   3
      Left            =   225
      TabIndex        =   17
      Top             =   5760
      Width           =   1485
   End
   Begin VB.Image Image1 
      Height          =   90
      Index           =   2
      Left            =   3360
      Picture         =   "frmarticulos.frx":799A
      Stretch         =   -1  'True
      Top             =   720
      Width           =   6075
   End
End
Attribute VB_Name = "frmarticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public execute As Boolean
Public Mytextbox As TextBox
Dim var_hubo_cambios As Boolean
Dim var_prorratear As String

Private Sub cbo_almacen_Click()
    var_hubo_cambios = True
End Sub

Private Sub cbo_almacen_KeyPress(KeyAscii As Integer)
    Call pro_no_teclas(KeyAscii)
    If KeyAscii = 13 Then
        var_hubo_cambios = True
        Call pro_combodrop(cbo_linea, True)
    End If

End Sub

Private Sub cbo_articulo_Click(Index As Integer)
    var_hubo_cambios = True
End Sub

Private Sub cbo_articulo_KeyPress(Index As Integer, KeyAscii As Integer)
    Call pro_no_teclas(KeyAscii)
    If KeyAscii = 13 Then
        var_hubo_cambios = True
    End If

End Sub

Private Sub cbo_linea_Click()
    var_hubo_cambios = True
End Sub

Private Sub cbo_linea_KeyPress(KeyAscii As Integer)
    Call pro_no_teclas(KeyAscii)
    If KeyAscii = 13 Then
        var_hubo_cambios = True
        Call pro_combodrop(cbo_sublinea, True)
    End If

End Sub

Private Sub cbo_proveedor_Click()
    var_hubo_cambios = True
End Sub

Private Sub cbo_proveedor_KeyPress(KeyAscii As Integer)
    Call pro_no_teclas(KeyAscii)
    If KeyAscii = 13 Then
        var_hubo_cambios = True
        Call pro_combodrop(cbo_almacen, True)
    End If
End Sub

Private Sub cbo_sublinea_Click()
    var_hubo_cambios = True
End Sub

Private Sub cbo_sublinea_KeyPress(KeyAscii As Integer)
    Call pro_no_teclas(KeyAscii)
    If KeyAscii = 13 Then
        var_hubo_cambios = True
        Call pro_combodrop(cbo_articulo(2), True)
    End If

End Sub



Private Sub chk_prorratear_Click()
    If chk_prorratear.Value = 1 Then
        var_prorratear = "1"
    Else
        var_prorratear = ""
    End If
    var_hubo_cambios = True
End Sub

Private Sub cmdbuscar_Click()
End Sub

Private Sub cmd_buscar_Click()
    
  '  Call pro_busca_registro(lv_articulos, txt_buscar)
    txt_buscar = ""
    pro_textos

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Call pro_enfoque(KeyAscii)
End Sub


Public Function pro_no_teclas(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Function

Private Sub Form_Load()
    
    lab_fecha_articulo = Format(Date, "dd/mm/yyyy")
   lv_articulos.SmallIcons = ImageList1
        fra_almacen.Visible = True
        lv_articulos.Visible = True
        
        rs.CursorLocation = adUseClient
        rs.Open "SELECT * FROM TB_PROVEEDOR_VIEW", cnn, adOpenKeyset, adLockOptimistic, adCmdText
        Call RecsetToCombo(CBO_PROVEEDOR.hwnd, rs, 1)
        rs.Close
        
        rs.Open "SELECT * FROM TB_ALMACENES_VIEW", cnn, adOpenKeyset, adLockOptimistic, adCmdText
        Call RecsetToCombo(cbo_almacen.hwnd, rs, 1)
        rs.Close
        
        rs.Open "SELECT DISTINCT VCHA_LIN_LINEA FROM TB_lineas", cnn, adOpenKeyset, adLockOptimistic, adCmdText
        Call RecsetToCombo(cbo_linea.hwnd, rs, 0)
        rs.Close
        
        rs.Open "SELECT DISTINCT VCHA_LIN_SUBLINEA FROM TB_lineas_VIEW WHERE VCHA_LIN_SUBLINEA <>''", cnn, adOpenKeyset, adLockOptimistic, adCmdText
        Call RecsetToCombo(cbo_sublinea.hwnd, rs, 0)
        rs.Close
        
        rs.Open "SELECT * FROM TB_UNIDAD_VIEW", cnn, adOpenKeyset, adLockOptimistic, adCmdText
        Call RecsetToCombo(cbo_articulo(2).hwnd, rs, 1)
        rs.Close
        
        Call pro_encabezadosView(Me, lv_articulos, False)
        Call pro_llena_listview1_1
        Call pro_textos
        Call pro_AsignarAViewColor(lv_articulos, Picture1, vbWhite, vbGray)
        txt_registros = lv_articulos.ListItems.Count
    
        If txt_registros.caption <> 0 Then
            var_modifica_registro = True
        Else
            var_modifica_registro = False
        End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    var_modifica_registro = False
    Call menuvisible(Frmmenu2, True)

End Sub

Private Sub lv_articulos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    Call pro_ardena_listas(lv_articulos, ColumnHeader)

End Sub

    

Private Sub lv_articulos_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Set lv_articulos.SelectedItem = Item

        var_modifica_registro = True
        pro_textos
        
End Sub

Private Sub lv_articulos2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    Call pro_ardena_listas(lv_articulos2, ColumnHeader)

End Sub

Private Sub lv_articulos2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    Set lv_articulos2.SelectedItem = Item

        var_modifica_registro = True
        pro_carga_textos_2

End Sub





Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
execute = True
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
    execute = False
End If

End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
       
    lv_articulos.SetFocus
    Call pro_avanzar(Me, lv_articulos, Button)
    pro_textos
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
        chk_prorratear.Value = 0
        Call pro_limpiatextos(Me)
        txt_articulos(0).Enabled = True
        txt_articulos(0).SetFocus
        var_modifica_registro = False
    Case 2
        Call pro_guardar_articulos
    Case 3
        If rut_valida_textos_vacios_1 And var_modifica_registro Then
            Call pro_elimina_articulos
        End If
    Case 5
        Call gPrintListView(lv_articulos, "LISTADO DE ARTICULOS")
    Case 7
        Unload Me
    End Select
End Sub


Sub pro_guardar_articulos()

Dim ok As Boolean
Dim var_proveedor As String, var_almacen As String, var_linea As String, var_sublinea As String _
, var_unidad As String
Set TB_Articulos = New TB_Articulos
ok = True

    If rut_valida_textos_vacios_1 Then
        If var_hubo_cambios Then
            var_proveedor = Obtener_llave(cnn, rsaux, "TB_PROVEEDOR_VIEW", "vcha_pro_nombre", CBO_PROVEEDOR, 0, "T")
            var_almacen = Obtener_llave(cnn, rsaux, "TB_ALMACENES_VIEW", "vcha_alm_descripcion", cbo_almacen, 0, "T")
            var_unidad = Obtener_llave(cnn, rsaux, "TB_UNIDAD_VIEW", "vcha_uni_descripcion", cbo_articulo(2), 0, "T")
            
            ok = TB_Articulos.Anadir(txt_articulos(0), txt_articulos(1), var_proveedor, cbo_linea, cbo_sublinea _
            , txt_articulos(5), txt_articulos(6), txt_articulos(7) _
            , txt_articulos(8), var_unidad, "", 0, 0, 0, 0, 0, 0, "1", var_almacen, var_prorratear, "A", lab_fecha_articulo _
            , fun_NombreUsuario, fun_NombrePc, var_numero_planta)
            If ok Then
                pro_actualiza_ListView_1
                MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
                txt_registros = lv_articulos.ListItems.Count
                var_modifica_registro = True
            Else
                MsgBox "No se puede grabar registro: " + TB_Articulos.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
        End If
    End If
    
Set TB_Articulos = Nothing: var_hubo_cambios = False

End Sub



Sub pro_elimina_articulos()
Dim var_llave_usuarios As String

Set TB_Articulos = New TB_Articulos
    
ok = True
    
    rs.Open "select * from TB_DETALLE where VCHA_ART_ARTICULO_ID = '" & txt_articulos(0) & "'", cnn, adOpenForwardOnly, adLockOptimistic
    If rs.RecordCount = 0 Then
        If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
            ok = TB_Articulos.Eliminar(txt_articulos(0))
        Else
            GoTo SALIR:
        End If
        If ok Then
            MsgBox "Se Elimino Correctamente el Registro", vbInformation
            Call pro_limpiatextos(Me)
            lv_articulos.ListItems.Remove (lv_articulos.SelectedItem.Index)
            txt_registros = lv_articulos.ListItems.Count
            lv_articulos.SelectedItem.Selected = True
            pro_textos
        Else
            MsgBox "No se puede grabar registro: " + TB_Articulos.MensajeError, vbOKOnly + vbCritical, "ATENCION"
        End If
    Else
        SetTimer hwnd, NV_CLOSEMSGBOX, 1800, AddressOf TimerProc
        MsgBox "No se Puede Borrar Articulo ... Existen Movimientos", , "TRANSACCIONES [ AVISO ]"
    End If
    
SALIR:
Set Tb_usuarios = Nothing: rs.Close

End Sub

Sub pro_llena_listview1_1()

Dim list_item As ListItem

Set TB_Articulos = New TB_Articulos

    rs.Open "select * from TB_ARTICULOS_VIEW", cnn, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
        If rs(16).Value = 1 Then
            Set list_item = lv_articulos.ListItems.Add(, , rs(1).Value): list_item.SmallIcon = 9
            For i = 1 To 7
                list_item.SubItems(i) = IIf(IsNull(rs(i + 1).Value), "", rs(i + 1).Value)
            Next i
            list_item.SubItems(8) = IIf(IsNull(rs(0).Value), "", rs(0).Value)
            list_item.SubItems(9) = IIf(IsNull(rs(17).Value), "", rs(17).Value)
            list_item.SubItems(10) = IIf(IsNull(rs(18).Value), "", rs(18).Value)
            list_item.SubItems(11) = IIf(IsNull(rs(19).Value), "", rs(19).Value)
        End If
    rs.MoveNext:
    Wend
    rs.Close

Set TB_Articulos = Nothing

End Sub

Sub pro_llena_listview1_2()

Dim list_item As ListItem

    rs.Open "select * from TB_ARTICULOS_VIEW", cnn, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
        If rs(18).Value = 0 Then
            Set list_item = lv_articulos2.ListItems.Add(, , rs(1).Value): list_item.SmallIcon = 9
            For i = 11 To 17
                list_item.SubItems(i - 10) = IIf(IsNull(rs(i).Value), "", rs(i).Value)
            Next i
            'list_item.SubItems(7) = IIf(IsNull(rs(11).Value), "", rs(11).Value)
            list_item.SubItems(8) = IIf(IsNull(rs(0).Value), "", rs(0).Value)
        End If
    rs.MoveNext:
    Wend
    rs.Close

End Sub



Private Sub pro_actualiza_ListView_1()
Dim list_item As ListItem
Dim var_temp_almacen As String, var_temp_proveedor As String, var_temp_unidad As String

    If var_modifica_registro = False Then
        Set list_item = lv_articulos.ListItems.Add(, , txt_articulos(1)): list_item.SmallIcon = 9
        list_item.SubItems(8) = txt_articulos(0)
        list_item.SubItems(4) = txt_articulos(5)
        list_item.SubItems(5) = txt_articulos(6)
        list_item.SubItems(6) = txt_articulos(7)
        list_item.SubItems(9) = txt_articulos(8)
        list_item.SubItems(2) = cbo_linea
        list_item.SubItems(3) = cbo_sublinea
        list_item.SubItems(10) = Obtener_llave(cnn, rsaux, "TB_ALMACENES_VIEW", "vcha_alm_descripcion", cbo_almacen, 0, "T")
        list_item.SubItems(1) = Obtener_llave(cnn, rsaux, "TB_PROVEEDOR_VIEW", "vcha_pro_nombre", CBO_PROVEEDOR, 0, "T")
        list_item.SubItems(7) = Obtener_llave(cnn, rsaux, "TB_UNIDAD_VIEW", "vcha_uni_descripcion", cbo_articulo(2), 0, "T")
        If chk_prorratear.Value = 1 Then
            list_item.SubItems(11) = "1"
        Else
            list_item.SubItems(11) = ""
        End If
        list_item.EnsureVisible
        list_item.Selected = True
    Else
        lv_articulos.ListItems.Item(lv_articulos.SelectedItem.Index).ListSubItems(8) = txt_articulos(0)
        lv_articulos.ListItems.Item(lv_articulos.SelectedItem.Index).ListSubItems(4) = txt_articulos(5)
        lv_articulos.ListItems.Item(lv_articulos.SelectedItem.Index).ListSubItems(5) = txt_articulos(6)
        lv_articulos.ListItems.Item(lv_articulos.SelectedItem.Index).ListSubItems(6) = txt_articulos(7)
        lv_articulos.ListItems.Item(lv_articulos.SelectedItem.Index).ListSubItems(9) = txt_articulos(8)
        
        lv_articulos.ListItems.Item(lv_articulos.SelectedItem.Index).ListSubItems(1) = Obtener_llave(cnn, rsaux, "TB_PROVEEDOR_VIEW", "vcha_pro_nombre", CBO_PROVEEDOR, 0, "T")
        lv_articulos.ListItems.Item(lv_articulos.SelectedItem.Index).ListSubItems(10) = Obtener_llave(cnn, rsaux, "TB_almacenes_VIEW", "vcha_alm_descripcion", cbo_almacen, 0, "T")
        lv_articulos.ListItems.Item(lv_articulos.SelectedItem.Index).ListSubItems(7) = Obtener_llave(cnn, rsaux, "TB_UNIDAD_VIEW", "vcha_uni_descripcion", cbo_articulo(2), 0, "T")
        
        lv_articulos.ListItems.Item(lv_articulos.SelectedItem.Index).ListSubItems(2) = cbo_linea
        lv_articulos.ListItems.Item(lv_articulos.SelectedItem.Index).ListSubItems(3) = cbo_sublinea
   
        lv_articulos.ListItems.Item(lv_articulos.SelectedItem.Index) = txt_articulos(1)
        lv_articulos.ListItems.Item(lv_articulos.SelectedItem.Index).ListSubItems(8) = txt_articulos(0)
        If chk_prorratear.Value = 1 Then
            lv_articulos.ListItems.Item(lv_articulos.SelectedItem.Index).ListSubItems(11) = "1"
        Else
            lv_articulos.ListItems.Item(lv_articulos.SelectedItem.Index).ListSubItems(11) = ""
        End If
        lv_articulos.ListItems.Item(lv_articulos.SelectedItem.Index).EnsureVisible
        lv_articulos.ListItems.Item(lv_articulos.SelectedItem.Index).Selected = True
    End If
    lv_articulos.SetFocus

End Sub

Private Sub pro_actualiza_ListView_2()
Dim list_item As ListItem

    If var_modifica_registro = False Then
        Set list_item = lv_articulos2.ListItems.Add(, , txt_articulo2(1)): list_item.SmallIcon = 9
        list_item.SubItems(8) = txt_articulo2(0)
        'For i = 2 To 8
        '    list_item.SubItems(i - 1) = txt_articulo2(i)
        'Next i
    Else
        
        'For i = 2 To 8
        '    lv_articulos2.ListItems.Item(lv_articulos2.SelectedItem.Index).ListSubItems(i - 1) = txt_articulo2(i)
        'Next i
        lv_articulos2.ListItems.Item(lv_articulos2.SelectedItem.Index) = txt_articulo2(1)
        lv_articulos2.ListItems.Item(lv_articulos2.SelectedItem.Index).ListSubItems(8) = txt_articulo2(0)
    End If

End Sub


Sub pro_textos()
On Error Resume Next
        txt_articulos(1) = lv_articulos.SelectedItem
        txt_articulos(0) = lv_articulos.SelectedItem.SubItems(8)
        txt_articulos(5) = lv_articulos.SelectedItem.SubItems(4)
        txt_articulos(6) = lv_articulos.SelectedItem.SubItems(5)
        txt_articulos(7) = lv_articulos.SelectedItem.SubItems(6)
                

        txt_articulos(8) = lv_articulos.SelectedItem.SubItems(9)
        
        cbo_almacen = Obtener_llave(cnn, rsaux, "TB_ALMACENES_VIEW", "vcha_alm_almacen_id", lv_articulos.SelectedItem.SubItems(10), 1, "T")
        CBO_PROVEEDOR = Obtener_llave(cnn, rsaux, "TB_PROVEEDOR_VIEW", "vcha_pro_proveedor_id", lv_articulos.SelectedItem.SubItems(1), 1, "T")
        cbo_linea = lv_articulos.SelectedItem.SubItems(2)
        cbo_sublinea = Format(lv_articulos.SelectedItem.SubItems(3), "00")
        cbo_articulo(2) = Obtener_llave(cnn, rsaux, "TB_UNIDAD_VIEW", "vcha_uni_unidad_id", lv_articulos.SelectedItem.SubItems(7), 1, "T")
        
        If Trim(lv_articulos.SelectedItem.SubItems(11)) = "1" Then
            chk_prorratear.Value = 1
        Else
            chk_prorratear.Value = 0
        End If

End Sub


Sub pro_carga_textos_2()
On Error Resume Next
        txt_articulo2(0) = lv_articulos2.SelectedItem.SubItems(8)
        txt_articulo2(1) = lv_articulos2.SelectedItem
        For i = 2 To 8
            txt_articulo2(i) = lv_articulos2.SelectedItem.SubItems(i - 1)
        Next i
        
End Sub

Function rut_valida_textos_vacios_1() As Boolean
Dim i As Byte

    For i = 0 To 1
        If txt_articulos(i) = "" Then
            MsgBox "El Campo " & Label1(i).caption & " No debe estar Vacio", vbCritical + vbOKOnly, "AVISO"
            txt_articulos(i).SetFocus
            Exit Function
            rut_valida_textos_vacios = False
        End If
    Next i
    
    If cbo_articulo(2) = "" Then
        MsgBox "El Campo " & Label1(2) & " No debe estar Vacio", vbCritical + vbOKOnly, "AVISO"
        cbo_articulo(2).SetFocus
        Exit Function
        rut_valida_textos_vacios_1 = False
    End If
    
    rut_valida_textos_vacios_1 = True
    
End Function


Function rut_valida_textos_vacios_2() As Boolean
Dim i As Byte

    For i = 0 To 8
        If txt_articulo2(i) = "" Then
            MsgBox "El Campo " & lab_articulos2(i).caption & " No debe estar Vacio", vbCritical + vbOKOnly, "AVISO"
            txt_articulo2(i).SetFocus
            Exit Function
            rut_valida_textos_vacios_2 = False
        End If
    Next i
    rut_valida_textos_vacios_2 = True
    
End Function



Public Sub Autocomplete(Lvw As ListView, sFind, Mytextbox As TextBox)
Dim Lvfindtm As ListItem
Dim TempSelStart As Integer
Dim strTemp As String

    Set Lvfindtm = Lvw.FindItem(sFind, lvwText, , lvwPartial)
    If Not Lvfindtm Is Nothing Then
        Lvfindtm.EnsureVisible
        Lvfindtm.Selected = True
    
    If execute Then
        TempSelStart = Mytextbox.SelStart
        Mytextbox.Text = CStr(Lvfindtm)
    If Not Mytextbox.Text = "" Then
        Mytextbox.SelStart = TempSelStart
        Mytextbox.SelLength = Len(Mytextbox.Text) - TempSelStart
    End If
        End If
            End If
End Sub

Private Sub txt_articulo2_GotFocus(Index As Integer)

    txt_articulo2(Index).BackColor = &HC0FFC0
    txt_articulo2(Index).SelStart = 0
    txt_articulo2(Index).SelLength = Len(txt_articulo2(Index).Text)

End Sub

Private Sub txt_articulo2_LostFocus(Index As Integer)
    txt_articulo2(Index).BackColor = &H80000005
End Sub


Private Sub txt_articulos_GotFocus(Index As Integer)
    
    txt_articulos(Index).BackColor = &HC0FFC0
    txt_articulos(Index).SelStart = 0
    txt_articulos(Index).SelLength = Len(txt_articulos(Index).Text)

End Sub

Private Sub txt_articulos_KeyPress(Index As Integer, KeyAscii As Integer)
    var_hubo_cambios = True
    Select Case Index
    Case 1
        If KeyAscii = 13 Then
            Call pro_combodrop(CBO_PROVEEDOR, True)
        End If
    Case 6, 7
        Call pro_valida_numeros(KeyAscii)
    Case 8, 9, 10
        Call pro_valida_numeros(KeyAscii)
    End Select
    
End Sub

Private Sub txt_articulos_LostFocus(Index As Integer)
    
    txt_articulos(Index).BackColor = &H80000005

End Sub

