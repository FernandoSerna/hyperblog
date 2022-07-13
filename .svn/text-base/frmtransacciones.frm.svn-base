VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#11.0#0"; "VBSKFREE.OCX"
Begin VB.Form frmtransacciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRANSACCIONES"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10080
   Icon            =   "frmtransacciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   10080
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fra_buscar 
      Caption         =   "Consultar Art."
      Height          =   4335
      Left            =   360
      TabIndex        =   43
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
      Begin VB.ListBox lstbox 
         BackColor       =   &H80000001&
         ForeColor       =   &H00FFFFFF&
         Height          =   3570
         ItemData        =   "frmtransacciones.frx":08CA
         Left            =   35
         List            =   "frmtransacciones.frx":08CC
         Sorted          =   -1  'True
         TabIndex        =   45
         Top             =   600
         Width           =   1140
      End
      Begin VB.TextBox txtbox 
         Height          =   285
         Left            =   35
         TabIndex        =   44
         Top             =   240
         Width           =   1140
      End
   End
   Begin MSComCtl2.MonthView mon_transacciones 
      Height          =   2370
      Left            =   6360
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   0
      BackColor       =   -2147483646
      Appearance      =   1
      StartOfWeek     =   50266113
      TitleBackColor  =   12632256
      CurrentDate     =   37466
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   7560
      TabIndex        =   30
      Top             =   840
      Width           =   1815
      Begin VB.Label lab_folio_transacciones 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   495
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Folio"
         Height          =   195
         Left            =   720
         TabIndex        =   31
         Top             =   360
         Width           =   330
      End
   End
   Begin vbskfree.Skinner Skinner1 
      Left            =   1320
      Top             =   7200
      _ExtentX        =   1270
      _ExtentY        =   1270
      MaxButton       =   0
      MinButton       =   0
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
   End
   Begin MSMask.MaskEdBox txt_movimientos 
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   10
      Top             =   3240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.TextBox txt_clave 
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   360
      Width           =   735
   End
   Begin VB.Frame fra_notas_entrada 
      Enabled         =   0   'False
      Height          =   1335
      Left            =   480
      TabIndex        =   8
      Top             =   720
      Width           =   4935
      Begin VB.ComboBox CBO_3 
         Height          =   315
         Left            =   2400
         TabIndex        =   41
         Top             =   960
         Width           =   2415
      End
      Begin VB.ComboBox CBO_2 
         Height          =   315
         Left            =   2400
         TabIndex        =   38
         Top             =   600
         Width           =   2415
      End
      Begin VB.ComboBox CBO_1 
         Height          =   315
         Left            =   2400
         TabIndex        =   37
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lab_titulo1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rerencia Extra"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   40
         Top             =   960
         Width           =   2235
      End
      Begin VB.Label lab_titulo1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   2235
      End
      Begin VB.Label lab_titulo1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   2235
      End
   End
   Begin VB.TextBox txt_fecha_transaccion 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   7680
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   1920
      ScaleHeight     =   105
      ScaleWidth      =   345
      TabIndex        =   6
      Top             =   6960
      Width           =   375
   End
   Begin MSComctlLib.ListView lv_transacciones 
      Height          =   3855
      Left            =   480
      TabIndex        =   4
      Top             =   3480
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   6800
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "x1"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "x2"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "x3"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "x4"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "x5"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "x6"
         Object.Width           =   159
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "x7"
         Object.Width           =   1288
      EndProperty
   End
   Begin VB.ComboBox cbo_tipo_movimiento 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   4335
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   6960
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
            Picture         =   "frmtransacciones.frx":08CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtransacciones.frx":11A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtransacciones.frx":1A82
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtransacciones.frx":235C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtransacciones.frx":2C36
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtransacciones.frx":3510
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtransacciones.frx":3DEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtransacciones.frx":4104
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtransacciones.frx":441E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtransacciones.frx":4738
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtransacciones.frx":4CD2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   1080
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   12
      ImageHeight     =   12
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtransacciones.frx":526C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtransacciones.frx":5B46
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtransacciones.frx":6420
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   540
      Left            =   480
      TabIndex        =   3
      Top             =   2280
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   953
      ButtonWidth     =   1217
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo"
            Object.ToolTipText     =   "Nuevo Registro"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Guardar"
            Object.ToolTipText     =   "Guardar Registro"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Eliminar"
            Object.ToolTipText     =   "Eliminar Registro"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Object.ToolTipText     =   "Salir de Esta Forma"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox txt_movimientos 
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   11
      Top             =   3240
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txt_movimientos 
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   12
      Top             =   3240
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   450
      _Version        =   393216
      Format          =   "0.0000"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txt_movimientos 
      Height          =   255
      Index           =   3
      Left            =   6555
      TabIndex        =   13
      Top             =   3240
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   450
      _Version        =   393216
      Format          =   "0.0000"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txt_movimientos 
      Height          =   255
      Index           =   4
      Left            =   7560
      TabIndex        =   19
      Top             =   3240
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txt_movimientos 
      Height          =   255
      Index           =   5
      Left            =   8640
      TabIndex        =   25
      Top             =   3240
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0.0000"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Height          =   330
      Left            =   120
      TabIndex        =   46
      Top             =   3240
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   7080
      TabIndex        =   42
      Top             =   7320
      Width           =   375
   End
   Begin VB.Label lab_almacen_id 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   5040
      TabIndex        =   39
      Top             =   2280
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   90
      Index           =   2
      Left            =   8640
      Picture         =   "frmtransacciones.frx":673A
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   795
   End
   Begin VB.Image Image1 
      Height          =   90
      Index           =   1
      Left            =   480
      Picture         =   "frmtransacciones.frx":6B77
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   7995
   End
   Begin VB.Image Image1 
      Height          =   90
      Index           =   0
      Left            =   480
      Picture         =   "frmtransacciones.frx":6FB4
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   8955
   End
   Begin VB.Label lab_almacen 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   5040
      TabIndex        =   36
      Top             =   2520
      Width           =   4335
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Almacen"
      Height          =   195
      Left            =   4320
      TabIndex        =   35
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lab_existencia 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   735
      Left            =   5520
      TabIndex        =   34
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Existencia"
      Height          =   195
      Index           =   5
      Left            =   6120
      TabIndex        =   33
      Top             =   720
      Width           =   720
   End
   Begin VB.Label lab_afectacion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   5880
      TabIndex        =   29
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Afectacion"
      Height          =   195
      Index           =   3
      Left            =   6240
      TabIndex        =   28
      Top             =   240
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      Height          =   195
      Index           =   2
      Left            =   8160
      TabIndex        =   27
      Top             =   240
      Width           =   450
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Unidad"
      Height          =   255
      Index           =   6
      Left            =   8640
      TabIndex        =   26
      Top             =   3000
      Width           =   765
   End
   Begin VB.Label Label3 
      Caption         =   "TOTALES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   22
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   7560
      TabIndex        =   21
      Top             =   7440
      Width           =   1005
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   1
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   5520
      TabIndex        =   20
      Top             =   7440
      Width           =   1050
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Importe"
      Height          =   255
      Index           =   4
      Left            =   7560
      TabIndex        =   18
      Top             =   3000
      Width           =   1005
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Precio"
      Height          =   255
      Index           =   3
      Left            =   6555
      TabIndex        =   17
      Top             =   3000
      Width           =   1005
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cantidad"
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   16
      Top             =   3000
      Width           =   1050
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripcion"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   15
      Top             =   3000
      Width           =   3975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clave"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   14
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Movimiento"
      Height          =   195
      Index           =   1
      Left            =   1320
      TabIndex        =   7
      Top             =   120
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clave"
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   5
      Top             =   120
      Width           =   405
   End
End
Attribute VB_Name = "frmtransacciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_costo_promedio As Double
Dim var_afecta As String, var_tabla As String, var_campo As String
Dim var_titulo1 As String, var_titulo2 As String, var_titcaption1 As String, var_titcaption2 As String
Dim var_tabla1 As String, var_tabla2 As String, var_campo1 As String, var_campo2 As String
Dim var_temp_folio As Long
Dim var_almacen_id As String
Dim lab_folio As String

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd _
As Long, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Const VB_SEARCHSTR = &H18F














Private Sub CBO_1_GotFocus()
    CBO_1.BackColor = &HC0FFC0
    CBO_1.SelStart = 0
    CBO_1.SelLength = Len(CBO_1.Text)

End Sub

Private Sub CBO_1_KeyPress(KeyAscii As Integer)
    If CBO_1.ListCount > 0 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
'    Call pro_valida_numeros(KeyAscii)
    If KeyAscii = 13 Then
        If CBO_1 <> "" And CBO_2.Visible = True Then
            If var_tabla2 <> "" Then
                Call pro_combodrop(CBO_2, True)
            End If
            CBO_2.SetFocus
        Else
            CBO_3.SetFocus
        End If
    End If
End Sub

Private Sub CBO_1_LostFocus()

    CBO_1.BackColor = &H80000005
    
End Sub

Private Sub CBO_2_GotFocus()
    CBO_2.BackColor = &HC0FFC0
    CBO_2.SelStart = 0
    CBO_2.SelLength = Len(CBO_2.Text)
End Sub

Private Sub CBO_2_KeyPress(KeyAscii As Integer)
    If CBO_2.ListCount > 0 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
    
'    Call pro_valida_numeros(KeyAscii)
    If KeyAscii = 13 And CBO_2 <> "" Then
        CBO_3.SetFocus
    End If

End Sub

Private Sub CBO_2_LostFocus()

    CBO_2.BackColor = &H80000005
    
End Sub

Private Sub CBO_3_GotFocus()
    CBO_3.BackColor = &HC0FFC0
    CBO_3.SelStart = 0
    CBO_3.SelLength = Len(CBO_3.Text)
End Sub

Private Sub CBO_3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txt_movimientos(0).SetFocus
    End If
End Sub

Private Sub CBO_3_LostFocus()

    CBO_3.BackColor = &H80000005
    
End Sub

Private Sub cbo_tipo_movimiento_Click()

On Error Resume Next
    
    txt_clave = Obtener_llave(cnn, rsaux, "tb_movimientos_view", "VCHA_MOV_DESCRIPCION", cbo_tipo_movimiento, 0, "T")
    
    var_afecta = Obtener_llave(cnn, rsaux, "tb_movimientos_view", "VCHA_MOV_DESCRIPCION", cbo_tipo_movimiento, 2, "T")
    
    lab_titulo1(0) = Obtener_llave(cnn, rsaux, "tb_movimientos_view", "VCHA_MOV_DESCRIPCION", cbo_tipo_movimiento, 4, "T")
    
    lab_folio_transacciones = Obtener_llave(cnn, rsaux, "tb_FOLIOS_view", "VCHA_MOV_MOVIMIENTO_ID", txt_clave, 1, "T")
    
    lab_titulo1(1) = Obtener_llave(cnn, rsaux, "tb_movimientos_view", "VCHA_MOV_DESCRIPCION", cbo_tipo_movimiento, 5, "T")
    
    If lab_titulo1(1) = "" Then
        lab_titulo1(1).Visible = False: CBO_2.Visible = False
    Else
        lab_titulo1(1).Visible = True: CBO_2.Visible = True
    End If
    
    var_tabla1 = Obtener_llave(cnn, rsaux, "tb_movimientos_view", "VCHA_MOV_DESCRIPCION", cbo_tipo_movimiento, 6, "T")
    var_tabla2 = Obtener_llave(cnn, rsaux, "tb_movimientos_view", "VCHA_MOV_DESCRIPCION", cbo_tipo_movimiento, 8, "T")
    
    If var_tabla1 <> "" Then
        If var_tabla1 = "TB_PLANTAS" Or var_tabla1 = "TB_ALMACENES" Then
            rs.Open "SELECT * FROM " & var_tabla1 & " where BINT_PLA_PLANTA_ID <> " & var_numero_planta, cnn, adOpenKeyset, adLockOptimistic, adCmdText
            Call RecsetToCombo(CBO_1.hwnd, rs, 1)
            rs.Close
        Else
            rs.Open "SELECT * FROM " & var_tabla1, cnn, adOpenKeyset, adLockOptimistic, adCmdText
            Call RecsetToCombo(CBO_1.hwnd, rs, 1)
            rs.Close
        End If
    Else
        CBO_1.Clear
    End If
        
    If var_tabla2 <> "" Then
        If var_tabla2 = "TB_PLANTAS" Or var_tabla1 = "TB_ALMACENES" Then
            rs.Open "SELECT * FROM " & var_tabla2 & " where BINT_PLA_PLANTA_ID <> " & var_numero_planta, cnn, adOpenKeyset, adLockOptimistic, adCmdText
            Call RecsetToCombo(CBO_2.hwnd, rs, 1)
            rs.Close
        Else
            rs.Open "SELECT * FROM " & var_tabla2, cnn, adOpenKeyset, adLockOptimistic, adCmdText
            Call RecsetToCombo(CBO_2.hwnd, rs, 1)
            rs.Close
        End If
    Else
        CBO_2.Clear
    End If
    
    lab_afectacion = var_afecta
    
    fra_notas_entrada.Enabled = True
    
    Call pro_limpiatextos2(Me)
    For i = 0 To 4
        txt_movimientos(i) = ""
    Next i
    
    
End Sub

Private Sub cbo_tipo_movimiento_GotFocus()
    
    cbo_tipo_movimiento.BackColor = &HC0FFC0
    cbo_tipo_movimiento.SelStart = 0
    cbo_tipo_movimiento.SelLength = Len(cbo_tipo_movimiento.Text)
    
    
    
    'Call pro_oculta_frames(Me)
    Call pro_limpiatextos2(Me)
   ' Call pro_combodrop(cbo_tipo_movimiento, True)
End Sub


Private Sub cbo_tipo_movimiento_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii <> 13 Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 And cbo_tipo_movimiento <> "" And var_usuario_global = "administrador" Then
        txt_fecha_transaccion.Enabled = True
        txt_fecha_transaccion.SetFocus
    Else
        If var_tabla1 <> "" Then
            Call pro_combodrop(CBO_1, True)
        End If
        CBO_1.SetFocus
    End If

End Sub

Private Sub cbo_tipo_movimiento_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        Call pro_combodrop(cbo_tipo_movimiento, True)
    End If
End Sub

Private Sub cbo_tipo_movimiento_LostFocus()
    cbo_tipo_movimiento.BackColor = &H80000005
    If cbo_tipo_movimiento = "" Then cbo_tipo_movimiento.SetFocus
   
End Sub


Private Sub Form_Load()


    frmtransacciones.caption = frmtransacciones.caption & " " & var_nombre_planta
    txt_fecha_transaccion = Format(Date, "dd/mm/yyyy")
    'lab_folio_transacciones = Siguiente(cnn, rsaux, "TB_TRANSACCIONES_VIEW", "BINT_TRA_TRANSACCIONES_ID")
    mon_transacciones.Value = Date
    Call pro_encabezadosView(Me, lv_transacciones, False)
    
    Call pro_AsignarAViewColor(lv_transacciones, Picture1, vbWhite, vbGray)
    
    rs.Open "SELECT * FROM TB_MOVIMIENTOS_VIEW where BINT_PLA_PLANTA_ID = " & var_numero_planta, cnn, adOpenDynamic, adLockOptimistic
    Call RecsetToCombo(cbo_tipo_movimiento.hwnd, rs, 1)
    rs.Close


    lv_transacciones.SmallIcons = ImageList1
End Sub


'Public Sub pro_Vacia_formulas(listado As Crystal.CrystalReport, iNumero As Integer)

   
'    Dim tiForm As Integer
    
'    For tiForm = 0 To iNumero
'        listado.Formulas(tiForm) = ""
'    Next tiForm
    
'    For tiForm = 0 To 10
'        listado.SortFields(tiForm) = ""
'    Next tiForm

'End Sub



Private Sub Form_Unload(Cancel As Integer)
    var_modifica_registro = False
    Call menuvisible(Frmmenu2, True)
    'rs.Close
End Sub

Private Sub lab_titcaption1_Change()
    If lab_titcaption1 <> "" Then
        Sh_titulo1.Visible = True
    Else
        Sh_titulo1.Visible = False
    End If
    Sh_titulo1.Top = lab_titcaption1.Top
    Sh_titulo1.Left = lab_titcaption1.Left - 50
    Sh_titulo1.Width = lab_titcaption1.Width + 100
End Sub

Private Sub lab_titcaption2_Change()
    If lab_titcaption2 <> "" Then
        Sh_titulo1.Visible = True
    Else
        Sh_titulo1.Visible = False
    End If
    Sh_titulo1.Top = lab_titcaption2.Top
    Sh_titulo1.Left = lab_titcaption2.Left - 50
    Sh_titulo1.Width = lab_titcaption2.Width + 100

End Sub

Private Sub lstbox_Click()
        txt_movimientos(0) = lstbox.Text

End Sub

Private Sub lstbox_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 And txt_movimientos(0) <> "" Then
        fra_buscar.Visible = False
        txt_movimientos(2).SetFocus
        txtbox = ""
    End If
    If KeyAscii = 13 Then
        fra_buscar.Visible = False
        txt_movimientos(2).SetFocus
        txtbox = ""
    End If

End Sub

Private Sub lv_transacciones_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'Dim e As Boolean, ite As Long
    'ite = CLng(Item)
'    e = ListView_DeleteItem(lv_transacciones.hwnd, Item)



End Sub




Private Sub mon_transacciones_Click()
    
    CBO_1.SetFocus
    txt_fecha_transaccion = Format(mon_transacciones.Value, "dd/mm/YYYy")
    mon_transacciones.Visible = False
    Call pro_combodrop(CBO_1, True)


End Sub

Private Sub mon_transacciones_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        txt_fecha_transaccion.SetFocus
        mon_transacciones.Visible = False
    End If
    If KeyAscii = 13 Then
        mon_transacciones_Click
    End If

End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)

    Call pro_valida_numeros(KeyAscii)

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
    Select Case Button.Index
    Case 1
        pro_limpia_todo
        var_modifica_registro = False
        Call pro_limpiatextos2(Me)
        For i = 0 To 4
            txt_movimientos(i) = ""
        Next i
        
    Case 2
           Call pro_guardar_articulos
    Case 4
        Unload Me
        'If rut_valida_textos_vacios_1 And var_modifica_registro Then
        'Call pro_elimina_articulos
        'End If
    Case 6
        
    End Select
End Sub

Sub pro_limpia_todo()
    
    lab_folio_transacciones = Siguiente(cnn, rsaux, "TB_TRANSACCIONES_VIEW", "BINT_TRA_TRANSACCIONES_ID")
    Call pro_limpiatextos(Me)
    pro_limpia_lista
    lab_titulo1(0) = "": lab_titulo1(1) = ""
    cbo_tipo_movimiento.SetFocus
    var_modifica_registro = False
    txt_fecha_transaccion = Format(Date, "dd/mm/yyyy")
    mon_transacciones.Visible = False
    lab_afectacion = ""
    Label2(5) = 0: Label2(7) = 0

End Sub


Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
        fra_buscar.Visible = True
        If lstbox.ListCount = 0 Then
            rs.Open "SELECT VCHA_ART_ARTICULO_ID FROM TB_ARTICULOS_VIEW ORDER BY VCHA_ART_ARTICULO_ID", cnn, adOpenDynamic, adLockOptimistic
            If rs.RecordCount <> 0 Then
                While Not rs.EOF
                    lstbox.AddItem (rs(0).Value)
                    rs.MoveNext
                Wend
                rs.Close
            End If
            txtbox.SetFocus
        End If
    End Select

End Sub

Private Sub txt_clave_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cbo_tipo_movimiento = Obtener_llave(cnn, rsaux, "tb_movimientos_view", "bint_mov_movimiento_id", txt_clave, 1, "N")
    End If

End Sub



Private Sub txt_fecha_transaccion_GotFocus()
    
    txt_fecha_transaccion.BackColor = &HC0FFC0
    txt_fecha_transaccion.SelStart = 0
    txt_fecha_transaccion.SelLength = Len(txt_fecha_transaccion.Text)

End Sub

Private Sub txt_fecha_transaccion_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        mon_transacciones.Visible = True
        mon_transacciones.SetFocus
    End If

End Sub

Private Sub txt_fecha_transaccion_LostFocus()
    
    txt_fecha_transaccion.BackColor = &H80000005

End Sub

Private Sub txt_movimientos_Change(Index As Integer)
Dim var_auxiliar_unidad As String
    Select Case Index
    Case 0
        txt_movimientos(1) = Obtener_llave(cnn, rsaux, "TB_ARTICULOS_view", "vcha_art_articulo_id", txt_movimientos(0), 1, "T")

        If var_afecta = "SUMA" Then
            txt_movimientos(3) = Obtener_llave(cnn, rsaux, "TB_ARTICULOS_view", "vcha_art_articulo_id", txt_movimientos(0), 6, "T")
        Else
            txt_movimientos(3) = Obtener_llave(cnn, rsaux, "TB_ARTICULOS_view", "vcha_art_articulo_id", txt_movimientos(0), 7, "T")
        End If
        var_auxiliar_unidad = Obtener_llave(cnn, rsaux, "TB_ARTICULOS_view", "vcha_art_articulo_id", txt_movimientos(0), 8, "T")
        txt_movimientos(5) = Obtener_llave(cnn, rsaux, "tb_unidad_view", "VCHA_UNI_UNIDAD_ID", var_auxiliar_unidad, 1, "T")
        lab_existencia = Obtener_llave(cnn, rsaux, "TB_ARTICULOS_view", "vcha_art_articulo_id", txt_movimientos(0), 17, "T") + " " + txt_movimientos(5)
        
        lab_almacen_id = Obtener_llave(cnn, rsaux, "TB_ARTICULOS_view", "vcha_art_articulo_id", txt_movimientos(0), 18, "T")
        lab_almacen = Obtener_llave(cnn, rsaux, "TB_ALMACENES_VIEW", "vcha_alm_almacen_id", lab_almacen_id, 1, "T")
    Case 2
        If Val(txt_movimientos(2)) > Val(lab_existencia) And lab_afectacion = "RESTA" Then
            SetTimer hwnd, NV_CLOSEMSGBOX, 1600, AddressOf TimerProc
            MsgBox "Cantidad Mayor a la Existencia", vbCritical, "TRANSACCIONES [ AVISO ]"
            txt_movimientos(2) = "": txt_movimientos(2).SetFocus
        End If
    End Select
End Sub

Private Sub txt_movimientos_GotFocus(Index As Integer)

    txt_movimientos(Index).BackColor = &HC0FFC0
    txt_movimientos(Index).SelStart = 0
    txt_movimientos(Index).SelLength = Len(txt_movimientos(Index).Text)

End Sub

Private Sub txt_movimientos_LostFocus(Index As Integer)
    txt_movimientos(Index).BackColor = &H80000005
End Sub


Private Sub txt_movimientos_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Select Case Index
    Case 0
        If KeyAscii = 13 And txt_movimientos(0) <> "" And txt_movimientos(1) <> "" Then
            txt_movimientos(2).SetFocus
        End If
    Case 1
    Case 2
        Call pro_valida_numeros(KeyAscii)
        If KeyAscii = 13 And txt_movimientos(2) <> "" Then
            If txt_movimientos(2) <= 0 Then Exit Sub
            txt_movimientos(4) = Val(txt_movimientos(2)) * Val(txt_movimientos(3))
            Label2(5) = Val(Label2(5)) + Val(txt_movimientos(2))
            txt_movimientos(3).SetFocus
        End If
    Case 3
        If lab_afectacion = "RESTA" And KeyAscii <> 13 Then
            KeyAscii = 0
        End If
        Call pro_valida_numeros(KeyAscii)
        If KeyAscii = 13 And txt_movimientos(Index) <> "" Then
            If txt_movimientos(Index) <= 0 And txt_movimientos(0) <> "VA04001" Then Exit Sub
            txt_movimientos(4) = Val(txt_movimientos(2)) * Val(txt_movimientos(3))
            pro_traspasa_datos
           
            Label2(7) = Val(Label2(7)) + Val(txt_movimientos(4))
            For i = 0 To 4
                txt_movimientos(i) = ""
            Next i
            txt_movimientos(0).SetFocus
            Toolbar1.Buttons.Item(2).Enabled = True
        End If
    End Select
    
End Sub


Sub pro_traspasa_datos()

Dim list_item As ListItem
    
    Set list_item = lv_transacciones.ListItems.Add(, , txt_movimientos(0)): list_item.SmallIcon = 11
        list_item.SubItems(1) = txt_movimientos(1)
        list_item.SubItems(2) = txt_movimientos(2)
        list_item.SubItems(3) = txt_movimientos(3)
        list_item.SubItems(4) = Format(txt_movimientos(4), "### ### ##0.00")
        list_item.SubItems(6) = txt_movimientos(5)
    Set list_item = Nothing

End Sub




Sub pro_guardar_articulos()


Set TB_TRANSACCIONES = New TB_TRANSACCIONES
ok = True
    If txt_clave <> "" And cbo_tipo_movimiento <> "" And CBO_1 <> "" Then
        lab_folio = Siguiente(cnn, rsaux, "TB_TRANSACCIONES", "BINT_TRA_TRANSACCIONES_ID")
        ok = TB_TRANSACCIONES.Anadir(lab_folio, txt_clave, CBO_1 _
        , CBO_2, lab_almacen_id, CBO_3, "A", txt_fecha_transaccion, fun_NombreUsuario, fun_NombrePc, var_numero_planta)
        If ok Then
            rs.Open "SELECT * FROM TB_FOLIOS_VIEW WHERE VCHA_MOV_MOVIMIENTO_ID = '" & txt_clave & "'", cnn, adOpenKeyset, adLockOptimistic, adCmdText
            If rs.RecordCount <> 0 Then
                rs(1).Value = rs(1).Value + 1
                rs.Update
            End If
            rs.Close
            pro_guarda_detalle
        Else
            MsgBox "No se puede grabar registro: " + TB_TRANSACCIONES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
        End If
    Else
        MsgBox "Verifica que los Datos Esten Completos !", , "TRANSACCIONES [ AVISO ]"
    End If
    
Set TB_TRANSACCIONES = Nothing


End Sub



Public Sub pro_guarda_detalle()

Dim ok As Boolean, var_unidad As String, txt_consecutico_deTALLE As String

Set TB_DETALLE = New TB_DETALLE
ok = True

        For i = 1 To lv_transacciones.ListItems.Count
            Call pro_costeo_entredas_salidas(lv_transacciones.ListItems.Item(i) _
            , lv_transacciones.ListItems.Item(i).SubItems(2) _
            , lv_transacciones.ListItems.Item(i).SubItems(4) _
            , lv_transacciones.ListItems.Item(i).SubItems(3))
            txt_consecutico_deTALLE = Siguiente(cnn, rsaux, "TB_DETALLE_VIEW", "BINT_DET_DETALLE_ID")
            ok = TB_DETALLE.Anadir(txt_consecutico_deTALLE, lv_transacciones.ListItems.Item(i) _
            , lv_transacciones.ListItems.Item(i).SubItems(2), lv_transacciones.ListItems.Item(i).SubItems(3) _
            , Format(lv_transacciones.ListItems.Item(i).SubItems(4), "###,###,##0.00"), lab_folio, lab_afectacion, lab_folio_transacciones, "A" _
            , txt_fecha_transaccion, fun_NombreUsuario, fun_NombrePc, var_numero_planta)
        Next i
        If ok Then
            SetTimer hwnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
            MsgBox "Se Guardo Exitosamente la Transaccion", , "TRANSACCIONES [ AVISO ]"
            If (MsgBox("Desea Imprimir Movimiento S / N ?", vbYesNo + vbInformation, "AVISO") = vbYes) Then
                pro_imprime
            End If
        Else
            MsgBox "No se puede grabar registro: " + TB_Articulos.MensajeError, vbOKOnly + vbCritical, "ATENCION"
        End If
        pro_limpia_todo
        
Set TB_DETALLE = Nothing

End Sub

Public Sub pro_imprime()
    With frmreportes
       ' .CR1.Destination = crptToPrinter
       ' '.CR1.WindowState = crptMaximized
       ' .CR1.CopiesToPrinter = 1
       ' .CR1.SelectionFormula = "{tb_transacciones.bint_tra_transacciones_id} = " + lab_folio
       ' .CR1.Formulas(0) = "usuario= """ & fun_NombreUsuario & """"
       ' .CR1.Formulas(1) = "maquina= """ & fun_NombrePc & """"
       ' .CR1.Formulas(2) = "hora= """ & Time & """"
       ' .CR1.Formulas(3) = "folio= """ & lab_folio_transacciones & """"
       ' .CR1.ReportFileName = App.Path + "\mov.rpt"
       ' .CR1.Action = 1
    End With
    'Call pro_Vacia_formulas(frmreportes.CR1, 4)
    
End Sub
Public Sub pro_limpia_lista()
    
    For i = 1 To lv_transacciones.ListItems.Count
        lv_transacciones.ListItems.Remove (1)
    Next i
    
End Sub



Sub pro_costeo_entredas_salidas(lvSelectedItem As String, lvsubitem1 As String, lvsubitem2 As String, lvsubitem3 As String)

Dim var_costo_promedio As String, var_existencia As String
Dim var_importe1 As Double, var_importe2 As Double, var_importe3 As Double
Dim var_nuevo_costo_promedio As Double, var_ultimo_costo
Dim var_ultima_compra As String, var_ultima_salida As String

    
   ' var_ultima_compra = Obtener_llave(cnn, rsaux, "tb_articulos_view", "vcha_art_articulo_id", lvSelectedItem, 6, "T")
   ' var_ultima_salida = Obtener_llave(cnn, rsaux, "tb_articulos_view", "vcha_art_articulo_id", lvSelectedItem, 7, "T")
    
    If var_afecta = "SUMA" Then
        var_costo_promedio = Obtener_llave(cnn, rsaux, "TB_ARTICULOS_view", "vcha_art_articulo_id", lvSelectedItem, 7, "T")
        var_existencia = Obtener_llave(cnn, rsaux, "TB_ARTICULOS_view", "vcha_art_articulo_id", lvSelectedItem, 17, "T")

            var_importe1 = Val(var_costo_promedio) * Val(var_existencia)
            
            var_importe2 = var_existencia + Val(lvsubitem1)    ' suma existencia
            
            var_importe3 = Val(var_importe1) + Val(lvsubitem2)
            var_nuevo_costo_promedio = var_importe3 / var_importe2 ' Costo Promedio

        If var_existencia <> 0 Then
            Call fun_Actualizar(lvSelectedItem, lvsubitem3, Str(var_nuevo_costo_promedio), Str(var_importe2))
        Else
            Call fun_Actualizar(lvSelectedItem, lvsubitem3, lvsubitem3, Str(var_importe2))
        End If
    Else
        var_costo_promedio = Obtener_llave(cnn, rsaux, "TB_ARTICULOS_view", "vcha_art_articulo_id", lvSelectedItem, 7, "T")
        var_existencia = Obtener_llave(cnn, rsaux, "TB_ARTICULOS_view", "vcha_art_articulo_id", lvSelectedItem, 17, "T")
        var_importe1 = Val(var_costo_promedio) * Val(var_existencia)
        
        var_importe2 = var_existencia - Val(lvsubitem1)    ' cantidad total
        
        'var_importe3 = Val(var_importe1) - Val(lvsubitem2)
        'var_nuevo_costo_promedio = var_importe3 / var_importe2
        var_ultimo_costo = Obtener_llave(cnn, rsaux, "TB_ARTICULOS_view", "vcha_art_articulo_id", lvSelectedItem, 6, "T")
        
        Call fun_Actualizar(lvSelectedItem, Str(var_ultimo_costo), Str(var_costo_promedio), Str(var_importe2))
    End If


End Sub



'________________ actualiza la existencia el costo y la fecha de la operacion_______________

Public Function fun_Actualizar(ByVal clVcha_art_articulo_id As String _
, clVcha_art_ultcosto As String, clVcha_art_cospromedio As String _
, clVcha_art_existencia As String) As Boolean
Dim cmd As New Command

fun_Actualizar = True
On Error GoTo HELL

Set cmd.ActiveConnection = cnn                      'Esta es la conexión activa
cmd.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    

    cmd.CommandText = "ARTICULOS_ACTUALIZA"                         'Abrir Procedimiento Almacenado para Actualizar Cambios

    cmd("@Vcha_art_articulo_id") = clVcha_art_articulo_id
    cmd("@FLOA_art_ultcosto") = clVcha_art_ultcosto
    cmd("@FLOA_art_cospromedio") = clVcha_art_cospromedio
    cmd("@FLOA_ART_EXISTENCIA") = clVcha_art_existencia

    
    
   
cmd.execute                                         'Ejecutar el PA
Set cmd = Nothing                                   'Liberar Memoria

SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    fun_Actualizar = False
    GoTo SIGUE
End Function


Private Sub txtbox_Change()
Dim prvsTxt As String
Dim txt As String, posCursor As Integer
    
    posCursor = txtbox.SelStart
    ' If Cursor does not stay on the beginning of text box.
    If posCursor <> 0 Then
        
        ' keep previous value of the text box
        prvsTxt = Trim(txtbox.Text)
        
        ' keep piece of text before cursor
        txt = Left(prvsTxt, posCursor)
        
        ' Call API function to find the appropriate entry in the list box
        lstbox.ListIndex = SendMessage(lstbox.hwnd, VB_SEARCHSTR, -1, ByVal txt)
        
        ' We have found appropriate value
        If lstbox.ListIndex <> -1 Then
            txtbox.Text = lstbox.Text
            txtbox.SelStart = posCursor
        ' We didn't find appropriate entry and return previous value to text box
        Else
            txtbox.Text = Left(prvsTxt, posCursor - 1) + Mid(prvsTxt, posCursor + 1)
            txtbox.SelStart = posCursor - 1
        End If
    
    End If

End Sub

Private Sub txtbox_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 And txt_movimientos(0) <> "" Then
        fra_buscar.Visible = False
        txt_movimientos(2).SetFocus
        txtbox = ""
    End If
    If KeyAscii = 13 Then
        fra_buscar.Visible = False
        txt_movimientos(2).SetFocus
        txtbox = ""
    End If

End Sub
