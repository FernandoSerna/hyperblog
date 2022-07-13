VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmarticulos2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Artículos"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   300
      Left            =   2730
      TabIndex        =   160
      Top             =   60
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   570
      TabIndex        =   127
      Top             =   1545
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   128
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
         TabIndex        =   129
         Top             =   135
         Width           =   5610
      End
   End
   Begin VB.Frame frm_almacen 
      Height          =   2940
      Left            =   555
      TabIndex        =   149
      Top             =   1365
      Width           =   5730
      Begin MSComctlLib.ListView lv_almacenes 
         Height          =   2460
         Left            =   45
         TabIndex        =   151
         Top             =   405
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   4339
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
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         Caption         =   " Almacenes"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   150
         Top             =   120
         Width           =   5655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Artículo "
      Height          =   6795
      Left            =   90
      TabIndex        =   132
      Top             =   525
      Width           =   6225
      Begin VB.TextBox txt_volumen 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4170
         MaxLength       =   9
         TabIndex        =   42
         Text            =   "36"
         Top             =   6045
         Width           =   1485
      End
      Begin VB.TextBox txt_peso 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   1470
         MaxLength       =   9
         TabIndex        =   41
         Text            =   "35"
         Top             =   6045
         Width           =   1485
      End
      Begin VB.CheckBox chk_numero_serie 
         Caption         =   "Requiere Número de Serie"
         Height          =   225
         Left            =   150
         TabIndex        =   43
         Top             =   6450
         Width           =   2280
      End
      Begin VB.TextBox txt_nombre_estampado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2685
         TabIndex        =   40
         Top             =   5700
         Width           =   3420
      End
      Begin VB.TextBox txt_estampado 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   1485
         TabIndex        =   39
         Top             =   5700
         Width           =   1170
      End
      Begin VB.TextBox txt_nombre_subdivision 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2685
         TabIndex        =   38
         Top             =   5355
         Width           =   3420
      End
      Begin VB.TextBox txt_subdivision 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   1485
         TabIndex        =   37
         Top             =   5355
         Width           =   1170
      End
      Begin VB.TextBox txt_nombre_division 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2685
         TabIndex        =   36
         Top             =   5010
         Width           =   3420
      End
      Begin VB.TextBox txt_division 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   1485
         TabIndex        =   35
         Top             =   5010
         Width           =   1170
      End
      Begin VB.TextBox txt_nombre_tipo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2685
         TabIndex        =   34
         Top             =   4665
         Width           =   3420
      End
      Begin VB.TextBox txt_tipo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   1485
         TabIndex        =   33
         Top             =   4665
         Width           =   1170
      End
      Begin VB.CheckBox chk_detenido 
         Caption         =   "Detenido"
         Height          =   210
         Left            =   4425
         TabIndex        =   32
         Top             =   4365
         Width           =   1410
      End
      Begin VB.TextBox txt_equivalente 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1485
         TabIndex        =   30
         Top             =   4320
         Width           =   1170
      End
      Begin MSComCtl2.MonthView mes 
         Height          =   2370
         Left            =   2430
         TabIndex        =   133
         Top             =   1065
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   81723393
         CurrentDate     =   38341
      End
      Begin VB.CheckBox chk_salida_masiva 
         Caption         =   "Salida masiva"
         Height          =   210
         Left            =   2775
         TabIndex        =   31
         Top             =   4365
         Width           =   1410
      End
      Begin VB.TextBox txt_nombre_linea 
         Height          =   315
         Left            =   2685
         TabIndex        =   23
         Top             =   2940
         Width           =   3420
      End
      Begin VB.TextBox txt_nombre_catalogo_fin 
         Height          =   315
         Left            =   2700
         TabIndex        =   16
         Top             =   1605
         Width           =   2925
      End
      Begin VB.TextBox txt_sublinea 
         Height          =   315
         Left            =   1485
         TabIndex        =   24
         Top             =   3285
         Width           =   1170
      End
      Begin VB.TextBox txt_fecha_inicio 
         Height          =   315
         Left            =   1485
         TabIndex        =   12
         Text            =   "4"
         Top             =   945
         Width           =   1335
      End
      Begin VB.TextBox txt_fecha_fin 
         Height          =   315
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "5"
         Top             =   945
         Width           =   1335
      End
      Begin VB.TextBox txt_catalogo_inicio 
         Height          =   315
         Left            =   1485
         TabIndex        =   13
         Text            =   "6"
         Top             =   1275
         Width           =   1170
      End
      Begin VB.TextBox txt_catalogo_fin 
         Height          =   315
         Left            =   1485
         TabIndex        =   15
         Text            =   "7"
         Top             =   1605
         Width           =   1170
      End
      Begin VB.TextBox txt_codigo 
         Height          =   315
         Left            =   1485
         TabIndex        =   7
         Text            =   "0"
         Top             =   285
         Width           =   1320
      End
      Begin VB.TextBox txt_descripcion 
         Height          =   315
         Left            =   2820
         MaxLength       =   50
         TabIndex        =   8
         Text            =   "1"
         Top             =   285
         Width           =   3300
      End
      Begin VB.TextBox txt_contrato 
         Height          =   315
         Left            =   1485
         TabIndex        =   19
         Text            =   "9"
         Top             =   2265
         Width           =   3195
      End
      Begin VB.TextBox txt_dueño_licencia 
         Height          =   315
         Left            =   1485
         TabIndex        =   17
         Top             =   1935
         Width           =   1170
      End
      Begin VB.TextBox txt_linea 
         Height          =   315
         Left            =   1485
         TabIndex        =   22
         Top             =   2940
         Width           =   1170
      End
      Begin VB.TextBox txt_familia 
         Height          =   315
         Left            =   1485
         TabIndex        =   20
         Top             =   2595
         Width           =   1170
      End
      Begin VB.CommandButton cmd_catalogos 
         Height          =   315
         Left            =   5655
         Picture         =   "frmarticulos2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Asignación de catalogo por canal de venta"
         Top             =   1635
         Width           =   330
      End
      Begin VB.TextBox txt_nombre_catalogo_inicio 
         Height          =   315
         Left            =   2685
         TabIndex        =   14
         Top             =   1275
         Width           =   3420
      End
      Begin VB.TextBox txt_nombre_dueño_licencia 
         Height          =   315
         Left            =   2685
         TabIndex        =   18
         Top             =   1935
         Width           =   3420
      End
      Begin VB.TextBox txt_nombre_familia 
         Height          =   315
         Left            =   2685
         TabIndex        =   21
         Top             =   2595
         Width           =   3420
      End
      Begin VB.TextBox txt_nombre_sublinea 
         Height          =   315
         Left            =   2685
         TabIndex        =   25
         Top             =   3285
         Width           =   3420
      End
      Begin VB.CommandButton cmd_fecha_inicio 
         Height          =   315
         Left            =   5805
         Picture         =   "frmarticulos2.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Fecha de Alta"
         Top             =   945
         Width           =   330
      End
      Begin VB.CommandButton cmd_fecha_fin 
         Height          =   315
         Left            =   2850
         Picture         =   "frmarticulos2.frx":1374
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Fecha de Baja"
         Top             =   945
         Width           =   330
      End
      Begin VB.TextBox txt_unidad 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1485
         TabIndex        =   28
         Top             =   3975
         Width           =   1170
      End
      Begin VB.TextBox txt_talla 
         Height          =   315
         Left            =   1485
         TabIndex        =   26
         Top             =   3630
         Width           =   1170
      End
      Begin VB.TextBox txt_nombre_talla 
         Height          =   315
         Left            =   2685
         TabIndex        =   27
         Top             =   3630
         Width           =   3420
      End
      Begin VB.TextBox txt_nombre_unidad 
         Height          =   315
         Left            =   2685
         TabIndex        =   29
         Top             =   3975
         Width           =   3420
      End
      Begin MSMask.MaskEdBox msk_precio 
         Height          =   315
         Left            =   1485
         TabIndex        =   9
         Top             =   615
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_costo 
         Height          =   315
         Left            =   4440
         TabIndex        =   10
         Top             =   615
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   556
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Volumen:"
         Height          =   195
         Index           =   30
         Left            =   3015
         TabIndex        =   159
         Top             =   6105
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "3"
         Height          =   195
         Index           =   35
         Left            =   5940
         TabIndex        =   158
         Top             =   6030
         Width           =   225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cm"
         Height          =   195
         Index           =   34
         Left            =   5700
         TabIndex        =   157
         Top             =   6105
         Width           =   225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Peso:"
         Height          =   195
         Index           =   25
         Left            =   165
         TabIndex        =   156
         Top             =   6105
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Estampado:"
         Height          =   195
         Index           =   49
         Left            =   150
         TabIndex        =   155
         Top             =   5745
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Subdivision:"
         Height          =   195
         Index           =   48
         Left            =   135
         TabIndex        =   154
         Top             =   5400
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Division:"
         Height          =   195
         Index           =   47
         Left            =   135
         TabIndex        =   153
         Top             =   5055
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo producto:"
         Height          =   195
         Index           =   46
         Left            =   135
         TabIndex        =   152
         Top             =   4710
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Equivalencia:"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   148
         Top             =   4380
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Contrato:"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   147
         Top             =   2340
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha alta:"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   146
         Top             =   1020
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha baja:"
         Height          =   195
         Index           =   2
         Left            =   3300
         TabIndex        =   145
         Top             =   1035
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cátalogo de inicio:"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   144
         Top             =   1335
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ultimo cátalogo:"
         Height          =   195
         Index           =   9
         Left            =   135
         TabIndex        =   143
         Top             =   1665
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Precio base:"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   142
         Top             =   675
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Costo estandar:"
         Height          =   195
         Index           =   3
         Left            =   3300
         TabIndex        =   141
         Top             =   675
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Artículo:"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   140
         Top             =   345
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Dueño de licencia:"
         Height          =   195
         Index           =   22
         Left            =   135
         TabIndex        =   139
         Top             =   1995
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Linea:"
         Height          =   195
         Index           =   7
         Left            =   135
         TabIndex        =   138
         Top             =   3000
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Familia:"
         Height          =   195
         Index           =   8
         Left            =   135
         TabIndex        =   137
         Top             =   2655
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sublinea:"
         Height          =   195
         Index           =   10
         Left            =   135
         TabIndex        =   136
         Top             =   3345
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Talla:"
         Height          =   195
         Index           =   20
         Left            =   135
         TabIndex        =   135
         Top             =   3675
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Unidad:"
         Height          =   195
         Index           =   21
         Left            =   135
         TabIndex        =   134
         Top             =   4020
         Width           =   555
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5910
      Picture         =   "frmarticulos2.frx":25E6
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   60
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1485
      Picture         =   "frmarticulos2.frx":2C20
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   60
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1155
      Picture         =   "frmarticulos2.frx":2D22
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   60
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   819
      Picture         =   "frmarticulos2.frx":2E24
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   60
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   492
      Picture         =   "frmarticulos2.frx":2EF6
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Guardar Alt + G"
      Top             =   60
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   165
      Picture         =   "frmarticulos2.frx":2FF8
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   60
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   6390
      TabIndex        =   123
      Top             =   495
      Visible         =   0   'False
      Width           =   15
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   1740
         TabIndex        =   124
         Top             =   150
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3600
         TabIndex        =   125
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
         Caption         =   "Busqueda de artículo:"
         Height          =   195
         Index           =   4
         Left            =   105
         TabIndex        =   126
         Top             =   210
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   12645
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   75
      Width           =   255
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   91
      Top             =   315
      Width           =   6165
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   1710
      Top             =   -60
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
            Picture         =   "frmarticulos2.frx":30FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmarticulos2.frx":39D4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   6255
      Left            =   6720
      TabIndex        =   107
      Top             =   5025
      Visible         =   0   'False
      Width           =   5550
      Begin MSComctlLib.ListView lv_articulos 
         Height          =   6000
         Left            =   195
         TabIndex        =   108
         Top             =   195
         Width           =   5130
         _ExtentX        =   9049
         _ExtentY        =   10583
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
         NumItems        =   68
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   5733
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "precio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "costo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "fecha baja"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "fecha alta"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "catalogo inicio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "nombre catalogo inicio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Catalogo vigente"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "nombre catalogo vigente"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "dueño lic"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "nombre dueño licencia"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "numero licencia"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "diseño"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "nombre diseño"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "linea"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "nombre linea"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "sublinea"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "nombre sublinea"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "producto"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "nombre producto"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "tipo producto"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "nombre tipo producto"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Text            =   "clase"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Text            =   "nombre clase"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   25
            Text            =   "estampado anverso"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   26
            Text            =   "nombre estampado anverso"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   27
            Text            =   "tipo estampado anverso"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   28
            Text            =   "nombre tipo estampado anverso"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   29
            Text            =   "estampado reverso"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   30
            Text            =   "nombre estampado reverso"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   31
            Text            =   "tipo estampado reverso"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   32
            Text            =   "nombre tipo estampado reverso"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   33
            Text            =   "color anverso"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   34
            Text            =   "nombre color anverso"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   35
            Text            =   "color reverso"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   36
            Text            =   "nombre color reverso"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   37
            Text            =   "tono anverso"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   38
            Text            =   "nombre tono anverso"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   39
            Text            =   "tono reverso"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   40
            Text            =   "nombre tono reverso"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   41
            Text            =   "decorativos"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   42
            Text            =   "fundas"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(44) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   43
            Text            =   "uso"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(45) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   44
            Text            =   "nombre uso"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(46) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   45
            Text            =   "sub tipo uso"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(47) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   46
            Text            =   "nombre subtipo uso"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(48) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   47
            Text            =   "talla"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(49) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   48
            Text            =   "nombre talla"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(50) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   49
            Text            =   "unidad"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(51) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   50
            Text            =   "nombre unidada"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(52) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   51
            Text            =   "volumen"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(53) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   52
            Text            =   "tela"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(54) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   53
            Text            =   "composicion"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(55) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   54
            Text            =   "peso"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(56) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   55
            Text            =   "tara"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(57) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   56
            Text            =   "caja"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(58) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   57
            Text            =   "nombre caja"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(59) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   58
            Text            =   "piezas caja"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(60) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   59
            Text            =   "maximo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(61) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   60
            Text            =   "minimo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(62) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   61
            Text            =   "punto reorden"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(63) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   62
            Text            =   "dias inventario"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(64) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   63
            Text            =   "ubicacion"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(65) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   64
            Text            =   "nombre ubicacion"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(66) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   65
            Text            =   "BULTO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(67) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   66
            Text            =   "SALIDA MASIVA"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(68) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   67
            Text            =   "Detenido"
            Object.Width           =   0
         EndProperty
      End
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
            Picture         =   "frmarticulos2.frx":42AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmarticulos2.frx":4B88
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmarticulos2.frx":5462
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmarticulos2.frx":59FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmarticulos2.frx":62DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmarticulos2.frx":6BB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmarticulos2.frx":748E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmarticulos2.frx":75A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmarticulos2.frx":76B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmarticulos2.frx":77C4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tabarticulos 
      Height          =   3210
      Left            =   6795
      TabIndex        =   90
      Top             =   1095
      Visible         =   0   'False
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   5662
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Descripción"
      TabPicture(0)   =   "frmarticulos2.frx":78D6
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txt_nombre_tipo_producto"
      Tab(0).Control(1)=   "txt_producto"
      Tab(0).Control(2)=   "txt_tipo_producto"
      Tab(0).Control(3)=   "txt_clase"
      Tab(0).Control(4)=   "txt_nombre_producto"
      Tab(0).Control(5)=   "txt_nombre_clase"
      Tab(0).Control(6)=   "Label2(12)"
      Tab(0).Control(7)=   "Label2(11)"
      Tab(0).Control(8)=   "Label2(4)"
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Caracteristicas"
      TabPicture(1)   =   "frmarticulos2.frx":78F2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt_nombre_tipo_estampado_anverso"
      Tab(1).Control(1)=   "txt_nombre_Estampado_anverso"
      Tab(1).Control(2)=   "txt_estampado_anverso"
      Tab(1).Control(3)=   "txt_tipo_estampado_anverso"
      Tab(1).Control(4)=   "txt_fundas"
      Tab(1).Control(5)=   "txt_decorativos"
      Tab(1).Control(6)=   "txt_tipo_estampado_reverso"
      Tab(1).Control(7)=   "txt_estampado_reverso"
      Tab(1).Control(8)=   "txt_uso_producto"
      Tab(1).Control(9)=   "txt_subtipo_uso_producto"
      Tab(1).Control(10)=   "txt_tono_reverso"
      Tab(1).Control(11)=   "txt_tono_anverso"
      Tab(1).Control(12)=   "txt_color_anverso"
      Tab(1).Control(13)=   "txt_color_reverso"
      Tab(1).Control(14)=   "txt_nombre_estampado_reverso"
      Tab(1).Control(15)=   "txt_nombre_color_anverso"
      Tab(1).Control(16)=   "txt_nombre_tono_anverso"
      Tab(1).Control(17)=   "txt_nombre_uso_producto"
      Tab(1).Control(18)=   "txt_nombre_tipo_estampado_reverso"
      Tab(1).Control(19)=   "txt_nombre_color_reverso"
      Tab(1).Control(20)=   "txt_nombre_tono_reverso"
      Tab(1).Control(21)=   "txt_nombre_subtipo_uso_producto"
      Tab(1).Control(22)=   "Label2(14)"
      Tab(1).Control(23)=   "Label2(13)"
      Tab(1).Control(24)=   "Label2(44)"
      Tab(1).Control(25)=   "Label2(43)"
      Tab(1).Control(26)=   "Label2(42)"
      Tab(1).Control(27)=   "Label2(41)"
      Tab(1).Control(28)=   "Label2(15)"
      Tab(1).Control(29)=   "Label2(16)"
      Tab(1).Control(30)=   "Label2(17)"
      Tab(1).Control(31)=   "Label1(2)"
      Tab(1).Control(32)=   "Label2(18)"
      Tab(1).Control(33)=   "Label2(19)"
      Tab(1).ControlCount=   34
      TabCaption(2)   =   "Inventarios"
      TabPicture(2)   =   "frmarticulos2.frx":790E
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label2(33)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label2(32)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label2(31)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label2(29)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label2(28)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label2(27)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label2(26)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label2(24)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label2(23)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Label2(40)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label2(39)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Label2(38)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Label2(37)"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Label2(36)"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Label2(1)"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Label2(45)"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "txt_tara"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "txt_minimo"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "txt_punto_reorden"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "txt_caja"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "txt_dias_inventario"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "txt_piezas_caja"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "txt_ubicacion"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "txt_maximo"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "txt_bulto"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "txt_volumen_compreso"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "txt_compresion"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "txt_tela"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "txt_composicion"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "txt_nombre_caja"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "txt_nombre_ubicacion"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).ControlCount=   31
      Begin VB.TextBox txt_nombre_tipo_estampado_anverso 
         Height          =   315
         Left            =   -66525
         TabIndex        =   56
         Top             =   450
         Width           =   2835
      End
      Begin VB.TextBox txt_nombre_Estampado_anverso 
         Height          =   315
         Left            =   -72360
         TabIndex        =   54
         Top             =   450
         Width           =   2865
      End
      Begin VB.TextBox txt_estampado_anverso 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73410
         TabIndex        =   53
         Top             =   450
         Width           =   1020
      End
      Begin VB.TextBox txt_tipo_estampado_anverso 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -67560
         TabIndex        =   55
         Top             =   450
         Width           =   1020
      End
      Begin VB.TextBox txt_nombre_ubicacion 
         Height          =   315
         Left            =   9345
         TabIndex        =   88
         Top             =   2115
         Width           =   1860
      End
      Begin VB.TextBox txt_nombre_caja 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2370
         TabIndex        =   81
         Top             =   1455
         Width           =   2040
      End
      Begin VB.TextBox txt_nombre_tipo_producto 
         Height          =   315
         Left            =   -72255
         TabIndex        =   50
         Top             =   2745
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txt_composicion 
         Height          =   315
         Left            =   5820
         TabIndex        =   78
         Text            =   "34"
         Top             =   795
         Width           =   1950
      End
      Begin VB.TextBox txt_tela 
         Height          =   315
         Left            =   1335
         TabIndex        =   77
         Text            =   "33"
         Top             =   795
         Width           =   1485
      End
      Begin VB.TextBox txt_compresion 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5820
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   75
         Text            =   "31"
         Top             =   465
         Width           =   1515
      End
      Begin VB.TextBox txt_volumen_compreso 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8640
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   76
         Text            =   "32"
         Top             =   480
         Width           =   1590
      End
      Begin VB.TextBox txt_producto 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -67980
         TabIndex        =   47
         Top             =   2415
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txt_tipo_producto 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73440
         TabIndex        =   49
         Top             =   2745
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txt_clase 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -67980
         TabIndex        =   51
         Top             =   2745
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txt_fundas 
         Height          =   315
         Left            =   -67560
         MaxLength       =   9
         TabIndex        =   70
         Text            =   "25"
         Top             =   1770
         Width           =   1020
      End
      Begin VB.TextBox txt_decorativos 
         Height          =   315
         Left            =   -73410
         MaxLength       =   9
         TabIndex        =   69
         Text            =   "24"
         Top             =   1770
         Width           =   1020
      End
      Begin VB.TextBox txt_tipo_estampado_reverso 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -67560
         TabIndex        =   59
         Top             =   780
         Width           =   1020
      End
      Begin VB.TextBox txt_estampado_reverso 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73410
         TabIndex        =   57
         Top             =   780
         Width           =   1020
      End
      Begin VB.TextBox txt_bulto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1335
         MaxLength       =   9
         TabIndex        =   89
         Text            =   "44"
         Top             =   2445
         Width           =   1485
      End
      Begin VB.TextBox txt_maximo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1335
         MaxLength       =   9
         TabIndex        =   83
         Text            =   "39"
         Top             =   1785
         Width           =   1485
      End
      Begin VB.TextBox txt_ubicacion 
         Enabled         =   0   'False
         Height          =   315
         Left            =   8310
         TabIndex        =   87
         Top             =   2115
         Width           =   1020
      End
      Begin VB.TextBox txt_piezas_caja 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5820
         Locked          =   -1  'True
         TabIndex        =   82
         Text            =   "38"
         Top             =   1455
         Width           =   1515
      End
      Begin VB.TextBox txt_dias_inventario 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5820
         MaxLength       =   9
         TabIndex        =   86
         Text            =   "42"
         Top             =   2115
         Width           =   1515
      End
      Begin VB.TextBox txt_caja 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   1335
         TabIndex        =   80
         Top             =   1455
         Width           =   1020
      End
      Begin VB.TextBox txt_punto_reorden 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1335
         MaxLength       =   9
         TabIndex        =   85
         Text            =   "41"
         Top             =   2115
         Width           =   1485
      End
      Begin VB.TextBox txt_minimo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5820
         MaxLength       =   9
         TabIndex        =   84
         Text            =   "40"
         Top             =   1785
         Width           =   1515
      End
      Begin VB.TextBox txt_tara 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5820
         MaxLength       =   9
         TabIndex        =   79
         Text            =   "36"
         Top             =   1125
         Width           =   1515
      End
      Begin VB.TextBox txt_uso_producto 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73410
         TabIndex        =   71
         Top             =   2100
         Width           =   1020
      End
      Begin VB.TextBox txt_subtipo_uso_producto 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -67560
         TabIndex        =   73
         Top             =   2100
         Width           =   1020
      End
      Begin VB.TextBox txt_tono_reverso 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -67560
         TabIndex        =   67
         Top             =   1440
         Width           =   1020
      End
      Begin VB.TextBox txt_tono_anverso 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73410
         TabIndex        =   65
         Top             =   1440
         Width           =   1020
      End
      Begin VB.TextBox txt_color_anverso 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73410
         TabIndex        =   61
         Top             =   1110
         Width           =   1020
      End
      Begin VB.TextBox txt_color_reverso 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -67560
         TabIndex        =   63
         Top             =   1110
         Width           =   1020
      End
      Begin VB.TextBox txt_nombre_producto 
         Height          =   315
         Left            =   -66795
         TabIndex        =   48
         Top             =   2415
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txt_nombre_clase 
         Height          =   315
         Left            =   -66795
         TabIndex        =   52
         Top             =   2745
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txt_nombre_estampado_reverso 
         Height          =   315
         Left            =   -72375
         TabIndex        =   58
         Top             =   780
         Width           =   2865
      End
      Begin VB.TextBox txt_nombre_color_anverso 
         Height          =   315
         Left            =   -72375
         TabIndex        =   62
         Top             =   1110
         Width           =   2865
      End
      Begin VB.TextBox txt_nombre_tono_anverso 
         Height          =   315
         Left            =   -72375
         TabIndex        =   66
         Top             =   1440
         Width           =   2865
      End
      Begin VB.TextBox txt_nombre_uso_producto 
         Height          =   315
         Left            =   -72375
         TabIndex        =   72
         Top             =   2100
         Width           =   2865
      End
      Begin VB.TextBox txt_nombre_tipo_estampado_reverso 
         Height          =   315
         Left            =   -66525
         TabIndex        =   60
         Top             =   780
         Width           =   2835
      End
      Begin VB.TextBox txt_nombre_color_reverso 
         Height          =   315
         Left            =   -66525
         TabIndex        =   64
         Top             =   1110
         Width           =   2835
      End
      Begin VB.TextBox txt_nombre_tono_reverso 
         Height          =   315
         Left            =   -66510
         TabIndex        =   68
         Top             =   1440
         Width           =   2835
      End
      Begin VB.TextBox txt_nombre_subtipo_uso_producto 
         Height          =   315
         Left            =   -66525
         TabIndex        =   74
         Top             =   2100
         Width           =   2835
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Estampado anverso:"
         Height          =   195
         Index           =   14
         Left            =   -69435
         TabIndex        =   131
         Top             =   510
         Width           =   1815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Estampado anverso:"
         Height          =   195
         Index           =   13
         Left            =   -74910
         TabIndex        =   130
         Top             =   510
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Composición:"
         Height          =   195
         Index           =   45
         Left            =   4485
         TabIndex        =   122
         Top             =   855
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tela:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   121
         Top             =   855
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Index           =   36
         Left            =   5850
         TabIndex        =   120
         Top             =   525
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Compresión:"
         Height          =   195
         Index           =   37
         Left            =   4500
         TabIndex        =   119
         Top             =   540
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Vol. Compreso:"
         Height          =   195
         Index           =   38
         Left            =   7530
         TabIndex        =   118
         Top             =   525
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "3"
         Height          =   195
         Index           =   39
         Left            =   10515
         TabIndex        =   117
         Top             =   465
         Width           =   225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cm"
         Height          =   195
         Index           =   40
         Left            =   10275
         TabIndex        =   116
         Top             =   525
         Width           =   225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clase:"
         Height          =   195
         Index           =   12
         Left            =   -69165
         TabIndex        =   115
         Top             =   2805
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Producto:"
         Height          =   195
         Index           =   11
         Left            =   -69165
         TabIndex        =   114
         Top             =   2475
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo producto:"
         Height          =   195
         Index           =   4
         Left            =   -74790
         TabIndex        =   113
         Top             =   2805
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fundas:"
         Height          =   195
         Index           =   44
         Left            =   -69420
         TabIndex        =   112
         Top             =   1830
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Decorativos:"
         Height          =   195
         Index           =   43
         Left            =   -74910
         TabIndex        =   111
         Top             =   1830
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Estampado reverso:"
         Height          =   195
         Index           =   42
         Left            =   -69420
         TabIndex        =   110
         Top             =   840
         Width           =   1770
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Estampado reverso:"
         Height          =   195
         Index           =   41
         Left            =   -74910
         TabIndex        =   109
         Top             =   840
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Pieza por caja:"
         Height          =   195
         Index           =   23
         Left            =   4485
         TabIndex        =   106
         Top             =   1515
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tara:"
         Height          =   195
         Index           =   24
         Left            =   4485
         TabIndex        =   105
         Top             =   1185
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Caja:"
         Height          =   195
         Index           =   26
         Left            =   180
         TabIndex        =   104
         Top             =   1515
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Máximo:"
         Height          =   195
         Index           =   27
         Left            =   180
         TabIndex        =   103
         Top             =   1845
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Dias de inventario:"
         Height          =   195
         Index           =   28
         Left            =   4485
         TabIndex        =   102
         Top             =   2160
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Pto. de reorden:"
         Height          =   195
         Index           =   29
         Left            =   180
         TabIndex        =   101
         Top             =   2160
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Bulto:"
         Height          =   195
         Index           =   31
         Left            =   180
         TabIndex        =   100
         Top             =   2490
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Mínimo:"
         Height          =   195
         Index           =   32
         Left            =   4485
         TabIndex        =   99
         Top             =   1845
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ubicación:"
         Height          =   195
         Index           =   33
         Left            =   7410
         TabIndex        =   98
         Top             =   2160
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tono reverso:"
         Height          =   195
         Index           =   15
         Left            =   -69420
         TabIndex        =   97
         Top             =   1500
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tono anverso:"
         Height          =   195
         Index           =   16
         Left            =   -74910
         TabIndex        =   96
         Top             =   1500
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Color reverso:"
         Height          =   195
         Index           =   17
         Left            =   -69420
         TabIndex        =   95
         Top             =   1170
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Color anverso:"
         Height          =   195
         Index           =   2
         Left            =   -74910
         TabIndex        =   94
         Top             =   1170
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Uso del prodcuto:"
         Height          =   195
         Index           =   18
         Left            =   -74910
         TabIndex        =   93
         Top             =   2145
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Subtipo de uso:"
         Height          =   195
         Index           =   19
         Left            =   -69420
         TabIndex        =   92
         Top             =   2145
         Width           =   1110
      End
   End
End
Attribute VB_Name = "frmarticulos2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_hubo_cambios As Boolean
Dim var_hubo_cambios_2 As Boolean
Dim numero_items_articulos As Integer
Dim var_guarda_cambios As Boolean
Dim var_falta As String
Dim var_campos As Integer
'Dim appl As New CRAXDRT.Application
'Dim reporte As New CRAXDRT.Report
Dim var_mismo_nombre_articulo As String
Dim var_ancho, var_largo, var_alto As Double
Dim var_tipo_lista As Integer
Dim var_tipo_fecha As Integer
Dim var_ventana As Integer
Dim var_tipo_mes As Integer




Private Sub chk_detenido_Click()
   var_hubo_cambios = True
End Sub

Private Sub chk_numero_serie_Click()
   var_hubo_cambios = True
End Sub

Private Sub chk_salida_masiva_Click()
   var_hubo_cambios = True
End Sub







Private Sub cmd_fecha_fin_Click()
   var_ventana = 1
   var_tipo_mes = 2
   If Trim(txt_fecha_inicio) <> "" Then
      If IsDate(txt_fecha_inicio) Then
         mes.Value = CDate(txt_fecha_inicio)
      Else
         mes.Value = Date
      End If
   Else
      mes.Value = Date
   End If
   mes.Visible = True
   mes.SetFocus
End Sub

Private Sub cmd_fecha_inicio_Click()
   var_ventana = 1
   var_tipo_mes = 1
   mes.Visible = True
   If Trim(txt_fecha_fin) <> "" Then
      If IsDate(Me.txt_fecha_fin) Then
         mes.Value = CDate(txt_fecha_fin)
      Else
         mes.Value = Date
      End If
   Else
      mes.Value = Date
   End If
   mes.SetFocus
End Sub

Private Sub Command1_Click()
   Dim sum1 As Integer
   Dim sum2 As Integer
   Dim icont As Integer
   Dim VERIFICADOR As Integer
   Dim verificador2 As Integer
   Dim var_codigo As String
   Dim longitud As Integer
   Dim msuma As Integer
   rs.Open "select * from codigos_050809", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
            txt_codigo = "646244" + Trim(CStr(rs!equivalencia))
            sum1 = 0
            sum2 = 0
            mcodigo = txt_codigo
            longitud = Len(mcodigo)
            For icont = 1 To longitud
                If ((icont / 2) - Int((icont / 2))) = 0 Then
                   sum2 = sum2 + Val(Mid(mcodigo, icont, 1))
                Else
                   sum1 = sum1 + Val(Mid(mcodigo, icont, 1))
                End If
            Next icont
            msuma = sum1 * 13 + sum2
            VERIFICADOR = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
            If VERIFICADOR = 10 Then
               VERIFICADOR = 0
            End If
            txt_codigo = txt_codigo + Trim(CStr(VERIFICADOR))
            rsaux.Open "update codigos_050809 set codigo = '" + txt_codigo + "' where equivalencia = '" + Trim(CStr(rs!equivalencia)) + "'", cnn, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      If var_ventana = 1 Then
         mes.Visible = False
         If var_tipo_mes = 1 Then
            txt_fecha_fin.SetFocus
         End If
         If var_tipo_mes = 2 Then
            txt_fecha_inicio.SetFocus
         End If
      End If
      If var_ventana = 0 Then
         frm_lista.Visible = False
         If var_tipo_lista = 1 Then
            txt_catalogo_inicio.SetFocus
         End If
         If var_tipo_lista = 2 Then
            txt_catalogo_fin.SetFocus
         End If
         If var_tipo_lista = 3 Then
            txt_dueño_licencia.SetFocus
         End If
         If var_tipo_lista = 4 Then
            txt_familia.SetFocus
         End If
         If var_tipo_lista = 5 Then
            txt_linea.SetFocus
         End If
         If var_tipo_lista = 6 Then
            txt_sublinea.SetFocus
         End If
         If var_tipo_lista = 7 Then
            txt_producto.SetFocus
         End If
         If var_tipo_lista = 8 Then
            txt_tipo_producto.SetFocus
         End If
         If var_tipo_lista = 9 Then
            txt_clase.SetFocus
         End If
         If var_tipo_lista = 10 Then
            txt_estampado_anverso.SetFocus
         End If
         If var_tipo_lista = 11 Then
            txt_tipo_estampado_anverso.SetFocus
         End If
         If var_tipo_lista = 12 Then
            txt_estampado_reverso.SetFocus
         End If
         If var_tipo_lista = 13 Then
            txt_tipo_estampado_reverso.SetFocus
         End If
         If var_tipo_lista = 14 Then
            txt_color_anverso.SetFocus
         End If
         If var_tipo_lista = 15 Then
            txt_color_reverso.SetFocus
         End If
         If var_tipo_lista = 16 Then
            txt_tono_anverso.SetFocus
         End If
         If var_tipo_lista = 17 Then
            txt_tono_reverso.SetFocus
         End If
         If var_tipo_lista = 18 Then
            txt_uso_producto.SetFocus
         End If
         If var_tipo_lista = 19 Then
            txt_subtipo_uso_producto.SetFocus
         End If
         If var_tipo_lista = 20 Then
            txt_talla.SetFocus
         End If
         If var_tipo_lista = 21 Then
            txt_unidad.SetFocus
         End If
         If var_tipo_lista = 22 Then
            txt_caja.SetFocus
         End If
      End If
   End If
End Sub

Private Sub lv_almacenes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_almacenes, ColumnHeader)
End Sub

Private Sub lv_almacenes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Set reporte = appl.OpenReport(App.Path + "\rep_catalogo_articulos_almacen.rpt")
      reporte.RecordSelectionFormula = "{VW_CATALOGO_ARTICULOS_ALMACEN.VCHA_ALM_ALMACEN_ID} = '" + lv_almacenes.selectedItem + "' or {VW_CATALOGO_ARTICULOS_ALMACEN.VCHA_ALM_ALMACEN_ID} = ''"
      frmvistasprevias.cr.ReportSource = reporte
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      frmvistasprevias.cr.ViewReport
      frmvistasprevias.Caption = "Reporte de Entradas concentrado"
      frmvistasprevias.Show 1
      Set reporte = Nothing
      var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         Set reporte = appl.OpenReport(App.Path + "\rep_catalogo_articulos_almacen.rpt")
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         reporte.RecordSelectionFormula = "{VW_CATALOGO_ARTICULOS_ALMACEN.VCHA_ALM_ALMACEN_ID} = '" + lv_almacenes.selectedItem + "'  or {VW_CATALOGO_ARTICULOS_ALMACEN.VCHA_ALM_ALMACEN_ID} = ''"
         reporte.ExportOptions.FormatType = crEFTExcel80
         reporte.ExportOptions.DestinationType = crEDTDiskFile
         archivo = "c:\reportessid\catalogo_articulos" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
         reporte.ExportOptions.DiskFileName = archivo
         reporte.Export False
         Set reporte = Nothing
         MsgBox "Se a terminado de guardar el archivo " + archivo
     End If
   End If
   If KeyAscii = 27 Then
      frm_almacen.Visible = False
   End If
End Sub

Private Sub lv_almacenes_LostFocus()
   frm_almacen.Visible = False
End Sub

Private Sub lv_articulos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_articulos, ColumnHeader)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub mes_DateDblClick(ByVal DateDblClicked As Date)
   If var_tipo_mes = 1 Then
      Me.txt_fecha_fin = mes.Value
      Me.txt_fecha_fin.SetFocus
   End If
   If var_tipo_mes = 2 Then
      Me.txt_fecha_inicio = mes.Value
      Me.txt_fecha_inicio.SetFocus
   End If
End Sub

Private Sub mes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_tipo_mes = 1 Then
         Me.txt_fecha_fin = mes.Value
         Me.txt_fecha_fin.SetFocus
      End If
      If var_tipo_mes = 2 Then
         Me.txt_fecha_inicio = mes.Value
         Me.txt_fecha_inicio.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      If var_tipo_mes = 1 Then
         Me.txt_fecha_fin.SetFocus
      End If
      If var_tipo_mes = 2 Then
         Me.txt_fecha_inicio.SetFocus
      End If
   End If
End Sub

Private Sub mes_LostFocus()
    mes.Visible = False
    var_ventana = 0
End Sub

Private Sub msk_costo_Change()
   var_hubo_cambios = True
End Sub

Private Sub msk_costo_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub msk_precio_Change()
   var_hubo_cambios = True
End Sub

Private Sub msk_precio_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub Text2_Change()

End Sub

Private Sub txt_bulto_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_bulto_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(txt_buscar) <> "" Then
         rs.Open "select * from tb_Articulos where vcha_art_articulo_id = '" + txt_buscar + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
         Else
            rsaux.Open "select * from tb_articulos where vcha_art_codigo_externo = '" + Me.txt_buscar + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               txt_buscar = IIf(IsNull(rsaux!vcha_Art_articulo_id), "", rsaux!vcha_Art_articulo_id)
            Else
               MsgBox "No existe el artículo", vbOKOnly, "ATENCION"
            End If
            rsaux.Close
         End If
         rs.Close
         Call pro_busca_registro(lv_articulos, txt_buscar, False)
         txt_buscar = ""
         pro_textos
      End If
   End If
End Sub

Private Sub txt_caja_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_caja_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_caja_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_catalogo_fin_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_catalogo_fin_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_catalogo_fin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_catalogo_inicio_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_catalogo_inicio_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_catalogo_inicio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_clase_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_clase_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_clase_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_codigo_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   'Select Case KeyAscii
   'Case 48 To 57, 52, 13, 8
   'Case Else
   '    KeyAscii = 0
   'End Select
   If var_empresa = 16 Then
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
      If KeyAscii = 39 Or KeyAscii = 61 Then
         KeyAscii = 0
      End If
      If KeyAscii = 13 Then
         Call pro_enfoque(KeyAscii)
      End If
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
      If KeyAscii = 39 Or KeyAscii = 61 Then
         KeyAscii = 0
      End If
      If KeyAscii = 13 Then
         Call pro_enfoque(KeyAscii)
      End If
   End If
End Sub

Private Sub txt_codigo_LostFocus()
   Dim sum1 As Integer
   Dim sum2 As Integer
   Dim icont As Integer
   Dim VERIFICADOR As Integer
   Dim verificador2 As Integer
   Dim var_codigo As String
   Dim longitud As Integer
   Dim msuma As Integer
   If var_empresa <> "16" Then
      If IsNumeric(Me.txt_codigo) Then
         If Len(Trim(txt_codigo)) = 5 Then
            If var_empresa <> "31" Then
               txt_codigo = "646244" + Trim(txt_codigo)
               sum1 = 0
               sum2 = 0
               mcodigo = txt_codigo
               longitud = Len(mcodigo)
               For icont = 1 To longitud
                   If ((icont / 2) - Int((icont / 2))) = 0 Then
                      sum2 = sum2 + Val(Mid(mcodigo, icont, 1))
                   Else
                      sum1 = sum1 + Val(Mid(mcodigo, icont, 1))
                   End If
               Next icont
               msuma = sum1 * 13 + sum2
               VERIFICADOR = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
               If VERIFICADOR = 10 Then
                  VERIFICADOR = 0
               End If
               txt_codigo = txt_codigo + Trim(CStr(VERIFICADOR))
            End If
         End If
         If Len(Trim(txt_codigo)) = 11 Then
            If var_empresa = "17" Or var_empresa = "31" Or var_empresa = "06" Or var_empresa = "15" Then
            Else
               sum1 = 0
               sum2 = 0
               mcodigo = txt_codigo
               longitud = Len(mcodigo)
               For icont = 1 To longitud
                   If ((icont / 2) - Int((icont / 2))) = 0 Then
                      sum2 = sum2 + Val(Mid(mcodigo, icont, 1))
                   Else
                     sum1 = sum1 + Val(Mid(mcodigo, icont, 1))
                   End If
               Next icont
               msuma = sum1 * 13 + sum2
               VERIFICADOR = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
               If VERIFICADOR = 10 Then
                  VERIFICADOR = 0
               End If
               txt_codigo = txt_codigo + Trim(CStr(VERIFICADOR))
            End If
         End If
      End If
      If rsaux5.State = 1 Then
         rsaux5.Close
      End If
      rsaux5.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux5.EOF Then
         var_empresa_articulo = IIf(IsNull(rsaux5!VCHA_EMP_EMPRESA_ID), "", rsaux5!VCHA_EMP_EMPRESA_ID)
         If var_empresa_articulo = "16" Then
            rsaux6.Open "select * from tb_empresas where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
            var_nombre_empresa = ""
            If Not rsaux6.EOF Then
               var_nombre_empresa = IIf(IsNull(rsaux6!VCHA_EMP_NOMBRE), "", rsaux6!VCHA_EMP_NOMBRE)
            End If
            rsaux6.Close
            MsgBox "Artículo incorrecto para la empresa " + var_nombre_empresa, vbOKOnly, "ATENCION"
         Else
            If Not rsaux5.EOF Then
               Call llena_registros
               Me.txt_codigo.Enabled = False
            Else
               var_codigo = Me.txt_codigo
               'Call pro_limpiatextos(Me)
               txt_codigo = var_codigo
            End If
         End If
      End If
      rsaux5.Close
   Else
      If Me.txt_codigo <> "" Then
         rsaux5.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux5.EOF Then
            var_empresa_articulo = IIf(IsNull(rsaux5!VCHA_EMP_EMPRESA_ID), "", rsaux5!VCHA_EMP_EMPRESA_ID)
            If var_empresa_articulo = "16" Then
               If Not rsaux5.EOF Then
                  Call llena_registros
                  Me.txt_codigo.Enabled = False
               Else
                  var_codigo = Me.txt_codigo
                  Call limpia_textos_2
                  txt_codigo = var_codigo
               End If
            Else
               rsaux6.Open "select * from tb_empresas where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
               var_nombre_empresa = ""
               If Not rsaux6.EOF Then
                  var_nombre_empresa = IIf(IsNull(rsaux6!VCHA_EMP_NOMBRE), "", rsaux6!VCHA_EMP_NOMBRE)
               End If
               MsgBox "Artículo incorrecto para la empresa " + var_nombre_empresa, vbOKOnly, "ATENCION"
               rsaux6.Close
            End If
         Else
            var_codigo = Me.txt_codigo
            Call limpia_textos_2
            txt_codigo = var_codigo
         End If
         rsaux5.Close
      Else
         Call limpia_textos_2
      End If
   End If
End Sub

Private Sub txt_color_anverso_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_color_anverso_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_color_anverso_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_color_reverso_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_color_reverso_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_color_reverso_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_composicion_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_composicion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_compresion_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_compresion_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_contrato_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_contrato_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_decorativos_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_decorativos_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_descripcion_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_descripcion_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_dias_inventario_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_dias_inventario_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_division_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_divisiones where vcha_tpr_tipo_producto_id = '" + Me.txt_tipo + "' order by vcha_DIV_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_div_division_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_div_nombre), "", rs!vcha_div_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "DIVISIONES"
      var_tipo_lista = 101
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
End Sub

Private Sub txt_division_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_division_LostFocus()
   If Trim(txt_tipo) <> "" Then
      If Trim(txt_division) <> "" Then
         rs.Open "select * from tb_articulos where vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_division = IIf(IsNull(rs!vcha_div_division_id), "", rs!vcha_div_division_id)
         End If
         rs.Close
         rs.Open "select * from tb_divisiones where vcha_tpr_tipo_producto_id = '" + Me.txt_tipo + "' and vcha_div_division_id = '" + Me.txt_division + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_nombre_division = IIf(IsNull(rs!vcha_div_nombre), "", rs!vcha_div_nombre)
            If var_division <> Me.txt_division Then
               Me.txt_subdivision = ""
               Me.txt_nombre_subdivision = ""
               Me.txt_estampado = ""
               Me.txt_nombre_estampado = ""
            End If
         Else
            Me.txt_division = ""
            Me.txt_nombre_division = ""
            Me.txt_subdivision = ""
            Me.txt_nombre_subdivision = ""
            Me.txt_estampado = ""
            Me.txt_nombre_estampado = ""
         End If
         rs.Close
      End If
   Else
      MsgBox "No se a seleccionado un tipo de producto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub txt_dueño_licencia_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_dueño_licencia_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_dueño_licencia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_equivalente_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_equivalente_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_estampado_anverso_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_estampado_anverso_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_estampado_anverso_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_estampado_anverso_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_estampado_anverso) <> "" Then
      rs.Open "SELECT * FROM TB_ESTAMPADOS WHERE VCHA_EST_ESTAMPADO_ID = '" + txt_estampado_anverso + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_Estampado_anverso = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
      Else
         MsgBox "Clave de estampado incorrecta", vbOKOnly, "ATENCION"
         Me.txt_estampado_anverso = ""
         txt_nombre_Estampado_anverso = ""
      End If
      rs.Close
   Else
      txt_nombre_Estampado_anverso = ""
   End If
End Sub

Private Sub txt_estampado_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_ESTAMPADOS order by vcha_EST_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_EST_ESTAMPADO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ESTAMPADOS"
      var_tipo_lista = 103
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
End Sub

Private Sub txt_estampado_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_estampado_LostFocus()
   If Trim(txt_estampado) <> "" Then
      rs.Open "SELECT * FROM TB_ESTAMPADOS WHERE VCHA_EST_ESTAMPADO_ID = '" + Me.txt_estampado + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_estampado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
      Else
         Me.txt_estampado = ""
         Me.txt_nombre_estampado = ""
      End If
      rs.Close
   End If
End Sub

Private Sub txt_estampado_reverso_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_estampado_reverso_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_estampado_reverso_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_familia_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_familia_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_familia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_fecha_fin_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_fecha_fin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_fecha_inicio_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_fecha_inicio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_fundas_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_fundas_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_linea_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_linea_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_linea_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_maximo_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_maximo_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_minimo_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_minimo_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_caja_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_caja_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_caja_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Then
      KeyCode = 0
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_CAJAS order by vcha_caj_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_caj_caja_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CAJ_NOMBRE), "", rs!VCHA_CAJ_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CAJAS"
      var_tipo_lista = 22
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
      var_activa_forma_cajas = Me.Name
      frmarticulos2.Enabled = False
      frmcajas.Show
   End If
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_tipo_lista = 1 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_catalogo_inicio = lv_lista.selectedItem
            txt_nombre_catalogo_inicio = lv_lista.selectedItem.SubItems(1)
         Else
            txt_catalogo_inicio = ""
            txt_nombre_catalogo_inicio = ""
         End If
         txt_catalogo_inicio.SetFocus
      End If
      If var_tipo_lista = 2 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_catalogo_fin = lv_lista.selectedItem
            txt_nombre_catalogo_fin = lv_lista.selectedItem.SubItems(1)
         Else
            txt_catalogo_fin = ""
            txt_nombre_catalogo_fin = ""
         End If
         txt_catalogo_fin.SetFocus
      End If
      If var_tipo_lista = 3 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_dueño_licencia = lv_lista.selectedItem
            txt_nombre_dueño_licencia = lv_lista.selectedItem.SubItems(1)
         Else
            txt_dueño_licencia = ""
            txt_nombre_dueño_licencia = ""
         End If
         txt_dueño_licencia.SetFocus
      End If
      If var_tipo_lista = 4 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_familia = lv_lista.selectedItem
            txt_nombre_familia = lv_lista.selectedItem.SubItems(1)
         Else
            txt_familia = ""
            txt_nombre_familia = ""
         End If
         txt_familia.SetFocus
      End If
      If var_tipo_lista = 5 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_linea = lv_lista.selectedItem
            txt_nombre_linea = lv_lista.selectedItem.SubItems(1)
         Else
            txt_linea = ""
            txt_nombre_linea = ""
         End If
         txt_linea.SetFocus
      End If
      If var_tipo_lista = 6 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_sublinea = lv_lista.selectedItem
            txt_nombre_sublinea = lv_lista.selectedItem.SubItems(1)
         Else
            txt_sublinea = ""
            txt_nombre_sublinea = ""
         End If
         txt_sublinea.SetFocus
      End If
      If var_tipo_lista = 7 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_producto = lv_lista.selectedItem
            txt_nombre_producto = lv_lista.selectedItem.SubItems(1)
         Else
            txt_producto = ""
            txt_nombre_producto = ""
         End If
         txt_producto.SetFocus
      End If
      If var_tipo_lista = 8 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_tipo_producto = lv_lista.selectedItem
            txt_nombre_tipo_producto = lv_lista.selectedItem.SubItems(1)
         Else
            txt_tipo_producto = ""
            txt_nombre_tipo_producto = ""
         End If
         txt_tipo_producto.SetFocus
      End If
      If var_tipo_lista = 9 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_clase = lv_lista.selectedItem
            txt_nombre_clase = lv_lista.selectedItem.SubItems(1)
         Else
            txt_clase = ""
            txt_nombre_clase = ""
         End If
         txt_clase.SetFocus
      End If
      If var_tipo_lista = 10 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_estampado_anverso = lv_lista.selectedItem
            txt_nombre_Estampado_anverso = lv_lista.selectedItem.SubItems(1)
         Else
           txt_estampado_anverso = ""
           txt_nombre_Estampado_anverso = ""
         End If
         txt_estampado_anverso.SetFocus
      End If
      If var_tipo_lista = 11 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_tipo_estampado_anverso = lv_lista.selectedItem
            txt_nombre_tipo_estampado_anverso = lv_lista.selectedItem.SubItems(1)
         Else
            txt_tipo_estampado_anverso = ""
            txt_nombre_tipo_estampado_anverso = ""
         End If
         txt_tipo_estampado_anverso.SetFocus
      End If
      If var_tipo_lista = 12 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_estampado_reverso = lv_lista.selectedItem
            txt_nombre_estampado_reverso = lv_lista.selectedItem.SubItems(1)
         Else
            txt_estampado_reverso = ""
            txt_nombre_estampado_reverso = ""
         End If
         txt_estampado_reverso.SetFocus
      End If
      If var_tipo_lista = 13 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_tipo_estampado_reverso = lv_lista.selectedItem
            txt_nombre_tipo_estampado_reverso = lv_lista.selectedItem.SubItems(1)
         Else
            txt_tipo_estampado_reverso = ""
            txt_nombre_tipo_estampado_reverso = ""
         End If
         txt_tipo_estampado_reverso.SetFocus
      End If
      If var_tipo_lista = 14 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_color_anverso = lv_lista.selectedItem
            txt_nombre_color_anverso = lv_lista.selectedItem.SubItems(1)
         Else
            txt_color_anverso = ""
            txt_nombre_color_anverso = ""
         End If
         txt_color_anverso.SetFocus
      End If
      If var_tipo_lista = 15 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_color_reverso = lv_lista.selectedItem
            txt_nombre_color_reverso = lv_lista.selectedItem.SubItems(1)
         Else
            txt_color_reverso = ""
            txt_nombre_color_reverso = ""
         End If
         txt_color_reverso.SetFocus
      End If
      If var_tipo_lista = 16 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_tono_anverso = lv_lista.selectedItem
            txt_nombre_tono_anverso = lv_lista.selectedItem.SubItems(1)
         Else
            txt_tono_anverso = ""
            txt_nombre_tono_anverso = ""
         End If
         txt_tono_anverso.SetFocus
      End If
      If var_tipo_lista = 17 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_tono_reverso = lv_lista.selectedItem
            txt_nombre_tono_reverso = lv_lista.selectedItem.SubItems(1)
         Else
            txt_tono_reverso = ""
            txt_nombre_tono_reverso = ""
         End If
         txt_tono_reverso.SetFocus
      End If
      If var_tipo_lista = 18 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_uso_producto = lv_lista.selectedItem
            txt_nombre_uso_producto = lv_lista.selectedItem.SubItems(1)
         Else
            txt_uso_producto = ""
            txt_nombre_uso_producto = ""
         End If
         txt_uso_producto.SetFocus
      End If
      If var_tipo_lista = 19 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_subtipo_uso_producto = lv_lista.selectedItem
            txt_nombre_subtipo_uso_producto = lv_lista.selectedItem.SubItems(1)
         Else
            txt_subtipo_uso_producto = ""
            txt_nombre_subtipo_uso_producto = ""
         End If
         txt_subtipo_uso_producto.SetFocus
      End If
      If var_tipo_lista = 20 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_talla = lv_lista.selectedItem
            txt_nombre_talla = lv_lista.selectedItem.SubItems(1)
         Else
            txt_talla = ""
            txt_nombre_talla = ""
         End If
         txt_talla.SetFocus
      End If
      If var_tipo_lista = 21 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_unidad = lv_lista.selectedItem
            txt_nombre_unidad = lv_lista.selectedItem.SubItems(1)
         Else
            txt_unidad = ""
            txt_nombre_unidad = ""
         End If
         txt_unidad.SetFocus
      End If
      If var_tipo_lista = 22 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_caja = lv_lista.selectedItem
            txt_nombre_caja = lv_lista.selectedItem.SubItems(1)
         Else
            txt_caja = ""
            txt_nombre_caja = ""
         End If
         txt_caja.SetFocus
      End If
       If var_tipo_lista = 100 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_tipo = lv_lista.selectedItem
            txt_nombre_tipo = lv_lista.selectedItem.SubItems(1)
            txt_tipo.SetFocus
         End If
      End If
      If var_tipo_lista = 101 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_division = lv_lista.selectedItem
            txt_nombre_division = lv_lista.selectedItem.SubItems(1)
            txt_division.SetFocus
         End If
      End If
      If var_tipo_lista = 102 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_subdivision = lv_lista.selectedItem
            txt_nombre_subdivision = lv_lista.selectedItem.SubItems(1)
            txt_subdivision.SetFocus
         End If
      End If
      If var_tipo_lista = 103 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_estampado = lv_lista.selectedItem
            txt_nombre_estampado = lv_lista.selectedItem.SubItems(1)
            txt_estampado.SetFocus
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

Private Sub txt_caja_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_CAJAS order by vcha_caj_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_caj_caja_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CAJ_NOMBRE), "", rs!VCHA_CAJ_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CAJAS"
      var_tipo_lista = 22
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
      var_activa_forma_cajas = Me.Name
      frmarticulos2.Enabled = False
      frmcajas.Show
   End If
End Sub

Private Sub txt_caja_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_caja) <> "" Then
      rs.Open "select * from tb_cajas where vcha_Caj_caja_id = '" + txt_caja + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_caja = IIf(IsNull(rs!VCHA_CAJ_NOMBRE), "", rs!VCHA_CAJ_NOMBRE)
         var_ancho = IIf(IsNull(rs!floa_caj_ancho), 0, rs!floa_caj_ancho)
         var_alto = IIf(IsNull(rs!floa_Caj_alto), 0, rs!floa_Caj_alto)
         var_largo = IIf(IsNull(rs!floa_caj_largo), 0, rs!floa_caj_largo)
         If Val(txt_volumen_compreso) > 0 Then
            txt_piezas_caja = (var_ancho * var_alto * var_largo) / txt_volumen_compreso
         Else
            txt_piezas_caja = 0
         End If
      Else
         MsgBox "Clave de caja incorrecta", vbOKOnly, "ATENCION"
         txt_caja = ""
         txt_nombre_caja = ""
      End If
      rs.Close
   Else
      txt_nombre_caja = ""
   End If
End Sub

Private Sub txt_catalogo_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_catalogos order by vcha_cat_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cat_catalogo_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_Cat_nombre), "", rs!vcha_Cat_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CATALOGOS"
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
      var_activa_forma_catalogos = Me.Name
      frmarticulos2.Enabled = False
      frmcatalogos.Show
   End If
End Sub

Private Sub txt_catalogo_fin_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_catalogo_fin) <> "" Then
      rs.Open "select * from tb_catalogos where vcha_cat_catalogo_id = '" + txt_catalogo_fin + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_catalogo_fin = IIf(IsNull(rs!vcha_Cat_nombre), "", rs!vcha_Cat_nombre)
      Else
         MsgBox "Clave de catálogo incorrecta", vbOKOnly, "ATENCION"
         Me.txt_catalogo_fin = ""
         txt_nombre_catalogo_fin = ""
      End If
      rs.Close
   Else
      txt_nombre_catalogo_fin = ""
   End If
End Sub

Private Sub txt_catalogo_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_catalogos order by vcha_cat_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cat_catalogo_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_Cat_nombre), "", rs!vcha_Cat_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CATALOGOS"
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
      frmarticulos2.Enabled = False
      var_activa_forma_catalogos = Me.Name
      frmcatalogos.Show
   End If
End Sub

Private Sub txt_catalogo_inicio_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_catalogo_inicio) <> "" Then
      rs.Open "select * from tb_catalogos where vcha_cat_catalogo_id = '" + txt_catalogo_inicio + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_catalogo_inicio = IIf(IsNull(rs!vcha_Cat_nombre), "", rs!vcha_Cat_nombre)
      Else
         MsgBox "Clave de catálogo incorrecta", vbOKOnly, "ATENCION"
         Me.txt_catalogo_inicio = ""
         txt_nombre_catalogo_inicio = ""
      End If
      rs.Close
   Else
      txt_nombre_catalogo_inicio = ""
   End If
End Sub

Private Sub txt_clase_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_CLASEARTICULOS order by vcha_caR_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_Car_clase_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_Car_nombre), "", rs!vcha_Car_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CLASE ARTICULOS"
      var_tipo_lista = 9
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
      var_activa_forma_clases = Me.Name
      frmarticulos2.Enabled = False
      frmclases.Show
   End If
End Sub

Private Sub txt_clase_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_clase) <> "" Then
      rs.Open "SELECT * FROM TB_CLASEARTICULOS WHERE VCHA_CAR_CLASE_ID = '" + txt_clase + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_clase = IIf(IsNull(rs!vcha_Car_nombre), "", rs!vcha_Car_nombre)
      Else
         txt_nombre_clase = ""
         Me.txt_clase = ""
         MsgBox "Clave de clase incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      txt_nombre_clase = ""
   End If
End Sub

Private Sub txt_color_anverso_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_COLORES order by vcha_clr_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_CLR_COLOR_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLR_NOMBRE), "", rs!VCHA_CLR_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "COLORES"
      var_tipo_lista = 14
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
      var_activa_forma_colores = Me.Name
      frmarticulos2.Enabled = False
      frmcolores.Show
   End If
End Sub

Private Sub txt_color_anverso_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_color_anverso) <> "" Then
      rs.Open "SELECT * FROM TB_COLORES WHERE VCHA_CLR_COLOR_ID = '" + txt_color_anverso + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_color_anverso = IIf(IsNull(rs!VCHA_CLR_NOMBRE), "", rs!VCHA_CLR_NOMBRE)
      Else
         MsgBox "Clave de color incorrecta", vbOKOnly, "ATENCION"
         Me.txt_color_anverso = ""
         txt_nombre_color_anverso = ""
      End If
      rs.Close
   Else
      txt_nombre_color_anverso = ""
   End If
End Sub

Private Sub txt_color_reverso_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_COLORES order by vcha_CLR_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_CLR_COLOR_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLR_NOMBRE), "", rs!VCHA_CLR_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "COLORES"
      var_tipo_lista = 15
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
      var_activa_forma_colores = Me.Name
      frmarticulos2.Enabled = False
      frmcolores.Show
   End If
End Sub

Private Sub txt_color_reverso_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_color_reverso) <> "" Then
      rs.Open "SELECT * FROM TB_COLORES WHERE VCHA_CLR_COLOR_ID = '" + txt_color_reverso + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_color_reverso = IIf(IsNull(rs!VCHA_CLR_NOMBRE), "", rs!VCHA_CLR_NOMBRE)
      Else
         MsgBox "Clave de color incorrecta", vbOKOnly, "ATENCION"
         Me.txt_color_reverso = ""
         txt_nombre_color_reverso = ""
      End If
      rs.Close
   Else
      txt_nombre_color_reverso = ""
   End If
End Sub

Private Sub txt_dueño_licencia_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_licencias order by vcha_lic_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_LIC_LICENCIA_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_LIC_NOMBRE), "", rs!VCHA_LIC_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "LICENCIAS"
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
      var_activa_forma_licencias = Me.Name
      frmarticulos2.Enabled = False
      frmlicencias.Show
   End If
End Sub

Private Sub txt_dueño_licencia_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_dueño_licencia) <> "" Then
      rs.Open "select * from tb_licencias where vcha_lic_licencia_id = '" + txt_dueño_licencia + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_dueño_licencia = IIf(IsNull(rs!VCHA_LIC_NOMBRE), "", rs!VCHA_LIC_NOMBRE)
      Else
         MsgBox "Clave de licencia incorrecta", vbOKOnly, "ATENCION"
         Me.txt_dueño_licencia = ""
         txt_nombre_dueño_licencia = ""
      End If
      rs.Close
   Else
      txt_nombre_dueño_licencia = ""
   End If
End Sub

Private Sub txt_estampado_anverso_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_estampados order by vcha_est_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_EST_ESTAMPADO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ESTAMPADOS"
      var_tipo_lista = 10
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
      var_activa_forma_estampados = Me.Name
      frmarticulos2.Enabled = False
      frmestampados.Show
   End If
End Sub

Private Sub txt_estampado_reverso_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_estampados order by vcha_est_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_EST_ESTAMPADO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ESTAMPADOS"
      var_tipo_lista = 12
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
      var_activa_forma_estampados = Me.Name
      frmarticulos2.Enabled = False
      frmestampados.Show
   End If
End Sub

Private Sub txt_estampado_reverso_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_estampado_reverso) <> "" Then
      rs.Open "SELECT * FROM TB_ESTAMPADOS WHERE VCHA_EST_ESTAMPADO_ID = '" + txt_estampado_reverso + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_estampado_reverso = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
      Else
         MsgBox "Clave de estampado incorrecta", vbOKOnly, "ATENCION"
         Me.txt_estampado_reverso = ""
         txt_nombre_estampado_reverso = ""
      End If
      rs.Close
   Else
      txt_nombre_estampado_reverso = ""
   End If
End Sub

Private Sub txt_familia_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_diseños order by vcha_dis_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_DIS_DISEÑO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_dis_NOMBRE), "", rs!VCHA_dis_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "FAMILIAS"
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
      var_activa_forma_diseños = Me.Name
      frmarticulos2.Enabled = False
      frmdiseños.Show
   End If
End Sub

Private Sub txt_familia_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_familia) <> "" Then
      rs.Open "select * from tb_diseños where vcha_dis_diseño_id = '" + txt_familia + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_familia = IIf(IsNull(rs!VCHA_dis_NOMBRE), "", rs!VCHA_dis_NOMBRE)
      Else
         MsgBox "Clave de familia incorrecta", vbOKOnly, "ATENCION"
         Me.txt_familia = ""
         txt_nombre_familia = ""
      End If
      rs.Close
   Else
      txt_nombre_familia = ""
   End If
End Sub

Private Sub txt_linea_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_lineas order by vcha_lin_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_lin_linea_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_lin_NOMBRE), "", rs!VCHA_lin_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "LINEAS"
      var_tipo_lista = 5
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
      var_activa_forma_lineas = Me.Name
      frmarticulos2.Enabled = False
      frmlineas.Show
   End If
End Sub

Private Sub txt_linea_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_linea) <> "" Then
      rs.Open "SELECT * FROM TB_LINEAS WHERE VCHA_LIN_LINEA_ID = '" + txt_linea + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_linea = IIf(IsNull(rs!VCHA_lin_NOMBRE), "", rs!VCHA_lin_NOMBRE)
      Else
         MsgBox "Clave de linea incorrecta", vbOKOnly, "ATENCION"
         Me.txt_linea = ""
         txt_nombre_linea = ""
      End If
      rs.Close
   Else
      txt_nombre_linea = ""
   End If
End Sub


Private Sub txt_nombre_caja_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_caja_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_catalogo_fin_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_catalogo_fin_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_catalogo_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 37 Then
      Me.txt_catalogo_fin.SetFocus
   End If
   If KeyCode = 46 Then
      KeyCode = 0
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_catalogos order by vcha_cat_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cat_catalogo_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_Cat_nombre), "", rs!vcha_Cat_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CATALOGOS"
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
      var_activa_forma_catalogos = Me.Name
      frmarticulos2.Enabled = False
      frmcatalogos.Show
   End If
End Sub


Private Sub txt_nombre_catalogo_fin_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_catalogo_fin_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_catalogo_inicio_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_catalogo_inicio_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_catalogo_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 37 Then
      Me.txt_catalogo_inicio.SetFocus
   End If
   If KeyCode = 46 Then
      KeyCode = 0
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_catalogos order by vcha_cat_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cat_catalogo_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_Cat_nombre), "", rs!vcha_Cat_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CATALOGOS"
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
      frmarticulos2.Enabled = False
      var_activa_forma_catalogos = Me.Name
      frmcatalogos.Show
   End If
End Sub


Private Sub txt_nombre_catalogo_inicio_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_catalogo_inicio_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_clase_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_clase_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_clase_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Then
      KeyCode = 0
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_CLASEARTICULOS order by vcha_caR_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_Car_clase_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_Car_nombre), "", rs!vcha_Car_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CLASE ARTICULOS"
      var_tipo_lista = 9
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
      var_activa_forma_clases = Me.Name
      frmarticulos2.Enabled = False
      frmclases.Show
   End If
End Sub


Private Sub txt_nombre_clase_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_clase_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_color_anverso_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_color_anverso_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_color_anverso_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Then
      KeyCode = 0
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_COLORES order by vcha_clr_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_CLR_COLOR_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLR_NOMBRE), "", rs!VCHA_CLR_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "COLORES"
      var_tipo_lista = 14
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
      var_activa_forma_colores = Me.Name
      frmarticulos2.Enabled = False
      frmcolores.Show
   End If
End Sub


Private Sub txt_nombre_color_anverso_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_color_anverso_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_color_reverso_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_color_reverso_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_color_reverso_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Then
      KeyCode = 0
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_COLORES order by vcha_CLR_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_CLR_COLOR_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLR_NOMBRE), "", rs!VCHA_CLR_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "COLORES"
      var_tipo_lista = 15
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
      var_activa_forma_colores = Me.Name
      frmarticulos2.Enabled = False
      frmcolores.Show
   End If
End Sub


Private Sub txt_nombre_color_reverso_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_color_reverso_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_division_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 37 Then
      Me.txt_division.SetFocus
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_divisiones where vcha_tpr_tipo_producto_id = '" + Me.txt_tipo + "' order by vcha_DIV_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_div_division_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_div_nombre), "", rs!vcha_div_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "DIVISIONES"
      var_tipo_lista = 101
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
End Sub

Private Sub txt_nombre_division_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 13
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_dueño_licencia_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_dueño_licencia_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_dueño_licencia_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 37 Then
      Me.txt_dueño_licencia.SetFocus
   End If
   If KeyCode = 46 Then
      KeyCode = 0
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_licencias order by vcha_lic_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_LIC_LICENCIA_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_LIC_NOMBRE), "", rs!VCHA_LIC_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "LICENCIAS"
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
      var_activa_forma_licencias = Me.Name
      frmarticulos2.Enabled = False
      frmlicencias.Show
   End If
End Sub


Private Sub txt_nombre_dueño_licencia_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_dueño_licencia_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_Estampado_anverso_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_Estampado_anverso_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_estampado_anverso_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Then
      KeyCode = 0
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_estampados order by vcha_est_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_EST_ESTAMPADO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ESTAMPADOS"
      var_tipo_lista = 10
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
      var_activa_forma_estampados = Me.Name
      frmarticulos2.Enabled = False
      frmestampados.Show
   End If
End Sub


Private Sub txt_nombre_Estampado_anverso_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_Estampado_anverso_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_estampado_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 37 Then
      Me.txt_estampado.SetFocus
   End If
End Sub

Private Sub txt_nombre_estampado_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 13
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub

Private Sub txt_nombre_estampado_reverso_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_estampado_reverso_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_estampado_reverso_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Then
      KeyCode = 0
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_estampados order by vcha_est_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_EST_ESTAMPADO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ESTAMPADOS"
      var_tipo_lista = 12
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
      var_activa_forma_estampados = Me.Name
      frmarticulos2.Enabled = False
      frmestampados.Show
   End If
End Sub


Private Sub txt_nombre_estampado_reverso_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_estampado_reverso_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_familia_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_familia_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_familia_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 37 Then
      Me.txt_familia.SetFocus
   End If
   If KeyCode = 46 Then
      KeyCode = 0
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_diseños order by vcha_dis_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_DIS_DISEÑO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_dis_NOMBRE), "", rs!VCHA_dis_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "FAMILIAS"
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
      var_activa_forma_diseños = Me.Name
      frmarticulos2.Enabled = False
      frmdiseños.Show
   End If
End Sub


Private Sub txt_nombre_familia_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_familia_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_linea_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_linea_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_linea_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 37 Then
      Me.txt_linea.SetFocus
   End If
   If KeyCode = 46 Then
      KeyCode = 0
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_lineas order by vcha_lin_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_lin_linea_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_lin_NOMBRE), "", rs!VCHA_lin_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "LINEAS"
      var_tipo_lista = 5
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
      var_activa_forma_lineas = Me.Name
      frmarticulos2.Enabled = False
      frmlineas.Show
   End If
End Sub


Private Sub txt_nombre_linea_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_linea_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_producto_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_producto_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_producto_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Then
      KeyCode = 0
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_productos order by vcha_pro_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_pro_producto_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_PRO_NOMBRE), "", rs!VCHA_PRO_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "PRODUCTOS"
      var_tipo_lista = 7
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
      var_activa_forma_productos = Me.Name
      frmarticulos2.Enabled = False
      frmproductos.Show
   End If
End Sub


Private Sub txt_nombre_producto_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_producto_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_subdivision_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 37 Then
      Me.txt_subdivision.SetFocus
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_SUBDIVISIONES where vcha_tpr_tipo_producto_id = '" + Me.txt_tipo + "' AND VCHA_DIV_DIVISION_ID = '" + Me.txt_division + "' order by vcha_SUB_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_SUB_SUBDIVISION_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_sub_nombre), "", rs!vcha_sub_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "SUBDIVISIONES"
      var_tipo_lista = 102
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
End Sub

Private Sub txt_nombre_subdivision_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 13
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_sublinea_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_sublinea_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_sublinea_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 37 Then
      Me.txt_sublinea.SetFocus
   End If
   If KeyCode = 46 Then
      KeyCode = 0
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_sublineas order by vcha_sli_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_SLI_SUBLINEA_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_SLI_NOMBRE), "", rs!VCHA_SLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "SUBLINEAS"
      var_tipo_lista = 6
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
      var_activa_forma_sublineas = Me.Name
      frmarticulos2.Enabled = False
      frmsublineas.Show
   End If
End Sub


Private Sub txt_nombre_sublinea_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_sublinea_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_subtipo_uso_producto_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_subtipo_uso_producto_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_subtipo_uso_producto_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Then
      KeyCode = 0
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_SUBTIPOSUSOS order by vcha_sus_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_SUS_SUBTIPO_USO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_SUS_NOMBRE), "", rs!VCHA_SUS_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "SUBTIPO DE USOS"
      var_tipo_lista = 19
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
      var_activa_forma_subtiposusos = Me.Name
      frmarticulos2.Enabled = False
      frmsubtiposusos.Show
   End If
End Sub


Private Sub txt_nombre_subtipo_uso_producto_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_subtipo_uso_producto_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_talla_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_talla_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_talla_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 37 Then
      Me.txt_talla.SetFocus
   End If
   If KeyCode = 46 Then
      KeyCode = 0
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_TALLAS order by vcha_tal_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_TAL_TALLA_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_tal_NOMBRE), "", rs!VCHA_tal_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TALLAS"
      var_tipo_lista = 20
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
      var_activa_forma_tallas = Me.Name
      frmarticulos2.Enabled = False
      frmtallas.Show
   End If
End Sub


Private Sub txt_nombre_talla_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_talla_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_tipo_estampado_anverso_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_tipo_estampado_anverso_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_tipo_estampado_anverso_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Then
      KeyCode = 0
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_TIPOESTAMPADOS order by vcha_tes_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_TES_TIPOESTAMPADO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_TES_NOMBRE), "", rs!VCHA_TES_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TIPO ESTAMPADOS"
      var_tipo_lista = 11
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
      var_activa_forma_tipoestampados = Me.Name
      frmarticulos2.Enabled = False
      frmtipoestampados.Show
   End If
End Sub


Private Sub txt_nombre_tipo_estampado_anverso_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_tipo_estampado_anverso_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_tipo_estampado_reverso_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_tipo_estampado_reverso_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_tipo_estampado_reverso_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Then
      KeyCode = 0
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_TIPOESTAMPADOS order by vcha_tes_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_TES_TIPOESTAMPADO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_TES_NOMBRE), "", rs!VCHA_TES_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TIPO ESTAMPADOS"
      var_tipo_lista = 13
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
      var_activa_forma_tipoestampados = Me.Name
      frmarticulos2.Enabled = False
      frmtipoestampados.Show
   End If
End Sub


Private Sub txt_nombre_tipo_estampado_reverso_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_tipo_estampado_reverso_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_tipo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 37 Then
      Me.txt_tipo.SetFocus
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_TIPOs_productos order by vcha_tpr_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_tpr_tipo_producto_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_tpr_nombre), "", rs!vcha_tpr_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TIPO PRODUCTO"
      var_tipo_lista = 100
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
End Sub

Private Sub txt_nombre_tipo_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 13
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_tipo_producto_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_tipo_producto_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_tipo_producto_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Then
      KeyCode = 0
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_TIPOARTICULOS order by vcha_tar_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_TAR_TIPO_ARTICULO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_TAR_NOMBRE), "", rs!VCHA_TAR_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TIPO PRODUCTO"
      var_tipo_lista = 8
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
      var_activa_forma_tipoarticulos = Me.Name
      frmarticulos2.Enabled = False
      frmtipoarticulos.Show
   End If
End Sub


Private Sub txt_nombre_tipo_producto_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_tipo_producto_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_tono_anverso_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_tono_anverso_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_tono_anverso_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Then
      KeyCode = 0
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_TONOS order by vcha_ton_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_TON_TONO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_TON_NOMBRE), "", rs!VCHA_TON_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TONOS"
      var_tipo_lista = 16
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
      var_activa_forma_tonos = Me.Name
      frmarticulos2.Enabled = False
      frmtonos.Show
   End If
End Sub


Private Sub txt_nombre_tono_anverso_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_tono_anverso_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_tono_reverso_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_tono_reverso_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_tono_reverso_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Then
      KeyCode = 0
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_TONOS order by vcha_ton_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_TON_TONO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_TON_NOMBRE), "", rs!VCHA_TON_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TONOS"
      var_tipo_lista = 17
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
      var_activa_forma_tonos = Me.Name
      frmarticulos2.Enabled = False
      frmtonos.Show
   End If
End Sub

Private Sub txt_nombre_tono_reverso_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_tono_reverso_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_ubicacion_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_ubicacion_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_unidad_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_unidad_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_unidad_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 37 Then
      Me.txt_unidad.SetFocus
   End If
   If KeyCode = 46 Then
      KeyCode = 0
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_UNIDADES order by vcha_uni_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_uni_unidad_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_UNI_NOMBRE), "", rs!VCHA_UNI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "UNIDADES"
      var_tipo_lista = 21
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
      var_activa_forma_unidades = Me.Name
      frmarticulos2.Enabled = False
      frmunidades.Show
   End If
End Sub


Private Sub txt_nombre_unidad_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_unidad_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_uso_producto_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_uso_producto_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_uso_producto_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Then
      KeyCode = 0
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_usos order by vcha_uso_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_USO_USO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_USO_NOMBRE), "", rs!VCHA_USO_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "USOS"
      var_tipo_lista = 18
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
      var_activa_forma_usos = Me.Name
      frmarticulos2.Enabled = False
      frmusos.Show
   End If
End Sub

Private Sub cmd_catalogos_Click()
   frmcatalogos_canales.Show 1
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
      Call pro_elimina_articulos
      'rs.Open "select * from tb_articulos", cnn, adOpenDynamic, adLockOptimistic
      'If rs.BOF Then
      '   cmd_guardar.Enabled = False
      '   cmd_deshacer.Enabled = False
      '   cmd_eliminar.Enabled = False
      'Else
      '   cmd_guardar.Enabled = True
      '   cmd_deshacer.Enabled = True
      '   cmd_eliminar.Enabled = True
      'End If
      'rs.Close
   Else
      MsgBox "Imposible ejecutar la acción solicitada", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_guardar_Click()
Dim var_posible As Boolean
Dim var_posible_equivalente As Boolean
Dim var_posible_textilera As Boolean
   var_posible = True
   var_posible_textilera = True
   If var_empresa = "18" Then
      var_codigo = Trim(Me.txt_tipo) + Trim(Me.txt_division) + Trim(Me.txt_subdivision) + Trim(Me.txt_estampado)
      If var_codigo <> Mid(txt_codigo, 1, 10) Then
         var_posible_textilera = False
      Else
         var_posible_textilera = True
      End If
   End If
   If var_posible_textilera = True Then
      If var_modifica_registro_articulo = False Then
         rs.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_aRT_ARTICULO_ID = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_posible = False
         End If
         rs.Close
      End If
      If var_posible = True Then
         var_campos = 0
         var_falta = ""
         If txt_codigo = "" Then
            var_falta = "Clave"
            var_campos = 0
         End If
         If txt_descripcion = "" Then
            If Len(var_falta) = 0 Then
               var_falta = "Nombre"
            Else
               var_falta = var_falta + ", Nombre"
            End If
            var_campos = var_campos + 1
         End If
         If msk_precio = "" Then
            If Len(var_falta) = 0 Then
               var_falta = "Precio Base"
            Else
               var_falta = var_falta + ", Precio Base"
            End If
            var_campos = var_campos + 1
         End If
         If msk_costo = "" Then
            If Len(var_falta) = 0 Then
               var_falta = "Costo Estandar"
            Else
               var_falta = var_falta + ", Costo Estandar"
            End If
            var_campos = var_campos + 1
         End If
         If txt_fecha_fin = "" Then
            If Len(var_falta) = 0 Then
               var_falta = "Fecha Alta"
            Else
               var_falta = var_falta + ", Fecha Alta"
            End If
            var_campos = var_campos + 1
         End If
         If txt_catalogo_inicio = "" Then
            If Len(var_falta) = 0 Then
               var_falta = "Catálogo de Inicio"
            Else
               var_falta = var_falta + ", Catálogo de Inicio"
            End If
            var_campos = var_campos + 1
         End If
         If txt_catalogo_fin = "" Then
            If Len(var_falta) = 0 Then
                var_falta = "Ultimo Catálogo"
            Else
               var_falta = var_falta + ", Ultimo Catálogo"
            End If
            var_campos = var_campos + 1
         End If
         If txt_dueño_licencia = "" Then
            If Len(var_falta) = 0 Then
               var_falta = "Licencia"
            Else
               var_falta = var_falta + ", Licencia"
            End If
            var_campos = var_campos + 1
         End If
         If txt_contrato = "" Then
            If Len(var_falta) = 0 Then
               var_falta = "Contrato"
            Else
              var_falta = var_falta + ", Contrato"
            End If
            var_campos = var_campos + 1
         End If
         If txt_familia = "" Then
            If Len(var_falta) = 0 Then
               var_falta = "Familia"
            Else
               var_falta = var_falta + ", Familia"
            End If
            var_campos = var_campos + 1
         End If
         If txt_linea = "" Then
            If Len(var_falta) = 0 Then
               var_falta = "Linea"
            Else
               var_falta = var_falta + ", Linea"
            End If
            var_campos = var_campos + 1
         End If
         If txt_sublinea = "" Then
            If Len(var_falta) = 0 Then
               var_falta = "Sublinea"
            Else
               var_falta = var_falta + ", Sublinea"
            End If
            var_campos = var_campos + 1
         End If
         var_posible_equivalente = False
         If Trim(Me.txt_equivalente) <> "" Then
            var_posible_equivalente = True
         Else
            var_posible_equivalente = False
         End If
         'If txt_producto = "" Then
         '   If Len(var_falta) = 0 Then
         '      var_falta = "Producto"
         '   Else
         '      var_falta = var_falta + ", Producto"
         '   End If
         '   var_campos = var_campos + 1
         'End If
         'If txt_tipo_producto = "" Then
         '   If Len(var_falta) = 0 Then
         '      var_falta = "Tipo Producto"
         '   Else
         '      var_falta = var_falta + ", Tipo producto"
         '   End If
         '   var_campos = var_campos + 1
         'End If
         'If txt_clase = "" Then
         '   If Len(var_falta) = 0 Then
         '      var_falta = "Clase"
         '   Else
         '      var_falta = var_falta + ", Clase"
         '   End If
         '   var_campos = var_campos + 1
         'End If
         'If txt_estampado_anverso = "" Then
         '   If Len(var_falta) = 0 Then
         '      var_falta = "Estampado Anverso"
         '   Else
         '      var_falta = var_falta + ", Estampado Anverso"
         '   End If
         '     var_campos = var_campos + 1
         'End If
         'If txt_tipo_estampado_anverso = "" Then
         '   If Len(var_falta) = 0 Then
         '      var_falta = "Tipo Estampado Anverso"
         '   Else
         '      var_falta = var_falta + ", Tipo Estampado Anverso"
         '   End If
         '   var_campos = var_campos + 1
         'End If
         'If txt_estampado_reverso = "" Then
         '   If Len(var_falta) = 0 Then
         '      var_falta = "Estampado Reverso"
         '   Else
         '      var_falta = var_falta + ", Estampado Reverso"
         '   End If
         '   var_campos = var_campos + 1
         'End If
         'If txt_tipo_estampado_reverso = "" Then
         '   If Len(var_falta) = 0 Then
         '      var_falta = "Tipo Estampado Reverso"
         '   Else
         '      var_falta = var_falta + ", Tipo Estampado Reverso"
         '   End If
         '   var_campos = var_campos + 1
         'End If
         'If  txt_color_anverso = "" Then
         '   If Len(var_falta) = 0 Then
         '      var_falta = "Color Anverso"
         '   Else
         '      var_falta = var_falta + ", Color Anverso"
         '   End If
         '   var_campos = var_campos + 1
         'End If
         'If txt_color_reverso = "" Then
         '   If Len(var_falta) = 0 Then
         '       var_falta = "Color Reverso"
         '   Else
         '      var_falta = var_falta + ", Color Reverso"
         '   End If
         '   var_campos = var_campos + 1
         'End If
         'If txt_tono_anverso = "" Then
         '   If Len(var_falta) = 0 Then
         '      var_falta = "Tono Anverso"
         '   Else
         '      var_falta = var_falta + ", Tono Anverso"
         '   End If
         '   var_campos = var_campos + 1
         'End If
         'If txt_tono_reverso = "" Then
         '   If Len(var_falta) = 0 Then
         '      var_falta = "Tono Reverso"
         '   Else
         '      var_falta = var_falta + ", Tono Reverso"
         '   End If
         '   var_campos = var_campos + 1
         'End If
         'If txt_decorativos = "" Then
         '   If Len(var_falta) = 0 Then
         '      var_falta = "Decorativos"
         '   Else
         '      var_falta = var_falta + ", Decorativos"
         '   End If
         '   var_campos = var_campos + 1
         'End If
         'If txt_fundas = "" Then
         '   If Len(var_falta) = 0 Then
         '      var_falta = "Fundas"
         '   Else
         '      var_falta = var_falta + ", Fundas"
         '   End If
         '   var_campos = var_campos + 1
         'End If
         'If txt_uso_producto = "" Then
         '   If Len(var_falta) = 0 Then
         '      var_falta = "Uso del Producto"
         '   Else
         '      var_falta = var_falta + ", Uso del Producto"
         '   End If
         '   var_campos = var_campos + 1
         'End If
         'If txt_subtipo_uso_producto = "" Then
         '   If Len(var_falta) = 0 Then
         '      var_falta = "Subtipo de Uso"
         '   Else
         '      var_falta = var_falta + ", Subtipo de Uso"
         '   End If
         '   var_campos = var_campos + 1
         'End If
         If txt_talla = "" Then
            If Len(var_falta) = 0 Then
               var_falta = "Talla"
            Else
               var_falta = var_falta + ", Talla"
            End If
            var_campos = var_campos + 1
         End If
         If txt_unidad = "" Then
            If Len(var_falta) = 0 Then
               var_falta = "Unidades"
            Else
               var_falta = var_falta + ", Unidades"
            End If
            var_campos = var_campos + 1
         End If
         'If   txt_volumen = "" Then
         '   If Len(var_falta) = 0 Then
         '      var_falta = "Volumen"
         '   Else
         '      var_falta = var_falta + ", Volumen"
         '   End If
         '   var_campos = var_campos + 1
         'End If
         'If txt_tela = "" Then
         '   If Len(var_falta) = 0 Then
         '       var_falta = "Tela"
         '   Else
         '      var_falta = var_falta + ", Tela"
         '   End If
         '   var_campos = var_campos + 1
         'End If
         'If txt_composicion = "" Then
         '   If Len(var_falta) = 0 Then
         '      var_falta = "Composición"
         '   Else
         '      var_falta = var_falta + ", Composición"
         '   End If
         '   var_campos = var_campos + 1
         'End If
         'If txt_peso = "" Then
         '   If Len(var_falta) = 0 Then
         '      var_falta = "Peso"
         '  Else
         '      var_falta = var_falta + ", Peso"
         '   End If
         '   var_campos = var_campos + 1
         'End If
         'If txt_tara = "" Then
         '   If Len(var_falta) = 0 Then
         '      var_falta = "Tara"
         '   Else
         '      var_falta = var_falta + ", Tara"
         '   End If
         '   var_campos = var_campos + 1
         'End If
         'If txt_caja = "" Then
         '   If Len(var_falta) = 0 Then
         '      var_falta = "Caja"
         '   Else
         '      var_falta = var_falta + ", Caja"
         '   End If
         '   var_campos = var_campos + 1
         'End If
         'If txt_piezas_caja = "" Then
         '   If Len(var_falta) = 0 Then
         '      var_falta = "Piezas por Caja"
         '   Else
         '      var_falta = var_falta + ", Piezas por Caja"
         '   End If
         '   var_campos = var_campos + 1
         'End If
         'If txt_maximo = "" Then
         '   If Len(var_falta) = 0 Then
         '      var_falta = "Máximo"
         '   Else
         '      var_falta = var_falta + ", Máximo"
         '   End If
         '   var_campos = var_campos + 1
         'End If
         'If txt_minimo = "" Then
         '   If Len(var_falta) = 0 Then
         '      var_falta = "Mínimo"
         '   Else
         '      var_falta = var_falta + ", Mínimo"
         '   End If
         '   var_campos = var_campos + 1
         'End If
         'If txt_punto_reorden = "" Then
         '   If Len(var_falta) = 0 Then
         '      var_falta = "Punto de Reorden"
         '   Else
         '      var_falta = var_falta + ", Punto de Reorden"
         '   End If
         '   var_campos = var_campos + 1
         'End If
         'If txt_dias_inventario = "" Then
         '   If Len(var_falta) = 0 Then
         '      var_falta = "Dias de Inventario"
         '   Else
         '      var_falta = var_falta + ", Dias de Inventario"
         '   End If
         '   var_campos = var_campos + 1
         'End If
         'If txt_bulto = "" Then
         '   If Len(var_falta) = 0 Then
         '      var_falta = "Bulto"
         '   Else
         '      var_falta = var_falta + ", Bulto"
         '   End If
         '   var_campos = var_campos + 1
         'End If
         si = 6
         'If txt_codigo = "" Or txt_descripcion = "" Or msk_precio = "" Or msk_costo = "" Or txt_fecha_inicio = "" Or txt_catalogo_inicio = "" Or txt_catalogo_fin = "" Or txt_dueño_licencia = "" Or txt_contrato = "" Or _
         '   txt_familia = "" Or txt_linea = "" Or txt_sublinea = "" Or txt_producto = "" Or txt_tipo_producto = "" Or txt_clase = "" Or txt_estampado_anverso = "" Or txt_tipo_estampado_anverso = "" Or txt_estampado_reverso = "" Or txt_tipo_estampado_reverso = "" Or _
         '   txt_color_anverso = "" Or txt_color_reverso = "" Or txt_tono_anverso = "" Or txt_tono_reverso = "" Or txt_decorativos = "" Or txt_fundas = "" Or txt_uso_producto = "" Or txt_subtipo_uso_producto = "" Or txt_talla = "" Or txt_unidad = "" Or _
         '   txt_volumen = "" Or txt_tela = "" Or txt_composicion = "" Or txt_peso = "" Or txt_tara = "" Or txt_caja = "" Or txt_piezas_caja = "" Or txt_maximo = "" Or _
         '   txt_minimo = "" Or txt_punto_reorden = "" Or txt_dias_inventario = "" Or txt_bulto = "" Then
         If var_posible_equivalente = True Then
            If txt_codigo = "" Or txt_descripcion = "" Or msk_precio = "" Or msk_costo = "" Or txt_fecha_inicio = "" Or txt_catalogo_inicio = "" Or txt_catalogo_fin = "" Or txt_dueño_licencia = "" Or txt_contrato = "" Or _
               txt_familia = "" Or txt_linea = "" Or txt_sublinea = "" Or txt_talla = "" Or txt_unidad = "" Then
               If var_campos = 1 Then
                  si = MsgBox("Información incompleta, falta el siguiente registro: " + var_falta + ", ¿Deseas guardar la información de cualquier modo?", vbYesNo, "ATENCION")
               End If
               If var_campos > 1 Then
                  si = MsgBox("Información incompleta, faltan los siguientes registros: " + var_falta + ", ¿Deseas guardar la información de cualquier modo?", vbYesNo, "ATENCION")
               End If
            End If
            If si = 6 Then
               var_guarda_cambios = False
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
                  Call pro_guardar_articulos
                  'rs.Open "select * from tb_articulos", cnn, adOpenDynamic, adLockOptimistic
                  'If rs.BOF Then
                  '   cmd_guardar.Enabled = False
                  '   cmd_deshacer.Enabled = False
                  '   cmd_eliminar.Enabled = False
                  'Else
                  '   cmd_guardar.Enabled = True
                  '   cmd_deshacer.Enabled = True
                  '   cmd_eliminar.Enabled = True
                  'End If
                  'rs.Close
               Else
                  MsgBox "Imposible ejecutar la acción solicitada", vbOKOnly, "ATENCION"
               End If
            End If
         Else
            MsgBox "Debe de indicar el codigo equivalente", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Clave de artículo ya existe", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Los conceptos deben de ser iguales al codigo del artículo", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
   lv_almacenes.ListItems.Clear
   rs.Open "select * from tb_almacenes where char_alm_tipo = 'A' order by vcha_alm_nombre ", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set list_item = lv_almacenes.ListItems.Add(, , rs!VCHA_ALM_ALMACEN_ID)
         list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
         rs.MoveNext
   Wend
   rs.Close
   Dim var_n As Integer
   var_n = lv_almacenes.ListItems.Count
   If var_n > 6 Then
      lv_almacenes.ColumnHeaders(2).Width = 4270.71
   Else
      lv_almacenes.ColumnHeaders(2).Width = 4499.71
   End If
   frm_almacen.Visible = True
   lv_almacenes.SetFocus
End Sub

Private Sub cmd_nuevo_Click()
         'Call pro_limpiatextos(Me)
         Me.chk_numero_serie = 0
         var_guarda_cambios = True
         msk_precio = ""
         msk_costo = ""
         txt_decorativos = 0#
         txt_fundas = 0#
         txt_volumen = 0#
         txt_compresion = 0#
         txt_volumen_compreso = 0#
         txt_tela = 0#
         txt_composicion = 0#
         txt_peso = 0#
         txt_tara = 0#
         txt_piezas_caja = 0#
         txt_maximo = 0#
         txt_minimo = 0#
         txt_punto_reorden = 0#
         txt_dias_inventario = 0#
         txt_bulto = 0#
         txt_codigo.Enabled = True
         txt_codigo.SetFocus
         var_modifica_registro_articulo = False
         'Toolbar1.Buttons.Item(2).Enabled = True
         'Toolbar1.Buttons.Item(3).Enabled = True
         txt_codigo.Enabled = True
         txt_descripcion.Enabled = True
         msk_precio.Enabled = True
         msk_costo.Enabled = True
         txt_fecha_inicio.Enabled = True
         txt_fecha_fin.Enabled = True
         txt_fecha_fin = Date
         txt_catalogo_inicio.Enabled = True
         txt_catalogo_fin.Enabled = True
         txt_dueño_licencia.Enabled = True
         txt_contrato.Enabled = True
         txt_familia.Enabled = True
         txt_linea.Enabled = True
         txt_sublinea.Enabled = True
         txt_producto.Enabled = True
         txt_tipo_producto.Enabled = True
         txt_clase.Enabled = True
         txt_estampado_anverso.Enabled = True
         txt_tipo_estampado_anverso.Enabled = True
         txt_estampado_reverso.Enabled = True
         txt_tipo_estampado_reverso.Enabled = True
         txt_color_anverso.Enabled = True
         txt_color_reverso.Enabled = True
         txt_tono_anverso.Enabled = True
         txt_tono_reverso.Enabled = True
         txt_decorativos.Enabled = True
         txt_fundas.Enabled = True
         txt_uso_producto.Enabled = True
         txt_subtipo_uso_producto.Enabled = True
         txt_talla.Enabled = True
         txt_unidad.Enabled = True
         txt_volumen.Enabled = True
         txt_compresion.Enabled = True
         txt_volumen_compreso.Enabled = True
         txt_tela.Enabled = True
         txt_composicion.Enabled = True
         txt_peso.Enabled = True
         txt_tara.Enabled = True
         txt_caja.Enabled = True
         mes.Enabled = True
         txt_nombre_catalogo_inicio.Enabled = True
         txt_nombre_catalogo_fin.Enabled = True
         txt_nombre_dueño_licencia.Enabled = True
         txt_nombre_familia.Enabled = True
         txt_nombre_linea.Enabled = True
         txt_nombre_sublinea.Enabled = True
         txt_nombre_producto.Enabled = True
         txt_nombre_tipo_producto.Enabled = True
         txt_nombre_clase.Enabled = True
         txt_nombre_Estampado_anverso.Enabled = True
         txt_nombre_tipo_estampado_anverso.Enabled = True
         txt_nombre_estampado_reverso.Enabled = True
         txt_nombre_tipo_estampado_reverso.Enabled = True
         txt_nombre_color_anverso.Enabled = True
         txt_nombre_color_reverso.Enabled = True
         txt_nombre_tono_anverso.Enabled = True
         txt_nombre_tono_reverso.Enabled = True
         txt_nombre_uso_producto.Enabled = True
         txt_nombre_subtipo_uso_producto.Enabled = True
         txt_nombre_talla.Enabled = True
         txt_nombre_unidad.Enabled = True
         txt_nombre_caja.Enabled = True
         txt_nombre_ubicacion.Enabled = True
         Me.txt_equivalente.Enabled = True
         txt_equivalente = ""
         
         
         txt_codigo = ""
         txt_descripcion = ""
         msk_precio = ""
         msk_costo = ""
         txt_fecha_inicio = Date
         txt_fecha_fin = ""
         txt_familia = ""
         txt_linea = ""
         txt_sublinea = ""
         txt_producto = ""
         txt_tipo_producto = ""
         txt_clase = ""
         txt_estampado_anverso = ""
         txt_tipo_estampado_anverso = ""
         txt_estampado_reverso = ""
         txt_tipo_estampado_reverso = ""
         txt_color_anverso = ""
         txt_color_reverso = ""
         txt_tono_anverso = ""
         txt_tono_reverso = ""
         txt_decorativos = ""
         txt_fundas = ""
         txt_uso_producto = ""
         txt_subtipo_uso_producto = ""
         txt_talla = ""
         txt_unidad = ""
         txt_volumen = ""
         txt_compresion = ""
         txt_volumen_compreso = ""
         txt_tela = ""
         txt_composicion = ""
         txt_peso = ""
         txt_tara = ""
         txt_caja = ""
         mes.Enabled = True
         txt_nombre_familia = ""
         txt_nombre_linea = ""
         txt_nombre_sublinea = ""
         txt_nombre_producto = ""
         txt_nombre_tipo_producto = ""
         txt_nombre_clase = ""
         txt_nombre_Estampado_anverso = ""
         txt_nombre_tipo_estampado_anverso = ""
         txt_nombre_estampado_reverso = ""
         txt_nombre_tipo_estampado_reverso = ""
         txt_nombre_color_anverso = ""
         txt_nombre_color_reverso = ""
         txt_nombre_tono_anverso = ""
         txt_nombre_tono_reverso = ""
         txt_nombre_uso_producto = ""
         txt_nombre_subtipo_uso_producto = ""
         txt_nombre_talla = ""
         txt_nombre_unidad = ""
         txt_nombre_caja = ""
         txt_nombre_ubicacion = ""
         Me.txt_equivalente = ""
         txt_equivalente = ""
         
         
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_articulo = False Then
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
   If KeyCode = 116 Then
      'Call pro_limpiatextos(Me)
      'var_modifica_registro_articulo = False
      'Me.msk_costo = 0
      'Me.msk_precio = 0
   End If
   
   'If KeyCode = 117 Then
   '   var_modifica_registro_articulo = True
   '   var_codigo = Me.txt_codigo
   '   Call pro_limpiatextos(Me)
   '   Me.txt_codigo = var_codigo
   '   Me.txt_codigo.Enabled = False
   '   Me.txt_descripcion.SetFocus
   '   Me.msk_costo = 0
   '   Me.msk_precio = 0
   '   rs.Open "select * from tb_equivalencias where vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
   '   If Not rs.EOF Then
   '      Me.txt_equivalente = IIf(IsNull(rs!VCHA_EQU_CODIGO_EQUIVALENTE), "", rs!VCHA_EQU_CODIGO_EQUIVALENTE)
   '      Me.txt_equivalente.Enabled = False
   '   Else
   '      Me.txt_equivalente.Enabled = True
   '   End If
   '   rs.Close
   'End If
   
End Sub

Private Sub Form_Load()
   
   directorio = Dir("c:\reportes_sid")
   'MsgBox directorio
   'If directorio = "" Then
   '   MkDir ("c:\reportes_sid")
   'End If
   frm_almacen.Visible = False
   var_ventana = 0
   var_cadena_seguridad = ""
   Top = 0
   Left = 2500
   frm_lista.Visible = False
   txt_codigo.Text = ""
   txt_descripcion.Text = ""
   msk_precio.Text = ""
   msk_costo.Text = ""
   txt_fecha_inicio.Text = ""
   txt_fecha_fin.Text = ""
   txt_catalogo_inicio.Text = ""
   txt_catalogo_fin.Text = ""
   txt_dueño_licencia.Text = ""
   txt_contrato.Text = ""
   txt_familia.Text = ""
   txt_linea.Text = ""
   txt_sublinea.Text = ""
   txt_producto.Text = ""
   txt_tipo_producto.Text = ""
   txt_clase.Text = ""
   txt_estampado_anverso.Text = ""
   txt_tipo_estampado_anverso.Text = ""
   txt_estampado_reverso.Text = ""
   txt_tipo_estampado_reverso.Text = ""
   txt_color_anverso.Text = ""
   txt_color_reverso.Text = ""
   txt_tono_anverso.Text = ""
   txt_tono_reverso.Text = ""
   txt_decorativos.Text = ""
   txt_fundas.Text = ""
   txt_uso_producto.Text = ""
   txt_subtipo_uso_producto.Text = ""
   txt_talla.Text = ""
   txt_unidad.Text = ""
   txt_volumen.Text = ""
   txt_compresion.Text = ""
   txt_volumen_compreso.Text = ""
   txt_tela.Text = ""
   txt_composicion.Text = ""
   txt_peso.Text = ""
   txt_tara.Text = ""
   txt_caja.Text = ""
   txt_piezas_caja.Text = ""
   txt_maximo.Text = ""
   txt_minimo.Text = ""
   txt_punto_reorden.Text = ""
   txt_dias_inventario.Text = ""
   txt_ubicacion.Text = ""
   txt_bulto.Text = ""
   txt_nombre_catalogo_inicio.Text = ""
   txt_nombre_catalogo_fin.Text = ""
   txt_nombre_dueño_licencia.Text = ""
   txt_nombre_familia.Text = ""
   txt_nombre_linea.Text = ""
   txt_nombre_sublinea.Text = ""
   txt_nombre_producto.Text = ""
   txt_nombre_tipo_producto.Text = ""
   txt_nombre_clase.Text = ""
   txt_nombre_Estampado_anverso.Text = ""
   txt_nombre_tipo_estampado_anverso.Text = ""
   txt_nombre_estampado_reverso.Text = ""
   txt_nombre_tipo_estampado_reverso.Text = ""
   txt_nombre_color_anverso.Text = ""
   txt_nombre_color_reverso.Text = ""
   txt_nombre_tono_anverso.Text = ""
   txt_nombre_tono_reverso.Text = ""
   txt_nombre_uso_producto.Text = ""
   txt_nombre_subtipo_uso_producto.Text = ""
   txt_nombre_talla.Text = ""
   txt_nombre_unidad.Text = ""
   txt_nombre_caja.Text = ""
   txt_nombre_ubicacion.Text = ""
   var_despliega_menu = False
   'Call pro_llena_listview1
   'var_modifica_registro_articulo = True
   'lv_articulos.SmallIcons = ImageList1
   'pro_textos
   'rs.Open "select * from tb_articulos", cnn, adOpenDynamic, adLockOptimistic
   'If rs.BOF Then
   '   cmd_guardar.Enabled = False
   '   cmd_deshacer.Enabled = False
   '   cmd_eliminar.Enabled = False
   '   txt_codigo.Enabled = False
   '   txt_descripcion.Enabled = False
   '   msk_precio.Enabled = False
   '   msk_costo.Enabled = False
   '   txt_fecha_inicio.Enabled = False
   '   txt_fecha_fin.Enabled = False
   '   txt_catalogo_inicio.Enabled = False
   '   txt_catalogo_fin.Enabled = False
   '   txt_dueño_licencia.Enabled = False
   '   txt_contrato.Enabled = False
   '   txt_familia.Enabled = False
   '   txt_linea.Enabled = False
   '   txt_sublinea.Enabled = False
   '   txt_producto.Enabled = False
   '   txt_tipo_producto.Enabled = False
   '   txt_clase.Enabled = False
   '   txt_estampado_anverso.Enabled = False
   '   txt_tipo_estampado_anverso.Enabled = False
   '   txt_estampado_reverso.Enabled = False
   '   txt_tipo_estampado_reverso.Enabled = False
   '   txt_color_anverso.Enabled = False
   '   txt_color_reverso.Enabled = False
   '   txt_tono_anverso.Enabled = False
   '   txt_tono_reverso.Enabled = False
   '   txt_decorativos.Enabled = False
   '   txt_fundas.Enabled = False
   '   txt_uso_producto.Enabled = False
   '   txt_subtipo_uso_producto.Enabled = False
   '   txt_talla.Enabled = False
   '   txt_unidad.Enabled = False
   '   txt_volumen.Enabled = False
   '   txt_compresion.Enabled = False
   '   txt_volumen_compreso.Enabled = False
   '   txt_tela.Enabled = False
   '   txt_composicion.Enabled = False
   '   txt_peso.Enabled = False
   '   txt_tara.Enabled = False
   '   txt_caja.Enabled = False
   '   txt_piezas_caja.Enabled = False
   '   txt_maximo.Enabled = False
   '   txt_minimo.Enabled = False
   '   txt_punto_reorden.Enabled = False
   '   txt_dias_inventario.Enabled = False
   '   txt_ubicacion.Enabled = False
   '   txt_bulto.Enabled = False
   '   txt_nombre_catalogo_inicio.Enabled = False
   '   txt_nombre_catalogo_fin.Enabled = False
   '   txt_nombre_dueño_licencia.Enabled = False
   '   txt_nombre_familia.Enabled = False
   '   txt_nombre_linea.Enabled = False
   '   txt_nombre_sublinea.Enabled = False
   '   txt_nombre_producto.Enabled = False
   '   txt_nombre_tipo_producto.Enabled = False
   '   txt_nombre_clase.Enabled = False
   '   txt_nombre_Estampado_anverso.Enabled = False
   '   txt_nombre_tipo_estampado_anverso.Enabled = False
   '   txt_nombre_estampado_reverso.Enabled = False
   '   txt_nombre_tipo_estampado_reverso.Enabled = False
   '   txt_nombre_color_anverso.Enabled = False
   '   txt_nombre_color_reverso.Enabled = False
   '   txt_nombre_tono_anverso.Enabled = False
   '   txt_nombre_tono_reverso.Enabled = False
   '   txt_nombre_uso_producto.Enabled = False
   '   txt_nombre_subtipo_uso_producto.Enabled = False
   '   txt_nombre_talla.Enabled = False
   '   txt_nombre_unidad.Enabled = False
   '   txt_nombre_caja.Enabled = False
   '   txt_nombre_ubicacion.Enabled = False
   '   txt_codigo.Text = ""
   '   txt_descripcion.Text = ""
   '   msk_precio.Text = ""
   '   msk_costo.Text = ""
   '   txt_fecha_inicio.Text = ""
   '   txt_fecha_fin.Text = ""
   '   txt_catalogo_inicio.Text = ""
   '   txt_catalogo_fin.Text = ""
   '   txt_dueño_licencia.Text = ""
   '   txt_contrato.Text = ""
   '   txt_familia.Text = ""
   '   txt_linea.Text = ""
   '   txt_sublinea.Text = ""
   '   txt_producto.Text = ""
   '   txt_tipo_producto.Text = ""
   '   txt_clase.Text = ""
   '   txt_estampado_anverso.Text = ""
   '   txt_tipo_estampado_anverso.Text = ""
   '   txt_estampado_reverso.Text = ""
   '   txt_tipo_estampado_reverso.Text = ""
   '   txt_color_anverso.Text = ""
   '   txt_color_reverso.Text = ""
   '   txt_tono_anverso.Text = ""
   '   txt_tono_reverso.Text = ""
   '   txt_decorativos.Text = ""
   '   txt_fundas.Text = ""
   '   txt_uso_producto.Text = ""
   '   txt_subtipo_uso_producto.Text = ""
   '   txt_talla.Text = ""
   '   txt_unidad.Text = ""
   '   txt_volumen.Text = ""
   '   txt_compresion.Text = ""
   '   txt_volumen_compreso.Text = ""
   '   txt_tela.Text = ""
   '   txt_composicion.Text = ""
   '   txt_peso.Text = ""
   '   txt_tara.Text = ""
   '   txt_caja.Text = ""
   '   txt_piezas_caja.Text = ""
   '   txt_maximo.Text = ""
   '   txt_minimo.Text = ""
   '   txt_punto_reorden.Text = ""
   '   txt_dias_inventario.Text = ""
   '   txt_ubicacion.Text = ""
   '   txt_bulto.Text = ""
   '   txt_nombre_catalogo_inicio.Text = ""
   '   txt_nombre_catalogo_fin.Text = ""
   '   txt_nombre_dueño_licencia.Text = ""
   '   txt_nombre_familia.Text = ""
   '   txt_nombre_linea.Text = ""
   '   txt_nombre_sublinea.Text = ""
   '   txt_nombre_producto.Text = ""
   '   txt_nombre_tipo_producto.Text = ""
   '   txt_nombre_clase.Text = ""
   '   txt_nombre_Estampado_anverso.Text = ""
   '   txt_nombre_tipo_estampado_anverso.Text = ""
   '   txt_nombre_estampado_reverso.Text = ""
   '   txt_nombre_tipo_estampado_reverso.Text = ""
   '   txt_nombre_color_anverso.Text = ""
   '   txt_nombre_color_reverso.Text = ""
   '   txt_nombre_tono_anverso.Text = ""
   '   txt_nombre_tono_reverso.Text = ""
   '   txt_nombre_uso_producto.Text = ""
   '   txt_nombre_subtipo_uso_producto.Text = ""
   '   txt_nombre_talla.Text = ""
   '   txt_nombre_unidad.Text = ""
   '   txt_nombre_caja.Text = ""
   '   txt_nombre_ubicacion.Text = ""
   'Else
   '   cmd_guardar.Enabled = True
   '   cmd_deshacer.Enabled = True
   '   cmd_eliminar.Enabled = True
   '   txt_codigo.Enabled = False
   '   txt_descripcion.Enabled = True
   '   msk_precio.Enabled = True
   '   msk_costo.Enabled = True
   '   txt_fecha_inicio.Enabled = True
   '   txt_fecha_fin.Enabled = True
   '   txt_catalogo_inicio.Enabled = True
   '   txt_catalogo_fin.Enabled = True
   '   txt_dueño_licencia.Enabled = True
   '   txt_contrato.Enabled = True
   '   txt_familia.Enabled = True
   '   txt_linea.Enabled = True
   '   txt_sublinea.Enabled = True
   '   txt_producto.Enabled = True
   '   txt_tipo_producto.Enabled = True
   '   txt_clase.Enabled = True
   '   txt_estampado_anverso.Enabled = True
   '   txt_tipo_estampado_anverso.Enabled = True
   '   txt_estampado_reverso.Enabled = True
   '   txt_tipo_estampado_reverso.Enabled = True
   '   txt_color_anverso.Enabled = True
   '   txt_color_reverso.Enabled = True
   '   txt_tono_anverso.Enabled = True
   '   txt_tono_reverso.Enabled = True
   '   txt_decorativos.Enabled = True
   '   txt_fundas.Enabled = True
   '   txt_uso_producto.Enabled = True
   '   txt_subtipo_uso_producto.Enabled = True
   '   txt_talla.Enabled = True
   '   txt_unidad.Enabled = True
   '   txt_volumen.Enabled = True
   '   txt_compresion.Enabled = True
   '   txt_volumen_compreso.Enabled = True
   '   txt_tela.Enabled = True
   '   txt_composicion.Enabled = True
   '   txt_peso.Enabled = True
   '   txt_tara.Enabled = True
   '   txt_caja.Enabled = True
   '   Me.cmd_fecha_fin.Enabled = True
   '   Me.cmd_fecha_inicio.Enabled = True
   '   txt_nombre_catalogo_inicio.Enabled = True
   '   txt_nombre_catalogo_fin.Enabled = True
   '   txt_nombre_dueño_licencia.Enabled = True
   '   txt_nombre_familia.Enabled = True
   '   txt_nombre_linea.Enabled = True
   '   txt_nombre_sublinea.Enabled = True
   '   txt_nombre_producto.Enabled = True
   '   txt_nombre_tipo_producto.Enabled = True
   '   txt_nombre_clase.Enabled = True
   '   txt_nombre_Estampado_anverso.Enabled = True
   '   txt_nombre_tipo_estampado_anverso.Enabled = True
   '   txt_nombre_estampado_reverso.Enabled = True
   '   txt_nombre_tipo_estampado_reverso.Enabled = True
   '   txt_nombre_color_anverso.Enabled = True
   '   txt_nombre_color_reverso.Enabled = True
   '   txt_nombre_tono_anverso.Enabled = True
   '   txt_nombre_tono_reverso.Enabled = True
   '   txt_nombre_uso_producto.Enabled = True
   '   txt_nombre_subtipo_uso_producto.Enabled = True
   '   txt_nombre_talla.Enabled = True
   '   txt_nombre_unidad.Enabled = True
   '   txt_nombre_caja.Enabled = True
   '   txt_nombre_ubicacion.Enabled = True
   'End If
   'rs.Close
   mes.Visible = False
   var_guarda_cambios = False
   Frmmenu2.StatusBar1.Panels(1).Text = ""
   Me.txt_codigo.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   var_swpassword = False
   var_modifica_registro_articulo = False
   Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub lv_articulos_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Set lv_articulos.selectedItem = Item
   pro_textos
   var_guarda_cambios = False
End Sub




Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_articulos.SetFocus
      Call pro_avanzar(Me, lv_articulos, Button)
      lv_articulos.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 1 Then
      If lv_articulos.ListItems.Count > 0 Then
         lv_articulos.ListItems(1).Selected = True
         pro_textos
         lv_articulos.ListItems(1).EnsureVisible
      End If
   End If
   If Button.Index = 4 Then
      numero_items_articulos = lv_articulos.ListItems.Count
      lv_articulos.ListItems(numero_items_articulos).Selected = True
      lv_articulos.selectedItem.EnsureVisible
      pro_textos
   End If
err0:
End Sub



Sub pro_guardar_articulos()
   Dim ok As Boolean
   Set TB_Articulos = New TB_Articulos
   Set TB_BITACORA_ARTICULOS = New TB_BITACORA_ARTICULOS
   ok = True
   If txt_codigo <> "" And txt_descripcion <> "" And msk_precio <> "" And msk_costo <> "" Then
         rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         ok = TB_Articulos.Anadir(txt_codigo, txt_descripcion, msk_precio, msk_costo, txt_fecha_fin, txt_fecha_inicio, txt_catalogo_inicio, txt_catalogo_fin, txt_dueño_licencia, txt_contrato, _
              txt_familia, txt_linea, txt_sublinea, txt_producto, txt_tipo_producto, txt_clase, txt_estampado_anverso, txt_tipo_estampado_anverso, txt_estampado_reverso, txt_tipo_estampado_reverso, _
              txt_color_anverso, txt_color_reverso, txt_tono_anverso, txt_tono_reverso, txt_decorativos, txt_fundas, txt_uso_producto, txt_subtipo_uso_producto, txt_talla, txt_unidad, _
              txt_volumen, txt_tela, txt_composicion, txt_peso, txt_tara, txt_caja, txt_piezas_caja, txt_maximo, _
              txt_minimo, txt_punto_reorden, txt_dias_inventario, txt_ubicacion, txt_bulto, chk_salida_masiva.Value)
              If ok Then
                 rsaux2.Open "update tb_articulos set inte_art_detenido = " + CStr(chk_detenido.Value) + ", VCHA_ART_codigo_EXTERNO = '" + Me.txt_equivalente + "', VCHA_TPR_TIPO_PRODUCTO_ID = '" + Me.txt_tipo + "', VCHA_DIV_DIVISION_ID = '" + Me.txt_division + "', VCHA_SUB_SUBDIVISION_ID = '" + Me.txt_subdivision + "', VCHA_EST_ESTAMPADO_ID = '" + Me.txt_estampado + "', inte_art_numero_Serie = " + CStr(Me.chk_numero_serie) + ", VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' where vcha_art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                 bitacora = True
                 cnn.BeginTrans
                 If var_modifica_registro_articulo = False Then
                    var_operacion_bitacora = "I"
                    bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "VCHA_ART_NOMBRE", var_operacion_bitacora, "", txt_descripcion, var_clave_usuario_global, fun_NombrePc, Date)
                    rsaux.Open "UPDATE TB_ARTICULOS SET NUM_INTER_TRANC_TYPE = 1, NUM_INTER_UPLOADED = 1 where VCHA_ART_ARTICULO_ID = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                 Else
                    rsaux.Open "UPDATE TB_ARTICULOS SET NUM_INTER_TRANC_TYPE = 2, NUM_INTER_UPLOADED = 2 where VCHA_ART_ARTICULO_ID = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                    var_operacion_bitacora = "M"
                    If rs(0) <> txt_codigo Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "VCHA_ART_ARTICULO_ID", var_operacion_bitacora, rs(0), txt_codigo, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(1) <> txt_descripcion Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "VCHA_ART_NOMBRE", var_operacion_bitacora, rs(1), txt_descripcion, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(2) <> msk_precio Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "VCHA_ART_PRECIO", var_operacion_bitacora, rs(2), msk_precio, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(3) <> msk_costo Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "VCHA_ART_COSTO", var_operacion_bitacora, rs(3), msk_costo, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(4) <> txt_fecha_inicio Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "DTIM_ART_FECHA_ALTA", var_operacion_bitacora, rs(4), txt_fecha_inicio, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(5) <> txt_fecha_fin Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "DTIM_ART_FECHA_BAJA", var_operacion_bitacora, rs(5), txt_fecha_fin, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(6) <> txt_catalogo_inicio Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "VCHA_ART_CATALOGO_INICIO", var_operacion_bitacora, rs(6), txt_catalogo_inicio, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(7) <> txt_catalogo_fin Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "VCHA_ART_CATALOGO_VIGENTE", var_operacion_bitacora, rs(7), txt_catalogo_fin, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(8) <> txt_dueño_licencia Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "VCHA_LIC_LICENCIA_ID", var_operacion_bitacora, rs(8), txt_dueño_licencia, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(9) <> txt_contrato Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "VCHA_ART_NUMERO_LIC", var_operacion_bitacora, rs(9), txt_contrato, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(10) <> txt_familia Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "VCHA_DIS_DISEÑO_ID", var_operacion_bitacora, rs(10), txt_familia, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(11) <> txt_linea Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "VCHA_LIN_LINEA_ID", var_operacion_bitacora, rs(11), txt_linea, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(12) <> txt_sublinea Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "VCHA_SLI_SUBLINEA_ID", var_operacion_bitacora, rs(12), txt_sublinea, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(13) <> txt_producto Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "VCHA_PRO_PRODUCTO_ID", var_operacion_bitacora, rs(13), txt_producto, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(14) <> txt_tipo_producto Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "VCHA_TAR_TIPO_ARTICULO_ID", var_operacion_bitacora, rs(14), txt_tipo_producto, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(15) <> txt_clase Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "VCHA_CAR_CLASE_ID", var_operacion_bitacora, rs(15), txt_clase, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(16) <> txt_estampado_anverso Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "VCHA_ART_ESTAMPADO1", var_operacion_bitacora, rs(16), txt_estampado_anverso, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(17) <> txt_tipo_estampado_anverso Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "VCHA_ART_TIPO_ESTAMPADO1", var_operacion_bitacora, rs(17), txt_tipo_estampado_anverso, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(18) <> txt_estampado_reverso Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "VCHA_ART_ESTAMPADO2", var_operacion_bitacora, rs(18), txt_estampado_reverso, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(19) <> txt_tipo_estampado_reverso Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "VCHA_ART_TIPO_ESTAMPADO2", var_operacion_bitacora, rs(19), txt_tipo_estampado_reverso, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(20) <> txt_color_anverso Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "VCHA_ART_COLOR1", var_operacion_bitacora, rs(20), txt_color_anverso, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(21) <> txt_color_reverso Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "VCHA_ART_COLOR2", var_operacion_bitacora, rs(21), txt_color_reverso, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(22) <> txt_tono_anverso Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "VCHA_ART_TONO1", var_operacion_bitacora, rs(22), txt_tono_anverso, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(23) <> txt_tono_reverso Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "VCHA_ART_TONO2", var_operacion_bitacora, rs(23), txt_tono_reverso, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(24) <> txt_decorativos Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "INTE_ART_NUMERO_DECORATIVOS", var_operacion_bitacora, rs(24), txt_decorativos, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(25) <> txt_fundas Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "INTE_ART_FUNDAS", var_operacion_bitacora, rs(25), txt_fundas, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(26) <> txt_uso_producto Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "VCHA_USO_USO_ID", var_operacion_bitacora, rs(26), txt_uso_producto, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(27) <> txt_subtipo_uso_producto Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "VCHA_SUS_SUBTIPO_USO_ID", var_operacion_bitacora, rs(27), txt_subtipo_uso_producto, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(28) <> txt_talla Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "VCHA_TAL_TALLA_ID", var_operacion_bitacora, rs(28), txt_talla, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(29) <> txt_unidad Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "VCHA_UNI_UNIDAD_ID", var_operacion_bitacora, rs(29), txt_unidad, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(30) <> txt_volumen Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "FLOA_ART_VOLUMEN", var_operacion_bitacora, rs(30), txt_volumen, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(31) <> txt_tela Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "FLOA_ART_TELA", var_operacion_bitacora, rs(31), txt_tela, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                 
                    If rs(32) <> txt_composicion Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "VCHA_ART_COMPOSICION", var_operacion_bitacora, rs(32), txt_composicion, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(33) <> txt_peso Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "FLOA_ART_PESO", var_operacion_bitacora, rs(33), txt_peso, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(34) <> txt_tara Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "FLOA_ART_TARA", var_operacion_bitacora, rs(34), txt_tara, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(35) <> txt_caja Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "VCHA_CAJ_CAJA_ID", var_operacion_bitacora, rs(35), txt_caja, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(36) <> txt_piezas_caja Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "FLOA_ART_PIEZAS_CAJA", var_operacion_bitacora, rs(36), txt_piezas_caja, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(37) <> txt_maximo Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "FLOA_ART_MAXIMO", var_operacion_bitacora, rs(37), txt_maximo, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                 
                    If rs(38) <> txt_minimo Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "FLOA_ART_MINIMO", var_operacion_bitacora, rs(32), txt_minimo, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(39) <> txt_punto_reorden Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "FLOA_ART_PUNTO_REORDEN", var_operacion_bitacora, rs(33), txt_punto_reorden, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(40) <> txt_dias_inventario Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "FLOA_ART_DIAS_INVENTARIO", var_operacion_bitacora, rs(34), txt_dias_inventario, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(41) <> txt_ubicacion Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "VCHA_UBI_UNICACION_ID", var_operacion_bitacora, rs(35), txt_ubicacion, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(42) <> txt_bulto Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "FLOA_ART_BULTO", var_operacion_bitacora, rs(42), txt_bulto, var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                    If rs(43) <> chk_salida_masiva Then
                       bitacora = TB_BITACORA_ARTICULOS.Anadir(txt_codigo, "INTE_ART_SALIDA_MASIVA", var_operacion_bitacora, CStr(IIf(IsNull(rs(37)), 0, rs(37))), CStr(chk_salida_masiva), var_clave_usuario_global, fun_NombrePc, Date)
                    End If
                 End If
                 rs.Close
                 rs.Open "SELECT * FROM TB_EQUIVALENCIAS WHERE VCHA_ART_ARTICULO_ID = '" + txt_codigo + "' AND VCHA_EQU_CODIGO_EQUIVALENTE = '" + txt_equivalente + "'", cnn, adOpenDynamic, adLockOptimistic
                 If rs.EOF Then
                    rsaux.Open "INSERT INTO TB_EQUIVALENCIAS (VCHA_ART_ARTICULO_ID, VCHA_EQU_CODIGO_EQUIVALENTE) VALUES ('" + Me.txt_codigo + "','" + txt_equivalente + "')", cnn, adOpenDynamic, adLockOptimistic
                 End If
                 rs.Close
                 cnn.CommitTrans
                 'pro_actualiza_ListView
                 txt_codigo.Enabled = False
                 MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
                 txt_registros = lv_articulos.ListItems.Count
                 var_modifica_registro_articulo = True
             Else
                 MsgBox "No se puede grabar registro: " + TB_Articulos.MensajeError, vbOKOnly + vbCritical, "ATENCION"
             End If
      Else
         MsgBox "Imposible guardar el registro debido a que la información esta incompleta", vbOKOnly + vbCritical, "ATENCION"
      End If
      Set TB_Articulos = Nothing: var_hubo_cambios = False: var_hubo_cambios_2 = False
End Sub

Sub pro_elimina_articulos()
   Dim var_llave_usuarios As String
   Set TB_Articulos = New TB_Articulos
   On Error GoTo salir:
   If txt_codigo <> "" And txt_descripcion <> "" And var_modifica_registro_articulo = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_Articulos.Eliminar(txt_codigo)
      Else
         GoTo salir:
      End If
      If ok Then
         MsgBox "Se Elimino Correctamente el Registro", vbInformation
         lv_articulos.ListItems.Remove (lv_articulos.selectedItem.Index)
         numero_items_articulos = numero_items_articulos - 1
         Call pro_limpiatextos(Me)
         txt_registros = lv_articulos.ListItems.Count
         lv_articulos.selectedItem.Selected = True
         pro_textos
      Else
         MsgBox "No se puede grabar registro: " + TB_Articulos.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
   Set TB_Articulos = Nothing
End Sub


Sub pro_llena_listview1()
Dim list_item As ListItem
Dim var_numero_items_articulos2 As Double
Dim var_contador_articulos As Integer
Dim var_porcentaje As Variant
   numero_items_articulos = 0
   rs.Open "select count(*) from tb_articulos", cnn, adOpenDynamic, adLockOptimistic
   var_contador_articulos = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
   rs.Close
   rs.Open "select * from vw_catalogo_articulos order by vcha_art_nombre_Español", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
      Frmmenu2.StatusBar1.Panels.Item(1).Text = "Cargando Catálogo de Artículos " + Str(var_porcentaje - 1) + " %. Favor de Esperar."
      Frmmenu2.Refresh
      Set list_item = lv_articulos.ListItems.Add(, , rs!vcha_Art_articulo_id)
      list_item.SubItems(1) = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
      list_item.SubItems(2) = IIf(IsNull(rs!mone_art_precio_base), 0, rs!mone_art_precio_base)
      list_item.SubItems(3) = IIf(IsNull(rs!mone_Art_costo_estandar), 0, rs!mone_Art_costo_estandar)
      list_item.SubItems(4) = IIf(IsNull(rs!DTIM_ART_FECHA_BAJA), "", rs!DTIM_ART_FECHA_BAJA)
      list_item.SubItems(5) = IIf(IsNull(rs!DTIM_ART_FECHA_alta), "", rs!DTIM_ART_FECHA_alta)
      list_item.SubItems(6) = IIf(IsNull(rs!CLAVE_CATALOGO_INICIO), "", rs!CLAVE_CATALOGO_INICIO)
      list_item.SubItems(7) = IIf(IsNull(rs!CATALOGO_INICIO), "", rs!CATALOGO_INICIO)
      list_item.SubItems(8) = IIf(IsNull(rs!CLAVE_CATALOGO_FINAL), "", rs!CLAVE_CATALOGO_FINAL)
      list_item.SubItems(9) = IIf(IsNull(rs!CATALOGO_FINAL), "", rs!CATALOGO_FINAL)
      list_item.SubItems(10) = IIf(IsNull(rs!VCHA_LIC_LICENCIA_ID), "", rs!VCHA_LIC_LICENCIA_ID)
      list_item.SubItems(11) = IIf(IsNull(rs!VCHA_LIC_NOMBRE), "", rs!VCHA_LIC_NOMBRE)
      list_item.SubItems(12) = IIf(IsNull(rs!VCHA_ART_NUMERO_LIC), "", rs!VCHA_ART_NUMERO_LIC)
      list_item.SubItems(13) = IIf(IsNull(rs!VCHA_DIS_DISEÑO_ID), "", rs!VCHA_DIS_DISEÑO_ID)
      list_item.SubItems(14) = IIf(IsNull(rs!VCHA_dis_NOMBRE), "", rs!VCHA_dis_NOMBRE)
      list_item.SubItems(15) = IIf(IsNull(rs!vcha_lin_linea_id), "", rs!vcha_lin_linea_id)
      list_item.SubItems(16) = IIf(IsNull(rs!VCHA_lin_NOMBRE), "", rs!VCHA_lin_NOMBRE)
      list_item.SubItems(17) = IIf(IsNull(rs!VCHA_SLI_SUBLINEA_ID), "", rs!VCHA_SLI_SUBLINEA_ID)
      list_item.SubItems(18) = IIf(IsNull(rs!VCHA_SLI_NOMBRE), "", rs!VCHA_SLI_NOMBRE)
      'list_item.SubItems(19) = IIf(IsNull(rs!VCHA_PRO_PRODUCTO_ID), "", rs!VCHA_PRO_PRODUCTO_ID)
      'list_item.SubItems(20) = IIf(IsNull(rs!VCHA_PRO_NOMBRE), "", rs!VCHA_PRO_NOMBRE)
      'list_item.SubItems(21) = IIf(IsNull(rs!VCHA_TAR_TIPO_ARTICULO_ID), "", rs!VCHA_TAR_TIPO_ARTICULO_ID)
      'list_item.SubItems(22) = IIf(IsNull(rs!VCHA_TAR_NOMBRE), "", rs!VCHA_TAR_NOMBRE)
      'list_item.SubItems(23) = IIf(IsNull(rs!vcha_car_clase_id), "", rs!vcha_car_clase_id)
      'list_item.SubItems(24) = IIf(IsNull(rs!vcha_car_nombre), "", rs!vcha_car_nombre)
      'list_item.SubItems(25) = IIf(IsNull(rs!VCHA_ART_ESTAMPADO1), "", rs!VCHA_ART_ESTAMPADO1)
      'list_item.SubItems(26) = IIf(IsNull(rs!NOMBRE_ESATAMPADO_1), "", rs!NOMBRE_ESATAMPADO_1)
      'list_item.SubItems(27) = IIf(IsNull(rs!VCHA_ART_TIPO_ESTAMPADO1), "", rs!VCHA_ART_TIPO_ESTAMPADO1)
      'list_item.SubItems(28) = IIf(IsNull(rs!NOMBRE_TIPO_ESTAMPADO_1), "", rs!NOMBRE_TIPO_ESTAMPADO_1)
      'list_item.SubItems(29) = IIf(IsNull(rs!VCHA_ART_ESTAMPADO2), "", rs!VCHA_ART_ESTAMPADO2)
      'list_item.SubItems(30) = IIf(IsNull(rs!NOMBRE_ESTAMPADO_2), "", rs!NOMBRE_ESTAMPADO_2)
      'list_item.SubItems(31) = IIf(IsNull(rs!VCHA_ART_TIPO_ESTAMPADO2), "", rs!VCHA_ART_TIPO_ESTAMPADO2)
      'list_item.SubItems(32) = IIf(IsNull(rs!NOMBRE_TIPO_ESTAMPADO_2), "", rs!NOMBRE_TIPO_ESTAMPADO_2)
      'list_item.SubItems(33) = IIf(IsNull(rs!VCHA_ART_COLOR1), "", rs!VCHA_ART_COLOR1)
      'list_item.SubItems(34) = IIf(IsNull(rs!NOMBRE_COLOR_1), "", rs!NOMBRE_COLOR_1)
      'list_item.SubItems(35) = IIf(IsNull(rs!VCHA_ART_COLOR2), "", rs!VCHA_ART_COLOR2)
      'list_item.SubItems(36) = IIf(IsNull(rs!NOMBRE_COLOR_2), "", rs!NOMBRE_COLOR_2)
      'list_item.SubItems(37) = IIf(IsNull(rs!VCHA_ART_TONO1), "", rs!VCHA_ART_TONO1)
      'list_item.SubItems(38) = IIf(IsNull(rs!NOMBRE_TONO1), "", rs!NOMBRE_TONO1)
      'list_item.SubItems(39) = IIf(IsNull(rs!VCHA_ART_TONO2), "", rs!VCHA_ART_TONO2)
      'list_item.SubItems(40) = IIf(IsNull(rs!NOMBRE_TONO_2), "", rs!NOMBRE_TONO_2)
      'list_item.SubItems(41) = IIf(IsNull(rs!INTE_ART_NUMERO_DECORATIVOS), 0, rs!INTE_ART_NUMERO_DECORATIVOS)
      'list_item.SubItems(42) = IIf(IsNull(rs!INTE_ART_FUNDAS), 0, rs!INTE_ART_FUNDAS)
      'list_item.SubItems(43) = IIf(IsNull(rs!VCHA_USO_USO_ID), "", rs!VCHA_USO_USO_ID)
      'list_item.SubItems(44) = IIf(IsNull(rs!VCHA_USO_NOMBRE), "", rs!VCHA_USO_NOMBRE)
      'list_item.SubItems(45) = IIf(IsNull(rs!VCHA_SUS_SUBTIPO_USO_ID), "", rs!VCHA_SUS_SUBTIPO_USO_ID)
      'list_item.SubItems(46) = IIf(IsNull(rs!VCHA_SUS_NOMBRE), "", rs!VCHA_SUS_NOMBRE)
      list_item.SubItems(47) = IIf(IsNull(rs!VCHA_TAL_TALLA_ID), "", rs!VCHA_TAL_TALLA_ID)
      list_item.SubItems(48) = IIf(IsNull(rs!VCHA_tal_NOMBRE), "", rs!VCHA_tal_NOMBRE)
      list_item.SubItems(49) = IIf(IsNull(rs!vcha_uni_unidad_id), "", rs!vcha_uni_unidad_id)
      list_item.SubItems(50) = IIf(IsNull(rs!VCHA_UNI_NOMBRE), "", rs!VCHA_UNI_NOMBRE)
      'list_item.SubItems(51) = IIf(IsNull(rs!FLOA_ART_VOLUMEN), 0, rs!FLOA_ART_VOLUMEN)
      'list_item.SubItems(52) = IIf(IsNull(rs!FLOA_ART_TELA), 0, rs!FLOA_ART_TELA)
      'list_item.SubItems(53) = IIf(IsNull(rs!VCHA_ART_COMPOSICION), "", rs!VCHA_ART_COMPOSICION)
      'list_item.SubItems(54) = IIf(IsNull(rs!FLOA_ART_PESO), 0, rs!FLOA_ART_PESO)
      'list_item.SubItems(55) = IIf(IsNull(rs!FLOA_ART_TARA), 0, rs!FLOA_ART_TARA)
      'list_item.SubItems(56) = IIf(IsNull(rs!VCHA_CAJ_CAJA_ID), "", rs!VCHA_CAJ_CAJA_ID)
      'list_item.SubItems(57) = IIf(IsNull(rs!vcha_caj_nombre), "", rs!vcha_caj_nombre)
      'list_item.SubItems(58) = IIf(IsNull(rs!FLOA_ART_PIEZAS_CAJA), 0, rs!FLOA_ART_PIEZAS_CAJA)
      'list_item.SubItems(59) = IIf(IsNull(rs!FLOA_ART_MAXIMO), 0, rs!FLOA_ART_MAXIMO)
      'list_item.SubItems(60) = IIf(IsNull(rs!FLOA_ART_MINIMO), 0, rs!FLOA_ART_MINIMO)
      'list_item.SubItems(61) = IIf(IsNull(rs!FLOA_ART_PUNTO_REORDEN), 0, rs!FLOA_ART_PUNTO_REORDEN)
      'list_item.SubItems(62) = IIf(IsNull(rs!FLOA_ART_DIAS_INVENTARIO), 0, rs!FLOA_ART_DIAS_INVENTARIO)
      'list_item.SubItems(63) = IIf(IsNull(rs!VCHA_UBI_UNICACION_ID), "", rs!VCHA_UBI_UNICACION_ID)
      list_item.SubItems(64) = ""
      'list_item.SubItems(65) = IIf(IsNull(rs!FLOA_ART_BULTO), 0, rs!FLOA_ART_BULTO)
      list_item.SubItems(66) = IIf(IsNull(rs!inte_Art_salida_masiva), 0, rs!inte_Art_salida_masiva)
      list_item.SubItems(67) = IIf(IsNull(rs!INTE_ART_detenido), 0, rs!INTE_ART_detenido)
      rs.MoveNext:
      numero_items_articulos = numero_items_articulos + 1
      var_numero_items_articulos2 = numero_items_articulos
      var_porcentaje = Round((var_numero_items_articulos2 * 100) / var_contador_articulos)
   Wend
   'If numero_items_articulos > 11 Then
   '   lv_articulos.ColumnHeaders(2).Width = 8900
   'Else
   '   lv_articulos.ColumnHeaders(2).Width = 9000
   'End If
   rs.Close
End Sub

Sub limpia_textos_2()
         Me.chk_numero_serie = 0
         var_guarda_cambios = True
         msk_precio = ""
         msk_costo = ""
         txt_decorativos = 0#
         txt_fundas = 0#
         txt_volumen = 0#
         txt_compresion = 0#
         txt_volumen_compreso = 0#
         txt_tela = 0#
         txt_composicion = 0#
         txt_peso = 0#
         txt_tara = 0#
         txt_piezas_caja = 0#
         txt_maximo = 0#
         txt_minimo = 0#
         txt_punto_reorden = 0#
         txt_dias_inventario = 0#
         txt_bulto = 0#
         txt_codigo.Enabled = True
         var_modifica_registro_articulo = False
         txt_codigo.Enabled = True
         txt_descripcion.Enabled = True
         msk_precio.Enabled = True
         msk_costo.Enabled = True
         txt_fecha_inicio.Enabled = True
         txt_fecha_fin.Enabled = True
         txt_fecha_fin = Date
         txt_catalogo_inicio.Enabled = True
         txt_catalogo_fin.Enabled = True
         txt_dueño_licencia.Enabled = True
         txt_contrato.Enabled = True
         txt_familia.Enabled = True
         txt_linea.Enabled = True
         txt_sublinea.Enabled = True
         txt_producto.Enabled = True
         txt_tipo_producto.Enabled = True
         txt_clase.Enabled = True
         txt_estampado_anverso.Enabled = True
         txt_tipo_estampado_anverso.Enabled = True
         txt_estampado_reverso.Enabled = True
         txt_tipo_estampado_reverso.Enabled = True
         txt_color_anverso.Enabled = True
         txt_color_reverso.Enabled = True
         txt_tono_anverso.Enabled = True
         txt_tono_reverso.Enabled = True
         txt_decorativos.Enabled = True
         txt_fundas.Enabled = True
         txt_uso_producto.Enabled = True
         txt_subtipo_uso_producto.Enabled = True
         txt_talla.Enabled = True
         txt_unidad.Enabled = True
         txt_volumen.Enabled = True
         txt_compresion.Enabled = True
         txt_volumen_compreso.Enabled = True
         txt_tela.Enabled = True
         txt_composicion.Enabled = True
         txt_peso.Enabled = True
         txt_tara.Enabled = True
         txt_caja.Enabled = True
         mes.Enabled = True
         txt_nombre_catalogo_inicio.Enabled = True
         txt_nombre_catalogo_fin.Enabled = True
         txt_nombre_dueño_licencia.Enabled = True
         txt_nombre_familia.Enabled = True
         txt_nombre_linea.Enabled = True
         txt_nombre_sublinea.Enabled = True
         txt_nombre_producto.Enabled = True
         txt_nombre_tipo_producto.Enabled = True
         txt_nombre_clase.Enabled = True
         txt_nombre_Estampado_anverso.Enabled = True
         txt_nombre_tipo_estampado_anverso.Enabled = True
         txt_nombre_estampado_reverso.Enabled = True
         txt_nombre_tipo_estampado_reverso.Enabled = True
         txt_nombre_color_anverso.Enabled = True
         txt_nombre_color_reverso.Enabled = True
         txt_nombre_tono_anverso.Enabled = True
         txt_nombre_tono_reverso.Enabled = True
         txt_nombre_uso_producto.Enabled = True
         txt_nombre_subtipo_uso_producto.Enabled = True
         txt_nombre_talla.Enabled = True
         txt_nombre_unidad.Enabled = True
         txt_nombre_caja.Enabled = True
         txt_nombre_ubicacion.Enabled = True
         Me.txt_equivalente.Enabled = True
         txt_equivalente = ""
         
         
         txt_codigo = ""
         txt_descripcion = ""
         msk_precio = ""
         msk_costo = ""
         txt_fecha_inicio = Date
         txt_fecha_fin = ""
         txt_familia = ""
         txt_linea = ""
         txt_sublinea = ""
         txt_producto = ""
         txt_tipo_producto = ""
         txt_clase = ""
         txt_estampado_anverso = ""
         txt_tipo_estampado_anverso = ""
         txt_estampado_reverso = ""
         txt_tipo_estampado_reverso = ""
         txt_color_anverso = ""
         txt_color_reverso = ""
         txt_tono_anverso = ""
         txt_tono_reverso = ""
         txt_decorativos = ""
         txt_fundas = ""
         txt_uso_producto = ""
         txt_subtipo_uso_producto = ""
         txt_talla = ""
         txt_unidad = ""
         txt_volumen = ""
         txt_compresion = ""
         txt_volumen_compreso = ""
         txt_tela = ""
         txt_composicion = ""
         txt_peso = ""
         txt_tara = ""
         txt_caja = ""
         mes.Enabled = True
         txt_nombre_familia = ""
         txt_nombre_linea = ""
         txt_nombre_sublinea = ""
         txt_nombre_producto = ""
         txt_nombre_tipo_producto = ""
         txt_nombre_clase = ""
         txt_nombre_Estampado_anverso = ""
         txt_nombre_tipo_estampado_anverso = ""
         txt_nombre_estampado_reverso = ""
         txt_nombre_tipo_estampado_reverso = ""
         txt_nombre_color_anverso = ""
         txt_nombre_color_reverso = ""
         txt_nombre_tono_anverso = ""
         txt_nombre_tono_reverso = ""
         txt_nombre_uso_producto = ""
         txt_nombre_subtipo_uso_producto = ""
         txt_nombre_talla = ""
         txt_nombre_unidad = ""
         txt_nombre_caja = ""
         txt_nombre_ubicacion = ""
         Me.txt_equivalente = ""
         txt_equivalente = ""
End Sub

Sub llena_registros()
'On Error GoTo err0:
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   rsaux.Open "select * from vw_catalogo_Articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rsaux.EOF Then
      txt_nombre_catalogo_inicio = ""
      txt_nombre_catalogo_fin = ""
      txt_nombre_dueño_licencia = ""
      txt_nombre_familia = ""
      txt_nombre_linea = ""
      txt_nombre_sublinea = ""
      txt_nombre_producto = ""
      txt_nombre_tipo_producto = ""
      txt_nombre_clase = ""
      txt_nombre_Estampado_anverso = ""
      txt_nombre_tipo_estampado_anverso = ""
      txt_nombre_estampado_reverso = ""
      txt_nombre_tipo_estampado_reverso = ""
      txt_nombre_color_anverso = ""
      txt_nombre_color_reverso = ""
      txt_nombre_tono_anverso = ""
      txt_nombre_tono_reverso = ""
      txt_nombre_uso_producto = ""
      txt_nombre_subtipo_uso_producto = ""
      txt_nombre_talla = ""
      txt_nombre_unidad = ""
      Me.chk_detenido = 0
      txt_descripcion = IIf(IsNull(rsaux!vcha_art_nombre_español), "", rsaux!vcha_art_nombre_español)
      msk_precio = IIf(IsNull(rsaux!mone_art_precio_base), 0, rsaux!mone_art_precio_base)
      msk_costo = IIf(IsNull(rsaux!mone_Art_costo_estandar), 0, rsaux!mone_Art_costo_estandar)
      txt_fecha_inicio = IIf(IsNull(rsaux!DTIM_ART_FECHA_alta), "", rsaux!DTIM_ART_FECHA_alta)
      txt_fecha_fin = IIf(IsNull(rsaux!DTIM_ART_FECHA_BAJA), "", rsaux!DTIM_ART_FECHA_BAJA)
      txt_catalogo_inicio = IIf(IsNull(rsaux!CLAVE_CATALOGO_INICIO), "", rsaux!CLAVE_CATALOGO_INICIO)
      txt_nombre_catalogo_inicio = IIf(IsNull(rsaux!CATALOGO_INICIO), "", rsaux!CATALOGO_INICIO)
      txt_catalogo_fin = IIf(IsNull(rsaux!CLAVE_CATALOGO_FINAL), "", rsaux!CLAVE_CATALOGO_FINAL)
      txt_nombre_catalogo_fin = IIf(IsNull(rsaux!CATALOGO_FINAL), "", rsaux!CATALOGO_FINAL)
      txt_dueño_licencia = IIf(IsNull(rsaux!VCHA_LIC_LICENCIA_ID), "", rsaux!VCHA_LIC_LICENCIA_ID)
      txt_nombre_dueño_licencia = IIf(IsNull(rsaux!VCHA_LIC_NOMBRE), "", rsaux!VCHA_LIC_NOMBRE)
      txt_contrato = IIf(IsNull(rsaux!VCHA_ART_NUMERO_LIC), "", rsaux!VCHA_ART_NUMERO_LIC)
      txt_familia = IIf(IsNull(rsaux!VCHA_DIS_DISEÑO_ID), "", rsaux!VCHA_DIS_DISEÑO_ID)
      txt_nombre_familia = IIf(IsNull(rsaux!VCHA_dis_NOMBRE), "", rsaux!VCHA_dis_NOMBRE)
      txt_linea = IIf(IsNull(rsaux!vcha_lin_linea_id), "", rsaux!vcha_lin_linea_id)
      txt_nombre_linea = IIf(IsNull(rsaux!VCHA_lin_NOMBRE), "", rsaux!VCHA_lin_NOMBRE)
      txt_sublinea = IIf(IsNull(rsaux!VCHA_SLI_SUBLINEA_ID), "", rsaux!VCHA_SLI_SUBLINEA_ID)
      txt_nombre_sublinea = IIf(IsNull(rsaux!VCHA_SLI_NOMBRE), "", rsaux!VCHA_SLI_NOMBRE)
      txt_talla = IIf(IsNull(rsaux!VCHA_TAL_TALLA_ID), "", rsaux!VCHA_TAL_TALLA_ID)
      txt_nombre_talla = IIf(IsNull(rsaux!VCHA_tal_NOMBRE), "", rsaux!VCHA_tal_NOMBRE)
      txt_unidad = IIf(IsNull(rsaux!vcha_uni_unidad_id), "", rsaux!vcha_uni_unidad_id)
      txt_nombre_unidad = IIf(IsNull(rsaux!VCHA_UNI_NOMBRE), "", rsaux!VCHA_UNI_NOMBRE)
      txt_nombre_ubicacion = ""
      chk_salida_masiva.Value = IIf(IsNull(rsaux!inte_Art_salida_masiva), 0, rsaux!inte_Art_salida_masiva)
      chk_detenido.Value = IIf(IsNull(rsaux!INTE_ART_detenido), 0, rsaux!INTE_ART_detenido)
      txt_nombre_ubicacion = ""
      Me.txt_tipo = IIf(IsNull(rsaux!VCHA_tpr_tipo_producto_ID), "", rsaux!VCHA_tpr_tipo_producto_ID)
      Me.txt_nombre_tipo = IIf(IsNull(rsaux!vcha_tpr_nombre), "", rsaux!vcha_tpr_nombre)
      Me.txt_division = IIf(IsNull(rsaux!vcha_div_division_id), "", rsaux!vcha_div_division_id)
      Me.txt_nombre_division = IIf(IsNull(rsaux!vcha_div_nombre), "", rsaux!vcha_div_nombre)
      Me.txt_subdivision = IIf(IsNull(rsaux!VCHA_SUB_SUBDIVISION_ID), "", rsaux!VCHA_SUB_SUBDIVISION_ID)
      Me.txt_nombre_subdivision = IIf(IsNull(rsaux!vcha_sub_nombre), "", rsaux!vcha_sub_nombre)
      Me.txt_estampado = IIf(IsNull(rsaux!VCHA_EST_ESTAMPADO_ID), "", rsaux!VCHA_EST_ESTAMPADO_ID)
      Me.txt_nombre_estampado = IIf(IsNull(rsaux!vcha_est_nombre), "", rsaux!vcha_est_nombre)
      Me.chk_numero_serie = IIf(IsNull(rsaux!inte_art_numero_serie), 0, rsaux!inte_art_numero_serie)
      Me.txt_peso = IIf(IsNull(rsaux!floa_art_peso), 0, rsaux!floa_art_peso)
      Me.txt_volumen = IIf(IsNull(rsaux!floa_art_volumen), 0, rsaux!floa_art_volumen)
      rsaux2.Open "SELECT * FROM TB_EQUIVALENCIAS WHERE VCHA_ART_ARTICULO_ID = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux2.EOF Then
         txt_equivalente = IIf(IsNull(rsaux2!vcha_equ_codigo_equivalente), "", rsaux2!vcha_equ_codigo_equivalente)
         txt_equivalente.Enabled = False
      Else
         txt_equivalente = ""
         txt_equivalente.Enabled = True
      End If
      rsaux2.Close
      var_guarda_cambios = False
      var_hubo_cambios = False
      var_modifica_registro_articulo = True
   Else
      txt_nombre_catalogo_inicio = ""
      txt_nombre_catalogo_fin = ""
      txt_nombre_dueño_licencia = ""
      txt_nombre_familia = ""
      txt_nombre_linea = ""
      txt_nombre_sublinea = ""
      txt_nombre_producto = ""
      txt_nombre_tipo_producto = ""
      txt_nombre_clase = ""
      txt_nombre_Estampado_anverso = ""
      txt_nombre_tipo_estampado_anverso = ""
      txt_nombre_estampado_reverso = ""
      txt_nombre_tipo_estampado_reverso = ""
      txt_nombre_color_anverso = ""
      txt_nombre_color_reverso = ""
      txt_nombre_tono_anverso = ""
      txt_nombre_tono_reverso = ""
      txt_nombre_uso_producto = ""
      txt_nombre_subtipo_uso_producto = ""
      txt_nombre_talla = ""
      txt_nombre_unidad = ""
      Me.chk_detenido = 0
   End If
   rsaux.Close
err0:

End Sub


Sub pro_textos()
'On Error GoTo err0:
   txt_nombre_catalogo_inicio = ""
   txt_nombre_catalogo_fin = ""
   txt_nombre_dueño_licencia = ""
   txt_nombre_familia = ""
   txt_nombre_linea = ""
   txt_nombre_sublinea = ""
   txt_nombre_producto = ""
   txt_nombre_tipo_producto = ""
   txt_nombre_clase = ""
   txt_nombre_Estampado_anverso = ""
   txt_nombre_tipo_estampado_anverso = ""
   txt_nombre_estampado_reverso = ""
   txt_nombre_tipo_estampado_reverso = ""
   txt_nombre_color_anverso = ""
   txt_nombre_color_reverso = ""
   txt_nombre_tono_anverso = ""
   txt_nombre_tono_reverso = ""
   txt_nombre_uso_producto = ""
   txt_nombre_subtipo_uso_producto = ""
   txt_nombre_talla = ""
   txt_nombre_unidad = ""
    Me.chk_detenido = 0
     
   txt_codigo = lv_articulos.selectedItem
   txt_descripcion = lv_articulos.selectedItem.SubItems(1)
   msk_precio = lv_articulos.selectedItem.SubItems(2)
   msk_costo = lv_articulos.selectedItem.SubItems(3)
   txt_fecha_inicio = lv_articulos.selectedItem.SubItems(4)
   txt_fecha_fin = lv_articulos.selectedItem.SubItems(5)
   txt_catalogo_inicio = lv_articulos.selectedItem.SubItems(6)
   txt_nombre_catalogo_inicio = lv_articulos.selectedItem.SubItems(7)
   txt_catalogo_fin = lv_articulos.selectedItem.SubItems(8)
   txt_nombre_catalogo_fin = lv_articulos.selectedItem.SubItems(9)
   txt_dueño_licencia = lv_articulos.selectedItem.SubItems(10)
   txt_nombre_dueño_licencia = lv_articulos.selectedItem.SubItems(11)
   txt_contrato = lv_articulos.selectedItem.SubItems(12)
   txt_familia = lv_articulos.selectedItem.SubItems(13)
   txt_nombre_familia = lv_articulos.selectedItem.SubItems(14)
   txt_linea = lv_articulos.selectedItem.SubItems(15)
   txt_nombre_linea = lv_articulos.selectedItem.SubItems(16)
   txt_sublinea = lv_articulos.selectedItem.SubItems(17)
   txt_nombre_sublinea = lv_articulos.selectedItem.SubItems(18)
   txt_producto = lv_articulos.selectedItem.SubItems(19)
   txt_nombre_producto = lv_articulos.selectedItem.SubItems(20)
   txt_tipo_producto = lv_articulos.selectedItem.SubItems(21)
   txt_nombre_tipo_producto = lv_articulos.selectedItem.SubItems(22)
   txt_clase = lv_articulos.selectedItem.SubItems(23)
   txt_nombre_clase = lv_articulos.selectedItem.SubItems(24)
   txt_estampado_anverso = lv_articulos.selectedItem.SubItems(25)
   txt_nombre_Estampado_anverso = lv_articulos.selectedItem.SubItems(26)
   txt_tipo_estampado_anverso = lv_articulos.selectedItem.SubItems(27)
   txt_nombre_tipo_estampado_anverso = lv_articulos.selectedItem.SubItems(28)
   txt_estampado_reverso = lv_articulos.selectedItem.SubItems(29)
   txt_nombre_estampado_reverso = lv_articulos.selectedItem.SubItems(30)
   txt_tipo_estampado_reverso = lv_articulos.selectedItem.SubItems(31)
   txt_nombre_tipo_estampado_reverso = lv_articulos.selectedItem.SubItems(32)
   txt_color_anverso = lv_articulos.selectedItem.SubItems(33)
   txt_nombre_color_anverso = lv_articulos.selectedItem.SubItems(34)
   txt_color_reverso = lv_articulos.selectedItem.SubItems(35)
   txt_nombre_color_reverso = lv_articulos.selectedItem.SubItems(36)
   txt_tono_anverso = lv_articulos.selectedItem.SubItems(37)
   txt_nombre_tono_anverso = lv_articulos.selectedItem.SubItems(38)
   txt_tono_renverso = lv_articulos.selectedItem.SubItems(39)
   txt_nombre_tono_reverso = lv_articulos.selectedItem.SubItems(40)
   txt_decorativos = lv_articulos.selectedItem.SubItems(41)
   txt_fundas = lv_articulos.selectedItem.SubItems(42)
   txt_uso_producto = lv_articulos.selectedItem.SubItems(43)
   txt_nombre_uso_producto = lv_articulos.selectedItem.SubItems(44)
   txt_subtipo_uso_producto = lv_articulos.selectedItem.SubItems(45)
   txt_nombre_subtipo_uso_producto = lv_articulos.selectedItem.SubItems(46)
   txt_talla = lv_articulos.selectedItem.SubItems(47)
   txt_nombre_talla = lv_articulos.selectedItem.SubItems(48)
   txt_unidad = lv_articulos.selectedItem.SubItems(49)
   txt_nombre_unidad = lv_articulos.selectedItem.SubItems(50)
   txt_volumen = lv_articulos.selectedItem.SubItems(51)
   txt_tela = lv_articulos.selectedItem.SubItems(52)
   txt_composicion = lv_articulos.selectedItem.SubItems(53)
   txt_peso = lv_articulos.selectedItem.SubItems(54)
   txt_tara = lv_articulos.selectedItem.SubItems(55)
   txt_caja = lv_articulos.selectedItem.SubItems(56)
   txt_nombre_caja = lv_articulos.selectedItem.SubItems(57)
   txt_piezas_caja = lv_articulos.selectedItem.SubItems(58)
   txt_maximo = lv_articulos.selectedItem.SubItems(59)
   txt_minimo = lv_articulos.selectedItem.SubItems(60)
   txt_punto_reorden = lv_articulos.selectedItem.SubItems(61)
   txt_dias_inventario = lv_articulos.selectedItem.SubItems(62)
   txt_ubicacion = lv_articulos.selectedItem.SubItems(63)
   txt_nombre_ubicacion = ""
   txt_bulto = lv_articulos.selectedItem.SubItems(65)
   chk_salida_masiva.Value = lv_articulos.selectedItem.SubItems(66)
   chk_detenido.Value = lv_articulos.selectedItem.SubItems(67)
   txt_nombre_ubicacion = ""
   rs.Open "SELECT * FROM TB_EQUIVALENCIAS WHERE VCHA_ART_ARTICULO_ID = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      txt_equivalente = IIf(IsNull(rs!vcha_equ_codigo_equivalente), "", rs!vcha_equ_codigo_equivalente)
      txt_equivalente.Enabled = False
   Else
      txt_equivalente = ""
      txt_equivalente.Enabled = True
   End If
   rs.Close
   'rs.Open "select * from tb_lineas where vcha_lin_linea_id = '" + txt_linea + "'", cnn, adOpenDynamic, adLockOptimistic
   'If Not rs.EOF Then
   '   txt_compresion = IIf(IsNull(rs!FLOA_LIN_COMPRECION), 0, rs!FLOA_LIN_COMPRECION)
   'Else
   '   txt_compresion = 0
   'End If
   'rs.Close
   'If txt_compresion = "" Then
   '   txt_compresion = 0
   'End If
   'txt_volumen_compreso = (txt_volumen * (1 - (txt_compresion / 100)))
   'rs.Open "select * from tb_Cajas where vcha_caj_caja_id = '" + txt_caja + "'", cnn, adOpenDynamic, adLockOptimistic
   'If Not rs.EOF Then
   '   var_ancho = IIf(IsNull(rs!floa_caj_ancho), 0, rs!floa_caj_ancho)
   '   var_alto = IIf(IsNull(rs!floa_Caj_alto), 0, rs!floa_Caj_alto)
   '   var_largo = IIf(IsNull(rs!floa_caj_largo), 0, rs!floa_caj_largo)
   'Else
   '   var_ancho = 0
   '   var_alto = 0
   '   var_largo = 0
   'End If
   'rs.Close
   'If (var_ancho * var_alto * var_largo) = 0 Then
   '   txt_piezas_caja = 0
   'Else
   '   If CStr(txt_volumen_compreso) <> 0 Then
   '      txt_piezas_caja = (var_ancho * var_alto * var_largo) / txt_volumen_compreso
   '   Else
   '      txt_piezas_caja = 0
   '   End If
   'End If
   var_guarda_cambios = False
   var_hubo_cambios = False
   var_modifica_registro_articulo = True
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro_articulo = False Then
        Set list_item = lv_articulos.ListItems.Add(, , txt_codigo)
        list_item.SubItems(1) = txt_descripcion
        list_item.SubItems(2) = msk_precio
        list_item.SubItems(3) = msk_costo
        list_item.SubItems(4) = txt_fecha_inicio
        list_item.SubItems(5) = txt_fecha_fin
        list_item.SubItems(6) = txt_catalogo_inicio
        list_item.SubItems(7) = txt_nombre_catalogo_inicio
        list_item.SubItems(8) = txt_catalogo_fin
        list_item.SubItems(9) = txt_nombre_catalogo_fin
        list_item.SubItems(10) = txt_dueño_licencia
        list_item.SubItems(11) = txt_nombre_dueño_licencia
        list_item.SubItems(12) = txt_contrato
        list_item.SubItems(13) = txt_familia
        list_item.SubItems(14) = txt_nombre_familia
        list_item.SubItems(15) = txt_linea
        list_item.SubItems(16) = txt_nombre_linea
        list_item.SubItems(17) = txt_sublinea
        list_item.SubItems(18) = txt_nombre_sublinea
        list_item.SubItems(19) = txt_producto
        list_item.SubItems(20) = txt_nombre_producto
        list_item.SubItems(21) = txt_tipo_producto
        list_item.SubItems(22) = txt_nombre_tipo_producto
        list_item.SubItems(23) = txt_clase
        list_item.SubItems(24) = txt_nombre_clase
        list_item.SubItems(25) = txt_estampado_anverso
        list_item.SubItems(26) = txt_nombre_Estampado_anverso
        list_item.SubItems(27) = txt_tipo_estampado_anverso
        list_item.SubItems(28) = txt_nombre_tipo_estampado_anverso
        list_item.SubItems(29) = txt_estampado_reverso
        list_item.SubItems(30) = txt_nombre_estampado_reverso
        list_item.SubItems(31) = txt_tipo_estampado_reverso
        list_item.SubItems(32) = txt_nombre_tipo_estampado_reverso
        list_item.SubItems(33) = txt_color_anverso
        list_item.SubItems(34) = txt_nombre_color_anverso
        list_item.SubItems(35) = txt_color_reverso
        list_item.SubItems(36) = txt_nombre_color_reverso
        list_item.SubItems(37) = txt_tono_anverso
        list_item.SubItems(38) = txt_nombre_tono_anverso
        list_item.SubItems(39) = txt_tono_renverso
        list_item.SubItems(40) = txt_nombre_tono_reverso
        list_item.SubItems(41) = txt_decorativos
        list_item.SubItems(42) = txt_fundas
        list_item.SubItems(43) = txt_uso_producto
        list_item.SubItems(44) = txt_nombre_uso_producto
        list_item.SubItems(45) = txt_subtipo_uso_producto
        list_item.SubItems(46) = txt_nombre_subtipo_uso_producto
        list_item.SubItems(47) = txt_talla
        list_item.SubItems(48) = txt_nombre_talla
        list_item.SubItems(49) = txt_unidad
        list_item.SubItems(50) = txt_nombre_unidad
        list_item.SubItems(51) = txt_volumen
        list_item.SubItems(52) = txt_tela
        list_item.SubItems(53) = txt_composicion
        list_item.SubItems(54) = txt_peso
        list_item.SubItems(55) = txt_tara
        list_item.SubItems(56) = txt_caja
        list_item.SubItems(57) = txt_nombre_caja
        list_item.SubItems(58) = txt_piezas_caja
        list_item.SubItems(59) = txt_maximo
        list_item.SubItems(60) = txt_minimo
        list_item.SubItems(61) = txt_punto_reorden
        list_item.SubItems(62) = txt_dias_inventario
        list_item.SubItems(63) = txt_ubicacion
        list_item.SubItems(64) = ""
        list_item.SubItems(65) = txt_bulto
        list_item.SubItems(66) = chk_salida_masiva.Value
        list_item.SubItems(67) = Me.chk_detenido
        list_item.EnsureVisible
        list_item.Selected = True
        numero_items_articulos = numero_items_articulos + 1
    Else
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).Checked = False
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index) = txt_codigo
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(1) = txt_descripcion
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(2) = msk_precio
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(3) = msk_costo
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(4) = txt_fecha_inicio
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(5) = txt_fecha_fin
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(6) = txt_catalogo_inicio
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(7) = txt_nombre_catalogo_inicio
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(8) = txt_catalogo_fin
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(9) = txt_nombre_catalogo_fin
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(10) = txt_dueño_licencia
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(11) = txt_nombre_dueño_licencia
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(12) = txt_contrato
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(13) = txt_familia
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(14) = txt_nombre_familia
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(15) = txt_linea
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(16) = txt_nombre_linea
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(17) = txt_sublinea
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(18) = txt_nombre_sublinea
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(19) = txt_producto
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(20) = txt_nombre_producto
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(21) = txt_tipo_producto
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(22) = txt_nombre_tipo_producto
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(23) = txt_clase
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(24) = txt_nombre_clase
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(25) = txt_estampado_anverso
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(26) = txt_nombre_Estampado_anverso
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(27) = txt_tipo_estampado_anverso
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(28) = txt_nombre_tipo_estampado_anverso
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(29) = txt_estampado_reverso
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(30) = txt_nombre_estampado_reverso
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(31) = txt_tipo_estampado_reverso
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(32) = txt_nombre_tipo_estampado_reverso
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(33) = txt_color_anverso
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(34) = txt_nombre_color_anverso
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(35) = txt_color_reverso
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(36) = txt_nombre_color_reverso
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(37) = txt_tono_anverso
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(38) = txt_nombre_tono_anverso
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(39) = txt_tono_renverso
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(40) = txt_nombre_tono_reverso
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(41) = txt_decorativos
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(42) = txt_fundas
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(43) = txt_uso_producto
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(44) = txt_nombre_uso_producto
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(45) = txt_subtipo_uso_producto
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(46) = txt_nombre_subtipo_uso_producto
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(47) = txt_talla
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(48) = txt_nombre_talla
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(49) = txt_unidad
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(50) = txt_nombre_unidad
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(51) = txt_volumen
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(52) = txt_tela
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(53) = txt_composicion
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(54) = txt_peso
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(55) = txt_tara
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(56) = txt_caja
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(57) = txt_nombre_caja
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(58) = txt_piezas_caja
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(59) = txt_maximo
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(60) = txt_minimo
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(61) = txt_punto_reorden
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(62) = txt_dias_inventario
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(63) = txt_ubicacion
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(64) = ""
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(65) = txt_bulto
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(66) = chk_salida_masiva.Value
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).ListSubItems(67) = Me.chk_detenido.Value
        lv_articulos.ListItems.Item(lv_articulos.selectedItem.Index).Selected = True
    End If
    'lv_articulos.SetFocus
End Sub

Private Sub txt_nombre_uso_producto_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_uso_producto_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_peso_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_peso_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_piezas_caja_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_piezas_caja_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_producto_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_producto_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_producto_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_productos order by vcha_pro_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_pro_producto_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_PRO_NOMBRE), "", rs!VCHA_PRO_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "PRODUCTOS"
      var_tipo_lista = 7
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
      var_activa_forma_productos = Me.Name
      frmarticulos2.Enabled = False
      frmproductos.Show
   End If
End Sub

Private Sub txt_producto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_producto_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_producto) <> "" Then
      rs.Open "select * from tb_productos where vcha_pro_producto_id = '" + txt_producto + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_producto = IIf(IsNull(rs!VCHA_PRO_NOMBRE), "", rs!VCHA_PRO_NOMBRE)
      Else
         txt_nombre_producto = ""
         Me.txt_producto = ""
         MsgBox "Clave de producto incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      txt_nombre_producto = ""
   End If
End Sub

Private Sub txt_punto_reorden_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_punto_reorden_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_subdivision_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_SUBDIVISIONES where vcha_tpr_tipo_producto_id = '" + Me.txt_tipo + "' AND VCHA_DIV_DIVISION_ID = '" + Me.txt_division + "' order by vcha_SUB_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_SUB_SUBDIVISION_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_sub_nombre), "", rs!vcha_sub_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "SUBDIVISIONES"
      var_tipo_lista = 102
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
End Sub

Private Sub txt_subdivision_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_subdivision_LostFocus()
   If Trim(txt_tipo) <> "" Then
      If Trim(Me.txt_division) <> "" Then
         If Trim(txt_subdivision) <> "" Then
            rs.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_aRT_ARTICULO_ID = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            var_subdivision = ""
            If Not rs.EOF Then
               var_subdivision = IIf(IsNull(rs!VCHA_SUB_SUBDIVISION_ID), "", rs!VCHA_SUB_SUBDIVISION_ID)
            End If
            rs.Close
            rs.Open "SELECT * FROM TB_SUBDIVISIONES WHERE VCHA_TPR_TIPO_PRODUCTO_ID = '" + txt_tipo + "' AND VCHA_DIV_DIVISION_ID = '" + Me.txt_division + "' AND VCHA_SUB_SUBDIVISION_ID = '" + Me.txt_subdivision + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               Me.txt_nombre_subdivision = IIf(IsNull(rs!vcha_sub_nombre), "", rs!vcha_sub_nombre)
               If var_subdivision <> Me.txt_subdivision Then
                  txt_estampado = ""
                  Me.txt_nombre_estampado = ""
               End If
            Else
               Me.txt_subdivision = ""
               Me.txt_nombre_subdivision = ""
               Me.txt_estampado = ""
               Me.txt_nombre_estampado = ""
            End If
            rs.Close
         Else
         End If
      Else
         MsgBox "No se a seleccionado una división", vbOKOnly, "ATENCION"
         Me.txt_subdivision = ""
         Me.txt_nombre_subdivision = ""
         Me.txt_estampado = ""
         Me.txt_nombre_estampado = ""
      End If
   Else
      MsgBox "No se a seleccionado un tipo de producto", vbOKOnly, "ATENCION"
      Me.txt_subdivision = ""
      Me.txt_nombre_subdivision = ""
      Me.txt_estampado = ""
      Me.txt_nombre_estampado = ""
   End If
End Sub

Private Sub txt_sublinea_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_sublinea_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_sublinea_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 37 Then
      Me.txt_sublinea.SetFocus
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_sublineas order by vcha_sli_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_SLI_SUBLINEA_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_SLI_NOMBRE), "", rs!VCHA_SLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "SUBLINEAS"
      var_tipo_lista = 6
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
      var_activa_forma_sublineas = Me.Name
      frmarticulos2.Enabled = False
      frmsublineas.Show
   End If
End Sub

Private Sub txt_sublinea_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_sublinea_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_sublinea) <> "" Then
      rs.Open "SELECT * FROM TB_SUBLINEAS WHERE VCHA_SLI_SUBLINEA_ID = '" + txt_sublinea + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_sublinea = IIf(IsNull(rs!VCHA_SLI_NOMBRE), "", rs!VCHA_SLI_NOMBRE)
      Else
         MsgBox "Clave de sublinea incorrecta", vbOKOnly, "ATENCION"
         Me.txt_sublinea = ""
         txt_nombre_sublinea = ""
      End If
      rs.Close
   Else
      txt_nombre_sublinea = ""
   End If
End Sub

Private Sub txt_subtipo_uso_producto_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_subtipo_uso_producto_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_subtipo_uso_producto_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_SUBTIPOSUSOS order by vcha_sus_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_SUS_SUBTIPO_USO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_SUS_NOMBRE), "", rs!VCHA_SUS_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "SUBTIPO DE USOS"
      var_tipo_lista = 19
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
      var_activa_forma_subtiposusos = Me.Name
      frmarticulos2.Enabled = False
      frmsubtiposusos.Show
   End If
End Sub

Private Sub txt_subtipo_uso_producto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_subtipo_uso_producto_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_subtipo_uso_producto) <> "" Then
      rs.Open "SELECT * FROM TB_SUBTIPOSUSOS WHERE VCHA_SUS_SUBTIPO_USO_ID = '" + txt_subtipo_uso_producto + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_subtipo_uso_producto = IIf(IsNull(rs!VCHA_SUS_NOMBRE), "", rs!VCHA_SUS_NOMBRE)
      Else
         MsgBox "Clave de subtipo de uso incorrecta", vbOKOnly, "ATENCION"
         Me.txt_subtipo_uso_producto = ""
         txt_nombre_subtipo_uso_producto = ""
      End If
      rs.Close
   Else
      txt_nombre_subtipo_uso_producto = ""
   End If
End Sub

Private Sub txt_talla_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_talla_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_talla_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_TALLAS order by vcha_tal_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_TAL_TALLA_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_tal_NOMBRE), "", rs!VCHA_tal_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TALLAS"
      var_tipo_lista = 20
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
      var_activa_forma_tallas = Me.Name
      frmarticulos2.Enabled = False
      frmtallas.Show
   End If
End Sub

Private Sub txt_talla_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_talla_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_talla) <> "" Then
      rs.Open "SELECT * FROM TB_TALLAS WHERE VCHA_TAL_TALLA_ID = '" + txt_talla + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_talla = IIf(IsNull(rs!VCHA_tal_NOMBRE), "", rs!VCHA_tal_NOMBRE)
      Else
         MsgBox "Clave de talla incorrecta", vbOKOnly, "ATENCION"
         Me.txt_talla = ""
         txt_nombre_talla = ""
      End If
      rs.Close
   Else
      txt_nombre_talla = ""
   End If
End Sub

Private Sub txt_tara_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_tara_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_tela_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_tela_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_tipo_estampado_anverso_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_tipo_estampado_anverso_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_tipo_estampado_anverso_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_TIPOESTAMPADOS order by vcha_tes_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_TES_TIPOESTAMPADO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_TES_NOMBRE), "", rs!VCHA_TES_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TIPO ESTAMPADOS"
      var_tipo_lista = 11
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
      var_activa_forma_tipoestampados = Me.Name
      frmarticulos2.Enabled = False
      frmtipoestampados.Show
   End If
End Sub

Private Sub txt_tipo_estampado_anverso_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_tipo_estampado_anverso_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_tipo_estampado_anverso) <> "" Then
      rs.Open "select * from TB_TIPOESTAMPADOS where VCHA_TES_TIPOESTAMPADO_ID = '" + txt_tipo_estampado_anverso + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_tipo_estampado_anverso = IIf(IsNull(rs!VCHA_TES_NOMBRE), "", rs!VCHA_TES_NOMBRE)
      Else
         MsgBox "Clave de tipo de estampado incorrecta", vbOKOnly, "ATENCION"
         Me.txt_tipo_estampado_anverso = ""
         txt_nombre_tipo_estampado_anverso = ""
      End If
      rs.Close
   Else
      txt_nombre_tipo_estampado_anverso = ""
   End If
End Sub

Private Sub txt_tipo_estampado_reverso_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_tipo_estampado_reverso_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_tipo_estampado_reverso_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_TIPOESTAMPADOS order by vcha_tes_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_TES_TIPOESTAMPADO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_TES_NOMBRE), "", rs!VCHA_TES_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TIPO ESTAMPADOS"
      var_tipo_lista = 13
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
      var_activa_forma_tipoestampados = Me.Name
      frmarticulos2.Enabled = False
      frmtipoestampados.Show
   End If
End Sub

Private Sub txt_tipo_estampado_reverso_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_tipo_estampado_reverso_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_tipo_estampado_reverso) <> "" Then
      rs.Open "select * from TB_TIPOESTAMPADOS where VCHA_TES_TIPOESTAMPADO_ID = '" + txt_tipo_estampado_reverso + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_tipo_estampado_reverso = IIf(IsNull(rs!VCHA_TES_NOMBRE), "", rs!VCHA_TES_NOMBRE)
      Else
         MsgBox "Clave de tipo de estampado incorrecta", vbOKOnly, "ATENCION"
         txt_nombre_tipo_estampado_reverso = ""
         Me.txt_tipo_estampado_reverso = ""
      End If
      rs.Close
   Else
      txt_nombre_tipo_estampado_reverso = ""
   End If
End Sub

Private Sub txt_tipo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_TIPOs_productos order by vcha_tpr_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_tpr_tipo_producto_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_tpr_nombre), "", rs!vcha_tpr_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TIPO PRODUCTO"
      var_tipo_lista = 100
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
End Sub

Private Sub txt_tipo_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_tipo_LostFocus()
   If Trim(txt_tipo) <> "" Then
      rs.Open "select * from tb_Articulos where vcha_art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
      var_tipo = ""
      If Not rs.EOF Then
         var_tipo = IIf(IsNull(rs!VCHA_tpr_tipo_producto_ID), "", rs!VCHA_tpr_tipo_producto_ID)
      End If
      rs.Close
      rs.Open "select * from tb_tipos_productos where vcha_tpr_tipo_producto_id = '" + Me.txt_tipo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_tipo = IIf(IsNull(rs!vcha_tpr_nombre), "", rs!vcha_tpr_nombre)
         If var_tipo <> Me.txt_tipo Then
            Me.txt_division = ""
            Me.txt_nombre_division = ""
            Me.txt_subdivision = ""
            Me.txt_nombre_subdivision = ""
            Me.txt_estampado = ""
            Me.txt_nombre_estampado = ""
         End If
      Else
         Me.txt_tipo = ""
         Me.txt_nombre_tipo = ""
         Me.txt_division = ""
         Me.txt_nombre_division = ""
         Me.txt_subdivision = ""
         Me.txt_nombre_subdivision = ""
         Me.txt_estampado = ""
         Me.txt_nombre_estampado = ""
      End If
      rs.Close
   End If
End Sub

Private Sub txt_tipo_producto_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_tipo_producto_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_tipo_producto_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_TIPOARTICULOS order by vcha_tar_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_TAR_TIPO_ARTICULO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_TAR_NOMBRE), "", rs!VCHA_TAR_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TIPO PRODUCTO"
      var_tipo_lista = 8
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
      var_activa_forma_tipoarticulos = Me.Name
      frmarticulos2.Enabled = False
      frmtipoarticulos.Show
   End If
End Sub

Private Sub txt_tipo_producto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_tipo_producto_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_tipo_producto) <> "" Then
      rs.Open "SELECT * FROM TB_TIPOARTICULOS WHERE VCHA_TAR_TIPO_ARTICULO_ID = '" + txt_tipo_producto + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_tipo_producto = IIf(IsNull(rs!VCHA_TAR_NOMBRE), "", rs!VCHA_TAR_NOMBRE)
      Else
         MsgBox "Clave de tipo de producto incorrecta", vbOKOnly, "ATENCION"
         Me.txt_tipo_producto = ""
         txt_nombre_tipo_producto = ""
      End If
      rs.Close
   Else
      txt_nombre_tipo_producto = ""
   End If
End Sub

Private Sub txt_tono_anverso_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_tono_anverso_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_tono_anverso_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_TONOS order by vcha_ton_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_TON_TONO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_TON_NOMBRE), "", rs!VCHA_TON_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TONOS"
      var_tipo_lista = 16
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
      var_activa_forma_tonos = Me.Name
      frmarticulos2.Enabled = False
      frmtonos.Show
   End If
End Sub

Private Sub txt_tono_anverso_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_tono_anverso_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_tono_anverso) <> "" Then
      rs.Open "SELECT * FROM TB_TONOS WHERE VCHA_TON_TONO_ID = '" + txt_tono_anverso + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_tono_anverso = IIf(IsNull(rs!VCHA_TON_NOMBRE), "", rs!VCHA_TON_NOMBRE)
      Else
         MsgBox "Clave de color incorrecta", vbOKOnly, "ATENCION"
         Me.txt_tono_anverso = ""
         txt_nombre_tono_anverso = ""
      End If
      rs.Close
   Else
      txt_nombre_tono_anverso = ""
   End If
End Sub

Private Sub txt_tono_reverso_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_tono_reverso_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_tono_reverso_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_TONOS order by vcha_ton_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_TON_TONO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_TON_NOMBRE), "", rs!VCHA_TON_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TONOS"
      var_tipo_lista = 17
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
      var_activa_forma_tonos = Me.Name
      frmarticulos2.Enabled = False
      frmtonos.Show
   End If
End Sub

Private Sub txt_tono_reverso_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_tono_reverso_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_tono_reverso) <> "" Then
      rs.Open "SELECT * FROM TB_TONOS WHERE VCHA_TON_TONO_ID = '" + txt_tono_anverso + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_tono_anverso = IIf(IsNull(rs!VCHA_TON_NOMBRE), "", rs!VCHA_TON_NOMBRE)
      Else
         MsgBox "Clave de color incorrecta", vbOKOnly, "ATENCION"
         txt_nombre_tono_anverso = ""
         Me.txt_tono_reverso = ""
      End If
      rs.Close
   Else
      txt_nombre_tono_anverso = ""
   End If
End Sub

Private Sub txt_ubicacion_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_ubicacion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_unidad_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_unidad_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_unidad_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_UNIDADES order by vcha_uni_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_uni_unidad_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_UNI_NOMBRE), "", rs!VCHA_UNI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "UNIDADES"
      var_tipo_lista = 21
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
      var_activa_forma_unidades = Me.Name
      frmarticulos2.Enabled = False
      frmunidades.Show
   End If
End Sub

Private Sub txt_unidad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_unidad_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_unidad) <> "" Then
      rs.Open "SELECT * FROM TB_UNIDADES WHERE VCHA_UNI_UNIDAD_ID = '" + txt_unidad + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_unidad = IIf(IsNull(rs!VCHA_UNI_NOMBRE), "", rs!VCHA_UNI_NOMBRE)
      Else
         MsgBox "Clave de unidad incorrecta", vbOKOnly, "ATENCION"
         Me.txt_unidad = ""
         txt_nombre_unidad = ""
      End If
      rs.Close
   Else
      txt_nombre_unidad = ""
   End If
End Sub

Private Sub txt_uso_producto_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_uso_producto_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_uso_producto_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_usos order by vcha_uso_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_USO_USO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_USO_NOMBRE), "", rs!VCHA_USO_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "USOS"
      var_tipo_lista = 18
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
      var_activa_forma_usos = Me.Name
      frmarticulos2.Enabled = False
      frmusos.Show
   End If
End Sub

Private Sub txt_uso_producto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_uso_producto_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_uso_producto) <> "" Then
      rs.Open "SELECT * FROM TB_USOS WHERE VCHA_USO_USO_ID = '" + txt_uso_producto + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_uso_producto = IIf(IsNull(rs!VCHA_USO_NOMBRE), "", rs!VCHA_USO_NOMBRE)
      Else
         MsgBox "Clave de uso incorrecta", vbOKOnly, "ATENCION"
         txt_nombre_uso_producto = ""
         Me.txt_uso_producto = ""
      End If
      rs.Close
   Else
      txt_nombre_uso_producto = ""
   End If
End Sub

Private Sub txt_volumen_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_volumen_compreso_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_volumen_compreso_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_volumen_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If Me.cmd_guardar.Enabled = True Then
         Me.cmd_guardar.SetFocus
      Else
         Call pro_enfoque(KeyAscii)
      End If
   End If
End Sub
