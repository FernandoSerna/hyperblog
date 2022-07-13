VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form froracle_asignacion_embarques 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignación de embarques"
   ClientHeight    =   11655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   582.75
   ScaleMode       =   0  'User
   ScaleWidth      =   9217.304
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_embarque 
      Height          =   405
      Left            =   45
      TabIndex        =   87
      Top             =   4710
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Frame Frame8 
      Height          =   510
      Left            =   150
      TabIndex        =   85
      Top             =   11085
      Width           =   15030
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "F5  Asignar máquinas a embarque.              F8 Gráfica de progreso de lotes.         F9 Agregar ruta a embarques"
         Height          =   195
         Left            =   165
         TabIndex        =   86
         Top             =   195
         Width           =   7800
      End
   End
   Begin VB.Frame Frame4 
      Height          =   3375
      Left            =   150
      TabIndex        =   65
      Top             =   7680
      Width           =   15030
      Begin VB.TextBox txt_transporte 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1365
         TabIndex        =   82
         Top             =   2805
         Width           =   3075
      End
      Begin VB.TextBox txt_total_volumen 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   13095
         TabIndex        =   75
         Top             =   2805
         Width           =   1065
      End
      Begin VB.TextBox txt_porcentaje 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9930
         TabIndex        =   74
         Top             =   2805
         Width           =   1020
      End
      Begin VB.TextBox txt_volumen_unidad 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6315
         TabIndex        =   73
         Top             =   2805
         Width           =   1050
      End
      Begin VB.Frame frm_orden 
         BorderStyle     =   0  'None
         Height          =   825
         Left            =   9660
         TabIndex        =   69
         Top             =   1395
         Width           =   2145
         Begin VB.Label lbl_orden 
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   795
            Left            =   30
            TabIndex        =   70
            Top             =   15
            Width           =   2100
         End
      End
      Begin MSComctlLib.ListView lv_pedidos 
         CausesValidation=   0   'False
         Height          =   2175
         Left            =   75
         TabIndex        =   67
         Top             =   450
         Width           =   14910
         _ExtentX        =   26300
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Pedido"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Agente"
            Object.Width           =   8114
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cliente"
            Object.Width           =   8820
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Piezas"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Orden de carga"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Volumen"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Clave cliente"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Clave establecimiento"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Paqueteria"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Transporte:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   81
         Top             =   2865
         Width           =   1215
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Total volumen:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   11490
         TabIndex        =   79
         Top             =   2865
         Width           =   1545
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Porcentaje ocupación:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7560
         TabIndex        =   78
         Top             =   2865
         Width           =   2340
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Volumen unidad:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4545
         TabIndex        =   77
         Top             =   2865
         Width           =   1740
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   11055
         TabIndex        =   76
         Top             =   2865
         Width           =   210
      End
      Begin VB.Label Label33 
         Caption         =   "Label33"
         Height          =   615
         Left            =   8055
         TabIndex        =   68
         Top             =   1365
         Width           =   3030
      End
      Begin VB.Label Label23 
         BackColor       =   &H000000C0&
         Caption         =   " Detalles del embarque"
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
         Height          =   270
         Left            =   30
         TabIndex        =   66
         Top             =   135
         Width           =   14955
      End
   End
   Begin VB.Frame Frame5 
      Height          =   2040
      Left            =   120
      TabIndex        =   20
      Top             =   4200
      Width           =   15030
      Begin VB.Frame Frame13 
         Caption         =   "Frame2"
         Height          =   2025
         Left            =   6030
         TabIndex        =   29
         Top             =   15
         Width           =   30
      End
      Begin VB.Frame Frame12 
         Caption         =   "Frame2"
         Height          =   2025
         Left            =   9020
         TabIndex        =   28
         Top             =   15
         Width           =   30
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   2025
         Index           =   0
         Left            =   3030
         TabIndex        =   27
         Top             =   0
         Width           =   30
      End
      Begin VB.CommandButton cmd_anden_6 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   690
         TabIndex        =   26
         Top             =   1530
         Width           =   1725
      End
      Begin VB.CommandButton cmd_anden_7 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3735
         TabIndex        =   25
         Top             =   1530
         Width           =   1725
      End
      Begin VB.CommandButton cmd_anden_8 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6750
         TabIndex        =   24
         Top             =   1530
         Width           =   1725
      End
      Begin VB.CommandButton cmd_anden_9 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9720
         TabIndex        =   23
         Top             =   1530
         Width           =   1725
      End
      Begin VB.CommandButton cmd_anden_10 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   12705
         TabIndex        =   22
         Top             =   1530
         Width           =   1725
      End
      Begin VB.Frame Frame11 
         Caption         =   "Frame2"
         Height          =   2025
         Left            =   12015
         TabIndex        =   21
         Top             =   15
         Width           =   30
      End
      Begin VB.Label lbl_cantidad_10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   14400
         TabIndex        =   54
         Top             =   885
         Width           =   540
      End
      Begin VB.Label lbl_cantidad_9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   11430
         TabIndex        =   53
         Top             =   885
         Width           =   540
      End
      Begin VB.Label lbl_cantidad_8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8415
         TabIndex        =   52
         Top             =   885
         Width           =   540
      End
      Begin VB.Label lbl_cantidad_7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5415
         TabIndex        =   51
         Top             =   885
         Width           =   540
      End
      Begin VB.Label lbl_cantidad_6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2430
         TabIndex        =   50
         Top             =   885
         Width           =   540
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Piezas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   12075
         TabIndex        =   44
         Top             =   915
         Width           =   780
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Piezas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9090
         TabIndex        =   43
         Top             =   915
         Width           =   780
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Piezas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6105
         TabIndex        =   42
         Top             =   915
         Width           =   780
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Piezas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3135
         TabIndex        =   41
         Top             =   915
         Width           =   780
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Piezas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   40
         Top             =   915
         Width           =   780
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Estación 6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   435
         Left            =   630
         TabIndex        =   34
         Top             =   180
         Width           =   1830
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Estación 7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   435
         Left            =   3675
         TabIndex        =   33
         Top             =   180
         Width           =   1830
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Estación 8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   435
         Left            =   6690
         TabIndex        =   32
         Top             =   180
         Width           =   1830
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Estación 9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   435
         Left            =   9660
         TabIndex        =   31
         Top             =   180
         Width           =   1830
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Estación 10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   435
         Left            =   12540
         TabIndex        =   30
         Top             =   180
         Width           =   2040
      End
   End
   Begin VB.Frame Frame10 
      Height          =   735
      Left            =   9180
      TabIndex        =   11
      Top             =   0
      Width           =   5970
      Begin VB.CommandButton cmd_hoja_carga 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5385
         Picture         =   "froracle_asignacion_embarques.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Generar hoja de carga para embarques en ruta"
         Top             =   270
         Width           =   330
      End
      Begin VB.TextBox txt_fecha 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1005
         TabIndex        =   13
         Top             =   172
         Width           =   2445
      End
      Begin VB.Label Label30 
         Caption         =   "Hoja de carga para embarques"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3540
         TabIndex        =   89
         Top             =   225
         Width           =   2010
      End
      Begin VB.Label Label13 
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   45
         TabIndex        =   12
         Top             =   225
         Width           =   885
      End
   End
   Begin VB.Frame Frame9 
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   9045
      Begin VB.CommandButton cmd_lotes_en_picking 
         Appearance      =   0  'Flat
         Caption         =   "Lotes en picking"
         Height          =   315
         Left            =   2880
         Picture         =   "froracle_asignacion_embarques.frx":0102
         TabIndex        =   96
         ToolTipText     =   "Imprimir ordenes de surtido ordenadas por prioridad"
         Top             =   240
         Width           =   1530
      End
      Begin VB.CommandButton cmd_fraccion_pedidos_prioridad 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5520
         Picture         =   "froracle_asignacion_embarques.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   95
         ToolTipText     =   "Imprimir ordenes de surtido ordenadas por prioridad"
         Top             =   240
         Width           =   330
      End
      Begin VB.CommandButton cmd_orden_ubicacion_310317 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5160
         Picture         =   "froracle_asignacion_embarques.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   94
         ToolTipText     =   "Imprimir ordenes de surtido con multiplos y agrupados por pasillo"
         Top             =   240
         Width           =   330
      End
      Begin VB.CommandButton cmd_imprimir_por_pasillo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4785
         Picture         =   "froracle_asignacion_embarques.frx":0408
         Style           =   1  'Graphical
         TabIndex        =   93
         ToolTipText     =   "Imprimir ordenes de surtido con multiplos y agrupados por pasillo"
         Top             =   240
         Width           =   330
      End
      Begin VB.ComboBox cmb_dia 
         Height          =   315
         ItemData        =   "froracle_asignacion_embarques.frx":050A
         Left            =   6960
         List            =   "froracle_asignacion_embarques.frx":0520
         TabIndex        =   92
         Top             =   240
         Width           =   1950
      End
      Begin VB.CommandButton cmd_actualiza_informacion_pedidos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6195
         Picture         =   "froracle_asignacion_embarques.frx":0557
         Style           =   1  'Graphical
         TabIndex        =   90
         ToolTipText     =   "Actualizar información "
         Top             =   255
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmd_imprimir_nuevo_metodo_divisiones 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4425
         Picture         =   "froracle_asignacion_embarques.frx":0659
         Style           =   1  'Graphical
         TabIndex        =   84
         ToolTipText     =   "Imprimir ordenes de surtido con multiplos y agrupados"
         Top             =   240
         Width           =   330
      End
      Begin VB.CommandButton cmd_imprimir_nuevo_metodo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1530
         Picture         =   "froracle_asignacion_embarques.frx":075B
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   "Imprimir ordenes de surtido con multiplos"
         Top             =   255
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   300
         Left            =   600
         TabIndex        =   80
         Top             =   255
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.CommandButton cmd_imprimir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1065
         Picture         =   "froracle_asignacion_embarques.frx":085D
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Imprimir ordenes de surtido"
         Top             =   255
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Command6"
         Height          =   330
         Left            =   90
         TabIndex        =   71
         Top             =   240
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Dia de carga:"
         Height          =   195
         Left            =   5940
         TabIndex        =   91
         Top             =   285
         Width           =   960
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Creación y asignacion de embarques"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   105
         TabIndex        =   10
         Top             =   210
         Visible         =   0   'False
         Width           =   4710
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2040
      Left            =   120
      TabIndex        =   0
      Top             =   675
      Width           =   15030
      Begin VB.Frame Frame3 
         Caption         =   "Frame2"
         Height          =   2025
         Left            =   12015
         TabIndex        =   19
         Top             =   15
         Width           =   30
      End
      Begin VB.CommandButton cmd_anden_5 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   12690
         TabIndex        =   18
         Top             =   1515
         Width           =   1725
      End
      Begin VB.CommandButton cmd_anden_4 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9720
         TabIndex        =   17
         Top             =   1515
         Width           =   1725
      End
      Begin VB.CommandButton cmd_anden_3 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6750
         TabIndex        =   16
         Top             =   1515
         Width           =   1725
      End
      Begin VB.CommandButton cmd_anden_2 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3750
         TabIndex        =   15
         Top             =   1500
         Width           =   1725
      End
      Begin VB.CommandButton cmd_anden_1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   765
         TabIndex        =   14
         Top             =   1515
         Width           =   1725
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   2025
         Index           =   1
         Left            =   3030
         TabIndex        =   3
         Top             =   0
         Width           =   30
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame2"
         Height          =   2025
         Left            =   9020
         TabIndex        =   2
         Top             =   15
         Width           =   30
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frame2"
         Height          =   2025
         Left            =   6030
         TabIndex        =   1
         Top             =   15
         Width           =   30
      End
      Begin VB.Label lbl_cantidad_5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   14385
         TabIndex        =   49
         Top             =   930
         Width           =   540
      End
      Begin VB.Label lbl_cantidad_4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   11415
         TabIndex        =   48
         Top             =   930
         Width           =   540
      End
      Begin VB.Label lbl_cantidad_3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8400
         TabIndex        =   47
         Top             =   930
         Width           =   540
      End
      Begin VB.Label lbl_cantidad_2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5400
         TabIndex        =   46
         Top             =   930
         Width           =   540
      End
      Begin VB.Label lbl_cantidad_1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2415
         TabIndex        =   45
         Top             =   930
         Width           =   540
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Piezas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   12120
         TabIndex        =   39
         Top             =   960
         Width           =   780
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Piezas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9135
         TabIndex        =   38
         Top             =   960
         Width           =   780
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Piezas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6150
         TabIndex        =   37
         Top             =   960
         Width           =   780
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Piezas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3180
         TabIndex        =   36
         Top             =   960
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Piezas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   195
         TabIndex        =   35
         Top             =   960
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Estación 5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   435
         Left            =   12630
         TabIndex        =   8
         Top             =   195
         Width           =   1830
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Estación 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   435
         Left            =   9660
         TabIndex        =   7
         Top             =   195
         Width           =   1830
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Estación 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   435
         Left            =   6690
         TabIndex        =   6
         Top             =   195
         Width           =   1830
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Estación 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   435
         Left            =   3690
         TabIndex        =   5
         Top             =   195
         Width           =   1830
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Estación 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   435
         Left            =   705
         TabIndex        =   4
         Top             =   195
         Width           =   1830
      End
   End
   Begin MSComctlLib.ListView lv_embarques_1 
      Height          =   1440
      Left            =   120
      TabIndex        =   55
      Top             =   2745
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   2540
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
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Embarques"
         Object.Width           =   5114
      EndProperty
   End
   Begin MSComctlLib.ListView lv_embarques_2 
      Height          =   1440
      Left            =   3150
      TabIndex        =   56
      Top             =   2745
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   2540
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
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Embarques"
         Object.Width           =   5114
      EndProperty
   End
   Begin MSComctlLib.ListView lv_embarques_3 
      Height          =   1440
      Left            =   6120
      TabIndex        =   57
      Top             =   2760
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   2540
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
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Embarques"
         Object.Width           =   5114
      EndProperty
   End
   Begin MSComctlLib.ListView lv_embarques_4 
      Height          =   1440
      Left            =   9135
      TabIndex        =   58
      Top             =   2745
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   2540
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
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Embarques"
         Object.Width           =   5114
      EndProperty
   End
   Begin MSComctlLib.ListView lv_embarques_5 
      Height          =   1440
      Left            =   12165
      TabIndex        =   59
      Top             =   2745
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   2540
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
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Embarque"
         Object.Width           =   5114
      EndProperty
   End
   Begin MSComctlLib.ListView lv_embarques_6 
      Height          =   1440
      Left            =   135
      TabIndex        =   60
      Top             =   6240
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   2540
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
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Embarques"
         Object.Width           =   5114
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Agente"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lv_embarques_7 
      Height          =   1440
      Left            =   3165
      TabIndex        =   61
      Top             =   6240
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   2540
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
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Embarques"
         Object.Width           =   5114
      EndProperty
   End
   Begin MSComctlLib.ListView lv_embarques_8 
      Height          =   1440
      Left            =   6165
      TabIndex        =   62
      Top             =   6240
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   2540
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
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Embarques"
         Object.Width           =   5114
      EndProperty
   End
   Begin MSComctlLib.ListView lv_embarques_9 
      Height          =   1440
      Left            =   9150
      TabIndex        =   63
      Top             =   6240
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   2540
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
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Embarques"
         Object.Width           =   5114
      EndProperty
   End
   Begin MSComctlLib.ListView lv_embarques_10 
      Height          =   1440
      Left            =   12180
      TabIndex        =   64
      Top             =   6240
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   2540
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
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Embarques"
         Object.Width           =   5114
      EndProperty
   End
End
Attribute VB_Name = "froracle_asignacion_embarques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_orden As String
Dim var_posicion As Integer
Dim var_consecutivo_general As Double
Dim var_imprime_pedidos As Integer
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter

Private Sub ilumina_grid()
   var_n = lv_pedidos.ListItems.Count
   For var_i = 1 To var_n
       Me.lv_pedidos.ListItems.Item(var_i).Selected = True
       If (lv_pedidos.ListItems.Item(var_i).ListSubItems(9) * 1) = 1 Then
          lv_pedidos.ListItems.Item(var_i).Bold = True
          lv_pedidos.ListItems.Item(var_i).ListSubItems(1).Bold = True
          lv_pedidos.ListItems.Item(var_i).ListSubItems(2).Bold = True
          lv_pedidos.ListItems.Item(var_i).ListSubItems(3).Bold = True
          lv_pedidos.ListItems.Item(var_i).ListSubItems(4).Bold = True
          lv_pedidos.ListItems.Item(var_i).ListSubItems(5).Bold = True
          lv_pedidos.ListItems.Item(var_i).ListSubItems(6).Bold = True
          lv_pedidos.ListItems.Item(var_i).ListSubItems(7).Bold = True
          lv_pedidos.ListItems.Item(var_i).ListSubItems(8).Bold = True
          lv_pedidos.ListItems.Item(var_i).ForeColor = &H8000000D
          lv_pedidos.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H8000000D
          lv_pedidos.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H8000000D
          lv_pedidos.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H8000000D
          lv_pedidos.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H8000000D
          lv_pedidos.ListItems.Item(var_i).ListSubItems(5).ForeColor = &H8000000D
          lv_pedidos.ListItems.Item(var_i).ListSubItems(6).ForeColor = &H8000000D
          lv_pedidos.ListItems.Item(var_i).ListSubItems(7).ForeColor = &H8000000D
          lv_pedidos.ListItems.Item(var_i).ListSubItems(8).ForeColor = &H8000000D
       Else
          lv_pedidos.ListItems.Item(var_i).Bold = False
          lv_pedidos.ListItems.Item(var_i).ListSubItems(1).Bold = False
          lv_pedidos.ListItems.Item(var_i).ListSubItems(2).Bold = False
          lv_pedidos.ListItems.Item(var_i).ListSubItems(3).Bold = False
          lv_pedidos.ListItems.Item(var_i).ListSubItems(4).Bold = False
          lv_pedidos.ListItems.Item(var_i).ListSubItems(5).Bold = False
          lv_pedidos.ListItems.Item(var_i).ListSubItems(6).Bold = False
          lv_pedidos.ListItems.Item(var_i).ListSubItems(7).Bold = False
          lv_pedidos.ListItems.Item(var_i).ListSubItems(8).Bold = False
          lv_pedidos.ListItems.Item(var_i).ForeColor = &H80000012
          lv_pedidos.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H80000012
          lv_pedidos.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H80000012
          lv_pedidos.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H80000012
          lv_pedidos.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H80000012
          lv_pedidos.ListItems.Item(var_i).ListSubItems(5).ForeColor = &H80000012
          lv_pedidos.ListItems.Item(var_i).ListSubItems(6).ForeColor = &H80000012
          lv_pedidos.ListItems.Item(var_i).ListSubItems(7).ForeColor = &H80000012
          lv_pedidos.ListItems.Item(var_i).ListSubItems(8).ForeColor = &H80000012
       End If
   Next var_i
   lv_pedidos.Refresh
End Sub



Private Sub PASILLOS_4()
                  var_grupo = 0
                  
'pasillo 0,1,2
                  var_consecutivo = var_consecutivo_general
                  var_contador = 0
                  var_veces_grupo = 0
                  var_lote = 0
                  rsaux15.Open "select pedido, ORDEN_PEDIDO from TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque = " + Me.txt_embarque + " order by orden_pedido", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux15.EOF
                  var_grupo = var_grupo + 1
                  var_lote = var_lote + 1
                  var_consecutivo = var_consecutivo_general
                  var_contador = 0
                  var_veces_grupo = 0
                  
                  
                  rsaux1.Open "select distinct source_header_number, ORDEN_PEDIDO from tb_Temp_oracle_orden_surtido_aux_2, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SOURCE_HEADER_NUMBER = pedido AND SUBSTRING(UBICACION,1,3) IN ('P00','B00','P02','P01')  order by ORDEN_pedido", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        
                        If var_veces_grupo = 1 Then
                           var_contador = 60
                           var_veces_grupo = 0
                        End If
                        If var_contador >= 60 Then
                           var_grupo = var_grupo + 1
                           var_contador = 0
                        End If
                        
                        rsaux2.Open "select * from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number = " + CStr(rsaux1!source_header_number) + " AND SUBSTRING(UBICACION,1,3) IN ('P00','B00','P02','P01')   order by ubicacion", cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux2.EOF
                              'If var_veces_grupo = 2 Then
                              '   var_veces_grupo = 1
                              '   var_grupo = var_grupo + 1
                              'End If
                              If var_contador >= 60 Then
                                 var_lote = var_lote + 1
                                 var_contador = 0
                                 var_grupo = var_grupo + 1
                                 'var_veces_grupo = var_veces_grupo + 1
                              End If
                              rsaux3.Open "SELECT * FROM tb_Temp_oracle_orden_surtido_aux_1 WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and delivery_detail_id = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 rsaux4.Open "UPDATE TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 SET SRC_REQUESTED_QUANTITY = SRC_REQUESTED_QUANTITY + " + CStr(rsaux2!src_requested_quantity) + " WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and DELIVERY_DETAIL_ID = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 var_dia = CStr(Day(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_mes = CStr(Month(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_año = CStr(Year(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 If Len(Trim(var_dia)) = 1 Then
                                    var_dia = "0" + var_dia
                                 End If
                                 If Len(Trim(var_mes)) = 1 Then
                                    var_mes = "0" + var_mes
                                 End If
                                 var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                 
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                 var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,LOTE, GRUPO, PASILLO) "
                                 var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux2!source_header_number + "', " + CStr(rsaux2!delivery_id) + "," + CStr(rsaux2!delivery_detail_id) + ", " + CStr(rsaux2!organization_id) + ",'" + IIf(IsNull(rsaux2!subinventory), "", rsaux2!subinventory) + "', " + CStr(rsaux2!delivery_line_id) + "," + CStr(rsaux2!inventory_item_id) + ",'" + rsaux2!item_description + "','" + CStr(rsaux2!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux2!src_requested_quantity) + ",'" + rsaux2!released_status + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!customer_name), "", rsaux2!customer_name) + "','" + rsaux2!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux2!collector_id), 0, rsaux2!collector_id)) + ",'" + IIf(IsNull(rsaux2!Name), "", rsaux2!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux2!ubicacion), "", rsaux2!ubicacion) + "','" + IIf(IsNull(rsaux2!establecimiento), "", rsaux2!establecimiento) + "','" + IIf(IsNull(rsaux2!nombre_Establecimiento), "", rsaux2!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux2!ORDENES), "", rsaux2!ORDENES) + "',"
                                 var_cadena = var_cadena + CStr(IIf(IsNull(rsaux2!CUST_ACCOUNT_ID), 0, rsaux2!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux2!source_header_type_name), "", rsaux2!source_header_type_name) + "', '" + IIf(IsNull(rsaux2!source_document_id), "", rsaux2!source_document_id) + "','" + IIf(IsNull(rsaux2!pais), "", rsaux2!pais) + "','" + IIf(IsNull(rsaux2!estado), "", rsaux2!estado) + "', '" + IIf(IsNull(rsaux2!municipio), "", rsaux2!municipio) + "', '" + IIf(IsNull(rsaux2!ciudad), "", rsaux2!ciudad) + "', '" + IIf(IsNull(rsaux2!colonia), "", rsaux2!colonia) + "','" + IIf(IsNull(rsaux2!DIRECCION), "", rsaux2!DIRECCION) + "', '" + IIf(IsNull(rsaux2!cp), "", rsaux2!cp) + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!site_use_id), "", rsaux2!site_use_id) + "','" + IIf(IsNull(rsaux2!paqueteria), "", rsaux2!paqueteria) + "','" + IIf(IsNull(rsaux2!attribute8), "", rsaux2!attribute8) + "','" + IIf(IsNull(rsaux2!ATTRIBUTE9), "", rsaux2!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux2!TELEFONO), "", rsaux2!TELEFONO) + "','" + IIf(IsNull(rsaux2!Linea), "", rsaux2!Linea) + "','" + CStr(IIf(IsNull(rsaux2!ruta), "", rsaux2!ruta)) + "','" + IIf(IsNull(rsaux2!nombre_ruta), "", rsaux2!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_SURTIDO), 0, rsaux2!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", '" + IIf(IsNull(rsaux2!estacion), "", rsaux2!estacion) + "'," + CStr(var_lote) + "," + CStr(var_grupo) + ",'B00')"
                                 rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux3.Close
                              If rsaux2!Linea <> "CATALOGOS" Then
                                 If rsaux2!Linea <> "CATALOGO" Then
                                    If rsaux2!Linea <> "POP" Then
                                       If rsaux2!Linea <> "EMPAQUE" Then
                                          If rsaux2!MULTIPLO > 1 Then
                                             var_contador = var_contador + rsaux2!src_requested_quantity
                                          Else
                                             var_contador = var_contador + 1
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                              rsaux2.MoveNext
                        Wend
                        rsaux2.Close
                        var_veces_grupo = var_veces_grupo + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  
                  var_consecutivo_ubicacion = 1
                  rsaux1.Open "select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_Consecutivo = " + CStr(var_consecutivo) + " and SUBSTRING(UBICACION,1,3) in ('B00','P00') ORDER BY UBICACION", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update tb_temp_oracle_orden_surtido_aux_1 set consecutivo_pasillo = " + CStr(var_consecutivo_ubicacion) + "  where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ubicacion = '" + rsaux1!ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_consecutivo_ubicacion = var_consecutivo_ubicacion + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close








'PASILLO 3
                  var_consecutivo = var_consecutivo_general
                  rsaux1.Open "select distinct source_header_number, ORDEN_pedido from tb_Temp_oracle_orden_surtido_aux_2, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SOURCE_HEADER_NUMBER = pedido AND SUBSTRING(UBICACION,1,3) IN ('P03')  order by ORDEN_pedido", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux1.EOF Then
                  var_grupo = var_grupo + 1
                  var_lote = var_lote + 1
                  var_contador = 0
                  var_veces_grupo = 0
                  While Not rsaux1.EOF
                        If var_veces_grupo = 1 Then
                           var_contador = 60
                           var_veces_grupo = 0
                        End If
                        If var_contador >= 60 Then
                           var_grupo = var_grupo + 1
                           var_contador = 0
                        End If
                        
                        rsaux2.Open "select * from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number = " + CStr(rsaux1!source_header_number) + " AND SUBSTRING(UBICACION,1,3) IN ('P03')   order by ubicacion", cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux2.EOF
                              If var_contador >= 60 Then
                                 var_lote = var_lote + 1
                                 var_contador = 0
                                 var_grupo = var_grupo + 1
                              End If
                              rsaux3.Open "SELECT * FROM tb_Temp_oracle_orden_surtido_aux_1 WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and delivery_detail_id = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 rsaux4.Open "UPDATE TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 SET SRC_REQUESTED_QUANTITY = SRC_REQUESTED_QUANTITY + " + CStr(rsaux2!src_requested_quantity) + " WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and DELIVERY_DETAIL_ID = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 var_dia = CStr(Day(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_mes = CStr(Month(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_año = CStr(Year(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 If Len(Trim(var_dia)) = 1 Then
                                    var_dia = "0" + var_dia
                                 End If
                                 If Len(Trim(var_mes)) = 1 Then
                                    var_mes = "0" + var_mes
                                 End If
                                 var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                 
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                 var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,LOTE, GRUPO, PASILLO) "
                                 var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux2!source_header_number + "', " + CStr(rsaux2!delivery_id) + "," + CStr(rsaux2!delivery_detail_id) + ", " + CStr(rsaux2!organization_id) + ",'" + IIf(IsNull(rsaux2!subinventory), "", rsaux2!subinventory) + "', " + CStr(rsaux2!delivery_line_id) + "," + CStr(rsaux2!inventory_item_id) + ",'" + rsaux2!item_description + "','" + CStr(rsaux2!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux2!src_requested_quantity) + ",'" + rsaux2!released_status + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!customer_name), "", rsaux2!customer_name) + "','" + rsaux2!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux2!collector_id), 0, rsaux2!collector_id)) + ",'" + IIf(IsNull(rsaux2!Name), "", rsaux2!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux2!ubicacion), "", rsaux2!ubicacion) + "','" + IIf(IsNull(rsaux2!establecimiento), "", rsaux2!establecimiento) + "','" + IIf(IsNull(rsaux2!nombre_Establecimiento), "", rsaux2!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux2!ORDENES), "", rsaux2!ORDENES) + "',"
                                 var_cadena = var_cadena + CStr(IIf(IsNull(rsaux2!CUST_ACCOUNT_ID), 0, rsaux2!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux2!source_header_type_name), "", rsaux2!source_header_type_name) + "', '" + IIf(IsNull(rsaux2!source_document_id), "", rsaux2!source_document_id) + "','" + IIf(IsNull(rsaux2!pais), "", rsaux2!pais) + "','" + IIf(IsNull(rsaux2!estado), "", rsaux2!estado) + "', '" + IIf(IsNull(rsaux2!municipio), "", rsaux2!municipio) + "', '" + IIf(IsNull(rsaux2!ciudad), "", rsaux2!ciudad) + "', '" + IIf(IsNull(rsaux2!colonia), "", rsaux2!colonia) + "','" + IIf(IsNull(rsaux2!DIRECCION), "", rsaux2!DIRECCION) + "', '" + IIf(IsNull(rsaux2!cp), "", rsaux2!cp) + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!site_use_id), "", rsaux2!site_use_id) + "','" + IIf(IsNull(rsaux2!paqueteria), "", rsaux2!paqueteria) + "','" + IIf(IsNull(rsaux2!attribute8), "", rsaux2!attribute8) + "','" + IIf(IsNull(rsaux2!ATTRIBUTE9), "", rsaux2!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux2!TELEFONO), "", rsaux2!TELEFONO) + "','" + IIf(IsNull(rsaux2!Linea), "", rsaux2!Linea) + "','" + CStr(IIf(IsNull(rsaux2!ruta), "", rsaux2!ruta)) + "','" + IIf(IsNull(rsaux2!nombre_ruta), "", rsaux2!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_SURTIDO), 0, rsaux2!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", '" + IIf(IsNull(rsaux2!estacion), "", rsaux2!estacion) + "'," + CStr(var_lote) + "," + CStr(var_grupo) + ",'P01')"
                                 rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux3.Close
                              If rsaux2!Linea <> "CATALOGOS" Then
                                 If rsaux2!Linea <> "CATALOGO" Then
                                    If rsaux2!Linea <> "POP" Then
                                       If rsaux2!Linea <> "EMPAQUE" Then
                                          If rsaux2!MULTIPLO > 1 Then
                                             var_contador = var_contador + rsaux2!src_requested_quantity
                                          Else
                                             var_contador = var_contador + 1
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                              rsaux2.MoveNext
                        Wend
                        rsaux2.Close
                        var_veces_grupo = var_veces_grupo + 1
                        rsaux1.MoveNext
                  Wend
                  
                  rsaux1.Close
                  
                  var_consecutivo_ubicacion = 1
                  rsaux1.Open "select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_Consecutivo = " + CStr(var_consecutivo) + " and SUBSTRING(UBICACION,1,3) in ('B01','P01') ORDER BY UBICACION", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update tb_temp_oracle_orden_surtido_aux_1 set consecutivo_pasillo = " + CStr(var_consecutivo_ubicacion) + "  where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ubicacion = '" + rsaux1!ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_consecutivo_ubicacion = var_consecutivo_ubicacion + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  
                  'var_consecutivo_ubicacion = 1
                  rsaux1.Open "select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_Consecutivo = " + CStr(var_consecutivo) + " and SUBSTRING(UBICACION,1,3) in ('P02') ORDER BY UBICACION DESC", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update tb_temp_oracle_orden_surtido_aux_1 set consecutivo_pasillo = " + CStr(var_consecutivo_ubicacion) + "  where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ubicacion = '" + rsaux1!ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_consecutivo_ubicacion = var_consecutivo_ubicacion + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  Else
                     rsaux1.Close
                  End If
                  
                  
'PASILLO 4
                  var_consecutivo = var_consecutivo_general
                  rsaux1.Open "select distinct source_header_number, ORDEN_pedido from tb_Temp_oracle_orden_surtido_aux_2, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SOURCE_HEADER_NUMBER = pedido AND SUBSTRING(UBICACION,1,3) IN ('P04')  order by ORDEN_pedido", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux1.EOF Then
                  var_grupo = var_grupo + 1
                  var_lote = var_lote + 1
                  var_contador = 0
                  var_veces_grupo = 0
                  While Not rsaux1.EOF
                        If var_veces_grupo = 1 Then
                           var_contador = 60
                           var_veces_grupo = 0
                        End If
                        If var_contador >= 60 Then
                           var_grupo = var_grupo + 1
                           var_contador = 0
                        End If
                        
                        rsaux2.Open "select * from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number = " + CStr(rsaux1!source_header_number) + " AND SUBSTRING(UBICACION,1,3) IN ('P04')   order by ubicacion", cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux2.EOF
                              If var_contador >= 60 Then
                                 var_lote = var_lote + 1
                                 var_contador = 0
                                 var_grupo = var_grupo + 1
                              End If
                              rsaux3.Open "SELECT * FROM tb_Temp_oracle_orden_surtido_aux_1 WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and delivery_detail_id = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 rsaux4.Open "UPDATE TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 SET SRC_REQUESTED_QUANTITY = SRC_REQUESTED_QUANTITY + " + CStr(rsaux2!src_requested_quantity) + " WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and DELIVERY_DETAIL_ID = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 var_dia = CStr(Day(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_mes = CStr(Month(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_año = CStr(Year(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 If Len(Trim(var_dia)) = 1 Then
                                    var_dia = "0" + var_dia
                                 End If
                                 If Len(Trim(var_mes)) = 1 Then
                                    var_mes = "0" + var_mes
                                 End If
                                 var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                 
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                 var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,LOTE, GRUPO, PASILLO) "
                                 var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux2!source_header_number + "', " + CStr(rsaux2!delivery_id) + "," + CStr(rsaux2!delivery_detail_id) + ", " + CStr(rsaux2!organization_id) + ",'" + IIf(IsNull(rsaux2!subinventory), "", rsaux2!subinventory) + "', " + CStr(rsaux2!delivery_line_id) + "," + CStr(rsaux2!inventory_item_id) + ",'" + rsaux2!item_description + "','" + CStr(rsaux2!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux2!src_requested_quantity) + ",'" + rsaux2!released_status + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!customer_name), "", rsaux2!customer_name) + "','" + rsaux2!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux2!collector_id), 0, rsaux2!collector_id)) + ",'" + IIf(IsNull(rsaux2!Name), "", rsaux2!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux2!ubicacion), "", rsaux2!ubicacion) + "','" + IIf(IsNull(rsaux2!establecimiento), "", rsaux2!establecimiento) + "','" + IIf(IsNull(rsaux2!nombre_Establecimiento), "", rsaux2!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux2!ORDENES), "", rsaux2!ORDENES) + "',"
                                 var_cadena = var_cadena + CStr(IIf(IsNull(rsaux2!CUST_ACCOUNT_ID), 0, rsaux2!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux2!source_header_type_name), "", rsaux2!source_header_type_name) + "', '" + IIf(IsNull(rsaux2!source_document_id), "", rsaux2!source_document_id) + "','" + IIf(IsNull(rsaux2!pais), "", rsaux2!pais) + "','" + IIf(IsNull(rsaux2!estado), "", rsaux2!estado) + "', '" + IIf(IsNull(rsaux2!municipio), "", rsaux2!municipio) + "', '" + IIf(IsNull(rsaux2!ciudad), "", rsaux2!ciudad) + "', '" + IIf(IsNull(rsaux2!colonia), "", rsaux2!colonia) + "','" + IIf(IsNull(rsaux2!DIRECCION), "", rsaux2!DIRECCION) + "', '" + IIf(IsNull(rsaux2!cp), "", rsaux2!cp) + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!site_use_id), "", rsaux2!site_use_id) + "','" + IIf(IsNull(rsaux2!paqueteria), "", rsaux2!paqueteria) + "','" + IIf(IsNull(rsaux2!attribute8), "", rsaux2!attribute8) + "','" + IIf(IsNull(rsaux2!ATTRIBUTE9), "", rsaux2!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux2!TELEFONO), "", rsaux2!TELEFONO) + "','" + IIf(IsNull(rsaux2!Linea), "", rsaux2!Linea) + "','" + CStr(IIf(IsNull(rsaux2!ruta), "", rsaux2!ruta)) + "','" + IIf(IsNull(rsaux2!nombre_ruta), "", rsaux2!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_SURTIDO), 0, rsaux2!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", '" + IIf(IsNull(rsaux2!estacion), "", rsaux2!estacion) + "'," + CStr(var_lote) + "," + CStr(var_grupo) + ",'P02')"
                                 rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux3.Close
                              If rsaux2!Linea <> "CATALOGOS" Then
                                 If rsaux2!Linea <> "CATALOGO" Then
                                    If rsaux2!Linea <> "POP" Then
                                       If rsaux2!Linea <> "EMPAQUE" Then
                                          If rsaux2!MULTIPLO > 1 Then
                                             var_contador = var_contador + rsaux2!src_requested_quantity
                                          Else
                                             var_contador = var_contador + 1
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                              rsaux2.MoveNext
                        Wend
                        rsaux2.Close
                        var_veces_grupo = var_veces_grupo + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  
                  var_consecutivo_ubicacion = 1
                  rsaux1.Open "select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_Consecutivo = " + CStr(var_consecutivo) + " and SUBSTRING(UBICACION,1,3) in ('B01','P01') ORDER BY UBICACION", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update tb_temp_oracle_orden_surtido_aux_1 set consecutivo_pasillo = " + CStr(var_consecutivo_ubicacion) + "  where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ubicacion = '" + rsaux1!ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_consecutivo_ubicacion = var_consecutivo_ubicacion + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  
                  'var_consecutivo_ubicacion = 1
                  rsaux1.Open "select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_Consecutivo = " + CStr(var_consecutivo) + " and SUBSTRING(UBICACION,1,3) in ('P02') ORDER BY UBICACION DESC", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update tb_temp_oracle_orden_surtido_aux_1 set consecutivo_pasillo = " + CStr(var_consecutivo_ubicacion) + "  where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ubicacion = '" + rsaux1!ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_consecutivo_ubicacion = var_consecutivo_ubicacion + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  Else
                     rsaux1.Close
                  End If
                  
                  
                  
                  
                  
                  
                  
                  
'PASILLO 5, 6
                  rsaux1.Open "select distinct source_header_number, ORDEN_pedido from tb_Temp_oracle_orden_surtido_aux_2, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SOURCE_HEADER_NUMBER = pedido AND SUBSTRING(UBICACION,1,3) IN ('P05','P06')  order by ORDEN_pedido", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux1.EOF Then
                  var_grupo = var_grupo + 1
                  var_lote = var_lote + 1
                  var_contador = 0
                  var_veces_grupo = 0
                  While Not rsaux1.EOF
                        If var_veces_grupo = 1 Then
                           var_contador = 30
                           var_veces_grupo = 0
                        End If
                        If var_contador >= 30 Then
                           var_grupo = var_grupo + 1
                           var_contador = 0
                        End If
                        rsaux2.Open "select * from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number = " + CStr(rsaux1!source_header_number) + " AND SUBSTRING(UBICACION,1,3) IN ('P05','P06')  order by ubicacion", cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux2.EOF
                              If var_contador >= 30 Then
                                 var_lote = var_lote + 1
                                 var_contador = 0
                                 var_grupo = var_grupo + 1
                              End If
                              rsaux3.Open "SELECT * FROM tb_Temp_oracle_orden_surtido_aux_1 WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and delivery_detail_id = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 rsaux4.Open "UPDATE TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 SET SRC_REQUESTED_QUANTITY = SRC_REQUESTED_QUANTITY + " + CStr(rsaux2!src_requested_quantity) + " WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and DELIVERY_DETAIL_ID = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 var_dia = CStr(Day(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_mes = CStr(Month(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_año = CStr(Year(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 If Len(Trim(var_dia)) = 1 Then
                                    var_dia = "0" + var_dia
                                 End If
                                 If Len(Trim(var_mes)) = 1 Then
                                    var_mes = "0" + var_mes
                                 End If
                                 var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                 
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                 var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,LOTE, GRUPO, PASILLO) "
                                 var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux2!source_header_number + "', " + CStr(rsaux2!delivery_id) + "," + CStr(rsaux2!delivery_detail_id) + ", " + CStr(rsaux2!organization_id) + ",'" + IIf(IsNull(rsaux2!subinventory), "", rsaux2!subinventory) + "', " + CStr(rsaux2!delivery_line_id) + "," + CStr(rsaux2!inventory_item_id) + ",'" + rsaux2!item_description + "','" + CStr(rsaux2!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux2!src_requested_quantity) + ",'" + rsaux2!released_status + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!customer_name), "", rsaux2!customer_name) + "','" + rsaux2!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux2!collector_id), 0, rsaux2!collector_id)) + ",'" + IIf(IsNull(rsaux2!Name), "", rsaux2!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux2!ubicacion), "", rsaux2!ubicacion) + "','" + IIf(IsNull(rsaux2!establecimiento), "", rsaux2!establecimiento) + "','" + IIf(IsNull(rsaux2!nombre_Establecimiento), "", rsaux2!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux2!ORDENES), "", rsaux2!ORDENES) + "',"
                                 var_cadena = var_cadena + CStr(IIf(IsNull(rsaux2!CUST_ACCOUNT_ID), 0, rsaux2!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux2!source_header_type_name), "", rsaux2!source_header_type_name) + "', '" + IIf(IsNull(rsaux2!source_document_id), "", rsaux2!source_document_id) + "','" + IIf(IsNull(rsaux2!pais), "", rsaux2!pais) + "','" + IIf(IsNull(rsaux2!estado), "", rsaux2!estado) + "', '" + IIf(IsNull(rsaux2!municipio), "", rsaux2!municipio) + "', '" + IIf(IsNull(rsaux2!ciudad), "", rsaux2!ciudad) + "', '" + IIf(IsNull(rsaux2!colonia), "", rsaux2!colonia) + "','" + IIf(IsNull(rsaux2!DIRECCION), "", rsaux2!DIRECCION) + "', '" + IIf(IsNull(rsaux2!cp), "", rsaux2!cp) + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!site_use_id), "", rsaux2!site_use_id) + "','" + IIf(IsNull(rsaux2!paqueteria), "", rsaux2!paqueteria) + "','" + IIf(IsNull(rsaux2!attribute8), "", rsaux2!attribute8) + "','" + IIf(IsNull(rsaux2!ATTRIBUTE9), "", rsaux2!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux2!TELEFONO), "", rsaux2!TELEFONO) + "','" + IIf(IsNull(rsaux2!Linea), "", rsaux2!Linea) + "','" + CStr(IIf(IsNull(rsaux2!ruta), "", rsaux2!ruta)) + "','" + IIf(IsNull(rsaux2!nombre_ruta), "", rsaux2!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_SURTIDO), 0, rsaux2!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", '" + IIf(IsNull(rsaux2!estacion), "", rsaux2!estacion) + "'," + CStr(var_lote) + "," + CStr(var_grupo) + ",'P03')"
                                 rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux3.Close
                              If rsaux2!Linea <> "CATALOGOS" Then
                                 If rsaux2!Linea <> "CATALOGO" Then
                                    If rsaux2!Linea <> "POP" Then
                                       If rsaux2!Linea <> "EMPAQUE" Then
                                          If rsaux2!MULTIPLO > 1 Then
                                             var_contador = var_contador + rsaux2!src_requested_quantity
                                          Else
                                             var_contador = var_contador + 1
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                              rsaux2.MoveNext
                        Wend
                        rsaux2.Close
                        var_veces_grupo = var_veces_grupo + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  
                  var_consecutivo_ubicacion = 1
                  rsaux1.Open "select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_Consecutivo = " + CStr(var_consecutivo) + " and SUBSTRING(UBICACION,1,3) in ('P03') ORDER BY UBICACION DESC", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update tb_temp_oracle_orden_surtido_aux_1 set consecutivo_pasillo = " + CStr(var_consecutivo_ubicacion) + "  where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ubicacion = '" + rsaux1!ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_consecutivo_ubicacion = var_consecutivo_ubicacion + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  Else
                     rsaux1.Close
                  End If
                  
                  
'PASILLO 8,9
                  rsaux1.Open "select distinct source_header_number, ORDEN_pedido from tb_Temp_oracle_orden_surtido_aux_2, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SOURCE_HEADER_NUMBER = pedido AND SUBSTRING(UBICACION,1,3) IN ('P08','P09')  order by ORDEN_pedido", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux1.EOF Then
                  var_grupo = var_grupo + 1
                  var_lote = var_lote + 1
                  var_contador = 0
                  var_veces_grupo = 0
                  While Not rsaux1.EOF
                        If var_veces_grupo = 1 Then
                           var_contador = 30
                           var_veces_grupo = 0
                        End If
                        If var_contador >= 30 Then
                           var_grupo = var_grupo + 1
                           var_contador = 0
                        End If
                        rsaux2.Open "select * from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number = " + CStr(rsaux1!source_header_number) + " AND SUBSTRING(UBICACION,1,3) IN ('P08','P09')   order by ubicacion", cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux2.EOF
                              If var_contador >= 30 Then
                                 var_lote = var_lote + 1
                                 var_contador = 0
                                 var_grupo = var_grupo + 1
                              End If
                              rsaux3.Open "SELECT * FROM tb_Temp_oracle_orden_surtido_aux_1 WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and delivery_detail_id = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 rsaux4.Open "UPDATE TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 SET SRC_REQUESTED_QUANTITY = SRC_REQUESTED_QUANTITY + " + CStr(rsaux2!src_requested_quantity) + " WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and DELIVERY_DETAIL_ID = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 var_dia = CStr(Day(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_mes = CStr(Month(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_año = CStr(Year(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 If Len(Trim(var_dia)) = 1 Then
                                    var_dia = "0" + var_dia
                                 End If
                                 If Len(Trim(var_mes)) = 1 Then
                                    var_mes = "0" + var_mes
                                 End If
                                 var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                 
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                 var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,LOTE, GRUPO, PASILLO) "
                                 var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux2!source_header_number + "', " + CStr(rsaux2!delivery_id) + "," + CStr(rsaux2!delivery_detail_id) + ", " + CStr(rsaux2!organization_id) + ",'" + IIf(IsNull(rsaux2!subinventory), "", rsaux2!subinventory) + "', " + CStr(rsaux2!delivery_line_id) + "," + CStr(rsaux2!inventory_item_id) + ",'" + rsaux2!item_description + "','" + CStr(rsaux2!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux2!src_requested_quantity) + ",'" + rsaux2!released_status + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!customer_name), "", rsaux2!customer_name) + "','" + rsaux2!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux2!collector_id), 0, rsaux2!collector_id)) + ",'" + IIf(IsNull(rsaux2!Name), "", rsaux2!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux2!ubicacion), "", rsaux2!ubicacion) + "','" + IIf(IsNull(rsaux2!establecimiento), "", rsaux2!establecimiento) + "','" + IIf(IsNull(rsaux2!nombre_Establecimiento), "", rsaux2!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux2!ORDENES), "", rsaux2!ORDENES) + "',"
                                 var_cadena = var_cadena + CStr(IIf(IsNull(rsaux2!CUST_ACCOUNT_ID), 0, rsaux2!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux2!source_header_type_name), "", rsaux2!source_header_type_name) + "', '" + IIf(IsNull(rsaux2!source_document_id), "", rsaux2!source_document_id) + "','" + IIf(IsNull(rsaux2!pais), "", rsaux2!pais) + "','" + IIf(IsNull(rsaux2!estado), "", rsaux2!estado) + "', '" + IIf(IsNull(rsaux2!municipio), "", rsaux2!municipio) + "', '" + IIf(IsNull(rsaux2!ciudad), "", rsaux2!ciudad) + "', '" + IIf(IsNull(rsaux2!colonia), "", rsaux2!colonia) + "','" + IIf(IsNull(rsaux2!DIRECCION), "", rsaux2!DIRECCION) + "', '" + IIf(IsNull(rsaux2!cp), "", rsaux2!cp) + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!site_use_id), "", rsaux2!site_use_id) + "','" + IIf(IsNull(rsaux2!paqueteria), "", rsaux2!paqueteria) + "','" + IIf(IsNull(rsaux2!attribute8), "", rsaux2!attribute8) + "','" + IIf(IsNull(rsaux2!ATTRIBUTE9), "", rsaux2!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux2!TELEFONO), "", rsaux2!TELEFONO) + "','" + IIf(IsNull(rsaux2!Linea), "", rsaux2!Linea) + "','" + CStr(IIf(IsNull(rsaux2!ruta), "", rsaux2!ruta)) + "','" + IIf(IsNull(rsaux2!nombre_ruta), "", rsaux2!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_SURTIDO), 0, rsaux2!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", '" + IIf(IsNull(rsaux2!estacion), "", rsaux2!estacion) + "'," + CStr(var_lote) + "," + CStr(var_grupo) + ",'P05')"
                                 rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux3.Close
                              If rsaux2!Linea <> "CATALOGOS" Then
                                 If rsaux2!Linea <> "CATALOGO" Then
                                    If rsaux2!Linea <> "POP" Then
                                       If rsaux2!Linea <> "EMPAQUE" Then
                                          If rsaux2!MULTIPLO > 1 Then
                                             var_contador = var_contador + rsaux2!src_requested_quantity
                                          Else
                                             var_contador = var_contador + 1
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                              rsaux2.MoveNext
                        Wend
                        rsaux2.Close
                        var_veces_grupo = var_veces_grupo + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close

                  var_consecutivo_ubicacion = 1
                  rsaux1.Open "select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_Consecutivo = " + CStr(var_consecutivo) + " and SUBSTRING(UBICACION,1,3) in ('P04') ORDER BY UBICACION DESC", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update tb_temp_oracle_orden_surtido_aux_1 set consecutivo_pasillo = " + CStr(var_consecutivo_ubicacion) + "  where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ubicacion = '" + rsaux1!ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_consecutivo_ubicacion = var_consecutivo_ubicacion + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  Else
                     rsaux1.Close
                  End If
                  
                  
                  
                  
                  
'PASILLO 7
                  rsaux1.Open "select distinct source_header_number, ORDEN_pedido from tb_Temp_oracle_orden_surtido_aux_2, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SOURCE_HEADER_NUMBER = pedido AND SUBSTRING(UBICACION,1,3) IN ('P07')  order by ORDEN_pedido", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux1.EOF Then
                  var_grupo = var_grupo + 1
                  var_lote = var_lote + 1
                  var_contador = 0
                  var_veces_grupo = 0
                  While Not rsaux1.EOF
                        If var_veces_grupo = 1 Then
                           var_contador = 30
                           var_veces_grupo = 0
                        End If
                        If var_contador >= 30 Then
                           var_grupo = var_grupo + 1
                           var_contador = 0
                        End If
                        rsaux2.Open "select * from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number = " + CStr(rsaux1!source_header_number) + " AND SUBSTRING(UBICACION,1,3) IN ('P07')  order by ubicacion", cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux2.EOF
                              If var_contador >= 30 Then
                                 var_lote = var_lote + 1
                                 var_contador = 0
                                 var_grupo = var_grupo + 1
                              End If
                              rsaux3.Open "SELECT * FROM tb_Temp_oracle_orden_surtido_aux_1 WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and delivery_detail_id = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 rsaux4.Open "UPDATE TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 SET SRC_REQUESTED_QUANTITY = SRC_REQUESTED_QUANTITY + " + CStr(rsaux2!src_requested_quantity) + " WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and DELIVERY_DETAIL_ID = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 var_dia = CStr(Day(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_mes = CStr(Month(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_año = CStr(Year(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 If Len(Trim(var_dia)) = 1 Then
                                    var_dia = "0" + var_dia
                                 End If
                                 If Len(Trim(var_mes)) = 1 Then
                                    var_mes = "0" + var_mes
                                 End If
                                 var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                 
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                 var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,LOTE, GRUPO, PASILLO) "
                                 var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux2!source_header_number + "', " + CStr(rsaux2!delivery_id) + "," + CStr(rsaux2!delivery_detail_id) + ", " + CStr(rsaux2!organization_id) + ",'" + IIf(IsNull(rsaux2!subinventory), "", rsaux2!subinventory) + "', " + CStr(rsaux2!delivery_line_id) + "," + CStr(rsaux2!inventory_item_id) + ",'" + rsaux2!item_description + "','" + CStr(rsaux2!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux2!src_requested_quantity) + ",'" + rsaux2!released_status + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!customer_name), "", rsaux2!customer_name) + "','" + rsaux2!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux2!collector_id), 0, rsaux2!collector_id)) + ",'" + IIf(IsNull(rsaux2!Name), "", rsaux2!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux2!ubicacion), "", rsaux2!ubicacion) + "','" + IIf(IsNull(rsaux2!establecimiento), "", rsaux2!establecimiento) + "','" + IIf(IsNull(rsaux2!nombre_Establecimiento), "", rsaux2!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux2!ORDENES), "", rsaux2!ORDENES) + "',"
                                 var_cadena = var_cadena + CStr(IIf(IsNull(rsaux2!CUST_ACCOUNT_ID), 0, rsaux2!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux2!source_header_type_name), "", rsaux2!source_header_type_name) + "', '" + IIf(IsNull(rsaux2!source_document_id), "", rsaux2!source_document_id) + "','" + IIf(IsNull(rsaux2!pais), "", rsaux2!pais) + "','" + IIf(IsNull(rsaux2!estado), "", rsaux2!estado) + "', '" + IIf(IsNull(rsaux2!municipio), "", rsaux2!municipio) + "', '" + IIf(IsNull(rsaux2!ciudad), "", rsaux2!ciudad) + "', '" + IIf(IsNull(rsaux2!colonia), "", rsaux2!colonia) + "','" + IIf(IsNull(rsaux2!DIRECCION), "", rsaux2!DIRECCION) + "', '" + IIf(IsNull(rsaux2!cp), "", rsaux2!cp) + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!site_use_id), "", rsaux2!site_use_id) + "','" + IIf(IsNull(rsaux2!paqueteria), "", rsaux2!paqueteria) + "','" + IIf(IsNull(rsaux2!attribute8), "", rsaux2!attribute8) + "','" + IIf(IsNull(rsaux2!ATTRIBUTE9), "", rsaux2!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux2!TELEFONO), "", rsaux2!TELEFONO) + "','" + IIf(IsNull(rsaux2!Linea), "", rsaux2!Linea) + "','" + CStr(IIf(IsNull(rsaux2!ruta), "", rsaux2!ruta)) + "','" + IIf(IsNull(rsaux2!nombre_ruta), "", rsaux2!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_SURTIDO), 0, rsaux2!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", '" + IIf(IsNull(rsaux2!estacion), "", rsaux2!estacion) + "'," + CStr(var_lote) + "," + CStr(var_grupo) + ",'P06')"
                                 rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux3.Close
                              If rsaux2!Linea <> "CATALOGOS" Then
                                 If rsaux2!Linea <> "CATALOGO" Then
                                    If rsaux2!Linea <> "POP" Then
                                       If rsaux2!Linea <> "EMPAQUE" Then
                                          If rsaux2!MULTIPLO > 1 Then
                                             var_contador = var_contador + rsaux2!src_requested_quantity
                                          Else
                                             var_contador = var_contador + 1
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                              rsaux2.MoveNext
                        Wend
                        rsaux2.Close
                        var_veces_grupo = var_veces_grupo + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  
                  var_consecutivo_ubicacion = 1
                  rsaux1.Open "select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_Consecutivo = " + CStr(var_consecutivo) + " and SUBSTRING(UBICACION,1,3) in ('P03') ORDER BY UBICACION DESC", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update tb_temp_oracle_orden_surtido_aux_1 set consecutivo_pasillo = " + CStr(var_consecutivo_ubicacion) + "  where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ubicacion = '" + rsaux1!ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_consecutivo_ubicacion = var_consecutivo_ubicacion + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  Else
                     rsaux1.Close
                  End If
                  
                  
                  rsaux1.Open "select distinct source_header_number, ORDEN_pedido from tb_Temp_oracle_orden_surtido_aux_2, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SOURCE_HEADER_NUMBER = pedido AND SUBSTRING(UBICACION,1,3) not IN ('B00', 'B01', 'P01', 'P02', 'P03', 'P04', 'P05', 'P06', 'P07', 'P08', 'P09','P00')  order by ORDEN_pedido", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux1.EOF Then
                  var_consecutivo = var_consecutivo_general
                  var_grupo = var_grupo + 1
                  var_lote = var_lote + 1
                  var_contador = 0
                  var_veces_grupo = 0
                  While Not rsaux1.EOF
                        'var_lote = 1
                        
                        If var_veces_grupo = 1 Then
                           var_contador = 60
                           'var_grupo = var_grupo + 1
                           var_veces_grupo = 0
                        End If
                        If var_contador >= 60 Then
                           var_grupo = var_grupo + 1
                           'var_veces_grupo = var_veces_grupo + 1
                           var_contador = 0
                        End If
                        
                        rsaux2.Open "select * from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number = " + CStr(rsaux1!source_header_number) + " AND SUBSTRING(UBICACION,1,3) NOT IN ('B00', 'B01', 'P01', 'P02', 'P03', 'P04', 'P05', 'P06', 'P07', 'P08', 'P09','P00')   order by ubicacion", cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux2.EOF
                              'If var_veces_grupo = 2 Then
                              '   var_veces_grupo = 1
                              '   var_grupo = var_grupo + 1
                              'End If
                              If var_contador >= 60 Then
                                 var_lote = var_lote + 1
                                 var_contador = 0
                                 var_grupo = var_grupo + 1
                                 'var_veces_grupo = var_veces_grupo + 1
                              End If
                              rsaux3.Open "SELECT * FROM tb_Temp_oracle_orden_surtido_aux_1 WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and delivery_detail_id = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 rsaux4.Open "UPDATE TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 SET SRC_REQUESTED_QUANTITY = SRC_REQUESTED_QUANTITY + " + CStr(rsaux2!src_requested_quantity) + " WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and DELIVERY_DETAIL_ID = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 var_dia = CStr(Day(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_mes = CStr(Month(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_año = CStr(Year(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 If Len(Trim(var_dia)) = 1 Then
                                    var_dia = "0" + var_dia
                                 End If
                                 If Len(Trim(var_mes)) = 1 Then
                                    var_mes = "0" + var_mes
                                 End If
                                 var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                 
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                 var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,LOTE, GRUPO, PASILLO) "
                                 var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux2!source_header_number + "', " + CStr(rsaux2!delivery_id) + "," + CStr(rsaux2!delivery_detail_id) + ", " + CStr(rsaux2!organization_id) + ",'" + IIf(IsNull(rsaux2!subinventory), "", rsaux2!subinventory) + "', " + CStr(rsaux2!delivery_line_id) + "," + CStr(rsaux2!inventory_item_id) + ",'" + rsaux2!item_description + "','" + CStr(rsaux2!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux2!src_requested_quantity) + ",'" + rsaux2!released_status + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!customer_name), "", rsaux2!customer_name) + "','" + rsaux2!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux2!collector_id), 0, rsaux2!collector_id)) + ",'" + IIf(IsNull(rsaux2!Name), "", rsaux2!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux2!ubicacion), "", rsaux2!ubicacion) + "','" + IIf(IsNull(rsaux2!establecimiento), "", rsaux2!establecimiento) + "','" + IIf(IsNull(rsaux2!nombre_Establecimiento), "", rsaux2!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux2!ORDENES), "", rsaux2!ORDENES) + "',"
                                 var_cadena = var_cadena + CStr(IIf(IsNull(rsaux2!CUST_ACCOUNT_ID), 0, rsaux2!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux2!source_header_type_name), "", rsaux2!source_header_type_name) + "', '" + IIf(IsNull(rsaux2!source_document_id), "", rsaux2!source_document_id) + "','" + IIf(IsNull(rsaux2!pais), "", rsaux2!pais) + "','" + IIf(IsNull(rsaux2!estado), "", rsaux2!estado) + "', '" + IIf(IsNull(rsaux2!municipio), "", rsaux2!municipio) + "', '" + IIf(IsNull(rsaux2!ciudad), "", rsaux2!ciudad) + "', '" + IIf(IsNull(rsaux2!colonia), "", rsaux2!colonia) + "','" + IIf(IsNull(rsaux2!DIRECCION), "", rsaux2!DIRECCION) + "', '" + IIf(IsNull(rsaux2!cp), "", rsaux2!cp) + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!site_use_id), "", rsaux2!site_use_id) + "','" + IIf(IsNull(rsaux2!paqueteria), "", rsaux2!paqueteria) + "','" + IIf(IsNull(rsaux2!attribute8), "", rsaux2!attribute8) + "','" + IIf(IsNull(rsaux2!ATTRIBUTE9), "", rsaux2!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux2!TELEFONO), "", rsaux2!TELEFONO) + "','" + IIf(IsNull(rsaux2!Linea), "", rsaux2!Linea) + "','" + CStr(IIf(IsNull(rsaux2!ruta), "", rsaux2!ruta)) + "','" + IIf(IsNull(rsaux2!nombre_ruta), "", rsaux2!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_SURTIDO), 0, rsaux2!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", '" + IIf(IsNull(rsaux2!estacion), "", rsaux2!estacion) + "'," + CStr(var_lote) + "," + CStr(var_grupo) + ",'P09')"
                                 rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux3.Close
                              If rsaux2!Linea <> "CATALOGOS" Then
                                 If rsaux2!Linea <> "CATALOGO" Then
                                    If rsaux2!Linea <> "POP" Then
                                       If rsaux2!Linea <> "EMPAQUE" Then
                                          If rsaux2!MULTIPLO > 1 Then
                                             var_contador = var_contador + rsaux2!src_requested_quantity
                                          Else
                                             var_contador = var_contador + 1
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                              rsaux2.MoveNext
                        Wend
                        rsaux2.Close
                        var_veces_grupo = var_veces_grupo + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  
                  var_consecutivo_ubicacion = 1
                  rsaux1.Open "select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_Consecutivo = " + CStr(var_consecutivo) + " and SUBSTRING(UBICACION,1,3) NOT in ('B00', 'B01', 'P01', 'P02', 'P03', 'P04', 'P05', 'P06', 'P07', 'P08', 'P09','P00') ORDER BY UBICACION", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update tb_temp_oracle_orden_surtido_aux_1 set consecutivo_pasillo = " + CStr(var_consecutivo_ubicacion) + "  where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ubicacion = '" + rsaux1!ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_consecutivo_ubicacion = var_consecutivo_ubicacion + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  Else
                     rsaux1.Close
                  End If
                  
                  
                  
                  
               rsaux15.MoveNext
               Wend
               rsaux15.Close
                  
                  

End Sub




Private Sub PASILLOS_2()
                  var_grupo = 0
                  
'pasillo 0
                  var_consecutivo = var_consecutivo_general
                  var_contador = 0
                  var_veces_grupo = 0
                  var_lote = 0
                  rsaux15.Open "select pedido, ORDEN_PEDIDO from TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque = " + Me.txt_embarque + " order by orden_pedido, pedido", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux15.EOF
                  var_grupo = var_grupo + 1
                  var_lote = var_lote + 1
                  var_consecutivo = var_consecutivo_general
                  var_contador = 0
                  var_veces_grupo = 0
                  
                  
                  rsaux1.Open "select distinct source_header_number, ORDEN_PEDIDO from tb_Temp_oracle_orden_surtido_aux_2, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SOURCE_HEADER_NUMBER = pedido AND SUBSTRING(UBICACION,1,3) IN ('P00','B00')  order by ORDEN_pedido", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        
                        If var_veces_grupo = 1 Then
                           var_contador = 30
                           var_veces_grupo = 0
                        End If
                        If var_contador >= 60 Then
                           var_grupo = var_grupo + 1
                           var_contador = 0
                        End If
                        
                        rsaux2.Open "select * from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number = " + CStr(rsaux1!source_header_number) + " AND SUBSTRING(UBICACION,1,3) IN ('P00','B00')   order by ubicacion", cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux2.EOF
                              'If var_veces_grupo = 2 Then
                              '   var_veces_grupo = 1
                              '   var_grupo = var_grupo + 1
                              'End If
                              If var_contador >= 60 Then
                                 var_lote = var_lote + 1
                                 var_contador = 0
                                 var_grupo = var_grupo + 1
                                 'var_veces_grupo = var_veces_grupo + 1
                              End If
                              rsaux3.Open "SELECT * FROM tb_Temp_oracle_orden_surtido_aux_1 WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and delivery_detail_id = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 rsaux4.Open "UPDATE TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 SET SRC_REQUESTED_QUANTITY = SRC_REQUESTED_QUANTITY + " + CStr(rsaux2!src_requested_quantity) + " WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and DELIVERY_DETAIL_ID = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 var_dia = CStr(Day(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_mes = CStr(Month(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_año = CStr(Year(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 If Len(Trim(var_dia)) = 1 Then
                                    var_dia = "0" + var_dia
                                 End If
                                 If Len(Trim(var_mes)) = 1 Then
                                    var_mes = "0" + var_mes
                                 End If
                                 var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                 
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                 var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,LOTE, GRUPO, PASILLO) "
                                 var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux2!source_header_number + "', " + CStr(rsaux2!delivery_id) + "," + CStr(rsaux2!delivery_detail_id) + ", " + CStr(rsaux2!organization_id) + ",'" + IIf(IsNull(rsaux2!subinventory), "", rsaux2!subinventory) + "', " + CStr(rsaux2!delivery_line_id) + "," + CStr(rsaux2!inventory_item_id) + ",'" + rsaux2!item_description + "','" + CStr(rsaux2!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux2!src_requested_quantity) + ",'" + rsaux2!released_status + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!customer_name), "", rsaux2!customer_name) + "','" + rsaux2!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux2!collector_id), 0, rsaux2!collector_id)) + ",'" + IIf(IsNull(rsaux2!Name), "", rsaux2!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux2!ubicacion), "", rsaux2!ubicacion) + "','" + IIf(IsNull(rsaux2!establecimiento), "", rsaux2!establecimiento) + "','" + IIf(IsNull(rsaux2!nombre_Establecimiento), "", rsaux2!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux2!ORDENES), "", rsaux2!ORDENES) + "',"
                                 var_cadena = var_cadena + CStr(IIf(IsNull(rsaux2!CUST_ACCOUNT_ID), 0, rsaux2!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux2!source_header_type_name), "", rsaux2!source_header_type_name) + "', '" + IIf(IsNull(rsaux2!source_document_id), "", rsaux2!source_document_id) + "','" + IIf(IsNull(rsaux2!pais), "", rsaux2!pais) + "','" + IIf(IsNull(rsaux2!estado), "", rsaux2!estado) + "', '" + IIf(IsNull(rsaux2!municipio), "", rsaux2!municipio) + "', '" + IIf(IsNull(rsaux2!ciudad), "", rsaux2!ciudad) + "', '" + IIf(IsNull(rsaux2!colonia), "", rsaux2!colonia) + "','" + IIf(IsNull(rsaux2!DIRECCION), "", rsaux2!DIRECCION) + "', '" + IIf(IsNull(rsaux2!cp), "", rsaux2!cp) + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!site_use_id), "", rsaux2!site_use_id) + "','" + IIf(IsNull(rsaux2!paqueteria), "", rsaux2!paqueteria) + "','" + IIf(IsNull(rsaux2!attribute8), "", rsaux2!attribute8) + "','" + IIf(IsNull(rsaux2!ATTRIBUTE9), "", rsaux2!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux2!TELEFONO), "", rsaux2!TELEFONO) + "','" + IIf(IsNull(rsaux2!Linea), "", rsaux2!Linea) + "','" + CStr(IIf(IsNull(rsaux2!ruta), "", rsaux2!ruta)) + "','" + IIf(IsNull(rsaux2!nombre_ruta), "", rsaux2!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_SURTIDO), 0, rsaux2!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", '" + IIf(IsNull(rsaux2!estacion), "", rsaux2!estacion) + "'," + CStr(var_lote) + "," + CStr(var_grupo) + ",'B00')"
                                 rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux3.Close
                              If rsaux2!Linea <> "CATALOGOS" Then
                                 If rsaux2!Linea <> "CATALOGO" Then
                                    If rsaux2!Linea <> "POP" Then
                                       If rsaux2!Linea <> "EMPAQUE" Then
                                          If rsaux2!MULTIPLO > 1 Then
                                             var_contador = var_contador + rsaux2!src_requested_quantity
                                          Else
                                             var_contador = var_contador + 1
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                              rsaux2.MoveNext
                        Wend
                        rsaux2.Close
                        var_veces_grupo = var_veces_grupo + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  
                  var_consecutivo_ubicacion = 1
                  rsaux1.Open "select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_Consecutivo = " + CStr(var_consecutivo) + " and SUBSTRING(UBICACION,1,3) in ('B00','P00') ORDER BY UBICACION", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update tb_temp_oracle_orden_surtido_aux_1 set consecutivo_pasillo = " + CStr(var_consecutivo_ubicacion) + "  where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ubicacion = '" + rsaux1!ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_consecutivo_ubicacion = var_consecutivo_ubicacion + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close




'pasillo 1, 2


                  var_grupo = var_grupo + 1
                  var_lote = var_lote + 1
                  var_consecutivo = var_consecutivo_general
                  rsaux1.Open "select distinct source_header_number, ORDEN_pedido from tb_Temp_oracle_orden_surtido_aux_2, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SOURCE_HEADER_NUMBER = pedido AND SUBSTRING(UBICACION,1,3) IN ('P01','P02')  order by ORDEN_pedido", cnn, adOpenDynamic, adLockOptimistic
                  var_contador = 0
                  var_veces_grupo = 0
                  While Not rsaux1.EOF
                        If var_veces_grupo = 1 Then
                           var_contador = 60
                           var_veces_grupo = 0
                        End If
                        If var_contador >= 60 Then
                           var_grupo = var_grupo + 1
                           var_contador = 0
                        End If
                        
                        rsaux2.Open "select * from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number = " + CStr(rsaux1!source_header_number) + " AND SUBSTRING(UBICACION,1,3) IN ('P01','P02')   order by ubicacion", cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux2.EOF
                              If var_contador >= 60 Then
                                 var_lote = var_lote + 1
                                 var_contador = 0
                                 var_grupo = var_grupo + 1
                              End If
                              rsaux3.Open "SELECT * FROM tb_Temp_oracle_orden_surtido_aux_1 WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and delivery_detail_id = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 rsaux4.Open "UPDATE TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 SET SRC_REQUESTED_QUANTITY = SRC_REQUESTED_QUANTITY + " + CStr(rsaux2!src_requested_quantity) + " WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and DELIVERY_DETAIL_ID = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 var_dia = CStr(Day(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_mes = CStr(Month(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_año = CStr(Year(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 If Len(Trim(var_dia)) = 1 Then
                                    var_dia = "0" + var_dia
                                 End If
                                 If Len(Trim(var_mes)) = 1 Then
                                    var_mes = "0" + var_mes
                                 End If
                                 var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                 
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                 var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,LOTE, GRUPO, PASILLO) "
                                 var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux2!source_header_number + "', " + CStr(rsaux2!delivery_id) + "," + CStr(rsaux2!delivery_detail_id) + ", " + CStr(rsaux2!organization_id) + ",'" + IIf(IsNull(rsaux2!subinventory), "", rsaux2!subinventory) + "', " + CStr(rsaux2!delivery_line_id) + "," + CStr(rsaux2!inventory_item_id) + ",'" + rsaux2!item_description + "','" + CStr(rsaux2!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux2!src_requested_quantity) + ",'" + rsaux2!released_status + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!customer_name), "", rsaux2!customer_name) + "','" + rsaux2!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux2!collector_id), 0, rsaux2!collector_id)) + ",'" + IIf(IsNull(rsaux2!Name), "", rsaux2!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux2!ubicacion), "", rsaux2!ubicacion) + "','" + IIf(IsNull(rsaux2!establecimiento), "", rsaux2!establecimiento) + "','" + IIf(IsNull(rsaux2!nombre_Establecimiento), "", rsaux2!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux2!ORDENES), "", rsaux2!ORDENES) + "',"
                                 var_cadena = var_cadena + CStr(IIf(IsNull(rsaux2!CUST_ACCOUNT_ID), 0, rsaux2!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux2!source_header_type_name), "", rsaux2!source_header_type_name) + "', '" + IIf(IsNull(rsaux2!source_document_id), "", rsaux2!source_document_id) + "','" + IIf(IsNull(rsaux2!pais), "", rsaux2!pais) + "','" + IIf(IsNull(rsaux2!estado), "", rsaux2!estado) + "', '" + IIf(IsNull(rsaux2!municipio), "", rsaux2!municipio) + "', '" + IIf(IsNull(rsaux2!ciudad), "", rsaux2!ciudad) + "', '" + IIf(IsNull(rsaux2!colonia), "", rsaux2!colonia) + "','" + IIf(IsNull(rsaux2!DIRECCION), "", rsaux2!DIRECCION) + "', '" + IIf(IsNull(rsaux2!cp), "", rsaux2!cp) + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!site_use_id), "", rsaux2!site_use_id) + "','" + IIf(IsNull(rsaux2!paqueteria), "", rsaux2!paqueteria) + "','" + IIf(IsNull(rsaux2!attribute8), "", rsaux2!attribute8) + "','" + IIf(IsNull(rsaux2!ATTRIBUTE9), "", rsaux2!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux2!TELEFONO), "", rsaux2!TELEFONO) + "','" + IIf(IsNull(rsaux2!Linea), "", rsaux2!Linea) + "','" + CStr(IIf(IsNull(rsaux2!ruta), "", rsaux2!ruta)) + "','" + IIf(IsNull(rsaux2!nombre_ruta), "", rsaux2!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_SURTIDO), 0, rsaux2!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", '" + IIf(IsNull(rsaux2!estacion), "", rsaux2!estacion) + "'," + CStr(var_lote) + "," + CStr(var_grupo) + ",'P01')"
                                 rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux3.Close
                              If rsaux2!Linea <> "CATALOGOS" Then
                                 If rsaux2!Linea <> "CATALOGO" Then
                                    If rsaux2!Linea <> "POP" Then
                                       If rsaux2!Linea <> "EMPAQUE" Then
                                          If rsaux2!MULTIPLO > 1 Then
                                             var_contador = var_contador + rsaux2!src_requested_quantity
                                          Else
                                             var_contador = var_contador + 1
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                              rsaux2.MoveNext
                        Wend
                        rsaux2.Close
                        var_veces_grupo = var_veces_grupo + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  
                  var_consecutivo_ubicacion = 1
                  rsaux1.Open "select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_Consecutivo = " + CStr(var_consecutivo) + " and SUBSTRING(UBICACION,1,3) in ('B01','P01') ORDER BY UBICACION", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update tb_temp_oracle_orden_surtido_aux_1 set consecutivo_pasillo = " + CStr(var_consecutivo_ubicacion) + "  where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ubicacion = '" + rsaux1!ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_consecutivo_ubicacion = var_consecutivo_ubicacion + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  
                  'var_consecutivo_ubicacion = 1
                  rsaux1.Open "select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_Consecutivo = " + CStr(var_consecutivo) + " and SUBSTRING(UBICACION,1,3) in ('P02') ORDER BY UBICACION DESC", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update tb_temp_oracle_orden_surtido_aux_1 set consecutivo_pasillo = " + CStr(var_consecutivo_ubicacion) + "  where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ubicacion = '" + rsaux1!ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_consecutivo_ubicacion = var_consecutivo_ubicacion + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  
                  var_grupo = var_grupo + 1




'PASILLO 3,4
                  var_grupo = var_grupo + 1
                  var_lote = var_lote + 1
                  var_consecutivo = var_consecutivo_general
                  rsaux1.Open "select distinct source_header_number, ORDEN_pedido from tb_Temp_oracle_orden_surtido_aux_2, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SOURCE_HEADER_NUMBER = pedido AND SUBSTRING(UBICACION,1,3) IN ('P03','P04')  order by ORDEN_pedido", cnn, adOpenDynamic, adLockOptimistic
                  var_contador = 0
                  var_veces_grupo = 0
                  While Not rsaux1.EOF
                        If var_veces_grupo = 1 Then
                           var_contador = 60
                           var_veces_grupo = 0
                        End If
                        If var_contador >= 60 Then
                           var_grupo = var_grupo + 1
                           var_contador = 0
                        End If
                        
                        rsaux2.Open "select * from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number = " + CStr(rsaux1!source_header_number) + " AND SUBSTRING(UBICACION,1,3) IN ('P03','P04')   order by ubicacion", cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux2.EOF
                              If var_contador >= 60 Then
                                 var_lote = var_lote + 1
                                 var_contador = 0
                                 var_grupo = var_grupo + 1
                              End If
                              rsaux3.Open "SELECT * FROM tb_Temp_oracle_orden_surtido_aux_1 WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and delivery_detail_id = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 rsaux4.Open "UPDATE TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 SET SRC_REQUESTED_QUANTITY = SRC_REQUESTED_QUANTITY + " + CStr(rsaux2!src_requested_quantity) + " WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and DELIVERY_DETAIL_ID = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 var_dia = CStr(Day(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_mes = CStr(Month(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_año = CStr(Year(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 If Len(Trim(var_dia)) = 1 Then
                                    var_dia = "0" + var_dia
                                 End If
                                 If Len(Trim(var_mes)) = 1 Then
                                    var_mes = "0" + var_mes
                                 End If
                                 var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                 
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                 var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,LOTE, GRUPO, PASILLO) "
                                 var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux2!source_header_number + "', " + CStr(rsaux2!delivery_id) + "," + CStr(rsaux2!delivery_detail_id) + ", " + CStr(rsaux2!organization_id) + ",'" + IIf(IsNull(rsaux2!subinventory), "", rsaux2!subinventory) + "', " + CStr(rsaux2!delivery_line_id) + "," + CStr(rsaux2!inventory_item_id) + ",'" + rsaux2!item_description + "','" + CStr(rsaux2!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux2!src_requested_quantity) + ",'" + rsaux2!released_status + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!customer_name), "", rsaux2!customer_name) + "','" + rsaux2!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux2!collector_id), 0, rsaux2!collector_id)) + ",'" + IIf(IsNull(rsaux2!Name), "", rsaux2!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux2!ubicacion), "", rsaux2!ubicacion) + "','" + IIf(IsNull(rsaux2!establecimiento), "", rsaux2!establecimiento) + "','" + IIf(IsNull(rsaux2!nombre_Establecimiento), "", rsaux2!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux2!ORDENES), "", rsaux2!ORDENES) + "',"
                                 var_cadena = var_cadena + CStr(IIf(IsNull(rsaux2!CUST_ACCOUNT_ID), 0, rsaux2!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux2!source_header_type_name), "", rsaux2!source_header_type_name) + "', '" + IIf(IsNull(rsaux2!source_document_id), "", rsaux2!source_document_id) + "','" + IIf(IsNull(rsaux2!pais), "", rsaux2!pais) + "','" + IIf(IsNull(rsaux2!estado), "", rsaux2!estado) + "', '" + IIf(IsNull(rsaux2!municipio), "", rsaux2!municipio) + "', '" + IIf(IsNull(rsaux2!ciudad), "", rsaux2!ciudad) + "', '" + IIf(IsNull(rsaux2!colonia), "", rsaux2!colonia) + "','" + IIf(IsNull(rsaux2!DIRECCION), "", rsaux2!DIRECCION) + "', '" + IIf(IsNull(rsaux2!cp), "", rsaux2!cp) + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!site_use_id), "", rsaux2!site_use_id) + "','" + IIf(IsNull(rsaux2!paqueteria), "", rsaux2!paqueteria) + "','" + IIf(IsNull(rsaux2!attribute8), "", rsaux2!attribute8) + "','" + IIf(IsNull(rsaux2!ATTRIBUTE9), "", rsaux2!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux2!TELEFONO), "", rsaux2!TELEFONO) + "','" + IIf(IsNull(rsaux2!Linea), "", rsaux2!Linea) + "','" + CStr(IIf(IsNull(rsaux2!ruta), "", rsaux2!ruta)) + "','" + IIf(IsNull(rsaux2!nombre_ruta), "", rsaux2!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_SURTIDO), 0, rsaux2!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", '" + IIf(IsNull(rsaux2!estacion), "", rsaux2!estacion) + "'," + CStr(var_lote) + "," + CStr(var_grupo) + ",'P01')"
                                 rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux3.Close
                              If rsaux2!Linea <> "CATALOGOS" Then
                                 If rsaux2!Linea <> "CATALOGO" Then
                                    If rsaux2!Linea <> "POP" Then
                                       If rsaux2!Linea <> "EMPAQUE" Then
                                          If rsaux2!MULTIPLO > 1 Then
                                             var_contador = var_contador + rsaux2!src_requested_quantity
                                          Else
                                             var_contador = var_contador + 1
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                              rsaux2.MoveNext
                        Wend
                        rsaux2.Close
                        var_veces_grupo = var_veces_grupo + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  
                  var_consecutivo_ubicacion = 1
                  rsaux1.Open "select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_Consecutivo = " + CStr(var_consecutivo) + " and SUBSTRING(UBICACION,1,3) in ('B01','P01') ORDER BY UBICACION", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update tb_temp_oracle_orden_surtido_aux_1 set consecutivo_pasillo = " + CStr(var_consecutivo_ubicacion) + "  where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ubicacion = '" + rsaux1!ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_consecutivo_ubicacion = var_consecutivo_ubicacion + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  
                  'var_consecutivo_ubicacion = 1
                  rsaux1.Open "select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_Consecutivo = " + CStr(var_consecutivo) + " and SUBSTRING(UBICACION,1,3) in ('P02') ORDER BY UBICACION DESC", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update tb_temp_oracle_orden_surtido_aux_1 set consecutivo_pasillo = " + CStr(var_consecutivo_ubicacion) + "  where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ubicacion = '" + rsaux1!ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_consecutivo_ubicacion = var_consecutivo_ubicacion + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  
                  var_grupo = var_grupo + 1
                  
'PASILLO 5, 6, 7
                  var_lote = var_lote + 1
                  rsaux1.Open "select distinct source_header_number, ORDEN_pedido from tb_Temp_oracle_orden_surtido_aux_2, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SOURCE_HEADER_NUMBER = pedido AND SUBSTRING(UBICACION,1,3) IN ('P05','P06','P07')  order by ORDEN_pedido", cnn, adOpenDynamic, adLockOptimistic
                  var_contador = 0
                  var_veces_grupo = 0
                  While Not rsaux1.EOF
                        If var_veces_grupo = 1 Then
                           var_contador = 30
                           var_veces_grupo = 0
                        End If
                        If var_contador >= 30 Then
                           var_grupo = var_grupo + 1
                           var_contador = 0
                        End If
                        rsaux2.Open "select * from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number = " + CStr(rsaux1!source_header_number) + " AND SUBSTRING(UBICACION,1,3) IN ('P05','P06','P07')  order by ubicacion", cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux2.EOF
                              If var_contador >= 30 Then
                                 var_lote = var_lote + 1
                                 var_contador = 0
                                 var_grupo = var_grupo + 1
                              End If
                              rsaux3.Open "SELECT * FROM tb_Temp_oracle_orden_surtido_aux_1 WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and delivery_detail_id = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 rsaux4.Open "UPDATE TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 SET SRC_REQUESTED_QUANTITY = SRC_REQUESTED_QUANTITY + " + CStr(rsaux2!src_requested_quantity) + " WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and DELIVERY_DETAIL_ID = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 var_dia = CStr(Day(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_mes = CStr(Month(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_año = CStr(Year(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 If Len(Trim(var_dia)) = 1 Then
                                    var_dia = "0" + var_dia
                                 End If
                                 If Len(Trim(var_mes)) = 1 Then
                                    var_mes = "0" + var_mes
                                 End If
                                 var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                 
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                 var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,LOTE, GRUPO, PASILLO) "
                                 var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux2!source_header_number + "', " + CStr(rsaux2!delivery_id) + "," + CStr(rsaux2!delivery_detail_id) + ", " + CStr(rsaux2!organization_id) + ",'" + IIf(IsNull(rsaux2!subinventory), "", rsaux2!subinventory) + "', " + CStr(rsaux2!delivery_line_id) + "," + CStr(rsaux2!inventory_item_id) + ",'" + rsaux2!item_description + "','" + CStr(rsaux2!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux2!src_requested_quantity) + ",'" + rsaux2!released_status + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!customer_name), "", rsaux2!customer_name) + "','" + rsaux2!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux2!collector_id), 0, rsaux2!collector_id)) + ",'" + IIf(IsNull(rsaux2!Name), "", rsaux2!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux2!ubicacion), "", rsaux2!ubicacion) + "','" + IIf(IsNull(rsaux2!establecimiento), "", rsaux2!establecimiento) + "','" + IIf(IsNull(rsaux2!nombre_Establecimiento), "", rsaux2!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux2!ORDENES), "", rsaux2!ORDENES) + "',"
                                 var_cadena = var_cadena + CStr(IIf(IsNull(rsaux2!CUST_ACCOUNT_ID), 0, rsaux2!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux2!source_header_type_name), "", rsaux2!source_header_type_name) + "', '" + IIf(IsNull(rsaux2!source_document_id), "", rsaux2!source_document_id) + "','" + IIf(IsNull(rsaux2!pais), "", rsaux2!pais) + "','" + IIf(IsNull(rsaux2!estado), "", rsaux2!estado) + "', '" + IIf(IsNull(rsaux2!municipio), "", rsaux2!municipio) + "', '" + IIf(IsNull(rsaux2!ciudad), "", rsaux2!ciudad) + "', '" + IIf(IsNull(rsaux2!colonia), "", rsaux2!colonia) + "','" + IIf(IsNull(rsaux2!DIRECCION), "", rsaux2!DIRECCION) + "', '" + IIf(IsNull(rsaux2!cp), "", rsaux2!cp) + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!site_use_id), "", rsaux2!site_use_id) + "','" + IIf(IsNull(rsaux2!paqueteria), "", rsaux2!paqueteria) + "','" + IIf(IsNull(rsaux2!attribute8), "", rsaux2!attribute8) + "','" + IIf(IsNull(rsaux2!ATTRIBUTE9), "", rsaux2!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux2!TELEFONO), "", rsaux2!TELEFONO) + "','" + IIf(IsNull(rsaux2!Linea), "", rsaux2!Linea) + "','" + CStr(IIf(IsNull(rsaux2!ruta), "", rsaux2!ruta)) + "','" + IIf(IsNull(rsaux2!nombre_ruta), "", rsaux2!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_SURTIDO), 0, rsaux2!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", '" + IIf(IsNull(rsaux2!estacion), "", rsaux2!estacion) + "'," + CStr(var_lote) + "," + CStr(var_grupo) + ",'P03')"
                                 rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux3.Close
                              If rsaux2!Linea <> "CATALOGOS" Then
                                 If rsaux2!Linea <> "CATALOGO" Then
                                    If rsaux2!Linea <> "POP" Then
                                       If rsaux2!Linea <> "EMPAQUE" Then
                                          If rsaux2!MULTIPLO > 1 Then
                                             var_contador = var_contador + rsaux2!src_requested_quantity
                                          Else
                                             var_contador = var_contador + 1
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                              rsaux2.MoveNext
                        Wend
                        rsaux2.Close
                        var_veces_grupo = var_veces_grupo + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  
                  var_consecutivo_ubicacion = 1
                  rsaux1.Open "select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_Consecutivo = " + CStr(var_consecutivo) + " and SUBSTRING(UBICACION,1,3) in ('P03') ORDER BY UBICACION DESC", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update tb_temp_oracle_orden_surtido_aux_1 set consecutivo_pasillo = " + CStr(var_consecutivo_ubicacion) + "  where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ubicacion = '" + rsaux1!ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_consecutivo_ubicacion = var_consecutivo_ubicacion + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  
                  
                  
'PASILLO 8,9
                  
                  
                  var_grupo = var_grupo + 1
                  var_lote = var_lote + 1
                  rsaux1.Open "select distinct source_header_number, ORDEN_pedido from tb_Temp_oracle_orden_surtido_aux_2, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SOURCE_HEADER_NUMBER = pedido AND SUBSTRING(UBICACION,1,3) IN ('P08','P09')  order by ORDEN_pedido", cnn, adOpenDynamic, adLockOptimistic
                  var_contador = 0
                  var_veces_grupo = 0
                  While Not rsaux1.EOF
                        If var_veces_grupo = 1 Then
                        ' se cambio de 30 a 60
                           var_contador = 60
                           var_veces_grupo = 0
                        End If
                        If var_contador >= 30 Then
                           var_grupo = var_grupo + 1
                           var_contador = 0
                        End If
                        rsaux2.Open "select * from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number = " + CStr(rsaux1!source_header_number) + " AND SUBSTRING(UBICACION,1,3) IN ('P08','P09')   order by ubicacion", cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux2.EOF
                              'se cambio de 30 a 60
                              If var_contador >= 60 Then
                                 var_lote = var_lote + 1
                                 var_contador = 0
                                 var_grupo = var_grupo + 1
                              End If
                              rsaux3.Open "SELECT * FROM tb_Temp_oracle_orden_surtido_aux_1 WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and delivery_detail_id = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 rsaux4.Open "UPDATE TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 SET SRC_REQUESTED_QUANTITY = SRC_REQUESTED_QUANTITY + " + CStr(rsaux2!src_requested_quantity) + " WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and DELIVERY_DETAIL_ID = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 var_dia = CStr(Day(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_mes = CStr(Month(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_año = CStr(Year(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 If Len(Trim(var_dia)) = 1 Then
                                    var_dia = "0" + var_dia
                                 End If
                                 If Len(Trim(var_mes)) = 1 Then
                                    var_mes = "0" + var_mes
                                 End If
                                 var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                 
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                 var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,LOTE, GRUPO, PASILLO) "
                                 var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux2!source_header_number + "', " + CStr(rsaux2!delivery_id) + "," + CStr(rsaux2!delivery_detail_id) + ", " + CStr(rsaux2!organization_id) + ",'" + IIf(IsNull(rsaux2!subinventory), "", rsaux2!subinventory) + "', " + CStr(rsaux2!delivery_line_id) + "," + CStr(rsaux2!inventory_item_id) + ",'" + rsaux2!item_description + "','" + CStr(rsaux2!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux2!src_requested_quantity) + ",'" + rsaux2!released_status + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!customer_name), "", rsaux2!customer_name) + "','" + rsaux2!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux2!collector_id), 0, rsaux2!collector_id)) + ",'" + IIf(IsNull(rsaux2!Name), "", rsaux2!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux2!ubicacion), "", rsaux2!ubicacion) + "','" + IIf(IsNull(rsaux2!establecimiento), "", rsaux2!establecimiento) + "','" + IIf(IsNull(rsaux2!nombre_Establecimiento), "", rsaux2!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux2!ORDENES), "", rsaux2!ORDENES) + "',"
                                 var_cadena = var_cadena + CStr(IIf(IsNull(rsaux2!CUST_ACCOUNT_ID), 0, rsaux2!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux2!source_header_type_name), "", rsaux2!source_header_type_name) + "', '" + IIf(IsNull(rsaux2!source_document_id), "", rsaux2!source_document_id) + "','" + IIf(IsNull(rsaux2!pais), "", rsaux2!pais) + "','" + IIf(IsNull(rsaux2!estado), "", rsaux2!estado) + "', '" + IIf(IsNull(rsaux2!municipio), "", rsaux2!municipio) + "', '" + IIf(IsNull(rsaux2!ciudad), "", rsaux2!ciudad) + "', '" + IIf(IsNull(rsaux2!colonia), "", rsaux2!colonia) + "','" + IIf(IsNull(rsaux2!DIRECCION), "", rsaux2!DIRECCION) + "', '" + IIf(IsNull(rsaux2!cp), "", rsaux2!cp) + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!site_use_id), "", rsaux2!site_use_id) + "','" + IIf(IsNull(rsaux2!paqueteria), "", rsaux2!paqueteria) + "','" + IIf(IsNull(rsaux2!attribute8), "", rsaux2!attribute8) + "','" + IIf(IsNull(rsaux2!ATTRIBUTE9), "", rsaux2!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux2!TELEFONO), "", rsaux2!TELEFONO) + "','" + IIf(IsNull(rsaux2!Linea), "", rsaux2!Linea) + "','" + CStr(IIf(IsNull(rsaux2!ruta), "", rsaux2!ruta)) + "','" + IIf(IsNull(rsaux2!nombre_ruta), "", rsaux2!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_SURTIDO), 0, rsaux2!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", '" + IIf(IsNull(rsaux2!estacion), "", rsaux2!estacion) + "'," + CStr(var_lote) + "," + CStr(var_grupo) + ",'P03')"
                                 rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux3.Close
                              If rsaux2!Linea <> "CATALOGOS" Then
                                 If rsaux2!Linea <> "CATALOGO" Then
                                    If rsaux2!Linea <> "POP" Then
                                       If rsaux2!Linea <> "EMPAQUE" Then
                                          If rsaux2!MULTIPLO > 1 Then
                                             var_contador = var_contador + rsaux2!src_requested_quantity
                                          Else
                                             var_contador = var_contador + 1
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                              rsaux2.MoveNext
                        Wend
                        rsaux2.Close
                        var_veces_grupo = var_veces_grupo + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close

'pasillo 10 y 11

                  var_grupo = var_grupo + 1
                  var_lote = var_lote + 1
                  rsaux1.Open "select distinct source_header_number, ORDEN_pedido from tb_Temp_oracle_orden_surtido_aux_2, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SOURCE_HEADER_NUMBER = pedido AND SUBSTRING(UBICACION,1,3) IN ('P10','P11')  order by ORDEN_pedido", cnn, adOpenDynamic, adLockOptimistic
                  var_contador = 0
                  var_veces_grupo = 0
                  While Not rsaux1.EOF
                        If var_veces_grupo = 1 Then
                        ' se cambio de 30 a 60
                           var_contador = 30
                           var_veces_grupo = 0
                        End If
                        If var_contador >= 30 Then
                           var_grupo = var_grupo + 1
                           var_contador = 0
                        End If
                        rsaux2.Open "select * from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number = " + CStr(rsaux1!source_header_number) + " AND SUBSTRING(UBICACION,1,3) IN ('P10','P11')   order by ubicacion", cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux2.EOF
                              'se cambio de 30 a 60
                              If var_contador >= 30 Then
                                 var_lote = var_lote + 1
                                 var_contador = 0
                                 var_grupo = var_grupo + 1
                              End If
                              rsaux3.Open "SELECT * FROM tb_Temp_oracle_orden_surtido_aux_1 WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and delivery_detail_id = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 rsaux4.Open "UPDATE TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 SET SRC_REQUESTED_QUANTITY = SRC_REQUESTED_QUANTITY + " + CStr(rsaux2!src_requested_quantity) + " WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and DELIVERY_DETAIL_ID = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 var_dia = CStr(Day(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_mes = CStr(Month(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_año = CStr(Year(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 If Len(Trim(var_dia)) = 1 Then
                                    var_dia = "0" + var_dia
                                 End If
                                 If Len(Trim(var_mes)) = 1 Then
                                    var_mes = "0" + var_mes
                                 End If
                                 var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                 
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                 var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,LOTE, GRUPO, PASILLO) "
                                 var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux2!source_header_number + "', " + CStr(rsaux2!delivery_id) + "," + CStr(rsaux2!delivery_detail_id) + ", " + CStr(rsaux2!organization_id) + ",'" + IIf(IsNull(rsaux2!subinventory), "", rsaux2!subinventory) + "', " + CStr(rsaux2!delivery_line_id) + "," + CStr(rsaux2!inventory_item_id) + ",'" + rsaux2!item_description + "','" + CStr(rsaux2!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux2!src_requested_quantity) + ",'" + rsaux2!released_status + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!customer_name), "", rsaux2!customer_name) + "','" + rsaux2!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux2!collector_id), 0, rsaux2!collector_id)) + ",'" + IIf(IsNull(rsaux2!Name), "", rsaux2!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux2!ubicacion), "", rsaux2!ubicacion) + "','" + IIf(IsNull(rsaux2!establecimiento), "", rsaux2!establecimiento) + "','" + IIf(IsNull(rsaux2!nombre_Establecimiento), "", rsaux2!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux2!ORDENES), "", rsaux2!ORDENES) + "',"
                                 var_cadena = var_cadena + CStr(IIf(IsNull(rsaux2!CUST_ACCOUNT_ID), 0, rsaux2!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux2!source_header_type_name), "", rsaux2!source_header_type_name) + "', '" + IIf(IsNull(rsaux2!source_document_id), "", rsaux2!source_document_id) + "','" + IIf(IsNull(rsaux2!pais), "", rsaux2!pais) + "','" + IIf(IsNull(rsaux2!estado), "", rsaux2!estado) + "', '" + IIf(IsNull(rsaux2!municipio), "", rsaux2!municipio) + "', '" + IIf(IsNull(rsaux2!ciudad), "", rsaux2!ciudad) + "', '" + IIf(IsNull(rsaux2!colonia), "", rsaux2!colonia) + "','" + IIf(IsNull(rsaux2!DIRECCION), "", rsaux2!DIRECCION) + "', '" + IIf(IsNull(rsaux2!cp), "", rsaux2!cp) + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!site_use_id), "", rsaux2!site_use_id) + "','" + IIf(IsNull(rsaux2!paqueteria), "", rsaux2!paqueteria) + "','" + IIf(IsNull(rsaux2!attribute8), "", rsaux2!attribute8) + "','" + IIf(IsNull(rsaux2!ATTRIBUTE9), "", rsaux2!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux2!TELEFONO), "", rsaux2!TELEFONO) + "','" + IIf(IsNull(rsaux2!Linea), "", rsaux2!Linea) + "','" + CStr(IIf(IsNull(rsaux2!ruta), "", rsaux2!ruta)) + "','" + IIf(IsNull(rsaux2!nombre_ruta), "", rsaux2!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_SURTIDO), 0, rsaux2!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", '" + IIf(IsNull(rsaux2!estacion), "", rsaux2!estacion) + "'," + CStr(var_lote) + "," + CStr(var_grupo) + ",'P03')"
                                 rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux3.Close
                              If rsaux2!Linea <> "CATALOGOS" Then
                                 If rsaux2!Linea <> "CATALOGO" Then
                                    If rsaux2!Linea <> "POP" Then
                                       If rsaux2!Linea <> "EMPAQUE" Then
                                          If rsaux2!MULTIPLO > 1 Then
                                             var_contador = var_contador + rsaux2!src_requested_quantity
                                          Else
                                             var_contador = var_contador + 1
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                              rsaux2.MoveNext
                        Wend
                        rsaux2.Close
                        var_veces_grupo = var_veces_grupo + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close


                  var_consecutivo_ubicacion = 1
                  rsaux1.Open "select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_Consecutivo = " + CStr(var_consecutivo) + " and SUBSTRING(UBICACION,1,3) in ('P04') ORDER BY UBICACION DESC", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update tb_temp_oracle_orden_surtido_aux_1 set consecutivo_pasillo = " + CStr(var_consecutivo_ubicacion) + "  where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ubicacion = '" + rsaux1!ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_consecutivo_ubicacion = var_consecutivo_ubicacion + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  

                  var_consecutivo = var_consecutivo_general
                  rsaux1.Open "select distinct source_header_number, ORDEN_pedido from tb_Temp_oracle_orden_surtido_aux_2, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SOURCE_HEADER_NUMBER = pedido AND SUBSTRING(UBICACION,1,3) not IN ('B00', 'B01', 'P01', 'P02', 'P03', 'P04', 'P05', 'P06', 'P07', 'P08', 'P09','P00','P10','P11')  order by ORDEN_pedido", cnn, adOpenDynamic, adLockOptimistic
                  var_contador = 0
                  var_veces_grupo = 0
                  var_lote = 1
                  While Not rsaux1.EOF
                        'var_lote = 1
                        
                        If var_veces_grupo = 1 Then
                           var_contador = 60
                           'var_grupo = var_grupo + 1
                           var_veces_grupo = 0
                        End If
                        If var_contador >= 60 Then
                           var_grupo = var_grupo + 1
                           'var_veces_grupo = var_veces_grupo + 1
                           var_contador = 0
                        End If
                        
                        rsaux2.Open "select * from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number = " + CStr(rsaux1!source_header_number) + " AND SUBSTRING(UBICACION,1,3) NOT IN ('B00', 'B01', 'P01', 'P02', 'P03', 'P04', 'P05', 'P06', 'P07', 'P08', 'P09','P00', 'P10','P11')   order by ubicacion", cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux2.EOF
                              'If var_veces_grupo = 2 Then
                              '   var_veces_grupo = 1
                              '   var_grupo = var_grupo + 1
                              'End If
                              If var_contador >= 60 Then
                                 var_lote = var_lote + 1
                                 var_contador = 0
                                 var_grupo = var_grupo + 1
                                 'var_veces_grupo = var_veces_grupo + 1
                              End If
                              rsaux3.Open "SELECT * FROM tb_Temp_oracle_orden_surtido_aux_1 WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and delivery_detail_id = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 rsaux4.Open "UPDATE TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 SET SRC_REQUESTED_QUANTITY = SRC_REQUESTED_QUANTITY + " + CStr(rsaux2!src_requested_quantity) + " WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and DELIVERY_DETAIL_ID = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 var_dia = CStr(Day(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_mes = CStr(Month(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_año = CStr(Year(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 If Len(Trim(var_dia)) = 1 Then
                                    var_dia = "0" + var_dia
                                 End If
                                 If Len(Trim(var_mes)) = 1 Then
                                    var_mes = "0" + var_mes
                                 End If
                                 var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                 
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                 var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,LOTE, GRUPO, PASILLO) "
                                 var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux2!source_header_number + "', " + CStr(rsaux2!delivery_id) + "," + CStr(rsaux2!delivery_detail_id) + ", " + CStr(rsaux2!organization_id) + ",'" + IIf(IsNull(rsaux2!subinventory), "", rsaux2!subinventory) + "', " + CStr(rsaux2!delivery_line_id) + "," + CStr(rsaux2!inventory_item_id) + ",'" + rsaux2!item_description + "','" + CStr(rsaux2!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux2!src_requested_quantity) + ",'" + rsaux2!released_status + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!customer_name), "", rsaux2!customer_name) + "','" + rsaux2!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux2!collector_id), 0, rsaux2!collector_id)) + ",'" + IIf(IsNull(rsaux2!Name), "", rsaux2!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux2!ubicacion), "", rsaux2!ubicacion) + "','" + IIf(IsNull(rsaux2!establecimiento), "", rsaux2!establecimiento) + "','" + IIf(IsNull(rsaux2!nombre_Establecimiento), "", rsaux2!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux2!ORDENES), "", rsaux2!ORDENES) + "',"
                                 var_cadena = var_cadena + CStr(IIf(IsNull(rsaux2!CUST_ACCOUNT_ID), 0, rsaux2!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux2!source_header_type_name), "", rsaux2!source_header_type_name) + "', '" + IIf(IsNull(rsaux2!source_document_id), "", rsaux2!source_document_id) + "','" + IIf(IsNull(rsaux2!pais), "", rsaux2!pais) + "','" + IIf(IsNull(rsaux2!estado), "", rsaux2!estado) + "', '" + IIf(IsNull(rsaux2!municipio), "", rsaux2!municipio) + "', '" + IIf(IsNull(rsaux2!ciudad), "", rsaux2!ciudad) + "', '" + IIf(IsNull(rsaux2!colonia), "", rsaux2!colonia) + "','" + IIf(IsNull(rsaux2!DIRECCION), "", rsaux2!DIRECCION) + "', '" + IIf(IsNull(rsaux2!cp), "", rsaux2!cp) + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!site_use_id), "", rsaux2!site_use_id) + "','" + IIf(IsNull(rsaux2!paqueteria), "", rsaux2!paqueteria) + "','" + IIf(IsNull(rsaux2!attribute8), "", rsaux2!attribute8) + "','" + IIf(IsNull(rsaux2!ATTRIBUTE9), "", rsaux2!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux2!TELEFONO), "", rsaux2!TELEFONO) + "','" + IIf(IsNull(rsaux2!Linea), "", rsaux2!Linea) + "','" + CStr(IIf(IsNull(rsaux2!ruta), "", rsaux2!ruta)) + "','" + IIf(IsNull(rsaux2!nombre_ruta), "", rsaux2!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_SURTIDO), 0, rsaux2!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", '" + IIf(IsNull(rsaux2!estacion), "", rsaux2!estacion) + "'," + CStr(var_lote) + "," + CStr(var_grupo) + ",'B00')"
                                 rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux3.Close
                              If rsaux2!Linea <> "CATALOGOS" Then
                                 If rsaux2!Linea <> "CATALOGO" Then
                                    If rsaux2!Linea <> "POP" Then
                                       If rsaux2!Linea <> "EMPAQUE" Then
                                          If rsaux2!MULTIPLO > 1 Then
                                             var_contador = var_contador + rsaux2!src_requested_quantity
                                          Else
                                             var_contador = var_contador + 1
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                              rsaux2.MoveNext
                        Wend
                        rsaux2.Close
                        var_veces_grupo = var_veces_grupo + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  
                  var_consecutivo_ubicacion = 1
                  rsaux1.Open "select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_Consecutivo = " + CStr(var_consecutivo) + " and SUBSTRING(UBICACION,1,3) NOT in ('B00', 'B01', 'P01', 'P02', 'P03', 'P04', 'P05', 'P06', 'P07','P08', 'P09', 'P10', 'P11') ORDER BY UBICACION", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update tb_temp_oracle_orden_surtido_aux_1 set consecutivo_pasillo = " + CStr(var_consecutivo_ubicacion) + "  where  source_header_number = '" + rsaux15!pedido + "' and  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ubicacion = '" + rsaux1!ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_consecutivo_ubicacion = var_consecutivo_ubicacion + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close

               rsaux15.MoveNext
               Wend
               rsaux15.Close

End Sub

Private Sub PASILLOS_3()

End Sub


Private Sub PASILLOS()
                  var_grupo = 1
                  
'pasillo 0
                  var_consecutivo = var_consecutivo_general
                  rsaux1.Open "select distinct source_header_number, ORDEN_PEDIDO from tb_Temp_oracle_orden_surtido_aux_2, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SOURCE_HEADER_NUMBER = pedido AND SUBSTRING(UBICACION,1,3) IN ('P00','B00')  order by ORDEN_pedido", cnn, adOpenDynamic, adLockOptimistic
                  var_contador = 0
                  var_veces_grupo = 0
                  var_lote = 1
                  While Not rsaux1.EOF
                        'var_lote = 1
                        
                        If var_veces_grupo = 2 Then
                           var_contador = 60
                           'var_grupo = var_grupo + 1
                           var_veces_grupo = 0
                        End If
                        If var_contador >= 60 Then
                           var_grupo = var_grupo + 1
                           'var_veces_grupo = var_veces_grupo + 1
                           var_contador = 0
                        End If
                        
                        rsaux2.Open "select * from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number = " + CStr(rsaux1!source_header_number) + " AND SUBSTRING(UBICACION,1,3) IN ('B00','P00')   order by ubicacion", cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux2.EOF
                              'If var_veces_grupo = 2 Then
                              '   var_veces_grupo = 1
                              '   var_grupo = var_grupo + 1
                              'End If
                              If var_contador >= 60 Then
                                 var_lote = var_lote + 1
                                 var_contador = 0
                                 var_grupo = var_grupo + 1
                                 'var_veces_grupo = var_veces_grupo + 1
                              End If
                              rsaux3.Open "SELECT * FROM tb_Temp_oracle_orden_surtido_aux_1 WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and delivery_detail_id = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 rsaux4.Open "UPDATE TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 SET SRC_REQUESTED_QUANTITY = SRC_REQUESTED_QUANTITY + " + CStr(rsaux2!src_requested_quantity) + " WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and DELIVERY_DETAIL_ID = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 var_dia = CStr(Day(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_mes = CStr(Month(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_año = CStr(Year(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 If Len(Trim(var_dia)) = 1 Then
                                    var_dia = "0" + var_dia
                                 End If
                                 If Len(Trim(var_mes)) = 1 Then
                                    var_mes = "0" + var_mes
                                 End If
                                 var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                 
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                 var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,LOTE, GRUPO, PASILLO) "
                                 var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux2!source_header_number + "', " + CStr(rsaux2!delivery_id) + "," + CStr(rsaux2!delivery_detail_id) + ", " + CStr(rsaux2!organization_id) + ",'" + IIf(IsNull(rsaux2!subinventory), "", rsaux2!subinventory) + "', " + CStr(rsaux2!delivery_line_id) + "," + CStr(rsaux2!inventory_item_id) + ",'" + rsaux2!item_description + "','" + CStr(rsaux2!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux2!src_requested_quantity) + ",'" + rsaux2!released_status + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!customer_name), "", rsaux2!customer_name) + "','" + rsaux2!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux2!collector_id), 0, rsaux2!collector_id)) + ",'" + IIf(IsNull(rsaux2!Name), "", rsaux2!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux2!ubicacion), "", rsaux2!ubicacion) + "','" + IIf(IsNull(rsaux2!establecimiento), "", rsaux2!establecimiento) + "','" + IIf(IsNull(rsaux2!nombre_Establecimiento), "", rsaux2!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux2!ORDENES), "", rsaux2!ORDENES) + "',"
                                 var_cadena = var_cadena + CStr(IIf(IsNull(rsaux2!CUST_ACCOUNT_ID), 0, rsaux2!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux2!source_header_type_name), "", rsaux2!source_header_type_name) + "', '" + IIf(IsNull(rsaux2!source_document_id), "", rsaux2!source_document_id) + "','" + IIf(IsNull(rsaux2!pais), "", rsaux2!pais) + "','" + IIf(IsNull(rsaux2!estado), "", rsaux2!estado) + "', '" + IIf(IsNull(rsaux2!municipio), "", rsaux2!municipio) + "', '" + IIf(IsNull(rsaux2!ciudad), "", rsaux2!ciudad) + "', '" + IIf(IsNull(rsaux2!colonia), "", rsaux2!colonia) + "','" + IIf(IsNull(rsaux2!DIRECCION), "", rsaux2!DIRECCION) + "', '" + IIf(IsNull(rsaux2!cp), "", rsaux2!cp) + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!site_use_id), "", rsaux2!site_use_id) + "','" + IIf(IsNull(rsaux2!paqueteria), "", rsaux2!paqueteria) + "','" + IIf(IsNull(rsaux2!attribute8), "", rsaux2!attribute8) + "','" + IIf(IsNull(rsaux2!ATTRIBUTE9), "", rsaux2!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux2!TELEFONO), "", rsaux2!TELEFONO) + "','" + IIf(IsNull(rsaux2!Linea), "", rsaux2!Linea) + "','" + CStr(IIf(IsNull(rsaux2!ruta), "", rsaux2!ruta)) + "','" + IIf(IsNull(rsaux2!nombre_ruta), "", rsaux2!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_SURTIDO), 0, rsaux2!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", '" + IIf(IsNull(rsaux2!estacion), "", rsaux2!estacion) + "'," + CStr(var_lote) + "," + CStr(var_grupo) + ",'B00')"
                                 rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux3.Close
                              If rsaux2!Linea <> "CATALOGOS" Then
                                 If rsaux2!Linea <> "CATALOGO" Then
                                    If rsaux2!Linea <> "POP" Then
                                       If rsaux2!Linea <> "EMPAQUE" Then
                                          If rsaux2!MULTIPLO > 1 Then
                                             var_contador = var_contador + rsaux2!src_requested_quantity
                                          Else
                                             var_contador = var_contador + 1
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                              rsaux2.MoveNext
                        Wend
                        rsaux2.Close
                        var_veces_grupo = var_veces_grupo + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  
                  var_consecutivo_ubicacion = 1
                  rsaux1.Open "select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where inte_tem_Consecutivo = " + CStr(var_consecutivo) + " and SUBSTRING(UBICACION,1,3) in ('B00','P00') ORDER BY UBICACION", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update tb_temp_oracle_orden_surtido_aux_1 set consecutivo_pasillo = " + CStr(var_consecutivo_ubicacion) + "  where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ubicacion = '" + rsaux1!ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_consecutivo_ubicacion = var_consecutivo_ubicacion + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close





'PASILLO 1
                  var_grupo = var_grupo + 1
                  var_lote = var_lote + 1
                  var_consecutivo = var_consecutivo_general
                  rsaux1.Open "select distinct source_header_number, ORDEN_pedido from tb_Temp_oracle_orden_surtido_aux_2, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SOURCE_HEADER_NUMBER = pedido AND SUBSTRING(UBICACION,1,3) IN ('B01','P01','P02')  order by ORDEN_pedido", cnn, adOpenDynamic, adLockOptimistic
                  var_contador = 0
                  var_veces_grupo = 0
                  While Not rsaux1.EOF
                        'var_lote = 1
                        
                        If var_veces_grupo = 2 Then
                           var_contador = 60
                           'var_grupo = var_grupo + 1
                           var_veces_grupo = 0
                        End If
                        If var_contador >= 60 Then
                           var_grupo = var_grupo + 1
                           'var_veces_grupo = var_veces_grupo + 1
                           var_contador = 0
                        End If
                        
                        rsaux2.Open "select * from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number = " + CStr(rsaux1!source_header_number) + " AND SUBSTRING(UBICACION,1,3) IN ('B01','P01','P02')   order by ubicacion", cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux2.EOF
                              'If var_veces_grupo = 2 Then
                              '   var_veces_grupo = 1
                              '   var_grupo = var_grupo + 1
                              'End If
                              If var_contador >= 60 Then
                                 var_lote = var_lote + 1
                                 var_contador = 0
                                 var_grupo = var_grupo + 1
                                 'var_veces_grupo = var_veces_grupo + 1
                              End If
                              rsaux3.Open "SELECT * FROM tb_Temp_oracle_orden_surtido_aux_1 WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and delivery_detail_id = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 rsaux4.Open "UPDATE TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 SET SRC_REQUESTED_QUANTITY = SRC_REQUESTED_QUANTITY + " + CStr(rsaux2!src_requested_quantity) + " WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and DELIVERY_DETAIL_ID = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 var_dia = CStr(Day(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_mes = CStr(Month(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_año = CStr(Year(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 If Len(Trim(var_dia)) = 1 Then
                                    var_dia = "0" + var_dia
                                 End If
                                 If Len(Trim(var_mes)) = 1 Then
                                    var_mes = "0" + var_mes
                                 End If
                                 var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                 
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                 var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,LOTE, GRUPO, PASILLO) "
                                 var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux2!source_header_number + "', " + CStr(rsaux2!delivery_id) + "," + CStr(rsaux2!delivery_detail_id) + ", " + CStr(rsaux2!organization_id) + ",'" + IIf(IsNull(rsaux2!subinventory), "", rsaux2!subinventory) + "', " + CStr(rsaux2!delivery_line_id) + "," + CStr(rsaux2!inventory_item_id) + ",'" + rsaux2!item_description + "','" + CStr(rsaux2!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux2!src_requested_quantity) + ",'" + rsaux2!released_status + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!customer_name), "", rsaux2!customer_name) + "','" + rsaux2!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux2!collector_id), 0, rsaux2!collector_id)) + ",'" + IIf(IsNull(rsaux2!Name), "", rsaux2!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux2!ubicacion), "", rsaux2!ubicacion) + "','" + IIf(IsNull(rsaux2!establecimiento), "", rsaux2!establecimiento) + "','" + IIf(IsNull(rsaux2!nombre_Establecimiento), "", rsaux2!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux2!ORDENES), "", rsaux2!ORDENES) + "',"
                                 var_cadena = var_cadena + CStr(IIf(IsNull(rsaux2!CUST_ACCOUNT_ID), 0, rsaux2!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux2!source_header_type_name), "", rsaux2!source_header_type_name) + "', '" + IIf(IsNull(rsaux2!source_document_id), "", rsaux2!source_document_id) + "','" + IIf(IsNull(rsaux2!pais), "", rsaux2!pais) + "','" + IIf(IsNull(rsaux2!estado), "", rsaux2!estado) + "', '" + IIf(IsNull(rsaux2!municipio), "", rsaux2!municipio) + "', '" + IIf(IsNull(rsaux2!ciudad), "", rsaux2!ciudad) + "', '" + IIf(IsNull(rsaux2!colonia), "", rsaux2!colonia) + "','" + IIf(IsNull(rsaux2!DIRECCION), "", rsaux2!DIRECCION) + "', '" + IIf(IsNull(rsaux2!cp), "", rsaux2!cp) + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!site_use_id), "", rsaux2!site_use_id) + "','" + IIf(IsNull(rsaux2!paqueteria), "", rsaux2!paqueteria) + "','" + IIf(IsNull(rsaux2!attribute8), "", rsaux2!attribute8) + "','" + IIf(IsNull(rsaux2!ATTRIBUTE9), "", rsaux2!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux2!TELEFONO), "", rsaux2!TELEFONO) + "','" + IIf(IsNull(rsaux2!Linea), "", rsaux2!Linea) + "','" + CStr(IIf(IsNull(rsaux2!ruta), "", rsaux2!ruta)) + "','" + IIf(IsNull(rsaux2!nombre_ruta), "", rsaux2!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_SURTIDO), 0, rsaux2!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", '" + IIf(IsNull(rsaux2!estacion), "", rsaux2!estacion) + "'," + CStr(var_lote) + "," + CStr(var_grupo) + ",'P01')"
                                 rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux3.Close
                              If rsaux2!Linea <> "CATALOGOS" Then
                                 If rsaux2!Linea <> "CATALOGO" Then
                                    If rsaux2!Linea <> "POP" Then
                                       If rsaux2!Linea <> "EMPAQUE" Then
                                          If rsaux2!MULTIPLO > 1 Then
                                             var_contador = var_contador + rsaux2!src_requested_quantity
                                          Else
                                             var_contador = var_contador + 1
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                              rsaux2.MoveNext
                        Wend
                        rsaux2.Close
                        var_veces_grupo = var_veces_grupo + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  
                  var_consecutivo_ubicacion = 1
                  rsaux1.Open "select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where inte_tem_Consecutivo = " + CStr(var_consecutivo) + " and SUBSTRING(UBICACION,1,3) in ('B01','P01') ORDER BY UBICACION", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update tb_temp_oracle_orden_surtido_aux_1 set consecutivo_pasillo = " + CStr(var_consecutivo_ubicacion) + "  where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ubicacion = '" + rsaux1!ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_consecutivo_ubicacion = var_consecutivo_ubicacion + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  
                  'var_consecutivo_ubicacion = 1
                  rsaux1.Open "select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where inte_tem_Consecutivo = " + CStr(var_consecutivo) + " and SUBSTRING(UBICACION,1,3) in ('P02') ORDER BY UBICACION DESC", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update tb_temp_oracle_orden_surtido_aux_1 set consecutivo_pasillo = " + CStr(var_consecutivo_ubicacion) + "  where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ubicacion = '" + rsaux1!ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_consecutivo_ubicacion = var_consecutivo_ubicacion + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  
                  var_grupo = var_grupo + 1
                  
'PASILLO 3
                  var_lote = var_lote + 1
                  rsaux1.Open "select distinct source_header_number, ORDEN_pedido from tb_Temp_oracle_orden_surtido_aux_2, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SOURCE_HEADER_NUMBER = pedido AND SUBSTRING(UBICACION,1,3) IN ('P03')  order by ORDEN_pedido", cnn, adOpenDynamic, adLockOptimistic
                  var_contador = 0
                  'VAR_GRUPO = 1
                  var_veces_grupo = 0
                  While Not rsaux1.EOF
                        If var_veces_grupo = 2 Then
                           var_contador = 60
                           'var_grupo = var_grupo + 1
                           var_veces_grupo = 0
                        End If
                        If var_contador >= 60 Then
                           var_grupo = var_grupo + 1
                           'var_veces_grupo = var_veces_grupo + 1
                           var_contador = 0
                        End If
                        rsaux2.Open "select * from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number = " + CStr(rsaux1!source_header_number) + " AND SUBSTRING(UBICACION,1,3) IN ('P03')  order by ubicacion", cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux2.EOF
                              'If var_veces_grupo = 2 Then
                              '   var_veces_grupo = 1
                              '   var_grupo = var_grupo + 1
                              'End If
                              If var_contador >= 60 Then
                                 var_lote = var_lote + 1
                                 var_contador = 0
                                 var_grupo = var_grupo + 1
                                 'var_veces_grupo = var_veces_grupo + 1
                              End If
                              rsaux3.Open "SELECT * FROM tb_Temp_oracle_orden_surtido_aux_1 WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and delivery_detail_id = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 rsaux4.Open "UPDATE TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 SET SRC_REQUESTED_QUANTITY = SRC_REQUESTED_QUANTITY + " + CStr(rsaux2!src_requested_quantity) + " WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and DELIVERY_DETAIL_ID = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 var_dia = CStr(Day(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_mes = CStr(Month(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_año = CStr(Year(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 If Len(Trim(var_dia)) = 1 Then
                                    var_dia = "0" + var_dia
                                 End If
                                 If Len(Trim(var_mes)) = 1 Then
                                    var_mes = "0" + var_mes
                                 End If
                                 var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                 
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                 var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,LOTE, GRUPO, PASILLO) "
                                 var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux2!source_header_number + "', " + CStr(rsaux2!delivery_id) + "," + CStr(rsaux2!delivery_detail_id) + ", " + CStr(rsaux2!organization_id) + ",'" + IIf(IsNull(rsaux2!subinventory), "", rsaux2!subinventory) + "', " + CStr(rsaux2!delivery_line_id) + "," + CStr(rsaux2!inventory_item_id) + ",'" + rsaux2!item_description + "','" + CStr(rsaux2!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux2!src_requested_quantity) + ",'" + rsaux2!released_status + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!customer_name), "", rsaux2!customer_name) + "','" + rsaux2!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux2!collector_id), 0, rsaux2!collector_id)) + ",'" + IIf(IsNull(rsaux2!Name), "", rsaux2!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux2!ubicacion), "", rsaux2!ubicacion) + "','" + IIf(IsNull(rsaux2!establecimiento), "", rsaux2!establecimiento) + "','" + IIf(IsNull(rsaux2!nombre_Establecimiento), "", rsaux2!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux2!ORDENES), "", rsaux2!ORDENES) + "',"
                                 var_cadena = var_cadena + CStr(IIf(IsNull(rsaux2!CUST_ACCOUNT_ID), 0, rsaux2!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux2!source_header_type_name), "", rsaux2!source_header_type_name) + "', '" + IIf(IsNull(rsaux2!source_document_id), "", rsaux2!source_document_id) + "','" + IIf(IsNull(rsaux2!pais), "", rsaux2!pais) + "','" + IIf(IsNull(rsaux2!estado), "", rsaux2!estado) + "', '" + IIf(IsNull(rsaux2!municipio), "", rsaux2!municipio) + "', '" + IIf(IsNull(rsaux2!ciudad), "", rsaux2!ciudad) + "', '" + IIf(IsNull(rsaux2!colonia), "", rsaux2!colonia) + "','" + IIf(IsNull(rsaux2!DIRECCION), "", rsaux2!DIRECCION) + "', '" + IIf(IsNull(rsaux2!cp), "", rsaux2!cp) + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!site_use_id), "", rsaux2!site_use_id) + "','" + IIf(IsNull(rsaux2!paqueteria), "", rsaux2!paqueteria) + "','" + IIf(IsNull(rsaux2!attribute8), "", rsaux2!attribute8) + "','" + IIf(IsNull(rsaux2!ATTRIBUTE9), "", rsaux2!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux2!TELEFONO), "", rsaux2!TELEFONO) + "','" + IIf(IsNull(rsaux2!Linea), "", rsaux2!Linea) + "','" + CStr(IIf(IsNull(rsaux2!ruta), "", rsaux2!ruta)) + "','" + IIf(IsNull(rsaux2!nombre_ruta), "", rsaux2!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_SURTIDO), 0, rsaux2!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", '" + IIf(IsNull(rsaux2!estacion), "", rsaux2!estacion) + "'," + CStr(var_lote) + "," + CStr(var_grupo) + ",'P03')"
                                 rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux3.Close
                              If rsaux2!Linea <> "CATALOGOS" Then
                                 If rsaux2!Linea <> "CATALOGO" Then
                                    If rsaux2!Linea <> "POP" Then
                                       If rsaux2!Linea <> "EMPAQUE" Then
                                          If rsaux2!MULTIPLO > 1 Then
                                             var_contador = var_contador + rsaux2!src_requested_quantity
                                          Else
                                             var_contador = var_contador + 1
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                              rsaux2.MoveNext
                        Wend
                        rsaux2.Close
                        var_veces_grupo = var_veces_grupo + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  
                  var_consecutivo_ubicacion = 1
                  rsaux1.Open "select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where inte_tem_Consecutivo = " + CStr(var_consecutivo) + " and SUBSTRING(UBICACION,1,3) in ('P03') ORDER BY UBICACION DESC", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update tb_temp_oracle_orden_surtido_aux_1 set consecutivo_pasillo = " + CStr(var_consecutivo_ubicacion) + "  where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ubicacion = '" + rsaux1!ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_consecutivo_ubicacion = var_consecutivo_ubicacion + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  
                  
                  
'PASILLO 4
                  var_grupo = var_grupo + 1
                  var_lote = var_lote + 1
                  rsaux1.Open "select distinct source_header_number, ORDEN_pedido from tb_Temp_oracle_orden_surtido_aux_2, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SOURCE_HEADER_NUMBER = pedido AND SUBSTRING(UBICACION,1,3) IN ('P04')  order by ORDEN_pedido", cnn, adOpenDynamic, adLockOptimistic
                  var_contador = 0
                  'VAR_GRUPO = 1
                  var_veces_grupo = 0
                  While Not rsaux1.EOF
                        If var_veces_grupo = 2 Then
                           var_contador = 60
                           'var_grupo = var_grupo + 1
                           var_veces_grupo = 0
                        End If
                        If var_contador >= 60 Then
                           var_grupo = var_grupo + 1
                           'var_veces_grupo = var_veces_grupo + 1
                           var_contador = 0
                        End If
                        rsaux2.Open "select * from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number = " + CStr(rsaux1!source_header_number) + " AND SUBSTRING(UBICACION,1,3) IN ('P04')   order by ubicacion", cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux2.EOF
                              'If var_veces_grupo = 2 Then
                              '   var_veces_grupo = 1
                              '   var_grupo = var_grupo + 1
                              'End If
                              If var_contador >= 60 Then
                                 var_lote = var_lote + 1
                                 var_contador = 0
                                 var_grupo = var_grupo + 1
                                 'var_veces_grupo = var_veces_grupo + 1
                              End If
                              rsaux3.Open "SELECT * FROM tb_Temp_oracle_orden_surtido_aux_1 WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and delivery_detail_id = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 rsaux4.Open "UPDATE TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 SET SRC_REQUESTED_QUANTITY = SRC_REQUESTED_QUANTITY + " + CStr(rsaux2!src_requested_quantity) + " WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and DELIVERY_DETAIL_ID = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 var_dia = CStr(Day(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_mes = CStr(Month(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_año = CStr(Year(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 If Len(Trim(var_dia)) = 1 Then
                                    var_dia = "0" + var_dia
                                 End If
                                 If Len(Trim(var_mes)) = 1 Then
                                    var_mes = "0" + var_mes
                                 End If
                                 var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                 
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                 var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,LOTE, GRUPO, PASILLO) "
                                 var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux2!source_header_number + "', " + CStr(rsaux2!delivery_id) + "," + CStr(rsaux2!delivery_detail_id) + ", " + CStr(rsaux2!organization_id) + ",'" + IIf(IsNull(rsaux2!subinventory), "", rsaux2!subinventory) + "', " + CStr(rsaux2!delivery_line_id) + "," + CStr(rsaux2!inventory_item_id) + ",'" + rsaux2!item_description + "','" + CStr(rsaux2!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux2!src_requested_quantity) + ",'" + rsaux2!released_status + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!customer_name), "", rsaux2!customer_name) + "','" + rsaux2!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux2!collector_id), 0, rsaux2!collector_id)) + ",'" + IIf(IsNull(rsaux2!Name), "", rsaux2!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux2!ubicacion), "", rsaux2!ubicacion) + "','" + IIf(IsNull(rsaux2!establecimiento), "", rsaux2!establecimiento) + "','" + IIf(IsNull(rsaux2!nombre_Establecimiento), "", rsaux2!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux2!ORDENES), "", rsaux2!ORDENES) + "',"
                                 var_cadena = var_cadena + CStr(IIf(IsNull(rsaux2!CUST_ACCOUNT_ID), 0, rsaux2!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux2!source_header_type_name), "", rsaux2!source_header_type_name) + "', '" + IIf(IsNull(rsaux2!source_document_id), "", rsaux2!source_document_id) + "','" + IIf(IsNull(rsaux2!pais), "", rsaux2!pais) + "','" + IIf(IsNull(rsaux2!estado), "", rsaux2!estado) + "', '" + IIf(IsNull(rsaux2!municipio), "", rsaux2!municipio) + "', '" + IIf(IsNull(rsaux2!ciudad), "", rsaux2!ciudad) + "', '" + IIf(IsNull(rsaux2!colonia), "", rsaux2!colonia) + "','" + IIf(IsNull(rsaux2!DIRECCION), "", rsaux2!DIRECCION) + "', '" + IIf(IsNull(rsaux2!cp), "", rsaux2!cp) + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!site_use_id), "", rsaux2!site_use_id) + "','" + IIf(IsNull(rsaux2!paqueteria), "", rsaux2!paqueteria) + "','" + IIf(IsNull(rsaux2!attribute8), "", rsaux2!attribute8) + "','" + IIf(IsNull(rsaux2!ATTRIBUTE9), "", rsaux2!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux2!TELEFONO), "", rsaux2!TELEFONO) + "','" + IIf(IsNull(rsaux2!Linea), "", rsaux2!Linea) + "','" + CStr(IIf(IsNull(rsaux2!ruta), "", rsaux2!ruta)) + "','" + IIf(IsNull(rsaux2!nombre_ruta), "", rsaux2!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_SURTIDO), 0, rsaux2!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", '" + IIf(IsNull(rsaux2!estacion), "", rsaux2!estacion) + "'," + CStr(var_lote) + "," + CStr(var_grupo) + ",'P03')"
                                 rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux3.Close
                              If rsaux2!Linea <> "CATALOGOS" Then
                                 If rsaux2!Linea <> "CATALOGO" Then
                                    If rsaux2!Linea <> "POP" Then
                                       If rsaux2!Linea <> "EMPAQUE" Then
                                          If rsaux2!MULTIPLO > 1 Then
                                             var_contador = var_contador + rsaux2!src_requested_quantity
                                          Else
                                             var_contador = var_contador + 1
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                              rsaux2.MoveNext
                        Wend
                        rsaux2.Close
                        var_veces_grupo = var_veces_grupo + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close

                  var_consecutivo_ubicacion = 1
                  rsaux1.Open "select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where inte_tem_Consecutivo = " + CStr(var_consecutivo) + " and SUBSTRING(UBICACION,1,3) in ('P04') ORDER BY UBICACION DESC", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update tb_temp_oracle_orden_surtido_aux_1 set consecutivo_pasillo = " + CStr(var_consecutivo_ubicacion) + "  where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ubicacion = '" + rsaux1!ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_consecutivo_ubicacion = var_consecutivo_ubicacion + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  
                  
                  
'PASILLO 5 Y 6
                  var_grupo = var_grupo + 1
                  var_lote = var_lote + 1
                  rsaux1.Open "select distinct source_header_number, ORDEN_pedido from tb_Temp_oracle_orden_surtido_aux_2, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SOURCE_HEADER_NUMBER = pedido AND SUBSTRING(UBICACION,1,3) IN ('P05','P06')  order by ORDEN_pedido", cnn, adOpenDynamic, adLockOptimistic
                  var_contador = 0
                  'VAR_GRUPO = 1
                  var_veces_grupo = 0
                  While Not rsaux1.EOF
                        If var_veces_grupo = 2 Then
                           var_contador = 30
                           'var_grupo = var_grupo + 1
                           var_veces_grupo = 0
                        End If
                        If var_contador >= 30 Then
                           var_grupo = var_grupo + 1
                           'var_veces_grupo = var_veces_grupo + 1
                           var_contador = 0
                        End If
                        rsaux2.Open "select * from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number = " + CStr(rsaux1!source_header_number) + " and SUBSTRING(UBICACION,1,3) IN ('P05','P06')   order by ubicacion", cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux2.EOF
                              'If var_veces_grupo = 2 Then
                              '   var_veces_grupo = 1
                              '   var_grupo = var_grupo + 1
                              'End If
                              If var_contador >= 30 Then
                                 var_lote = var_lote + 1
                                 var_contador = 0
                                 var_grupo = var_grupo + 1
                                 'var_veces_grupo = var_veces_grupo + 1
                              End If
                              rsaux3.Open "SELECT * FROM tb_Temp_oracle_orden_surtido_aux_1 WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and delivery_detail_id = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 rsaux4.Open "UPDATE TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 SET SRC_REQUESTED_QUANTITY = SRC_REQUESTED_QUANTITY + " + CStr(rsaux2!src_requested_quantity) + " WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and DELIVERY_DETAIL_ID = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 var_dia = CStr(Day(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_mes = CStr(Month(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_año = CStr(Year(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 If Len(Trim(var_dia)) = 1 Then
                                    var_dia = "0" + var_dia
                                 End If
                                 If Len(Trim(var_mes)) = 1 Then
                                    var_mes = "0" + var_mes
                                 End If
                                 var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                 
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                 var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,LOTE, GRUPO, PASILLO) "
                                 var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux2!source_header_number + "', " + CStr(rsaux2!delivery_id) + "," + CStr(rsaux2!delivery_detail_id) + ", " + CStr(rsaux2!organization_id) + ",'" + IIf(IsNull(rsaux2!subinventory), "", rsaux2!subinventory) + "', " + CStr(rsaux2!delivery_line_id) + "," + CStr(rsaux2!inventory_item_id) + ",'" + rsaux2!item_description + "','" + CStr(rsaux2!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux2!src_requested_quantity) + ",'" + rsaux2!released_status + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!customer_name), "", rsaux2!customer_name) + "','" + rsaux2!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux2!collector_id), 0, rsaux2!collector_id)) + ",'" + IIf(IsNull(rsaux2!Name), "", rsaux2!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux2!ubicacion), "", rsaux2!ubicacion) + "','" + IIf(IsNull(rsaux2!establecimiento), "", rsaux2!establecimiento) + "','" + IIf(IsNull(rsaux2!nombre_Establecimiento), "", rsaux2!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux2!ORDENES), "", rsaux2!ORDENES) + "',"
                                 var_cadena = var_cadena + CStr(IIf(IsNull(rsaux2!CUST_ACCOUNT_ID), 0, rsaux2!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux2!source_header_type_name), "", rsaux2!source_header_type_name) + "', '" + IIf(IsNull(rsaux2!source_document_id), "", rsaux2!source_document_id) + "','" + IIf(IsNull(rsaux2!pais), "", rsaux2!pais) + "','" + IIf(IsNull(rsaux2!estado), "", rsaux2!estado) + "', '" + IIf(IsNull(rsaux2!municipio), "", rsaux2!municipio) + "', '" + IIf(IsNull(rsaux2!ciudad), "", rsaux2!ciudad) + "', '" + IIf(IsNull(rsaux2!colonia), "", rsaux2!colonia) + "','" + IIf(IsNull(rsaux2!DIRECCION), "", rsaux2!DIRECCION) + "', '" + IIf(IsNull(rsaux2!cp), "", rsaux2!cp) + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!site_use_id), "", rsaux2!site_use_id) + "','" + IIf(IsNull(rsaux2!paqueteria), "", rsaux2!paqueteria) + "','" + IIf(IsNull(rsaux2!attribute8), "", rsaux2!attribute8) + "','" + IIf(IsNull(rsaux2!ATTRIBUTE9), "", rsaux2!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux2!TELEFONO), "", rsaux2!TELEFONO) + "','" + IIf(IsNull(rsaux2!Linea), "", rsaux2!Linea) + "','" + CStr(IIf(IsNull(rsaux2!ruta), "", rsaux2!ruta)) + "','" + IIf(IsNull(rsaux2!nombre_ruta), "", rsaux2!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_SURTIDO), 0, rsaux2!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", '" + IIf(IsNull(rsaux2!estacion), "", rsaux2!estacion) + "'," + CStr(var_lote) + "," + CStr(var_grupo) + ",'P03')"
                                 rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux3.Close
                              If rsaux2!Linea <> "CATALOGOS" Then
                                 If rsaux2!Linea <> "CATALOGO" Then
                                    If rsaux2!Linea <> "POP" Then
                                       If rsaux2!Linea <> "EMPAQUE" Then
                                          If rsaux2!MULTIPLO > 1 Then
                                             var_contador = var_contador + rsaux2!src_requested_quantity
                                          Else
                                             var_contador = var_contador + 1
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                              rsaux2.MoveNext
                        Wend
                        rsaux2.Close
                        var_veces_grupo = var_veces_grupo + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close


                  var_consecutivo_ubicacion = 1
                  rsaux1.Open "select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where inte_tem_Consecutivo = " + CStr(var_consecutivo) + " and SUBSTRING(UBICACION,1,3) in ('P05') ORDER BY UBICACION", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update tb_temp_oracle_orden_surtido_aux_1 set consecutivo_pasillo = " + CStr(var_consecutivo_ubicacion) + "  where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ubicacion = '" + rsaux1!ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_consecutivo_ubicacion = var_consecutivo_ubicacion + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close


                  rsaux1.Open "select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where inte_tem_Consecutivo = " + CStr(var_consecutivo) + " and SUBSTRING(UBICACION,1,3) in ('P06') ORDER BY UBICACION DESC", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update tb_temp_oracle_orden_surtido_aux_1 set consecutivo_pasillo = " + CStr(var_consecutivo_ubicacion) + "  where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ubicacion = '" + rsaux1!ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_consecutivo_ubicacion = var_consecutivo_ubicacion + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close




'PASILLO 7
                  var_grupo = var_grupo + 1
                  rsaux1.Open "select distinct source_header_number, ORDEN_pedido from tb_Temp_oracle_orden_surtido_aux_2, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SOURCE_HEADER_NUMBER = pedido AND SUBSTRING(UBICACION,1,3) IN ('P07')  order by ORDEN_pedido", cnn, adOpenDynamic, adLockOptimistic
                  var_lote = var_lote + 1
                  var_contador = 0
                  'VAR_GRUPO = 1
                  var_veces_grupo = 0
                  While Not rsaux1.EOF
                        If var_veces_grupo = 2 Then
                           var_contador = 60
                           'var_grupo = var_grupo + 1
                           var_veces_grupo = 0
                        End If
                        If var_contador >= 60 Then
                           var_grupo = var_grupo + 1
                           'var_veces_grupo = var_veces_grupo + 1
                           var_contador = 0
                        End If
                        rsaux2.Open "select * from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number = " + CStr(rsaux1!source_header_number) + "  AND SUBSTRING(UBICACION,1,3) IN ('P07')  order by ubicacion", cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux2.EOF
                              'If var_veces_grupo = 2 Then
                              '   var_veces_grupo = 1
                              '   var_grupo = var_grupo + 1
                              'End If
                              If var_contador >= 60 Then
                                 var_lote = var_lote + 1
                                 var_contador = 0
                                 var_grupo = var_grupo + 1
                                 'var_veces_grupo = var_veces_grupo + 1
                              End If
                              rsaux3.Open "SELECT * FROM tb_Temp_oracle_orden_surtido_aux_1 WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and delivery_detail_id = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 rsaux4.Open "UPDATE TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 SET SRC_REQUESTED_QUANTITY = SRC_REQUESTED_QUANTITY + " + CStr(rsaux2!src_requested_quantity) + " WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and DELIVERY_DETAIL_ID = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 var_dia = CStr(Day(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_mes = CStr(Month(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_año = CStr(Year(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 If Len(Trim(var_dia)) = 1 Then
                                    var_dia = "0" + var_dia
                                 End If
                                 If Len(Trim(var_mes)) = 1 Then
                                    var_mes = "0" + var_mes
                                 End If
                                 var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                 
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                 var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,LOTE, GRUPO, PASILLO) "
                                 var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux2!source_header_number + "', " + CStr(rsaux2!delivery_id) + "," + CStr(rsaux2!delivery_detail_id) + ", " + CStr(rsaux2!organization_id) + ",'" + IIf(IsNull(rsaux2!subinventory), "", rsaux2!subinventory) + "', " + CStr(rsaux2!delivery_line_id) + "," + CStr(rsaux2!inventory_item_id) + ",'" + rsaux2!item_description + "','" + CStr(rsaux2!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux2!src_requested_quantity) + ",'" + rsaux2!released_status + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!customer_name), "", rsaux2!customer_name) + "','" + rsaux2!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux2!collector_id), 0, rsaux2!collector_id)) + ",'" + IIf(IsNull(rsaux2!Name), "", rsaux2!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux2!ubicacion), "", rsaux2!ubicacion) + "','" + IIf(IsNull(rsaux2!establecimiento), "", rsaux2!establecimiento) + "','" + IIf(IsNull(rsaux2!nombre_Establecimiento), "", rsaux2!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux2!ORDENES), "", rsaux2!ORDENES) + "',"
                                 var_cadena = var_cadena + CStr(IIf(IsNull(rsaux2!CUST_ACCOUNT_ID), 0, rsaux2!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux2!source_header_type_name), "", rsaux2!source_header_type_name) + "', '" + IIf(IsNull(rsaux2!source_document_id), "", rsaux2!source_document_id) + "','" + IIf(IsNull(rsaux2!pais), "", rsaux2!pais) + "','" + IIf(IsNull(rsaux2!estado), "", rsaux2!estado) + "', '" + IIf(IsNull(rsaux2!municipio), "", rsaux2!municipio) + "', '" + IIf(IsNull(rsaux2!ciudad), "", rsaux2!ciudad) + "', '" + IIf(IsNull(rsaux2!colonia), "", rsaux2!colonia) + "','" + IIf(IsNull(rsaux2!DIRECCION), "", rsaux2!DIRECCION) + "', '" + IIf(IsNull(rsaux2!cp), "", rsaux2!cp) + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!site_use_id), "", rsaux2!site_use_id) + "','" + IIf(IsNull(rsaux2!paqueteria), "", rsaux2!paqueteria) + "','" + IIf(IsNull(rsaux2!attribute8), "", rsaux2!attribute8) + "','" + IIf(IsNull(rsaux2!ATTRIBUTE9), "", rsaux2!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux2!TELEFONO), "", rsaux2!TELEFONO) + "','" + IIf(IsNull(rsaux2!Linea), "", rsaux2!Linea) + "','" + CStr(IIf(IsNull(rsaux2!ruta), "", rsaux2!ruta)) + "','" + IIf(IsNull(rsaux2!nombre_ruta), "", rsaux2!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_SURTIDO), 0, rsaux2!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", '" + IIf(IsNull(rsaux2!estacion), "", rsaux2!estacion) + "'," + CStr(var_lote) + "," + CStr(var_grupo) + ",'P03')"
                                 rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux3.Close
                              If rsaux2!Linea <> "CATALOGOS" Then
                                 If rsaux2!Linea <> "CATALOGO" Then
                                    If rsaux2!Linea <> "POP" Then
                                       If rsaux2!Linea <> "EMPAQUE" Then
                                          If rsaux2!MULTIPLO > 1 Then
                                             var_contador = var_contador + rsaux2!src_requested_quantity
                                          Else
                                             var_contador = var_contador + 1
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                              rsaux2.MoveNext
                        Wend
                        var_veces_grupo = var_veces_grupo + 1
                        rsaux2.Close
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close


                  var_consecutivo_ubicacion = 1
                  rsaux1.Open "select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where inte_tem_Consecutivo = " + CStr(var_consecutivo) + " and SUBSTRING(UBICACION,1,3) in ('P07') ORDER BY UBICACION DESC", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update tb_temp_oracle_orden_surtido_aux_1 set consecutivo_pasillo = " + CStr(var_consecutivo_ubicacion) + "  where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ubicacion = '" + rsaux1!ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_consecutivo_ubicacion = var_consecutivo_ubicacion + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close

                  
'PASILLO 8 Y 9
                  var_grupo = var_grupo + 1
                  rsaux1.Open "select distinct source_header_number, ORDEN_pedido from tb_Temp_oracle_orden_surtido_aux_2, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SOURCE_HEADER_NUMBER = pedido AND SUBSTRING(UBICACION,1,3) IN ('P08','P09')  order by ORDEN_pedido", cnn, adOpenDynamic, adLockOptimistic
                  var_lote = var_lote + 1
                  var_contador = 0
                  'VAR_GRUPO = 1
                  var_veces_grupo = 0
                  While Not rsaux1.EOF
                        If var_veces_grupo = 2 Then
                           var_contador = 60
                           'var_grupo = var_grupo + 1
                           var_veces_grupo = 0
                        End If
                        If var_contador >= 60 Then
                           var_grupo = var_grupo + 1
                           'var_veces_grupo = var_veces_grupo + 1
                           var_contador = 0
                        End If
                        rsaux2.Open "select * from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number = " + CStr(rsaux1!source_header_number) + "  AND SUBSTRING(UBICACION,1,3) IN ('P08','P09')   order by ubicacion", cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux2.EOF
                              'If var_veces_grupo = 2 Then
                              '   var_veces_grupo = 1
                              '   var_grupo = var_grupo + 1
                              'End If
                              If var_contador >= 60 Then
                                 var_lote = var_lote + 1
                                 var_contador = 0
                                 var_grupo = var_grupo + 1
                                 'var_veces_grupo = var_veces_grupo + 1
                              End If
                              rsaux3.Open "SELECT * FROM tb_Temp_oracle_orden_surtido_aux_1 WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and delivery_detail_id = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 rsaux4.Open "UPDATE TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 SET SRC_REQUESTED_QUANTITY = SRC_REQUESTED_QUANTITY + " + CStr(rsaux2!src_requested_quantity) + " WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and DELIVERY_DETAIL_ID = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 var_dia = CStr(Day(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_mes = CStr(Month(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_año = CStr(Year(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 If Len(Trim(var_dia)) = 1 Then
                                    var_dia = "0" + var_dia
                                 End If
                                 If Len(Trim(var_mes)) = 1 Then
                                    var_mes = "0" + var_mes
                                 End If
                                 var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                 
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                 var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,LOTE, GRUPO, PASILLO) "
                                 var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux2!source_header_number + "', " + CStr(rsaux2!delivery_id) + "," + CStr(rsaux2!delivery_detail_id) + ", " + CStr(rsaux2!organization_id) + ",'" + IIf(IsNull(rsaux2!subinventory), "", rsaux2!subinventory) + "', " + CStr(rsaux2!delivery_line_id) + "," + CStr(rsaux2!inventory_item_id) + ",'" + rsaux2!item_description + "','" + CStr(rsaux2!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux2!src_requested_quantity) + ",'" + rsaux2!released_status + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!customer_name), "", rsaux2!customer_name) + "','" + rsaux2!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux2!collector_id), 0, rsaux2!collector_id)) + ",'" + IIf(IsNull(rsaux2!Name), "", rsaux2!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux2!ubicacion), "", rsaux2!ubicacion) + "','" + IIf(IsNull(rsaux2!establecimiento), "", rsaux2!establecimiento) + "','" + IIf(IsNull(rsaux2!nombre_Establecimiento), "", rsaux2!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux2!ORDENES), "", rsaux2!ORDENES) + "',"
                                 var_cadena = var_cadena + CStr(IIf(IsNull(rsaux2!CUST_ACCOUNT_ID), 0, rsaux2!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux2!source_header_type_name), "", rsaux2!source_header_type_name) + "', '" + IIf(IsNull(rsaux2!source_document_id), "", rsaux2!source_document_id) + "','" + IIf(IsNull(rsaux2!pais), "", rsaux2!pais) + "','" + IIf(IsNull(rsaux2!estado), "", rsaux2!estado) + "', '" + IIf(IsNull(rsaux2!municipio), "", rsaux2!municipio) + "', '" + IIf(IsNull(rsaux2!ciudad), "", rsaux2!ciudad) + "', '" + IIf(IsNull(rsaux2!colonia), "", rsaux2!colonia) + "','" + IIf(IsNull(rsaux2!DIRECCION), "", rsaux2!DIRECCION) + "', '" + IIf(IsNull(rsaux2!cp), "", rsaux2!cp) + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!site_use_id), "", rsaux2!site_use_id) + "','" + IIf(IsNull(rsaux2!paqueteria), "", rsaux2!paqueteria) + "','" + IIf(IsNull(rsaux2!attribute8), "", rsaux2!attribute8) + "','" + IIf(IsNull(rsaux2!ATTRIBUTE9), "", rsaux2!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux2!TELEFONO), "", rsaux2!TELEFONO) + "','" + IIf(IsNull(rsaux2!Linea), "", rsaux2!Linea) + "','" + CStr(IIf(IsNull(rsaux2!ruta), "", rsaux2!ruta)) + "','" + IIf(IsNull(rsaux2!nombre_ruta), "", rsaux2!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_SURTIDO), 0, rsaux2!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", '" + IIf(IsNull(rsaux2!estacion), "", rsaux2!estacion) + "'," + CStr(var_lote) + "," + CStr(var_grupo) + ",'P03')"
                                 rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux3.Close
                              If rsaux2!Linea <> "CATALOGOS" Then
                                 If rsaux2!Linea <> "CATALOGO" Then
                                    If rsaux2!Linea <> "POP" Then
                                       If rsaux2!Linea <> "EMPAQUE" Then
                                          If rsaux2!MULTIPLO > 1 Then
                                             var_contador = var_contador + rsaux2!src_requested_quantity
                                          Else
                                             var_contador = var_contador + 1
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                              rsaux2.MoveNext
                        Wend
                        rsaux2.Close
                        var_veces_grupo = var_veces_grupo + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  
                  
                  var_consecutivo_ubicacion = 1
                  rsaux1.Open "select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where inte_tem_Consecutivo = " + CStr(var_consecutivo) + " and SUBSTRING(UBICACION,1,3) in ('P08') ORDER BY UBICACION", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update tb_temp_oracle_orden_surtido_aux_1 set consecutivo_pasillo = " + CStr(var_consecutivo_ubicacion) + "  where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ubicacion = '" + rsaux1!ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_consecutivo_ubicacion = var_consecutivo_ubicacion + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  
                  
                  rsaux1.Open "select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where inte_tem_Consecutivo = " + CStr(var_consecutivo) + " and SUBSTRING(UBICACION,1,3) in ('P09') ORDER BY UBICACION DESC", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update tb_temp_oracle_orden_surtido_aux_1 set consecutivo_pasillo = " + CStr(var_consecutivo_ubicacion) + "  where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ubicacion = '" + rsaux1!ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_consecutivo_ubicacion = var_consecutivo_ubicacion + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  

                  var_consecutivo = var_consecutivo_general
                  rsaux1.Open "select distinct source_header_number, ORDEN_pedido from tb_Temp_oracle_orden_surtido_aux_2, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SOURCE_HEADER_NUMBER = pedido AND SUBSTRING(UBICACION,1,3) not IN ('B00', 'B01', 'P01', 'P02', 'P03', 'P04', 'P05', 'P06', 'P07', 'P08', 'P09','P00')  order by ORDEN_pedido", cnn, adOpenDynamic, adLockOptimistic
                  var_contador = 0
                  var_veces_grupo = 0
                  var_lote = 1
                  While Not rsaux1.EOF
                        'var_lote = 1
                        
                        If var_veces_grupo = 2 Then
                           var_contador = 60
                           'var_grupo = var_grupo + 1
                           var_veces_grupo = 0
                        End If
                        If var_contador >= 60 Then
                           var_grupo = var_grupo + 1
                           'var_veces_grupo = var_veces_grupo + 1
                           var_contador = 0
                        End If
                        
                        rsaux2.Open "select * from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number = " + CStr(rsaux1!source_header_number) + " AND SUBSTRING(UBICACION,1,3) NOT IN ('B00', 'B01', 'P01', 'P02', 'P03', 'P04', 'P05', 'P06', 'P07', 'P08', 'P09','P00')   order by ubicacion", cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux2.EOF
                              'If var_veces_grupo = 2 Then
                              '   var_veces_grupo = 1
                              '   var_grupo = var_grupo + 1
                              'End If
                              If var_contador >= 60 Then
                                 var_lote = var_lote + 1
                                 var_contador = 0
                                 var_grupo = var_grupo + 1
                                 'var_veces_grupo = var_veces_grupo + 1
                              End If
                              rsaux3.Open "SELECT * FROM tb_Temp_oracle_orden_surtido_aux_1 WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and delivery_detail_id = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 rsaux4.Open "UPDATE TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 SET SRC_REQUESTED_QUANTITY = SRC_REQUESTED_QUANTITY + " + CStr(rsaux2!src_requested_quantity) + " WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and DELIVERY_DETAIL_ID = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 var_dia = CStr(Day(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_mes = CStr(Month(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_año = CStr(Year(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 If Len(Trim(var_dia)) = 1 Then
                                    var_dia = "0" + var_dia
                                 End If
                                 If Len(Trim(var_mes)) = 1 Then
                                    var_mes = "0" + var_mes
                                 End If
                                 var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                 
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                 var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,LOTE, GRUPO, PASILLO) "
                                 var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux2!source_header_number + "', " + CStr(rsaux2!delivery_id) + "," + CStr(rsaux2!delivery_detail_id) + ", " + CStr(rsaux2!organization_id) + ",'" + IIf(IsNull(rsaux2!subinventory), "", rsaux2!subinventory) + "', " + CStr(rsaux2!delivery_line_id) + "," + CStr(rsaux2!inventory_item_id) + ",'" + rsaux2!item_description + "','" + CStr(rsaux2!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux2!src_requested_quantity) + ",'" + rsaux2!released_status + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!customer_name), "", rsaux2!customer_name) + "','" + rsaux2!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux2!collector_id), 0, rsaux2!collector_id)) + ",'" + IIf(IsNull(rsaux2!Name), "", rsaux2!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux2!ubicacion), "", rsaux2!ubicacion) + "','" + IIf(IsNull(rsaux2!establecimiento), "", rsaux2!establecimiento) + "','" + IIf(IsNull(rsaux2!nombre_Establecimiento), "", rsaux2!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux2!ORDENES), "", rsaux2!ORDENES) + "',"
                                 var_cadena = var_cadena + CStr(IIf(IsNull(rsaux2!CUST_ACCOUNT_ID), 0, rsaux2!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux2!source_header_type_name), "", rsaux2!source_header_type_name) + "', '" + IIf(IsNull(rsaux2!source_document_id), "", rsaux2!source_document_id) + "','" + IIf(IsNull(rsaux2!pais), "", rsaux2!pais) + "','" + IIf(IsNull(rsaux2!estado), "", rsaux2!estado) + "', '" + IIf(IsNull(rsaux2!municipio), "", rsaux2!municipio) + "', '" + IIf(IsNull(rsaux2!ciudad), "", rsaux2!ciudad) + "', '" + IIf(IsNull(rsaux2!colonia), "", rsaux2!colonia) + "','" + IIf(IsNull(rsaux2!DIRECCION), "", rsaux2!DIRECCION) + "', '" + IIf(IsNull(rsaux2!cp), "", rsaux2!cp) + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!site_use_id), "", rsaux2!site_use_id) + "','" + IIf(IsNull(rsaux2!paqueteria), "", rsaux2!paqueteria) + "','" + IIf(IsNull(rsaux2!attribute8), "", rsaux2!attribute8) + "','" + IIf(IsNull(rsaux2!ATTRIBUTE9), "", rsaux2!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux2!TELEFONO), "", rsaux2!TELEFONO) + "','" + IIf(IsNull(rsaux2!Linea), "", rsaux2!Linea) + "','" + CStr(IIf(IsNull(rsaux2!ruta), "", rsaux2!ruta)) + "','" + IIf(IsNull(rsaux2!nombre_ruta), "", rsaux2!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_SURTIDO), 0, rsaux2!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", '" + IIf(IsNull(rsaux2!estacion), "", rsaux2!estacion) + "'," + CStr(var_lote) + "," + CStr(var_grupo) + ",'B00')"
                                 rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux3.Close
                              If rsaux2!Linea <> "CATALOGOS" Then
                                 If rsaux2!Linea <> "CATALOGO" Then
                                    If rsaux2!Linea <> "POP" Then
                                       If rsaux2!Linea <> "EMPAQUE" Then
                                          If rsaux2!MULTIPLO > 1 Then
                                             var_contador = var_contador + rsaux2!src_requested_quantity
                                          Else
                                             var_contador = var_contador + 1
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                              rsaux2.MoveNext
                        Wend
                        rsaux2.Close
                        var_veces_grupo = var_veces_grupo + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  
                  var_consecutivo_ubicacion = 1
                  rsaux1.Open "select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where inte_tem_Consecutivo = " + CStr(var_consecutivo) + " and SUBSTRING(UBICACION,1,3) NOT in ('B00', 'B01', 'P01', 'P02', 'P03', 'P04', 'P05', 'P06', 'P07', 'P08', 'P09','P00') ORDER BY UBICACION", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        rsaux2.Open "update tb_temp_oracle_orden_surtido_aux_1 set consecutivo_pasillo = " + CStr(var_consecutivo_ubicacion) + "  where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and ubicacion = '" + rsaux1!ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_consecutivo_ubicacion = var_consecutivo_ubicacion + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close



End Sub





Private Sub ACTUALIZA_INFORMACION()
    Dim var_j As Integer
    rs.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
    If Me.lv_pedidos.ListItems.Count > 0 Then
       For var_j = 1 To Me.lv_pedidos.ListItems.Count
           Me.lv_pedidos.ListItems.Item(var_j).Selected = True
           If Me.lv_pedidos.selectedItem <> "10000000" Then
              var_pedido = CDbl(Me.lv_pedidos.selectedItem)
              strconsulta = "SELECT oh.order_number,sum(NVL(ol.requested_quantity, 0)) As CANTIDAD_PEDIDA FROM oe_order_lines_all ola, oe_order_headers_all oh, WSH_DELIVERABLES_V ol, xxvia_system_items_b g, xxvia_vw_articulos_cat h WHERE order_number = ? AND oh.header_id = ol.source_header_id AND ol.organization_id = 93 AND ol.inventory_item_id = g.inventory_item_id AND g.organization_id = ol.organization_id AND h.item_id = g.inventory_item_id AND h.organization_id = g.organization_id AND ola.line_id = ol.source_line_id and LINEA IN ('CATALOGO','POP','CATALOGOS') and released_status = 'Y' group by oh.order_number"
              With comandoORA
                  .ActiveConnection = cnnoracle_4
                   .CommandType = adCmdText
                   .CommandText = strconsulta
                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(var_pedido))
                   .Parameters.Append parametro
              End With
              Set rsaux6 = comandoORA.execute
              Set comandoORA = Nothing
              Set parametro = Nothing
              If rsaux6.EOF Then
                 VAR_CANTIDAD_CATALOGOS = 0
              Else
                 VAR_CANTIDAD_CATALOGOS = IIf(IsNull(rsaux6!CANTIDAD_PEDIDA), 0, rsaux6!CANTIDAD_PEDIDA)
              End If
              rsaux6.Close
              strconsulta = "SELECT oh.order_number,sum(NVL(ol.requested_quantity, 0)) As CANTIDAD_PEDIDA FROM oe_order_lines_all ola, oe_order_headers_all oh, WSH_DELIVERABLES_V ol, xxvia_system_items_b g, xxvia_vw_articulos_cat h WHERE order_number = ? AND oh.header_id = ol.source_header_id AND ol.organization_id = 93 AND ol.inventory_item_id = g.inventory_item_id AND g.organization_id = ol.organization_id AND h.item_id = g.inventory_item_id AND h.organization_id = g.organization_id AND ola.line_id = ol.source_line_id and LINEA not IN ('CATALOGO','POP','CATALOGOS') and released_status = 'Y' group by oh.order_number"
              With comandoORA
                   .ActiveConnection = cnnoracle_4
                   .CommandType = adCmdText
                   .CommandText = strconsulta
                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(var_pedido))
                   .Parameters.Append parametro
              End With
              Set rsaux6 = comandoORA.execute
              Set comandoORA = Nothing
              Set parametro = Nothing
              If rsaux6.EOF Then
                 VAR_CANTIDAD_SIN_CATALOGOS = 0
              Else
                 VAR_CANTIDAD_SIN_CATALOGOS = IIf(IsNull(rsaux6!CANTIDAD_PEDIDA), 0, rsaux6!CANTIDAD_PEDIDA)
              End If
              rsaux6.Close
           Else
              VAR_CANTIDAD_CATALOGOS = 0
              VAR_CANTIDAD_SIN_CATALOGOS = 0
           End If
           rsaux6.Open "SELECT * FROM tb_oracle_pedidos_asignados_embarques WHERE PEDIDO = '" + Me.lv_pedidos.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
           If Not rsaux6.EOF Then
              VAR_A_CANTIDAD_SIN_CATALOGOS = IIf(IsNull(rsaux6!CANTIDAD_SIN_CATALOGOS), 0, rsaux6!CANTIDAD_SIN_CATALOGOS)
              VAR_A_CANTIDAD_CATALOGOS = IIf(IsNull(rsaux6!CANTIDAD_CATALOGOS), 0, rsaux6!CANTIDAD_CATALOGOS)
              If VAR_CANTIDAD_SIN_CATALOGOS > VAR_A_CANTIDAD_SIN_CATALOGOS Then
                 rsaux7.Open "UPDATE tb_oracle_pedidos_asignados_embarques SET CANTIDAD_SIN_CATALOGOS = " + CStr(VAR_CANTIDAD_SIN_CATALOGOS) + " WHERE PEDIDO = " + CStr(var_pedido), cnn, adOpenDynamic, adLockOptimistic
              End If
              If VAR_CANTIDAD_CATALOGOS > VAR_A_CANTIDAD_CATALOGOS Then
                 rsaux7.Open "UPDATE tb_oracle_pedidos_asignados_embarques SET CANTIDAD_CATALOGOS = " + CStr(VAR_CANTIDAD_CATALOGOS) + " WHERE PEDIDO = " + CStr(var_pedido), cnn, adOpenDynamic, adLockOptimistic
              End If
              If rsaux6!PIEZAS < VAR_CANTIDAD_CATALOGOS + VAR_CANTIDAD_SIN_CATALOGOS Then
                 rsaux7.Open "UPDATE tb_oracle_pedidos_asignados_embarques SET PIEZAS = " + CStr(VAR_CANTIDAD_CATALOGOS + VAR_CANTIDAD_SIN_CATALOGOS) + " WHERE PEDIDO = " + CStr(var_pedido), cnn, adOpenDynamic, adLockOptimistic
              End If
           End If
           rsaux6.Close
       Next var_j
    End If
End Sub


Private Sub crea_tablas()
        var_imprime_pedidos = 1
        rsaux2.Open "select * from tb_temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo_general), cnn, adOpenDynamic, adLockOptimistic
              If cnnoracle_4.State = 7 Then
                 cnnoracle_4.Close
                 If var_prueba = 2 Then
                    cnnoracle_4.Open "Provider=OraOLEDB.Oracle.1;User ID=apps;Data Source=tvia;Extended Properties=;Persist Security Info=True;Password=apps"
                 Else
                    cnnoracle_4.Open "Provider=OraOLEDB.Oracle.1;User ID=apps;Data Source=pvia;Extended Properties=;Persist Security Info=True;Password=apps"
                 End If
              End If
        
        While Not rsaux2.EOF
              rsaux3.Open "select * from xxvia_tb_pedidos_divididos  where SOURCE_HEADER_NUMBER = " + rsaux2!source_header_number + " and DELIVERY_ID = " + CStr(rsaux2!delivery_id) + " and DELIVERY_DETAIL_ID = " + CStr(rsaux2!delivery_detail_id) + " and ORGANIZATION_ID = " + CStr(rsaux2!organization_id) + " and DELIVERY_LINE_ID = " + CStr(rsaux2!delivery_line_id) + " and INVENTORY_ITEM_ID = " + CStr(rsaux2!inventory_item_id) + " and SOURCE_LINE_NUMBER = '" + rsaux2!SOURCE_LINE_NUMBER + "' and lote = " + CStr(rsaux2!LOTE), cnnoracle_4, adOpenDynamic, adLockOptimistic
              
              If rsaux3.EOF Then
                 'MsgBox CStr(rsaux2!source_header_number) + " " + rsaux2!segment1 + " " + CStr(rsaux2!src_requested_quantity)
                 var_cadena = "INSERT INTO XXVIA_TB_PEDIDOS_DIVIDIDOS (SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,SEGMENT1,LOTE,CUSTOMER_NAME,COLLECTOR_ID,NAME,DATE_REQUESTED,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,CUST_ACCOUNT_ID,SOURCE_HEADER_TYPE_NAME,SOURCE_DOCUMENT_ID,SITE_USE_ID,linea,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,Embarque,ESTACION,ORDEN_CARGA) VALUES "
                 var_cadena = var_cadena + "(" + CStr(rsaux2!source_header_number) + "," + CStr(rsaux2!delivery_id) + "," + CStr(rsaux2!delivery_detail_id) + "," + CStr(rsaux2!organization_id) + ",'" + rsaux2!subinventory + "'," + CStr(rsaux2!delivery_line_id) + "," + CStr(rsaux2!inventory_item_id) + ",'" + rsaux2!item_description + "','" + rsaux2!SOURCE_LINE_NUMBER + "'," + CStr(rsaux2!src_requested_quantity) + ",'" + rsaux2!released_status + "','" + rsaux2!SEGMENT1 + "'," + CStr(rsaux2!LOTE) + ",'" + rsaux2!customer_name + "'," + CStr(rsaux2!collector_id) + ",'" + rsaux2!Name + "','" + CStr(rsaux2!DATE_REQUESTED) + "'," + CStr(rsaux2!establecimiento) + ",'" + rsaux2!nombre_Establecimiento + "'," + CStr(rsaux2!CUST_ACCOUNT_ID) + ",'" + rsaux2!source_header_type_name + "','" + CStr(IIf(IsNull(rsaux2!source_document_id), "", rsaux2!source_document_id)) + "'," + CStr(rsaux2!site_use_id) + ",'" + IIf(IsNull(rsaux2!Linea), "", rsaux2!Linea) + "'," + CStr(IIf(IsNull(rsaux2!ruta), 0, rsaux2!ruta))
                 var_cadena = var_cadena + ",'" + IIf(IsNull(rsaux2!nombre_ruta), "", rsaux2!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_SURTIDO), 0, rsaux2!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ",'" + IIf(IsNull(rsaux2!estacion), "", rsaux2!estacion) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_CARGA), 0, rsaux2!ORDEN_CARGA)) + ")"
                 'MsgBox var_cadena
                 rsaux4.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
              End If
              rsaux3.Close
              rsaux2.MoveNext
        Wend
        rsaux2.Close
        
        rsaux2.Open "DELETE FROM TB_TEMP_ORACLE_COMPARACION_PEDIDO_DIVIDIDO WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo_general), cnn, adOpenDynamic, adLockOptimistic
        rsaux2.Open "INSERT INTO TB_TEMP_ORACLE_COMPARACION_PEDIDO_DIVIDIDO (INTE_TEM_CONSECUTIVO, PEDIDO, CODIGO, LINEA, CANTIDAD, DIVIDIDO) SELECT INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER, SEGMENT1, SOURCE_LINE_NUMBER, SUM(SRC_REQUESTED_QUANTITY) AS CANTIDAD, 0 FROM tb_temp_oracle_orden_surtido WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo_general) + " GROUP BY INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER, SEGMENT1, SOURCE_LINE_NUMBER ", cnn, adOpenDynamic, adLockOptimistic
        rsaux2.Open "SELECT DISTINCT PEDIDO FROM TB_TEMP_ORACLE_COMPARACION_PEDIDO_DIVIDIDO WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo_general), cnn, adOpenDynamic, adLockOptimistic
        While Not rsaux2.EOF
              rsaux4.Open "select SOURCE_HEADER_NUMBER, SEGMENT1, SOURCE_LINE_NUMBER, SUM(SRC_REQUESTED_QUANTITY) AS CANTIDAD from xxvia_tb_pedidos_divididos where source_header_number =  " + CStr(rsaux2!pedido) + "  GROUP BY SOURCE_HEADER_NUMBER, SEGMENT1, SOURCE_LINE_NUMBER", cnnoracle_4, adOpenDynamic, adLockOptimistic
              While Not rsaux4.EOF
                    rsaux5.Open "UPDATE TB_TEMP_ORACLE_COMPARACION_PEDIDO_DIVIDIDO SET DIVIDIDO = " + CStr(rsaux4!cantidad) + " WHERE INTE_TEM_CONSECUTIVO =" + CStr(var_consecutivo_general) + " AND PEDIDO = " + CStr(rsaux4!source_header_number) + " AND CODIGO = '" + rsaux4!SEGMENT1 + "' AND LINEA = '" + rsaux4!SOURCE_LINE_NUMBER + "'", cnn, adOpenDynamic, adLockOptimistic
                    rsaux4.MoveNext
              Wend
              rsaux2.MoveNext
              rsaux4.Close
        Wend
        rsaux2.Close
        
        rsaux2.Open "SELECT distinct pedido FROM TB_TEMP_ORACLE_COMPARACION_PEDIDO_DIVIDIDO WHERE CANTIDAD <> DIVIDIDO AND INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo_general), cnn, adOpenDynamic, adLockOptimistic
        If Not rsaux2.EOF Then
           var_imprime_pedidos = 0
           MsgBox "Existen diferencias en el pedido " + CStr(rsaux2!pedido) + " con la división del pedido y la impresiòn de los documentos, se eliminara la informaciòn y se tendra que volver a generar la impresiòn de los pedidos", vbOKOnly, "ATENCION"
           While Not rsaux2.EOF
                 rsaux6.Open "select * from xxvia_tb_salidas_cajas where source_header_number = " + CStr(rsaux2!pedido), cnnoracle_4, adOpenDynamic, adLockOptimistic
                 If rsaux6.EOF Then
                    
                     strconsulta = "select * from xxvia_tb_pedidos_divididos where source_header_number = ? AND RELEASED_STATUS = 'Y'"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(rsaux2!pedido))
                          .Parameters.Append parametro
                     End With
                     Set rsaux10 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     If rsaux10.EOF Then
                        rsaux5.Open "DELETE FROM xxvia_tb_pedidos_divididos where SOURCE_HEADER_NUMBER = " + CStr(rsaux2!pedido), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux10.Close
                 Else
                    MsgBox "El pedido " + CStr(rsaux2!pedido) + " ya esta leido", vbOKOnly, "ATENCION"
                 End If
                 rsaux6.Close
                 rsaux2.MoveNext
           Wend
        End If
        rsaux2.Close
        rsaux2.Open "DELETE TB_TEMP_ORACLE_COMPARACION_PEDIDO_DIVIDIDO WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo_general), cnn, adOpenDynamic, adLockOptimistic
        
End Sub

Private Sub cmd_actualiza_informacion_pedidos_Click()
   Call ACTUALIZA_INFORMACION
End Sub

Private Sub cmd_anden_1_Click()
   froracle_asignacion_embarques.Enabled = False
   var_activa_forma_embarques = "froracle_asignacion_embarques"
   var_anden_asignar = 1
   var_anden_global = 1
   frmoracle_embarques.Show 1
   'frmoracle_asignar_pedidos_embarque.Show 1
    
   Me.lv_embarques_1.ListItems.Clear
   var_dia_s = CStr(Day(Date))
   If Len(var_dia_s) = 1 Then
      var_dia_s = "0" + var_dia_s
   End If
   var_mes_s = CStr(Month(Date))
   If Len(var_mes_s) = 1 Then
      var_mes_s = "0" + var_mes_s
   End If
   rs.Open "select distinct embarque as embarque from xxvia_Tb_encabezado_embarques where jaula = 1 and to_char(fecha_inicio,'DD')  = '" + CStr(var_dia_s) + "' and to_char(fecha_inicio,'MM')  = '" + CStr(var_mes_s) + "' and to_char(fecha_inicio,'yyyy')  = '" + CStr(Year(Date)) + "' and organizacion = " + CStr(var_unidad_organizacional), cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_Cadena_embarques = ""
   While Not rs.EOF
         If var_Cadena_embarques = "" Then
            var_Cadena_embarques = CStr(rs!Embarque)
         Else
            var_Cadena_embarques = var_Cadena_embarques + "," + CStr(rs!Embarque)
         End If
         rs.MoveNext
   Wend
   rs.Close
   If var_Cadena_embarques <> "" Then
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open "select distinct EMBARQUE from tb_oracle_pedidos_asignados_embarques where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_embarques_1.ListItems.Add(, , rs!Embarque)
            'list_item.SubItems(1) = rs!NOMBRE_AGENTE
            rs.MoveNext
      Wend
      rs.Close
      If lv_embarques_1.ListItems.Count > 0 Then
         lv_embarques_1.ListItems(1).Selected = True
         rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_1.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         lv_pedidos.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
               'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
               list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
               'rsaux.Close
               list_item.SubItems(2) = rs!Cliente
               list_item.SubItems(3) = rs!PIEZAS
               list_item.SubItems(5) = rs!orden_pedido
               rs.MoveNext
         Wend
         rs.Close
         If lv_pedidos.ListItems.Count > 11 Then
            Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
         Else
            Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
         End If
         rs.Open "SELECT SUM(PIEZAS) FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
         Me.lbl_cantidad_1 = Format(rs(0).Value, "###,###,##0.00")
         rs.Close
         
         
         
         
      End If
   End If
End Sub

Private Sub cmd_anden_10_Click()
   froracle_asignacion_embarques.Enabled = False
   var_activa_forma_embarques = "froracle_asignacion_embarques"
   var_anden_asignar = 10
   var_anden_global = 10
   frmoracle_embarques.Show 1
   'frmoracle_asignar_pedidos_embarque.Show 1
    
   Me.lv_embarques_10.ListItems.Clear
   var_dia_s = CStr(Day(Date))
   If Len(var_dia_s) = 1 Then
      var_dia_s = "0" + var_dia_s
   End If
   var_mes_s = CStr(Month(Date))
   If Len(var_mes_s) = 1 Then
      var_mes_s = "0" + var_mes_s
   End If
   rs.Open "select distinct embarque as embarque from xxvia_Tb_encabezado_embarques where jaula = 10 and to_char(fecha_inicio,'DD')  = '" + CStr(var_dia_s) + "' and to_char(fecha_inicio,'MM')  = '" + CStr(var_mes_s) + "' and to_char(fecha_inicio,'yyyy')  = '" + CStr(Year(Date)) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_Cadena_embarques = ""
   While Not rs.EOF
         If var_Cadena_embarques = "" Then
            var_Cadena_embarques = CStr(rs!Embarque)
         Else
            var_Cadena_embarques = var_Cadena_embarques + "," + CStr(rs!Embarque)
         End If
         rs.MoveNext
   Wend
   rs.Close
   If var_Cadena_embarques <> "" Then
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open "select distinct EMBARQUE from tb_oracle_pedidos_asignados_embarques where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
      
            Set list_item = lv_embarques_10.ListItems.Add(, , rs!Embarque)
            'list_item.SubItems(1) = rs!NOMBRE_AGENTE
            rs.MoveNext
      Wend
      rs.Close
      If lv_embarques_10.ListItems.Count > 0 Then
         lv_embarques_10.ListItems(1).Selected = True
         rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_10.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         lv_pedidos.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
               'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
               list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
               'rsaux.Close
               list_item.SubItems(2) = rs!Cliente
               list_item.SubItems(3) = rs!PIEZAS
               list_item.SubItems(5) = rs!orden_pedido
               rs.MoveNext
         Wend
         rs.Close
         If lv_pedidos.ListItems.Count > 11 Then
            Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
         Else
            Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
         End If
         rs.Open "SELECT SUM(PIEZAS) FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
         Me.lbl_cantidad_10 = Format(rs(0).Value, "###,###,##0.00")
         rs.Close
      End If
   End If
End Sub

Private Sub cmd_anden_2_Click()
   froracle_asignacion_embarques.Enabled = False
   var_activa_forma_embarques = "froracle_asignacion_embarques"
   var_anden_asignar = 2
   var_anden_global = 2
   frmoracle_embarques.Show 1
   'frmoracle_asignar_pedidos_embarque.Show 1
    
   Me.lv_embarques_2.ListItems.Clear
   var_dia_s = CStr(Day(Date))
   If Len(var_dia_s) = 1 Then
      var_dia_s = "0" + var_dia_s
   End If
   var_mes_s = CStr(Month(Date))
   If Len(var_mes_s) = 1 Then
      var_mes_s = "0" + var_mes_s
   End If
   rs.Open "select distinct embarque as embarque from xxvia_Tb_encabezado_embarques where jaula = 2 and to_char(fecha_inicio,'DD')  = '" + CStr(var_dia_s) + "' and to_char(fecha_inicio,'MM')  = '" + CStr(var_mes_s) + "' and to_char(fecha_inicio,'yyyy')  = '" + CStr(Year(Date)) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_Cadena_embarques = ""
   While Not rs.EOF
         If var_Cadena_embarques = "" Then
            var_Cadena_embarques = CStr(rs!Embarque)
         Else
            var_Cadena_embarques = var_Cadena_embarques + "," + CStr(rs!Embarque)
         End If
         rs.MoveNext
   Wend
   rs.Close
   If var_Cadena_embarques <> "" Then
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open "select distinct EMBARQUE from tb_oracle_pedidos_asignados_embarques where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_embarques_2.ListItems.Add(, , rs!Embarque)
            'list_item.SubItems(1) = rs!NOMBRE_AGENTE
            rs.MoveNext
      Wend
      rs.Close
      If lv_embarques_2.ListItems.Count > 0 Then
         lv_embarques_2.ListItems(1).Selected = True
         rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_2.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         lv_pedidos.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
               'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
               list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
               'rsaux.Close
               list_item.SubItems(2) = rs!Cliente
               list_item.SubItems(3) = rs!PIEZAS
               list_item.SubItems(5) = rs!orden_pedido
               rs.MoveNext
         Wend
         rs.Close
         If lv_pedidos.ListItems.Count > 11 Then
            Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
         Else
            Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
         End If
         rs.Open "SELECT SUM(PIEZAS) FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
         Me.lbl_cantidad_2 = Format(rs(0).Value, "###,###,##0.00")
         rs.Close
         
         
         
         
      End If
   End If
End Sub

Private Sub cmd_anden_3_Click()
   froracle_asignacion_embarques.Enabled = False
   var_activa_forma_embarques = "froracle_asignacion_embarques"
   var_anden_asignar = 3
   var_anden_global = 3
   frmoracle_embarques.Show 1
   'frmoracle_asignar_pedidos_embarque.Show 1
    
   Me.lv_embarques_3.ListItems.Clear
   var_dia_s = CStr(Day(Date))
   If Len(var_dia_s) = 1 Then
      var_dia_s = "0" + var_dia_s
   End If
   var_mes_s = CStr(Month(Date))
   If Len(var_mes_s) = 1 Then
      var_mes_s = "0" + var_mes_s
   End If
   rs.Open "select distinct embarque as embarque from xxvia_Tb_encabezado_embarques where jaula = 3 and to_char(fecha_inicio,'DD')  = '" + CStr(var_dia_s) + "' and to_char(fecha_inicio,'MM')  = '" + CStr(var_mes_s) + "' and to_char(fecha_inicio,'yyyy')  = '" + CStr(Year(Date)) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_Cadena_embarques = ""
   While Not rs.EOF
         If var_Cadena_embarques = "" Then
            var_Cadena_embarques = CStr(rs!Embarque)
         Else
            var_Cadena_embarques = var_Cadena_embarques + "," + CStr(rs!Embarque)
         End If
         rs.MoveNext
   Wend
   rs.Close
   If var_Cadena_embarques <> "" Then
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open "select distinct EMBARQUE from tb_oracle_pedidos_asignados_embarques where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_embarques_3.ListItems.Add(, , rs!Embarque)
            'list_item.SubItems(1) = rs!NOMBRE_AGENTE
            rs.MoveNext
      Wend
      rs.Close
      If lv_embarques_3.ListItems.Count > 0 Then
         lv_embarques_3.ListItems(1).Selected = True
         rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_3.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         lv_pedidos.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
               'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
               list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
               'rsaux.Close
               list_item.SubItems(2) = rs!Cliente
               list_item.SubItems(3) = rs!PIEZAS
               list_item.SubItems(5) = rs!orden_pedido
               rs.MoveNext
         Wend
         rs.Close
         If lv_pedidos.ListItems.Count > 11 Then
            Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
         Else
            Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
         End If
         rs.Open "SELECT SUM(PIEZAS) FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
         Me.lbl_cantidad_3 = Format(rs(0).Value, "###,###,##0.00")
         rs.Close
         
         
         
         
      End If
   End If
End Sub

Private Sub cmd_anden_4_Click()
   froracle_asignacion_embarques.Enabled = False
   var_activa_forma_embarques = "froracle_asignacion_embarques"
   var_anden_asignar = 4
   var_anden_global = 4
   frmoracle_embarques.Show 1
   'frmoracle_asignar_pedidos_embarque.Show 1
    
   Me.lv_embarques_4.ListItems.Clear
   var_dia_s = CStr(Day(Date))
   If Len(var_dia_s) = 1 Then
      var_dia_s = "0" + var_dia_s
   End If
   var_mes_s = CStr(Month(Date))
   If Len(var_mes_s) = 1 Then
      var_mes_s = "0" + var_mes_s
   End If
   rs.Open "select distinct embarque as embarque from xxvia_Tb_encabezado_embarques where jaula = 4 and to_char(fecha_inicio,'DD')  = '" + CStr(var_dia_s) + "' and to_char(fecha_inicio,'MM')  = '" + CStr(var_mes_s) + "' and to_char(fecha_inicio,'yyyy')  = '" + CStr(Year(Date)) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_Cadena_embarques = ""
   While Not rs.EOF
         If var_Cadena_embarques = "" Then
            var_Cadena_embarques = CStr(rs!Embarque)
         Else
            var_Cadena_embarques = var_Cadena_embarques + "," + CStr(rs!Embarque)
         End If
         rs.MoveNext
   Wend
   rs.Close
   If var_Cadena_embarques <> "" Then
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open "select distinct EMBARQUE from tb_oracle_pedidos_asignados_embarques where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_embarques_4.ListItems.Add(, , rs!Embarque)
            'list_item.SubItems(1) = rs!NOMBRE_AGENTE
            rs.MoveNext
      Wend
      rs.Close
      If lv_embarques_4.ListItems.Count > 0 Then
         lv_embarques_4.ListItems(1).Selected = True
         rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_4.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         lv_pedidos.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
               'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
               list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
               'rsaux.Close
               list_item.SubItems(2) = rs!Cliente
               list_item.SubItems(3) = rs!PIEZAS
               list_item.SubItems(5) = rs!orden_pedido
               rs.MoveNext
         Wend
         rs.Close
         If lv_pedidos.ListItems.Count > 11 Then
            Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
         Else
            Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
         End If
         rs.Open "SELECT SUM(PIEZAS) FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
         Me.lbl_cantidad_4 = Format(rs(0).Value, "###,###,##0.00")
         rs.Close
         
         
         
         
      End If
   End If
End Sub

Private Sub cmd_anden_5_Click()
   froracle_asignacion_embarques.Enabled = False
   var_activa_forma_embarques = "froracle_asignacion_embarques"
   var_anden_asignar = 5
   var_anden_global = 5
   frmoracle_embarques.Show 1
   'frmoracle_asignar_pedidos_embarque.Show 1
    
   Me.lv_embarques_5.ListItems.Clear
   var_dia_s = CStr(Day(Date))
   If Len(var_dia_s) = 1 Then
      var_dia_s = "0" + var_dia_s
   End If
   var_mes_s = CStr(Month(Date))
   If Len(var_mes_s) = 1 Then
      var_mes_s = "0" + var_mes_s
   End If
   rs.Open "select distinct embarque as embarque from xxvia_Tb_encabezado_embarques where jaula = 5 and to_char(fecha_inicio,'DD')  = '" + CStr(var_dia_s) + "' and to_char(fecha_inicio,'MM')  = '" + CStr(var_mes_s) + "' and to_char(fecha_inicio,'yyyy')  = '" + CStr(Year(Date)) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_Cadena_embarques = ""
   While Not rs.EOF
         If var_Cadena_embarques = "" Then
            var_Cadena_embarques = CStr(rs!Embarque)
         Else
            var_Cadena_embarques = var_Cadena_embarques + "," + CStr(rs!Embarque)
         End If
         rs.MoveNext
   Wend
   rs.Close
   If var_Cadena_embarques <> "" Then
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open "select distinct EMBARQUE from tb_oracle_pedidos_asignados_embarques where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_embarques_5.ListItems.Add(, , rs!Embarque)
            'list_item.SubItems(1) = rs!NOMBRE_AGENTE
            rs.MoveNext
      Wend
      rs.Close
      If lv_embarques_5.ListItems.Count > 0 Then
         lv_embarques_5.ListItems(1).Selected = True
         rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_5.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         lv_pedidos.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
               'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
               list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
               'rsaux.Close
               list_item.SubItems(2) = rs!Cliente
               list_item.SubItems(3) = rs!PIEZAS
               list_item.SubItems(5) = rs!orden_pedido
               rs.MoveNext
         Wend
         rs.Close
         If lv_pedidos.ListItems.Count > 11 Then
            Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
         Else
            Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
         End If
         rs.Open "SELECT SUM(PIEZAS) FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
         Me.lbl_cantidad_5 = Format(rs(0).Value, "###,###,##0.00")
         rs.Close
         
         
         
         
      End If
   End If
End Sub


Private Sub cmd_anden_6_Click()
   froracle_asignacion_embarques.Enabled = False
   var_activa_forma_embarques = "froracle_asignacion_embarques"
   var_anden_asignar = 6
   var_anden_global = 6
   frmoracle_embarques.Show 1
   'frmoracle_asignar_pedidos_embarque.Show 1
    
   Me.lv_embarques_6.ListItems.Clear
   var_dia_s = CStr(Day(Date))
   If Len(var_dia_s) = 1 Then
      var_dia_s = "0" + var_dia_s
   End If
   var_mes_s = CStr(Month(Date))
   If Len(var_mes_s) = 1 Then
      var_mes_s = "0" + var_mes_s
   End If
   rs.Open "select distinct embarque as embarque from xxvia_Tb_encabezado_embarques where jaula = 6 and to_char(fecha_inicio,'DD')  = '" + CStr(var_dia_s) + "' and to_char(fecha_inicio,'MM')  = '" + CStr(var_mes_s) + "' and to_char(fecha_inicio,'yyyy')  = '" + CStr(Year(Date)) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_Cadena_embarques = ""
   While Not rs.EOF
         If var_Cadena_embarques = "" Then
            var_Cadena_embarques = CStr(rs!Embarque)
         Else
            var_Cadena_embarques = var_Cadena_embarques + "," + CStr(rs!Embarque)
         End If
         rs.MoveNext
   Wend
   rs.Close
   If var_Cadena_embarques <> "" Then
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open "select distinct EMBARQUE from tb_oracle_pedidos_asignados_embarques where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_embarques_6.ListItems.Add(, , rs!Embarque)
            'list_item.SubItems(1) = rs!NOMBRE_AGENTE
            rs.MoveNext
      Wend
      rs.Close
      If lv_embarques_6.ListItems.Count > 0 Then
         lv_embarques_6.ListItems(1).Selected = True
         rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_6.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         lv_pedidos.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
               'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
               list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
               'rsaux.Close
               list_item.SubItems(2) = rs!Cliente
               list_item.SubItems(3) = rs!PIEZAS
               list_item.SubItems(5) = rs!orden_pedido
               rs.MoveNext
         Wend
         rs.Close
         If lv_pedidos.ListItems.Count > 11 Then
            Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
         Else
            Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
         End If
         rs.Open "SELECT SUM(PIEZAS) FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
         Me.lbl_cantidad_1 = Format(rs(0).Value, "###,###,##0.00")
         rs.Close
      End If
   End If
End Sub

Private Sub cmd_anden_7_Click()
   froracle_asignacion_embarques.Enabled = False
   var_activa_forma_embarques = "froracle_asignacion_embarques"
   var_anden_asignar = 7
   var_anden_global = 7
   frmoracle_embarques.Show 1
   'frmoracle_asignar_pedidos_embarque.Show 1
    
   Me.lv_embarques_7.ListItems.Clear
   var_dia_s = CStr(Day(Date))
   If Len(var_dia_s) = 1 Then
      var_dia_s = "0" + var_dia_s
   End If
   var_mes_s = CStr(Month(Date))
   If Len(var_mes_s) = 1 Then
      var_mes_s = "0" + var_mes_s
   End If
   rs.Open "select distinct embarque as embarque from xxvia_Tb_encabezado_embarques where jaula = 7 and to_char(fecha_inicio,'DD')  = '" + CStr(var_dia_s) + "' and to_char(fecha_inicio,'MM')  = '" + CStr(var_mes_s) + "' and to_char(fecha_inicio,'yyyy')  = '" + CStr(Year(Date)) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_Cadena_embarques = ""
   While Not rs.EOF
         If var_Cadena_embarques = "" Then
            var_Cadena_embarques = CStr(rs!Embarque)
         Else
            var_Cadena_embarques = var_Cadena_embarques + "," + CStr(rs!Embarque)
         End If
         rs.MoveNext
   Wend
   rs.Close
   If var_Cadena_embarques <> "" Then
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open "select distinct EMBARQUE from tb_oracle_pedidos_asignados_embarques where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_embarques_7.ListItems.Add(, , rs!Embarque)
            'list_item.SubItems(1) = rs!NOMBRE_AGENTE
            rs.MoveNext
      Wend
      rs.Close
      If lv_embarques_7.ListItems.Count > 0 Then
         lv_embarques_7.ListItems(1).Selected = True
         rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_7.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         lv_pedidos.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
               'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
               list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
               'rsaux.Close
               list_item.SubItems(2) = rs!Cliente
               list_item.SubItems(3) = rs!PIEZAS
               list_item.SubItems(5) = rs!orden_pedido
               rs.MoveNext
         Wend
         rs.Close
         If lv_pedidos.ListItems.Count > 11 Then
            Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
         Else
            Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
         End If
         rs.Open "SELECT SUM(PIEZAS) FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
         Me.lbl_cantidad_7 = Format(rs(0).Value, "###,###,##0.00")
         rs.Close
      End If
   End If
End Sub

Private Sub cmd_anden_8_Click()
   froracle_asignacion_embarques.Enabled = False
   var_activa_forma_embarques = "froracle_asignacion_embarques"
   var_anden_asignar = 8
   var_anden_global = 8
   frmoracle_embarques.Show 1
   'frmoracle_asignar_pedidos_embarque.Show 1
    
   Me.lv_embarques_8.ListItems.Clear
   var_dia_s = CStr(Day(Date))
   If Len(var_dia_s) = 1 Then
      var_dia_s = "0" + var_dia_s
   End If
   var_mes_s = CStr(Month(Date))
   If Len(var_mes_s) = 1 Then
      var_mes_s = "0" + var_mes_s
   End If
   rs.Open "select distinct embarque as embarque from xxvia_Tb_encabezado_embarques where jaula = 8 and to_char(fecha_inicio,'DD')  = '" + CStr(var_dia_s) + "' and to_char(fecha_inicio,'MM')  = '" + CStr(var_mes_s) + "' and to_char(fecha_inicio,'yyyy')  = '" + CStr(Year(Date)) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_Cadena_embarques = ""
   While Not rs.EOF
         If var_Cadena_embarques = "" Then
            var_Cadena_embarques = CStr(rs!Embarque)
         Else
            var_Cadena_embarques = var_Cadena_embarques + "," + CStr(rs!Embarque)
         End If
         rs.MoveNext
   Wend
   rs.Close
   If var_Cadena_embarques <> "" Then
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open "select distinct EMBARQUE from tb_oracle_pedidos_asignados_embarques where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_embarques_8.ListItems.Add(, , rs!Embarque)
            'list_item.SubItems(1) = rs!NOMBRE_AGENTE
            rs.MoveNext
      Wend
      rs.Close
      If lv_embarques_8.ListItems.Count > 0 Then
         lv_embarques_8.ListItems(1).Selected = True
         rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_8.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         lv_pedidos.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
               'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
               list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
               'rsaux.Close
               list_item.SubItems(2) = rs!Cliente
               list_item.SubItems(3) = rs!PIEZAS
               list_item.SubItems(5) = rs!orden_pedido
               rs.MoveNext
         Wend
         rs.Close
         If lv_pedidos.ListItems.Count > 11 Then
            Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
         Else
            Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
         End If
         rs.Open "SELECT SUM(PIEZAS) FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
         Me.lbl_cantidad_8 = Format(rs(0).Value, "###,###,##0.00")
         rs.Close
      End If
   End If
End Sub

Private Sub cmd_anden_9_Click()
   froracle_asignacion_embarques.Enabled = False
   var_activa_forma_embarques = "froracle_asignacion_embarques"
   var_anden_asignar = 9
   var_anden_global = 9
   frmoracle_embarques.Show 1
   'frmoracle_asignar_pedidos_embarque.Show 1
    
   Me.lv_embarques_9.ListItems.Clear
   var_dia_s = CStr(Day(Date))
   If Len(var_dia_s) = 1 Then
      var_dia_s = "0" + var_dia_s
   End If
   var_mes_s = CStr(Month(Date))
   If Len(var_mes_s) = 1 Then
      var_mes_s = "0" + var_mes_s
   End If
   rs.Open "select distinct embarque as embarque from xxvia_Tb_encabezado_embarques where jaula = 9 and to_char(fecha_inicio,'DD')  = '" + CStr(var_dia_s) + "' and to_char(fecha_inicio,'MM')  = '" + CStr(var_mes_s) + "' and to_char(fecha_inicio,'yyyy')  = '" + CStr(Year(Date)) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_Cadena_embarques = ""
   While Not rs.EOF
         If var_Cadena_embarques = "" Then
            var_Cadena_embarques = CStr(rs!Embarque)
         Else
            var_Cadena_embarques = var_Cadena_embarques + "," + CStr(rs!Embarque)
         End If
         rs.MoveNext
   Wend
   rs.Close
   If var_Cadena_embarques <> "" Then
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open "select distinct EMBARQUE from tb_oracle_pedidos_asignados_embarques where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_embarques_9.ListItems.Add(, , rs!Embarque)
            'list_item.SubItems(1) = rs!NOMBRE_AGENTE
            rs.MoveNext
      Wend
      rs.Close
      If lv_embarques_9.ListItems.Count > 0 Then
         lv_embarques_9.ListItems(1).Selected = True
         rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_9.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         lv_pedidos.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
               'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
               list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
               'rsaux.Close
               list_item.SubItems(2) = rs!Cliente
               list_item.SubItems(3) = rs!PIEZAS
               list_item.SubItems(5) = rs!orden_pedido
               rs.MoveNext
         Wend
         rs.Close
         If lv_pedidos.ListItems.Count > 11 Then
            Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
         Else
            Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
         End If
         rs.Open "SELECT SUM(PIEZAS) FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
         Me.lbl_cantidad_9 = Format(rs(0).Value, "###,###,##0.00")
         rs.Close
      End If
   End If
End Sub

Private Sub cmd_fraccion_pedidos_prioridad_Click()
'GoTo x:
      If rs.State = 1 Then
         rs.Close
      End If
      If rsaux.State = 1 Then
         rsaux.Close
      End If
      If rsaux1.State = 1 Then
         rsaux1.Close
      End If
      If rsaux2.State = 1 Then
         rsaux2.Close
      End If
      If rsaux3.State = 1 Then
         rsaux3.Close
      End If
      If rsaux4.State = 1 Then
         rsaux4.Close
      End If
      If rsaux5.State = 1 Then
         rsaux5.Close
      End If
      If rsaux6.State = 1 Then
         rsaux6.Close
      End If
      If rsaux7.State = 1 Then
         rsaux7.Close
      End If
      If rsaux8.State = 1 Then
         rsaux8.Close
      End If
      If rsaux9.State = 1 Then
         rsaux9.Close
      End If
      var_si_pasa_exportacion = 1
      If Me.lv_pedidos.ListItems.Count > 0 Then
         For var_j = 1 To Me.lv_pedidos.ListItems.Count
             Me.lv_pedidos.ListItems.Item(var_j).Selected = True
             var_pedido = Me.lv_pedidos.selectedItem
             
             
             strconsulta = "select * from oe_order_headers_all where order_number = ?"
             With comandoORA
                         .ActiveConnection = cnnoracle_4
                         .CommandType = adCmdText
                         .CommandText = strconsulta
                         Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, CStr(var_pedido))
                         .Parameters.Append parametro
             End With
             Set rs = comandoORA.execute
             Set comandoORA = Nothing
             Set parametro = Nothing
             If Not rs.EOF Then
                var_tipo_pedido = rs!ORDER_TYPE_ID
                If var_tipo_pedido = 1048 Or var_tipo_pedido = 1051 Or var_tipo_pedido = 1464 Or var_tipo_pedido = 2121 Then
                   var_si_pasa_exportacion = 0
                End If
             End If
             rs.Close
         Next var_j
         var_si_pasa_exportacion = 1
         strconsulta = "select embarque, clave, exportaciones from  xxvia_Tb_Encabezado_embarques, xxvia_tb_Transportes where transporte = clave and embarque = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_embarque))
              .Parameters.Append parametro
         End With
         Set rs = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         
         VAR_EXPORTACIONES = IIf(IsNull(rs!EXPORTACIONES), 0, rs!EXPORTACIONES)
         If var_si_pasa_exportacion = 0 Then
            If VAR_EXPORTACIONES = 1 Then
               var_si_pasa_exportacion = 1
            End If
         End If
         
         
      End If
     
     
      If rs.State = 1 Then
         rs.Close
      End If
      If rsaux.State = 1 Then
         rsaux.Close
      End If
      If rsaux1.State = 1 Then
         rsaux1.Close
      End If
      If rsaux2.State = 1 Then
         rsaux2.Close
      End If
      If rsaux3.State = 1 Then
         rsaux3.Close
      End If
      If rsaux4.State = 1 Then
         rsaux4.Close
      End If
      If rsaux5.State = 1 Then
         rsaux5.Close
      End If
      If rsaux6.State = 1 Then
         rsaux6.Close
      End If
      If rsaux7.State = 1 Then
         rsaux7.Close
      End If
      If rsaux8.State = 1 Then
         rsaux8.Close
      End If
      If rsaux9.State = 1 Then
         rsaux9.Close
      End If
     
     
     
     
     
      If Me.lv_pedidos.ListItems.Count > 0 Then
         If var_si_pasa_exportacion = 1 Then
         var_si = MsgBox("Desea imprimir las ordenes de surtido?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            If rs.State = 1 Then
               rs.Close
            End If
            var_Cadena_pedidos = ""
            For var_j = 1 To Me.lv_pedidos.ListItems.Count
                Me.lv_pedidos.ListItems.Item(var_j).Selected = True
                If Me.lv_pedidos.selectedItem <> "10000000" Then
                   If var_Cadena_pedidos = "" Then
                      var_Cadena_pedidos = Me.lv_pedidos.selectedItem
                   Else
                      var_Cadena_pedidos = var_Cadena_pedidos + "," + Me.lv_pedidos.selectedItem
                   End If
                End If
            Next var_j
            'var_cadena_pedidos = "105208"
            rs.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_cadena = "SELECT to_char(a.LAST_UPDATE_DATE,'day') DIA_SEMANA, CAT.LINEA, HCSU.SITE_USE_ID, HCAS.CUST_ACCOUNT_ID, OHA.SHIP_TO_ORG_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id, a.LAST_UPDATE_DATE,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,c.description,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, a.source_header_type_name, oha.source_document_id, C.ATTRIBUTE2, oha.attribute8, oha.attribute9, j.NAME as nombre_ruta, j.salesrep_id as clave_ruta from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, xxvia_vw_articulos_cat cat, XXVIA_VENDEDORES j Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND "
            var_cadena = var_cadena + " to_number(source_header_number)  IN (" + var_Cadena_pedidos + ")"
            var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID "
            var_cadena = var_cadena + " AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND released_status = 'Y' AND A.inventory_item_id  = cat.item_id AND A.ORGANIZATION_ID = Cat.organization_id and oha.salesrep_id = j.salesrep_id "
            rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
'--------------------------
            If Not rs.EOF Then
               cnn.BeginTrans
               rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM tb_Temp_oracle_orden_surtido", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
               Else
                  var_consecutivo = 1
               End If
               rsaux.Close
               rsaux1.Open "insert into tb_Temp_oracle_orden_surtido(inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
               cnn.CommitTrans
               While Not rs.EOF
                     var_establecimiento = rs!SHIP_TO_ORG_ID
                     var_clave_cliente = rs!site_use_id
                     rsaux.Open "SELECT csu.site_use_id AS VCHA_ESB_ESTABLECIMIENTO_ID, ps.party_site_number, lo.address1 AS VCHA_eSB_NOMBRE FROM hz_cust_site_uses_all csu, hz_cust_acct_sites_all cas, hz_party_sites ps, hz_locations lo Where csu.cust_acct_site_id = cas.cust_acct_site_id AND cas.party_site_id = ps.party_site_id AND ps.location_id = lo.location_id AND csu.site_use_code = 'SHIP_TO' AND csu.LOCATION = ps.party_site_number and csu.site_use_id = " + CStr(var_establecimiento), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        VAR_NOMBRE_ESTABLECIMIENTO = IIf(IsNull(rsaux!party_site_number), "", rsaux!party_site_number) + " " + IIf(IsNull(rsaux!vcha_esb_nombre), "", rsaux!vcha_esb_nombre)
                     Else
                        VAR_NOMBRE_ESTABLECIMIENTO = ""
                     End If
                     rsaux.Close
                     
                     rsaux.Open "SELECT csu.site_use_id AS VCHA_ESB_ESTABLECIMIENTO_ID, ps.party_site_number, lo.address1 AS VCHA_eSB_NOMBRE FROM hz_cust_site_uses_all csu, hz_cust_acct_sites_all cas, hz_party_sites ps, hz_locations lo Where csu.cust_acct_site_id = cas.cust_acct_site_id AND cas.party_site_id = ps.party_site_id AND ps.location_id = lo.location_id AND csu.site_use_code = 'BILL_TO' AND csu.LOCATION = ps.party_site_number and csu.site_use_id = " + CStr(var_clave_cliente), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        VAR_CLAVE_CLIENTE_BCP = IIf(IsNull(rsaux!party_site_number), "", rsaux!party_site_number)
                     Else
                        VAR_CLAVE_CLIENTE_BCP = ""
                     End If
                     rsaux.Close
                     
                     
                     
                     var_dia = CStr(Day(CDate(rs!LAST_UPDATE_DATE)))
                     var_mes = CStr(Month(CDate(rs!LAST_UPDATE_DATE)))
                     var_año = CStr(Year(CDate(rs!LAST_UPDATE_DATE)))
                     If Len(Trim(var_dia)) = 1 Then
                        var_dia = "0" + var_dia
                     End If
                     If Len(Trim(var_mes)) = 1 Then
                        var_mes = "0" + var_mes
                     End If
                     var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                     rsaux1.Open "select * from tb_oracle_multiplos where segment1 = '" + IIf(IsNull(rs!SEGMENT1), "", rs!SEGMENT1) + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux1.EOF Then
                        VAR_MULTIPLO = IIf(IsNull(rsaux1!MULTIPLO), 1, rsaux1!MULTIPLO)
                     Else
                        VAR_MULTIPLO = 1
                     End If
                     rsaux1.Close
'''''
                     If rs!source_header_type_name = "VIA_PEDIDO_INTERNO" Then
                        rsaux1.Open "SELECT * FROM TB_ORACLE_ARTICULOS_MOTOR_LOGISTICO WHERE CODIGO = '" + rs!SEGMENT1 + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux1.EOF Then
                           strconsulta = "SELECT secondary_inventory_name, A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rs!source_document_id)
                                .Parameters.Append parametro
                           End With
                           Set rsaux8 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           If rsaux8.EOF Then
                              var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                           Else
                              var_almacen = rsaux8!secondary_inventory_name
                              rsaux9.Open "SELECT * FROM TB_ORACLE_UBICACIONES_MOTOR_LOGISTICO WHERE CLAVE = '" + var_almacen + "' AND CODIGO = '" + rs!SEGMENT1 + "'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux9.EOF Then
                                 var_ubicacion = ""
                                 If Me.cmb_dia.Text = "Lunes" Then
                                    var_ubicacion = rsaux9!ubicacion_1
                                 End If
                                 If Me.cmb_dia.Text = "Martes" Then
                                    var_ubicacion = rsaux9!ubicacion_2
                                 End If
                                 If Me.cmb_dia.Text = "Miercoles" Then
                                    var_ubicacion = rsaux9!ubicacion_3
                                 End If
                                 If Me.cmb_dia.Text = "Jueves" Then
                                    var_ubicacion = rsaux9!ubicacion_4
                                 End If
                                 If Me.cmb_dia.Text = "Viernes" Then
                                    var_ubicacion = rsaux9!ubicacion_5
                                 End If
                                 If Me.cmb_dia.Text = "Sabado" Then
                                    var_ubicacion = rsaux9!ubicacion_6
                                 End If
                                 If IIf(IsNull(var_ubicacion), "", var_ubicacion) = "" Then
                                    var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                                 End If
                              Else
                                 var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                              End If
                              rsaux9.Close
                           End If
                           rsaux8.Close
                        Else
                           var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                        End If
                        rsaux1.Close
                     Else
                        var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                     End If
                     
                     
'''''
                     var_cadena = "insert into tb_Temp_oracle_orden_surtido(INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER, DELIVERY_ID, DELIVERY_DETAIL_ID, ORGANIZATION_ID, SUBINVENTORY, DELIVERY_LINE_ID, INVENTORY_ITEM_ID, ITEM_DESCRIPTION, SOURCE_LINE_NUMBER, src_requested_quantity, RELEASED_STATUS, CUSTOMER_NAME, SEGMENT1, COLLECTOR_ID, NAME, date_requested, UBICACION, ESTABLECIMIENTO, NOMBRE_ESTABLECIMIENTO, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, SITE_USE_ID, attribute8, attribute9, LINEA, RUTA, NOMBRE_RUTA,MULTIPLO)  values "
                     var_cadena = var_cadena + " (" + CStr(var_consecutivo) + ", '" + IIf(IsNull(rs!source_header_number), "", rs!source_header_number) + "', " + CStr(IIf(IsNull(rs!delivery_id), 0, rs!delivery_id)) + ", " + CStr(IIf(IsNull(rs!delivery_detail_id), 0, rs!delivery_detail_id)) + ", " + CStr(IIf(IsNull(rs!organization_id), 0, rs!organization_id)) + ", '" + IIf(IsNull(rs!subinventory), "", rs!subinventory) + "', " + CStr(IIf(IsNull(rs!delivery_line_id), 0, rs!delivery_line_id)) + ", " + CStr(IIf(IsNull(rs!inventory_item_id), "", rs!inventory_item_id)) + ", '" + IIf(IsNull(rs!Description), "", rs!Description) + "', '" + IIf(IsNull(rs!SOURCE_LINE_NUMBER), "", rs!SOURCE_LINE_NUMBER) + "', " + CStr(IIf(IsNull(rs!requested_quantity), 0, rs!requested_quantity)) + ", '" + IIf(IsNull(rs!released_status), "", rs!released_status) + "', '" + VAR_CLAVE_CLIENTE_BCP + " " + Replace(IIf(IsNull(rs!customer_name), "", rs!customer_name), "'", " ") + "', '" + IIf(IsNull(rs!SEGMENT1), "", rs!SEGMENT1) + "'"
                     'var_cadena = var_cadena + ", 0,''," + var_fecha + ",'" + IIf(IsNull(rs!ATTRIBUTE2), "", rs!ATTRIBUTE2) + "','" + CStr(VAR_ESTABLECIMIENTO) + "','" + Replace(VAR_NOMBRE_ESTABLECIMIENTO, "'", " ") + "'," + CStr(rs!CUST_ACCOUNT_ID) + ",'" + rs!source_header_type_name + "','" + CStr(IIf(IsNull(rs!source_document_id), "", rs!source_document_id)) + "'," + CStr(IIf(IsNull(rs!SITE_USE_ID), 0, rs!SITE_USE_ID)) + ",'" + IIf(IsNull(rs!attribute8), "", rs!attribute8) + "','" + IIf(IsNull(rs!ATTRIBUTE9), "", rs!ATTRIBUTE9) + "','" + IIf(IsNull(rs!linea), "", rs!linea) + "','" + CStr(rs!CLAVE_RUTA) + "','" + IIf(IsNull(rs!nombre_ruta), "", rs!nombre_ruta) + "'," + CStr(VAR_MULTIPLO) + ")"
                     var_cadena = var_cadena + ", 0,''," + var_fecha + ",'" + var_ubicacion + "','" + CStr(var_establecimiento) + "','" + Replace(VAR_NOMBRE_ESTABLECIMIENTO, "'", " ") + "'," + CStr(rs!CUST_ACCOUNT_ID) + ",'" + rs!source_header_type_name + "','" + CStr(IIf(IsNull(rs!source_document_id), "", rs!source_document_id)) + "'," + CStr(IIf(IsNull(rs!site_use_id), 0, rs!site_use_id)) + ",'" + IIf(IsNull(rs!attribute8), "", rs!attribute8) + "','" + IIf(IsNull(rs!ATTRIBUTE9), "", rs!ATTRIBUTE9) + "','" + IIf(IsNull(rs!Linea), "", rs!Linea) + "','" + CStr(rs!CLAVE_RUTA) + "','" + IIf(IsNull(rs!nombre_ruta), "", rs!nombre_ruta) + "'," + CStr(VAR_MULTIPLO) + ")"
                     rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               rsaux1.Open "delete from tb_Temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number is null", cnn, adOpenDynamic, adLockOptimistic
               If rsaux1.State = 1 Then
                  rsaux1.Close
               End If
               
               var_cadena_pedidos_diferencias = ""
               rsaux1.Open "select source_header_number, sum(src_requested_quantity) as cantidad from tb_Temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo) + " group by source_header_number", cnn, adOpenDynamic, adLockOptimistic
               While Not rsaux1.EOF
                     strconsulta = "select sum(requested_quantity)  as cantidad from WSH_DELIVERABLES_V where source_header_number = ? AND RELEASED_STATUS = 'Y'"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(rsaux1!source_header_number))
                          .Parameters.Append parametro
                     End With
                     Set rsaux10 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     If rsaux1!cantidad <> rsaux10!cantidad Then
                        If var_cadena_pedidos_diferencias = "" Then
                           var_cadena_pedidos_diferencias = CStr(rsaux1!source_header_number)
                        Else
                           var_cadena_pedidos_diferencias = var_cadena_pedidos_diferencias + ", " + CStr(rsaux1!source_header_number)
                        End If
                     End If
                     rsaux10.Close
                     rsaux1.MoveNext
               Wend
               rsaux1.Close
               
               If var_cadena_pedidos_diferencias = "" Then
                  rsaux1.Open "select min(DATE_REQUESTED) as DATE_REQUESTED, source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME, RUTA, NOMBRE_RUTA, SUM(SRC_REQUESTED_QUANTITY) AS CANTIDAD from tb_Temp_oracle_orden_surtido where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " group by source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME, RUTA, NOMBRE_RUTA", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        If rsaux1!source_header_type_name = "VIA_PEDIDO_INTERNO" Or rsaux1!source_header_type_name = "TEX_PEDIDO_INTERNO" Then
                           If var_pedido_tienda = 0 Then
                              If rsaux2.State = 1 Then
                                 rsaux2.Close
                              End If
                              rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(rsaux1!source_document_id) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux2.EOF Then
                                 If rsaux4.State = 1 Then
                                    rsaux4.Close
                                 End If
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                                 rsaux4.Close
                              Else
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                                 rsaux4.Close
                              End If
                              rsaux2.Close
                           Else
                              rsaux2.Open "select a.attribute8, B.description from oe_order_headers_all a, MTL_SECONDARY_INVENTORIES b where order_number = " + CStr(rsaux1!source_header_number) + " and a.attribute8 = b.secondary_inventory_name", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux2.EOF Then
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                                 rsaux4.Close
                              Else
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                                 rsaux4.Close
                              End If
                              rsaux2.Close
                           End If
                        Else
                           rsaux6.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                           VAR_PROVEEDOR = rsaux6!collector_id
                           VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux6!Name), "", rsaux6!Name)
                           rsaux6.Close
                        End If
                        var_año_str = CStr(Year(rsaux1!DATE_REQUESTED))
                        If Len(var_año_str) < 2 Then
                           var_año_str = "20" + var_año_str
                        End If
                        var_mes_str = CStr(Month(rsaux1!DATE_REQUESTED))
                        If Len(var_mes_str) < 2 Then
                           var_mes_str = "0" + var_mes_str
                        End If
                        var_dia_str = CStr(Day(rsaux1!DATE_REQUESTED))
                        If Len(var_dia_str) < 2 Then
                           var_dia_str = "0" + var_dia_str
                        End If
                        var_hora_str = CStr(Hour(rsaux1!DATE_REQUESTED))
                        If Len(var_hora_str) < 2 Then
                           var_hora_str = "0" + var_hora_str
                        End If
                        VAR_MINUTO_STR = CStr(Minute(rsaux1!DATE_REQUESTED))
                        If Len(VAR_MINUTO_STR) < 2 Then
                           VAR_MINUTO_STR = "0" + VAR_MINUTO_STR
                        End If
                        VAR_SEGUNDO_STR = CStr(Second(rsaux1!DATE_REQUESTED))
                        If Len(VAR_SEGUNDO_STR) < 2 Then
                           VAR_SEGUNDO_STR = "0" + VAR_SEGUNDO_STR
                        End If
                        var_fecha_pedido = var_año_str + "-" + var_mes_str + "-" + var_dia_str + " " + var_hora_str + ":" + VAR_MINUTO_STR + ":" + VAR_SEGUNDO_STR
                        rsaux6.Open "update tb_Temp_oracle_orden_surtido set COLLECTOR_ID = " + CStr(VAR_PROVEEDOR) + ", NAME = '" + VAR_NOMBRE_PROVEEDOR + "' where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " and CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and source_header_number = " + CStr(rsaux1!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                        rsaux6.Open "SELECT * FROM XXVIA_TB_ORDENES_GRAFICA WHERE ORGANIZACION = " + var_unidad_organizacional + " AND PEDIDO = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If rsaux6.EOF Then
                           rsaux5.Open "INSERT INTO XXVIA_TB_ORDENES_GRAFICA (ORGANIZACION, PEDIDO, PROVEEDOR_ID, NOMBRE_PROVEEDOR, CLIENTE, ESTABLECIMIENTO, CANTIDAD, FECHA, TIPO_PEDIDO, CANTIDAD_LEIDA, FECHA_PEDIDO, RUTA, NOMBRE_RUTA) VALUES (" + var_unidad_organizacional + ", " + CStr(rsaux1!source_header_number) + "," + CStr(VAR_PROVEEDOR) + ",'" + VAR_NOMBRE_PROVEEDOR + "','" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "'," + CStr(rsaux1!cantidad) + ", '" + CStr(rsaux1!DATE_REQUESTED) + "','" + rsaux1!source_header_type_name + "',0, TO_DATE('" + var_fecha_pedido + "','YYYY-MM-DD HH24:MI:SS'),'" + rsaux1!ruta + "', '" + rsaux1!nombre_ruta + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        Else
                           rsaux5.Open "UPDATE XXVIA_TB_ORDENES_GRAFICA SET RUTA = '" + rsaux1!ruta + "', NOMBRE_RUTA = '" + rsaux1!nombre_ruta + "' WHERE PEDIDO = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        End If
                        rsaux6.Close
                        rsaux6.Open "select * from tb_oracle_pedidos_asignados_embarques where pedido = " + CStr(rsaux1!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                        If rsaux6.EOF Then
                           rsaux5.Open "insert into tb_oracle_pedidos_asignados_embarques (AGENTE, NOMBRE_AGENTE, PEDIDO, CLIENTE, PIEZAS, embarque, dia,  mes, AÑO, ORGANIZACION) values (" + CStr(VAR_PROVEEDOR) + ",'" + VAR_NOMBRE_PROVEEDOR + "'," + CStr(rsaux1!source_header_number) + ",'" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "', " + CStr(rsaux1!cantidad) + ",0," + var_dia_str + "," + var_mes_str + "," + var_año_str + "," + CStr(var_unidad_organizacional) + ")", cnn, adOpenDynamic, adLockOptimistic
                        End If
                        rsaux6.Close
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  rsaux1.Open "select min(DATE_REQUESTED) as DATE_REQUESTED, source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME, SUM(SRC_REQUESTED_QUANTITY) AS CANTIDAD from tb_Temp_oracle_orden_surtido where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " AND (LINEA NOT IN ('CATALOGOS','CATALOGO','POP') OR LINEA IS NULL) group by source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME"
                  While Not rsaux1.EOF
                        rsaux5.Open "UPDATE XXVIA_TB_ORDENES_GRAFICA SET CANTIDAD_SIN_CATALOGOS = " + CStr(IIf(IsNull(rsaux1!cantidad), 0, rsaux1!cantidad)) + " WHERE PEDIDO = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  rsaux1.Open "DELETE from tb_Temp_oracle_orden_surtido_aux_2", cnn, adOpenDynamic, adLockOptimistic
                  rsaux1.Open "SELECT * FROM tb_Temp_oracle_orden_surtido where inte_tem_consecutivo =  " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        If rsaux1!Linea = "CATALOGOS" Or rsaux1!Linea = "CATALOGO" Or rsaux1!Linea = "POP" Or rsaux1!Linea = "EMPAQUE" Then
                           var_dia = CStr(Day(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                           var_mes = CStr(Month(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                           var_año = CStr(Year(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                           If Len(Trim(var_dia)) = 1 Then
                              var_dia = "0" + var_dia
                           End If
                           If Len(Trim(var_mes)) = 1 Then
                              var_mes = "0" + var_mes
                           End If
                           var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                           var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_2 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                           var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,MULTIPLO) "
                           var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux1!source_header_number + "', " + CStr(rsaux1!delivery_id) + "," + CStr(rsaux1!delivery_detail_id) + ", " + CStr(rsaux1!organization_id) + ",'" + IIf(IsNull(rsaux1!subinventory), "", rsaux1!subinventory) + "', " + CStr(rsaux1!delivery_line_id) + "," + CStr(rsaux1!inventory_item_id) + ",'" + rsaux1!item_description + "','" + CStr(rsaux1!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux1!src_requested_quantity) + ",'" + rsaux1!released_status + "',"
                           var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + rsaux1!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux1!collector_id), 0, rsaux1!collector_id)) + ",'" + IIf(IsNull(rsaux1!Name), "", rsaux1!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux1!ubicacion), "", rsaux1!ubicacion) + "','" + IIf(IsNull(rsaux1!establecimiento), "", rsaux1!establecimiento) + "','" + IIf(IsNull(rsaux1!nombre_Establecimiento), "", rsaux1!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux1!ORDENES), "", rsaux1!ORDENES) + "',"
                           var_cadena = var_cadena + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux1!source_header_type_name), "", rsaux1!source_header_type_name) + "', '" + IIf(IsNull(rsaux1!source_document_id), "", rsaux1!source_document_id) + "','" + IIf(IsNull(rsaux1!pais), "", rsaux1!pais) + "','" + IIf(IsNull(rsaux1!estado), "", rsaux1!estado) + "', '" + IIf(IsNull(rsaux1!municipio), "", rsaux1!municipio) + "', '" + IIf(IsNull(rsaux1!ciudad), "", rsaux1!ciudad) + "', '" + IIf(IsNull(rsaux1!colonia), "", rsaux1!colonia) + "','" + IIf(IsNull(rsaux1!DIRECCION), "", rsaux1!DIRECCION) + "', '" + IIf(IsNull(rsaux1!cp), "", rsaux1!cp) + "',"
                           var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!site_use_id), "", rsaux1!site_use_id) + "','" + IIf(IsNull(rsaux1!paqueteria), "", rsaux1!paqueteria) + "','" + IIf(IsNull(rsaux1!attribute8), "", rsaux1!attribute8) + "','" + IIf(IsNull(rsaux1!ATTRIBUTE9), "", rsaux1!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux1!TELEFONO), "", rsaux1!TELEFONO) + "','" + IIf(IsNull(rsaux1!Linea), "", rsaux1!Linea) + "','" + CStr(IIf(IsNull(rsaux1!ruta), "", rsaux1!ruta)) + "','" + IIf(IsNull(rsaux1!nombre_ruta), "", rsaux1!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux1!ORDEN_SURTIDO), 0, rsaux1!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux1!Embarque), 0, rsaux1!Embarque)) + ", '" + IIf(IsNull(rsaux1!estacion), "", rsaux1!estacion) + "'," + CStr(IIf(IsNull(rsaux1!MULTIPLO), "", rsaux1!MULTIPLO)) + ")"
                           rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        Else
                           var_cantidad_total = rsaux1!src_requested_quantity
                           If rsaux1!MULTIPLO > 1 Then
                              While var_cantidad_total > 0
                                    If var_cantidad_total < rsaux1!MULTIPLO Then
                                       var_cantidad = var_cantidad_total
                                    Else
                                       var_cantidad = rsaux1!MULTIPLO
                                    End If
                                    
                                    var_dia = CStr(Day(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                    var_mes = CStr(Month(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                    var_año = CStr(Year(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                    If Len(Trim(var_dia)) = 1 Then
                                       var_dia = "0" + var_dia
                                    End If
                                    If Len(Trim(var_mes)) = 1 Then
                                      var_mes = "0" + var_mes
                                    End If
                                    
                                    var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                    var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_2 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                    var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,MULTIPLO, PASILLO) "
                                    var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux1!source_header_number + "', " + CStr(rsaux1!delivery_id) + "," + CStr(rsaux1!delivery_detail_id) + ", " + CStr(rsaux1!organization_id) + ",'" + IIf(IsNull(rsaux1!subinventory), "", rsaux1!subinventory) + "', " + CStr(rsaux1!delivery_line_id) + "," + CStr(rsaux1!inventory_item_id) + ",'" + rsaux1!item_description + "','" + CStr(rsaux1!SOURCE_LINE_NUMBER) + "'," + CStr(var_cantidad) + ",'" + rsaux1!released_status + "',"
                                    var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + rsaux1!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux1!collector_id), 0, rsaux1!collector_id)) + ",'" + IIf(IsNull(rsaux1!Name), "", rsaux1!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux1!ubicacion), "", rsaux1!ubicacion) + "','" + IIf(IsNull(rsaux1!establecimiento), "", rsaux1!establecimiento) + "','" + IIf(IsNull(rsaux1!nombre_Establecimiento), "", rsaux1!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux1!ORDENES), "", rsaux1!ORDENES) + "',"
                                    var_cadena = var_cadena + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux1!source_header_type_name), "", rsaux1!source_header_type_name) + "', '" + IIf(IsNull(rsaux1!source_document_id), "", rsaux1!source_document_id) + "','" + IIf(IsNull(rsaux1!pais), "", rsaux1!pais) + "','" + IIf(IsNull(rsaux1!estado), "", rsaux1!estado) + "', '" + IIf(IsNull(rsaux1!municipio), "", rsaux1!municipio) + "', '" + IIf(IsNull(rsaux1!ciudad), "", rsaux1!ciudad) + "', '" + IIf(IsNull(rsaux1!colonia), "", rsaux1!colonia) + "','" + IIf(IsNull(rsaux1!DIRECCION), "", rsaux1!DIRECCION) + "', '" + IIf(IsNull(rsaux1!cp), "", rsaux1!cp) + "',"
                                    var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!site_use_id), "", rsaux1!site_use_id) + "','" + IIf(IsNull(rsaux1!paqueteria), "", rsaux1!paqueteria) + "','" + IIf(IsNull(rsaux1!attribute8), "", rsaux1!attribute8) + "','" + IIf(IsNull(rsaux1!ATTRIBUTE9), "", rsaux1!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux1!TELEFONO), "", rsaux1!TELEFONO) + "','" + IIf(IsNull(rsaux1!Linea), "", rsaux1!Linea) + "','" + CStr(IIf(IsNull(rsaux1!ruta), "", rsaux1!ruta)) + "','" + IIf(IsNull(rsaux1!nombre_ruta), "", rsaux1!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux1!ORDEN_SURTIDO), 0, rsaux1!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux1!Embarque), 0, rsaux1!Embarque)) + ", '" + IIf(IsNull(rsaux1!estacion), "", rsaux1!estacion) + "'," + CStr(IIf(IsNull(rsaux1!MULTIPLO), "", rsaux1!MULTIPLO)) + ",'" + IIf(IsNull(rsaux1!pasillo), "", rsaux1!pasillo) + "')"
                                    rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                    var_cantidad_total = var_cantidad_total - rsaux1!MULTIPLO
                              Wend
                           Else
                              var_cantidad = var_cantidad_total
                              While var_cantidad > 0
                                    var_dia = CStr(Day(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                    var_mes = CStr(Month(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                    var_año = CStr(Year(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                    If Len(Trim(var_dia)) = 1 Then
                                       var_dia = "0" + var_dia
                                    End If
                                    If Len(Trim(var_mes)) = 1 Then
                                      var_mes = "0" + var_mes
                                    End If
                                    var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                    var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_2 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                    var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,MULTIPLO,PASILLO) "
                                    var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux1!source_header_number + "', " + CStr(rsaux1!delivery_id) + "," + CStr(rsaux1!delivery_detail_id) + ", " + CStr(rsaux1!organization_id) + ",'" + IIf(IsNull(rsaux1!subinventory), "", rsaux1!subinventory) + "', " + CStr(rsaux1!delivery_line_id) + "," + CStr(rsaux1!inventory_item_id) + ",'" + rsaux1!item_description + "','" + CStr(rsaux1!SOURCE_LINE_NUMBER) + "'," + CStr(1) + ",'" + rsaux1!released_status + "',"
                                    var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + rsaux1!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux1!collector_id), 0, rsaux1!collector_id)) + ",'" + IIf(IsNull(rsaux1!Name), "", rsaux1!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux1!ubicacion), "", rsaux1!ubicacion) + "','" + IIf(IsNull(rsaux1!establecimiento), "", rsaux1!establecimiento) + "','" + IIf(IsNull(rsaux1!nombre_Establecimiento), "", rsaux1!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux1!ORDENES), "", rsaux1!ORDENES) + "',"
                                    var_cadena = var_cadena + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux1!source_header_type_name), "", rsaux1!source_header_type_name) + "', '" + IIf(IsNull(rsaux1!source_document_id), "", rsaux1!source_document_id) + "','" + IIf(IsNull(rsaux1!pais), "", rsaux1!pais) + "','" + IIf(IsNull(rsaux1!estado), "", rsaux1!estado) + "', '" + IIf(IsNull(rsaux1!municipio), "", rsaux1!municipio) + "', '" + IIf(IsNull(rsaux1!ciudad), "", rsaux1!ciudad) + "', '" + IIf(IsNull(rsaux1!colonia), "", rsaux1!colonia) + "','" + IIf(IsNull(rsaux1!DIRECCION), "", rsaux1!DIRECCION) + "', '" + IIf(IsNull(rsaux1!cp), "", rsaux1!cp) + "',"
                                    var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!site_use_id), "", rsaux1!site_use_id) + "','" + IIf(IsNull(rsaux1!paqueteria), "", rsaux1!paqueteria) + "','" + IIf(IsNull(rsaux1!attribute8), "", rsaux1!attribute8) + "','" + IIf(IsNull(rsaux1!ATTRIBUTE9), "", rsaux1!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux1!TELEFONO), "", rsaux1!TELEFONO) + "','" + IIf(IsNull(rsaux1!Linea), "", rsaux1!Linea) + "','" + CStr(IIf(IsNull(rsaux1!ruta), "", rsaux1!ruta)) + "','" + IIf(IsNull(rsaux1!nombre_ruta), "", rsaux1!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux1!ORDEN_SURTIDO), 0, rsaux1!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux1!Embarque), 0, rsaux1!Embarque)) + ", '" + IIf(IsNull(rsaux1!estacion), "", rsaux1!estacion) + "'," + CStr(IIf(IsNull(rsaux1!MULTIPLO), "", rsaux1!MULTIPLO)) + ",'" + IIf(IsNull(rsaux1!pasillo), "", rsaux1!pasillo) + "')"
                                    rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                    var_cantidad = var_cantidad - 1
                              Wend
                           End If
                        End If
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  
                  
                  
                  
                  rsaux1.Open "delete from tb_Temp_oracle_orden_surtido_aux_1", cnn, adOpenDynamic, adLockOptimistic
                  var_consecutivo_general = var_consecutivo
                  
Call PASILLOS_2
                  'rsaux1.Open "select distinct source_header_number, ORDEN_SURTIDO from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " order by ORDEN_SURTIDO", cnn, adOpenDynamic, adLockOptimistic

                  
                 ' MsgBox var_consecutivo
                  rsaux1.Open "insert TB_TEMP_ORACLE_ORDEN_SURTIDO (inte_tem_consecutivo, segment1) values (" + CStr(var_consecutivo) + ",'---------')", cnn, adOpenDynamic, adLockOptimistic
                  rsaux1.Open "delete from TB_TEMP_ORACLE_ORDEN_SURTIDO where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and segment1 <> '---------'", cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.Open "insert into TB_TEMP_ORACLE_ORDEN_SURTIDO select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.Open "delete from TB_TEMP_ORACLE_ORDEN_SURTIDO where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and segment1 = '---------'", cnn, adOpenDynamic, adLockOptimistic
                  var_consecutivo_general = var_consecutivo
                  Call crea_tablas
                  If rsaux.State = 1 Then
                     rsaux.Close
                  End If
                  rsaux.Open "select distinct a.source_header_number from tb_Temp_oracle_orden_surtido_aux_1 a, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES  b where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and a.source_header_number = b.pedido", cnn, adOpenDynamic, adLockOptimistic
                  var_cadena_pedidos_diferencias = ""
                  While Not rsaux.EOF
                        strconsulta = "select sum(requested_quantity)  as cantidad from WSH_DELIVERABLES_V where source_header_number = ? AND RELEASED_STATUS = 'Y'"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(rsaux!source_header_number))
                             .Parameters.Append parametro
                        End With
                        Set rsaux10 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        
                        strconsulta = "SELECT SUM(SRC_REQUESTED_QUANTITY) AS CANTIDAD FROM XXVIA_TB_PEDIDOS_DIVIDIDOS WHERE SOURCE_HEADER_NUMBER = ?"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(rsaux!source_header_number))
                             .Parameters.Append parametro
                        End With
                        Set rsaux11 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                     
                     
                        If IIf(IsNull(rsaux11!cantidad), 0, rsaux11!cantidad) <> rsaux10!cantidad Then
                           If var_cadena_pedidos_diferencias = "" Then
                              var_cadena_pedidos_diferencias = CStr(rsaux!source_header_number)
                           Else
                              var_cadena_pedidos_diferencias = var_cadena_pedidos_diferencias + ", " + CStr(rsaux!source_header_number)
                           End If
                        End If
                        rsaux10.Close
                        rsaux11.Close
                        rsaux.MoveNext
                  Wend
                  rsaux.Close
                  If var_cadena_pedidos_diferencias = "" Then
                     If var_imprime_pedidos = 1 Then
                        ' orden
x:
                        'var_consecutivo = 1360
                        'CHECAR QUERY
                        rsaux.Open "select distinct a.grupo from tb_Temp_oracle_orden_surtido_aux_1 a, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES  b where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and a.source_header_number = b.pedido order by a.grupo", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           While Not rsaux.EOF
                                 x = 1
                                 If x = 1 Then
                                    
                                    'strconsulta = "select shipping_method_code, packing_instructions from oe_order_headers_all where order_number = ?"
                                    'With comandoORA
                                    '     .ActiveConnection = cnnoracle_4
                                    '     .CommandType = adCmdText
                                    '     .CommandText = strconsulta
                                    '     Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rsaux!source_header_number))
                                    '     .Parameters.Append parametro
                                    'End With
                                    'Set rs = comandoORA.execute
                                    'Set comandoORA = Nothing
                                    'Set parametro = Nothing
                                    
                                    var_paqueteria = ""
                                    'If Not rs.EOF Then
                                    '   VAR_COMENTARIOS = IIf(IsNull(rs!packing_instructions), "", rs!packing_instructions)
                                    '   var_tipo_metodo = IIf(IsNull(rs(0).Value), "", rs(0).Value)
                                    '   If var_tipo_metodo <> "" Then
                                    '
                                    '      strconsulta = "SELECT description FROM fnd_lookup_values where lookup_type = 'SHIP_METHOD' and lookup_code = ? AND LANGUAGE = 'ESA'"
                                    '      With comandoORA
                                    '           .ActiveConnection = cnno-racle_4
                                    '           .CommandType = adCmdText
                                    '           .CommandText = strconsulta
                                    '           Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_tipo_metodo)
                                    '           .Parameters.Append parametro
                                    '      End With
                                    '      Set rsaux1 = comandoORA.execute
                                    '      Set comandoORA = Nothing
                                    '      Set parametro = Nothing
                                          
                                    '      If Not rsaux1.EOF Then
                                    '         var_paqueteria = IIf(IsNull(rsaux1(0).Value), "", rsaux1(0).Value)
                                    '      End If
                                    '      rsaux1.Close
                                    '   End If
                                    'End If
                                    'rs.Close
                                    
                                    VAR_ZZ = 0
                                    If VAR_ZZ = 1 Then
                                       strconsulta = "SELECT  HPS.party_site_id as tel, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1, OHA.packing_instructions from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.ship_TO_ORG_ID AND oha.order_number = ? AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                                       With comandoORA
                                            .ActiveConnection = cnnoracle_4
                                            .CommandType = adCmdText
                                            .CommandText = strconsulta
                                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rsaux!source_header_number))
                                            .Parameters.Append parametro
                                       End With
                                       Set rsaux6 = comandoORA.execute
                                       Set comandoORA = Nothing
                                       Set parametro = Nothing
                                       
                                       If Not rsaux6.EOF Then
                                       
                                          
                                          strconsulta = "SELECT  HPS.party_site_id as tel, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1, OHA.packing_instructions from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.ship_TO_ORG_ID AND oha.order_number = ? AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                                          With comandoORA
                                               .ActiveConnection = cnnoracle_4
                                               .CommandType = adCmdText
                                               .CommandText = strconsulta
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, IIf(IsNull(rsaux!source_header_number), "", rsaux!source_header_number))
                                               .Parameters.Append parametro
                                          End With
                                          Set rsaux5 = comandoORA.execute
                                          Set comandoORA = Nothing
                                          Set parametro = Nothing
                                          
                                          'rsaux5.Open "SELECT  HPS.party_site_id as tel, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1, OHA.packing_instructions from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.ship_TO_ORG_ID AND oha.order_number = '" + IIf(IsNull(rsaux!source_header_number), "", rsaux!source_header_number) + "' AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                          If Not rsaux5.EOF Then
                                             var_nombre = IIf(IsNull(rsaux5!customer_name), "", rsaux5!customer_name)
                                             var_tel = IIf(IsNull(rsaux5!tel), 0, rsaux5!tel)
                                             VAR_DIRECCION = IIf(IsNull(rsaux5!calle), "", rsaux5!calle) + " " + IIf(IsNull(rsaux5!numero), "", rsaux5!numero)
                                             VAR_COLONIA = IIf(IsNull(rsaux5!colonia), "", rsaux5!colonia)
                                             var_ciudad = IIf(IsNull(rsaux5!ciudad), "", rsaux5!ciudad)
                                             VAR_MUNICIPIO = IIf(IsNull(rsaux5!municipio), "", rsaux5!municipio)
                                             var_estado = IIf(IsNull(rsaux5!estado), "", rsaux5!estado)
                                             var_pais = IIf(IsNull(rsaux5!pais), "", rsaux5!pais)
                                             VAR_CP = IIf(IsNull(rsaux5!cp), "", rsaux5!cp)
                                             rsaux5.Close
                                          Else
                                             rsaux5.Close
                                             var_nombre = IIf(IsNull(rsaux6!customer_name), "", rsaux6!customer_name)
                                             var_tel = IIf(IsNull(rsaux6!tel), 0, rsaux6!tel)
                                             VAR_DIRECCION = IIf(IsNull(rsaux6!calle), "", rsaux6!calle) + " " + IIf(IsNull(rsaux6!numero), "", rsaux6!numero)
                                             VAR_COLONIA = IIf(IsNull(rsaux6!colonia), "", rsaux6!colonia)
                                             var_ciudad = IIf(IsNull(rsaux6!ciudad), "", rsaux6!ciudad)
                                             VAR_MUNICIPIO = IIf(IsNull(rsaux6!municipio), "", rsaux6!municipio)
                                             var_estado = IIf(IsNull(rsaux6!estado), "", rsaux6!estado)
                                             var_pais = IIf(IsNull(rsaux6!pais), "", rsaux6!pais)
                                             VAR_CP = IIf(IsNull(rsaux6!cp), "", rsaux6!cp)
                                          End If
                                       Else
                                          var_tel = 0
                                          VAR_DIRECCION = ""
                                          VAR_COLONIA = ""
                                          var_ciudad = ""
                                          VAR_MUNICIPIO = ""
                                          var_estado = ""
                                          var_pais = ""
                                          VAR_CP = ""
                                       End If
                                       rsaux6.Close
                                       If var_tel > 0 Then
                                          
                                          strconsulta = "select Phone_Number from hz_contact_points where owner_table_id = ?"
                                          With comandoORA
                                               .ActiveConnection = cnnoracle_4
                                               .CommandType = adCmdText
                                               .CommandText = strconsulta
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(var_tel))
                                               .Parameters.Append parametro
                                          End With
                                          Set rsaux6 = comandoORA.execute
                                          Set comandoORA = Nothing
                                          Set parametro = Nothing
                                          If Not rsaux6.EOF Then
                                             var_telefono = CStr(IIf(IsNull(rsaux6(0).Value), "", rsaux6(0).Value))
                                          Else
                                             var_telefono = ""
                                          End If
                                          rsaux6.Close
                                       Else
                                          var_telefono = ""
                                       End If
                                    Else
                                       var_tel = 0
                                       VAR_DIRECCION = ""
                                       VAR_COLONIA = ""
                                       var_ciudad = ""
                                       VAR_MUNICIPIO = ""
                                       var_estado = ""
                                       var_pais = ""
                                       VAR_CP = ""
                                       var_telefono = ""
                                    End If
                                    
                                    If IsNumeric(txt_total_volumen) Then
                                       var_cubicaje_EMBARQUE = CDbl(Me.txt_total_volumen)
                                    Else
                                       var_cubicaje_EMBARQUE = 0
                                    End If
                                    rsaux4.Open "select distinct source_header_number from tb_Temp_oracle_orden_surtido where  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and grupo = " + CStr(rsaux(0).Value), cnn, adOpenDynamic, adLockOptimistic
                                    While Not rsaux4.EOF
                                          rsaux2.Open "SELECT * FROM tb_oracle_pedidos_asignados_embarques WHERE PEDIDO = " + CStr(rsaux4(0).Value), cnn, adOpenDynamic, adLockOptimistic
                                          If Not rsaux2.EOF Then
                                             var_transporte = ""
                                             rsaux3.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                             If Not rsaux3.EOF Then
                                                var_transporte = IIf(IsNull(rsaux3!vehiculo), "", rsaux3!vehiculo)
                                             End If
                                             rsaux3.Close
                                             strconsulta = "SELECT ORDER_TYPE_ID FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ?"
                                             With comandoORA
                                                  .ActiveConnection = cnnoracle_4
                                                  .CommandType = adCmdText
                                                  .CommandText = strconsulta
                                                  Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rsaux4(0).Value))
                                                  .Parameters.Append parametro
                                             End With
                                             Set rsaux5 = comandoORA.execute
                                             Set comandoORA = Nothing
                                             Set parametro = Nothing
                                             If rsaux5!ORDER_TYPE_ID = 1002 Then
                                                rsaux5.Close
                                                strconsulta = "select ADDRESS_LINE_1||', '||ADDRESS_LINE_2||', '||TOWN_OR_CITY||', '||REGION_1||', '||COUNTRY||' CP:'||POSTAL_CODE DIRECCION, EMAIL from mtl_secondary_inventories a, hr_locations_all b, xxvia_jv_tb_agentes c, po_requisition_headers_ALL D, OE_ORDER_HEADERS_ALL E Where A.location_id = b.location_id and a.secondary_inventory_name = c.subinventory_code AND E.source_document_id = D.requisition_header_id AND A.secondary_inventory_name = D.ATTRIBUTE1 AND E.ORDER_NUMBER = ?"
                                             Else
                                                rsaux5.Close
                                                strconsulta = "SELECT CALLE||' '||NUM_CALLE||' '||NVL(num_interior,'')||', '||colonia||', '||ciudad||', '||estado||', '||pais||' CP: '||codigo_postal DIRECCION FROM oe_order_headers_all a, xxvia_vw_CLIENTES_BCP B WHERE A.SHIP_TO_ORG_ID = B.SITE_USE_ID AND A.ORDER_NUMBER     = ?"
                                             End If
                                             
                                             With comandoORA
                                                  .ActiveConnection = cnnoracle_4
                                                  .CommandType = adCmdText
                                                  .CommandText = strconsulta
                                                  Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rsaux4(0).Value))
                                                  .Parameters.Append parametro
                                             End With
                                             Set rsaux5 = comandoORA.execute
                                             Set comandoORA = Nothing
                                             Set parametro = Nothing
                                             If Not rsaux5.EOF Then
                                                VAR_DIRECCION = IIf(IsNull(rsaux5!DIRECCION), "", rsaux5!DIRECCION)
                                             Else
                                                VAR_DIRECCION = ""
                                             End If
                                             var_cadena = "UPDATE tb_Temp_oracle_orden_surtido SET EMBARQUE = " + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", CUBICAJE = " + CStr(var_cubicaje_EMBARQUE) + " , ORDEN_SURTIDO = " + CStr(IIf(IsNull(rsaux2!orden_pedido), 0, rsaux2!orden_pedido)) + ", ANDEN = '" + CStr(IIf(IsNull(rsaux2!estacion), 0, rsaux2!estacion)) + "', TRANSPORTE = '" + var_transporte + "',"
                                             var_cadena = var_cadena + " pais= '" + var_pais + "', estado = '" + var_estado + "', municipio = '" + VAR_MUNICIPIO + "', ciudad = '" + var_ciudad + "', colonia = '" + VAR_COLONIA + "', direccion = '" + VAR_DIRECCION + "', cp = '" + VAR_CP + "', paqueteria = '" + var_paqueteria + "'"
                                             var_cadena = var_cadena + " WHERE SOURCE_HEADER_NUMBER = " + CStr(rsaux4(0).Value) + " AND inte_tem_consecutivo = " + CStr(var_consecutivo)
                                             rsaux3.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                          End If
                                          rsaux2.Close
                                          rsaux4.MoveNext
                                    Wend
                                    rsaux4.Close
                                    
                                                                  
                                    x = 1
                                    If x = 1 Then
                                       If rsaux9.State = 1 Then
                                          rsaux9.Close
                                       End If
                                       rsaux9.Open "SELECT DISTINCT  GRUPO, SOURCE_HEADER_NUMBER from tb_Temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and grupo = " + CStr(rsaux(0).Value), cnn, adOpenDynamic, adLockOptimistic
                                       VAR_TOTAL_GRUPOS = rsaux9.RecordCount
                                       rsaux9.Close
                                       'VAR_TOTAL_GRUPOS = 1
                                       rsaux10.Open "update tb_Temp_oracle_orden_surtido set cubicaje = " + CStr(var_cubicaje_EMBARQUE) + "  where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                                       If VAR_TOTAL_GRUPOS = 1 Then
                                          Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos_060417.rpt")
                                          reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.grupo} = " + CStr(rsaux(0).Value) + " and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                                          'frmvistasprevias.cr.ReportSource = reporte
                                          For ntablas = 1 To reporte.Database.Tables.Count
                                              reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                          Next ntablas
                                          reporte.PrintOut False
                                          Set reporte = Nothing
                                          
                                       Else
                                          
                                          Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos_ENCABEZADOS.rpt")
                                          reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.grupo} = " + CStr(rsaux(0).Value) + " and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                                          frmvista_previa_auxiliar.cr2.ReportSource = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos_ENCABEZADOS.rpt")
                                          For ntablas = 1 To reporte.Database.Tables.Count
                                              reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                          Next ntablas
                                          reporte.PrintOut False
                                          frmvista_previa_auxiliar.Show 1
                                          
            'frmvistasprevias.cr.ViewReport
            'frmvistasprevias.Caption = "Packing List"
            'frmvistasprevias.Show 1
                                          
                                          Set reporte = Nothing
                                          
                                          rsaux1.Open "select distinct source_header_number from tb_Temp_oracle_orden_surtido where  inte_Tem_Consecutivo = " + CStr(var_consecutivo) + " and grupo = " + CStr(rsaux(0).Value), cnn, adOpenDynamic, adLockOptimistic
                                          var_i = 0
                                          While Not rsaux1.EOF
                                                If var_i = 0 Then
                                                   rsaux10.Open "update tb_Temp_oracle_orden_surtido set pasillo = 'plata' where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number = " + CStr(rsaux1!source_header_number) + " and grupo = " + CStr(rsaux(0).Value), cnn, adOpenDynamic, adLockOptimistic
                                                   var_i = 1
                                                End If
                                                rsaux1.MoveNext
                                          Wend
                                          rsaux1.Close
                                          
                                          
                                          Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos_DETALLE_060417.rpt")
                                          reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.grupo} = " + CStr(rsaux(0).Value) + " and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                                          frmvista_previa_auxiliar.cr2.ReportSource = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos_DETALLE.rpt")
                                          
                                          
                                          For ntablas = 1 To reporte.Database.Tables.Count
                                              reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                          Next ntablas
                                          reporte.PrintOut False
            'frmvistasprevias.cr.ViewReport
            'frmvistasprevias.Caption = "Packing List"
            'frmvistasprevias.Show 1
                                          
                                          Set reporte = Nothing
                                          frmvista_previa_auxiliar.Show 1
                                       
                                       End If
                                       
                                       
                                       'Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos.rpt")
                                       'reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.grupo} = " + CStr(rsaux(0).Value) + " and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                                       'For ntablas = 1 To reporte.Database.Tables.Count
                                       '    reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                       'Next ntablas
                                       'reporte.PrintOut False
                                       'Set reporte = Nothing
                                    Else
                                       If rsaux9.State = 1 Then
                                          rsaux9.Close
                                       End If
                                       rsaux9.Open "SELECT DISTINCT  GRUPO, SOURCE_HEADER_NUMBER from tb_Temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and grupo = " + CStr(rsaux(0).Value), cnn, adOpenDynamic, adLockOptimistic
                                       VAR_TOTAL_GRUPOS = rsaux9.RecordCount
                                       rsaux9.Close
                                       'VAR_TOTAL_GRUPOS = 1
                                       If VAR_TOTAL_GRUPOS = 1 Then
                                          Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos_060417.rpt")
                                          reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.grupo} = " + CStr(rsaux(0).Value) + " and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                                          frmvistasprevias.cr.ReportSource = reporte
                                          For ntablas = 1 To reporte.Database.Tables.Count
                                              reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                          Next ntablas
                                          frmvistasprevias.cr.ViewReport
                                          frmvistasprevias.Caption = "Ordenes de surtido "
                                          frmvistasprevias.Show 1
                                          Set reporte = Nothing
                                       Else
                                       
                                          rsaux1.Open "select distinct source_header_number from tb_Temp_oracle_orden_surtido where  inte_Tem_Consecutivo = " + CStr(var_consecutivo) + " and grupo = " + CStr(rsaux(0).Value), cnn, adOpenDynamic, adLockOptimistic
                                          var_i = 0
                                          While Not rsaux1.EOF
                                                If var_i = 0 Then
                                                   rsaux10.Open "update tb_Temp_oracle_orden_surtido set pasillo = 'plata' where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number = " + CStr(rsaux1!source_header_number) + " and grupo = " + CStr(rsaux(0).Value), cnn, adOpenDynamic, adLockOptimistic
                                                   var_i = 1
                                                End If
                                                rsaux1.MoveNext
                                          Wend
                                          rsaux1.Close
                                          
                                          Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos_ENCABEZADOS.rpt")
                                          reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.grupo} = " + CStr(rsaux(0).Value) + " and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                                          frmvistasprevias.cr.ReportSource = reporte
                                          For ntablas = 1 To reporte.Database.Tables.Count
                                              reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                          Next ntablas
                                          frmvistasprevias.cr.ViewReport
                                          frmvistasprevias.Caption = "Ordenes de surtido "
                                          frmvistasprevias.Show 1
                                          Set reporte = Nothing
                                          
                                          Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos_DETALLE_060417.rpt")
                                          reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.grupo} = " + CStr(rsaux(0).Value) + " and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                                          frmvistasprevias.cr.ReportSource = reporte
                                          For ntablas = 1 To reporte.Database.Tables.Count
                                              reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                          Next ntablas
                                          frmvistasprevias.cr.ViewReport
                                          frmvistasprevias.Caption = "Ordenes de surtido "
                                          frmvistasprevias.Show 1
                                          Set reporte = Nothing
                                       
                                       
                                       End If
                                       
                                    End If
                                 End If
                                 rsaux.MoveNext
                           Wend
                        End If
                        rsaux.Close
                     Else
                        MsgBox "No se pueden imprimir los pedidos, vuelva a intentar la impresión", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "No se pueden imprimir los pedidos, vuelva a intentar la impresión", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "No se pueden imprimir los pedidos, vuelva a intentar la impresión", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "No existen ordenes de surtido", vbOKOnly, "ATENCION"
            End If
            If rs.State = 1 Then
               rs.Close
            End If
            If var_consecutivo > 0 Then
               rs.Open "delete from tb_Temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(IIf(IsNull(var_consecutivo), 0, var_consecutivo)), cnn, adOpenDynamic, adLockOptimistic
            End If
         Else
            'MsgBox "Número superior incorrecto", vbOKOnly, "ATENCION"
         End If
         Else
            MsgBox "El embarque de exportaciones no contiene una caja de transporte adecuada, favor de indicar una para poder imprimir las ordenes de surtido", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "No existen pedidos a imprimir"
      End If

End Sub

Private Sub cmd_hoja_carga_Click()
   If IsNumeric(Me.txt_embarque) Then
      Call ACTUALIZA_INFORMACION
      var_embarque_global = CDbl(Me.txt_embarque)
      frmoracle_hoja_carga_embarques_ruta.Show 1
   Else
      MsgBox "No se a seleccionado un embarque", vbOKOnly
   End If
End Sub

Private Sub cmd_imprimir_Click()
     
      If Me.lv_pedidos.ListItems.Count > 0 Then
         var_si = MsgBox("Desea imprimir las ordenes de surtido?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            If rs.State = 1 Then
               rs.Close
            End If
            var_Cadena_pedidos = ""
            For var_j = 1 To Me.lv_pedidos.ListItems.Count
                Me.lv_pedidos.ListItems.Item(var_j).Selected = True
                If var_Cadena_pedidos = "" Then
                   var_Cadena_pedidos = Me.lv_pedidos.selectedItem
                Else
                   var_Cadena_pedidos = var_Cadena_pedidos + "," + Me.lv_pedidos.selectedItem
                End If
            Next var_j
            'var_cadena_pedidos = "105208"
            rs.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_cadena = "SELECT CAT.LINEA, HCSU.SITE_USE_ID, HCAS.CUST_ACCOUNT_ID, OHA.SHIP_TO_ORG_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id, a.LAST_UPDATE_DATE,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,c.description,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, a.source_header_type_name, oha.source_document_id, C.ATTRIBUTE2, oha.attribute8, oha.attribute9, j.NAME as nombre_ruta, j.salesrep_id as clave_ruta from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, xxvia_vw_articulos_cat cat, XXVIA_VENDEDORES j Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND "
            var_cadena = var_cadena + " to_number(source_header_number)  IN (" + var_Cadena_pedidos + ")"
            var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID "
            var_cadena = var_cadena + " AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND released_status = 'Y' AND A.inventory_item_id  = cat.item_id AND A.ORGANIZATION_ID = Cat.organization_id and oha.salesrep_id = j.salesrep_id "
            rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
'--------------------------
            If Not rs.EOF Then
               cnn.BeginTrans
               rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM tb_Temp_oracle_orden_surtido", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
               Else
                  var_consecutivo = 1
               End If
               rsaux.Close
               rsaux1.Open "insert into tb_Temp_oracle_orden_surtido(inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
               cnn.CommitTrans
               While Not rs.EOF
                     var_establecimiento = rs!SHIP_TO_ORG_ID
                     rsaux.Open "SELECT csu.site_use_id AS VCHA_ESB_ESTABLECIMIENTO_ID, ps.party_site_number, lo.address1 AS VCHA_eSB_NOMBRE FROM hz_cust_site_uses_all csu, hz_cust_acct_sites_all cas, hz_party_sites ps, hz_locations lo Where csu.cust_acct_site_id = cas.cust_acct_site_id AND cas.party_site_id = ps.party_site_id AND ps.location_id = lo.location_id AND csu.site_use_code = 'SHIP_TO' AND csu.LOCATION = ps.party_site_number and csu.site_use_id = " + CStr(var_establecimiento), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        VAR_NOMBRE_ESTABLECIMIENTO = IIf(IsNull(rsaux!vcha_esb_nombre), "", rsaux!vcha_esb_nombre)
                     Else
                        VAR_NOMBRE_ESTABLECIMIENTO = ""
                     End If
                     rsaux.Close
                     var_dia = CStr(Day(CDate(rs!LAST_UPDATE_DATE)))
                     var_mes = CStr(Month(CDate(rs!LAST_UPDATE_DATE)))
                     var_año = CStr(Year(CDate(rs!LAST_UPDATE_DATE)))
                     If Len(Trim(var_dia)) = 1 Then
                        var_dia = "0" + var_dia
                     End If
                     If Len(Trim(var_mes)) = 1 Then
                        var_mes = "0" + var_mes
                     End If
                     var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                     var_cadena = "insert into tb_Temp_oracle_orden_surtido(INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER, DELIVERY_ID, DELIVERY_DETAIL_ID, ORGANIZATION_ID, SUBINVENTORY, DELIVERY_LINE_ID, INVENTORY_ITEM_ID, ITEM_DESCRIPTION, SOURCE_LINE_NUMBER, src_requested_quantity, RELEASED_STATUS, CUSTOMER_NAME, SEGMENT1, COLLECTOR_ID, NAME, date_requested, UBICACION, ESTABLECIMIENTO, NOMBRE_ESTABLECIMIENTO, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, SITE_USE_ID, attribute8, attribute9, LINEA, RUTA, NOMBRE_RUTA)  values "
                     var_cadena = var_cadena + " (" + CStr(var_consecutivo) + ", '" + IIf(IsNull(rs!source_header_number), "", rs!source_header_number) + "', " + CStr(IIf(IsNull(rs!delivery_id), 0, rs!delivery_id)) + ", " + CStr(IIf(IsNull(rs!delivery_detail_id), 0, rs!delivery_detail_id)) + ", " + CStr(IIf(IsNull(rs!organization_id), 0, rs!organization_id)) + ", '" + IIf(IsNull(rs!subinventory), "", rs!subinventory) + "', " + CStr(IIf(IsNull(rs!delivery_line_id), 0, rs!delivery_line_id)) + ", " + CStr(IIf(IsNull(rs!inventory_item_id), "", rs!inventory_item_id)) + ", '" + IIf(IsNull(rs!Description), "", rs!Description) + "', '" + IIf(IsNull(rs!SOURCE_LINE_NUMBER), "", rs!SOURCE_LINE_NUMBER) + "', " + CStr(IIf(IsNull(rs!requested_quantity), 0, rs!requested_quantity)) + ", '" + IIf(IsNull(rs!released_status), "", rs!released_status) + "', '" + Replace(IIf(IsNull(rs!customer_name), "", rs!customer_name), "'", " ") + "', '" + IIf(IsNull(rs!SEGMENT1), "", rs!SEGMENT1) + "'"
                     var_cadena = var_cadena + ", 0,''," + var_fecha + ",'" + IIf(IsNull(rs!attribute2), "", rs!attribute2) + "','" + CStr(var_establecimiento) + "','" + Replace(VAR_NOMBRE_ESTABLECIMIENTO, "'", " ") + "'," + CStr(rs!CUST_ACCOUNT_ID) + ",'" + rs!source_header_type_name + "','" + CStr(IIf(IsNull(rs!source_document_id), "", rs!source_document_id)) + "'," + CStr(IIf(IsNull(rs!site_use_id), 0, rs!site_use_id)) + ",'" + IIf(IsNull(rs!attribute8), "", rs!attribute8) + "','" + IIf(IsNull(rs!ATTRIBUTE9), "", rs!ATTRIBUTE9) + "','" + IIf(IsNull(rs!Linea), "", rs!Linea) + "','" + CStr(rs!CLAVE_RUTA) + "','" + IIf(IsNull(rs!nombre_ruta), "", rs!nombre_ruta) + "')"
                     rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               rsaux1.Open "delete from tb_Temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number is null", cnn, adOpenDynamic, adLockOptimistic
               If rsaux1.State = 1 Then
                  rsaux1.Close
               End If
               
               var_cadena_pedidos_diferencias = ""
               rsaux1.Open "select source_header_number, sum(src_requested_quantity) as cantidad from tb_Temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo) + " group by source_header_number", cnn, adOpenDynamic, adLockOptimistic
               While Not rsaux1.EOF
                     strconsulta = "select sum(requested_quantity)  as cantidad from WSH_DELIVERABLES_V where source_header_number = ? AND RELEASED_STATUS = 'Y'"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(rsaux1!source_header_number))
                          .Parameters.Append parametro
                     End With
                     Set rsaux10 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     If rsaux1!cantidad <> rsaux10!cantidad Then
                        If var_cadena_pedidos_diferencias = "" Then
                           var_cadena_pedidos_diferencias = CStr(rsaux1!source_header_number)
                        Else
                           var_cadena_pedidos_diferencias = var_cadena_pedidos_diferencias + ", " + CStr(rsaux1!source_header_number)
                        End If
                     End If
                     rsaux10.Close
                     rsaux1.MoveNext
               Wend
               rsaux1.Close
               
               If var_cadena_pedidos_diferencias = "" Then
                  rsaux1.Open "select DATE_REQUESTED, source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME, RUTA, NOMBRE_RUTA, SUM(SRC_REQUESTED_QUANTITY) AS CANTIDAD from tb_Temp_oracle_orden_surtido where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " group by DATE_REQUESTED, source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME, RUTA, NOMBRE_RUTA", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        If rsaux1!source_header_type_name = "VIA_PEDIDO_INTERNO" Or rsaux1!source_header_type_name = "TEX_PEDIDO_INTERNO" Then
                           If var_pedido_tienda = 0 Then
                              If rsaux2.State = 1 Then
                                 rsaux2.Close
                              End If
                              rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(rsaux1!source_document_id) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux2.EOF Then
                                 If rsaux4.State = 1 Then
                                    rsaux4.Close
                                 End If
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                                 rsaux4.Close
                              Else
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                                 rsaux4.Close
                              End If
                              rsaux2.Close
                           Else
                              rsaux2.Open "select a.attribute8, B.description from oe_order_headers_all a, MTL_SECONDARY_INVENTORIES b where order_number = " + CStr(rsaux1!source_header_number) + " and a.attribute8 = b.secondary_inventory_name", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux2.EOF Then
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                                 rsaux4.Close
                              Else
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                                 rsaux4.Close
                              End If
                              rsaux2.Close
                           End If
                        Else
                           rsaux6.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                           VAR_PROVEEDOR = rsaux6!collector_id
                           VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux6!Name), "", rsaux6!Name)
                           rsaux6.Close
                        End If
                        var_año_str = CStr(Year(rsaux1!DATE_REQUESTED))
                        If Len(var_año_str) < 2 Then
                           var_año_str = "20" + var_año_str
                        End If
                        var_mes_str = CStr(Month(rsaux1!DATE_REQUESTED))
                        If Len(var_mes_str) < 2 Then
                           var_mes_str = "0" + var_mes_str
                        End If
                        var_dia_str = CStr(Day(rsaux1!DATE_REQUESTED))
                        If Len(var_dia_str) < 2 Then
                           var_dia_str = "0" + var_dia_str
                        End If
                        var_hora_str = CStr(Hour(rsaux1!DATE_REQUESTED))
                        If Len(var_hora_str) < 2 Then
                           var_hora_str = "0" + var_hora_str
                        End If
                        VAR_MINUTO_STR = CStr(Minute(rsaux1!DATE_REQUESTED))
                        If Len(VAR_MINUTO_STR) < 2 Then
                           VAR_MINUTO_STR = "0" + VAR_MINUTO_STR
                        End If
                        VAR_SEGUNDO_STR = CStr(Second(rsaux1!DATE_REQUESTED))
                        If Len(VAR_SEGUNDO_STR) < 2 Then
                           VAR_SEGUNDO_STR = "0" + VAR_SEGUNDO_STR
                        End If
                        var_fecha_pedido = var_año_str + "-" + var_mes_str + "-" + var_dia_str + " " + var_hora_str + ":" + VAR_MINUTO_STR + ":" + VAR_SEGUNDO_STR
                        rsaux6.Open "update tb_Temp_oracle_orden_surtido set COLLECTOR_ID = " + CStr(VAR_PROVEEDOR) + ", NAME = '" + VAR_NOMBRE_PROVEEDOR + "' where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " and CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and source_header_number = " + CStr(rsaux1!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                        rsaux6.Open "SELECT * FROM XXVIA_TB_ORDENES_GRAFICA WHERE ORGANIZACION = " + var_unidad_organizacional + " AND PEDIDO = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If rsaux6.EOF Then
                           rsaux5.Open "INSERT INTO XXVIA_TB_ORDENES_GRAFICA (ORGANIZACION, PEDIDO, PROVEEDOR_ID, NOMBRE_PROVEEDOR, CLIENTE, ESTABLECIMIENTO, CANTIDAD, FECHA, TIPO_PEDIDO, CANTIDAD_LEIDA, FECHA_PEDIDO, RUTA, NOMBRE_RUTA) VALUES (" + var_unidad_organizacional + ", " + CStr(rsaux1!source_header_number) + "," + CStr(VAR_PROVEEDOR) + ",'" + VAR_NOMBRE_PROVEEDOR + "','" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "'," + CStr(rsaux1!cantidad) + ", '" + CStr(rsaux1!DATE_REQUESTED) + "','" + rsaux1!source_header_type_name + "',0, TO_DATE('" + var_fecha_pedido + "','YYYY-MM-DD HH24:MI:SS'),'" + rsaux1!ruta + "', '" + rsaux1!nombre_ruta + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        Else
                           rsaux5.Open "UPDATE XXVIA_TB_ORDENES_GRAFICA SET RUTA = '" + rsaux1!ruta + "', NOMBRE_RUTA = '" + rsaux1!nombre_ruta + "' WHERE PEDIDO = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        End If
                        rsaux6.Close
                        rsaux6.Open "select * from tb_oracle_pedidos_asignados_embarques where pedido = " + CStr(rsaux1!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                        If rsaux6.EOF Then
                           rsaux5.Open "insert into tb_oracle_pedidos_asignados_embarques (AGENTE, NOMBRE_AGENTE, PEDIDO, CLIENTE, PIEZAS, embarque, dia,  mes, AÑO, ORGANIZACION) values (" + CStr(VAR_PROVEEDOR) + ",'" + VAR_NOMBRE_PROVEEDOR + "'," + CStr(rsaux1!source_header_number) + ",'" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "', " + CStr(rsaux1!cantidad) + ",0," + var_dia_str + "," + var_mes_str + "," + var_año_str + "," + CStr(var_unidad_organizacional) + ")", cnn, adOpenDynamic, adLockOptimistic
                        End If
                        rsaux6.Close
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  rsaux1.Open "select DATE_REQUESTED, source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME, SUM(SRC_REQUESTED_QUANTITY) AS CANTIDAD from tb_Temp_oracle_orden_surtido where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " AND (LINEA NOT IN ('CATALOGOS','CATALOGO','POP') OR LINEA IS NULL) group by DATE_REQUESTED, source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME"
                  While Not rsaux1.EOF
                        rsaux5.Open "UPDATE XXVIA_TB_ORDENES_GRAFICA SET CANTIDAD_SIN_CATALOGOS = " + CStr(IIf(IsNull(rsaux1!cantidad), 0, rsaux1!cantidad)) + " WHERE PEDIDO = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  rsaux1.Open "DELETE from tb_Temp_oracle_orden_surtido_aux_2", cnn, adOpenDynamic, adLockOptimistic
                  rsaux1.Open "SELECT * FROM tb_Temp_oracle_orden_surtido where inte_tem_consecutivo =  " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        If rsaux1!Linea = "CATALOGOS" Then
                           var_dia = CStr(Day(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                           var_mes = CStr(Month(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                           var_año = CStr(Year(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                           If Len(Trim(var_dia)) = 1 Then
                              var_dia = "0" + var_dia
                           End If
                           If Len(Trim(var_mes)) = 1 Then
                              var_mes = "0" + var_mes
                           End If
                           var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                           var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_2 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                           var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION) "
                           var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux1!source_header_number + "', " + CStr(rsaux1!delivery_id) + "," + CStr(rsaux1!delivery_detail_id) + ", " + CStr(rsaux1!organization_id) + ",'" + IIf(IsNull(rsaux1!subinventory), "", rsaux1!subinventory) + "', " + CStr(rsaux1!delivery_line_id) + "," + CStr(rsaux1!inventory_item_id) + ",'" + rsaux1!item_description + "','" + CStr(rsaux1!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux1!src_requested_quantity) + ",'" + rsaux1!released_status + "',"
                           var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + rsaux1!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux1!collector_id), 0, rsaux1!collector_id)) + ",'" + IIf(IsNull(rsaux1!Name), "", rsaux1!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux1!ubicacion), "", rsaux1!ubicacion) + "','" + IIf(IsNull(rsaux1!establecimiento), "", rsaux1!establecimiento) + "','" + IIf(IsNull(rsaux1!nombre_Establecimiento), "", rsaux1!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux1!ORDENES), "", rsaux1!ORDENES) + "',"
                           var_cadena = var_cadena + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux1!source_header_type_name), "", rsaux1!source_header_type_name) + "', '" + IIf(IsNull(rsaux1!source_document_id), "", rsaux1!source_document_id) + "','" + IIf(IsNull(rsaux1!pais), "", rsaux1!pais) + "','" + IIf(IsNull(rsaux1!estado), "", rsaux1!estado) + "', '" + IIf(IsNull(rsaux1!municipio), "", rsaux1!municipio) + "', '" + IIf(IsNull(rsaux1!ciudad), "", rsaux1!ciudad) + "', '" + IIf(IsNull(rsaux1!colonia), "", rsaux1!colonia) + "','" + IIf(IsNull(rsaux1!DIRECCION), "", rsaux1!DIRECCION) + "', '" + IIf(IsNull(rsaux1!cp), "", rsaux1!cp) + "',"
                           var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!site_use_id), "", rsaux1!site_use_id) + "','" + IIf(IsNull(rsaux1!paqueteria), "", rsaux1!paqueteria) + "','" + IIf(IsNull(rsaux1!attribute8), "", rsaux1!attribute8) + "','" + IIf(IsNull(rsaux1!ATTRIBUTE9), "", rsaux1!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux1!TELEFONO), "", rsaux1!TELEFONO) + "','" + IIf(IsNull(rsaux1!Linea), "", rsaux1!Linea) + "','" + CStr(IIf(IsNull(rsaux1!ruta), "", rsaux1!ruta)) + "','" + IIf(IsNull(rsaux1!nombre_ruta), "", rsaux1!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux1!ORDEN_SURTIDO), 0, rsaux1!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux1!Embarque), 0, rsaux1!Embarque)) + ", '" + IIf(IsNull(rsaux1!estacion), "", rsaux1!estacion) + "')"
                           rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        Else
                           var_cantidad = rsaux1!src_requested_quantity
                           While var_cantidad > 0
                                 var_dia = CStr(Day(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                 var_mes = CStr(Month(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                 var_año = CStr(Year(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                 If Len(Trim(var_dia)) = 1 Then
                                    var_dia = "0" + var_dia
                                 End If
                                 If Len(Trim(var_mes)) = 1 Then
                                    var_mes = "0" + var_mes
                                 End If
                                 var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_2 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                 var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION) "
                                 var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux1!source_header_number + "', " + CStr(rsaux1!delivery_id) + "," + CStr(rsaux1!delivery_detail_id) + ", " + CStr(rsaux1!organization_id) + ",'" + IIf(IsNull(rsaux1!subinventory), "", rsaux1!subinventory) + "', " + CStr(rsaux1!delivery_line_id) + "," + CStr(rsaux1!inventory_item_id) + ",'" + rsaux1!item_description + "','" + CStr(rsaux1!SOURCE_LINE_NUMBER) + "'," + CStr(1) + ",'" + rsaux1!released_status + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + rsaux1!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux1!collector_id), 0, rsaux1!collector_id)) + ",'" + IIf(IsNull(rsaux1!Name), "", rsaux1!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux1!ubicacion), "", rsaux1!ubicacion) + "','" + IIf(IsNull(rsaux1!establecimiento), "", rsaux1!establecimiento) + "','" + IIf(IsNull(rsaux1!nombre_Establecimiento), "", rsaux1!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux1!ORDENES), "", rsaux1!ORDENES) + "',"
                                 var_cadena = var_cadena + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux1!source_header_type_name), "", rsaux1!source_header_type_name) + "', '" + IIf(IsNull(rsaux1!source_document_id), "", rsaux1!source_document_id) + "','" + IIf(IsNull(rsaux1!pais), "", rsaux1!pais) + "','" + IIf(IsNull(rsaux1!estado), "", rsaux1!estado) + "', '" + IIf(IsNull(rsaux1!municipio), "", rsaux1!municipio) + "', '" + IIf(IsNull(rsaux1!ciudad), "", rsaux1!ciudad) + "', '" + IIf(IsNull(rsaux1!colonia), "", rsaux1!colonia) + "','" + IIf(IsNull(rsaux1!DIRECCION), "", rsaux1!DIRECCION) + "', '" + IIf(IsNull(rsaux1!cp), "", rsaux1!cp) + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!site_use_id), "", rsaux1!site_use_id) + "','" + IIf(IsNull(rsaux1!paqueteria), "", rsaux1!paqueteria) + "','" + IIf(IsNull(rsaux1!attribute8), "", rsaux1!attribute8) + "','" + IIf(IsNull(rsaux1!ATTRIBUTE9), "", rsaux1!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux1!TELEFONO), "", rsaux1!TELEFONO) + "','" + IIf(IsNull(rsaux1!Linea), "", rsaux1!Linea) + "','" + CStr(IIf(IsNull(rsaux1!ruta), "", rsaux1!ruta)) + "','" + IIf(IsNull(rsaux1!nombre_ruta), "", rsaux1!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux1!ORDEN_SURTIDO), 0, rsaux1!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux1!Embarque), 0, rsaux1!Embarque)) + ", '" + IIf(IsNull(rsaux1!estacion), "", rsaux1!estacion) + "')"
                                 rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                 var_cantidad = var_cantidad - 1
                           Wend
                        End If
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  rsaux1.Open "delete from tb_Temp_oracle_orden_surtido_aux_1", cnn, adOpenDynamic, adLockOptimistic
                  rsaux1.Open "select distinct source_header_number from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        var_lote = 1
                        var_contador = 0
                        rsaux2.Open "select * from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number = " + CStr(rsaux1!source_header_number) + " order by ubicacion", cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux2.EOF
                              If var_contador = 50 Then
                                 var_lote = var_lote + 1
                                 var_contador = 0
                              End If
                              rsaux3.Open "SELECT * FROM tb_Temp_oracle_orden_surtido_aux_1 WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and delivery_detail_id = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 rsaux4.Open "UPDATE TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 SET SRC_REQUESTED_QUANTITY = SRC_REQUESTED_QUANTITY + " + CStr(rsaux2!src_requested_quantity) + " WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and DELIVERY_DETAIL_ID = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 var_dia = CStr(Day(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_mes = CStr(Month(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_año = CStr(Year(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 If Len(Trim(var_dia)) = 1 Then
                                    var_dia = "0" + var_dia
                                 End If
                                 If Len(Trim(var_mes)) = 1 Then
                                    var_mes = "0" + var_mes
                                 End If
                                 var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                 
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                 var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,LOTE) "
                                 var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux2!source_header_number + "', " + CStr(rsaux2!delivery_id) + "," + CStr(rsaux2!delivery_detail_id) + ", " + CStr(rsaux2!organization_id) + ",'" + IIf(IsNull(rsaux2!subinventory), "", rsaux2!subinventory) + "', " + CStr(rsaux2!delivery_line_id) + "," + CStr(rsaux2!inventory_item_id) + ",'" + rsaux2!item_description + "','" + CStr(rsaux2!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux2!src_requested_quantity) + ",'" + rsaux2!released_status + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!customer_name), "", rsaux2!customer_name) + "','" + rsaux2!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux2!collector_id), 0, rsaux2!collector_id)) + ",'" + IIf(IsNull(rsaux2!Name), "", rsaux2!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux2!ubicacion), "", rsaux2!ubicacion) + "','" + IIf(IsNull(rsaux2!establecimiento), "", rsaux2!establecimiento) + "','" + IIf(IsNull(rsaux2!nombre_Establecimiento), "", rsaux2!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux2!ORDENES), "", rsaux2!ORDENES) + "',"
                                 var_cadena = var_cadena + CStr(IIf(IsNull(rsaux2!CUST_ACCOUNT_ID), 0, rsaux2!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux2!source_header_type_name), "", rsaux2!source_header_type_name) + "', '" + IIf(IsNull(rsaux2!source_document_id), "", rsaux2!source_document_id) + "','" + IIf(IsNull(rsaux2!pais), "", rsaux2!pais) + "','" + IIf(IsNull(rsaux2!estado), "", rsaux2!estado) + "', '" + IIf(IsNull(rsaux2!municipio), "", rsaux2!municipio) + "', '" + IIf(IsNull(rsaux2!ciudad), "", rsaux2!ciudad) + "', '" + IIf(IsNull(rsaux2!colonia), "", rsaux2!colonia) + "','" + IIf(IsNull(rsaux2!DIRECCION), "", rsaux2!DIRECCION) + "', '" + IIf(IsNull(rsaux2!cp), "", rsaux2!cp) + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!site_use_id), "", rsaux2!site_use_id) + "','" + IIf(IsNull(rsaux2!paqueteria), "", rsaux2!paqueteria) + "','" + IIf(IsNull(rsaux2!attribute8), "", rsaux2!attribute8) + "','" + IIf(IsNull(rsaux2!ATTRIBUTE9), "", rsaux2!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux2!TELEFONO), "", rsaux2!TELEFONO) + "','" + IIf(IsNull(rsaux2!Linea), "", rsaux2!Linea) + "','" + CStr(IIf(IsNull(rsaux2!ruta), "", rsaux2!ruta)) + "','" + IIf(IsNull(rsaux2!nombre_ruta), "", rsaux2!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_SURTIDO), 0, rsaux2!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", '" + IIf(IsNull(rsaux2!estacion), "", rsaux2!estacion) + "'," + CStr(var_lote) + ")"
                                 rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux3.Close
                              If rsaux2!Linea <> "CATALOGOS" Then
                                 var_contador = var_contador + 1
                              End If
                              rsaux2.MoveNext
                        Wend
                        rsaux2.Close
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  rsaux1.Open "insert TB_TEMP_ORACLE_ORDEN_SURTIDO (inte_tem_consecutivo, segment1) values (" + CStr(var_consecutivo) + ",'---------')", cnn, adOpenDynamic, adLockOptimistic
                  rsaux1.Open "delete from TB_TEMP_ORACLE_ORDEN_SURTIDO where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and segment1 <> '---------'", cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.Open "insert into TB_TEMP_ORACLE_ORDEN_SURTIDO select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.Open "delete from TB_TEMP_ORACLE_ORDEN_SURTIDO where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and segment1 = '---------'", cnn, adOpenDynamic, adLockOptimistic
                  var_consecutivo_general = var_consecutivo
                  Call crea_tablas
                  If rsaux.State = 1 Then
                     rsaux.Close
                  End If
                  rsaux.Open "select distinct a.source_header_number from tb_Temp_oracle_orden_surtido_aux_1 a, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES  b where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and a.source_header_number = b.pedido", cnn, adOpenDynamic, adLockOptimistic
                  var_cadena_pedidos_diferencias = ""
                  While Not rsaux.EOF
                        strconsulta = "select sum(requested_quantity)  as cantidad from WSH_DELIVERABLES_V where source_header_number = ? AND RELEASED_STATUS = 'Y'"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(rsaux!source_header_number))
                             .Parameters.Append parametro
                        End With
                        Set rsaux10 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        
                        strconsulta = "SELECT SUM(SRC_REQUESTED_QUANTITY) AS CANTIDAD FROM XXVIA_TB_PEDIDOS_DIVIDIDOS WHERE SOURCE_HEADER_NUMBER = ?"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(rsaux!source_header_number))
                             .Parameters.Append parametro
                        End With
                        Set rsaux11 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                     
                     
                        If rsaux11!cantidad <> rsaux10!cantidad Then
                           If var_cadena_pedidos_diferencias = "" Then
                              var_cadena_pedidos_diferencias = CStr(rsaux1!source_header_number)
                           Else
                              var_cadena_pedidos_diferencias = var_cadena_pedidos_diferencias + ", " + CStr(rsaux1!source_header_number)
                           End If
                        End If
                        rsaux10.Close
                        rsaux11.Close
                        rsaux.MoveNext
                  Wend
                  rsaux.Close
                  If var_cadena_pedidos_diferencias = "" Then
                     If var_imprime_pedidos = 1 Then
                        rsaux.Open "select distinct a.source_header_number, a.lote, orden_pedido from tb_Temp_oracle_orden_surtido_aux_1 a, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES  b where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and a.source_header_number = b.pedido order by orden_pedido", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           While Not rsaux.EOF
                                 x = 1
                                 If x = 1 Then
                                    
                                    strconsulta = "select shipping_method_code, packing_instructions from oe_order_headers_all where order_number = ?"
                                    With comandoORA
                                         .ActiveConnection = cnnoracle_4
                                         .CommandType = adCmdText
                                         .CommandText = strconsulta
                                         Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rsaux!source_header_number))
                                         .Parameters.Append parametro
                                    End With
                                    Set rs = comandoORA.execute
                                    Set comandoORA = Nothing
                                    Set parametro = Nothing
                                    
                                    var_paqueteria = ""
                                    If Not rs.EOF Then
                                       VAR_COMENTARIOS = IIf(IsNull(rs!packing_instructions), "", rs!packing_instructions)
                                       var_tipo_metodo = IIf(IsNull(rs(0).Value), "", rs(0).Value)
                                       If var_tipo_metodo <> "" Then
                                          
                                          strconsulta = "SELECT description FROM fnd_lookup_values where lookup_type = 'SHIP_METHOD' and lookup_code = ? AND LANGUAGE = 'ESA'"
                                          With comandoORA
                                               .ActiveConnection = cnnoracle_4
                                               .CommandType = adCmdText
                                               .CommandText = strconsulta
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_tipo_metodo)
                                               .Parameters.Append parametro
                                          End With
                                          Set rsaux1 = comandoORA.execute
                                          Set comandoORA = Nothing
                                          Set parametro = Nothing
                                          
                                          If Not rsaux1.EOF Then
                                             var_paqueteria = IIf(IsNull(rsaux1(0).Value), "", rsaux1(0).Value)
                                          End If
                                          rsaux1.Close
                                       End If
                                    End If
                                    rs.Close
                                    
                                    
                                    
                                    strconsulta = "SELECT  HPS.party_site_id as tel, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1, OHA.packing_instructions from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.invoice_TO_ORG_ID AND oha.order_number = ? AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                                    With comandoORA
                                         .ActiveConnection = cnnoracle_4
                                         .CommandType = adCmdText
                                         .CommandText = strconsulta
                                         Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rsaux!source_header_number))
                                         .Parameters.Append parametro
                                    End With
                                    Set rsaux6 = comandoORA.execute
                                    Set comandoORA = Nothing
                                    Set parametro = Nothing
                                    
                                    If Not rsaux6.EOF Then
                                    
                                       
                                       strconsulta = "SELECT  HPS.party_site_id as tel, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1, OHA.packing_instructions from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.ship_TO_ORG_ID AND oha.order_number = ? AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                                       With comandoORA
                                            .ActiveConnection = cnnoracle_4
                                            .CommandType = adCmdText
                                            .CommandText = strconsulta
                                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, IIf(IsNull(rsaux!source_header_number), "", rsaux!source_header_number))
                                            .Parameters.Append parametro
                                       End With
                                       Set rsaux5 = comandoORA.execute
                                       Set comandoORA = Nothing
                                       Set parametro = Nothing
                                       
                                       'rsaux5.Open "SELECT  HPS.party_site_id as tel, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1, OHA.packing_instructions from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.ship_TO_ORG_ID AND oha.order_number = '" + IIf(IsNull(rsaux!source_header_number), "", rsaux!source_header_number) + "' AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                       If Not rsaux5.EOF Then
                                          var_nombre = IIf(IsNull(rsaux5!customer_name), "", rsaux5!customer_name)
                                          var_tel = IIf(IsNull(rsaux5!tel), 0, rsaux5!tel)
                                          VAR_DIRECCION = IIf(IsNull(rsaux5!calle), "", rsaux5!calle) + " " + IIf(IsNull(rsaux5!numero), "", rsaux5!numero)
                                          VAR_COLONIA = IIf(IsNull(rsaux5!colonia), "", rsaux5!colonia)
                                          var_ciudad = IIf(IsNull(rsaux5!ciudad), "", rsaux5!ciudad)
                                          VAR_MUNICIPIO = IIf(IsNull(rsaux5!municipio), "", rsaux5!municipio)
                                          var_estado = IIf(IsNull(rsaux5!estado), "", rsaux5!estado)
                                          var_pais = IIf(IsNull(rsaux5!pais), "", rsaux5!pais)
                                          VAR_CP = IIf(IsNull(rsaux5!cp), "", rsaux5!cp)
                                          rsaux5.Close
                                       Else
                                          rsaux5.Close
                                          var_nombre = IIf(IsNull(rsaux6!customer_name), "", rsaux6!customer_name)
                                          var_tel = IIf(IsNull(rsaux6!tel), 0, rsaux6!tel)
                                          VAR_DIRECCION = IIf(IsNull(rsaux6!calle), "", rsaux6!calle) + " " + IIf(IsNull(rsaux6!numero), "", rsaux6!numero)
                                          VAR_COLONIA = IIf(IsNull(rsaux6!colonia), "", rsaux6!colonia)
                                          var_ciudad = IIf(IsNull(rsaux6!ciudad), "", rsaux6!ciudad)
                                          VAR_MUNICIPIO = IIf(IsNull(rsaux6!municipio), "", rsaux6!municipio)
                                          var_estado = IIf(IsNull(rsaux6!estado), "", rsaux6!estado)
                                          var_pais = IIf(IsNull(rsaux6!pais), "", rsaux6!pais)
                                          VAR_CP = IIf(IsNull(rsaux6!cp), "", rsaux6!cp)
                                       End If
                                    Else
                                       var_tel = 0
                                       VAR_DIRECCION = ""
                                       VAR_COLONIA = ""
                                       var_ciudad = ""
                                       VAR_MUNICIPIO = ""
                                       var_estado = ""
                                       var_pais = ""
                                       VAR_CP = ""
                                    End If
                                    rsaux6.Close
                                    If var_tel > 0 Then
                                       
                                       strconsulta = "select Phone_Number from hz_contact_points where owner_table_id = ?"
                                       With comandoORA
                                            .ActiveConnection = cnnoracle_4
                                            .CommandType = adCmdText
                                            .CommandText = strconsulta
                                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(var_tel))
                                            .Parameters.Append parametro
                                       End With
                                       Set rsaux6 = comandoORA.execute
                                       Set comandoORA = Nothing
                                       Set parametro = Nothing
                                       
                                       
                                       'rsaux6.Open "select Phone_Number from hz_contact_points where owner_table_id = " + CStr(var_tel), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                       If Not rsaux6.EOF Then
                                          var_telefono = CStr(IIf(IsNull(rsaux6(0).Value), "", rsaux6(0).Value))
                                       Else
                                          var_telefono = ""
                                       End If
                                       rsaux6.Close
                                    Else
                                       var_telefono = ""
                                    End If
                                       
                                    
                                    
                                    
                                    rsaux2.Open "SELECT * FROM tb_oracle_pedidos_asignados_embarques WHERE PEDIDO = " + CStr(rsaux(0).Value), cnn, adOpenDynamic, adLockOptimistic
                                    If Not rsaux2.EOF Then
                                       var_transporte = ""
                                       rsaux3.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                       If Not rsaux3.EOF Then
                                          var_transporte = IIf(IsNull(rsaux3!vehiculo), "", rsaux3!vehiculo)
                                       End If
                                       rsaux3.Close
                                       var_cadena = "UPDATE tb_Temp_oracle_orden_surtido SET EMBARQUE = " + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", ORDEN_SURTIDO = " + CStr(IIf(IsNull(rsaux2!orden_pedido), 0, rsaux2!orden_pedido)) + ", ANDEN = '" + CStr(IIf(IsNull(rsaux2!estacion), 0, rsaux2!estacion)) + "', TRANSPORTE = '" + var_transporte + "',"
                                       var_cadena = var_cadena + " pais= '" + var_pais + "', estado = '" + var_estado + "', municipio = '" + VAR_MUNICIPIO + "', ciudad = '" + var_ciudad + "', colonia = '" + VAR_COLONIA + "', direccion = '" + VAR_DIRECCION + "', cp = '" + VAR_CP + "', paqueteria = '" + var_paqueteria + "'"
                                       var_cadena = var_cadena + " WHERE SOURCE_HEADER_NUMBER = " + CStr(rsaux(0).Value) + " AND inte_tem_consecutivo = " + CStr(var_consecutivo)
                                       rsaux3.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                    End If
                                    rsaux2.Close
                                 
                                    
                                    
                                    x = 1
                                    If x = 1 Then
                                       Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA.rpt")
                                       reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.SOURCE_HEADER_NUMBER} = '" + rsaux(0).Value + "' and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_ORACLE_ORDEN_SURTIDO.LOTE} = " + CStr(rsaux(1).Value)
                                       For ntablas = 1 To reporte.Database.Tables.Count
                                           reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                       Next ntablas
                                       reporte.PrintOut False
                                       Set reporte = Nothing
                                    Else
                                    
                                       Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA.rpt")
                                       reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.SOURCE_HEADER_NUMBER} = '" + rsaux(0).Value + "' and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_ORACLE_ORDEN_SURTIDO.LOTE} = " + CStr(rsaux(1).Value)
                                       frmvistasprevias.cr.ReportSource = reporte
                                       For ntablas = 1 To reporte.Database.Tables.Count
                                           reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                       Next ntablas
                                       frmvistasprevias.cr.ViewReport
                                       frmvistasprevias.Caption = "Ordenes de surtido pendientes de empacar o facturar"
                                       frmvistasprevias.Show 1
                                       Set reporte = Nothing
                                    End If
                                 End If
                                 rsaux.MoveNext
                           Wend
                        End If
                        rsaux.Close
                     Else
                        MsgBox "No se pueden imprimir los pedidos, vuelva a intentar la impresión", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "No se pueden imprimir los pedidos, vuelva a intentar la impresión", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "No se pueden imprimir los pedidos, vuelva a intentar la impresión", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "No existen ordenes de surtido", vbOKOnly, "ATENCION"
            End If
            If rs.State = 1 Then
               rs.Close
            End If
            If var_consecutivo Then
               rs.Open "delete from tb_Temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(IIf(IsNull(var_consecutivo), 0, var_consecutivo)), cnn, adOpenDynamic, adLockOptimistic
            End If
         Else
            'MsgBox "Número superior incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "No existen pedidos a imprimir"
      End If
End Sub

Private Sub cmd_imprimir_nuevo_metodo_Click()
     
      If Me.lv_pedidos.ListItems.Count > 0 Then
         var_si = MsgBox("Desea imprimir las ordenes de surtido?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            If rs.State = 1 Then
               rs.Close
            End If
            var_Cadena_pedidos = ""
            For var_j = 1 To Me.lv_pedidos.ListItems.Count
                Me.lv_pedidos.ListItems.Item(var_j).Selected = True
                If var_Cadena_pedidos = "" Then
                   var_Cadena_pedidos = Me.lv_pedidos.selectedItem
                Else
                   var_Cadena_pedidos = var_Cadena_pedidos + "," + Me.lv_pedidos.selectedItem
                End If
            Next var_j
            'var_cadena_pedidos = "105208"
            rs.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_cadena = "SELECT CAT.LINEA, HCSU.SITE_USE_ID, HCAS.CUST_ACCOUNT_ID, OHA.SHIP_TO_ORG_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id, a.LAST_UPDATE_DATE,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,c.description,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, a.source_header_type_name, oha.source_document_id, C.ATTRIBUTE2, oha.attribute8, oha.attribute9, j.NAME as nombre_ruta, j.salesrep_id as clave_ruta from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, xxvia_vw_articulos_cat cat, XXVIA_VENDEDORES j Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND "
            var_cadena = var_cadena + " to_number(source_header_number)  IN (" + var_Cadena_pedidos + ")"
            var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID "
            var_cadena = var_cadena + " AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND released_status = 'Y' AND A.inventory_item_id  = cat.item_id AND A.ORGANIZATION_ID = Cat.organization_id and oha.salesrep_id = j.salesrep_id "
            rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
'--------------------------
            If Not rs.EOF Then
               cnn.BeginTrans
               rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM tb_Temp_oracle_orden_surtido", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
               Else
                  var_consecutivo = 1
               End If
               rsaux.Close
               rsaux1.Open "insert into tb_Temp_oracle_orden_surtido(inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
               cnn.CommitTrans
               While Not rs.EOF
                     var_establecimiento = rs!SHIP_TO_ORG_ID
                     rsaux.Open "SELECT csu.site_use_id AS VCHA_ESB_ESTABLECIMIENTO_ID, ps.party_site_number, lo.address1 AS VCHA_eSB_NOMBRE FROM hz_cust_site_uses_all csu, hz_cust_acct_sites_all cas, hz_party_sites ps, hz_locations lo Where csu.cust_acct_site_id = cas.cust_acct_site_id AND cas.party_site_id = ps.party_site_id AND ps.location_id = lo.location_id AND csu.site_use_code = 'SHIP_TO' AND csu.LOCATION = ps.party_site_number and csu.site_use_id = " + CStr(var_establecimiento), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        VAR_NOMBRE_ESTABLECIMIENTO = IIf(IsNull(rsaux!vcha_esb_nombre), "", rsaux!vcha_esb_nombre)
                     Else
                        VAR_NOMBRE_ESTABLECIMIENTO = ""
                     End If
                     rsaux.Close
                     var_dia = CStr(Day(CDate(rs!LAST_UPDATE_DATE)))
                     var_mes = CStr(Month(CDate(rs!LAST_UPDATE_DATE)))
                     var_año = CStr(Year(CDate(rs!LAST_UPDATE_DATE)))
                     If Len(Trim(var_dia)) = 1 Then
                        var_dia = "0" + var_dia
                     End If
                     If Len(Trim(var_mes)) = 1 Then
                        var_mes = "0" + var_mes
                     End If
                     var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                     rsaux1.Open "select * from tb_oracle_multiplos where segment1 = '" + IIf(IsNull(rs!SEGMENT1), "", rs!SEGMENT1) + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux1.EOF Then
                        VAR_MULTIPLO = IIf(IsNull(rsaux1!MULTIPLO), 1, rsaux1!MULTIPLO)
                     Else
                        VAR_MULTIPLO = 1
                     End If
                     rsaux1.Close
                     var_cadena = "insert into tb_Temp_oracle_orden_surtido(INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER, DELIVERY_ID, DELIVERY_DETAIL_ID, ORGANIZATION_ID, SUBINVENTORY, DELIVERY_LINE_ID, INVENTORY_ITEM_ID, ITEM_DESCRIPTION, SOURCE_LINE_NUMBER, src_requested_quantity, RELEASED_STATUS, CUSTOMER_NAME, SEGMENT1, COLLECTOR_ID, NAME, date_requested, UBICACION, ESTABLECIMIENTO, NOMBRE_ESTABLECIMIENTO, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, SITE_USE_ID, attribute8, attribute9, LINEA, RUTA, NOMBRE_RUTA,MULTIPLO)  values "
                     var_cadena = var_cadena + " (" + CStr(var_consecutivo) + ", '" + IIf(IsNull(rs!source_header_number), "", rs!source_header_number) + "', " + CStr(IIf(IsNull(rs!delivery_id), 0, rs!delivery_id)) + ", " + CStr(IIf(IsNull(rs!delivery_detail_id), 0, rs!delivery_detail_id)) + ", " + CStr(IIf(IsNull(rs!organization_id), 0, rs!organization_id)) + ", '" + IIf(IsNull(rs!subinventory), "", rs!subinventory) + "', " + CStr(IIf(IsNull(rs!delivery_line_id), 0, rs!delivery_line_id)) + ", " + CStr(IIf(IsNull(rs!inventory_item_id), "", rs!inventory_item_id)) + ", '" + IIf(IsNull(rs!Description), "", rs!Description) + "', '" + IIf(IsNull(rs!SOURCE_LINE_NUMBER), "", rs!SOURCE_LINE_NUMBER) + "', " + CStr(IIf(IsNull(rs!requested_quantity), 0, rs!requested_quantity)) + ", '" + IIf(IsNull(rs!released_status), "", rs!released_status) + "', '" + Replace(IIf(IsNull(rs!customer_name), "", rs!customer_name), "'", " ") + "', '" + IIf(IsNull(rs!SEGMENT1), "", rs!SEGMENT1) + "'"
                     var_cadena = var_cadena + ", 0,''," + var_fecha + ",'" + IIf(IsNull(rs!attribute2), "", rs!attribute2) + "','" + CStr(var_establecimiento) + "','" + Replace(VAR_NOMBRE_ESTABLECIMIENTO, "'", " ") + "'," + CStr(rs!CUST_ACCOUNT_ID) + ",'" + rs!source_header_type_name + "','" + CStr(IIf(IsNull(rs!source_document_id), "", rs!source_document_id)) + "'," + CStr(IIf(IsNull(rs!site_use_id), 0, rs!site_use_id)) + ",'" + IIf(IsNull(rs!attribute8), "", rs!attribute8) + "','" + IIf(IsNull(rs!ATTRIBUTE9), "", rs!ATTRIBUTE9) + "','" + IIf(IsNull(rs!Linea), "", rs!Linea) + "','" + CStr(rs!CLAVE_RUTA) + "','" + IIf(IsNull(rs!nombre_ruta), "", rs!nombre_ruta) + "'," + CStr(VAR_MULTIPLO) + ")"
                     rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               rsaux1.Open "delete from tb_Temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number is null", cnn, adOpenDynamic, adLockOptimistic
               If rsaux1.State = 1 Then
                  rsaux1.Close
               End If
               
               var_cadena_pedidos_diferencias = ""
               rsaux1.Open "select source_header_number, sum(src_requested_quantity) as cantidad from tb_Temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo) + " group by source_header_number", cnn, adOpenDynamic, adLockOptimistic
               While Not rsaux1.EOF
                     strconsulta = "select sum(requested_quantity)  as cantidad from WSH_DELIVERABLES_V where source_header_number = ? AND RELEASED_STATUS = 'Y'"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(rsaux1!source_header_number))
                          .Parameters.Append parametro
                     End With
                     Set rsaux10 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     If rsaux1!cantidad <> rsaux10!cantidad Then
                        If var_cadena_pedidos_diferencias = "" Then
                           var_cadena_pedidos_diferencias = CStr(rsaux1!source_header_number)
                        Else
                           var_cadena_pedidos_diferencias = var_cadena_pedidos_diferencias + ", " + CStr(rsaux1!source_header_number)
                        End If
                     End If
                     rsaux10.Close
                     rsaux1.MoveNext
               Wend
               rsaux1.Close
               
               If var_cadena_pedidos_diferencias = "" Then
                  rsaux1.Open "select DATE_REQUESTED, source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME, RUTA, NOMBRE_RUTA, SUM(SRC_REQUESTED_QUANTITY) AS CANTIDAD from tb_Temp_oracle_orden_surtido where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " group by DATE_REQUESTED, source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME, RUTA, NOMBRE_RUTA", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        If rsaux1!source_header_type_name = "VIA_PEDIDO_INTERNO" Or rsaux1!source_header_type_name = "TEX_PEDIDO_INTERNO" Then
                           If var_pedido_tienda = 0 Then
                              If rsaux2.State = 1 Then
                                 rsaux2.Close
                              End If
                              rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(rsaux1!source_document_id) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux2.EOF Then
                                 If rsaux4.State = 1 Then
                                    rsaux4.Close
                                 End If
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                                 rsaux4.Close
                              Else
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                                 rsaux4.Close
                              End If
                              rsaux2.Close
                           Else
                              rsaux2.Open "select a.attribute8, B.description from oe_order_headers_all a, MTL_SECONDARY_INVENTORIES b where order_number = " + CStr(rsaux1!source_header_number) + " and a.attribute8 = b.secondary_inventory_name", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux2.EOF Then
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                                 rsaux4.Close
                              Else
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                                 rsaux4.Close
                              End If
                              rsaux2.Close
                           End If
                        Else
                           rsaux6.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                           VAR_PROVEEDOR = rsaux6!collector_id
                           VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux6!Name), "", rsaux6!Name)
                           rsaux6.Close
                        End If
                        var_año_str = CStr(Year(rsaux1!DATE_REQUESTED))
                        If Len(var_año_str) < 2 Then
                           var_año_str = "20" + var_año_str
                        End If
                        var_mes_str = CStr(Month(rsaux1!DATE_REQUESTED))
                        If Len(var_mes_str) < 2 Then
                           var_mes_str = "0" + var_mes_str
                        End If
                        var_dia_str = CStr(Day(rsaux1!DATE_REQUESTED))
                        If Len(var_dia_str) < 2 Then
                           var_dia_str = "0" + var_dia_str
                        End If
                        var_hora_str = CStr(Hour(rsaux1!DATE_REQUESTED))
                        If Len(var_hora_str) < 2 Then
                           var_hora_str = "0" + var_hora_str
                        End If
                        VAR_MINUTO_STR = CStr(Minute(rsaux1!DATE_REQUESTED))
                        If Len(VAR_MINUTO_STR) < 2 Then
                           VAR_MINUTO_STR = "0" + VAR_MINUTO_STR
                        End If
                        VAR_SEGUNDO_STR = CStr(Second(rsaux1!DATE_REQUESTED))
                        If Len(VAR_SEGUNDO_STR) < 2 Then
                           VAR_SEGUNDO_STR = "0" + VAR_SEGUNDO_STR
                        End If
                        var_fecha_pedido = var_año_str + "-" + var_mes_str + "-" + var_dia_str + " " + var_hora_str + ":" + VAR_MINUTO_STR + ":" + VAR_SEGUNDO_STR
                        rsaux6.Open "update tb_Temp_oracle_orden_surtido set COLLECTOR_ID = " + CStr(VAR_PROVEEDOR) + ", NAME = '" + VAR_NOMBRE_PROVEEDOR + "' where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " and CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and source_header_number = " + CStr(rsaux1!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                        rsaux6.Open "SELECT * FROM XXVIA_TB_ORDENES_GRAFICA WHERE ORGANIZACION = " + var_unidad_organizacional + " AND PEDIDO = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If rsaux6.EOF Then
                           rsaux5.Open "INSERT INTO XXVIA_TB_ORDENES_GRAFICA (ORGANIZACION, PEDIDO, PROVEEDOR_ID, NOMBRE_PROVEEDOR, CLIENTE, ESTABLECIMIENTO, CANTIDAD, FECHA, TIPO_PEDIDO, CANTIDAD_LEIDA, FECHA_PEDIDO, RUTA, NOMBRE_RUTA) VALUES (" + var_unidad_organizacional + ", " + CStr(rsaux1!source_header_number) + "," + CStr(VAR_PROVEEDOR) + ",'" + VAR_NOMBRE_PROVEEDOR + "','" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "'," + CStr(rsaux1!cantidad) + ", '" + CStr(rsaux1!DATE_REQUESTED) + "','" + rsaux1!source_header_type_name + "',0, TO_DATE('" + var_fecha_pedido + "','YYYY-MM-DD HH24:MI:SS'),'" + rsaux1!ruta + "', '" + rsaux1!nombre_ruta + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        Else
                           rsaux5.Open "UPDATE XXVIA_TB_ORDENES_GRAFICA SET RUTA = '" + rsaux1!ruta + "', NOMBRE_RUTA = '" + rsaux1!nombre_ruta + "' WHERE PEDIDO = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        End If
                        rsaux6.Close
                        rsaux6.Open "select * from tb_oracle_pedidos_asignados_embarques where pedido = " + CStr(rsaux1!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                        If rsaux6.EOF Then
                           'ORGANIZACION
                           rsaux5.Open "insert into tb_oracle_pedidos_asignados_embarques (AGENTE, NOMBRE_AGENTE, PEDIDO, CLIENTE, PIEZAS, embarque, dia,  mes, AÑO, ORGANIZACION) values (" + CStr(VAR_PROVEEDOR) + ",'" + VAR_NOMBRE_PROVEEDOR + "'," + CStr(rsaux1!source_header_number) + ",'" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "', " + CStr(rsaux1!cantidad) + ",0," + var_dia_str + "," + var_mes_str + "," + var_año_str + "," + CStr(var_unidad_organizacional) + ")", cnn, adOpenDynamic, adLockOptimistic
                        End If
                        rsaux6.Close
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  rsaux1.Open "select DATE_REQUESTED, source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME, SUM(SRC_REQUESTED_QUANTITY) AS CANTIDAD from tb_Temp_oracle_orden_surtido where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " AND (LINEA <> 'CATALOGOS' OR LINEA IS NULL) group by DATE_REQUESTED, source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME"
                  While Not rsaux1.EOF
                        rsaux5.Open "UPDATE XXVIA_TB_ORDENES_GRAFICA SET CANTIDAD_SIN_CATALOGOS = " + CStr(IIf(IsNull(rsaux1!cantidad), 0, rsaux1!cantidad)) + " WHERE PEDIDO = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  rsaux1.Open "DELETE from tb_Temp_oracle_orden_surtido_aux_2", cnn, adOpenDynamic, adLockOptimistic
                  rsaux1.Open "SELECT * FROM tb_Temp_oracle_orden_surtido where inte_tem_consecutivo =  " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        If rsaux1!Linea = "CATALOGOS" Then
                           var_dia = CStr(Day(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                           var_mes = CStr(Month(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                           var_año = CStr(Year(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                           If Len(Trim(var_dia)) = 1 Then
                              var_dia = "0" + var_dia
                           End If
                           If Len(Trim(var_mes)) = 1 Then
                              var_mes = "0" + var_mes
                           End If
                           var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                           var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_2 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                           var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,MULTIPLO) "
                           var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux1!source_header_number + "', " + CStr(rsaux1!delivery_id) + "," + CStr(rsaux1!delivery_detail_id) + ", " + CStr(rsaux1!organization_id) + ",'" + IIf(IsNull(rsaux1!subinventory), "", rsaux1!subinventory) + "', " + CStr(rsaux1!delivery_line_id) + "," + CStr(rsaux1!inventory_item_id) + ",'" + rsaux1!item_description + "','" + CStr(rsaux1!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux1!src_requested_quantity) + ",'" + rsaux1!released_status + "',"
                           var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + rsaux1!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux1!collector_id), 0, rsaux1!collector_id)) + ",'" + IIf(IsNull(rsaux1!Name), "", rsaux1!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux1!ubicacion), "", rsaux1!ubicacion) + "','" + IIf(IsNull(rsaux1!establecimiento), "", rsaux1!establecimiento) + "','" + IIf(IsNull(rsaux1!nombre_Establecimiento), "", rsaux1!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux1!ORDENES), "", rsaux1!ORDENES) + "',"
                           var_cadena = var_cadena + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux1!source_header_type_name), "", rsaux1!source_header_type_name) + "', '" + IIf(IsNull(rsaux1!source_document_id), "", rsaux1!source_document_id) + "','" + IIf(IsNull(rsaux1!pais), "", rsaux1!pais) + "','" + IIf(IsNull(rsaux1!estado), "", rsaux1!estado) + "', '" + IIf(IsNull(rsaux1!municipio), "", rsaux1!municipio) + "', '" + IIf(IsNull(rsaux1!ciudad), "", rsaux1!ciudad) + "', '" + IIf(IsNull(rsaux1!colonia), "", rsaux1!colonia) + "','" + IIf(IsNull(rsaux1!DIRECCION), "", rsaux1!DIRECCION) + "', '" + IIf(IsNull(rsaux1!cp), "", rsaux1!cp) + "',"
                           var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!site_use_id), "", rsaux1!site_use_id) + "','" + IIf(IsNull(rsaux1!paqueteria), "", rsaux1!paqueteria) + "','" + IIf(IsNull(rsaux1!attribute8), "", rsaux1!attribute8) + "','" + IIf(IsNull(rsaux1!ATTRIBUTE9), "", rsaux1!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux1!TELEFONO), "", rsaux1!TELEFONO) + "','" + IIf(IsNull(rsaux1!Linea), "", rsaux1!Linea) + "','" + CStr(IIf(IsNull(rsaux1!ruta), "", rsaux1!ruta)) + "','" + IIf(IsNull(rsaux1!nombre_ruta), "", rsaux1!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux1!ORDEN_SURTIDO), 0, rsaux1!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux1!Embarque), 0, rsaux1!Embarque)) + ", '" + IIf(IsNull(rsaux1!estacion), "", rsaux1!estacion) + "'," + CStr(IIf(IsNull(rsaux1!MULTIPLO), "", rsaux1!MULTIPLO)) + ")"
                           rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        Else
                           var_cantidad_total = rsaux1!src_requested_quantity
                           If rsaux1!MULTIPLO > 1 Then
                              While var_cantidad_total > 0
                                    If var_cantidad_total < rsaux1!MULTIPLO Then
                                       var_cantidad = var_cantidad_total
                                    Else
                                       var_cantidad = rsaux1!MULTIPLO
                                    End If
                                    
                                    var_dia = CStr(Day(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                    var_mes = CStr(Month(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                    var_año = CStr(Year(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                    If Len(Trim(var_dia)) = 1 Then
                                       var_dia = "0" + var_dia
                                    End If
                                    If Len(Trim(var_mes)) = 1 Then
                                      var_mes = "0" + var_mes
                                    End If
                                    
                                    var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                    var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_2 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                    var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,MULTIPLO) "
                                    var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux1!source_header_number + "', " + CStr(rsaux1!delivery_id) + "," + CStr(rsaux1!delivery_detail_id) + ", " + CStr(rsaux1!organization_id) + ",'" + IIf(IsNull(rsaux1!subinventory), "", rsaux1!subinventory) + "', " + CStr(rsaux1!delivery_line_id) + "," + CStr(rsaux1!inventory_item_id) + ",'" + rsaux1!item_description + "','" + CStr(rsaux1!SOURCE_LINE_NUMBER) + "'," + CStr(var_cantidad) + ",'" + rsaux1!released_status + "',"
                                    var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + rsaux1!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux1!collector_id), 0, rsaux1!collector_id)) + ",'" + IIf(IsNull(rsaux1!Name), "", rsaux1!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux1!ubicacion), "", rsaux1!ubicacion) + "','" + IIf(IsNull(rsaux1!establecimiento), "", rsaux1!establecimiento) + "','" + IIf(IsNull(rsaux1!nombre_Establecimiento), "", rsaux1!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux1!ORDENES), "", rsaux1!ORDENES) + "',"
                                    var_cadena = var_cadena + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux1!source_header_type_name), "", rsaux1!source_header_type_name) + "', '" + IIf(IsNull(rsaux1!source_document_id), "", rsaux1!source_document_id) + "','" + IIf(IsNull(rsaux1!pais), "", rsaux1!pais) + "','" + IIf(IsNull(rsaux1!estado), "", rsaux1!estado) + "', '" + IIf(IsNull(rsaux1!municipio), "", rsaux1!municipio) + "', '" + IIf(IsNull(rsaux1!ciudad), "", rsaux1!ciudad) + "', '" + IIf(IsNull(rsaux1!colonia), "", rsaux1!colonia) + "','" + IIf(IsNull(rsaux1!DIRECCION), "", rsaux1!DIRECCION) + "', '" + IIf(IsNull(rsaux1!cp), "", rsaux1!cp) + "',"
                                    var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!site_use_id), "", rsaux1!site_use_id) + "','" + IIf(IsNull(rsaux1!paqueteria), "", rsaux1!paqueteria) + "','" + IIf(IsNull(rsaux1!attribute8), "", rsaux1!attribute8) + "','" + IIf(IsNull(rsaux1!ATTRIBUTE9), "", rsaux1!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux1!TELEFONO), "", rsaux1!TELEFONO) + "','" + IIf(IsNull(rsaux1!Linea), "", rsaux1!Linea) + "','" + CStr(IIf(IsNull(rsaux1!ruta), "", rsaux1!ruta)) + "','" + IIf(IsNull(rsaux1!nombre_ruta), "", rsaux1!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux1!ORDEN_SURTIDO), 0, rsaux1!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux1!Embarque), 0, rsaux1!Embarque)) + ", '" + IIf(IsNull(rsaux1!estacion), "", rsaux1!estacion) + "'," + CStr(IIf(IsNull(rsaux1!MULTIPLO), "", rsaux1!MULTIPLO)) + ")"
                                    rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                    var_cantidad_total = var_cantidad_total - rsaux1!MULTIPLO
                              Wend
                           Else
                              var_cantidad = var_cantidad_total
                              While var_cantidad > 0
                                    var_dia = CStr(Day(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                    var_mes = CStr(Month(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                    var_año = CStr(Year(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                    If Len(Trim(var_dia)) = 1 Then
                                       var_dia = "0" + var_dia
                                    End If
                                    If Len(Trim(var_mes)) = 1 Then
                                      var_mes = "0" + var_mes
                                    End If
                                    var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                    var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_2 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                    var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,MULTIPLO) "
                                    var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux1!source_header_number + "', " + CStr(rsaux1!delivery_id) + "," + CStr(rsaux1!delivery_detail_id) + ", " + CStr(rsaux1!organization_id) + ",'" + IIf(IsNull(rsaux1!subinventory), "", rsaux1!subinventory) + "', " + CStr(rsaux1!delivery_line_id) + "," + CStr(rsaux1!inventory_item_id) + ",'" + rsaux1!item_description + "','" + CStr(rsaux1!SOURCE_LINE_NUMBER) + "'," + CStr(1) + ",'" + rsaux1!released_status + "',"
                                    var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + rsaux1!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux1!collector_id), 0, rsaux1!collector_id)) + ",'" + IIf(IsNull(rsaux1!Name), "", rsaux1!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux1!ubicacion), "", rsaux1!ubicacion) + "','" + IIf(IsNull(rsaux1!establecimiento), "", rsaux1!establecimiento) + "','" + IIf(IsNull(rsaux1!nombre_Establecimiento), "", rsaux1!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux1!ORDENES), "", rsaux1!ORDENES) + "',"
                                    var_cadena = var_cadena + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux1!source_header_type_name), "", rsaux1!source_header_type_name) + "', '" + IIf(IsNull(rsaux1!source_document_id), "", rsaux1!source_document_id) + "','" + IIf(IsNull(rsaux1!pais), "", rsaux1!pais) + "','" + IIf(IsNull(rsaux1!estado), "", rsaux1!estado) + "', '" + IIf(IsNull(rsaux1!municipio), "", rsaux1!municipio) + "', '" + IIf(IsNull(rsaux1!ciudad), "", rsaux1!ciudad) + "', '" + IIf(IsNull(rsaux1!colonia), "", rsaux1!colonia) + "','" + IIf(IsNull(rsaux1!DIRECCION), "", rsaux1!DIRECCION) + "', '" + IIf(IsNull(rsaux1!cp), "", rsaux1!cp) + "',"
                                    var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!site_use_id), "", rsaux1!site_use_id) + "','" + IIf(IsNull(rsaux1!paqueteria), "", rsaux1!paqueteria) + "','" + IIf(IsNull(rsaux1!attribute8), "", rsaux1!attribute8) + "','" + IIf(IsNull(rsaux1!ATTRIBUTE9), "", rsaux1!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux1!TELEFONO), "", rsaux1!TELEFONO) + "','" + IIf(IsNull(rsaux1!Linea), "", rsaux1!Linea) + "','" + CStr(IIf(IsNull(rsaux1!ruta), "", rsaux1!ruta)) + "','" + IIf(IsNull(rsaux1!nombre_ruta), "", rsaux1!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux1!ORDEN_SURTIDO), 0, rsaux1!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux1!Embarque), 0, rsaux1!Embarque)) + ", '" + IIf(IsNull(rsaux1!estacion), "", rsaux1!estacion) + "'," + CStr(IIf(IsNull(rsaux1!MULTIPLO), "", rsaux1!MULTIPLO)) + ")"
                                    rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                    var_cantidad = var_cantidad - 1
                              Wend
                           End If
                        End If
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  rsaux1.Open "delete from tb_Temp_oracle_orden_surtido_aux_1", cnn, adOpenDynamic, adLockOptimistic
                  rsaux1.Open "select distinct source_header_number from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        var_lote = 1
                        var_contador = 0
                        rsaux2.Open "select * from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number = " + CStr(rsaux1!source_header_number) + " order by ubicacion", cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux2.EOF
                              If var_contador >= 50 Then
                                 var_lote = var_lote + 1
                                 var_contador = 0
                              End If
                              rsaux3.Open "SELECT * FROM tb_Temp_oracle_orden_surtido_aux_1 WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and delivery_detail_id = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 rsaux4.Open "UPDATE TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 SET SRC_REQUESTED_QUANTITY = SRC_REQUESTED_QUANTITY + " + CStr(rsaux2!src_requested_quantity) + " WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and DELIVERY_DETAIL_ID = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 var_dia = CStr(Day(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_mes = CStr(Month(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_año = CStr(Year(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 If Len(Trim(var_dia)) = 1 Then
                                    var_dia = "0" + var_dia
                                 End If
                                 If Len(Trim(var_mes)) = 1 Then
                                    var_mes = "0" + var_mes
                                 End If
                                 var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                 
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                 var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,LOTE) "
                                 var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux2!source_header_number + "', " + CStr(rsaux2!delivery_id) + "," + CStr(rsaux2!delivery_detail_id) + ", " + CStr(rsaux2!organization_id) + ",'" + IIf(IsNull(rsaux2!subinventory), "", rsaux2!subinventory) + "', " + CStr(rsaux2!delivery_line_id) + "," + CStr(rsaux2!inventory_item_id) + ",'" + rsaux2!item_description + "','" + CStr(rsaux2!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux2!src_requested_quantity) + ",'" + rsaux2!released_status + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!customer_name), "", rsaux2!customer_name) + "','" + rsaux2!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux2!collector_id), 0, rsaux2!collector_id)) + ",'" + IIf(IsNull(rsaux2!Name), "", rsaux2!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux2!ubicacion), "", rsaux2!ubicacion) + "','" + IIf(IsNull(rsaux2!establecimiento), "", rsaux2!establecimiento) + "','" + IIf(IsNull(rsaux2!nombre_Establecimiento), "", rsaux2!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux2!ORDENES), "", rsaux2!ORDENES) + "',"
                                 var_cadena = var_cadena + CStr(IIf(IsNull(rsaux2!CUST_ACCOUNT_ID), 0, rsaux2!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux2!source_header_type_name), "", rsaux2!source_header_type_name) + "', '" + IIf(IsNull(rsaux2!source_document_id), "", rsaux2!source_document_id) + "','" + IIf(IsNull(rsaux2!pais), "", rsaux2!pais) + "','" + IIf(IsNull(rsaux2!estado), "", rsaux2!estado) + "', '" + IIf(IsNull(rsaux2!municipio), "", rsaux2!municipio) + "', '" + IIf(IsNull(rsaux2!ciudad), "", rsaux2!ciudad) + "', '" + IIf(IsNull(rsaux2!colonia), "", rsaux2!colonia) + "','" + IIf(IsNull(rsaux2!DIRECCION), "", rsaux2!DIRECCION) + "', '" + IIf(IsNull(rsaux2!cp), "", rsaux2!cp) + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!site_use_id), "", rsaux2!site_use_id) + "','" + IIf(IsNull(rsaux2!paqueteria), "", rsaux2!paqueteria) + "','" + IIf(IsNull(rsaux2!attribute8), "", rsaux2!attribute8) + "','" + IIf(IsNull(rsaux2!ATTRIBUTE9), "", rsaux2!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux2!TELEFONO), "", rsaux2!TELEFONO) + "','" + IIf(IsNull(rsaux2!Linea), "", rsaux2!Linea) + "','" + CStr(IIf(IsNull(rsaux2!ruta), "", rsaux2!ruta)) + "','" + IIf(IsNull(rsaux2!nombre_ruta), "", rsaux2!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_SURTIDO), 0, rsaux2!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", '" + IIf(IsNull(rsaux2!estacion), "", rsaux2!estacion) + "'," + CStr(var_lote) + ")"
                                 rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux3.Close
                              If rsaux2!Linea <> "CATALOGOS" Then
                                 If rsaux2!MULTIPLO > 1 Then
                                    var_contador = var_contador + rsaux2!src_requested_quantity
                                 Else
                                    var_contador = var_contador + 1
                                 End If
                              End If
                              rsaux2.MoveNext
                        Wend
                        rsaux2.Close
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  rsaux1.Open "insert TB_TEMP_ORACLE_ORDEN_SURTIDO (inte_tem_consecutivo, segment1) values (" + CStr(var_consecutivo) + ",'---------')", cnn, adOpenDynamic, adLockOptimistic
                  rsaux1.Open "delete from TB_TEMP_ORACLE_ORDEN_SURTIDO where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and segment1 <> '---------'", cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.Open "insert into TB_TEMP_ORACLE_ORDEN_SURTIDO select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.Open "delete from TB_TEMP_ORACLE_ORDEN_SURTIDO where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and segment1 = '---------'", cnn, adOpenDynamic, adLockOptimistic
                  var_consecutivo_general = var_consecutivo
                  Call crea_tablas
                  If rsaux.State = 1 Then
                     rsaux.Close
                  End If
                  rsaux.Open "select distinct a.source_header_number from tb_Temp_oracle_orden_surtido_aux_1 a, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES  b where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and a.source_header_number = b.pedido", cnn, adOpenDynamic, adLockOptimistic
                  var_cadena_pedidos_diferencias = ""
                  While Not rsaux.EOF
                        strconsulta = "select sum(requested_quantity)  as cantidad from WSH_DELIVERABLES_V where source_header_number = ? AND RELEASED_STATUS = 'Y'"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(rsaux!source_header_number))
                             .Parameters.Append parametro
                        End With
                        Set rsaux10 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        
                        strconsulta = "SELECT SUM(SRC_REQUESTED_QUANTITY) AS CANTIDAD FROM XXVIA_TB_PEDIDOS_DIVIDIDOS WHERE SOURCE_HEADER_NUMBER = ?"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(rsaux!source_header_number))
                             .Parameters.Append parametro
                        End With
                        Set rsaux11 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                     
                     
                        If rsaux11!cantidad <> rsaux10!cantidad Then
                           If var_cadena_pedidos_diferencias = "" Then
                              var_cadena_pedidos_diferencias = CStr(rsaux!source_header_number)
                           Else
                              var_cadena_pedidos_diferencias = var_cadena_pedidos_diferencias + ", " + CStr(rsaux1!source_header_number)
                           End If
                        End If
                        rsaux10.Close
                        rsaux11.Close
                        rsaux.MoveNext
                  Wend
                  rsaux.Close
                  If var_cadena_pedidos_diferencias = "" Then
                     If var_imprime_pedidos = 1 Then
                        rsaux.Open "select distinct a.source_header_number, a.lote, orden_pedido from tb_Temp_oracle_orden_surtido_aux_1 a, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES  b where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and a.source_header_number = b.pedido order by orden_pedido", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           While Not rsaux.EOF
                                 x = 1
                                 If x = 1 Then
                                    
                                    strconsulta = "select shipping_method_code, packing_instructions from oe_order_headers_all where order_number = ?"
                                    With comandoORA
                                         .ActiveConnection = cnnoracle_4
                                         .CommandType = adCmdText
                                         .CommandText = strconsulta
                                         Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rsaux!source_header_number))
                                         .Parameters.Append parametro
                                    End With
                                    Set rs = comandoORA.execute
                                    Set comandoORA = Nothing
                                    Set parametro = Nothing
                                    
                                    var_paqueteria = ""
                                    If Not rs.EOF Then
                                       VAR_COMENTARIOS = IIf(IsNull(rs!packing_instructions), "", rs!packing_instructions)
                                       var_tipo_metodo = IIf(IsNull(rs(0).Value), "", rs(0).Value)
                                       If var_tipo_metodo <> "" Then
                                          
                                          strconsulta = "SELECT description FROM fnd_lookup_values where lookup_type = 'SHIP_METHOD' and lookup_code = ? AND LANGUAGE = 'ESA'"
                                          With comandoORA
                                               .ActiveConnection = cnnoracle_4
                                               .CommandType = adCmdText
                                               .CommandText = strconsulta
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_tipo_metodo)
                                               .Parameters.Append parametro
                                          End With
                                          Set rsaux1 = comandoORA.execute
                                          Set comandoORA = Nothing
                                          Set parametro = Nothing
                                          
                                          If Not rsaux1.EOF Then
                                             var_paqueteria = IIf(IsNull(rsaux1(0).Value), "", rsaux1(0).Value)
                                          End If
                                          rsaux1.Close
                                       End If
                                    End If
                                    rs.Close
                                    
                                    
                                    
                                    strconsulta = "SELECT  HPS.party_site_id as tel, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1, OHA.packing_instructions from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.invoice_TO_ORG_ID AND oha.order_number = ? AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                                    With comandoORA
                                         .ActiveConnection = cnnoracle_4
                                         .CommandType = adCmdText
                                         .CommandText = strconsulta
                                         Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rsaux!source_header_number))
                                         .Parameters.Append parametro
                                    End With
                                    Set rsaux6 = comandoORA.execute
                                    Set comandoORA = Nothing
                                    Set parametro = Nothing
                                    
                                    If Not rsaux6.EOF Then
                                    
                                       
                                       strconsulta = "SELECT  HPS.party_site_id as tel, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1, OHA.packing_instructions from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.ship_TO_ORG_ID AND oha.order_number = ? AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                                       With comandoORA
                                            .ActiveConnection = cnnoracle_4
                                            .CommandType = adCmdText
                                            .CommandText = strconsulta
                                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, IIf(IsNull(rsaux!source_header_number), "", rsaux!source_header_number))
                                            .Parameters.Append parametro
                                       End With
                                       Set rsaux5 = comandoORA.execute
                                       Set comandoORA = Nothing
                                       Set parametro = Nothing
                                       
                                       'rsaux5.Open "SELECT  HPS.party_site_id as tel, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1, OHA.packing_instructions from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.ship_TO_ORG_ID AND oha.order_number = '" + IIf(IsNull(rsaux!source_header_number), "", rsaux!source_header_number) + "' AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                       If Not rsaux5.EOF Then
                                          var_nombre = IIf(IsNull(rsaux5!customer_name), "", rsaux5!customer_name)
                                          var_tel = IIf(IsNull(rsaux5!tel), 0, rsaux5!tel)
                                          VAR_DIRECCION = IIf(IsNull(rsaux5!calle), "", rsaux5!calle) + " " + IIf(IsNull(rsaux5!numero), "", rsaux5!numero)
                                          VAR_COLONIA = IIf(IsNull(rsaux5!colonia), "", rsaux5!colonia)
                                          var_ciudad = IIf(IsNull(rsaux5!ciudad), "", rsaux5!ciudad)
                                          VAR_MUNICIPIO = IIf(IsNull(rsaux5!municipio), "", rsaux5!municipio)
                                          var_estado = IIf(IsNull(rsaux5!estado), "", rsaux5!estado)
                                          var_pais = IIf(IsNull(rsaux5!pais), "", rsaux5!pais)
                                          VAR_CP = IIf(IsNull(rsaux5!cp), "", rsaux5!cp)
                                          rsaux5.Close
                                       Else
                                          rsaux5.Close
                                          var_nombre = IIf(IsNull(rsaux6!customer_name), "", rsaux6!customer_name)
                                          var_tel = IIf(IsNull(rsaux6!tel), 0, rsaux6!tel)
                                          VAR_DIRECCION = IIf(IsNull(rsaux6!calle), "", rsaux6!calle) + " " + IIf(IsNull(rsaux6!numero), "", rsaux6!numero)
                                          VAR_COLONIA = IIf(IsNull(rsaux6!colonia), "", rsaux6!colonia)
                                          var_ciudad = IIf(IsNull(rsaux6!ciudad), "", rsaux6!ciudad)
                                          VAR_MUNICIPIO = IIf(IsNull(rsaux6!municipio), "", rsaux6!municipio)
                                          var_estado = IIf(IsNull(rsaux6!estado), "", rsaux6!estado)
                                          var_pais = IIf(IsNull(rsaux6!pais), "", rsaux6!pais)
                                          VAR_CP = IIf(IsNull(rsaux6!cp), "", rsaux6!cp)
                                       End If
                                    Else
                                       var_tel = 0
                                       VAR_DIRECCION = ""
                                       VAR_COLONIA = ""
                                       var_ciudad = ""
                                       VAR_MUNICIPIO = ""
                                       var_estado = ""
                                       var_pais = ""
                                       VAR_CP = ""
                                    End If
                                    rsaux6.Close
                                    If var_tel > 0 Then
                                       
                                       strconsulta = "select Phone_Number from hz_contact_points where owner_table_id = ?"
                                       With comandoORA
                                            .ActiveConnection = cnnoracle_4
                                            .CommandType = adCmdText
                                            .CommandText = strconsulta
                                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(var_tel))
                                            .Parameters.Append parametro
                                       End With
                                       Set rsaux6 = comandoORA.execute
                                       Set comandoORA = Nothing
                                       Set parametro = Nothing
                                       
                                       
                                       'rsaux6.Open "select Phone_Number from hz_contact_points where owner_table_id = " + CStr(var_tel), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                       If Not rsaux6.EOF Then
                                          var_telefono = CStr(IIf(IsNull(rsaux6(0).Value), "", rsaux6(0).Value))
                                       Else
                                          var_telefono = ""
                                       End If
                                       rsaux6.Close
                                    Else
                                       var_telefono = ""
                                    End If
                                       
                                    
                                    If IsNumeric(txt_total_volumen) Then
                                       var_cubicaje_EMBARQUE = CDbl(Me.txt_total_volumen)
                                    Else
                                       var_cubicaje_EMBARQUE = 0
                                    End If
                                    
                                    rsaux2.Open "SELECT * FROM tb_oracle_pedidos_asignados_embarques WHERE PEDIDO = " + CStr(rsaux(0).Value), cnn, adOpenDynamic, adLockOptimistic
                                    If Not rsaux2.EOF Then
                                       var_transporte = ""
                                       rsaux3.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                       If Not rsaux3.EOF Then
                                          var_transporte = IIf(IsNull(rsaux3!vehiculo), "", rsaux3!vehiculo)
                                       End If
                                       rsaux3.Close
                                       var_cadena = "UPDATE tb_Temp_oracle_orden_surtido SET EMBARQUE = " + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", CUBICAJE = " + CStr(var_cubicaje_EMBARQUE) + " , ORDEN_SURTIDO = " + CStr(IIf(IsNull(rsaux2!orden_pedido), 0, rsaux2!orden_pedido)) + ", ANDEN = '" + CStr(IIf(IsNull(rsaux2!estacion), 0, rsaux2!estacion)) + "', TRANSPORTE = '" + var_transporte + "',"
                                       var_cadena = var_cadena + " pais= '" + var_pais + "', estado = '" + var_estado + "', municipio = '" + VAR_MUNICIPIO + "', ciudad = '" + var_ciudad + "', colonia = '" + VAR_COLONIA + "', direccion = '" + VAR_DIRECCION + "', cp = '" + VAR_CP + "', paqueteria = '" + var_paqueteria + "'"
                                       var_cadena = var_cadena + " WHERE SOURCE_HEADER_NUMBER = " + CStr(rsaux(0).Value) + " AND inte_tem_consecutivo = " + CStr(var_consecutivo)
                                       rsaux3.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                    End If
                                    rsaux2.Close
                                 
                                    
                                    
                                    x = 1
                                    If x = 1 Then
                                       Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA.rpt")
                                       reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.SOURCE_HEADER_NUMBER} = '" + rsaux(0).Value + "' and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_ORACLE_ORDEN_SURTIDO.LOTE} = " + CStr(rsaux(1).Value)
                                       For ntablas = 1 To reporte.Database.Tables.Count
                                           reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                       Next ntablas
                                       reporte.PrintOut False
                                       Set reporte = Nothing
                                    Else
                                    
                                       Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA.rpt")
                                       reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.SOURCE_HEADER_NUMBER} = '" + rsaux(0).Value + "' and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_ORACLE_ORDEN_SURTIDO.LOTE} = " + CStr(rsaux(1).Value)
                                       frmvistasprevias.cr.ReportSource = reporte
                                       For ntablas = 1 To reporte.Database.Tables.Count
                                           reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                       Next ntablas
                                       frmvistasprevias.cr.ViewReport
                                       frmvistasprevias.Caption = "Ordenes de surtido pendientes de empacar o facturar"
                                       frmvistasprevias.Show 1
                                       Set reporte = Nothing
                                    End If
                                 End If
                                 rsaux.MoveNext
                           Wend
                        End If
                        rsaux.Close
                     Else
                        MsgBox "No se pueden imprimir los pedidos, vuelva a intentar la impresión", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "No se pueden imprimir los pedidos, vuelva a intentar la impresión", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "No se pueden imprimir los pedidos, vuelva a intentar la impresión", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "No existen ordenes de surtido", vbOKOnly, "ATENCION"
            End If
            If rs.State = 1 Then
               rs.Close
            End If
            If var_consecutivo Then
               rs.Open "delete from tb_Temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(IIf(IsNull(var_consecutivo), 0, var_consecutivo)), cnn, adOpenDynamic, adLockOptimistic
            End If
         Else
            'MsgBox "Número superior incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "No existen pedidos a imprimir"
      End If
End Sub

Private Sub cmd_imprimir_nuevo_metodo_divisiones_Click()
'GoTo x:
      If rs.State = 1 Then
         rs.Close
      End If
      If rsaux.State = 1 Then
         rsaux.Close
      End If
      If rsaux1.State = 1 Then
         rsaux1.Close
      End If
      If rsaux2.State = 1 Then
         rsaux2.Close
      End If
      If rsaux3.State = 1 Then
         rsaux3.Close
      End If
      If rsaux4.State = 1 Then
         rsaux4.Close
      End If
      If rsaux5.State = 1 Then
         rsaux5.Close
      End If
      If rsaux6.State = 1 Then
         rsaux6.Close
      End If
      If rsaux7.State = 1 Then
         rsaux7.Close
      End If
      If rsaux8.State = 1 Then
         rsaux8.Close
      End If
      If rsaux9.State = 1 Then
         rsaux9.Close
      End If
     
      If Me.lv_pedidos.ListItems.Count > 0 Then
         var_si = MsgBox("Desea imprimir las ordenes de surtido?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            If rs.State = 1 Then
               rs.Close
            End If
            var_Cadena_pedidos = ""
            For var_j = 1 To Me.lv_pedidos.ListItems.Count
                Me.lv_pedidos.ListItems.Item(var_j).Selected = True
                If Me.lv_pedidos.selectedItem <> "10000000" Then
                   If var_Cadena_pedidos = "" Then
                      var_Cadena_pedidos = Me.lv_pedidos.selectedItem
                   Else
                      var_Cadena_pedidos = var_Cadena_pedidos + "," + Me.lv_pedidos.selectedItem
                   End If
                End If
            Next var_j
            'var_cadena_pedidos = "105208"
            rs.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_cadena = "SELECT to_char(a.LAST_UPDATE_DATE,'day') DIA_SEMANA, CAT.LINEA, HCSU.SITE_USE_ID, HCAS.CUST_ACCOUNT_ID, OHA.SHIP_TO_ORG_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id, a.LAST_UPDATE_DATE,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,c.description,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, a.source_header_type_name, oha.source_document_id, C.ATTRIBUTE2, oha.attribute8, oha.attribute9, j.NAME as nombre_ruta, j.salesrep_id as clave_ruta from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, xxvia_vw_articulos_cat cat, XXVIA_VENDEDORES j Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND "
            var_cadena = var_cadena + " to_number(source_header_number)  IN (" + var_Cadena_pedidos + ")"
            var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID "
            var_cadena = var_cadena + " AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND released_status = 'Y' AND A.inventory_item_id  = cat.item_id AND A.ORGANIZATION_ID = Cat.organization_id and oha.salesrep_id = j.salesrep_id "
            rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
'--------------------------
            If Not rs.EOF Then
               cnn.BeginTrans
               rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM tb_Temp_oracle_orden_surtido", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
               Else
                  var_consecutivo = 1
               End If
               rsaux.Close
               rsaux1.Open "insert into tb_Temp_oracle_orden_surtido(inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
               cnn.CommitTrans
               While Not rs.EOF
                     var_establecimiento = rs!SHIP_TO_ORG_ID
                     var_clave_cliente = rs!site_use_id
                     rsaux.Open "SELECT csu.site_use_id AS VCHA_ESB_ESTABLECIMIENTO_ID, ps.party_site_number, lo.address1 AS VCHA_eSB_NOMBRE FROM hz_cust_site_uses_all csu, hz_cust_acct_sites_all cas, hz_party_sites ps, hz_locations lo Where csu.cust_acct_site_id = cas.cust_acct_site_id AND cas.party_site_id = ps.party_site_id AND ps.location_id = lo.location_id AND csu.site_use_code = 'SHIP_TO' AND csu.LOCATION = ps.party_site_number and csu.site_use_id = " + CStr(var_establecimiento), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        VAR_NOMBRE_ESTABLECIMIENTO = IIf(IsNull(rsaux!party_site_number), "", rsaux!party_site_number) + " " + IIf(IsNull(rsaux!vcha_esb_nombre), "", rsaux!vcha_esb_nombre)
                     Else
                        VAR_NOMBRE_ESTABLECIMIENTO = ""
                     End If
                     rsaux.Close
                     
                     rsaux.Open "SELECT csu.site_use_id AS VCHA_ESB_ESTABLECIMIENTO_ID, ps.party_site_number, lo.address1 AS VCHA_eSB_NOMBRE FROM hz_cust_site_uses_all csu, hz_cust_acct_sites_all cas, hz_party_sites ps, hz_locations lo Where csu.cust_acct_site_id = cas.cust_acct_site_id AND cas.party_site_id = ps.party_site_id AND ps.location_id = lo.location_id AND csu.site_use_code = 'BILL_TO' AND csu.LOCATION = ps.party_site_number and csu.site_use_id = " + CStr(var_clave_cliente), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        VAR_CLAVE_CLIENTE_BCP = IIf(IsNull(rsaux!party_site_number), "", rsaux!party_site_number)
                     Else
                        VAR_CLAVE_CLIENTE_BCP = ""
                     End If
                     rsaux.Close
                     
                     
                     
                     var_dia = CStr(Day(CDate(rs!LAST_UPDATE_DATE)))
                     var_mes = CStr(Month(CDate(rs!LAST_UPDATE_DATE)))
                     var_año = CStr(Year(CDate(rs!LAST_UPDATE_DATE)))
                     If Len(Trim(var_dia)) = 1 Then
                        var_dia = "0" + var_dia
                     End If
                     If Len(Trim(var_mes)) = 1 Then
                        var_mes = "0" + var_mes
                     End If
                     var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                     rsaux1.Open "select * from tb_oracle_multiplos where segment1 = '" + IIf(IsNull(rs!SEGMENT1), "", rs!SEGMENT1) + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux1.EOF Then
                        VAR_MULTIPLO = IIf(IsNull(rsaux1!MULTIPLO), 1, rsaux1!MULTIPLO)
                     Else
                        VAR_MULTIPLO = 1
                     End If
                     rsaux1.Close
'''''
                     If rs!source_header_type_name = "VIA_PEDIDO_INTERNO" Then
                        rsaux1.Open "SELECT * FROM TB_ORACLE_ARTICULOS_MOTOR_LOGISTICO WHERE CODIGO = '" + rs!SEGMENT1 + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux1.EOF Then
                           strconsulta = "SELECT secondary_inventory_name, A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rs!source_document_id)
                                .Parameters.Append parametro
                           End With
                           Set rsaux8 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           If rsaux8.EOF Then
                              var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                           Else
                              var_almacen = rsaux8!secondary_inventory_name
                              rsaux9.Open "SELECT * FROM TB_ORACLE_UBICACIONES_MOTOR_LOGISTICO WHERE CLAVE = '" + var_almacen + "' AND CODIGO = '" + rs!SEGMENT1 + "'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux9.EOF Then
                                 var_ubicacion = ""
                                 If Me.cmb_dia.Text = "Lunes" Then
                                    var_ubicacion = rsaux9!ubicacion_1
                                 End If
                                 If Me.cmb_dia.Text = "Martes" Then
                                    var_ubicacion = rsaux9!ubicacion_2
                                 End If
                                 If Me.cmb_dia.Text = "Miercoles" Then
                                    var_ubicacion = rsaux9!ubicacion_3
                                 End If
                                 If Me.cmb_dia.Text = "Jueves" Then
                                    var_ubicacion = rsaux9!ubicacion_4
                                 End If
                                 If Me.cmb_dia.Text = "Viernes" Then
                                    var_ubicacion = rsaux9!ubicacion_5
                                 End If
                                 If Me.cmb_dia.Text = "Sabado" Then
                                    var_ubicacion = rsaux9!ubicacion_6
                                 End If
                                 If IIf(IsNull(var_ubicacion), "", var_ubicacion) = "" Then
                                    var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                                 End If
                              Else
                                 var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                              End If
                              rsaux9.Close
                           End If
                           rsaux8.Close
                        Else
                           var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                        End If
                        rsaux1.Close
                     Else
                        var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                     End If
                     
                     
'''''
                     var_cadena = "insert into tb_Temp_oracle_orden_surtido(INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER, DELIVERY_ID, DELIVERY_DETAIL_ID, ORGANIZATION_ID, SUBINVENTORY, DELIVERY_LINE_ID, INVENTORY_ITEM_ID, ITEM_DESCRIPTION, SOURCE_LINE_NUMBER, src_requested_quantity, RELEASED_STATUS, CUSTOMER_NAME, SEGMENT1, COLLECTOR_ID, NAME, date_requested, UBICACION, ESTABLECIMIENTO, NOMBRE_ESTABLECIMIENTO, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, SITE_USE_ID, attribute8, attribute9, LINEA, RUTA, NOMBRE_RUTA,MULTIPLO)  values "
                     var_cadena = var_cadena + " (" + CStr(var_consecutivo) + ", '" + IIf(IsNull(rs!source_header_number), "", rs!source_header_number) + "', " + CStr(IIf(IsNull(rs!delivery_id), 0, rs!delivery_id)) + ", " + CStr(IIf(IsNull(rs!delivery_detail_id), 0, rs!delivery_detail_id)) + ", " + CStr(IIf(IsNull(rs!organization_id), 0, rs!organization_id)) + ", '" + IIf(IsNull(rs!subinventory), "", rs!subinventory) + "', " + CStr(IIf(IsNull(rs!delivery_line_id), 0, rs!delivery_line_id)) + ", " + CStr(IIf(IsNull(rs!inventory_item_id), "", rs!inventory_item_id)) + ", '" + IIf(IsNull(rs!Description), "", rs!Description) + "', '" + IIf(IsNull(rs!SOURCE_LINE_NUMBER), "", rs!SOURCE_LINE_NUMBER) + "', " + CStr(IIf(IsNull(rs!requested_quantity), 0, rs!requested_quantity)) + ", '" + IIf(IsNull(rs!released_status), "", rs!released_status) + "', '" + VAR_CLAVE_CLIENTE_BCP + " " + Replace(IIf(IsNull(rs!customer_name), "", rs!customer_name), "'", " ") + "', '" + IIf(IsNull(rs!SEGMENT1), "", rs!SEGMENT1) + "'"
                     'var_cadena = var_cadena + ", 0,''," + var_fecha + ",'" + IIf(IsNull(rs!ATTRIBUTE2), "", rs!ATTRIBUTE2) + "','" + CStr(VAR_ESTABLECIMIENTO) + "','" + Replace(VAR_NOMBRE_ESTABLECIMIENTO, "'", " ") + "'," + CStr(rs!CUST_ACCOUNT_ID) + ",'" + rs!source_header_type_name + "','" + CStr(IIf(IsNull(rs!source_document_id), "", rs!source_document_id)) + "'," + CStr(IIf(IsNull(rs!SITE_USE_ID), 0, rs!SITE_USE_ID)) + ",'" + IIf(IsNull(rs!attribute8), "", rs!attribute8) + "','" + IIf(IsNull(rs!ATTRIBUTE9), "", rs!ATTRIBUTE9) + "','" + IIf(IsNull(rs!linea), "", rs!linea) + "','" + CStr(rs!CLAVE_RUTA) + "','" + IIf(IsNull(rs!nombre_ruta), "", rs!nombre_ruta) + "'," + CStr(VAR_MULTIPLO) + ")"
                     var_cadena = var_cadena + ", 0,''," + var_fecha + ",'" + var_ubicacion + "','" + CStr(var_establecimiento) + "','" + Replace(VAR_NOMBRE_ESTABLECIMIENTO, "'", " ") + "'," + CStr(rs!CUST_ACCOUNT_ID) + ",'" + rs!source_header_type_name + "','" + CStr(IIf(IsNull(rs!source_document_id), "", rs!source_document_id)) + "'," + CStr(IIf(IsNull(rs!site_use_id), 0, rs!site_use_id)) + ",'" + IIf(IsNull(rs!attribute8), "", rs!attribute8) + "','" + IIf(IsNull(rs!ATTRIBUTE9), "", rs!ATTRIBUTE9) + "','" + IIf(IsNull(rs!Linea), "", rs!Linea) + "','" + CStr(rs!CLAVE_RUTA) + "','" + IIf(IsNull(rs!nombre_ruta), "", rs!nombre_ruta) + "'," + CStr(VAR_MULTIPLO) + ")"
                     rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               rsaux1.Open "delete from tb_Temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number is null", cnn, adOpenDynamic, adLockOptimistic
               If rsaux1.State = 1 Then
                  rsaux1.Close
               End If
               
               var_cadena_pedidos_diferencias = ""
               rsaux1.Open "select source_header_number, sum(src_requested_quantity) as cantidad from tb_Temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo) + " group by source_header_number", cnn, adOpenDynamic, adLockOptimistic
               While Not rsaux1.EOF
                     strconsulta = "select sum(requested_quantity)  as cantidad from WSH_DELIVERABLES_V where source_header_number = ? AND RELEASED_STATUS = 'Y'"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(rsaux1!source_header_number))
                          .Parameters.Append parametro
                     End With
                     Set rsaux10 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     If rsaux1!cantidad <> rsaux10!cantidad Then
                        If var_cadena_pedidos_diferencias = "" Then
                           var_cadena_pedidos_diferencias = CStr(rsaux1!source_header_number)
                        Else
                           var_cadena_pedidos_diferencias = var_cadena_pedidos_diferencias + ", " + CStr(rsaux1!source_header_number)
                        End If
                     End If
                     rsaux10.Close
                     rsaux1.MoveNext
               Wend
               rsaux1.Close
               
               If var_cadena_pedidos_diferencias = "" Then
                  rsaux1.Open "select min(DATE_REQUESTED) as DATE_REQUESTED, source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME, RUTA, NOMBRE_RUTA, SUM(SRC_REQUESTED_QUANTITY) AS CANTIDAD from tb_Temp_oracle_orden_surtido where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " group by source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME, RUTA, NOMBRE_RUTA", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        If rsaux1!source_header_type_name = "VIA_PEDIDO_INTERNO" Or rsaux1!source_header_type_name = "TEX_PEDIDO_INTERNO" Then
                           If var_pedido_tienda = 0 Then
                              If rsaux2.State = 1 Then
                                 rsaux2.Close
                              End If
                              rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(rsaux1!source_document_id) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux2.EOF Then
                                 If rsaux4.State = 1 Then
                                    rsaux4.Close
                                 End If
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                                 rsaux4.Close
                              Else
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                                 rsaux4.Close
                              End If
                              rsaux2.Close
                           Else
                              rsaux2.Open "select a.attribute8, B.description from oe_order_headers_all a, MTL_SECONDARY_INVENTORIES b where order_number = " + CStr(rsaux1!source_header_number) + " and a.attribute8 = b.secondary_inventory_name", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux2.EOF Then
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                                 rsaux4.Close
                              Else
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                                 rsaux4.Close
                              End If
                              rsaux2.Close
                           End If
                        Else
                           rsaux6.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                           VAR_PROVEEDOR = rsaux6!collector_id
                           VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux6!Name), "", rsaux6!Name)
                           rsaux6.Close
                        End If
                        var_año_str = CStr(Year(rsaux1!DATE_REQUESTED))
                        If Len(var_año_str) < 2 Then
                           var_año_str = "20" + var_año_str
                        End If
                        var_mes_str = CStr(Month(rsaux1!DATE_REQUESTED))
                        If Len(var_mes_str) < 2 Then
                           var_mes_str = "0" + var_mes_str
                        End If
                        var_dia_str = CStr(Day(rsaux1!DATE_REQUESTED))
                        If Len(var_dia_str) < 2 Then
                           var_dia_str = "0" + var_dia_str
                        End If
                        var_hora_str = CStr(Hour(rsaux1!DATE_REQUESTED))
                        If Len(var_hora_str) < 2 Then
                           var_hora_str = "0" + var_hora_str
                        End If
                        VAR_MINUTO_STR = CStr(Minute(rsaux1!DATE_REQUESTED))
                        If Len(VAR_MINUTO_STR) < 2 Then
                           VAR_MINUTO_STR = "0" + VAR_MINUTO_STR
                        End If
                        VAR_SEGUNDO_STR = CStr(Second(rsaux1!DATE_REQUESTED))
                        If Len(VAR_SEGUNDO_STR) < 2 Then
                           VAR_SEGUNDO_STR = "0" + VAR_SEGUNDO_STR
                        End If
                        var_fecha_pedido = var_año_str + "-" + var_mes_str + "-" + var_dia_str + " " + var_hora_str + ":" + VAR_MINUTO_STR + ":" + VAR_SEGUNDO_STR
                        rsaux6.Open "update tb_Temp_oracle_orden_surtido set COLLECTOR_ID = " + CStr(VAR_PROVEEDOR) + ", NAME = '" + VAR_NOMBRE_PROVEEDOR + "' where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " and CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and source_header_number = " + CStr(rsaux1!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                        rsaux6.Open "SELECT * FROM XXVIA_TB_ORDENES_GRAFICA WHERE ORGANIZACION = " + var_unidad_organizacional + " AND PEDIDO = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If rsaux6.EOF Then
                           rsaux5.Open "INSERT INTO XXVIA_TB_ORDENES_GRAFICA (ORGANIZACION, PEDIDO, PROVEEDOR_ID, NOMBRE_PROVEEDOR, CLIENTE, ESTABLECIMIENTO, CANTIDAD, FECHA, TIPO_PEDIDO, CANTIDAD_LEIDA, FECHA_PEDIDO, RUTA, NOMBRE_RUTA) VALUES (" + var_unidad_organizacional + ", " + CStr(rsaux1!source_header_number) + "," + CStr(VAR_PROVEEDOR) + ",'" + VAR_NOMBRE_PROVEEDOR + "','" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "'," + CStr(rsaux1!cantidad) + ", '" + CStr(rsaux1!DATE_REQUESTED) + "','" + rsaux1!source_header_type_name + "',0, TO_DATE('" + var_fecha_pedido + "','YYYY-MM-DD HH24:MI:SS'),'" + rsaux1!ruta + "', '" + rsaux1!nombre_ruta + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        Else
                           rsaux5.Open "UPDATE XXVIA_TB_ORDENES_GRAFICA SET RUTA = '" + rsaux1!ruta + "', NOMBRE_RUTA = '" + rsaux1!nombre_ruta + "' WHERE PEDIDO = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        End If
                        rsaux6.Close
                        rsaux6.Open "select * from tb_oracle_pedidos_asignados_embarques where pedido = " + CStr(rsaux1!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                        If rsaux6.EOF Then
                           rsaux5.Open "insert into tb_oracle_pedidos_asignados_embarques (AGENTE, NOMBRE_AGENTE, PEDIDO, CLIENTE, PIEZAS, embarque, dia,  mes, AÑO, ORGANIZACION) values (" + CStr(VAR_PROVEEDOR) + ",'" + VAR_NOMBRE_PROVEEDOR + "'," + CStr(rsaux1!source_header_number) + ",'" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "', " + CStr(rsaux1!cantidad) + ",0," + var_dia_str + "," + var_mes_str + "," + var_año_str + "," + CStr(var_unidad_organizacional) + ")", cnn, adOpenDynamic, adLockOptimistic
                        End If
                        rsaux6.Close
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  rsaux1.Open "select min(DATE_REQUESTED) as DATE_REQUESTED, source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME, SUM(SRC_REQUESTED_QUANTITY) AS CANTIDAD from tb_Temp_oracle_orden_surtido where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " AND (LINEA NOT IN ('CATALOGOS','CATALOGO','POP') OR LINEA IS NULL) group by source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME"
                  While Not rsaux1.EOF
                        rsaux5.Open "UPDATE XXVIA_TB_ORDENES_GRAFICA SET CANTIDAD_SIN_CATALOGOS = " + CStr(IIf(IsNull(rsaux1!cantidad), 0, rsaux1!cantidad)) + " WHERE PEDIDO = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  rsaux1.Open "DELETE from tb_Temp_oracle_orden_surtido_aux_2", cnn, adOpenDynamic, adLockOptimistic
                  rsaux1.Open "SELECT * FROM tb_Temp_oracle_orden_surtido where inte_tem_consecutivo =  " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        If rsaux1!Linea = "CATALOGOS" Or rsaux1!Linea = "CATALOGO" Or rsaux1!Linea = "POP" Or rsaux1!Linea = "EMPAQUE" Then
                           var_dia = CStr(Day(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                           var_mes = CStr(Month(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                           var_año = CStr(Year(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                           If Len(Trim(var_dia)) = 1 Then
                              var_dia = "0" + var_dia
                           End If
                           If Len(Trim(var_mes)) = 1 Then
                              var_mes = "0" + var_mes
                           End If
                           var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                           var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_2 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                           var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,MULTIPLO) "
                           var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux1!source_header_number + "', " + CStr(rsaux1!delivery_id) + "," + CStr(rsaux1!delivery_detail_id) + ", " + CStr(rsaux1!organization_id) + ",'" + IIf(IsNull(rsaux1!subinventory), "", rsaux1!subinventory) + "', " + CStr(rsaux1!delivery_line_id) + "," + CStr(rsaux1!inventory_item_id) + ",'" + rsaux1!item_description + "','" + CStr(rsaux1!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux1!src_requested_quantity) + ",'" + rsaux1!released_status + "',"
                           var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + rsaux1!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux1!collector_id), 0, rsaux1!collector_id)) + ",'" + IIf(IsNull(rsaux1!Name), "", rsaux1!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux1!ubicacion), "", rsaux1!ubicacion) + "','" + IIf(IsNull(rsaux1!establecimiento), "", rsaux1!establecimiento) + "','" + IIf(IsNull(rsaux1!nombre_Establecimiento), "", rsaux1!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux1!ORDENES), "", rsaux1!ORDENES) + "',"
                           var_cadena = var_cadena + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux1!source_header_type_name), "", rsaux1!source_header_type_name) + "', '" + IIf(IsNull(rsaux1!source_document_id), "", rsaux1!source_document_id) + "','" + IIf(IsNull(rsaux1!pais), "", rsaux1!pais) + "','" + IIf(IsNull(rsaux1!estado), "", rsaux1!estado) + "', '" + IIf(IsNull(rsaux1!municipio), "", rsaux1!municipio) + "', '" + IIf(IsNull(rsaux1!ciudad), "", rsaux1!ciudad) + "', '" + IIf(IsNull(rsaux1!colonia), "", rsaux1!colonia) + "','" + IIf(IsNull(rsaux1!DIRECCION), "", rsaux1!DIRECCION) + "', '" + IIf(IsNull(rsaux1!cp), "", rsaux1!cp) + "',"
                           var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!site_use_id), "", rsaux1!site_use_id) + "','" + IIf(IsNull(rsaux1!paqueteria), "", rsaux1!paqueteria) + "','" + IIf(IsNull(rsaux1!attribute8), "", rsaux1!attribute8) + "','" + IIf(IsNull(rsaux1!ATTRIBUTE9), "", rsaux1!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux1!TELEFONO), "", rsaux1!TELEFONO) + "','" + IIf(IsNull(rsaux1!Linea), "", rsaux1!Linea) + "','" + CStr(IIf(IsNull(rsaux1!ruta), "", rsaux1!ruta)) + "','" + IIf(IsNull(rsaux1!nombre_ruta), "", rsaux1!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux1!ORDEN_SURTIDO), 0, rsaux1!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux1!Embarque), 0, rsaux1!Embarque)) + ", '" + IIf(IsNull(rsaux1!estacion), "", rsaux1!estacion) + "'," + CStr(IIf(IsNull(rsaux1!MULTIPLO), "", rsaux1!MULTIPLO)) + ")"
                           rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        Else
                           var_cantidad_total = rsaux1!src_requested_quantity
                           If rsaux1!MULTIPLO > 1 Then
                              While var_cantidad_total > 0
                                    If var_cantidad_total < rsaux1!MULTIPLO Then
                                       var_cantidad = var_cantidad_total
                                    Else
                                       var_cantidad = rsaux1!MULTIPLO
                                    End If
                                    
                                    var_dia = CStr(Day(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                    var_mes = CStr(Month(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                    var_año = CStr(Year(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                    If Len(Trim(var_dia)) = 1 Then
                                       var_dia = "0" + var_dia
                                    End If
                                    If Len(Trim(var_mes)) = 1 Then
                                      var_mes = "0" + var_mes
                                    End If
                                    
                                    var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                    var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_2 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                    var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,MULTIPLO) "
                                    var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux1!source_header_number + "', " + CStr(rsaux1!delivery_id) + "," + CStr(rsaux1!delivery_detail_id) + ", " + CStr(rsaux1!organization_id) + ",'" + IIf(IsNull(rsaux1!subinventory), "", rsaux1!subinventory) + "', " + CStr(rsaux1!delivery_line_id) + "," + CStr(rsaux1!inventory_item_id) + ",'" + rsaux1!item_description + "','" + CStr(rsaux1!SOURCE_LINE_NUMBER) + "'," + CStr(var_cantidad) + ",'" + rsaux1!released_status + "',"
                                    var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + rsaux1!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux1!collector_id), 0, rsaux1!collector_id)) + ",'" + IIf(IsNull(rsaux1!Name), "", rsaux1!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux1!ubicacion), "", rsaux1!ubicacion) + "','" + IIf(IsNull(rsaux1!establecimiento), "", rsaux1!establecimiento) + "','" + IIf(IsNull(rsaux1!nombre_Establecimiento), "", rsaux1!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux1!ORDENES), "", rsaux1!ORDENES) + "',"
                                    var_cadena = var_cadena + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux1!source_header_type_name), "", rsaux1!source_header_type_name) + "', '" + IIf(IsNull(rsaux1!source_document_id), "", rsaux1!source_document_id) + "','" + IIf(IsNull(rsaux1!pais), "", rsaux1!pais) + "','" + IIf(IsNull(rsaux1!estado), "", rsaux1!estado) + "', '" + IIf(IsNull(rsaux1!municipio), "", rsaux1!municipio) + "', '" + IIf(IsNull(rsaux1!ciudad), "", rsaux1!ciudad) + "', '" + IIf(IsNull(rsaux1!colonia), "", rsaux1!colonia) + "','" + IIf(IsNull(rsaux1!DIRECCION), "", rsaux1!DIRECCION) + "', '" + IIf(IsNull(rsaux1!cp), "", rsaux1!cp) + "',"
                                    var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!site_use_id), "", rsaux1!site_use_id) + "','" + IIf(IsNull(rsaux1!paqueteria), "", rsaux1!paqueteria) + "','" + IIf(IsNull(rsaux1!attribute8), "", rsaux1!attribute8) + "','" + IIf(IsNull(rsaux1!ATTRIBUTE9), "", rsaux1!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux1!TELEFONO), "", rsaux1!TELEFONO) + "','" + IIf(IsNull(rsaux1!Linea), "", rsaux1!Linea) + "','" + CStr(IIf(IsNull(rsaux1!ruta), "", rsaux1!ruta)) + "','" + IIf(IsNull(rsaux1!nombre_ruta), "", rsaux1!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux1!ORDEN_SURTIDO), 0, rsaux1!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux1!Embarque), 0, rsaux1!Embarque)) + ", '" + IIf(IsNull(rsaux1!estacion), "", rsaux1!estacion) + "'," + CStr(IIf(IsNull(rsaux1!MULTIPLO), "", rsaux1!MULTIPLO)) + ")"
                                    rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                    var_cantidad_total = var_cantidad_total - rsaux1!MULTIPLO
                              Wend
                           Else
                              var_cantidad = var_cantidad_total
                              While var_cantidad > 0
                                    var_dia = CStr(Day(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                    var_mes = CStr(Month(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                    var_año = CStr(Year(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                    If Len(Trim(var_dia)) = 1 Then
                                       var_dia = "0" + var_dia
                                    End If
                                    If Len(Trim(var_mes)) = 1 Then
                                      var_mes = "0" + var_mes
                                    End If
                                    var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                    var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_2 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                    var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,MULTIPLO) "
                                    var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux1!source_header_number + "', " + CStr(rsaux1!delivery_id) + "," + CStr(rsaux1!delivery_detail_id) + ", " + CStr(rsaux1!organization_id) + ",'" + IIf(IsNull(rsaux1!subinventory), "", rsaux1!subinventory) + "', " + CStr(rsaux1!delivery_line_id) + "," + CStr(rsaux1!inventory_item_id) + ",'" + rsaux1!item_description + "','" + CStr(rsaux1!SOURCE_LINE_NUMBER) + "'," + CStr(1) + ",'" + rsaux1!released_status + "',"
                                    var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + rsaux1!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux1!collector_id), 0, rsaux1!collector_id)) + ",'" + IIf(IsNull(rsaux1!Name), "", rsaux1!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux1!ubicacion), "", rsaux1!ubicacion) + "','" + IIf(IsNull(rsaux1!establecimiento), "", rsaux1!establecimiento) + "','" + IIf(IsNull(rsaux1!nombre_Establecimiento), "", rsaux1!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux1!ORDENES), "", rsaux1!ORDENES) + "',"
                                    var_cadena = var_cadena + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux1!source_header_type_name), "", rsaux1!source_header_type_name) + "', '" + IIf(IsNull(rsaux1!source_document_id), "", rsaux1!source_document_id) + "','" + IIf(IsNull(rsaux1!pais), "", rsaux1!pais) + "','" + IIf(IsNull(rsaux1!estado), "", rsaux1!estado) + "', '" + IIf(IsNull(rsaux1!municipio), "", rsaux1!municipio) + "', '" + IIf(IsNull(rsaux1!ciudad), "", rsaux1!ciudad) + "', '" + IIf(IsNull(rsaux1!colonia), "", rsaux1!colonia) + "','" + IIf(IsNull(rsaux1!DIRECCION), "", rsaux1!DIRECCION) + "', '" + IIf(IsNull(rsaux1!cp), "", rsaux1!cp) + "',"
                                    var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!site_use_id), "", rsaux1!site_use_id) + "','" + IIf(IsNull(rsaux1!paqueteria), "", rsaux1!paqueteria) + "','" + IIf(IsNull(rsaux1!attribute8), "", rsaux1!attribute8) + "','" + IIf(IsNull(rsaux1!ATTRIBUTE9), "", rsaux1!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux1!TELEFONO), "", rsaux1!TELEFONO) + "','" + IIf(IsNull(rsaux1!Linea), "", rsaux1!Linea) + "','" + CStr(IIf(IsNull(rsaux1!ruta), "", rsaux1!ruta)) + "','" + IIf(IsNull(rsaux1!nombre_ruta), "", rsaux1!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux1!ORDEN_SURTIDO), 0, rsaux1!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux1!Embarque), 0, rsaux1!Embarque)) + ", '" + IIf(IsNull(rsaux1!estacion), "", rsaux1!estacion) + "'," + CStr(IIf(IsNull(rsaux1!MULTIPLO), "", rsaux1!MULTIPLO)) + ")"
                                    rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                    var_cantidad = var_cantidad - 1
                              Wend
                           End If
                        End If
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  rsaux1.Open "delete from tb_Temp_oracle_orden_surtido_aux_1", cnn, adOpenDynamic, adLockOptimistic
                  'rsaux1.Open "select distinct source_header_number, ORDEN_SURTIDO from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " order by ORDEN_SURTIDO", cnn, adOpenDynamic, adLockOptimistic
                  rsaux1.Open "select distinct source_header_number, ORDEN_pedido from tb_Temp_oracle_orden_surtido_aux_2, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SOURCE_HEADER_NUMBER = pedido  order by ORDEN_pedido", cnn, adOpenDynamic, adLockOptimistic
                  var_contador = 0
                  var_grupo = 1
                  While Not rsaux1.EOF
                        var_lote = 1
                        If var_contador >= 1500 Then
                           var_grupo = var_grupo + 1
                           var_contador = 0
                        End If
                        rsaux2.Open "select * from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number = " + CStr(rsaux1!source_header_number) + " order by ubicacion", cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux2.EOF
                              If var_contador >= 1500 Then
                                 var_lote = var_lote + 1
                                 var_contador = 0
                                 var_grupo = var_grupo + 1
                              End If
                              rsaux3.Open "SELECT * FROM tb_Temp_oracle_orden_surtido_aux_1 WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and delivery_detail_id = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 rsaux4.Open "UPDATE TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 SET SRC_REQUESTED_QUANTITY = SRC_REQUESTED_QUANTITY + " + CStr(rsaux2!src_requested_quantity) + " WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and DELIVERY_DETAIL_ID = " + CStr(rsaux2!delivery_detail_id), cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 var_dia = CStr(Day(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_mes = CStr(Month(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_año = CStr(Year(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 If Len(Trim(var_dia)) = 1 Then
                                    var_dia = "0" + var_dia
                                 End If
                                 If Len(Trim(var_mes)) = 1 Then
                                    var_mes = "0" + var_mes
                                 End If
                                 var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                 
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                 var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,LOTE, GRUPO) "
                                 var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux2!source_header_number + "', " + CStr(rsaux2!delivery_id) + "," + CStr(rsaux2!delivery_detail_id) + ", " + CStr(rsaux2!organization_id) + ",'" + IIf(IsNull(rsaux2!subinventory), "", rsaux2!subinventory) + "', " + CStr(rsaux2!delivery_line_id) + "," + CStr(rsaux2!inventory_item_id) + ",'" + rsaux2!item_description + "','" + CStr(rsaux2!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux2!src_requested_quantity) + ",'" + rsaux2!released_status + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!customer_name), "", rsaux2!customer_name) + "','" + rsaux2!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux2!collector_id), 0, rsaux2!collector_id)) + ",'" + IIf(IsNull(rsaux2!Name), "", rsaux2!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux2!ubicacion), "", rsaux2!ubicacion) + "','" + IIf(IsNull(rsaux2!establecimiento), "", rsaux2!establecimiento) + "','" + IIf(IsNull(rsaux2!nombre_Establecimiento), "", rsaux2!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux2!ORDENES), "", rsaux2!ORDENES) + "',"
                                 var_cadena = var_cadena + CStr(IIf(IsNull(rsaux2!CUST_ACCOUNT_ID), 0, rsaux2!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux2!source_header_type_name), "", rsaux2!source_header_type_name) + "', '" + IIf(IsNull(rsaux2!source_document_id), "", rsaux2!source_document_id) + "','" + IIf(IsNull(rsaux2!pais), "", rsaux2!pais) + "','" + IIf(IsNull(rsaux2!estado), "", rsaux2!estado) + "', '" + IIf(IsNull(rsaux2!municipio), "", rsaux2!municipio) + "', '" + IIf(IsNull(rsaux2!ciudad), "", rsaux2!ciudad) + "', '" + IIf(IsNull(rsaux2!colonia), "", rsaux2!colonia) + "','" + IIf(IsNull(rsaux2!DIRECCION), "", rsaux2!DIRECCION) + "', '" + IIf(IsNull(rsaux2!cp), "", rsaux2!cp) + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!site_use_id), "", rsaux2!site_use_id) + "','" + IIf(IsNull(rsaux2!paqueteria), "", rsaux2!paqueteria) + "','" + IIf(IsNull(rsaux2!attribute8), "", rsaux2!attribute8) + "','" + IIf(IsNull(rsaux2!ATTRIBUTE9), "", rsaux2!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux2!TELEFONO), "", rsaux2!TELEFONO) + "','" + IIf(IsNull(rsaux2!Linea), "", rsaux2!Linea) + "','" + CStr(IIf(IsNull(rsaux2!ruta), "", rsaux2!ruta)) + "','" + IIf(IsNull(rsaux2!nombre_ruta), "", rsaux2!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_SURTIDO), 0, rsaux2!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", '" + IIf(IsNull(rsaux2!estacion), "", rsaux2!estacion) + "'," + CStr(var_lote) + "," + CStr(var_grupo) + ")"
                                 rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux3.Close
                              If rsaux2!Linea <> "CATALOGOS" Then
                                 If rsaux2!Linea <> "CATALOGO" Then
                                    If rsaux2!Linea <> "POP" Then
                                       If rsaux2!Linea <> "EMPAQUE" Then
                                          If rsaux2!MULTIPLO > 1 Then
                                             var_contador = var_contador + rsaux2!src_requested_quantity
                                          Else
                                             var_contador = var_contador + 1
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                              rsaux2.MoveNext
                        Wend
                        rsaux2.Close
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  
                  rsaux1.Open "insert TB_TEMP_ORACLE_ORDEN_SURTIDO (inte_tem_consecutivo, segment1) values (" + CStr(var_consecutivo) + ",'---------')", cnn, adOpenDynamic, adLockOptimistic
                  rsaux1.Open "delete from TB_TEMP_ORACLE_ORDEN_SURTIDO where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and segment1 <> '---------'", cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.Open "insert into TB_TEMP_ORACLE_ORDEN_SURTIDO select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.Open "delete from TB_TEMP_ORACLE_ORDEN_SURTIDO where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and segment1 = '---------'", cnn, adOpenDynamic, adLockOptimistic
                  var_consecutivo_general = var_consecutivo
                  Call crea_tablas
                  If rsaux.State = 1 Then
                     rsaux.Close
                  End If
                  rsaux.Open "select distinct a.source_header_number from tb_Temp_oracle_orden_surtido_aux_1 a, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES  b where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and a.source_header_number = b.pedido", cnn, adOpenDynamic, adLockOptimistic
                  var_cadena_pedidos_diferencias = ""
                  While Not rsaux.EOF
                        strconsulta = "select sum(requested_quantity)  as cantidad from WSH_DELIVERABLES_V where source_header_number = ? AND RELEASED_STATUS = 'Y'"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(rsaux!source_header_number))
                             .Parameters.Append parametro
                        End With
                        Set rsaux10 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        
                        strconsulta = "SELECT SUM(SRC_REQUESTED_QUANTITY) AS CANTIDAD FROM XXVIA_TB_PEDIDOS_DIVIDIDOS WHERE SOURCE_HEADER_NUMBER = ?"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(rsaux!source_header_number))
                             .Parameters.Append parametro
                        End With
                        Set rsaux11 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                     
                     
                        If IIf(IsNull(rsaux11!cantidad), 0, rsaux11!cantidad) <> rsaux10!cantidad Then
                           If var_cadena_pedidos_diferencias = "" Then
                              var_cadena_pedidos_diferencias = CStr(rsaux!source_header_number)
                           Else
                              var_cadena_pedidos_diferencias = var_cadena_pedidos_diferencias + ", " + CStr(rsaux!source_header_number)
                           End If
                        End If
                        rsaux10.Close
                        rsaux11.Close
                        rsaux.MoveNext
                  Wend
                  rsaux.Close
                  If var_cadena_pedidos_diferencias = "" Then
                     If var_imprime_pedidos = 1 Then
                        ' orden
x:
                        'var_consecutivo = 1360
                        rsaux.Open "select distinct a.grupo from tb_Temp_oracle_orden_surtido_aux_1 a, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES  b where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and a.source_header_number = b.pedido order by a.grupo", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           While Not rsaux.EOF
                                 x = 1
                                 If x = 1 Then
                                    
                                    'strconsulta = "select shipping_method_code, packing_instructions from oe_order_headers_all where order_number = ?"
                                    'With comandoORA
                                    '     .ActiveConnection = cnnoracle_4
                                    '     .CommandType = adCmdText
                                    '     .CommandText = strconsulta
                                    '     Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rsaux!source_header_number))
                                    '     .Parameters.Append parametro
                                    'End With
                                    'Set rs = comandoORA.execute
                                    'Set comandoORA = Nothing
                                    'Set parametro = Nothing
                                    
                                    var_paqueteria = ""
                                    'If Not rs.EOF Then
                                    '   VAR_COMENTARIOS = IIf(IsNull(rs!packing_instructions), "", rs!packing_instructions)
                                    '   var_tipo_metodo = IIf(IsNull(rs(0).Value), "", rs(0).Value)
                                    '   If var_tipo_metodo <> "" Then
                                    '
                                    '      strconsulta = "SELECT description FROM fnd_lookup_values where lookup_type = 'SHIP_METHOD' and lookup_code = ? AND LANGUAGE = 'ESA'"
                                    '      With comandoORA
                                    '           .ActiveConnection = cnno-racle_4
                                    '           .CommandType = adCmdText
                                    '           .CommandText = strconsulta
                                    '           Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_tipo_metodo)
                                    '           .Parameters.Append parametro
                                    '      End With
                                    '      Set rsaux1 = comandoORA.execute
                                    '      Set comandoORA = Nothing
                                    '      Set parametro = Nothing
                                          
                                    '      If Not rsaux1.EOF Then
                                    '         var_paqueteria = IIf(IsNull(rsaux1(0).Value), "", rsaux1(0).Value)
                                    '      End If
                                    '      rsaux1.Close
                                    '   End If
                                    'End If
                                    'rs.Close
                                    
                                    VAR_ZZ = 0
                                    If VAR_ZZ = 1 Then
                                        
                                    
                                    
                                    
                                       strconsulta = "SELECT  HPS.party_site_id as tel, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1, OHA.packing_instructions from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.ship_TO_ORG_ID AND oha.order_number = ? AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                                       With comandoORA
                                            .ActiveConnection = cnnoracle_4
                                            .CommandType = adCmdText
                                            .CommandText = strconsulta
                                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rsaux!source_header_number))
                                            .Parameters.Append parametro
                                       End With
                                       Set rsaux6 = comandoORA.execute
                                       Set comandoORA = Nothing
                                       Set parametro = Nothing
                                       
                                       If Not rsaux6.EOF Then
                                       
                                          
                                          strconsulta = "SELECT  HPS.party_site_id as tel, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1, OHA.packing_instructions from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.ship_TO_ORG_ID AND oha.order_number = ? AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                                          With comandoORA
                                               .ActiveConnection = cnnoracle_4
                                               .CommandType = adCmdText
                                               .CommandText = strconsulta
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, IIf(IsNull(rsaux!source_header_number), "", rsaux!source_header_number))
                                               .Parameters.Append parametro
                                          End With
                                          Set rsaux5 = comandoORA.execute
                                          Set comandoORA = Nothing
                                          Set parametro = Nothing
                                          
                                          'rsaux5.Open "SELECT  HPS.party_site_id as tel, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1, OHA.packing_instructions from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.ship_TO_ORG_ID AND oha.order_number = '" + IIf(IsNull(rsaux!source_header_number), "", rsaux!source_header_number) + "' AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                          If Not rsaux5.EOF Then
                                             var_nombre = IIf(IsNull(rsaux5!customer_name), "", rsaux5!customer_name)
                                             var_tel = IIf(IsNull(rsaux5!tel), 0, rsaux5!tel)
                                             VAR_DIRECCION = IIf(IsNull(rsaux5!calle), "", rsaux5!calle) + " " + IIf(IsNull(rsaux5!numero), "", rsaux5!numero)
                                             VAR_COLONIA = IIf(IsNull(rsaux5!colonia), "", rsaux5!colonia)
                                             var_ciudad = IIf(IsNull(rsaux5!ciudad), "", rsaux5!ciudad)
                                             VAR_MUNICIPIO = IIf(IsNull(rsaux5!municipio), "", rsaux5!municipio)
                                             var_estado = IIf(IsNull(rsaux5!estado), "", rsaux5!estado)
                                             var_pais = IIf(IsNull(rsaux5!pais), "", rsaux5!pais)
                                             VAR_CP = IIf(IsNull(rsaux5!cp), "", rsaux5!cp)
                                             rsaux5.Close
                                          Else
                                             rsaux5.Close
                                             var_nombre = IIf(IsNull(rsaux6!customer_name), "", rsaux6!customer_name)
                                             var_tel = IIf(IsNull(rsaux6!tel), 0, rsaux6!tel)
                                             VAR_DIRECCION = IIf(IsNull(rsaux6!calle), "", rsaux6!calle) + " " + IIf(IsNull(rsaux6!numero), "", rsaux6!numero)
                                             VAR_COLONIA = IIf(IsNull(rsaux6!colonia), "", rsaux6!colonia)
                                             var_ciudad = IIf(IsNull(rsaux6!ciudad), "", rsaux6!ciudad)
                                             VAR_MUNICIPIO = IIf(IsNull(rsaux6!municipio), "", rsaux6!municipio)
                                             var_estado = IIf(IsNull(rsaux6!estado), "", rsaux6!estado)
                                             var_pais = IIf(IsNull(rsaux6!pais), "", rsaux6!pais)
                                             VAR_CP = IIf(IsNull(rsaux6!cp), "", rsaux6!cp)
                                          End If
                                       Else
                                          var_tel = 0
                                          VAR_DIRECCION = ""
                                          VAR_COLONIA = ""
                                          var_ciudad = ""
                                          VAR_MUNICIPIO = ""
                                          var_estado = ""
                                          var_pais = ""
                                          VAR_CP = ""
                                       End If
                                       rsaux6.Close
                                       If var_tel > 0 Then
                                          
                                          strconsulta = "select Phone_Number from hz_contact_points where owner_table_id = ?"
                                          With comandoORA
                                               .ActiveConnection = cnnoracle_4
                                               .CommandType = adCmdText
                                               .CommandText = strconsulta
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(var_tel))
                                               .Parameters.Append parametro
                                          End With
                                          Set rsaux6 = comandoORA.execute
                                          Set comandoORA = Nothing
                                          Set parametro = Nothing
                                          If Not rsaux6.EOF Then
                                             var_telefono = CStr(IIf(IsNull(rsaux6(0).Value), "", rsaux6(0).Value))
                                          Else
                                             var_telefono = ""
                                          End If
                                          rsaux6.Close
                                       Else
                                          var_telefono = ""
                                       End If
                                    Else
                                       var_tel = 0
                                       VAR_DIRECCION = ""
                                       VAR_COLONIA = ""
                                       var_ciudad = ""
                                       VAR_MUNICIPIO = ""
                                       var_estado = ""
                                       var_pais = ""
                                       VAR_CP = ""
                                       var_telefono = ""
                                    End If
                                    
                                    If IsNumeric(txt_total_volumen) Then
                                       var_cubicaje_EMBARQUE = CDbl(Me.txt_total_volumen)
                                    Else
                                       var_cubicaje_EMBARQUE = 0
                                    End If
                                    
                                    rsaux4.Open "select distinct source_header_number from tb_Temp_oracle_orden_surtido where  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and grupo = " + CStr(rsaux(0).Value), cnn, adOpenDynamic, adLockOptimistic
                                    While Not rsaux4.EOF
                                          rsaux2.Open "SELECT * FROM tb_oracle_pedidos_asignados_embarques WHERE PEDIDO = " + CStr(rsaux4(0).Value), cnn, adOpenDynamic, adLockOptimistic
                                          If Not rsaux2.EOF Then
                                             var_transporte = ""
                                             rsaux3.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                             If Not rsaux3.EOF Then
                                                var_transporte = IIf(IsNull(rsaux3!vehiculo), "", rsaux3!vehiculo)
                                             End If
                                             rsaux3.Close
                                             
                                             
                                             strconsulta = "SELECT ORDER_TYPE_ID FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ?"
                                             With comandoORA
                                                  .ActiveConnection = cnnoracle_4
                                                  .CommandType = adCmdText
                                                  .CommandText = strconsulta
                                                  Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rsaux4(0).Value))
                                                  .Parameters.Append parametro
                                             End With
                                             Set rsaux5 = comandoORA.execute
                                             Set comandoORA = Nothing
                                             Set parametro = Nothing
                                             If rsaux5!ORDER_TYPE_ID = 1002 Then
                                                rsaux5.Close
                                                strconsulta = "select ADDRESS_LINE_1||', '||ADDRESS_LINE_2||', '||TOWN_OR_CITY||', '||REGION_1||', '||COUNTRY||' CP:'||POSTAL_CODE DIRECCION, EMAIL from mtl_secondary_inventories a, hr_locations_all b, xxvia_jv_tb_agentes c, po_requisition_headers_ALL D, OE_ORDER_HEADERS_ALL E Where A.location_id = b.location_id and a.secondary_inventory_name = c.subinventory_code AND E.source_document_id = D.requisition_header_id AND A.secondary_inventory_name = D.ATTRIBUTE1 AND E.ORDER_NUMBER = ?"
                                             Else
                                                rsaux5.Close
                                                strconsulta = "SELECT CALLE||' '||NUM_CALLE||' '||NVL(num_interior,'')||', '||colonia||', '||ciudad||', '||estado||', '||pais||' CP: '||codigo_postal DIRECCION FROM oe_order_headers_all a, xxvia_vw_CLIENTES_BCP B WHERE A.SHIP_TO_ORG_ID = B.SITE_USE_ID AND A.ORDER_NUMBER     = ?"
                                             End If
                                             
                                             With comandoORA
                                                  .ActiveConnection = cnnoracle_4
                                                  .CommandType = adCmdText
                                                  .CommandText = strconsulta
                                                  Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rsaux4(0).Value))
                                                  .Parameters.Append parametro
                                             End With
                                             Set rsaux5 = comandoORA.execute
                                             Set comandoORA = Nothing
                                             Set parametro = Nothing
                                             If Not rsaux5.EOF Then
                                                VAR_DIRECCION = IIf(IsNull(rsaux5!DIRECCION), "", rsaux5!DIRECCION)
                                             Else
                                                VAR_DIRECCION = ""
                                             End If
                                             
                                             
                                             
                                             var_cadena = "UPDATE tb_Temp_oracle_orden_surtido SET EMBARQUE = " + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", CUBICAJE = " + CStr(var_cubicaje_EMBARQUE) + " , ORDEN_SURTIDO = " + CStr(IIf(IsNull(rsaux2!orden_pedido), 0, rsaux2!orden_pedido)) + ", ANDEN = '" + CStr(IIf(IsNull(rsaux2!estacion), 0, rsaux2!estacion)) + "', TRANSPORTE = '" + var_transporte + "',"
                                             var_cadena = var_cadena + " pais= '" + var_pais + "', estado = '" + var_estado + "', municipio = '" + VAR_MUNICIPIO + "', ciudad = '" + var_ciudad + "', colonia = '" + VAR_COLONIA + "', direccion = '" + VAR_DIRECCION + "', cp = '" + VAR_CP + "', paqueteria = '" + var_paqueteria + "'"
                                             var_cadena = var_cadena + " WHERE SOURCE_HEADER_NUMBER = " + CStr(rsaux4(0).Value) + " AND inte_tem_consecutivo = " + CStr(var_consecutivo)
                                             rsaux3.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                          End If
                                          rsaux2.Close
                                          rsaux4.MoveNext
                                    Wend
                                    rsaux4.Close
                                    
                                                                  
                                    x = 1
                                    If x = 1 Then
                                       If rsaux9.State = 1 Then
                                          rsaux9.Close
                                       End If
                                       rsaux9.Open "SELECT DISTINCT  GRUPO, SOURCE_HEADER_NUMBER from tb_Temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and grupo = " + CStr(rsaux(0).Value), cnn, adOpenDynamic, adLockOptimistic
                                       VAR_TOTAL_GRUPOS = rsaux9.RecordCount
                                       rsaux9.Close
                                       'VAR_TOTAL_GRUPOS = 1
                                       rsaux10.Open "update tb_Temp_oracle_orden_surtido set cubicaje = " + CStr(var_cubicaje_EMBARQUE) + "  where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                                       If VAR_TOTAL_GRUPOS = 1 Then
                                          Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos.rpt")
                                          reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.grupo} = " + CStr(rsaux(0).Value) + " and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                                          'frmvistasprevias.cr.ReportSource = reporte
                                          For ntablas = 1 To reporte.Database.Tables.Count
                                              reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                          Next ntablas
                                          'MsgBox cnn.ConnectionString
                                          reporte.PrintOut False
                                          Set reporte = Nothing
                                          
                                       Else
                                          
                                          Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos_ENCABEZADOS.rpt")
                                          reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.grupo} = " + CStr(rsaux(0).Value) + " and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                                          frmvista_previa_auxiliar.cr2.ReportSource = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos_ENCABEZADOS.rpt")
                                          For ntablas = 1 To reporte.Database.Tables.Count
                                              reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                          Next ntablas
                                          reporte.PrintOut False
                                          frmvista_previa_auxiliar.Show 1
                                          
            'frmvistasprevias.cr.ViewReport
            'frmvistasprevias.Caption = "Packing List"
            'frmvistasprevias.Show 1
                                          
                                          Set reporte = Nothing
                                          
                                          Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos_DETALLE.rpt")
                                          reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.grupo} = " + CStr(rsaux(0).Value) + " and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                                          frmvista_previa_auxiliar.cr2.ReportSource = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos_DETALLE.rpt")
                                          
                                          
                                          For ntablas = 1 To reporte.Database.Tables.Count
                                              reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                          Next ntablas
                                          reporte.PrintOut False
            'frmvistasprevias.cr.ViewReport
            'frmvistasprevias.Caption = "Packing List"
            'frmvistasprevias.Show 1
                                          
                                          Set reporte = Nothing
                                          frmvista_previa_auxiliar.Show 1
                                       
                                       End If
                                       
                                       
                                       'Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos.rpt")
                                       'reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.grupo} = " + CStr(rsaux(0).Value) + " and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                                       'For ntablas = 1 To reporte.Database.Tables.Count
                                       '    reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                       'Next ntablas
                                       'reporte.PrintOut False
                                       'Set reporte = Nothing
                                    Else
                                       If rsaux9.State = 1 Then
                                          rsaux9.Close
                                       End If
                                       rsaux9.Open "SELECT DISTINCT  GRUPO, SOURCE_HEADER_NUMBER from tb_Temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and grupo = " + CStr(rsaux(0).Value), cnn, adOpenDynamic, adLockOptimistic
                                       VAR_TOTAL_GRUPOS = rsaux9.RecordCount
                                       rsaux9.Close
                                       'VAR_TOTAL_GRUPOS = 1
                                       If VAR_TOTAL_GRUPOS = 1 Then
                                          Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos.rpt")
                                          reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.grupo} = " + CStr(rsaux(0).Value) + " and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                                          frmvistasprevias.cr.ReportSource = reporte
                                          For ntablas = 1 To reporte.Database.Tables.Count
                                              reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                          Next ntablas
                                          frmvistasprevias.cr.ViewReport
                                          frmvistasprevias.Caption = "Ordenes de surtido "
                                          frmvistasprevias.Show 1
                                          Set reporte = Nothing
                                       Else
                                          Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos_ENCABEZADOS.rpt")
                                          reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.grupo} = " + CStr(rsaux(0).Value) + " and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                                          frmvistasprevias.cr.ReportSource = reporte
                                          For ntablas = 1 To reporte.Database.Tables.Count
                                              reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                          Next ntablas
                                          frmvistasprevias.cr.ViewReport
                                          frmvistasprevias.Caption = "Ordenes de surtido "
                                          frmvistasprevias.Show 1
                                          Set reporte = Nothing
                                          
                                          Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos_DETALLE.rpt")
                                          reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.grupo} = " + CStr(rsaux(0).Value) + " and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                                          frmvistasprevias.cr.ReportSource = reporte
                                          For ntablas = 1 To reporte.Database.Tables.Count
                                              reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                          Next ntablas
                                          frmvistasprevias.cr.ViewReport
                                          frmvistasprevias.Caption = "Ordenes de surtido "
                                          frmvistasprevias.Show 1
                                          Set reporte = Nothing
                                       
                                       
                                       End If
                                       
                                    End If
                                 End If
                                 rsaux.MoveNext
                           Wend
                        End If
                        rsaux.Close
                     Else
                        MsgBox "No se pueden imprimir los pedidos, vuelva a intentar la impresión", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "No se pueden imprimir los pedidos, vuelva a intentar la impresión", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "No se pueden imprimir los pedidos, vuelva a intentar la impresión", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "No existen ordenes de surtido", vbOKOnly, "ATENCION"
            End If
            If rs.State = 1 Then
               rs.Close
            End If
            If var_consecutivo > 0 Then
               rs.Open "delete from tb_Temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(IIf(IsNull(var_consecutivo), 0, var_consecutivo)), cnn, adOpenDynamic, adLockOptimistic
            End If
         Else
            'MsgBox "Número superior incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "No existen pedidos a imprimir"
      End If
End Sub

Private Sub cmd_orden_carga_Click()
   If Me.lv_pedidos.ListItems.Count > 0 Then
      For var_j = 1 To Me.lv_pedidos.ListItems.Count
          Me.lv_pedidos.ListItems.Item(var_j).Selected = True
          rs.Open "select * from XXVIA_TB_CLIENTES_RUTAS_DISTR where ESTABLECIMIENTO = '" + Me.lv_pedidos.selectedItem.SubItems(8) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
          If Not rs.EOF Then
             Me.lv_pedidos.selectedItem.SubItems(5) = IIf(IsNull(rs!prioridad), 0, rs!prioridad)
             'rsaux.Open "update tb_oracle_pedidos_asignados_embarques set orden_pedido = '" + CStr(IIf(IsNull(rs!prioridad), 0, rs!prioridad)) + "' where pedido = " + Me.lv_pedidos.selectedItem, cnn, adOpenDynamic, adLockOptimistic
          Else
             Me.lv_pedidos.selectedItem.SubItems(5) = 0
             'rsaux.Open "update tb_oracle_pedidos_asignados_embarques set orden_pedido = '0' where pedido = " + Me.lv_pedidos.selectedItem, cnn, adOpenDynamic, adLockOptimistic
          End If
          rs.Close
      Next var_j
   Else
      MsgBox "No existen pedidos", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_por_pasillo_Click()
   var_si = MsgBox("¿Desea imprimir las ordenes de surtido por pasillo", vbYesNo, "ATENCION")
   If var_si = 6 Then
      var_si = MsgBox("Confirmar la impresión de las ordenes de surtido por pasillo", vbYesNo, "ATENCION")
      If var_si = 6 Then
'GoTo x:
      If rs.State = 1 Then
         rs.Close
      End If
      If rsaux.State = 1 Then
         rsaux.Close
      End If
      If rsaux1.State = 1 Then
         rsaux1.Close
      End If
      If rsaux2.State = 1 Then
         rsaux2.Close
      End If
      If rsaux3.State = 1 Then
         rsaux3.Close
      End If
      If rsaux4.State = 1 Then
         rsaux4.Close
      End If
      If rsaux5.State = 1 Then
         rsaux5.Close
      End If
      If rsaux6.State = 1 Then
         rsaux6.Close
      End If
      If rsaux7.State = 1 Then
         rsaux7.Close
      End If
      If rsaux8.State = 1 Then
         rsaux8.Close
      End If
      If rsaux9.State = 1 Then
         rsaux9.Close
      End If
     
      If Me.lv_pedidos.ListItems.Count > 0 Then
         var_si = MsgBox("Desea imprimir las ordenes de surtido?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            If rs.State = 1 Then
               rs.Close
            End If
            var_Cadena_pedidos = ""
            For var_j = 1 To Me.lv_pedidos.ListItems.Count
                Me.lv_pedidos.ListItems.Item(var_j).Selected = True
                If Me.lv_pedidos.selectedItem <> "10000000" Then
                   If var_Cadena_pedidos = "" Then
                      var_Cadena_pedidos = Me.lv_pedidos.selectedItem
                   Else
                      var_Cadena_pedidos = var_Cadena_pedidos + "," + Me.lv_pedidos.selectedItem
                   End If
                End If
            Next var_j
            'var_cadena_pedidos = "105208"
            rs.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_cadena = "SELECT to_char(a.LAST_UPDATE_DATE,'day') DIA_SEMANA, CAT.LINEA, HCSU.SITE_USE_ID, HCAS.CUST_ACCOUNT_ID, OHA.SHIP_TO_ORG_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id, a.LAST_UPDATE_DATE,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,c.description,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, a.source_header_type_name, oha.source_document_id, C.ATTRIBUTE2, oha.attribute8, oha.attribute9, j.NAME as nombre_ruta, j.salesrep_id as clave_ruta from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, xxvia_vw_articulos_cat cat, XXVIA_VENDEDORES j Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND "
            var_cadena = var_cadena + " to_number(source_header_number)  IN (" + var_Cadena_pedidos + ")"
            var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID "
            var_cadena = var_cadena + " AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND released_status = 'Y' AND A.inventory_item_id  = cat.item_id AND A.ORGANIZATION_ID = Cat.organization_id and oha.salesrep_id = j.salesrep_id "
            rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
'--------------------------
            If Not rs.EOF Then
               cnn.BeginTrans
               rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM tb_Temp_oracle_orden_surtido", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
               Else
                  var_consecutivo = 1
               End If
               rsaux.Close
               rsaux1.Open "insert into tb_Temp_oracle_orden_surtido(inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
               cnn.CommitTrans
               While Not rs.EOF
                     var_establecimiento = rs!SHIP_TO_ORG_ID
                     rsaux.Open "SELECT csu.site_use_id AS VCHA_ESB_ESTABLECIMIENTO_ID, ps.party_site_number, lo.address1 AS VCHA_eSB_NOMBRE FROM hz_cust_site_uses_all csu, hz_cust_acct_sites_all cas, hz_party_sites ps, hz_locations lo Where csu.cust_acct_site_id = cas.cust_acct_site_id AND cas.party_site_id = ps.party_site_id AND ps.location_id = lo.location_id AND csu.site_use_code = 'SHIP_TO' AND csu.LOCATION = ps.party_site_number and csu.site_use_id = " + CStr(var_establecimiento), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        VAR_NOMBRE_ESTABLECIMIENTO = IIf(IsNull(rsaux!vcha_esb_nombre), "", rsaux!vcha_esb_nombre)
                     Else
                        VAR_NOMBRE_ESTABLECIMIENTO = ""
                     End If
                     rsaux.Close
                     var_dia = CStr(Day(CDate(rs!LAST_UPDATE_DATE)))
                     var_mes = CStr(Month(CDate(rs!LAST_UPDATE_DATE)))
                     var_año = CStr(Year(CDate(rs!LAST_UPDATE_DATE)))
                     If Len(Trim(var_dia)) = 1 Then
                        var_dia = "0" + var_dia
                     End If
                     If Len(Trim(var_mes)) = 1 Then
                        var_mes = "0" + var_mes
                     End If
                     var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                     rsaux1.Open "select * from tb_oracle_multiplos where segment1 = '" + IIf(IsNull(rs!SEGMENT1), "", rs!SEGMENT1) + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux1.EOF Then
                        VAR_MULTIPLO = IIf(IsNull(rsaux1!MULTIPLO), 1, rsaux1!MULTIPLO)
                     Else
                        VAR_MULTIPLO = 1
                     End If
                     rsaux1.Close
'''''
                     If rs!source_header_type_name = "VIA_PEDIDO_INTERNO" Then
                        rsaux1.Open "SELECT * FROM TB_ORACLE_ARTICULOS_MOTOR_LOGISTICO WHERE CODIGO = '" + rs!SEGMENT1 + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux1.EOF Then
                           strconsulta = "SELECT secondary_inventory_name, A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rs!source_document_id)
                                .Parameters.Append parametro
                           End With
                           Set rsaux8 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           If rsaux8.EOF Then
                              var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                           Else
                              var_almacen = rsaux8!secondary_inventory_name
                              rsaux9.Open "SELECT * FROM TB_ORACLE_UBICACIONES_MOTOR_LOGISTICO WHERE CLAVE = '" + var_almacen + "' AND CODIGO = '" + rs!SEGMENT1 + "'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux9.EOF Then
                                 var_ubicacion = ""
                                 If Me.cmb_dia.Text = "Lunes" Then
                                    var_ubicacion = rsaux9!ubicacion_1
                                 End If
                                 If Me.cmb_dia.Text = "Martes" Then
                                    var_ubicacion = rsaux9!ubicacion_2
                                 End If
                                 If Me.cmb_dia.Text = "Miercoles" Then
                                    var_ubicacion = rsaux9!ubicacion_3
                                 End If
                                 If Me.cmb_dia.Text = "Jueves" Then
                                    var_ubicacion = rsaux9!ubicacion_4
                                 End If
                                 If Me.cmb_dia.Text = "Viernes" Then
                                    var_ubicacion = rsaux9!ubicacion_5
                                 End If
                                 If Me.cmb_dia.Text = "Sabado" Then
                                    var_ubicacion = rsaux9!ubicacion_6
                                 End If
                                 If IIf(IsNull(var_ubicacion), "", var_ubicacion) = "" Then
                                    var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                                 End If
                              Else
                                 var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                              End If
                              rsaux9.Close
                           End If
                           rsaux8.Close
                        Else
                           var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                        End If
                        rsaux1.Close
                     Else
                        var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                     End If
                     
                     
'''''
                     var_cadena = "insert into tb_Temp_oracle_orden_surtido(INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER, DELIVERY_ID, DELIVERY_DETAIL_ID, ORGANIZATION_ID, SUBINVENTORY, DELIVERY_LINE_ID, INVENTORY_ITEM_ID, ITEM_DESCRIPTION, SOURCE_LINE_NUMBER, src_requested_quantity, RELEASED_STATUS, CUSTOMER_NAME, SEGMENT1, COLLECTOR_ID, NAME, date_requested, UBICACION, ESTABLECIMIENTO, NOMBRE_ESTABLECIMIENTO, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, SITE_USE_ID, attribute8, attribute9, LINEA, RUTA, NOMBRE_RUTA,MULTIPLO)  values "
                     var_cadena = var_cadena + " (" + CStr(var_consecutivo) + ", '" + IIf(IsNull(rs!source_header_number), "", rs!source_header_number) + "', " + CStr(IIf(IsNull(rs!delivery_id), 0, rs!delivery_id)) + ", " + CStr(IIf(IsNull(rs!delivery_detail_id), 0, rs!delivery_detail_id)) + ", " + CStr(IIf(IsNull(rs!organization_id), 0, rs!organization_id)) + ", '" + IIf(IsNull(rs!subinventory), "", rs!subinventory) + "', " + CStr(IIf(IsNull(rs!delivery_line_id), 0, rs!delivery_line_id)) + ", " + CStr(IIf(IsNull(rs!inventory_item_id), "", rs!inventory_item_id)) + ", '" + IIf(IsNull(rs!Description), "", rs!Description) + "', '" + IIf(IsNull(rs!SOURCE_LINE_NUMBER), "", rs!SOURCE_LINE_NUMBER) + "', " + CStr(IIf(IsNull(rs!requested_quantity), 0, rs!requested_quantity)) + ", '" + IIf(IsNull(rs!released_status), "", rs!released_status) + "', '" + Replace(IIf(IsNull(rs!customer_name), "", rs!customer_name), "'", " ") + "', '" + IIf(IsNull(rs!SEGMENT1), "", rs!SEGMENT1) + "'"
                     'var_cadena = var_cadena + ", 0,''," + var_fecha + ",'" + IIf(IsNull(rs!ATTRIBUTE2), "", rs!ATTRIBUTE2) + "','" + CStr(VAR_ESTABLECIMIENTO) + "','" + Replace(VAR_NOMBRE_ESTABLECIMIENTO, "'", " ") + "'," + CStr(rs!CUST_ACCOUNT_ID) + ",'" + rs!source_header_type_name + "','" + CStr(IIf(IsNull(rs!source_document_id), "", rs!source_document_id)) + "'," + CStr(IIf(IsNull(rs!SITE_USE_ID), 0, rs!SITE_USE_ID)) + ",'" + IIf(IsNull(rs!attribute8), "", rs!attribute8) + "','" + IIf(IsNull(rs!ATTRIBUTE9), "", rs!ATTRIBUTE9) + "','" + IIf(IsNull(rs!linea), "", rs!linea) + "','" + CStr(rs!CLAVE_RUTA) + "','" + IIf(IsNull(rs!nombre_ruta), "", rs!nombre_ruta) + "'," + CStr(VAR_MULTIPLO) + ")"
                     var_cadena = var_cadena + ", 0,''," + var_fecha + ",'" + var_ubicacion + "','" + CStr(var_establecimiento) + "','" + Replace(VAR_NOMBRE_ESTABLECIMIENTO, "'", " ") + "'," + CStr(rs!CUST_ACCOUNT_ID) + ",'" + rs!source_header_type_name + "','" + CStr(IIf(IsNull(rs!source_document_id), "", rs!source_document_id)) + "'," + CStr(IIf(IsNull(rs!site_use_id), 0, rs!site_use_id)) + ",'" + IIf(IsNull(rs!attribute8), "", rs!attribute8) + "','" + IIf(IsNull(rs!ATTRIBUTE9), "", rs!ATTRIBUTE9) + "','" + IIf(IsNull(rs!Linea), "", rs!Linea) + "','" + CStr(rs!CLAVE_RUTA) + "','" + IIf(IsNull(rs!nombre_ruta), "", rs!nombre_ruta) + "'," + CStr(VAR_MULTIPLO) + ")"
                     rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               rsaux1.Open "delete from tb_Temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number is null", cnn, adOpenDynamic, adLockOptimistic
               If rsaux1.State = 1 Then
                  rsaux1.Close
               End If
               
               var_cadena_pedidos_diferencias = ""
               rsaux1.Open "select source_header_number, sum(src_requested_quantity) as cantidad from tb_Temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo) + " group by source_header_number", cnn, adOpenDynamic, adLockOptimistic
               While Not rsaux1.EOF
                     strconsulta = "select sum(requested_quantity)  as cantidad from WSH_DELIVERABLES_V where source_header_number = ? AND RELEASED_STATUS = 'Y'"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(rsaux1!source_header_number))
                          .Parameters.Append parametro
                     End With
                     Set rsaux10 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     If rsaux1!cantidad <> rsaux10!cantidad Then
                        If var_cadena_pedidos_diferencias = "" Then
                           var_cadena_pedidos_diferencias = CStr(rsaux1!source_header_number)
                        Else
                           var_cadena_pedidos_diferencias = var_cadena_pedidos_diferencias + ", " + CStr(rsaux1!source_header_number)
                        End If
                     End If
                     rsaux10.Close
                     rsaux1.MoveNext
               Wend
               rsaux1.Close
               
               If var_cadena_pedidos_diferencias = "" Then
                  rsaux1.Open "select min(DATE_REQUESTED) as DATE_REQUESTED, source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME, RUTA, NOMBRE_RUTA, SUM(SRC_REQUESTED_QUANTITY) AS CANTIDAD from tb_Temp_oracle_orden_surtido where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " group by source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME, RUTA, NOMBRE_RUTA", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        If rsaux1!source_header_type_name = "VIA_PEDIDO_INTERNO" Or rsaux1!source_header_type_name = "TEX_PEDIDO_INTERNO" Then
                           If var_pedido_tienda = 0 Then
                              If rsaux2.State = 1 Then
                                 rsaux2.Close
                              End If
                              rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(rsaux1!source_document_id) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux2.EOF Then
                                 If rsaux4.State = 1 Then
                                    rsaux4.Close
                                 End If
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                                 rsaux4.Close
                              Else
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                                 rsaux4.Close
                              End If
                              rsaux2.Close
                           Else
                              rsaux2.Open "select a.attribute8, B.description from oe_order_headers_all a, MTL_SECONDARY_INVENTORIES b where order_number = " + CStr(rsaux1!source_header_number) + " and a.attribute8 = b.secondary_inventory_name", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux2.EOF Then
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                                 rsaux4.Close
                              Else
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                                 rsaux4.Close
                              End If
                              rsaux2.Close
                           End If
                        Else
                           rsaux6.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                           VAR_PROVEEDOR = rsaux6!collector_id
                           VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux6!Name), "", rsaux6!Name)
                           rsaux6.Close
                        End If
                        var_año_str = CStr(Year(rsaux1!DATE_REQUESTED))
                        If Len(var_año_str) < 2 Then
                           var_año_str = "20" + var_año_str
                        End If
                        var_mes_str = CStr(Month(rsaux1!DATE_REQUESTED))
                        If Len(var_mes_str) < 2 Then
                           var_mes_str = "0" + var_mes_str
                        End If
                        var_dia_str = CStr(Day(rsaux1!DATE_REQUESTED))
                        If Len(var_dia_str) < 2 Then
                           var_dia_str = "0" + var_dia_str
                        End If
                        var_hora_str = CStr(Hour(rsaux1!DATE_REQUESTED))
                        If Len(var_hora_str) < 2 Then
                           var_hora_str = "0" + var_hora_str
                        End If
                        VAR_MINUTO_STR = CStr(Minute(rsaux1!DATE_REQUESTED))
                        If Len(VAR_MINUTO_STR) < 2 Then
                           VAR_MINUTO_STR = "0" + VAR_MINUTO_STR
                        End If
                        VAR_SEGUNDO_STR = CStr(Second(rsaux1!DATE_REQUESTED))
                        If Len(VAR_SEGUNDO_STR) < 2 Then
                           VAR_SEGUNDO_STR = "0" + VAR_SEGUNDO_STR
                        End If
                        var_fecha_pedido = var_año_str + "-" + var_mes_str + "-" + var_dia_str + " " + var_hora_str + ":" + VAR_MINUTO_STR + ":" + VAR_SEGUNDO_STR
                        rsaux6.Open "update tb_Temp_oracle_orden_surtido set COLLECTOR_ID = " + CStr(VAR_PROVEEDOR) + ", NAME = '" + VAR_NOMBRE_PROVEEDOR + "' where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " and CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and source_header_number = " + CStr(rsaux1!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                        rsaux6.Open "SELECT * FROM XXVIA_TB_ORDENES_GRAFICA WHERE ORGANIZACION = " + var_unidad_organizacional + " AND PEDIDO = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If rsaux6.EOF Then
                           rsaux5.Open "INSERT INTO XXVIA_TB_ORDENES_GRAFICA (ORGANIZACION, PEDIDO, PROVEEDOR_ID, NOMBRE_PROVEEDOR, CLIENTE, ESTABLECIMIENTO, CANTIDAD, FECHA, TIPO_PEDIDO, CANTIDAD_LEIDA, FECHA_PEDIDO, RUTA, NOMBRE_RUTA) VALUES (" + var_unidad_organizacional + ", " + CStr(rsaux1!source_header_number) + "," + CStr(VAR_PROVEEDOR) + ",'" + VAR_NOMBRE_PROVEEDOR + "','" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "'," + CStr(rsaux1!cantidad) + ", '" + CStr(rsaux1!DATE_REQUESTED) + "','" + rsaux1!source_header_type_name + "',0, TO_DATE('" + var_fecha_pedido + "','YYYY-MM-DD HH24:MI:SS'),'" + rsaux1!ruta + "', '" + rsaux1!nombre_ruta + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        Else
                           rsaux5.Open "UPDATE XXVIA_TB_ORDENES_GRAFICA SET RUTA = '" + rsaux1!ruta + "', NOMBRE_RUTA = '" + rsaux1!nombre_ruta + "' WHERE PEDIDO = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        End If
                        rsaux6.Close
                        rsaux6.Open "select * from tb_oracle_pedidos_asignados_embarques where pedido = " + CStr(rsaux1!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                        If rsaux6.EOF Then
                           rsaux5.Open "insert into tb_oracle_pedidos_asignados_embarques (AGENTE, NOMBRE_AGENTE, PEDIDO, CLIENTE, PIEZAS, embarque, dia,  mes, AÑO, ORGANIZACION) values (" + CStr(VAR_PROVEEDOR) + ",'" + VAR_NOMBRE_PROVEEDOR + "'," + CStr(rsaux1!source_header_number) + ",'" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "', " + CStr(rsaux1!cantidad) + ",0," + var_dia_str + "," + var_mes_str + "," + var_año_str + "," + CStr(var_unidad_organizacional) + ")", cnn, adOpenDynamic, adLockOptimistic
                        End If
                        rsaux6.Close
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  rsaux1.Open "select min(DATE_REQUESTED) as DATE_REQUESTED, source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME, SUM(SRC_REQUESTED_QUANTITY) AS CANTIDAD from tb_Temp_oracle_orden_surtido where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " AND (LINEA NOT IN ('CATALOGOS','CATALOGO','POP') OR LINEA IS NULL) group by source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME"
                  While Not rsaux1.EOF
                        rsaux5.Open "UPDATE XXVIA_TB_ORDENES_GRAFICA SET CANTIDAD_SIN_CATALOGOS = " + CStr(IIf(IsNull(rsaux1!cantidad), 0, rsaux1!cantidad)) + " WHERE PEDIDO = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  rsaux1.Open "DELETE from tb_Temp_oracle_orden_surtido_aux_2", cnn, adOpenDynamic, adLockOptimistic
                  rsaux1.Open "SELECT * FROM tb_Temp_oracle_orden_surtido where inte_tem_consecutivo =  " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        If rsaux1!Linea = "CATALOGOS" Or rsaux1!Linea = "CATALOGO" Or rsaux1!Linea = "POP" Or rsaux1!Linea = "EMPAQUE" Then
                           var_dia = CStr(Day(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                           var_mes = CStr(Month(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                           var_año = CStr(Year(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                           If Len(Trim(var_dia)) = 1 Then
                              var_dia = "0" + var_dia
                           End If
                           If Len(Trim(var_mes)) = 1 Then
                              var_mes = "0" + var_mes
                           End If
                           var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                           var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_2 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                           var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,MULTIPLO) "
                           var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux1!source_header_number + "', " + CStr(rsaux1!delivery_id) + "," + CStr(rsaux1!delivery_detail_id) + ", " + CStr(rsaux1!organization_id) + ",'" + IIf(IsNull(rsaux1!subinventory), "", rsaux1!subinventory) + "', " + CStr(rsaux1!delivery_line_id) + "," + CStr(rsaux1!inventory_item_id) + ",'" + rsaux1!item_description + "','" + CStr(rsaux1!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux1!src_requested_quantity) + ",'" + rsaux1!released_status + "',"
                           var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + rsaux1!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux1!collector_id), 0, rsaux1!collector_id)) + ",'" + IIf(IsNull(rsaux1!Name), "", rsaux1!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux1!ubicacion), "", rsaux1!ubicacion) + "','" + IIf(IsNull(rsaux1!establecimiento), "", rsaux1!establecimiento) + "','" + IIf(IsNull(rsaux1!nombre_Establecimiento), "", rsaux1!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux1!ORDENES), "", rsaux1!ORDENES) + "',"
                           var_cadena = var_cadena + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux1!source_header_type_name), "", rsaux1!source_header_type_name) + "', '" + IIf(IsNull(rsaux1!source_document_id), "", rsaux1!source_document_id) + "','" + IIf(IsNull(rsaux1!pais), "", rsaux1!pais) + "','" + IIf(IsNull(rsaux1!estado), "", rsaux1!estado) + "', '" + IIf(IsNull(rsaux1!municipio), "", rsaux1!municipio) + "', '" + IIf(IsNull(rsaux1!ciudad), "", rsaux1!ciudad) + "', '" + IIf(IsNull(rsaux1!colonia), "", rsaux1!colonia) + "','" + IIf(IsNull(rsaux1!DIRECCION), "", rsaux1!DIRECCION) + "', '" + IIf(IsNull(rsaux1!cp), "", rsaux1!cp) + "',"
                           var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!site_use_id), "", rsaux1!site_use_id) + "','" + IIf(IsNull(rsaux1!paqueteria), "", rsaux1!paqueteria) + "','" + IIf(IsNull(rsaux1!attribute8), "", rsaux1!attribute8) + "','" + IIf(IsNull(rsaux1!ATTRIBUTE9), "", rsaux1!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux1!TELEFONO), "", rsaux1!TELEFONO) + "','" + IIf(IsNull(rsaux1!Linea), "", rsaux1!Linea) + "','" + CStr(IIf(IsNull(rsaux1!ruta), "", rsaux1!ruta)) + "','" + IIf(IsNull(rsaux1!nombre_ruta), "", rsaux1!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux1!ORDEN_SURTIDO), 0, rsaux1!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux1!Embarque), 0, rsaux1!Embarque)) + ", '" + IIf(IsNull(rsaux1!estacion), "", rsaux1!estacion) + "'," + CStr(IIf(IsNull(rsaux1!MULTIPLO), "", rsaux1!MULTIPLO)) + ")"
                           rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        Else
                           var_cantidad_total = rsaux1!src_requested_quantity
                           If rsaux1!MULTIPLO > 1 Then
                              While var_cantidad_total > 0
                                    If var_cantidad_total < rsaux1!MULTIPLO Then
                                       var_cantidad = var_cantidad_total
                                    Else
                                       var_cantidad = rsaux1!MULTIPLO
                                    End If
                                    
                                    var_dia = CStr(Day(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                    var_mes = CStr(Month(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                    var_año = CStr(Year(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                    If Len(Trim(var_dia)) = 1 Then
                                       var_dia = "0" + var_dia
                                    End If
                                    If Len(Trim(var_mes)) = 1 Then
                                      var_mes = "0" + var_mes
                                    End If
                                    
                                    var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                    var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_2 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                    var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,MULTIPLO) "
                                    var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux1!source_header_number + "', " + CStr(rsaux1!delivery_id) + "," + CStr(rsaux1!delivery_detail_id) + ", " + CStr(rsaux1!organization_id) + ",'" + IIf(IsNull(rsaux1!subinventory), "", rsaux1!subinventory) + "', " + CStr(rsaux1!delivery_line_id) + "," + CStr(rsaux1!inventory_item_id) + ",'" + rsaux1!item_description + "','" + CStr(rsaux1!SOURCE_LINE_NUMBER) + "'," + CStr(var_cantidad) + ",'" + rsaux1!released_status + "',"
                                    var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + rsaux1!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux1!collector_id), 0, rsaux1!collector_id)) + ",'" + IIf(IsNull(rsaux1!Name), "", rsaux1!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux1!ubicacion), "", rsaux1!ubicacion) + "','" + IIf(IsNull(rsaux1!establecimiento), "", rsaux1!establecimiento) + "','" + IIf(IsNull(rsaux1!nombre_Establecimiento), "", rsaux1!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux1!ORDENES), "", rsaux1!ORDENES) + "',"
                                    var_cadena = var_cadena + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux1!source_header_type_name), "", rsaux1!source_header_type_name) + "', '" + IIf(IsNull(rsaux1!source_document_id), "", rsaux1!source_document_id) + "','" + IIf(IsNull(rsaux1!pais), "", rsaux1!pais) + "','" + IIf(IsNull(rsaux1!estado), "", rsaux1!estado) + "', '" + IIf(IsNull(rsaux1!municipio), "", rsaux1!municipio) + "', '" + IIf(IsNull(rsaux1!ciudad), "", rsaux1!ciudad) + "', '" + IIf(IsNull(rsaux1!colonia), "", rsaux1!colonia) + "','" + IIf(IsNull(rsaux1!DIRECCION), "", rsaux1!DIRECCION) + "', '" + IIf(IsNull(rsaux1!cp), "", rsaux1!cp) + "',"
                                    var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!site_use_id), "", rsaux1!site_use_id) + "','" + IIf(IsNull(rsaux1!paqueteria), "", rsaux1!paqueteria) + "','" + IIf(IsNull(rsaux1!attribute8), "", rsaux1!attribute8) + "','" + IIf(IsNull(rsaux1!ATTRIBUTE9), "", rsaux1!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux1!TELEFONO), "", rsaux1!TELEFONO) + "','" + IIf(IsNull(rsaux1!Linea), "", rsaux1!Linea) + "','" + CStr(IIf(IsNull(rsaux1!ruta), "", rsaux1!ruta)) + "','" + IIf(IsNull(rsaux1!nombre_ruta), "", rsaux1!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux1!ORDEN_SURTIDO), 0, rsaux1!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux1!Embarque), 0, rsaux1!Embarque)) + ", '" + IIf(IsNull(rsaux1!estacion), "", rsaux1!estacion) + "'," + CStr(IIf(IsNull(rsaux1!MULTIPLO), "", rsaux1!MULTIPLO)) + ")"
                                    rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                    var_cantidad_total = var_cantidad_total - rsaux1!MULTIPLO
                              Wend
                           Else
                              var_cantidad = var_cantidad_total
                              While var_cantidad > 0
                                    var_dia = CStr(Day(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                    var_mes = CStr(Month(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                    var_año = CStr(Year(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                    If Len(Trim(var_dia)) = 1 Then
                                       var_dia = "0" + var_dia
                                    End If
                                    If Len(Trim(var_mes)) = 1 Then
                                      var_mes = "0" + var_mes
                                    End If
                                    var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                    var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_2 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                    var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,MULTIPLO) "
                                    var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux1!source_header_number + "', " + CStr(rsaux1!delivery_id) + "," + CStr(rsaux1!delivery_detail_id) + ", " + CStr(rsaux1!organization_id) + ",'" + IIf(IsNull(rsaux1!subinventory), "", rsaux1!subinventory) + "', " + CStr(rsaux1!delivery_line_id) + "," + CStr(rsaux1!inventory_item_id) + ",'" + rsaux1!item_description + "','" + CStr(rsaux1!SOURCE_LINE_NUMBER) + "'," + CStr(1) + ",'" + rsaux1!released_status + "',"
                                    var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + rsaux1!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux1!collector_id), 0, rsaux1!collector_id)) + ",'" + IIf(IsNull(rsaux1!Name), "", rsaux1!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux1!ubicacion), "", rsaux1!ubicacion) + "','" + IIf(IsNull(rsaux1!establecimiento), "", rsaux1!establecimiento) + "','" + IIf(IsNull(rsaux1!nombre_Establecimiento), "", rsaux1!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux1!ORDENES), "", rsaux1!ORDENES) + "',"
                                    var_cadena = var_cadena + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux1!source_header_type_name), "", rsaux1!source_header_type_name) + "', '" + IIf(IsNull(rsaux1!source_document_id), "", rsaux1!source_document_id) + "','" + IIf(IsNull(rsaux1!pais), "", rsaux1!pais) + "','" + IIf(IsNull(rsaux1!estado), "", rsaux1!estado) + "', '" + IIf(IsNull(rsaux1!municipio), "", rsaux1!municipio) + "', '" + IIf(IsNull(rsaux1!ciudad), "", rsaux1!ciudad) + "', '" + IIf(IsNull(rsaux1!colonia), "", rsaux1!colonia) + "','" + IIf(IsNull(rsaux1!DIRECCION), "", rsaux1!DIRECCION) + "', '" + IIf(IsNull(rsaux1!cp), "", rsaux1!cp) + "',"
                                    var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!site_use_id), "", rsaux1!site_use_id) + "','" + IIf(IsNull(rsaux1!paqueteria), "", rsaux1!paqueteria) + "','" + IIf(IsNull(rsaux1!attribute8), "", rsaux1!attribute8) + "','" + IIf(IsNull(rsaux1!ATTRIBUTE9), "", rsaux1!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux1!TELEFONO), "", rsaux1!TELEFONO) + "','" + IIf(IsNull(rsaux1!Linea), "", rsaux1!Linea) + "','" + CStr(IIf(IsNull(rsaux1!ruta), "", rsaux1!ruta)) + "','" + IIf(IsNull(rsaux1!nombre_ruta), "", rsaux1!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux1!ORDEN_SURTIDO), 0, rsaux1!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux1!Embarque), 0, rsaux1!Embarque)) + ", '" + IIf(IsNull(rsaux1!estacion), "", rsaux1!estacion) + "'," + CStr(IIf(IsNull(rsaux1!MULTIPLO), "", rsaux1!MULTIPLO)) + ")"
                                    rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                    var_cantidad = var_cantidad - 1
                              Wend
                           End If
                        End If
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  rsaux1.Open "delete from tb_Temp_oracle_orden_surtido_aux_1", cnn, adOpenDynamic, adLockOptimistic
                  'rsaux1.Open "select distinct source_header_number, ORDEN_pedido from tb_Temp_oracle_orden_surtido_aux_2, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SOURCE_HEADER_NUMBER = pedido  order by ORDEN_pedido", cnn, adOpenDynamic, adLockOptimistic
                  'rsaux1.Open "select distinct  source_header_number, SUBSTRING(ubicacion,1,3) as pasillo from tb_Temp_oracle_orden_surtido_aux_2, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SOURCE_HEADER_NUMBER = pedido order by SUBSTRING(ubicacion,1,3)", cnn, adOpenDynamic, adLockOptimistic
                  rsaux1.Open "select distinct  SUBSTRING(ubicacion,1,3) as pasillo from tb_Temp_oracle_orden_surtido_aux_2, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SOURCE_HEADER_NUMBER = pedido order by SUBSTRING(ubicacion,1,3)", cnn, adOpenDynamic, adLockOptimistic
                  var_contador = 0
                  var_grupo = 1
                  var_lote = 1
                  While Not rsaux1.EOF
                        'var_lote = 1
                        
                        If var_contador >= 60 Then
                           var_grupo = var_grupo + 1
                           var_contador = 0
                        End If
                        'rsaux2.Open "select * from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SUBSTRING(ubicacion,1,3) = '" + rsaux1!pasillo + "'  and source_header_number = " + CStr(rsaux1!source_header_number) + " order by ubicacion", cnn, adOpenDynamic, adLockOptimistic
                        rsaux2.Open "select * from tb_Temp_oracle_orden_surtido_aux_2 a, tb_oracle_pedidos_asignados_embarques b where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and SUBSTRING(ubicacion,1,3) = '" + rsaux1!pasillo + "' and a.source_header_number = b.pedido order by orden_pedido, source_header_number, ubicacion ", cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux2.EOF
                              If var_contador >= 60 Then
                                 var_lote = var_lote + 1
                                 var_contador = 0
                                 var_grupo = var_grupo + 1
                              End If
                              'rsaux3.Open "SELECT * FROM tb_Temp_oracle_orden_surtido_aux_1 WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and delivery_detail_id = " + CStr(rsaux2!delivery_detail_id) + " and SUBSTRING(ubicacion,1,3) = '" + rsaux1!pasillo + "'", cnn, adOpenDynamic, adLockOptimistic
                              rsaux3.Open "SELECT * FROM tb_Temp_oracle_orden_surtido_aux_1 WHERE  segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and delivery_detail_id = " + CStr(rsaux2!delivery_detail_id) + " and SUBSTRING(ubicacion,1,3) = '" + rsaux1!pasillo + "'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 rsaux4.Open "UPDATE TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 SET SRC_REQUESTED_QUANTITY = SRC_REQUESTED_QUANTITY + " + CStr(rsaux2!src_requested_quantity) + " WHERE source_header_number = '" + CStr(rsaux2!source_header_number) + "' AND segment1 = '" + rsaux2!SEGMENT1 + "' AND LOTE = " + CStr(var_lote) + " and DELIVERY_DETAIL_ID = " + CStr(rsaux2!delivery_detail_id) + " and SUBSTRING(ubicacion,1,3) = '" + rsaux1!pasillo + "'", cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 var_dia = CStr(Day(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_mes = CStr(Month(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 var_año = CStr(Year(IIf(IsNull(rsaux2!DATE_REQUESTED), Now, rsaux2!DATE_REQUESTED)))
                                 If Len(Trim(var_dia)) = 1 Then
                                    var_dia = "0" + var_dia
                                 End If
                                 If Len(Trim(var_mes)) = 1 Then
                                    var_mes = "0" + var_mes
                                 End If
                                 var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                 
                                 var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                 var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,LOTE, GRUPO) "
                                 var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux2!source_header_number + "', " + CStr(rsaux2!delivery_id) + "," + CStr(rsaux2!delivery_detail_id) + ", " + CStr(rsaux2!organization_id) + ",'" + IIf(IsNull(rsaux2!subinventory), "", rsaux2!subinventory) + "', " + CStr(rsaux2!delivery_line_id) + "," + CStr(rsaux2!inventory_item_id) + ",'" + rsaux2!item_description + "','" + CStr(rsaux2!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux2!src_requested_quantity) + ",'" + rsaux2!released_status + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!customer_name), "", rsaux2!customer_name) + "','" + rsaux2!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux2!collector_id), 0, rsaux2!collector_id)) + ",'" + IIf(IsNull(rsaux2!Name), "", rsaux2!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux2!ubicacion), "", rsaux2!ubicacion) + "','" + IIf(IsNull(rsaux2!establecimiento), "", rsaux2!establecimiento) + "','" + IIf(IsNull(rsaux2!nombre_Establecimiento), "", rsaux2!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux2!ORDENES), "", rsaux2!ORDENES) + "',"
                                 var_cadena = var_cadena + CStr(IIf(IsNull(rsaux2!CUST_ACCOUNT_ID), 0, rsaux2!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux2!source_header_type_name), "", rsaux2!source_header_type_name) + "', '" + IIf(IsNull(rsaux2!source_document_id), "", rsaux2!source_document_id) + "','" + IIf(IsNull(rsaux2!pais), "", rsaux2!pais) + "','" + IIf(IsNull(rsaux2!estado), "", rsaux2!estado) + "', '" + IIf(IsNull(rsaux2!municipio), "", rsaux2!municipio) + "', '" + IIf(IsNull(rsaux2!ciudad), "", rsaux2!ciudad) + "', '" + IIf(IsNull(rsaux2!colonia), "", rsaux2!colonia) + "','" + IIf(IsNull(rsaux2!DIRECCION), "", rsaux2!DIRECCION) + "', '" + IIf(IsNull(rsaux2!cp), "", rsaux2!cp) + "',"
                                 var_cadena = var_cadena + "'" + IIf(IsNull(rsaux2!site_use_id), "", rsaux2!site_use_id) + "','" + IIf(IsNull(rsaux2!paqueteria), "", rsaux2!paqueteria) + "','" + IIf(IsNull(rsaux2!attribute8), "", rsaux2!attribute8) + "','" + IIf(IsNull(rsaux2!ATTRIBUTE9), "", rsaux2!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux2!TELEFONO), "", rsaux2!TELEFONO) + "','" + IIf(IsNull(rsaux2!Linea), "", rsaux2!Linea) + "','" + CStr(IIf(IsNull(rsaux2!ruta), "", rsaux2!ruta)) + "','" + IIf(IsNull(rsaux2!nombre_ruta), "", rsaux2!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux2!ORDEN_SURTIDO), 0, rsaux2!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", '" + CStr(IIf(IsNull(rsaux2!estacion), "", rsaux2!estacion)) + "'," + CStr(var_lote) + "," + CStr(var_grupo) + ")"
                                 rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux3.Close
                              If rsaux2!Linea <> "CATALOGOS" Then
                                 If rsaux2!Linea <> "CATALOGO" Then
                                    If rsaux2!Linea <> "POP" Then
                                       If rsaux2!Linea <> "EMPAQUE" Then
                                          If rsaux2!MULTIPLO > 1 Then
                                             var_contador = var_contador + rsaux2!src_requested_quantity
                                          Else
                                             var_contador = var_contador + 1
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                              rsaux2.MoveNext
                        Wend
                        rsaux2.Close
                        var_contador = 60
                        var_lote = var_lote + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  
                  rsaux1.Open "insert TB_TEMP_ORACLE_ORDEN_SURTIDO (inte_tem_consecutivo, segment1) values (" + CStr(var_consecutivo) + ",'---------')", cnn, adOpenDynamic, adLockOptimistic
                  rsaux1.Open "delete from TB_TEMP_ORACLE_ORDEN_SURTIDO where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and segment1 <> '---------'", cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.Open "insert into TB_TEMP_ORACLE_ORDEN_SURTIDO select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.Open "delete from TB_TEMP_ORACLE_ORDEN_SURTIDO where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and segment1 = '---------'", cnn, adOpenDynamic, adLockOptimistic
                  var_consecutivo_general = var_consecutivo
                  
                  
                  
                  
                  Call crea_tablas
                  If rsaux.State = 1 Then
                     rsaux.Close
                  End If
                  rsaux.Open "select distinct a.source_header_number from tb_Temp_oracle_orden_surtido_aux_1 a, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES  b where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and a.source_header_number = b.pedido", cnn, adOpenDynamic, adLockOptimistic
                  var_cadena_pedidos_diferencias = ""
                  While Not rsaux.EOF
                        strconsulta = "select sum(requested_quantity)  as cantidad from WSH_DELIVERABLES_V where source_header_number = ? AND RELEASED_STATUS = 'Y'"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(rsaux!source_header_number))
                             .Parameters.Append parametro
                        End With
                        Set rsaux10 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        
                        strconsulta = "SELECT SUM(SRC_REQUESTED_QUANTITY) AS CANTIDAD FROM XXVIA_TB_PEDIDOS_DIVIDIDOS WHERE SOURCE_HEADER_NUMBER = ?"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(rsaux!source_header_number))
                             .Parameters.Append parametro
                        End With
                        Set rsaux11 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                     
                     
                        If IIf(IsNull(rsaux11!cantidad), 0, rsaux11!cantidad) <> rsaux10!cantidad Then
                           If var_cadena_pedidos_diferencias = "" Then
                              var_cadena_pedidos_diferencias = CStr(rsaux!source_header_number)
                           Else
                              var_cadena_pedidos_diferencias = var_cadena_pedidos_diferencias + ", " + CStr(rsaux!source_header_number)
                           End If
                        End If
                        rsaux10.Close
                        rsaux11.Close
                        rsaux.MoveNext
                  Wend
                  rsaux.Close
                  If var_cadena_pedidos_diferencias = "" Then
                     If var_imprime_pedidos = 1 Then
                        ' orden
x:
                        'var_consecutivo = 1360
                        rsaux.Open "select distinct a.grupo from tb_Temp_oracle_orden_surtido_aux_1 a, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES  b where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and a.source_header_number = b.pedido order by a.grupo", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           While Not rsaux.EOF
                                 x = 1
                                 If x = 1 Then
                                    
                                    'strconsulta = "select shipping_method_code, packing_instructions from oe_order_headers_all where order_number = ?"
                                    'With comandoORA
                                    '     .ActiveConnection = cnnoracle_4
                                    '     .CommandType = adCmdText
                                    '     .CommandText = strconsulta
                                    '     Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rsaux!source_header_number))
                                    '     .Parameters.Append parametro
                                    'End With
                                    'Set rs = comandoORA.execute
                                    'Set comandoORA = Nothing
                                    'Set parametro = Nothing
                                    
                                    var_paqueteria = ""
                                    'If Not rs.EOF Then
                                    '   VAR_COMENTARIOS = IIf(IsNull(rs!packing_instructions), "", rs!packing_instructions)
                                    '   var_tipo_metodo = IIf(IsNull(rs(0).Value), "", rs(0).Value)
                                    '   If var_tipo_metodo <> "" Then
                                    '
                                    '      strconsulta = "SELECT description FROM fnd_lookup_values where lookup_type = 'SHIP_METHOD' and lookup_code = ? AND LANGUAGE = 'ESA'"
                                    '      With comandoORA
                                    '           .ActiveConnection = cnno-racle_4
                                    '           .CommandType = adCmdText
                                    '           .CommandText = strconsulta
                                    '           Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_tipo_metodo)
                                    '           .Parameters.Append parametro
                                    '      End With
                                    '      Set rsaux1 = comandoORA.execute
                                    '      Set comandoORA = Nothing
                                    '      Set parametro = Nothing
                                          
                                    '      If Not rsaux1.EOF Then
                                    '         var_paqueteria = IIf(IsNull(rsaux1(0).Value), "", rsaux1(0).Value)
                                    '      End If
                                    '      rsaux1.Close
                                    '   End If
                                    'End If
                                    'rs.Close
                                    
                                    VAR_ZZ = 0
                                    If VAR_ZZ = 1 Then
                                        
                                    
                                    
                                    
                                       strconsulta = "SELECT  HPS.party_site_id as tel, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1, OHA.packing_instructions from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.ship_TO_ORG_ID AND oha.order_number = ? AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                                       With comandoORA
                                            .ActiveConnection = cnnoracle_4
                                            .CommandType = adCmdText
                                            .CommandText = strconsulta
                                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rsaux!source_header_number))
                                            .Parameters.Append parametro
                                       End With
                                       Set rsaux6 = comandoORA.execute
                                       Set comandoORA = Nothing
                                       Set parametro = Nothing
                                       
                                       If Not rsaux6.EOF Then
                                       
                                          
                                          strconsulta = "SELECT  HPS.party_site_id as tel, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1, OHA.packing_instructions from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.ship_TO_ORG_ID AND oha.order_number = ? AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                                          With comandoORA
                                               .ActiveConnection = cnnoracle_4
                                               .CommandType = adCmdText
                                               .CommandText = strconsulta
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, IIf(IsNull(rsaux!source_header_number), "", rsaux!source_header_number))
                                               .Parameters.Append parametro
                                          End With
                                          Set rsaux5 = comandoORA.execute
                                          Set comandoORA = Nothing
                                          Set parametro = Nothing
                                          
                                          'rsaux5.Open "SELECT  HPS.party_site_id as tel, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1, OHA.packing_instructions from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.ship_TO_ORG_ID AND oha.order_number = '" + IIf(IsNull(rsaux!source_header_number), "", rsaux!source_header_number) + "' AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                          If Not rsaux5.EOF Then
                                             var_nombre = IIf(IsNull(rsaux5!customer_name), "", rsaux5!customer_name)
                                             var_tel = IIf(IsNull(rsaux5!tel), 0, rsaux5!tel)
                                             VAR_DIRECCION = IIf(IsNull(rsaux5!calle), "", rsaux5!calle) + " " + IIf(IsNull(rsaux5!numero), "", rsaux5!numero)
                                             VAR_COLONIA = IIf(IsNull(rsaux5!colonia), "", rsaux5!colonia)
                                             var_ciudad = IIf(IsNull(rsaux5!ciudad), "", rsaux5!ciudad)
                                             VAR_MUNICIPIO = IIf(IsNull(rsaux5!municipio), "", rsaux5!municipio)
                                             var_estado = IIf(IsNull(rsaux5!estado), "", rsaux5!estado)
                                             var_pais = IIf(IsNull(rsaux5!pais), "", rsaux5!pais)
                                             VAR_CP = IIf(IsNull(rsaux5!cp), "", rsaux5!cp)
                                             rsaux5.Close
                                          Else
                                             rsaux5.Close
                                             var_nombre = IIf(IsNull(rsaux6!customer_name), "", rsaux6!customer_name)
                                             var_tel = IIf(IsNull(rsaux6!tel), 0, rsaux6!tel)
                                             VAR_DIRECCION = IIf(IsNull(rsaux6!calle), "", rsaux6!calle) + " " + IIf(IsNull(rsaux6!numero), "", rsaux6!numero)
                                             VAR_COLONIA = IIf(IsNull(rsaux6!colonia), "", rsaux6!colonia)
                                             var_ciudad = IIf(IsNull(rsaux6!ciudad), "", rsaux6!ciudad)
                                             VAR_MUNICIPIO = IIf(IsNull(rsaux6!municipio), "", rsaux6!municipio)
                                             var_estado = IIf(IsNull(rsaux6!estado), "", rsaux6!estado)
                                             var_pais = IIf(IsNull(rsaux6!pais), "", rsaux6!pais)
                                             VAR_CP = IIf(IsNull(rsaux6!cp), "", rsaux6!cp)
                                          End If
                                       Else
                                          var_tel = 0
                                          VAR_DIRECCION = ""
                                          VAR_COLONIA = ""
                                          var_ciudad = ""
                                          VAR_MUNICIPIO = ""
                                          var_estado = ""
                                          var_pais = ""
                                          VAR_CP = ""
                                       End If
                                       rsaux6.Close
                                       If var_tel > 0 Then
                                          
                                          strconsulta = "select Phone_Number from hz_contact_points where owner_table_id = ?"
                                          With comandoORA
                                               .ActiveConnection = cnnoracle_4
                                               .CommandType = adCmdText
                                               .CommandText = strconsulta
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(var_tel))
                                               .Parameters.Append parametro
                                          End With
                                          Set rsaux6 = comandoORA.execute
                                          Set comandoORA = Nothing
                                          Set parametro = Nothing
                                          If Not rsaux6.EOF Then
                                             var_telefono = CStr(IIf(IsNull(rsaux6(0).Value), "", rsaux6(0).Value))
                                          Else
                                             var_telefono = ""
                                          End If
                                          rsaux6.Close
                                       Else
                                          var_telefono = ""
                                       End If
                                    Else
                                       var_tel = 0
                                       VAR_DIRECCION = ""
                                       VAR_COLONIA = ""
                                       var_ciudad = ""
                                       VAR_MUNICIPIO = ""
                                       var_estado = ""
                                       var_pais = ""
                                       VAR_CP = ""
                                       var_telefono = ""
                                    End If
                                    
                                    If IsNumeric(txt_total_volumen) Then
                                       var_cubicaje_EMBARQUE = CDbl(Me.txt_total_volumen)
                                    Else
                                       var_cubicaje_EMBARQUE = 0
                                    End If
                                    
                                    rsaux4.Open "select distinct source_header_number from tb_Temp_oracle_orden_surtido where  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and grupo = " + CStr(rsaux(0).Value), cnn, adOpenDynamic, adLockOptimistic
                                    While Not rsaux4.EOF
                                          rsaux2.Open "SELECT * FROM tb_oracle_pedidos_asignados_embarques WHERE PEDIDO = " + CStr(rsaux4(0).Value), cnn, adOpenDynamic, adLockOptimistic
                                          If Not rsaux2.EOF Then
                                             var_transporte = ""
                                             rsaux3.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                             If Not rsaux3.EOF Then
                                                var_transporte = IIf(IsNull(rsaux3!vehiculo), "", rsaux3!vehiculo)
                                             End If
                                             rsaux3.Close
                                             
                                             
                                             strconsulta = "SELECT ORDER_TYPE_ID FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ?"
                                             With comandoORA
                                                  .ActiveConnection = cnnoracle_4
                                                  .CommandType = adCmdText
                                                  .CommandText = strconsulta
                                                  Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rsaux4(0).Value))
                                                  .Parameters.Append parametro
                                             End With
                                             Set rsaux5 = comandoORA.execute
                                             Set comandoORA = Nothing
                                             Set parametro = Nothing
                                             If rsaux5!ORDER_TYPE_ID = 1002 Then
                                                rsaux5.Close
                                                strconsulta = "select ADDRESS_LINE_1||', '||ADDRESS_LINE_2||', '||TOWN_OR_CITY||', '||REGION_1||', '||COUNTRY||' CP:'||POSTAL_CODE DIRECCION, EMAIL from mtl_secondary_inventories a, hr_locations_all b, xxvia_jv_tb_agentes c, po_requisition_headers_ALL D, OE_ORDER_HEADERS_ALL E Where A.location_id = b.location_id and a.secondary_inventory_name = c.subinventory_code AND E.source_document_id = D.requisition_header_id AND A.secondary_inventory_name = D.ATTRIBUTE1 AND E.ORDER_NUMBER = ?"
                                             Else
                                                rsaux5.Close
                                                strconsulta = "SELECT CALLE||' '||NUM_CALLE||' '||NVL(num_interior,'')||', '||colonia||', '||ciudad||', '||estado||', '||pais||' CP: '||codigo_postal DIRECCION FROM oe_order_headers_all a, xxvia_vw_CLIENTES_BCP B WHERE A.SHIP_TO_ORG_ID = B.SITE_USE_ID AND A.ORDER_NUMBER     = ?"
                                             End If
                                             
                                             With comandoORA
                                                  .ActiveConnection = cnnoracle_4
                                                  .CommandType = adCmdText
                                                  .CommandText = strconsulta
                                                  Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rsaux4(0).Value))
                                                  .Parameters.Append parametro
                                             End With
                                             Set rsaux5 = comandoORA.execute
                                             Set comandoORA = Nothing
                                             Set parametro = Nothing
                                             If Not rsaux5.EOF Then
                                                VAR_DIRECCION = IIf(IsNull(rsaux5!DIRECCION), "", rsaux5!DIRECCION)
                                             Else
                                                VAR_DIRECCION = ""
                                             End If
                                             
                                             
                                             
                                             var_cadena = "UPDATE tb_Temp_oracle_orden_surtido SET EMBARQUE = " + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", CUBICAJE = " + CStr(var_cubicaje_EMBARQUE) + " , ORDEN_SURTIDO = " + CStr(IIf(IsNull(rsaux2!orden_pedido), 0, rsaux2!orden_pedido)) + ", ANDEN = '" + CStr(IIf(IsNull(rsaux2!estacion), 0, rsaux2!estacion)) + "', TRANSPORTE = '" + var_transporte + "',"
                                             var_cadena = var_cadena + " pais= '" + var_pais + "', estado = '" + var_estado + "', municipio = '" + VAR_MUNICIPIO + "', ciudad = '" + var_ciudad + "', colonia = '" + VAR_COLONIA + "', direccion = '" + VAR_DIRECCION + "', cp = '" + VAR_CP + "', paqueteria = '" + var_paqueteria + "'"
                                             var_cadena = var_cadena + " WHERE SOURCE_HEADER_NUMBER = " + CStr(rsaux4(0).Value) + " AND inte_tem_consecutivo = " + CStr(var_consecutivo)
                                             rsaux3.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                          End If
                                          rsaux2.Close
                                          rsaux4.MoveNext
                                    Wend
                                    rsaux4.Close
                                    
                                                                  
                                    x = 1
                                    If x = 1 Then
                                       If rsaux9.State = 1 Then
                                          rsaux9.Close
                                       End If
                                       rsaux9.Open "SELECT DISTINCT  GRUPO, SOURCE_HEADER_NUMBER from tb_Temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and grupo = " + CStr(rsaux(0).Value), cnn, adOpenDynamic, adLockOptimistic
                                       VAR_TOTAL_GRUPOS = rsaux9.RecordCount
                                       rsaux9.Close
                                       'VAR_TOTAL_GRUPOS = 1
                                       rsaux10.Open "update tb_Temp_oracle_orden_surtido set cubicaje = " + CStr(var_cubicaje_EMBARQUE) + "  where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                                       If VAR_TOTAL_GRUPOS = 1 Then
                                          Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos_pasillos.rpt")
                                          reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.grupo} = " + CStr(rsaux(0).Value) + " and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                                          'frmvista_previa_auxiliar.cr2.ReportSource = reporte
                                          For ntablas = 1 To reporte.Database.Tables.Count
                                              reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                          Next ntablas
                                          
                                          
                                          reporte.PrintOut False
                                          'frmvista_previa_auxiliar.cr2.ViewReport
                                          
                                          'frmvista_previa_auxiliar.Show 1
                                          
                                          Set reporte = Nothing
                                          
                                       Else
                                          
                                          Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos_ENCABEZADOS.rpt")
                                          reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.grupo} = " + CStr(rsaux(0).Value) + " and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                                          'frmvista_previa_auxiliar.cr2.ReportSource = reporte
                                          
                                          'frmvista_previa_auxiliar.cr2.ReportSource = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos_ENCABEZADOS.rpt")
                                          For ntablas = 1 To reporte.Database.Tables.Count
                                              reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                          Next ntablas
                                          reporte.PrintOut False
                                          'frmvista_previa_auxiliar.cr2.ViewReport
                                          
                                          'frmvista_previa_auxiliar.Show 1
                                          
            'frmvistasprevias.cr.ViewReport
            'frmvistasprevias.Caption = "Packing List"
            'frmvistasprevias.Show 1
                                          
                                          Set reporte = Nothing
                                          
                                          Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos_DETALLE_pasillo.rpt")
                                          reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.grupo} = " + CStr(rsaux(0).Value) + " and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                                          
                                          'frmvista_previa_auxiliar.cr2.ReportSource = reporte
                                          'frmvista_previa_auxiliar.cr2.ReportSource = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos_DETALLE.rpt")
                                          
                                          
                                          
                                          For ntablas = 1 To reporte.Database.Tables.Count
                                              reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                          Next ntablas
                                          reporte.PrintOut False
            'frmvistasprevias.cr.ViewReport
            'frmvistasprevias.Caption = "Packing List"
            'frmvistasprevias.Show 1
                                          'frmvista_previa_auxiliar.cr2.ViewReport
                                          'frmvista_previa_auxiliar.Show 1
                                          Set reporte = Nothing
                                       
                                       End If
                                       
                                       
                                       'Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos.rpt")
                                       'reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.grupo} = " + CStr(rsaux(0).Value) + " and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                                       'For ntablas = 1 To reporte.Database.Tables.Count
                                       '    reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                       'Next ntablas
                                       'reporte.PrintOut False
                                       'Set reporte = Nothing
                                    Else
                                       If rsaux9.State = 1 Then
                                          rsaux9.Close
                                       End If
                                       rsaux9.Open "SELECT DISTINCT  GRUPO, SOURCE_HEADER_NUMBER from tb_Temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and grupo = " + CStr(rsaux(0).Value), cnn, adOpenDynamic, adLockOptimistic
                                       VAR_TOTAL_GRUPOS = rsaux9.RecordCount
                                       rsaux9.Close
                                       'VAR_TOTAL_GRUPOS = 1
                                       If VAR_TOTAL_GRUPOS = 1 Then
                                          Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos.rpt")
                                          reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.grupo} = " + CStr(rsaux(0).Value) + " and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                                          frmvistasprevias.cr.ReportSource = reporte
                                          For ntablas = 1 To reporte.Database.Tables.Count
                                              reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                          Next ntablas
                                          frmvistasprevias.cr.ViewReport
                                          frmvistasprevias.Caption = "Ordenes de surtido "
                                          frmvistasprevias.Show 1
                                          Set reporte = Nothing
                                       Else
                                          Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos_ENCABEZADOS.rpt")
                                          reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.grupo} = " + CStr(rsaux(0).Value) + " and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                                          frmvistasprevias.cr.ReportSource = reporte
                                          For ntablas = 1 To reporte.Database.Tables.Count
                                              reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                          Next ntablas
                                          frmvistasprevias.cr.ViewReport
                                          frmvistasprevias.Caption = "Ordenes de surtido "
                                          frmvistasprevias.Show 1
                                          Set reporte = Nothing
                                          
                                          Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos_DETALLE_pasillo.rpt")
                                          reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.grupo} = " + CStr(rsaux(0).Value) + " and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                                          frmvistasprevias.cr.ReportSource = reporte
                                          For ntablas = 1 To reporte.Database.Tables.Count
                                              reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                          Next ntablas
                                          frmvistasprevias.cr.ViewReport
                                          frmvistasprevias.Caption = "Ordenes de surtido "
                                          frmvistasprevias.Show 1
                                          Set reporte = Nothing
                                       
                                       
                                       End If
                                       
                                    End If
                                 End If
                                 rsaux.MoveNext
                           Wend
                        End If
                        rsaux.Close
                     Else
                        MsgBox "No se pueden imprimir los pedidos, vuelva a intentar la impresión", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "No se pueden imprimir los pedidos, vuelva a intentar la impresión", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "No se pueden imprimir los pedidos, vuelva a intentar la impresión", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "No existen ordenes de surtido", vbOKOnly, "ATENCION"
            End If
            If rs.State = 1 Then
               rs.Close
            End If
            If var_consecutivo > 0 Then
               rs.Open "delete from tb_Temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(IIf(IsNull(var_consecutivo), 0, var_consecutivo)), cnn, adOpenDynamic, adLockOptimistic
            End If
         Else
            'MsgBox "Número superior incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "No existen pedidos a imprimir"
      End If
   End If
   End If
   
End Sub

Private Sub cmd_lotes_en_picking_Click()
   If IsNumeric(Me.txt_embarque) Then
      rs.Open "select * from tb_oracle_pedidos_asignados_embarqueS WHERE EMBARQUE = '" + Me.txt_embarque + "'", cnn, adOpenDynamic, adLockOptimistic
      var_Cadena_pedidos = ""
      While Not rs.EOF
            If var_Cadena_pedidos = "" Then
               var_Cadena_pedidos = CStr(rs!pedido)
            Else
               var_Cadena_pedidos = var_Cadena_pedidos + "," + CStr(rs!pedido)
            End If
            rs.MoveNext
      Wend
      rs.Close
      If var_Cadena_pedidos <> "" Then
         rsaux10.Open "select * from xxvia_tb_pedidos_divididos where source_header_number in (" + var_Cadena_pedidos + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rsaux10.EOF Then
            var_si = MsgBox("¿Desea imprimir el reporte de lotes en picking?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               
               cnn.BeginTrans
               rs.Open "select isnull(max(inte_tem_consecutivo),0) as consecutivo from TB_TEMPORAL_LOTES_EN_PICKING", cnn, adOpenDynamic, adLockOptimistic
               rsaux.Open "insert into TB_TEMPORAL_LOTES_EN_PICKING (inte_tem_consecutivo) values (" + CStr(rs!CONSECUTIVO + 1) + ")", cnn, adOpenDynamic, adLockOptimistic
               var_consecutivo = rs!CONSECUTIVO + 1
               rs.Close
               cnn.CommitTrans
               
               rs.Open "select * from tb_ORACLE_pedidos_asignados_embarques where embarque = " + Me.txt_embarque + " ORDER BY ORDEN_PEDIDO", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  While Not rs.EOF
                        If rs!Cliente = "VIANNEY TEXTIL HOGAR SA DE CV" Then
                           var_cliente = rs!NOMBRE_AGENTE
                        Else
                           var_cliente = rs!Cliente
                        End If
                        strconsulta = "SELECT DISTINCT LOTE FROM XXVIA_TB_PEDIDOS_DIVIDIDOS WHERE SOURCE_HEADER_NUMBER = ? ORDER BY LOTE"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(rs!pedido))
                             .Parameters.Append parametro
                        End With
                        Set rsaux2 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        If Not rsaux2.EOF Then
                           var_total_lotes = 0
                           While Not rsaux2.EOF
                                 var_total_lotes = var_total_lotes + 1
                                 
                                 var_cadena = "select lote, NVL(max(attribute2),' ') ubicacion_inicio, NVL(min(attribute2), ' ') ubicacion_final from xxvia_tb_pedidos_divididos a, xxvia_system_items_b b where source_header_number = ? and a.segment1 = b.segment1 and b.organization_id= ? AND LOTE = ? group by lote"
                                 With comandoORA
                                      .ActiveConnection = cnnoracle_4
                                      .CommandType = adCmdText
                                      .CommandText = var_cadena
                                      Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(rs!pedido))
                                      .Parameters.Append parametro
                                      Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(var_unidad_organizacional))
                                      .Parameters.Append parametro
                                      Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(rsaux2!LOTE))
                                      .Parameters.Append parametro
                                 End With
                                 Set rsaux14 = comandoORA.execute
                                 Set comandoORA = Nothing
                                 Set parametro = Nothing
                                 If Not rsaux14.EOF Then
                                    VAR_UBICACION_INICIO = rsaux14!UBICACION_INICIO
                                    VAR_UBICACION_FINAL = rsaux14!UBICACION_FINAL
                                 Else
                                    VAR_UBICACION_INICIO = ""
                                    VAR_UBICACION_FINAL = ""
                                 End If
                                 rsaux14.Close
                                 var_cadena = "INSERT INTO TB_TEMPORAL_LOTES_EN_PICKING (INTE_TEM_CONSECUTIVO,EMBARQUE,PEDIDO,CLIENTE,TOTAL_LOTES,LOTE, ORDEN_PEDIDO, UBICACION_INICIO, UBICACION_FINAL) Values (" + CStr(var_consecutivo) + ",'" + Me.txt_embarque + "','" + rs!pedido + "','" + Mid(var_cliente, 1, 50) + "'," + CStr(var_total_lotes) + " ," + CStr(rsaux2!LOTE) + "," + CStr(IIf(IsNull(rs!orden_pedido), 0, rs!orden_pedido)) + ",'" + VAR_UBICACION_INICIO + "','" + VAR_UBICACION_FINAL + "')"
                              
                                 rsaux3.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              
                                 rsaux2.MoveNext
                           Wend
                           rsaux2.Close
                           rsaux2.Open "update TB_TEMPORAL_LOTES_EN_PICKING set total_lotes = " + CStr(var_total_lotes) + " where inte_tem_consecutivo =" + CStr(var_consecutivo) + " and pedido = " + CStr(rs!pedido), cnn, adOpenDynamic, adLockOptimistic
                        End If
                        rs.MoveNext
                       
                  Wend
               
                  Set reporte = appl.OpenReport(App.Path + "\rep_oracle_lotes_en_picking.rpt")
                  reporte.RecordSelectionFormula = "{VW_ORACLE_LOTES_EN_PICKING.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  reporte.PrintOut False
                  Set reporte = Nothing
               End If
            End If
            rsaux10.Close
         Else
            MsgBox "No se han generado los lotes del embarque " + Me.txt_embarque, vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
      End If
      If rs.State = 1 Then
         rs.Close
      End If
   Else
      MsgBox "Debe de seleccionar un embarque", vbOKOnly, "ATENCION"
   End If
   
   
   
End Sub

Private Sub cmd_orden_ubicacion_310317_Click()
'GoTo x:
      If rs.State = 1 Then
         rs.Close
      End If
      If rsaux.State = 1 Then
         rsaux.Close
      End If
      If rsaux1.State = 1 Then
         rsaux1.Close
      End If
      If rsaux2.State = 1 Then
         rsaux2.Close
      End If
      If rsaux3.State = 1 Then
         rsaux3.Close
      End If
      If rsaux4.State = 1 Then
         rsaux4.Close
      End If
      If rsaux5.State = 1 Then
         rsaux5.Close
      End If
      If rsaux6.State = 1 Then
         rsaux6.Close
      End If
      If rsaux7.State = 1 Then
         rsaux7.Close
      End If
      If rsaux8.State = 1 Then
         rsaux8.Close
      End If
      If rsaux9.State = 1 Then
         rsaux9.Close
      End If
     
      If Me.lv_pedidos.ListItems.Count > 0 Then
         var_si = MsgBox("Desea imprimir las ordenes de surtido?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            If rs.State = 1 Then
               rs.Close
            End If
            var_Cadena_pedidos = ""
            For var_j = 1 To Me.lv_pedidos.ListItems.Count
                Me.lv_pedidos.ListItems.Item(var_j).Selected = True
                If Me.lv_pedidos.selectedItem <> "10000000" Then
                   If var_Cadena_pedidos = "" Then
                      var_Cadena_pedidos = Me.lv_pedidos.selectedItem
                   Else
                      var_Cadena_pedidos = var_Cadena_pedidos + "," + Me.lv_pedidos.selectedItem
                   End If
                End If
            Next var_j
            'var_cadena_pedidos = "105208"
            rs.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_cadena = "SELECT to_char(a.LAST_UPDATE_DATE,'day') DIA_SEMANA, CAT.LINEA, HCSU.SITE_USE_ID, HCAS.CUST_ACCOUNT_ID, OHA.SHIP_TO_ORG_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id, a.LAST_UPDATE_DATE,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,c.description,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, a.source_header_type_name, oha.source_document_id, C.ATTRIBUTE2, oha.attribute8, oha.attribute9, j.NAME as nombre_ruta, j.salesrep_id as clave_ruta from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, xxvia_vw_articulos_cat cat, XXVIA_VENDEDORES j Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND "
            var_cadena = var_cadena + " to_number(source_header_number)  IN (" + var_Cadena_pedidos + ")"
            var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID "
            var_cadena = var_cadena + " AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND released_status = 'Y' AND A.inventory_item_id  = cat.item_id AND A.ORGANIZATION_ID = Cat.organization_id and oha.salesrep_id = j.salesrep_id "
            rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
'--------------------------
            If Not rs.EOF Then
               cnn.BeginTrans
               rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM tb_Temp_oracle_orden_surtido", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
               Else
                  var_consecutivo = 1
               End If
               rsaux.Close
               rsaux1.Open "insert into tb_Temp_oracle_orden_surtido(inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
               cnn.CommitTrans
               While Not rs.EOF
                     var_establecimiento = rs!SHIP_TO_ORG_ID
                     var_clave_cliente = rs!site_use_id
                     rsaux.Open "SELECT csu.site_use_id AS VCHA_ESB_ESTABLECIMIENTO_ID, ps.party_site_number, lo.address1 AS VCHA_eSB_NOMBRE FROM hz_cust_site_uses_all csu, hz_cust_acct_sites_all cas, hz_party_sites ps, hz_locations lo Where csu.cust_acct_site_id = cas.cust_acct_site_id AND cas.party_site_id = ps.party_site_id AND ps.location_id = lo.location_id AND csu.site_use_code = 'SHIP_TO' AND csu.LOCATION = ps.party_site_number and csu.site_use_id = " + CStr(var_establecimiento), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        VAR_NOMBRE_ESTABLECIMIENTO = IIf(IsNull(rsaux!party_site_number), "", rsaux!party_site_number) + " " + IIf(IsNull(rsaux!vcha_esb_nombre), "", rsaux!vcha_esb_nombre)
                     Else
                        VAR_NOMBRE_ESTABLECIMIENTO = ""
                     End If
                     rsaux.Close
                     
                     rsaux.Open "SELECT csu.site_use_id AS VCHA_ESB_ESTABLECIMIENTO_ID, ps.party_site_number, lo.address1 AS VCHA_eSB_NOMBRE FROM hz_cust_site_uses_all csu, hz_cust_acct_sites_all cas, hz_party_sites ps, hz_locations lo Where csu.cust_acct_site_id = cas.cust_acct_site_id AND cas.party_site_id = ps.party_site_id AND ps.location_id = lo.location_id AND csu.site_use_code = 'BILL_TO' AND csu.LOCATION = ps.party_site_number and csu.site_use_id = " + CStr(var_clave_cliente), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        VAR_CLAVE_CLIENTE_BCP = IIf(IsNull(rsaux!party_site_number), "", rsaux!party_site_number)
                     Else
                        VAR_CLAVE_CLIENTE_BCP = ""
                     End If
                     rsaux.Close
                     
                     
                     
                     var_dia = CStr(Day(CDate(rs!LAST_UPDATE_DATE)))
                     var_mes = CStr(Month(CDate(rs!LAST_UPDATE_DATE)))
                     var_año = CStr(Year(CDate(rs!LAST_UPDATE_DATE)))
                     If Len(Trim(var_dia)) = 1 Then
                        var_dia = "0" + var_dia
                     End If
                     If Len(Trim(var_mes)) = 1 Then
                        var_mes = "0" + var_mes
                     End If
                     var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                     rsaux1.Open "select * from tb_oracle_multiplos where segment1 = '" + IIf(IsNull(rs!SEGMENT1), "", rs!SEGMENT1) + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux1.EOF Then
                        VAR_MULTIPLO = IIf(IsNull(rsaux1!MULTIPLO), 1, rsaux1!MULTIPLO)
                     Else
                        VAR_MULTIPLO = 1
                     End If
                     rsaux1.Close
'''''
                     If rs!source_header_type_name = "VIA_PEDIDO_INTERNO" Then
                        rsaux1.Open "SELECT * FROM TB_ORACLE_ARTICULOS_MOTOR_LOGISTICO WHERE CODIGO = '" + rs!SEGMENT1 + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux1.EOF Then
                           strconsulta = "SELECT secondary_inventory_name, A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rs!source_document_id)
                                .Parameters.Append parametro
                           End With
                           Set rsaux8 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           If rsaux8.EOF Then
                              var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                           Else
                              var_almacen = rsaux8!secondary_inventory_name
                              rsaux9.Open "SELECT * FROM TB_ORACLE_UBICACIONES_MOTOR_LOGISTICO WHERE CLAVE = '" + var_almacen + "' AND CODIGO = '" + rs!SEGMENT1 + "'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux9.EOF Then
                                 var_ubicacion = ""
                                 If Me.cmb_dia.Text = "Lunes" Then
                                    var_ubicacion = rsaux9!ubicacion_1
                                 End If
                                 If Me.cmb_dia.Text = "Martes" Then
                                    var_ubicacion = rsaux9!ubicacion_2
                                 End If
                                 If Me.cmb_dia.Text = "Miercoles" Then
                                    var_ubicacion = rsaux9!ubicacion_3
                                 End If
                                 If Me.cmb_dia.Text = "Jueves" Then
                                    var_ubicacion = rsaux9!ubicacion_4
                                 End If
                                 If Me.cmb_dia.Text = "Viernes" Then
                                    var_ubicacion = rsaux9!ubicacion_5
                                 End If
                                 If Me.cmb_dia.Text = "Sabado" Then
                                    var_ubicacion = rsaux9!ubicacion_6
                                 End If
                                 If IIf(IsNull(var_ubicacion), "", var_ubicacion) = "" Then
                                    var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                                 End If
                              Else
                                 var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                              End If
                              rsaux9.Close
                           End If
                           rsaux8.Close
                        Else
                           var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                        End If
                        rsaux1.Close
                     Else
                        var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
                     End If
                     
                     
'''''
                     var_cadena = "insert into tb_Temp_oracle_orden_surtido(INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER, DELIVERY_ID, DELIVERY_DETAIL_ID, ORGANIZATION_ID, SUBINVENTORY, DELIVERY_LINE_ID, INVENTORY_ITEM_ID, ITEM_DESCRIPTION, SOURCE_LINE_NUMBER, src_requested_quantity, RELEASED_STATUS, CUSTOMER_NAME, SEGMENT1, COLLECTOR_ID, NAME, date_requested, UBICACION, ESTABLECIMIENTO, NOMBRE_ESTABLECIMIENTO, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, SITE_USE_ID, attribute8, attribute9, LINEA, RUTA, NOMBRE_RUTA,MULTIPLO)  values "
                     var_cadena = var_cadena + " (" + CStr(var_consecutivo) + ", '" + IIf(IsNull(rs!source_header_number), "", rs!source_header_number) + "', " + CStr(IIf(IsNull(rs!delivery_id), 0, rs!delivery_id)) + ", " + CStr(IIf(IsNull(rs!delivery_detail_id), 0, rs!delivery_detail_id)) + ", " + CStr(IIf(IsNull(rs!organization_id), 0, rs!organization_id)) + ", '" + IIf(IsNull(rs!subinventory), "", rs!subinventory) + "', " + CStr(IIf(IsNull(rs!delivery_line_id), 0, rs!delivery_line_id)) + ", " + CStr(IIf(IsNull(rs!inventory_item_id), "", rs!inventory_item_id)) + ", '" + IIf(IsNull(rs!Description), "", rs!Description) + "', '" + IIf(IsNull(rs!SOURCE_LINE_NUMBER), "", rs!SOURCE_LINE_NUMBER) + "', " + CStr(IIf(IsNull(rs!requested_quantity), 0, rs!requested_quantity)) + ", '" + IIf(IsNull(rs!released_status), "", rs!released_status) + "', '" + VAR_CLAVE_CLIENTE_BCP + " " + Replace(IIf(IsNull(rs!customer_name), "", rs!customer_name), "'", " ") + "', '" + IIf(IsNull(rs!SEGMENT1), "", rs!SEGMENT1) + "'"
                     'var_cadena = var_cadena + ", 0,''," + var_fecha + ",'" + IIf(IsNull(rs!ATTRIBUTE2), "", rs!ATTRIBUTE2) + "','" + CStr(VAR_ESTABLECIMIENTO) + "','" + Replace(VAR_NOMBRE_ESTABLECIMIENTO, "'", " ") + "'," + CStr(rs!CUST_ACCOUNT_ID) + ",'" + rs!source_header_type_name + "','" + CStr(IIf(IsNull(rs!source_document_id), "", rs!source_document_id)) + "'," + CStr(IIf(IsNull(rs!SITE_USE_ID), 0, rs!SITE_USE_ID)) + ",'" + IIf(IsNull(rs!attribute8), "", rs!attribute8) + "','" + IIf(IsNull(rs!ATTRIBUTE9), "", rs!ATTRIBUTE9) + "','" + IIf(IsNull(rs!linea), "", rs!linea) + "','" + CStr(rs!CLAVE_RUTA) + "','" + IIf(IsNull(rs!nombre_ruta), "", rs!nombre_ruta) + "'," + CStr(VAR_MULTIPLO) + ")"
                     var_cadena = var_cadena + ", 0,''," + var_fecha + ",'" + var_ubicacion + "','" + CStr(var_establecimiento) + "','" + Replace(VAR_NOMBRE_ESTABLECIMIENTO, "'", " ") + "'," + CStr(rs!CUST_ACCOUNT_ID) + ",'" + rs!source_header_type_name + "','" + CStr(IIf(IsNull(rs!source_document_id), "", rs!source_document_id)) + "'," + CStr(IIf(IsNull(rs!site_use_id), 0, rs!site_use_id)) + ",'" + IIf(IsNull(rs!attribute8), "", rs!attribute8) + "','" + IIf(IsNull(rs!ATTRIBUTE9), "", rs!ATTRIBUTE9) + "','" + IIf(IsNull(rs!Linea), "", rs!Linea) + "','" + CStr(rs!CLAVE_RUTA) + "','" + IIf(IsNull(rs!nombre_ruta), "", rs!nombre_ruta) + "'," + CStr(VAR_MULTIPLO) + ")"
                     rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               rsaux1.Open "delete from tb_Temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number is null", cnn, adOpenDynamic, adLockOptimistic
               If rsaux1.State = 1 Then
                  rsaux1.Close
               End If
               
               var_cadena_pedidos_diferencias = ""
               rsaux1.Open "select source_header_number, sum(src_requested_quantity) as cantidad from tb_Temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo) + " group by source_header_number", cnn, adOpenDynamic, adLockOptimistic
               While Not rsaux1.EOF
                     strconsulta = "select sum(requested_quantity)  as cantidad from WSH_DELIVERABLES_V where source_header_number = ? AND RELEASED_STATUS = 'Y'"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(rsaux1!source_header_number))
                          .Parameters.Append parametro
                     End With
                     Set rsaux10 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     If rsaux1!cantidad <> rsaux10!cantidad Then
                        If var_cadena_pedidos_diferencias = "" Then
                           var_cadena_pedidos_diferencias = CStr(rsaux1!source_header_number)
                        Else
                           var_cadena_pedidos_diferencias = var_cadena_pedidos_diferencias + ", " + CStr(rsaux1!source_header_number)
                        End If
                     End If
                     rsaux10.Close
                     rsaux1.MoveNext
               Wend
               rsaux1.Close
               
               If var_cadena_pedidos_diferencias = "" Then
                  rsaux1.Open "select min(DATE_REQUESTED) as DATE_REQUESTED, source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME, RUTA, NOMBRE_RUTA, SUM(SRC_REQUESTED_QUANTITY) AS CANTIDAD from tb_Temp_oracle_orden_surtido where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " group by source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME, RUTA, NOMBRE_RUTA", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        If rsaux1!source_header_type_name = "VIA_PEDIDO_INTERNO" Or rsaux1!source_header_type_name = "TEX_PEDIDO_INTERNO" Then
                           If var_pedido_tienda = 0 Then
                              If rsaux2.State = 1 Then
                                 rsaux2.Close
                              End If
                              rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(rsaux1!source_document_id) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux2.EOF Then
                                 If rsaux4.State = 1 Then
                                    rsaux4.Close
                                 End If
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                                 rsaux4.Close
                              Else
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                                 rsaux4.Close
                              End If
                              rsaux2.Close
                           Else
                              rsaux2.Open "select a.attribute8, B.description from oe_order_headers_all a, MTL_SECONDARY_INVENTORIES b where order_number = " + CStr(rsaux1!source_header_number) + " and a.attribute8 = b.secondary_inventory_name", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux2.EOF Then
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                                 rsaux4.Close
                              Else
                                 rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 VAR_PROVEEDOR = rsaux4!collector_id
                                 VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                                 rsaux4.Close
                              End If
                              rsaux2.Close
                           End If
                        Else
                           rsaux6.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + " and SITE_USE_ID = " + CStr(IIf(IsNull(rsaux1!site_use_id), 0, rsaux1!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                           VAR_PROVEEDOR = rsaux6!collector_id
                           VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux6!Name), "", rsaux6!Name)
                           rsaux6.Close
                        End If
                        var_año_str = CStr(Year(rsaux1!DATE_REQUESTED))
                        If Len(var_año_str) < 2 Then
                           var_año_str = "20" + var_año_str
                        End If
                        var_mes_str = CStr(Month(rsaux1!DATE_REQUESTED))
                        If Len(var_mes_str) < 2 Then
                           var_mes_str = "0" + var_mes_str
                        End If
                        var_dia_str = CStr(Day(rsaux1!DATE_REQUESTED))
                        If Len(var_dia_str) < 2 Then
                           var_dia_str = "0" + var_dia_str
                        End If
                        var_hora_str = CStr(Hour(rsaux1!DATE_REQUESTED))
                        If Len(var_hora_str) < 2 Then
                           var_hora_str = "0" + var_hora_str
                        End If
                        VAR_MINUTO_STR = CStr(Minute(rsaux1!DATE_REQUESTED))
                        If Len(VAR_MINUTO_STR) < 2 Then
                           VAR_MINUTO_STR = "0" + VAR_MINUTO_STR
                        End If
                        VAR_SEGUNDO_STR = CStr(Second(rsaux1!DATE_REQUESTED))
                        If Len(VAR_SEGUNDO_STR) < 2 Then
                           VAR_SEGUNDO_STR = "0" + VAR_SEGUNDO_STR
                        End If
                        var_fecha_pedido = var_año_str + "-" + var_mes_str + "-" + var_dia_str + " " + var_hora_str + ":" + VAR_MINUTO_STR + ":" + VAR_SEGUNDO_STR
                        rsaux6.Open "update tb_Temp_oracle_orden_surtido set COLLECTOR_ID = " + CStr(VAR_PROVEEDOR) + ", NAME = '" + VAR_NOMBRE_PROVEEDOR + "' where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " and CUST_ACCOUNT_ID = " + CStr(rsaux1!CUST_ACCOUNT_ID) + " and source_header_number = " + CStr(rsaux1!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                        rsaux6.Open "SELECT * FROM XXVIA_TB_ORDENES_GRAFICA WHERE ORGANIZACION = " + var_unidad_organizacional + " AND PEDIDO = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If rsaux6.EOF Then
                           rsaux5.Open "INSERT INTO XXVIA_TB_ORDENES_GRAFICA (ORGANIZACION, PEDIDO, PROVEEDOR_ID, NOMBRE_PROVEEDOR, CLIENTE, ESTABLECIMIENTO, CANTIDAD, FECHA, TIPO_PEDIDO, CANTIDAD_LEIDA, FECHA_PEDIDO, RUTA, NOMBRE_RUTA) VALUES (" + var_unidad_organizacional + ", " + CStr(rsaux1!source_header_number) + "," + CStr(VAR_PROVEEDOR) + ",'" + VAR_NOMBRE_PROVEEDOR + "','" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "'," + CStr(rsaux1!cantidad) + ", '" + CStr(rsaux1!DATE_REQUESTED) + "','" + rsaux1!source_header_type_name + "',0, TO_DATE('" + var_fecha_pedido + "','YYYY-MM-DD HH24:MI:SS'),'" + rsaux1!ruta + "', '" + rsaux1!nombre_ruta + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        Else
                           rsaux5.Open "UPDATE XXVIA_TB_ORDENES_GRAFICA SET RUTA = '" + rsaux1!ruta + "', NOMBRE_RUTA = '" + rsaux1!nombre_ruta + "' WHERE PEDIDO = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        End If
                        rsaux6.Close
                        rsaux6.Open "select * from tb_oracle_pedidos_asignados_embarques where pedido = " + CStr(rsaux1!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                        If rsaux6.EOF Then
                           rsaux5.Open "insert into tb_oracle_pedidos_asignados_embarques (AGENTE, NOMBRE_AGENTE, PEDIDO, CLIENTE, PIEZAS, embarque, dia,  mes, AÑO, ORGANIZACION) values (" + CStr(VAR_PROVEEDOR) + ",'" + VAR_NOMBRE_PROVEEDOR + "'," + CStr(rsaux1!source_header_number) + ",'" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "', " + CStr(rsaux1!cantidad) + ",0," + var_dia_str + "," + var_mes_str + "," + var_año_str + "," + CStr(var_unidad_organizacional) + ")", cnn, adOpenDynamic, adLockOptimistic
                        End If
                        rsaux6.Close
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  rsaux1.Open "select min(DATE_REQUESTED) as DATE_REQUESTED, source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME, SUM(SRC_REQUESTED_QUANTITY) AS CANTIDAD from tb_Temp_oracle_orden_surtido where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " AND (LINEA NOT IN ('CATALOGOS','CATALOGO','POP') OR LINEA IS NULL) group by source_header_number, CUST_ACCOUNT_ID, source_header_type_name, source_document_id, site_use_id, NOMBRE_ESTABLECIMIENTO, CUSTOMER_NAME"
                  While Not rsaux1.EOF
                        rsaux5.Open "UPDATE XXVIA_TB_ORDENES_GRAFICA SET CANTIDAD_SIN_CATALOGOS = " + CStr(IIf(IsNull(rsaux1!cantidad), 0, rsaux1!cantidad)) + " WHERE PEDIDO = " + CStr(rsaux1!source_header_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  rsaux1.Open "DELETE from tb_Temp_oracle_orden_surtido_aux_2", cnn, adOpenDynamic, adLockOptimistic
                  rsaux1.Open "SELECT * FROM tb_Temp_oracle_orden_surtido where inte_tem_consecutivo =  " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux1.EOF
                        If rsaux1!Linea = "CATALOGOS" Or rsaux1!Linea = "CATALOGO" Or rsaux1!Linea = "POP" Or rsaux1!Linea = "EMPAQUE" Then
                           var_dia = CStr(Day(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                           var_mes = CStr(Month(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                           var_año = CStr(Year(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                           If Len(Trim(var_dia)) = 1 Then
                              var_dia = "0" + var_dia
                           End If
                           If Len(Trim(var_mes)) = 1 Then
                              var_mes = "0" + var_mes
                           End If
                           var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                           var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_2 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                           var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,MULTIPLO) "
                           var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux1!source_header_number + "', " + CStr(rsaux1!delivery_id) + "," + CStr(rsaux1!delivery_detail_id) + ", " + CStr(rsaux1!organization_id) + ",'" + IIf(IsNull(rsaux1!subinventory), "", rsaux1!subinventory) + "', " + CStr(rsaux1!delivery_line_id) + "," + CStr(rsaux1!inventory_item_id) + ",'" + rsaux1!item_description + "','" + CStr(rsaux1!SOURCE_LINE_NUMBER) + "'," + CStr(rsaux1!src_requested_quantity) + ",'" + rsaux1!released_status + "',"
                           var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + rsaux1!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux1!collector_id), 0, rsaux1!collector_id)) + ",'" + IIf(IsNull(rsaux1!Name), "", rsaux1!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux1!ubicacion), "", rsaux1!ubicacion) + "','" + IIf(IsNull(rsaux1!establecimiento), "", rsaux1!establecimiento) + "','" + IIf(IsNull(rsaux1!nombre_Establecimiento), "", rsaux1!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux1!ORDENES), "", rsaux1!ORDENES) + "',"
                           var_cadena = var_cadena + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux1!source_header_type_name), "", rsaux1!source_header_type_name) + "', '" + IIf(IsNull(rsaux1!source_document_id), "", rsaux1!source_document_id) + "','" + IIf(IsNull(rsaux1!pais), "", rsaux1!pais) + "','" + IIf(IsNull(rsaux1!estado), "", rsaux1!estado) + "', '" + IIf(IsNull(rsaux1!municipio), "", rsaux1!municipio) + "', '" + IIf(IsNull(rsaux1!ciudad), "", rsaux1!ciudad) + "', '" + IIf(IsNull(rsaux1!colonia), "", rsaux1!colonia) + "','" + IIf(IsNull(rsaux1!DIRECCION), "", rsaux1!DIRECCION) + "', '" + IIf(IsNull(rsaux1!cp), "", rsaux1!cp) + "',"
                           var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!site_use_id), "", rsaux1!site_use_id) + "','" + IIf(IsNull(rsaux1!paqueteria), "", rsaux1!paqueteria) + "','" + IIf(IsNull(rsaux1!attribute8), "", rsaux1!attribute8) + "','" + IIf(IsNull(rsaux1!ATTRIBUTE9), "", rsaux1!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux1!TELEFONO), "", rsaux1!TELEFONO) + "','" + IIf(IsNull(rsaux1!Linea), "", rsaux1!Linea) + "','" + CStr(IIf(IsNull(rsaux1!ruta), "", rsaux1!ruta)) + "','" + IIf(IsNull(rsaux1!nombre_ruta), "", rsaux1!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux1!ORDEN_SURTIDO), 0, rsaux1!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux1!Embarque), 0, rsaux1!Embarque)) + ", '" + IIf(IsNull(rsaux1!estacion), "", rsaux1!estacion) + "'," + CStr(IIf(IsNull(rsaux1!MULTIPLO), "", rsaux1!MULTIPLO)) + ")"
                           rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        Else
                           var_cantidad_total = rsaux1!src_requested_quantity
                           If rsaux1!MULTIPLO > 1 Then
                              While var_cantidad_total > 0
                                    If var_cantidad_total < rsaux1!MULTIPLO Then
                                       var_cantidad = var_cantidad_total
                                    Else
                                       var_cantidad = rsaux1!MULTIPLO
                                    End If
                                    
                                    var_dia = CStr(Day(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                    var_mes = CStr(Month(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                    var_año = CStr(Year(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                    If Len(Trim(var_dia)) = 1 Then
                                       var_dia = "0" + var_dia
                                    End If
                                    If Len(Trim(var_mes)) = 1 Then
                                      var_mes = "0" + var_mes
                                    End If
                                    
                                    var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                    var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_2 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                    var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,MULTIPLO, PASILLO) "
                                    var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux1!source_header_number + "', " + CStr(rsaux1!delivery_id) + "," + CStr(rsaux1!delivery_detail_id) + ", " + CStr(rsaux1!organization_id) + ",'" + IIf(IsNull(rsaux1!subinventory), "", rsaux1!subinventory) + "', " + CStr(rsaux1!delivery_line_id) + "," + CStr(rsaux1!inventory_item_id) + ",'" + rsaux1!item_description + "','" + CStr(rsaux1!SOURCE_LINE_NUMBER) + "'," + CStr(var_cantidad) + ",'" + rsaux1!released_status + "',"
                                    var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + rsaux1!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux1!collector_id), 0, rsaux1!collector_id)) + ",'" + IIf(IsNull(rsaux1!Name), "", rsaux1!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux1!ubicacion), "", rsaux1!ubicacion) + "','" + IIf(IsNull(rsaux1!establecimiento), "", rsaux1!establecimiento) + "','" + IIf(IsNull(rsaux1!nombre_Establecimiento), "", rsaux1!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux1!ORDENES), "", rsaux1!ORDENES) + "',"
                                    var_cadena = var_cadena + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux1!source_header_type_name), "", rsaux1!source_header_type_name) + "', '" + IIf(IsNull(rsaux1!source_document_id), "", rsaux1!source_document_id) + "','" + IIf(IsNull(rsaux1!pais), "", rsaux1!pais) + "','" + IIf(IsNull(rsaux1!estado), "", rsaux1!estado) + "', '" + IIf(IsNull(rsaux1!municipio), "", rsaux1!municipio) + "', '" + IIf(IsNull(rsaux1!ciudad), "", rsaux1!ciudad) + "', '" + IIf(IsNull(rsaux1!colonia), "", rsaux1!colonia) + "','" + IIf(IsNull(rsaux1!DIRECCION), "", rsaux1!DIRECCION) + "', '" + IIf(IsNull(rsaux1!cp), "", rsaux1!cp) + "',"
                                    var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!site_use_id), "", rsaux1!site_use_id) + "','" + IIf(IsNull(rsaux1!paqueteria), "", rsaux1!paqueteria) + "','" + IIf(IsNull(rsaux1!attribute8), "", rsaux1!attribute8) + "','" + IIf(IsNull(rsaux1!ATTRIBUTE9), "", rsaux1!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux1!TELEFONO), "", rsaux1!TELEFONO) + "','" + IIf(IsNull(rsaux1!Linea), "", rsaux1!Linea) + "','" + CStr(IIf(IsNull(rsaux1!ruta), "", rsaux1!ruta)) + "','" + IIf(IsNull(rsaux1!nombre_ruta), "", rsaux1!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux1!ORDEN_SURTIDO), 0, rsaux1!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux1!Embarque), 0, rsaux1!Embarque)) + ", '" + IIf(IsNull(rsaux1!estacion), "", rsaux1!estacion) + "'," + CStr(IIf(IsNull(rsaux1!MULTIPLO), "", rsaux1!MULTIPLO)) + ",'" + IIf(IsNull(rsaux1!pasillo), "", rsaux1!pasillo) + "')"
                                    rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                    var_cantidad_total = var_cantidad_total - rsaux1!MULTIPLO
                              Wend
                           Else
                              var_cantidad = var_cantidad_total
                              While var_cantidad > 0
                                    var_dia = CStr(Day(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                    var_mes = CStr(Month(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                    var_año = CStr(Year(IIf(IsNull(rsaux1!DATE_REQUESTED), Now, rsaux1!DATE_REQUESTED)))
                                    If Len(Trim(var_dia)) = 1 Then
                                       var_dia = "0" + var_dia
                                    End If
                                    If Len(Trim(var_mes)) = 1 Then
                                      var_mes = "0" + var_mes
                                    End If
                                    var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                    var_cadena = "INSERT INTO TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_2 (INTE_TEM_CONSECUTIVO, SOURCE_HEADER_NUMBER,DELIVERY_ID,DELIVERY_DETAIL_ID,ORGANIZATION_ID,SUBINVENTORY,DELIVERY_LINE_ID,INVENTORY_ITEM_ID,ITEM_DESCRIPTION,SOURCE_LINE_NUMBER,SRC_REQUESTED_QUANTITY,RELEASED_STATUS,CUSTOMER_NAME,SEGMENT1,COLLECTOR_ID,NAME,DATE_REQUESTED,UBICACION,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,ORDENES,CUST_ACCOUNT_ID,source_header_type_name,source_document_id,PAIS,ESTADO,MUNICIPIO,CIUDAD,COLONIA,DIRECCION,CP,SITE_USE_ID,PAQUETERIA,ATTRIBUTE8,ATTRIBUTE9"
                                    var_cadena = var_cadena + ",TELEFONO,LINEA,RUTA,NOMBRE_RUTA,ORDEN_SURTIDO,EMBARQUE,ESTACION,MULTIPLO,PASILLO) "
                                    var_cadena = var_cadena + "Values (" + CStr(var_consecutivo) + ",'" + rsaux1!source_header_number + "', " + CStr(rsaux1!delivery_id) + "," + CStr(rsaux1!delivery_detail_id) + ", " + CStr(rsaux1!organization_id) + ",'" + IIf(IsNull(rsaux1!subinventory), "", rsaux1!subinventory) + "', " + CStr(rsaux1!delivery_line_id) + "," + CStr(rsaux1!inventory_item_id) + ",'" + rsaux1!item_description + "','" + CStr(rsaux1!SOURCE_LINE_NUMBER) + "'," + CStr(1) + ",'" + rsaux1!released_status + "',"
                                    var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!customer_name), "", rsaux1!customer_name) + "','" + rsaux1!SEGMENT1 + "'," + CStr(IIf(IsNull(rsaux1!collector_id), 0, rsaux1!collector_id)) + ",'" + IIf(IsNull(rsaux1!Name), "", rsaux1!Name) + "', " + CStr(var_fecha) + ", '" + IIf(IsNull(rsaux1!ubicacion), "", rsaux1!ubicacion) + "','" + IIf(IsNull(rsaux1!establecimiento), "", rsaux1!establecimiento) + "','" + IIf(IsNull(rsaux1!nombre_Establecimiento), "", rsaux1!nombre_Establecimiento) + "', '" + IIf(IsNull(rsaux1!ORDENES), "", rsaux1!ORDENES) + "',"
                                    var_cadena = var_cadena + CStr(IIf(IsNull(rsaux1!CUST_ACCOUNT_ID), 0, rsaux1!CUST_ACCOUNT_ID)) + ",'" + IIf(IsNull(rsaux1!source_header_type_name), "", rsaux1!source_header_type_name) + "', '" + IIf(IsNull(rsaux1!source_document_id), "", rsaux1!source_document_id) + "','" + IIf(IsNull(rsaux1!pais), "", rsaux1!pais) + "','" + IIf(IsNull(rsaux1!estado), "", rsaux1!estado) + "', '" + IIf(IsNull(rsaux1!municipio), "", rsaux1!municipio) + "', '" + IIf(IsNull(rsaux1!ciudad), "", rsaux1!ciudad) + "', '" + IIf(IsNull(rsaux1!colonia), "", rsaux1!colonia) + "','" + IIf(IsNull(rsaux1!DIRECCION), "", rsaux1!DIRECCION) + "', '" + IIf(IsNull(rsaux1!cp), "", rsaux1!cp) + "',"
                                    var_cadena = var_cadena + "'" + IIf(IsNull(rsaux1!site_use_id), "", rsaux1!site_use_id) + "','" + IIf(IsNull(rsaux1!paqueteria), "", rsaux1!paqueteria) + "','" + IIf(IsNull(rsaux1!attribute8), "", rsaux1!attribute8) + "','" + IIf(IsNull(rsaux1!ATTRIBUTE9), "", rsaux1!ATTRIBUTE9) + "','" + IIf(IsNull(rsaux1!TELEFONO), "", rsaux1!TELEFONO) + "','" + IIf(IsNull(rsaux1!Linea), "", rsaux1!Linea) + "','" + CStr(IIf(IsNull(rsaux1!ruta), "", rsaux1!ruta)) + "','" + IIf(IsNull(rsaux1!nombre_ruta), "", rsaux1!nombre_ruta) + "'," + CStr(IIf(IsNull(rsaux1!ORDEN_SURTIDO), 0, rsaux1!ORDEN_SURTIDO)) + "," + CStr(IIf(IsNull(rsaux1!Embarque), 0, rsaux1!Embarque)) + ", '" + IIf(IsNull(rsaux1!estacion), "", rsaux1!estacion) + "'," + CStr(IIf(IsNull(rsaux1!MULTIPLO), "", rsaux1!MULTIPLO)) + ",'" + IIf(IsNull(rsaux1!pasillo), "", rsaux1!pasillo) + "')"
                                    rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                    var_cantidad = var_cantidad - 1
                              Wend
                           End If
                        End If
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  
                  
                  
                  
                  rsaux1.Open "delete from tb_Temp_oracle_orden_surtido_aux_1", cnn, adOpenDynamic, adLockOptimistic
                  var_consecutivo_general = var_consecutivo
Call PASILLOS
                  'rsaux1.Open "select distinct source_header_number, ORDEN_SURTIDO from tb_Temp_oracle_orden_surtido_aux_2 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " order by ORDEN_SURTIDO", cnn, adOpenDynamic, adLockOptimistic

                  
                 ' MsgBox var_consecutivo
                  rsaux1.Open "insert TB_TEMP_ORACLE_ORDEN_SURTIDO (inte_tem_consecutivo, segment1) values (" + CStr(var_consecutivo) + ",'---------')", cnn, adOpenDynamic, adLockOptimistic
                  rsaux1.Open "delete from TB_TEMP_ORACLE_ORDEN_SURTIDO where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and segment1 <> '---------'", cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.Open "insert into TB_TEMP_ORACLE_ORDEN_SURTIDO select * from TB_TEMP_ORACLE_ORDEN_SURTIDO_AUX_1 where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.Open "delete from TB_TEMP_ORACLE_ORDEN_SURTIDO where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and segment1 = '---------'", cnn, adOpenDynamic, adLockOptimistic
                  var_consecutivo_general = var_consecutivo
                  Call crea_tablas
                  If rsaux.State = 1 Then
                     rsaux.Close
                  End If
                  rsaux.Open "select distinct a.source_header_number from tb_Temp_oracle_orden_surtido_aux_1 a, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES  b where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and a.source_header_number = b.pedido", cnn, adOpenDynamic, adLockOptimistic
                  var_cadena_pedidos_diferencias = ""
                  While Not rsaux.EOF
                        strconsulta = "select sum(requested_quantity)  as cantidad from WSH_DELIVERABLES_V where source_header_number = ? AND RELEASED_STATUS = 'Y'"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(rsaux!source_header_number))
                             .Parameters.Append parametro
                        End With
                        Set rsaux10 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        
                        strconsulta = "SELECT SUM(SRC_REQUESTED_QUANTITY) AS CANTIDAD FROM XXVIA_TB_PEDIDOS_DIVIDIDOS WHERE SOURCE_HEADER_NUMBER = ?"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(rsaux!source_header_number))
                             .Parameters.Append parametro
                        End With
                        Set rsaux11 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                     
                     
                        If IIf(IsNull(rsaux11!cantidad), 0, rsaux11!cantidad) <> rsaux10!cantidad Then
                           If var_cadena_pedidos_diferencias = "" Then
                              var_cadena_pedidos_diferencias = CStr(rsaux!source_header_number)
                           Else
                              var_cadena_pedidos_diferencias = var_cadena_pedidos_diferencias + ", " + CStr(rsaux!source_header_number)
                           End If
                        End If
                        rsaux10.Close
                        rsaux11.Close
                        rsaux.MoveNext
                  Wend
                  rsaux.Close
                  If var_cadena_pedidos_diferencias = "" Then
                     If var_imprime_pedidos = 1 Then
                        ' orden
x:
                        'var_consecutivo = 1360
                        rsaux.Open "select distinct a.grupo from tb_Temp_oracle_orden_surtido_aux_1 a, TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES  b where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and a.source_header_number = b.pedido order by a.grupo", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           While Not rsaux.EOF
                                 x = 1
                                 If x = 1 Then
                                    
                                    'strconsulta = "select shipping_method_code, packing_instructions from oe_order_headers_all where order_number = ?"
                                    'With comandoORA
                                    '     .ActiveConnection = cnnoracle_4
                                    '     .CommandType = adCmdText
                                    '     .CommandText = strconsulta
                                    '     Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rsaux!source_header_number))
                                    '     .Parameters.Append parametro
                                    'End With
                                    'Set rs = comandoORA.execute
                                    'Set comandoORA = Nothing
                                    'Set parametro = Nothing
                                    
                                    var_paqueteria = ""
                                    'If Not rs.EOF Then
                                    '   VAR_COMENTARIOS = IIf(IsNull(rs!packing_instructions), "", rs!packing_instructions)
                                    '   var_tipo_metodo = IIf(IsNull(rs(0).Value), "", rs(0).Value)
                                    '   If var_tipo_metodo <> "" Then
                                    '
                                    '      strconsulta = "SELECT description FROM fnd_lookup_values where lookup_type = 'SHIP_METHOD' and lookup_code = ? AND LANGUAGE = 'ESA'"
                                    '      With comandoORA
                                    '           .ActiveConnection = cnno-racle_4
                                    '           .CommandType = adCmdText
                                    '           .CommandText = strconsulta
                                    '           Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_tipo_metodo)
                                    '           .Parameters.Append parametro
                                    '      End With
                                    '      Set rsaux1 = comandoORA.execute
                                    '      Set comandoORA = Nothing
                                    '      Set parametro = Nothing
                                          
                                    '      If Not rsaux1.EOF Then
                                    '         var_paqueteria = IIf(IsNull(rsaux1(0).Value), "", rsaux1(0).Value)
                                    '      End If
                                    '      rsaux1.Close
                                    '   End If
                                    'End If
                                    'rs.Close
                                    
                                    VAR_ZZ = 0
                                    If VAR_ZZ = 1 Then
                                       strconsulta = "SELECT  HPS.party_site_id as tel, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1, OHA.packing_instructions from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.ship_TO_ORG_ID AND oha.order_number = ? AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                                       With comandoORA
                                            .ActiveConnection = cnnoracle_4
                                            .CommandType = adCmdText
                                            .CommandText = strconsulta
                                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rsaux!source_header_number))
                                            .Parameters.Append parametro
                                       End With
                                       Set rsaux6 = comandoORA.execute
                                       Set comandoORA = Nothing
                                       Set parametro = Nothing
                                       
                                       If Not rsaux6.EOF Then
                                       
                                          
                                          strconsulta = "SELECT  HPS.party_site_id as tel, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1, OHA.packing_instructions from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.ship_TO_ORG_ID AND oha.order_number = ? AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                                          With comandoORA
                                               .ActiveConnection = cnnoracle_4
                                               .CommandType = adCmdText
                                               .CommandText = strconsulta
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, IIf(IsNull(rsaux!source_header_number), "", rsaux!source_header_number))
                                               .Parameters.Append parametro
                                          End With
                                          Set rsaux5 = comandoORA.execute
                                          Set comandoORA = Nothing
                                          Set parametro = Nothing
                                          
                                          'rsaux5.Open "SELECT  HPS.party_site_id as tel, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1, OHA.packing_instructions from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.ship_TO_ORG_ID AND oha.order_number = '" + IIf(IsNull(rsaux!source_header_number), "", rsaux!source_header_number) + "' AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                          If Not rsaux5.EOF Then
                                             var_nombre = IIf(IsNull(rsaux5!customer_name), "", rsaux5!customer_name)
                                             var_tel = IIf(IsNull(rsaux5!tel), 0, rsaux5!tel)
                                             VAR_DIRECCION = IIf(IsNull(rsaux5!calle), "", rsaux5!calle) + " " + IIf(IsNull(rsaux5!numero), "", rsaux5!numero)
                                             VAR_COLONIA = IIf(IsNull(rsaux5!colonia), "", rsaux5!colonia)
                                             var_ciudad = IIf(IsNull(rsaux5!ciudad), "", rsaux5!ciudad)
                                             VAR_MUNICIPIO = IIf(IsNull(rsaux5!municipio), "", rsaux5!municipio)
                                             var_estado = IIf(IsNull(rsaux5!estado), "", rsaux5!estado)
                                             var_pais = IIf(IsNull(rsaux5!pais), "", rsaux5!pais)
                                             VAR_CP = IIf(IsNull(rsaux5!cp), "", rsaux5!cp)
                                             rsaux5.Close
                                          Else
                                             rsaux5.Close
                                             var_nombre = IIf(IsNull(rsaux6!customer_name), "", rsaux6!customer_name)
                                             var_tel = IIf(IsNull(rsaux6!tel), 0, rsaux6!tel)
                                             VAR_DIRECCION = IIf(IsNull(rsaux6!calle), "", rsaux6!calle) + " " + IIf(IsNull(rsaux6!numero), "", rsaux6!numero)
                                             VAR_COLONIA = IIf(IsNull(rsaux6!colonia), "", rsaux6!colonia)
                                             var_ciudad = IIf(IsNull(rsaux6!ciudad), "", rsaux6!ciudad)
                                             VAR_MUNICIPIO = IIf(IsNull(rsaux6!municipio), "", rsaux6!municipio)
                                             var_estado = IIf(IsNull(rsaux6!estado), "", rsaux6!estado)
                                             var_pais = IIf(IsNull(rsaux6!pais), "", rsaux6!pais)
                                             VAR_CP = IIf(IsNull(rsaux6!cp), "", rsaux6!cp)
                                          End If
                                       Else
                                          var_tel = 0
                                          VAR_DIRECCION = ""
                                          VAR_COLONIA = ""
                                          var_ciudad = ""
                                          VAR_MUNICIPIO = ""
                                          var_estado = ""
                                          var_pais = ""
                                          VAR_CP = ""
                                       End If
                                       rsaux6.Close
                                       If var_tel > 0 Then
                                          
                                          strconsulta = "select Phone_Number from hz_contact_points where owner_table_id = ?"
                                          With comandoORA
                                               .ActiveConnection = cnnoracle_4
                                               .CommandType = adCmdText
                                               .CommandText = strconsulta
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(var_tel))
                                               .Parameters.Append parametro
                                          End With
                                          Set rsaux6 = comandoORA.execute
                                          Set comandoORA = Nothing
                                          Set parametro = Nothing
                                          If Not rsaux6.EOF Then
                                             var_telefono = CStr(IIf(IsNull(rsaux6(0).Value), "", rsaux6(0).Value))
                                          Else
                                             var_telefono = ""
                                          End If
                                          rsaux6.Close
                                       Else
                                          var_telefono = ""
                                       End If
                                    Else
                                       var_tel = 0
                                       VAR_DIRECCION = ""
                                       VAR_COLONIA = ""
                                       var_ciudad = ""
                                       VAR_MUNICIPIO = ""
                                       var_estado = ""
                                       var_pais = ""
                                       VAR_CP = ""
                                       var_telefono = ""
                                    End If
                                    
                                    If IsNumeric(txt_total_volumen) Then
                                       var_cubicaje_EMBARQUE = CDbl(Me.txt_total_volumen)
                                    Else
                                       var_cubicaje_EMBARQUE = 0
                                    End If
                                    rsaux4.Open "select distinct source_header_number from tb_Temp_oracle_orden_surtido where  inte_tem_consecutivo = " + CStr(var_consecutivo) + " and grupo = " + CStr(rsaux(0).Value), cnn, adOpenDynamic, adLockOptimistic
                                    While Not rsaux4.EOF
                                          rsaux2.Open "SELECT * FROM tb_oracle_pedidos_asignados_embarques WHERE PEDIDO = " + CStr(rsaux4(0).Value), cnn, adOpenDynamic, adLockOptimistic
                                          If Not rsaux2.EOF Then
                                             var_transporte = ""
                                             rsaux3.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                             If Not rsaux3.EOF Then
                                                var_transporte = IIf(IsNull(rsaux3!vehiculo), "", rsaux3!vehiculo)
                                             End If
                                             rsaux3.Close
                                             strconsulta = "SELECT ORDER_TYPE_ID FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ?"
                                             With comandoORA
                                                  .ActiveConnection = cnnoracle_4
                                                  .CommandType = adCmdText
                                                  .CommandText = strconsulta
                                                  Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rsaux4(0).Value))
                                                  .Parameters.Append parametro
                                             End With
                                             Set rsaux5 = comandoORA.execute
                                             Set comandoORA = Nothing
                                             Set parametro = Nothing
                                             If rsaux5!ORDER_TYPE_ID = 1002 Then
                                                rsaux5.Close
                                                strconsulta = "select ADDRESS_LINE_1||', '||ADDRESS_LINE_2||', '||TOWN_OR_CITY||', '||REGION_1||', '||COUNTRY||' CP:'||POSTAL_CODE DIRECCION, EMAIL from mtl_secondary_inventories a, hr_locations_all b, xxvia_jv_tb_agentes c, po_requisition_headers_ALL D, OE_ORDER_HEADERS_ALL E Where A.location_id = b.location_id and a.secondary_inventory_name = c.subinventory_code AND E.source_document_id = D.requisition_header_id AND A.secondary_inventory_name = D.ATTRIBUTE1 AND E.ORDER_NUMBER = ?"
                                             Else
                                                rsaux5.Close
                                                strconsulta = "SELECT CALLE||' '||NUM_CALLE||' '||NVL(num_interior,'')||', '||colonia||', '||ciudad||', '||estado||', '||pais||' CP: '||codigo_postal DIRECCION FROM oe_order_headers_all a, xxvia_vw_CLIENTES_BCP B WHERE A.SHIP_TO_ORG_ID = B.SITE_USE_ID AND A.ORDER_NUMBER     = ?"
                                             End If
                                             
                                             With comandoORA
                                                  .ActiveConnection = cnnoracle_4
                                                  .CommandType = adCmdText
                                                  .CommandText = strconsulta
                                                  Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rsaux4(0).Value))
                                                  .Parameters.Append parametro
                                             End With
                                             Set rsaux5 = comandoORA.execute
                                             Set comandoORA = Nothing
                                             Set parametro = Nothing
                                             If Not rsaux5.EOF Then
                                                VAR_DIRECCION = IIf(IsNull(rsaux5!DIRECCION), "", rsaux5!DIRECCION)
                                             Else
                                                VAR_DIRECCION = ""
                                             End If
                                             var_cadena = "UPDATE tb_Temp_oracle_orden_surtido SET EMBARQUE = " + CStr(IIf(IsNull(rsaux2!Embarque), 0, rsaux2!Embarque)) + ", CUBICAJE = " + CStr(var_cubicaje_EMBARQUE) + " , ORDEN_SURTIDO = " + CStr(IIf(IsNull(rsaux2!orden_pedido), 0, rsaux2!orden_pedido)) + ", ANDEN = '" + CStr(IIf(IsNull(rsaux2!estacion), 0, rsaux2!estacion)) + "', TRANSPORTE = '" + var_transporte + "',"
                                             var_cadena = var_cadena + " pais= '" + var_pais + "', estado = '" + var_estado + "', municipio = '" + VAR_MUNICIPIO + "', ciudad = '" + var_ciudad + "', colonia = '" + VAR_COLONIA + "', direccion = '" + VAR_DIRECCION + "', cp = '" + VAR_CP + "', paqueteria = '" + var_paqueteria + "'"
                                             var_cadena = var_cadena + " WHERE SOURCE_HEADER_NUMBER = " + CStr(rsaux4(0).Value) + " AND inte_tem_consecutivo = " + CStr(var_consecutivo)
                                             rsaux3.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                          End If
                                          rsaux2.Close
                                          rsaux4.MoveNext
                                    Wend
                                    rsaux4.Close
                                    
                                                                  
                                    x = 1
                                    If x = 1 Then
                                       If rsaux9.State = 1 Then
                                          rsaux9.Close
                                       End If
                                       rsaux9.Open "SELECT DISTINCT  GRUPO, SOURCE_HEADER_NUMBER from tb_Temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and grupo = " + CStr(rsaux(0).Value), cnn, adOpenDynamic, adLockOptimistic
                                       VAR_TOTAL_GRUPOS = rsaux9.RecordCount
                                       rsaux9.Close
                                       'VAR_TOTAL_GRUPOS = 1
                                       rsaux10.Open "update tb_Temp_oracle_orden_surtido set cubicaje = " + CStr(var_cubicaje_EMBARQUE) + "  where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                                       If VAR_TOTAL_GRUPOS = 1 Then
                                          Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos_060417.rpt")
                                          reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.grupo} = " + CStr(rsaux(0).Value) + " and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                                          'frmvistasprevias.cr.ReportSource = reporte
                                          For ntablas = 1 To reporte.Database.Tables.Count
                                              reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                          Next ntablas
                                          reporte.PrintOut False
                                          Set reporte = Nothing
                                          
                                       Else
                                          
                                          Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos_ENCABEZADOS.rpt")
                                          reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.grupo} = " + CStr(rsaux(0).Value) + " and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                                          frmvista_previa_auxiliar.cr2.ReportSource = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos_ENCABEZADOS.rpt")
                                          For ntablas = 1 To reporte.Database.Tables.Count
                                              reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                          Next ntablas
                                          reporte.PrintOut False
                                          frmvista_previa_auxiliar.Show 1
                                          
            'frmvistasprevias.cr.ViewReport
            'frmvistasprevias.Caption = "Packing List"
            'frmvistasprevias.Show 1
                                          
                                          Set reporte = Nothing
                                          
                                          rsaux1.Open "select distinct source_header_number from tb_Temp_oracle_orden_surtido where  inte_Tem_Consecutivo = " + CStr(var_consecutivo) + " and grupo = " + CStr(rsaux(0).Value), cnn, adOpenDynamic, adLockOptimistic
                                          var_i = 0
                                          While Not rsaux1.EOF
                                                If var_i = 0 Then
                                                   rsaux10.Open "update tb_Temp_oracle_orden_surtido set pasillo = 'plata' where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number = " + CStr(rsaux1!source_header_number) + " and grupo = " + CStr(rsaux(0).Value), cnn, adOpenDynamic, adLockOptimistic
                                                   var_i = 1
                                                End If
                                                rsaux1.MoveNext
                                          Wend
                                          rsaux1.Close
                                          
                                          
                                          Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos_DETALLE_060417.rpt")
                                          reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.grupo} = " + CStr(rsaux(0).Value) + " and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                                          frmvista_previa_auxiliar.cr2.ReportSource = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos_DETALLE.rpt")
                                          
                                          
                                          For ntablas = 1 To reporte.Database.Tables.Count
                                              reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                          Next ntablas
                                          reporte.PrintOut False
            'frmvistasprevias.cr.ViewReport
            'frmvistasprevias.Caption = "Packing List"
            'frmvistasprevias.Show 1
                                          
                                          Set reporte = Nothing
                                          frmvista_previa_auxiliar.Show 1
                                       
                                       End If
                                       
                                       
                                       'Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos.rpt")
                                       'reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.grupo} = " + CStr(rsaux(0).Value) + " and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                                       'For ntablas = 1 To reporte.Database.Tables.Count
                                       '    reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                       'Next ntablas
                                       'reporte.PrintOut False
                                       'Set reporte = Nothing
                                    Else
                                       If rsaux9.State = 1 Then
                                          rsaux9.Close
                                       End If
                                       rsaux9.Open "SELECT DISTINCT  GRUPO, SOURCE_HEADER_NUMBER from tb_Temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and grupo = " + CStr(rsaux(0).Value), cnn, adOpenDynamic, adLockOptimistic
                                       VAR_TOTAL_GRUPOS = rsaux9.RecordCount
                                       rsaux9.Close
                                       'VAR_TOTAL_GRUPOS = 1
                                       If VAR_TOTAL_GRUPOS = 1 Then
                                          Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos_060417.rpt")
                                          reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.grupo} = " + CStr(rsaux(0).Value) + " and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                                          frmvistasprevias.cr.ReportSource = reporte
                                          For ntablas = 1 To reporte.Database.Tables.Count
                                              reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                          Next ntablas
                                          frmvistasprevias.cr.ViewReport
                                          frmvistasprevias.Caption = "Ordenes de surtido "
                                          frmvistasprevias.Show 1
                                          Set reporte = Nothing
                                       Else
                                       
                                          rsaux1.Open "select distinct source_header_number from tb_Temp_oracle_orden_surtido where  inte_Tem_Consecutivo = " + CStr(var_consecutivo) + " and grupo = " + CStr(rsaux(0).Value), cnn, adOpenDynamic, adLockOptimistic
                                          var_i = 0
                                          While Not rsaux1.EOF
                                                If var_i = 0 Then
                                                   rsaux10.Open "update tb_Temp_oracle_orden_surtido set pasillo = 'plata' where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " and source_header_number = " + CStr(rsaux1!source_header_number) + " and grupo = " + CStr(rsaux(0).Value), cnn, adOpenDynamic, adLockOptimistic
                                                   var_i = 1
                                                End If
                                                rsaux1.MoveNext
                                          Wend
                                          rsaux1.Close
                                          
                                          Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos_ENCABEZADOS.rpt")
                                          reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.grupo} = " + CStr(rsaux(0).Value) + " and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                                          frmvistasprevias.cr.ReportSource = reporte
                                          For ntablas = 1 To reporte.Database.Tables.Count
                                              reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                          Next ntablas
                                          frmvistasprevias.cr.ViewReport
                                          frmvistasprevias.Caption = "Ordenes de surtido "
                                          frmvistasprevias.Show 1
                                          Set reporte = Nothing
                                          
                                          Set reporte = appl.OpenReport(App.Path + "\rep_oracle_orden_surtido_DIVIDIDA_grupos_DETALLE_060417.rpt")
                                          reporte.RecordSelectionFormula = "{VW_ORACLE_ORDEN_SURTIDO.grupo} = " + CStr(rsaux(0).Value) + " and {VW_ORACLE_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                                          frmvistasprevias.cr.ReportSource = reporte
                                          For ntablas = 1 To reporte.Database.Tables.Count
                                              reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                                          Next ntablas
                                          frmvistasprevias.cr.ViewReport
                                          frmvistasprevias.Caption = "Ordenes de surtido "
                                          frmvistasprevias.Show 1
                                          Set reporte = Nothing
                                       
                                       
                                       End If
                                       
                                    End If
                                 End If
                                 rsaux.MoveNext
                           Wend
                        End If
                        rsaux.Close
                     Else
                        MsgBox "No se pueden imprimir los pedidos, vuelva a intentar la impresión", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "No se pueden imprimir los pedidos, vuelva a intentar la impresión", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "No se pueden imprimir los pedidos, vuelva a intentar la impresión", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "No existen ordenes de surtido", vbOKOnly, "ATENCION"
            End If
            If rs.State = 1 Then
               rs.Close
            End If
            If var_consecutivo > 0 Then
               rs.Open "delete from tb_Temp_oracle_orden_surtido where inte_tem_consecutivo = " + CStr(IIf(IsNull(var_consecutivo), 0, var_consecutivo)), cnn, adOpenDynamic, adLockOptimistic
            End If
         Else
            'MsgBox "Número superior incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "No existen pedidos a imprimir"
      End If

End Sub

Private Sub Command1_Click()
   rs.Open "select source_header_number, sum(src_requested_quantity * to_number(nvl(a.unit_volume,'0'))) as volumen from xxvia_system_items_b a, xxvia_tb_pedidos_divididos b where a.attribute11 is not null and a.inventory_item_id = b.inventory_item_id and source_header_number > 220000 group by source_header_number", cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_i = 0
   While Not rs.EOF
         rsaux.Open "update tb_oracle_pedidos_asignados_embarques set volumen = " + CStr(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN)) + " where pedido = " + CStr(rs!source_header_number), cnn, adOpenDynamic, adLockOptimistic
         var_i = var_i + 1
         If var_i = 100 Then
            var_i = 0
         End If
         rs.MoveNext
   Wend
   rs.Close
   MsgBox "termino"
End Sub

Private Sub Command2_Click()
   
End Sub

Private Sub Command6_Click()
   var_i = 0
   For var_j = 1 To 10
       VAR_Z = CInt(Int((500 * Rnd()) + 1))
       rs.Open "select TO_CHAR(dbms_random.value(1,100), '999') as numb from dual", cnnoracle_4, adOpenDynamic, adLockOptimistic
       If Not rs.EOF Then
          If rs(0).Value = VAR_Z Then
             var_i = var_i + 1
          End If
       End If
       rs.Close
   Next var_j
   MsgBox var_i
End Sub

Private Sub Form_Load()
   Me.frm_orden.Visible = False
   Me.txt_fecha = Date
   Me.lv_embarques_1.ListItems.Clear
   var_dia_s = CStr(Day(Date))
   If Len(var_dia_s) = 1 Then
      var_dia_s = "0" + var_dia_s
   End If
   var_mes_s = CStr(Month(Date))
   If Len(var_mes_s) = 1 Then
      var_mes_s = "0" + var_mes_s
   End If
   var_año = CStr(Year(Date))
   
   rs.Open "select distinct embarque as embarque from xxvia_Tb_encabezado_embarques where jaula = 1 and to_char(fecha_inicio,'DD')  = '" + CStr(var_dia_s) + "' and to_char(fecha_inicio,'MM')  = '" + CStr(var_mes_s) + "' and (to_char(fecha_inicio,'yyyy')  = '" + CStr(Year(Date)) + "' or to_char(fecha_inicio,'YY') = '" + Mid(var_año, 1, 2) + "') and organizacion = " + CStr(var_unidad_organizacional), cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_Cadena_embarques = ""
   While Not rs.EOF
         If var_Cadena_embarques = "" Then
            var_Cadena_embarques = CStr(rs!Embarque)
         Else
            var_Cadena_embarques = var_Cadena_embarques + "," + CStr(rs!Embarque)
         End If
         rs.MoveNext
   Wend
   rs.Close
   If var_Cadena_embarques <> "" Then
      If rs.State = 1 Then
         rs.Close
      End If
      'MsgBox cnn.ConnectionString
      rs.Open "select distinct EMBARQUE from tb_oracle_pedidos_asignados_embarques where embarque in (" + var_Cadena_embarques + ") union all select distinct EMBARQUE from tb_oracle_pedidos_asignados_embarques_vad where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_embarques_1.ListItems.Add(, , rs!Embarque)
            strconsulta = "select nvl(char_emb_estatus,' ') as estatus from xxvia_Tb_encabezado_embarques where embarque = ?"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rs!Embarque)
                 .Parameters.Append parametro
            End With
            Set rsaux8 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            If rsaux8!estatus = " " Or rsaux8!estatus = "E" Then
               list_item.Bold = False
               list_item.ForeColor = &H80000012
            Else
               list_item.Bold = True
               list_item.ForeColor = &H8000&
            End If
            'list_item.SubItems(1) = rs!NOMBRE_AGENTE
            rs.MoveNext
      Wend
      rs.Close
      If lv_embarques_1.ListItems.Count > 0 Then
         Me.txt_embarque = Me.lv_embarques_1.selectedItem
         lv_embarques_1.ListItems(1).Selected = True
         x = 0
         If x = 1 Then
         
         rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria, 'Oracle' origen  " & _
                "from tb_oracle_pedidos_asignados_embarques_vad " & _
                "where embarque = " + CStr(Me.lv_embarques_1.selectedItem) & _
                " select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria 'VAD' origen" & _
                "from tb_oracle_pedidos_asignados_embarques " & _
                "where embarque = " + CStr(Me.lv_embarques_1.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         lv_pedidos.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
               strconsulta = "select * from oe_order_headers_all where order_number = ? union all "
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                    .Parameters.Append parametro
               End With
               Set rsaux6 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
               list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
               'rsaux.Close
               list_item.SubItems(2) = rs!Cliente
               list_item.SubItems(3) = rs!PIEZAS
               list_item.SubItems(5) = rs!orden_pedido
               list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
               list_item.SubItems(7) = rsaux6!INVOICE_TO_ORG_ID
               list_item.SubItems(9) = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
               If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
                  strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!pedido)
                       .Parameters.Append parametro
                  End With
                  Set rsaux8 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
                  rsaux8.Close
                  strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                       .Parameters.Append parametro
                  End With
                  Set rsaux7 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  list_item.SubItems(8) = rsaux7!attribute1
                  rsaux7.Close
               Else
                  list_item.SubItems(8) = rsaux6!SHIP_TO_ORG_ID
               End If
               rsaux6.Close
               rs.MoveNext
         Wend
         rs.Close
         End If
         rs.Open "select sum(pzas) piezas, sum(VOLUMEN) volumen from (select PIEZAS pzas, isnull(VOLUMEN,0) VOLUMEN  from TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque in (" + var_Cadena_embarques + ") union all select PIEZAS pzas, isnull(VOLUMEN,0) VOLUMEN from tb_oracle_pedidos_asignados_embarques_vad where embarque in (" + var_Cadena_embarques + ") ) ped", cnn, adOpenDynamic, adLockOptimistic
         Me.lbl_cantidad_1 = Format(rs(0).Value, "###,###,##0.00")
         'rs.Close
         
         'rs.Open "select sum(volumen) as volumen from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_1.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_total_volumen = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
         Else
            Me.txt_total_volumen = 0
         End If
         rs.Close
         
         strconsulta = "select * from xxvia_tb_encabezado_embarques where embarque = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.lv_embarques_1.selectedItem))
              .Parameters.Append parametro
         End With
         Set rsaux6 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         var_volumen = 0
         If Not rsaux6.EOF Then
            var_transporte = IIf(IsNull(rsaux6!transporte), "", rsaux6!transporte)
         Else
            var_transporte = ""
         End If
         rsaux6.Close
         
         rs.Open "select * from tb_oracle_transportes where clave = '" + var_transporte + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_volumen_unidad = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
            Me.txt_transporte = IIf(IsNull(rs!NOMBRE), "", rs!NOMBRE)
         Else
            Me.txt_transporte = ""
            Me.txt_volumen_unidad = "0.00"
         End If
         rs.Close
         
         
         
      End If
   End If
   
   rs.Open "select distinct embarque as embarque from xxvia_Tb_encabezado_embarques where jaula = 2 and to_char(fecha_inicio,'DD')  = '" + CStr(var_dia_s) + "' and to_char(fecha_inicio,'MM')  = '" + CStr(var_mes_s) + "' and to_char(fecha_inicio,'yyyy')  = '" + CStr(Year(Date)) + "' and organizacion = " + CStr(var_unidad_organizacional), cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_Cadena_embarques = ""
   While Not rs.EOF
         If var_Cadena_embarques = "" Then
            var_Cadena_embarques = CStr(rs!Embarque)
         Else
            var_Cadena_embarques = var_Cadena_embarques + "," + CStr(rs!Embarque)
         End If
         rs.MoveNext
   Wend
   rs.Close
   If var_Cadena_embarques <> "" Then
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open "select distinct EMBARQUE from tb_oracle_pedidos_asignados_embarques where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_embarques_2.ListItems.Add(, , rs!Embarque)
            strconsulta = "select nvl(char_emb_estatus,' ') as estatus from xxvia_Tb_encabezado_embarques where embarque = ?"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rs!Embarque)
                 .Parameters.Append parametro
            End With
            Set rsaux8 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            If rsaux8!estatus = " " Or rsaux8!estatus = "E" Then
               list_item.Bold = False
               list_item.ForeColor = &H80000012
            Else
               list_item.Bold = True
               list_item.ForeColor = &H8000&
            End If
            'list_item.SubItems(1) = rs!NOMBRE_AGENTE
            rs.MoveNext
      Wend
      rs.Close
      If lv_embarques_2.ListItems.Count > 0 Then
         lv_embarques_2.ListItems(1).Selected = True
         Me.txt_embarque = Me.lv_embarques_2.selectedItem
         x = 0
         If x = 1 Then
         
         rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria   from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_2.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         lv_pedidos.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
               strconsulta = "select * from oe_order_headers_all where order_number = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                    .Parameters.Append parametro
               End With
               Set rsaux6 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
               list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
               'rsaux.Close
               list_item.SubItems(2) = rs!Cliente
               list_item.SubItems(3) = rs!PIEZAS
               list_item.SubItems(5) = rs!orden_pedido
               list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
               list_item.SubItems(7) = rsaux6!INVOICE_TO_ORG_ID
               list_item.SubItems(9) = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
               
               If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
                  strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!pedido)
                       .Parameters.Append parametro
                  End With
                  Set rsaux8 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
                  rsaux8.Close
                  strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                       .Parameters.Append parametro
                  End With
                  Set rsaux7 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  list_item.SubItems(8) = rsaux7!attribute1
                  rsaux7.Close
               Else
                  list_item.SubItems(8) = rsaux6!SHIP_TO_ORG_ID
               End If
               rsaux6.Close
               rs.MoveNext
         Wend
         rs.Close
         End If
         'If lv_pedidos.ListItems.Count > 11 Then
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
         'Else
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
         'End If
         rs.Open "SELECT SUM(PIEZAS) FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
         Me.lbl_cantidad_2 = Format(rs(0).Value, "###,###,##0.00")
         rs.Close
      End If
   End If
   
   
   rs.Open "select distinct embarque as embarque from xxvia_Tb_encabezado_embarques where jaula = 3 and to_char(fecha_inicio,'DD')  = '" + CStr(var_dia_s) + "' and to_char(fecha_inicio,'MM')  = '" + CStr(var_mes_s) + "' and to_char(fecha_inicio,'yyyy')  = '" + CStr(Year(Date)) + "' and organizacion = " + CStr(var_unidad_organizacional), cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_Cadena_embarques = ""
   While Not rs.EOF
         If var_Cadena_embarques = "" Then
            var_Cadena_embarques = CStr(rs!Embarque)
         Else
            var_Cadena_embarques = var_Cadena_embarques + "," + CStr(rs!Embarque)
         End If
         rs.MoveNext
   Wend
   rs.Close
   If var_Cadena_embarques <> "" Then
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open "select distinct EMBARQUE from tb_oracle_pedidos_asignados_embarques where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_embarques_3.ListItems.Add(, , rs!Embarque)
            strconsulta = "select nvl(char_emb_estatus,' ') as estatus from xxvia_Tb_encabezado_embarques where embarque = ?"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rs!Embarque)
                 .Parameters.Append parametro
            End With
            Set rsaux8 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            If rsaux8!estatus = " " Or rsaux8!estatus = "E" Then
               list_item.Bold = False
               list_item.ForeColor = &H80000012
            Else
               list_item.Bold = True
               list_item.ForeColor = &H8000&
            End If
            'list_item.SubItems(1) = rs!NOMBRE_AGENTE
            rs.MoveNext
      Wend
      rs.Close
      If lv_embarques_3.ListItems.Count > 0 Then
         lv_embarques_3.ListItems(1).Selected = True
         Me.txt_embarque = Me.lv_embarques_3.selectedItem
                  x = 0
         If x = 1 Then

         rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria   from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_3.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         lv_pedidos.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
               strconsulta = "select * from oe_order_headers_all where order_number = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                    .Parameters.Append parametro
               End With
               Set rsaux6 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
               list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
               'rsaux.Close
               list_item.SubItems(2) = rs!Cliente
               list_item.SubItems(3) = rs!PIEZAS
               list_item.SubItems(5) = rs!orden_pedido
               list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
               list_item.SubItems(7) = rsaux6!INVOICE_TO_ORG_ID
               list_item.SubItems(9) = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
               If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
                  strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!pedido)
                       .Parameters.Append parametro
                  End With
                  Set rsaux8 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
                  rsaux8.Close
                  strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                       .Parameters.Append parametro
                  End With
                  Set rsaux7 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  list_item.SubItems(8) = rsaux7!attribute1
                  rsaux7.Close
               Else
                  list_item.SubItems(8) = rsaux6!SHIP_TO_ORG_ID
               End If
               rsaux6.Close
               rs.MoveNext
         Wend
         rs.Close
End If
         'If lv_pedidos.ListItems.Count > 11 Then
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
         'Else
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
         'End If
         rs.Open "SELECT SUM(PIEZAS) FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
         Me.lbl_cantidad_3 = Format(rs(0).Value, "###,###,##0.00")
         rs.Close
      End If
   End If
   
   rs.Open "select distinct embarque as embarque from xxvia_Tb_encabezado_embarques where jaula = 4 and to_char(fecha_inicio,'DD')  = '" + CStr(var_dia_s) + "' and to_char(fecha_inicio,'MM')  = '" + CStr(var_mes_s) + "' and to_char(fecha_inicio,'yyyy')  = '" + CStr(Year(Date)) + "' and organizacion = " + CStr(var_unidad_organizacional), cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_Cadena_embarques = ""
   While Not rs.EOF
         If var_Cadena_embarques = "" Then
            var_Cadena_embarques = CStr(rs!Embarque)
         Else
            var_Cadena_embarques = var_Cadena_embarques + "," + CStr(rs!Embarque)
         End If
         rs.MoveNext
   Wend
   rs.Close
   If var_Cadena_embarques <> "" Then
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open "select distinct EMBARQUE from tb_oracle_pedidos_asignados_embarques where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_embarques_4.ListItems.Add(, , rs!Embarque)
            strconsulta = "select nvl(char_emb_estatus,' ') as estatus from xxvia_Tb_encabezado_embarques where embarque = ?"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rs!Embarque)
                 .Parameters.Append parametro
            End With
            Set rsaux8 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            If rsaux8!estatus = " " Or rsaux8!estatus = "E" Then
               list_item.Bold = False
               list_item.ForeColor = &H80000012
            Else
               list_item.Bold = True
               list_item.ForeColor = &H8000&
            End If
            'list_item.SubItems(1) = rs!NOMBRE_AGENTE
            rs.MoveNext
      Wend
      rs.Close
      If lv_embarques_4.ListItems.Count > 0 Then
         lv_embarques_4.ListItems(1).Selected = True
         Me.txt_embarque = Me.lv_embarques_4.selectedItem
         x = 0
         If x = 1 Then
         rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria   from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_4.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         lv_pedidos.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
               strconsulta = "select * from oe_order_headers_all where order_number = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                    .Parameters.Append parametro
               End With
               Set rsaux6 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
               list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
               'rsaux.Close
               list_item.SubItems(2) = rs!Cliente
               list_item.SubItems(3) = rs!PIEZAS
               list_item.SubItems(5) = rs!orden_pedido
               list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
               list_item.SubItems(7) = rsaux6!INVOICE_TO_ORG_ID
               list_item.SubItems(9) = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
               If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
                  strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!pedido)
                       .Parameters.Append parametro
                  End With
                  Set rsaux8 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
                  rsaux8.Close
                  strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                       .Parameters.Append parametro
                  End With
                  Set rsaux7 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  list_item.SubItems(8) = rsaux7!attribute1
                  rsaux7.Close
               Else
                  list_item.SubItems(8) = rsaux6!SHIP_TO_ORG_ID
               End If
               rsaux6.Close
               rs.MoveNext
         Wend
         rs.Close
End If
         'If lv_pedidos.ListItems.Count > 11 Then
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
         'Else
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
         'End If
         rs.Open "SELECT SUM(PIEZAS) FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
         Me.lbl_cantidad_4 = Format(rs(0).Value, "###,###,##0.00")
         rs.Close
      End If
   End If
   
   
   rs.Open "select distinct embarque as embarque from xxvia_Tb_encabezado_embarques where jaula = 5 and to_char(fecha_inicio,'DD')  = '" + CStr(var_dia_s) + "' and to_char(fecha_inicio,'MM')  = '" + CStr(var_mes_s) + "' and to_char(fecha_inicio,'yyyy')  = '" + CStr(Year(Date)) + "' and organizacion = " + CStr(var_unidad_organizacional), cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_Cadena_embarques = ""
   While Not rs.EOF
         If var_Cadena_embarques = "" Then
            var_Cadena_embarques = CStr(rs!Embarque)
         Else
            var_Cadena_embarques = var_Cadena_embarques + "," + CStr(rs!Embarque)
         End If
         rs.MoveNext
   Wend
   rs.Close
   If var_Cadena_embarques <> "" Then
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open "select distinct EMBARQUE from tb_oracle_pedidos_asignados_embarques where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_embarques_5.ListItems.Add(, , rs!Embarque)
            strconsulta = "select nvl(char_emb_estatus,' ') as estatus from xxvia_Tb_encabezado_embarques where embarque = ?"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rs!Embarque)
                 .Parameters.Append parametro
            End With
            Set rsaux8 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            If rsaux8!estatus = " " Or rsaux8!estatus = "E" Then
               list_item.Bold = False
               list_item.ForeColor = &H80000012
            Else
               list_item.Bold = True
               list_item.ForeColor = &H8000&
            End If
            'list_item.SubItems(1) = rs!NOMBRE_AGENTE
            rs.MoveNext
      Wend
      rs.Close
      If lv_embarques_5.ListItems.Count > 0 Then
         lv_embarques_5.ListItems(1).Selected = True
         Me.txt_embarque = Me.lv_embarques_5.selectedItem
                  x = 0
         If x = 1 Then
         rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria   from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_5.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         lv_pedidos.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
               strconsulta = "select * from oe_order_headers_all where order_number = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                    .Parameters.Append parametro
               End With
               Set rsaux6 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
               list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
               'rsaux.Close
               list_item.SubItems(2) = rs!Cliente
               list_item.SubItems(3) = rs!PIEZAS
               list_item.SubItems(5) = rs!orden_pedido
               list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
               list_item.SubItems(7) = rsaux6!INVOICE_TO_ORG_ID
               list_item.SubItems(9) = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
               If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
                  strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!pedido)
                       .Parameters.Append parametro
                  End With
                  Set rsaux8 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
                  rsaux8.Close
                  strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                       .Parameters.Append parametro
                  End With
                  Set rsaux7 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  list_item.SubItems(8) = rsaux7!attribute1
                  rsaux7.Close
               Else
                  list_item.SubItems(8) = rsaux6!SHIP_TO_ORG_ID
               End If
               rsaux6.Close
               rs.MoveNext
         Wend
         rs.Close
         End If
         'If lv_pedidos.ListItems.Count > 11 Then
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
         'Else
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
         'End If
         rs.Open "SELECT SUM(PIEZAS) FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
         
         Me.lbl_cantidad_5 = Format(rs(0).Value, "###,###,##0.00")
         rs.Close
      End If
   End If
   
   rs.Open "select distinct embarque as embarque from xxvia_Tb_encabezado_embarques where jaula = 6 and to_char(fecha_inicio,'DD')  = '" + CStr(var_dia_s) + "' and to_char(fecha_inicio,'MM')  = '" + CStr(var_mes_s) + "' and to_char(fecha_inicio,'yyyy')  = '" + CStr(Year(Date)) + "' and organizacion = " + CStr(var_unidad_organizacional), cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_Cadena_embarques = ""
   While Not rs.EOF
         If var_Cadena_embarques = "" Then
            var_Cadena_embarques = CStr(rs!Embarque)
         Else
            var_Cadena_embarques = var_Cadena_embarques + "," + CStr(rs!Embarque)
         End If
         rs.MoveNext
   Wend
   rs.Close
   If var_Cadena_embarques <> "" Then
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open "select distinct EMBARQUE from tb_oracle_pedidos_asignados_embarques where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_embarques_6.ListItems.Add(, , rs!Embarque)
            strconsulta = "select nvl(char_emb_estatus,' ') as estatus from xxvia_Tb_encabezado_embarques where embarque = ?"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rs!Embarque)
                 .Parameters.Append parametro
            End With
            Set rsaux8 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            If rsaux8!estatus = " " Or rsaux8!estatus = "E" Then
               list_item.Bold = False
               list_item.ForeColor = &H80000012
            Else
               list_item.Bold = True
               list_item.ForeColor = &H8000&
            End If
            'list_item.SubItems(1) = rs!NOMBRE_AGENTE
            rs.MoveNext
      Wend
      rs.Close
      If lv_embarques_6.ListItems.Count > 0 Then
         lv_embarques_6.ListItems(1).Selected = True
         Me.txt_embarque = Me.lv_embarques_6.selectedItem
         x = 0
         If x = 1 Then
         rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria   from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_6.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         lv_pedidos.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
               strconsulta = "select * from oe_order_headers_all where order_number = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                    .Parameters.Append parametro
               End With
               Set rsaux6 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
               list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
               'rsaux.Close
               list_item.SubItems(2) = rs!Cliente
               list_item.SubItems(3) = rs!PIEZAS
               list_item.SubItems(5) = rs!orden_pedido
               list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
               list_item.SubItems(7) = rsaux6!INVOICE_TO_ORG_ID
               list_item.SubItems(9) = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
               If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
                  strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!pedido)
                       .Parameters.Append parametro
                  End With
                  Set rsaux8 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
                  rsaux8.Close
                  strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                       .Parameters.Append parametro
                  End With
                  Set rsaux7 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  list_item.SubItems(8) = rsaux7!attribute1
                  rsaux7.Close
               Else
                  list_item.SubItems(8) = rsaux6!SHIP_TO_ORG_ID
               End If
               rsaux6.Close
               rs.MoveNext
         Wend
         rs.Close
End If
         'If lv_pedidos.ListItems.Count > 11 Then
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
         'Else
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
         'End If
         rs.Open "SELECT SUM(PIEZAS) FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
         Me.lbl_cantidad_6 = Format(rs(0).Value, "###,###,##0.00")
         rs.Close
      End If
   End If
   
   rs.Open "select distinct embarque as embarque from xxvia_Tb_encabezado_embarques where jaula = 7 and to_char(fecha_inicio,'DD')  = '" + CStr(var_dia_s) + "' and to_char(fecha_inicio,'MM')  = '" + CStr(var_mes_s) + "' and to_char(fecha_inicio,'yyyy')  = '" + CStr(Year(Date)) + "' and organizacion = " + CStr(var_unidad_organizacional), cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_Cadena_embarques = ""
   While Not rs.EOF
         If var_Cadena_embarques = "" Then
            var_Cadena_embarques = CStr(rs!Embarque)
         Else
            var_Cadena_embarques = var_Cadena_embarques + "," + CStr(rs!Embarque)
         End If
         rs.MoveNext
   Wend
   rs.Close
   If var_Cadena_embarques <> "" Then
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open "select distinct EMBARQUE from tb_oracle_pedidos_asignados_embarques where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_embarques_7.ListItems.Add(, , rs!Embarque)
            strconsulta = "select nvl(char_emb_estatus,' ') as estatus from xxvia_Tb_encabezado_embarques where embarque = ?"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rs!Embarque)
                 .Parameters.Append parametro
            End With
            Set rsaux8 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            If rsaux8!estatus = " " Or rsaux8!estatus = "E" Then
               list_item.Bold = False
               list_item.ForeColor = &H80000012
            Else
               list_item.Bold = True
               list_item.ForeColor = &H8000&
            End If
            'list_item.SubItems(1) = rs!NOMBRE_AGENTE
            rs.MoveNext
      Wend
      rs.Close
      If lv_embarques_7.ListItems.Count > 0 Then
         lv_embarques_7.ListItems(1).Selected = True
         Me.txt_embarque = Me.lv_embarques_7.selectedItem
         x = 0
         If x = 1 Then
         rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria   from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_7.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         lv_pedidos.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
               strconsulta = "select * from oe_order_headers_all where order_number = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                    .Parameters.Append parametro
               End With
               Set rsaux6 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
               list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
               'rsaux.Close
               list_item.SubItems(2) = rs!Cliente
               list_item.SubItems(3) = rs!PIEZAS
               list_item.SubItems(5) = rs!orden_pedido
               list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
               list_item.SubItems(7) = rsaux6!INVOICE_TO_ORG_ID
               list_item.SubItems(9) = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
               If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
                  strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!pedido)
                       .Parameters.Append parametro
                  End With
                  Set rsaux8 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
                  rsaux8.Close
                  strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                       .Parameters.Append parametro
                  End With
                  Set rsaux7 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  list_item.SubItems(8) = rsaux7!attribute1
                  rsaux7.Close
               Else
                  list_item.SubItems(8) = rsaux6!SHIP_TO_ORG_ID
               End If
               rsaux6.Close
               rs.MoveNext
         Wend
         rs.Close
End If
         'If lv_pedidos.ListItems.Count > 11 Then
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
         'Else
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
         'End If
         rs.Open "SELECT SUM(PIEZAS) FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
         Me.lbl_cantidad_7 = Format(rs(0).Value, "###,###,##0.00")
         rs.Close
      End If
   End If
   
   
   
   rs.Open "select distinct embarque as embarque from xxvia_Tb_encabezado_embarques where jaula = 8 and to_char(fecha_inicio,'DD')  = '" + CStr(var_dia_s) + "' and to_char(fecha_inicio,'MM')  = '" + CStr(var_mes_s) + "' and to_char(fecha_inicio,'yyyy')  = '" + CStr(Year(Date)) + "' and organizacion = " + CStr(var_unidad_organizacional), cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_Cadena_embarques = ""
   While Not rs.EOF
         If var_Cadena_embarques = "" Then
            var_Cadena_embarques = CStr(rs!Embarque)
         Else
            var_Cadena_embarques = var_Cadena_embarques + "," + CStr(rs!Embarque)
         End If
         rs.MoveNext
   Wend
   rs.Close
   If var_Cadena_embarques <> "" Then
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open "select distinct EMBARQUE from tb_oracle_pedidos_asignados_embarques where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_embarques_8.ListItems.Add(, , rs!Embarque)
            strconsulta = "select nvl(char_emb_estatus,' ') as estatus from xxvia_Tb_encabezado_embarques where embarque = ?"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rs!Embarque)
                 .Parameters.Append parametro
            End With
            Set rsaux8 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            If rsaux8!estatus = " " Or rsaux8!estatus = "E" Then
               list_item.Bold = False
               list_item.ForeColor = &H80000012
            Else
               list_item.Bold = True
               list_item.ForeColor = &H8000&
            End If
            'list_item.SubItems(1) = rs!NOMBRE_AGENTE
            rs.MoveNext
      Wend
      rs.Close
      If lv_embarques_8.ListItems.Count > 0 Then
         lv_embarques_8.ListItems(1).Selected = True
         Me.txt_embarque = Me.lv_embarques_8.selectedItem
         x = 0
         If x = 1 Then
         
         rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria   from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_8.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         lv_pedidos.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
               strconsulta = "select * from oe_order_headers_all where order_number = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                    .Parameters.Append parametro
               End With
               Set rsaux6 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
               list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
               'rsaux.Close
               list_item.SubItems(2) = rs!Cliente
               list_item.SubItems(3) = rs!PIEZAS
               list_item.SubItems(5) = rs!orden_pedido
               list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
               list_item.SubItems(7) = rsaux6!INVOICE_TO_ORG_ID
               list_item.SubItems(9) = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
               If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
                  strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!pedido)
                       .Parameters.Append parametro
                  End With
                  Set rsaux8 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
                  rsaux8.Close
                  strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                       .Parameters.Append parametro
                  End With
                  Set rsaux7 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  list_item.SubItems(8) = rsaux7!attribute1
                  rsaux7.Close
               Else
                  list_item.SubItems(8) = rsaux6!SHIP_TO_ORG_ID
               End If
               rsaux6.Close
               rs.MoveNext
         Wend
         rs.Close
End If
         'If lv_pedidos.ListItems.Count > 11 Then
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
         'Else
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
         'End If
         rs.Open "SELECT SUM(PIEZAS) FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
         Me.lbl_cantidad_8 = Format(rs(0).Value, "###,###,##0.00")
         rs.Close
      End If
   End If
   
   rs.Open "select distinct embarque as embarque from xxvia_Tb_encabezado_embarques where jaula = 9 and to_char(fecha_inicio,'DD')  = '" + CStr(var_dia_s) + "' and to_char(fecha_inicio,'MM')  = '" + CStr(var_mes_s) + "' and to_char(fecha_inicio,'yyyy')  = '" + CStr(Year(Date)) + "' and organizacion = " + CStr(var_unidad_organizacional), cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_Cadena_embarques = ""
   While Not rs.EOF
         If var_Cadena_embarques = "" Then
            var_Cadena_embarques = CStr(rs!Embarque)
         Else
            var_Cadena_embarques = var_Cadena_embarques + "," + CStr(rs!Embarque)
         End If
         rs.MoveNext
   Wend
   rs.Close
   If var_Cadena_embarques <> "" Then
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open "select distinct EMBARQUE from tb_oracle_pedidos_asignados_embarques where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_embarques_9.ListItems.Add(, , rs!Embarque)
            strconsulta = "select nvl(char_emb_estatus,' ') as estatus from xxvia_Tb_encabezado_embarques where embarque = ?"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rs!Embarque)
                 .Parameters.Append parametro
            End With
            Set rsaux8 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            If rsaux8!estatus = " " Or rsaux8!estatus = "E" Then
               list_item.Bold = False
               list_item.ForeColor = &H80000012
            Else
               list_item.Bold = True
               list_item.ForeColor = &H8000&
            End If
            'list_item.SubItems(1) = rs!NOMBRE_AGENTE
            rs.MoveNext
      Wend
      rs.Close
      If lv_embarques_9.ListItems.Count > 0 Then
         lv_embarques_9.ListItems(1).Selected = True
         Me.txt_embarque = Me.lv_embarques_9.selectedItem
         x = 0
         If x = 1 Then
         rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria   from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_9.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         lv_pedidos.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
               strconsulta = "select * from oe_order_headers_all where order_number = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                    .Parameters.Append parametro
               End With
               Set rsaux6 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
               list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
               'rsaux.Close
               list_item.SubItems(2) = rs!Cliente
               list_item.SubItems(3) = rs!PIEZAS
               list_item.SubItems(5) = rs!orden_pedido
               list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
               list_item.SubItems(7) = rsaux6!INVOICE_TO_ORG_ID
               list_item.SubItems(9) = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
               If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
                  strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!pedido)
                       .Parameters.Append parametro
                  End With
                  Set rsaux8 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
                  rsaux8.Close
                  strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                       .Parameters.Append parametro
                  End With
                  Set rsaux7 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  list_item.SubItems(8) = rsaux7!attribute1
                  rsaux7.Close
               Else
                  list_item.SubItems(8) = rsaux6!SHIP_TO_ORG_ID
               End If
               rsaux6.Close
               rs.MoveNext
         Wend
         rs.Close
End If
         'If lv_pedidos.ListItems.Count > 11 Then
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
         'Else
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
         'End If
         rs.Open "SELECT SUM(PIEZAS) FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
         Me.lbl_cantidad_9 = Format(rs(0).Value, "###,###,##0.00")
         rs.Close
      End If
   End If
   
   
   rs.Open "select distinct embarque as embarque from xxvia_Tb_encabezado_embarques where jaula = 10 and to_char(fecha_inicio,'DD')  = '" + CStr(var_dia_s) + "' and to_char(fecha_inicio,'MM')  = '" + CStr(var_mes_s) + "' and to_char(fecha_inicio,'yyyy')  = '" + CStr(Year(Date)) + "' and organizacion = " + CStr(var_unidad_organizacional), cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_Cadena_embarques = ""
   While Not rs.EOF
         If var_Cadena_embarques = "" Then
            var_Cadena_embarques = CStr(rs!Embarque)
         Else
            var_Cadena_embarques = var_Cadena_embarques + "," + CStr(rs!Embarque)
         End If
         rs.MoveNext
   Wend
   rs.Close
   If var_Cadena_embarques <> "" Then
      If rs.State = 1 Then
         rs.Close
      End If
      
      rs.Open "select distinct EMBARQUE from tb_oracle_pedidos_asignados_embarques where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_embarques_10.ListItems.Add(, , rs!Embarque)
            strconsulta = "select nvl(char_emb_estatus,' ') as estatus from xxvia_Tb_encabezado_embarques where embarque = ?"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rs!Embarque)
                 .Parameters.Append parametro
            End With
            Set rsaux8 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            If rsaux8!estatus = " " Or rsaux8!estatus = "E" Then
               list_item.Bold = False
               list_item.ForeColor = &H80000012
            Else
               list_item.Bold = True
               list_item.ForeColor = &H8000&
            End If
            'list_item.SubItems(1) = rs!NOMBRE_AGENTE
            rs.MoveNext
      Wend
      rs.Close
      If lv_embarques_10.ListItems.Count > 0 Then
         lv_embarques_10.ListItems(1).Selected = True
         Me.txt_embarque = Me.lv_embarques_10.selectedItem
         x = 0
         If x = 1 Then
         
         rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria   from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_10.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         lv_pedidos.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
               strconsulta = "select * from oe_order_headers_all where order_number = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                    .Parameters.Append parametro
               End With
               Set rsaux6 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
               list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
               'rsaux.Close
               list_item.SubItems(2) = rs!Cliente
               list_item.SubItems(3) = rs!PIEZAS
               list_item.SubItems(5) = rs!orden_pedido
               list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
               list_item.SubItems(7) = rsaux6!INVOICE_TO_ORG_ID
               list_item.SubItems(9) = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
               If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
                  strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!pedido)
                       .Parameters.Append parametro
                  End With
                  Set rsaux8 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
                  rsaux8.Close
                  strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                       .Parameters.Append parametro
                  End With
                  Set rsaux7 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  list_item.SubItems(8) = rsaux7!attribute1
                  rsaux7.Close
               Else
                  list_item.SubItems(8) = rsaux6!SHIP_TO_ORG_ID
               End If
               rsaux6.Close
               rs.MoveNext
         Wend
         rs.Close
End If
         'If lv_pedidos.ListItems.Count > 11 Then
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
         'Else
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
         'End If
         rs.Open "SELECT SUM(PIEZAS) FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
         Me.lbl_cantidad_10 = Format(rs(0).Value, "###,###,##0.00")
         rs.Close
      End If
   End If
   Call ilumina_grid
   
   
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_lineas)
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub lv_embarques_1_ItemClick(ByVal Item As MSComctlLib.ListItem)
         Me.txt_embarque = Me.lv_embarques_1.selectedItem
         If rs.State = 1 Then
            rs.Close
         End If
         
         rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria  from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_1.selectedItem) + " union all select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria  from tb_oracle_pedidos_asignados_embarques_vad where embarque = " + CStr(Me.lv_embarques_1.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         lv_pedidos.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
               If rs!pedido < 10000000 Then
                  strconsulta = "select * from oe_order_headers_all where order_number = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                       .Parameters.Append parametro
                  End With
                  Set rsaux6 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
               
                  list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
                  list_item.SubItems(2) = rs!Cliente
                  list_item.SubItems(3) = rs!PIEZAS
                  list_item.SubItems(5) = rs!orden_pedido
                  list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                  list_item.SubItems(7) = IIf(IsNull(rsaux6!INVOICE_TO_ORG_ID), "", rsaux6!INVOICE_TO_ORG_ID)
                  list_item.SubItems(9) = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
                  
                  If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
                     strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!pedido)
                          .Parameters.Append parametro
                     End With
                     Set rsaux8 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
                     rsaux8.Close
                     strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                          .Parameters.Append parametro
                     End With
                     Set rsaux7 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     list_item.SubItems(8) = rsaux7!attribute1
                     rsaux7.Close
                  Else
                     list_item.SubItems(8) = rsaux6!SHIP_TO_ORG_ID
                  End If
                  rsaux6.Close
                  Call ilumina_grid
               Else
                  list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
                  list_item.SubItems(2) = rs!Cliente
                  list_item.SubItems(3) = rs!PIEZAS
                  list_item.SubItems(5) = rs!orden_pedido
                  list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                  list_item.SubItems(7) = ""
                  list_item.SubItems(8) = ""
                  list_item.SubItems(9) = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
                  ilumina_grid
               End If
               rs.MoveNext
            Wend
            rs.Close
         'If lv_pedidos.ListItems.Count > 11 Then
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
         'Else
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
         'End If
         rs.Open "select sum(volumen) VOLUMEN from (select isnull(VOLUMEN,0) as volumen from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_1.selectedItem) + " union all select isnull(VOLUMEN,0) VOLUMEN from tb_oracle_pedidos_asignados_embarques_vad where embarque = " + CStr(Me.lv_embarques_1.selectedItem) + ") ped ", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_total_volumen = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
         Else
            Me.txt_total_volumen = 0
         End If
         rs.Close
         
         strconsulta = "select * from xxvia_tb_encabezado_embarques where embarque = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.lv_embarques_1.selectedItem))
              .Parameters.Append parametro
         End With
         Set rsaux6 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         var_volumen = 0
         If Not rsaux6.EOF Then
            var_transporte = IIf(IsNull(rsaux6!transporte), "", rsaux6!transporte)
         Else
            var_transporte = ""
         End If
         rsaux6.Close
         rs.Open "select * from tb_oracle_transportes where clave = '" + var_transporte + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_volumen_unidad = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
            Me.txt_transporte = IIf(IsNull(rs!NOMBRE), "", rs!NOMBRE)
         Else
            Me.txt_transporte = ""
            Me.txt_volumen_unidad = "0.00"
         End If
         rs.Close

End Sub

Private Sub lv_embarques_1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_embarque_asignar = CDbl(Me.lv_embarques_1.selectedItem)
      frmoracle_asignar_maquinas.Show 1
   End If
   If KeyCode = 119 Then
      var_embarque_lotes = Me.lv_embarques_1.selectedItem
      frmoracle_progreso_lotes.Show 1
   End If
   If KeyCode = 120 Then
      If Me.lv_embarques_1.ListItems.Count > 0 Then
         var_embarque_ruta = CDbl(Me.lv_embarques_1.selectedItem)
         frmoracle_asignar_ruta.Show 1
         var_tipo_asigna_ruta = 1
      End If
   End If
   If KeyCode = 121 And Shift = 1 Then
      If Me.cmd_imprimir_nuevo_metodo_divisiones.Enabled = True Then
         cmd_imprimir_nuevo_metodo_divisiones.Enabled = False
      Else
         Me.cmd_imprimir_nuevo_metodo_divisiones.Enabled = True
      End If
   End If
   
End Sub

Private Sub lv_embarques_10_ItemClick(ByVal Item As MSComctlLib.ListItem)
         Me.txt_embarque = Me.lv_embarques_10.selectedItem
         rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria  from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_10.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         lv_pedidos.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
               If rs!pedido < 10000000 Then
                  strconsulta = "select * from oe_order_headers_all where order_number = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                       .Parameters.Append parametro
                  End With
                  Set rsaux6 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
                  'rsaux.Close
                  list_item.SubItems(2) = rs!Cliente
                  list_item.SubItems(3) = rs!PIEZAS
                  list_item.SubItems(5) = rs!orden_pedido
                  list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                  list_item.SubItems(7) = rsaux6!INVOICE_TO_ORG_ID
                  list_item.SubItems(9) = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
                  If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
                     strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!pedido)
                          .Parameters.Append parametro
                     End With
                     Set rsaux8 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
                     rsaux8.Close
                     strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                          .Parameters.Append parametro
                     End With
                     Set rsaux7 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     list_item.SubItems(8) = rsaux7!attribute1
                     rsaux7.Close
                  Else
                     list_item.SubItems(8) = rsaux6!SHIP_TO_ORG_ID
                  End If
                  rsaux6.Close
                  Call ilumina_grid
               Else
                  list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
                  list_item.SubItems(2) = rs!Cliente
                  list_item.SubItems(3) = rs!PIEZAS
                  list_item.SubItems(5) = rs!orden_pedido
                  list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                  list_item.SubItems(7) = ""
                  list_item.SubItems(8) = ""
               End If
               rs.MoveNext
         Wend
         rs.Close
         'If lv_pedidos.ListItems.Count > 11 Then
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
         'Else
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
         'End If
         rs.Open "select sum(volumen) as volumen from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_10.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_total_volumen = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
         Else
            Me.txt_total_volumen = 0
         End If
         rs.Close

         strconsulta = "select * from xxvia_tb_encabezado_embarques where embarque = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.lv_embarques_10.selectedItem))
              .Parameters.Append parametro
         End With
         Set rsaux6 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         var_volumen = 0
         If Not rsaux6.EOF Then
            var_transporte = IIf(IsNull(rsaux6!transporte), "", rsaux6!transporte)
         Else
            var_transporte = ""
         End If
         rsaux6.Close
         rs.Open "select * from tb_oracle_transportes where clave = '" + var_transporte + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_volumen_unidad = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
            Me.txt_transporte = IIf(IsNull(rs!NOMBRE), "", rs!NOMBRE)
         Else
            Me.txt_transporte = ""
            Me.txt_volumen_unidad = "0.00"
         End If
         rs.Close

End Sub

Private Sub lv_embarques_10_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_embarque_asignar = CDbl(Me.lv_embarques_10.selectedItem)
      frmoracle_asignar_maquinas.Show 1
   End If
   If KeyCode = 119 Then
      var_embarque_lotes = Me.lv_embarques_10.selectedItem
      frmoracle_progreso_lotes.Show 1
   End If
   If KeyCode = 120 Then
      If Me.lv_embarques_10.ListItems.Count > 0 Then
         var_embarque_ruta = CDbl(Me.lv_embarques_10.selectedItem)
         frmoracle_asignar_ruta.Show 1
         var_tipo_asigna_ruta = 1
      End If
   End If
End Sub

Private Sub lv_embarques_2_ItemClick(ByVal Item As MSComctlLib.ListItem)
         Me.txt_embarque = Me.lv_embarques_2.selectedItem
         rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria  from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_2.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         lv_pedidos.ListItems.Clear
         While Not rs.EOF
         
               Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
               If IsNumeric(rs!pedido) Then
               
                If rs!pedido < 1000000 Then
                   strconsulta = "select * from oe_order_headers_all where order_number = ?"
                   With comandoORA
                        .ActiveConnection = cnnoracle_4
                        .CommandType = adCmdText
                        .CommandText = strconsulta
                        Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                        .Parameters.Append parametro
                   End With
                   Set rsaux6 = comandoORA.execute
                   Set comandoORA = Nothing
                   Set parametro = Nothing
                   'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
                   list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
                   'rsaux.Close
                   list_item.SubItems(2) = rs!Cliente
                   list_item.SubItems(3) = rs!PIEZAS
                   list_item.SubItems(5) = rs!orden_pedido
                   list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                   list_item.SubItems(7) = rsaux6!INVOICE_TO_ORG_ID
                   list_item.SubItems(9) = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
                   If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
                      strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                      With comandoORA
                           .ActiveConnection = cnnoracle_4
                           .CommandType = adCmdText
                           .CommandText = strconsulta
                           Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!pedido)
                           .Parameters.Append parametro
                      End With
                      Set rsaux8 = comandoORA.execute
                      Set comandoORA = Nothing
                      Set parametro = Nothing
                      var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
                      rsaux8.Close
                      strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                      With comandoORA
                           .ActiveConnection = cnnoracle_4
                           .CommandType = adCmdText
                           .CommandText = strconsulta
                           Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                           .Parameters.Append parametro
                      End With
                      Set rsaux7 = comandoORA.execute
                      Set comandoORA = Nothing
                      Set parametro = Nothing
                      list_item.SubItems(8) = rsaux7!attribute1
                      rsaux7.Close
                   Else
                      list_item.SubItems(8) = rsaux6!SHIP_TO_ORG_ID
                   End If
                   rsaux6.Close
                   Call ilumina_grid
                Else
                   list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
                   list_item.SubItems(2) = rs!Cliente
                   list_item.SubItems(3) = rs!PIEZAS
                   list_item.SubItems(5) = rs!orden_pedido
                   list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                   list_item.SubItems(7) = ""
                   list_item.SubItems(8) = ""
                End If
            Else
                list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
                list_item.SubItems(2) = rs!Cliente
                list_item.SubItems(3) = rs!PIEZAS
                list_item.SubItems(5) = rs!orden_pedido
                list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                list_item.SubItems(7) = ""
                list_item.SubItems(8) = ""
            End If
               rs.MoveNext
         Wend
         rs.Close
         'If lv_pedidos.ListItems.Count > 11 Then
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
         'Else
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
         'End If
         rs.Open "select sum(volumen) as volumen from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_2.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_total_volumen = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
         Else
            Me.txt_total_volumen = 0
         End If
         rs.Close
         
         
         
         strconsulta = "select * from xxvia_tb_encabezado_embarques where embarque = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.lv_embarques_2.selectedItem))
              .Parameters.Append parametro
         End With
         Set rsaux6 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         var_volumen = 0
         If Not rsaux6.EOF Then
            var_transporte = IIf(IsNull(rsaux6!transporte), "", rsaux6!transporte)
         Else
            var_transporte = ""
         End If
         rsaux6.Close
         rs.Open "select * from tb_oracle_transportes where clave = '" + var_transporte + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_volumen_unidad = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
            Me.txt_transporte = IIf(IsNull(rs!NOMBRE), "", rs!NOMBRE)
         Else
            Me.txt_transporte = ""
            Me.txt_volumen_unidad = "0.00"
         End If
         rs.Close
End Sub

Private Sub lv_embarques_2_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_embarque_asignar = CDbl(Me.lv_embarques_2.selectedItem)
      frmoracle_asignar_maquinas.Show 1
   End If
   If KeyCode = 119 Then
      var_embarque_lotes = Me.lv_embarques_2.selectedItem
      frmoracle_progreso_lotes.Show 1
   End If
   If KeyCode = 120 Then
      If Me.lv_embarques_2.ListItems.Count > 0 Then
         var_embarque_ruta = CDbl(Me.lv_embarques_2.selectedItem)
         frmoracle_asignar_ruta.Show 1
         var_tipo_asigna_ruta = 1
      End If
   End If
End Sub

Private Sub lv_embarques_3_ItemClick(ByVal Item As MSComctlLib.ListItem)
         Me.txt_embarque = Me.lv_embarques_3.selectedItem
         rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria  from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_3.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         lv_pedidos.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
               If rs!pedido < 10000000 Then
                  strconsulta = "select * from oe_order_headers_all where order_number = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                       .Parameters.Append parametro
                  End With
                  Set rsaux6 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
                  'rsaux.Close
                  list_item.SubItems(2) = rs!Cliente
                  list_item.SubItems(3) = rs!PIEZAS
                  list_item.SubItems(5) = rs!orden_pedido
                  list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                  list_item.SubItems(7) = rsaux6!INVOICE_TO_ORG_ID
                  list_item.SubItems(9) = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
                  If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
                     strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!pedido)
                          .Parameters.Append parametro
                     End With
                     Set rsaux8 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
                     rsaux8.Close
                     strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                          .Parameters.Append parametro
                     End With
                     Set rsaux7 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     list_item.SubItems(8) = rsaux7!attribute1
                     rsaux7.Close
                  Else
                     list_item.SubItems(8) = rsaux6!SHIP_TO_ORG_ID
                  End If
                  rsaux6.Close
                  Call ilumina_grid
               Else
                  list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
                  list_item.SubItems(2) = rs!Cliente
                  list_item.SubItems(3) = rs!PIEZAS
                  list_item.SubItems(5) = rs!orden_pedido
                  list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                  list_item.SubItems(7) = ""
                  list_item.SubItems(8) = ""
               End If
               rs.MoveNext
         Wend
         rs.Close
         'If lv_pedidos.ListItems.Count > 11 Then
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
         'Else
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
         'End If
         rs.Open "select sum(volumen) as volumen from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_3.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_total_volumen = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
         Else
            Me.txt_total_volumen = 0
         End If
         rs.Close

         strconsulta = "select * from xxvia_tb_encabezado_embarques where embarque = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.lv_embarques_3.selectedItem))
              .Parameters.Append parametro
         End With
         Set rsaux6 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         var_volumen = 0
         If Not rsaux6.EOF Then
            var_transporte = IIf(IsNull(rsaux6!transporte), "", rsaux6!transporte)
         Else
            var_transporte = ""
         End If
         rsaux6.Close
         rs.Open "select * from tb_oracle_transportes where clave = '" + var_transporte + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_volumen_unidad = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
            Me.txt_transporte = IIf(IsNull(rs!NOMBRE), "", rs!NOMBRE)
         Else
            Me.txt_transporte = ""
            Me.txt_volumen_unidad = "0.00"
         End If
         rs.Close

End Sub

Private Sub lv_embarques_3_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_embarque_asignar = CDbl(Me.lv_embarques_3.selectedItem)
      frmoracle_asignar_maquinas.Show 1
   End If
   If KeyCode = 119 Then
      var_embarque_lotes = Me.lv_embarques_3.selectedItem
      frmoracle_progreso_lotes.Show 1
   End If
   If KeyCode = 120 Then
      If Me.lv_embarques_3.ListItems.Count > 0 Then
         var_embarque_ruta = CDbl(Me.lv_embarques_3.selectedItem)
         frmoracle_asignar_ruta.Show 1
         var_tipo_asigna_ruta = 1
      End If
   End If
End Sub

Private Sub lv_embarques_4_ItemClick(ByVal Item As MSComctlLib.ListItem)
         Me.txt_embarque = Me.lv_embarques_4.selectedItem
         rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria  from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_4.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         lv_pedidos.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
               If rs!pedido < 10000000 Then
                  strconsulta = "select * from oe_order_headers_all where order_number = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                       .Parameters.Append parametro
                  End With
                  Set rsaux6 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
                  'rsaux.Close
                  list_item.SubItems(2) = rs!Cliente
                  list_item.SubItems(3) = rs!PIEZAS
                  list_item.SubItems(5) = rs!orden_pedido
                  list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                  list_item.SubItems(7) = rsaux6!INVOICE_TO_ORG_ID
                  list_item.SubItems(9) = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
                  If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
                     strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!pedido)
                          .Parameters.Append parametro
                     End With
                     Set rsaux8 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
                     rsaux8.Close
                     strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                          .Parameters.Append parametro
                     End With
                     Set rsaux7 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     list_item.SubItems(8) = rsaux7!attribute1
                     rsaux7.Close
                  Else
                     list_item.SubItems(8) = rsaux6!SHIP_TO_ORG_ID
                  End If
                  rsaux6.Close
                  Call ilumina_grid
               Else
                  list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
                  list_item.SubItems(2) = rs!Cliente
                  list_item.SubItems(3) = rs!PIEZAS
                  list_item.SubItems(5) = rs!orden_pedido
                  list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                  list_item.SubItems(7) = ""
                  list_item.SubItems(8) = ""
               End If
               rs.MoveNext
         Wend
         rs.Close
         'If lv_pedidos.ListItems.Count > 11 Then
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
         'Else
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
         'End If
         rs.Open "select sum(volumen) as volumen from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_4.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_total_volumen = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
         Else
            Me.txt_total_volumen = 0
         End If
         rs.Close

         strconsulta = "select * from xxvia_tb_encabezado_embarques where embarque = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.lv_embarques_4.selectedItem))
              .Parameters.Append parametro
         End With
         Set rsaux6 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         var_volumen = 0
         If Not rsaux6.EOF Then
            var_transporte = IIf(IsNull(rsaux6!transporte), "", rsaux6!transporte)
         Else
            var_transporte = ""
         End If
         rsaux6.Close
         rs.Open "select * from tb_oracle_transportes where clave = '" + var_transporte + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_volumen_unidad = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
            Me.txt_transporte = IIf(IsNull(rs!NOMBRE), "", rs!NOMBRE)
         Else
            Me.txt_transporte = ""
            Me.txt_volumen_unidad = "0.00"
         End If
         rs.Close

End Sub

Private Sub lv_embarques_4_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_embarque_asignar = CDbl(Me.lv_embarques_4.selectedItem)
      frmoracle_asignar_maquinas.Show 1
   End If
   If KeyCode = 119 Then
      var_embarque_lotes = Me.lv_embarques_4.selectedItem
      frmoracle_progreso_lotes.Show 1
   End If
   If KeyCode = 120 Then
      If Me.lv_embarques_4.ListItems.Count > 0 Then
         var_embarque_ruta = CDbl(Me.lv_embarques_4.selectedItem)
         frmoracle_asignar_ruta.Show 1
         var_tipo_asigna_ruta = 1
      End If
   End If
End Sub

Private Sub lv_embarques_5_ItemClick(ByVal Item As MSComctlLib.ListItem)
         Me.txt_embarque = Me.lv_embarques_5.selectedItem
         rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria  from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_5.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         lv_pedidos.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
               If rs!pedido < 10000000 Then
                  strconsulta = "select * from oe_order_headers_all where order_number = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                       .Parameters.Append parametro
                  End With
                  Set rsaux6 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
                  'rsaux.Close
                  list_item.SubItems(2) = rs!Cliente
                  list_item.SubItems(3) = rs!PIEZAS
                  list_item.SubItems(5) = rs!orden_pedido
                  list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                  list_item.SubItems(7) = rsaux6!INVOICE_TO_ORG_ID
                  list_item.SubItems(9) = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
                  If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
                     strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!pedido)
                          .Parameters.Append parametro
                     End With
                     Set rsaux8 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
                     rsaux8.Close
                     strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                          .Parameters.Append parametro
                     End With
                     Set rsaux7 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     list_item.SubItems(8) = rsaux7!attribute1
                     rsaux7.Close
                  Else
                     list_item.SubItems(8) = rsaux6!SHIP_TO_ORG_ID
                  End If
                  rsaux6.Close
                  Call ilumina_grid
               Else
                  list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
                  list_item.SubItems(2) = rs!Cliente
                  list_item.SubItems(3) = rs!PIEZAS
                  list_item.SubItems(5) = rs!orden_pedido
                  list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                  list_item.SubItems(7) = ""
                  list_item.SubItems(8) = ""
               End If
               rs.MoveNext
         Wend
         rs.Close
         'If lv_pedidos.ListItems.Count > 11 Then
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
         'Else
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
         'End If
         rs.Open "select sum(volumen) as volumen from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_5.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_total_volumen = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
         Else
            Me.txt_total_volumen = 0
         End If
         rs.Close

         strconsulta = "select * from xxvia_tb_encabezado_embarques where embarque = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.lv_embarques_5.selectedItem))
              .Parameters.Append parametro
         End With
         Set rsaux6 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         var_volumen = 0
         If Not rsaux6.EOF Then
            var_transporte = IIf(IsNull(rsaux6!transporte), "", rsaux6!transporte)
         Else
            var_transporte = ""
         End If
         rsaux6.Close
         rs.Open "select * from tb_oracle_transportes where clave = '" + var_transporte + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_volumen_unidad = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
            Me.txt_transporte = IIf(IsNull(rs!NOMBRE), "", rs!NOMBRE)
         Else
            Me.txt_transporte = ""
            Me.txt_volumen_unidad = "0.00"
         End If
         rs.Close

End Sub

Private Sub lv_embarques_5_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_embarque_asignar = CDbl(Me.lv_embarques_5.selectedItem)
      frmoracle_asignar_maquinas.Show 1
   End If
   If KeyCode = 119 Then
      var_embarque_lotes = Me.lv_embarques_5.selectedItem
      frmoracle_progreso_lotes.Show 1
   End If
   If KeyCode = 120 Then
      If Me.lv_embarques_5.ListItems.Count > 0 Then
         var_embarque_ruta = CDbl(Me.lv_embarques_5.selectedItem)
         frmoracle_asignar_ruta.Show 1
         var_tipo_asigna_ruta = 1
      End If
   End If
End Sub

Private Sub lv_embarques_6_ItemClick(ByVal Item As MSComctlLib.ListItem)
         Me.txt_embarque = Me.lv_embarques_6.selectedItem
         rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria  from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_6.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         lv_pedidos.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
               If rs!pedido < 10000000 Then
                  strconsulta = "select * from oe_order_headers_all where order_number = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                       .Parameters.Append parametro
                  End With
                  Set rsaux6 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
                  'rsaux.Close
                  list_item.SubItems(2) = rs!Cliente
                  list_item.SubItems(3) = rs!PIEZAS
                  list_item.SubItems(5) = rs!orden_pedido
                  list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                  list_item.SubItems(7) = rsaux6!INVOICE_TO_ORG_ID
                  list_item.SubItems(9) = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
                  If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
                     strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!pedido)
                          .Parameters.Append parametro
                     End With
                     Set rsaux8 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
                     rsaux8.Close
                     strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                          .Parameters.Append parametro
                     End With
                     Set rsaux7 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     list_item.SubItems(8) = rsaux7!attribute1
                     rsaux7.Close
                  Else
                     list_item.SubItems(8) = rsaux6!SHIP_TO_ORG_ID
                  End If
                  rsaux6.Close
                  Call ilumina_grid
               Else
                  list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
                  list_item.SubItems(2) = rs!Cliente
                  list_item.SubItems(3) = rs!PIEZAS
                  list_item.SubItems(5) = rs!orden_pedido
                  list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                  list_item.SubItems(7) = ""
                  list_item.SubItems(8) = ""
               End If
               rs.MoveNext
         Wend
         rs.Close
         'If lv_pedidos.ListItems.Count > 11 Then
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
         'Else
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
         'End If
         rs.Open "select sum(volumen) as volumen from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_6.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_total_volumen = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
         Else
            Me.txt_total_volumen = 0
         End If
         rs.Close

         strconsulta = "select * from xxvia_tb_encabezado_embarques where embarque = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.lv_embarques_6.selectedItem))
              .Parameters.Append parametro
         End With
         Set rsaux6 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         var_volumen = 0
         If Not rsaux6.EOF Then
            var_transporte = IIf(IsNull(rsaux6!transporte), "", rsaux6!transporte)
         Else
            var_transporte = ""
         End If
         rsaux6.Close
         rs.Open "select * from tb_oracle_transportes where clave = '" + var_transporte + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_volumen_unidad = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
            Me.txt_transporte = IIf(IsNull(rs!NOMBRE), "", rs!NOMBRE)
         Else
            Me.txt_transporte = ""
            Me.txt_volumen_unidad = "0.00"
         End If
         rs.Close

End Sub

Private Sub lv_embarques_6_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_embarque_asignar = CDbl(Me.lv_embarques_6.selectedItem)
      frmoracle_asignar_maquinas.Show 1
   End If
   If KeyCode = 119 Then
      var_embarque_lotes = Me.lv_embarques_6.selectedItem
      frmoracle_progreso_lotes.Show 1
   End If
   If KeyCode = 120 Then
      If Me.lv_embarques_6.ListItems.Count > 0 Then
         var_embarque_ruta = CDbl(Me.lv_embarques_6.selectedItem)
         frmoracle_asignar_ruta.Show 1
         var_tipo_asigna_ruta = 1
      End If
   End If
End Sub

Private Sub lv_embarques_7_ItemClick(ByVal Item As MSComctlLib.ListItem)
         Me.txt_embarque = Me.lv_embarques_7.selectedItem
         rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria  from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_7.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         lv_pedidos.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
               If rs!pedido < 10000000 Then
                  strconsulta = "select * from oe_order_headers_all where order_number = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                       .Parameters.Append parametro
                  End With
                  Set rsaux6 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
                  'rsaux.Close
                  list_item.SubItems(2) = rs!Cliente
                  list_item.SubItems(3) = rs!PIEZAS
                  list_item.SubItems(5) = rs!orden_pedido
                  list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                  list_item.SubItems(7) = rsaux6!INVOICE_TO_ORG_ID
                  list_item.SubItems(9) = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
                  If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
                     strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!pedido)
                          .Parameters.Append parametro
                     End With
                     Set rsaux8 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
                     rsaux8.Close
                     strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                          .Parameters.Append parametro
                     End With
                     Set rsaux7 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     list_item.SubItems(8) = rsaux7!attribute1
                     rsaux7.Close
                  Else
                     list_item.SubItems(8) = rsaux6!SHIP_TO_ORG_ID
                  End If
                  rsaux6.Close
                  Call ilumina_grid
               Else
                  list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
                  list_item.SubItems(2) = rs!Cliente
                  list_item.SubItems(3) = rs!PIEZAS
                  list_item.SubItems(5) = rs!orden_pedido
                  list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                  list_item.SubItems(7) = ""
                  list_item.SubItems(8) = ""
               End If
               rs.MoveNext
         Wend
         rs.Close
         'If lv_pedidos.ListItems.Count > 11 Then
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
         'Else
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
         'End If
         rs.Open "select sum(volumen) as volumen from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_7.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_total_volumen = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
         Else
            Me.txt_total_volumen = 0
         End If
         rs.Close

         strconsulta = "select * from xxvia_tb_encabezado_embarques where embarque = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.lv_embarques_7.selectedItem))
              .Parameters.Append parametro
         End With
         Set rsaux6 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         var_volumen = 0
         If Not rsaux6.EOF Then
            var_transporte = IIf(IsNull(rsaux6!transporte), "", rsaux6!transporte)
         Else
            var_transporte = ""
         End If
         rsaux6.Close
         rs.Open "select * from tb_oracle_transportes where clave = '" + var_transporte + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_volumen_unidad = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
            Me.txt_transporte = IIf(IsNull(rs!NOMBRE), "", rs!NOMBRE)
         Else
            Me.txt_transporte = ""
            Me.txt_volumen_unidad = "0.00"
         End If
         rs.Close

End Sub

Private Sub lv_embarques_7_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_embarque_asignar = CDbl(Me.lv_embarques_7.selectedItem)
      frmoracle_asignar_maquinas.Show 1
   End If
   If KeyCode = 119 Then
      var_embarque_lotes = Me.lv_embarques_7.selectedItem
      frmoracle_progreso_lotes.Show 1
   End If
   If KeyCode = 120 Then
      If Me.lv_embarques_7.ListItems.Count > 0 Then
         var_embarque_ruta = CDbl(Me.lv_embarques_7.selectedItem)
         frmoracle_asignar_ruta.Show 1
         var_tipo_asigna_ruta = 1
      End If
   End If
End Sub

Private Sub lv_embarques_8_ItemClick(ByVal Item As MSComctlLib.ListItem)
         Me.txt_embarque = Me.lv_embarques_8.selectedItem
         rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria  from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_8.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         lv_pedidos.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
               If rs!pedido < 10000000 Then
                  strconsulta = "select * from oe_order_headers_all where order_number = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                       .Parameters.Append parametro
                  End With
                  Set rsaux6 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
                  'rsaux.Close
                  list_item.SubItems(2) = rs!Cliente
                  list_item.SubItems(3) = rs!PIEZAS
                  list_item.SubItems(5) = rs!orden_pedido
                  list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                  list_item.SubItems(7) = rsaux6!INVOICE_TO_ORG_ID
                  list_item.SubItems(9) = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
                  If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
                     strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!pedido)
                          .Parameters.Append parametro
                     End With
                     Set rsaux8 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
                     rsaux8.Close
                     strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                          .Parameters.Append parametro
                     End With
                     Set rsaux7 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     list_item.SubItems(8) = rsaux7!attribute1
                     rsaux7.Close
                  Else
                     list_item.SubItems(8) = rsaux6!SHIP_TO_ORG_ID
                  End If
                  rsaux6.Close
                  Call ilumina_grid
               Else
                  list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
                  list_item.SubItems(2) = rs!Cliente
                  list_item.SubItems(3) = rs!PIEZAS
                  list_item.SubItems(5) = rs!orden_pedido
                  list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                  list_item.SubItems(7) = ""
                  list_item.SubItems(8) = ""
               End If
               rs.MoveNext
         Wend
         rs.Close
         'If lv_pedidos.ListItems.Count > 11 Then
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
         'Else
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
         'End If
         rs.Open "select sum(volumen) as volumen from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_8.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_total_volumen = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
         Else
            Me.txt_total_volumen = 0
         End If
         rs.Close

         strconsulta = "select * from xxvia_tb_encabezado_embarques where embarque = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.lv_embarques_8.selectedItem))
              .Parameters.Append parametro
         End With
         Set rsaux6 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         var_volumen = 0
         If Not rsaux6.EOF Then
            var_transporte = IIf(IsNull(rsaux6!transporte), "", rsaux6!transporte)
         Else
            var_transporte = ""
         End If
         rsaux6.Close
         rs.Open "select * from tb_oracle_transportes where clave = '" + var_transporte + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_volumen_unidad = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
            Me.txt_transporte = IIf(IsNull(rs!NOMBRE), "", rs!NOMBRE)
         Else
            Me.txt_transporte = ""
            Me.txt_volumen_unidad = "0.00"
         End If
         rs.Close

End Sub

Private Sub lv_embarques_8_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_embarque_asignar = CDbl(Me.lv_embarques_8.selectedItem)
      frmoracle_asignar_maquinas.Show 1
   End If
   If KeyCode = 119 Then
      var_embarque_lotes = Me.lv_embarques_8.selectedItem
      frmoracle_progreso_lotes.Show 1
   End If
   If KeyCode = 120 Then
      If Me.lv_embarques_8.ListItems.Count > 0 Then
         var_embarque_ruta = CDbl(Me.lv_embarques_8.selectedItem)
         frmoracle_asignar_ruta.Show 1
         var_tipo_asigna_ruta = 1
      End If
   End If
End Sub

Private Sub lv_embarques_9_ItemClick(ByVal Item As MSComctlLib.ListItem)
         Me.txt_embarque = Me.lv_embarques_9.selectedItem
         rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria  from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_9.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         lv_pedidos.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
               If rs!pedido < 10000000 Then
                  strconsulta = "select * from oe_order_headers_all where order_number = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                       .Parameters.Append parametro
                  End With
                  Set rsaux6 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
                  'rsaux.Close
                  list_item.SubItems(2) = rs!Cliente
                  list_item.SubItems(3) = rs!PIEZAS
                  list_item.SubItems(5) = rs!orden_pedido
                  list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                  list_item.SubItems(7) = rsaux6!INVOICE_TO_ORG_ID
                  list_item.SubItems(9) = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
                  If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
                     strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!pedido)
                          .Parameters.Append parametro
                     End With
                     Set rsaux8 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
                     rsaux8.Close
                     strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                          .Parameters.Append parametro
                     End With
                     Set rsaux7 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     list_item.SubItems(8) = rsaux7!attribute1
                     rsaux7.Close
                  Else
                     list_item.SubItems(8) = rsaux6!SHIP_TO_ORG_ID
                  End If
                  rsaux6.Close
                  Call ilumina_grid
               Else
                  list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
                  list_item.SubItems(2) = rs!Cliente
                  list_item.SubItems(3) = rs!PIEZAS
                  list_item.SubItems(5) = rs!orden_pedido
                  list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                  list_item.SubItems(7) = ""
                  list_item.SubItems(8) = ""
               End If
               rs.MoveNext
         Wend
         rs.Close
         'If lv_pedidos.ListItems.Count > 11 Then
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
         'Else
         '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
         'End If
         rs.Open "select sum(volumen) as volumen from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_9.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_total_volumen = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
         Else
            Me.txt_total_volumen = 0
         End If
         rs.Close

         strconsulta = "select * from xxvia_tb_encabezado_embarques where embarque = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.lv_embarques_9.selectedItem))
              .Parameters.Append parametro
         End With
         Set rsaux6 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         var_volumen = 0
         If Not rsaux6.EOF Then
            var_transporte = IIf(IsNull(rsaux6!transporte), "", rsaux6!transporte)
         Else
            var_transporte = ""
         End If
         rsaux6.Close
         rs.Open "select * from tb_oracle_transportes where clave = '" + var_transporte + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_volumen_unidad = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
            Me.txt_transporte = IIf(IsNull(rs!NOMBRE), "", rs!NOMBRE)
         Else
            Me.txt_transporte = ""
            Me.txt_volumen_unidad = "0.00"
         End If
         rs.Close

End Sub

Private Sub lv_embarques_9_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_embarque_asignar = CDbl(Me.lv_embarques_9.selectedItem)
      frmoracle_asignar_maquinas.Show 1
   End If
   If KeyCode = 119 Then
      var_embarque_lotes = Me.lv_embarques_9.selectedItem
      frmoracle_progreso_lotes.Show 1
   End If
   If KeyCode = 120 Then
      If Me.lv_embarques_9.ListItems.Count > 0 Then
         var_embarque_ruta = CDbl(Me.lv_embarques_9.selectedItem)
         frmoracle_asignar_ruta.Show 1
         var_tipo_asigna_ruta = 1
      End If
   End If
End Sub

Private Sub lv_pedidos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  
If ColumnHeader = "Orden de carga" Then
   Call pro_ordena_listas(Me.lv_pedidos, ColumnHeader)
Else
   Call pro_ordena_listas(Me.lv_pedidos, ColumnHeader)
End If
   
End Sub

Private Sub lv_pedidos_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 117 Then
      var_si = MsgBox("¿Desea cambiar el transporte?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         var_transporte_global = ""
         frmoracle_transortes.Show 1
         If var_transporte_global <> "" Then
            rs.Open "UPDATE XXVIA_TB_ENCABEZADO_EMBARQUES SET TRANSPORTE = '" + var_transporte_global + "' WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
            MsgBox "Se a actualizado el transporte", vbOKOnly, "ATENCION"
         Else
            MsgBox "No se selecciono un transporte", vbOKOnly, "ATENCION"
         End If
      End If
   End If
   
   If Shift = 1 And KeyCode = 80 Then
      var_i = Me.lv_pedidos.selectedItem.Index
      If Me.lv_pedidos.ListItems.Count > 0 Then
         rs.Open "SELECT * FROM TB_oracle_PEDIDOS_ASIGNADOS_EMBARQUES WHERE PEDIDO = " + CStr(Me.lv_pedidos.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_estatus_paqueteria = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
            If var_estatus_paqueteria = 0 Then
               var_transporte_global = ""
               If var_transporte_global = "" Then
               
               rsaux.Open "UPDATE TB_oracle_PEDIDOS_ASIGNADOS_EMBARQUES SET PAQUETERIA = 1 WHERE PEDIDO = " + Me.lv_pedidos.selectedItem, cnn, adOpenDynamic, adLockOptimistic
               Me.lv_pedidos.selectedItem.SubItems(9) = "1"
               lv_pedidos.ListItems.Item(var_i).Bold = True
               lv_pedidos.ListItems.Item(var_i).ListSubItems(1).Bold = True
               lv_pedidos.ListItems.Item(var_i).ListSubItems(2).Bold = True
               lv_pedidos.ListItems.Item(var_i).ListSubItems(3).Bold = True
               lv_pedidos.ListItems.Item(var_i).ListSubItems(4).Bold = True
               lv_pedidos.ListItems.Item(var_i).ListSubItems(5).Bold = True
               lv_pedidos.ListItems.Item(var_i).ListSubItems(6).Bold = True
               lv_pedidos.ListItems.Item(var_i).ListSubItems(7).Bold = True
               lv_pedidos.ListItems.Item(var_i).ListSubItems(8).Bold = True
               lv_pedidos.ListItems.Item(var_i).ListSubItems(9).Bold = True
               lv_pedidos.ListItems.Item(var_i).ForeColor = &H8000000D
               lv_pedidos.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H8000000D
               lv_pedidos.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H8000000D
               lv_pedidos.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H8000000D
               lv_pedidos.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H8000000D
               lv_pedidos.ListItems.Item(var_i).ListSubItems(5).ForeColor = &H8000000D
               lv_pedidos.ListItems.Item(var_i).ListSubItems(6).ForeColor = &H8000000D
               lv_pedidos.ListItems.Item(var_i).ListSubItems(7).ForeColor = &H8000000D
               lv_pedidos.ListItems.Item(var_i).ListSubItems(8).ForeColor = &H8000000D
               lv_pedidos.ListItems.Item(var_i).ListSubItems(9).ForeColor = &H8000000D
               Else
                  MsgBox "Debe de seleccionar el transporte", vbOKOnly, "ATENCION"
               
               End If
               
               
               
            Else
               rsaux.Open "UPDATE TB_oracle_PEDIDOS_ASIGNADOS_EMBARQUES SET PAQUETERIA = 0 WHERE PEDIDO = " + Me.lv_pedidos.selectedItem, cnn, adOpenDynamic, adLockOptimistic
               Me.lv_pedidos.selectedItem.SubItems(9) = "0"
               lv_pedidos.ListItems.Item(var_i).Bold = False
               lv_pedidos.ListItems.Item(var_i).ListSubItems(1).Bold = False
               lv_pedidos.ListItems.Item(var_i).ListSubItems(2).Bold = False
               lv_pedidos.ListItems.Item(var_i).ListSubItems(3).Bold = False
               lv_pedidos.ListItems.Item(var_i).ListSubItems(4).Bold = False
               lv_pedidos.ListItems.Item(var_i).ListSubItems(5).Bold = False
               lv_pedidos.ListItems.Item(var_i).ListSubItems(6).Bold = False
               lv_pedidos.ListItems.Item(var_i).ListSubItems(7).Bold = False
               lv_pedidos.ListItems.Item(var_i).ListSubItems(8).Bold = False
               lv_pedidos.ListItems.Item(var_i).ListSubItems(9).Bold = False
               lv_pedidos.ListItems.Item(var_i).ForeColor = &H80000007
               lv_pedidos.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H80000007
               lv_pedidos.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H80000007
               lv_pedidos.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H80000007
               lv_pedidos.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H80000007
               lv_pedidos.ListItems.Item(var_i).ListSubItems(5).ForeColor = &H80000007
               lv_pedidos.ListItems.Item(var_i).ListSubItems(6).ForeColor = &H80000007
               lv_pedidos.ListItems.Item(var_i).ListSubItems(7).ForeColor = &H80000007
               lv_pedidos.ListItems.Item(var_i).ListSubItems(8).ForeColor = &H80000007
               lv_pedidos.ListItems.Item(var_i).ListSubItems(9).ForeColor = &H80000007
            End If
         End If
         rs.Close
      End If
   End If
   If KeyCode = 114 Then
      var_si = MsgBox("¿Desea eliminar el pedido del embarque?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         var_si = MsgBox("Confirmar la eliminación del pedido del embarque?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            If rsaux.State = 1 Then
               rsaux.Close
            End If
            rsaux.Open "SELECT * FROM XXVIA_TB_SALIDAS_CAJAS WHERE SOURCE_HEADER_NUMBER = " + Me.lv_pedidos.selectedItem, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If rsaux.EOF Then
               rs.Open "update TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES SET EMBARQUE = 0 WHERE PEDIDO = " + Me.lv_pedidos.selectedItem, cnn, adOpenDynamic, adLockOptimistic
               lv_pedidos.ListItems.Remove (lv_pedidos.selectedItem.Index)
            Else
               MsgBox "El pedido no puede ser eliminado ya que se leyeron piezas", vbOKOnly, "ATENCION"
            End If
            rsaux.Close
         End If
      End If
   End If
End Sub

Private Sub lv_pedidos_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 13
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii > 0 Then
      If KeyAscii = 13 Then
      
         If var_orden <> "" Then
            var_j = Len(Trim(var_orden))
            If var_j <= 3 Then
               Me.lv_pedidos.ListItems.Item(var_posicion).Selected = True
               x = 1
               If x = 1 Then 'para que ya no se pueda cambiar el orden de carga
                  If var_clave_usuario_global = "U0000000011" Or var_clave_usuario_global = "U0000000003" Or var_clave_usuario_global = "U00000000032" Or var_clave_usuario_global = "U0000000346" Then
                     rs.Open "update TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES set orden_pedido = " + var_orden + " where pedido = " + Me.lv_pedidos.selectedItem, cnn, adOpenDynamic, adLockOptimistic
                     Me.lv_pedidos.selectedItem.SubItems(5) = var_orden
                     var_orden = ""
                     Me.frm_orden.Visible = False
                     Me.lbl_orden = ""
                     var_posicion = 0
                  End If
               End If
             Else
                MsgBox "Número incorrecto " + var_orden, vbOKOnly
                Me.lbl_orden = ""
                var_orden = ""
                Me.frm_orden.Visible = False
                var_posicion = 0
             End If
         End If
      Else
         x = 1
         If x = 1 Then 'para que no puedan cambiar el orden de la carga
            If var_clave_usuario_global = "U0000000011" Or var_clave_usuario_global = "U0000000346" Or var_clave_usuario_global = "U0000000003" Then
               Me.frm_orden.Visible = True
               var_orden = var_orden + Chr(KeyAscii)
               Me.lbl_orden.Caption = var_orden
               If var_posicion = 0 Then
                  var_posicion = Me.lv_pedidos.selectedItem.Index
               End If
            End If
         End If
      End If
   End If
End Sub

Private Sub lv_pedidos_LostFocus()
   Me.frm_orden.Visible = False
End Sub

Private Sub txt_fecha_Change()
   Me.txt_porcentaje = ""
   Me.txt_total_volumen = ""
   Me.txt_transporte = ""
   Me.txt_volumen_unidad = ""
End Sub
Private Sub txt_fecha_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsDate(Me.txt_fecha) Then
         Me.lv_pedidos.ListItems.Clear
         Me.lv_embarques_1.ListItems.Clear
         Me.lv_embarques_2.ListItems.Clear
         Me.lv_embarques_3.ListItems.Clear
         Me.lv_embarques_4.ListItems.Clear
         Me.lv_embarques_5.ListItems.Clear
         Me.lv_embarques_6.ListItems.Clear
         Me.lv_embarques_7.ListItems.Clear
         Me.lv_embarques_8.ListItems.Clear
         Me.lv_embarques_9.ListItems.Clear
         Me.lv_embarques_10.ListItems.Clear
         Me.lbl_cantidad_1 = "0.00"
         Me.lbl_cantidad_2 = "0.00"
         Me.lbl_cantidad_3 = "0.00"
         Me.lbl_cantidad_4 = "0.00"
         Me.lbl_cantidad_5 = "0.00"
         Me.lbl_cantidad_6 = "0.00"
         Me.lbl_cantidad_7 = "0.00"
         Me.lbl_cantidad_8 = "0.00"
         Me.lbl_cantidad_9 = "0.00"
         Me.lbl_cantidad_10 = "0.00"
         Me.frm_orden.Visible = False
         Me.lv_embarques_1.ListItems.Clear
         var_dia_s = CStr(Day(CDate(Me.txt_fecha)))
         If Len(var_dia_s) = 1 Then
            var_dia_s = "0" + var_dia_s
         End If
         var_mes_s = CStr(Month(CDate(Me.txt_fecha)))
         If Len(var_mes_s) = 1 Then
            var_mes_s = "0" + var_mes_s
         End If
         rs.Open "select distinct embarque as embarque from xxvia_Tb_encabezado_embarques where jaula = 1 and to_char(fecha_inicio,'DD')  = '" + CStr(var_dia_s) + "' and to_char(fecha_inicio,'MM')  = '" + CStr(var_mes_s) + "' and to_char(fecha_inicio,'yyyy')  = '" + CStr(Year(CDate(Me.txt_fecha))) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         var_Cadena_embarques = ""
         While Not rs.EOF
               If var_Cadena_embarques = "" Then
                  var_Cadena_embarques = CStr(rs!Embarque)
               Else
                  var_Cadena_embarques = var_Cadena_embarques + "," + CStr(rs!Embarque)
               End If
               rs.MoveNext
         Wend
         rs.Close
         If var_Cadena_embarques <> "" Then
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open "select distinct EMBARQUE from tb_oracle_pedidos_asignados_embarques where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  Set list_item = lv_embarques_1.ListItems.Add(, , rs!Embarque)
                  strconsulta = "select nvl(char_emb_estatus,' ') as estatus from xxvia_Tb_encabezado_embarques where embarque = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rs!Embarque)
                       .Parameters.Append parametro
                  End With
                  Set rsaux8 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  If rsaux8!estatus = " " Or rsaux8!estatus = "E" Then
                     list_item.Bold = False
                     list_item.ForeColor = &H80000012
                  Else
                     list_item.Bold = True
                     list_item.ForeColor = &H8000&
                  End If
                  rs.MoveNext
            Wend
            rs.Close
            If lv_embarques_1.ListItems.Count > 0 Then
               lv_embarques_1.ListItems(1).Selected = True
               rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria  from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_1.selectedItem), cnn, adOpenDynamic, adLockOptimistic
               lv_pedidos.ListItems.Clear
               While Not rs.EOF
                     Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
                     strconsulta = "select * from oe_order_headers_all where order_number = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                          .Parameters.Append parametro
                     End With
                     Set rsaux6 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
                     'rsaux.Close
                     list_item.SubItems(2) = rs!Cliente
                     list_item.SubItems(3) = rs!PIEZAS
                     list_item.SubItems(5) = rs!orden_pedido
                     list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                     list_item.SubItems(7) = rsaux6!INVOICE_TO_ORG_ID
                     list_item.SubItems(9) = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
                     If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
                        strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                        With comandoORA
                            .ActiveConnection = cnnoracle_4
                            .CommandType = adCmdText
                            .CommandText = strconsulta
                            Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!pedido)
                            .Parameters.Append parametro
                        End With
                        Set rsaux8 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
                        rsaux8.Close
                        strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                             .Parameters.Append parametro
                        End With
                        Set rsaux7 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        list_item.SubItems(8) = rsaux7!attribute1
                        rsaux7.Close
                     Else
                        list_item.SubItems(8) = rsaux6!SHIP_TO_ORG_ID
                     End If
                     rsaux6.Close
                     rs.MoveNext
               Wend
               rs.Close
               'If lv_pedidos.ListItems.Count > 11 Then
               '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
               'Else
               '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
               'End If
               
               rs.Open "SELECT SUM(PIEZAS) FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
               Me.lbl_cantidad_1 = Format(rs(0).Value, "###,###,##0.00")
               rs.Close
            
               rs.Open "select sum(volumen) as volumen from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_1.selectedItem), cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  Me.txt_total_volumen = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
               Else
                  Me.txt_total_volumen = 0
               End If
               rs.Close
         
               strconsulta = "select * from xxvia_tb_encabezado_embarques where embarque = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.lv_embarques_1.selectedItem))
                    .Parameters.Append parametro
               End With
               Set rsaux6 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               var_volumen = 0
               If Not rsaux6.EOF Then
                  var_transporte = IIf(IsNull(rsaux6!transporte), "", rsaux6!transporte)
               Else
                  var_transporte = ""
               End If
               rsaux6.Close
         
               rs.Open "select * from tb_oracle_transportes where clave = '" + var_transporte + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  Me.txt_volumen_unidad = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                  Me.txt_transporte = IIf(IsNull(rs!NOMBRE), "", rs!NOMBRE)
               Else
                  Me.txt_transporte = ""
                  Me.txt_volumen_unidad = "0.00"
               End If
               rs.Close
            End If
         End If
      
         rs.Open "select distinct embarque as embarque from xxvia_Tb_encabezado_embarques where jaula = 2 and to_char(fecha_inicio,'DD')  = '" + CStr(var_dia_s) + "' and to_char(fecha_inicio,'MM')  = '" + CStr(var_mes_s) + "' and to_char(fecha_inicio,'yyyy')  = '" + CStr(Year(CDate(Me.txt_fecha))) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         var_Cadena_embarques = ""
         While Not rs.EOF
               If var_Cadena_embarques = "" Then
                  var_Cadena_embarques = CStr(rs!Embarque)
               Else
                  var_Cadena_embarques = var_Cadena_embarques + "," + CStr(rs!Embarque)
               End If
               rs.MoveNext
         Wend
         rs.Close
         If var_Cadena_embarques <> "" Then
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open "select distinct EMBARQUE from tb_oracle_pedidos_asignados_embarques where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  Set list_item = lv_embarques_2.ListItems.Add(, , rs!Embarque)
                  strconsulta = "select nvl(char_emb_estatus,' ') as estatus from xxvia_Tb_encabezado_embarques where embarque = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rs!Embarque)
                       .Parameters.Append parametro
                  End With
                  Set rsaux8 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  If rsaux8!estatus = " " Or rsaux8!estatus = "E" Then
                     list_item.Bold = False
                     list_item.ForeColor = &H80000012
                  Else
                     list_item.Bold = True
                     list_item.ForeColor = &H8000&
                  End If
                  
                  'list_item.SubItems(1) = rs!NOMBRE_AGENTE
                  rs.MoveNext
            Wend
            rs.Close
            If lv_embarques_2.ListItems.Count > 0 Then
               lv_embarques_2.ListItems(1).Selected = True
               rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria  from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_2.selectedItem), cnn, adOpenDynamic, adLockOptimistic
               lv_pedidos.ListItems.Clear
               While Not rs.EOF
                     Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
                     strconsulta = "select * from oe_order_headers_all where order_number = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                          .Parameters.Append parametro
                     End With
                     Set rsaux6 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
                     'rsaux.Close
                     list_item.SubItems(2) = rs!Cliente
                     list_item.SubItems(3) = rs!PIEZAS
                     list_item.SubItems(5) = rs!orden_pedido
                     list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                     list_item.SubItems(7) = rsaux6!INVOICE_TO_ORG_ID
                     list_item.SubItems(9) = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
                     If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
                        strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                        With comandoORA
                            .ActiveConnection = cnnoracle_4
                            .CommandType = adCmdText
                            .CommandText = strconsulta
                            Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!pedido)
                            .Parameters.Append parametro
                        End With
                        Set rsaux8 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
                        rsaux8.Close
                        strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                             .Parameters.Append parametro
                        End With
                        Set rsaux7 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        list_item.SubItems(8) = rsaux7!attribute1
                        rsaux7.Close
                     Else
                        list_item.SubItems(8) = rsaux6!SHIP_TO_ORG_ID
                     End If
                     rsaux6.Close
                     rs.MoveNext
               Wend
               rs.Close
               'If lv_pedidos.ListItems.Count > 11 Then
               '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
               'Else
               '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
               'End If
               rs.Open "SELECT SUM(PIEZAS) FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
               Me.lbl_cantidad_2 = Format(rs(0).Value, "###,###,##0.00")
               rs.Close
            End If
         End If
      
         
         rs.Open "select distinct embarque as embarque from xxvia_Tb_encabezado_embarques where jaula = 3 and to_char(fecha_inicio,'DD')  = '" + CStr(var_dia_s) + "' and to_char(fecha_inicio,'MM')  = '" + CStr(var_mes_s) + "' and to_char(fecha_inicio,'yyyy')  = '" + CStr(Year(CDate(Me.txt_fecha))) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         var_Cadena_embarques = ""
         While Not rs.EOF
               If var_Cadena_embarques = "" Then
                  var_Cadena_embarques = CStr(rs!Embarque)
               Else
                  var_Cadena_embarques = var_Cadena_embarques + "," + CStr(rs!Embarque)
               End If
               rs.MoveNext
         Wend
         rs.Close
         If var_Cadena_embarques <> "" Then
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open "select distinct EMBARQUE from tb_oracle_pedidos_asignados_embarques where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  Set list_item = lv_embarques_3.ListItems.Add(, , rs!Embarque)
                  strconsulta = "select nvl(char_emb_estatus,' ') as estatus from xxvia_Tb_encabezado_embarques where embarque = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rs!Embarque)
                       .Parameters.Append parametro
                  End With
                  Set rsaux8 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  If rsaux8!estatus = " " Or rsaux8!estatus = "E" Then
                     list_item.Bold = False
                     list_item.ForeColor = &H80000012
                  Else
                     list_item.Bold = True
                     list_item.ForeColor = &H8000&
                  End If
                  'list_item.SubItems(1) = rs!NOMBRE_AGENTE
                  rs.MoveNext
            Wend
            rs.Close
            If lv_embarques_3.ListItems.Count > 0 Then
               lv_embarques_3.ListItems(1).Selected = True
               rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria  from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_3.selectedItem), cnn, adOpenDynamic, adLockOptimistic
               lv_pedidos.ListItems.Clear
               While Not rs.EOF
                     Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
                     strconsulta = "select * from oe_order_headers_all where order_number = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                          .Parameters.Append parametro
                     End With
                     Set rsaux6 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
                     'rsaux.Close
                     list_item.SubItems(2) = rs!Cliente
                     list_item.SubItems(3) = rs!PIEZAS
                     list_item.SubItems(5) = rs!orden_pedido
                     list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                     list_item.SubItems(7) = rsaux6!INVOICE_TO_ORG_ID
                     list_item.SubItems(9) = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
                     If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
                        strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                        With comandoORA
                            .ActiveConnection = cnnoracle_4
                            .CommandType = adCmdText
                            .CommandText = strconsulta
                            Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!pedido)
                            .Parameters.Append parametro
                        End With
                        Set rsaux8 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
                        rsaux8.Close
                        strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                             .Parameters.Append parametro
                        End With
                        Set rsaux7 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        list_item.SubItems(8) = rsaux7!attribute1
                        rsaux7.Close
                     Else
                        list_item.SubItems(8) = rsaux6!SHIP_TO_ORG_ID
                     End If
                     rsaux6.Close
                     rs.MoveNext
               Wend
               rs.Close
               'If lv_pedidos.ListItems.Count > 11 Then
               '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
               'Else
               '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
               'End If
               rs.Open "SELECT SUM(PIEZAS) FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
               Me.lbl_cantidad_3 = Format(rs(0).Value, "###,###,##0.00")
               rs.Close
            End If
         End If
         
         rs.Open "select distinct embarque as embarque from xxvia_Tb_encabezado_embarques where jaula = 4 and to_char(fecha_inicio,'DD')  = '" + CStr(var_dia_s) + "' and to_char(fecha_inicio,'MM')  = '" + CStr(var_mes_s) + "' and to_char(fecha_inicio,'yyyy')  = '" + CStr(Year(CDate(Me.txt_fecha))) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         var_Cadena_embarques = ""
         While Not rs.EOF
               If var_Cadena_embarques = "" Then
                  var_Cadena_embarques = CStr(rs!Embarque)
               Else
                  var_Cadena_embarques = var_Cadena_embarques + "," + CStr(rs!Embarque)
               End If
               rs.MoveNext
         Wend
         rs.Close
         If var_Cadena_embarques <> "" Then
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open "select distinct EMBARQUE from tb_oracle_pedidos_asignados_embarques where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  Set list_item = lv_embarques_4.ListItems.Add(, , rs!Embarque)
                  strconsulta = "select nvl(char_emb_estatus,' ') as estatus from xxvia_Tb_encabezado_embarques where embarque = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rs!Embarque)
                       .Parameters.Append parametro
                  End With
                  Set rsaux8 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  If rsaux8!estatus = " " Or rsaux8!estatus = "E" Then
                     list_item.Bold = False
                     list_item.ForeColor = &H80000012
                  Else
                     list_item.Bold = True
                     list_item.ForeColor = &H8000&
                  End If
                  'list_item.SubItems(1) = rs!NOMBRE_AGENTE
                  rs.MoveNext
            Wend
            rs.Close
            If lv_embarques_4.ListItems.Count > 0 Then
               lv_embarques_4.ListItems(1).Selected = True
               rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria  from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_4.selectedItem), cnn, adOpenDynamic, adLockOptimistic
               lv_pedidos.ListItems.Clear
               While Not rs.EOF
                     Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
                     strconsulta = "select * from oe_order_headers_all where order_number = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                          .Parameters.Append parametro
                     End With
                     Set rsaux6 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
                     'rsaux.Close
                     list_item.SubItems(2) = rs!Cliente
                     list_item.SubItems(3) = rs!PIEZAS
                     list_item.SubItems(5) = rs!orden_pedido
                     list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                     list_item.SubItems(7) = rsaux6!INVOICE_TO_ORG_ID
                     list_item.SubItems(9) = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
                     If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
                        strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                        With comandoORA
                            .ActiveConnection = cnnoracle_4
                            .CommandType = adCmdText
                            .CommandText = strconsulta
                            Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!pedido)
                            .Parameters.Append parametro
                        End With
                        Set rsaux8 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
                        rsaux8.Close
                        strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                             .Parameters.Append parametro
                        End With
                        Set rsaux7 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        list_item.SubItems(8) = rsaux7!attribute1
                        rsaux7.Close
                     Else
                        list_item.SubItems(8) = rsaux6!SHIP_TO_ORG_ID
                     End If
                     rsaux6.Close
                     rs.MoveNext
               Wend
               rs.Close
               'If lv_pedidos.ListItems.Count > 11 Then
               '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
               'Else
               '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
               'End If
               rs.Open "SELECT SUM(PIEZAS) FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
               Me.lbl_cantidad_4 = Format(rs(0).Value, "###,###,##0.00")
               rs.Close
            End If
         End If
   
   
         rs.Open "select distinct embarque as embarque from xxvia_Tb_encabezado_embarques where jaula = 5 and to_char(fecha_inicio,'DD')  = '" + CStr(var_dia_s) + "' and to_char(fecha_inicio,'MM')  = '" + CStr(var_mes_s) + "' and to_char(fecha_inicio,'yyyy')  = '" + CStr(Year(CDate(Me.txt_fecha))) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         var_Cadena_embarques = ""
         While Not rs.EOF
               If var_Cadena_embarques = "" Then
                  var_Cadena_embarques = CStr(rs!Embarque)
               Else
                  var_Cadena_embarques = var_Cadena_embarques + "," + CStr(rs!Embarque)
               End If
               rs.MoveNext
         Wend
         rs.Close
         If var_Cadena_embarques <> "" Then
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open "select distinct EMBARQUE from tb_oracle_pedidos_asignados_embarques where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  Set list_item = lv_embarques_5.ListItems.Add(, , rs!Embarque)
                  strconsulta = "select nvl(char_emb_estatus,' ') as estatus from xxvia_Tb_encabezado_embarques where embarque = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rs!Embarque)
                       .Parameters.Append parametro
                  End With
                  Set rsaux8 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  If rsaux8!estatus = " " Or rsaux8!estatus = "E" Then
                     list_item.Bold = False
                     list_item.ForeColor = &H80000012
                  Else
                     list_item.Bold = True
                     list_item.ForeColor = &H8000&
                  End If
                  'list_item.SubItems(1) = rs!NOMBRE_AGENTE
                  rs.MoveNext
            Wend
            rs.Close
            If lv_embarques_5.ListItems.Count > 0 Then
               lv_embarques_5.ListItems(1).Selected = True
               rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria  from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_5.selectedItem), cnn, adOpenDynamic, adLockOptimistic
               lv_pedidos.ListItems.Clear
               While Not rs.EOF
                     Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
                     strconsulta = "select * from oe_order_headers_all where order_number = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                          .Parameters.Append parametro
                     End With
                     Set rsaux6 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
                     'rsaux.Close
                     list_item.SubItems(2) = rs!Cliente
                     list_item.SubItems(3) = rs!PIEZAS
                     list_item.SubItems(5) = rs!orden_pedido
                     list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                     list_item.SubItems(7) = rsaux6!INVOICE_TO_ORG_ID
                     list_item.SubItems(9) = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
                     If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
                        strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                        With comandoORA
                            .ActiveConnection = cnnoracle_4
                            .CommandType = adCmdText
                            .CommandText = strconsulta
                            Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!pedido)
                            .Parameters.Append parametro
                        End With
                        Set rsaux8 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
                        rsaux8.Close
                        strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                             .Parameters.Append parametro
                        End With
                        Set rsaux7 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        list_item.SubItems(8) = rsaux7!attribute1
                        rsaux7.Close
                     Else
                        list_item.SubItems(8) = rsaux6!SHIP_TO_ORG_ID
                     End If
                     rsaux6.Close
                     rs.MoveNext
               Wend
               rs.Close
               'If lv_pedidos.ListItems.Count > 11 Then
               '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
               'Else
               '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
               'End If
               rs.Open "SELECT SUM(PIEZAS) FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
               Me.lbl_cantidad_5 = Format(rs(0).Value, "###,###,##0.00")
               rs.Close
            End If
         End If
   
         rs.Open "select distinct embarque as embarque from xxvia_Tb_encabezado_embarques where jaula = 6 and to_char(fecha_inicio,'DD')  = '" + CStr(var_dia_s) + "' and to_char(fecha_inicio,'MM')  = '" + CStr(var_mes_s) + "' and to_char(fecha_inicio,'yyyy')  = '" + CStr(Year(CDate(Me.txt_fecha))) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         var_Cadena_embarques = ""
         While Not rs.EOF
               If var_Cadena_embarques = "" Then
                  var_Cadena_embarques = CStr(rs!Embarque)
               Else
                  var_Cadena_embarques = var_Cadena_embarques + "," + CStr(rs!Embarque)
               End If
               rs.MoveNext
          Wend
         rs.Close
         If var_Cadena_embarques <> "" Then
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open "select distinct EMBARQUE from tb_oracle_pedidos_asignados_embarques where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  Set list_item = lv_embarques_6.ListItems.Add(, , rs!Embarque)
                  'list_item.SubItems(1) = rs!NOMBRE_AGENTE
                  strconsulta = "select nvl(char_emb_estatus,' ') as estatus from xxvia_Tb_encabezado_embarques where embarque = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rs!Embarque)
                       .Parameters.Append parametro
                  End With
                  Set rsaux8 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  If rsaux8!estatus = " " Or rsaux8!estatus = "E" Then
                     list_item.Bold = False
                     list_item.ForeColor = &H80000012
                  Else
                     list_item.Bold = True
                     list_item.ForeColor = &H8000&
                  End If
                  rs.MoveNext
            Wend
            rs.Close
            If lv_embarques_6.ListItems.Count > 0 Then
               lv_embarques_6.ListItems(1).Selected = True
               rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria  from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_6.selectedItem), cnn, adOpenDynamic, adLockOptimistic
               lv_pedidos.ListItems.Clear
               While Not rs.EOF
                     Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
                     strconsulta = "select * from oe_order_headers_all where order_number = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                          .Parameters.Append parametro
                     End With
                     Set rsaux6 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
                     'rsaux.Close
                     list_item.SubItems(2) = rs!Cliente
                     list_item.SubItems(3) = rs!PIEZAS
                     list_item.SubItems(5) = rs!orden_pedido
                     list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                     list_item.SubItems(7) = rsaux6!INVOICE_TO_ORG_ID
                     list_item.SubItems(9) = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
                     If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
                        strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                        With comandoORA
                            .ActiveConnection = cnnoracle_4
                            .CommandType = adCmdText
                            .CommandText = strconsulta
                            Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!pedido)
                            .Parameters.Append parametro
                        End With
                        Set rsaux8 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
                        rsaux8.Close
                        strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                             .Parameters.Append parametro
                        End With
                        Set rsaux7 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        list_item.SubItems(8) = rsaux7!attribute1
                        rsaux7.Close
                     Else
                        list_item.SubItems(8) = rsaux6!SHIP_TO_ORG_ID
                     End If
                     rsaux6.Close
                     rs.MoveNext
               Wend
               rs.Close
               'If lv_pedidos.ListItems.Count > 11 Then
               '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
               'Else
               '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
               'End If
               rs.Open "SELECT SUM(PIEZAS) FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
               Me.lbl_cantidad_6 = Format(rs(0).Value, "###,###,##0.00")
               rs.Close
            End If
         End If
      
         rs.Open "select distinct embarque as embarque from xxvia_Tb_encabezado_embarques where jaula = 7 and to_char(fecha_inicio,'DD')  = '" + CStr(var_dia_s) + "' and to_char(fecha_inicio,'MM')  = '" + CStr(var_mes_s) + "' and to_char(fecha_inicio,'yyyy')  = '" + CStr(Year(CDate(Me.txt_fecha))) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         var_Cadena_embarques = ""
         While Not rs.EOF
               If var_Cadena_embarques = "" Then
                  var_Cadena_embarques = CStr(rs!Embarque)
               Else
                  var_Cadena_embarques = var_Cadena_embarques + "," + CStr(rs!Embarque)
               End If
               rs.MoveNext
         Wend
         rs.Close
         If var_Cadena_embarques <> "" Then
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open "select distinct EMBARQUE from tb_oracle_pedidos_asignados_embarques where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  Set list_item = lv_embarques_7.ListItems.Add(, , rs!Embarque)
                  strconsulta = "select nvl(char_emb_estatus,' ') as estatus from xxvia_Tb_encabezado_embarques where embarque = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rs!Embarque)
                       .Parameters.Append parametro
                  End With
                  Set rsaux8 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  If rsaux8!estatus = " " Or rsaux8!estatus = "E" Then
                     list_item.Bold = False
                     list_item.ForeColor = &H80000012
                  Else
                     list_item.Bold = True
                     list_item.ForeColor = &H8000&
                  End If
                  'list_item.SubItems(1) = rs!NOMBRE_AGENTE
                  rs.MoveNext
            Wend
            rs.Close
            If lv_embarques_7.ListItems.Count > 0 Then
               lv_embarques_7.ListItems(1).Selected = True
               rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria  from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_7.selectedItem), cnn, adOpenDynamic, adLockOptimistic
               lv_pedidos.ListItems.Clear
               While Not rs.EOF
                     Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
                     strconsulta = "select * from oe_order_headers_all where order_number = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                          .Parameters.Append parametro
                     End With
                     Set rsaux6 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
                     'rsaux.Close
                     list_item.SubItems(2) = rs!Cliente
                     list_item.SubItems(3) = rs!PIEZAS
                     list_item.SubItems(5) = rs!orden_pedido
                     list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                     list_item.SubItems(7) = rsaux6!INVOICE_TO_ORG_ID
                     list_item.SubItems(9) = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
                     If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
                        strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                        With comandoORA
                            .ActiveConnection = cnnoracle_4
                            .CommandType = adCmdText
                            .CommandText = strconsulta
                            Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!pedido)
                            .Parameters.Append parametro
                        End With
                        Set rsaux8 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
                        rsaux8.Close
                        strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                             .Parameters.Append parametro
                        End With
                        Set rsaux7 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        list_item.SubItems(8) = rsaux7!attribute1
                        rsaux7.Close
                     Else
                        list_item.SubItems(8) = rsaux6!SHIP_TO_ORG_ID
                     End If
                     rsaux6.Close
                     rs.MoveNext
               Wend
               rs.Close
               'If lv_pedidos.ListItems.Count > 11 Then
               '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
               'Else
               '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
               'End If
               rs.Open "SELECT SUM(PIEZAS) FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
               Me.lbl_cantidad_7 = Format(rs(0).Value, "###,###,##0.00")
               rs.Close
            End If
         End If
      
      
         
         rs.Open "select distinct embarque as embarque from xxvia_Tb_encabezado_embarques where jaula = 8 and to_char(fecha_inicio,'DD')  = '" + CStr(var_dia_s) + "' and to_char(fecha_inicio,'MM')  = '" + CStr(var_mes_s) + "' and to_char(fecha_inicio,'yyyy')  = '" + CStr(Year(CDate(Me.txt_fecha))) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         var_Cadena_embarques = ""
         While Not rs.EOF
               If var_Cadena_embarques = "" Then
                  var_Cadena_embarques = CStr(rs!Embarque)
               Else
                  var_Cadena_embarques = var_Cadena_embarques + "," + CStr(rs!Embarque)
               End If
               rs.MoveNext
         Wend
         rs.Close
         If var_Cadena_embarques <> "" Then
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open "select distinct EMBARQUE from tb_oracle_pedidos_asignados_embarques where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  Set list_item = lv_embarques_8.ListItems.Add(, , rs!Embarque)
                  strconsulta = "select nvl(char_emb_estatus,' ') as estatus from xxvia_Tb_encabezado_embarques where embarque = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rs!Embarque)
                       .Parameters.Append parametro
                  End With
                  Set rsaux8 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  If rsaux8!estatus = " " Or rsaux8!estatus = "E" Then
                     list_item.Bold = False
                     list_item.ForeColor = &H80000012
                  Else
                     list_item.Bold = True
                     list_item.ForeColor = &H8000&
                  End If
                  'list_item.SubItems(1) = rs!NOMBRE_AGENTE
                  rs.MoveNext
            Wend
            rs.Close
            If lv_embarques_8.ListItems.Count > 0 Then
               lv_embarques_8.ListItems(1).Selected = True
               rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria  from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_8.selectedItem), cnn, adOpenDynamic, adLockOptimistic
               lv_pedidos.ListItems.Clear
               While Not rs.EOF
                     Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
                     strconsulta = "select * from oe_order_headers_all where order_number = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                          .Parameters.Append parametro
                     End With
                     Set rsaux6 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
                     'rsaux.Close
                     list_item.SubItems(2) = rs!Cliente
                     list_item.SubItems(3) = rs!PIEZAS
                     list_item.SubItems(5) = rs!orden_pedido
                     list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                     list_item.SubItems(7) = rsaux6!INVOICE_TO_ORG_ID
                     list_item.SubItems(9) = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
                     If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
                        strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                        With comandoORA
                            .ActiveConnection = cnnoracle_4
                            .CommandType = adCmdText
                            .CommandText = strconsulta
                            Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!pedido)
                            .Parameters.Append parametro
                        End With
                        Set rsaux8 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
                        rsaux8.Close
                        strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                             .Parameters.Append parametro
                        End With
                        Set rsaux7 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        list_item.SubItems(8) = rsaux7!attribute1
                        rsaux7.Close
                     Else
                        list_item.SubItems(8) = rsaux6!SHIP_TO_ORG_ID
                     End If
                     rsaux6.Close
                     rs.MoveNext
               Wend
               rs.Close
               'If lv_pedidos.ListItems.Count > 11 Then
               '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
               'Else
               '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
               'End If
               rs.Open "SELECT SUM(PIEZAS) FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
               Me.lbl_cantidad_8 = Format(rs(0).Value, "###,###,##0.00")
               rs.Close
            End If
         End If
      
         rs.Open "select distinct embarque as embarque from xxvia_Tb_encabezado_embarques where jaula = 9 and to_char(fecha_inicio,'DD')  = '" + CStr(var_dia_s) + "' and to_char(fecha_inicio,'MM')  = '" + CStr(var_mes_s) + "' and to_char(fecha_inicio,'yyyy')  = '" + CStr(Year(CDate(Me.txt_fecha))) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         var_Cadena_embarques = ""
         While Not rs.EOF
               If var_Cadena_embarques = "" Then
                  var_Cadena_embarques = CStr(rs!Embarque)
               Else
                  var_Cadena_embarques = var_Cadena_embarques + "," + CStr(rs!Embarque)
               End If
               rs.MoveNext
         Wend
         rs.Close
         If var_Cadena_embarques <> "" Then
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open "select distinct EMBARQUE from tb_oracle_pedidos_asignados_embarques where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  Set list_item = lv_embarques_9.ListItems.Add(, , rs!Embarque)
                  strconsulta = "select nvl(char_emb_estatus,' ') as estatus from xxvia_Tb_encabezado_embarques where embarque = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rs!Embarque)
                       .Parameters.Append parametro
                  End With
                  Set rsaux8 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  If rsaux8!estatus = " " Or rsaux8!estatus = "E" Then
                     list_item.Bold = False
                     list_item.ForeColor = &H80000012
                  Else
                     list_item.Bold = True
                     list_item.ForeColor = &H8000&
                  End If
                  'list_item.SubItems(1) = rs!NOMBRE_AGENTE
                  rs.MoveNext
            Wend
            rs.Close
            If lv_embarques_9.ListItems.Count > 0 Then
               lv_embarques_9.ListItems(1).Selected = True
               rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria  from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_9.selectedItem), cnn, adOpenDynamic, adLockOptimistic
               lv_pedidos.ListItems.Clear
               While Not rs.EOF
                     Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
                     strconsulta = "select * from oe_order_headers_all where order_number = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                          .Parameters.Append parametro
                     End With
                     Set rsaux6 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
                     'rsaux.Close
                     list_item.SubItems(2) = rs!Cliente
                     list_item.SubItems(3) = rs!PIEZAS
                     list_item.SubItems(5) = rs!orden_pedido
                     list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                     list_item.SubItems(7) = rsaux6!INVOICE_TO_ORG_ID
                     list_item.SubItems(9) = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
                     If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
                        strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                        With comandoORA
                            .ActiveConnection = cnnoracle_4
                            .CommandType = adCmdText
                            .CommandText = strconsulta
                            Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!pedido)
                            .Parameters.Append parametro
                        End With
                        Set rsaux8 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
                        rsaux8.Close
                        strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                             .Parameters.Append parametro
                        End With
                        Set rsaux7 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        list_item.SubItems(8) = rsaux7!attribute1
                        rsaux7.Close
                     Else
                        list_item.SubItems(8) = rsaux6!SHIP_TO_ORG_ID
                     End If
                     rsaux6.Close
                     rs.MoveNext
               Wend
               rs.Close
               'If lv_pedidos.ListItems.Count > 11 Then
               '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
               'Else
               '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
               'End If
               rs.Open "SELECT SUM(PIEZAS) FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
               Me.lbl_cantidad_9 = Format(rs(0).Value, "###,###,##0.00")
               rs.Close
            End If
         End If
   
     
         rs.Open "select distinct embarque as embarque from xxvia_Tb_encabezado_embarques where jaula = 10 and to_char(fecha_inicio,'DD')  = '" + CStr(var_dia_s) + "' and to_char(fecha_inicio,'MM')  = '" + CStr(var_mes_s) + "' and to_char(fecha_inicio,'yyyy')  = '" + CStr(Year(CDate(Me.txt_fecha))) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         var_Cadena_embarques = ""
         While Not rs.EOF
               If var_Cadena_embarques = "" Then
                  var_Cadena_embarques = CStr(rs!Embarque)
               Else
                  var_Cadena_embarques = var_Cadena_embarques + "," + CStr(rs!Embarque)
               End If
               rs.MoveNext
         Wend
         rs.Close
         If var_Cadena_embarques <> "" Then
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open "select distinct EMBARQUE from tb_oracle_pedidos_asignados_embarques where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  Set list_item = lv_embarques_10.ListItems.Add(, , rs!Embarque)
                  strconsulta = "select nvl(char_emb_estatus,' ') as estatus from xxvia_Tb_encabezado_embarques where embarque = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rs!Embarque)
                       .Parameters.Append parametro
                  End With
                  Set rsaux8 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  If rsaux8!estatus = " " Or rsaux8!estatus = "E" Then
                     list_item.Bold = False
                     list_item.ForeColor = &H80000012
                  Else
                     list_item.Bold = True
                     list_item.ForeColor = &H8000&
                  End If
                  'list_item.SubItems(1) = rs!NOMBRE_AGENTE
                  rs.MoveNext
            Wend
            rs.Close
            If lv_embarques_10.ListItems.Count > 0 Then
               lv_embarques_10.ListItems(1).Selected = True
               rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, paqueteria  from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(Me.lv_embarques_10.selectedItem), cnn, adOpenDynamic, adLockOptimistic
               lv_pedidos.ListItems.Clear
               While Not rs.EOF
                     Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
                     strconsulta = "select * from oe_order_headers_all where order_number = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                          .Parameters.Append parametro
                     End With
                     Set rsaux6 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     'rsaux.Open "select * from ar_collectors where collector_id = " + CStr(rs!Agente), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
                     'rsaux.Close
                     list_item.SubItems(2) = rs!Cliente
                     list_item.SubItems(3) = rs!PIEZAS
                     list_item.SubItems(5) = rs!orden_pedido
                     list_item.SubItems(6) = Format(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN), "###,###,##0.00")
                     list_item.SubItems(7) = rsaux6!INVOICE_TO_ORG_ID
                     list_item.SubItems(9) = IIf(IsNull(rs!paqueteria), 0, rs!paqueteria)
                     If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
                        strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                        With comandoORA
                            .ActiveConnection = cnnoracle_4
                            .CommandType = adCmdText
                            .CommandText = strconsulta
                            Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!pedido)
                            .Parameters.Append parametro
                        End With
                        Set rsaux8 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
                        rsaux8.Close
                        strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                             .Parameters.Append parametro
                        End With
                        Set rsaux7 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        list_item.SubItems(8) = rsaux7!attribute1
                        rsaux7.Close
                     Else
                        list_item.SubItems(8) = rsaux6!SHIP_TO_ORG_ID
                     End If
                     rsaux6.Close
                     rs.MoveNext
               Wend
               rs.Close
               'If lv_pedidos.ListItems.Count > 11 Then
               '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5000.22
               'Else
               '   Me.lv_pedidos.ColumnHeaders.Item(2).Width = 5300.22
               'End If
               rs.Open "SELECT SUM(PIEZAS) FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
               Me.lbl_cantidad_10 = Format(rs(0).Value, "###,###,##0.00")
               rs.Close
            End If
         End If
         
         If Me.lv_embarques_1.ListItems.Count > 0 Then
            Me.lv_embarques_1.SetFocus
         Else
            If Me.lv_embarques_2.ListItems.Count > 0 Then
               Me.lv_embarques_2.SetFocus
            Else
               If Me.lv_embarques_3.ListItems.Count > 0 Then
                  Me.lv_embarques_3.SetFocus
               Else
                  If Me.lv_embarques_4.ListItems.Count > 0 Then
                     Me.lv_embarques_4.SetFocus
                  Else
                     If Me.lv_embarques_5.ListItems.Count > 0 Then
                        Me.lv_embarques_5.SetFocus
                     Else
                        If Me.lv_embarques_6.ListItems.Count > 0 Then
                           Me.lv_embarques_6.SetFocus
                        Else
                           If Me.lv_embarques_7.ListItems.Count > 0 Then
                              Me.lv_embarques_7.SetFocus
                           Else
                              If Me.lv_embarques_8.ListItems.Count > 0 Then
                                 Me.lv_embarques_8.SetFocus
                              Else
                                 If Me.lv_embarques_9.ListItems.Count > 0 Then
                                    Me.lv_embarques_9.SetFocus
                                 Else
                                    If Me.lv_embarques_10.ListItems.Count > 0 Then
                                       Me.lv_embarques_10.SetFocus
                                    End If
                                 End If
                              End If
                           End If
                        End If
                     End If
                  End If
               End If
            End If
         End If
         
         
      Else
         MsgBox "Fecha incorrecta", vbOKOnly, "ATENCION"
      End If
      Call ilumina_grid
   End If
End Sub

Private Sub txt_fecha_LostFocus()
   If Not IsDate(Me.txt_fecha) Then
      MsgBox "Fecha incorrecta", vbOKOnly, "ATENCION"
      Me.txt_fecha = Date
   End If
End Sub

Private Sub txt_porcentaje_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_total_volumen_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_transporte_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_volumen_unidad_Change()
    If Not IsNumeric(Me.txt_volumen_unidad) Then
       Me.txt_volumen_unidad = 0
    End If
    If Not IsNumeric(Me.txt_total_volumen) Then
       Me.txt_total_volumen = 0
    End If
    If Me.txt_volumen_unidad > 0 Then
       Me.txt_porcentaje = Format((CDbl(Me.txt_total_volumen) / CDbl(Me.txt_volumen_unidad)) * 100, "###,##0.000")
    Else
       Me.txt_porcentaje = 0
    End If
End Sub

Private Sub txt_volumen_unidad_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub
