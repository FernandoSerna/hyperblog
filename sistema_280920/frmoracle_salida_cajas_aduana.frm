VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmoracle_salida_cajas_aduana 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Salida cajas"
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.Frame Frame9 
      Caption         =   "Frame9"
      Height          =   2055
      Left            =   4680
      TabIndex        =   44
      Top             =   5400
      Visible         =   0   'False
      Width           =   6855
      Begin VB.TextBox txt_volumen_carga 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txt_porcentaje_carga 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txt_volumen_carga_lectores 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txt_porcentaje_carga_lectores 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txt_clave_unidad 
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txt_nombre_unidad 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   360
         Width           =   5535
      End
      Begin VB.TextBox txt_volumen_unidad 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.Frame frm_sellos 
      Height          =   2460
      Left            =   1200
      TabIndex        =   21
      Top             =   240
      Width           =   4005
      Begin VB.CommandButton cmd_cerrar_embarque_nuevo_metodo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   690
         Picture         =   "frmoracle_salida_cajas_aduana.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Cerrar Alt + C"
         Top             =   330
         Width           =   330
      End
      Begin VB.CommandButton cmd_cancelar_sello 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   375
         Picture         =   "frmoracle_salida_cajas_aduana.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Cancelar Alt + C"
         Top             =   330
         Width           =   330
      End
      Begin VB.CommandButton cmd_aceptar_sello 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   45
         Picture         =   "frmoracle_salida_cajas_aduana.frx":024C
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Aceptar Alt + A"
         Top             =   330
         Width           =   330
      End
      Begin VB.TextBox txt_sello 
         Height          =   315
         Left            =   585
         TabIndex        =   24
         Top             =   795
         Width           =   3345
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   690
         Picture         =   "frmoracle_salida_cajas_aduana.frx":0396
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Cerrar Alt + C"
         Top             =   330
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Frame Frame5 
         Height          =   75
         Left            =   30
         TabIndex        =   22
         Top             =   645
         Width           =   3930
      End
      Begin MSComctlLib.ListView lv_sellos 
         Height          =   1200
         Left            =   30
         TabIndex        =   27
         Top             =   1110
         Width           =   3915
         _ExtentX        =   6906
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Número de Sello"
            Object.Width           =   5115
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   " Sellos"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   7
         Left            =   45
         TabIndex        =   29
         Top             =   135
         Width           =   3930
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sello:"
         Height          =   195
         Left            =   90
         TabIndex        =   28
         Top             =   840
         Width           =   390
      End
   End
   Begin VB.Frame Frame1 
      Height          =   810
      Left            =   120
      TabIndex        =   40
      Top             =   2160
      Width           =   15030
      Begin ComctlLib.ProgressBar ProgressBar2 
         Height          =   255
         Left            =   1080
         TabIndex        =   41
         Top             =   120
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   8520
         TabIndex        =   53
         Top             =   120
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lbl_porcentaje_aduana 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11220
         TabIndex        =   54
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lbl_porcentaje_lectores 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3780
         TabIndex        =   52
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "% Aduana:"
         Height          =   195
         Left            =   7680
         TabIndex        =   43
         Top             =   120
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "% Lectores:"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   120
         Width           =   825
      End
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2160
      Picture         =   "frmoracle_salida_cajas_aduana.frx":0498
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Actualizar "
      Top             =   0
      Width           =   375
   End
   Begin VB.Frame Frame8 
      Height          =   810
      Left            =   120
      TabIndex        =   35
      Top             =   1350
      Width           =   15030
      Begin VB.Label lbl_tipo_bulto 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   45
         TabIndex        =   36
         Top             =   120
         Width           =   14895
      End
   End
   Begin VB.CommandButton cmd_mensaje_4 
      Caption         =   "mensaje 4"
      Height          =   195
      Left            =   2910
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.CommandButton cmd_mensaje_2 
      Caption         =   "mensaje 2"
      Height          =   195
      Left            =   2745
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "Detener"
      Height          =   375
      Left            =   840
      Picture         =   "frmoracle_salida_cajas_aduana.frx":059A
      TabIndex        =   20
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmd_cerrar_pedido 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   480
      Picture         =   "frmoracle_salida_cajas_aduana.frx":069C
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Cerrar Alt + C"
      Top             =   0
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   12315
      Top             =   -45
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   14760
      Picture         =   "frmoracle_salida_cajas_aduana.frx":079E
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame7 
      Height          =   960
      Left            =   11280
      TabIndex        =   11
      Top             =   390
      Width           =   3855
      Begin VB.TextBox txt_cantidad 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   3930
         TabIndex        =   14
         Top             =   165
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lbl_semaforo 
         BorderStyle     =   1  'Fixed Single
         Height          =   630
         Left            =   1185
         TabIndex        =   31
         Top             =   210
         Width           =   2520
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Semáforo:"
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
         Left            =   105
         TabIndex        =   30
         Top             =   450
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad leida:"
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
         Left            =   105
         TabIndex        =   12
         Top             =   450
         Visible         =   0   'False
         Width           =   1500
      End
   End
   Begin VB.Frame Frame6 
      Height          =   960
      Left            =   7380
      TabIndex        =   9
      Top             =   390
      Width           =   3855
      Begin VB.TextBox txt_embarque 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   1770
         TabIndex        =   13
         Top             =   165
         Width           =   2010
      End
      Begin VB.Label Label3 
         Caption         =   "Embarque:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1605
      End
   End
   Begin VB.CommandButton cmd_cerrar_embarque 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1800
      Picture         =   "frmoracle_salida_cajas_aduana.frx":0DD8
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Cerrar Embarque"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   120
      Picture         =   "frmoracle_salida_cajas_aduana.frx":0EDA
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Actualizar"
      Top             =   0
      Width           =   375
   End
   Begin VB.Frame Frame4 
      Height          =   60
      Left            =   60
      TabIndex        =   5
      Top             =   330
      Width           =   15105
   End
   Begin VB.Frame Frame3 
      Height          =   2760
      Left            =   105
      TabIndex        =   2
      Top             =   6345
      Width           =   15090
      Begin MSComctlLib.ListView lv_cajas_siguientes 
         Height          =   2175
         Left            =   45
         TabIndex        =   17
         Top             =   540
         Width           =   14880
         _ExtentX        =   26247
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
            Text            =   "          Código"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Pedido"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Agente"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cliente"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Cantidad"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "estatus"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Caja"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Tipo empaque"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Caja"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Sello"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   "  Pedidos pendientes"
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   1
         Left            =   30
         TabIndex        =   19
         Top             =   165
         Width           =   15000
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3330
      Left            =   105
      TabIndex        =   1
      Top             =   3000
      Width           =   15060
      Begin MSComctlLib.ListView lv_cajas 
         Height          =   2700
         Left            =   0
         TabIndex        =   4
         Top             =   480
         Width           =   14880
         _ExtentX        =   26247
         _ExtentY        =   4763
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
         NumItems        =   14
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "          Código"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Pedido"
            Object.Width           =   1605
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Agente"
            Object.Width           =   2434
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cliente"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Cantidad"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "estatus"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Caja"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Tipo empaque"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Caja"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Sello"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Guia"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "C. Nueva"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Marca cambiar embarque"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Embarque Nuevo"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   "  Pedido actual"
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   0
         Left            =   30
         TabIndex        =   18
         Top             =   135
         Width           =   14970
      End
   End
   Begin VB.Frame frm_codigo 
      Height          =   960
      Left            =   105
      TabIndex        =   0
      Top             =   390
      Width           =   7245
      Begin VB.TextBox txt_codigo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2820
         TabIndex        =   3
         Top             =   180
         Width           =   4305
      End
      Begin VB.Label lbl_tipo_codigo 
         AutoSize        =   -1  'True
         Caption         =   "Código de la Caja:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   105
         TabIndex        =   8
         Top             =   330
         Width           =   2595
      End
   End
   Begin VB.CommandButton cmd_cambiar_embarque 
      Appearance      =   0  'Flat
      Caption         =   "Cambiar embarque"
      Height          =   375
      Left            =   2520
      Picture         =   "frmoracle_salida_cajas_aduana.frx":0FDC
      TabIndex        =   56
      ToolTipText     =   "Actualizar "
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label lbl_transporte 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   55
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label lbl_paqueteria 
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
      Left            =   6480
      TabIndex        =   39
      Top             =   0
      Width           =   1935
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp4 
      Height          =   135
      Left            =   5340
      TabIndex        =   34
      Top             =   120
      Visible         =   0   'False
      Width           =   630
      URL             =   "C:\sistemas\desarrollo\integral\type.wma"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   0   'False
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   100
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   1111
      _cy             =   238
   End
End
Attribute VB_Name = "frmoracle_salida_cajas_aduana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter
Dim objConn As New ADODB.Connection
Dim objCmd As New ADODB.Command
Dim objParm As ADODB.Parameter




Private Sub ilumina_grid()
    var_n = lv_cajas.ListItems.Count
    For var_i = 1 To var_n
        lv_cajas.ListItems.Item(var_i).Selected = True
        If Trim(lv_cajas.selectedItem.SubItems(5)) = "L" Then
           lv_cajas.ListItems.Item(var_i).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(1).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(2).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(3).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(4).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(5).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(6).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(7).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(8).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(9).Bold = False
           lv_cajas.ListItems.Item(var_i).ForeColor = &HFF&
           lv_cajas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &HFF&
           lv_cajas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &HFF&
           lv_cajas.ListItems.Item(var_i).ListSubItems(3).ForeColor = &HFF&
           lv_cajas.ListItems.Item(var_i).ListSubItems(4).ForeColor = &HFF&
           lv_cajas.ListItems.Item(var_i).ListSubItems(5).ForeColor = &HFF&
           lv_cajas.ListItems.Item(var_i).ListSubItems(6).ForeColor = &HFF&
           lv_cajas.ListItems.Item(var_i).ListSubItems(7).ForeColor = &HFF&
           lv_cajas.ListItems.Item(var_i).ListSubItems(8).ForeColor = &HFF&
           lv_cajas.ListItems.Item(var_i).ListSubItems(9).ForeColor = &HFF&
           If Me.lv_cajas.ListItems.Item(var_i).ListSubItems(11) <> "" Then
           
           
              lv_cajas.ListItems.Item(var_i).Bold = True
              lv_cajas.ListItems.Item(var_i).ListSubItems(1).Bold = True
              lv_cajas.ListItems.Item(var_i).ListSubItems(2).Bold = True
              lv_cajas.ListItems.Item(var_i).ListSubItems(3).Bold = True
              lv_cajas.ListItems.Item(var_i).ListSubItems(4).Bold = True
              lv_cajas.ListItems.Item(var_i).ListSubItems(5).Bold = True
              lv_cajas.ListItems.Item(var_i).ListSubItems(6).Bold = True
              lv_cajas.ListItems.Item(var_i).ListSubItems(7).Bold = True
              lv_cajas.ListItems.Item(var_i).ListSubItems(8).Bold = True
              lv_cajas.ListItems.Item(var_i).ListSubItems(9).Bold = True
              lv_cajas.ListItems.Item(var_i).ListSubItems(10).Bold = True
              lv_cajas.ListItems.Item(var_i).ListSubItems(11).Bold = True
              lv_cajas.ListItems.Item(var_i).ForeColor = &H8000000D
              lv_cajas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H8000000D
              lv_cajas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H8000000D
              lv_cajas.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H8000000D
              lv_cajas.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H8000000D
              lv_cajas.ListItems.Item(var_i).ListSubItems(5).ForeColor = &H8000000D
              lv_cajas.ListItems.Item(var_i).ListSubItems(6).ForeColor = &H8000000D
              lv_cajas.ListItems.Item(var_i).ListSubItems(7).ForeColor = &H8000000D
              lv_cajas.ListItems.Item(var_i).ListSubItems(8).ForeColor = &H8000000D
              lv_cajas.ListItems.Item(var_i).ListSubItems(9).ForeColor = &H8000000D
              lv_cajas.ListItems.Item(var_i).ListSubItems(10).ForeColor = &H8000000D
              lv_cajas.ListItems.Item(var_i).ListSubItems(11).ForeColor = &H8000000D
           
           
           
           End If
        Else
           lv_cajas.ListItems.Item(var_i).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(1).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(2).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(3).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(4).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(5).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(6).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(7).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(8).Bold = False
           lv_cajas.ListItems.Item(var_i).ListSubItems(9).Bold = False
           lv_cajas.ListItems.Item(var_i).ForeColor = &H80000008
           lv_cajas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H80000008
           lv_cajas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H80000008
           lv_cajas.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H80000008
           lv_cajas.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H80000008
           lv_cajas.ListItems.Item(var_i).ListSubItems(5).ForeColor = &H80000008
           lv_cajas.ListItems.Item(var_i).ListSubItems(6).ForeColor = &H80000008
           lv_cajas.ListItems.Item(var_i).ListSubItems(7).ForeColor = &H80000008
           lv_cajas.ListItems.Item(var_i).ListSubItems(8).ForeColor = &H80000008
           lv_cajas.ListItems.Item(var_i).ListSubItems(9).ForeColor = &H80000008
        End If
        If Trim(lv_cajas.selectedItem.SubItems(12)) = "*" Then
        
         lv_cajas.ListItems.Item(var_i).Bold = True
         rs.Open "UPDATE tb_oracle_cajas_aduana SET MARCA_CAMBIO_EMBARQUE = '*' WHERE EMBARQUE = '" + Me.txt_embarque + "' AND CAJA = '" + Me.lv_cajas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
         
         lv_cajas.ListItems.Item(var_i).ListSubItems(1).Bold = True
         lv_cajas.ListItems.Item(var_i).ListSubItems(2).Bold = True
         lv_cajas.ListItems.Item(var_i).ListSubItems(3).Bold = True
         lv_cajas.ListItems.Item(var_i).ListSubItems(4).Bold = True
         lv_cajas.ListItems.Item(var_i).ListSubItems(5).Bold = True
         lv_cajas.ListItems.Item(var_i).ListSubItems(6).Bold = True
         lv_cajas.ListItems.Item(var_i).ListSubItems(7).Bold = True
         lv_cajas.ListItems.Item(var_i).ListSubItems(8).Bold = True
         lv_cajas.ListItems.Item(var_i).ListSubItems(9).Bold = True
         lv_cajas.ListItems.Item(var_i).ForeColor = &HC000&
         lv_cajas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &HC000&
         lv_cajas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &HC000&
         lv_cajas.ListItems.Item(var_i).ListSubItems(3).ForeColor = &HC000&
         lv_cajas.ListItems.Item(var_i).ListSubItems(4).ForeColor = &HC000&
         lv_cajas.ListItems.Item(var_i).ListSubItems(5).ForeColor = &HC000&
         lv_cajas.ListItems.Item(var_i).ListSubItems(6).ForeColor = &HC000&
         lv_cajas.ListItems.Item(var_i).ListSubItems(7).ForeColor = &HC000&
         lv_cajas.ListItems.Item(var_i).ListSubItems(8).ForeColor = &HC000&
         lv_cajas.ListItems.Item(var_i).ListSubItems(9).ForeColor = &HC000&
        
        
        End If
    Next var_i
    If var_renglon > 0 Then
       If var_renglon <= var_n Then
          var_i = var_renglon
          lv_cajas.ListItems.Item(var_i).Bold = True
          lv_cajas.ListItems.Item(var_i).ListSubItems(1).Bold = True
          lv_cajas.ListItems.Item(var_i).ListSubItems(2).Bold = True
          lv_cajas.ListItems.Item(var_i).ListSubItems(3).Bold = True
          lv_cajas.ListItems.Item(var_i).ListSubItems(4).Bold = True
          lv_cajas.ListItems.Item(var_i).ListSubItems(5).Bold = True
          lv_cajas.ListItems.Item(var_i).ListSubItems(6).Bold = True
          lv_cajas.ListItems.Item(var_i).ListSubItems(7).Bold = True
          lv_cajas.ListItems.Item(var_i).ListSubItems(8).Bold = True
          lv_cajas.ListItems.Item(var_i).ListSubItems(9).Bold = True
          lv_cajas.ListItems.Item(var_i).ForeColor = &H8000&
          lv_cajas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H8000&
          lv_cajas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H8000&
          lv_cajas.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H8000&
          lv_cajas.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H8000&
          lv_cajas.ListItems.Item(var_i).ListSubItems(5).ForeColor = &H8000&
          lv_cajas.ListItems.Item(var_i).ListSubItems(6).ForeColor = &H8000&
          lv_cajas.ListItems.Item(var_i).ListSubItems(7).ForeColor = &H8000&
          lv_cajas.ListItems.Item(var_i).ListSubItems(8).ForeColor = &H8000&
          lv_cajas.ListItems.Item(var_i).ListSubItems(9).ForeColor = &H8000&
       End If
    End If
    lv_cajas.Refresh
End Sub






Private Sub cmd_aceptar_sello_Click()
   If Trim(txt_sello) <> "" Then
      rs.Open "insert into tb_Sellos (inte_emb_embarque, vcha_Sel_Sello) values (" + Me.txt_embarque + ",'" + Me.txt_sello + "')", cnn, adOpenDynamic, adLockOptimistic
      Set list_item = lv_sellos.ListItems.Add(, , txt_sello)
      Me.txt_sello = ""
      Me.txt_sello.SetFocus
   Else
      MsgBox "No se indico un sello", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_cambiar_embarque_Click()
   var_posible = 0
   For var_j = 1 To Me.lv_cajas.ListItems.Count
       Me.lv_cajas.ListItems.Item(var_j).Selected = True
       If Me.lv_cajas.selectedItem.SubItems(12) = "*" Then
          var_pedido_cambio_embarque = Me.lv_cajas.selectedItem.SubItems(1)
          var_posible = 1
       End If
   Next var_j
   If var_posible = 1 Then
      
      var_activa_forma_embarques = "froracle_asignacion_embarques"
      var_anden_asignar = 1
      var_anden_global = 1
      var_cambio_embarque = 1
      frmoracle_embarques.Show 1
      var_pedido_cambio_embarque = ""
      var_cambio_embarque = 0
   Else
      MsgBox "No se a seleccionado algun bulto a cambiar de embarque.", vbOKOnly, "ATENCION"
   End If

End Sub

Private Sub cmd_cancelar_sello_Click()
   Me.frm_sellos.Visible = False
End Sub

Private Sub cmd_cerrar_embarque_Click()
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   rsaux.Open "SELECT PEDIDO FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
   var_Cadena_pedidos = ""
   While Not rsaux.EOF
         If var_Cadena_pedidos = "" Then
            var_Cadena_pedidos = CStr(rsaux!pedido)
         Else
            var_Cadena_pedidos = var_Cadena_pedidos + "," + CStr(rsaux!pedido)
         End If
         rsaux.MoveNext
   Wend
   rsaux.Close
   rsaux.Open "select distinct source_header_number, lote from XXVIA_TB_PEDIDOS_DIVIDIDOS where source_header_number in (" + var_Cadena_pedidos + ") and nvl(estatus_lote,0) = 0", cnnoracle_4, adOpenDynamic, adLockOptimistic
   If rsaux.EOF Then
      rs.Open "select * from XXVIA_tB_ENCABEZADO_EMBARQUES  where embarque = " + Me.txt_embarque + " and CHAR_EMB_ESTATUS = 'E'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         If rs!char_emb_estatus = "E" Then
            rs.Close
            rs.Open "select * from tb_sellos where inte_emb_embarque = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  Set list_item = lv_sellos.ListItems.Add(, , IIf(IsNull(rs!vcha_sel_Sello), "", rs!vcha_sel_Sello))
                  rs.MoveNext
            Wend
            rs.Close
            Me.frm_sellos.Visible = True
         Else
            rs.Close
         End If
      Else
         MsgBox "El embarque ya se habia cerrado con anterioridad", vbOKOnly, "ATENCION"
         rs.Close
      End If
   Else
      var_cadena_lotes = ""
      While Not rsaux.EOF
            If var_cadena_lotes = "" Then
               var_cadena_lotes = "Pedido: " + CStr(rsaux!source_header_number) + " Lote: " + CStr(rsaux!lote)
            Else
               var_cadena_lotes = var_cadena_lotes + ", Pedido: " + CStr(rsaux!source_header_number) + " Lote: " + CStr(rsaux!lote)
            End If
            rsaux.MoveNext
      Wend
      MsgBox "Faltan por cerrar los siguientes lotes " + var_cadena_lotes, vbOKOnly, "ATENCION"
   End If
   rsaux.Close
End Sub

Private Sub cmd_cerrar_embarque_nuevo_metodo_Click()
   var_transorte_global = ""
   frmoracle_transortes.Show 1
   If var_transporte_global <> "" Then
   
   
   
   
   
   
   rs.Open "UPDATE XXVIA_TB_ENCABEZADO_EMBARQUES SET CHAR_EMB_ESTATUS = 'I', FECHA_FIN = SYSDATE, USUARIO_CERRO = '" + var_clave_usuario_global + "', TRANSPORTE = '" + var_transporte_global + "' WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
   rs.Open "update TB_ORACLE_TIEMPO_EMBARQUE_ADUANAS set HORA_FIN = GETDATE() where embarque = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
   rs.Open "UPDATE TB_ORACLE_EMBARQUES_ORDENES SET estatus = 'I' WHERE inte_emb_EMBARQUE = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
   rs.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_embarque_costales = CDbl(Me.txt_embarque)
   'frmoracle_crear_pedidos_costales.Show 1
'inicio pedidos costales
   x = 1
   If x = 1 Then
   If IsNumeric(Me.txt_embarque) Then
      rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select * from xxvia_tb_encabezado_embarques where embarque = ?"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_embarque))
           .Parameters.Append parametro
      End With
      Set rsaux4 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux4.EOF Then
         VAR_eSTATUS_ = IIf(IsNull(rsaux4!char_emb_estatus), "", rsaux4!char_emb_estatus)
         If VAR_eSTATUS_ <> "" Then
            strconsulta = "select distinct source_header_number from XXVIA_TB_SALIDAS_cAJAS where inte_emb_embarque = ?"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_embarque))
                 .Parameters.Append parametro
            End With
            Set rsaux5 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            While Not rsaux5.EOF
                  strconsulta = "SELECT TIPO_CAJA, COUNT(*) AS CANTIDAD FROM XXVIA_VW_CAJAS_POR_PEDIDO WHERE SOURCE_HEADER_NUMBER = ? AND (TIPO_CAJA LIKE '%COSTAL%' OR TIPO_CAJA LIKE 'CAJA BIASI') GROUP BY TIPO_CAJA"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux5!source_header_number)
                       .Parameters.Append parametro
                  End With
                  Set rsaux6 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  
                  
                  If Not rsaux6.EOF Then
                     strconsulta = "select * from oe_order_headers_all where order_number = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, rsaux5!source_header_number)
                          .Parameters.Append parametro
                     End With
                     Set rsaux9 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     var_posible_pedido = 1
                     If rsaux9!order_type_id = 1002 Then
                        var_posible_pedido = 0
                        var_pedido_tienda = IIf(IsNull(rsaux9!order_number), "", rsaux9!order_number)
                     End If
                     rsaux9.Close
                     'se inhabilita la facturacion de costaes
                     If var_posible_pedido = 111111 Then
                        strconsulta = "SELECT * FROM OE_ORDER_HEADERS_ALL WHERE ORIG_SYS_DOCUMENT_REF = ?"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, "SIDCOSTALES_" + Trim(CStr((rsaux5!source_header_number))))
                             .Parameters.Append parametro
                        End With
                        Set rsaux11 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        If rsaux11.EOF Then
                           strconsulta = "SELECT  distinct hp.party_name as nombre_titular,  account_number as vcha_tit_titular_id, hcas.cust_account_id AS VCHA_CLI_CLAVE_ID FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, OE_ORDER_HEADERS_ALL OHA Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND hcas.cust_account_id = OHA.SOLD_TO_ORG_ID AND ORDER_NUMBER = ? ORDER BY hp.party_name"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, Trim(CStr((rsaux5!source_header_number))))
                                .Parameters.Append parametro
                           End With
                           Set rsaux12 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           If Not rsaux12.EOF Then
                              rsaux13.Open "SELECT * FROM TB_ORACLE_TITULARES_FACTURA_COSTALES WHERE TITULAR = '" + CStr(rsaux12!vcha_tit_titular_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux13.EOF Or rsaux13.EOF Then
                                 var_posible_pedido = 1
                              Else
                                 var_posible_pedido = 0
                              End If
                              rsaux13.Close
                           Else
                              var_posible_pedido = 0
                           End If
                           rsaux12.Close
                           If var_posible_pedido = 1 Then
                              strconsulta = "SELECT SOLD_TO_ORG_ID AS TITULAR, SHIP_TO_ORG_ID AS ESTABLECIMIENTO, INVOICE_TO_ORG_ID AS CLIENTE, PRICE_LIST_ID AS LISTA_PRECIOS FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ?"
                              With comandoORA
                                  .ActiveConnection = cnnoracle_4
                                   .CommandType = adCmdText
                                   .CommandText = strconsulta
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rsaux5!source_header_number))
                                   .Parameters.Append parametro
                              End With
                              Set rsaux7 = comandoORA.execute
                              Set comandoORA = Nothing
                              Set parametro = Nothing
                              
                              strconsulta = "select name from qp_secu_list_headers_v where list_header_id = ?"
                              With comandoORA
                                   .ActiveConnection = cnnoracle_4
                                   .CommandType = adCmdText
                                   .CommandText = strconsulta
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, rsaux7!LISTA_PRECIOS)
                                   .Parameters.Append parametro
                              End With
                              Set rsaux8 = comandoORA.execute
                              Set comandoORA = Nothing
                              Set parametro = Nothing
                              var_lista_precios = rsaux8!Name
                              rsaux8.Close
                              var_clave_tipo_pedido = 1681
                              strconsulta = "INSERT INTO oe_headers_iface_all (ORDER_SOURCE_ID, orig_sys_document_ref, creation_date, created_by, last_update_date, last_updated_by, operation_code , sold_to_org_id        , SHIP_TO_ORG_id                   ,INVOICE_TO_ORG_ID     , Order_type_ID, PRICE_LIST, SHIP_FROM_ORG_ID, attribute7)"
                              strconsulta = strconsulta + "  VALUES (1001,?,SYSDATE,-1,SYSDATE,-1,'INSERT', ?,?,?,?,?,?,?)"
                              With comandoORA
                                   .ActiveConnection = cnnoracle_4
                                   .CommandType = adCmdText
                                   .CommandText = strconsulta
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, "SIDCOSTALES_" + Trim(CStr((rsaux5!source_header_number))))
                                   .Parameters.Append parametro
                                   Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux7!TITULAR)
                                   .Parameters.Append parametro
                                   Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux7!ESTABLECIMIENTO)
                                   .Parameters.Append parametro
                                   Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux7!Cliente)
                                   .Parameters.Append parametro
                                   Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, var_clave_tipo_pedido)
                                   .Parameters.Append parametro
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_lista_precios)
                                   .Parameters.Append parametro
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_unidad_organizacional)
                                   .Parameters.Append parametro
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, "FACT. DE COSTALES")
                                   .Parameters.Append parametro
                              End With
                              Set rsaux8 = comandoORA.execute
                              Set comandoORA = Nothing
                              Set parametro = Nothing
                           
                           
                           
                              var_i = 0
                              While Not rsaux6.EOF
                                    var_i = var_i + 1
                                    rs.Open "select * from tb_oracle_empaques where empaque = '" + rsaux6!tipo_caja + "'", cnn, adOpenDynamic, adLockOptimistic
                                    If Not rs.EOF Then
                                       strconsulta = "select PRIMARY_UOM_CODE, INVENTORY_ITEM_ID from xxvia_system_items_b where SEGMENT1 = ? AND ORGANIZATION_ID = ?"
                                       With comandoORA
                                            .ActiveConnection = cnnoracle_4
                                            .CommandType = adCmdText
                                            .CommandText = strconsulta
                                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, rs!codigo)
                                            .Parameters.Append parametro
                                            Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, var_unidad_organizacional)
                                            .Parameters.Append parametro
                                       End With
                                       Set rsaux8 = comandoORA.execute
                                       Set comandoORA = Nothing
                                       Set parametro = Nothing
                                       var_inventory_item_id = rsaux8!inventory_item_id
                                       VAR_MEDIDA = rsaux8!PRIMARY_UOM_CODE
                                       rsaux8.Close
                                    
                                    
                                       
                                       strconsulta = "INSERT INTO oe_lines_iface_all (ORDER_SOURCE_ID, orig_sys_document_ref,orig_sys_line_ref,inventory_item_id,ordered_quantity, operation_code, created_by, creation_date, last_updated_by, last_update_date, unit_selling_price, unit_list_price, calculate_price_flag, PRICING_QUANTITY, PRICING_QUANTITY_UOM, ATTRIBUTE1, subinventory, org_id, ship_from_org_id)"
                                       strconsulta = strconsulta + " VALUES (1001,?,?,?, ?,'INSERT', -1,SYSDATE, -1,SYSDATE,0,0,'Y', ?, ?,'0','CDI_ALMPT',?,?)"
                                       With comandoORA
                                            .ActiveConnection = cnnoracle_4
                                            .CommandType = adCmdText
                                            .CommandText = strconsulta
                                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, "SIDCOSTALES_" + Trim(CStr((rsaux5!source_header_number))))
                                            .Parameters.Append parametro
                                            Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, var_i)
                                            .Parameters.Append parametro
                                            Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, var_inventory_item_id)
                                            .Parameters.Append parametro
                                            Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux6!cantidad)
                                            .Parameters.Append parametro
                                            Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux6!cantidad)
                                            .Parameters.Append parametro
                                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, VAR_MEDIDA)
                                            .Parameters.Append parametro
                                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_empresa)
                                            .Parameters.Append parametro
                                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_unidad_organizacional)
                                            .Parameters.Append parametro
                                       End With
                                       Set rsaux8 = comandoORA.execute
                                       Set comandoORA = Nothing
                                       Set parametro = Nothing
                                    End If
                                    rs.Close
                                    rsaux6.MoveNext
                              Wend
                              On Error GoTo salir2
                              rsaux8.Open "INSERT INTO oe_actions_iface_all (order_source_ID, orig_sys_document_ref, operation_code) VALUES (1001, 'SIDCOSTALES_" + Trim(CStr((rsaux5!source_header_number))) + "','BOOK_ORDER')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              rsaux8.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              rsaux8.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              rsaux8.Open "CALL XXVIA_PK_INTERFACES_OM.importar_pedido('SIDCOSTALES_" + Trim(CStr((rsaux5!source_header_number))) + "'," + var_empresa + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              var_cadena = "select * from oe_order_headers_all where orig_sys_document_ref = 'SIDCOSTALES_" + Trim(CStr((rsaux5!source_header_number))) + "'"
                              'MsgBox var_cadena
                              rsaux8.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux8.EOF Then
                                 rsaux9.Open "INSERT INTO TB_ORACLE_PEDIDOS_CERRADOS (PEDIDO, REQUEST_ID) VALUES (" + CStr(rsaux8!order_number) + ",0)", cnn, adOpenDynamic, adLockOptimistic
                                 rsaux8.Close
                              
                              Else
                                 rsaux8.Close
                              End If
                           End If
                        End If
                        rsaux11.Close
                     End If
                  End If
                  rsaux6.Close
                  rsaux5.MoveNext
            Wend
            rsaux5.Close
            MsgBox "Se a terminado el proceso de insercion de costales"
         End If
      Else
         MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
      End If
      rsaux4.Close
   End If
   End If
'fin pedidos costales




   MsgBox "Se a cerrado el embarque", vbOKOnly, "ATENCION"
   Me.frm_sellos.Visible = False
   Else
      MsgBox "No se selecciono un transporte", vbOKOnly, "ATENCION"
   End If
   Exit Sub
salir2:
   If Err.Number = -2147217900 Then
      rsaux10.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      rsaux10.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      'MsgBox Err.Description
      Resume
   Else
      'MsgBox Err.Description
      Resume Next
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
   End If
   Exit Sub
salir_factura:
   MsgBox "No se pudo generar el documento electrónico", vbOKOnly, "ATENCION"
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


End Sub

Private Sub cmd_cerrar_pedido_Click()
   If Me.lv_cajas.ListItems.Count > 0 Then
      rs.Open "select pedido from TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque in (select embarque  from TB_ORACLE_GRUPOS_EMBARQUES where grupo = (select top 1 grupo from TB_ORACLE_GRUPOS_EMBARQUES where embarque = '" + Me.txt_embarque + "')) and PEDIDO not in (select pedido from tb_oracle_pedidos_Cerrados_cn union all select pedido from tb_oracle_pedidos_Cerrados)", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_ultimo_pedido = rs.RecordCount
      Else
         var_ultimo_pedido = 0
      End If
      rs.Close
      
      var_posible_Cerrar = 0
      If var_ultimo_pedido = 1 Then
         var_porcentaje_aduana = CDbl(Replace(Me.lbl_porcentaje_aduana, "%", ""))
         If var_porcentaje_aduana < 100 Then
            var_posible_Cerrar = 0
         Else
            var_posible_Cerrar = 1
         End If
      Else
         var_posible_Cerrar = 1
      End If
      If var_posible_Cerrar = 1 Then
         rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         'strconsulta = "SELECT * FROM XXVIA_COMP_LEIDA_VS_PED_AFEC WHERE inte_emb_embarque = ? AND SOURCE_HEADER_NUMBER =  ? AND CANTIDAD_LEIDA > CANTIDAD_PEDIDA_AFECTADA"
         strconsulta = "SELECT * FROM XXVIA_COMP_LEIDA_VS_PED_AFEC WHERE  SOURCE_HEADER_NUMBER =  ? AND CANTIDAD_LEIDA > CANTIDAD_PEDIDA_AFECTADA"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              'Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_embarque))
              '.Parameters.Append parametro
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.lv_cajas.selectedItem.SubItems(1)))
              .Parameters.Append parametro
         End With
         Set rsaux3 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         var_posible_cerrado_comparacion = 1
         If Not rsaux3.EOF Then
            var_posible_cerrado_comparacion = 0
            cnn.BeginTrans
            rs.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM  TB_TEMP_ORACLE_COMPARACION_CANTIDADES_LEIDAS_VS_PEDIDAS_AFECTADAS", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
            Else
               var_consecutivo = 1
            End If
            rs.Close
            rsaux.Open "INSERT INTO TB_TEMP_ORACLE_COMPARACION_CANTIDADES_LEIDAS_VS_PEDIDAS_AFECTADAS (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            While Not rsaux3.EOF
                  strconsulta = "SELECT * FROM XXVIA_TB_SALIDAS_CAJAS WHERE inte_emb_embarque = ? and SOURCE_HEADER_NUMBER = ? AND SEGMENT1 = ? AND FLOA_SAL_cANTIDAD_LEIDA > 0"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rsaux3!inte_Emb_Embarque))
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rsaux3!source_header_number))
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, rsaux3!segment1)
                       .Parameters.Append parametro
                  End With
                  Set rsaux2 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  VAR_CAJAS = ""
                  While Not rsaux2.EOF
                        If VAR_CAJAS = "" Then
                           VAR_CAJAS = "CAJA: " + CStr(IIf(IsNull(rsaux2!INTE_PAQ_CAJA), 0, rsaux2!INTE_PAQ_CAJA)) + " CANTIDAD: " + CStr(IIf(IsNull(rsaux2!floa_Sal_cantidad_leida), 0, rsaux2!floa_Sal_cantidad_leida))
                        Else
                           VAR_CAJAS = VAR_CAJAS + ", CAJA: " + CStr(IIf(IsNull(rsaux2!INTE_PAQ_CAJA), 0, rsaux2!INTE_PAQ_CAJA)) + " CANTIDAD: " + CStr(IIf(IsNull(rsaux2!floa_Sal_cantidad_leida), 0, rsaux2!floa_Sal_cantidad_leida))
                        End If
                        rsaux2.MoveNext
                  Wend
                  rsaux2.Close
                  var_cadena = "INSERT INTO TB_TEMP_ORACLE_COMPARACION_CANTIDADES_LEIDAS_VS_PEDIDAS_AFECTADAS (INTE_TEM_CONSECUTIVO, EMBARQUE, PEDIDO, CODIGO, DESCRIPCION, CANTIDAD_PEDIDA, CANTIDAD_LEIDA, CAJAS )"
                  var_cadena = var_cadena + " VALUES (" + CStr(var_consecutivo) + "," + CStr(rsaux3!inte_Emb_Embarque) + "," + CStr(rsaux3!source_header_number) + ",'" + rsaux3!segment1 + "', '" + rsaux3!item_description + "'," + CStr(rsaux3!CANTIDAD_PEDIDA_AFECTADA) + "," + CStr(rsaux3!cantidad_leida) + ",'" + VAR_CAJAS + "')"
                  rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  rsaux3.MoveNext
            Wend
         End If
         rsaux3.Close
         'var_posible_cerrado_comparacion = 1
         If var_posible_cerrado_comparacion = 1 Then
            var_pedido = CDbl(Me.lv_cajas.selectedItem.SubItems(1))
            If var_pedido >= 10000002 Then
               If rs.State = 1 Then
                  rs.Close
               End If
               Call cerrar_pedido
            Else
               rs.Open "SELECT NVL(ESTATUS_PEDIDO,0) AS ESTATUS_PEDIDO FROM XXVIA_TB_SALIDAS_CAJAS WHERE SOURCE_HEADER_NUMBER = " + Me.lv_cajas.selectedItem.SubItems(1) + " and floa_sal_Cantidad_leida > 0", cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  If rs!estatus_pedido = 1 Then
                     var_pedido = CDbl(Me.lv_cajas.selectedItem.SubItems(1))
                     rsaux.Open "SELECT * FROM tb_oracle_cajas_aduana WHERE pedido = " + CStr(var_pedido) + " AND isnull(estatus,'') <> 'L' and embarque = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        MsgBox "No se puede cerrar el pedido debido a que no se han despachado todas las cajas", vbOKOnly, "ATENCION"
                     Else
                        'Call cmd_nuevo_Click
                        Call cerrar_pedido
                     End If
                     If rsaux.State = 1 Then
                        rsaux.Close
                     End If
                  Else
                     MsgBox "El pedido no a sido cerrado por parte del lector", vbOKOnly, "ATENCION"
                  End If
               End If
               If rs.State = 1 Then
                  rs.Close
               End If
            End If
         Else
            rsaux.Open "DELETE FROM TB_TEMP_ORACLE_COMPARACION_CANTIDADES_LEIDAS_VS_PEDIDAS_AFECTADAS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " and codigo is null", cnn, adOpenDynamic, adLockOptimistic
            MsgBox "El embarque tiene diferencias entre las piezas pedidas y las leidas", vbOKOnly, "ATENCION"
            Set reporte = appl.OpenReport(App.Path + "\rep_oracle_diferencias_pedido_leido.rpt")
            var_cadena = "{TB_TEMP_ORACLE_COMPARACION_CANTIDADES_LEIDAS_VS_PEDIDAS_AFECTADAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            reporte.RecordSelectionFormula = var_cadena
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de diferencias pedido contra leido"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            rsaux.Open "DELETE FROM TB_TEMP_ORACLE_COMPARACION_CANTIDADES_LEIDAS_VS_PEDIDAS_AFECTADAS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         End If
      Else
         MsgBox "El ultimo pedido no puede ser cerrado ya que no a sido completado el volumen de la unidad", vbOKOnly, "ATENCION"
      End If
    End If
    If Me.txt_codigo.Enabled = True Then
       Me.txt_codigo.SetFocus
    End If
End Sub

Private Sub cmd_mensaje_4_Click()
   Me.wmp4.Controls.play
End Sub

Private Sub cmd_nuevo_Click()
   If rs.State = 1 Then
      rs.Close
   End If
   rs.Open "SELECT SUM(PIEZAS) FROM tb_oracle_cajas_aduana WHERE ESTATUS = 'L' AND EMBARQUE = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
   Me.txt_cantidad = Format(IIf(IsNull(rs(0).Value), 0, rs(0).Value), "###,###,##0.00")
   rs.Close
   var_contador = 1
   rs.Open "select AGENTE, nombre_agente, pedido, cliente, orden_pedido, ESTATUS from tb_oracle_pedidos_asignados_embarques WHERE EMBARQUE = " + Me.txt_embarque + " AND ISNULL(ESTATUS,'') = '' order by orden_pedido, pedido", cnn, adOpenDynamic, adLockOptimistic
   Me.lv_cajas.ListItems.Clear
   Me.lv_cajas_siguientes.ListItems.Clear
   While Not rs.EOF
         If var_contador = 1 Then
            var_pedido = rs!pedido
            rsaux10.Open "select char_paq_estatus, inte_paq_caja, sum(floa_sal_cantidad_leida) as cantidad from xxvia_tb_Salidas_cajas where source_header_number = " + CStr(rs!pedido) + " and char_paq_estatus in ('I','S') group by char_paq_estatus, inte_paq_caja", cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rsaux10.EOF
                  var_numero_caja = IIf(IsNull(rsaux10!INTE_PAQ_CAJA), 0, rsaux10!INTE_PAQ_CAJA)
                  If Len(Trim(Str(var_numero_caja))) = 1 Then
                     var_referencia_caja = "00" + Trim(Str(var_numero_caja))
                  End If
                  If Len(Trim(Str(var_numero_caja))) = 2 Then
                     var_referencia_caja = "0" + Trim(Str(var_numero_caja))
                  End If
                  If Len(Trim(Str(var_numero_caja))) = 3 Then
                     var_referencia_caja = Trim(Str(var_numero_caja))
                  End If
                  If Len(Trim(Str(txt_embarque))) = 1 Then
                     var_referencia_embarque = "00000" + Trim(Str(txt_embarque))
                  End If
                  If Len(Trim(Str(txt_embarque))) = 2 Then
                     var_referencia_embarque = "0000" + Trim(Str(txt_embarque))
                  End If
                  If Len(Trim(Str(txt_embarque))) = 3 Then
                     var_referencia_embarque = "000" + Trim(Str(txt_embarque))
                  End If
                  If Len(Trim(Str(txt_embarque))) = 4 Then
                      var_referencia_embarque = "00" + Trim(Str(txt_embarque))
                  End If
                  If Len(Trim(Str(txt_embarque))) = 5 Then
                     var_referencia_embarque = "0" + Trim(Str(txt_embarque))
                  End If
                  If Len(Trim(Str(txt_embarque))) = 6 Then
                     var_referencia_embarque = "" + Trim(Str(txt_embarque))
                  End If
                  var_codigo_caja = "C" + var_referencia_embarque + var_referencia_caja
                  Set list_item = Me.lv_cajas.ListItems.Add(, , var_codigo_caja)
                  list_item.SubItems(1) = IIf(IsNull(rs!pedido), "", rs!pedido)
                  list_item.SubItems(2) = IIf(IsNull(rs!nombre_agente), "", rs!nombre_agente)
                  list_item.SubItems(3) = IIf(IsNull(rs!Cliente), "", rs!Cliente)
                  list_item.SubItems(4) = Format(rsaux10!cantidad, "###,###,##0.00")
                  list_item.SubItems(5) = IIf(IsNull(rsaux10!char_paq_estatus), "", rsaux10!char_paq_estatus)
                  list_item.SubItems(6) = IIf(IsNull(rsaux10!INTE_PAQ_CAJA), "", rsaux10!INTE_PAQ_CAJA)
                   
                  rsaux10.MoveNext
            Wend
            rsaux10.Close
         Else
            rsaux10.Open "select char_paq_estatus, inte_paq_caja, sum(floa_sal_cantidad_leida) as cantidad from xxvia_tb_Salidas_cajas where source_header_number = " + CStr(rs!pedido) + " and char_paq_estatus = 'I' and source_header_number <> " + CStr(var_pedido) + " group by char_paq_estatus, inte_paq_caja", cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rsaux10.EOF
                  var_numero_caja = IIf(IsNull(rsaux10!INTE_PAQ_CAJA), 0, rsaux10!INTE_PAQ_CAJA)
                  If Len(Trim(Str(var_numero_caja))) = 1 Then
                     var_referencia_caja = "00" + Trim(Str(var_numero_caja))
                  End If
                  If Len(Trim(Str(var_numero_caja))) = 2 Then
                     var_referencia_caja = "0" + Trim(Str(var_numero_caja))
                  End If
                  If Len(Trim(Str(var_numero_caja))) = 3 Then
                     var_referencia_caja = Trim(Str(var_numero_caja))
                  End If
                  If Len(Trim(Str(txt_embarque))) = 1 Then
                     var_referencia_embarque = "00000" + Trim(Str(txt_embarque))
                  End If
                  If Len(Trim(Str(txt_embarque))) = 2 Then
                     var_referencia_embarque = "0000" + Trim(Str(txt_embarque))
                  End If
                  If Len(Trim(Str(txt_embarque))) = 3 Then
                     var_referencia_embarque = "000" + Trim(Str(txt_embarque))
                  End If
                  If Len(Trim(Str(txt_embarque))) = 4 Then
                      var_referencia_embarque = "00" + Trim(Str(txt_embarque))
                  End If
                  If Len(Trim(Str(txt_embarque))) = 5 Then
                     var_referencia_embarque = "0" + Trim(Str(txt_embarque))
                  End If
                  If Len(Trim(Str(txt_embarque))) = 6 Then
                     var_referencia_embarque = "" + Trim(Str(txt_embarque))
                  End If
                  var_codigo_caja = "C" + var_referencia_embarque + var_referencia_caja
                  Set list_item = Me.lv_cajas_siguientes.ListItems.Add(, , var_codigo_caja)
                  list_item.SubItems(1) = IIf(IsNull(rs!pedido), "", rs!pedido)
                  list_item.SubItems(2) = IIf(IsNull(rs!nombre_agente), "", rs!nombre_agente)
                  list_item.SubItems(3) = IIf(IsNull(rs!Cliente), "", rs!Cliente)
                  list_item.SubItems(4) = Format(rsaux10!cantidad, "###,###,##0.00")
                  list_item.SubItems(5) = IIf(IsNull(rsaux10!char_paq_estatus), "", rsaux10!char_paq_estatus)
                  list_item.SubItems(6) = IIf(IsNull(rsaux10!INTE_PAQ_CAJA), "", rsaux10!INTE_PAQ_CAJA)
                  rsaux10.MoveNext
            Wend
            rsaux10.Close
         End If
         var_contador = var_contador + 1
         rs.MoveNext
   Wend
   rs.Close
   Call ilumina_grid
End Sub

Private Sub Command1_Click()
   Unload Me
End Sub


Private Sub Command2_Click()
Dim clnt As New SoapClient30

Dim var_arreglo() As String
Dim var_container_id As String
Dim var_trip_id As String
Dim var_b As Boolean
VAR_ESTATUS = "E"
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
   If rsaux10.State = 1 Then
      rsaux10.Close
   End If
   rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   strconsulta = "SELECT * FROM XXVIA_COMP_LEIDA_VS_PED_AFEC WHERE inte_emb_embarque = ? and CANTIDAD_LEIDA > CANTIDAD_PEDIDA_AFECTADA"
   With comandoORA
        .ActiveConnection = cnnoracle_4
        .CommandType = adCmdText
        .CommandText = strconsulta
        Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_embarque))
        .Parameters.Append parametro
   End With
   Set rsaux3 = comandoORA.execute
   Set comandoORA = Nothing
   Set parametro = Nothing
   var_posible_cerrado_comparacion = 1
   If Not rsaux3.EOF Then
      var_posible_cerrado_comparacion = 0
      cnn.BeginTrans
      rs.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM  TB_TEMP_ORACLE_COMPARACION_CANTIDADES_LEIDAS_VS_PEDIDAS_AFECTADAS", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
      Else
         var_consecutivo = 1
      End If
      rs.Close
      rsaux.Open "INSERT INTO TB_TEMP_ORACLE_COMPARACION_CANTIDADES_LEIDAS_VS_PEDIDAS_AFECTADAS (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
      cnn.CommitTrans
      While Not rsaux3.EOF
            strconsulta = "SELECT * FROM XXVIA_TB_SALIDAS_CAJAS WHERE inte_emb_embarque = ? and SOURCE_HEADER_NUMBER = ? AND SEGMENT1 = ? AND FLOA_SAL_cANTIDAD_LEIDA > 0"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rsaux3!inte_Emb_Embarque))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rsaux3!source_header_number))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, rsaux3!segment1)
                 .Parameters.Append parametro
            End With
            Set rsaux2 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            VAR_CAJAS = ""
            While Not rsaux2.EOF
                  If VAR_CAJAS = "" Then
                     VAR_CAJAS = "CAJA: " + CStr(IIf(IsNull(rsaux2!INTE_PAQ_CAJA), 0, rsaux2!INTE_PAQ_CAJA)) + " CANTIDAD: " + CStr(IIf(IsNull(rsaux2!floa_Sal_cantidad_leida), 0, rsaux2!floa_Sal_cantidad_leida))
                  Else
                     VAR_CAJAS = VAR_CAJAS + ", CAJA: " + CStr(IIf(IsNull(rsaux2!INTE_PAQ_CAJA), 0, rsaux2!INTE_PAQ_CAJA)) + " CANTIDAD: " + CStr(IIf(IsNull(rsaux2!floa_Sal_cantidad_leida), 0, rsaux2!floa_Sal_cantidad_leida))
                  End If
                  rsaux2.MoveNext
            Wend
            rsaux2.Close
            var_cadena = "INSERT INTO TB_TEMP_ORACLE_COMPARACION_CANTIDADES_LEIDAS_VS_PEDIDAS_AFECTADAS (INTE_TEM_CONSECUTIVO, EMBARQUE, PEDIDO, CODIGO, DESCRIPCION, CANTIDAD_PEDIDA, CANTIDAD_LEIDA, CAJAS )"
            var_cadena = var_cadena + " VALUES (" + CStr(var_consecutivo) + "," + CStr(rsaux3!inte_Emb_Embarque) + "," + CStr(rsaux3!source_header_number) + ",'" + rsaux3!segment1 + "', '" + rsaux3!item_description + "'," + CStr(rsaux3!CANTIDAD_PEDIDA_AFECTADA) + "," + CStr(rsaux3!cantidad_leida) + ",'" + VAR_CAJAS + "')"
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            rsaux3.MoveNext
      Wend
   End If
   rsaux3.Close
   If var_posible_cerrado_comparacion = 1 Then
      If VAR_ESTATUS = "E" Then
         var_si = MsgBox("¿Desea cerrar el embarque?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_si = MsgBox("Confirmar el cerrado del embarque", vbYesNo, "ATENCION")
            If var_si = 6 Then
               x = 1
            Else
               x = 0
            End If
         Else
            x = 0
         End If
         If x = 1 Then
            rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
            VAR_X_TRIP_ID = rs!ARREGLO_0
            var_x_trip_name = rs!ARREGLO_1
            rs.Close
            If var_x_trip_name <> "" Then
               rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If rs!tipo_embarque = 2 Then
                  rsaux.Open "select distinct source_header_number from xxvia_tb_salidas_CAJAS where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
               End If
               var_Cadena_pedidos = ""
               var_j = 0
               While Not rsaux.EOF
                     If var_Cadena_pedidos = "" Then
                        var_Cadena_pedidos = "'" + CStr(rsaux!source_header_number) + "'"
                     Else
                        var_Cadena_pedidos = var_Cadena_pedidos + ", '" + CStr(rsaux!source_header_number) + "'"
                     End If
                     var_j = var_j + 1
                     rsaux.MoveNext
               Wend
               rsaux.Close
               var_cadena = "SELECT HCAS.CUST_ACCOUNT_ID, HCAS.CUST_ACCT_SITE_ID as customer_id, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND to_number(source_header_number) IN (" + var_Cadena_pedidos + ")"
               var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID  AND released_status = 'Y'"
               rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               'rsaux2.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux!cust_account_id), cnnoracle_4, adOpenDynamic, adLockOptimistic
               'VAR_AGENTE_str = rsaux2!COLLECTOR_ID
               'var_nombre_agente_str = rsaux2!Name
               'rsaux2.Close
               While Not rsaux.EOF
                     'rsaux3.Open "SELECT * FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " AND source_header_number = " + CStr(CDbl(rsaux!SOURCE_HEADER_NUMBER)) + " AND DELIVERY_DETAIL_ID = " + CStr(rsaux!DELIVERY_DETAIL_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     strconsulta = "SELECT * FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = ? AND source_header_number = ? AND DELIVERY_DETAIL_ID = ?"
                     With comandoORA
                         .ActiveConnection = cnnoracle_4
                         .CommandType = adCmdText
                         .CommandText = strconsulta
                         Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_embarque))
                         .Parameters.Append parametro
                         Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rsaux!source_header_number))
                         .Parameters.Append parametro
                         Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(rsaux!delivery_detail_id))
                         .Parameters.Append parametro
                     End With
                     Set rsaux3 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     If rsaux3.EOF Then
                        var_cadena = "INSERT INTO XXVIA_TB_SALIDAS_CAJAS (INTE_EMB_EMBARQUE, SOURCE_HEADER_NUMBER, SEGMENT1, FLOA_SAL_CANTIDAD_LEIDA, INVENTORY_ITEM_ID, DELIVERY_DETAIL_ID, SOURCE_LINE_NUMBER, DELIVERY_ID, INTE_PAQ_CAJA, CUSTOMER_ID, SUBINVENTORY, NAME, COLLECTOR_ID, ITEM_DESCRIPTION, CUSTOMER_NAME)"
                        var_cadena = var_cadena + " values (" + Me.txt_embarque + "," + CStr(CDbl(rsaux!source_header_number)) + ",'" + rsaux!segment1 + "',0," + CStr(rsaux!inventory_item_id) + "," + CStr(rsaux!delivery_detail_id) + ",'" + CStr(rsaux!SOURCE_LINE_NUMBER) + "'," + CStr(IIf(IsNull(rsaux!delivery_id), 0, rsaux!delivery_id)) + ",0," + CStr(rsaux!CUSTOMER_ID) + ",'" + CStr(rsaux!subinventory) + "', '" + var_nombre_agente_str + "','" + CStr(VAR_AGENTE_str) + "','" + CStr(rsaux!Description) + "','" + rsaux!customer_name + "')"
                        rsaux4.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux3.Close
                     rsaux.MoveNext
               Wend
               rsaux.Close
               If rsaux9.State = 1 Then
                  rsaux9.Close
               End If
               x = 1
               If x = 0 Then
                  rsaux9.Open "select * from tb_unidadesorganizacionales where vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux9.EOF Then
                     VAR_USER_ID = rsaux9!user_id
                     VAR_RESP_ID = rsaux9!resp_id
                     VAR_RESP_APPL_ID = rsaux9!resp_appl_id
                  End If
                  rsaux9.Close
                  var_cadena = "SELECT a.source_line_id, OHA.HEADER_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, E.NAME from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND A.SOURCE_HEADER_NUMBER in (" + var_Cadena_pedidos + ")"
                  var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id = e.collector_id AND released_status = 'B' order by A.source_header_number"
                  rsaux9.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  While Not rsaux9.EOF
                        rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        'MsgBox "call XXVIA_SP_DEPURA_ORDEN_SURTIDO (" + CStr(CDbl(rsaux9!header_id)) + ", " + CStr(CDbl(rsaux9!SOURCE_LINE_ID)) + ", 'PRODUCCION')"
                        On Error GoTo salir2:
                        rsaux7.Open "call XXVIA_SP_DEPURA_ORDEN_SURTIDO (" + CStr(CDbl(rsaux9!header_id)) + ", " + CStr(CDbl(rsaux9!source_LINE_ID)) + ", 'PRODUCCION'," + CStr(VAR_USER_ID) + "," + CStr(VAR_RESP_ID) + "," + CStr(VAR_RESP_APPL_ID) + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        rsaux9.MoveNext
                  Wend
                  rsaux9.Close
                  rs.Close
               End If
               clnt.MSSoapInit var_webservice
               If rs.State = 1 Then
                  rs.Close
               End If
               rs.Open "SELECT delivery_detail_id, sum(floa_sal_Cantidad_leida) as floa_sal_Cantidad_leida FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " group by delivery_detail_id", cnnoracle_4, adOpenDynamic, adLockOptimistic
               VAR_ZZ = 0
               While Not rs.EOF
                     VAR_ZZ = VAR_ZZ + 1
                     'VAR_ZZZ = 0
                     If VAR_ZZ = 50 Then
                        var_zzz = var_zzz + 1
                        VAR_ZZ = 0
                     End If
                     'rsaux.Open "SELECT * FROM WSH_DELIVERABLES_V WHERE delivery_detail_id = " + CStr(rs!DELIVERY_DETAIL_ID) + " AND RELEASED_STATUS = 'Y'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     strconsulta = "SELECT * FROM WSH_DELIVERABLES_V WHERE delivery_detail_id = ? AND RELEASED_STATUS = 'Y'"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!delivery_detail_id)
                          .Parameters.Append parametro
                     End With
                     Set rsaux = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     
                     If Not rsaux.EOF Then
                        'var_b = clnt.actualizar_detalle(Val(rs!delivery_detail_id), CDbl(rs!FLOA_sAL_cANTIDAD_LEIDA), "OE", 0)
                        On Error GoTo salir2:
                        rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        'rsaux6.Open "select max(inte_paq_caja) as inte_paq_caja  from xxvia_tb_Salidas_cajas where delivery_detail_id = " + CStr(rs!DELIVERY_DETAIL_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        strconsulta = "select max(inte_paq_caja) as inte_paq_caja  from xxvia_tb_Salidas_cajas where delivery_detail_id = ?"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!delivery_detail_id)
                             .Parameters.Append parametro
                        End With
                        Set rsaux6 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        var_consecutivo = rsaux6!INTE_PAQ_CAJA
                        rsaux6.Close
                        rsaux6.Open "CALL xxvia_pk_interfaces_om.actualizar_detalle (1.0, " + CStr(rs!delivery_detail_id) + "," + CStr(rs!floa_Sal_cantidad_leida) + ",'OE'," + CStr(var_consecutivo) + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux.Close
                     rs.MoveNext
               Wend
               rs.Close
               Set clnt = Nothing
               'clnt.MSSoapInit var_webservice
               'rs.Open "SELECT DISTINCT DELIVERY_ID FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
               'While Not rs.EOF
               '
               '      var_arreglo = clnt.ASIGNAR_embarque(rs!delivery_id, Val(VAR_X_TRIP_ID), "CONFIRM")
               '      rs.MoveNext
               'Wend
               'rs.Close
               'Set clint = Nothing
               rs.Open "SELECT * FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
               While Not rs.EOF
                     If IIf(IsNull(rs!floa_Sal_cantidad_leida), 0, rs!floa_Sal_cantidad_leida) > 0 Then
                        var_cadena = "INSERT INTO XXVIA_TB_DETALLE_CAJAS (EMBARQUE, PEDIDO,AGENTE, NOMBRE_AGENTE,CLIENTE,NOMBRE_CLIENTE,CODIGO, DESCRIPCION, CANTIDAD, PESO, CAJA, INVENTORY_ITEM_ID, CAJA_PEDIDO)"
                        var_cadena = var_cadena + " values (" + Me.txt_embarque + ", " + CStr(rs!source_header_number) + ",'" + CStr(IIf(IsNull(rs!collector_id), 0, rs!collector_id)) + "', '" + IIf(IsNull(rs!Name), "", rs!Name) + "',  '" + CStr(rs!CUSTOMER_ID) + "','" + IIf(IsNull(rs!customer_name), "", rs!customer_name) + "','" + rs!segment1 + "','" + rs!item_description + "'," + CStr(rs!floa_Sal_cantidad_leida) + ",0," + CStr(rs!INTE_PAQ_CAJA) + "," + CStr(rs!inventory_item_id) + "," + CStr(IIf(IsNull(rs!caja_pedido), 0, rs!caja_pedido)) + ")"
                       rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     End If
                     rs.MoveNext
               Wend
               rs.Close
               rs.Open "UPDATE XXVIA_TB_ENCABEZADO_EMBARQUES SET CHAR_EMB_ESTATUS = 'I', FECHA_FIN = SYSDATE, USUARIO_CERRO = '" + var_clave_usuario_global + "' WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
               rs.Open "UPDATE TB_ORACLE_EMBARQUES_ORDENES SET estatus = 'I' WHERE inte_emb_EMBARQUE = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
               x = 0
               If x = 1 Then
                  rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     If IIf(IsNull(rs!char_emb_estatus), "", rs!char_emb_estatus) = "I" Then
                        If rs!tipo_embarque = 2 Then
                           rsaux.Open "select distinct source_header_number from xxvia_tb_salidas_cAJAS where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        End If
                        var_Cadena_pedidos = ""
                        var_j = 0
                        While Not rsaux.EOF
                              If var_Cadena_pedidos = "" Then
                                 var_Cadena_pedidos = "'" + CStr(rsaux!source_header_number) + "'"
                              Else
                                 var_Cadena_pedidos = var_Cadena_pedidos + ", '" + CStr(rsaux!source_header_number) + "'"
                              End If
                              var_j = var_j + 1
                              rsaux.MoveNext
                        Wend
                        rsaux.Close
                        var_i = 0
                        If var_i = 1 Then
                           While var_j <> var_i
                                 var_i = 0
                                 var_cadena = "SELECT e.collector_id, A.SOURCE_HEADER_NUMBER,  HL.ADDRESS1 AS CUSTOMER_NAME,  A.released_status,  E.NAME , sum(shipped_quantity) as cantidad from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, hz_customer_profiles D, ar_collectors e Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id "
                                 var_cadena = var_cadena + " AND A.SOURCE_HEADER_NUMBER in (" + var_Cadena_pedidos + ") AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND D.collector_id = e.collector_id AND released_status = 'C' group by  e.collector_id, A.SOURCE_HEADER_NUMBER, HL.ADDRESS1,  A.released_status,  E.NAME"
                                 'MsgBox var_cadena_pedidos
                                 rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 While Not rsaux.EOF
                                       var_i = var_i + 1
                                       rsaux.MoveNext
                                 Wend
                                 rsaux.Close
                           Wend
                           x = 1
                           If x = 0 Then
                              var_cadena_pedidos_global = var_Cadena_pedidos
                              var_cadena = "SELECT a.source_line_id, OHA.HEADER_ID from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND A.SOURCE_HEADER_NUMBER in (" + var_cadena_pedidos_global + ") "
                              var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id = e.collector_id AND released_status = 'B' order by A.source_header_number"
                              rsaux7.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux7.EOF Then
                                 var_tipo_depurado = 1
                                 frmoracle_depurar_pedidos.Show 1
                              End If
                              rsaux7.Close
                              var_tipo_depurado = 0
                              var_cadena = "SELECT a.source_line_id, OHA.HEADER_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, E.NAME from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND A.SOURCE_HEADER_NUMBER in (" + var_cadena_pedidos_global + ")"
                              var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id = e.collector_id AND released_status = 'B' order by A.source_header_number"
                              rsaux9.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux9.EOF Then
                                 rsaux9.Close
                                 var_sigue = 1
                                 While var_sigue = 1
                                       If rsaux8.State = 1 Then
                                          rsaux8.Close
                                       End If
                                       var_cadena = "SELECT a.source_line_id, OHA.HEADER_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, E.NAME from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND A.SOURCE_HEADER_NUMBER in (" + var_cadena_pedidos_global + ")"
                                       var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id = e.collector_id AND released_status = 'B' order by A.source_header_number"
                                       rsaux8.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                       If rsaux8.EOF Then
                                          var_sigue = 0
                                       Else
                                          While Not rsaux8.EOF
                                                rsaux7.Open "SELECT * FROM TB_ORACLE_NEGADO WHERE PEDIDO IN (" + CStr(rsaux8!source_header_number) + ") AND INVENTORY_ITEM_ID = " + CStr(rsaux8!inventory_item_id), cnn, adOpenDynamic, adLockOptimistic
                                                rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                                Set clnt = Nothing
                                                clnt.MSSoapInit var_webservice
                                                var_s = clnt.cancelar_back_order(CDbl(rsaux8!header_id), CDbl(rsaux8!source_LINE_ID), rsaux7!CAUSA_NEGADO)
                                                Set clnt = Nothing
                                                rsaux7.Close
                                                rsaux8.MoveNext
                                          Wend
                                       End If
                                       rsaux8.Close
                                 Wend
                              Else
                                 rsaux9.Close
                              End If
                           End If 'x
                        End If
                     End If
                  End If
               End If
               '--------------- confirmar pedidos
               x = 1
               If x = 1 Then
                  rsaux.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     VAR_X_TRIP_ID = rs!ARREGLO_0
                     var_x_trip_name = rs!ARREGLO_1
                     VAR_ESTATUS = IIf(IsNull(rs!char_emb_estatus), "", rs!char_emb_estatus)
                     If IIf(IsNull(rs!char_emb_estatus), "", rs!char_emb_estatus) = "I" Then
                        If rs!tipo_embarque = 1 Then
                           rsaux.Open "select distinct source_header_number from xxvia_tb_salidas where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        End If
                        If rs!tipo_embarque = 2 Then
                           rsaux.Open "select distinct source_header_number from xxvia_tb_SAlidas_cajas where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        End If
                        VAR_CADENA_PEDIDOS_M = ""
                        While Not rsaux.EOF
                              If VAR_CADENA_PEDIDOS_M = "" Then
                                 VAR_CADENA_PEDIDOS_M = CStr(rsaux!source_header_number)
                              Else
                                 VAR_CADENA_PEDIDOS_M = VAR_CADENA_PEDIDOS_M + ", " + CStr(rsaux!source_header_number)
                              End If
                              rsaux.MoveNext
                        Wend
                        var_Cadena_pedidos = ""
                        rsaux.MoveFirst
                        While Not rsaux.EOF
                              rsaux1.Open "select distinct delivery_id from wsh_deliverables_v where SOURCE_HEADER_NUMBER = " + CStr(rsaux!source_header_number) + " AND delivery_id IS NOT NULL", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              VAR_ENTREGA = rsaux1!delivery_id
                              rsaux1.Close
                              rsaux1.Open "select distinct source_header_number from wsh_deliverables_v where delivery_id = " + CStr(VAR_ENTREGA), cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux1.EOF Then
                                 var_j = 0
                                 While Not rsaux1.EOF
                                       var_j = var_j + 1
                                       rsaux1.MoveNext
                                 Wend
                                 If var_j > 1 Then
                                    If var_Cadena_pedidos = "" Then
                                       var_Cadena_pedidos = CStr(rsaux!source_header_number) + " ENTREGA: " + CStr(VAR_ENTREGA)
                                    Else
                                       var_Cadena_pedidos = var_Cadena_pedidos + ", " + CStr(rsaux!source_header_number) + " ENTREGA: " + CStr(VAR_ENTREGA)
                                    End If
                                 End If
                              End If
                              rsaux1.Close
                              rsaux.MoveNext
                        Wend
                        rsaux.MoveFirst
                        If var_Cadena_pedidos <> "" Then
                           MsgBox "Los pedidos siguientes tienen dos entregas " + var_Cadena_pedidos
                        Else
                           cnn.BeginTrans
                           rsaux8.Open "SELECT MAX(CONSECUTIVO) FROM TB_ORACLE_COMPARACION_PEDIDO_AFECTACION", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux8.EOF Then
                              var_consecutivo = IIf(IsNull(rsaux8(0).Value), 0, rsaux8(0).Value) + 1
                           Else
                              var_consecutivo = 1
                           End If
                           rsaux8.Close
                           rsaux8.Open "insert into TB_ORACLE_COMPARACION_PEDIDO_AFECTACION (consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                           cnn.CommitTrans
                           rsaux8.Open "SELECT inte_emb_embarque, SOURCE_HEADER_NUMBER, SUM(FLOA_sal_cANTIDAD_LEIDA) AS CANTIDAD FROM XXVIA_TB_sALIDAS where source_header_number in (" + VAR_CADENA_PEDIDOS_M + ") GROUP BY inte_emb_embarque, SOURCE_HEADER_NUMBER", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           While Not rsaux8.EOF
                                 rsaux2.Open "select * from xxvia_tb_encabezado_embarques where embarque = " + CStr(rsaux8!inte_Emb_Embarque), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 If Not rsaux2.EOF Then
                                    rsaux10.Open "insert INTO TB_ORACLE_COMPARACION_PEDIDO_AFECTACION (PEDIDO, CANTIDAD_LEIDA, CANTIDAD_AFECTADA, fecha, CONSECUTIVO) VALUES (" + CStr(rsaux8!source_header_number) + "," + CStr(rsaux8!cantidad) + ",0, '" + CStr(rsaux2!FECHA_INICIO) + "'," + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                                 Else
                                    rsaux10.Open "insert INTO TB_ORACLE_COMPARACION_PEDIDO_AFECTACION (PEDIDO, CANTIDAD_LEIDA, CANTIDAD_AFECTADA, fecha, CONSECUTIVO) VALUES (" + CStr(rsaux8!source_header_number) + "," + CStr(rsaux8!cantidad) + ",0, ''," + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                                 End If
                                 rsaux2.Close
                                 rsaux8.MoveNext
                           Wend
                           rsaux8.Close
                           rsaux8.Open "SELECT inte_emb_embarque, SOURCE_HEADER_NUMBER, SUM(FLOA_sal_cANTIDAD_LEIDA) AS CANTIDAD FROM XXVIA_TB_sALIDAS_CAJAS where source_header_number in (" + VAR_CADENA_PEDIDOS_M + ") GROUP BY inte_emb_embarque, SOURCE_HEADER_NUMBER", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           While Not rsaux8.EOF
                                 rsaux2.Open "select * from xxvia_tb_encabezado_embarques where embarque = " + CStr(rsaux8!inte_Emb_Embarque), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 If Not rsaux2.EOF Then
                                    rsaux10.Open "insert INTO TB_ORACLE_COMPARACION_PEDIDO_AFECTACION (PEDIDO, CANTIDAD_LEIDA, CANTIDAD_AFECTADA, fecha, CONSECUTIVO) VALUES (" + CStr(rsaux8!source_header_number) + "," + CStr(rsaux8!cantidad) + ",0, '" + CStr(rsaux2!FECHA_INICIO) + "'," + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                                 Else
                                    rsaux10.Open "insert INTO TB_ORACLE_COMPARACION_PEDIDO_AFECTACION (PEDIDO, CANTIDAD_LEIDA, CANTIDAD_AFECTADA, fecha, CONSECUTIVO) VALUES (" + CStr(rsaux8!source_header_number) + "," + CStr(rsaux8!cantidad) + ",0, ''," + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                                 End If
                                 rsaux2.Close
                                 rsaux8.MoveNext
                           Wend
                           rsaux8.Close
                           rsaux8.Open "SELECT pedido FROM TB_ORACLE_COMPARACION_PEDIDO_AFECTACION WHERE CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                           While Not rsaux8.EOF
                                 rsaux1.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 rsaux10.Open "SELECT SOURCE_HEADER_NUMBER, SUM(SHIPPED_QUANTITY) AS CANTIDAD FROM WSH_DELIVERABLES_V WHERE SOURCE_HEADER_NUMBER = " + CStr(IIf(IsNull(rsaux8!pedido), 0, rsaux8!pedido)) + " GROUP BY SOURCE_HEADER_NUMBER", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 If Not rsaux10.EOF Then
                                    rsaux1.Open "UPDATE TB_ORACLE_COMPARACION_PEDIDO_AFECTACION SET CANTIDAD_AFECTADA = " + CStr(IIf(IsNull(rsaux10!cantidad), 0, rsaux10!cantidad)) + " WHERE PEDIDO = " + CStr(rsaux8!pedido) + " AND CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                                 End If
                                 rsaux10.Close
                                 rsaux8.MoveNext
                           Wend
                           rsaux8.Close
                           rsaux8.Open "SELECT *  FROM TB_ORACLE_COMPARACION_PEDIDO_AFECTACION where cantidad_afectada > 0 and CANTIDAD_LEIDA <> cantidad_afectada AND CONSECUTIVO = " + CStr(var_consecutivo) + " order by PEDIDO desc "
                           If Not rsaux8.EOF Then
                              var_cadena_pedidos_mal = ""
                              While Not rsaux8.EOF
                                    If var_cadena_pedidos_mal = "" Then
                                       var_cadena_pedidos_mal = CStr(IIf(IsNull(rsaux8!pedido), 0, rsaux8!pedido))
                                    Else
                                       var_cadena_pedidos_mal = var_cadena_pedidos_mal + ", " + CStr(IIf(IsNull(rsaux8!pedido), 0, rsaux8!pedido))
                                    End If
                                    rsaux8.MoveNext
                              Wend
                              MsgBox "Los siguientes pedidos tienen errores entra la cantidad leida y la cantidad afectada: " + CStr(var_cadena_pedidos_mal), vbOKOnly, "ATENCION"
                           Else
                              clnt.MSSoapInit "http://intranet/WsEBS12Prod/wsInterfaceOM.asmx?wsdl"
                              While Not rsaux.EOF
                                    rsaux2.Open "select distinct delivery_id from wsh_deliverables_v where SOURCE_HEADER_NUMBER = " + CStr(rsaux!source_header_number) + " AND delivery_id IS NOT NULL", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                    While Not rsaux2.EOF
                                          VAR_ENTREGA = rsaux2!delivery_id
                                          rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                          VAR_ESTATUS = 0
                                          On Error GoTo salirc:
                                          var_arreglo = clnt.ASIGNAR_embarque(VAR_ENTREGA, Val(VAR_X_TRIP_ID), "CONFIRM")
                                          rsaux1.Open "insert into tb_oracle_pedidos_confirmados (pedido, fecha, maquina, error) values (" + CStr(rsaux!source_header_number) + ", getdate(), '" + fun_NombrePc + "'," + CStr(VAR_ESTATUS) + ")", cnn, adOpenDynamic, adLockOptimistic
                                          rsaux2.MoveNext
                                    Wend
                                    rsaux2.Close
                                    rsaux.MoveNext
                              Wend
                              Set clnt = Nothing
                              MsgBox "Se termino de cerrar el embarque", vbOKOnly, "ATENCION"
                           End If
                           rsaux8.Close
                        End If
                        rsaux.Close
                     Else
                        If VAR_ESTATUS = "F" Then
                           MsgBox "EL embarque ya fue facturado"
                        Else
                           MsgBox "El embarque NO a sido cerrado", vbOKOnly, "ATENCION"
                        End If
                     End If
                  End If
                  rs.Close
               End If
               '--------------- fin de confirmar pedidos
               MsgBox "Se a cerrado el embarque", vbOKOnly, "ATENCION"
               Me.frm_sellos.Visible = False
               Me.txt_codigo.Enabled = False
            Else
               MsgBox "No se pudo crear el embarque en oracle", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "No se cerro el embarque", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El embarque ya habia sido cerrado", vbOKOnly, "ATENCION"
      End If
   Else
      rsaux.Open "DELETE FROM TB_TEMP_ORACLE_COMPARACION_CANTIDADES_LEIDAS_VS_PEDIDAS_AFECTADAS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " and codigo is null", cnn, adOpenDynamic, adLockOptimistic
      MsgBox "El embarque tiene diferencias entre las piezas pedidas y las leidas", vbOKOnly, "ATENCION"
      Set reporte = appl.OpenReport(App.Path + "\rep_oracle_diferencias_pedido_leido.rpt")
      var_cadena = "{TB_TEMP_ORACLE_COMPARACION_CANTIDADES_LEIDAS_VS_PEDIDAS_AFECTADAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
      reporte.RecordSelectionFormula = var_cadena
      frmvistasprevias.cr.ReportSource = reporte
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      frmvistasprevias.cr.ViewReport
      frmvistasprevias.Caption = "Reporte de diferencias pedido contra leido"
      frmvistasprevias.Show 1
      Set reporte = Nothing
      
      rsaux.Open "DELETE FROM TB_TEMP_ORACLE_COMPARACION_CANTIDADES_LEIDAS_VS_PEDIDAS_AFECTADAS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
   End If
   Exit Sub
salir2:
   'MsgBox Err.Description
   If Err.Number = -2147217900 Then
      'MsgBox Err.Description
      rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      Resume
   End If
salirc:
   If Err.Number = -2147467259 Then
      'MsgBox Err.Description
      Resume Next
      VAR_ESTATUS = 1
   End If
End Sub

Private Sub Command3_Click()
  If UCase(Me.Command3.Caption) = "DETENER" Then
     Me.Timer1.Enabled = False
     Me.Command3.Caption = "SEGUIR"
  Else
     If Me.Command3.Caption = "SEGUIR" Then
        Me.Timer1.Enabled = True
        Me.Command3.Caption = "DETENER"
     End If
  End If
  
End Sub


Private Sub cerrar_pedido()
Dim clnt As New SoapClient30
Dim clnt2 As New SoapClient30

Dim var_arreglo() As String
Dim var_container_id As String
Dim var_trip_id As String
Dim var_b As Boolean
'VAR_ESTATUS = "E"

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
   If rsaux10.State = 1 Then
      rsaux10.Close
   End If
   If CDbl(Me.lv_cajas.selectedItem.SubItems(1)) < 10000002 Then
   rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   'strconsulta = "SELECT * FROM XXVIA_COMP_LEIDA_VS_PED_AFEC WHERE inte_emb_embarque = ? and CANTIDAD_LEIDA > CANTIDAD_PEDIDA_AFECTADA and source_header_number = ?"
   strconsulta = "SELECT * FROM XXVIA_COMP_LEIDA_VS_PED_AFEC WHERE CANTIDAD_LEIDA > CANTIDAD_PEDIDA_AFECTADA and source_header_number = ?"
   With comandoORA
        .ActiveConnection = cnnoracle_4
        .CommandType = adCmdText
        .CommandText = strconsulta
        'Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_embarque))
        '.Parameters.Append parametro
        Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.lv_cajas.selectedItem.SubItems(1)))
        .Parameters.Append parametro
   End With
   Set rsaux3 = comandoORA.execute
   Set comandoORA = Nothing
   Set parametro = Nothing
   var_posible_cerrado_comparacion = 1
   If Not rsaux3.EOF Then
      var_posible_cerrado_comparacion = 0
      cnn.BeginTrans
      rs.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM  TB_TEMP_ORACLE_COMPARACION_CANTIDADES_LEIDAS_VS_PEDIDAS_AFECTADAS", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
      Else
         var_consecutivo = 1
      End If
      rs.Close
      rsaux.Open "INSERT INTO TB_TEMP_ORACLE_COMPARACION_CANTIDADES_LEIDAS_VS_PEDIDAS_AFECTADAS (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
      cnn.CommitTrans
      While Not rsaux3.EOF
            strconsulta = "SELECT * FROM XXVIA_TB_SALIDAS_CAJAS WHERE inte_emb_embarque = ? and SOURCE_HEADER_NUMBER = ? AND SEGMENT1 = ? AND FLOA_SAL_cANTIDAD_LEIDA > 0"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rsaux3!inte_Emb_Embarque))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rsaux3!source_header_number))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, rsaux3!segment1)
                 .Parameters.Append parametro
            End With
            Set rsaux2 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            VAR_CAJAS = ""
            While Not rsaux2.EOF
                  If VAR_CAJAS = "" Then
                     VAR_CAJAS = "CAJA: " + CStr(IIf(IsNull(rsaux2!INTE_PAQ_CAJA), 0, rsaux2!INTE_PAQ_CAJA)) + " CANTIDAD: " + CStr(IIf(IsNull(rsaux2!floa_Sal_cantidad_leida), 0, rsaux2!floa_Sal_cantidad_leida))
                  Else
                     VAR_CAJAS = VAR_CAJAS + ", CAJA: " + CStr(IIf(IsNull(rsaux2!INTE_PAQ_CAJA), 0, rsaux2!INTE_PAQ_CAJA)) + " CANTIDAD: " + CStr(IIf(IsNull(rsaux2!floa_Sal_cantidad_leida), 0, rsaux2!floa_Sal_cantidad_leida))
                  End If
                  rsaux2.MoveNext
            Wend
            rsaux2.Close
            var_cadena = "INSERT INTO TB_TEMP_ORACLE_COMPARACION_CANTIDADES_LEIDAS_VS_PEDIDAS_AFECTADAS (INTE_TEM_CONSECUTIVO, EMBARQUE, PEDIDO, CODIGO, DESCRIPCION, CANTIDAD_PEDIDA, CANTIDAD_LEIDA, CAJAS )"
            var_cadena = var_cadena + " VALUES (" + CStr(var_consecutivo) + "," + CStr(rsaux3!inte_Emb_Embarque) + "," + CStr(rsaux3!source_header_number) + ",'" + rsaux3!segment1 + "', '" + rsaux3!item_description + "'," + CStr(rsaux3!CANTIDAD_PEDIDA_AFECTADA) + "," + CStr(rsaux3!cantidad_leida) + ",'" + VAR_CAJAS + "')"
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            rsaux3.MoveNext
      Wend
   End If
   rsaux3.Close
   If var_posible_cerrado_comparacion = 1 Then
      VAR_ESTATUS = "I"
      If VAR_ESTATUS = "I" Then
         var_si = MsgBox("¿Desea cerrar el pedido?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_si = MsgBox("Confirmar el cerrado del pedido", vbYesNo, "ATENCION")
            If var_si = 6 Then
               x = 1
            Else
               x = 0
            End If
         Else
            x = 0
         End If
         If x = 1 Then
            rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
            VAR_X_TRIP_ID = rs!ARREGLO_0
            var_x_trip_name = rs!ARREGLO_1
            rs.Close
            If "X" <> "" Then
               rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If rs!tipo_embarque = 2 Then
                  rsaux.Open "select distinct source_header_number from xxvia_tb_salidas_CAJAS where inte_emb_embarque = " + Me.txt_embarque + " and source_header_number = " + Me.lv_cajas.selectedItem.SubItems(1), cnnoracle_4, adOpenDynamic, adLockOptimistic
               End If
               var_Cadena_pedidos = ""
               var_j = 0
               While Not rsaux.EOF
                     If var_Cadena_pedidos = "" Then
                        var_Cadena_pedidos = "'" + CStr(rsaux!source_header_number) + "'"
                     Else
                        var_Cadena_pedidos = var_Cadena_pedidos + ", '" + CStr(rsaux!source_header_number) + "'"
                     End If
                     var_j = var_j + 1
                     rsaux.MoveNext
               Wend
               rsaux.Close
               
               var_cadena = "SELECT HCAS.CUST_ACCOUNT_ID, HCAS.CUST_ACCT_SITE_ID as customer_id, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND to_number(source_header_number) IN (" + var_Cadena_pedidos + ")"
               var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID  AND released_status = 'Y'"
               rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               'rsaux2.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux!cust_account_id), cnnoracle_4, adOpenDynamic, adLockOptimistic
               'VAR_AGENTE_str = rsaux2!COLLECTOR_ID
               'var_nombre_agente_str = rsaux2!Name
               'rsaux2.Close
               While Not rsaux.EOF
                     'rsaux3.Open "SELECT * FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " AND source_header_number = " + CStr(CDbl(rsaux!SOURCE_HEADER_NUMBER)) + " AND DELIVERY_DETAIL_ID = " + CStr(rsaux!DELIVERY_DETAIL_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     'strconsulta = "SELECT * FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = ? AND source_header_number = ? AND DELIVERY_DETAIL_ID = ?"
                     strconsulta = "SELECT * FROM XXVIA_TB_SALIDAS_CAJAS WHERE source_header_number = ? AND DELIVERY_DETAIL_ID = ?"
                     With comandoORA
                         .ActiveConnection = cnnoracle_4
                         .CommandType = adCmdText
                         .CommandText = strconsulta
                         'Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_embarque))
                         '.Parameters.Append parametro
                         Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rsaux!source_header_number))
                         .Parameters.Append parametro
                         Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(rsaux!delivery_detail_id))
                         .Parameters.Append parametro
                     End With
                     Set rsaux3 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     If rsaux3.EOF Then
                        var_cadena = "INSERT INTO XXVIA_TB_SALIDAS_CAJAS (INTE_EMB_EMBARQUE, SOURCE_HEADER_NUMBER, SEGMENT1, FLOA_SAL_CANTIDAD_LEIDA, INVENTORY_ITEM_ID, DELIVERY_DETAIL_ID, SOURCE_LINE_NUMBER, DELIVERY_ID, INTE_PAQ_CAJA, CUSTOMER_ID, SUBINVENTORY, NAME, COLLECTOR_ID, ITEM_DESCRIPTION, CUSTOMER_NAME, char_paq_estatus)"
                        var_cadena = var_cadena + " values (" + Me.txt_embarque + "," + CStr(CDbl(rsaux!source_header_number)) + ",'" + rsaux!segment1 + "',0," + CStr(rsaux!inventory_item_id) + "," + CStr(rsaux!delivery_detail_id) + ",'" + CStr(rsaux!SOURCE_LINE_NUMBER) + "'," + CStr(IIf(IsNull(rsaux!delivery_id), 0, rsaux!delivery_id)) + ",0," + CStr(rsaux!CUSTOMER_ID) + ",'" + CStr(IIf(IsNull(rsaux!subinventory), "", rsaux!subinventory)) + "', '" + var_nombre_agente_str + "','" + CStr(VAR_AGENTE_str) + "','" + CStr(rsaux!Description) + "','" + rsaux!customer_name + "','S')"
                        rsaux4.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux3.Close
                     rsaux.MoveNext
               Wend
               rsaux.Close
               If rsaux9.State = 1 Then
                  rsaux9.Close
               End If
               x = 1
               If x = 0 Then
                  rsaux9.Open "select * from tb_unidadesorganizacionales where vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux9.EOF Then
                     VAR_USER_ID = rsaux9!user_id
                     VAR_RESP_ID = rsaux9!resp_id
                     VAR_RESP_APPL_ID = rsaux9!resp_appl_id
                  End If
                  rsaux9.Close
                  var_cadena = "SELECT a.source_line_id, OHA.HEADER_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, E.NAME from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND A.SOURCE_HEADER_NUMBER in (" + var_Cadena_pedidos + ")"
                  var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id = e.collector_id AND released_status = 'B' order by A.source_header_number"
                  rsaux9.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  While Not rsaux9.EOF
                        rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        'MsgBox "call XXVIA_SP_DEPURA_ORDEN_SURTIDO (" + CStr(CDbl(rsaux9!header_id)) + ", " + CStr(CDbl(rsaux9!SOURCE_LINE_ID)) + ", 'PRODUCCION')"
                        On Error GoTo salir2:
                        rsaux7.Open "call XXVIA_SP_DEPURA_ORDEN_SURTIDO (" + CStr(CDbl(rsaux9!header_id)) + ", " + CStr(CDbl(rsaux9!source_LINE_ID)) + ", 'PRODUCCION'," + CStr(VAR_USER_ID) + "," + CStr(VAR_RESP_ID) + "," + CStr(VAR_RESP_APPL_ID) + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        rsaux9.MoveNext
                  Wend
                  rsaux9.Close
                  rs.Close
               End If
               clnt.MSSoapInit var_webservice
               If rs.State = 1 Then
                  rs.Close
               End If
               x = 0
               If x = 1 Then
                  'rs.Open "SELECT delivery_detail_id, sum(floa_sal_Cantidad_leida) as floa_sal_Cantidad_leida FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " and source_header_number = " + Me.lv_cajas.selectedItem.SubItems(1) + " group by delivery_detail_id", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  'cambio por pedidos divididos
                  rs.Open "SELECT delivery_detail_id, sum(floa_sal_Cantidad_leida) as floa_sal_Cantidad_leida FROM XXVIA_TB_SALIDAS_CAJAS WHERE source_header_number = " + Me.lv_cajas.selectedItem.SubItems(1) + " group by delivery_detail_id", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  VAR_ZZ = 0
                  While Not rs.EOF
                        VAR_ZZ = VAR_ZZ + 1
                        'VAR_ZZZ = 0
                        If VAR_ZZ = 50 Then
                           var_zzz = var_zzz + 1
                           VAR_ZZ = 0
                        End If
                        'rsaux.Open "SELECT * FROM WSH_DELIVERABLES_V WHERE delivery_detail_id = " + CStr(rs!DELIVERY_DETAIL_ID) + " AND RELEASED_STATUS = 'Y'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        strconsulta = "SELECT * FROM WSH_DELIVERABLES_V WHERE delivery_detail_id = ? AND RELEASED_STATUS = 'Y'"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!delivery_detail_id)
                             .Parameters.Append parametro
                        End With
                        Set rsaux = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                     
                        If Not rsaux.EOF Then
                           'var_b = clnt.actualizar_detalle(Val(rs!delivery_detail_id), CDbl(rs!FLOA_sAL_cANTIDAD_LEIDA), "OE", 0)
                           On Error GoTo salir2:
                           rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           'rsaux6.Open "select max(inte_paq_caja) as inte_paq_caja  from xxvia_tb_Salidas_cajas where delivery_detail_id = " + CStr(rs!DELIVERY_DETAIL_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                           strconsulta = "select max(inte_paq_caja) as inte_paq_caja  from xxvia_tb_Salidas_cajas where delivery_detail_id = ?"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!delivery_detail_id)
                                .Parameters.Append parametro
                           End With
                           Set rsaux6 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           var_consecutivo = rsaux6!INTE_PAQ_CAJA
                           rsaux6.Close
                           rsaux6.Open "CALL xxvia_pk_interfaces_om.actualizar_detalle (1.0, " + CStr(rs!delivery_detail_id) + "," + CStr(rs!floa_Sal_cantidad_leida) + ",'OE'," + CStr(var_consecutivo) + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        End If
                        rsaux.Close
                        rs.MoveNext
                  Wend
                  rs.Close
               Else
' METODO DE ACTUALIZACION DE DETALLE DE PEDIDOS NUEVO

                  On Error GoTo salir2:
                  
                  var_cadena = "select distinct source_header_number from xxvia_tb_salidas_cajas where  SOURCE_HEADER_NUMBER = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = var_cadena
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.lv_cajas.selectedItem.SubItems(1)))
                       .Parameters.Append parametro
                  End With
                  Set rsaux6 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                                          
                  While Not rsaux6.EOF
                        On Error GoTo salir2:
                        var_cadena = "call xxvia_sp_act_det_pedido_2 (?)"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = var_cadena
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rsaux6!source_header_number))
                             .Parameters.Append parametro
                        End With
                        Set rsaux7 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        
                        var_cadena = "select order_type_id from oe_order_headers_all where order_number = ?"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = var_cadena
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rsaux6!source_header_number))
                             .Parameters.Append parametro
                        End With
                        Set rsaux7 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        If rsaux7!order_type_id = 1002 Or rsaux7!order_type_id = 1023 Then
                           rsaux9.Open "SELECT * FROM TB_ORACLE_PEDIDOS_CERRADOS_CN WHERE PEDIDO = " + CStr(rsaux6!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                           If rsaux9.EOF Then
                              rsaux10.Open "INSERT INTO TB_ORACLE_PEDIDOS_CERRADOS_CN (PEDIDO) VALUES ('" + CStr(rsaux6!source_header_number) + "')", cnn, adOpenDynamic, adLockOptimistic
                           End If
                           rsaux9.Close
                           
                           rsaux7.Close
                           var_cadena = "select  A.SECONDARY_INVENTORY_NAME, A.DESCRIPTION, ADDRESS_LINE_1, ADDRESS_LINE_2, TOWN_OR_CITY, REGION_1, COUNTRY, POSTAL_CODE, loc_information13 EMAIL from mtl_secondary_inventories a, hr_locations_all b, xxvia_jv_tb_agentes c, po_requisition_headers_ALL D, OE_ORDER_HEADERS_ALL E Where A.location_id = b.location_id and a.secondary_inventory_name = c.subinventory_code AND E.source_document_id = D.requisition_header_id AND A.secondary_inventory_name = D.ATTRIBUTE1 AND E.ORDER_NUMBER = ?                 "
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = var_cadena
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rsaux6!source_header_number))
                                .Parameters.Append parametro
                           End With
                           Set rsaux7 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           If Not rsaux7.EOF Then
                              var_contingencia = 0
                              If var_contingencia = 1 Then
                              VAR_CORREO = ""
                              Else
                              VAR_CORREO = IIf(IsNull(rsaux7!Email), "", rsaux7!Email)
                              End If
                              If VAR_CORREO <> "" Then
                                 ' SE ENVIA CORREO A TIENDA
                                 clnt2.MSSoapInit "http://serviciowebcedisdesa.vianney.com.mx/EnviarCorreos.asmx?wsdl"
                                 var_s = clnt2.CorreoAdjunto(VAR_CORREO, "Packing List del pedido " + CStr(rsaux6!source_header_number), "Se anexa packing list de pedido " + CStr(rsaux6!source_header_number) + " del CN. " + rsaux7!Description, CStr(rsaux6!source_header_number), "1002")
                              End If
                           End If
                           rsaux7.Close
                        Else
                           rsaux9.Open "SELECT * FROM TB_ORACLE_PEDIDOS_CERRADOS WHERE PEDIDO = " + CStr(rsaux6!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                           If rsaux9.EOF Then
                              rsaux10.Open "INSERT INTO TB_ORACLE_PEDIDOS_CERRADOS (PEDIDO, REQUEST_ID) VALUES (" + CStr(rsaux6!source_header_number) + ",0)", cnn, adOpenDynamic, adLockOptimistic
                           End If
                           rsaux9.Close
                           
                           var_cadena = "select razon_social_cliente as description, email_Address as email from xxvia_vw_clientes_bcp a, oe_order_headers_all b where order_number = ? and a.site_use_id = b.invoice_to_org_id"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = var_cadena
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rsaux6!source_header_number))
                                .Parameters.Append parametro
                           End With
                           Set rsaux7 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           If Not rsaux7.EOF Then
                              VAR_CORREO = IIf(IsNull(rsaux7!Email), "", rsaux7!Email)
                              If VAR_CORREO <> "" Then
                                 ' SE ENVIA CORREO A TIENDA
                                 clnt2.MSSoapInit "http://serviciowebcedisdesa.vianney.com.mx/EnviarCorreos.asmx?wsdl"
                                 var_s = clnt2.CorreoAdjunto(VAR_CORREO, "Packing List del pedido " + CStr(rsaux6!source_header_number), "Se anexa packing list de pedido " + CStr(rsaux6!source_header_number) + " del cliente " + rsaux7!Description, CStr(rsaux6!source_header_number), "9999")
                              End If
                           End If
                           rsaux7.Close
                                                     
                                                     
                                                     
                        End If
                        
                        rsaux6.MoveNext
                  Wend
                  rsaux6.Close
                  Set comandoORA = Nothing
                  Set parametro = Nothing

'FIN DE METODO DE ACTUALIZACION DE DETALLE DE PEDIDOS NUEVO
               End If
               Set clnt = Nothing
               'clnt.MSSoapInit var_webservice
               'rs.Open "SELECT DISTINCT DELIVERY_ID FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
               'While Not rs.EOF
               '
               '      var_arreglo = clnt.ASIGNAR_embarque(rs!delivery_id, Val(VAR_X_TRIP_ID), "CONFIRM")
               '      rs.MoveNext
               'Wend
               'rs.Close
               'Set clint = Nothing
               
               var_cadena = "SELECT b.vcha_tit_titular_id, D.vcha_esb_establecimient_id FROM XXVIA_VW_CLIENTES_PEDIDOS B, OE_ORDER_HEADERS_ALL C, XXVIA_VW_ESTABLECIMIENTOS_PED D, XXVIA_VW_ESTABLECIMIENTOS_PED E Where c.order_number = ? AND c.SOLD_TO_ORG_ID = B.CUST_ACCOUNT_ID AND D.SITE_USE_ID    = C.SHIP_TO_ORG_ID AND E.SITE_USE_ID    = C.INVOICE_TO_ORG_ID"
               'MsgBox var_cadena
               strconsulta = var_cadena
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.lv_cajas.selectedItem.SubItems(1))
                    .Parameters.Append parametro
               End With
               Set rsaux9 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               
               
               var_Titular_vianney_catalog = rsaux9!vcha_tit_titular_id
               var_si_vianney_Catalog = 0
               var_almacen_Destino = rsaux9!vcha_esb_establecimient_id
                
               If var_Titular_vianney_catalog = "T000000343" Then
                  var_si_vianney_Catalog = 1
               End If
               
               var_dia = CStr(Day(CDate(Date)))
               var_mes = CStr(Month(CDate(Date)))
               var_año = CStr(Year(CDate(Date)))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               If Len(Trim(var_hora)) = 1 Then
                  var_hora = "0" + var_hora
               End If
                  
                  
               var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
               
               x = 0
               If x = 1 Then
               If var_si_vianney_Catalog = 1 Then
                  rs.Open "SELECT floa_sal_Cantidad_leida as cantidad, organizacion,  inte_emb_embarque as embarque, inte_paq_caja as caja, source_header_number  as pedido, a.segment1 as codigo, collector_id as agente,name as nombre_agente, customer_id as cliente, customer_name as nombre_cliente, a.inventory_item_id, caja_pedido, sello, UNIT_WEIGHT as peso,  item_description as descripcion    FROM XXVIA_TB_salidas_cajas a, xxvia_tb_encabezado_embarques, xxvia_system_items_b b, oe_order_headers_all oh where inte_emb_embarque = embarque and organizacion = b.organization_id and a.inventory_item_id = b.inventory_item_id and order_number = a.source_header_number and oh.ship_from_org_id = organizacion and inte_emb_embarque = " + Me.txt_embarque + " and floa_sal_Cantidad_leida >0 AND a.source_header_number = " + Me.lv_cajas.selectedItem.SubItems(1), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     strconsulta = "select SECONDARY_INVENTORY_NAME from mtl_secondary_inventories where ATTRIBUTE8 = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_almacen_Destino)
                          .Parameters.Append parametro
                      End With
                      Set rsaux = comandoORA.execute
                      Set comandoORA = Nothing
                      Set parametro = Nothing
                      If Not rsaux.EOF Then
                         var_almacen_icg = rsaux(0).Value
                      Else
                         var_almacen_icg = ""
                      End If
                     rsaux.Close
                     rsaux.Open "alter session set nls_language= 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     While Not rs.EOF
                
                           var_cadena = "SELECT * FROM OPENQUERY(ICGCENTRAL, 'SELECT * FROM BDICGDESA_USA.DBO.IT_PEDIDOCOMPRA where no_embarque = ''" + Me.txt_embarque + "'' and no_pedido = ''" + CStr(rs!pedido) + "'' and no_caja =  ''" + CStr(rs!Caja) + "'' and codigo = ''" + rs!codigo + "''')A"
                           rsaux10.Open var_cadena, cnn_icg_usa, adOpenDynamic, adLockOptimistic
                           If rsaux10.EOF Then
                           
                              var_pedido = rs!pedido
                              strconsulta = "select unit_selling_price from oe_order_headers_all oh, oe_order_lines_all ol where order_number = ? and oh.header_id = ol.header_id and oh.ship_from_org_id = ? and ol.inventory_item_id = ?"
                              With comandoORA
                                   .ActiveConnection = cnnoracle_4
                                   .CommandType = adCmdText
                                   .CommandText = strconsulta
                                   Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(var_pedido))
                                   .Parameters.Append parametro
                                   Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(var_unidad_organizacional))
                                   .Parameters.Append parametro
                                   Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, IIf(IsNull(rs!inventory_item_id), 0, rs!inventory_item_id))
                                   .Parameters.Append parametro
                              End With
                              Set rsaux = comandoORA.execute
                              Set comandoORA = Nothing
                              Set parametro = Nothing
                           
                              If Not rsaux.EOF Then
                                 var_precio = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
                              Else
                                 var_precio = 0
                              End If
                              rsaux.Close
                           
                           
                           
                           
                           
                              rsaux11.Open "insert into IT_PEDIDOCOMPRA (OU, SUBINVENTORY_CODE, TRANSFER_SUBINVENTORY, FECHA, NO_EMBARQUE, NO_PEDIDO, NO_CAJA, CODIGO, CANTIDAD, PRECIO, DESCRIPCION) values (381,'" + var_almacen_icg + "','CDI_ALMPT', " + var_fecha + ", " + Me.txt_embarque + ",'" + CStr(rs!pedido) + "','" + CStr(rs!Caja) + "','" + rs!codigo + "'," + CStr(rs!cantidad) + "," + CStr(var_precio) + ",'" + rs!descripcion + "')", cnn, adOpenDynamic, adLockOptimistic
                           End If
                           rsaux10.Close
                           rs.MoveNext
                     Wend
                     rs.Close
                     rs.Open "INSERT INTO ICGCENTRAL.BDICGDESA_USA.DBO.IT_PEDIDOCOMPRA (OU, SUBINVENTORY_CODE, TRANSFER_SUBINVENTORY, FECHA, NO_EMBARQUE, NO_PEDIDO, NO_CAJA, CODIGO, CANTIDAD, PRECIO, DESCRIPCION,status ) select OU, SUBINVENTORY_CODE, TRANSFER_SUBINVENTORY, FECHA, NO_EMBARQUE, NO_PEDIDO, NO_CAJA, CODIGO, CANTIDAD, PRECIO, DESCRIPCION,3 from IT_PEDIDOCOMPRA where NO_EMBARQUE = " + Me.txt_embarque + " AND NO_PEDIDO = " + CStr(var_pedido), cnn_icg_usa, adOpenDynamic, adLockOptimistic
                     'rs.Open "INSERT OPENQUERY (icgcentral, 'SELECT OU, SUBINVENTORY_CODE, TRANSFER_SUBINVENTORY, FECHA, NO_EMBARQUE, NO_PEDIDO, NO_CAJA, CODIGO, CANTIDAD, PRECIO, DESCRIPCION FROM SIDAlmacenBkp_USA.DBO.it_pedidocompra') select OU, SUBINVENTORY_CODE, TRANSFER_SUBINVENTORY, FECHA, NO_EMBARQUE, NO_PEDIDO, NO_CAJA, CODIGO, CANTIDAD, PRECIO, DESCRIPCION from IT_PEDIDOCOMPRA where NO_EMBARQUE = " + Me.txt_embarque + " AND NO_PEDIDO = " + CStr(var_pedido), cnnicg_sql, adOpenDynamic, adLockOptimistic
                  End If
               End If
               End If
               
               'rs.Open "SELECT * FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " and source_header_number = " + Me.lv_cajas.selectedItem.SubItems(1), cnnoracle_4, adOpenDynamic, adLockOptimistic
               rs.Open "SELECT * FROM XXVIA_TB_SALIDAS_CAJAS WHERE source_header_number = " + Me.lv_cajas.selectedItem.SubItems(1), cnnoracle_4, adOpenDynamic, adLockOptimistic
               While Not rs.EOF
                     If IIf(IsNull(rs!floa_Sal_cantidad_leida), 0, rs!floa_Sal_cantidad_leida) > 0 Then
                        var_cadena = "INSERT INTO XXVIA_TB_DETALLE_CAJAS (EMBARQUE, PEDIDO,AGENTE, NOMBRE_AGENTE,CLIENTE,NOMBRE_CLIENTE,CODIGO, DESCRIPCION, CANTIDAD, PESO, CAJA, INVENTORY_ITEM_ID, CAJA_PEDIDO)"
                        var_cadena = var_cadena + " values (" + Me.txt_embarque + ", " + CStr(rs!source_header_number) + ",'" + CStr(IIf(IsNull(rs!collector_id), 0, rs!collector_id)) + "', '" + IIf(IsNull(rs!Name), "", rs!Name) + "',  '" + CStr(rs!CUSTOMER_ID) + "','" + IIf(IsNull(rs!customer_name), "", rs!customer_name) + "','" + rs!segment1 + "','" + rs!item_description + "'," + CStr(rs!floa_Sal_cantidad_leida) + ",0," + CStr(rs!INTE_PAQ_CAJA) + "," + CStr(rs!inventory_item_id) + "," + CStr(IIf(IsNull(rs!caja_pedido), 0, rs!caja_pedido)) + ")"
                       rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     End If
                     rs.MoveNext
               Wend
               rs.Close
               'rs.Open "UPDATE XXVIA_TB_ENCABEZADO_EMBARQUES SET CHAR_EMB_ESTATUS = 'I', FECHA_FIN = SYSDATE, USUARIO_CERRO = '" + var_clave_usuario_global + "' WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
               'rs.Open "UPDATE TB_ORACLE_EMBARQUES_ORDENES SET estatus = 'I' WHERE inte_emb_EMBARQUE = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
               x = 0
               If x = 1 Then
                  rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     If IIf(IsNull(rs!char_emb_estatus), "", rs!char_emb_estatus) = "I" Then
                        If rs!tipo_embarque = 2 Then
                           rsaux.Open "select distinct source_header_number from xxvia_tb_salidas_cAJAS where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        End If
                        var_Cadena_pedidos = ""
                        var_j = 0
                        While Not rsaux.EOF
                              If var_Cadena_pedidos = "" Then
                                 var_Cadena_pedidos = "'" + CStr(rsaux!source_header_number) + "'"
                              Else
                                 var_Cadena_pedidos = var_Cadena_pedidos + ", '" + CStr(rsaux!source_header_number) + "'"
                              End If
                              var_j = var_j + 1
                              rsaux.MoveNext
                        Wend
                        rsaux.Close
                        var_i = 0
                        If var_i = 1 Then
                           While var_j <> var_i
                                 var_i = 0
                                 var_cadena = "SELECT e.collector_id, A.SOURCE_HEADER_NUMBER,  HL.ADDRESS1 AS CUSTOMER_NAME,  A.released_status,  E.NAME , sum(shipped_quantity) as cantidad from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, hz_customer_profiles D, ar_collectors e Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id "
                                 var_cadena = var_cadena + " AND A.SOURCE_HEADER_NUMBER in (" + var_Cadena_pedidos + ") AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND D.collector_id = e.collector_id AND released_status = 'C' group by  e.collector_id, A.SOURCE_HEADER_NUMBER, HL.ADDRESS1,  A.released_status,  E.NAME"
                                 'MsgBox var_cadena_pedidos
                                 rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 While Not rsaux.EOF
                                       var_i = var_i + 1
                                       rsaux.MoveNext
                                 Wend
                                 rsaux.Close
                           Wend
                           x = 1
                           If x = 0 Then
                              var_cadena_pedidos_global = var_Cadena_pedidos
                              var_cadena = "SELECT a.source_line_id, OHA.HEADER_ID from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND A.SOURCE_HEADER_NUMBER in (" + var_cadena_pedidos_global + ") "
                              var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id = e.collector_id AND released_status = 'B' order by A.source_header_number"
                              rsaux7.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux7.EOF Then
                                 var_tipo_depurado = 1
                                 frmoracle_depurar_pedidos.Show 1
                              End If
                              rsaux7.Close
                              var_tipo_depurado = 0
                              var_cadena = "SELECT a.source_line_id, OHA.HEADER_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, E.NAME from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND A.SOURCE_HEADER_NUMBER in (" + var_cadena_pedidos_global + ")"
                              var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id = e.collector_id AND released_status = 'B' order by A.source_header_number"
                              rsaux9.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux9.EOF Then
                                 rsaux9.Close
                                 var_sigue = 1
                                 While var_sigue = 1
                                       If rsaux8.State = 1 Then
                                          rsaux8.Close
                                       End If
                                       var_cadena = "SELECT a.source_line_id, OHA.HEADER_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, E.NAME from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND A.SOURCE_HEADER_NUMBER in (" + var_cadena_pedidos_global + ")"
                                       var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id = e.collector_id AND released_status = 'B' order by A.source_header_number"
                                       rsaux8.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                       If rsaux8.EOF Then
                                          var_sigue = 0
                                       Else
                                          While Not rsaux8.EOF
                                                rsaux7.Open "SELECT * FROM TB_ORACLE_NEGADO WHERE PEDIDO IN (" + CStr(rsaux8!source_header_number) + ") AND INVENTORY_ITEM_ID = " + CStr(rsaux8!inventory_item_id), cnn, adOpenDynamic, adLockOptimistic
                                                rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                                Set clnt = Nothing
                                                clnt.MSSoapInit var_webservice
                                                var_s = clnt.cancelar_back_order(CDbl(rsaux8!header_id), CDbl(rsaux8!source_LINE_ID), rsaux7!CAUSA_NEGADO)
                                                Set clnt = Nothing
                                                rsaux7.Close
                                                rsaux8.MoveNext
                                          Wend
                                       End If
                                       rsaux8.Close
                                 Wend
                              Else
                                 rsaux9.Close
                              End If
                           End If 'x
                        End If
                     End If
                  End If
               End If
               '--------------- confirmar pedidos
               
               x = 1
               If x = 1 Then
                  rsaux.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     VAR_X_TRIP_ID = IIf(IsNull(rs!ARREGLO_0), 0, rs!ARREGLO_0)
                     var_x_trip_name = IIf(IsNull(rs!ARREGLO_1), 0, rs!ARREGLO_1)
                     VAR_ESTATUS = IIf(IsNull(rs!char_emb_estatus), "", rs!char_emb_estatus)
                     'If IIf(IsNull(rs!char_emb_estatus), "", rs!char_emb_estatus) = "I" Then
                     If "I" = "I" Then
                        If rs!tipo_embarque = 1 Then
                           rsaux.Open "select distinct source_header_number from xxvia_tb_salidas where inte_emb_embarque = " + Me.txt_embarque + " and source_header_number = " + Me.lv_cajas.selectedItem.SubItems(1), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        End If
                        If rs!tipo_embarque = 2 Then
                           'rsaux.Open "select distinct source_header_number from xxvia_tb_SAlidas_cajas where inte_emb_embarque = " + Me.txt_embarque + " and source_header_number = " + Me.lv_cajas.selectedItem.SubItems(1), cnnoracle_4, adOpenDynamic, adLockOptimistic
                           rsaux.Open "select distinct source_header_number from xxvia_tb_SAlidas_cajas where source_header_number = " + Me.lv_cajas.selectedItem.SubItems(1), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        End If
                        VAR_CADENA_PEDIDOS_M = ""
                        While Not rsaux.EOF
                              If VAR_CADENA_PEDIDOS_M = "" Then
                                 VAR_CADENA_PEDIDOS_M = CStr(rsaux!source_header_number)
                              Else
                                 VAR_CADENA_PEDIDOS_M = VAR_CADENA_PEDIDOS_M + ", " + CStr(rsaux!source_header_number)
                              End If
                              rsaux.MoveNext
                        Wend
                        var_Cadena_pedidos = ""
                        rsaux.MoveFirst
                        While Not rsaux.EOF
                              rsaux1.Open "select distinct delivery_id from wsh_deliverables_v where SOURCE_HEADER_NUMBER = " + CStr(rsaux!source_header_number) + " AND delivery_id IS NOT NULL", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              VAR_ENTREGA = rsaux1!delivery_id
                              rsaux1.Close
                              rsaux1.Open "select distinct source_header_number from wsh_deliverables_v where delivery_id = " + CStr(VAR_ENTREGA), cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux1.EOF Then
                                 var_j = 0
                                 While Not rsaux1.EOF
                                       var_j = var_j + 1
                                       rsaux1.MoveNext
                                 Wend
                                 If var_j > 1 Then
                                    If var_Cadena_pedidos = "" Then
                                       var_Cadena_pedidos = CStr(rsaux!source_header_number) + " ENTREGA: " + CStr(VAR_ENTREGA)
                                    Else
                                       var_Cadena_pedidos = var_Cadena_pedidos + ", " + CStr(rsaux!source_header_number) + " ENTREGA: " + CStr(VAR_ENTREGA)
                                    End If
                                 End If
                              End If
                              rsaux1.Close
                              rsaux.MoveNext
                        Wend
                        rsaux.MoveFirst
                        If var_Cadena_pedidos <> "" Then
                           MsgBox "Los pedidos siguientes tienen dos entregas " + var_Cadena_pedidos
                        Else
                           cnn.BeginTrans
                           rsaux8.Open "SELECT MAX(CONSECUTIVO) FROM TB_ORACLE_COMPARACION_PEDIDO_AFECTACION", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux8.EOF Then
                              var_consecutivo = IIf(IsNull(rsaux8(0).Value), 0, rsaux8(0).Value) + 1
                           Else
                              var_consecutivo = 1
                           End If
                           rsaux8.Close
                           rsaux8.Open "insert into TB_ORACLE_COMPARACION_PEDIDO_AFECTACION (consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                           cnn.CommitTrans
                           'rsaux8.Open "SELECT inte_emb_embarque, SOURCE_HEADER_NUMBER, SUM(FLOA_sal_cANTIDAD_LEIDA) AS CANTIDAD FROM XXVIA_TB_sALIDAS where source_header_number in (" + VAR_CADENA_PEDIDOS_M + ") GROUP BY inte_emb_embarque, SOURCE_HEADER_NUMBER", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           'cambio por pedido dividido
                           rsaux8.Open "SELECT SOURCE_HEADER_NUMBER, SUM(FLOA_sal_cANTIDAD_LEIDA) AS CANTIDAD FROM XXVIA_TB_sALIDAS where source_header_number in (" + VAR_CADENA_PEDIDOS_M + ") GROUP BY SOURCE_HEADER_NUMBER", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           While Not rsaux8.EOF
                                 rsaux2.Open "select * from xxvia_tb_encabezado_embarques where embarque = " + CStr(rsaux8!inte_Emb_Embarque), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 If Not rsaux2.EOF Then
                                    rsaux10.Open "insert INTO TB_ORACLE_COMPARACION_PEDIDO_AFECTACION (PEDIDO, CANTIDAD_LEIDA, CANTIDAD_AFECTADA, fecha, CONSECUTIVO) VALUES (" + CStr(rsaux8!source_header_number) + "," + CStr(rsaux8!cantidad) + ",0, '" + CStr(rsaux2!FECHA_INICIO) + "'," + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                                 Else
                                    rsaux10.Open "insert INTO TB_ORACLE_COMPARACION_PEDIDO_AFECTACION (PEDIDO, CANTIDAD_LEIDA, CANTIDAD_AFECTADA, fecha, CONSECUTIVO) VALUES (" + CStr(rsaux8!source_header_number) + "," + CStr(rsaux8!cantidad) + ",0, ''," + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                                 End If
                                 rsaux2.Close
                                 rsaux8.MoveNext
                           Wend
                           rsaux8.Close
                           'rsaux8.Open "SELECT inte_emb_embarque, SOURCE_HEADER_NUMBER, SUM(FLOA_sal_cANTIDAD_LEIDA) AS CANTIDAD FROM XXVIA_TB_sALIDAS_CAJAS where source_header_number in (" + VAR_CADENA_PEDIDOS_M + ") GROUP BY inte_emb_embarque, SOURCE_HEADER_NUMBER", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           'cambio por pedido dividido
                           rsaux8.Open "SELECT SOURCE_HEADER_NUMBER, SUM(FLOA_sal_cANTIDAD_LEIDA) AS CANTIDAD FROM XXVIA_TB_sALIDAS_CAJAS where source_header_number in (" + VAR_CADENA_PEDIDOS_M + ") GROUP BY SOURCE_HEADER_NUMBER", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           While Not rsaux8.EOF
                                 'rsaux2.Open "select * from xxvia_tb_encabezado_embarques where embarque = " + CStr(rsaux8!inte_Emb_Embarque), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 'cambio por pedido dividido
                                 rsaux2.Open "select * from xxvia_tb_encabezado_embarques where embarque = " + CStr(Me.txt_embarque), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 If Not rsaux2.EOF Then
                                    rsaux10.Open "insert INTO TB_ORACLE_COMPARACION_PEDIDO_AFECTACION (PEDIDO, CANTIDAD_LEIDA, CANTIDAD_AFECTADA, fecha, CONSECUTIVO) VALUES (" + CStr(rsaux8!source_header_number) + "," + CStr(rsaux8!cantidad) + ",0, '" + CStr(rsaux2!FECHA_INICIO) + "'," + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                                 Else
                                    rsaux10.Open "insert INTO TB_ORACLE_COMPARACION_PEDIDO_AFECTACION (PEDIDO, CANTIDAD_LEIDA, CANTIDAD_AFECTADA, fecha, CONSECUTIVO) VALUES (" + CStr(rsaux8!source_header_number) + "," + CStr(rsaux8!cantidad) + ",0, ''," + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                                 End If
                                 rsaux2.Close
                                 rsaux8.MoveNext
                           Wend
                           rsaux8.Close
                           rsaux8.Open "SELECT pedido FROM TB_ORACLE_COMPARACION_PEDIDO_AFECTACION WHERE CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                           While Not rsaux8.EOF
                                 If rsaux1.State = 1 Then
                                    rsaux1.Close
                                 End If
                                 rsaux1.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 rsaux10.Open "SELECT SOURCE_HEADER_NUMBER, SUM(SHIPPED_QUANTITY) AS CANTIDAD FROM WSH_DELIVERABLES_V WHERE SOURCE_HEADER_NUMBER = " + CStr(IIf(IsNull(rsaux8!pedido), 0, rsaux8!pedido)) + " GROUP BY SOURCE_HEADER_NUMBER", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 If Not rsaux10.EOF Then
                                    rsaux1.Open "UPDATE TB_ORACLE_COMPARACION_PEDIDO_AFECTACION SET CANTIDAD_AFECTADA = " + CStr(IIf(IsNull(rsaux10!cantidad), 0, rsaux10!cantidad)) + " WHERE PEDIDO = " + CStr(rsaux8!pedido) + " AND CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                                 End If
                                 rsaux10.Close
                                 rsaux8.MoveNext
                           Wend
                           rsaux8.Close
                           rsaux8.Open "SELECT *  FROM TB_ORACLE_COMPARACION_PEDIDO_AFECTACION where cantidad_afectada > 0 and CANTIDAD_LEIDA <> cantidad_afectada AND CONSECUTIVO = " + CStr(var_consecutivo) + " order by PEDIDO desc "
                           If Not rsaux8.EOF Then
                              var_cadena_pedidos_mal = ""
                              While Not rsaux8.EOF
                                    If var_cadena_pedidos_mal = "" Then
                                       var_cadena_pedidos_mal = CStr(IIf(IsNull(rsaux8!pedido), 0, rsaux8!pedido))
                                    Else
                                       var_cadena_pedidos_mal = var_cadena_pedidos_mal + ", " + CStr(IIf(IsNull(rsaux8!pedido), 0, rsaux8!pedido))
                                    End If
                                    rsaux8.MoveNext
                              Wend
                              MsgBox "Los siguientes pedidos tienen errores entra la cantidad leida y la cantidad afectada: " + CStr(var_cadena_pedidos_mal), vbOKOnly, "ATENCION"
                           Else
                              If UCase(parametros(1)) = "SIDEBS12BKP" Then
                                 clnt.MSSoapInit "http://intranet/WsEBS12TEST/wsInterfaceOM.asmx?wsdl"
                              Else
                                 clnt.MSSoapInit "http://intranet/WsEBS12Prod/wsInterfaceOM.asmx?wsdl"
                              End If
                              rsaux.MoveFirst
                              While Not rsaux.EOF
                                    rsaux2.Open "select distinct delivery_id from wsh_deliverables_v where SOURCE_HEADER_NUMBER = " + CStr(rsaux!source_header_number) + " AND delivery_id IS NOT NULL", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                    While Not rsaux2.EOF
                                          VAR_ENTREGA = rsaux2!delivery_id
                                          rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                          VAR_ESTATUS = 0
                                          On Error GoTo salirc:
                                          
'-----
                                          strconsulta = "select * from wsh_deliverables_v where delivery_id = ? and released_status = 'Y' and SHIPPED_QUANTITY is null"
                                          With comandoORA
                                               .ActiveConnection = cnnoracle_4
                                               .CommandType = adCmdText
                                               .CommandText = strconsulta
                                               Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, VAR_ENTREGA)
                                               .Parameters.Append parametro
                                          End With
                                          Set rsaux6 = comandoORA.execute
                                          Set comandoORA = Nothing
                                          Set parametro = Nothing
                                                 
                                          var_posible_entrega = 1
                                          If Not rsaux6.EOF Then
                                             var_posible_entrega = 0
                                          End If
                                          rsaux6.Close
                                          If var_posible_entrega = 1 Then
                                             strconsulta = "select sum(SHIPPED_QUANTITY) as cantidad from wsh_deliverables_v where source_header_number  = ?"
                                             With comandoORA
                                                  .ActiveConnection = cnnoracle_4
                                                  .CommandType = adCmdText
                                                  .CommandText = strconsulta
                                                  Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux!source_header_number)
                                                  .Parameters.Append parametro
                                             End With
                                             Set rsaux6 = comandoORA.execute
                                             Set comandoORA = Nothing
                                             Set parametro = Nothing
                                             var_cantidad_oracle = IIf(IsNull(rsaux6!cantidad), 0, rsaux6!cantidad)
                                             rsaux6.Close
                                                       
                                             strconsulta = "select sum(floa_sal_cantidad_leida) as cantidad from xxvia_tb_Salidas_cajas where source_header_number  = ?"
                                             With comandoORA
                                                  .ActiveConnection = cnnoracle_4
                                                  .CommandType = adCmdText
                                                  .CommandText = strconsulta
                                                  Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux!source_header_number)
                                                  .Parameters.Append parametro
                                             End With
                                             Set rsaux6 = comandoORA.execute
                                             Set comandoORA = Nothing
                                             Set parametro = Nothing
                                             var_cantidad_sid = IIf(IsNull(rsaux6!cantidad), 0, rsaux6!cantidad)
                                             rsaux6.Close
                                             If var_cantidad_oracle <> var_cantidad_sid Then
                                                var_posible_entrega = 0
                                             End If
                                                                                                  
                                           End If
'fin de la validacion
                                           If var_posible_entrega = 1 Then
'-----
                                              var_arreglo = clnt.ASIGNAR_embarque(VAR_ENTREGA, Val(VAR_X_TRIP_ID), "CONFIRM")
                                              var_pedido = rsaux!source_header_number
                                              rsaux1.Open "insert into tb_oracle_pedidos_confirmados (pedido, fecha, maquina, error) values (" + CStr(rsaux!source_header_number) + ", getdate(), '" + fun_NombrePc + "'," + CStr(VAR_ESTATUS) + ")", cnn, adOpenDynamic, adLockOptimistic
                                              rsaux14.Open "UPDATE TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES SET ESTATUS = 'S' WHERE PEDIDO  = " + CStr(var_pedido) + " AND EMBARQUE = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
                                              rsaux14.Open "update tb_oracle_cajas_aduana set estatus = 'S' where PEDIDO  = " + CStr(var_pedido) + " and embarque = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
                                              rsaux14.Open "update XXVIA_TB_SALIDAS_CAJAS set estatus_pedido = 2, char_paq_estatus = 'S' WHERE SOURCE_HEADER_NUMBER = " + CStr(var_pedido), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                              rsaux14.Open "update TB_ORACLE_TIEMPO_PEDIDO_ADUANAS set HORA_FIN = GETDATE() where pedido = " + CStr(var_pedido), cnn, adOpenDynamic, adLockOptimistic
                                           
                                           Else
                                              MsgBox "El pedido no puede ser cerrado porque hay diferencias entre las cantidades del oracle y el SID", vbOKOnly, "ATENCION"
                                           
                                           End If
                                           rsaux2.MoveNext
                                    Wend
                                    rsaux2.Close
                                    rsaux.MoveNext
                              Wend
                              Set clnt = Nothing
                              'MsgBox "Se termino de cerrar el pedido", vbOKOnly, "ATENCION"
                           End If
                           If rsaux8.State = 1 Then
                              rsaux8.Close
                           End If
                        End If
                        rsaux.Close
                     Else
                        If VAR_ESTATUS = "F" Then
                           MsgBox "EL embarque ya fue facturado"
                        Else
                           MsgBox "El embarque NO a sido cerrado", vbOKOnly, "ATENCION"
                        End If
                     End If
                  End If
                  rs.Close
               End If
               '--------------- fin de confirmar pedidos
               'MsgBox "Se a cerrado el embarque", vbOKOnly, "ATENCION"
               'Me.frm_sellos.Visible = False
               'Me.txt_codigo.Enabled = False
            Else
               MsgBox "No se pudo crear el embarque en oracle", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "No se cerro el embarque", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El embarque ya habia sido cerrado", vbOKOnly, "ATENCION"
      End If
   Else
      rsaux.Open "DELETE FROM TB_TEMP_ORACLE_COMPARACION_CANTIDADES_LEIDAS_VS_PEDIDAS_AFECTADAS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " and codigo is null", cnn, adOpenDynamic, adLockOptimistic
      MsgBox "El embarque tiene diferencias entre las piezas pedidas y las leidas", vbOKOnly, "ATENCION"
      Set reporte = appl.OpenReport(App.Path + "\rep_oracle_diferencias_pedido_leido.rpt")
      var_cadena = "{TB_TEMP_ORACLE_COMPARACION_CANTIDADES_LEIDAS_VS_PEDIDAS_AFECTADAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
      reporte.RecordSelectionFormula = var_cadena
      frmvistasprevias.cr.ReportSource = reporte
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      frmvistasprevias.cr.ViewReport
      frmvistasprevias.Caption = "Reporte de diferencias pedido contra leido"
      frmvistasprevias.Show 1
      Set reporte = Nothing
      
      rsaux.Open "DELETE FROM TB_TEMP_ORACLE_COMPARACION_CANTIDADES_LEIDAS_VS_PEDIDAS_AFECTADAS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
   End If
   If rsaux6.State = 1 Then
      rsaux6.Close
   End If
   Else
      var_pedido = CDbl(Me.lv_cajas.selectedItem.SubItems(1))
                                              rsaux1.Open "insert into tb_oracle_pedidos_confirmados (pedido, fecha, maquina, error) values (" + CStr(Me.lv_cajas.selectedItem.SubItems(1)) + ", getdate(), '" + fun_NombrePc + "',1)", cnn, adOpenDynamic, adLockOptimistic
                                              rsaux14.Open "UPDATE TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES SET ESTATUS = 'S' WHERE PEDIDO  = " + CStr(var_pedido) + " AND EMBARQUE = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
                                              rsaux14.Open "update tb_oracle_cajas_aduana set estatus = 'S' where PEDIDO  = " + CStr(var_pedido), cnn, adOpenDynamic, adLockOptimistic
                                              rsaux14.Open "update XXVIA_TB_SALIDAS_CAJAS set estatus_pedido = 2, char_paq_estatus = 'S' WHERE SOURCE_HEADER_NUMBER = " + CStr(var_pedido), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                              rsaux14.Open "update TB_ORACLE_TIEMPO_PEDIDO_ADUANAS set HORA_FIN = GETDATE() where pedido = " + CStr(var_pedido), cnn, adOpenDynamic, adLockOptimistic
      
   End If
   Exit Sub
salir2:
   'MsgBox Err.Description
   If Err.Number = -2147217900 Then
      'MsgBox Err.Description
      rsaux14.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      rsaux14.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      Resume
   End If
salirc:
   If Err.Number = -2147467259 Then
      'MsgBox Err.Description
      Resume Next
      VAR_ESTATUS = 1
   End If
   'MsgBox Err.Number
   If Err.Number = 5415 Then
      Resume Next
   End If
   MsgBox Err.Description


End Sub

Private Sub Command4_Click()
   If rs.State = 1 Then
      rs.Close
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
   If rsaux10.State = 1 Then
      rsaux10.Close
   End If
   If rsaux11.State = 1 Then
      rsaux11.Close
   End If
   If rsaux12.State = 1 Then
      rsaux12.Close
   End If
   If rsaux13.State = 1 Then
      rsaux13.Close
   End If
   If rsaux14.State = 1 Then
      rsaux14.Close
   End If
   If rsaux15.State = 1 Then
      rsaux15.Close
   End If
   
   
   If IsNumeric(Me.txt_embarque) Then
      'rs.Open "select * from TB_ORACLE_GRUPOS_EMBARQUES where grupo = '" + Me.txt_embarque + "'", cnn, adOpenDynamic, adLockOptimistic
      rs.Open "select * from TB_ORACLE_GRUPOS_EMBARQUES where grupo = (select top 1 grupo from TB_ORACLE_GRUPOS_EMBARQUES where embarque = '" + Me.txt_embarque + "' )", cnn, adOpenDynamic, adLockOptimistic
      var_Cadena_embarques = ""
      If Not rs.EOF Then
         var_transporte = IIf(IsNull(rs!unidad), "", rs!unidad)
         While Not rs.EOF
               If var_Cadena_embarques = "" Then
                  var_Cadena_embarques = rs!Embarque
               Else
                  var_Cadena_embarques = var_Cadena_embarques + "," + rs!Embarque
               End If
               rs.MoveNext
         Wend
         rs.Close
         'Me.txt_embarques = var_Cadena_embarques
          
         strconsulta = "SELECT * FROM XXVIA_tB_ENCABEZADO_EMBARQUES WHERE embarque in (" + var_Cadena_embarques + ") "
         rs.Open strconsulta, cnnoracle_4, adOpenDynamic, adLockOptimistic
         'With comandoORA
         '     .ActiveConnection = cnnoracle_4
         '     .CommandType = adCmdText
         '     .CommandText = strconsulta
         '     Set parametro = .CreateParameter(, adVarChar, adParamInput, 1000, var_Cadena_embarques)
         '     .Parameters.Append parametro
         'End With
         'MsgBox var_Cadena_embarques
         'Set rs = comandoORA.execute
         'Set comandoORA = Nothing
         'Set parametro = Nothing
         
         rsaux1.Open "SELECT SUM(TB_ORACLE_EMPAQUES.VOLUMEN) FROM TB_ORACLE_CAJAS_ADUANA A, TB_ORACLE_EMPAQUES WHERE EMBARQUE in (" + var_Cadena_embarques + ") AND TIPO_EMPAQUE = TB_ORACLE_EMPAQUES.EMPAQUE AND ESTATUS in ('S','L')", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux1.EOF Then
            Me.txt_volumen_carga = Round(IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value), 2)
         Else
            Me.txt_volumen_carga = 0
         End If
         rsaux1.Close
         
        
         
         rsaux1.Open "SELECT SUM(TB_ORACLE_EMPAQUES.VOLUMEN) FROM TB_ORACLE_CAJAS_ADUANA A, TB_ORACLE_EMPAQUES WHERE EMBARQUE in (" + var_Cadena_embarques + ") AND TIPO_EMPAQUE = TB_ORACLE_EMPAQUES.EMPAQUE", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux1.EOF Then
            Me.txt_volumen_carga_lectores = Round(IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value), 2)
         Else
            Me.txt_volumen_carga_lectores = 0
         End If
         rsaux1.Close
         
         
         
         
         If Not rs.EOF Then
            rsaux.Open "select * from tb_oracle_transportes where clave = '" + var_transporte + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               Me.txt_clave_unidad = IIf(IsNull(rsaux!clave), "", rsaux!clave)
               Me.txt_nombre_unidad = IIf(IsNull(rsaux!nombre), "", rsaux!nombre)
               Me.txt_volumen_unidad = Round(IIf(IsNull(rsaux!VOLUMEN), 0, rsaux!VOLUMEN), 2)
               
               If CDbl(Me.txt_volumen_unidad) > 0 Then
                  var_porcentaje = (CDbl(Me.txt_volumen_carga) * 100) / CDbl(Me.txt_volumen_unidad)
                  Me.txt_porcentaje_carga = Round(var_porcentaje, 2)
               Else
                  var_porcentaje = 0
                  Me.txt_porcentaje_carga = Round(var_porcentaje, 2)
               End If
               
               Me.lbl_porcentaje_aduana = txt_porcentaje_carga + " %"
               
               
               If var_porcentaje = 0 Then
                  Me.ProgressBar1.Value = 0
               End If
               
               If var_porcentaje > 0 And var_porcentaje < 5.01 Then
                  Me.ProgressBar1.Value = 5
               End If
               
               If var_porcentaje > 5 And var_porcentaje < 10.01 Then
                  Me.ProgressBar1.Value = 10
               End If
               
               If var_porcentaje > 10 And var_porcentaje < 15.01 Then
                  Me.ProgressBar1.Value = 15
               End If
               
               If var_porcentaje > 15 And var_porcentaje < 20.01 Then
                  Me.ProgressBar1.Value = 20
               End If
               
               If var_porcentaje > 20 And var_porcentaje < 25.01 Then
                  Me.ProgressBar1.Value = 25
               End If
               
               
               If var_porcentaje > 25 And var_porcentaje < 30.01 Then
                  Me.ProgressBar1.Value = 30
               End If
               
               
               If var_porcentaje > 30 And var_porcentaje < 35.01 Then
                  Me.ProgressBar1.Value = 35
               End If
               
               
               If var_porcentaje > 35 And var_porcentaje < 40.01 Then
                  Me.ProgressBar1.Value = 40
               End If
               
               If var_porcentaje > 40 And var_porcentaje < 45.01 Then
                  Me.ProgressBar1.Value = 45
               End If
               
               If var_porcentaje > 45 And var_porcentaje < 50.01 Then
                  Me.ProgressBar1.Value = 50
               End If
               
               If var_porcentaje > 50 And var_porcentaje < 55.01 Then
                  Me.ProgressBar1.Value = 55
               End If
               
               If var_porcentaje > 55 And var_porcentaje < 60.01 Then
                  Me.ProgressBar1.Value = 60
               End If
               
               
               If var_porcentaje > 60 And var_porcentaje < 65.01 Then
                  Me.ProgressBar1.Value = 65
               End If
               
               If var_porcentaje > 65 And var_porcentaje < 70.01 Then
                  Me.ProgressBar1.Value = 70
               End If
               
               If var_porcentaje > 70 And var_porcentaje < 75.01 Then
                  Me.ProgressBar1.Value = 75
               End If
               
               
               If var_porcentaje > 75 And var_porcentaje < 80.01 Then
                  Me.ProgressBar1.Value = 80
               End If
               
               If var_porcentaje > 80 And var_porcentaje < 85.01 Then
                  Me.ProgressBar1.Value = 85
               End If
               
               If var_porcentaje > 85 And var_porcentaje < 90.01 Then
                  Me.ProgressBar1.Value = 90
               End If
               
               
               If var_porcentaje > 90 And var_porcentaje < 95.01 Then
                  Me.ProgressBar1.Value = 95
               End If
               
               If var_porcentaje > 95 Then
                  Me.ProgressBar1.Value = 100
               End If
'------------------
               If CDbl(Me.txt_volumen_unidad) > 0 Then
                  var_porcentaje = (CDbl(Me.txt_volumen_carga_lectores) * 100) / CDbl(Me.txt_volumen_unidad)
                  Me.txt_porcentaje_carga_lectores = Round(var_porcentaje, 2)
               Else
                  var_porcentaje = 0
                  Me.txt_porcentaje_carga_lectores = Round(var_porcentaje, 2)
               End If
               Me.lbl_porcentaje_lectores = txt_porcentaje_carga_lectores + " %"
               
               
               If var_porcentaje = 0 Then
                  Me.ProgressBar2.Value = 0
               End If
               
               If var_porcentaje > 0 And var_porcentaje < 5.01 Then
                  Me.ProgressBar2.Value = 5
               End If
               
               If var_porcentaje > 5 And var_porcentaje < 10.01 Then
                  Me.ProgressBar2.Value = 10
               End If
               
               If var_porcentaje > 10 And var_porcentaje < 15.01 Then
                  Me.ProgressBar2.Value = 15
               End If
               
               If var_porcentaje > 15 And var_porcentaje < 20.01 Then
                  Me.ProgressBar2.Value = 20
               End If
               
               If var_porcentaje > 20 And var_porcentaje < 25.01 Then
                  Me.ProgressBar2.Value = 25
               End If
               
               
               If var_porcentaje > 25 And var_porcentaje < 30.01 Then
                  Me.ProgressBar2.Value = 30
               End If
               
               
               If var_porcentaje > 30 And var_porcentaje < 35.01 Then
                  Me.ProgressBar2.Value = 35
               End If
               
               
               If var_porcentaje > 35 And var_porcentaje < 40.01 Then
                  Me.ProgressBar2.Value = 40
               End If
               
               If var_porcentaje > 40 And var_porcentaje < 45.01 Then
                  Me.ProgressBar2.Value = 45
               End If
               
               If var_porcentaje > 45 And var_porcentaje < 50.01 Then
                  Me.ProgressBar2.Value = 50
               End If
               
               If var_porcentaje > 50 And var_porcentaje < 55.01 Then
                  Me.ProgressBar2.Value = 55
               End If
               
               If var_porcentaje > 55 And var_porcentaje < 60.01 Then
                  Me.ProgressBar2.Value = 60
               End If
               
               
               If var_porcentaje > 60 And var_porcentaje < 65.01 Then
                  Me.ProgressBar2.Value = 65
               End If
               
               If var_porcentaje > 65 And var_porcentaje < 70.01 Then
                  Me.ProgressBar2.Value = 70
               End If
               
               If var_porcentaje > 70 And var_porcentaje < 75.01 Then
                  Me.ProgressBar2.Value = 75
               End If
               
               
               If var_porcentaje > 75 And var_porcentaje < 80.01 Then
                  Me.ProgressBar2.Value = 80
               End If
               
               If var_porcentaje > 80 And var_porcentaje < 85.01 Then
                  Me.ProgressBar2.Value = 85
               End If
               
               If var_porcentaje > 85 And var_porcentaje < 90.01 Then
                  Me.ProgressBar2.Value = 90
               End If
               
               
               If var_porcentaje > 90 And var_porcentaje < 95.01 Then
                  Me.ProgressBar2.Value = 95
               End If
               
               If var_porcentaje > 95 Then
                  Me.ProgressBar2.Value = 100
               End If
               
            Else
               MsgBox "El grupo de embarques no tienen una unidad seleccionada", vbOKOnly, "ATENCION"
               Me.txt_clave_unidad = ""
               Me.txt_nombre_unidad = ""
               Me.txt_volumen_unidad = 0
               Me.txt_porcentaje_carga = 0
               Me.txt_porcentaje_carga_lectores = 0
               
               'Me.txt_clave_unidad.SetFocus
            End If
            rsaux.Close
         Else
            'MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         rs.Close
         'MsgBox "El grupo de embarques no existe"
      End If
   Else
      'MsgBox "Número de embarque incorrecto", vbOKOnly, "ATENCION"
   End If
   
   
   
   
   
   
   
   
   Me.Command3.Caption = "DETENER"
   rs.Open "SELECT SUM(PIEZAS) FROM tb_oracle_cajas_aduana WHERE ESTATUS = 'L' AND EMBARQUE = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
   Me.txt_cantidad = Format(IIf(IsNull(rs(0).Value), 0, rs(0).Value), "###,###,##0.00")
   rs.Close
   var_contador = 1
   rs.Open "select AGENTE, nombre_agente, pedido, cliente, orden_pedido, ESTATUS, estatus_pedido, ISNULL(PAQUETERIA,0) PAQUETERIA from tb_oracle_pedidos_asignados_embarques WHERE EMBARQUE = " + Me.txt_embarque + " AND ((ISNULL(ESTATUS,'') = '')) order by orden_pedido, pedido", cnn, adOpenDynamic, adLockOptimistic
   Me.lv_cajas.ListItems.Clear
   Me.lv_cajas_siguientes.ListItems.Clear
   While Not rs.EOF
         
         If var_contador = 1 Then
            var_pedido = rs!pedido
            rsaux10.Open "select * from tb_oracle_cajas_aduana where embarque = " + Me.txt_embarque + " and pedido = " + CStr(var_pedido) + " and (ISNULL(estatus,'') = '' OR ESTATUS = 'L')", cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux10.EOF
                  rsaux12.Open "select AGENTE, nombre_agente, pedido, cliente, orden_pedido, ESTATUS, estatus_pedido, isnull(paqueteria,0) paqueteria from tb_oracle_pedidos_asignados_embarques WHERE EMBARQUE = " + Me.txt_embarque + " AND ((ISNULL(ESTATUS,'') = '')) and pedido = " + CStr(rs!pedido) + " order by orden_pedido, pedido", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux12.EOF Then
                     If IIf(IsNull(rsaux12!paqueteria), 0, rsaux12!paqueteria) = 1 Then
                        Me.lbl_paqueteria = "PAQUETERIA"
                     Else
                        Me.lbl_paqueteria = ""
                     End If
                  End If
                  rsaux12.Close
                 
                 
                 If IIf(IsNull(rs!estatus_pedido), 0, rs!estatus_pedido) = 1 Then
                    Me.lbl_semaforo.BackColor = &HC000&
                 Else
                    Me.lbl_semaforo.BackColor = &HC0&
                  End If
                  Set list_item = Me.lv_cajas.ListItems.Add(, , rsaux10!Caja)
                  list_item.SubItems(1) = IIf(IsNull(rsaux10!pedido), "", rsaux10!pedido)
                  list_item.SubItems(2) = IIf(IsNull(rsaux10!Agente), "", rsaux10!Agente)
                  list_item.SubItems(3) = IIf(IsNull(rsaux10!Cliente), "", rsaux10!Cliente)
                  list_item.SubItems(4) = Format(rsaux10!PIEZAS, "###,###,##0.00")
                  list_item.SubItems(5) = IIf(IsNull(rsaux10!estatus), "", rsaux10!estatus)
                  list_item.SubItems(6) = IIf(IsNull(rsaux10!numero_caja), "", rsaux10!numero_caja)
                  list_item.SubItems(7) = IIf(IsNull(rsaux10!TIPO_EMPAQUE), "", rsaux10!TIPO_EMPAQUE)
                  list_item.SubItems(8) = IIf(IsNull(rsaux10!caja_pedido), "", rsaux10!caja_pedido)
                  list_item.SubItems(9) = IIf(IsNull(rsaux10!sello), "", rsaux10!sello)
                  list_item.SubItems(10) = ""
                  list_item.SubItems(11) = IIf(IsNull(rsaux10!caja_actual), "", rsaux10!caja_actual)
                  list_item.SubItems(12) = IIf(IsNull(rsaux10!MARCA_CAMBIO_EMBARQUE), "", rsaux10!MARCA_CAMBIO_EMBARQUE)
                   
                  rsaux10.MoveNext
            Wend
            rsaux10.Close
         Else
            rsaux10.Open "select * from tb_oracle_cajas_aduana where embarque = " + Me.txt_embarque + " and pedido = " + CStr(rs!pedido) + " and pedido <>" + CStr(var_pedido) + " and estatus = ''", cnn, adOpenDynamic, adLockOptimistic
            'rsaux10.Open "select char_paq_estatus, inte_paq_caja, sum(floa_sal_cantidad_leida) as cantidad from xxvia_tb_Salidas_cajas where source_header_number = " + CStr(rs!pedido) + " and char_paq_estatus = '' and source_header_number <> " + CStr(var_pedido) + "  and inte_emb_embarque = " + Me.txt_embarque + " group by char_paq_estatus, inte_paq_caja", cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rsaux10.EOF
                  Set list_item = Me.lv_cajas_siguientes.ListItems.Add(, , rsaux10!Caja)
                  list_item.SubItems(1) = IIf(IsNull(rsaux10!pedido), "", rsaux10!pedido)
                  list_item.SubItems(2) = IIf(IsNull(rsaux10!Agente), "", rsaux10!Agente)
                  list_item.SubItems(3) = IIf(IsNull(rsaux10!Cliente), "", rsaux10!Cliente)
                  list_item.SubItems(4) = Format(rsaux10!PIEZAS, "###,###,##0.00")
                  list_item.SubItems(5) = IIf(IsNull(rsaux10!estatus), "", rsaux10!estatus)
                  list_item.SubItems(6) = IIf(IsNull(rsaux10!numero_caja), "", rsaux10!numero_caja)
                  list_item.SubItems(7) = IIf(IsNull(rsaux10!TIPO_EMPAQUE), "", rsaux10!TIPO_EMPAQUE)
                  list_item.SubItems(8) = IIf(IsNull(rsaux10!caja_pedido), "", rsaux10!caja_pedido)
                  list_item.SubItems(9) = IIf(IsNull(rsaux10!sello), "", rsaux10!sello)
                  rsaux10.MoveNext
            Wend
            rsaux10.Close
         End If
         var_contador = var_contador + 1
         rs.MoveNext
   Wend
   rs.Close
   Call ilumina_grid
End Sub

Private Sub Form_Load()
   Me.Timer1.Enabled = True
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   Me.frm_sellos.Visible = False

End Sub

Private Sub lv_cajas_GotFocus()
   Me.Timer1.Enabled = False
End Sub

Private Sub lv_cajas_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = 119 Then
      If Me.lv_cajas.selectedItem.SubItems(5) = "L" Then
         var_referencia_embarque = Mid(Me.lv_cajas.selectedItem, 2, 6)
         var_referencia_caja = Mid(Me.lv_cajas.selectedItem, 8, 3)
         If IsNumeric(var_referencia_embarque) Then
            If IsNumeric(var_referencia_caja) Then
               If CDbl(Me.txt_embarque) = CDbl(var_referencia_embarque) Then
                  If rs.State = 1 Then
                     rs.Close
                  End If
                  var_embarque_auditar = CDbl(Me.txt_embarque)
                  var_caja_auditar = CDbl(var_referencia_caja)
                  frmoracle_sello_caja.Show 1
                  If rsaux1.State = 1 Then
                     rsaux1.Close
                  End If
                  rsaux1.Open "update xxvia_tb_salidas_cajas set sello = '" + var_sello_caja + "' where inte_emb_embarque = " + CStr(Me.txt_embarque) + " and inte_paq_caja = " + CStr(var_caja_auditar), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  rsaux1.Open "update tb_oracle_cajas_aduana set sello = '" + var_sello_caja + "' where embarque = " + CStr(Me.txt_embarque) + " and numero_caja = " + CStr(var_caja_auditar), cnn, adOpenDynamic, adLockOptimistic

                  MsgBox "Se a cambiado el sello correctamente", vbOKOnly, "ATENCION"
               Else
                  MsgBox "No es posible cambiar el sello de la caja", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "No es posible cambiar el sello de la caja", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "No es posible cambiar el sello de la caja", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "No es posible cambiar el sello ya que la caja no a sido auditada", vbOKOnly, "ATENCION"
      End If
   End If
   If Shift = 1 And KeyCode = 116 Then
      If Me.lv_cajas.selectedItem.SubItems(5) = "L" Then
         MsgBox "La caja ya no puede ser auditada", vbOKOnly, "ATENCION"
      Else
         Me.Timer1.Enabled = False
         var_si = MsgBox("¿Desea auditar el bulto?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_si = MsgBox("Confirmar auditar el bulto", vbYesNo, "ATENCION")
            If var_si = 6 Then
               var_referencia_embarque = Mid(Me.lv_cajas.selectedItem, 2, 6)
               var_referencia_caja = Mid(Me.lv_cajas.selectedItem, 8, 3)
               If IsNumeric(var_referencia_embarque) Then
                  If IsNumeric(var_referencia_caja) Then
                     If CDbl(Me.txt_embarque) = CDbl(var_referencia_embarque) Then
                        If rs.State = 1 Then
                           rs.Close
                        End If
                        rs.Open "select * from tb_oracle_cajas_aduana where embarque  = " + CStr(var_referencia_embarque) + " and CAJA = '" + CStr(Me.lv_cajas.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           var_estatus_caja = IIf(IsNull(rs!estatus), "", rs!estatus)
                           If var_estatus_caja = "" Then
                              var_embarque_auditar = CDbl(Me.txt_embarque)
                              var_caja_auditar = CDbl(var_referencia_caja)
               
                              frmoracle_audita_caja.Show 1
                              rsaux.Open "SELECT * FROM XXVIA_TB_cAJAS_AUDITADAS WHERE EMBARQUE = " + CStr(var_embarque_auditar) + " AND CAJA = " + CStr(var_caja_auditar) + " AND CANTIDAD_ORIGINAL <> CANTIDAD_AUDITADA", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If rsaux.EOF Then
                                 If rsaux1.State = 1 Then
                                    rsaux1.Close
                                 End If
                                 var_observaciones_auditoria = "NO HUBO DIFERENCIAS"
                                 'frmoracle_audita_observaciones.Show 1
                                 frmoracle_sello_caja.Show 1
                                        
                                 rsaux1.Open "update xxvia_tb_salidas_cajas set sello = '" + var_sello_caja + "', AUDITADA = 1, OBSERVACIONES_AUDITORIA = '" + var_observaciones_auditoria + "' where inte_emb_embarque = " + CStr(Me.txt_embarque) + " and inte_paq_caja = " + CStr(var_caja_auditar), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 rsaux1.Open "update tb_oracle_cajas_aduana set estatus = 'L', sello = '" + var_sello_caja + "' where embarque = " + CStr(Me.txt_embarque) + " and numero_caja = " + CStr(var_caja_auditar), cnn, adOpenDynamic, adLockOptimistic
                                 rsaux1.Open "select source_header_number, inte_paq_caja, char_paq_estatus, sum(floa_Sal_Cantidad_leida) as cantidad from xxvia_tb_salidas_cajas where inte_emb_embarque = " + CStr(Me.txt_embarque) + " and inte_paq_caja = " + CStr(var_caja_auditar) + " group by source_header_number, inte_paq_caja, char_paq_estatus", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 var_codigo_caja = "C" + var_referencia_embarque + var_referencia_caja
                                 var_estatus_caja = IIf(IsNull(rsaux1!char_paq_estatus), "", rsaux1!char_paq_estatus)
                                 var_cantidad_enviada = var_cantidad_enviada + rsaux1!cantidad
                                 var_cantidad_leida = 0
                                 Me.txt_codigo = ""
                                 Me.lv_cajas.selectedItem.SubItems(5) = "L"
                                 'rsaux2.Open "update xxvia_tb_salidas_cajas SET CHAR_PAQ_ESTATUS = 'S' where inte_emb_embarque = " + Me.txt_embarque + " and source_header_number = " + Me.lv_cajas.selectedItem.SubItems(1) + " and inte_paq_caja = " + Me.lv_cajas.selectedItem.SubItems(6), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 Call ilumina_grid
                              Else
                                 frmmensaje.lbl_mensaje = "Existen diferencias en la caja auditada"
                                 frmmensaje.Show 1
                                 var_observaciones_auditoria = ""
                                 frmoracle_audita_observaciones.Show 1
                                 frmoracle_sello_caja.Show 1
                                 rsaux8.Open "UPDATE XXVIA_TB_SALIDAS_CAJAS SET AUDITADA = 1, SELLO = '" + var_sello_caja + "', OBSERVACIONES_AUDITORIA = '" + var_observaciones_auditoria + "'  where inte_emb_embarque  = " + CStr(var_referencia_embarque) + " and inte_paq_caja = " + CStr(var_referencia_caja), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 rsaux1.Open "update tb_oracle_cajas_aduana set estatus = 'L' where embarque = " + CStr(Me.txt_embarque) + " and numero_caja = " + CStr(var_caja_auditar), cnn, adOpenDynamic, adLockOptimistic
                                 txt_codigo.SetFocus
                                 var_orden_surtido = 0
                                 var_caja = 0
                                 var_factura_ceros = 0
                                 var_tipo_pedido = ""
                                 Me.txt_codigo = ""
                                 Me.txt_codigo.SetFocus
                              End If
                              rsaux.Close
                           Else
                              MsgBox "La caja ya no puede ser auditada", vbOKOnly, "ATENCION"
                           End If
                        End If
                        If rs.State = 1 Then
                           rs.Close
                        End If
                     Else
                     End If
                  Else
                  End If
               End If
            End If
         End If
      End If
   End If
   If Shift = 1 And KeyCode = 117 Then
      var_renglon = Me.lv_cajas.selectedItem.Index
      
      var_si_permiso = 0
      frmoracle_permiso_cerrar_pedidos.Show 1
      If var_si_permiso = 1 Then
         Me.Timer1.Enabled = False
         frmoracle_tipo_cajas.Show 1
         txt_nombre_caja = var_nombre_caja
         var_si = MsgBox("¿Desea cambiar el tipo de bulto?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            Me.Timer1.Enabled = False
            Me.lv_cajas.ListItems.Item(var_renglon).Selected = True
            var_referencia_embarque = Mid(Me.lv_cajas.selectedItem, 2, 6)
            var_referencia_caja = Mid(Me.lv_cajas.selectedItem, 8, 3)
            rsaux8.Open "UPDATE XXVIA_TB_SALIDAS_CAJAS SET TIPO_CAJA = '" + var_nombre_caja + "', OBSERVACIONES_AUDITORIA = 'CAMBIO DE TIPO DE BULTO'  where inte_emb_embarque  = " + CStr(Me.txt_embarque) + " and inte_paq_caja = " + CStr(var_referencia_caja), cnnoracle_4, adOpenDynamic, adLockOptimistic
            rsaux10.Open "UPDATE tb_oracle_cajas_aduana SET TIPO_EMPAQUE = '" + var_nombre_caja + "' WHERE EMBARQUE = " + CStr(var_referencia_embarque) + " AND NUMERO_CAJA = " + CStr(var_referencia_caja), cnn, adOpenDynamic, adLockOptimistic
            Me.lv_cajas.selectedItem.SubItems(7) = var_nombre_caja
            Me.txt_codigo = ""
            Me.txt_codigo.SetFocus
         End If
         Me.Timer1.Enabled = True
      Else
         MsgBox "No tiene autorizacion para cambiar tipos de bultos", vbOKOnly, "ATENCION"
         Me.txt_codigo.SetFocus
      End If
   End If
   
   If Shift = 1 And KeyCode = 113 Then
      var_referencia_embarque = Mid(Me.lv_cajas.selectedItem, 2, 6)
      var_referencia_caja = Mid(Me.lv_cajas.selectedItem, 8, 3)
      If IsNumeric(var_referencia_embarque) Then
         If IsNumeric(var_referencia_caja) Then
            If CDbl(Me.txt_embarque) = CDbl(var_referencia_embarque) Then
               var_embarque_auditar = CDbl(Me.txt_embarque)
               var_caja_auditar = CDbl(var_referencia_caja)
               If rs.State = 1 Then
                  rs.Close
               End If
               rs.Open "select b.MAQUINA, usuario from xxvia_Tb_salidas_cajas a, xxvia_Tb_bitacora_lectura  b where a.inte_emb_embarque = " + CStr(var_embarque_auditar) + " and inte_paq_caja = " + CStr(var_caja_auditar) + " and a.source_header_number = b.pedido and inte_paq_caja = caja", cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_maquina_ad = IIf(IsNull(rs!maquina), "", rs!maquina)
                  var_usuario_ad = IIf(IsNull(rs!USUARIO), "", rs!USUARIO)
                  rsaux10.Open "select * from tb_usuarios where vcha_usu_usuario_id = '" + var_usuario_id0 + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux10.EOF Then
                     var_nombre_usuario_ad = IIf(IsNull(rsaux10!vcha_usu_nombre), "", rsaux10!vcha_usu_nombre) + " " + IIf(IsNull(rsaux10!vcha_usu_apellidos), "", rsaux10!vcha_usu_apellidos)
                  Else
                     var_nombre_usuario_ad = ""
                  End If
                  rsaux10.Close
               Else
                  var_maquina_ad = ""
                  var_usuario_ad = ""
               End If
               rs.Close
               MsgBox "Caja leida en la maquina " + var_maquina_ad + " por el usuario " + var_nombre_usuario_ad, vbOKOnly, "ATENCION"
            End If
         End If
      End If
   End If
End Sub

Private Sub lv_cajas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.lv_cajas.selectedItem.SubItems(12) = "" Then
         Me.lv_cajas.selectedItem.SubItems(12) = "*"
         var_i = Me.lv_cajas.selectedItem.Index
         lv_cajas.ListItems.Item(var_i).Bold = True
         rs.Open "UPDATE tb_oracle_cajas_aduana SET MARCA_CAMBIO_EMBARQUE = '*' WHERE EMBARQUE = '" + Me.txt_embarque + "' AND CAJA = '" + Me.lv_cajas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
         
         lv_cajas.ListItems.Item(var_i).ListSubItems(1).Bold = True
         lv_cajas.ListItems.Item(var_i).ListSubItems(2).Bold = True
         lv_cajas.ListItems.Item(var_i).ListSubItems(3).Bold = True
         lv_cajas.ListItems.Item(var_i).ListSubItems(4).Bold = True
         lv_cajas.ListItems.Item(var_i).ListSubItems(5).Bold = True
         lv_cajas.ListItems.Item(var_i).ListSubItems(6).Bold = True
         lv_cajas.ListItems.Item(var_i).ListSubItems(7).Bold = True
         lv_cajas.ListItems.Item(var_i).ListSubItems(8).Bold = True
         lv_cajas.ListItems.Item(var_i).ListSubItems(9).Bold = True
         lv_cajas.ListItems.Item(var_i).ForeColor = &HC000&
         lv_cajas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &HC000&
         lv_cajas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &HC000&
         lv_cajas.ListItems.Item(var_i).ListSubItems(3).ForeColor = &HC000&
         lv_cajas.ListItems.Item(var_i).ListSubItems(4).ForeColor = &HC000&
         lv_cajas.ListItems.Item(var_i).ListSubItems(5).ForeColor = &HC000&
         lv_cajas.ListItems.Item(var_i).ListSubItems(6).ForeColor = &HC000&
         lv_cajas.ListItems.Item(var_i).ListSubItems(7).ForeColor = &HC000&
         lv_cajas.ListItems.Item(var_i).ListSubItems(8).ForeColor = &HC000&
         lv_cajas.ListItems.Item(var_i).ListSubItems(9).ForeColor = &HC000&
      Else
         var_i = Me.lv_cajas.selectedItem.Index
         Me.lv_cajas.selectedItem.SubItems(12) = ""
         lv_cajas.ListItems.Item(var_i).Bold = False
         rs.Open "UPDATE tb_oracle_cajas_aduana SET MARCA_CAMBIO_EMBARQUE = '' WHERE EMBARQUE = '" + Me.txt_embarque + "' AND CAJA = '" + Me.lv_cajas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
         lv_cajas.ListItems.Item(var_i).ListSubItems(1).Bold = False
         lv_cajas.ListItems.Item(var_i).ListSubItems(2).Bold = False
         lv_cajas.ListItems.Item(var_i).ListSubItems(3).Bold = False
         lv_cajas.ListItems.Item(var_i).ListSubItems(4).Bold = False
         lv_cajas.ListItems.Item(var_i).ListSubItems(5).Bold = False
         lv_cajas.ListItems.Item(var_i).ListSubItems(6).Bold = False
         lv_cajas.ListItems.Item(var_i).ListSubItems(7).Bold = False
         lv_cajas.ListItems.Item(var_i).ListSubItems(8).Bold = False
         lv_cajas.ListItems.Item(var_i).ListSubItems(9).Bold = False
         lv_cajas.ListItems.Item(var_i).ForeColor = &H80000008
         lv_cajas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H80000008
         lv_cajas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H80000008
         lv_cajas.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H80000008
         lv_cajas.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H80000008
         lv_cajas.ListItems.Item(var_i).ListSubItems(5).ForeColor = &H80000008
         lv_cajas.ListItems.Item(var_i).ListSubItems(6).ForeColor = &H80000008
         lv_cajas.ListItems.Item(var_i).ListSubItems(7).ForeColor = &H80000008
         lv_cajas.ListItems.Item(var_i).ListSubItems(8).ForeColor = &H80000008
         lv_cajas.ListItems.Item(var_i).ListSubItems(9).ForeColor = &H80000008
      End If
   End If
End Sub

Private Sub lv_cajas_LostFocus()
   Me.Timer1.Enabled = True
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_ventana = 2
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_oracle_transportes", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!clave)
            list_item.SubItems(1) = IIf(IsNull(rs!nombre), "", rs!nombre)
            list_item.SubItems(2) = IIf(IsNull(rs!VOLUMEN), "", rs!VOLUMEN)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TRANSPORTES"
      VAR_TIPO_LISTA = 100
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

Private Sub Timer1_Timer()
   If rs.State = 1 Then
      rs.Close
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
   If rsaux10.State = 1 Then
      rsaux10.Close
   End If
   If rsaux11.State = 1 Then
      rsaux11.Close
   End If
   If rsaux12.State = 1 Then
      rsaux12.Close
   End If
   If rsaux13.State = 1 Then
      rsaux13.Close
   End If
   If rsaux14.State = 1 Then
      rsaux14.Close
   End If
   If rsaux15.State = 1 Then
      rsaux15.Close
   End If
   Me.Command3.Caption = "DETENER"
   
   
   
   If IsNumeric(Me.txt_embarque) Then
   
      strconsulta = "SELECT TRANSPORTE FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = ?"
      With comandoORA
           'MsgBox cnnoracle_4.ConnectionString
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_embarque))
           .Parameters.Append parametro
      End With
      Set rsaux8 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      var_clave_transporte = ""
      If Not rsaux8.EOF Then
         var_clave_transporte = IIf(IsNull(rsaux8!transporte), "", rsaux8!transporte)
      End If
      rsaux8.Close
      rsaux8.Open "SELECT isnull(NOMBRE,'') nombre FROM TB_ORACLE_TRANSPORTES WHERE CLAVE = '" + var_clave_transporte + "'", cnn, adOpenDynamic, adLockOptimistic
      Me.lbl_transporte = ""
      If Not rsaux8.EOF Then
         Me.lbl_transporte = rsaux8!nombre
      Else
         Me.lbl_transporte = ""
      End If
      rsaux8.Close

   
   
   
      'rs.Open "select * from TB_ORACLE_GRUPOS_EMBARQUES where grupo = '" + Me.txt_embarque + "'", cnn, adOpenDynamic, adLockOptimistic
      rs.Open "select * from TB_ORACLE_GRUPOS_EMBARQUES where grupo = (select top 1 grupo from TB_ORACLE_GRUPOS_EMBARQUES where embarque = '" + Me.txt_embarque + "' )", cnn, adOpenDynamic, adLockOptimistic
      var_Cadena_embarques = ""
      If Not rs.EOF Then
         var_transporte = IIf(IsNull(rs!unidad), "", rs!unidad)
         While Not rs.EOF
               If var_Cadena_embarques = "" Then
                  var_Cadena_embarques = rs!Embarque
               Else
                  var_Cadena_embarques = var_Cadena_embarques + "," + rs!Embarque
               End If
               rs.MoveNext
         Wend
         rs.Close
         'Me.txt_embarques = var_Cadena_embarques
          
         strconsulta = "SELECT * FROM XXVIA_tB_ENCABEZADO_EMBARQUES WHERE embarque in (" + var_Cadena_embarques + ") "
         rs.Open strconsulta, cnnoracle_4, adOpenDynamic, adLockOptimistic
         'With comandoORA
         '     .ActiveConnection = cnnoracle_4
         '     .CommandType = adCmdText
         '     .CommandText = strconsulta
         '     Set parametro = .CreateParameter(, adVarChar, adParamInput, 1000, var_Cadena_embarques)
         '     .Parameters.Append parametro
         'End With
         'MsgBox var_Cadena_embarques
         'Set rs = comandoORA.execute
         'Set comandoORA = Nothing
         'Set parametro = Nothing
         
         rsaux1.Open "SELECT SUM(TB_ORACLE_EMPAQUES.VOLUMEN) FROM TB_ORACLE_CAJAS_ADUANA A, TB_ORACLE_EMPAQUES WHERE EMBARQUE in (" + var_Cadena_embarques + ") AND TIPO_EMPAQUE = TB_ORACLE_EMPAQUES.EMPAQUE AND ESTATUS in ('S','L')", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux1.EOF Then
            Me.txt_volumen_carga = Round(IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value), 2)
         Else
            Me.txt_volumen_carga = 0
         End If
         rsaux1.Close
         
        
         
         rsaux1.Open "SELECT SUM(TB_ORACLE_EMPAQUES.VOLUMEN) FROM TB_ORACLE_CAJAS_ADUANA A, TB_ORACLE_EMPAQUES WHERE EMBARQUE in (" + var_Cadena_embarques + ") AND TIPO_EMPAQUE = TB_ORACLE_EMPAQUES.EMPAQUE", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux1.EOF Then
            Me.txt_volumen_carga_lectores = Round(IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value), 2)
         Else
            Me.txt_volumen_carga_lectores = 0
         End If
         rsaux1.Close
         
         
         
         
         If Not rs.EOF Then
            rsaux.Open "select * from tb_oracle_transportes where clave = '" + var_transporte + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               Me.txt_clave_unidad = IIf(IsNull(rsaux!clave), "", rsaux!clave)
               Me.txt_nombre_unidad = IIf(IsNull(rsaux!nombre), "", rsaux!nombre)
               Me.txt_volumen_unidad = Round(IIf(IsNull(rsaux!VOLUMEN), 0, rsaux!VOLUMEN), 2)
               Me.lbl_transporte = IIf(IsNull(rsaux!nombre), "", rsaux!nombre)
               
               If CDbl(Me.txt_volumen_unidad) > 0 Then
                  var_porcentaje = (CDbl(Me.txt_volumen_carga) * 100) / CDbl(Me.txt_volumen_unidad)
                  Me.txt_porcentaje_carga = Round(var_porcentaje, 2)
               Else
                  var_porcentaje = 0
                  Me.txt_porcentaje_carga = Round(var_porcentaje, 2)
               End If
               Me.lbl_porcentaje_aduana = txt_porcentaje_carga + " %"
               
               
               
               If var_porcentaje = 0 Then
                  Me.ProgressBar1.Value = 0
               End If
               
               If var_porcentaje > 0 And var_porcentaje < 5.01 Then
                  Me.ProgressBar1.Value = 5
               End If
               
               If var_porcentaje > 5 And var_porcentaje < 10.01 Then
                  Me.ProgressBar1.Value = 10
               End If
               
               If var_porcentaje > 10 And var_porcentaje < 15.01 Then
                  Me.ProgressBar1.Value = 15
               End If
               
               If var_porcentaje > 15 And var_porcentaje < 20.01 Then
                  Me.ProgressBar1.Value = 20
               End If
               
               If var_porcentaje > 20 And var_porcentaje < 25.01 Then
                  Me.ProgressBar1.Value = 25
               End If
               
               
               If var_porcentaje > 25 And var_porcentaje < 30.01 Then
                  Me.ProgressBar1.Value = 30
               End If
               
               
               If var_porcentaje > 30 And var_porcentaje < 35.01 Then
                  Me.ProgressBar1.Value = 35
               End If
               
               
               If var_porcentaje > 35 And var_porcentaje < 40.01 Then
                  Me.ProgressBar1.Value = 40
               End If
               
               If var_porcentaje > 40 And var_porcentaje < 45.01 Then
                  Me.ProgressBar1.Value = 45
               End If
               
               If var_porcentaje > 45 And var_porcentaje < 50.01 Then
                  Me.ProgressBar1.Value = 50
               End If
               
               If var_porcentaje > 50 And var_porcentaje < 55.01 Then
                  Me.ProgressBar1.Value = 55
               End If
               
               If var_porcentaje > 55 And var_porcentaje < 60.01 Then
                  Me.ProgressBar1.Value = 60
               End If
               
               
               If var_porcentaje > 60 And var_porcentaje < 65.01 Then
                  Me.ProgressBar1.Value = 65
               End If
               
               If var_porcentaje > 65 And var_porcentaje < 70.01 Then
                  Me.ProgressBar1.Value = 70
               End If
               
               If var_porcentaje > 70 And var_porcentaje < 75.01 Then
                  Me.ProgressBar1.Value = 75
               End If
               
               
               If var_porcentaje > 75 And var_porcentaje < 80.01 Then
                  Me.ProgressBar1.Value = 80
               End If
               
               If var_porcentaje > 80 And var_porcentaje < 85.01 Then
                  Me.ProgressBar1.Value = 85
               End If
               
               If var_porcentaje > 85 And var_porcentaje < 90.01 Then
                  Me.ProgressBar1.Value = 90
               End If
               
               
               If var_porcentaje > 90 And var_porcentaje < 95.01 Then
                  Me.ProgressBar1.Value = 95
               End If
               
               If var_porcentaje > 95 Then
                  Me.ProgressBar1.Value = 100
               End If
'------------------
               If CDbl(Me.txt_volumen_unidad) > 0 Then
                  var_porcentaje = (CDbl(Me.txt_volumen_carga_lectores) * 100) / CDbl(Me.txt_volumen_unidad)
                  Me.txt_porcentaje_carga_lectores = Round(var_porcentaje, 2)
               Else
                  var_porcentaje = 0
                  Me.txt_porcentaje_carga_lectores = Round(var_porcentaje, 2)
               End If
               Me.lbl_porcentaje_lectores = txt_porcentaje_carga_lectores + " %"
               
               
               If var_porcentaje = 0 Then
                  Me.ProgressBar2.Value = 0
               End If
               
               If var_porcentaje > 0 And var_porcentaje < 5.01 Then
                  Me.ProgressBar2.Value = 5
               End If
               
               If var_porcentaje > 5 And var_porcentaje < 10.01 Then
                  Me.ProgressBar2.Value = 10
               End If
               
               If var_porcentaje > 10 And var_porcentaje < 15.01 Then
                  Me.ProgressBar2.Value = 15
               End If
               
               If var_porcentaje > 15 And var_porcentaje < 20.01 Then
                  Me.ProgressBar2.Value = 20
               End If
               
               If var_porcentaje > 20 And var_porcentaje < 25.01 Then
                  Me.ProgressBar2.Value = 25
               End If
               
               
               If var_porcentaje > 25 And var_porcentaje < 30.01 Then
                  Me.ProgressBar2.Value = 30
               End If
               
               
               If var_porcentaje > 30 And var_porcentaje < 35.01 Then
                  Me.ProgressBar2.Value = 35
               End If
               
               
               If var_porcentaje > 35 And var_porcentaje < 40.01 Then
                  Me.ProgressBar2.Value = 40
               End If
               
               If var_porcentaje > 40 And var_porcentaje < 45.01 Then
                  Me.ProgressBar2.Value = 45
               End If
               
               If var_porcentaje > 45 And var_porcentaje < 50.01 Then
                  Me.ProgressBar2.Value = 50
               End If
               
               If var_porcentaje > 50 And var_porcentaje < 55.01 Then
                  Me.ProgressBar2.Value = 55
               End If
               
               If var_porcentaje > 55 And var_porcentaje < 60.01 Then
                  Me.ProgressBar2.Value = 60
               End If
               
               
               If var_porcentaje > 60 And var_porcentaje < 65.01 Then
                  Me.ProgressBar2.Value = 65
               End If
               
               If var_porcentaje > 65 And var_porcentaje < 70.01 Then
                  Me.ProgressBar2.Value = 70
               End If
               
               If var_porcentaje > 70 And var_porcentaje < 75.01 Then
                  Me.ProgressBar2.Value = 75
               End If
               
               
               If var_porcentaje > 75 And var_porcentaje < 80.01 Then
                  Me.ProgressBar2.Value = 80
               End If
               
               If var_porcentaje > 80 And var_porcentaje < 85.01 Then
                  Me.ProgressBar2.Value = 85
               End If
               
               If var_porcentaje > 85 And var_porcentaje < 90.01 Then
                  Me.ProgressBar2.Value = 90
               End If
               
               
               If var_porcentaje > 90 And var_porcentaje < 95.01 Then
                  Me.ProgressBar2.Value = 95
               End If
               
               If var_porcentaje > 95 Then
                  Me.ProgressBar2.Value = 100
               End If
               
            Else
               MsgBox "El grupo de embarques no tienen una unidad seleccionada", vbOKOnly, "ATENCION"
               Me.txt_clave_unidad = ""
               Me.txt_nombre_unidad = ""
               Me.txt_volumen_unidad = 0
               Me.txt_porcentaje_carga = 0
               Me.txt_porcentaje_carga_lectores = 0
               
               'Me.txt_clave_unidad.SetFocus
            End If
            rsaux.Close
         Else
            'MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         rs.Close
         'MsgBox "El grupo de embarques no existe"
      End If
   Else
      'MsgBox "Número de embarque incorrecto", vbOKOnly, "ATENCION"
   End If
   
   
   
   
   
   
   
   
   
   
   
   rs.Open "SELECT SUM(PIEZAS) FROM tb_oracle_cajas_aduana WHERE ESTATUS = 'L' AND EMBARQUE = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
   Me.txt_cantidad = Format(IIf(IsNull(rs(0).Value), 0, rs(0).Value), "###,###,##0.00")
   rs.Close
   var_contador = 1
   rs.Open "select AGENTE, nombre_agente, pedido, cliente, orden_pedido, ESTATUS, estatus_pedido, isnull(paqueteria,0) paqueteria from tb_oracle_pedidos_asignados_embarques WHERE EMBARQUE = " + Me.txt_embarque + " AND ((ISNULL(ESTATUS,'') = '')) order by orden_pedido, pedido", cnn, adOpenDynamic, adLockOptimistic
   Me.lv_cajas.ListItems.Clear
   Me.lv_cajas_siguientes.ListItems.Clear
   While Not rs.EOF
         If var_contador = 1 Then
            var_pedido = rs!pedido
            rsaux10.Open "select * from tb_oracle_cajas_aduana where embarque = " + Me.txt_embarque + " and pedido = " + CStr(var_pedido) + " and (ISNULL(estatus,'') = '' OR ESTATUS = 'L')", cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux10.EOF
                  rsaux12.Open "select AGENTE, nombre_agente, pedido, cliente, orden_pedido, ESTATUS, estatus_pedido, isnull(paqueteria,0) paqueteria from tb_oracle_pedidos_asignados_embarques WHERE EMBARQUE = " + Me.txt_embarque + " AND ((ISNULL(ESTATUS,'') = '')) and pedido = " + CStr(rs!pedido) + " order by orden_pedido, pedido", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux12.EOF Then
                     If IIf(IsNull(rsaux12!paqueteria), 0, rsaux12!paqueteria) = 1 Then
                        Me.lbl_paqueteria = "PAQUETERIA"
                     Else
                        Me.lbl_paqueteria = ""
                     End If
                  End If
                  rsaux12.Close
                 
                 If IIf(IsNull(rs!estatus_pedido), 0, rs!estatus_pedido) = 1 Then
                    Me.lbl_semaforo.BackColor = &HC000&
                 Else
                    Me.lbl_semaforo.BackColor = &HC0&
                  End If
                  If rsaux10!Caja = "C313593163" Then
                     x = x
                  End If
                  Set list_item = Me.lv_cajas.ListItems.Add(, , rsaux10!Caja)
                  list_item.SubItems(1) = IIf(IsNull(rsaux10!pedido), "", rsaux10!pedido)
                  list_item.SubItems(2) = IIf(IsNull(rsaux10!Agente), "", rsaux10!Agente)
                  list_item.SubItems(3) = IIf(IsNull(rsaux10!Cliente), "", rsaux10!Cliente)
                  list_item.SubItems(4) = Format(rsaux10!PIEZAS, "###,###,##0.00")
                  list_item.SubItems(5) = IIf(IsNull(rsaux10!estatus), "", rsaux10!estatus)
                  list_item.SubItems(6) = IIf(IsNull(rsaux10!numero_caja), "", rsaux10!numero_caja)
                  list_item.SubItems(7) = IIf(IsNull(rsaux10!TIPO_EMPAQUE), "", rsaux10!TIPO_EMPAQUE)
                  list_item.SubItems(8) = IIf(IsNull(rsaux10!caja_pedido), "", rsaux10!caja_pedido)
                  list_item.SubItems(9) = IIf(IsNull(rsaux10!sello), "", rsaux10!sello)
                  list_item.SubItems(11) = IIf(IsNull(rsaux10!caja_actual), "", rsaux10!caja_actual)
                  list_item.SubItems(12) = IIf(IsNull(rsaux10!MARCA_CAMBIO_EMBARQUE), "", rsaux10!MARCA_CAMBIO_EMBARQUE)
                   
                  rsaux10.MoveNext
            Wend
            rsaux10.Close
         Else
            rsaux10.Open "select * from tb_oracle_cajas_aduana where embarque = " + Me.txt_embarque + " and pedido = " + CStr(rs!pedido) + " and pedido <>" + CStr(var_pedido) + " and estatus = ''", cnn, adOpenDynamic, adLockOptimistic
            'rsaux10.Open "select char_paq_estatus, inte_paq_caja, sum(floa_sal_cantidad_leida) as cantidad from xxvia_tb_Salidas_cajas where source_header_number = " + CStr(rs!pedido) + " and char_paq_estatus = '' and source_header_number <> " + CStr(var_pedido) + "  and inte_emb_embarque = " + Me.txt_embarque + " group by char_paq_estatus, inte_paq_caja", cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rsaux10.EOF
                  Set list_item = Me.lv_cajas_siguientes.ListItems.Add(, , rsaux10!Caja)
                  list_item.SubItems(1) = IIf(IsNull(rsaux10!pedido), "", rsaux10!pedido)
                  list_item.SubItems(2) = IIf(IsNull(rsaux10!Agente), "", rsaux10!Agente)
                  list_item.SubItems(3) = IIf(IsNull(rsaux10!Cliente), "", rsaux10!Cliente)
                  list_item.SubItems(4) = Format(rsaux10!PIEZAS, "###,###,##0.00")
                  list_item.SubItems(5) = IIf(IsNull(rsaux10!estatus), "", rsaux10!estatus)
                  list_item.SubItems(6) = IIf(IsNull(rsaux10!numero_caja), "", rsaux10!numero_caja)
                  list_item.SubItems(7) = IIf(IsNull(rsaux10!TIPO_EMPAQUE), "", rsaux10!TIPO_EMPAQUE)
                  list_item.SubItems(8) = IIf(IsNull(rsaux10!caja_pedido), "", rsaux10!caja_pedido)
                  list_item.SubItems(9) = IIf(IsNull(rsaux10!sello), "", rsaux10!sello)
                  rsaux10.MoveNext
            Wend
            rsaux10.Close
         End If
         var_contador = var_contador + 1
         rs.MoveNext
   Wend
   rs.Close
   Call ilumina_grid
End Sub

Private Sub txt_codigo_Change()
   Me.Timer1.Enabled = True
   
End Sub

Private Sub txt_codigo_GotFocus()
   Me.txt_codigo = ""
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   Dim var_pedido_tiempo As Double
   Dim clnt As New SoapClient30
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      var_encontro = 0
      If Trim(Me.txt_codigo) <> "" Then
         For var_j = 1 To lv_cajas.ListItems.Count
             Me.lv_cajas.ListItems.Item(var_j).Selected = True
             If Me.txt_codigo = Me.lv_cajas.selectedItem Then
                var_pedido_tiempo = CDbl(Me.lv_cajas.selectedItem.SubItems(1))
                var_encontro = 1
                If Me.lv_cajas.selectedItem.SubItems(5) = "L" Then
                   'MsgBox "La caja ya fue leida con anterioridad", vbOKOnly, "ATENCION"
                   If Me.lv_cajas.selectedItem.SubItems(11) <> "" Then
                      frmmensaje.lbl_mensaje = "La caja esta incluida en el bulto " + Me.lv_cajas.selectedItem.SubItems(11)
                   Else
                      frmmensaje.lbl_mensaje = "La caja ya fue leida con anterioridad"
                   End If
                   frmmensaje.Show 1
                   txt_codigo.SetFocus
                   
                   Call ilumina_grid
                Else
                   '----------
                   If var_pedido_tiempo >= 10000002 Then
                      rs.Open "select * from TB_ORACLE_CAJAS_ADUANA where pedido = " + CStr(var_pedido_tiempo), cnn, adOpenDynamic, adLockOptimistic
                      var_referencia_embarque = CStr(rs!Embarque)
                      var_referencia_caja = CStr(rs!numero_caja)
                      rsaux14.Open "UPDATE TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES SET ESTATUS = 'S' WHERE PEDIDO  = " + CStr(var_pedido_tiempo), cnn, adOpenDynamic, adLockOptimistic
                      rsaux14.Open "update tb_oracle_cajas_aduana set estatus = 'S' where PEDIDO  = " + CStr(var_pedido_tiempo), cnn, adOpenDynamic, adLockOptimistic
                      rsaux14.Open "update TB_ORACLE_TIEMPO_PEDIDO_ADUANAS set HORA_FIN = GETDATE() where pedido = " + CStr(var_pedido_tiempo), cnn, adOpenDynamic, adLockOptimistic
                      
                      Me.txt_codigo = ""
                      Me.lv_cajas.selectedItem.SubItems(5) = "L"
                      'rsaux2.Open "update xxvia_tb_salidas_cajas SET CHAR_PAQ_ESTATUS = 'S' where inte_emb_embarque = " + Me.txt_embarque + " and source_header_number = " + Me.lv_cajas.selectedItem.SubItems(1) + " and inte_paq_caja = " + Me.lv_cajas.selectedItem.SubItems(6), cnnoracle_4, adOpenDynamic, adLockOptimistic
                      Call ilumina_grid
                      Call cmd_mensaje_4_Click
                      rs.Close
                   Else
                      var_referencia_embarque = Mid(Me.txt_codigo, 2, 6)
                      var_referencia_caja = Mid(Me.txt_codigo, 8, 3)
                      var_referencia_embarque = Me.txt_embarque.Text
                      
                   If IsNumeric(var_referencia_embarque) Then
                       If IsNumeric(var_referencia_caja) Then
                          If CDbl(Me.txt_embarque) = CDbl(var_referencia_embarque) Then
                             strconsulta = "SELECT TRANSPORTE FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = ?"
                             With comandoORA
                                  'MsgBox cnnoracle_4.ConnectionString
                                  .ActiveConnection = cnnoracle_4
                                  .CommandType = adCmdText
                                  .CommandText = strconsulta
                                  Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(var_referencia_embarque))
                                  .Parameters.Append parametro
                             End With
                             Set rsaux8 = comandoORA.execute
                             Set comandoORA = Nothing
                             Set parametro = Nothing
                             var_clave_transporte = ""
                             If Not rsaux8.EOF Then
                                var_clave_transporte = IIf(IsNull(rsaux8!transporte), "", rsaux8!transporte)
                             End If
                             rsaux8.Close
                             rsaux8.Open "SELECT isnull(NOMBRE,'') nombre FROM TB_ORACLE_TRANSPORTES WHERE CLAVE = '" + var_clave_transporte + "'", cnn, adOpenDynamic, adLockOptimistic
                             var_nombre_transporte = ""
                             If Not rsaux8.EOF Then
                                var_nombre_transporte = rsaux8!nombre
                             End If
                             rsaux8.Close
                             var_posible_peso = 1
                             strconsulta = "SELECT source_header_number, inte_paq_caja, tipo_caja,  sum(floa_sal_cantidad_leida * (CASE WHEN PESO = 0 THEN 0.55 ELSE PESO END)) peso, sum(floa_sal_cantidad_leida) as cantidad FROM XXVIA_TB_SALIDAS_CAJAS where inte_emb_embarque = ? and inte_paq_caja = ? group by source_header_number, inte_paq_caja, tipo_caja"
                             With comandoORA
                                  'MsgBox cnnoracle_4.ConnectionString
                                  .ActiveConnection = cnnoracle_4
                                  .CommandType = adCmdText
                                  .CommandText = strconsulta
                                  Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(var_referencia_embarque))
                                  .Parameters.Append parametro
                                  Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(var_referencia_caja))
                                  .Parameters.Append parametro
                             End With
                             Set rsaux8 = comandoORA.execute
                             Set comandoORA = Nothing
                             Set parametro = Nothing
                             If Not rsaux8.EOF Then
                                var_tipo_empaque = rsaux8!tipo_caja
                                var_peso_empaque = IIf(IsNull(rsaux8!PESO), 0, rsaux8!PESO)
                                var_cantidad_permitida = IIf(IsNull(rsaux8!cantidad), 0, rsaux8!cantidad)
                                rsaux8.Close
                                rsaux8.Open "select * from TB_ORACLE_EMPAQUES where EMPAQUE = '" + var_tipo_empaque + "'", cnn, adOpenDynamic, adLockOptimistic
                                If Not rsaux8.EOF Then
                                   If rsaux8!peso_permitido = 0 Then
                                      var_posible_peso = 1
                                   Else
                                      If var_peso_empaque > rsaux8!peso_permitido Then
                                         var_posible_peso = 1
                                      Else
                                         var_posible_peso = 0
                                      End If
                                   End If
                                Else
                                   var_posible_peso = 1
                                End If
                                rsaux8.Close
                             Else
                                var_posible_peso = 1
                                rsaux8.Close
                             End If
                             If var_cantidad_permitida < 2 Then
                                'var_posible_peso = 0
                             End If
                             If var_posible_peso = 1 Then
                             rs.Open "select * from tb_oracle_cajas_aduana where embarque  = " + CStr(var_referencia_embarque) + " and CAJA = '" + CStr(Me.txt_codigo) + "'", cnn, adOpenDynamic, adLockOptimistic
                             If Not rs.EOF Then
                                var_estatus_caja = IIf(IsNull(rs!estatus), "", rs!estatus)
                                If var_estatus_caja = "" Then
                                   var_embarque_auditar = CDbl(Me.txt_embarque)
                                   var_caja_auditar = CDbl(var_referencia_caja)
                                   Me.lbl_tipo_bulto = Me.lv_cajas.selectedItem.SubItems(7) + " " + Me.lv_cajas.selectedItem.SubItems(8)
                                   If Me.lv_cajas.selectedItem.SubItems(7) = "COSTAL GRANDE" Or Me.lv_cajas.selectedItem.SubItems(7) = "COSTAL CHICO" Or Me.lv_cajas.selectedItem.SubItems(7) = "CAJA  SOBRE-CAJA" Or Me.lv_cajas.selectedItem.SubItems(7) = "CAJA CHICA" Or Me.lv_cajas.selectedItem.SubItems(7) = "CAJA CORTINERO" Or Me.lv_cajas.selectedItem.SubItems(7) = "CAJA EXTRAGRANDE" Or Me.lv_cajas.selectedItem.SubItems(7) = "CAJA GRANDE" Or Me.lv_cajas.selectedItem.SubItems(7) = "CAJA MEDIANA" Or Me.lv_cajas.selectedItem.SubItems(7) = "CAJA MINI/CATALOGO" Or Me.lv_cajas.selectedItem.SubItems(7) = "COSTAL CHICO" Or Me.lv_cajas.selectedItem.SubItems(7) = "COSTAL GRANDE" Or Me.lv_cajas.selectedItem.SubItems(7) = "EMPLAYE CORTINEROS" Or Me.lv_cajas.selectedItem.SubItems(7) = "OTROS" Or Me.lv_cajas.selectedItem.SubItems(7) = "OTROS MUEBLES" Or Me.lv_cajas.selectedItem.SubItems(7) = "PAQUETE BOLSA" Or Me.lv_cajas.selectedItem.SubItems(7) = "PAQUETE PUBLICIDAD" Then
                                   
                                      rsaux10.Open "select TO_CHAR(dbms_random.value(1,100), '999') as numb from dual", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                      If rsaux10(0).Value <= 5 Then
                                         var_auditar = 1
                                      Else
                                         var_auditar = 0
                                      End If
                                      rsaux10.Close
                                   Else
                                      var_auditar = 0
                                   End If
                                   If var_auditar = 1 Then
                                      rsaux.Open "SELECT NVL(AUDITADA,0) AS AUDITADA FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE  = " + CStr(var_referencia_embarque) + " AND INTE_PAQ_CAJA = " + CStr(var_referencia_caja), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                      If rsaux!auditada = 1 Then
                                         var_auditar = 0
                                      End If
                                      rsaux.Close
                                    End If
                                    On Error GoTo SALIR:
                                    If var_prueba = 0 Then
                                       Set clnt = Nothing
                                       clnt.MSSoapInit var_webservice_texto
                                       var_s = clnt.insertar_texto(CStr(var_dvr_texto_ip), CStr(var_puerto_texto), "MAQUINA: " + fun_NombrePc + ", USUARIO: " + var_nombre_usuario + Chr(13) + "EMBARQUE: " + Me.txt_embarque + "-CAJA: " + Me.txt_codigo + "-TIPO BIULTO: " + Me.lbl_tipo_bulto)
                                    Set clnt = Nothing
                                    End If
                                    'var_auditar = 1
                                   
                                   If var_auditar = 1 Then
                                   
                                      var_guia_aduana = ""
                                      var_posible_guia = 1
                                      If Me.lbl_paqueteria <> "" Then
                                         frmoracle_guia.Show 1
                                         
                                         var_cadena = "select formNo AS GUIA from xxvia_Tb_guias_generadas_2 WHERE formNo  = '" + var_guia_aduana + "' Union All select NumeroGuia as GUIA from xxvia_tb_guias_estafeta WHERE NumeroGuia = '" + var_guia_aduana + "'"
                                         var_cadena = "select formNo as GUIA from xxvia_Tb_guias_generadas_2 where substring(formno,1,7)+'0'+SUBSTRING(formno,8,6) = substring('" + var_guia_aduana + "',1,14)  UNION ALL "
                                         var_cadena = var_cadena + " select numeroguia as GUIA from xxvia_tb_guias_estafeta where NumeroGuia = '" + var_guia_aduana + "'"
                                         rsaux13.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                         If rsaux13.EOF Then
                                            strconsulta = "SELECT * FROM XXVIA.XXVIA_TB_PAQUETERIAS_GUIAS where vcha_guia=? or numb_rastreo = ?"
                                            With comandoORA
                                                 'MsgBox cnnoracle_4.ConnectionString
                                                 .ActiveConnection = cnnoracle_4
                                                 .CommandType = adCmdText
                                                 .CommandText = strconsulta
                                                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_guia_aduana)
                                                 .Parameters.Append parametro
                                                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_guia_aduana)
                                                 .Parameters.Append parametro
                                            End With
                                            Set rsaux8 = comandoORA.execute
                                            Set comandoORA = Nothing
                                            Set parametro = Nothing
                                            If Not rsaux8.EOF Then
                                               var_guia_caja = IIf(IsNull(rsaux8!NUMB_RASTREO), "", rsaux8!NUMB_RASTREO)
                                               var_posible_guia = 1
                                            Else
                                               var_posible_guia = 2
                                            End If
                                            rsaux8.Close
                                         Else
                                            
                                            var_guia_caja_aduana = IIf(IsNull(rsaux13!Guia), "", rsaux13!Guia)
                                            strconsulta = "SELECT * FROM XXVIA.XXVIA_TB_PAQUETERIAS_GUIAS where vcha_guia=? or numb_rastreo = ?"
                                            With comandoORA
                                                 'MsgBox cnnoracle_4.ConnectionString
                                                 .ActiveConnection = cnnoracle_4
                                                 .CommandType = adCmdText
                                                 .CommandText = strconsulta
                                                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_guia_caja_aduana)
                                                 .Parameters.Append parametro
                                                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_guia_caja_aduana)
                                                 .Parameters.Append parametro
                                            End With
                                            Set rsaux8 = comandoORA.execute
                                            Set comandoORA = Nothing
                                            Set parametro = Nothing
                                            If Not rsaux8.EOF Then
                                               var_posible_guia = 3
                                               While Not rsaux8.EOF
                                                     var_caja_aduana = "C" + IIf(IsNull(rsaux8!vcha_caja_id), "", rsaux8!vcha_caja_id)
                                                     If var_caja_aduana = Me.txt_codigo Then
                                                        var_posible_guia = 1
                                                     End If
                                                     rsaux8.MoveNext
                                               Wend
                                            Else
                                               var_posible_guia = 2
                                            End If
                                            rsaux8.Close
                                         End If
                                         rsaux13.Close
                                      End If
                                      If Me.lbl_paqueteria = "" Then
                                         var_posible_guia = 1
                                      Else
                                         If var_guia_aduana <> "" Then
                                            If var_posible_guia = 2 Or var_posible_guia = 3 Then
                                            Else
                                               var_posible_guia = 1
                                            End If
                                         Else
                                            var_posible_guia = 0
                                         End If
                                      End If
                                      If Me.lv_cajas.selectedItem.SubItems(7) <> "OTROS" Then
                                         If var_posible_guia = 2 Or var_posible_guia = 3 Then
                                            var_guia_aduana = ""
                                         End If
                                      Else
                                         var_posible_guia = 1
                                      End If
                                      
                                      
                                      
                                      
                                      If var_posible_guia = 1 Then
                                         rsaux1.Open "update tb_oracle_cajas_aduana set estatus = 'L', guia = '" + var_guia_aduana + "' where embarque = " + CStr(Me.txt_embarque) + " and numero_caja = " + CStr(var_caja_auditar), cnn, adOpenDynamic, adLockOptimistic
                                         strconsulta = "update xxvia_Tb_salidas_cajas set guia = ?, transporte = ? where inte_emb_embarque = ? and inte_paq_caja = ?"
                                         With comandoORA
                                              'MsgBox cnnoracle_4.ConnectionString
                                              .ActiveConnection = cnnoracle_4
                                              .CommandType = adCmdText
                                              .CommandText = strconsulta
                                              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_guia_aduana)
                                              .Parameters.Append parametro
                                              
                                              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_nombre_transporte)
                                              .Parameters.Append parametro
                                              
                                              
                                              Set parametro = .CreateParameter(, adNumeric, adParamInput, 8, CDbl(Me.txt_embarque))
                                              .Parameters.Append parametro
                                              Set parametro = .CreateParameter(, adNumeric, adParamInput, 8, CDbl(var_caja_auditar))
                                              .Parameters.Append parametro
                                         End With
                                         Set rsaux8 = comandoORA.execute
                                         Set comandoORA = Nothing
                                         Set parametro = Nothing
                                   
                                         frmoracle_audita_caja.Show 1
                                         rsaux.Open "SELECT * FROM XXVIA_TB_cAJAS_AUDITADAS WHERE EMBARQUE = " + CStr(var_embarque_auditar) + " AND CAJA = " + CStr(var_caja_auditar) + " AND CANTIDAD_ORIGINAL <> CANTIDAD_AUDITADA", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                         If rsaux.EOF Then
                                            If rsaux1.State = 1 Then
                                               rsaux1.Close
                                            End If
                                            var_observaciones_auditoria = "NO HUBO DIFERENCIAS"
                                            'frmoracle_audita_observaciones.Show 1
                                            frmoracle_sello_caja.Show 1
                                            
                                            rsaux1.Open "update xxvia_tb_salidas_cajas set sello = '" + var_sello_caja + "', AUDITADA = 1, OBSERVACIONES_AUDITORIA = '" + var_observaciones_auditoria + "' where inte_emb_embarque = " + CStr(Me.txt_embarque) + " and inte_paq_caja = " + CStr(var_caja_auditar), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                            rsaux1.Open "update tb_oracle_cajas_aduana set estatus = 'L' where embarque = " + CStr(Me.txt_embarque) + " and numero_caja = " + CStr(var_caja_auditar), cnn, adOpenDynamic, adLockOptimistic
                                            rsaux1.Open "select * from tb_oracle_tiempo_pedido_aduanas where pedido = " + CStr(var_pedido_tiempo), cnn, adOpenDynamic, adLockOptimistic
                                            If rsaux1.EOF Then
                                               rsaux10.Open "insert into tb_oracle_tiempo_pedido_aduanas (pedido, hora_inicio) values (" + CStr(var_pedido_tiempo) + ",getdate())", cnn, adOpenDynamic, adLockOptimistic
                                            End If
                                            rsaux1.Close
                                            rsaux1.Open "select * from tb_oracle_tiempo_embarque_aduanas where embarque = " + CStr(Me.txt_embarque), cnn, adOpenDynamic, adLockOptimistic
                                            If rsaux1.EOF Then
                                               rsaux10.Open "insert into tb_oracle_tiempo_embarque_aduanas (embarque, hora_inicio) values (" + CStr(Me.txt_embarque) + ",getdate())", cnn, adOpenDynamic, adLockOptimistic
                                            End If
                                            rsaux1.Close
                                            
                                            rsaux1.Open "select source_header_number, inte_paq_caja, char_paq_estatus, sum(floa_Sal_Cantidad_leida) as cantidad from xxvia_tb_salidas_cajas where inte_emb_embarque = " + CStr(Me.txt_embarque) + " and inte_paq_caja = " + CStr(var_caja_auditar) + " group by source_header_number, inte_paq_caja, char_paq_estatus", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                            var_codigo_caja = "C" + var_referencia_embarque + var_referencia_caja
                                            var_estatus_caja = IIf(IsNull(rsaux1!char_paq_estatus), "", rsaux1!char_paq_estatus)
                                            var_cantidad_enviada = var_cantidad_enviada + rsaux1!cantidad
                                            var_cantidad_leida = 0
                                            Me.txt_codigo = ""
                                            Me.lv_cajas.selectedItem.SubItems(5) = "L"
                                            'rsaux2.Open "update xxvia_tb_salidas_cajas SET CHAR_PAQ_ESTATUS = 'S' where inte_emb_embarque = " + Me.txt_embarque + " and source_header_number = " + Me.lv_cajas.selectedItem.SubItems(1) + " and inte_paq_caja = " + Me.lv_cajas.selectedItem.SubItems(6), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                            Call ilumina_grid
                                            Call cmd_mensaje_4_Click
                                         Else
                                            frmmensaje.lbl_mensaje = "Existen diferencias en la caja auditada"
                                            frmmensaje.Show 1
                                            var_observaciones_auditoria = ""
                                            frmoracle_audita_observaciones.Show 1
                                            frmoracle_sello_caja.Show 1
                                            rsaux8.Open "UPDATE XXVIA_TB_SALIDAS_CAJAS SET AUDITADA = 1, SELLO = '" + var_sello_caja + "', OBSERVACIONES_AUDITORIA = '" + var_observaciones_auditoria + "'  where inte_emb_embarque  = " + CStr(var_referencia_embarque) + " and inte_paq_caja = " + CStr(var_referencia_caja), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                            rsaux1.Open "update tb_oracle_cajas_aduana set estatus = 'L' where embarque = " + CStr(Me.txt_embarque) + " and numero_caja = " + CStr(var_caja_auditar), cnn, adOpenDynamic, adLockOptimistic
                                            
                                            rsaux1.Open "select * from tb_oracle_tiempo_pedido_aduanas where pedido = " + CStr(var_pedido_tiempo), cnn, adOpenDynamic, adLockOptimistic
                                            If rsaux1.EOF Then
                                               rsaux10.Open "insert into tb_oracle_tiempo_pedido_aduanas (pedido, hora_inicio) values (" + CStr(var_pedido_tiempo) + ",getdate())", cnn, adOpenDynamic, adLockOptimistic
                                            End If
                                            rsaux1.Close
                                            rsaux1.Open "select * from tb_oracle_tiempo_embarque_aduanas where embarque = " + CStr(Me.txt_embarque), cnn, adOpenDynamic, adLockOptimistic
                                            If rsaux1.EOF Then
                                               rsaux10.Open "insert into tb_oracle_tiempo_embarque_aduanas (embarque, hora_inicio) values (" + CStr(Me.txt_embarque) + ",getdate())", cnn, adOpenDynamic, adLockOptimistic
                                            End If
                                            rsaux1.Close
                                         
                                         
                                            txt_codigo.SetFocus
                                            var_orden_surtido = 0
                                            var_caja = 0
                                            var_factura_ceros = 0
                                            var_tipo_pedido = ""
                                            Me.txt_codigo = ""
                                            Me.txt_codigo.SetFocus
                                            Call cmd_mensaje_4_Click
                                         End If
                                         rsaux.Close
                                      Else
                                         If var_posible_guia = 2 Then
                                            frmmensaje.lbl_mensaje = "Guia incorrecta."
                                         Else
                                            If var_posible_guia = 3 Then
                                               frmmensaje.lbl_mensaje = "La guia no corresponde a la caja seleccionada"
                                            Else
                                               frmmensaje.lbl_mensaje = "No se selecciono la guia."
                                            End If
                                         End If
                                         Me.lbl_tipo_bulto = ""
                                         frmmensaje.Show 1
                                         Me.lbl_tipo_bulto = ""
                                         frmmensaje.Show 1
                                         txt_codigo.SetFocus
                                         var_orden_surtido = 0
                                         var_caja = 0
                                         var_factura_ceros = 0
                                         var_tipo_pedido = ""
                                         Me.txt_codigo = ""
                                         Me.txt_codigo.SetFocus
                                      End If
                                   Else
                                      If rsaux1.State = 1 Then
                                         rsaux1.Close
                                      End If
                                      'rsaux1.Open "update xxvia_tb_salidas_cajas set char_paq_estatus = 'S'where inte_emb_embarque = " + CStr(Me.txt_embarque) + " and inte_paq_caja = " + CStr(var_caja_auditar), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                      var_guia_aduana = ""
                                      var_posible_guia = 1
                                      If Me.lbl_paqueteria <> "" Then
                                         frmoracle_guia.Show 1
                                         
                                         var_cadena = "select formNo AS GUIA from xxvia_Tb_guias_generadas_2 WHERE formNo  = '" + var_guia_aduana + "' Union All select NumeroGuia as GUIA from xxvia_tb_guias_estafeta WHERE NumeroGuia = '" + var_guia_aduana + "'"
                                         var_cadena = "select formNo as GUIA from xxvia_Tb_guias_generadas_2 where substring(formno,1,7)+'0'+SUBSTRING(formno,8,6) = substring('" + var_guia_aduana + "',1,14)  UNION ALL "
                                         var_cadena = var_cadena + " select numeroguia as GUIA from xxvia_tb_guias_estafeta where NumeroGuia = '" + var_guia_aduana + "'"
                                         rsaux13.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                         If rsaux13.EOF Then
                                            strconsulta = "SELECT * FROM XXVIA.XXVIA_TB_PAQUETERIAS_GUIAS where vcha_guia=? or numb_rastreo = ?"
                                            With comandoORA
                                                 'MsgBox cnnoracle_4.ConnectionString
                                                 .ActiveConnection = cnnoracle_4
                                                 .CommandType = adCmdText
                                                 .CommandText = strconsulta
                                                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_guia_aduana)
                                                 .Parameters.Append parametro
                                                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_guia_aduana)
                                                 .Parameters.Append parametro
                                            End With
                                            Set rsaux8 = comandoORA.execute
                                            Set comandoORA = Nothing
                                            Set parametro = Nothing
                                            If Not rsaux8.EOF Then
                                               var_guia_caja = IIf(IsNull(rsaux8!NUMB_RASTREO), "", rsaux8!NUMB_RASTREO)
                                               var_posible_guia = 1
                                            Else
                                               var_posible_guia = 2
                                            End If
                                            rsaux8.Close
                                         Else
                                            
                                            var_guia_caja_aduana = IIf(IsNull(rsaux13!Guia), "", rsaux13!Guia)
                                            strconsulta = "SELECT * FROM XXVIA.XXVIA_TB_PAQUETERIAS_GUIAS where vcha_guia=? or numb_rastreo = ?"
                                            With comandoORA
                                                 'MsgBox cnnoracle_4.ConnectionString
                                                 .ActiveConnection = cnnoracle_4
                                                 .CommandType = adCmdText
                                                 .CommandText = strconsulta
                                                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_guia_caja_aduana)
                                                 .Parameters.Append parametro
                                                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_guia_caja_aduana)
                                                 .Parameters.Append parametro
                                            End With
                                            Set rsaux8 = comandoORA.execute
                                            Set comandoORA = Nothing
                                            Set parametro = Nothing
                                            If Not rsaux8.EOF Then
                                               var_posible_guia = 3
                                               While Not rsaux8.EOF
                                                     var_caja_aduana = "C" + IIf(IsNull(rsaux8!vcha_caja_id), "", rsaux8!vcha_caja_id)
                                                     If var_caja_aduana = Me.txt_codigo Then
                                                        var_posible_guia = 1
                                                     End If
                                                     rsaux8.MoveNext
                                               Wend
                                            Else
                                               var_posible_guia = 2
                                            End If
                                            rsaux8.Close
                                         End If
                                         rsaux13.Close
                                      End If
                                      If Me.lbl_paqueteria = "" Then
                                         var_posible_guia = 1
                                      Else
                                         If var_guia_aduana <> "" Then
                                            If var_posible_guia = 2 Or var_posible_guia = 3 Then
                                            Else
                                               var_posible_guia = 1
                                            End If
                                         Else
                                            var_posible_guia = 0
                                         End If
                                      End If
                                      
                                      If Me.lv_cajas.selectedItem.SubItems(7) <> "OTROS" Then
                                         If var_posible_guia = 2 Or var_posible_guia = 3 Then
                                            var_guia_aduana = ""
                                         End If
                                      Else
                                         var_posible_guia = 1
                                      End If
                                      
                                      If var_posible_guia = 1 Then
                                         rsaux1.Open "update tb_oracle_cajas_aduana set estatus = 'L', guia = '" + var_guia_aduana + "' where embarque = " + CStr(Me.txt_embarque) + " and numero_caja = " + CStr(var_caja_auditar), cnn, adOpenDynamic, adLockOptimistic
                                         
                                         
                                         strconsulta = "update xxvia_Tb_salidas_cajas set guia = ?, transporte = ? where inte_emb_embarque = ? and inte_paq_caja = ?"
                                         With comandoORA
                                              'MsgBox cnnoracle_4.ConnectionString
                                              .ActiveConnection = cnnoracle_4
                                              .CommandType = adCmdText
                                              .CommandText = strconsulta
                                              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_guia_aduana)
                                              .Parameters.Append parametro
                                              
                                              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_nombre_transporte)
                                              .Parameters.Append parametro
                                              
                                              
                                              Set parametro = .CreateParameter(, adNumeric, adParamInput, 8, CDbl(Me.txt_embarque))
                                              .Parameters.Append parametro
                                              Set parametro = .CreateParameter(, adNumeric, adParamInput, 8, CDbl(var_caja_auditar))
                                              .Parameters.Append parametro
                                         End With
                                         Set rsaux8 = comandoORA.execute
                                         Set comandoORA = Nothing
                                         Set parametro = Nothing
                                         
                                         
                                         
                                         
                                         
                                         
                                         
                                         rsaux1.Open "select * from tb_oracle_tiempo_pedido_aduanas where pedido = " + CStr(var_pedido_tiempo), cnn, adOpenDynamic, adLockOptimistic
                                         If rsaux1.EOF Then
                                            rsaux10.Open "insert into tb_oracle_tiempo_pedido_aduanas (pedido, hora_inicio) values (" + CStr(var_pedido_tiempo) + ",getdate())", cnn, adOpenDynamic, adLockOptimistic
                                         End If
                                         rsaux1.Close
                                         rsaux1.Open "select * from tb_oracle_tiempo_embarque_aduanas where embarque = " + CStr(Me.txt_embarque), cnn, adOpenDynamic, adLockOptimistic
                                         If rsaux1.EOF Then
                                            rsaux10.Open "insert into tb_oracle_tiempo_embarque_aduanas (embarque, hora_inicio) values (" + CStr(Me.txt_embarque) + ",getdate())", cnn, adOpenDynamic, adLockOptimistic
                                         End If
                                         rsaux1.Close
                                         
                                         'rsaux1.Open "select source_header_number, inte_paq_caja, char_paq_estatus, sum(floa_Sal_Cantidad_leida) as cantidad from xxvia_tb_salidas_cajas where inte_emb_embarque = " + CStr(Me.txt_embarque) + " and inte_paq_caja = " + CStr(var_caja_auditar) + " group by source_header_number, inte_paq_caja, char_paq_estatus", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                         'var_codigo_caja = "C" + var_referencia_embarque + var_referencia_caja
                                         'var_estatus_caja = IIf(IsNull(rsaux1!char_paq_estatus), "", rsaux1!char_paq_estatus)
                                         'var_cantidad_enviada = var_cantidad_enviada + rsaux1!Cantidad
                                         'var_cantidad_leida = 0
                                         Me.txt_codigo = ""
                                         Me.lv_cajas.selectedItem.SubItems(5) = "L"
                                         'rsaux2.Open "update xxvia_tb_salidas_cajas SET CHAR_PAQ_ESTATUS = 'S' where inte_emb_embarque = " + Me.txt_embarque + " and source_header_number = " + Me.lv_cajas.selectedItem.SubItems(1) + " and inte_paq_caja = " + Me.lv_cajas.selectedItem.SubItems(6), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                         Call ilumina_grid
                                         Call cmd_mensaje_4_Click
                                      Else
                                         If var_posible_guia = 2 Then
                                            frmmensaje.lbl_mensaje = "Guia incorrecta."
                                         Else
                                            If var_posible_guia = 3 Then
                                               frmmensaje.lbl_mensaje = "La guia no corresponde a la caja seleccionada"
                                            Else
                                               frmmensaje.lbl_mensaje = "No se selecciono la guia."
                                            End If
                                         End If
                                         Me.lbl_tipo_bulto = ""
                                         frmmensaje.Show 1
                                         txt_codigo.SetFocus
                                         var_orden_surtido = 0
                                         var_caja = 0
                                         var_factura_ceros = 0
                                         var_tipo_pedido = ""
                                         Me.txt_codigo = ""
                                         Me.txt_codigo.SetFocus
                                      End If
                                   End If
                                Else
                                   If var_estatus_caja = "E" Then
                                      frmmensaje.lbl_mensaje = "La caja no contiene información"
                                      Me.lbl_tipo_bulto = ""
                                      frmmensaje.Show 1
                                      txt_codigo.SetFocus
                                      var_orden_surtido = 0
                                      var_caja = 0
                                      var_factura_ceros = 0
                                      var_tipo_pedido = ""
                                      Me.txt_codigo = ""
                                      Me.txt_codigo.SetFocus
                                   End If
                                   If var_estatus_caja = "S" Then
                                      frmmensaje.lbl_mensaje = "La caja ya fue leida"
                                      frmmensaje.Show 1
                                      Me.lbl_tipo_bulto = ""
                                      txt_codigo.SetFocus
                                      var_orden_surtido = 0
                                      var_caja = 0
                                      var_factura_ceros = 0
                                      var_tipo_pedido = ""
                                      Me.txt_codigo = ""
                                      Me.txt_codigo.SetFocus
                                   End If
                                   If var_estatus_caja = "" Then
                                      frmmensaje.lbl_mensaje = "La caja no a sido cerrada aun"
                                      frmmensaje.Show 1
                                      Me.lbl_tipo_bulto = ""
                                      txt_codigo.SetFocus
                                      var_orden_surtido = 0
                                      var_caja = 0
                                      var_factura_ceros = 0
                                      var_tipo_pedido = ""
                                      Me.txt_codigo = ""
                                      Me.txt_codigo.SetFocus
                                   End If
                                End If
                             Else
                                frmmensaje.lbl_mensaje = "La caja no contiene información"
                                frmmensaje.Show 1
                                Me.lbl_tipo_bulto = ""
                                txt_codigo.SetFocus
                                var_orden_surtido = 0
                                var_caja = 0
                                var_factura_ceros = 0
                                var_tipo_pedido = ""
                                Me.txt_codigo = ""
                                Me.txt_codigo.SetFocus
                             End If
                             If rs.State = 1 Then
                                rs.Close
                             End If
                             Else
                                frmmensaje.lbl_mensaje = "El contenido del bulto es menor al permitido"
                                frmmensaje.Show 1
                                txt_codigo.SetFocus
                                Me.lbl_tipo_bulto = ""
                                var_orden_surtido = 0
                                var_caja = 0
                                var_factura_ceros = 0
                                var_tipo_pedido = ""
                                Me.txt_codigo = ""
                                Me.txt_codigo.SetFocus
                             
                             End If
                          Else
                             frmmensaje.lbl_mensaje = "Número de embarque incorrecto"
                             frmmensaje.Show 1
                             txt_codigo.SetFocus
                             Me.lbl_tipo_bulto = ""
                             var_orden_surtido = 0
                             var_caja = 0
                             var_factura_ceros = 0
                             var_tipo_pedido = ""
                             Me.txt_codigo = ""
                             Me.txt_codigo.SetFocus
                          End If
                       Else
                          frmmensaje.lbl_mensaje = "Número de caja incorrecto"
                          frmmensaje.Show 1
                          txt_codigo.SetFocus
                          Me.lbl_tipo_bulto = ""
                          var_orden_surtido = 0
                          var_caja = 0
                          var_factura_ceros = 0
                          var_tipo_pedido = ""
                          Me.txt_codigo = ""
                          Me.txt_codigo.SetFocus
                       End If
                    Else
                       frmmensaje.lbl_mensaje = "Número de embarque incorrecto"
                       frmmensaje.Show 1
                       txt_codigo.SetFocus
                       Me.lbl_tipo_bulto = ""
                       var_orden_surtido = 0
                       var_caja = 0
                       var_factura_ceros = 0
                       var_tipo_pedido = ""
                       Me.txt_codigo = ""
                       Me.txt_codigo.SetFocus
                    End If
                    End If 'de insertar cn textilera
                   '----------
                End If
             End If
         Next var_j
      End If
      var_encontro_2 = 0
      If var_encontro = 0 Then
         For var_j = 1 To lv_cajas_siguientes.ListItems.Count
             Me.lv_cajas_siguientes.ListItems.Item(var_j).Selected = True
             If Me.txt_codigo = Me.lv_cajas_siguientes.selectedItem Then
                var_encontro_2 = 1
             End If
         Next var_j
      End If
      If var_encontro_2 = 1 Then
         MsgBox "La caja no puede ser despachada ya que no corresponde al orden establecido", vbOKOnly, "ATENCION"
         Me.lbl_tipo_bulto = ""
      Else
         If var_encontro = 0 Then
            MsgBox "La caja no se encuentra", vbOKOnly, "ATENCION"
            Me.txt_codigo = ""
            Me.lbl_tipo_bulto = ""
         End If
      End If
      'Me.Timer1.Enabled = True
   End If
   Exit Sub
SALIR:
    'MsgBox Err.Description
    Resume Next
   
End Sub

Private Sub txt_sello_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_aceptar_sello.SetFocus
   End If
End Sub

