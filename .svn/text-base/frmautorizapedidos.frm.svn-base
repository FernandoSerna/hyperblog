VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmautorizapedidos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autorización de Pedidos"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   Icon            =   "frmautorizapedidos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11670
   Begin MSComCtl2.MonthView mes_filtro 
      Height          =   2370
      Left            =   3045
      TabIndex        =   52
      Top             =   -90
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   75563009
      CurrentDate     =   37727
   End
   Begin VB.Frame frm_facturas 
      Height          =   2925
      Left            =   5505
      TabIndex        =   45
      Top             =   2040
      Width           =   5400
      Begin VB.TextBox txt_total_facturas_venciadas 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   3240
         TabIndex        =   74
         Top             =   2460
         Width           =   2070
      End
      Begin MSComctlLib.ListView lv_facturas 
         Height          =   2025
         Left            =   75
         TabIndex        =   46
         Top             =   405
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   3572
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Factura"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Vencimiento"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Importe"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
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
         Left            =   2355
         TabIndex        =   75
         Top             =   2475
         Width           =   795
      End
      Begin VB.Label Label16 
         BackColor       =   &H8000000D&
         Caption         =   "Detalle de Facturación Vencida"
         ForeColor       =   &H80000005&
         Height          =   225
         Left            =   30
         TabIndex        =   47
         Top             =   120
         Width           =   5325
      End
   End
   Begin VB.Frame frm_correo 
      Height          =   1335
      Left            =   2220
      TabIndex        =   63
      Top             =   240
      Width           =   4440
      Begin VB.CommandButton cmd_aceptar_correo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   30
         Picture         =   "frmautorizapedidos.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Aceptar el envio de correo"
         Top             =   345
         Width           =   330
      End
      Begin VB.CommandButton cmd_Cancelar_corro 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   360
         Picture         =   "frmautorizapedidos.frx":0A14
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Cancelar el envio de correo"
         Top             =   345
         Width           =   330
      End
      Begin VB.Frame Frame5 
         Height          =   45
         Left            =   15
         TabIndex        =   71
         Top             =   660
         Width           =   4395
      End
      Begin VB.TextBox txt_inicio_correo 
         Height          =   330
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   810
         Width           =   1080
      End
      Begin VB.TextBox txt_fin_correo 
         Height          =   330
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   810
         Width           =   1080
      End
      Begin VB.CommandButton Command11 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1725
         Picture         =   "frmautorizapedidos.frx":0B5E
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Ejecutar Calendario"
         Top             =   825
         Width           =   330
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3780
         Picture         =   "frmautorizapedidos.frx":1DD0
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Ejecutar Calendario Alt + E"
         Top             =   825
         Width           =   330
      End
      Begin VB.Label Label24 
         BackColor       =   &H8000000D&
         Caption         =   " Rango de fecha para envio de correo"
         ForeColor       =   &H80000005&
         Height          =   225
         Left            =   30
         TabIndex        =   70
         Top             =   120
         Width           =   4365
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Del:"
         Height          =   195
         Left            =   300
         TabIndex        =   69
         Top             =   885
         Width           =   285
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Al:"
         Height          =   195
         Left            =   2340
         TabIndex        =   68
         Top             =   885
         Width           =   180
      End
   End
   Begin VB.CommandButton cmd_correo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1125
      Picture         =   "frmautorizapedidos.frx":3042
      Style           =   1  'Graphical
      TabIndex        =   62
      ToolTipText     =   "Enviar correo de pedidos autorizados"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame frm_cheques_devueltos 
      Height          =   2700
      Left            =   3960
      TabIndex        =   25
      Top             =   1965
      Width           =   6705
      Begin MSComctlLib.ListView lv_cheques 
         Height          =   2160
         Left            =   60
         TabIndex        =   27
         Top             =   450
         Width           =   6570
         _ExtentX        =   11589
         _ExtentY        =   3810
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cheque"
            Object.Width           =   1605
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Banco"
            Object.Width           =   4762
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Importe"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000D&
         Caption         =   " Cheques Devueltos"
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   30
         TabIndex        =   26
         Top             =   120
         Width           =   6630
      End
   End
   Begin VB.Frame frm_detalle_pedido 
      Height          =   4605
      Left            =   1710
      TabIndex        =   53
      Top             =   2640
      Width           =   7050
      Begin MSComctlLib.ListView lv_detalle_pedido 
         Height          =   4125
         Left            =   75
         TabIndex        =   54
         Top             =   405
         Width           =   6900
         _ExtentX        =   12171
         _ExtentY        =   7276
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   2293
         EndProperty
      End
      Begin VB.Label Label20 
         BackColor       =   &H8000000D&
         Caption         =   "Detalle del Pedido"
         ForeColor       =   &H80000005&
         Height          =   225
         Left            =   30
         TabIndex        =   55
         Top             =   120
         Width           =   6990
      End
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   795
      Picture         =   "frmautorizapedidos.frx":32C4
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11145
      Picture         =   "frmautorizapedidos.frx":33C6
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   465
      Picture         =   "frmautorizapedidos.frx":3A00
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Actualizar Alt + A"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      Picture         =   "frmautorizapedidos.frx":3B02
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Filtro Alt + F"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame frm_filtro 
      Height          =   1485
      Left            =   120
      TabIndex        =   48
      Top             =   390
      Width           =   4440
      Begin VB.CommandButton Command4 
         Caption         =   "Cerrar"
         Height          =   435
         Left            =   3270
         TabIndex        =   59
         Top             =   885
         Width           =   1035
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cancelados"
         Height          =   435
         Left            =   2220
         TabIndex        =   58
         Top             =   885
         Width           =   1035
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Autorizados"
         Height          =   435
         Left            =   1170
         TabIndex        =   57
         Top             =   885
         Width           =   1035
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Sin Autorizar"
         Height          =   435
         Left            =   120
         TabIndex        =   56
         Top             =   885
         Width           =   1035
      End
      Begin VB.CommandButton cmd_calendario_2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3780
         Picture         =   "frmautorizapedidos.frx":3BFC
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Ejecutar Calendario Alt + E"
         Top             =   495
         Width           =   330
      End
      Begin VB.CommandButton cmd_calendario_1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1725
         Picture         =   "frmautorizapedidos.frx":4E6E
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Ejecutar Calendario"
         Top             =   495
         Width           =   330
      End
      Begin VB.TextBox txt_fecha_fin 
         Height          =   330
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   480
         Width           =   1080
      End
      Begin VB.TextBox txt_fecha_inicio 
         Height          =   330
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   480
         Width           =   1080
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Al:"
         Height          =   195
         Left            =   2340
         TabIndex        =   51
         Top             =   555
         Width           =   180
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Del:"
         Height          =   195
         Left            =   300
         TabIndex        =   50
         Top             =   555
         Width           =   285
      End
      Begin VB.Label Label17 
         BackColor       =   &H8000000D&
         Caption         =   " Filtrado por:"
         ForeColor       =   &H80000005&
         Height          =   225
         Left            =   30
         TabIndex        =   49
         Top             =   120
         Width           =   4365
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Pedido "
      Height          =   3090
      Left            =   135
      TabIndex        =   28
      Top             =   390
      Width           =   11415
      Begin VB.TextBox txt_observacion 
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   2640
         Width           =   9885
      End
      Begin VB.TextBox txt_importe_vencido 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8205
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   2295
         Width           =   2040
      End
      Begin VB.TextBox txt_numero 
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   210
         Width           =   1485
      End
      Begin VB.TextBox txt_agente 
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   900
         Width           =   5145
      End
      Begin VB.TextBox txt_titular 
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1245
         Width           =   5145
      End
      Begin VB.TextBox txt_establecimiento 
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1590
         Width           =   5145
      End
      Begin VB.TextBox txt_cliente 
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1950
         Width           =   5145
      End
      Begin VB.TextBox txt_descuento1 
         Height          =   315
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2295
         Width           =   1020
      End
      Begin VB.TextBox txt_descuento2 
         Height          =   315
         Left            =   5445
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2295
         Width           =   1020
      End
      Begin VB.TextBox txt_fecha 
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   555
         Width           =   1935
      End
      Begin VB.CheckBox chk_autorizado 
         Caption         =   "Autorizado"
         Enabled         =   0   'False
         Height          =   315
         Left            =   8205
         TabIndex        =   16
         Top             =   1260
         Width           =   1140
      End
      Begin VB.TextBox txt_cantidad 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8205
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   555
         Width           =   795
      End
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   8190
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   900
         Width           =   1545
      End
      Begin VB.TextBox txt_fecha_autorizacion 
         Height          =   315
         Left            =   8205
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1950
         Width           =   2040
      End
      Begin VB.TextBox txt_autorizo 
         Height          =   315
         Left            =   8205
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1590
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8205
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   210
         Width           =   2055
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   10830
         TabIndex        =   29
         Top             =   1935
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
               Object.ToolTipText     =   "Cheques Devueltos"
               ImageIndex      =   16
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1095
         Top             =   -210
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   20
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmautorizapedidos.frx":60E0
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmautorizapedidos.frx":69BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmautorizapedidos.frx":7294
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmautorizapedidos.frx":7830
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmautorizapedidos.frx":810C
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmautorizapedidos.frx":89E6
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmautorizapedidos.frx":92C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmautorizapedidos.frx":93D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmautorizapedidos.frx":94E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmautorizapedidos.frx":95F6
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmautorizapedidos.frx":9708
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmautorizapedidos.frx":981A
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmautorizapedidos.frx":992C
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmautorizapedidos.frx":9E6E
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmautorizapedidos.frx":A3B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmautorizapedidos.frx":A4C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmautorizapedidos.frx":A5D4
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmautorizapedidos.frx":A6E6
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmautorizapedidos.frx":A7F0
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmautorizapedidos.frx":A902
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar4 
         Height          =   330
         Left            =   10845
         TabIndex        =   44
         Top             =   2280
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
               Object.ToolTipText     =   "Detalle de Facturas Vencidas"
               ImageIndex      =   17
            EndProperty
         EndProperty
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Observación:"
         Height          =   195
         Left            =   150
         TabIndex        =   61
         Top             =   2700
         Width           =   945
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Importe Vencido:"
         Height          =   195
         Left            =   6540
         TabIndex        =   43
         Top             =   2355
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   150
         TabIndex        =   42
         Top             =   270
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   150
         TabIndex        =   41
         Top             =   960
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Titular:"
         Height          =   195
         Left            =   150
         TabIndex        =   40
         Top             =   1305
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Establecimiento:"
         Height          =   195
         Left            =   150
         TabIndex        =   39
         Top             =   1650
         Width           =   1155
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   150
         TabIndex        =   38
         Top             =   2010
         Width           =   525
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Descuento por Volumen:"
         Height          =   195
         Left            =   150
         TabIndex        =   37
         Top             =   2355
         Width           =   1755
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Descuento por Pago Correcto:"
         Height          =   195
         Left            =   3270
         TabIndex        =   36
         Top             =   2355
         Width           =   2160
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Piezas:"
         Height          =   195
         Left            =   6540
         TabIndex        =   35
         Top             =   615
         Width           =   510
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
         Height          =   210
         Left            =   7575
         TabIndex        =   34
         Top             =   960
         Width           =   570
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   150
         TabIndex        =   33
         Top             =   615
         Width           =   495
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Autorización:"
         Height          =   195
         Left            =   6525
         TabIndex        =   32
         Top             =   2010
         Width           =   1635
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Autorizo:"
         Height          =   195
         Left            =   6540
         TabIndex        =   31
         Top             =   1650
         Width           =   615
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Limite de Crédito:"
         Height          =   195
         Left            =   6540
         TabIndex        =   30
         Top             =   270
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   105
      TabIndex        =   24
      Top             =   255
      Width           =   11430
   End
   Begin MSComctlLib.ListView lv_pedidos 
      Height          =   3690
      Left            =   150
      TabIndex        =   0
      Top             =   3540
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   6509
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
      NumItems        =   20
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Numero"
         Object.Width           =   1640
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Agente"
         Object.Width           =   3440
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Titular"
         Object.Width           =   3440
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Establecimiento"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Cliente     "
         Object.Width           =   3819
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Piezas           "
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Importe          "
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Autorizado"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "autorizo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Fecha autorizo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Fecha Pedido"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Estatus"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Descuento 1"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Descuento 2"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Nombre Autorizo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Clave cliente"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Importe vencido"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Grupo Actual"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Observaciones"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "No activo"
         Object.Width           =   0
      EndProperty
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   615
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Menu mnu_menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu mnu_autorizar 
         Caption         =   "Autorizar"
      End
      Begin VB.Menu mnu_cancelar_autorizacion 
         Caption         =   "Cancelar Autorización"
      End
      Begin VB.Menu mnu_separador2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_cancelar_pedido 
         Caption         =   "Cancelar Pedido"
      End
      Begin VB.Menu mnu_reactivar_pedido 
         Caption         =   "Reactivar Pedido"
      End
   End
End
Attribute VB_Name = "frmautorizapedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_filtro As Boolean
Dim var_mes As Integer
Dim var_fecha_inicio As Date
Dim var_fecha_fin As Date
Dim var_clave_cliente As String

Private Sub cmd_aceptar_correo_Click()
   Dim var_dia_1 As String
   Dim var_mes_1 As String
   Dim var_año_1 As String
   
   var_dia_1 = CStr(Day(Me.txt_inicio_correo))
   var_mes_1 = CStr(Month(txt_inicio_correo))
   var_año_1 = CStr(Year(txt_inicio_correo))
   If Len(Trim(var_dia_1)) = 1 Then
      var_dia_1 = "0" + var_dia_1
   End If
   If Len(Trim(var_mes_1)) = 1 Then
      var_mes_1 = "0" + var_mes_1
   End If
   var_fecha_inicio_1 = "{d '" + var_año_1 + "-" + var_mes_1 + "-" + var_dia_1 + "'}"
   
   var_dia_1 = CStr(Day(Me.txt_fin_correo))
   var_mes_1 = CStr(Month(txt_fin_correo))
   var_año_1 = CStr(Year(txt_fin_correo))
   If Len(Trim(var_dia_1)) = 1 Then
      var_dia_1 = "0" + var_dia_1
   End If
   If Len(Trim(var_mes_1)) = 1 Then
      var_mes_1 = "0" + var_mes_1
   End If
   var_fecha_fin_1 = "{d '" + var_año_1 + "-" + var_mes_1 + "-" + var_dia_1 + "'}"

   
   rs.Open "select distinct vcha_age_Agente_id from tb_encabezado_pedidos where dtim_ped_autorizo >= " + var_fecha_inicio_1 + " and dtim_ped_autorizo <= " + var_fecha_fin_1 + "+1-.0000001", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         If rsaux.State = 1 Then
            rsaux.Close
         End If
         rsaux.Open "select vcha_Age_nombre, VCHA_AGE_EMAIL from tb_agentes where vcha_Age_agente_id = '" + rs!VCHA_AGE_AGENTE_ID + "'", cnn, adOpenDynamic, adLockOptimistic
         var_si = MsgBox("¿Desea enviar el correo al agente " + CStr(rsaux!VCHA_AGE_NOMBRE) + "?", vbYesNo, "ATENCION")
         If var_si = 6 Then
             var_correo_electronico = IIf(IsNull(rsaux!VCHA_AGE_EMAIL), "", rsaux!VCHA_AGE_EMAIL)
              If Trim(var_correo_electronico) <> "" Then
                 If MAPISession1.SessionID = 0 Then
                    MAPISession1.SignOn
                 End If
                 'var_correo_electronico = "fserna@vianney.com.mx"
                 MAPIMessages1.SessionID = MAPISession1.SessionID
                 MAPIMessages1.Compose
                 MAPIMessages1.RecipDisplayName = var_correo_electronico
                 MAPIMessages1.RecipAddress = var_correo_electronico
                 MAPIMessages1.AddressResolveUI = True
                 MAPIMessages1.ResolveName
                 MAPIMessages1.MsgSubject = "Pedidos autorizados del  " + Me.txt_inicio_correo + " al " + Me.txt_fin_correo
                 MAPIMessages1.MsgNoteText = "Se anexa archivo con información de los pedidos autorizados del " + Me.txt_inicio_correo + " al " + Me.txt_fin_correo
                 var_Archivo = App.Path & "\Pedidos_agente_" + Trim(CStr(IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID))) + ".txt"
                 Open (App.Path & "\Pedidos_agente_" + Trim(CStr(IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID))) + ".txt") For Output As #1
                 Print #1, "PEDIDOS AUTORIZADOS DEL " + Me.txt_inicio_correo + " AL " + Me.txt_fin_correo
                 Print #1, "========================================================================================================"
                 Print #1, "CLIENTE    NOMBRE                                                      PEDIDO    ATURIZADO    PIEZAS    "
                 Print #1, "========================================================================================================"
                 rsaux2.Open "select distinct inte_ped_numero, vcha_cli_clave_id, vcha_Cli_nombre, cantidad, importe, ISNULL(inte_ped_autorizo,0) AS AUTORIZO  from vw_suma_pedidos_2 where dtim_ped_autorizo >= " + var_fecha_inicio_1 + " and dtim_ped_autorizo <= " + var_fecha_fin_1 + "+1-.0000001 and vcha_age_agente_id = '" + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
                 While Not rsaux2.EOF
                       var_cadena = rsaux2!vcha_cli_clave_id + " " + Mid(rsaux2!VCHA_CLI_NOMBRE, 1, 50)
                       For var_j = Len(var_cadena) To 70
                            var_cadena = var_cadena + " "
                       Next var_j
                       var_cadena = var_cadena + CStr(rsaux2!inte_ped_numero)
                       If rsaux2!AUTORIZO = 1 Then
                          var_cadena = var_cadena + "        *"
                       Else
                          var_cadena = var_cadena + "      NO AUT."
                       End If
                       var_cantidad = Format(CStr(IIf(IsNull(rsaux2!Cantidad), 0, rsaux2!Cantidad)), "###,###,##0.00")
                       For var_j = Len(var_cantidad) To 14
                           var_cantidad = " " + var_cantidad
                       Next var_j
                       var_cadena = var_cadena + var_cantidad
                       Print #1, var_cadena
                       rsaux2.MoveNext
                 Wend
                 Print #1, "========================================================================================================"
                 Close #1
                 MAPIMessages1.AttachmentPathName = var_Archivo
                 MAPIMessages1.Send True
                 If MAPISession1.SessionID > 0 Then
                    MAPISession1.SignOff
                 End If
             End If
             rsaux2.Close
         End If
         rs.MoveNext
         rsaux.Close
   Wend
   rs.Close
   
   
   
   
   
   frm_correo.Visible = False
End Sub

Private Sub cmd_calendario_1_Click()
   var_mes = 1
   mes_filtro.Value = Date
   mes_filtro.Visible = True
   mes_filtro.SetFocus
End Sub

Private Sub cmd_calendario_2_Click()
   var_mes = 2
   mes_filtro.Value = Date
   mes_filtro.Visible = True
   mes_filtro.SetFocus
End Sub

Private Sub cmd_Cancelar_corro_Click()
   frm_correo.Visible = False
End Sub

Private Sub cmd_correo_Click()
   Me.txt_inicio_correo = Date
   Me.txt_fin_correo = Date
   Me.frm_correo.Visible = True
End Sub

Private Sub cmd_guardar_Click()
   Dim var_dia As String
   Dim var_mes As String
   Dim var_año As String
   txt_inicio = Date
   txt_fin = Date
   var_fecha_fin_1 = CDate(txt_fin) + 1
   var_dia = CStr(Day(CDate(txt_inicio)))
   var_mes = CStr(Month(CDate(txt_inicio)))
   var_año = CStr(Year(CDate(txt_inicio)))
   If Len(Trim(var_dia)) = 1 Then
      var_dia = "0" + var_dia
   End If
   If Len(Trim(var_mes)) = 1 Then
      var_mes = "0" + var_mes
   End If
   var_fecha_inicio_str = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            
   var_dia = CStr(Day(var_fecha_fin_1))
   var_mes = CStr(Month(var_fecha_fin_1))
   var_año = CStr(Year(var_fecha_fin_1))
   If Len(Trim(var_dia)) = 1 Then
      var_dia = "0" + var_dia
   End If
   If Len(Trim(var_mes)) = 1 Then
      var_mes = "0" + var_mes
   End If
   var_fecha_fin_str = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
 
   
   
   rs.Open "select * from vw_suma_pedidos where dtim_ped_fecha >= " + var_fecha_inicio_str + " and dtim_ped_fecha <= " + var_fecha_fin_str + " and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
      lv_pedidos.SmallIcons = ImageList1
      lv_pedidos.ListItems.Clear
      Dim list_item As ListItem
      While Not rs.EOF
         If IsNull(rs!inte_ped_autorizo) Then
            var_a = 0
         Else
            var_a = rs!inte_ped_autorizo
         End If
         If var_a <> 1 And (Len(Trim(rs!char_ped_estatus)) = 0 Or rs!char_ped_estatus = "I") Then
            Set list_item = lv_pedidos.ListItems.Add(, , rs!inte_ped_numero)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            list_item.SubItems(2) = IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE)
            list_item.SubItems(3) = IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE)
            list_item.SubItems(4) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            If IsNull(rs!Cantidad) Then
               list_item.SubItems(5) = Format(0, "###,###,##0.00")
            Else
               list_item.SubItems(5) = Format(rs!Cantidad, "###,###,##0.00")
            End If
            If IsNull(rs!Importe) Then
               list_item.SubItems(6) = Format(0, "###,###,##0.00")
            Else
               list_item.SubItems(6) = Format(rs!Importe, "###,###,##0.00")
            End If
            If IsNull(rs!inte_ped_autorizo) Then
               list_item.SubItems(7) = 0
            Else
               list_item.SubItems(7) = rs!inte_ped_autorizo
               If rs!inte_ped_autorizo = 1 Then
                  list_item.SmallIcon = 13
                  list_item.Bold = True
                  list_item.ForeColor = &H8000&
                  list_item.ListSubItems.item(1).Bold = True
                  list_item.ListSubItems.item(2).Bold = True
                  list_item.ListSubItems.item(3).Bold = True
                  list_item.ListSubItems.item(4).Bold = True
                  list_item.ListSubItems.item(5).Bold = True
                  list_item.ListSubItems.item(6).Bold = True
                  list_item.ListSubItems.item(1).ForeColor = &H8000&
                  list_item.ListSubItems.item(2).ForeColor = &H8000&
                  list_item.ListSubItems.item(3).ForeColor = &H8000&
                  list_item.ListSubItems.item(4).ForeColor = &H8000&
                  list_item.ListSubItems.item(5).ForeColor = &H8000&
                  list_item.ListSubItems.item(6).ForeColor = &H8000&
               End If
            End If
            If IsNull(rs!VCHA_PED_AUTORIZO) Then
               list_item.SubItems(8) = ""
            Else
               list_item.SubItems(8) = rs!VCHA_PED_AUTORIZO
            End If
            If IsNull(rs!DTIM_PED_AUTORIZO) Then
               list_item.SubItems(9) = ""
            Else
               list_item.SubItems(9) = rs!DTIM_PED_AUTORIZO
            End If
            If IsNull(rs!dtim_ped_fecha) Then
               list_item.SubItems(10) = ""
            Else
               list_item.SubItems(10) = rs!dtim_ped_fecha
            End If
            If IsNull(rs!char_ped_estatus) Then
               list_item.SubItems(11) = ""
            Else
               list_item.SubItems(11) = rs!char_ped_estatus
            End If
            If IsNull(rs!floa_ped_descuento_1) Then
               list_item.SubItems(12) = 0
            Else
               list_item.SubItems(12) = rs!floa_ped_descuento_1
            End If
            If IsNull(rs!FLOA_PED_DESCUENTO_2) Then
               list_item.SubItems(13) = 0
            Else
               list_item.SubItems(13) = rs!FLOA_PED_DESCUENTO_2
            End If
            If IsNull(rs!VCHA_USU_NOMBRE) Then
               list_item.SubItems(14) = ""
            Else
               If IsNull(rs!vcha_usu_apellidos) Then
                  list_item.SubItems(14) = rs!VCHA_USU_NOMBRE
               Else
                  list_item.SubItems(14) = Trim(rs!VCHA_USU_NOMBRE) + " " + Trim(rs!vcha_usu_apellidos)
               End If
            End If
         End If
         
         rsaux4.Open "select * from tb_titulares where vcha_tit_titular_id = '" + rs!vcha_tit_titular_id + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux4.EOF Then
            var_clave_grupo_real = IIf(IsNull(rsaux4!vcha_gre_grupo_real_id), "", rsaux4!vcha_gre_grupo_real_id)
         Else
            var_clave_grupo_real = ""
         End If
         rsaux4.Close
         If var_clave_grupo_real <> "" Then
            rsaux4.Open "select * from tb_gruposreales where vcha_gre_grupo_real_id = '" + var_clave_grupo_real + "' ", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux4.EOF Then
               var_clave_grupo_actual = IIf(IsNull(rsaux4!VCHA_GAC_GRUPO_aCTUAL_ID), "", rsaux4!VCHA_GAC_GRUPO_aCTUAL_ID)
            Else
               var_clave_grupo_actual = ""
            End If
            rsaux4.Close
            list_item.SubItems(17) = var_clave_grupo_actual
            If Trim(var_clave_grupo_actual) <> "" Then
               rsaux2.Open "select sum(floa_sal_importe) from VW_FACTURAS_VENCIADAS where vcha_gac_grupo_actual_id = '" + var_clave_grupo_actual + "'  and dias > 0", cnn, adOpenDynamic, adLockOptimistic
               'var_cadena = "SELECT dbo.TB_SALDOS.VCHA_EMP_EMPRESA_ID, dbo.TB_SALDOS.VCHA_SER_SERIE_ID, dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO, dbo.TB_SALDOS.VCHA_CLI_CLAVE_ID, dbo.TB_SALDOS.INTE_CAR_NUMERO, dbo.TB_SALDOS.FLOA_SAL_IMPORTE, dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA, dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_PLAZO, DATEDIFF([day], dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA + dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_PLAZO, GETDATE()) AS dias, dbo.TB_ENCABEZADO_CARTERA.VCHA_MON_MONEDA_ID, dbo.TB_MONEDAS.VCHA_MON_NOMBRE_PLURAL, dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_IMPORTE_NETO / dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_TIPO_CAMBIO AS IMPORTE_FACTURA, dbo.TB_ENCABEZADO_CARTERA.VCHA_AGE_AGENTE_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_GAC_GRUPO_ACTUAL_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_GRE_GRUPO_REAL_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_TIT_TITULAR_ID, dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_TIPO_CAMBIO FROM  dbo.TB_SALDOS INNER JOIN "
               'var_cadena = var_cadena + " dbo.TB_ENCABEZADO_CARTERA ON dbo.TB_SALDOS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_SALDOS.VCHA_SER_SERIE_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID AND dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO = dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO AND dbo.TB_SALDOS.INTE_CAR_NUMERO = dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO INNER JOIN dbo.TB_MONEDAS ON dbo.TB_SALDOS.VCHA_MON_MONEDA_ID = dbo.TB_MONEDAS.VCHA_MON_MONEDA_ID WHERE (dbo.TB_SALDOS.FLOA_SAL_IMPORTE > 0) AND (dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO = 'FA') AND (dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS <> 'C' OR dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS IS NULL and dbo.TB_ENCABEZADO_CARTERA.vcha_gac_grupo_Actual_id = '" + var_clave_grupo_actual + "' and DATEDIFF([day], dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA + dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_PLAZO, GETDATE()) > 0) "
               'rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               
               
               
               If Not rsaux2.EOF Then
                  list_item.SubItems(16) = Format(IIf(IsNull(rsaux2(0).Value), 0, rsaux2(0).Value), "###,###,##0.00")
               End If
               rsaux2.Close
             End If
         End If
         list_item.SubItems(18) = IIf(IsNull(rs!VCHA_PED_OBSERVACION), "", rs!VCHA_PED_OBSERVACION)
         rs.MoveNext:
      Wend
      rs.Close
      pro_textos
End Sub

Private Sub cmd_imprimir_Click()
    x = 1 + 1
End Sub

Private Sub cmd_nuevo_Click()
         var_filtro = False
         frm_filtro.Visible = True
         If Len(Trim(txt_fecha_inicio)) = 0 Then
            txt_fecha_inicio = Date
            var_fecha_inicio = Date
         End If
         If Len(Trim(txt_fecha_fin)) = 0 Then
            txt_fecha_fin = Date
            var_fecha_fin = Date
         End If
         'If opt_filtro(0).Value = False And opt_filtro(1).Value = False And opt_filtro(2).Value = False Then
         '   opt_filtro(0) = 1
         '   opt_filtro(0).SetFocus
         'Else
         '   If opt_filtro(0).Value = True Then
         '      opt_filtro(0).SetFocus
         '   End If
         '   If opt_filtro(1).Value = True Then
         '      opt_filtro(1).SetFocus
         '   End If
         '   If opt_filtro(2).Value = True Then
         '      opt_filtro(2).SetFocus
         '   End If
         'End If

End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   Dim var_fecha_inicio_s As String
   Dim var_fecha_fin_s As String
   Dim var_dia As String
   Dim var_mes As String
   Dim var_año As String
   
   
   var_fecha_fin_1 = CDate(txt_fecha_fin) + 1
   var_dia = CStr(Day(CDate(txt_fecha_inicio)))
   var_mes = CStr(Month(CDate(txt_fecha_inicio)))
   var_año = CStr(Year(CDate(txt_fecha_inicio)))
   If Len(Trim(var_dia)) = 1 Then
      var_dia = "0" + var_dia
   End If
   If Len(Trim(var_mes)) = 1 Then
      var_mes = "0" + var_mes
   End If
   var_fecha_inicio_s = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
           
           
   var_dia = CStr(Day(var_fecha_fin_1))
   var_mes = CStr(Month(var_fecha_fin_1))
   var_año = CStr(Year(var_fecha_fin_1))
   If Len(Trim(var_dia)) = 1 Then
      var_dia = "0" + var_dia
   End If
   If Len(Trim(var_mes)) = 1 Then
      var_mes = "0" + var_mes
   End If
   var_fecha_fin_s = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"


   Dim var_clave_grupo_real As String
   Dim var_clave_grupo_actual As String
   var_cadena_seguridad = ""
   Frmmenu2.StatusBar1.Panels(1).Text = "Autorizar con F8, cancelar Autorización con F9, cancelar pedido con F10 y reactivar con F11"
   Top = 0
   Left = 0
   frm_detalle_pedido.Visible = False
   var_fecha_inicio = Date
   var_decha_fin = Date
   mes_filtro.Visible = False
   frm_filtro.Visible = False
   frm_cheques_devueltos.Visible = False
   frm_facturas.Visible = False
   
   rs.Open "select * from vw_suma_pedidos where (inte_ped_autorizo = 0 or inte_ped_autorizo is null or inte_ped_pedido_credito = 1) AND DTIM_PED_FECHA >= " + var_fecha_inicio_s + " AND DTIM_PED_FECHA <= " + var_fecha_fin_s + "-.000001", cnn, adOpenDynamic, adLockOptimistic
   
   lv_pedidos.SmallIcons = ImageList1
   lv_pedidos.ListItems.Clear
   Dim list_item As ListItem
   While Not rs.EOF
      If IsNull(rs!inte_ped_autorizo) Then
         var_a = 0
      Else
         var_a = rs!inte_ped_autorizo
      End If
      If var_a = 0 Then
         Set list_item = lv_pedidos.ListItems.Add(, , rs!inte_ped_numero)
         var_clave_cliente = rs!vcha_cli_clave_id
         rsaux4.Open "select * from tb_titulares where vcha_tit_titular_id = '" + IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id) + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux4.EOF Then
            var_clave_grupo_real = IIf(IsNull(rsaux4!vcha_gre_grupo_real_id), "", rsaux4!vcha_gre_grupo_real_id)
         Else
            var_clave_grupo_real = ""
         End If
         rsaux4.Close
         If var_clave_grupo_real <> "" Then
            rsaux4.Open "select * from tb_gruposreales where vcha_gre_grupo_real_id = '" + var_clave_grupo_real + "' ", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux4.EOF Then
               var_clave_grupo_actual = IIf(IsNull(rsaux4!VCHA_GAC_GRUPO_aCTUAL_ID), "", rsaux4!VCHA_GAC_GRUPO_aCTUAL_ID)
            Else
               var_clave_grupo_actual = ""
            End If
            rsaux4.Close
            list_item.SubItems(17) = var_clave_grupo_actual
            If Trim(var_clave_grupo_actual) <> "" Then
               rsaux2.Open "select sum(floa_sal_importe) from VW_FACTURAS_VENCIADAS where vcha_gac_grupo_actual_id = '" + var_clave_grupo_actual + "'  and dias > 0", cnn, adOpenDynamic, adLockOptimistic
               'var_cadena = "SELECT dbo.TB_SALDOS.VCHA_EMP_EMPRESA_ID, dbo.TB_SALDOS.VCHA_SER_SERIE_ID, dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO, dbo.TB_SALDOS.VCHA_CLI_CLAVE_ID, dbo.TB_SALDOS.INTE_CAR_NUMERO, dbo.TB_SALDOS.FLOA_SAL_IMPORTE, dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA, dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_PLAZO, DATEDIFF([day], dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA + dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_PLAZO, GETDATE()) AS dias, dbo.TB_ENCABEZADO_CARTERA.VCHA_MON_MONEDA_ID, dbo.TB_MONEDAS.VCHA_MON_NOMBRE_PLURAL, dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_IMPORTE_NETO / dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_TIPO_CAMBIO AS IMPORTE_FACTURA, dbo.TB_ENCABEZADO_CARTERA.VCHA_AGE_AGENTE_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_GAC_GRUPO_ACTUAL_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_GRE_GRUPO_REAL_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_TIT_TITULAR_ID, dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_TIPO_CAMBIO FROM  dbo.TB_SALDOS INNER JOIN "
               'var_cadena = var_cadena + " dbo.TB_ENCABEZADO_CARTERA ON dbo.TB_SALDOS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_SALDOS.VCHA_SER_SERIE_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID AND dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO = dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO AND dbo.TB_SALDOS.INTE_CAR_NUMERO = dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO INNER JOIN dbo.TB_MONEDAS ON dbo.TB_SALDOS.VCHA_MON_MONEDA_ID = dbo.TB_MONEDAS.VCHA_MON_MONEDA_ID WHERE (dbo.TB_SALDOS.FLOA_SAL_IMPORTE > 0) AND (dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO = 'FA') AND (dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS <> 'C' OR dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS IS NULL and dbo.TB_ENCABEZADO_CARTERA.vcha_gac_grupo_Actual_id = '" + var_clave_grupo_actual + "' and DATEDIFF([day], dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA + dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_PLAZO, GETDATE()) > 0) "
               'rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              
               If Not rsaux2.EOF Then
                  list_item.SubItems(16) = Format(IIf(IsNull(rsaux2(0).Value), 0, rsaux2(0).Value), "###,###,##0.00")
               End If
               rsaux2.Close
             End If
         End If
         list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
         list_item.SubItems(2) = IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE)
         list_item.SubItems(3) = IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE)
         list_item.SubItems(4) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
         If IsNull(rs!Cantidad) Then
            list_item.SubItems(5) = Format(0, "###,###,##0.00")
         Else
            list_item.SubItems(5) = Format(rs!Cantidad, "###,###,##0.00")
         End If
         If IsNull(rs!Importe) Then
            list_item.SubItems(6) = Format(0, "###,###,##0.00")
         Else
            list_item.SubItems(6) = Format(rs!Importe, "###,###,##0.00")
         End If
         If IsNull(rs!inte_ped_autorizo) Then
            list_item.SubItems(7) = 0
         Else
            list_item.SubItems(7) = rs!inte_ped_autorizo
            If rs!inte_ped_autorizo = 1 Then
               list_item.SmallIcon = 13
               list_item.Bold = True
               list_item.ForeColor = &H8000&
               list_item.ListSubItems.item(1).Bold = True
               list_item.ListSubItems.item(2).Bold = True
               list_item.ListSubItems.item(3).Bold = True
               list_item.ListSubItems.item(4).Bold = True
               list_item.ListSubItems.item(5).Bold = True
               list_item.ListSubItems.item(6).Bold = True
               list_item.ListSubItems.item(1).ForeColor = &H8000&
               list_item.ListSubItems.item(2).ForeColor = &H8000&
               list_item.ListSubItems.item(3).ForeColor = &H8000&
               list_item.ListSubItems.item(4).ForeColor = &H8000&
               list_item.ListSubItems.item(5).ForeColor = &H8000&
               list_item.ListSubItems.item(6).ForeColor = &H8000&
            End If
         End If
         If IsNull(rs!VCHA_PED_AUTORIZO) Then
            list_item.SubItems(8) = ""
         Else
            list_item.SubItems(8) = rs!VCHA_PED_AUTORIZO
         End If
         If IsNull(rs!DTIM_PED_AUTORIZO) Then
            list_item.SubItems(9) = ""
         Else
            list_item.SubItems(9) = rs!DTIM_PED_AUTORIZO
         End If
         If IsNull(rs!dtim_ped_fecha) Then
            list_item.SubItems(10) = ""
         Else
            list_item.SubItems(10) = rs!dtim_ped_fecha
         End If
         If IsNull(rs!char_ped_estatus) Then
            list_item.SubItems(11) = ""
         Else
            list_item.SubItems(11) = rs!char_ped_estatus
         End If
         If IsNull(rs!floa_ped_descuento_1) Then
            list_item.SubItems(12) = 0
         Else
            list_item.SubItems(12) = rs!floa_ped_descuento_1
         End If
         If IsNull(rs!FLOA_PED_DESCUENTO_2) Then
            list_item.SubItems(13) = 0
         Else
            list_item.SubItems(13) = rs!FLOA_PED_DESCUENTO_2
         End If
         If IsNull(rs!VCHA_USU_NOMBRE) Then
            list_item.SubItems(14) = ""
         Else
            If IsNull(rs!vcha_usu_apellidos) Then
               list_item.SubItems(14) = rs!VCHA_USU_NOMBRE
            Else
               list_item.SubItems(14) = Trim(rs!VCHA_USU_NOMBRE) + " " + Trim(rs!vcha_usu_apellidos)
            End If
         End If
         list_item.SubItems(15) = Trim(rs!vcha_cli_clave_id) + "" + vcha_cli_clave_id
         list_item.SubItems(18) = IIf(IsNull(rs!VCHA_PED_OBSERVACION), "", rs!VCHA_PED_OBSERVACION)
         list_item.SubItems(19) = IIf(IsNull(rs!inte_cli_activo), "", rs!inte_cli_activo)
         If rs!inte_cli_activo = 1 Then
            list_item.Bold = True
            list_item.ForeColor = &HC000C0
            list_item.ListSubItems.item(1).Bold = True
            list_item.ListSubItems.item(2).Bold = True
            list_item.ListSubItems.item(3).Bold = True
            list_item.ListSubItems.item(4).Bold = True
            list_item.ListSubItems.item(5).Bold = True
            list_item.ListSubItems.item(6).Bold = True
            list_item.ListSubItems.item(1).ForeColor = &HC000C0
            list_item.ListSubItems.item(2).ForeColor = &HC000C0
            list_item.ListSubItems.item(3).ForeColor = &HC000C0
            list_item.ListSubItems.item(4).ForeColor = &HC000C0
            list_item.ListSubItems.item(5).ForeColor = &HC000C0
            list_item.ListSubItems.item(6).ForeColor = &HC000C0
         End If
         
      End If
      rs.MoveNext:
   Wend
   rs.Close
   pro_textos
End Sub

Private Sub Command10_Click()
   var_mes = 4
   mes_filtro.Value = Date
   mes_filtro.Visible = True
   mes_filtro.SetFocus
End Sub

Private Sub Command11_Click()
   var_mes = 3
   mes_filtro.Value = Date
   mes_filtro.Visible = True
   mes_filtro.SetFocus
End Sub

Private Sub Command2_Click()
   
   Dim var_fecha_inicio_s As String
   Dim var_fecha_fin_s As String
   Dim var_dia As String
   Dim var_mes As String
   Dim var_año As String
   
   
   var_fecha_fin_1 = CDate(txt_fecha_fin) + 1
   var_dia = CStr(Day(CDate(txt_fecha_inicio)))
   var_mes = CStr(Month(CDate(txt_fecha_inicio)))
   var_año = CStr(Year(CDate(txt_fecha_inicio)))
   If Len(Trim(var_dia)) = 1 Then
      var_dia = "0" + var_dia
   End If
   If Len(Trim(var_mes)) = 1 Then
      var_mes = "0" + var_mes
   End If
   var_fecha_inicio_s = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
           
           
   var_dia = CStr(Day(var_fecha_fin_1))
   var_mes = CStr(Month(var_fecha_fin_1))
   var_año = CStr(Year(var_fecha_fin_1))
   If Len(Trim(var_dia)) = 1 Then
      var_dia = "0" + var_dia
   End If
   If Len(Trim(var_mes)) = 1 Then
      var_mes = "0" + var_mes
   End If
   var_fecha_fin_s = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
   
   
   Dim var_clave_grupo_real As String
   Dim var_clave_grupo_actual As String
   var_cadena_seguridad = ""
   Frmmenu2.StatusBar1.Panels(1).Text = "Autorizar con F8, cancelar Autorización con F9, cancelar pedido con F10 y reactivar con F11"
   Top = 0
   Left = 0
   frm_detalle_pedido.Visible = False
   var_fecha_inicio = Date
   var_decha_fin = Date
   mes_filtro.Visible = False
   frm_filtro.Visible = False
   frm_cheques_devueltos.Visible = False
   frm_facturas.Visible = False
   rs.Open "select * from vw_suma_pedidos_2 where inte_ped_autorizo = 1 and dtim_ped_autorizo >= " + var_fecha_inicio_s + " and dtim_ped_autorizo <= " + var_fecha_fin_s + "-.0000001", cnn, adOpenDynamic, adLockOptimistic
   lv_pedidos.SmallIcons = ImageList1
   lv_pedidos.ListItems.Clear
   Dim list_item As ListItem
   While Not rs.EOF
      If IsNull(rs!inte_ped_autorizo) Then
         var_a = 0
      Else
         var_a = rs!inte_ped_autorizo
      End If
      If var_a = 1 Then
         Set list_item = lv_pedidos.ListItems.Add(, , rs!inte_ped_numero)
         var_clave_cliente = rs!vcha_cli_clave_id
         rsaux4.Open "select * from tb_titulares where vcha_tit_titular_id = '" + IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id) + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux4.EOF Then
            var_clave_grupo_real = IIf(IsNull(rsaux4!vcha_gre_grupo_real_id), "", rsaux4!vcha_gre_grupo_real_id)
         Else
            var_clave_grupo_real = ""
         End If
         rsaux4.Close
         If var_clave_grupo_real <> "" Then
            rsaux4.Open "select * from tb_gruposreales where vcha_gre_grupo_real_id = '" + var_clave_grupo_real + "' ", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux4.EOF Then
               var_clave_grupo_actual = IIf(IsNull(rsaux4!VCHA_GAC_GRUPO_aCTUAL_ID), "", rsaux4!VCHA_GAC_GRUPO_aCTUAL_ID)
            Else
               var_clave_grupo_actual = ""
            End If
            rsaux4.Close
            list_item.SubItems(17) = var_clave_grupo_actual
            If Trim(var_clave_grupo_actual) <> "" Then
               cnn.CommandTimeout = 360
               ' se inivio por peticion de victor de luna
               rsaux2.Open "select sum(floa_sal_importe) from VW_FACTURAS_VENCIADAS where vcha_gac_grupo_actual_id = '" + var_clave_grupo_actual + "'  and dias > 0", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux2.EOF Then
                  list_item.SubItems(16) = Format(IIf(IsNull(rsaux2(0).Value), 0, rsaux2(0).Value), "###,###,##0.00")
               End If
               rsaux2.Close
             End If
         End If
         list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
         list_item.SubItems(2) = IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE)
         list_item.SubItems(3) = IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE)
         list_item.SubItems(4) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
         If IsNull(rs!Cantidad) Then
            list_item.SubItems(5) = Format(0, "###,###,##0.00")
         Else
            list_item.SubItems(5) = Format(rs!Cantidad, "###,###,##0.00")
         End If
         If IsNull(rs!Importe) Then
            list_item.SubItems(6) = Format(0, "###,###,##0.00")
         Else
            list_item.SubItems(6) = Format(rs!Importe, "###,###,##0.00")
         End If
         If IsNull(rs!inte_ped_autorizo) Then
            list_item.SubItems(7) = 0
         Else
            list_item.SubItems(7) = rs!inte_ped_autorizo
            If rs!inte_ped_autorizo = 1 Then
               list_item.SmallIcon = 13
               list_item.Bold = True
               list_item.ForeColor = &H8000&
               list_item.ListSubItems.item(1).Bold = True
               list_item.ListSubItems.item(2).Bold = True
               list_item.ListSubItems.item(3).Bold = True
               list_item.ListSubItems.item(4).Bold = True
               list_item.ListSubItems.item(5).Bold = True
               list_item.ListSubItems.item(6).Bold = True
               list_item.ListSubItems.item(1).ForeColor = &H8000&
               list_item.ListSubItems.item(2).ForeColor = &H8000&
               list_item.ListSubItems.item(3).ForeColor = &H8000&
               list_item.ListSubItems.item(4).ForeColor = &H8000&
               list_item.ListSubItems.item(5).ForeColor = &H8000&
               list_item.ListSubItems.item(6).ForeColor = &H8000&
            End If
         End If
         If IsNull(rs!VCHA_PED_AUTORIZO) Then
            list_item.SubItems(8) = ""
         Else
            list_item.SubItems(8) = rs!VCHA_PED_AUTORIZO
         End If
         If IsNull(rs!DTIM_PED_AUTORIZO) Then
            list_item.SubItems(9) = ""
         Else
            list_item.SubItems(9) = rs!DTIM_PED_AUTORIZO
         End If
         If IsNull(rs!dtim_ped_fecha) Then
            list_item.SubItems(10) = ""
         Else
            list_item.SubItems(10) = rs!dtim_ped_fecha
         End If
         If IsNull(rs!char_ped_estatus) Then
            list_item.SubItems(11) = ""
         Else
            list_item.SubItems(11) = rs!char_ped_estatus
         End If
         If IsNull(rs!floa_ped_descuento_1) Then
            list_item.SubItems(12) = 0
         Else
            list_item.SubItems(12) = rs!floa_ped_descuento_1
         End If
         If IsNull(rs!FLOA_PED_DESCUENTO_2) Then
            list_item.SubItems(13) = 0
         Else
            list_item.SubItems(13) = rs!FLOA_PED_DESCUENTO_2
         End If
         If IsNull(rs!VCHA_USU_NOMBRE) Then
            list_item.SubItems(14) = ""
         Else
            If IsNull(rs!vcha_usu_apellidos) Then
               list_item.SubItems(14) = rs!VCHA_USU_NOMBRE
            Else
               list_item.SubItems(14) = Trim(rs!VCHA_USU_NOMBRE) + " " + Trim(rs!vcha_usu_apellidos)
            End If
         End If
         list_item.SubItems(15) = Trim(rs!vcha_cli_clave_id) + "" + vcha_cli_clave_id
         list_item.SubItems(18) = IIf(IsNull(rs!VCHA_PED_OBSERVACION), "", rs!VCHA_PED_OBSERVACION)
      End If
      rs.MoveNext:
   Wend
   rs.Close
   pro_textos
   var_fecha_fin_1 = CStr(Now)
   'MsgBox var_fecha_inicio_1 + " " + var_fecha_fin_1, 32, ""
End Sub

Private Sub Command3_Click()
   Dim var_fecha_inicio_s As String
   Dim var_fecha_fin_s As String
   Dim var_dia As String
   Dim var_mes As String
   Dim var_año As String
   
   
   var_fecha_fin_1 = CDate(txt_fecha_fin) + 1
   var_dia = CStr(Day(CDate(txt_fecha_inicio)))
   var_mes = CStr(Month(CDate(txt_fecha_inicio)))
   var_año = CStr(Year(CDate(txt_fecha_inicio)))
   If Len(Trim(var_dia)) = 1 Then
      var_dia = "0" + var_dia
   End If
   If Len(Trim(var_mes)) = 1 Then
      var_mes = "0" + var_mes
   End If
   var_fecha_inicio_s = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
           
           
   var_dia = CStr(Day(var_fecha_fin_1))
   var_mes = CStr(Month(var_fecha_fin_1))
   var_año = CStr(Year(var_fecha_fin_1))
   If Len(Trim(var_dia)) = 1 Then
      var_dia = "0" + var_dia
   End If
   If Len(Trim(var_mes)) = 1 Then
      var_mes = "0" + var_mes
   End If
   var_fecha_fin_s = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"

   Dim var_clave_grupo_real As String
   Dim var_clave_grupo_actual As String
   Dim var_fecha_inicio_1 As String
   var_cadena_seguridad = ""
   Frmmenu2.StatusBar1.Panels(1).Text = "Autorizar con F8, cancelar Autorización con F9, cancelar pedido con F10 y reactivar con F11"
   Top = 0
   Left = 0
   frm_detalle_pedido.Visible = False
   var_fecha_inicio = Date
   var_decha_fin = Date
   mes_filtro.Visible = False
   frm_filtro.Visible = False
   frm_cheques_devueltos.Visible = False
   frm_facturas.Visible = False
   rs.Open "select * from vw_suma_pedidos_CANCELADOS where dtim_ped_fecha >= " + var_fecha_inicio_s + " and dtim_ped_fecha <= " + var_fecha_fin_s + "-.000001", cnn, adOpenDynamic, adLockOptimistic
   lv_pedidos.SmallIcons = ImageList1
   lv_pedidos.ListItems.Clear
   Dim list_item As ListItem
   While Not rs.EOF
      If rs!char_ped_estatus = "C" Then
         Set list_item = lv_pedidos.ListItems.Add(, , rs!inte_ped_numero)
         var_clave_cliente = rs!vcha_cli_clave_id
         rsaux4.Open "select * from tb_titulares where vcha_tit_titular_id = '" + IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id) + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux4.EOF Then
            var_clave_grupo_real = IIf(IsNull(rsaux4!vcha_gre_grupo_real_id), "", rsaux4!vcha_gre_grupo_real_id)
         Else
            var_clave_grupo_real = ""
         End If
         rsaux4.Close
         If var_clave_grupo_real <> "" Then
            rsaux4.Open "select * from tb_gruposreales where vcha_gre_grupo_real_id = '" + var_clave_grupo_real + "' ", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux4.EOF Then
               var_clave_grupo_actual = IIf(IsNull(rsaux4!VCHA_GAC_GRUPO_aCTUAL_ID), "", rsaux4!VCHA_GAC_GRUPO_aCTUAL_ID)
            Else
               var_clave_grupo_actual = ""
            End If
            rsaux4.Close
            list_item.SubItems(17) = var_clave_grupo_actual
            If Trim(var_clave_grupo_actual) <> "" Then
               rsaux2.Open "select sum(floa_sal_importe) from VW_FACTURAS_VENCIADAS where vcha_gac_grupo_actual_id = '" + var_clave_grupo_actual + "'  and dias > 0", cnn, adOpenDynamic, adLockOptimistic
               'var_cadena = "SELECT dbo.TB_SALDOS.VCHA_EMP_EMPRESA_ID, dbo.TB_SALDOS.VCHA_SER_SERIE_ID, dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO, dbo.TB_SALDOS.VCHA_CLI_CLAVE_ID, dbo.TB_SALDOS.INTE_CAR_NUMERO, dbo.TB_SALDOS.FLOA_SAL_IMPORTE, dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA, dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_PLAZO, DATEDIFF([day], dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA + dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_PLAZO, GETDATE()) AS dias, dbo.TB_ENCABEZADO_CARTERA.VCHA_MON_MONEDA_ID, dbo.TB_MONEDAS.VCHA_MON_NOMBRE_PLURAL, dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_IMPORTE_NETO / dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_TIPO_CAMBIO AS IMPORTE_FACTURA, dbo.TB_ENCABEZADO_CARTERA.VCHA_AGE_AGENTE_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_GAC_GRUPO_ACTUAL_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_GRE_GRUPO_REAL_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_TIT_TITULAR_ID, dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_TIPO_CAMBIO FROM  dbo.TB_SALDOS INNER JOIN "
               'var_cadena = var_cadena + " dbo.TB_ENCABEZADO_CARTERA ON dbo.TB_SALDOS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_SALDOS.VCHA_SER_SERIE_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID AND dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO = dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO AND dbo.TB_SALDOS.INTE_CAR_NUMERO = dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO INNER JOIN dbo.TB_MONEDAS ON dbo.TB_SALDOS.VCHA_MON_MONEDA_ID = dbo.TB_MONEDAS.VCHA_MON_MONEDA_ID WHERE (dbo.TB_SALDOS.FLOA_SAL_IMPORTE > 0) AND (dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO = 'FA') AND (dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS <> 'C' OR dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS IS NULL and dbo.TB_ENCABEZADO_CARTERA.vcha_gac_grupo_Actual_id = '" + var_clave_grupo_actual + "' and DATEDIFF([day], dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA + dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_PLAZO, GETDATE()) > 0) "
               'rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               
               
               If Not rsaux2.EOF Then
                  list_item.SubItems(16) = Format(IIf(IsNull(rsaux2(0).Value), 0, rsaux2(0).Value), "###,###,##0.00")
               End If
               rsaux2.Close
             End If
         End If
         list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
         list_item.SubItems(2) = IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE)
         list_item.SubItems(3) = IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE)
         list_item.SubItems(4) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
         If IsNull(rs!Cantidad) Then
            list_item.SubItems(5) = Format(0, "###,###,##0.00")
         Else
            list_item.SubItems(5) = Format(rs!Cantidad, "###,###,##0.00")
         End If
         If IsNull(rs!Importe) Then
            list_item.SubItems(6) = Format(0, "###,###,##0.00")
         Else
            list_item.SubItems(6) = Format(rs!Importe, "###,###,##0.00")
         End If
         If IsNull(rs!inte_ped_autorizo) Then
            list_item.SubItems(7) = 0
         Else
            list_item.SubItems(7) = rs!inte_ped_autorizo
            If rs!inte_ped_autorizo = 1 Then
               list_item.SmallIcon = 13
               list_item.Bold = True
               list_item.ForeColor = &H8000&
               list_item.ListSubItems.item(1).Bold = True
               list_item.ListSubItems.item(2).Bold = True
               list_item.ListSubItems.item(3).Bold = True
               list_item.ListSubItems.item(4).Bold = True
               list_item.ListSubItems.item(5).Bold = True
               list_item.ListSubItems.item(6).Bold = True
               list_item.ListSubItems.item(1).ForeColor = &H8000&
               list_item.ListSubItems.item(2).ForeColor = &H8000&
               list_item.ListSubItems.item(3).ForeColor = &H8000&
               list_item.ListSubItems.item(4).ForeColor = &H8000&
               list_item.ListSubItems.item(5).ForeColor = &H8000&
               list_item.ListSubItems.item(6).ForeColor = &H8000&
            End If
         End If
         If IsNull(rs!VCHA_PED_AUTORIZO) Then
            list_item.SubItems(8) = ""
         Else
            list_item.SubItems(8) = rs!VCHA_PED_AUTORIZO
         End If
         If IsNull(rs!DTIM_PED_AUTORIZO) Then
            list_item.SubItems(9) = ""
         Else
            list_item.SubItems(9) = rs!DTIM_PED_AUTORIZO
         End If
         If IsNull(rs!dtim_ped_fecha) Then
            list_item.SubItems(10) = ""
         Else
            list_item.SubItems(10) = rs!dtim_ped_fecha
         End If
         If IsNull(rs!char_ped_estatus) Then
            list_item.SubItems(11) = ""
         Else
            list_item.SubItems(11) = rs!char_ped_estatus
         End If
         If IsNull(rs!floa_ped_descuento_1) Then
            list_item.SubItems(12) = 0
         Else
            list_item.SubItems(12) = rs!floa_ped_descuento_1
         End If
         If IsNull(rs!FLOA_PED_DESCUENTO_2) Then
            list_item.SubItems(13) = 0
         Else
            list_item.SubItems(13) = rs!FLOA_PED_DESCUENTO_2
         End If
         If IsNull(rs!VCHA_USU_NOMBRE) Then
            list_item.SubItems(14) = ""
         Else
            If IsNull(rs!vcha_usu_apellidos) Then
               list_item.SubItems(14) = rs!VCHA_USU_NOMBRE
            Else
               list_item.SubItems(14) = Trim(rs!VCHA_USU_NOMBRE) + " " + Trim(rs!vcha_usu_apellidos)
            End If
         End If
         list_item.SubItems(15) = Trim(rs!vcha_cli_clave_id) + "" + vcha_cli_clave_id
         list_item.SubItems(18) = IIf(IsNull(rs!VCHA_PED_OBSERVACION), "", rs!VCHA_PED_OBSERVACION)
      End If
      rs.MoveNext:
   Wend
   rs.Close
   pro_textos
   var_fecha_fin_1 = CStr(Now)
   'MsgBox var_fecha_inicio_1 + " " + var_fecha_fin_1, 32, ""
End Sub

Private Sub Command4_Click()
   Me.frm_filtro.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   
   If Shift = 4 And KeyCode = 70 Then
      cmd_nuevo_Click
   End If
   If Shift = 4 And KeyCode = 65 Then
      cmd_guardar_Click
   End If
   If Shift = 4 And KeyCode = 73 Then
      cmd_imprimir_Click
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   Dim var_dia As String
   Dim var_mes As String
   Dim var_año As String
   Dim var_clave_grupo_real As String
   Dim var_clave_grupo_actual As String
   Dim var_fecha_inicio_1 As String
   Dim var_fecha_fin_1 As String
   var_fecha_inicio_1 = CStr(Now)
   var_cadena_seguridad = ""
   Frmmenu2.StatusBar1.Panels(1).Text = "Autorizar con F8, cancelar Autorización con F9, cancelar pedido con F10 y reactivar con F11"
   Top = 0
   Left = 0
   frm_detalle_pedido.Visible = False
   var_fecha_inicio = Date
   var_decha_fin = Date
   mes_filtro.Visible = False
   frm_filtro.Visible = False
   frm_cheques_devueltos.Visible = False
   frm_facturas.Visible = False
   txt_inicio = Date
   txt_fin = Date
   var_fecha_fin_1 = CDate(txt_fin) + 1
   var_dia = CStr(Day(CDate(txt_inicio)))
   var_mes = CStr(Month(CDate(txt_inicio)))
   var_año = CStr(Year(CDate(txt_inicio)))
   If Len(Trim(var_dia)) = 1 Then
      var_dia = "0" + var_dia
   End If
   If Len(Trim(var_mes)) = 1 Then
      var_mes = "0" + var_mes
   End If
   var_fecha_inicio_str = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            
   var_dia = CStr(Day(var_fecha_fin_1))
   var_mes = CStr(Month(var_fecha_fin_1))
   var_año = CStr(Year(var_fecha_fin_1))
   If Len(Trim(var_dia)) = 1 Then
      var_dia = "0" + var_dia
   End If
   If Len(Trim(var_mes)) = 1 Then
      var_mes = "0" + var_mes
   End If
   var_fecha_fin_str = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
 
   cnn.CommandTimeout = 360
   Text2 = "select * from vw_suma_pedidos where dtim_ped_fecha >= " + var_fecha_inicio_str + " and dtim_ped_fecha <= " + var_fecha_fin_str + " and vcha_emp_empresa_id = '" + var_empresa + "'"
   rs.Open "select * from vw_suma_pedidos where dtim_ped_fecha >= " + var_fecha_inicio_str + " and dtim_ped_fecha <= " + var_fecha_fin_str + " and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
   
   
   
   'rs.Open "select * from pedido_187841", cnn, adOpenDynamic, adLockOptimistic

   lv_pedidos.SmallIcons = ImageList1
   lv_pedidos.ListItems.Clear
   Dim list_item As ListItem
   While Not rs.EOF
      If IsNull(rs!inte_ped_autorizo) Then
         var_a = 0
      Else
         var_a = rs!inte_ped_autorizo
      End If
      
      If var_a <> 1 And (Len(Trim(rs!char_ped_estatus)) = 0 Or rs!char_ped_estatus = "I") Or (rs!inte_ped_pedido_credito = 1) Then
         Set list_item = lv_pedidos.ListItems.Add(, , rs!inte_ped_numero)
         var_clave_cliente = rs!vcha_cli_clave_id
         rsaux4.Open "select * from tb_titulares where vcha_tit_titular_id = '" + IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id) + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux4.EOF Then
            var_clave_grupo_real = IIf(IsNull(rsaux4!vcha_gre_grupo_real_id), "", rsaux4!vcha_gre_grupo_real_id)
         Else
            var_clave_grupo_real = ""
         End If
         rsaux4.Close
         If var_clave_grupo_real <> "" Then
            rsaux4.Open "select * from tb_gruposreales where vcha_gre_grupo_real_id = '" + var_clave_grupo_real + "' ", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux4.EOF Then
               var_clave_grupo_actual = IIf(IsNull(rsaux4!VCHA_GAC_GRUPO_aCTUAL_ID), "", rsaux4!VCHA_GAC_GRUPO_aCTUAL_ID)
            Else
               var_clave_grupo_actual = ""
            End If
            rsaux4.Close
            list_item.SubItems(17) = var_clave_grupo_actual
            If Trim(var_clave_grupo_actual) <> "" Then
               rsaux2.Open "select sum(floa_sal_importe) from VW_FACTURAS_VENCIADAS where vcha_gac_grupo_actual_id = '" + var_clave_grupo_actual + "'  and dias > 0", cnn, adOpenDynamic, adLockOptimistic
               'var_cadena = "SELECT dbo.TB_SALDOS.VCHA_EMP_EMPRESA_ID, dbo.TB_SALDOS.VCHA_SER_SERIE_ID, dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO, dbo.TB_SALDOS.VCHA_CLI_CLAVE_ID, dbo.TB_SALDOS.INTE_CAR_NUMERO, dbo.TB_SALDOS.FLOA_SAL_IMPORTE, dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA, dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_PLAZO, DATEDIFF([day], dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA + dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_PLAZO, GETDATE()) AS dias, dbo.TB_ENCABEZADO_CARTERA.VCHA_MON_MONEDA_ID, dbo.TB_MONEDAS.VCHA_MON_NOMBRE_PLURAL, dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_IMPORTE_NETO / dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_TIPO_CAMBIO AS IMPORTE_FACTURA, dbo.TB_ENCABEZADO_CARTERA.VCHA_AGE_AGENTE_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_GAC_GRUPO_ACTUAL_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_GRE_GRUPO_REAL_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_TIT_TITULAR_ID, dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_TIPO_CAMBIO FROM  dbo.TB_SALDOS INNER JOIN "
               'var_cadena = var_cadena + " dbo.TB_ENCABEZADO_CARTERA ON dbo.TB_SALDOS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_SALDOS.VCHA_SER_SERIE_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID AND dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO = dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO AND dbo.TB_SALDOS.INTE_CAR_NUMERO = dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO INNER JOIN dbo.TB_MONEDAS ON dbo.TB_SALDOS.VCHA_MON_MONEDA_ID = dbo.TB_MONEDAS.VCHA_MON_MONEDA_ID WHERE (dbo.TB_SALDOS.FLOA_SAL_IMPORTE > 0) AND (dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO = 'FA') AND (dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS <> 'C' OR dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS IS NULL and dbo.TB_ENCABEZADO_CARTERA.vcha_gac_grupo_Actual_id = '" + var_clave_grupo_actual + "' and DATEDIFF([day], dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA + dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_PLAZO, GETDATE()) > 0) "
               'rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               
               
               If Not rsaux2.EOF Then
                  list_item.SubItems(16) = Format(IIf(IsNull(rsaux2(0).Value), 0, rsaux2(0).Value), "###,###,##0.00")
               End If
               rsaux2.Close
             End If
         End If
         list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
         list_item.SubItems(2) = IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE)
         list_item.SubItems(3) = IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE)
         list_item.SubItems(4) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
         If IsNull(rs!Cantidad) Then
            list_item.SubItems(5) = Format(0, "###,###,##0.00")
         Else
            list_item.SubItems(5) = Format(rs!Cantidad, "###,###,##0.00")
         End If
         If IsNull(rs!Importe) Then
            list_item.SubItems(6) = Format(0, "###,###,##0.00")
         Else
            list_item.SubItems(6) = Format(rs!Importe, "###,###,##0.00")
         End If
         If IsNull(rs!inte_ped_autorizo) Then
            list_item.SubItems(7) = 0
         Else
            list_item.SubItems(7) = rs!inte_ped_autorizo
            If rs!inte_ped_autorizo = 1 Then
               list_item.SmallIcon = 13
               list_item.Bold = True
               list_item.ForeColor = &H8000&
               list_item.ListSubItems.item(1).Bold = True
               list_item.ListSubItems.item(2).Bold = True
               list_item.ListSubItems.item(3).Bold = True
               list_item.ListSubItems.item(4).Bold = True
               list_item.ListSubItems.item(5).Bold = True
               list_item.ListSubItems.item(6).Bold = True
               list_item.ListSubItems.item(1).ForeColor = &H8000&
               list_item.ListSubItems.item(2).ForeColor = &H8000&
               list_item.ListSubItems.item(3).ForeColor = &H8000&
               list_item.ListSubItems.item(4).ForeColor = &H8000&
               list_item.ListSubItems.item(5).ForeColor = &H8000&
               list_item.ListSubItems.item(6).ForeColor = &H8000&
            End If
         End If
         If IsNull(rs!VCHA_PED_AUTORIZO) Then
            list_item.SubItems(8) = ""
         Else
            list_item.SubItems(8) = rs!VCHA_PED_AUTORIZO
         End If
         If IsNull(rs!DTIM_PED_AUTORIZO) Then
            list_item.SubItems(9) = ""
         Else
            list_item.SubItems(9) = rs!DTIM_PED_AUTORIZO
         End If
         If IsNull(rs!dtim_ped_fecha) Then
            list_item.SubItems(10) = ""
         Else
            list_item.SubItems(10) = rs!dtim_ped_fecha
         End If
         If IsNull(rs!char_ped_estatus) Then
            list_item.SubItems(11) = ""
         Else
            list_item.SubItems(11) = rs!char_ped_estatus
         End If
         If IsNull(rs!floa_ped_descuento_1) Then
            list_item.SubItems(12) = 0
         Else
            list_item.SubItems(12) = rs!floa_ped_descuento_1
         End If
         If IsNull(rs!FLOA_PED_DESCUENTO_2) Then
            list_item.SubItems(13) = 0
         Else
            list_item.SubItems(13) = rs!FLOA_PED_DESCUENTO_2
         End If
         If IsNull(rs!VCHA_USU_NOMBRE) Then
            list_item.SubItems(14) = ""
         Else
            If IsNull(rs!vcha_usu_apellidos) Then
               list_item.SubItems(14) = rs!VCHA_USU_NOMBRE
            Else
               list_item.SubItems(14) = Trim(rs!VCHA_USU_NOMBRE) + " " + Trim(rs!vcha_usu_apellidos)
            End If
         End If
         list_item.SubItems(15) = Trim(rs!vcha_cli_clave_id) + "" + vcha_cli_clave_id
         list_item.SubItems(18) = IIf(IsNull(rs!VCHA_PED_OBSERVACION), "", rs!VCHA_PED_OBSERVACION)
         list_item.SubItems(19) = IIf(IsNull(rs!inte_cli_activo), "", rs!inte_cli_activo)
         If rs!inte_cli_activo = 1 Then
            list_item.Bold = True
            list_item.ForeColor = &HC000C0
            list_item.ListSubItems.item(1).Bold = True
            list_item.ListSubItems.item(2).Bold = True
            list_item.ListSubItems.item(3).Bold = True
            list_item.ListSubItems.item(4).Bold = True
            list_item.ListSubItems.item(5).Bold = True
            list_item.ListSubItems.item(6).Bold = True
            list_item.ListSubItems.item(1).ForeColor = &HC000C0
            list_item.ListSubItems.item(2).ForeColor = &HC000C0
            list_item.ListSubItems.item(3).ForeColor = &HC000C0
            list_item.ListSubItems.item(4).ForeColor = &HC000C0
            list_item.ListSubItems.item(5).ForeColor = &HC000C0
            list_item.ListSubItems.item(6).ForeColor = &HC000C0
         End If
      
      End If
      rs.MoveNext:
   Wend
   rs.Close
   pro_textos
   var_fecha_fin_1 = CStr(Now)
   'MsgBox var_fecha_inicio_1 + " " + var_fecha_fin_1, 32, ""
   Me.frm_correo.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Frmmenu2.StatusBar1.Panels(1).Text = ""
   Call activa_forma(var_activa_forma_autorizapedidos)
End Sub

Private Sub lv_cheques_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_cheques, ColumnHeader)
End Sub

Private Sub lv_cheques_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      frm_cheques_devueltos.Visible = False
   End If
End Sub

Private Sub lv_cheques_LostFocus()
   frm_cheques_devueltos.Visible = False
End Sub

Private Sub lv_detalle_pedido_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_detalle_pedido, ColumnHeader)
End Sub

Private Sub lv_detalle_pedido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      frm_detalle_pedido.Visible = False
   End If
End Sub

Private Sub lv_facturas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      frm_facturas.Visible = False
   End If
End Sub

Private Sub lv_facturas_LostFocus()
   frm_facturas.Visible = False
End Sub

Private Sub lv_pedidos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_pedidos, ColumnHeader)
End Sub

Private Sub lv_pedidos_ItemClick(ByVal item As MSComctlLib.ListItem)
    Call pro_textos
    'Me.txt_importe_vencido = ""
End Sub

Private Sub lv_pedidos_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 115 Then
      Dim list_item As ListItem
      frm_detalle_pedido.Visible = True
      lv_detalle_pedido.ListItems.Clear
      rs.Open "select * from vw_pedidos where inte_ped_numero = " + lv_pedidos.selectedItem, cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
         Set list_item = lv_detalle_pedido.ListItems.Add(, , rs!VCHA_ART_ARTICULO_ID)
         list_item.SubItems(1) = IIf(IsNull(rs!vcha_Art_nombre_español), "", rs!vcha_Art_nombre_español)
         list_item.SubItems(2) = IIf(IsNull(rs!FLOA_PED_CANTIDAD), 0, rs!FLOA_PED_CANTIDAD)
         lv_detalle_pedido.SetFocus
         rs.MoveNext
      Wend
      rs.Close
   End If
   If KeyCode = 119 Then
      Set TB_AUTORIZA_PEDIDOS = New TB_AUTORIZA_PEDIDOS
      ok = TB_AUTORIZA_PEDIDOS.Anadir(lv_pedidos.selectedItem, 1, var_clave_usuario_global, Now)
      rsaux8.Open "SELECT INTE_ORS_ORDEN_SURTIDO FROM TB_ENC_ORDEN_SURTIDO WHERE INTE_PED_NUMERO = " + lv_pedidos.selectedItem, cnn, adOpenDynamic, adLockOptimistic
      While Not rsaux8.EOF
            rsaux10.Open "select isnull(inte_ped_pedido_credito,0) from tb_encabezado_pedidos where INTE_PED_NUMERO = " + lv_pedidos.selectedItem, cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux10.EOF Then
               var_pedido_credito = IIf(IsNull(rsaux10(0).Value), 0, rsaux10(0).Value)
            End If
            rsaux10.Close
            If var_pedido_credito = 1 Then
               rsaux9.Open "UPDATE TB_ENC_ORDEN_SURTIDO SET DTIM_ORS_FECHA_LIBERACION =  GETDATE() WHERE INTE_ORS_ORDEN_SURTIDO = " + CStr(rsaux8(0).Value), cnn, adOpenDynamic, adLockOptimistic
            End If
            rsaux8.MoveNext
      Wend
      rsaux8.Close
      If ok Then
         lv_pedidos.selectedItem.SubItems(7) = 1
         lv_pedidos.selectedItem.SmallIcon = 13
         lv_pedidos.selectedItem.Bold = True
         lv_pedidos.selectedItem.ForeColor = &H8000&
         lv_pedidos.selectedItem.ListSubItems.item(1).Bold = True
         lv_pedidos.selectedItem.ListSubItems.item(2).Bold = True
         lv_pedidos.selectedItem.ListSubItems.item(3).Bold = True
         lv_pedidos.selectedItem.ListSubItems.item(4).Bold = True
         lv_pedidos.selectedItem.ListSubItems.item(5).Bold = True
         lv_pedidos.selectedItem.ListSubItems.item(6).Bold = True
         lv_pedidos.selectedItem.ListSubItems.item(1).ForeColor = &H8000&
         lv_pedidos.selectedItem.ListSubItems.item(2).ForeColor = &H8000&
         lv_pedidos.selectedItem.ListSubItems.item(3).ForeColor = &H8000&
         lv_pedidos.selectedItem.ListSubItems.item(4).ForeColor = &H8000&
         lv_pedidos.selectedItem.ListSubItems.item(5).ForeColor = &H8000&
         lv_pedidos.selectedItem.ListSubItems.item(6).ForeColor = &H8000&
         'MsgBox "El pedido a sido autorizado", vbOKOnly, "ATENCION"
         pro_textos
         lv_pedidos.selectedItem.SubItems(14) = var_nombre_usuario_global + " " + var_apellidos_usuario_global
         lv_pedidos.SetFocus
      End If
   End If
   If KeyCode = 120 Then
      Set TB_AUTORIZA_PEDIDOS = New TB_AUTORIZA_PEDIDOS
      ok = TB_AUTORIZA_PEDIDOS.Anadir(lv_pedidos.selectedItem, 0, var_clave_usuario_global, Now)
      If ok Then
         lv_pedidos.selectedItem.SubItems(7) = 0
         lv_pedidos.selectedItem.SmallIcon = 0
         lv_pedidos.selectedItem.SmallIcon = 0
         lv_pedidos.selectedItem.Bold = False
         lv_pedidos.selectedItem.ForeColor = &H0&
         lv_pedidos.selectedItem.ListSubItems.item(1).Bold = False
         lv_pedidos.selectedItem.ListSubItems.item(2).Bold = False
         lv_pedidos.selectedItem.ListSubItems.item(3).Bold = False
         lv_pedidos.selectedItem.ListSubItems.item(4).Bold = False
         lv_pedidos.selectedItem.ListSubItems.item(5).Bold = False
         lv_pedidos.selectedItem.ListSubItems.item(6).Bold = False
         lv_pedidos.selectedItem.ListSubItems.item(1).ForeColor = &H0&
         lv_pedidos.selectedItem.ListSubItems.item(2).ForeColor = &H0&
         lv_pedidos.selectedItem.ListSubItems.item(3).ForeColor = &H0&
         lv_pedidos.selectedItem.ListSubItems.item(4).ForeColor = &H0&
         lv_pedidos.selectedItem.ListSubItems.item(5).ForeColor = &H0&
         lv_pedidos.selectedItem.ListSubItems.item(6).ForeColor = &H0&
         pro_textos
         lv_pedidos.selectedItem.SubItems(14) = var_nombre_usuario_global + " " + var_apellidos_usuario_global
         MsgBox "Se a cancelado la autorización", vbOKOnly, "ATENCION"
         lv_pedidos.SetFocus
      End If
   End If
   If KeyCode = 121 Then
      Set TB_CANCELA_PEDIDOS = New TB_CANCELA_PEDIDOS
      ok = TB_CANCELA_PEDIDOS.Anadir(lv_pedidos.selectedItem, "C")
      If ok Then
         lv_pedidos.selectedItem.SubItems(7) = 1
         lv_pedidos.selectedItem.SmallIcon = 15
         lv_pedidos.selectedItem.Bold = True
         lv_pedidos.selectedItem.ForeColor = &HFF&
         lv_pedidos.selectedItem.ListSubItems.item(1).Bold = True
         lv_pedidos.selectedItem.ListSubItems.item(2).Bold = True
         lv_pedidos.selectedItem.ListSubItems.item(3).Bold = True
         lv_pedidos.selectedItem.ListSubItems.item(4).Bold = True
         lv_pedidos.selectedItem.ListSubItems.item(5).Bold = True
         lv_pedidos.selectedItem.ListSubItems.item(6).Bold = True
         lv_pedidos.selectedItem.ListSubItems.item(1).ForeColor = &HFF&
         lv_pedidos.selectedItem.ListSubItems.item(2).ForeColor = &HFF&
         lv_pedidos.selectedItem.ListSubItems.item(3).ForeColor = &HFF&
         lv_pedidos.selectedItem.ListSubItems.item(4).ForeColor = &HFF&
         lv_pedidos.selectedItem.ListSubItems.item(5).ForeColor = &HFF&
         lv_pedidos.selectedItem.ListSubItems.item(6).ForeColor = &HFF&
         'MsgBox "El pedido a sido cancelado", vbOKOnly, "ATENCION"
         lv_pedidos.SetFocus
      End If
   End If
   If KeyCode = 122 Then
      Set TB_CANCELA_PEDIDOS = New TB_CANCELA_PEDIDOS
      ok = TB_CANCELA_PEDIDOS.Anadir(lv_pedidos.selectedItem, " ")
      If ok Then
         lv_pedidos.selectedItem.SubItems(7) = 0
         lv_pedidos.selectedItem.SmallIcon = 0
         lv_pedidos.selectedItem.SmallIcon = 0
         lv_pedidos.selectedItem.Bold = False
         lv_pedidos.selectedItem.ForeColor = &H0&
         lv_pedidos.selectedItem.ListSubItems.item(1).Bold = False
         lv_pedidos.selectedItem.ListSubItems.item(2).Bold = False
         lv_pedidos.selectedItem.ListSubItems.item(3).Bold = False
         lv_pedidos.selectedItem.ListSubItems.item(4).Bold = False
         lv_pedidos.selectedItem.ListSubItems.item(5).Bold = False
         lv_pedidos.selectedItem.ListSubItems.item(6).Bold = False
         lv_pedidos.selectedItem.ListSubItems.item(1).ForeColor = &H0&
         lv_pedidos.selectedItem.ListSubItems.item(2).ForeColor = &H0&
         lv_pedidos.selectedItem.ListSubItems.item(3).ForeColor = &H0&
         lv_pedidos.selectedItem.ListSubItems.item(4).ForeColor = &H0&
         lv_pedidos.selectedItem.ListSubItems.item(5).ForeColor = &H0&
         lv_pedidos.selectedItem.ListSubItems.item(6).ForeColor = &H0&
         MsgBox "Se a reactivado el pedido", vbOKOnly, "ATENCION"
         lv_pedidos.SetFocus
      End If
   End If
   If KeyCode = 116 Then
      If Me.lv_pedidos.ListItems.Count > 0 Then
         var_clave_cliente = lv_pedidos.selectedItem.SubItems(15)
         var_clave_grupo_actual = lv_pedidos.selectedItem.SubItems(17)
         rsaux2.Open "select sum(floa_sal_importe) from VW_FACTURAS_VENCIADAS where vcha_gac_grupo_actual_id = '" + var_clave_grupo_actual + "'  and dias > 0", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux2.EOF Then
            txt_importe_vencido = Format(IIf(IsNull(rsaux2(0).Value), 0, rsaux2(0).Value), "###,###,##0.00")
            Me.lv_pedidos.selectedItem.SubItems(16) = Format(IIf(IsNull(rsaux2(0).Value), 0, rsaux2(0).Value), "###,###,##0.00")
         End If
         rsaux2.Close
      End If
   End If
  lv_pedidos.SetFocus
End Sub

Private Sub lv_pedidos_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'   If Button = 2 Then
'      If lv_pedidos.selectedItem <> "" Then
'         PopupMenu mnu_menu
'      End If
'   End If
End Sub

Private Sub mes_filtro_DateDblClick(ByVal DateDblClicked As Date)
   If var_mes = 1 Then
      txt_fecha_inicio = mes_filtro.Value
      var_fecha_inicio = mes_filtro.Value
      txt_fecha_inicio.SetFocus
      mes_filtro.Visible = False
   End If
   If var_mes = 2 Then
      txt_fecha_fin = mes_filtro.Value
      var_fecha_fin = mes_filtro.Value
      txt_fecha_fin.SetFocus
      mes_filtro.Visible = False
   End If
   If var_mes = 3 Then
      Me.txt_inicio_correo = mes_filtro.Value
      Me.txt_inicio_correo.SetFocus
      mes_filtro.Visible = False
   End If
   If var_mes = 4 Then
      Me.txt_fin_correo = mes_filtro.Value
      Me.txt_fin_correo.SetFocus
      mes_filtro.Visible = False
   End If
   
End Sub



Private Sub mes_filtro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      If var_mes = 1 Then
         Me.txt_fecha_inicio.SetFocus
      Else
         Me.txt_fecha_fin.SetFocus
      End If
   End If
End Sub

Private Sub mes_filtro_LostFocus()

   Me.mes_filtro.Visible = False
   If var_mes = 1 Then
      Me.txt_fecha_inicio.SetFocus
   End If
   If var_mes = 2 Then
      Me.txt_fecha_fin.SetFocus
   End If
   If var_mes = 3 Then
      Me.txt_inicio_correo.SetFocus
   End If
   If var_mes = 4 Then
      Me.txt_fin_correo.SetFocus
   End If
End Sub

Private Sub opt_filtro_Click(Index As Integer)
   Dim list_item As ListItem
   If Index = 0 Then
      var_cadena = "SELECT     dbo.TB_DETALLE_PEDIDOS.INTE_PED_NUMERO, SUM(((dbo.TB_DETALLE_PEDIDOS.FLOA_PED_PRECIO * ((100 - dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_1) / 100)) * (100 - dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_2) / 100) * (1 + dbo.TB_TIPOPEDIDOS.FLOA_TPE_IVA / 100) * dbo.TB_DETALLE_PEDIDOS.FLOA_PED_CANTIDAD) AS importe, SUM(dbo.TB_DETALLE_PEDIDOS.FLOA_PED_CANTIDAD) AS cantidad, dbo.TB_ENCABEZADO_PEDIDOS.CHAR_TPE_TIPO_PEDIDO_ID, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_REFERENCIA, dbo.TB_ENCABEZADO_PEDIDOS.DTIM_PED_REFERENCIA, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_CLI_CLAVE_ID,"
      var_cadena = var_cadena + " dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_ESPECIALES, dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_ENCABEZADO_PEDIDOS.DTIM_PED_FECHA, dbo.TB_TITULARES.VCHA_TIT_NOMBRE, dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_NOMBRE, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_AUTORIZO, dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_1, dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_2, dbo.TB_ENCABEZADO_PEDIDOS.CHAR_PED_ESTATUS, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_PED_AUTORIZO, dbo.TB_ENCABEZADO_PEDIDOS.DTIM_PED_AUTORIZO, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_RESURTIBLE, dbo.TB_USUARIOS.VCHA_USU_NOMBRE, dbo.TB_USUARIOS.VCHA_USU_APELLIDOS, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_TIT_TITULAR_ID, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_ESB_ESTABLECIMIENTO_ID, dbo.TB_CLIENTES.CHAR_PRI_PRIORIDAD_ID, dbo.TB_PRIORIDADES.FLOA_PRI_NUMERO_ORDEN, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_DIAS_CONDICIONES,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_DIAS_CADUCIDAD, dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_3, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_ALM_ALMACEN_ID, dbo.TB_TIPOPEDIDOS.FLOA_TPE_IVA, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_MON_MONEDA_ID, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_FACTURA_CEROS, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_SUGERIDO, dbo.TB_ENCABEZADO_PEDIDOS.CHAR_PED_TIPO, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_PED_CLIENTE_REFERENCIA FROM         dbo.TB_TITULARES RIGHT OUTER JOIN dbo.TB_ENCABEZADO_PEDIDOS INNER JOIN dbo.TB_DETALLE_PEDIDOS ON dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_NUMERO = dbo.TB_DETALLE_PEDIDOS.INTE_PED_NUMERO INNER JOIN dbo.TB_TIPOPEDIDOS ON dbo.TB_ENCABEZADO_PEDIDOS.CHAR_TPE_TIPO_PEDIDO_ID = dbo.TB_TIPOPEDIDOS.CHAR_TPE_TIPO_PEDIDO_ID LEFT OUTER JOIN dbo.TB_USUARIOS ON dbo.TB_ENCABEZADO_PEDIDOS.VCHA_PED_AUTORIZO = dbo.TB_USUARIOS.VCHA_USU_USUARIO_ID LEFT OUTER JOIN"
      var_cadena = var_cadena + " dbo.TB_ESTABLECIMIENTOS ON"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.VCHA_ESB_ESTABLECIMIENTO_ID = dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_ESTABLECIMIENTO_ID ON"
      var_cadena = var_cadena + " dbo.TB_TITULARES.VCHA_TIT_TITULAR_ID = dbo.TB_ENCABEZADO_PEDIDOS.VCHA_TIT_TITULAR_ID LEFT OUTER JOIN"
      var_cadena = var_cadena + " dbo.TB_CLIENTES LEFT OUTER JOIN"
      var_cadena = var_cadena + " dbo.TB_PRIORIDADES ON dbo.TB_CLIENTES.CHAR_PRI_PRIORIDAD_ID = dbo.TB_PRIORIDADES.CHAR_PRI_PRIORIDAD_ID ON"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID LEFT OUTER JOIN"
      var_cadena = var_cadena + " dbo.TB_AGENTES ON dbo.TB_CLIENTES.VCHA_AGE_AGENTE_ID = dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID"
      var_cadena = var_cadena + " GROUP BY dbo.TB_DETALLE_PEDIDOS.INTE_PED_NUMERO, dbo.TB_ENCABEZADO_PEDIDOS.CHAR_TPE_TIPO_PEDIDO_ID,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_REFERENCIA, dbo.TB_ENCABEZADO_PEDIDOS.DTIM_PED_REFERENCIA,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.VCHA_CLI_CLAVE_ID, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_ESPECIALES, dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.DTIM_PED_FECHA, dbo.TB_TITULARES.VCHA_TIT_NOMBRE, dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_NOMBRE,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_AUTORIZO, dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_1,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_2, dbo.TB_ENCABEZADO_PEDIDOS.CHAR_PED_ESTATUS,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.VCHA_PED_AUTORIZO, dbo.TB_ENCABEZADO_PEDIDOS.DTIM_PED_AUTORIZO,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_RESURTIBLE, dbo.TB_USUARIOS.VCHA_USU_NOMBRE, dbo.TB_USUARIOS.VCHA_USU_APELLIDOS,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.VCHA_TIT_TITULAR_ID, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_ESB_ESTABLECIMIENTO_ID,"
      var_cadena = var_cadena + " dbo.TB_CLIENTES.CHAR_PRI_PRIORIDAD_ID, dbo.TB_PRIORIDADES.FLOA_PRI_NUMERO_ORDEN,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_DIAS_CONDICIONES, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_DIAS_CADUCIDAD,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_3, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_EMP_EMPRESA_ID,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_ALM_ALMACEN_ID,"
      var_cadena = var_cadena + " dbo.TB_TIPOPEDIDOS.FLOA_TPE_IVA, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_MON_MONEDA_ID,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_FACTURA_CEROS, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_SUGERIDO,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.CHAR_PED_TIPO , dbo.TB_ENCABEZADO_PEDIDOS.VCHA_PED_CLIENTE_REFERENCIA"
      rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
      lv_pedidos.SmallIcons = ImageList1
      lv_pedidos.ListItems.Clear
      While Not rs.EOF
         var_a = 0
         If IsNull(rs!inte_ped_autorizo) Then
            var_a = 0
         Else
            var_a = rs!inte_ped_autorizo
         End If
         var_fecha_fin = (CDate(Me.txt_fecha_fin) + 1)
         If var_a <> 1 And (rs!char_ped_estatus = "I" Or rs!char_ped_estatus = "S" Or rs!char_ped_estatus = "E") And (rs!dtim_ped_fecha >= var_fecha_inicio And rs!dtim_ped_fecha <= var_fecha_fin) Then
            Set list_item = lv_pedidos.ListItems.Add(, , rs!inte_ped_numero)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            list_item.SubItems(2) = IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE)
            list_item.SubItems(3) = IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE)
            list_item.SubItems(4) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            If IsNull(rs!Cantidad) Then
               list_item.SubItems(5) = Format(0, "###,###,##0.00")
            Else
               list_item.SubItems(5) = Format(rs!Cantidad, "###,###,##0.00")
            End If
            If IsNull(rs!Importe) Then
               list_item.SubItems(6) = Format(0, "###,###,##0.00")
            Else
               list_item.SubItems(6) = Format(rs!Importe, "###,###,##0.00")
            End If
            If IsNull(rs!inte_ped_autorizo) Then
               list_item.SubItems(7) = 0
            Else
               list_item.SubItems(7) = rs!inte_ped_autorizo
            End If
            If IsNull(rs!VCHA_PED_AUTORIZO) Then
               list_item.SubItems(8) = ""
            Else
               list_item.SubItems(8) = rs!VCHA_PED_AUTORIZO
            End If
            If IsNull(rs!DTIM_PED_AUTORIZO) Then
               list_item.SubItems(9) = ""
            Else
               list_item.SubItems(9) = rs!DTIM_PED_AUTORIZO
            End If
            If IsNull(rs!dtim_ped_fecha) Then
               list_item.SubItems(10) = ""
            Else
               list_item.SubItems(10) = rs!dtim_ped_fecha
            End If
            If IsNull(rs!char_ped_estatus) Then
               list_item.SubItems(11) = ""
            Else
               list_item.SubItems(11) = rs!char_ped_estatus
            End If
            If IsNull(rs!floa_ped_descuento_1) Then
               list_item.SubItems(12) = 0
            Else
               list_item.SubItems(12) = rs!floa_ped_descuento_1
            End If
            If IsNull(rs!FLOA_PED_DESCUENTO_2) Then
               list_item.SubItems(13) = 0
            Else
               list_item.SubItems(13) = rs!FLOA_PED_DESCUENTO_2
            End If
            If IsNull(rs!VCHA_USU_NOMBRE) Then
               list_item.SubItems(14) = ""
            Else
               If IsNull(rs!vcha_usu_apellidos) Then
                  list_item.SubItems(14) = rs!VCHA_USU_NOMBRE
               Else
                  list_item.SubItems(14) = Trim(rs!VCHA_USU_NOMBRE) + " " + Trim(rs!vcha_usu_apellidos)
               End If
            End If
         
            rsaux4.Open "select * from tb_titulares where vcha_tit_titular_id = '" + rs!vcha_tit_titular_id + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux4.EOF Then
               var_clave_grupo_real = IIf(IsNull(rsaux4!vcha_gre_grupo_real_id), "", rsaux4!vcha_gre_grupo_real_id)
            Else
               var_clave_grupo_real = ""
            End If
            rsaux4.Close
            If var_clave_grupo_real <> "" Then
               rsaux4.Open "select * from tb_gruposreales where vcha_gre_grupo_real_id = '" + var_clave_grupo_real + "' ", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux4.EOF Then
                  var_clave_grupo_actual = IIf(IsNull(rsaux4!VCHA_GAC_GRUPO_aCTUAL_ID), "", rsaux4!VCHA_GAC_GRUPO_aCTUAL_ID)
                Else
                   var_clave_grupo_actual = ""
               End If
               rsaux4.Close
               list_item.SubItems(17) = var_clave_grupo_actual
               If Trim(var_clave_grupo_actual) <> "" Then
                  rsaux2.Open "select sum(floa_sal_importe) from VW_FACTURAS_VENCIADAS where vcha_gac_grupo_actual_id = '" + var_clave_grupo_actual + "'  and dias > 0", cnn, adOpenDynamic, adLockOptimistic
                  'var_cadena = "SELECT dbo.TB_SALDOS.VCHA_EMP_EMPRESA_ID, dbo.TB_SALDOS.VCHA_SER_SERIE_ID, dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO, dbo.TB_SALDOS.VCHA_CLI_CLAVE_ID, dbo.TB_SALDOS.INTE_CAR_NUMERO, dbo.TB_SALDOS.FLOA_SAL_IMPORTE, dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA, dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_PLAZO, DATEDIFF([day], dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA + dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_PLAZO, GETDATE()) AS dias, dbo.TB_ENCABEZADO_CARTERA.VCHA_MON_MONEDA_ID, dbo.TB_MONEDAS.VCHA_MON_NOMBRE_PLURAL, dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_IMPORTE_NETO / dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_TIPO_CAMBIO AS IMPORTE_FACTURA, dbo.TB_ENCABEZADO_CARTERA.VCHA_AGE_AGENTE_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_GAC_GRUPO_ACTUAL_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_GRE_GRUPO_REAL_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_TIT_TITULAR_ID, dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_TIPO_CAMBIO FROM  dbo.TB_SALDOS INNER JOIN "
                  'var_cadena = var_cadena + " dbo.TB_ENCABEZADO_CARTERA ON dbo.TB_SALDOS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_SALDOS.VCHA_SER_SERIE_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID AND dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO = dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO AND dbo.TB_SALDOS.INTE_CAR_NUMERO = dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO INNER JOIN dbo.TB_MONEDAS ON dbo.TB_SALDOS.VCHA_MON_MONEDA_ID = dbo.TB_MONEDAS.VCHA_MON_MONEDA_ID WHERE (dbo.TB_SALDOS.FLOA_SAL_IMPORTE > 0) AND (dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO = 'FA') AND (dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS <> 'C' OR dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS IS NULL and dbo.TB_ENCABEZADO_CARTERA.vcha_gac_grupo_Actual_id = '" + var_clave_grupo_actual + "' and DATEDIFF([day], dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA + dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_PLAZO, GETDATE()) > 0) "
                  'rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  
                  
                  
                  If Not rsaux2.EOF Then
                     list_item.SubItems(16) = Format(IIf(IsNull(rsaux2(0).Value), 0, rsaux2(0).Value), "###,###,##0.00")
                  End If
                  rsaux2.Close
                End If
            End If
         
         End If
         
         rs.MoveNext:
      Wend
      rs.Close
      pro_textos
   End If
   If Index = 1 Then
      var_cadena = "SELECT     dbo.TB_DETALLE_PEDIDOS.INTE_PED_NUMERO, SUM(((dbo.TB_DETALLE_PEDIDOS.FLOA_PED_PRECIO * ((100 - dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_1) / 100)) * (100 - dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_2) / 100) * (1 + dbo.TB_TIPOPEDIDOS.FLOA_TPE_IVA / 100) * dbo.TB_DETALLE_PEDIDOS.FLOA_PED_CANTIDAD) AS importe, SUM(dbo.TB_DETALLE_PEDIDOS.FLOA_PED_CANTIDAD) AS cantidad, dbo.TB_ENCABEZADO_PEDIDOS.CHAR_TPE_TIPO_PEDIDO_ID, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_REFERENCIA, dbo.TB_ENCABEZADO_PEDIDOS.DTIM_PED_REFERENCIA, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_CLI_CLAVE_ID,"
      var_cadena = var_cadena + " dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_ESPECIALES, dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_ENCABEZADO_PEDIDOS.DTIM_PED_FECHA, dbo.TB_TITULARES.VCHA_TIT_NOMBRE, dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_NOMBRE, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_AUTORIZO, dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_1, dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_2, dbo.TB_ENCABEZADO_PEDIDOS.CHAR_PED_ESTATUS, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_PED_AUTORIZO, dbo.TB_ENCABEZADO_PEDIDOS.DTIM_PED_AUTORIZO, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_RESURTIBLE, dbo.TB_USUARIOS.VCHA_USU_NOMBRE, dbo.TB_USUARIOS.VCHA_USU_APELLIDOS, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_TIT_TITULAR_ID, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_ESB_ESTABLECIMIENTO_ID, dbo.TB_CLIENTES.CHAR_PRI_PRIORIDAD_ID, dbo.TB_PRIORIDADES.FLOA_PRI_NUMERO_ORDEN, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_DIAS_CONDICIONES,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_DIAS_CADUCIDAD, dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_3, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_ALM_ALMACEN_ID, dbo.TB_TIPOPEDIDOS.FLOA_TPE_IVA, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_MON_MONEDA_ID, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_FACTURA_CEROS, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_SUGERIDO, dbo.TB_ENCABEZADO_PEDIDOS.CHAR_PED_TIPO, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_PED_CLIENTE_REFERENCIA FROM         dbo.TB_TITULARES RIGHT OUTER JOIN dbo.TB_ENCABEZADO_PEDIDOS INNER JOIN dbo.TB_DETALLE_PEDIDOS ON dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_NUMERO = dbo.TB_DETALLE_PEDIDOS.INTE_PED_NUMERO INNER JOIN dbo.TB_TIPOPEDIDOS ON dbo.TB_ENCABEZADO_PEDIDOS.CHAR_TPE_TIPO_PEDIDO_ID = dbo.TB_TIPOPEDIDOS.CHAR_TPE_TIPO_PEDIDO_ID LEFT OUTER JOIN dbo.TB_USUARIOS ON dbo.TB_ENCABEZADO_PEDIDOS.VCHA_PED_AUTORIZO = dbo.TB_USUARIOS.VCHA_USU_USUARIO_ID LEFT OUTER JOIN"
      var_cadena = var_cadena + " dbo.TB_ESTABLECIMIENTOS ON"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.VCHA_ESB_ESTABLECIMIENTO_ID = dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_ESTABLECIMIENTO_ID ON"
      var_cadena = var_cadena + " dbo.TB_TITULARES.VCHA_TIT_TITULAR_ID = dbo.TB_ENCABEZADO_PEDIDOS.VCHA_TIT_TITULAR_ID LEFT OUTER JOIN"
      var_cadena = var_cadena + " dbo.TB_CLIENTES LEFT OUTER JOIN"
      var_cadena = var_cadena + " dbo.TB_PRIORIDADES ON dbo.TB_CLIENTES.CHAR_PRI_PRIORIDAD_ID = dbo.TB_PRIORIDADES.CHAR_PRI_PRIORIDAD_ID ON"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID LEFT OUTER JOIN"
      var_cadena = var_cadena + " dbo.TB_AGENTES ON dbo.TB_CLIENTES.VCHA_AGE_AGENTE_ID = dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID"
      var_cadena = var_cadena + " GROUP BY dbo.TB_DETALLE_PEDIDOS.INTE_PED_NUMERO, dbo.TB_ENCABEZADO_PEDIDOS.CHAR_TPE_TIPO_PEDIDO_ID,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_REFERENCIA, dbo.TB_ENCABEZADO_PEDIDOS.DTIM_PED_REFERENCIA,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.VCHA_CLI_CLAVE_ID, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_ESPECIALES, dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.DTIM_PED_FECHA, dbo.TB_TITULARES.VCHA_TIT_NOMBRE, dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_NOMBRE,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_AUTORIZO, dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_1,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_2, dbo.TB_ENCABEZADO_PEDIDOS.CHAR_PED_ESTATUS,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.VCHA_PED_AUTORIZO, dbo.TB_ENCABEZADO_PEDIDOS.DTIM_PED_AUTORIZO,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_RESURTIBLE, dbo.TB_USUARIOS.VCHA_USU_NOMBRE, dbo.TB_USUARIOS.VCHA_USU_APELLIDOS,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.VCHA_TIT_TITULAR_ID, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_ESB_ESTABLECIMIENTO_ID,"
      var_cadena = var_cadena + " dbo.TB_CLIENTES.CHAR_PRI_PRIORIDAD_ID, dbo.TB_PRIORIDADES.FLOA_PRI_NUMERO_ORDEN,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_DIAS_CONDICIONES, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_DIAS_CADUCIDAD,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_3, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_EMP_EMPRESA_ID,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_ALM_ALMACEN_ID,"
      var_cadena = var_cadena + " dbo.TB_TIPOPEDIDOS.FLOA_TPE_IVA, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_MON_MONEDA_ID,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_FACTURA_CEROS, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_SUGERIDO,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.CHAR_PED_TIPO , dbo.TB_ENCABEZADO_PEDIDOS.VCHA_PED_CLIENTE_REFERENCIA"
      rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
      lv_pedidos.SmallIcons = ImageList1
      lv_pedidos.ListItems.Clear
      While Not rs.EOF
         var_fecha_fin = (CDate(Me.txt_fecha_fin) + 1)
         If rs!inte_ped_autorizo = 1 And (rs!char_ped_estatus = "I" Or rs!char_ped_estatus = "S" Or rs!char_ped_estatus = "E") And rs!dtim_ped_fecha >= var_fecha_inicio And rs!dtim_ped_fecha <= var_fecha_fin Then
            Set list_item = lv_pedidos.ListItems.Add(, , rs!inte_ped_numero)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            list_item.SubItems(2) = IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE)
            list_item.SubItems(3) = IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE)
            list_item.SubItems(4) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            If IsNull(rs!Cantidad) Then
               list_item.SubItems(5) = Format(0, "###,###,##0.00")
            Else
               list_item.SubItems(5) = Format(rs!Cantidad, "###,###,##0.00")
            End If
            If IsNull(rs!Importe) Then
               list_item.SubItems(6) = Format(0, "###,###,##0.00")
            Else
               list_item.SubItems(6) = Format(rs!Importe, "###,###,##0.00")
            End If
            If IsNull(rs!inte_ped_autorizo) Then
               list_item.SubItems(7) = 0
            Else
               list_item.SubItems(7) = rs!inte_ped_autorizo
               If rs!inte_ped_autorizo = 1 Then
                  list_item.SmallIcon = 13
                  list_item.Bold = True
                  list_item.ForeColor = &H8000&
                  list_item.ListSubItems.item(1).Bold = True
                  list_item.ListSubItems.item(2).Bold = True
                  list_item.ListSubItems.item(3).Bold = True
                  list_item.ListSubItems.item(4).Bold = True
                  list_item.ListSubItems.item(5).Bold = True
                  list_item.ListSubItems.item(6).Bold = True
                  list_item.ListSubItems.item(1).ForeColor = &H8000&
                  list_item.ListSubItems.item(2).ForeColor = &H8000&
                  list_item.ListSubItems.item(3).ForeColor = &H8000&
                  list_item.ListSubItems.item(4).ForeColor = &H8000&
                  list_item.ListSubItems.item(5).ForeColor = &H8000&
                  list_item.ListSubItems.item(6).ForeColor = &H8000&
               End If
            End If
            If IsNull(rs!VCHA_PED_AUTORIZO) Then
               list_item.SubItems(8) = ""
            Else
               list_item.SubItems(8) = rs!VCHA_PED_AUTORIZO
            End If
            If IsNull(rs!DTIM_PED_AUTORIZO) Then
               list_item.SubItems(9) = ""
            Else
               list_item.SubItems(9) = rs!DTIM_PED_AUTORIZO
            End If
            If IsNull(rs!dtim_ped_fecha) Then
               list_item.SubItems(10) = ""
            Else
               list_item.SubItems(10) = rs!dtim_ped_fecha
            End If
            If IsNull(rs!char_ped_estatus) Then
               list_item.SubItems(11) = ""
            Else
               list_item.SubItems(11) = rs!char_ped_estatus
            End If
            If IsNull(rs!floa_ped_descuento_1) Then
               list_item.SubItems(12) = 0
            Else
               list_item.SubItems(12) = rs!floa_ped_descuento_1
            End If
            If IsNull(rs!FLOA_PED_DESCUENTO_2) Then
               list_item.SubItems(13) = 0
            Else
               list_item.SubItems(13) = rs!FLOA_PED_DESCUENTO_2
            End If
            If IsNull(rs!VCHA_USU_NOMBRE) Then
               list_item.SubItems(14) = ""
            Else
               If IsNull(rs!vcha_usu_apellidos) Then
                  list_item.SubItems(14) = rs!VCHA_USU_NOMBRE
               Else
                  list_item.SubItems(14) = Trim(rs!VCHA_USU_NOMBRE) + " " + Trim(rs!vcha_usu_apellidos)
               End If
            End If
         
         
            rsaux4.Open "select * from tb_titulares where vcha_tit_titular_id = '" + rs!vcha_tit_titular_id + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux4.EOF Then
               var_clave_grupo_real = IIf(IsNull(rsaux4!vcha_gre_grupo_real_id), "", rsaux4!vcha_gre_grupo_real_id)
            Else
               var_clave_grupo_real = ""
            End If
            rsaux4.Close
            If var_clave_grupo_real <> "" Then
               rsaux4.Open "select * from tb_gruposreales where vcha_gre_grupo_real_id = '" + var_clave_grupo_real + "' ", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux4.EOF Then
                  var_clave_grupo_actual = IIf(IsNull(rsaux4!VCHA_GAC_GRUPO_aCTUAL_ID), "", rsaux4!VCHA_GAC_GRUPO_aCTUAL_ID)
               Else
                  var_clave_grupo_actual = ""
               End If
               rsaux4.Close
               list_item.SubItems(17) = var_clave_grupo_actual
               If Trim(var_clave_grupo_actual) <> "" Then
                  rsaux2.Open "select sum(floa_sal_importe) from VW_FACTURAS_VENCIADAS where vcha_gac_grupo_actual_id = '" + var_clave_grupo_actual + "'  and dias > 0", cnn, adOpenDynamic, adLockOptimistic
                  'var_cadena = "SELECT dbo.TB_SALDOS.VCHA_EMP_EMPRESA_ID, dbo.TB_SALDOS.VCHA_SER_SERIE_ID, dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO, dbo.TB_SALDOS.VCHA_CLI_CLAVE_ID, dbo.TB_SALDOS.INTE_CAR_NUMERO, dbo.TB_SALDOS.FLOA_SAL_IMPORTE, dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA, dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_PLAZO, DATEDIFF([day], dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA + dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_PLAZO, GETDATE()) AS dias, dbo.TB_ENCABEZADO_CARTERA.VCHA_MON_MONEDA_ID, dbo.TB_MONEDAS.VCHA_MON_NOMBRE_PLURAL, dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_IMPORTE_NETO / dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_TIPO_CAMBIO AS IMPORTE_FACTURA, dbo.TB_ENCABEZADO_CARTERA.VCHA_AGE_AGENTE_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_GAC_GRUPO_ACTUAL_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_GRE_GRUPO_REAL_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_TIT_TITULAR_ID, dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_TIPO_CAMBIO FROM  dbo.TB_SALDOS INNER JOIN "
                  'var_cadena = var_cadena + " dbo.TB_ENCABEZADO_CARTERA ON dbo.TB_SALDOS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_SALDOS.VCHA_SER_SERIE_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID AND dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO = dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO AND dbo.TB_SALDOS.INTE_CAR_NUMERO = dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO INNER JOIN dbo.TB_MONEDAS ON dbo.TB_SALDOS.VCHA_MON_MONEDA_ID = dbo.TB_MONEDAS.VCHA_MON_MONEDA_ID WHERE (dbo.TB_SALDOS.FLOA_SAL_IMPORTE > 0) AND (dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO = 'FA') AND (dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS <> 'C' OR dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS IS NULL and dbo.TB_ENCABEZADO_CARTERA.vcha_gac_grupo_Actual_id = '" + var_clave_grupo_actual + "' and DATEDIFF([day], dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA + dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_PLAZO, GETDATE()) > 0) "
                  'rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  
                  
                  If Not rsaux2.EOF Then
                     list_item.SubItems(16) = Format(IIf(IsNull(rsaux2(0).Value), 0, rsaux2(0).Value), "###,###,##0.00")
                  End If
                  rsaux2.Close
                End If
            End If
         
         
         
         End If
         rs.MoveNext:
      Wend
      rs.Close
      pro_textos
   End If
   If Index = 2 Then
      var_cadena = "SELECT     dbo.TB_DETALLE_PEDIDOS.INTE_PED_NUMERO, SUM(((dbo.TB_DETALLE_PEDIDOS.FLOA_PED_PRECIO * ((100 - dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_1) / 100)) * (100 - dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_2) / 100) * (1 + dbo.TB_TIPOPEDIDOS.FLOA_TPE_IVA / 100) * dbo.TB_DETALLE_PEDIDOS.FLOA_PED_CANTIDAD) AS importe, SUM(dbo.TB_DETALLE_PEDIDOS.FLOA_PED_CANTIDAD) AS cantidad, dbo.TB_ENCABEZADO_PEDIDOS.CHAR_TPE_TIPO_PEDIDO_ID, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_REFERENCIA, dbo.TB_ENCABEZADO_PEDIDOS.DTIM_PED_REFERENCIA, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_CLI_CLAVE_ID,"
      var_cadena = var_cadena + " dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_ESPECIALES, dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_ENCABEZADO_PEDIDOS.DTIM_PED_FECHA, dbo.TB_TITULARES.VCHA_TIT_NOMBRE, dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_NOMBRE, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_AUTORIZO, dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_1, dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_2, dbo.TB_ENCABEZADO_PEDIDOS.CHAR_PED_ESTATUS, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_PED_AUTORIZO, dbo.TB_ENCABEZADO_PEDIDOS.DTIM_PED_AUTORIZO, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_RESURTIBLE, dbo.TB_USUARIOS.VCHA_USU_NOMBRE, dbo.TB_USUARIOS.VCHA_USU_APELLIDOS, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_TIT_TITULAR_ID, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_ESB_ESTABLECIMIENTO_ID, dbo.TB_CLIENTES.CHAR_PRI_PRIORIDAD_ID, dbo.TB_PRIORIDADES.FLOA_PRI_NUMERO_ORDEN, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_DIAS_CONDICIONES,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_DIAS_CADUCIDAD, dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_3, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_EMP_EMPRESA_ID, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_ALM_ALMACEN_ID, dbo.TB_TIPOPEDIDOS.FLOA_TPE_IVA, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_MON_MONEDA_ID, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_FACTURA_CEROS, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_SUGERIDO, dbo.TB_ENCABEZADO_PEDIDOS.CHAR_PED_TIPO, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_PED_CLIENTE_REFERENCIA FROM         dbo.TB_TITULARES RIGHT OUTER JOIN dbo.TB_ENCABEZADO_PEDIDOS INNER JOIN dbo.TB_DETALLE_PEDIDOS ON dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_NUMERO = dbo.TB_DETALLE_PEDIDOS.INTE_PED_NUMERO INNER JOIN dbo.TB_TIPOPEDIDOS ON dbo.TB_ENCABEZADO_PEDIDOS.CHAR_TPE_TIPO_PEDIDO_ID = dbo.TB_TIPOPEDIDOS.CHAR_TPE_TIPO_PEDIDO_ID LEFT OUTER JOIN dbo.TB_USUARIOS ON dbo.TB_ENCABEZADO_PEDIDOS.VCHA_PED_AUTORIZO = dbo.TB_USUARIOS.VCHA_USU_USUARIO_ID LEFT OUTER JOIN"
      var_cadena = var_cadena + " dbo.TB_ESTABLECIMIENTOS ON"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.VCHA_ESB_ESTABLECIMIENTO_ID = dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_ESTABLECIMIENTO_ID ON"
      var_cadena = var_cadena + " dbo.TB_TITULARES.VCHA_TIT_TITULAR_ID = dbo.TB_ENCABEZADO_PEDIDOS.VCHA_TIT_TITULAR_ID LEFT OUTER JOIN"
      var_cadena = var_cadena + " dbo.TB_CLIENTES LEFT OUTER JOIN"
      var_cadena = var_cadena + " dbo.TB_PRIORIDADES ON dbo.TB_CLIENTES.CHAR_PRI_PRIORIDAD_ID = dbo.TB_PRIORIDADES.CHAR_PRI_PRIORIDAD_ID ON"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID LEFT OUTER JOIN"
      var_cadena = var_cadena + " dbo.TB_AGENTES ON dbo.TB_CLIENTES.VCHA_AGE_AGENTE_ID = dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID"
      var_cadena = var_cadena + " GROUP BY dbo.TB_DETALLE_PEDIDOS.INTE_PED_NUMERO, dbo.TB_ENCABEZADO_PEDIDOS.CHAR_TPE_TIPO_PEDIDO_ID,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_REFERENCIA, dbo.TB_ENCABEZADO_PEDIDOS.DTIM_PED_REFERENCIA,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.VCHA_CLI_CLAVE_ID, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_ESPECIALES, dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.DTIM_PED_FECHA, dbo.TB_TITULARES.VCHA_TIT_NOMBRE, dbo.TB_ESTABLECIMIENTOS.VCHA_ESB_NOMBRE,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_AUTORIZO, dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_1,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_2, dbo.TB_ENCABEZADO_PEDIDOS.CHAR_PED_ESTATUS,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.VCHA_PED_AUTORIZO, dbo.TB_ENCABEZADO_PEDIDOS.DTIM_PED_AUTORIZO,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_RESURTIBLE, dbo.TB_USUARIOS.VCHA_USU_NOMBRE, dbo.TB_USUARIOS.VCHA_USU_APELLIDOS,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.VCHA_TIT_TITULAR_ID, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_ESB_ESTABLECIMIENTO_ID,"
      var_cadena = var_cadena + " dbo.TB_CLIENTES.CHAR_PRI_PRIORIDAD_ID, dbo.TB_PRIORIDADES.FLOA_PRI_NUMERO_ORDEN,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_DIAS_CONDICIONES, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_DIAS_CADUCIDAD,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.FLOA_PED_DESCUENTO_3, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_EMP_EMPRESA_ID,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.VCHA_UOR_UNIDAD_ID, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_ALM_ALMACEN_ID,"
      var_cadena = var_cadena + " dbo.TB_TIPOPEDIDOS.FLOA_TPE_IVA, dbo.TB_ENCABEZADO_PEDIDOS.VCHA_MON_MONEDA_ID,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_FACTURA_CEROS, dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_SUGERIDO,"
      var_cadena = var_cadena + " dbo.TB_ENCABEZADO_PEDIDOS.CHAR_PED_TIPO , dbo.TB_ENCABEZADO_PEDIDOS.VCHA_PED_CLIENTE_REFERENCIA"
      rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
      lv_pedidos.SmallIcons = ImageList1
      lv_pedidos.ListItems.Clear
      While Not rs.EOF
         If rs!char_ped_estatus = "C" And rs!dtim_ped_fecha >= var_fecha_inicio And rs!dtim_ped_fecha <= var_fecha_fin Then
            Set list_item = lv_pedidos.ListItems.Add(, , rs!inte_ped_numero)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            list_item.SubItems(2) = IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE)
            list_item.SubItems(3) = IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE)
            list_item.SubItems(4) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            If IsNull(rs!Cantidad) Then
               list_item.SubItems(5) = Format(0, "###,###,##0.00")
            Else
               list_item.SubItems(5) = Format(rs!Cantidad, "###,###,##0.00")
            End If
            If IsNull(rs!Importe) Then
               list_item.SubItems(6) = Format(0, "###,###,##0.00")
            Else
               list_item.SubItems(6) = Format(rs!Importe, "###,###,##0.00")
            End If
            If IsNull(rs!inte_ped_autorizo) Then
               list_item.SubItems(7) = 0
            Else
               list_item.SubItems(7) = rs!inte_ped_autorizo
            End If
            If IsNull(rs!VCHA_PED_AUTORIZO) Then
               list_item.SubItems(8) = ""
            Else
               list_item.SubItems(8) = rs!VCHA_PED_AUTORIZO
            End If
            If IsNull(rs!DTIM_PED_AUTORIZO) Then
               list_item.SubItems(9) = ""
            Else
               list_item.SubItems(9) = rs!DTIM_PED_AUTORIZO
            End If
            If IsNull(rs!dtim_ped_fecha) Then
               list_item.SubItems(10) = ""
            Else
               list_item.SubItems(10) = rs!dtim_ped_fecha
            End If
            If IsNull(rs!char_ped_estatus) Then
               list_item.SubItems(11) = ""
            Else
               list_item.SubItems(11) = rs!char_ped_estatus
            End If
            If IsNull(rs!floa_ped_descuento_1) Then
               list_item.SubItems(12) = 0
            Else
               list_item.SubItems(12) = rs!floa_ped_descuento_1
            End If
            If IsNull(rs!FLOA_PED_DESCUENTO_2) Then
               list_item.SubItems(13) = 0
            Else
               list_item.SubItems(13) = rs!FLOA_PED_DESCUENTO_2
            End If
            If IsNull(rs!VCHA_USU_NOMBRE) Then
               list_item.SubItems(14) = ""
            Else
               If IsNull(rs!vcha_usu_apellidos) Then
                  list_item.SubItems(14) = rs!VCHA_USU_NOMBRE
               Else
                  list_item.SubItems(14) = Trim(rs!VCHA_USU_NOMBRE) + " " + Trim(rs!vcha_usu_apellidos)
               End If
            End If
            list_item.SmallIcon = 15
            list_item.Bold = True
            list_item.ForeColor = &HFF&
            list_item.ListSubItems.item(1).Bold = True
            list_item.ListSubItems.item(2).Bold = True
            list_item.ListSubItems.item(3).Bold = True
            list_item.ListSubItems.item(4).Bold = True
            list_item.ListSubItems.item(5).Bold = True
            list_item.ListSubItems.item(6).Bold = True
            list_item.ListSubItems.item(1).ForeColor = &HFF&
            list_item.ListSubItems.item(2).ForeColor = &HFF&
            list_item.ListSubItems.item(3).ForeColor = &HFF&
            list_item.ListSubItems.item(4).ForeColor = &HFF&
            list_item.ListSubItems.item(5).ForeColor = &HFF&
            list_item.ListSubItems.item(6).ForeColor = &HFF&
         
         
            rsaux4.Open "select * from tb_titulares where vcha_tit_titular_id = '" + rs!vcha_tit_titular_id + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux4.EOF Then
               var_clave_grupo_real = IIf(IsNull(rsaux4!vcha_gre_grupo_real_id), "", rsaux4!vcha_gre_grupo_real_id)
            Else
               var_clave_grupo_real = ""
            End If
            rsaux4.Close
            If var_clave_grupo_real <> "" Then
               rsaux4.Open "select * from tb_gruposreales where vcha_gre_grupo_real_id = '" + var_clave_grupo_real + "' ", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux4.EOF Then
                  var_clave_grupo_actual = IIf(IsNull(rsaux4!VCHA_GAC_GRUPO_aCTUAL_ID), "", rsaux4!VCHA_GAC_GRUPO_aCTUAL_ID)
               Else
                  var_clave_grupo_actual = ""
               End If
               rsaux4.Close
               list_item.SubItems(17) = var_clave_grupo_actual
               If Trim(var_clave_grupo_actual) <> "" Then
                  rsaux2.Open "select sum(floa_sal_importe) from VW_FACTURAS_VENCIADAS where vcha_gac_grupo_actual_id = '" + var_clave_grupo_actual + "'  and dias > 0", cnn, adOpenDynamic, adLockOptimistic
                  'var_cadena = "SELECT dbo.TB_SALDOS.VCHA_EMP_EMPRESA_ID, dbo.TB_SALDOS.VCHA_SER_SERIE_ID, dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO, dbo.TB_SALDOS.VCHA_CLI_CLAVE_ID, dbo.TB_SALDOS.INTE_CAR_NUMERO, dbo.TB_SALDOS.FLOA_SAL_IMPORTE, dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA, dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_PLAZO, DATEDIFF([day], dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA + dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_PLAZO, GETDATE()) AS dias, dbo.TB_ENCABEZADO_CARTERA.VCHA_MON_MONEDA_ID, dbo.TB_MONEDAS.VCHA_MON_NOMBRE_PLURAL, dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_IMPORTE_NETO / dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_TIPO_CAMBIO AS IMPORTE_FACTURA, dbo.TB_ENCABEZADO_CARTERA.VCHA_AGE_AGENTE_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_GAC_GRUPO_ACTUAL_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_GRE_GRUPO_REAL_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_TIT_TITULAR_ID, dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_TIPO_CAMBIO FROM  dbo.TB_SALDOS INNER JOIN "
                  'var_cadena = var_cadena + " dbo.TB_ENCABEZADO_CARTERA ON dbo.TB_SALDOS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_SALDOS.VCHA_SER_SERIE_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID AND dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO = dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO AND dbo.TB_SALDOS.INTE_CAR_NUMERO = dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO INNER JOIN dbo.TB_MONEDAS ON dbo.TB_SALDOS.VCHA_MON_MONEDA_ID = dbo.TB_MONEDAS.VCHA_MON_MONEDA_ID WHERE (dbo.TB_SALDOS.FLOA_SAL_IMPORTE > 0) AND (dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO = 'FA') AND (dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS <> 'C' OR dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS IS NULL and dbo.TB_ENCABEZADO_CARTERA.vcha_gac_grupo_Actual_id = '" + var_clave_grupo_actual + "' and DATEDIFF([day], dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA + dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_PLAZO, GETDATE()) > 0) "
                  'rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  
                  
                  
                  If Not rsaux2.EOF Then
                      list_item.SubItems(16) = Format(IIf(IsNull(rsaux2(0).Value), 0, rsaux2(0).Value), "###,###,##0.00")
                  End If
                  rsaux2.Close
                End If
            End If
         
         
         
         End If
         rs.MoveNext:
      Wend
      rs.Close
      pro_textos
   End If
End Sub

Private Sub opt_filtro_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 27 Then
      frm_filtro.Visible = False
   End If
End Sub


Private Sub mnu_autorizar_click()
   Set TB_AUTORIZA_PEDIDOS = New TB_AUTORIZA_PEDIDOS
   ok = TB_AUTORIZA_PEDIDOS.Anadir(lv_pedidos.selectedItem, 1, var_clave_usuario_global, Now)
   If ok Then
      lv_pedidos.selectedItem.SubItems(7) = 1
      lv_pedidos.selectedItem.SmallIcon = 13
      lv_pedidos.selectedItem.Bold = True
      lv_pedidos.selectedItem.ForeColor = &H8000&
      lv_pedidos.selectedItem.ListSubItems.item(1).Bold = True
      lv_pedidos.selectedItem.ListSubItems.item(2).Bold = True
      lv_pedidos.selectedItem.ListSubItems.item(3).Bold = True
      lv_pedidos.selectedItem.ListSubItems.item(4).Bold = True
      lv_pedidos.selectedItem.ListSubItems.item(5).Bold = True
      lv_pedidos.selectedItem.ListSubItems.item(6).Bold = True
      lv_pedidos.selectedItem.ListSubItems.item(1).ForeColor = &H8000&
      lv_pedidos.selectedItem.ListSubItems.item(2).ForeColor = &H8000&
      lv_pedidos.selectedItem.ListSubItems.item(3).ForeColor = &H8000&
      lv_pedidos.selectedItem.ListSubItems.item(4).ForeColor = &H8000&
      lv_pedidos.selectedItem.ListSubItems.item(5).ForeColor = &H8000&
      lv_pedidos.selectedItem.ListSubItems.item(6).ForeColor = &H8000&
      MsgBox "El pedido a sido autorizado", vbOKOnly, "ATENCION"
      pro_textos
      lv_pedidos.selectedItem.SubItems(14) = var_nombre_usuario_global + " " + var_apellidos_usuario_global
   End If
End Sub

Private Sub mnu_cancelar_autorizacion_click()
   Set TB_AUTORIZA_PEDIDOS = New TB_AUTORIZA_PEDIDOS
   ok = TB_AUTORIZA_PEDIDOS.Anadir(lv_pedidos.selectedItem, 0, var_clave_usuario_global, Now)
   If ok Then
      lv_pedidos.selectedItem.SubItems(7) = 0
      lv_pedidos.selectedItem.SmallIcon = 0
      lv_pedidos.selectedItem.SmallIcon = 0
      lv_pedidos.selectedItem.Bold = False
      lv_pedidos.selectedItem.ForeColor = &H0&
      lv_pedidos.selectedItem.ListSubItems.item(1).Bold = False
      lv_pedidos.selectedItem.ListSubItems.item(2).Bold = False
      lv_pedidos.selectedItem.ListSubItems.item(3).Bold = False
      lv_pedidos.selectedItem.ListSubItems.item(4).Bold = False
      lv_pedidos.selectedItem.ListSubItems.item(5).Bold = False
      lv_pedidos.selectedItem.ListSubItems.item(6).Bold = False
      lv_pedidos.selectedItem.ListSubItems.item(1).ForeColor = &H0&
      lv_pedidos.selectedItem.ListSubItems.item(2).ForeColor = &H0&
      lv_pedidos.selectedItem.ListSubItems.item(3).ForeColor = &H0&
      lv_pedidos.selectedItem.ListSubItems.item(4).ForeColor = &H0&
      lv_pedidos.selectedItem.ListSubItems.item(5).ForeColor = &H0&
      lv_pedidos.selectedItem.ListSubItems.item(6).ForeColor = &H0&
      pro_textos
      lv_pedidos.selectedItem.SubItems(14) = var_nombre_usuario_global + " " + var_apellidos_usuario_global
      MsgBox "Se a cancelado la autorización", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Text3_Change()

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frm_correo.Visible = False
   End If
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
   Dim list_item As ListItem
   rs.Open "select * from vw_cheques_devueltos_con_saldo where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_cli_clave_id = '" + var_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
   lv_cheques.ListItems.Clear
   If Not rs.EOF Then
      While Not rs.EOF
         If Not IsNull(rs!vcha_car_cheque) Then
            Set list_item = lv_cheques.ListItems.Add(, , rs!vcha_car_cheque)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_BAN_NOMBRE), "", rs!VCHA_BAN_NOMBRE)
            list_item.SubItems(2) = IIf(IsNull(rs!dtim_Car_fecha), "", Format(rs!dtim_Car_fecha, "Short Date"))
            list_item.SubItems(3) = IIf(IsNull(rs!FLOA_sAL_IMPORTE), 0, rs!FLOA_sAL_IMPORTE)
         End If
         rs.MoveNext
      Wend
   Else
      MsgBox "El cliente no tiene cheques devueltos con saldo", vbOKOnly, "ATENCION"
   End If
   rs.Close
   frm_cheques_devueltos.Visible = True
   lv_cheques.SetFocus
End Sub

Private Sub mnu_cancelar_pedido_click()
   Set TB_CANCELA_PEDIDOS = New TB_CANCELA_PEDIDOS
   ok = TB_CANCELA_PEDIDOS.Anadir(lv_pedidos.selectedItem, "C")
   lv_pedidos.selectedItem.SubItems(7) = 1
   lv_pedidos.selectedItem.SmallIcon = 15
   lv_pedidos.selectedItem.Bold = True
   lv_pedidos.selectedItem.ForeColor = &HFF&
   lv_pedidos.selectedItem.ListSubItems.item(1).Bold = True
   lv_pedidos.selectedItem.ListSubItems.item(2).Bold = True
   lv_pedidos.selectedItem.ListSubItems.item(3).Bold = True
   lv_pedidos.selectedItem.ListSubItems.item(4).Bold = True
   lv_pedidos.selectedItem.ListSubItems.item(5).Bold = True
   lv_pedidos.selectedItem.ListSubItems.item(6).Bold = True
   lv_pedidos.selectedItem.ListSubItems.item(1).ForeColor = &HFF&
   lv_pedidos.selectedItem.ListSubItems.item(2).ForeColor = &HFF&
   lv_pedidos.selectedItem.ListSubItems.item(3).ForeColor = &HFF&
   lv_pedidos.selectedItem.ListSubItems.item(4).ForeColor = &HFF&
   lv_pedidos.selectedItem.ListSubItems.item(5).ForeColor = &HFF&
   lv_pedidos.selectedItem.ListSubItems.item(6).ForeColor = &HFF&
   MsgBox "El pedido a sido cancelado", vbOKOnly, "ATENCION"
End Sub

Private Sub mnu_reactivar_pedido_click()
   Set TB_CANCELA_PEDIDOS = New TB_CANCELA_PEDIDOS
   ok = TB_CANCELA_PEDIDOS.Anadir(lv_pedidos.selectedItem, " ")
      lv_pedidos.selectedItem.SubItems(7) = 0
      lv_pedidos.selectedItem.SmallIcon = 0
      lv_pedidos.selectedItem.SmallIcon = 0
      lv_pedidos.selectedItem.Bold = False
      lv_pedidos.selectedItem.ForeColor = &H0&
      lv_pedidos.selectedItem.ListSubItems.item(1).Bold = False
      lv_pedidos.selectedItem.ListSubItems.item(2).Bold = False
      lv_pedidos.selectedItem.ListSubItems.item(3).Bold = False
      lv_pedidos.selectedItem.ListSubItems.item(4).Bold = False
      lv_pedidos.selectedItem.ListSubItems.item(5).Bold = False
      lv_pedidos.selectedItem.ListSubItems.item(6).Bold = False
      lv_pedidos.selectedItem.ListSubItems.item(1).ForeColor = &H0&
      lv_pedidos.selectedItem.ListSubItems.item(2).ForeColor = &H0&
      lv_pedidos.selectedItem.ListSubItems.item(3).ForeColor = &H0&
      lv_pedidos.selectedItem.ListSubItems.item(4).ForeColor = &H0&
      lv_pedidos.selectedItem.ListSubItems.item(5).ForeColor = &H0&
      lv_pedidos.selectedItem.ListSubItems.item(6).ForeColor = &H0&
      MsgBox "Se a reactivado el pedido", vbOKOnly, "ATENCION"
End Sub



Private Sub pro_textos()
Dim var_autorizado As Integer
   On Error GoTo n:
   txt_numero = lv_pedidos.selectedItem
   txt_fecha = lv_pedidos.selectedItem.SubItems(10)
   txt_agente = lv_pedidos.selectedItem.SubItems(1)
   txt_titular = lv_pedidos.selectedItem.SubItems(2)
   txt_establecimiento = lv_pedidos.selectedItem.SubItems(3)
   txt_cliente = lv_pedidos.selectedItem.SubItems(4)
   txt_descuento1 = lv_pedidos.selectedItem.SubItems(12)
   txt_descuento2 = lv_pedidos.selectedItem.SubItems(13)
   txt_Cantidad = lv_pedidos.selectedItem.SubItems(5)
   txt_importe = lv_pedidos.selectedItem.SubItems(6)
   If lv_pedidos.selectedItem.SubItems(7) = 1 Then
      chk_autorizado = 1
      var_autorizado = 1
   Else
      chk_autorizado = 0
      var_autorizado = 0
   End If
   If var_autorizado = 1 Then
      txt_autorizo = lv_pedidos.selectedItem.SubItems(14)
      txt_fecha_autorizacion = lv_pedidos.selectedItem.SubItems(9)
   Else
      txt_autorizo = ""
      txt_fecha_autorizacion = ""
   End If
   var_clave_cliente = lv_pedidos.selectedItem.SubItems(15)
   txt_importe_vencido = lv_pedidos.selectedItem.SubItems(16)
   var_n = lv_pedidos.ListItems.Count
   var_numero_renglones = lv_pedidos.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_pedidos.ColumnHeaders(6).Width = 950
   Else
      lv_pedidos.ColumnHeaders(6).Width = 1150
   End If
   Me.txt_observacion = lv_pedidos.selectedItem.SubItems(18)
n:
End Sub

Private Sub Toolbar4_ButtonClick(ByVal Button As MSComctlLib.Button)
   Dim list_item As ListItem
   lv_facturas.ListItems.Clear
   var_clave_cliente = lv_pedidos.selectedItem.SubItems(15)
   var_clave_grupo_actual = lv_pedidos.selectedItem.SubItems(17)
   rs.Open "select * from vw_facturas_venciadas where vcha_gac_grupo_actual_id = '" + var_clave_grupo_actual + "' and dias > 0", cnn, adOpenDynamic, adLockOptimistic
   'var_cadena = "SELECT dbo.TB_SALDOS.VCHA_EMP_EMPRESA_ID, dbo.TB_SALDOS.VCHA_SER_SERIE_ID, dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO, dbo.TB_SALDOS.VCHA_CLI_CLAVE_ID, dbo.TB_SALDOS.INTE_CAR_NUMERO, dbo.TB_SALDOS.FLOA_SAL_IMPORTE, dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA, dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_PLAZO, DATEDIFF([day], dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA + dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_PLAZO, GETDATE()) AS dias, dbo.TB_ENCABEZADO_CARTERA.VCHA_MON_MONEDA_ID, dbo.TB_MONEDAS.VCHA_MON_NOMBRE_PLURAL, dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_IMPORTE_NETO / dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_TIPO_CAMBIO AS IMPORTE_FACTURA, dbo.TB_ENCABEZADO_CARTERA.VCHA_AGE_AGENTE_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_GAC_GRUPO_ACTUAL_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_GRE_GRUPO_REAL_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_TIT_TITULAR_ID, dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_TIPO_CAMBIO FROM  dbo.TB_SALDOS INNER JOIN "
   'var_cadena = var_cadena + " dbo.TB_ENCABEZADO_CARTERA ON dbo.TB_SALDOS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_SALDOS.VCHA_SER_SERIE_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID AND dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO = dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO AND dbo.TB_SALDOS.INTE_CAR_NUMERO = dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO INNER JOIN dbo.TB_MONEDAS ON dbo.TB_SALDOS.VCHA_MON_MONEDA_ID = dbo.TB_MONEDAS.VCHA_MON_MONEDA_ID WHERE (dbo.TB_SALDOS.FLOA_SAL_IMPORTE > 0) AND (dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO = 'FA') AND (dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS <> 'C' OR dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS IS NULL and dbo.TB_ENCABEZADO_CARTERA.vcha_gac_grupo_Actual_id = '" + var_clave_grupo_actual + "' and DATEDIFF([day], dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA + dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_PLAZO, GETDATE()) > 0) "
   'rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
   txt_total_facturas_venciadas = "0"
   If Not rs.EOF Then
      While Not rs.EOF
         Set list_item = lv_facturas.ListItems.Add(, , rs!inte_Car_numero)
         list_item.SubItems(1) = IIf(IsNull(rs!dtim_Car_fecha), "", Format(rs!dtim_Car_fecha, "Short Date"))
         list_item.SubItems(2) = IIf(IsNull(rs!dias), 0, rs!dias)
         list_item.SubItems(3) = Format(IIf(IsNull(rs!FLOA_sAL_IMPORTE), 0, rs!FLOA_sAL_IMPORTE), "###,###,##0.00")
         txt_total_facturas_venciadas = CDbl(txt_total_facturas_venciadas) + IIf(IsNull(rs!FLOA_sAL_IMPORTE), 0, rs!FLOA_sAL_IMPORTE)
         rs.MoveNext
      Wend
      txt_total_facturas_venciadas = Format(CDbl(txt_total_facturas_venciadas), "###,###,##0.00")
   Else
      MsgBox "El cliente no tiene facturas vencidas", vbOKOnly, "ATENCION"
   End If
   
   If lv_facturas.ListItems.Count > 9 Then
      lv_facturas.ColumnHeaders(1).Width = 850
   Else
      lv_facturas.ColumnHeaders(1).Width = 1099.84
   End If

   rs.Close
   frm_facturas.Visible = True
   lv_facturas.SetFocus
End Sub



Private Sub txt_fecha_fin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      frm_filtro.Visible = False
   End If
End Sub

Private Sub txt_fecha_inicio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      frm_filtro.Visible = False
   End If
End Sub

Private Sub txt_fin_correo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frm_correo.Visible = False
   End If
End Sub

Private Sub txt_inicio_correo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frm_correo.Visible = False
   End If
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
End Sub
