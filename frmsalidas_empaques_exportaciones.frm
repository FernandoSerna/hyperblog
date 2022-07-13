VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmsalidas_empaques_exportaciones_no_usar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   45
      Picture         =   "frmsalidas_empaques_exportaciones.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   570
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   375
      Picture         =   "frmsalidas_empaques_exportaciones.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Buscar Movimiento"
      Top             =   570
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   705
      Picture         =   "frmsalidas_empaques_exportaciones.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cerrar Caja e Imprimir las Etiquetas"
      Top             =   570
      Width           =   330
   End
   Begin VB.CommandButton cmd_cerrar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1035
      Picture         =   "frmsalidas_empaques_exportaciones.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Cerrar para surtir Alt + C"
      Top             =   570
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11220
      Picture         =   "frmsalidas_empaques_exportaciones.frx":0408
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   570
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   0
      Left            =   60
      TabIndex        =   66
      Top             =   435
      Width           =   11580
   End
   Begin VB.TextBox txt_clave_movimiento 
      Height          =   285
      Left            =   2205
      TabIndex        =   51
      Top             =   660
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox txt_foco 
      Height          =   285
      Left            =   12330
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   1800
      Width           =   2100
   End
   Begin VB.Frame frm_detalle 
      Height          =   2190
      Left            =   2220
      TabIndex        =   12
      Top             =   975
      Width           =   5730
      Begin VB.TextBox txt_pedido 
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1785
         Width           =   2190
      End
      Begin VB.TextBox txt_establecimiento 
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1125
         Width           =   4230
      End
      Begin VB.TextBox txt_agente 
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   465
         Width           =   4230
      End
      Begin VB.TextBox txt_ruta 
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1455
         Width           =   4230
      End
      Begin VB.TextBox txt_titular 
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   795
         Width           =   4230
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Pedido:"
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Top             =   1845
         Width           =   540
      End
      Begin VB.Label lbl_origen 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Establecimiento:"
         Height          =   195
         Left            =   180
         TabIndex        =   22
         Top             =   1185
         Width           =   1155
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   525
         Width           =   555
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ruta:"
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Top             =   1515
         Width           =   390
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Titular:"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   855
         Width           =   480
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Detalle "
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   6
         Left            =   30
         TabIndex        =   18
         Top             =   120
         Width           =   5655
      End
   End
   Begin VB.Frame frm_busqueda 
      Height          =   975
      Left            =   345
      TabIndex        =   6
      Top             =   945
      Width           =   3150
      Begin VB.TextBox txt_busqueda_embarque 
         Height          =   315
         Left            =   1290
         TabIndex        =   8
         Top             =   495
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.TextBox txt_busqueda_caja 
         Height          =   315
         Left            =   1290
         TabIndex        =   7
         Top             =   495
         Width           =   1485
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Orden Surtido:"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   555
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Busqueda de Caja"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   2
         Left            =   30
         TabIndex        =   10
         Top             =   120
         Width           =   3075
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Caja:"
         Height          =   195
         Left            =   195
         TabIndex        =   9
         Top             =   555
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   60
      TabIndex        =   5
      Top             =   825
      Width           =   11580
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   1215
      Top             =   15
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
      Left            =   645
      Top             =   15
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   75
      Top             =   15
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
            Picture         =   "frmsalidas_empaques_exportaciones.frx":0A42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_empaques_exportaciones.frx":131C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_empaques_exportaciones.frx":1BF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_empaques_exportaciones.frx":2192
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_empaques_exportaciones.frx":2A6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_empaques_exportaciones.frx":3348
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_empaques_exportaciones.frx":3C22
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_empaques_exportaciones.frx":3D34
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_empaques_exportaciones.frx":3E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_empaques_exportaciones.frx":3F58
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_empaques_exportaciones.frx":406A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_empaques_exportaciones.frx":417C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_empaques_exportaciones.frx":428E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_empaques_exportaciones.frx":4430
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_empaques_exportaciones.frx":5282
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_empaques_exportaciones.frx":5458
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_empaques_exportaciones.frx":556A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_empaques_exportaciones.frx":67EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_empaques_exportaciones.frx":68FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_empaques_exportaciones.frx":7B80
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog comdialog 
      Left            =   -15
      Top             =   15
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Height          =   1200
      Index           =   1
      Left            =   90
      TabIndex        =   24
      Top             =   1800
      Width           =   11430
      Begin VB.TextBox txt_descuento2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   8745
         TabIndex        =   28
         Top             =   750
         Width           =   1155
      End
      Begin VB.TextBox txt_descuento1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6930
         TabIndex        =   27
         Top             =   750
         Width           =   1170
      End
      Begin VB.TextBox txt_cliente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1035
         TabIndex        =   26
         Top             =   750
         Width           =   4590
      End
      Begin VB.TextBox txt_origen 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1035
         TabIndex        =   25
         Top             =   420
         Width           =   4590
      End
      Begin MSComctlLib.Toolbar tool_detalle 
         Height          =   330
         Left            =   10785
         TabIndex        =   29
         Top             =   750
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ver Detalle"
               ImageIndex      =   11
            EndProperty
         EndProperty
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   9945
         TabIndex        =   35
         Top             =   810
         Width           =   120
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   8190
         TabIndex        =   34
         Top             =   810
         Width           =   120
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Descuentos:"
         Height          =   195
         Left            =   6000
         TabIndex        =   33
         Top             =   810
         Width           =   900
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   165
         TabIndex        =   32
         Top             =   810
         Width           =   525
      End
      Begin VB.Label label 
         BackColor       =   &H8000000D&
         Caption         =   " Datos del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   30
         TabIndex        =   31
         Top             =   120
         Width           =   11355
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Almacen:"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   30
         Top             =   480
         Width           =   660
      End
   End
   Begin VB.Frame Frame3 
      Height          =   885
      Index           =   0
      Left            =   90
      TabIndex        =   43
      Top             =   915
      Width           =   5790
      Begin VB.TextBox txt_caja 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   4905
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   405
         Width           =   825
      End
      Begin VB.TextBox txt_archivo 
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
         Left            =   3375
         TabIndex        =   45
         Top             =   405
         Width           =   1080
      End
      Begin VB.TextBox txt_embarque 
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
         Left            =   915
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   405
         Width           =   1125
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   1
         Left            =   30
         TabIndex        =   50
         Top             =   120
         Width           =   5715
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Caja:"
         Height          =   195
         Left            =   4515
         TabIndex        =   49
         Top             =   495
         Width           =   360
      End
      Begin VB.Label lbl_archivo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Orden de Surtido:"
         Height          =   195
         Left            =   2100
         TabIndex        =   48
         Top             =   495
         Width           =   1245
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Embarque:"
         Height          =   195
         Left            =   120
         TabIndex        =   47
         Top             =   488
         Width           =   765
      End
   End
   Begin VB.Frame Frame3 
      Height          =   885
      Index           =   3
      Left            =   5940
      TabIndex        =   40
      Top             =   915
      Width           =   1815
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Cantidad a Surtir"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   4
         Left            =   30
         TabIndex        =   42
         Top             =   120
         Width           =   1740
      End
      Begin VB.Label lbl_enviados 
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
         Height          =   405
         Left            =   120
         TabIndex        =   41
         Top             =   390
         Width           =   1635
      End
   End
   Begin VB.Frame Frame3 
      Height          =   885
      Index           =   4
      Left            =   7815
      TabIndex        =   37
      Top             =   915
      Width           =   1815
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Cantidad en Caja"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   5
         Left            =   30
         TabIndex        =   39
         Top             =   120
         Width           =   1740
      End
      Begin VB.Label lbl_recibidos 
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
         Height          =   345
         Left            =   105
         TabIndex        =   38
         Top             =   420
         Width           =   1500
      End
   End
   Begin VB.Frame Frame3 
      Height          =   885
      Index           =   2
      Left            =   9705
      TabIndex        =   67
      Top             =   915
      Width           =   1815
      Begin VB.Label lbl_empacados 
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
         Height          =   345
         Left            =   105
         TabIndex        =   69
         Top             =   420
         Width           =   1500
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Cantidad Empacada"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   7
         Left            =   30
         TabIndex        =   68
         Top             =   120
         Width           =   1740
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4260
      Left            =   105
      TabIndex        =   52
      Top             =   2985
      Width           =   11415
      Begin VB.TextBox txt_codigo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1575
         TabIndex        =   60
         Top             =   465
         Width           =   6660
      End
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   3960
         TabIndex        =   57
         Top             =   2100
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar 
            Height          =   330
            Left            =   75
            TabIndex        =   58
            Top             =   375
            Width           =   2745
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000D&
            Caption         =   "Código a eliminar"
            ForeColor       =   &H8000000E&
            Height          =   225
            Index           =   3
            Left            =   0
            TabIndex        =   59
            Top             =   15
            Width           =   2895
         End
      End
      Begin VB.TextBox txt_cantidad 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5160
         TabIndex        =   56
         Top             =   495
         Width           =   1890
      End
      Begin VB.Frame frm_refacturacion 
         Height          =   540
         Left            =   8985
         TabIndex        =   53
         Top             =   435
         Visible         =   0   'False
         Width           =   2325
         Begin VB.CommandButton cmd_refacturacion 
            Height          =   345
            Left            =   60
            Picture         =   "frmsalidas_empaques_exportaciones.frx":8E02
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   120
            Width           =   345
         End
         Begin VB.TextBox txt_archivo_refacturar 
            Height          =   315
            Left            =   495
            TabIndex        =   54
            Top             =   135
            Width           =   1485
         End
      End
      Begin MSComctlLib.ListView lv_salidas 
         Height          =   3060
         Left            =   30
         TabIndex        =   61
         Top             =   1110
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   5398
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
            Text            =   "Código"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   7576
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Posibles    "
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Surtidos    "
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Empacado "
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Caja      "
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Faltan      "
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Costo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "precio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Tipo"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código del Artículo:"
         Height          =   195
         Left            =   120
         TabIndex        =   65
         Top             =   615
         Width           =   1395
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Lectura de Artículos"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   0
         Left            =   30
         TabIndex        =   64
         Top             =   120
         Width           =   11340
      End
      Begin VB.Label lbl_cantidad 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   4410
         TabIndex        =   63
         Top             =   615
         Width           =   675
      End
      Begin VB.Label lbl_estatus 
         Caption         =   "cancelada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   7200
         TabIndex        =   62
         Top             =   450
         Width           =   4095
      End
   End
   Begin VB.Label lblnombremovimiento 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   165
      TabIndex        =   70
      Top             =   0
      Width           =   11445
   End
End
Attribute VB_Name = "frmsalidas_empaques_exportaciones_no_usar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_planta As String
Dim var_lote As String
Dim var_codigo As String
Dim var_nodo As Integer
Dim var_cantidad_total_empacada As Double
Dim var_agente_embarque As String
Dim var_estatus_embarque As String
Dim var_numero_caja As Integer
Dim var_origen As String
Dim var_transporto As String
Dim var_tipo_proveedor As String
Dim var_primera_vez As Boolean
Dim var_numero_folio As Double
Dim VAR_TABLA_NOMBRE_ORIGEN As String
Dim VAR_RUTA_TABLA_ORIGEN As String
Dim VAR_CAMPO_CODIGO_ORIGEN As String
Dim VAR_CAMPO_DESCRIPCION_ORIGEN As String
Dim VAR_CAMPO_COSTO_ORIGEN As String
Dim VAR_CAMPO_CANTIDAD_ORIGEN As String
Dim VAR_CAMPO_CANTIDAD_ENTRADA As String
Dim VAR_TABLA_DESTINO As String
Dim VAR_CAMPO_CODIGO_DESTINO As String
Dim VAR_CAMPO_DESCRIPCION_DESTINO As String
Dim VAR_CAMPO_COSTO_DESTINO As String
Dim VAR_CAMPO_CANTIDAD_DESTINO  As String
Dim VAR_CAMPO_NUMERO  As String
Dim var_cantidad_enviada As Double
Dim var_cantidad_recibida As Double
Dim var_articulo_enviado As String
Dim var_costo_enviado As Double
Dim var_almacen_Destino As String
Dim var_almacen_origen As String
Dim var_proveedor As String
Dim var_estatus_movimiento As String
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_modifica As Boolean
Dim var_factura As String
Dim var_cantidad_leida As Double
Dim var_tabla As ADODB.Connection
Dim var_ruta As String
Dim var_folio_enviado As Integer
Dim var_referencia As String
Dim var_suma_cantidad_enviada As Double
Dim var_suma_cantidad_recibida As Double
Dim var_orden_surtido As Double
Dim var_clave_agente As String
Dim var_clave_establecimiento As String
Dim var_clave_titular As String
Dim var_clave_cliente As String
Dim var_clave_ruta As String
Dim var_plazo As Integer
Dim var_descuento_1 As Variant
Dim var_descuento_3 As Variant
Dim var_descuento_2 As Variant
Dim var_iva As Variant
Dim var_agrupador As String
Dim var_correo_electronico As String
Dim var_autorizo_embarque As Boolean
Dim var_renglon As Double
Sub ilumina_grid()
   var_n = lv_salidas.ListItems.Count
   For var_i = 1 To var_n
       If var_i = var_renglon Then
          lv_salidas.ListItems.Item(var_i).Bold = True
          lv_salidas.ListItems.Item(var_i).ListSubItems(1).Bold = True
          lv_salidas.ListItems.Item(var_i).ListSubItems(2).Bold = True
          lv_salidas.ListItems.Item(var_i).ListSubItems(3).Bold = True
          lv_salidas.ListItems.Item(var_i).ListSubItems(4).Bold = True
          lv_salidas.ListItems.Item(var_i).ListSubItems(5).Bold = True
          lv_salidas.ListItems.Item(var_i).ListSubItems(6).Bold = True
          lv_salidas.ListItems.Item(var_i).ListSubItems(7).Bold = True
          lv_salidas.ListItems.Item(var_i).ForeColor = &H8000&
          lv_salidas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H8000&
          lv_salidas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H8000&
          lv_salidas.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H8000&
          lv_salidas.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H8000&
          lv_salidas.ListItems.Item(var_i).ListSubItems(5).ForeColor = &H8000&
          lv_salidas.ListItems.Item(var_i).ListSubItems(6).ForeColor = &H8000&
          lv_salidas.ListItems.Item(var_i).ListSubItems(7).ForeColor = &H8000&
       Else
          If (lv_salidas.ListItems.Item(var_i).ListSubItems(6) * 1) = 0 Then
             lv_salidas.ListItems.Item(var_i).Bold = False
             lv_salidas.ListItems.Item(var_i).ListSubItems(1).Bold = False
             lv_salidas.ListItems.Item(var_i).ListSubItems(2).Bold = False
             lv_salidas.ListItems.Item(var_i).ListSubItems(3).Bold = False
             lv_salidas.ListItems.Item(var_i).ListSubItems(4).Bold = False
             lv_salidas.ListItems.Item(var_i).ListSubItems(5).Bold = False
             lv_salidas.ListItems.Item(var_i).ListSubItems(6).Bold = False
             lv_salidas.ListItems.Item(var_i).ListSubItems(7).Bold = False
             lv_salidas.ListItems.Item(var_i).ForeColor = &HFF&
             lv_salidas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &HFF&
             lv_salidas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &HFF&
             lv_salidas.ListItems.Item(var_i).ListSubItems(3).ForeColor = &HFF&
             lv_salidas.ListItems.Item(var_i).ListSubItems(4).ForeColor = &HFF&
             lv_salidas.ListItems.Item(var_i).ListSubItems(5).ForeColor = &HFF&
             lv_salidas.ListItems.Item(var_i).ListSubItems(6).ForeColor = &HFF&
             lv_salidas.ListItems.Item(var_i).ListSubItems(7).ForeColor = &HFF&
          Else
             lv_salidas.ListItems.Item(var_i).Bold = False
             lv_salidas.ListItems.Item(var_i).ListSubItems(1).Bold = False
             lv_salidas.ListItems.Item(var_i).ListSubItems(2).Bold = False
             lv_salidas.ListItems.Item(var_i).ListSubItems(3).Bold = False
             lv_salidas.ListItems.Item(var_i).ListSubItems(4).Bold = False
             lv_salidas.ListItems.Item(var_i).ListSubItems(5).Bold = False
             lv_salidas.ListItems.Item(var_i).ListSubItems(6).Bold = False
             lv_salidas.ListItems.Item(var_i).ListSubItems(7).Bold = False
             lv_salidas.ListItems.Item(var_i).ForeColor = &H80000012
             lv_salidas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H80000012
             lv_salidas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H80000012
             lv_salidas.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H80000012
             lv_salidas.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H80000012
             lv_salidas.ListItems.Item(var_i).ListSubItems(5).ForeColor = &H80000012
             lv_salidas.ListItems.Item(var_i).ListSubItems(6).ForeColor = &H80000012
             lv_salidas.ListItems.Item(var_i).ListSubItems(7).ForeColor = &H80000012
          End If
       End If
   Next var_i
   If var_renglon > 0 Then
      lv_salidas.ListItems.Item(var_renglon).Selected = True
      lv_salidas.selectedItem.EnsureVisible
   End If
   lv_salidas.Refresh
End Sub



Private Sub cmd_buscar_Click()
            txt_busqueda_caja = ""
            txt_busqueda_caja.Enabled = True
            txt_busqueda_embarque = ""
            txt_busqueda_embarque.Enabled = True
            frm_busqueda.Visible = True
            txt_busqueda_caja.SetFocus
End Sub

Private Sub cmd_cerrar_Click()
   If var_estatus_embarque = "" Then
      var_si = MsgBox("¿Deseas ya cerrar el embarque?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         var_si = MsgBox("Confirmar el cerrado de el embarque", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_estatus_embarque = "E"
            rs.Open "UPDATE TB_ENCABEZADO_EMBARQUEs SET CHAR_EMB_ESTATUS = 'E' WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_EMB_EMBARQUE = " + txt_embarque, cnn, adOpenDynamic, adLockOptimistic
            MsgBox "El embarque a sido cerrado", vbOKOnly, "ATENCION"
         End If
      End If
   Else
      MsgBox "El embarque ya no puede ser cerrado", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
   Set TB_DETALLE_CAJAS_M = New TB_DETALLE_CAJAS_M
   Dim var_referencia_vi As String
   Dim var_contador_renglones As Integer
   Dim var_numero_etiqueta As Integer
   Dim var_longitud As Integer
   Dim var_articulo As String
   Dim var_referencia_caja As String
   Dim var_contador As Integer
   Dim var_cantidad_total As String
   Dim var_cantidad_caja_impresion As Double
   Dim var_posible_cerrar_movimiento As Integer
   var_posible_cerrar_movimiento = 1
   
   Dim dl As Long                                 ' Valor devuelto por la función API
   Dim sAttributes As String                  ' Aributos
   Dim sDriver As String                       ' Nombre del controlador
   Dim sDescription As String                ' Descripción del DSN
   Dim sDsnName As String                  ' Nombre del DSN

   Const ODBC_ADD_SYS_DSN As Long = 4         ' Se creará un DSN de sistema
   Const vbAPINull As Long = 0&                         ' Puntero NULL

   ' se elimina
   Const ODBC_REMOVE_SYS_DSN As Long = 6    ' Se eliminará un DSN de sistema
   sDsnName = "DSN=sqlsistema"
   sDriver = "SQL Server"
   dl = SQLConfigDataSource(0&, ODBC_REMOVE_SYS_DSN, sDriver, sDsnName)

   'se crea
   sDsnName = "sqlsistema"
   sDescription = "sqlsistema"
   sDriver = "SQL Server"
   sAttributes = "DSN=" & sDsnName & Chr(0)
   sAttributes = sAttributes & "Server=" + parametros(0) & Chr$(0)
   sAttributes = sAttributes & "Description=" & sDescription & Chr(0)
   sAttributes = sAttributes & "Database=" + var_bd_movimientos & Chr(0)
   strAttributes = strAttributes & "UID=sa" & Chr$(0)
   strAttributes = strAttributes & "PWD=elia" & Chr$(0)
   dl = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYS_DSN, sDriver, sAttributes)
   
   If var_numero_caja > 0 Then
      var_cantidad_caja_impresion = 0
      If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Or var_estatus_movimiento = "S" Then
         rs.Open "select * from tb_detalle_cajas with (nolock)  where vcha_emp_empresa_id ='" + var_empresa + "' and vcha_uor_unidad_id ='" + var_unidad_organizacional + "' and inte_ors_orden_surtido = " + txt_archivo + " and inte_paq_caja = " + Str(var_numero_caja) + " AND INTE_EMB_EMBARQUE = " + Me.txt_embarque + " and floa_paq_cantidad > 0 ", cnn, adOpenDynamic, adLockOptimistic
         If IsNumeric(Me.lbl_recibidos) Then
            var_cantidad_total = CStr(CInt(Me.lbl_recibidos))
         Else
            var_cantidad_total = ""
         End If
         If Not rs.EOF Then
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set a = fs.CreateTextFile(App.Path + "\etiquetas.txt", True)
            var_numero_caja = rs!inte_paq_caja
            var_referencia_caja = ""
            var_contador = 0
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
               var_referencia_embarque = Trim(Str(txt_embarque))
            End If
            var_numero_etiqueta = 1
            While Not rs.EOF
                  var_articulo = ""
                  If var_numero_etiqueta = 7 Then
                     var_numero_etiqueta = 1
                  End If
                  If var_numero_etiqueta = 1 Then
                     a.writeline ("")
                     a.writeline ("US")
                     a.writeline ("N")
                     a.writeline ("q816")
                     a.writeline ("Q1015,20+0")
                     a.writeline ("S2")
                     a.writeline ("D8")
                     a.writeline ("ZT")
                     a.writeline ("TTh:m")
                     a.writeline ("TDy2.mn.dd")
                  End If
                  rsaux3.Open "select vcha_art_nombre_español from tb_articulos where vcha_art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                  var_longitud = Len(Trim(rsaux3!vcha_art_nombre_español))
                  If var_longitud >= 35 Then
                     var_articulo = Left(Trim(rsaux3!vcha_art_nombre_español), 35) + "  "
                  End If
                  If var_longitud < 35 Then
                     var_articulo = Trim(rsaux3!vcha_art_nombre_español)
                     While Not var_longitud = 38
                           var_articulo = var_articulo + " "
                           var_longitud = var_longitud + 1
                     Wend
                  End If
                  rsaux3.Close
                  var_cantidad_caja_impresion = var_cantidad_caja_impresion + rs!floa_paq_cantidad
                  var_articulo = var_articulo + Trim(Str(rs!floa_paq_cantidad))
                  If var_numero_etiqueta = 1 Then
                     a.writeline ("A782,20,1,4,2,1,N,""" + var_articulo + """")
                  End If
                  If var_numero_etiqueta = 2 Then
                     a.writeline ("A696,20,1,4,2,1,N,""" + var_articulo + """")
                  End If
                  If var_numero_etiqueta = 3 Then
                     a.writeline ("A627,20,1,4,2,1,N,""" + var_articulo + """")
                  End If
                  If var_numero_etiqueta = 4 Then
                     a.writeline ("A554,20,1,4,2,1,N,""" + var_articulo + """")
                  End If
                  If var_numero_etiqueta = 5 Then
                     a.writeline ("A475,20,1,4,2,1,N,""" + var_articulo + """")
                  End If
                  If var_numero_etiqueta = 6 Then
                     a.writeline ("A390,20,1,4,2,1,N,""" + var_articulo + """")
                  End If
                  var_articulo = ""
                  rs.MoveNext
                  If rs.EOF Then
                     var_numero_etiqueta = 6
                  End If
                  If var_numero_etiqueta = 6 Then
                     a.writeline ("A270,20,1,5,1,1,N,""CAJA     :""")
                     a.writeline ("A168,20,1,5,1,1,N,""EMBARQUE :""")
                     a.writeline ("A116,20,1,4,2,1,N,""" + txt_cliente + """")
                     a.writeline ("A282,459,1,5,1,1,N,""" + var_referencia_caja + "/" + CStr(var_cantidad_caja_impresion) + "/" + var_cantidad_total + """")
                     a.writeline ("A187,459,1,5,1,1,N,""" + var_referencia_embarque + """")
                     If var_contador = 0 Then
                        a.writeline ("B77,782,0,3,4,8,101,B,""C" + var_referencia_embarque + var_referencia_caja + """")
                     End If
                     var_contador = var_contador + 1
                     a.writeline ("P1")
                  End If
                  var_numero_etiqueta = var_numero_etiqueta + 1
            Wend
            If var_numero_etiqueta < 6 Then
               While Not var_numero_etiqueta = 7
                     If var_numero_etiqueta = 1 Then
                        a.writeline ("A782,20,1,4,2,1,N,""" + var_articulo + """")
                     End If
                     If var_numero_etiqueta = 2 Then
                        a.writeline ("A696,20,1,4,2,1,N,""" + var_articulo + """")
                     End If
                     If var_numero_etiqueta = 3 Then
                        a.writeline ("A627,20,1,4,2,1,N,""" + var_articulo + """")
                     End If
                     If var_numero_etiqueta = 4 Then
                        a.writeline ("A554,20,1,4,2,1,N,""" + var_articulo + """")
                     End If
                     If var_numero_etiqueta = 5 Then
                        a.writeline ("A475,20,1,4,2,1,N,""" + var_articulo + """")
                     End If
                     If var_numero_etiqueta = 6 Then
                        a.writeline ("A390,20,1,4,2,1,N,""" + var_articulo + """")
                     End If
                     var_articulo = ""
                     If var_numero_etiqueta = 6 Then
                        a.writeline ("A270,20,1,5,1,1,N,""CAJA     :""")
                        a.writeline ("A168,20,1,5,1,1,N,""EMBARQUE :""")
                        a.writeline ("A116,20,1,4,2,1,N,""" + txt_cliente + """")
                        a.writeline ("A282,459,1,5,1,1,N,""" + var_referencia_caja + "/" + CStr(var_cantidad_caja_impresion) + "/" + var_cantidad_total + """")
                        a.writeline ("A187,459,1,5,1,1,N,""" + var_referencia_embarque + """")
                        If var_contador = 0 Then
                           a.writeline ("B77,782,0,3,4,8,101,B,""C" + var_referencia_embarque + var_referencia_caja + """")
                        End If
                        var_contador = var_contador + 1
                        a.writeline ("P1")
                     End If
                     If var_numero_etiqueta = 6 Then
                        'a.writeline ("")
                        'a.writeline ("O")
                        'a.writeline ("q816<")
                        'a.writeline ("Q1015,20+0")
                        'a.writeline ("S2")
                        'a.writeline ("D8")
                        'a.writeline ("ZT")
                        'a.writeline ("TTh: m")
                        'a.writeline ("TDy2.mn.dd")
                     End If
                     var_numero_etiqueta = var_numero_etiqueta + 1
               Wend
            End If
            
            rsaux4.Open "select * from vw_clientes where vcha_cli_clave_id = '" + var_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux4.EOF Then
               rsaux5.Open "select * from vw_establecimientos_2 where vcha_esb_establecimiento_id = '" + var_clave_establecimiento + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux5.EOF Then
                  var_direccion = IIf(IsNull(rsaux5!vcha_esb_domicilio), "", rsaux5!vcha_esb_domicilio)
                  var_colonia = IIf(IsNull(rsaux5!vcha_col_nombre), "", rsaux5!vcha_col_nombre)
                  var_ciudad = IIf(IsNull(rsaux5!vcha_ciu_nombre), "", rsaux5!vcha_ciu_nombre)
                  var_municipio = IIf(IsNull(rsaux5!vcha_mun_nombre), "", rsaux5!vcha_mun_nombre)
                  var_estado = IIf(IsNull(rsaux5!vcha_est_nombre), "", rsaux5!vcha_est_nombre)
                  var_pais = IIf(IsNull(rsaux5!vcha_pai_nombre), "", rsaux5!vcha_pai_nombre)
                  var_cp = IIf(IsNull(rsaux5!vcha_esb_cp), "", rsaux5!vcha_esb_cp)
                  rsaux5.Close
               Else
                  rsaux5.Close
                  var_direccion = IIf(IsNull(rsaux4!VCHA_CLI_DIRECCION), "", rsaux4!VCHA_CLI_DIRECCION)
                  var_colonia = IIf(IsNull(rsaux4!vcha_col_nombre), "", rsaux4!vcha_col_nombre)
                  var_ciudad = IIf(IsNull(rsaux4!vcha_ciu_nombre), "", rsaux4!vcha_ciu_nombre)
                  var_municipio = IIf(IsNull(rsaux4!vcha_mun_nombre), "", rsaux4!vcha_mun_nombre)
                  var_estado = IIf(IsNull(rsaux4!vcha_est_nombre), "", rsaux4!vcha_est_nombre)
                  var_pais = IIf(IsNull(rsaux4!vcha_pai_nombre), "", rsaux4!vcha_pai_nombre)
                  var_cp = IIf(IsNull(rsaux4!VCHA_CLI_CP), "", rsaux4!VCHA_CLI_CP)
               End If
               
               
               a.writeline ("")
               a.writeline ("US")
               a.writeline ("N")
               a.writeline ("q816")
               a.writeline ("Q1015,20+0")
               a.writeline ("S2")
               a.writeline ("D8")
               a.writeline ("ZT")
               a.writeline ("TTh:m")
               a.writeline ("TDy2.mn.dd")
               a.writeline ("A782,20,1,4,2,1,N,""Cliente: " + txt_cliente + """")
               a.writeline ("A696,20,1,4,2,1,N,""Dirección: " + var_direccion + """")
               a.writeline ("A627,20,1,4,2,1,N,""Colonia: " + var_colonia + """")
               a.writeline ("A554,20,1,4,2,1,N,""C.P: " + var_cp + """")
               a.writeline ("A475,20,1,4,2,1,N,""Ciudad: " + var_ciudad + """")
               a.writeline ("A390,20,1,4,2,1,N,""Municipio : " + var_municipio + """")
               a.writeline ("A305,20,1,4,2,1,N,""Estado: " + var_estado + ", " + var_pais + """")
               If var_clave_movimiento = "FT" Then
                  rsaux8.Open "SELECT * FROM VW_PAQUETERIA_IMPRESION_ETIQUETA WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " AND VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux8.EOF Then
                     var_paqueteria = IIf(IsNull(rsaux8!vcha_paq_nombre), "", rsaux8!vcha_paq_nombre)
                     a.writeline ("A220,20,1,4,8,3,N,""" + var_paqueteria + """")
                  End If
                  rsaux8.Close
               End If
               a.writeline ("P1")
            End If
            rsaux4.Close
            a.Close
            Open (App.Path & "\etiquetas.bat") For Output As #2
            var_Archivo = App.Path & "\etiquetas.bat"
            Print #2, "copy " + App.Path + "\etiquetas.txt lpt1"
            Close #2
            x = Shell(var_Archivo, vbHide)
         End If
         rs.Close
      Else
         var_si = MsgBox("¿Desea imprimir las etiquetas de la caja " + Trim(txt_caja) + "?", vbOKCancel, "ATENCION")
         If var_si = 1 Then
            rsaux4.Open "update tb_detalle_cajas set CHAR_PAQ_ESTATUS = 'I' where vcha_emp_empresa_id = '" + var_empresa + "' and INTE_EMB_EMBARQUE = " + txt_embarque + " AND  INTE_PAQ_CAJA = " + CStr(var_numero_caja), cnn, adOpenDynamic, adLockOptimistic
            'var_inserta = TB_DETALLE_CAJAS_M.Anadir(var_orden_surtido, var_numero_caja, var_empresa, var_unidad_organizacional, var_almacen_origen, "I", "", 0, 0)
            var_estatus_movimiento = "I"
            rs.Open "select * from tb_detalle_cajas with (nolock) where vcha_emp_empresa_id ='" + var_empresa + "' and vcha_uor_unidad_id ='" + var_unidad_organizacional + "' and inte_ors_orden_surtido = " + txt_archivo + " and inte_paq_caja = " + Str(var_numero_caja) + " AND INTE_EMB_EMBARQUE = " + Me.txt_embarque + " and floa_paq_cantidad > 0 ", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_contador = 0
               Set fs = CreateObject("Scripting.FileSystemObject")
               Set a = fs.CreateTextFile(App.Path + "\etiquetas.txt", True)
               var_numero_caja = rs!inte_paq_caja
               If IsNumeric(Me.lbl_recibidos) Then
                   var_cantidad_total = CStr(CInt(Me.lbl_recibidos))
               Else
                  var_cantidad_total = ""
               End If
               var_referencia_caja = ""
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
                 var_referencia_embarque = Trim(Str(txt_embarque))
              End If
              var_numero_etiqueta = 1
              While Not rs.EOF
                    var_articulo = ""
                    If var_numero_etiqueta = 7 Then
                       var_numero_etiqueta = 1
                    End If
                    If var_numero_etiqueta = 1 Then
                       a.writeline ("")
                       a.writeline ("US")
                       a.writeline ("N")
                       a.writeline ("q816")
                       a.writeline ("Q1015,20+0")
                       a.writeline ("S2")
                       a.writeline ("D8")
                       a.writeline ("ZT")
                       a.writeline ("TTh: m")
                       a.writeline ("TDy2.mn.dd")
                    End If
                    rsaux3.Open "select vcha_art_nombre_español from tb_articulos where vcha_art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                    var_longitud = Len(Trim(rsaux3!vcha_art_nombre_español))
                    If var_longitud >= 35 Then
                       var_articulo = Left(Trim(rsaux3!vcha_art_nombre_español), 35) + "  "
                    End If
                    If var_longitud < 35 Then
                       var_articulo = Trim(rsaux3!vcha_art_nombre_español)
                       While Not var_longitud = 38
                             var_articulo = var_articulo + " "
                             var_longitud = var_longitud + 1
                       Wend
                    End If
                    rsaux3.Close
                    var_cantidad_caja_impresion = var_cantidad_caja_impresion + rs!floa_paq_cantidad
                    var_articulo = var_articulo + Trim(Str(rs!floa_paq_cantidad))
                    If var_numero_etiqueta = 1 Then
                       a.writeline ("A782,20,1,4,2,1,N,""" + var_articulo + """")
                    End If
                    If var_numero_etiqueta = 2 Then
                       a.writeline ("A696,20,1,4,2,1,N,""" + var_articulo + """")
                    End If
                    If var_numero_etiqueta = 3 Then
                       a.writeline ("A627,20,1,4,2,1,N,""" + var_articulo + """")
                    End If
                    If var_numero_etiqueta = 4 Then
                       a.writeline ("A554,20,1,4,2,1,N,""" + var_articulo + """")
                    End If
                    If var_numero_etiqueta = 5 Then
                       a.writeline ("A475,20,1,4,2,1,N,""" + var_articulo + """")
                    End If
                    If var_numero_etiqueta = 6 Then
                       a.writeline ("A390,20,1,4,2,1,N,""" + var_articulo + """")
                    End If
                    var_articulo = ""
                    rs.MoveNext
                    If rs.EOF Then
                       var_numero_etiqueta = 6
                    End If
                    If var_numero_etiqueta = 6 Then
                       a.writeline ("A270,20,1,5,1,1,N,""CAJA     :""")
                       a.writeline ("A168,20,1,5,1,1,N,""EMBARQUE :""")
                       a.writeline ("A116,20,1,4,2,1,N,""" + txt_cliente + """")
                       a.writeline ("A282,459,1,5,1,1,N,""" + var_referencia_caja + "/" + CStr(var_cantidad_caja_impresion) + "/" + var_cantidad_total + """")
                       a.writeline ("A187,459,1,5,1,1,N,""" + var_referencia_embarque + """")
                       If var_contador = 0 Then
                          a.writeline ("B77,782,0,3,4,8,101,B,""C" + var_referencia_embarque + var_referencia_caja + """")
                       End If
                       var_contador = var_contador + 1
                       a.writeline ("P1")
                    End If
                    var_numero_etiqueta = var_numero_etiqueta + 1
              Wend
              If var_numero_etiqueta < 6 Then
                 While Not var_numero_etiqueta = 7
                       If var_numero_etiqueta = 1 Then
                          a.writeline ("A782,20,1,4,2,1,N,""" + var_articulo + """")
                       End If
                       If var_numero_etiqueta = 2 Then
                          a.writeline ("A696,20,1,4,2,1,N,""" + var_articulo + """")
                       End If
                       If var_numero_etiqueta = 3 Then
                          a.writeline ("A627,20,1,4,2,1,N,""" + var_articulo + """")
                       End If
                       If var_numero_etiqueta = 4 Then
                          a.writeline ("A554,20,1,4,2,1,N,""" + var_articulo + """")
                       End If
                       If var_numero_etiqueta = 5 Then
                          a.writeline ("A475,20,1,4,2,1,N,""" + var_articulo + """")
                       End If
                       If var_numero_etiqueta = 6 Then
                          a.writeline ("A390,20,1,4,2,1,N,""" + var_articulo + """")
                       End If
                       var_articulo = ""
                       If var_numero_etiqueta = 6 Then
                          a.writeline ("A270,20,1,5,1,1,N,""CAJA     :""")
                          a.writeline ("A168,20,1,5,1,1,N,""EMBARQUE :""")
                          a.writeline ("A116,20,1,4,2,1,N,""" + txt_cliente + """")
                          a.writeline ("A282,459,1,5,1,1,N,""" + var_referencia_caja + "/" + CStr(var_cantidad_caja_impresion) + "/" + var_cantidad_total + """")
                          a.writeline ("A187,459,1,5,1,1,N,""" + var_referencia_embarque + """")
                          If var_contador = 0 Then
                             a.writeline ("B77,782,0,3,4,8,101,B,""C" + var_referencia_embarque + var_referencia_caja + """")
                          End If
                          var_contador = var_contador + 1
                          a.writeline ("P1")
                       End If
                       If var_numero_etiqueta = 6 Then
                          'a.writeline ("")
                          'a.writeline ("O")
                          'a.writeline ("q816<")
                          'a.writeline ("Q1015,20+0")
                          'a.writeline ("S2")
                          'a.writeline ("D8")
                          'a.writeline ("ZT")
                          'a.writeline ("TTh: m")
                          'a.writeline ("TDy2.mn.dd")
                       End If
                       var_numero_etiqueta = var_numero_etiqueta + 1
                  Wend
               End If
               rs.Close
               rs.Open "select * from vw_clientes where vcha_cli_clave_id = '" + var_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  rsaux5.Open "select * from vw_establecimientos_2 where vcha_esb_establecimiento_id = '" + var_clave_establecimiento + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux5.EOF Then
                     var_direccion = IIf(IsNull(rsaux5!vcha_esb_domicilio), "", rsaux5!vcha_esb_domicilio)
                     var_colonia = IIf(IsNull(rsaux5!vcha_col_nombre), "", rsaux5!vcha_col_nombre)
                     var_ciudad = IIf(IsNull(rsaux5!vcha_ciu_nombre), "", rsaux5!vcha_ciu_nombre)
                     var_municipio = IIf(IsNull(rsaux5!vcha_mun_nombre), "", rsaux5!vcha_mun_nombre)
                     var_estado = IIf(IsNull(rsaux5!vcha_est_nombre), "", rsaux5!vcha_est_nombre)
                     var_pais = IIf(IsNull(rsaux5!vcha_pai_nombre), "", rsaux5!vcha_pai_nombre)
                     var_cp = IIf(IsNull(rsaux5!vcha_esb_cp), "", rsaux5!vcha_esb_cp)
                     rsaux5.Close
                  Else
                     rsaux5.Close
                     var_direccion = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION)
                     var_colonia = IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
                     var_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                     var_municipio = IIf(IsNull(rs!vcha_mun_nombre), "", rs!vcha_mun_nombre)
                     var_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                     var_pais = IIf(IsNull(rs!vcha_pai_nombre), "", rs!vcha_pai_nombre)
                     var_cp = IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
                  End If
                  
                  
                  a.writeline ("")
                  a.writeline ("US")
                  a.writeline ("N")
                  a.writeline ("q816")
                  a.writeline ("Q1015,20+0")
                  a.writeline ("S2")
                  a.writeline ("D8")
                  a.writeline ("ZT")
                  a.writeline ("TTh:m")
                  a.writeline ("TDy2.mn.dd")
                  a.writeline ("A782,20,1,4,2,1,N,""Cliente: " + txt_cliente + """")
                  a.writeline ("A696,20,1,4,2,1,N,""Dirección: " + var_direccion + """")
                  a.writeline ("A627,20,1,4,2,1,N,""Colonia: " + var_colonia + """")
                  a.writeline ("A554,20,1,4,2,1,N,""C.P: " + var_cp + """")
                  a.writeline ("A475,20,1,4,2,1,N,""Ciudad: " + var_ciudad + """")
                  a.writeline ("A390,20,1,4,2,1,N,""Municipio : " + var_municipio + """")
                  a.writeline ("A305,20,1,4,2,1,N,""Estado: " + var_estado + ", " + var_pais + """")
                  If var_clave_movimiento = "FT" Then
                     rsaux8.Open "SELECT * FROM VW_PAQUETERIA_IMPRESION_ETIQUETA WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " AND VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux8.EOF Then
                        var_paqueteria = IIf(IsNull(rsaux8!vcha_paq_nombre), "", rsaux8!vcha_paq_nombre)
                        a.writeline ("A220,20,1,4,8,3,N,""" + var_paqueteria + """")
                     End If
                     rsaux8.Close
                  End If
                  a.writeline ("P1")
               End If
               rs.Close
               a.Close
               
               Open (App.Path & "\etiquetas.bat") For Output As #2
               var_Archivo = App.Path & "\etiquetas.bat"
               Print #2, "copy " + App.Path + "\etiquetas.txt lpt1"
               Close #2
               x = Shell(var_Archivo, vbHide)
            End If
            txt_codigo.Enabled = False
            txt_foco.Enabled = False
         End If
      End If
   Else
      MsgBox "No se a seleccionado ningún movimiento", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   If Trim(var_estatus_embarque) = "" Then
      lbl_estatus = ""
      lv_salidas.ListItems.Clear
      var_primera_vez = True
      txt_origen = ""
      'txt_archivo = ""
      txt_titular = ""
      txt_agente = ""
      txt_establecimiento = ""
      txt_cliente = ""
      txt_ruta = ""
      txt_pedido = ""
      txt_descuento1 = ""
      txt_descuento2 = ""
      lv_salidas.ListItems.Clear
      'txt_archivo.Enabled = True
      var_cantidad_enviada = 0
      var_cantidad_recibida = 0
      var_cantidad_total_empacada = 0
      var_numero_folio = 0
      var_factura = ""
      txt_factura = ""
      txt_proveedor = ""
      txt_numero = ""
      lbl_recibidos = ""
      lbl_enviados = ""
      lbl_empacados = ""
      txt_folio = ""
      txt_codigo = ""
      var_estatus_movimiento = ""
      txt_codigo.Enabled = False
      txt_foco.Enabled = False
      txt_caja = ""
      'txt_archivo = ""
      'Me.txt_archivo.SetFocus
      Call ejecuta
   Else
      MsgBox "El embarque ya no puede ser modificado", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_refacturacion_Click()
If rsaux5.State = 1 Then
   rsaux5.Close
End If
   rsaux5.Open "select * from tb_detalle_cajas with (nolock)  where inte_emb_embarque = 470 and vcha_emp_empresa_id = '" + var_empresa + "' and inte_ors_orden_surtido = " + Me.txt_archivo_refacturar + " order by inte_paq_caja", cnn, adOpenDynamic, adLockOptimistic
   While Not rsaux5.EOF
   Dim var_actualiza As Boolean
   Dim var_inserta As Boolean
   Dim bandera_suma As Boolean
   Dim var_cantidad As Variant
   Dim var_costo As Double
   Dim var_precio As Double
   Dim var_cantidad_posible As Variant
   Dim var_encontrado As Integer
   Dim var_i As Integer
   Dim var_n As Integer
   Dim var_j As Integer
   Dim var_tipo_pedido As String
   Set TB_DET_ORDEN_SURTIDO_M = New TB_DET_ORDEN_SURTIDO_M
   Set TB_DETALLE_CAJAS_I = New TB_DETALLE_CAJAS_I
   Set TB_DETALLE_CAJA_MOD_CANT = New TB_DETALLE_CAJA_MOD_CANT
   'On Error GoTo salir:
   cnn.CommandTimeout = 360
   Me.txt_codigo = rsaux5!vcha_Art_articulo_id
   var_cantidad_leida = rsaux5!floa_paq_cantidad
   If Trim(txt_codigo.Text) <> "" Then
      var_numero_caja = rsaux5!inte_paq_caja
      bandera_suma = False
      txt_caja = var_numero_caja
      rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux.EOF Then
         rsaux.Close
         Cadena = "select * from tb_det_orden_surtido where inte_ors_orden_surtido = " + txt_archivo_refacturar + " and vcha_art_articulo_id = '" + txt_codigo + "'"
         rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            valor = txt_codigo
            'Set itmfound = lv_salidas.FindItem(valor, lvwText, , lvwPartial)
            'itmfound.EnsureVisible
            'itmfound.Selected = True
             var_n = lv_salidas.ListItems.Count
             var_encontro = 0
             var_i = 1
             While (var_i <= var_n)
                   lv_salidas.ListItems.Item(var_i).Selected = True
                   valor = Trim(lv_salidas.selectedItem)
                   If txt_codigo = valor Then
                      var_cantidad_posible = lv_salidas.selectedItem.SubItems(2)
                      If var_cantidad_posible < (lv_salidas.selectedItem.SubItems(3) * 1) + (lv_salidas.selectedItem.SubItems(4) * 1) + var_cantidad_leida Then
                         var_encontro = 0
                      Else
                         var_encontro = 1
                         var_i = var_n + 1
                      End If
                   End If
                   var_i = var_i + 1
            Wend
            If var_encontro = 1 Then
               var_tipo_pedido = lv_salidas.selectedItem.SubItems(9)
               var_cantidad_posible = lv_salidas.selectedItem.SubItems(2)
               If var_cantidad_posible < (lv_salidas.selectedItem.SubItems(3) + 0) + (lv_salidas.selectedItem.SubItems(4) + var_cantidad_leida) Then
                  txt_codigo = ""
                  frmmensaje.lbl_mensaje = "Cantidad supera a la posible a empaquetar"
                  frmmensaje.Show 1
                  'MsgBox "Cantidad supera a la posible a empaquetar", vbOKOnly, "ATENCION"
               Else
                  lv_salidas.selectedItem.SubItems(3) = lv_salidas.selectedItem.SubItems(3)
                  lv_salidas.selectedItem.SubItems(5) = Format(lv_salidas.selectedItem.SubItems(5) + var_cantidad_leida, "###,###,##0.00")
                  lv_salidas.selectedItem.SubItems(6) = Format(lv_salidas.selectedItem.SubItems(2) - (var_cantidad_leida + lv_salidas.selectedItem.SubItems(3) + lv_salidas.selectedItem.SubItems(4)), "###,###,##0.00")
                  lv_salidas.selectedItem.SubItems(4) = Format(lv_salidas.selectedItem.SubItems(4) + var_cantidad_leida, "###,###,##0.00")
                  var_renglon = lv_salidas.selectedItem.Index
                  Call ilumina_grid
                  var_precio = lv_salidas.selectedItem.SubItems(8)
                  var_costo = lv_salidas.selectedItem.SubItems(7)
                  var_cantidad = lv_salidas.selectedItem.SubItems(4)
                  lbl_recibidos = Format(Int(lbl_recibidos) + var_cantidad_leida, "###,###,##0.00")
                  lbl_empacados = Format(Int(lbl_empacados) + var_cantidad_leida, "###,###,##0.00")
                  var_cantidad_recibida = var_cantidad_recibida + var_cantidad_leida
                  'cnn.BeginTrans
                  'rs.Open "update tb_det_orden_surtido set "
                  'var_actualiza = TB_DET_ORDEN_SURTIDO_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_orden_surtido, txt_codigo, 0, var_cantidad_leida, var_precio, var_tipo_pedido)
                  
                  bandera_suma = True
                  If bandera_suma = True Then
                     rsaux4.Open "select * from tb_detalle_cajas with (nolock)  where inte_ors_orden_surtido = " + txt_archivo + " and inte_paq_caja = " + Str(var_numero_caja) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_Art_articulo_id = '" + txt_codigo + "' AND INTE_EMB_EMBARQUE = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
                     If rsaux4.EOF Then
                        var_inserta = TB_DETALLE_CAJAS_I.Anadir(txt_archivo, var_numero_caja, var_empresa, var_unidad_organizacional, var_almacen_origen, txt_codigo, var_cantidad_leida, "", "", 0, var_costo, var_precio, var_tipo_pedido, CDbl(Me.txt_embarque))
                     Else
                        rsaux2.Open "UPDATE TB_DETALLE_CAJAS SET FLOA_PAQ_CANTIDAD = FLOA_PAQ_CANTIDAD + " + CStr(var_cantidad_leida) + " Where INTE_ORS_ORDEN_SURTIDO   = " + txt_archivo + " AND INTE_PAQ_CAJA = " + CStr(var_numero_caja) + "  and VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_ART_ARTICULO_ID  = '" + lv_salidas.selectedItem + "' AND INTE_EMB_EMBARQUE = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux4.Close
                  End If
                  rsaux4.Open "update tb_det_orden_surtido set floa_ors_cantidad_empacada = floa_ors_cantidad_empacada + " + CStr(var_cantidad_leida) + " where inte_ors_orden_surtido = " + txt_archivo + " and vcha_Art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                  'cnn.CommitTrans
               End If
            Else
               rsaux4.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               VAR_descripcion_no = IIf(IsNull(rsaux4!vcha_art_nombre_español), "", rsaux4!vcha_art_nombre_español)
               rsaux4.Close
               frmmensaje.lbl_articulo = VAR_descripcion_no
               txt_codigo = ""
               frmmensaje.lbl_mensaje = "Cantidad supera a la posible a surtir"
               frmmensaje.Show 1
               'MsgBox "Cantidad supera a la la posible a surtir", vbOKOnly, "ATENCION"
            End If
         Else
            rsaux4.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            VAR_descripcion_no = IIf(IsNull(rsaux4!vcha_art_nombre_español), "", rsaux4!vcha_art_nombre_español)
            rsaux4.Close
            frmmensaje.lbl_articulo = VAR_descripcion_no
            txt_codigo = ""
            frmmensaje.lbl_mensaje = "El artículo no se encuentra dentro de la Orden de Surtido"
            frmmensaje.Show 1
            'MsgBox "El artículo no se encuentra dentro de la Orden de Surtido", vbOKOnly, "ATENCION"
            bandera_suma = False
         End If
         rs.Close
      Else
         frmmensaje.lbl_mensaje = "El artículo no existe"
         frmmensaje.Show 1
         'MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
         rsaux.Close
      End If
         End If
         rsaux5.MoveNext
   Wend
End Sub

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = 116 Then
      frmexisten_rapidas.Show 1
   End If
   If Shift = 1 And KeyCode = 117 Then
      Set reporte = appl.OpenReport(App.Path + "\rep_PROGRESO_EQUIPOS.rpt")
      reporte.RecordSelectionFormula = "{VW_PROGRESO_EQUIPOS.DTIM_EQU_FECHA} = CURRENTDATE"
      frmvistasprevias.cr.ReportSource = reporte
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), "chiquiblancos_sid", parametros(4), parametros(5)
      Next ntablas
      frmvistasprevias.cr.ViewReport
      frmvistasprevias.Caption = "Reporte de Progreso de Surtido"
      frmvistasprevias.Show 1
      Set reporte = Nothing
   End If
   If Shift = 4 And KeyCode = 78 Then
      cmd_nuevo_Click
   End If
   If Shift = 4 And KeyCode = 66 Then
      cmd_buscar_Click
   End If
   If Shift = 4 And KeyCode = 73 Then
      cmd_imprimir_Click
   End If
   If Shift = 4 And KeyCode = 67 Then
      
   End If
   If Shift = 4 And KeyCode = 77 Then
   End If
End Sub

Private Sub Form_Load()
   var_posible_kanban = 0
   lbl_estatus = ""
   var_clave_moviento = var_clave_movimiento
   var_estatus_embarque = ""
   var_cadena_seguridad = ""
   Top = 0
   Left = 0
   var_autorizo_embarque = False
   var_iva = 0
   var_agrupador = ""
   var_cantidad_leida = 1#
   var_estatus_movimiento = ""
   var_almacen_Destino = ""
   var_almacen_origen = ""
   var_proveedor = ""
   var_factura = ""
   var_correo_electronico = "2"
   frm_eliminar.Visible = False
   var_modifica = False
   txt_cantidad.Visible = False
   lbl_cantidad.Visible = False
   txt_codigo.Enabled = False
   var_primera_vez = True
   frm_busqueda.Visible = False
   Set var_tabla = CreateObject("ADODB.connection")
   var_suma_cantidad_enviada = 0
   var_suma_cantidad_recibida = 0
   frm_detalle.Visible = False
   txt_archivo = var_numero_embarque_paquete
   txt_embarque = var_numero_embarque
   rs.Open "SELECT * FROM TB_ENCABEZADO_EMBARQUES WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND INTE_EMB_EMBARQUE = " + txt_embarque, cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_estatus_embarque = Trim(IIf(IsNull(rs!CHAR_EMB_ESTATUS), "", rs!CHAR_EMB_ESTATUS))
      var_agente_embarque = rs!VCHA_AGE_AGENTE_ID
   End If
   If var_estatus_embarque = "" Then
      Me.cmd_nuevo.Enabled = True
   Else
      MsgBox "El embarque ya fue cerrado", vbOKOnly, "ATENCION"
      Me.cmd_nuevo.Enabled = False
   End If
   rs.Close
   'var_clave_movimiento = Me.txt_clave_movimiento
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_salidas_empaques)
End Sub

Private Sub lv_salidas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_salidas, ColumnHeader)
End Sub

Private Sub lv_salidas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
         MsgBox "Imposible modificar el movimiento", vbOKOnly, "ATENCION"
      Else
         frm_eliminar.Visible = True
         txt_cantidad_eliminar.SetFocus
      End If
   End If
End Sub

Private Sub tool_detalle_ButtonClick(ByVal Button As MSComctlLib.Button)
   frm_detalle.Visible = True
   txt_agente.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
End Sub

Private Sub txt_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      frm_detalle.Visible = False
   End If
End Sub

Private Sub txt_agente_LostFocus()
   frm_detalle.Visible = False
End Sub

Private Sub txt_archivo_KeyPress(KeyAscii As Integer)
   Dim var_clave_movimiento_tem As String
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      ejecuta
   End If
End Sub

Private Sub txt_busqueda_caja_KeyPress(KeyAscii As Integer)
Dim var_busqueda_folio As Integer
Dim var_busqueda_movimiento As String
Dim var_busqueda_numero As Integer
Dim var_busqueda_referencia As String
Dim var_posible As Boolean
Dim var_falta As Double
Dim var_surtir As Double
Dim var_surtido As Double
Dim var_empacada As Double
Dim var_encontro As Boolean
Dim var_encontrado As Integer
Dim var_i As Integer
Dim var_n As Integer
Dim var_j As Integer
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 27 Then
      Me.frm_busqueda.Visible = False
   End If
   If KeyAscii = 13 Then
      var_posible = False
      If Trim(txt_busqueda_caja) <> "" Then
         rs.Open "select * from TB_DETALLE_CAJAS with (nolock)  where VCHA_EMP_EMPRESA_ID ='" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and INTE_EMB_EMBARQUE = " + txt_embarque + " and inte_paq_caja = " + txt_busqueda_caja, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            txt_busqueda_embarque = CStr(IIf(IsNull(rs!INTE_ORS_ORDEN_SURTIDO), 0, rs!INTE_ORS_ORDEN_SURTIDO))
            var_estatus_movimiento = IIf(IsNull(rs!char_paq_estatus), "", rs!char_paq_estatus)
            If Not rs.EOF Then
               var_posible = True
            Else
               var_posible = False
            End If
            rs.Close
            If var_posible = True Then
               
               rs.Open "select * from vw_orden_surtido where  inte_ors_orden_surtido = " + txt_busqueda_embarque + " and floa_ors_cantidad_surtir > 0", cnn, adOpenDynamic, adLockOptimistic
               var_orden_surtido = rs!INTE_ORS_ORDEN_SURTIDO
               var_numero_caja = txt_busqueda_caja
               txt_caja = txt_busqueda_caja
               var_primera_vez = False
               txt_archivo = var_orden_surtido
               var_suma_cantidad_enviada = 0
               var_suma_cantidad_recibida = 0
               lbl_enviados.Caption = Format("0", "###,###,##0.00")
               lbl_recibidos.Caption = Format("0", "###,###,##0.00")
               lbl_empacados.Caption = Format("0", "###,###,##0.00")
               lv_salidas.ListItems.Clear
               If IsNull(rs!VCHA_ALM_NOMBRE) Then
                  GoTo no_almacen:
               Else
                  var_almacen_origen = rs!VCHA_ALM_ALMACEN_ID
                  txt_origen = rs!VCHA_ALM_NOMBRE
               End If
               If IsNull(rs!VCHA_TIT_NOMBRE) Then
                  GoTo no_titular:
               Else
                  txt_titular = rs!VCHA_TIT_NOMBRE
                  var_clave_titular = rs!vcha_tit_titular_id
               End If
               If IsNull(rs!inte_ped_dias_condiciones) Then
                  var_plazo = 0
               Else
                  var_plazo = rs!inte_ped_dias_condiciones
               End If
               If IsNull(rs!VCHA_ESB_NOMBRE) Then
                  GoTo no_establecimiento:
               Else
                  txt_establecimiento = rs!VCHA_ESB_NOMBRE
                  var_clave_establecimiento = rs!vcha_ESB_ESTABLECIMIENTO_id
               End If
               If IsNull(rs!VCHA_AGE_NOMBRE) Then
                  GoTo no_agente:
               Else
                  txt_agente = rs!VCHA_AGE_NOMBRE
                  var_clave_agente = rs!VCHA_AGE_AGENTE_ID
               End If
               If IsNull(rs!VCHA_CLI_NOMBRE) Then
                  GoTo no_cliente:
               Else
                  txt_cliente = rs!VCHA_CLI_NOMBRE
                  var_clave_cliente = rs!vcha_cli_clave_id
               End If
               If IsNull(rs!vcha_rut_nombre) Then
                  txt_ruta = ""
                  var_clave_ruta = ""
               Else
                  txt_ruta = rs!vcha_rut_nombre
                  var_clave_ruta = rs!VCHA_RUT_RUTA_ID
               End If
               If IsNull(rs!inte_ped_numero) Then
                  GoTo no_Pedido:
               Else
                  txt_pedido = rs!inte_ped_numero
               End If
               If IsNull(rs!FLOA_ORS_DESCUENTO_1) Then
                  txt_descuento1 = 0
                  var_descuento_1 = 0
               Else
                  txt_descuento1 = rs!FLOA_ORS_DESCUENTO_1
                  var_descuento_1 = rs!FLOA_ORS_DESCUENTO_1
               End If
               If IsNull(rs!FLOA_ORS_DESCUENTO_2) Then
                  txt_descuento2 = 0
                  var_descuento_2 = 0
               Else
                  txt_descuento2 = rs!FLOA_ORS_DESCUENTO_2
                  var_descuento_2 = rs!FLOA_ORS_DESCUENTO_2
               End If
               var_descuento_3 = 0
               var_cantidad_total_empacada = 0
               While Not rs.EOF
                     Set list_item = lv_salidas.ListItems.Add(, , rs!vcha_Art_articulo_id)
                     list_item.SubItems(1) = IIf(IsNull(rs!vcha_art_nombre_español), "", Trim(rs!vcha_art_nombre_español))
                     var_surtir = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), 0, rs!FLOA_ORS_CANTIDAD_SURTIR)
                     list_item.SubItems(2) = Format(IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), 0, rs!FLOA_ORS_CANTIDAD_SURTIR), "###,###,##0.00")
                     var_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
                     list_item.SubItems(3) = Format(IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA), "###,###,##0.00")
                     var_empacada = IIf(IsNull(rs!floa_ors_Cantidad_empacada), 0, rs!floa_ors_Cantidad_empacada)
                     list_item.SubItems(4) = Format(IIf(IsNull(rs!floa_ors_Cantidad_empacada), 0, rs!floa_ors_Cantidad_empacada), "###,###,##0.00")
                     list_item.SubItems(5) = Format(0, "###,###,##0.00")
                     var_falta = (var_empacada + var_surtida)
                     list_item.SubItems(6) = Format(var_surtir - var_falta, "###,###,##0.00")
                     list_item.SubItems(7) = IIf(IsNull(rs!floa_ors_costo), "", rs!floa_ors_costo)
                     list_item.SubItems(8) = IIf(IsNull(rs!floa_ors_precio), "", rs!floa_ors_precio)
                     list_item.SubItems(9) = IIf(IsNull(rs!char_ped_tipo), "", rs!char_ped_tipo)
                     var_suma_cantidad_enviada = var_suma_cantidad_enviada + rs!FLOA_ORS_CANTIDAD_SURTIR
                     var_cantidad_total_empacada = var_cantidad_total_empacada + IIf(IsNull(rs!floa_ors_Cantidad_empacada), 0, rs!floa_ors_Cantidad_empacada)
                  rs.MoveNext:
               Wend
               var_numero_renglones = lv_salidas.Height / 312.5
               var_n = lv_salidas.ListItems.Count
               If var_n > var_numero_renglones Then
                  lv_salidas.ColumnHeaders(2).Width = 4045.05
               Else
                  lv_salidas.ColumnHeaders(2).Width = 4295.05
               End If
               
               rs.Close
               rs.Open "select * from TB_DETALLE_CAJAS with (nolock)  where VCHA_EMP_EMPRESA_ID ='" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and inte_emb_embarque = " + Me.txt_embarque + " and inte_paq_caja = " + txt_busqueda_caja, cnn, adOpenDynamic, adLockOptimistic
               While Not rs.EOF
                   var_n = lv_salidas.ListItems.Count
                   var_encontro = 0
                   var_i = 1
                   var_tipo_pedido = rs!char_ped_tipo
                   While (var_i <= var_n)
                         lv_salidas.ListItems.Item(var_i).Selected = True
                         valor = rs!vcha_Art_articulo_id
                         var_precio = rs!floa_paq_precio
                         var_tipo_pedido = rs!char_ped_tipo
                         If lv_salidas.selectedItem.SubItems(8) * 1 = var_precio And lv_salidas.selectedItem = valor And var_tipo_pedido = lv_salidas.selectedItem.SubItems(9) Then
                            var_encontro = 1
                            var_i = var_n + 1
                         End If
                         var_i = var_i + 1
                   Wend
                   lv_salidas.selectedItem.SubItems(5) = Format(IIf(IsNull(rs!floa_paq_cantidad), 0, rs!floa_paq_cantidad), "###,###,##0.00")
                   var_suma_cantidad_recibida = var_suma_cantidad_recibida + IIf(IsNull(rs!floa_paq_cantidad), 0, rs!floa_paq_cantidad)
                   rs.MoveNext
               Wend
               rs.Close
               lbl_enviados = Format(Str(var_suma_cantidad_enviada), "###,###,##0.00")
               lbl_recibidos = Format(Str(var_suma_cantidad_recibida), "###,###,##0.00")
               lbl_empacados = Format(Str(var_cantidad_total_empacada), "###,###,##0.00")
               txt_archivo.Enabled = False
               frm_busqueda.Visible = False
               If Trim(var_estatus_movimiento) = "S" Then
                  lbl_estatus = "SURTIDA"
               End If
               If Trim(var_estatus_movimiento) = "C" Then
                  lbl_estatus = "CANCELADA"
               End If
               If Trim(var_estatus_movimiento) = "I" Then
                  lbl_estatus = "IMPRESA"
               End If
               If Trim(var_estatus_movimiento) = "" Then
                  lbl_estatus = ""
               End If
               If var_estatus_embarque = "" Then
                  If Trim(var_estatus_movimiento) = "" Then
                     txt_codigo.Enabled = True
                  Else
                     txt_codigo.Enabled = False
                  End If
                  If Me.txt_codigo.Enabled = True Then
                     Me.txt_codigo.SetFocus
                  End If
               Else
                  MsgBox "El embarque ya no puede ser modificado", vbOKOnly, "ATENCION"
                  Me.txt_codigo.Enabled = False
               End If
            Else
               MsgBox "La caja no existe", vbOKOnly, "ATENCION"
            End If
         Else
            rs.Close
            MsgBox "La caja no existe", vbOKOnly, "ATENCION"
         End If
      End If
   End If
   If KeyAscii = 27 Then
      frm_busqueda.Visible = False
   End If
   Exit Sub
no_almacen:
   MsgBox "Almacen Incorrecto", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   Exit Sub
no_almacen_empaque:
   MsgBox "Almacen de Empaque Incorrecto", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   Exit Sub
no_Pedido:
   MsgBox "Pedido Incorrecto", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   Exit Sub
no_establecimiento:
   MsgBox "Establecimiento Incorrecto", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   Exit Sub
no_agente:
   MsgBox "Agente Incorrecto", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   Exit Sub
no_cliente:
   MsgBox "Cliente Incorrecto", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   Exit Sub
no_ruta:
   MsgBox "Ruta Incorrecta", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   Exit Sub
no_titular:
   MsgBox "Titular incorrecto", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   Exit Sub
no_almacen_agente:
   MsgBox "No existe un almacen relacionado a este agente", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   Exit Sub
End Sub

Private Sub txt_busqueda_embarque_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      rs.Open "select * from tb_enc_orden_surtido where inte_ors_orden_surtido = " + txt_busqueda_embarque, cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         rs.Close
         txt_busqueda_caja.Enabled = True
         txt_busqueda_caja.SetFocus
      Else
         MsgBox "No existe la orden de surtido", vbOKOnly, "ATENCION"
         rs.Close
      End If
   End If
   If KeyAscii = 27 Then
      If rs.State = 1 Then
         rs.Close
      End If
      frm_busqueda.Visible = False
   End If
End Sub

Private Sub txt_busqueda_embarque_LostFocus()
'      frm_busqueda.Visible = False
End Sub

Private Sub txt_cantidad_eliminar_GotFocus()
   txt_cantidad_eliminar = ""
End Sub

Private Sub txt_cantidad_eliminar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim var_precio As Variant
      Set TB_DET_ORDEN_SURTIDO_M = New TB_DET_ORDEN_SURTIDO_M
      Set TB_DETALLE_CAJA_MOD_CANT = New TB_DETALLE_CAJA_MOD_CANT
      var_cantidad_eliminar = 1
      var_cantidad_eliminar_arch = lv_salidas.selectedItem.SubItems(5) * 1
      If var_cantidad_eliminar_arch >= var_cantidad_eliminar Then
         If rsaux10.State = 1 Then
            rsaux10.Close
         End If
         rsaux10.Open "select * from tb_trazabilidad_codigos where VCHA_eMP_EMPRESA_ID = '" + var_empresa + "' AND inte_emb_embarque = " + Me.txt_embarque + " and inte_ors_orden_surtido = " + Me.txt_archivo + " and vcha_tra_codigo = '" + Me.txt_cantidad_eliminar + "' AND INTE_PAQ_CAJA = " + CStr(var_numero_caja), cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux10.EOF Then
            var_precio = lv_salidas.selectedItem.SubItems(8)
            var_tipo_pedido = lv_salidas.selectedItem.SubItems(9)
            
            If rsaux4.State = 1 Then
               rsaux4.Close
            End If
            If rsaux5.State = 1 Then
               rsaux5.Close
            End If
            rsaux5.Open "UPDATE TB_TRAZABILIDAD_CODIGOS SET FLOA_TRA_CANTIDAD = ISNULL(FLOA_TRA_CANTIDAD,0) - 1 WHERE  VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND inte_emb_embarque = " + Me.txt_embarque + " and inte_ors_orden_surtido = " + Me.txt_archivo + " and vcha_tra_codigo = '" + Me.txt_cantidad_eliminar + "' AND INTE_PAQ_CAJA = " + CStr(var_numero_caja), cnn, adOpenDynamic, adLockOptimistic
            rsaux5.Open "update TB_DETALLE_EQUIPOS_ORDEN_SURTIDO set FLOA_ORS_CANTIDAD_SURTIDA = isnull(FLOA_ORS_CANTIDAD_SURTIDA,0) - 1 where INTE_ORS_ORDEN_SURTIDO = " + CStr(var_orden_surtido), cnn, adOpenDynamic, adLockOptimistic
            rsaux4.Open "update tb_det_orden_surtido set FLOA_ORS_CANTIDAD_EMPACADA = FLOA_ORS_CANTIDAD_EMPACADA - 1 where INTE_ORS_ORDEN_SURTIDO = " + CStr(var_orden_surtido) + " and VCHA_ART_ARTICULO_ID = '" + Trim(lv_salidas.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
            lv_salidas.selectedItem.SubItems(4) = Format(lv_salidas.selectedItem.SubItems(4) - 1, "###,###,##0.00")
            lv_salidas.selectedItem.SubItems(5) = Format(lv_salidas.selectedItem.SubItems(5) - 1, "###,###,##0.00")
            lv_salidas.selectedItem.SubItems(6) = Format(lv_salidas.selectedItem.SubItems(6) + 1, "###,###,##0.00")
            var_renglon = lv_salidas.selectedItem.Index
            Call ilumina_grid
            var_precio = lv_salidas.selectedItem.SubItems(8)
            var_inserta = False
            rsaux2.Open "UPDATE TB_DETALLE_CAJAS SET FLOA_PAQ_CANTIDAD = FLOA_PAQ_CANTIDAD - 1 Where inte_emb_embarque = " + txt_embarque + " AND INTE_PAQ_CAJA = " + CStr(var_numero_caja) + "  and VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND  VCHA_ALM_ALMACEN_ID   = '" + var_almacen_origen + "' AND VCHA_ART_ARTICULO_ID  = '" + lv_salidas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
            lbl_recibidos = Format(Int(lbl_recibidos) - 1, "###,###,##0.00")
            lbl_empacados = Format(Int(lbl_empacados) - 1, "###,###,##0.00")
            frm_eliminar.Visible = False
            If rsaux4.State = 1 Then
               rsaux4.Close
            End If
            If rsaux2.State = 1 Then
               rsaux2.Close
            End If
         Else
            MsgBox "El artículo no se encuentra dentro del embarque", vbOKOnly, "ATENCION"
         End If
         rsaux10.Close
         var_planta = ""
         var_lote = ""
         var_codigo = ""
         var_nodo = 0
         txt_codigo.SetFocus
      Else
         MsgBox "No esposible eliminar esta cantidad", vbOKOnly, "ATENCION"
      End If
   End If
   If KeyAscii = 27 Then
      frm_eliminar.Visible = False
      var_planta = ""
      var_lote = ""
      var_codigo = ""
      var_nodo = 0
      txt_codigo.SetFocus
   End If
End Sub

Private Sub txt_cantidad_eliminar_LostFocus()
   frm_eliminar.Visible = False
   txt_codigo.SetFocus
End Sub

Private Sub txt_cantidad_GotFocus()
   txt_cantidad = "1"
End Sub

Private Sub txt_cantidad_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If Trim(txt_cantidad) <> "" Then
         var_cantidad_leida = txt_cantidad
         txt_foco.Enabled = True
         txt_foco.SetFocus
         lbl_cantidad.Visible = False
         txt_cantidad.Visible = False
      End If
   End If
End Sub

Private Sub txt_codigo_GotFocus()
   var_planta = ""
   var_lote = ""
   var_codigo = ""
   var_nodo = 0
   txt_codigo = ""
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   Dim var_recontable As Integer
   Dim var_caja As String
   Dim var_cantidad_caja As Integer
   Dim var_si_exporta As Integer
   txt_codigo = Trim(txt_codigo)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Trim(Me.txt_codigo) <> "" Then
         var_si_exporta = 0
         If Mid(Me.txt_codigo, 1, 1) = "<" Then
            var_si_planta = 1
            var_si_lote = 0
            var_si_codigo = 0
            var_planta = ""
            var_lote = ""
            var_codigo = ""
            var_nodo = 0
            For var_j = 1 To Len(Me.txt_codigo)
                If var_si_planta = 1 Then
                   If Mid(Me.txt_codigo, var_j, 1) <> "<" Then
                      If Mid(Me.txt_codigo, var_j, 1) <> ">" Then
                         var_planta = var_planta + Mid(Me.txt_codigo, var_j, 1)
                      End If
                   End If
                End If
                If var_si_lote = 1 Then
                   If Mid(Me.txt_codigo, var_j, 1) <> "[" Then
                      If Mid(Me.txt_codigo, var_j, 1) <> "]" Then
                         var_lote = var_lote + Mid(Me.txt_codigo, var_j, 1)
                      End If
                   End If
                End If
                If var_si_codigo = 1 Then
                   If Mid(Me.txt_codigo, var_j, 1) <> "{" Then
                      If Mid(Me.txt_codigo, var_j, 1) <> "}" Then
                         var_codigo = var_codigo + Mid(Me.txt_codigo, var_j, 1)
                      End If
                   End If
                End If
                If Mid(Me.txt_codigo, var_j, 1) = ">" Then
                   var_si_planta = 0
                   var_si_lote = 1
                   var_si_codigo = 0
                End If
                If Mid(Me.txt_codigo, var_j, 1) = "]" Then
                   var_si_planta = 0
                   var_si_lote = 0
                   var_si_codigo = 1
                End If
            Next var_j
            'MsgBox var_planta
            'MsgBox var_lote
            'MsgBox var_codigo
            rs.Open "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE VCHA_UOR_UNIDAD_ID = '" + var_planta + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               If IsNumeric(var_lote) Then
                  rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     rsaux1.Open "select * from TB_TRAZABILIDAD_LOTES where vcha_tra_planta = '" + var_planta + "' and inte_tra_lote = " + CStr(var_lote) + " and vcha_Art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux1.EOF Then
                        var_si_exporta = IIf(IsNull(rsaux1!inte_tra_permite_exportar), 0, rsaux1!inte_tra_permite_exportar)
                        var_nodo = IIf(IsNull(rsaux1!inte_tra_nodo_identificador), 0, rsaux1!inte_tra_nodo_identificador)
                     Else
                        txt_codigo = ""
                        frmmensaje.lbl_mensaje = "Código de artículo incorrecto"
                        frmmensaje.Show 1
                     End If
                     rsaux1.Close
                  Else
                     txt_codigo = ""
                     frmmensaje.lbl_mensaje = "Código de artículo incorrecto"
                     frmmensaje.Show 1
                  End If
                  rsaux.Close
               Else
                  txt_codigo = ""
                  frmmensaje.lbl_mensaje = "Número de lote incorrecto"
                  frmmensaje.Show 1
               End If
            Else
               txt_codigo = ""
               frmmensaje.lbl_mensaje = "Clave de planta incorrecta"
               frmmensaje.Show 1
            End If
            rs.Close
         Else
            txt_codigo = ""
            frmmensaje.lbl_mensaje = "Código incorrecto"
            frmmensaje.Show 1
         End If
         If var_si_exporta = 1 Then
            Me.txt_foco.Enabled = True
            Me.txt_foco.SetFocus
         End If
         If var_si_exporta = 0 Then
            If var_codigo <> "" Then
               rsaux4.Open "select * from tb_articulos where vcha_art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               VAR_descripcion_no = IIf(IsNull(rsaux4!vcha_art_nombre_español), "", rsaux4!vcha_art_nombre_español)
               rsaux4.Close
               frmmensaje.lbl_articulo = VAR_descripcion_no
               var_codigo = ""
               frmmensaje.lbl_mensaje = "El artículo no es exportable"
               frmmensaje.Show 1
            End If
         End If
      End If
   End If
End Sub




Private Sub txt_foco_GotFocus()
   Dim var_actualiza As Boolean
   Dim var_inserta As Boolean
   Dim bandera_suma As Boolean
   Dim var_cantidad As Variant
   Dim var_costo As Double
   Dim var_precio As Double
   Dim var_cantidad_posible As Variant
   Dim var_encontrado As Integer
   Dim var_i As Integer
   Dim var_n As Integer
   Dim var_j As Integer
   Dim var_tipo_pedido As String
   Set TB_DET_ORDEN_SURTIDO_M = New TB_DET_ORDEN_SURTIDO_M
   Set TB_DETALLE_CAJAS_I = New TB_DETALLE_CAJAS_I
   Set TB_DETALLE_CAJA_MOD_CANT = New TB_DETALLE_CAJA_MOD_CANT
   'On Error GoTo salir:
   cnn.CommandTimeout = 360
   If Trim(Me.txt_codigo) <> "" Then
      If var_primera_vez = True Then
         rs.Open "select maximo_caja from vw_maximo_caja where inte_emb_embarque = " + txt_embarque + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_numero_caja = rs(0).Value + 1
         Else
            var_numero_caja = 1
         End If
         var_primera_vez = False
         rs.Close
      End If
      bandera_suma = False
      txt_caja = var_numero_caja
      rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux.EOF Then
         rsaux.Close
         Cadena = "select * from tb_det_orden_surtido where inte_ors_orden_surtido = " + txt_archivo + " and vcha_art_articulo_id = '" + var_codigo + "'"
         rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            valor = var_codigo
            'Set itmfound = lv_salidas.FindItem(valor, lvwText, , lvwPartial)
            'itmfound.EnsureVisible
            'itmfound.Selected = True
             var_n = lv_salidas.ListItems.Count
             var_encontro = 0
             var_i = 1
             While (var_i <= var_n)
                   lv_salidas.ListItems.Item(var_i).Selected = True
                   valor = Trim(lv_salidas.selectedItem)
                   If var_codigo = valor Then
                      var_cantidad_posible = lv_salidas.selectedItem.SubItems(2)
                      If var_cantidad_posible < (lv_salidas.selectedItem.SubItems(3) * 1) + (lv_salidas.selectedItem.SubItems(4) * 1) + var_cantidad_leida Then
                         var_encontro = 0
                      Else
                         var_encontro = 1
                         var_i = var_n + 1
                      End If
                   End If
                   var_i = var_i + 1
            Wend
            If var_encontro = 1 Then
               var_tipo_pedido = lv_salidas.selectedItem.SubItems(9)
               var_cantidad_posible = lv_salidas.selectedItem.SubItems(2)
               If var_cantidad_posible < (lv_salidas.selectedItem.SubItems(3) + 0) + (lv_salidas.selectedItem.SubItems(4) + var_cantidad_leida) Then
                  var_codigo = ""
                  frmmensaje.lbl_mensaje = "Cantidad supera a la posible a empaquetar"
                  frmmensaje.Show 1
                  'MsgBox "Cantidad supera a la posible a empaquetar", vbOKOnly, "ATENCION"
               Else
                  lv_salidas.selectedItem.SubItems(3) = lv_salidas.selectedItem.SubItems(3)
                  lv_salidas.selectedItem.SubItems(5) = Format(lv_salidas.selectedItem.SubItems(5) + var_cantidad_leida, "###,###,##0.00")
                  lv_salidas.selectedItem.SubItems(6) = Format(lv_salidas.selectedItem.SubItems(2) - (var_cantidad_leida + lv_salidas.selectedItem.SubItems(3) + lv_salidas.selectedItem.SubItems(4)), "###,###,##0.00")
                  lv_salidas.selectedItem.SubItems(4) = Format(lv_salidas.selectedItem.SubItems(4) + var_cantidad_leida, "###,###,##0.00")
                  var_renglon = lv_salidas.selectedItem.Index
                  Call ilumina_grid
                  var_precio = lv_salidas.selectedItem.SubItems(8)
                  var_costo = lv_salidas.selectedItem.SubItems(7)
                  var_cantidad = lv_salidas.selectedItem.SubItems(4)
                  lbl_recibidos = Format(Int(lbl_recibidos) + var_cantidad_leida, "###,###,##0.00")
                  lbl_empacados = Format(Int(lbl_empacados) + var_cantidad_leida, "###,###,##0.00")
                  var_cantidad_recibida = var_cantidad_recibida + var_cantidad_leida
                  'cnn.BeginTrans
                  'rs.Open "update tb_det_orden_surtido set "
                  'var_actualiza = TB_DET_ORDEN_SURTIDO_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_orden_surtido, var_codigo, 0, var_cantidad_leida, var_precio, var_tipo_pedido)
                  
                  bandera_suma = True
                  If bandera_suma = True Then
                     If rsaux4.State = 1 Then
                        rsaux4.Close
                     End If
                     rsaux4.Open "select * from tb_detalle_cajas with (nolock) where inte_ors_orden_surtido = " + txt_archivo + " and inte_paq_caja = " + Str(var_numero_caja) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_Art_articulo_id = '" + var_codigo + "' AND INTE_EMB_EMBARQUE = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
                     rsaux5.Open "update TB_DETALLE_EQUIPOS_ORDEN_SURTIDO set FLOA_ORS_CANTIDAD_SURTIDA = isnull(FLOA_ORS_CANTIDAD_SURTIDA,0) + " + CStr(var_cantidad_leida) + " where INTE_ORS_ORDEN_SURTIDO = " + CStr(rs!INTE_ORS_ORDEN_SURTIDO), cnn, adOpenDynamic, adLockOptimistic
                     If rsaux4.EOF Then
                        var_inserta = TB_DETALLE_CAJAS_I.Anadir(txt_archivo, var_numero_caja, var_empresa, var_unidad_organizacional, var_almacen_origen, var_codigo, var_cantidad_leida, "", "", 0, var_costo, var_precio, var_tipo_pedido, CDbl(Me.txt_embarque))
                     Else
                        rsaux2.Open "UPDATE TB_DETALLE_CAJAS SET FLOA_PAQ_CANTIDAD = FLOA_PAQ_CANTIDAD + " + CStr(var_cantidad_leida) + " Where INTE_ORS_ORDEN_SURTIDO   = " + txt_archivo + " AND INTE_PAQ_CAJA = " + CStr(var_numero_caja) + "  and VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_ART_ARTICULO_ID  = '" + lv_salidas.selectedItem + "' AND INTE_EMB_EMBARQUE = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux4.Close
                  End If
                  rsaux4.Open "update tb_det_orden_surtido set floa_ors_cantidad_empacada = floa_ors_cantidad_empacada + " + CStr(var_cantidad_leida) + " where inte_ors_orden_surtido = " + CStr(var_orden_surtido) + " and vcha_Art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                  'cnn.CommitTrans
                  rsaux9.Open "select * from tb_trazabilidad_codigos where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND inte_emb_embarque = " + Me.txt_embarque + " and inte_ors_orden_surtido = " + Me.txt_archivo + " and vcha_tra_codigo = '" + Me.txt_codigo + "' AND INTE_PAQ_CAJA = " + CStr(var_numero_caja), cnn, adOpenDynamic, adLockOptimistic
                  If rsaux10.State = 1 Then
                     rsaux10.Close
                  End If
                  If Not rsaux9.EOF Then
                     rsaux10.Open "update tb_trazabilidad_codigos set floa_Tra_cantidad = isnull(floa_Tra_cantidad,0) + 1  where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND inte_emb_embarque = " + Me.txt_embarque + " and inte_ors_orden_surtido = " + Me.txt_archivo + " and vcha_tra_codigo = '" + Me.txt_codigo + "' AND INTE_PAQ_CAJA = " + CStr(var_numero_caja), cnn, adOpenDynamic, adLockOptimistic
                  Else
                     rsaux10.Open "insert into tb_trazabilidad_codigos (inte_emb_embarque, inte_ors_orden_surtido, vcha_tra_codigo, floa_tra_cantidad, vcha_tra_planta, inte_tra_lote, inte_tra_nodo_identificador, vcha_art_articulo_id, VCHA_EMP_EMPRESA_ID, INTE_PAQ_CAJA) values (" + Me.txt_embarque + "," + Me.txt_archivo + ",'" + Me.txt_codigo + "',1,'" + var_planta + "', " + var_lote + "," + CStr(var_nodo) + ",'" + var_codigo + "', '" + var_empresa + "'," + CStr(var_numero_caja) + ")", cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux9.Close
               End If
            Else
               rsaux4.Open "select * from tb_articulos where vcha_art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               VAR_descripcion_no = IIf(IsNull(rsaux4!vcha_art_nombre_español), "", rsaux4!vcha_art_nombre_español)
               rsaux4.Close
               frmmensaje.lbl_articulo = VAR_descripcion_no
               var_codigo = ""
               frmmensaje.lbl_mensaje = "Cantidad supera a la posible a surtir"
               frmmensaje.Show 1
               'MsgBox "Cantidad supera a la la posible a surtir", vbOKOnly, "ATENCION"
            End If
         Else
            rsaux4.Open "select * from tb_articulos where vcha_art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            VAR_descripcion_no = IIf(IsNull(rsaux4!vcha_art_nombre_español), "", rsaux4!vcha_art_nombre_español)
            rsaux4.Close
            frmmensaje.lbl_articulo = VAR_descripcion_no
            var_codigo = ""
            frmmensaje.lbl_mensaje = "El artículo no se encuentra dentro de la Orden de Surtido"
            frmmensaje.Show 1
            'MsgBox "El artículo no se encuentra dentro de la Orden de Surtido", vbOKOnly, "ATENCION"
            bandera_suma = False
         End If
         rs.Close
      Else
         frmmensaje.lbl_mensaje = "El artículo no existe"
         frmmensaje.Show 1
         'MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
         rsaux.Close
      End If
      var_planta = ""
      var_lote = ""
      var_codigo = ""
      var_nodo = 0
      txt_codigo.SetFocus
   End If
   Exit Sub
salir:
Resume
End Sub


Sub ejecuta()
   Dim var_embarque_agente As String
   Dim var_embarque_almacen As String
   Dim var_movimiento_agente As String
   Dim var_embarque_cerrado As String
   Dim var_clave_cliente_paquete As String
   Dim var_falta As Double
   Dim var_surtir As Double
   Dim var_surtido As Double
   Dim var_empacada As Double
   Dim var_posible As Boolean
   Dim var_asignado As Boolean
   Dim var_cerrado_embarque As Boolean
   Dim var_estatus_embarque As String
   var_autorizo_embarque = False
   var_clave_movimiento = txt_clave_movimiento
   var_posible = False
   If Trim(txt_archivo) <> "" Then
      rs.Open "select * from vw_orden_surtido where inte_ors_orden_surtido = " + txt_archivo + " and floa_ors_cantidad_surtir > 0 and floa_ors_cantidad_surtir > isnull(floa_ors_cantidad_surtida,0)+isnull(floa_ors_cantidad_negada,0) ", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         If var_clave_movimiento = rs!VCHA_MOV_MOVIMIENTO_ID Then
            If var_agente_embarque = rs!VCHA_AGE_AGENTE_ID Then
               var_liberada = IIf(IsNull(rs!inte_ors_liberada), 0, rs!inte_ors_liberada)
               If var_liberada = 1 Then
                  If IsNull(rs!VCHA_CLI_NOMBRE) Then
                     GoTo no_cliente:
                  Else
                     var_clave_cliente = rs!vcha_cli_clave_id
                  End If
                  If IsNull(rs!VCHA_CLI_NOMBRE) Then
                     GoTo no_cliente:
                  Else
                     txt_cliente = rs!VCHA_CLI_NOMBRE
                     var_clave_cliente = rs!vcha_cli_clave_id
                  End If
                  var_orden_surtido = txt_archivo
                  var_suma_cantidad_enviada = 0
                  var_suma_cantidad_recibida = 0
                  lbl_enviados.Caption = Format("0", "###,###,##0.00")
                  lbl_recibidos.Caption = Format("0", "###,###,##0.00")
                  lbl_empacados.Caption = Format("0", "###,###,##0.00")
                  lv_salidas.ListItems.Clear
                  If IsNull(rs!VCHA_cLI_EMAIL) Then
                     var_correo_electronico = ""
                  Else
                     var_correo_electronico = rs!VCHA_cLI_EMAIL
                  End If
                  If IsNull(rs!VCHA_ALM_NOMBRE) Then
                     GoTo no_almacen:
                  Else
                     var_almacen_origen = rs!VCHA_ALM_ALMACEN_ID
                     txt_origen = rs!VCHA_ALM_NOMBRE
                  End If
                  If IsNull(rs!VCHA_TIT_NOMBRE) Then
                     GoTo no_titular:
                  Else
                     txt_titular = rs!VCHA_TIT_NOMBRE
                     var_clave_titular = rs!vcha_tit_titular_id
                  End If
                  If IsNull(rs!inte_ped_dias_condiciones) Then
                     var_plazo = 0
                  Else
                     var_plazo = rs!inte_ped_dias_condiciones
                  End If
                  If IsNull(rs!VCHA_ESB_NOMBRE) Then
                     GoTo no_establecimiento:
                  Else
                     txt_establecimiento = rs!VCHA_ESB_NOMBRE
                     var_clave_establecimiento = rs!vcha_ESB_ESTABLECIMIENTO_id
                  End If
                  If IsNull(rs!VCHA_AGE_NOMBRE) Then
                     GoTo no_agente:
                  Else
                     txt_agente = rs!VCHA_AGE_NOMBRE
                     var_clave_agente = rs!VCHA_AGE_AGENTE_ID
                  End If
                  If IsNull(rs!vcha_rut_nombre) Then
                     txt_ruta = ""
                     var_clave_ruta = ""
                  Else
                     txt_ruta = rs!vcha_rut_nombre
                     var_clave_ruta = rs!VCHA_RUT_RUTA_ID
                  End If
                  If IsNull(rs!inte_ped_numero) Then
                     GoTo no_Pedido:
                  Else
                     txt_pedido = rs!inte_ped_numero
                  End If
                  If IsNull(rs!FLOA_ORS_DESCUENTO_1) Then
                     txt_descuento1 = 0
                     var_descuento_1 = 0
                  Else
                     txt_descuento1 = rs!FLOA_ORS_DESCUENTO_1
                     var_descuento_1 = rs!FLOA_ORS_DESCUENTO_1
                  End If
                  If IsNull(rs!FLOA_ORS_DESCUENTO_2) Then
                     txt_descuento2 = 0
                     var_descuento_2 = 0
                  Else
                     txt_descuento2 = rs!FLOA_ORS_DESCUENTO_2
                     var_descuento_2 = rs!FLOA_ORS_DESCUENTO_2
                  End If
                  var_descuento_3 = 0
                  var_cantidad_total_empacada = 0
                  While Not rs.EOF
                     'If IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), 0, rs!FLOA_ORS_CANTIDAD_SURTIR) < IIf(IsNull(rs!FLOA_ORS_CANTIDAD_NEGADA), 0, rs!FLOA_ORS_CANTIDAD_NEGADA) Then
                        Set list_item = lv_salidas.ListItems.Add(, , rs!vcha_Art_articulo_id)
                        list_item.SubItems(1) = IIf(IsNull(rs!vcha_art_nombre_español), "", Trim(rs!vcha_art_nombre_español))
                        list_item.SubItems(2) = Format(IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), 0, rs!FLOA_ORS_CANTIDAD_SURTIR), "###,###,##0.00")
                        var_surtir = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), 0, rs!FLOA_ORS_CANTIDAD_SURTIR)
                        list_item.SubItems(3) = Format(IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA), "###,###,##0.00")
                        var_surtido = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
                        list_item.SubItems(4) = Format(IIf(IsNull(rs!floa_ors_Cantidad_empacada), 0, rs!floa_ors_Cantidad_empacada), "###,###,##0.00")
                        var_empacada = IIf(IsNull(rs!floa_ors_Cantidad_empacada), 0, rs!floa_ors_Cantidad_empacada)
                        list_item.SubItems(5) = Format(0, "###,###,##0.00")
                        var_falta = var_surtir - (var_surtido + var_empacada)
                        list_item.SubItems(6) = Format(var_falta, "###,###,##0.00")
                        list_item.SubItems(7) = IIf(IsNull(rs!floa_ors_costo), "", rs!floa_ors_costo)
                        list_item.SubItems(8) = IIf(IsNull(rs!floa_ors_precio), "", rs!floa_ors_precio)
                        list_item.SubItems(9) = IIf(IsNull(rs!char_ped_tipo), "", rs!char_ped_tipo)
                        var_suma_cantidad_enviada = var_suma_cantidad_enviada + rs!FLOA_ORS_CANTIDAD_SURTIR
                        var_suma_cantidad_recibida = var_suma_cantidad_recibida + rs!FLOA_ORS_CANTIDAD_SURTIDA
                        var_cantidad_total_empacada = var_cantidad_total_empacada + IIf(IsNull(rs!floa_ors_Cantidad_empacada), 0, rs!floa_ors_Cantidad_empacada)
                        rs.MoveNext:
                     'End If
                  Wend
                  lbl_enviados = Format(Str(var_suma_cantidad_enviada), "###,###,##0.00")
                  lbl_recibidos = Format("0", "###,###,##0.00")
                  lbl_empacados = Format(Str(var_cantidad_total_empacada), "###,###,##0.00")
                  txt_codigo.Enabled = True
                  txt_codigo.SetFocus
               Else
                  MsgBox "La orden de surtido no a sido liberada", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "El agente de la orden de surtido no pertence al agente del embarque", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Orden de surtido incorrecta para el movimiento seleccionado", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Numero de Orden de surtido no existe", vbOKOnly, "ATENCION"
      End If
      rs.Close
      var_n = lv_salidas.ListItems.Count
      var_numero_renglones = Me.lv_salidas.Height / 312.5
      If var_n > var_numero_renglones Then
         lv_salidas.ColumnHeaders(2).Width = 4050.05
      Else
         lv_salidas.ColumnHeaders(2).Width = 4290
      End If
   Else
      txt_archivo.Enabled = True
      txt_archivo.SetFocus
   End If
Exit Sub
no_almacen:
   MsgBox "Almacen Incorrecto", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   Exit Sub
no_almacen_empaque:
   MsgBox "Almacen de Empaque Incorrecto", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   Exit Sub
no_Pedido:
   MsgBox "Pedido Incorrecto", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   Exit Sub
no_establecimiento:
   MsgBox "Establecimiento Incorrecto", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   Exit Sub
no_agente:
   MsgBox "Agente Incorrecto", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   Exit Sub
no_cliente:
   MsgBox "Cliente Incorrecto", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   Exit Sub
no_ruta:
   MsgBox "Ruta Incorrecta", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   Exit Sub
no_titular:
   MsgBox "Titular incorrecto", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   Exit Sub
no_almacen_agente:
   MsgBox "No existe un almacen relacionado a este agente", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   Exit Sub
End Sub



