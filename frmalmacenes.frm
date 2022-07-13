VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmalmacenes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Almacenes"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   ControlBox      =   0   'False
   Icon            =   "frmalmacenes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11670
   Begin VB.Frame frm_colonias 
      Height          =   2400
      Left            =   5610
      TabIndex        =   66
      Top             =   1305
      Width           =   5685
      Begin MSComctlLib.ListView lv_colonias 
         Height          =   1830
         Left            =   45
         TabIndex        =   67
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
         TabIndex        =   68
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   5550
      TabIndex        =   60
      Top             =   1320
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1875
         Left            =   30
         TabIndex        =   61
         Top             =   435
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   3307
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
         TabIndex        =   62
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmalmacenes.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   32
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frmalmacenes.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Guardar Alt + G"
      Top             =   33
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   720
      Picture         =   "frmalmacenes.frx":0ACE
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   34
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1050
      Picture         =   "frmalmacenes.frx":0BA0
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   35
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1380
      Picture         =   "frmalmacenes.frx":0CA2
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   36
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11205
      Picture         =   "frmalmacenes.frx":0DA4
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Salir"
      Top             =   37
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   6870
      Left            =   60
      TabIndex        =   43
      Top             =   390
      Width           =   5835
      Begin VB.TextBox txt_colonia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   9
         Top             =   2160
         Width           =   900
      End
      Begin VB.TextBox txt_nombre_colonia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   10
         Top             =   2160
         Width           =   3360
      End
      Begin VB.TextBox txt_nombre_pais 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   18
         Top             =   3480
         Width           =   3360
      End
      Begin VB.TextBox txt_nombre_estado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   16
         Top             =   3150
         Width           =   3360
      End
      Begin VB.TextBox txt_municipio 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   13
         Top             =   2820
         Width           =   900
      End
      Begin VB.TextBox txt_nombre_municipio 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   14
         Top             =   2820
         Width           =   3360
      End
      Begin VB.ComboBox cmb_tipo_almacenes 
         Height          =   315
         ItemData        =   "frmalmacenes.frx":13DE
         Left            =   2385
         List            =   "frmalmacenes.frx":13EB
         TabIndex        =   25
         Top             =   5145
         Width           =   3390
      End
      Begin VB.TextBox txt_tipo_reempaque 
         Height          =   315
         Left            =   2055
         MaxLength       =   50
         TabIndex        =   30
         Top             =   6150
         Width           =   900
      End
      Begin VB.TextBox txt_nombre_ciudad 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   12
         Top             =   2490
         Width           =   3360
      End
      Begin VB.TextBox txt_nombre_unidad 
         Height          =   315
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         Top             =   510
         Width           =   3360
      End
      Begin VB.TextBox txt_nombre_empresa 
         Height          =   315
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   2
         Top             =   180
         Width           =   3360
      End
      Begin VB.CheckBox chk_sobrantes 
         Caption         =   "Almacen de Sobrantes"
         Height          =   315
         Left            =   1470
         TabIndex        =   31
         Top             =   6510
         Width           =   3225
      End
      Begin VB.CheckBox chk_reempaque 
         Caption         =   "Almacen de Reempaque"
         Height          =   315
         Left            =   3390
         TabIndex        =   29
         Top             =   5790
         Width           =   2145
      End
      Begin VB.CheckBox chk_costeo 
         Caption         =   "Almacen de Costeo"
         Height          =   315
         Left            =   1470
         TabIndex        =   28
         Top             =   5790
         Width           =   1815
      End
      Begin VB.CheckBox chk_calidad 
         Caption         =   "Almacen de Calidad"
         Height          =   315
         Left            =   3390
         TabIndex        =   27
         Top             =   5475
         Width           =   1755
      End
      Begin VB.CheckBox chk_rechazo 
         Caption         =   "Almacen de Rechazo"
         Height          =   315
         Left            =   1470
         TabIndex        =   26
         Top             =   5475
         Width           =   1845
      End
      Begin VB.TextBox txt_clave_empresa 
         Height          =   315
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   1
         Top             =   180
         Width           =   900
      End
      Begin VB.TextBox txt_clave_unidad 
         Height          =   315
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   3
         Top             =   510
         Width           =   900
      End
      Begin VB.TextBox txt_direccion 
         Height          =   315
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1500
         Width           =   4290
      End
      Begin VB.TextBox txt_ciudad 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   11
         Top             =   2490
         Width           =   900
      End
      Begin VB.TextBox txt_estado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   15
         Top             =   3150
         Width           =   900
      End
      Begin VB.TextBox txt_pais 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   17
         Top             =   3480
         Width           =   900
      End
      Begin VB.TextBox txt_nombre_almacen 
         Height          =   315
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1170
         Width           =   4290
      End
      Begin VB.TextBox txt_clave_almacen 
         Height          =   315
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   5
         Top             =   840
         Width           =   900
      End
      Begin VB.TextBox txt_codigo_postal 
         Height          =   315
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1830
         Width           =   900
      End
      Begin VB.TextBox txt_afectacion 
         Height          =   315
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   19
         Top             =   3810
         Width           =   3285
      End
      Begin VB.TextBox txt_neteable 
         Height          =   315
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   21
         Top             =   4485
         Width           =   900
      End
      Begin VB.TextBox txt_prioridad 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4605
         TabIndex        =   22
         Top             =   4485
         Width           =   900
      End
      Begin VB.TextBox txt_correo 
         Height          =   315
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   23
         Top             =   4815
         Width           =   4290
      End
      Begin VB.CheckBox chk_surtir 
         Caption         =   "Surtir"
         Height          =   315
         Left            =   1470
         TabIndex        =   20
         Top             =   4140
         Width           =   750
      End
      Begin VB.TextBox txt_tipo 
         Height          =   315
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   24
         Top             =   5145
         Width           =   900
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Colonia:"
         Height          =   195
         Index           =   16
         Left            =   210
         TabIndex        =   65
         Top             =   2220
         Width           =   570
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Municipio:"
         Height          =   195
         Index           =   15
         Left            =   210
         TabIndex        =   64
         Top             =   2880
         Width           =   720
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Tipo entrada reempaque:"
         Height          =   195
         Index           =   6
         Left            =   210
         TabIndex        =   63
         Top             =   6210
         Width           =   1785
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Index           =   14
         Left            =   210
         TabIndex        =   59
         Top             =   5205
         Width           =   360
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Correo:"
         Height          =   195
         Index           =   13
         Left            =   210
         TabIndex        =   58
         Top             =   4875
         Width           =   510
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Prioridad:"
         Height          =   195
         Index           =   12
         Left            =   3345
         TabIndex        =   57
         Top             =   4545
         Width           =   660
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Codigo Postal:"
         Height          =   195
         Index           =   11
         Left            =   210
         TabIndex        =   56
         Top             =   1890
         Width           =   1020
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad:"
         Height          =   195
         Index           =   10
         Left            =   210
         TabIndex        =   55
         Top             =   2550
         Width           =   540
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Index           =   9
         Left            =   210
         TabIndex        =   54
         Top             =   3210
         Width           =   540
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Pais:"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   53
         Top             =   3540
         Width           =   345
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Empresa:"
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   50
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   49
         Top             =   900
         Width           =   405
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   48
         Top             =   1230
         Width           =   600
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Dirección:"
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   47
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Contable:"
         Height          =   195
         Index           =   7
         Left            =   210
         TabIndex        =   46
         Top             =   3870
         Width           =   1230
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Planta:"
         Height          =   195
         Index           =   8
         Left            =   210
         TabIndex        =   45
         Top             =   570
         Width           =   495
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Neteable:"
         Height          =   195
         Index           =   5
         Left            =   210
         TabIndex        =   44
         Top             =   4545
         Width           =   690
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   5940
      TabIndex        =   41
      Top             =   390
      Width           =   5625
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   1905
         TabIndex        =   38
         Top             =   165
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3570
         TabIndex        =   39
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
         Caption         =   "Busqueda de Almacen:"
         Height          =   195
         Left            =   195
         TabIndex        =   42
         Top             =   195
         Width           =   1650
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2655
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   75
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   2580
      Top             =   45
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
            Picture         =   "frmalmacenes.frx":1408
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmalmacenes.frx":1CE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmalmacenes.frx":25BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmalmacenes.frx":2B58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmalmacenes.frx":3434
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmalmacenes.frx":3D0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmalmacenes.frx":45E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmalmacenes.frx":46FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmalmacenes.frx":480C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmalmacenes.frx":491E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmalmacenes.frx":4A30
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   0
      TabIndex        =   52
      Top             =   270
      Width           =   11550
   End
   Begin VB.Frame Frame3 
      Height          =   6390
      Left            =   5955
      TabIndex        =   51
      Top             =   870
      Width           =   5640
      Begin MSComctlLib.ImageList icono_encabezado 
         Left            =   2475
         Top             =   -75
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
               Picture         =   "frmalmacenes.frx":4B42
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmalmacenes.frx":541C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lv_almacenes 
         Height          =   6195
         Left            =   30
         TabIndex        =   40
         Top             =   150
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   10927
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   23
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "empresa"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "planta"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Clave"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nombre"
            Object.Width           =   6879
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Pais"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "estado"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ciudad"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "direccion"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "cp"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "cuenta contable"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "surtir"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Neteable"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Prioridad"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Correo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "tipo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "rechazo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "calidad"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "costeo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Reempaque"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "tipo reempaque"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "sobrantes"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "COLONIA"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "MUNICIPIO"
            Object.Width           =   0
         EndProperty
      End
   End
End
Attribute VB_Name = "frmalmacenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim numero_items_ALMACENES As Integer
Dim bitacora As Boolean
Dim var_pais As String
Dim var_estado As String
Dim var_bit_pais As String
Dim var_bit_estado As String
Dim var_ciudad As String
Dim varnombrepais As String
Dim varnombreestado As String
Dim var_tipo_lista As Integer






Private Sub chk_calidad_Click()
   var_hubo_cambios = True
End Sub

Private Sub chk_costeo_Click()
   var_hubo_cambios = True
End Sub

Private Sub chk_rechazo_Click()
   var_hubo_cambios = True
End Sub

Private Sub chk_reempaque_Click()
   var_hubo_cambios = True
End Sub

Private Sub chk_sobrantes_Click()
   var_hubo_cambios = True
End Sub

Private Sub chk_surtir_Click()
   var_hubo_cambios = True
End Sub

Private Sub cmb_tipo_almacenes_Click()
   var_hubo_cambios = True
   If Trim(cmb_tipo_almacenes = "ALMACEN") Then
      txt_tipo = "A"
   End If
   If Trim(cmb_tipo_almacenes = "TIENDA") Then
      txt_tipo = "T"
   End If
   If Trim(cmb_tipo_almacenes = "AGENTE") Then
      txt_tipo = "G"
   End If
End Sub

Private Sub cmb_tipo_almacenes_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub cmd_deshacer_Click()
   txt_clave_empresa.Enabled = False
   txt_clave_unidad.Enabled = False
   txt_clave_almacen.Enabled = False
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
      txt_clave_empresa.Enabled = False
      txt_clave_unidad.Enabled = False
      txt_clave_almacen.Enabled = False
      Call pro_elimina_ALMACENES
      rs.Open "select * from tb_ALMACENES", cnn, adOpenDynamic, adLockOptimistic
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
   If var_modifica_registro_almacen = False Then
      rs.Open "select * from tb_almacenes where vcha_alm_almacen_id = '" + Me.txt_clave_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_posible = False
      End If
      rs.Close
   End If
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
         If txt_clave_empresa = "" Or txt_clave_unidad = "" Or txt_clave_almacen = "" Or txt_nombre_almacen = "" Or txt_pais = "" Or _
            txt_estado = "" Or txt_pais = "" Or txt_direccion = "" Or txt_codigo_postal = "" Or txt_afectacion = "" Or _
            txt_tipo = "" Then
            MsgBox "Información Incompleta", vbOKOnly, "ATENCION"
         Else
            txt_clave_empresa.Enabled = False
            txt_clave_unidad.Enabled = False
            txt_clave_almacen.Enabled = False
            Call pro_guardar_ALMACENES
            rs.Open "select * from tb_ALMACENES", cnn, adOpenDynamic, adLockOptimistic
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
      End If
   Else
      MsgBox "La clave del almacen ya existe", vbOKOnly, "ATENCION"
   End If

End Sub

Private Sub cmd_imprimir_Click()
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call gPrintListView(lv_almacenes, "LISTADO DE ALMACENES")
        End If

End Sub

Private Sub cmd_nuevo_Click()
   Call pro_limpiatextos(Me)
   txt_clave_empresa.Enabled = True
   txt_clave_empresa.SetFocus: var_modifica_registro_almacen = False
   cmd_guardar.Enabled = True
   cmd_deshacer.Enabled = True
   txt_clave_empresa.Enabled = True
   txt_clave_unidad.Enabled = True
   txt_clave_almacen.Enabled = True
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_almacen = False Then
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
   var_swpassword = False
   var_modifica_registro_almacen = False
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
   frm_colonias.Visible = False
   frm_lista.Visible = False
   var_modifica_registro_almacen = True
   lv_almacenes.SmallIcons = ImageList
   Call pro_encabezadosView(Me, lv_almacenes, False)
   Call pro_llena_listview1
   pro_textos
   rs.Open "select * from tb_ALMACENES", cnn, adOpenDynamic, adLockOptimistic
   txt_clave_empresa.Enabled = False
   txt_clave_unidad.Enabled = False
   txt_clave_almacen.Enabled = False
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
   Call activa_forma(var_activa_forma_almacenes)
End Sub

Private Sub lv_almacenes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_almacenes, ColumnHeader)
End Sub

Private Sub lv_almacenes_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Set lv_almacenes.selectedItem = Item
   pro_textos
   var_modifica_registro_almacen = True
   txt_clave_empresa.Enabled = False
End Sub

Private Sub lv_colonias_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_colonias, ColumnHeader)
End Sub

Private Sub lv_colonias_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_colonias.ListItems.Count > 0 Then
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
      frm_colonias.Visible = False
      Me.txt_afectacion.SetFocus
   End If
   If KeyAscii = 27 Then
      frm_colonias.Visible = False
      Me.txt_codigo_postal.SetFocus
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
      If var_tipo_lista = 3 Then
         If lv_lista.ListItems.Count > 0 Then
         rs.Open "select * from tb_ciudades where vcha_ciu_ciudad_id = '" + lv_lista.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
               txt_ciudad = lv_lista.selectedItem
               txt_nombre_ciudad = rs!vcha_ciu_nombre
               varpais = IIf(IsNull(rs(0).Value), "", rs(0).Value)
               varestado = IIf(IsNull(rs(1).Value), "", rs(1).Value)
               var_bit_pais = IIf(IsNull(rs(0).Value), "", rs(0).Value)
               var_bit_estado = IIf(IsNull(rs(1).Value), "", rs(1).Value)
               varciudad = IIf(IsNull(rs(2).Value), "", rs(2).Value)
               txt_pais = IIf(IsNull(rs(2).Value), "", rs(2).Value)
               rs.Close
               rs.Open "select * from tb_paises where vcha_pai_pais_id = '" & varpais & "'", cnn, adOpenDynamic, adLockBatchOptimistic
               If Not rs.EOF Then
                  txt_pais = IIf(IsNull(rs(1).Value), "", rs(1).Value)
                  varnombrepais = IIf(IsNull(rs(1).Value), "", rs(1).Value)
               Else
                  txt_pais = ""
                  varnombrepais = ""
               End If
               rs.Close
               rs.Open "select * from tb_estados where vcha_pai_pais_id = '" & varpais & "' and vcha_est_estado_id = '" & varestado & "'", cnn, adOpenDynamic, adLockBatchOptimistic
               If Not rs.EOF Then
                  txt_estado = IIf(IsNull(rs(2).Value), "", rs(2).Value)
                  varnombreestado = IIf(IsNull(rs(2).Value), "", rs(2).Value)
               Else
                  txt_estado = ""
                  varnombreestado = ""
               End If
               rs.Close
               Call pro_enfoque(KeyAscii)
            Else
               rs.Close
            End If
         Else
            txt_ciudad = ""
            txt_nombre_ciudad = ""
            varpais = ""
            varestado = ""
            var_bit_pais = ""
            var_bit_estado = ""
            varciudad = ""
            txt_pais = ""
            txt_pais = ""
            varnombrepais = ""
            txt_estado = ""
            varnombreestado = ""
         End If
      End If
      If var_tipo_lista = 1 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_clave_empresa = lv_lista.selectedItem
            txt_nombre_empresa = lv_lista.selectedItem.SubItems(1)
         Else
            txt_clave_empresa = ""
            txt_nombre_empresa = ""
         End If
      End If
      If var_tipo_lista = 2 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_clave_unidad = lv_lista.selectedItem
            txt_nombre_unidad = lv_lista.selectedItem.SubItems(1)
         Else
            txt_clave_unidad = ""
            txt_nombre_unidad = ""
         End If
      End If
      frm_lista.Visible = False
   End If
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_almacenes.SetFocus
      Call pro_avanzar(Me, lv_almacenes, Button)
      lv_almacenes.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_almacenes.ListItems(1).Selected = True
      lv_almacenes.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 4 Then
      numero_items_ALMACENES = lv_almacenes.ListItems.Count
      lv_almacenes.ListItems(numero_items_ALMACENES).Selected = True
      lv_almacenes.selectedItem.EnsureVisible
      pro_textos
   End If
err0:
End Sub

Sub pro_guardar_ALMACENES()
   Dim ok As Boolean
   Set TB_ALMACENES = New TB_ALMACENES
   Set TB_BITACORA_ALMACENES = New TB_BITACORA_ALMACENES
   ok = True
   If txt_clave_empresa <> "" And txt_clave_unidad <> "" Then
      If var_hubo_cambios Then
         rs.Open "select * from tb_ALMACENES where vcha_alm_almacen_id = '" + txt_clave_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
         ok = TB_ALMACENES.Anadir(txt_clave_empresa, txt_clave_unidad, txt_clave_almacen, txt_nombre_almacen, txt_pais, txt_estado, txt_ciudad, txt_direccion, txt_codigo_postal, txt_afectacion, chk_surtir, txt_neteable, txt_prioridad, txt_correo, txt_tipo, chk_rechazo, chk_calidad, chk_costeo, chk_reempaque, Val(txt_tipo_reempaque), chk_sobrantes, txt_colonia, txt_municipio)
         If ok Then
            bitacora = True
            If var_modifica_registro_almacen = False Then
               var_operacion_bitacora = "I"
               bitacora = TB_BITACORA_ALMACENES.Anadir(txt_clave_almacen, "VCHA_ALM_NOMBRE", var_operacion_bitacora, "", txt_clave_unidad, var_clave_usuario_global, fun_NombrePc, Date)
            Else
               var_operacion_bitacora = "M"
               If rs!vcha_emp_empresa_id <> txt_clave_empresa Then
                  bitacora = TB_BITACORA_ALMACENES.Anadir(txt_clave_almacen, "VCHA_EMP_EMPRESA_ID", var_operacion_bitacora, rs!vcha_empresa_id, txt_clave_empresa, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!VCHA_UOR_UNIDAD_ID <> txt_clave_unidad Then
                  bitacora = TB_BITACORA_ALMACENES.Anadir(txt_clave_almacen, "VCHA_UOR_UNIDAD_ID", var_operacion_bitacora, rs!VCHA_UOR_UNIDAD_ID, txt_clave_unidad, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!vcha_alm_almacen_id <> txt_clave_almacen Then
                  bitacora = TB_BITACORA_ALMACENES.Anadir(txt_clave_almacen, "VCHA_ALM_ALMACEN_ID", var_operacion_bitacora, rs!vcha_alm_almacen_id, txt_clave_almacen, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!VCHA_ALM_NOMBRE <> txt_nombre_almacen Then
                  bitacora = TB_BITACORA_ALMACENES.Anadir(txt_clave_almacen, "VCHA_ALM_NOMBRE", var_operacion_bitacora, rs!VCHA_ALM_NOMBRE, txt_nombre_almacen, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!vcha_pai_pais_id <> varpais Then
                  bitacora = TB_BITACORA_ALMACENES.Anadir(txt_clave_almacen, "VCHA_PAI_PAIS_ID", var_operacion_bitacora, rs!vcha_pai_pais_id, varpais, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!vcha_est_estado_id <> varestado Then
                  bitacora = TB_BITACORA_ALMACENES.Anadir(txt_clave_almacen, "VCHA_EST_ESTADO_ID", var_operacion_bitacora, rs!vcha_est_estado_id, varestado, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!vcha_ciu_ciudad_id <> txt_ciudad Then
                  bitacora = TB_BITACORA_ALMACENES.Anadir(txt_clave_almacen, "VCHA_CIU_CIUDAD_ID", var_operacion_bitacora, rs!vcha_ciu_ciudad_id, txt_ciudad, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!VCHA_ALM_DIRECCION <> txt_direccion Then
                  bitacora = TB_BITACORA_ALMACENES.Anadir(txt_clave_almacen, "VCHA_ALM_DIRECCION", var_operacion_bitacora, rs!VCHA_ALM_DIRECCION, txt_direccion, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!VCHA_ALM_CP <> txt_codigo_postal Then
                  bitacora = TB_BITACORA_ALMACENES.Anadir(txt_clave_almacen, "VCHA_ALM_CP", var_operacion_bitacora, rs!VCHA_ALM_CP, txt_codigo_postal, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!VCHA_ALM_AFECTACION <> txt_afectacion Then
                  bitacora = TB_BITACORA_ALMACENES.Anadir(txt_clave_almacen, "VCHA_ALM_AFECTACION_CONTABLE", var_operacion_bitacora, rs!VCHA_ALM_AFECTACION, txt_afectacion, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!INTE_ALM_SURTIR <> chk_surtir Then
                  bitacora = TB_BITACORA_ALMACENES.Anadir(txt_clave_almacen, "VCHA_ALM_SURTIR", var_operacion_bitacora, rs!INTE_ALM_SURTIR, chk_surtir, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!VCHA_ALM_NETEABLE <> txt_neteable Then
                  bitacora = TB_BITACORA_ALMACENES.Anadir(txt_clave_almacen, "VCHA_ALM_NETEABLE", var_operacion_bitacora, rs!VCHA_ALM_NETEABLE, txt_neteable, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!INTE_ALM_PRIORIDAD <> txt_prioridad Then
                  bitacora = TB_BITACORA_ALMACENES.Anadir(txt_clave_almacen, "VCHA_ALM_PRIORIDAD", var_operacion_bitacora, rs!INTE_ALM_PRIORIDAD, txt_prioridad, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!vcha_alm_correo <> txt_correo Then
                  bitacora = TB_BITACORA_ALMACENES.Anadir(txt_clave_almacen, "VCHA_ALM_CORREO", var_operacion_bitacora, rs!vcha_alm_correo, txt_correo, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!char_alm_tipo <> txt_tipo Then
                  bitacora = TB_BITACORA_ALMACENES.Anadir(txt_clave_almacen, "VCHA_ALM_TIPO", var_operacion_bitacora, rs!char_alm_tipo, txt_tipo, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!inte_alm_rechazo <> chk_rechazo Then
                  bitacora = TB_BITACORA_ALMACENES.Anadir(txt_clave_almacen, "INTE_ALM_RECHAZO", var_operacion_bitacora, rs!inte_alm_rechazo, chk_rechazo, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!INTE_ALM_CALIDAD <> chk_calidad Then
                  bitacora = TB_BITACORA_ALMACENES.Anadir(txt_clave_almacen, "INTE_ALM_CALIDAD", var_operacion_bitacora, rs!INTE_ALM_CALIDAD, chk_calidad, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!INTE_ALM_COSTEO <> chk_costeo Then
                  bitacora = TB_BITACORA_ALMACENES.Anadir(txt_clave_almacen, "INTE_ALM_COSTEO", var_operacion_bitacora, rs!INTE_ALM_COSTEO, chk_costeo, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!INTE_ALM_REEMPAQUE <> chk_reempaque Then
                  bitacora = TB_BITACORA_ALMACENES.Anadir(txt_clave_almacen, "INTE_ALM_REEMPAQUE", var_operacion_bitacora, rs!INTE_ALM_REEMPAQUE, chk_reempaque, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!INTE_ALM_TIPO_ENTRADA_REEMPAQUE <> Val(txt_tipo_reempaque) Then
                  bitacora = TB_BITACORA_ALMACENES.Anadir(txt_clave_almacen, "INTE_ALM_TIPO_ENTRADA_REEMPAQUE", var_operacion_bitacora, rs!INTE_ALM_TIPO_ENTRADA_REEMPAQUE, txt_tipo_reempaque, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!INTE_ALM_SOBRANTES <> chk_sobrantes Then
                  bitacora = TB_BITACORA_ALMACENES.Anadir(txt_clave_almacen, "INTE_ALM_SOBRANTES", var_operacion_bitacora, rs!INTE_ALM_SOBRANTES, chk_sobrantes, var_clave_usuario_global, fun_NombrePc, Date)
               End If
            End If
            rs.Close
            pro_actualiza_ListView
            txt_clave_empresa.Enabled = False
            MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
            txt_registros = lv_almacenes.ListItems.Count
            var_modifica_registro_almacen = True
         Else
            MsgBox "No se puede grabar registro: " + TB_ALMACENES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
         End If
      End If
   End If
   Set TB_ALMACENES = Nothing: var_hubo_cambios = False
End Sub

Sub pro_elimina_ALMACENES()
   Dim var_llave_usuarios As String
   Set TB_ALMACENES = New TB_ALMACENES
   Set TB_BITACORA_ALMACENES = New TB_BITACORA_ALMACENES
   On Error GoTo salir:
   ok = True
   If txt_clave_empresa <> "" And txt_clave_unidad <> "" And var_modifica_registro_almacen = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_ALMACENES.Eliminar(txt_clave_almacen)
      Else
         GoTo salir:
      End If
      If ok Then
         var_operacion_bitacora = "E"
         bitacora = TB_BITACORA_ALMACENES.Anadir(txt_clave_empresa, "VCHA_ALM_NOMBRE", var_operacion_bitacora, txt_clave_unidad, "", var_clave_usuario_global, fun_NombrePc, Date)
         numero_items_ALMACENES = numero_items_ALMACENES - 1
         MsgBox "Se Elimino Correctamente el Registro", vbInformation
         lv_almacenes.ListItems.Remove (lv_almacenes.selectedItem.Index)
         Call pro_limpiatextos(Me)
         txt_registros = lv_almacenes.ListItems.Count
         lv_almacenes.selectedItem.Selected = True
         pro_textos
      Else
        MsgBox "No se puede eliminar registro: " + TB_ALMACENES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
   Set TB_ALMACENES = Nothing
End Sub


Sub pro_llena_listview1()
   Dim list_item As ListItem
   rs.Open "select * from TB_ALMACENES", cnn, adOpenDynamic, adLockOptimistic
   numero_items_ALMACENES = 0
   While Not rs.EOF
      Set list_item = lv_almacenes.ListItems.Add(, , rs!vcha_emp_empresa_id)
      list_item.SubItems(1) = IIf(IsNull(rs!VCHA_UOR_UNIDAD_ID), "", rs!VCHA_UOR_UNIDAD_ID)
      list_item.SubItems(2) = IIf(IsNull(rs!vcha_alm_almacen_id), "", rs!vcha_alm_almacen_id)
      list_item.SubItems(3) = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
      list_item.SubItems(4) = IIf(IsNull(rs!vcha_pai_pais_id), "", rs!vcha_pai_pais_id)
      list_item.SubItems(5) = IIf(IsNull(rs!vcha_est_estado_id), "", rs!vcha_est_estado_id)
      list_item.SubItems(6) = IIf(IsNull(rs!vcha_ciu_ciudad_id), "", rs!vcha_ciu_ciudad_id)
      list_item.SubItems(7) = IIf(IsNull(rs!VCHA_ALM_DIRECCION), "", rs!VCHA_ALM_DIRECCION)
      list_item.SubItems(8) = IIf(IsNull(rs!VCHA_ALM_CP), "", rs!VCHA_ALM_CP)
      list_item.SubItems(9) = IIf(IsNull(rs!VCHA_ALM_AFECTACION), "", rs!VCHA_ALM_AFECTACION)
      list_item.SubItems(10) = IIf(IsNull(rs!INTE_ALM_SURTIR), 0, rs!INTE_ALM_SURTIR)
      list_item.SubItems(11) = IIf(IsNull(rs!VCHA_ALM_NETEABLE), "", rs!VCHA_ALM_NETEABLE)
      list_item.SubItems(12) = IIf(IsNull(rs!INTE_ALM_PRIORIDAD), 0, rs!INTE_ALM_PRIORIDAD)
      list_item.SubItems(13) = IIf(IsNull(rs!vcha_alm_correo), "", rs!vcha_alm_correo)
      list_item.SubItems(14) = IIf(IsNull(rs!char_alm_tipo), "", rs!char_alm_tipo)
      list_item.SubItems(15) = IIf(IsNull(rs!inte_alm_rechazo), 0, rs!inte_alm_rechazo)
      list_item.SubItems(16) = IIf(IsNull(rs!INTE_ALM_CALIDAD), 0, rs!INTE_ALM_CALIDAD)
      list_item.SubItems(17) = IIf(IsNull(rs!INTE_ALM_COSTEO), 0, rs!INTE_ALM_COSTEO)
      list_item.SubItems(18) = IIf(IsNull(rs!INTE_ALM_REEMPAQUE), 0, rs!INTE_ALM_REEMPAQUE)
      list_item.SubItems(19) = IIf(IsNull(rs!INTE_ALM_TIPO_ENTRADA_REEMPAQUE), 0, rs!INTE_ALM_TIPO_ENTRADA_REEMPAQUE)
      list_item.SubItems(20) = IIf(IsNull(rs!INTE_ALM_SOBRANTES), 0, rs!INTE_ALM_SOBRANTES)
      list_item.SubItems(21) = IIf(IsNull(rs!VCHA_COL_COLONIA_ID), "", rs!VCHA_COL_COLONIA_ID)
      list_item.SubItems(22) = IIf(IsNull(rs!vcha_mun_municipio_id), "", rs!vcha_mun_municipio_id)
      
      rs.MoveNext:
      numero_items_ALMACENES = numero_items_ALMACENES + 1
    Wend
    If numero_items_ALMACENES > 11 Then
       lv_almacenes.ColumnHeaders(4).Width = 4050
    Else
       lv_almacenes.ColumnHeaders(4).Width = 4250.04
    End If

    rs.Close
End Sub


Sub pro_textos()
'On Error GoTo err0:
Dim var_n As Double
var_n = lv_almacenes.ListItems.Count
   If var_n > 0 Then
      txt_clave_empresa = lv_almacenes.selectedItem
      txt_clave_unidad = lv_almacenes.selectedItem.SubItems(1)
      txt_clave_almacen = lv_almacenes.selectedItem.SubItems(2)
      txt_nombre_almacen = lv_almacenes.selectedItem.SubItems(3)
      txt_pais = lv_almacenes.selectedItem.SubItems(4)
      txt_estado = lv_almacenes.selectedItem.SubItems(5)
      txt_ciudad = lv_almacenes.selectedItem.SubItems(6)
      txt_direccion = lv_almacenes.selectedItem.SubItems(7)
      txt_codigo_postal = lv_almacenes.selectedItem.SubItems(8)
      txt_afectacion = lv_almacenes.selectedItem.SubItems(9)
      chk_surtir = lv_almacenes.selectedItem.SubItems(10)
      txt_neteable = lv_almacenes.selectedItem.SubItems(11)
      txt_prioridad = lv_almacenes.selectedItem.SubItems(12)
      txt_correo = lv_almacenes.selectedItem.SubItems(13)
      txt_tipo = lv_almacenes.selectedItem.SubItems(14)
      
      chk_rechazo = lv_almacenes.selectedItem.SubItems(15)
      chk_calidad = lv_almacenes.selectedItem.SubItems(16)
      chk_costeo = lv_almacenes.selectedItem.SubItems(17)
      chk_reempaque = lv_almacenes.selectedItem.SubItems(18)
      txt_tipo_reempaque = lv_almacenes.selectedItem.SubItems(19)
      chk_sobrantes = lv_almacenes.selectedItem.SubItems(20)
      txt_colonia = lv_almacenes.selectedItem.SubItems(21)
      txt_municipio = lv_almacenes.selectedItem.SubItems(22)
      
      If Trim(txt_tipo = "A") Then
         cmb_tipo_almacenes = "ALMACEN"
      End If
      If Trim(txt_tipo = "T") Then
         cmb_tipo_almacenes = "TIENDA"
      End If
      If Trim(txt_tipo = "") Then
         cmb_tipo_almacenes = ""
      End If
      rs.Open "SELECT * FROM TB_EMPRESAS WHERE VCHA_EMP_EMPRESA_ID = '" + txt_clave_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_empresa = IIf(IsNull(rs!vcha_emp_nombre), "", rs!vcha_emp_nombre)
      Else
         txt_nombre_empresa = ""
      End If
      rs.Close
      rs.Open "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE VCHA_UOR_UNIDAD_ID = '" + txt_clave_unidad + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_unidad = IIf(IsNull(rs!vcha_uor_nombre), "", rs!vcha_uor_nombre)
      Else
         txt_nombre_unidad = ""
      End If
      rs.Close
      rs.Open "SELECT * FROM TB_PAISES WHERE VCHA_PAI_PAIS_ID = '" + txt_pais + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_pais = IIf(IsNull(rs!VCHA_PAI_NOMBRE), "", rs!VCHA_PAI_NOMBRE)
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
         txt_nombre_municipio = IIf(IsNull(rs!VCHA_MUN_NOMBRE), "", rs!VCHA_MUN_NOMBRE)
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
      rs.Open "SELECT * FROM TB_COLONIAS WHERE VCHA_COL_COLONIA_ID = '" + txt_colonia + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_colonia = IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
      Else
         txt_nombre_colonia = ""
      End If
      rs.Close
   End If
   var_numero_renglones = lv_almacenes.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_almacenes.ColumnHeaders(4).Width = 3850
   Else
      lv_almacenes.ColumnHeaders(4).Width = 4099.71
   End If
   var_hubo_cambios = False
   var_modifica_registro_almacen = True
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem

   If var_modifica_registro_almacen = False Then
      Set list_item = lv_almacenes.ListItems.Add(, , txt_clave_empresa)
      list_item.SubItems(1) = txt_clave_unidad
      list_item.SubItems(2) = txt_clave_almacen
      list_item.SubItems(3) = txt_nombre_almacen
      list_item.SubItems(4) = txt_pais
      list_item.SubItems(5) = txt_estado
      list_item.SubItems(6) = txt_ciudad
      list_item.SubItems(7) = txt_direccion
      list_item.SubItems(8) = txt_codigo_postal
      list_item.SubItems(9) = txt_afectacion
      list_item.SubItems(10) = chk_surtir
      list_item.SubItems(11) = txt_neteable
      list_item.SubItems(12) = txt_prioridad
      list_item.SubItems(13) = txt_correo
      list_item.SubItems(14) = txt_tipo
      list_item.SubItems(15) = chk_rechazo
      list_item.SubItems(16) = chk_calidad
      list_item.SubItems(17) = chk_costeo
      list_item.SubItems(18) = chk_reempaque
      list_item.SubItems(19) = txt_tipo_reempaque
      list_item.SubItems(20) = chk_sobrantes
      list_item.SubItems(21) = txt_colonia
      list_item.SubItems(22) = txt_municipio
      list_item.EnsureVisible
      list_item.Selected = True
      numero_items_ALMACENES = numero_items_ALMACENES + 1
   Else
      lv_almacenes.ListItems.Item(lv_almacenes.selectedItem.Index).Checked = False
      lv_almacenes.ListItems.Item(lv_almacenes.selectedItem.Index) = txt_clave_empresa
      lv_almacenes.ListItems.Item(lv_almacenes.selectedItem.Index).ListSubItems(1) = txt_clave_unidad
      lv_almacenes.ListItems.Item(lv_almacenes.selectedItem.Index).ListSubItems(2) = txt_clave_almacen
      lv_almacenes.ListItems.Item(lv_almacenes.selectedItem.Index).ListSubItems(3) = txt_nombre_almacen
      lv_almacenes.ListItems.Item(lv_almacenes.selectedItem.Index).ListSubItems(4) = txt_pais
      lv_almacenes.ListItems.Item(lv_almacenes.selectedItem.Index).ListSubItems(5) = txt_estado
      lv_almacenes.ListItems.Item(lv_almacenes.selectedItem.Index).ListSubItems(6) = txt_ciudad
      lv_almacenes.ListItems.Item(lv_almacenes.selectedItem.Index).ListSubItems(7) = txt_direccion
      lv_almacenes.ListItems.Item(lv_almacenes.selectedItem.Index).ListSubItems(8) = txt_codigo_postal
      lv_almacenes.ListItems.Item(lv_almacenes.selectedItem.Index).ListSubItems(9) = txt_afectacion
      lv_almacenes.ListItems.Item(lv_almacenes.selectedItem.Index).ListSubItems(10) = chk_surtir
      lv_almacenes.ListItems.Item(lv_almacenes.selectedItem.Index).ListSubItems(11) = txt_neteable
      lv_almacenes.ListItems.Item(lv_almacenes.selectedItem.Index).ListSubItems(12) = txt_prioridad
      lv_almacenes.ListItems.Item(lv_almacenes.selectedItem.Index).ListSubItems(13) = txt_correo
      lv_almacenes.ListItems.Item(lv_almacenes.selectedItem.Index).ListSubItems(14) = txt_tipo
      lv_almacenes.ListItems.Item(lv_almacenes.selectedItem.Index).ListSubItems(15) = chk_rechazo
      lv_almacenes.ListItems.Item(lv_almacenes.selectedItem.Index).ListSubItems(16) = chk_calidad
      lv_almacenes.ListItems.Item(lv_almacenes.selectedItem.Index).ListSubItems(17) = chk_costeo
      lv_almacenes.ListItems.Item(lv_almacenes.selectedItem.Index).ListSubItems(18) = chk_reempaque
      lv_almacenes.ListItems.Item(lv_almacenes.selectedItem.Index).ListSubItems(19) = txt_tipo_reempaque
      lv_almacenes.ListItems.Item(lv_almacenes.selectedItem.Index).ListSubItems(20) = chk_sobrantes
      lv_almacenes.ListItems.Item(lv_almacenes.selectedItem.Index).ListSubItems(21) = txt_colonia
      lv_almacenes.ListItems.Item(lv_almacenes.selectedItem.Index).ListSubItems(22) = txt_municipio
      lv_almacenes.ListItems.Item(lv_almacenes.selectedItem.Index).Selected = True
   End If
End Sub



Private Sub txt_afectacion_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_afectacion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_almacenes, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub

Private Sub txt_ciudad_Change()
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub

Private Sub txt_ciudad_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_CIUDADES order by VCHA_CIU_NOMBRE", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_ciu_ciudad_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Plantas"
      var_tipo_lista = 3
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_ciudad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
End Sub

Private Sub txt_ciudad_LostFocus()
   If Trim(txt_ciudad) <> "" Then
      rs.Open "SELECT * FROM tb_ciudades where vcha_ciu_ciudad_id = '" + txt_ciudad + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_ciudad = rs!vcha_ciu_nombre
         varpais = IIf(IsNull(rs(0).Value), "", rs(0).Value)
         varestado = IIf(IsNull(rs(1).Value), "", rs(1).Value)
         var_bit_pais = IIf(IsNull(rs(0).Value), "", rs(0).Value)
         var_bit_estado = IIf(IsNull(rs(1).Value), "", rs(1).Value)
         varciudad = IIf(IsNull(rs(2).Value), "", rs(2).Value)
         txt_pais = IIf(IsNull(rs(2).Value), "", rs(2).Value)
         rs.Close
         rs.Open "select * from tb_paises where vcha_pai_pais_id = '" & varpais & "'", cnn, adOpenDynamic, adLockBatchOptimistic
         If Not rs.EOF Then
            txt_pais = IIf(IsNull(rs(1).Value), "", rs(1).Value)
            varnombrepais = IIf(IsNull(rs(1).Value), "", rs(1).Value)
         Else
            txt_pais = ""
            varnombrepais = ""
         End If
         rs.Close
         rs.Open "select * from tb_estados where vcha_pai_pais_id = '" & varpais & "' and vcha_est_estado_id = '" & varestado & "'", cnn, adOpenDynamic, adLockBatchOptimistic
         If Not rs.EOF Then
            txt_estado = IIf(IsNull(rs(2).Value), "", rs(2).Value)
            varnombreestado = IIf(IsNull(rs(2).Value), "", rs(2).Value)
         Else
            txt_estado = ""
            varnombreestado = ""
         End If
         rs.Close
      Else
         rs.Close
         MsgBox "Clave de ciudad incorrecta", vbOKOnly, "ATENCION"
         txt_ciudad = ""
         txt_nombre_ciudad = ""
         txt_estado = ""
         txt_pais = ""
      End If
   End If
End Sub

Private Sub txt_clave_almacen_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_clave_almacen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_clave_empresa_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_clave_empresa_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_empresas order by vcha_emp_empresa_id", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_emp_empresa_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_emp_nombre), "", rs!vcha_emp_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Empresas"
      var_tipo_lista = 1
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_clave_empresa_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_clave_empresa_LostFocus()
   If Trim(txt_clave_empresa) <> "" Then
      rs.Open "select * from tb_empresas where vcha_emp_empresa_id ='" + Me.txt_clave_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_empresa = rs!vcha_emp_nombre
      Else
         MsgBox "Clave de empresa incorrecta", vbOKOnly, "ATENCION"
         txt_clave_empresa = ""
         txt_nombre_empresa = ""
      End If
      rs.Close
   End If
End Sub

Private Sub txt_clave_unidad_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_clave_unidad_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_UNIDADESORGANIZACIONALES where vcha_emp_empresa_id = '" + txt_clave_empresa + "' order by VCHA_UOR_NOMBRE", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_UOR_UNIDAD_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_uor_nombre), "", rs!vcha_uor_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Plantas"
      var_tipo_lista = 2
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_clave_unidad_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_clave_unidad_LostFocus()
   If Trim(txt_clave_unidad) <> "" Then
      rs.Open "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE VCHA_EMP_EMPRESA_ID = '" + txt_clave_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + txt_clave_unidad + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_unidad = rs!vcha_uor_nombre
      Else
         MsgBox "Clave de planta incorrecta", vbOKOnly, "ATENCION"
         txt_clave_unidad = ""
         txt_nombre_unidad = ""
      End If
      rs.Close
   End If
End Sub

Private Sub txt_correo_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_correo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
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
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_colonias where vcha_pai_pais_id = '" + txt_pais + "' and vcha_est_estado_id = '" + txt_estado + "' order by vcha_col_nombre", cnn, adOpenDynamic, adLockOptimistic
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
      If var_n > 6 Then
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

Private Sub txt_codigo_postal_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Trim(txt_codigo_postal) <> "" Then
         rs.Open "select * from vw_colonias where vcha_col_cp = '" + txt_codigo_postal + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            lv_colonias.ListItems.Clear
            While Not rs.EOF
                  Set list_item = lv_colonias.ListItems.Add(, , rs!VCHA_COL_COLONIA_ID)
                  list_item.SubItems(1) = IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
                  list_item.SubItems(2) = IIf(IsNull(rs!vcha_pai_pais_id), "", rs!vcha_pai_pais_id)
                  list_item.SubItems(3) = IIf(IsNull(rs!VCHA_PAI_NOMBRE), "", rs!VCHA_PAI_NOMBRE)
                  list_item.SubItems(4) = IIf(IsNull(rs!vcha_est_estado_id), "", rs!vcha_est_estado_id)
                  list_item.SubItems(5) = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                  list_item.SubItems(6) = IIf(IsNull(rs!vcha_mun_municipio_id), "", rs!vcha_mun_municipio_id)
                  list_item.SubItems(7) = IIf(IsNull(rs!VCHA_MUN_NOMBRE), "", rs!VCHA_MUN_NOMBRE)
                  list_item.SubItems(8) = IIf(IsNull(rs!vcha_ciu_ciudad_id), "", rs!vcha_ciu_ciudad_id)
                  list_item.SubItems(9) = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                  rs.MoveNext
            Wend
            lbl_colonias = "COLONIAS DEL C.P. " + txt_codigo_postal
            Dim var_n As Integer
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
   End If
End Sub

Private Sub txt_codigo_postal_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_colonia) <> "" Then
      rs.Open "select * from TB_COLONIAS where VCHA_COL_COLONIA_ID = '" + txt_colonia + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_colonia = IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
      Else
         MsgBox "Clave de colonias incorrecta", vbOKOnly, "ATENCION"
         txt_colonia = ""
         txt_nombre_colonia = ""
      End If
      rs.Close
   Else
      txt_nombre_colonia = ""
   End If
End Sub

Private Sub txt_direccion_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_direccion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_estado_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
End Sub

Private Sub txt_neteable_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_neteable_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_almacen_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_almacen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_empresa_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Then
      KeyCode = 0
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_empresas order by vcha_emp_empresa_id", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_emp_empresa_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_emp_nombre), "", rs!vcha_emp_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Empresas"
      var_tipo_lista = 1
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_empresa_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 13
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_unidad_Change()
   If KeyCode = 46 Then
      KeyCode = 0
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_UNIDADESORGANIZACIONALES where vcha_emp_empresa_id = '" + txt_clave_empresa + "' order by VCHA_UOR_NOMBRE", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_UOR_UNIDAD_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_uor_nombre), "", rs!vcha_uor_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Plantas"
      var_tipo_lista = 2
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_unidad_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 13
   Case Else
       KeyAscii = 0
   End Select
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_pais_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
End Sub

Private Sub txt_prioridad_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_prioridad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_tipo_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_tipo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_tipo_reempaque_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_tipo_reempaque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub
