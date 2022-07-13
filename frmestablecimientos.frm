VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmestablecimientos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Establecimientos"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   Icon            =   "frmestablecimientos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5955
   Begin VB.Frame frm_cambio_titulares 
      Height          =   1860
      Left            =   150
      TabIndex        =   46
      Top             =   510
      Width           =   5700
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   2265
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   1305
         Width           =   3270
      End
      Begin VB.TextBox txt 
         Height          =   315
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   1305
         Width           =   975
      End
      Begin VB.TextBox txt_nombre_titular_actual 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2265
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   960
         Width           =   3270
      End
      Begin VB.TextBox txt_titular_actual 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmd_aceptar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   60
         Picture         =   "frmestablecimientos.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Aceptar Alt + A"
         Top             =   405
         Width           =   330
      End
      Begin VB.Frame Frame2 
         Height          =   30
         Left            =   0
         TabIndex        =   49
         Top             =   750
         Width           =   5670
      End
      Begin VB.CommandButton cmd_cancelar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   390
         Picture         =   "frmestablecimientos.frx":0A14
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Cancelar Alt + C"
         Top             =   405
         Width           =   330
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Titular Nuevo:"
         Height          =   195
         Left            =   180
         TabIndex        =   52
         Top             =   1365
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Titular Actual:"
         Height          =   195
         Left            =   180
         TabIndex        =   51
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   47
         Top             =   120
         Width           =   5625
      End
   End
   Begin VB.CommandButton cmd_cambiar_titulares 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1470
      Picture         =   "frmestablecimientos.frx":0B5E
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Cambiar de Titular"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   120
      TabIndex        =   40
      Top             =   555
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   41
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
         TabIndex        =   42
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame frm_colonias 
      Height          =   2400
      Left            =   135
      TabIndex        =   37
      Top             =   585
      Width           =   5685
      Begin MSComctlLib.ListView lv_colonias 
         Height          =   1830
         Left            =   45
         TabIndex        =   38
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
         TabIndex        =   39
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5445
      Picture         =   "frmestablecimientos.frx":1050
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1800
      Picture         =   "frmestablecimientos.frx":168A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      Picture         =   "frmestablecimientos.frx":178C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   810
      Picture         =   "frmestablecimientos.frx":188E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      Picture         =   "frmestablecimientos.frx":1960
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Guardar Alt + G"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmestablecimientos.frx":1A62
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2775
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   33
      Top             =   45
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   " Establecimientos "
      Height          =   3960
      Left            =   150
      TabIndex        =   0
      Top             =   420
      Width           =   5655
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5250
         Picture         =   "frmestablecimientos.frx":1B64
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Generar pedido "
         Top             =   3585
         Width           =   330
      End
      Begin VB.CheckBox chk_franquicia 
         Caption         =   "Franquicia"
         Height          =   210
         Left            =   1245
         TabIndex        =   57
         Top             =   3615
         Width           =   1515
      End
      Begin VB.TextBox txt_referencia 
         Height          =   315
         Left            =   3105
         MaxLength       =   50
         TabIndex        =   43
         Top             =   3210
         Width           =   1560
      End
      Begin VB.TextBox txt_nombre_pais 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   2880
         Width           =   4320
      End
      Begin VB.TextBox txt_nombre_estado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2550
         Width           =   4320
      End
      Begin VB.TextBox txt_nombre_ciudad 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1890
         Width           =   4320
      End
      Begin VB.TextBox txt_nombre_colonia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1560
         Width           =   4320
      End
      Begin VB.TextBox txt_nombre_municipio 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   2220
         Width           =   4320
      End
      Begin VB.TextBox txt_municipio 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1245
         TabIndex        =   15
         Top             =   2220
         Width           =   1020
      End
      Begin VB.TextBox txt_codigo_postal 
         Height          =   315
         Left            =   1245
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1230
         Width           =   1020
      End
      Begin VB.TextBox txt_telefono 
         Height          =   315
         Left            =   1245
         MaxLength       =   50
         TabIndex        =   21
         Top             =   3210
         Width           =   1560
      End
      Begin VB.TextBox txt_nombre_establecimiento 
         Height          =   315
         Left            =   1245
         MaxLength       =   50
         TabIndex        =   8
         Top             =   570
         Width           =   4305
      End
      Begin VB.TextBox txt_domicilio 
         Height          =   315
         Left            =   1245
         MaxLength       =   100
         TabIndex        =   9
         Top             =   900
         Width           =   4305
      End
      Begin VB.TextBox txt_colonia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1245
         TabIndex        =   11
         Top             =   1560
         Width           =   1020
      End
      Begin VB.TextBox txt_ciudad 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1245
         TabIndex        =   13
         Top             =   1890
         Width           =   1020
      End
      Begin VB.TextBox txt_estado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   2550
         Width           =   1020
      End
      Begin VB.TextBox txt_pais 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   2880
         Width           =   1020
      End
      Begin VB.TextBox txt_establecimiento 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1245
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   7
         Top             =   240
         Width           =   1140
      End
      Begin MSComctlLib.Toolbar tool_grupos 
         Height          =   330
         Left            =   5145
         TabIndex        =   22
         Top             =   3225
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Detalle de agrupador"
               ImageIndex      =   11
            EndProperty
         EndProperty
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono:"
         Height          =   195
         Index           =   10
         Left            =   1995
         TabIndex        =   44
         Top             =   3285
         Width           =   675
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Municipio:"
         Height          =   195
         Index           =   9
         Left            =   135
         TabIndex        =   36
         Top             =   2280
         Width           =   720
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "C.P.:"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   35
         Top             =   1290
         Width           =   345
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono:"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   31
         Top             =   3270
         Width           =   675
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   8
         Left            =   135
         TabIndex        =   30
         Top             =   630
         Width           =   600
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Colonia:"
         Height          =   195
         Index           =   7
         Left            =   135
         TabIndex        =   29
         Top             =   1620
         Width           =   570
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Domicilio:"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   28
         Top             =   960
         Width           =   675
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad:"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   27
         Top             =   1950
         Width           =   540
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   26
         Top             =   2610
         Width           =   540
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Pais:"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   25
         Top             =   2940
         Width           =   345
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   23
         Top             =   300
         Width           =   450
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2910
      Left            =   150
      TabIndex        =   24
      Top             =   4380
      Width           =   5655
      Begin MSComctlLib.ImageList icono_encabezado 
         Left            =   -285
         Top             =   90
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
               Picture         =   "frmestablecimientos.frx":1C66
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmestablecimientos.frx":2540
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lv_establecimientos 
         Height          =   2700
         Left            =   45
         TabIndex        =   32
         Top             =   135
         Width           =   5565
         _ExtentX        =   9816
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
         NumItems        =   13
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
            Text            =   "municipio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ciudad"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "colonia"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "domicilio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Codigo Postal"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "telefono"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "titular"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "referencia"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Franquicia"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   -60
      Top             =   4125
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
            Picture         =   "frmestablecimientos.frx":2E1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmestablecimientos.frx":36F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmestablecimientos.frx":3FCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmestablecimientos.frx":456A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmestablecimientos.frx":4E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmestablecimientos.frx":5720
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmestablecimientos.frx":5FFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmestablecimientos.frx":610C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmestablecimientos.frx":621E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmestablecimientos.frx":6330
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmestablecimientos.frx":6442
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   34
      Top             =   285
      Width           =   5655
   End
End
Attribute VB_Name = "frmestablecimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim numero_items_establecimientos As Integer
Dim var_bit_pais As String
Dim var_bit_estado As String

Private Sub chk_franquicia_Click()
   var_hubo_cambios = True
End Sub

Private Sub cmd_aceptar_Click()
   Me.frm_cambio_titulares.Visible = False
End Sub

Private Sub cmd_cambiar_titulares_Click()
   If Trim(Me.txt_establecimiento) <> "" Then
      frmcambiar_titular.txt_tipo = 2
      frmcambiar_titular.txt_clave_establecimiento = Me.txt_establecimiento
      frmcambiar_titular.Show 1
      Call pro_limpiatextos(Me)
      lv_establecimientos.ListItems.Clear
      Call pro_llena_listview1
   Else
      MsgBox "No se a seleccionado un establecimiento", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_cancelar_Click()
   Me.frm_cambio_titulares.Visible = False
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
         'rsaux5.Open "select vcha_emp_conexion from tb_empresas where len(vcha_emp_conexion) > 0", cnn_distribucion, adOpenDynamic, adLockOptimistic
         'While Not rsaux5.EOF
         '      var_conexion_importacion = IIf(IsNull(rsaux5(0).Value), "", rsaux5(0).Value)
         '      If Trim(var_conexion_importacion) <> "" Then
         '         If cnn.State = 1 Then
         '            cnn.Close
         '         End If
         '         cnn.Open var_conexion_importacion
                  Call pro_elimina_establecimientos
         '      End If
         '      rsaux5.MoveNext
         'Wend
         'rsaux5.Close
         
         numero_items_establecimientos = numero_items_establecimientos - 1
         MsgBox "Se Elimino Correctamente el Registro", vbInformation
         lv_establecimientos.ListItems.Remove (lv_establecimientos.selectedItem.Index)
         Call pro_limpiatextos(Me)
         txt_registros = lv_establecimientos.ListItems.Count
         If lv_establecimientos.ListItems.Count > 0 Then
            lv_establecimientos.selectedItem.Selected = True
         End If
         pro_textos
      
      
      
         rs.Open "select * from tb_establecimientos", cnn_distribucion, adOpenDynamic, adLockOptimistic
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
End Sub

Private Sub cmd_guardar_Click()
Dim var_posible As Boolean
   var_posible = True
   If var_modifica_regsitro_establecimientos = False Then
      rs.Open "select * from tb_establecimientos where vcha_esb_establecimiento_id =  '" + Me.txt_establecimiento + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
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
         If rsaux5.State = 1 Then
            rsaux5.Close
         End If
         rsaux5.Open "select vcha_emp_conexion from tb_empresas where len(vcha_emp_conexion) > 0 and vcha_emp_empresa_id = '02' ORDER BY INTE_EMP_ORDEN_CONEXION", cnn_distribucion, adOpenDynamic, adLockOptimistic
         While Not rsaux5.EOF
               var_conexion_importacion = IIf(IsNull(rsaux5(0).Value), "", rsaux5(0).Value)
               If Trim(var_conexion_importacion) <> "" Then
                  If cnn_importacion.State = 1 Then
                     cnn_importacion.Close
                  End If
                  cnn_importacion.Open var_conexion_importacion
                  Call pro_guardar_establecimientos
               End If
               rsaux5.MoveNext
         Wend
         rsaux5.Close
         If Trim(var_establecimiento_regreso) <> "" Then
            txt_establecimiento = var_establecimiento_regreso
         End If
         
         If (UCase(parametros(0)) = "sqlquezada2" Or UCase(parametros(0)) = "DBPRUEBAS") And var_empresa = "31" Then
            If var_cliente_pedido_internet <> "" Then
               cnn_distribucion.BeginTrans
               cnn.BeginTrans
               rs.Open "SELECT * FROM tb_Establecimientos WHERE vcha_esb_establecimiento_id = '" + Me.txt_establecimiento + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
               If rs.EOF Then
                  var_cadena = "INSERT INTO TB_ESTABLECIMIENTOS (VCHA_TIT_TITULAR_ID, VCHA_ESB_ESTABLECIMIENTO_ID, VCHA_ESB_NOMBRE, VCHA_PAI_PAIS_ID, VCHA_EST_ESTADO_ID, VCHA_CIU_CIUDAD_ID, VCHA_COL_COLONIA_ID, VCHA_ESB_DOMICILIO, VCHA_ESB_TELEFONO, CHAR_ESB_FACTURA_CATALOGOS, VCHA_MUN_MUNICIPIO_ID, VCHA_ESB_CP, vcha_emp_empresa_id)"
                  var_cadena = var_cadena + " Values ('" + vartitular + "','" + txt_establecimiento + "', '" + Me.txt_nombre_establecimiento + "','" + Me.txt_pais + "', '" + Me.txt_estado + "','" + Me.txt_ciudad + "', '" + Me.txt_colonia + "','" + Me.txt_domicilio + "', '" + Me.txt_telefono + "',0,'" + Me.txt_municipio + "','" + Me.txt_codigo_postal + "','" + var_empresa + "')"
                  rsaux1.Open var_cadena, cnn_distribucion, adOpenDynamic, adLockOptimistic
               Else
                  var_cadena = "Update tb_establecimientos set VCHA_TIT_TITULAR_ID = '" + vartitular + "', VCHA_ESB_NOMBRE = '" + Me.txt_nombre_establecimiento + "', VCHA_PAI_PAIS_ID = '" + Me.txt_pais + "', VCHA_EST_ESTADO_ID = '" + Me.txt_estado + "', VCHA_CIU_CIUDAD_ID = '" + Me.txt_ciudad + "', VCHA_COL_COLONIA_ID = '" + Me.txt_colonia + "', VCHA_ESB_DOMICILIO = '" + Me.txt_domicilio + "', VCHA_ESB_TELEFONO = '" + Me.txt_telefono + "', CHAR_ESB_FACTURA_CATALOGOS = '', VCHA_MUN_MUNICIPIO_ID = '" + Me.txt_municipio + "', VCHA_ESB_CP = '" + Me.txt_codigo_postal + "' where VCHA_ESB_ESTABLECIMIENTO_ID = '" + Me.txt_establecimiento + "'"
                  rsaux1.Open var_cadena, cnn_distribucion, adOpenDynamic, adLockOptimistic
               End If
               rs.Close
               rs.Open "select * from tb_detalle_establecimientos where vcha_Esb_Establecimiento_id = '" + Me.txt_establecimiento + "' and vcha_cli_clave_id = '" + var_cliente_pedido_internet + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
               If rs.EOF Then
                  rsaux.Open "insert into tb_Detalle_establecimientos (vcha_Cli_clave_id, vcha_esb_establecimiento_id) values ('" + var_cliente_pedido_internet + "', '" + Me.txt_establecimiento + "' )", cnn_distribucion, adOpenDynamic, adLockOptimistic
               End If
               rs.Close
               
               rs.Open "select * from tb_detalle_establecimientos where vcha_Esb_Establecimiento_id = '" + Me.txt_establecimiento + "' and vcha_cli_clave_id = '" + var_cliente_pedido_internet + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
               If rs.EOF Then
                  rsaux.Open "insert into tb_Detalle_establecimientos (vcha_Cli_clave_id, vcha_esb_establecimiento_id) values ('" + var_cliente_pedido_internet + "', '" + Me.txt_establecimiento + "' )", cnn_distribucion, adOpenDynamic, adLockOptimistic
               End If
               rs.Close
               
               cnn.CommitTrans
               cnn_distribucion.CommitTrans
            End If
         End If
         
         pro_actualiza_ListView
         txt_establecimiento.Enabled = False
         MsgBox "Informacion Guardada Correctamente!", vbOKOnly + vbInformation, "Aviso"
         txt_registros = lv_establecimientos.ListItems.Count
         var_modifica_regsitro_establecimientos = True
         var_hubo_cambios = False
         
         rs.Open "select * from tb_establecimientos", cnn_distribucion, adOpenDynamic, adLockOptimistic
         If rs.BOF Then
            tool_grupos.Enabled = False
            cmd_guardar.Enabled = False
            cmd_deshacer.Enabled = False
            cmd_eliminar.Enabled = False
         Else
            cmd_guardar.Enabled = True
            cmd_deshacer.Enabled = True
            cmd_eliminar.Enabled = True
            tool_grupos.Enabled = True
         End If
         rs.Close
      End If
   Else
      MsgBox "Clave de establecimiento ya existe", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call gPrintListView(lv_establecimientos, "LISTADO DE establecimientos")
        End If

End Sub

Private Sub cmd_nuevo_Click()
        Call pro_limpiatextos(Me)
        txt_establecimiento.Enabled = False
        txt_nombre_establecimiento.Enabled = True
        txt_nombre_establecimiento.SetFocus: var_modifica_regsitro_establecimientos = False
        cmd_guardar.Enabled = True
        cmd_deshacer.Enabled = True
        txt_pais.Enabled = False
        txt_estado.Enabled = False
        txt_domicilio.Enabled = True
        txt_telefono.Enabled = True
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_regsitro_establecimientos = False Then
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
      var_tipo_datos_adicionales = 2
      var_hubo_cambios = True
      If Trim(Me.txt_establecimiento) <> "" Then
         rs.Open "select * from tb_establecimientos where vcha_esb_establecimiento_id = '" + Me.txt_establecimiento + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_clave_establecimiento_global = Me.txt_establecimiento
            var_nombre_cliente_ad = IIf(IsNull(rs!vcha_Esb_nombre_2), "", rs!vcha_Esb_nombre_2)
            var_paterno_cliente_ad = IIf(IsNull(rs!vcha_esb_paterno), "", rs!vcha_esb_paterno)
            var_materno_cliente_ad = IIf(IsNull(rs!vcha_Esb_materno), "", rs!vcha_Esb_materno)
            var_numero_cliente_ad = IIf(IsNull(rs!vcha_esb_numero), "", rs!vcha_esb_numero)
            var_clave_tel_pais_ad = IIf(IsNull(rs!vcha_esb_clave_tel_pais), "", rs!vcha_esb_clave_tel_pais)
            var_clave_tel_estado_ad = IIf(IsNull(rs!vcha_esb_clave_tel_estado), "", rs!vcha_esb_clave_tel_estado)
            var_calle_cliente_ad = IIf(IsNull(rs!vcha_Esb_calle), "", rs!vcha_Esb_calle)
            var_numero_interno_cliente_ad = IIf(IsNull(rs!vcha_esb_numero_interno), "", rs!vcha_esb_numero_interno)
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
   frm_colonias.Visible = False
   frm_lista.Visible = False
   Me.frm_cambio_titulares.Visible = False
   vartitular = frmlistatitulares.lv_listatitulares.selectedItem
   'vartitular = 9
   var_modifica_regsitro_establecimientos = True
   lv_establecimientos.SmallIcons = ImageList
   'Call pro_encabezadosView(Me, lv_establecimientos, False)
   If var_cliente_pedido_internet <> "" Then
      rs.Open "select a.vcha_esb_establecimiento_id, b.vcha_esb_nombre from TB_DETALLE_establecimientos a, TB_establecimientos b where a.vcha_cli_clave_id = '" & var_cliente_pedido_internet & "' and a.vcha_esb_establecimiento_id = b.vcha_esb_establecimiento_id", cnn_distribucion, adOpenDynamic, adLockOptimistic
   Else
      rs.Open "select * from tb_establecimientos where vcha_tit_titular_id = '" + vartitular + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
   End If
   If Not rs.EOF Then
      cmd_guardar.Enabled = True
      cmd_deshacer.Enabled = True
      cmd_eliminar.Enabled = True
      txt_nombre_establecimiento.Enabled = True
      txt_pais.Enabled = False
      txt_estado.Enabled = False
      txt_domicilio.Enabled = True
      txt_telefono.Enabled = True
      tool_grupos.Enabled = True
      rs.Close
      Call pro_llena_listview1
      pro_textos
   Else
      tool_grupos.Enabled = False
      txt_establecimiento.Enabled = False
      txt_nombre_establecimiento.Enabled = False
      txt_pais.Enabled = False
      txt_estado.Enabled = False
      txt_domicilio.Enabled = False
      txt_telefono.Enabled = False
      rs.Close
      cmd_guardar.Enabled = False
      cmd_deshacer.Enabled = False
      cmd_eliminar.Enabled = False
   End If
   If var_cliente_pedido_internet <> "" Then
      Me.cmd_cambiar_titulares.Enabled = False
      Me.tool_grupos.Enabled = False
      Me.chk_franquicia.Enabled = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    var_swpassword = False
    var_modifica_regsitro_establecimientos = False
    Call activa_forma(var_activa_forma_establecimientos)
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
      txt_telefono.SetFocus
      frm_colonias.Visible = False
   End If
   If KeyAscii = 27 Then
      frm_colonias.Visible = False
      Me.txt_codigo_postal.SetFocus
   End If
End Sub

Private Sub lv_establecimientos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_establecimientos, ColumnHeader)
End Sub

Private Sub lv_establecimientos_ItemClick(ByVal item As MSComctlLib.ListItem)
   Set lv_establecimientos.selectedItem = item
   pro_textos
   var_modifica_regsitro_establecimientos = True
End Sub



Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
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
         frm_lista.Visible = False
      Else
         frm_lista.Visible = False
      End If
   End If
   If KeyAscii = 27 Then
      Me.txt_codigo_postal.SetFocus
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub tool_grupos_ButtonClick(ByVal Button As MSComctlLib.Button)
   frmdetalle_establecimientos.Caption = "Detalle de establecimientos de:  " + lv_establecimientos.selectedItem.SubItems(1)
   frmdetalle_establecimientos.Show
End Sub


Sub pro_guardar_establecimientos()
   Dim ok As Boolean
   Set TB_ESTABLECIMIENTOS = New TB_ESTABLECIMIENTOS
   Set TB_BITACORA_ESTABLECIMIENTOS = New TB_BITACORA_ESTABLECIMIENTOS
   ok = True
   If txt_nombre_establecimiento <> "" And Me.txt_nombre_establecimiento <> "" Then
      If var_hubo_cambios Then
         rs.Open "select * from tb_establecimientos where vcha_esb_establecimiento_id = '" + txt_establecimiento + "'", cnn_distribucion, adOpenDynamic, adLockBatchOptimistic
         var_establecimiento_regreso = txt_establecimiento
         ok = TB_ESTABLECIMIENTOS.Anadir(vartitular, txt_establecimiento, txt_nombre_establecimiento, txt_pais, txt_estado, txt_ciudad, txt_colonia, Trim(txt_domicilio), txt_telefono, "", txt_municipio, txt_codigo_postal)
         rsaux8.Open "update tb_establecimientos set INTE_ESB_FRANQUICIA = " + CStr(Me.chk_franquicia) + " where vcha_Esb_establecimiento_id = '" + Me.txt_establecimiento + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
         If Trim(var_establecimiento_regreso) <> "" Then
            txt_establecimiento = var_establecimiento_regreso
         End If
         If Trim(Me.txt_referencia) = "" Then
            Me.txt_referencia = Me.txt_establecimiento
         End If
         If ok Then
            rsaux4.Open "update tb_establecimientos set vcha_esb_establecimiento_anterior_id = '" + txt_referencia + "', vcha_emp_empresa_id = '" + var_empresa + "' where vcha_esb_establecimiento_id = '" + Me.txt_establecimiento + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
            bitacora = True
            If var_modifica_regsitro_establecimientos = False Then
               var_operacion_bitacora = "I"
               bitacora = TB_BITACORA_ESTABLECIMIENTOS.Anadir(vartitular, txt_establecimiento, "VCHA_ESB_NOMBRE", var_operacion_bitacora, "", txt_nombre_establecimiento, var_clave_usuario_global, fun_NombrePc, Date)
            Else
               If Not rs.EOF Then
                  var_operacion_bitacora = "M"
                  If rs!vcha_ESB_ESTABLECIMIENTO_id <> txt_establecimiento Then
                     bitacora = TB_BITACORA_ESTABLECIMIENTOS.Anadir(vartitular, txt_establecimiento, "VCHA_ESB_ESTABLECIMIENTO_ID", var_operacion_bitacora, rs!vcha_ESB_ESTABLECIMIENTO_id, txt_establecimiento, var_clave_usuario_global, fun_NombrePc, Date)
                  End If
                  If rs!VCHA_ESB_NOMBRE <> txt_nombre_establecimiento Then
                     bitacora = TB_BITACORA_ESTABLECIMIENTOS.Anadir(vartitular, txt_establecimiento, "VCHA_ESB_NOMBRE", var_operacion_bitacora, rs!VCHA_ESB_NOMBRE, txt_nombre_establecimiento, var_clave_usuario_global, fun_NombrePc, Date)
                  End If
                  If rs!VCHA_PAI_PAIS_ID <> txt_pais Then
                     bitacora = TB_BITACORA_ESTABLECIMIENTOS.Anadir(vartitular, txt_establecimiento, "VCHA_PAI_PAIS_ID", var_operacion_bitacora, rs!VCHA_PAI_PAIS_ID, var_bit_pais, var_clave_usuario_global, fun_NombrePc, Date)
                  End If
                  If rs!VCHA_EST_ESTADO_ID <> txt_estado Then
                     bitacora = TB_BITACORA_ESTABLECIMIENTOS.Anadir(vartitular, txt_establecimiento, "VCHA_EST_ESTADO_ID", var_operacion_bitacora, rs!VCHA_EST_ESTADO_ID, var_bit_estado, var_clave_usuario_global, fun_NombrePc, Date)
                  End If
                  If rs!VCHA_CIU_CIUDAD_ID <> txt_ciudad Then
                     bitacora = TB_BITACORA_ESTABLECIMIENTOS.Anadir(vartitular, txt_establecimiento, "VCHA_CIU_CIUDAD_ID", var_operacion_bitacora, rs!VCHA_CIU_CIUDAD_ID, txt_ciudad, var_clave_usuario_global, fun_NombrePc, Date)
                  End If
                  If rs!VCHA_COL_COLONIA_ID <> txt_colonia Then
                     bitacora = TB_BITACORA_ESTABLECIMIENTOS.Anadir(vartitular, txt_establecimiento, "VCHA_COL_COLONIA_ID", var_operacion_bitacora, rs!VCHA_COL_COLONIA_ID, txt_colonia, var_clave_usuario_global, fun_NombrePc, Date)
                  End If
                  If rs!vcha_esb_domicilio <> txt_domicilio Then
                     bitacora = TB_BITACORA_ESTABLECIMIENTOS.Anadir(vartitular, txt_establecimiento, "VCHA_ESB_DOMICILIO", var_operacion_bitacora, rs!vcha_esb_domicilio, txt_domicilio, var_clave_usuario_global, fun_NombrePc, Date)
                  End If
                  If rs!vcha_esb_telefono <> txt_telefono Then
                     bitacora = TB_BITACORA_ESTABLECIMIENTOS.Anadir(vartitular, txt_establecimiento, "VCHA_ESB_TELEFONO", var_operacion_bitacora, rs!vcha_esb_telefono, txt_telefono, var_clave_usuario_global, fun_NombrePc, Date)
                  End If
               End If
            End If
            rs.Close
         Else
            MsgBox "No se puede grabar registro: " + TB_ESTABLECIMIENTOS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
         End If
      End If
   End If
   Set TB_ESTABLECIMIENTOS = Nothing
End Sub

Sub pro_elimina_establecimientos()
   Dim var_llave_usuarios As String
   Set TB_ESTABLECIMIENTOS = New TB_ESTABLECIMIENTOS
   Set TB_BITACORA_ESTABLECIMIENTOS = New TB_BITACORA_ESTABLECIMIENTOS
   On Error GoTo salir:
   ok = True
   If txt_establecimiento <> "" And txt_nombre_establecimiento <> "" And var_modifica_regsitro_establecimientos = True Then
      'If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_ESTABLECIMIENTOS.Eliminar(vartitular, txt_establecimiento)
      'Else
      '   GoTo salir:
      'End If
      If ok Then
         bitacora = True
         var_operacion_bitacora = "E"
         bitacora = TB_BITACORA_ESTABLECIMIENTOS.Anadir(vartitular, txt_establecimiento, "VCHA_ESB_NOMBRE", var_operacion_bitacora, "", txt_nombre_establecimiento, var_clave_usuario_global, fun_NombrePc, Date)
      Else
        MsgBox "No se puede eliminar registro: " + TB_ESTABLECIMIENTOS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
   Set TB_ESTABLECIMIENTOS = Nothing
End Sub


Sub pro_llena_listview1()
   Dim list_item As ListItem
   If var_cliente_pedido_internet <> "" Then
      rs.Open "SELECT b.vcha_tit_titular_id, a.VCHA_ESB_ESTABLECIMIENTO_ID, b.VCHA_ESB_NOMBRE, b.VCHA_PAI_PAIS_ID, b.VCHA_EST_ESTADO_ID, b.VCHA_CIU_CIUDAD_ID,  b.VCHA_COL_COLONIA_ID, b.VCHA_ESB_DOMICILIO, b.VCHA_ESB_TELEFONO, b.CHAR_ESB_FACTURA_CATALOGOS, b.VCHA_MUN_MUNICIPIO_ID, b.VCHA_ESB_CP, b.VCHA_ESB_ESTABLECIMIENTO_ANTERIOR_ID, b.VCHA_EMP_EMPRESA_ID, b.TEXTILERA, b.DTIM_INT_FECHA, b.INTE_INT_INTERFACE , b.INTE_ESB_FRANQUICIA, b.VCHA_UOR_UNIDAD_ID FROM dbo.TB_DETALLE_ESTABLECIMIENTOS a INNER JOIN dbo.TB_ESTABLECIMIENTOS b ON a.VCHA_ESB_ESTABLECIMIENTO_ID = b.VCHA_ESB_ESTABLECIMIENTO_ID wHERE (a.VCHA_CLI_CLAVE_ID = '" + var_cliente_pedido_internet + "') ", cnn_distribucion, adOpenDynamic, adLockOptimistic
   Else
      rs.Open "SELECT * FROM TB_ESTABLECIMIENTOS Where vcha_tit_titular_id = '" + vartitular + "'  and vcha_emp_empresa_id = '" + var_empresa + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
   End If
   numero_items_establecimientos = 0
   While Not rs.EOF
      Set list_item = lv_establecimientos.ListItems.Add(, , rs!vcha_ESB_ESTABLECIMIENTO_id)
      list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE)
      list_item.SubItems(2) = IIf(IsNull(rs!VCHA_PAI_PAIS_ID), "", rs!VCHA_PAI_PAIS_ID)
      list_item.SubItems(3) = IIf(IsNull(rs!VCHA_EST_ESTADO_ID), "", rs!VCHA_EST_ESTADO_ID)
      list_item.SubItems(4) = IIf(IsNull(rs!VCHA_MUN_MUNICIPIO_ID), "", rs!VCHA_MUN_MUNICIPIO_ID)
      list_item.SubItems(5) = IIf(IsNull(rs!VCHA_CIU_CIUDAD_ID), "", rs!VCHA_CIU_CIUDAD_ID)
      list_item.SubItems(6) = IIf(IsNull(rs!VCHA_COL_COLONIA_ID), "", rs!VCHA_COL_COLONIA_ID)
      list_item.SubItems(7) = IIf(IsNull(rs!vcha_esb_domicilio), "", rs!vcha_esb_domicilio)
      list_item.SubItems(8) = IIf(IsNull(rs!vcha_esb_cp), "", rs!vcha_esb_cp)
      list_item.SubItems(9) = IIf(IsNull(rs!vcha_esb_telefono), "", rs!vcha_esb_telefono)
      list_item.SubItems(10) = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
      list_item.SubItems(11) = IIf(IsNull(rs!vcha_esb_establecimiento_anterior_id), "", rs!vcha_esb_establecimiento_anterior_id)
      list_item.SubItems(12) = IIf(IsNull(rs!INTE_ESB_FRANQUICIA), 0, rs!INTE_ESB_FRANQUICIA)
      rs.MoveNext:
      numero_items_establecimientos = numero_items_establecimientos + 1
    Wend
    rs.Close
End Sub


Sub pro_textos()
On Error GoTo err0:
Dim var_n As Double
var_n = lv_establecimientos.ListItems.Count
   If var_n > 0 Then
      txt_establecimiento = lv_establecimientos.selectedItem
      txt_nombre_establecimiento = lv_establecimientos.selectedItem.SubItems(1)
      txt_pais = lv_establecimientos.selectedItem.SubItems(2)
      txt_estado = lv_establecimientos.selectedItem.SubItems(3)
      txt_municipio = lv_establecimientos.selectedItem.SubItems(4)
      txt_ciudad = lv_establecimientos.selectedItem.SubItems(5)
      txt_colonia = lv_establecimientos.selectedItem.SubItems(6)
      txt_domicilio = lv_establecimientos.selectedItem.SubItems(7)
      txt_codigo_postal = lv_establecimientos.selectedItem.SubItems(8)
      txt_telefono = lv_establecimientos.selectedItem.SubItems(9)
      txt_referencia = lv_establecimientos.selectedItem.SubItems(11)
      Me.chk_franquicia = Me.lv_establecimientos.selectedItem.SubItems(12)
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
   End If
   var_numero_renglones = lv_establecimientos.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_establecimientos.ColumnHeaders(2).Width = 3850
   Else
      lv_establecimientos.ColumnHeaders(2).Width = 4099.9
   End If
   var_modifica_regsitro_establecimientos = True
   var_hubo_cambios = False
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem

    If var_modifica_regsitro_establecimientos = False Then
        Set list_item = lv_establecimientos.ListItems.Add(, , txt_establecimiento)
        list_item.SubItems(1) = txt_nombre_establecimiento
        list_item.SubItems(2) = txt_pais
        list_item.SubItems(3) = txt_estado
        list_item.SubItems(4) = txt_municipio
        list_item.SubItems(5) = txt_ciudad
        list_item.SubItems(6) = txt_colonia
        list_item.SubItems(7) = txt_domicilio
        list_item.SubItems(8) = txt_codigo_postal
        list_item.SubItems(9) = txt_telefono
        list_item.SubItems(10) = vartitular
        list_item.SubItems(11) = txt_referencia
        list_item.SubItems(12) = Me.chk_franquicia
        list_item.EnsureVisible
        list_item.Selected = True
        numero_items_establecimientos = numero_items_establecimientos + 1
    Else
        lv_establecimientos.ListItems.item(lv_establecimientos.selectedItem.Index).Checked = False
        lv_establecimientos.ListItems.item(lv_establecimientos.selectedItem.Index) = txt_establecimiento
        lv_establecimientos.ListItems.item(lv_establecimientos.selectedItem.Index).ListSubItems(1) = txt_nombre_establecimiento
        lv_establecimientos.ListItems.item(lv_establecimientos.selectedItem.Index).ListSubItems(2) = txt_pais
        lv_establecimientos.ListItems.item(lv_establecimientos.selectedItem.Index).ListSubItems(3) = txt_estado
        lv_establecimientos.ListItems.item(lv_establecimientos.selectedItem.Index).ListSubItems(4) = txt_municipio
        lv_establecimientos.ListItems.item(lv_establecimientos.selectedItem.Index).ListSubItems(5) = txt_ciudad
        lv_establecimientos.ListItems.item(lv_establecimientos.selectedItem.Index).ListSubItems(6) = txt_colonia
        lv_establecimientos.ListItems.item(lv_establecimientos.selectedItem.Index).ListSubItems(7) = txt_domicilio
        lv_establecimientos.ListItems.item(lv_establecimientos.selectedItem.Index).ListSubItems(8) = txt_codigo_postal
        lv_establecimientos.ListItems.item(lv_establecimientos.selectedItem.Index).ListSubItems(9) = txt_telefono
        lv_establecimientos.ListItems.item(lv_establecimientos.selectedItem.Index).ListSubItems(10) = vartitular
        lv_establecimientos.ListItems.item(lv_establecimientos.selectedItem.Index).ListSubItems(11) = txt_referencia
        lv_establecimientos.ListItems.item(lv_establecimientos.selectedItem.Index).ListSubItems(12) = Me.chk_franquicia
        lv_establecimientos.ListItems.item(lv_establecimientos.selectedItem.Index).Selected = True
    End If
'    lv_establecimientos.SetFocus
End Sub




Private Sub txt_ciudad_Change()
   var_hubo_cambios = True
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
      var_activa_forma_direcciones = Me.Name
      frmestablecimientos.Enabled = False
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
            var_tipo_lista = 1
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

Private Sub txt_colonia_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para accesar al catálogo"
End Sub

Private Sub txt_colonia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_colonia_LostFocus()
   If Trim(txt_colonia) <> "" Then
      rs.Open "select * from tb_colonias where vcha_col_colonia_id = '" + txt_colonia + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_colonia = IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
      Else
         txt_nombre_colonia = ""
         txt_colonia = ""
         MsgBox "Clave de colonia incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      txt_nombre_colonia = ""
   End If
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

Private Sub txt_establecimiento_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_establecimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_estado_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_estado_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_municipio_Change()
   var_hubo_cambios = True
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

Private Sub txt_nombre_colonia_KeyDown(KeyCode As Integer, Shift As Integer)
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para accesar al catálogo"
End Sub

Private Sub txt_nombre_colonia_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_establecimiento_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_establecimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
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

Private Sub txt_pais_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_pais_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_referencia_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_telefono_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_telefono_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Me.cmd_guardar.Enabled = True Then
         Me.cmd_guardar.SetFocus
      End If
   End If
End Sub

