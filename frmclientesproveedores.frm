VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#11.0#0"; "VBSKFREE.OCX"
Begin VB.Form frmclientesproveedores 
   Caption         =   "Clientes y/o proveedores"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9735
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   8595
   ScaleWidth      =   9735
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   135
      TabIndex        =   31
      Top             =   15
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo Registro"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Guardar Registro"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Deshacer cambios"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar Registro"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Lista"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir de Esta Ventana"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   75
      TabIndex        =   32
      Top             =   285
      Width           =   9615
   End
   Begin VB.Frame Frame3 
      Height          =   3885
      Left            =   120
      TabIndex        =   29
      Top             =   4665
      Width           =   9510
      Begin MSComctlLib.ListView lv_clientesproveedores 
         Height          =   3660
         Left            =   60
         TabIndex        =   30
         Top             =   150
         Width           =   9360
         _ExtentX        =   16510
         _ExtentY        =   6456
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   18
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   13758
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "representante"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "proveedor"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "fecha"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "vendedor"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ruta"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "curp"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "rfc"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "moneda"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "plazo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "tipo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "lista"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "canal"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "descripcion"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "transporte"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "estatus"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "TITULAR"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Clientes y/o proveedores"
      Height          =   3555
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   9540
      Begin MSComCtl2.MonthView mes 
         Height          =   2370
         Index           =   0
         Left            =   5955
         TabIndex        =   11
         Top             =   705
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   50724865
         CurrentDate     =   37581
      End
      Begin VB.ComboBox cmb_clientesproveedores 
         Height          =   315
         Index           =   15
         Left            =   1515
         TabIndex        =   52
         Top             =   3135
         Width           =   2325
      End
      Begin VB.ComboBox cmb_clientesproveedores 
         Height          =   315
         Index           =   14
         Left            =   6150
         TabIndex        =   51
         Top             =   2790
         Width           =   2325
      End
      Begin VB.ComboBox cmb_clientesproveedores 
         Height          =   315
         Index           =   13
         Left            =   1515
         TabIndex        =   50
         Top             =   2805
         Width           =   3255
      End
      Begin VB.ComboBox cmb_clientesproveedores 
         Height          =   315
         Index           =   12
         Left            =   7365
         TabIndex        =   49
         Top             =   2475
         Width           =   2085
      End
      Begin VB.ComboBox cmb_clientesproveedores 
         Height          =   315
         Index           =   11
         Left            =   3840
         TabIndex        =   48
         Top             =   2475
         Width           =   2325
      End
      Begin VB.ComboBox cmb_clientesproveedores 
         Height          =   315
         Index           =   10
         Left            =   1515
         TabIndex        =   47
         Top             =   2475
         Width           =   1785
      End
      Begin VB.ComboBox cmb_clientesproveedores 
         Height          =   315
         Index           =   9
         Left            =   7590
         TabIndex        =   46
         Top             =   2145
         Width           =   1845
      End
      Begin VB.ComboBox cmb_clientesproveedores 
         Height          =   315
         Index           =   6
         Left            =   1515
         TabIndex        =   45
         Top             =   1830
         Width           =   4185
      End
      Begin VB.ComboBox cmb_clientesproveedores 
         Height          =   315
         Index           =   5
         Left            =   1515
         TabIndex        =   44
         Top             =   1515
         Width           =   4185
      End
      Begin VB.CheckBox chk_clientesproveedores 
         Caption         =   "Estatus"
         Height          =   210
         Index           =   16
         Left            =   4080
         TabIndex        =   43
         Top             =   3195
         Width           =   855
      End
      Begin VB.CheckBox chk_clientesproveedores 
         Caption         =   "Proveedor"
         Height          =   210
         Index           =   3
         Left            =   1515
         TabIndex        =   42
         Top             =   1260
         Width           =   1170
      End
      Begin VB.CommandButton cmdfecha 
         Height          =   285
         Index           =   0
         Left            =   5640
         Picture         =   "frmclientesproveedores.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Seleccione la fecha"
         Top             =   1230
         Width           =   315
      End
      Begin VB.TextBox txt_clientesproveedores 
         Height          =   285
         Index           =   13
         Left            =   1545
         MaxLength       =   50
         TabIndex        =   22
         Top             =   2835
         Width           =   2760
      End
      Begin VB.TextBox txt_clientesproveedores 
         Height          =   285
         Index           =   9
         Left            =   7605
         MaxLength       =   50
         TabIndex        =   18
         Top             =   2160
         Width           =   1830
      End
      Begin VB.TextBox txt_clientesproveedores 
         Height          =   285
         Index           =   14
         Left            =   6135
         MaxLength       =   50
         TabIndex        =   23
         Top             =   2805
         Width           =   2325
      End
      Begin VB.TextBox txt_clientesproveedores 
         Height          =   285
         Index           =   12
         Left            =   7440
         MaxLength       =   50
         TabIndex        =   21
         Top             =   2475
         Width           =   1830
      End
      Begin VB.TextBox txt_clientesproveedores 
         Height          =   285
         Index           =   15
         Left            =   1530
         MaxLength       =   50
         TabIndex        =   24
         Top             =   3150
         Width           =   2085
      End
      Begin VB.TextBox txt_clientesproveedores 
         Height          =   285
         Index           =   8
         Left            =   4425
         MaxLength       =   50
         TabIndex        =   17
         Top             =   2160
         Width           =   2325
      End
      Begin VB.TextBox txt_clientesproveedores 
         Height          =   285
         Index           =   10
         Left            =   1530
         MaxLength       =   50
         TabIndex        =   19
         Top             =   2490
         Width           =   1545
      End
      Begin VB.TextBox txt_clientesproveedores 
         Height          =   285
         Index           =   11
         Left            =   3900
         MaxLength       =   50
         TabIndex        =   20
         Top             =   2505
         Width           =   1455
      End
      Begin VB.TextBox txt_clientesproveedores 
         Height          =   285
         Index           =   0
         Left            =   1545
         MaxLength       =   3
         TabIndex        =   3
         Top             =   300
         Width           =   1320
      End
      Begin VB.TextBox txt_clientesproveedores 
         Height          =   285
         Index           =   1
         Left            =   1530
         MaxLength       =   50
         TabIndex        =   5
         Top             =   630
         Width           =   7950
      End
      Begin VB.TextBox txt_clientesproveedores 
         Height          =   285
         Index           =   2
         Left            =   1530
         MaxLength       =   50
         TabIndex        =   7
         Top             =   930
         Width           =   7950
      End
      Begin VB.TextBox txt_clientesproveedores 
         Height          =   285
         Index           =   5
         Left            =   1530
         MaxLength       =   50
         TabIndex        =   13
         Top             =   1515
         Width           =   1605
      End
      Begin VB.TextBox txt_clientesproveedores 
         Height          =   285
         Index           =   6
         Left            =   1530
         MaxLength       =   50
         TabIndex        =   14
         Top             =   1830
         Width           =   1605
      End
      Begin VB.TextBox txt_clientesproveedores 
         Height          =   285
         Index           =   7
         Left            =   1530
         MaxLength       =   50
         TabIndex        =   16
         Top             =   2160
         Width           =   2325
      End
      Begin VB.TextBox txt_clientesproveedores 
         Height          =   285
         Index           =   4
         Left            =   4380
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   9
         Top             =   1230
         Width           =   1230
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Transporte:"
         Height          =   195
         Index           =   13
         Left            =   675
         TabIndex        =   41
         Top             =   3195
         Width           =   810
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Fam. agrupadores:"
         Height          =   195
         Index           =   4
         Left            =   4830
         TabIndex        =   40
         Top             =   2850
         Width           =   1320
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Moneda:"
         Height          =   195
         Index           =   15
         Left            =   6960
         TabIndex        =   39
         Top             =   2205
         Width           =   630
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "CURP:"
         Height          =   195
         Index           =   14
         Left            =   1020
         TabIndex        =   38
         Top             =   2205
         Width           =   495
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Plazo:"
         Height          =   195
         Index           =   12
         Left            =   1065
         TabIndex        =   37
         Top             =   2535
         Width           =   435
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Canal de venta:"
         Height          =   195
         Index           =   11
         Left            =   375
         TabIndex        =   36
         Top             =   2850
         Width           =   1125
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Lista de precios:"
         Height          =   195
         Index           =   10
         Left            =   6195
         TabIndex        =   35
         Top             =   2535
         Width           =   1155
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Index           =   9
         Left            =   3405
         TabIndex        =   34
         Top             =   2535
         Width           =   360
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "RFC:"
         Height          =   195
         Index           =   8
         Left            =   4050
         TabIndex        =   33
         Top             =   2190
         Width           =   360
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   0
         Left            =   1050
         TabIndex        =   25
         Top             =   345
         Width           =   450
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   1
         Left            =   900
         TabIndex        =   15
         Top             =   630
         Width           =   600
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Representante:"
         Height          =   195
         Index           =   2
         Left            =   405
         TabIndex        =   12
         Top             =   945
         Width           =   1095
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Fecha captura:"
         Height          =   195
         Index           =   5
         Left            =   3225
         TabIndex        =   8
         Top             =   1260
         Width           =   1080
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Ruta:"
         Height          =   195
         Index           =   6
         Left            =   1110
         TabIndex        =   6
         Top             =   1845
         Width           =   390
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor:"
         Height          =   195
         Index           =   7
         Left            =   765
         TabIndex        =   4
         Top             =   1545
         Width           =   735
      End
   End
   Begin VB.TextBox txt_buscar 
      Height          =   285
      Left            =   3135
      TabIndex        =   1
      Top             =   4245
      Width           =   1350
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2775
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   90
      Width           =   255
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   600
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
            Picture         =   "frmclientesproveedores.frx":0102
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclientesproveedores.frx":09DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclientesproveedores.frx":12B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclientesproveedores.frx":1852
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclientesproveedores.frx":212C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclientesproveedores.frx":2A06
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclientesproveedores.frx":32E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclientesproveedores.frx":35FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclientesproveedores.frx":3914
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclientesproveedores.frx":3EB0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   0
      Top             =   15
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
            Picture         =   "frmclientesproveedores.frx":41CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclientesproveedores.frx":4AA4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin vbskfree.Skinner Skinner1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1270
      _ExtentY        =   1270
      MaxButton       =   0
      MinButton       =   0
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   120
      TabIndex        =   26
      Top             =   4095
      Width           =   9525
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   6480
         TabIndex        =   27
         Top             =   165
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Un Registro Atras"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Un Registro Adelante"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Buscar"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Busqueda de cliente y/o proveedor:"
         Height          =   195
         Left            =   195
         TabIndex        =   28
         Top             =   195
         Width           =   2550
      End
   End
End
Attribute VB_Name = "frmclientesproveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean



Private Sub cmd_buscar_Click()
    Call pro_busca_registro(lv_clientesproveedores, txt_buscar, False)
    txt_buscar = ""
    pro_textos

End Sub

Private Sub Combo1_Click()
   txt_clientesproveedores(0) = Obtener_llave(cnn, rs, "TB_EMPRESAS", "VCHA_EMP_NOMBRE", Combo1, 0, "T")
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_clientesproveedores(1).SetFocus
   Else
      KeyAscii = 0
   End If
   
End Sub

Private Sub Combo2_Click()
   txt_clientesproveedores(4) = Obtener_llave(cnn, rs, "TB_TIPOclientesproveedores", "Vcha_tag_descripcion", Combo2, 0, "T")
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Combo3.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub Combo3_Click()
   txt_clientesproveedores(5) = Obtener_llave(cnn, rs, "TB_ZONAS", "Vcha_zon_descripcion", Combo3, 0, "T")
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_clientesproveedores(6).SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub cmb_clientesproveedores_Click(index As Integer)
   
   If index = 5 Then
      txt_clientesproveedores(5) = Obtener_llave(cnn, rs, "TB_VENDEDORES", "VCHA_VEN_NOMBRE", cmb_clientesproveedores(5), 0, "T")
   End If
   If index = 6 Then
      txt_clientesproveedores(6) = Obtener_llave(cnn, rs, "TB_RUTAS", "VCHA_RUT_NOMBRE", cmb_clientesproveedores(6), 0, "T")
   End If
   If index = 9 Then
      txt_clientesproveedores(9) = Obtener_llave(cnn, rs, "TB_MONEDAS", "VCHA_MON_DESCRIPCION", cmb_clientesproveedores(9), 0, "T")
   End If
   If index = 10 Then
      txt_clientesproveedores(10) = Obtener_llave(cnn, rs, "TB_PLAZOS", "VCHA_PLA_NOMBRE", cmb_clientesproveedores(10), 0, "T")
   End If
   If index = 11 Then
      txt_clientesproveedores(11) = Obtener_llave(cnn, rs, "TB_TIPOSCLIENTES", "VCHA_TCL_NOMBRE", cmb_clientesproveedores(11), 0, "T")
   End If
   If index = 12 Then
      txt_clientesproveedores(12) = Obtener_llave(cnn, rs, "TB_LISTADEPRECIOS", "VCHA_LIS_NOMBRE", cmb_clientesproveedores(12), 0, "T")
   End If
   If index = 13 Then
      txt_clientesproveedores(13) = Obtener_llave(cnn, rs, "TB_CANALESVENTAS", "VCHA_CAN_NOMBRE", cmb_clientesproveedores(13), 0, "T")
   End If
   If index = 14 Then
      txt_clientesproveedores(14) = Obtener_llave(cnn, rs, "TB_FAMILIA_AGRUPADORES", "VCHA_FAG_NOMBRE", cmb_clientesproveedores(14), 0, "T")
   End If
   If index = 15 Then
      txt_clientesproveedores(15) = Obtener_llave(cnn, rs, "TB_TRANSPORTES", "VCHA_TRN_NOMBRE", cmb_clientesproveedores(15), 0, "T")
   End If
   If index = 17 Then
      txt_clientesproveedores(17) = Obtener_llave(cnn, rs, "TB_TITULARES", "VCHA_TIT_NOMBRE", cmb_clientesproveedores(17), 1, "T")
   End If
   var_hubo_cambios = True
End Sub

Private Sub cmb_clientesproveedores_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
   If index = 5 And KeyCode = 116 Then
      frmvendedores.Show 1
      rs.Open "select * from tb_vendedores order by vcha_ven_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
      Call RecsetToCombo(cmb_clientesproveedores(5).hwnd, rs, 1)
      rs.Close
   End If
   If index = 6 And KeyCode = 116 Then
      frmrutas.Show 1
      rs.Open "select * from tb_rutas", cnn, adOpenDynamic, adLockBatchOptimistic
      Call RecsetToCombo(cmb_clientesproveedores(6).hwnd, rs, 1)
      rs.Close
   End If
   Call pro_textos
   var_hubo_cambios = True

End Sub

Private Sub cmdfecha_Click(index As Integer)
   mes(0).Visible = True
End Sub

Private Sub Form_Activate()
Dim var_resultado As Variant
Dim mientras As Integer
mientras = 0
If mientras = 0 Then
    If sw_primera_validacion = False Then
    
        If var_swpassword = False Then
        Call menuvisible(Frmmenu2, False)
            var_resultado = InStr(1, var_menus, Me.caption & "*1")
            If var_resultado <> 0 Then
                Set var_forma = frmclientesproveedores
                var_swpassword = True
                sw_primera_validacion = True
                frmclientesproveedores.Hide
                frmpasswords.Show 1
            End If
        End If
        If var_swpassword = False Then
            var_resultado = InStr(1, var_menus, Me.caption & "*01")
            If var_resultado <> 0 Then
                Set var_forma = frmclientesproveedores
                var_swpassword = True
                sw_primera_validacion = True
                frmclientesproveedores.Hide
                frmpasswords2.txt_supervisor = var_supervisor
                frmpasswords2.Show
            End If
        End If
    End If
End If
mes(0).Visible = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  Call pro_enfoque(KeyAscii)
End Sub

Private Sub Form_Load()
    var_modifica_registro = True
    lv_clientesproveedores.SmallIcons = ImageList1
    Call pro_encabezadosView(Me, lv_clientesproveedores, False)
    Call pro_llena_listview1
    rs.Open "select * from tb_vendedores order by vcha_ven_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
       Call RecsetToCombo(cmb_clientesproveedores(5).hwnd, rs, 1)
    rs.Close
    rs.Open "select * from tb_rutas", cnn, adOpenDynamic, adLockBatchOptimistic
       Call RecsetToCombo(cmb_clientesproveedores(6).hwnd, rs, 1)
    rs.Close
    rs.Open "select * from tb_monedas", cnn, adOpenDynamic, adLockBatchOptimistic
       Call RecsetToCombo(cmb_clientesproveedores(9).hwnd, rs, 1)
    rs.Close
    rs.Open "select * from tb_plazos", cnn, adOpenDynamic, adLockBatchOptimistic
       Call RecsetToCombo(cmb_clientesproveedores(10).hwnd, rs, 1)
    rs.Close
    rs.Open "select * from tb_tiposclientes", cnn, adOpenDynamic, adLockBatchOptimistic
       Call RecsetToCombo(cmb_clientesproveedores(11).hwnd, rs, 1)
    rs.Close
    rs.Open "select * from tb_listadeprecios", cnn, adOpenDynamic, adLockBatchOptimistic
       Call RecsetToCombo(cmb_clientesproveedores(12).hwnd, rs, 1)
    rs.Close
    rs.Open "select * from tb_canalesventas", cnn, adOpenDynamic, adLockBatchOptimistic
       Call RecsetToCombo(cmb_clientesproveedores(13).hwnd, rs, 1)
    rs.Close
    rs.Open "select * from tb_familia_agrupadores", cnn, adOpenDynamic, adLockBatchOptimistic
       Call RecsetToCombo(cmb_clientesproveedores(14).hwnd, rs, 1)
    rs.Close
    rs.Open "select * from tb_transportes", cnn, adOpenDynamic, adLockBatchOptimistic
       Call RecsetToCombo(cmb_clientesproveedores(15).hwnd, rs, 1)
    rs.Close
    rs.Open "select * from tb_TITULARES", cnn, adOpenDynamic, adLockBatchOptimistic
       Call RecsetToCombo(cmb_clientesproveedores(17).hwnd, rs, 2)
    rs.Close

    Call pro_AsignarAViewColor(lv_clientesproveedores, Picture1, vbWhite, vbGray)
    rs.Open "select * from tb_clientesproveedores", cnn, adOpenDynamic, adLockOptimistic
    If rs.BOF Then
       Toolbar1.Buttons.Item(2).Enabled = False
       Toolbar1.Buttons.Item(3).Enabled = False
       Toolbar1.Buttons.Item(4).Enabled = False
    Else
       Toolbar1.Buttons.Item(2).Enabled = True
       Toolbar1.Buttons.Item(3).Enabled = True
       Toolbar1.Buttons.Item(4).Enabled = True
    End If
    rs.Close
    Call pro_textos
End Sub

Private Sub Form_Unload(Cancel As Integer)
    var_swpassword = False
    var_modifica_registro = False
    Call menuvisible(Frmmenu2, True)
End Sub

Private Sub lv_clientesproveedores_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_clientesproveedores.SelectedItem = Item
        Call pro_textos
        var_modifica_registro = True
        txt_clientesproveedores(0).Enabled = True
End Sub

Private Sub mes_DblClick(index As Integer)
   txt_clientesproveedores(4) = mes(0).Value
   mes(0).Visible = False
End Sub

Private Sub mes_KeyPress(index As Integer, KeyAscii As Integer)
   If KeyAscii = 27 Then
      mes(0).Visible = False
   End If
End Sub

Private Sub mes_LostFocus(index As Integer)
   mes(0).Visible = False
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
   If Button.index < 3 Then
      lv_clientesproveedores.SetFocus
      Call pro_avanzar(Me, lv_clientesproveedores, Button)
      pro_textos
   Else
      Call pro_busca_registro(lv_clientesproveedores, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.index
    Case 1
        Call pro_limpiatextos(Me)
        txt_clientesproveedores(0).Enabled = True
        txt_clientesproveedores(0).SetFocus: var_modifica_registro = False
        Toolbar1.Buttons.Item(2).Enabled = True
        Toolbar1.Buttons.Item(3).Enabled = True
    Case 2
        var_resultado = InStr(1, var_menus, Me.caption)
        var_inicio = var_resultado + Len(Me.caption) + 3
        If Mid(var_menus, var_inicio, 1) = "1" Then
            Set var_forma = frmclientesproveedores
            var_swpassword = True
            sw_primera_validacion = False
            frmpasswords.Show 1
        Else
            If Mid(var_menus, var_inicio, 2) = "01" Then
                Set var_forma = frmclientesproveedores
                var_swpassword = True
                sw_primera_validacion = False
                frmpasswords2.txt_supervisor = var_supervisor
                frmpasswords2.Show 1
            Else
            Call pro_guardar_clientesproveedores
            rs.Open "select * from tb_clientesproveedores", cnn, adOpenDynamic, adLockOptimistic
            If rs.BOF Then
               Toolbar1.Buttons.Item(2).Enabled = False
               Toolbar1.Buttons.Item(3).Enabled = False
               Toolbar1.Buttons.Item(4).Enabled = False
            Else
               Toolbar1.Buttons.Item(2).Enabled = True
               Toolbar1.Buttons.Item(3).Enabled = True
               Toolbar1.Buttons.Item(4).Enabled = True
            End If
            rs.Close
            
            End If
        End If
    Case 3
        Call pro_textos
    Case 4
        var_resultado = InStr(1, var_menus, Me.caption)
        var_inicio = var_resultado + Len(Me.caption) + 3
        If Mid(var_menus, var_inicio, 1) = "1" Then
            Set var_forma = frmclientesproveedores
            var_swpassword = True
            sw_primera_validacion = False
            frmpasswords.Show 1
        Else
            If Mid(var_menus, var_inicio, 2) = "01" Then
                Set var_forma = frmclientesproveedores
                var_swpassword = True
                sw_primera_validacion = False
                frmpasswords2.txt_supervisor = var_supervisor
                frmpasswords2.Show 1
            Else
               Call pro_elimina_clientesproveedores
               rs.Open "select * from tb_clientesproveedores", cnn, adOpenDynamic, adLockOptimistic
               If rs.BOF Then
                  Toolbar1.Buttons.Item(2).Enabled = False
                  Toolbar1.Buttons.Item(3).Enabled = False
                  Toolbar1.Buttons.Item(4).Enabled = False
               Else
                  Toolbar1.Buttons.Item(2).Enabled = True
                  Toolbar1.Buttons.Item(3).Enabled = True
                  Toolbar1.Buttons.Item(4).Enabled = True
               End If
               rs.Close
            End If
        End If
    Case 6
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call gPrintListView(lv_clientesproveedores, "LISTADO DE clientesproveedores")
        End If
    Case 8
        Unload Me
    End Select

End Sub

Sub pro_guardar_clientesproveedores()

Dim ok As Boolean

Set TB_CLIENTESPROVEEDORES = New TB_CLIENTESPROVEEDORES
    
    
    If txt_clientesproveedores(0) <> "" And txt_clientesproveedores(1) <> "" Then
        If var_hubo_cambios Then
            ok = TB_CLIENTESPROVEEDORES.Anadir(txt_clientesproveedores(0), txt_clientesproveedores(1), txt_clientesproveedores(2), chk_clientesproveedores(3), txt_clientesproveedores(4), txt_clientesproveedores(5), txt_clientesproveedores(6), txt_clientesproveedores(7), txt_clientesproveedores(8), txt_clientesproveedores(9), txt_clientesproveedores(10), txt_clientesproveedores(11), txt_clientesproveedores(12), txt_clientesproveedores(13), txt_clientesproveedores(14), txt_clientesproveedores(15), chk_clientesproveedores(16), txt_clientesproveedores(17))
            If ok Then
                pro_actualiza_ListView
                txt_clientesproveedores(0).Enabled = False
                MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
                txt_registros = lv_clientesproveedores.ListItems.Count
                var_modifica_registro = True
            Else
                MsgBox "No se puede grabar registro: " + TB_CLIENTESPROVEEDORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
        End If
    End If
    
Set TB_CLIENTESPROVEEDORES = Nothing: var_hubo_cambios = False

End Sub

Sub pro_elimina_clientesproveedores()
Dim var_llave_usuarios As String

Set TB_CLIENTESPROVEEDORES = New TB_CLIENTESPROVEEDORES
On Error GoTo SALIR
  
    ok = True
        If txt_clientesproveedores(0) <> "" And txt_clientesproveedores(1) <> "" Then
            If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
                ok = TB_CLIENTESPROVEEDORES.Eliminar(txt_clientesproveedores(0))
            Else
                GoTo SALIR:
            End If
            If ok Then
                MsgBox "Se Elimino Correctamente el Registro", vbInformation
                lv_clientesproveedores.ListItems.Remove (lv_clientesproveedores.SelectedItem.index)
                Call pro_limpiatextos(Me)
                txt_registros = lv_clientesproveedores.ListItems.Count
                lv_clientesproveedores.SelectedItem.Selected = True
                pro_textos
           Else
                MsgBox "No se puede grabar registro: " + TB_CLIENTESPROVEEDORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
        End If
SALIR:
Set TB_CLIENTESPROVEEDORES = Nothing

End Sub


Sub pro_llena_listview1()

Dim list_item As ListItem

    rs.Open "select * from TB_clientesproveedores", cnn, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
        Set list_item = lv_clientesproveedores.ListItems.Add(, , rs(0).Value): list_item.SmallIcon = 9
        list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
        list_item.SubItems(2) = IIf(IsNull(rs(2).Value), "", rs(2).Value)
        list_item.SubItems(3) = IIf(IsNull(rs(3).Value), "", rs(3).Value)
        list_item.SubItems(4) = IIf(IsNull(rs(4).Value), "", rs(4).Value)
        list_item.SubItems(5) = IIf(IsNull(rs(5).Value), "", rs(5).Value)
        list_item.SubItems(6) = IIf(IsNull(rs(6).Value), "", rs(6).Value)
        list_item.SubItems(7) = IIf(IsNull(rs(1).Value), "", rs(7).Value)
        list_item.SubItems(8) = IIf(IsNull(rs(2).Value), "", rs(8).Value)
        list_item.SubItems(9) = IIf(IsNull(rs(3).Value), "", rs(9).Value)
        list_item.SubItems(10) = IIf(IsNull(rs(10).Value), "", rs(10).Value)
        list_item.SubItems(11) = IIf(IsNull(rs(11).Value), "", rs(11).Value)
        list_item.SubItems(12) = IIf(IsNull(rs(12).Value), "", rs(12).Value)
        list_item.SubItems(13) = IIf(IsNull(rs(13).Value), "", rs(13).Value)
        list_item.SubItems(14) = IIf(IsNull(rs(14).Value), "", rs(14).Value)
        list_item.SubItems(15) = IIf(IsNull(rs(15).Value), "", rs(15).Value)
        list_item.SubItems(16) = IIf(IsNull(rs(16).Value), "", rs(16).Value)
        list_item.SubItems(17) = IIf(IsNull(rs(17).Value), "", rs(17).Value)
    rs.MoveNext:
    Wend
    rs.Close

End Sub


Sub pro_textos()
'On Error GoTo err0:
        txt_clientesproveedores(0) = lv_clientesproveedores.SelectedItem
        txt_clientesproveedores(1) = lv_clientesproveedores.SelectedItem.SubItems(1)
        txt_clientesproveedores(2) = lv_clientesproveedores.SelectedItem.SubItems(2)
        chk_clientesproveedores(3) = lv_clientesproveedores.SelectedItem.SubItems(3)
        txt_clientesproveedores(4) = lv_clientesproveedores.SelectedItem.SubItems(4)
        txt_clientesproveedores(5) = lv_clientesproveedores.SelectedItem.SubItems(5)
        txt_clientesproveedores(6) = lv_clientesproveedores.SelectedItem.SubItems(6)
        txt_clientesproveedores(7) = lv_clientesproveedores.SelectedItem.SubItems(7)
        txt_clientesproveedores(8) = lv_clientesproveedores.SelectedItem.SubItems(8)
        txt_clientesproveedores(9) = lv_clientesproveedores.SelectedItem.SubItems(9)
        txt_clientesproveedores(10) = lv_clientesproveedores.SelectedItem.SubItems(10)
        txt_clientesproveedores(11) = lv_clientesproveedores.SelectedItem.SubItems(11)
        txt_clientesproveedores(12) = lv_clientesproveedores.SelectedItem.SubItems(12)
        txt_clientesproveedores(13) = lv_clientesproveedores.SelectedItem.SubItems(13)
        txt_clientesproveedores(14) = lv_clientesproveedores.SelectedItem.SubItems(14)
        txt_clientesproveedores(15) = lv_clientesproveedores.SelectedItem.SubItems(15)
        chk_clientesproveedores(16) = lv_clientesproveedores.SelectedItem.SubItems(16)
        txt_clientesproveedores(17) = lv_clientesproveedores.SelectedItem.SubItems(17)
        cmb_clientesproveedores(5) = Obtener_llave(cnn, rs, "TB_VENDEDORES", "VCHA_VEN_VENDEDOR_ID", txt_clientesproveedores(5), 1, "T")
        cmb_clientesproveedores(6) = Obtener_llave(cnn, rs, "TB_RUTAS", "VCHA_RUT_RUTA_ID", txt_clientesproveedores(6), 1, "T")
        cmb_clientesproveedores(9) = Obtener_llave(cnn, rs, "TB_MONEDAS", "VCHA_MON_MONEDA_ID", txt_clientesproveedores(9), 1, "T")
        cmb_clientesproveedores(10) = Obtener_llave(cnn, rs, "TB_PLAZOS", "VCHA_PLA_PLAZO_ID", txt_clientesproveedores(10), 1, "T")
        cmb_clientesproveedores(11) = Obtener_llave(cnn, rs, "TB_TIPOSCLIENTES", "VCHA_TCL_TIPO_CLIENTE_ID", txt_clientesproveedores(11), 1, "T")
        cmb_clientesproveedores(12) = Obtener_llave(cnn, rs, "TB_LISTADEPRECIOS", "VCHA_LIS_LISTA_ID ", txt_clientesproveedores(12), 1, "T")
        cmb_clientesproveedores(13) = Obtener_llave(cnn, rs, "TB_CANALESVENTAS", "VCHA_CAN_CANAL_VENTA_ID", txt_clientesproveedores(13), 1, "T")
        cmb_clientesproveedores(14) = Obtener_llave(cnn, rs, "TB_FAMILIA_AGRUPADORES", "VCHA_FAG_FAMILIA_AGRUPADOR_ID", txt_clientesproveedores(14), 1, "T")
        cmb_clientesproveedores(15) = Obtener_llave(cnn, rs, "TB_TRANSPORTES", "VCHA_TRN_TRANSPORTE_ID", txt_clientesproveedores(15), 1, "T")
        cmb_clientesproveedores(17) = Obtener_llave(cnn, rs, "TB_TITULARES", "VCHA_TIT_TITULAR_ID", txt_clientesproveedores(17), 2, "T")
        
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro = False Then
        Set list_item = lv_clientesproveedores.ListItems.Add(, , txt_clientesproveedores(0)): list_item.SmallIcon = 9
        list_item.SubItems(1) = txt_clientesproveedores(1)
        list_item.SubItems(2) = txt_clientesproveedores(2)
        list_item.SubItems(3) = chk_clientesproveedores(3)
        list_item.SubItems(4) = txt_clientesproveedores(4)
        list_item.SubItems(5) = txt_clientesproveedores(5)
        list_item.SubItems(6) = txt_clientesproveedores(6)
        list_item.SubItems(7) = txt_clientesproveedores(7)
        list_item.SubItems(8) = txt_clientesproveedores(8)
        list_item.SubItems(9) = txt_clientesproveedores(9)
        list_item.SubItems(10) = txt_clientesproveedores(10)
        list_item.SubItems(11) = txt_clientesproveedores(11)
        list_item.SubItems(12) = txt_clientesproveedores(12)
        list_item.SubItems(13) = txt_clientesproveedores(13)
        list_item.SubItems(14) = txt_clientesproveedores(14)
        list_item.SubItems(14) = txt_clientesproveedores(15)
        list_item.SubItems(16) = chk_clientesproveedores(16)
        list_item.SubItems(17) = txt_clientesproveedores(17)
        list_item.EnsureVisible
        list_item.Selected = True
    Else
        lv_clientesproveedores.ListItems.Item(lv_clientesproveedores.SelectedItem.index).Checked = False
        lv_clientesproveedores.ListItems.Item(lv_clientesproveedores.SelectedItem.index) = txt_clientesproveedores(0)
        lv_clientesproveedores.ListItems.Item(lv_clientesproveedores.SelectedItem.index).ListSubItems(1) = txt_clientesproveedores(1)
        lv_clientesproveedores.ListItems.Item(lv_clientesproveedores.SelectedItem.index).ListSubItems(2) = txt_clientesproveedores(2)
        lv_clientesproveedores.ListItems.Item(lv_clientesproveedores.SelectedItem.index).ListSubItems(3) = chk_clientesproveedores(3)
        lv_clientesproveedores.ListItems.Item(lv_clientesproveedores.SelectedItem.index).ListSubItems(4) = txt_clientesproveedores(4)
        lv_clientesproveedores.ListItems.Item(lv_clientesproveedores.SelectedItem.index).ListSubItems(5) = txt_clientesproveedores(5)
        lv_clientesproveedores.ListItems.Item(lv_clientesproveedores.SelectedItem.index).ListSubItems(6) = txt_clientesproveedores(6)
        lv_clientesproveedores.ListItems.Item(lv_clientesproveedores.SelectedItem.index).ListSubItems(7) = txt_clientesproveedores(7)
        lv_clientesproveedores.ListItems.Item(lv_clientesproveedores.SelectedItem.index).ListSubItems(8) = txt_clientesproveedores(8)
        lv_clientesproveedores.ListItems.Item(lv_clientesproveedores.SelectedItem.index).ListSubItems(9) = txt_clientesproveedores(9)
        lv_clientesproveedores.ListItems.Item(lv_clientesproveedores.SelectedItem.index).ListSubItems(10) = txt_clientesproveedores(10)
        lv_clientesproveedores.ListItems.Item(lv_clientesproveedores.SelectedItem.index).ListSubItems(11) = txt_clientesproveedores(11)
        lv_clientesproveedores.ListItems.Item(lv_clientesproveedores.SelectedItem.index).ListSubItems(12) = txt_clientesproveedores(12)
        lv_clientesproveedores.ListItems.Item(lv_clientesproveedores.SelectedItem.index).ListSubItems(13) = txt_clientesproveedores(13)
        lv_clientesproveedores.ListItems.Item(lv_clientesproveedores.SelectedItem.index).ListSubItems(14) = txt_clientesproveedores(14)
        lv_clientesproveedores.ListItems.Item(lv_clientesproveedores.SelectedItem.index).ListSubItems(15) = txt_clientesproveedores(15)
        lv_clientesproveedores.ListItems.Item(lv_clientesproveedores.SelectedItem.index).ListSubItems(16) = chk_clientesproveedores(16)
        lv_clientesproveedores.ListItems.Item(lv_clientesproveedores.SelectedItem.index).ListSubItems(17) = txt_clientesproveedores(17)
    End If
    lv_clientesproveedores.SetFocus
End Sub

Private Sub txt_clientesproveedores_Change(index As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   var_modifica_registro = True
End Sub

Private Sub txt_clientesproveedores_KeyPress(index As Integer, KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub

