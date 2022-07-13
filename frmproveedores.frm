VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmproveedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proveedores"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   11670
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   1590
      TabIndex        =   53
      Top             =   2145
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   30
         TabIndex        =   54
         Top             =   495
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
         TabIndex        =   55
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame frm_colonias 
      Height          =   2400
      Left            =   1590
      TabIndex        =   50
      Top             =   2160
      Width           =   5685
      Begin MSComctlLib.ListView lv_colonias 
         Height          =   1830
         Left            =   30
         TabIndex        =   51
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
         TabIndex        =   52
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_importar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1695
      Picture         =   "frmproveedores.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Importar Información  Alt + M"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      Picture         =   "frmproveedores.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   375
      Picture         =   "frmproveedores.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Guardar Alt + G"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   705
      Picture         =   "frmproveedores.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1035
      Picture         =   "frmproveedores.frx":03D8
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1365
      Picture         =   "frmproveedores.frx":04DA
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11265
      Picture         =   "frmproveedores.frx":05DC
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame frm_ruta 
      Height          =   3615
      Left            =   6690
      TabIndex        =   43
      Top             =   -90
      Width           =   3330
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   75
         TabIndex        =   48
         Top             =   2835
         Width           =   3195
      End
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   75
         TabIndex        =   47
         Top             =   870
         Width           =   3180
      End
      Begin VB.TextBox txt_path 
         Height          =   330
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   450
         Width           =   3180
      End
      Begin VB.CommandButton cmd_aceptar 
         Caption         =   "&Aceptar"
         Height          =   330
         Left            =   90
         TabIndex        =   45
         Top             =   3210
         Width           =   1605
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Cancelar"
         Height          =   330
         Left            =   1680
         TabIndex        =   44
         Top             =   3210
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Seleccione la Ruta"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   5
         Left            =   30
         TabIndex        =   49
         Top             =   120
         Width           =   3255
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   5460
      TabIndex        =   39
      Top             =   450
      Width           =   6090
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   2100
         TabIndex        =   40
         Top             =   150
         Width           =   1815
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   4230
         TabIndex        =   41
         Top             =   150
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
         Caption         =   "Busqueda de Proveedor:"
         Height          =   195
         Left            =   195
         TabIndex        =   42
         Top             =   210
         Width           =   1770
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7140
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   38
      Top             =   60
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   "  Proveedor "
      Height          =   4710
      Left            =   45
      TabIndex        =   24
      Top             =   450
      Width           =   5370
      Begin VB.TextBox txt_nombre_municipio 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   18
         Top             =   3180
         Width           =   4020
      End
      Begin MSComCtl2.MonthView mes 
         Height          =   2370
         Left            =   1530
         TabIndex        =   25
         Top             =   255
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   49741825
         CurrentDate     =   37581
      End
      Begin VB.TextBox txt_fecha_alta 
         Height          =   315
         Left            =   1290
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1200
         Width           =   1230
      End
      Begin VB.TextBox txt_rfc 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   13
         Top             =   1530
         Width           =   1980
      End
      Begin VB.TextBox txt_representante 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   10
         Top             =   870
         Width           =   4020
      End
      Begin VB.TextBox txt_nombre 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   9
         Top             =   540
         Width           =   4020
      End
      Begin VB.TextBox txt_clave 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   8
         Top             =   210
         Width           =   1815
      End
      Begin VB.CommandButton cmdfecha 
         Height          =   285
         Index           =   0
         Left            =   2550
         Picture         =   "frmproveedores.frx":0C16
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Seleccione la fecha"
         Top             =   1215
         Width           =   315
      End
      Begin VB.TextBox txt_telefono 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   21
         Top             =   4170
         Width           =   2055
      End
      Begin VB.TextBox txt_nombre_pais 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   20
         Top             =   3840
         Width           =   4035
      End
      Begin VB.TextBox txt_nombre_estado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   19
         Top             =   3510
         Width           =   4020
      End
      Begin VB.TextBox txt_nombre_ciudad 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   17
         Top             =   2850
         Width           =   4020
      End
      Begin VB.TextBox txt_codigo_postal 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   15
         Top             =   2190
         Width           =   1095
      End
      Begin VB.TextBox txt_domicilio 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   14
         Top             =   1860
         Width           =   4020
      End
      Begin VB.TextBox txt_nombre_colonia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   16
         Top             =   2520
         Width           =   4020
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Municipio:"
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   56
         Top             =   3240
         Width           =   720
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Fecha captura:"
         Height          =   195
         Index           =   5
         Left            =   150
         TabIndex        =   37
         Top             =   1215
         Width           =   1080
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Representante:"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   36
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   35
         Top             =   585
         Width           =   600
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   34
         Top             =   240
         Width           =   450
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "RFC"
         Height          =   195
         Index           =   14
         Left            =   150
         TabIndex        =   33
         Top             =   1575
         Width           =   315
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono:"
         Height          =   195
         Index           =   16
         Left            =   165
         TabIndex        =   32
         Top             =   4230
         Width           =   675
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Pais:"
         Height          =   195
         Index           =   17
         Left            =   165
         TabIndex        =   31
         Top             =   3900
         Width           =   345
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Index           =   18
         Left            =   165
         TabIndex        =   30
         Top             =   3570
         Width           =   540
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad:"
         Height          =   195
         Index           =   19
         Left            =   150
         TabIndex        =   29
         Top             =   2910
         Width           =   540
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Colonia:"
         Height          =   195
         Index           =   20
         Left            =   150
         TabIndex        =   28
         Top             =   2580
         Width           =   570
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Domicilio:"
         Height          =   195
         Index           =   21
         Left            =   150
         TabIndex        =   27
         Top             =   1905
         Width           =   675
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "CP:"
         Height          =   195
         Index           =   22
         Left            =   150
         TabIndex        =   26
         Top             =   2250
         Width           =   255
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4170
      Left            =   5460
      TabIndex        =   22
      Top             =   990
      Width           =   6105
      Begin MSComctlLib.ListView lv_proveedores 
         Height          =   3915
         Left            =   45
         TabIndex        =   23
         Top             =   165
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   6906
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
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7691
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "representante"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "fecha"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "rfc"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "pais"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "estado"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "municipio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "ciudad"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "colonia"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "domicilio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "cp"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "telefono"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   7620
      Top             =   -30
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
            Picture         =   "frmproveedores.frx":0D18
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmproveedores.frx":15F2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3960
      Top             =   -135
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
            Picture         =   "frmproveedores.frx":1ECC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmproveedores.frx":27A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmproveedores.frx":3080
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmproveedores.frx":361C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmproveedores.frx":3EF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmproveedores.frx":47D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmproveedores.frx":50AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmproveedores.frx":51BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmproveedores.frx":52D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmproveedores.frx":53E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmproveedores.frx":54F4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   60
      Left            =   45
      TabIndex        =   0
      Top             =   330
      Width           =   11580
   End
End
Attribute VB_Name = "frmproveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim var_guardar_cambios As Boolean
Dim varpais As String
Dim varestado As String
Dim varmunicipio As String
Dim varciudad As String
Dim varcolonia As String
Dim bitacora As Boolean
Dim numero_items_proveedores As Integer
Dim var_tabla As ADODB.Connection

Private Sub cmd_aceptar_Click()
   On Error GoTo salir:
   Dim var_Archivo As String
   Dim var_fecha As String
   Dim var_nombre As String
   Dim var_representante As String
   Dim var_colonia As String
   Dim var_direccion As String
   Dim var_ciudad As String
   Dim i, l As Integer
   Dim var_cadena As String
   var_Archivo = txt_path + "\MGP10013.DBF"
   If var_tabla.State = 1 Then
      var_tabla.Close
   End If
   var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + txt_path + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
   rs.Open "select * from MGP10013.DBF", var_tabla, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
         l = Len(Trim(IIf(IsNull(rs!nombprovee), "", rs!nombprovee)))
         var_nombre = Trim(IIf(IsNull(rs!nombprovee), "", rs!nombprovee))
         Cadena = ""
         For i = 1 To l
             If Mid(var_nombre, i, 1) = "'" Then
                Cadena = Cadena + "'" + Mid(var_nombre, i, 1)
             Else
                Cadena = Cadena + Mid(var_nombre, i, 1)
             End If
         Next i
         var_nombre = Cadena
         
         l = Len(Trim(IIf(IsNull(rs!represprov), "", rs!represprov)))
         var_representante = Trim(IIf(IsNull(rs!represprov), "", rs!represprov))
         Cadena = ""
         For i = 1 To l
             If Mid(var_representante, i, 1) = "'" Then
                Cadena = Cadena + "'" + Mid(var_representante, i, 1)
             Else
                Cadena = Cadena + Mid(var_representante, i, 1)
             End If
         Next i
         var_representante = Cadena
         
         l = Len(Trim(IIf(IsNull(rs!direcprove), "", rs!direcprove)))
         var_direccion = Trim(IIf(IsNull(rs!direcprove), "", rs!direcprove))
         Cadena = ""
         For i = 1 To l
             If Mid(var_direccion, i, 1) = "'" Then
                Cadena = Cadena + "'" + Mid(var_direccion, i, 1)
             Else
                Cadena = Cadena + Mid(var_direccion, i, 1)
             End If
         Next i
         var_direccion = Cadena
         
         l = Len(Trim(IIf(IsNull(rs!coloniapro), "", rs!coloniapro)))
         var_colonia = Trim(IIf(IsNull(rs!coloniapro), "", rs!coloniapro))
         Cadena = ""
         For i = 1 To l
             If Mid(var_colonia, i, 1) = "'" Then
                Cadena = Cadena + "'" + Mid(var_colonia, i, 1)
             Else
                Cadena = Cadena + Mid(var_colonia, i, 1)
             End If
         Next i
         var_colonia = Cadena
         
         var_fecha = CStr(fecaltprov)
         If Trim(var_fecha) = "" Then
            var_fecha = Date
         End If
         rsaux2.Open "exec importar_proveedores '" + IIf(IsNull(rs!codcteprov), "", Trim(rs!codcteprov)) + "', '" + var_nombre + "', '" + var_representante + "', '" + var_fecha + "', '" + IIf(IsNull(rs!rfcproved), "", rs!rfcproved) + "', '', '', '', '" + var_colonia + "', '" + var_direccion + "', '" + IIf(IsNull(rs!codpostpro), "", Trim(rs!codpostpro)) + "', '" + IIf(IsNull(rs!telprov1), "", Trim(rs!telprov1)) + "'", cnn, adOpenDynamic, adLockOptimistic
         rs.MoveNext
      Wend
   End If
   rs.Close
   lv_proveedores.ListItems.Clear
   Call pro_llena_listview1
   Call pro_textos
   frm_ruta.Visible = False
   Exit Sub
salir:
   If var_tabla.State = 1 Then
      var_tabla.Close
   End If
   MsgBox "Error al leer el archivo, puede que no exista o este siendo utilizado por otro usuario", vbOKOnly, "ATENCIO"
   frm_ruta.Visible = False
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
      Call pro_elimina_proveedores
      rs.Open "select * from tb_proveedores", cnn, adOpenDynamic, adLockOptimistic
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
   Else
      MsgBox "Imposible realizar la acción solicitada", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_guardar_Click()
Dim var_posible As Boolean
   var_posible = True
   If var_modifica_registro_proveedor = False Then
      rs.Open "select * from tb_proveedores where vcha_pro_proveedor_id = '" + Me.txt_clave + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_posible = False
      End If
      rs.Close
   End If
   If var_posible = True Then
      If txt_clave = "" Or txt_nombre = "" Or txt_representante = "" Or txt_fecha_alta = "" Then
         MsgBox "Información incompleta", vbOKOnly, "ATENCION"
      Else
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
            Call pro_guardar_proveedores
            rs.Open "select * from tb_proveedores", cnn, adOpenDynamic, adLockOptimistic
            If rs.BOF Then
               var_guardar_cambios = False
               cmd_guardar.Enabled = False
               cmd_deshacer.Enabled = False
               cmd_eliminar.Enabled = False
            Else
               cmd_guardar.Enabled = True
               cmd_deshacer.Enabled = True
               cmd_eliminar.Enabled = True
               var_guardar_cambios = False
            End If
            rs.Close
         Else
            MsgBox "Imposible realizar la acción solicitada", vbOKOnly, "ATENCION"
         End If
      End If
   Else
      MsgBox "Clave de proveedor ya existe", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_importar_Click()
        txt_path = App.Path
        Dir1.Path = App.Path
        frm_ruta.Visible = True

End Sub

Private Sub cmd_imprimir_Click()
         If vector_valida_passwords(var_indice_menu) = "*" Then
            frmpasswords.Show
         Else
            Call gPrintListView(lv_proveedores, "LISTADO DE proveedores")
         End If

End Sub

Private Sub cmd_nuevo_Click()
         Call pro_limpiatextos(Me)
         txt_clave.Enabled = True
         txt_clave.SetFocus: var_modifica_registro_proveedor = False
         cmd_guardar.Enabled = True
         cmd_deshacer.Enabled = True
         txt_fecha_alta = Date
         var_guardar_cambios = True

End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_proveedor = False Then
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

Private Sub cmdfecha_Click(Index As Integer)
   If Trim(Me.txt_fecha_alta) <> "" Then
      If IsDate(Me.txt_fecha_alta) Then
         mes.Value = Me.txt_fecha_alta
      Else
         mes.Value = Date
      End If
   Else
      mes.Value = Date
   End If
   mes.Visible = True
   mes.SetFocus
End Sub

Private Sub Command1_Click()
   frm_ruta.Visible = False
End Sub

Private Sub Dir1_Change()
   txt_path = Dir1.Path
End Sub

Private Sub Dir1_Click()
   txt_path = Dir1.Path
End Sub

Private Sub Dir1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      frm_ruta.Visible = False
   Else
      txt_path = Dir1.Path
   End If
End Sub

Private Sub Drive1_Change()
   On Error GoTo salir:
   Dir1.Path = Drive1.Drive
   Exit Sub
salir:
   MsgBox "La unidad " + Drive1.Drive + " no esta disponible", vbOKOnly, "ATENCION"
   Drive1.Drive = "c:"
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
   frm_colonias.Visible = False
   var_cadena_seguridad = ""
   Top = 1200
   Left = 0
   frm_ruta.Visible = False
   mes.Visible = False
   var_modifica_registro_proveedor = True
   'lv_proveedores.SmallIcons = ImageList1
   Set var_tabla = CreateObject("ADODB.connection")
   rs.Open "select * from tb_proveedores", cnn, adOpenDynamic, adLockOptimistic
   frm_lista.Visible = False
   If rs.BOF Then
      rs.Close
      cmd_guardar.Enabled = False
      cmd_deshacer.Enabled = False
      cmd_eliminar.Enabled = False
   Else
      cmd_guardar.Enabled = True
      cmd_deshacer.Enabled = True
      cmd_eliminar.Enabled = True
      rs.Close
      Call pro_llena_listview1
      Call pro_textos
      'lv_proveedores.SetFocus
   End If
   pro_textos
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_proveedores)
End Sub

Private Sub lv_colonias_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_colonias, ColumnHeader)
End Sub

Private Sub lv_colonias_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_colonias.ListItems.Count > 0 Then
         varcolonia = lv_colonias.selectedItem
         txt_nombre_colonia = lv_colonias.selectedItem.SubItems(1)
         varpais = lv_colonias.selectedItem.SubItems(2)
         txt_nombre_pais = lv_colonias.selectedItem.SubItems(3)
         varestado = lv_colonias.selectedItem.SubItems(4)
         txt_nombre_estado = lv_colonias.selectedItem.SubItems(5)
         varmunicipio = lv_colonias.selectedItem.SubItems(6)
         txt_nombre_municipio = lv_colonias.selectedItem.SubItems(7)
         varciudad = lv_colonias.selectedItem.SubItems(8)
         txt_nombre_ciudad = lv_colonias.selectedItem.SubItems(9)
      Else
         varcolonia = ""
         txt_nombre_colonia = ""
         varpais = ""
         txt_nombre_pais = ""
         varpais = ""
         txt_nombre_estado = ""
         varmunicipio = ""
         txt_nombre_municipio = ""
         varciudad = ""
         txt_nombre_ciudad = ""
      End If
      frm_colonias.Visible = False
      Me.txt_telefono.SetFocus
   End If
   If KeyAscii = 27 Then
      frm_colonias.Visible = False
   End If
End Sub

Private Sub lv_colonias_LostFocus()
   frm_colonias.Visible = False
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         rs.Open "select * from vw_colonias where vcha_col_cp = '" + Me.txt_codigo_postal + "' AND VCHA_PAI_PAIS_ID = '" + lv_lista.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
         If rs.RecordCount = 1 Then
            txt_colonia = IIf(IsNull(rs!VCHA_COL_COLONIA_ID), "", rs!VCHA_COL_COLONIA_ID)
            txt_nombre_colonia = IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
            txt_pais = IIf(IsNull(rs!vcha_pai_pais_id), "", rs!vcha_pai_pais_id)
            txt_nombre_pais = IIf(IsNull(rs!VCHA_PAI_NOMBRE), "", rs!VCHA_PAI_NOMBRE)
            txt_nombre_estado = IIf(IsNull(rs!vcha_est_estado_id), "", rs!vcha_est_estado_id)
            txt_nombre_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
            txt_municipio = IIf(IsNull(rs!vcha_mun_municipio_id), "", rs!vcha_mun_municipio_id)
            txt_nombre_municipio = IIf(IsNull(rs!VCHA_MUN_NOMBRE), "", rs!VCHA_MUN_NOMBRE)
            txt_ciudad = IIf(IsNull(rs!vcha_ciu_ciudad_id), "", rs!vcha_ciu_ciudad_id)
            txt_nombre_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
            rs.Close
            txt_telefono.SetFocus
            frm_lista.Visible = False
         Else
            'rs.MoveFirst
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
      Else
         txt_colonia = ""
         txt_nombre_colonia = ""
         txt_pais = ""
         txt_nombre_pais = ""
         txt_nombre_estado = ""
         txt_nombre_estado = ""
         txt_municipio = ""
         txt_nombre_municipio = ""
         txt_ciudad = ""
         txt_nombre_ciudad = ""
      End If
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub lv_proveedores_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_proveedores, ColumnHeader)
End Sub

Private Sub lv_proveedores_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_proveedores.selectedItem = Item
        Call pro_textos
        var_modifica_registro_proveedor = True
End Sub

Private Sub mes_DblClick()
   txt_fecha_alta = mes.Value
   mes.Visible = False
   txt_fecha_alta.SetFocus
End Sub

Private Sub mes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      mes.Visible = False
      txt_fecha_alta.SetFocus
   End If
End Sub

Private Sub mes_LostFocus()
   mes.Visible = False
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_proveedores.SetFocus
      Call pro_avanzar(Me, lv_proveedores, Button)
      lv_proveedores.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_proveedores.ListItems(1).Selected = True
      lv_proveedores.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 4 Then
      numero_items_proveedores = lv_proveedores.ListItems.Count
      lv_proveedores.ListItems(numero_items_proveedores).Selected = True
      lv_proveedores.selectedItem.EnsureVisible
      pro_textos
   End If
err0:
End Sub


Sub pro_guardar_proveedores()
   Dim ok As Boolean
   Set TB_PROVEEDORES = New TB_PROVEEDORES
   If var_hubo_cambios Then
      ok = TB_PROVEEDORES.Anadir(txt_clave, txt_nombre, txt_representante, txt_fecha_alta, txt_rfc, varpais, varestado, varmunicipio, varciudad, varcolonia, txt_domicilio, txt_codigo_postal, txt_telefono)
      If ok Then
         pro_actualiza_ListView
         txt_clave.Enabled = False
         MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
         var_modifica_registro_proveedor = True
      Else
         MsgBox "No se puede grabar registro: " + TB_PROVEEDORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
   Set TB_PROVEEDORES = Nothing: var_hubo_cambios = False
End Sub

Sub pro_elimina_proveedores()
Dim var_llave_usuarios As String

Set TB_PROVEEDORES = New TB_PROVEEDORES
On Error GoTo salir
  
    ok = True
        If txt_clave <> "" Then
            If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
                ok = TB_PROVEEDORES.Eliminar(txt_clave)
            Else
                GoTo salir:
            End If
            If ok Then
                numero_items_proveedores = numero_items_proveedores - 1
                var_operacion_bitacora = "E"
                MsgBox "Se Elimino Correctamente el Registro", vbInformation
                lv_proveedores.ListItems.Remove (lv_proveedores.selectedItem.Index)
                Call pro_limpiatextos(Me)
                txt_registros = lv_proveedores.ListItems.Count
                lv_proveedores.selectedItem.Selected = True
                pro_textos
           Else
                MsgBox "No se puede grabar registro: " + TB_PROVEEDORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
        End If
salir:
Set TB_PROVEEDORES = Nothing

End Sub


Sub pro_llena_listview1()
   
   Dim list_item As ListItem
   numero_items_proveedores = 0
   rs.Open "select * from TB_PROVEEDORES order by vcha_pro_nombre", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
      Set list_item = lv_proveedores.ListItems.Add(, , rs!VCHA_PRO_PROVEEDOR_ID)
      list_item.SubItems(1) = IIf(IsNull(rs!VCHA_PRO_NOMBRE), "", rs!VCHA_PRO_NOMBRE)
      list_item.SubItems(2) = IIf(IsNull(rs!vcha_pro_representante), "", rs!vcha_pro_representante)
      list_item.SubItems(3) = IIf(IsNull(rs!dtim_pro_fecha_alta), "", rs!dtim_pro_fecha_alta)
      list_item.SubItems(4) = IIf(IsNull(rs!vcha_pro_rfc), "", rs!vcha_pro_rfc)
      list_item.SubItems(5) = IIf(IsNull(rs!vcha_pai_pais_id), "", rs!vcha_pai_pais_id)
      list_item.SubItems(6) = IIf(IsNull(rs!vcha_est_estado_id), "", rs!vcha_est_estado_id)
      list_item.SubItems(7) = IIf(IsNull(rs!vcha_mun_municipio_id), "", rs!vcha_mun_municipio_id)
      list_item.SubItems(8) = IIf(IsNull(rs!vcha_ciu_ciudad_id), "", rs!vcha_ciu_ciudad_id)
      list_item.SubItems(9) = IIf(IsNull(rs!vcha_pro_colonia), "", rs!vcha_pro_colonia)
      list_item.SubItems(10) = IIf(IsNull(rs!vcha_pro_direccion), "", rs!vcha_pro_direccion)
      list_item.SubItems(11) = IIf(IsNull(rs!vcha_pro_cp), "", rs!vcha_pro_cp)
      list_item.SubItems(12) = IIf(IsNull(rs!vcha_pro_telefono), "", rs!vcha_pro_telefono)
      numero_items_proveedores = numero_items_proveedores + 1
      varpais = IIf(IsNull(rs!vcha_pai_pais_id), "", rs!vcha_pai_pais_id)
      varestado = IIf(IsNull(rs!vcha_est_estado_id), "", rs!vcha_est_estado_id)
      varmunicipio = IIf(IsNull(rs!vcha_mun_municipio_id), "", rs!vcha_mun_municipio_id)
      varciudad = IIf(IsNull(rs!vcha_ciu_ciudad_id), "", rs!vcha_ciu_ciudad_id)
      varcolonia = IIf(IsNull(rs!vcha_pro_colonia), "", rs!vcha_pro_colonia)
      
      rs.MoveNext:
   Wend
   rs.Close
   If numero_items_proveedores > 10 Then
      lv_proveedores.ColumnHeaders(2).Width = 4160
   Else
      lv_proveedores.ColumnHeaders(2).Width = 4360.25
   End If
End Sub


Sub pro_textos()
On Error GoTo err0:
        txt_clave = lv_proveedores.selectedItem
        txt_nombre = lv_proveedores.selectedItem.SubItems(1)
        txt_representante = lv_proveedores.selectedItem.SubItems(2)
        txt_fecha_alta = lv_proveedores.selectedItem.SubItems(3)
        txt_rfc = lv_proveedores.selectedItem.SubItems(4)
        varpais = lv_proveedores.selectedItem.SubItems(5)
        varestado = lv_proveedores.selectedItem.SubItems(6)
        varmunicipio = lv_proveedores.selectedItem.SubItems(7)
        varciudad = lv_proveedores.selectedItem.SubItems(8)
        varcolonia = lv_proveedores.selectedItem.SubItems(9)
        txt_domicilio = lv_proveedores.selectedItem.SubItems(10)
        txt_codigo_postal = lv_proveedores.selectedItem.SubItems(11)
        txt_telefono = lv_proveedores.selectedItem.SubItems(12)
        var_hubo_cambios = False
        
      rs.Open "SELECT * FROM TB_PAISES WHERE VCHA_PAI_PAIS_ID = '" + varpais + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_pais = IIf(IsNull(rs!VCHA_PAI_NOMBRE), "", rs!VCHA_PAI_NOMBRE)
      Else
         txt_nombre_pais = ""
      End If
      rs.Close
      rs.Open "SELECT * FROM TB_ESTADOS WHERE VCHA_EST_ESTADO_ID = '" + varestado + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
      Else
         txt_nombre_estado = ""
      End If
      rs.Close
      rs.Open "SELECT * FROM TB_MUNICIPIOS WHERE VCHA_MUN_MUNICIPIO_ID = '" + varmunicipio + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_municipio = IIf(IsNull(rs!VCHA_MUN_NOMBRE), "", rs!VCHA_MUN_NOMBRE)
      Else
         txt_nombre_municipio = ""
      End If
      rs.Close
      rs.Open "SELECT * FROM TB_CIUDADES WHERE VCHA_CIU_CIUDAD_ID = '" + varciudad + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
      Else
         txt_nombre_ciudad = ""
      End If
      rs.Close
      rs.Open "SELECT * FROM TB_COLONIAS WHERE VCHA_COL_COLONIA_ID = '" + varcolonia + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_colonia = IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
      Else
         txt_nombre_colonia = ""
      End If
      rs.Close
      Me.txt_clave.Enabled = False
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro_proveedor = False Then
        Set list_item = lv_proveedores.ListItems.Add(, , txt_clave)
        list_item.SubItems(1) = txt_nombre
        list_item.SubItems(2) = txt_representante
        list_item.SubItems(3) = txt_fecha_alta
        list_item.SubItems(4) = txt_rfc
        list_item.SubItems(5) = varpais
        list_item.SubItems(6) = varestado
        list_item.SubItems(7) = varmunicipio
        list_item.SubItems(8) = varciudad
        list_item.SubItems(9) = varcolonia
        list_item.SubItems(10) = Me.txt_domicilio
        list_item.SubItems(11) = txt_codigo_postal
        list_item.SubItems(12) = txt_telefono
        list_item.EnsureVisible
        list_item.Selected = True
        numero_items_proveedores = numero_items_proveedores + 1
    Else
        lv_proveedores.ListItems.Item(lv_proveedores.selectedItem.Index).Checked = False
        lv_proveedores.ListItems.Item(lv_proveedores.selectedItem.Index) = txt_clave
        lv_proveedores.ListItems.Item(lv_proveedores.selectedItem.Index).ListSubItems(1) = txt_nombre
        lv_proveedores.ListItems.Item(lv_proveedores.selectedItem.Index).ListSubItems(2) = txt_representante
        lv_proveedores.ListItems.Item(lv_proveedores.selectedItem.Index).ListSubItems(3) = txt_fecha_alta
        lv_proveedores.ListItems.Item(lv_proveedores.selectedItem.Index).ListSubItems(4) = txt_rfc
        lv_proveedores.ListItems.Item(lv_proveedores.selectedItem.Index).ListSubItems(5) = varpais
        lv_proveedores.ListItems.Item(lv_proveedores.selectedItem.Index).ListSubItems(6) = varestado
        lv_proveedores.ListItems.Item(lv_proveedores.selectedItem.Index).ListSubItems(7) = varmunicipio
        lv_proveedores.ListItems.Item(lv_proveedores.selectedItem.Index).ListSubItems(8) = varciudad
        lv_proveedores.ListItems.Item(lv_proveedores.selectedItem.Index).ListSubItems(9) = varcolonia
        lv_proveedores.ListItems.Item(lv_proveedores.selectedItem.Index).ListSubItems(10) = txt_domicilio
        lv_proveedores.ListItems.Item(lv_proveedores.selectedItem.Index).ListSubItems(11) = txt_codigo_postal
        lv_proveedores.ListItems.Item(lv_proveedores.selectedItem.Index).ListSubItems(12) = txt_telefono
    End If
   If numero_items_proveedores > 10 Then
      lv_proveedores.ColumnHeaders(2).Width = 4160
   Else
      lv_proveedores.ColumnHeaders(2).Width = 4360.25
   End If
    lv_proveedores.SetFocus
End Sub





Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_proveedores, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub





Private Sub txt_clave_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_clave_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_colonia_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_colonia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   var_hubo_cambios = True
   var_guardar_cambios = True
End Sub

Private Sub txt_codigo_postal_Change()
   var_hubo_cambios = True
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
         rs.Open "select distinct vcha_pai_pais_id from tb_colonias where vcha_col_cp = '" + txt_codigo_postal + "'", cnn, adOpenDynamic, adLockOptimistic
         Dim var_ren As Integer
         var_ren = rs.RecordCount
         rs.Close
         If var_ren > 1 Then
            lv_lista.ListItems.Clear
            rsaux.Open "select DISTINCT VCHA_PAI_PAIS_ID, VCHA_PAI_NOMBRE from vw_colonias order by vcha_pai_nombre", cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux.EOF
                  Set list_item = lv_lista.ListItems.Add(, , rsaux!vcha_pai_pais_id)
                  list_item.SubItems(1) = IIf(IsNull(rsaux!VCHA_PAI_NOMBRE), "", rsaux!VCHA_PAI_NOMBRE)
                  rsaux.MoveNext
            Wend
            rsaux.Close
            lbl_lista = "SELECCIONE EL PAIS"
            var_tipo_lista = 16
             var_n = lv_lista.ListItems.Count
            If var_n > 6 Then
               lv_lista.ColumnHeaders(2).Width = 4270.71
            Else
               lv_lista.ColumnHeaders(2).Width = 4499.71
            End If
            frm_lista.Visible = True
            lv_lista.SetFocus
         Else
            rs.Open "select * from vw_colonias where vcha_col_cp = '" + txt_codigo_postal + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               If rs.RecordCount = 1 Then
                  varcolonia = IIf(IsNull(rs!VCHA_COL_COLONIA_ID), "", rs!VCHA_COL_COLONIA_ID)
                  txt_nombre_colonia = IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
                  varpais = IIf(IsNull(rs!vcha_pai_pais_id), "", rs!vcha_pai_pais_id)
                  txt_nombre_pais = IIf(IsNull(rs!VCHA_PAI_NOMBRE), "", rs!VCHA_PAI_NOMBRE)
                  varestado = IIf(IsNull(rs!vcha_est_estado_id), "", rs!vcha_est_estado_id)
                  txt_nombre_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                  varmunicipio = IIf(IsNull(rs!vcha_mun_municipio_id), "", rs!vcha_mun_municipio_id)
                  txt_nombre_municipio = IIf(IsNull(rs!VCHA_MUN_NOMBRE), "", rs!VCHA_MUN_NOMBRE)
                  varciudad = IIf(IsNull(rs!vcha_ciu_ciudad_id), "", rs!vcha_ciu_ciudad_id)
                  txt_nombre_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                  txt_telefono.SetFocus
               Else
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

Private Sub txt_domicilio_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_domicilio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
   var_hubo_cambios = True
   var_guardar_cambios = True
End Sub

Private Sub txt_fecha_alta_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_fecha_alta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
   var_hubo_cambios = True
   var_guardar_cambios = True
End Sub

Private Sub txt_nombre_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
   var_hubo_cambios = True
   var_guardar_cambios = True
End Sub


Private Sub txt_representante_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_representante_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
   var_hubo_cambios = True
   var_guardar_cambios = True
End Sub

Private Sub txt_rfc_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_rfc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
   var_hubo_cambios = True
   var_guardar_cambios = True
End Sub

Private Sub txt_telefono_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_telefono_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   var_hubo_cambios = True
   var_guardar_cambios = True
   If KeyAscii = 13 Then
      cmd_guardar.SetFocus
   End If
End Sub
