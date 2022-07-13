VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmagentes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo de Agentes"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   ControlBox      =   0   'False
   Icon            =   "frmagentes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11670
   Begin VB.Frame Frame5 
      Height          =   615
      Left            =   5880
      TabIndex        =   80
      Top             =   480
      Width           =   5640
      Begin VB.TextBox txt_zona 
         Height          =   315
         Left            =   660
         TabIndex        =   33
         Top             =   195
         Width           =   900
      End
      Begin VB.TextBox txt_nombre_zona 
         Height          =   315
         Left            =   1575
         MaxLength       =   50
         TabIndex        =   34
         Top             =   195
         Width           =   3975
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Zona:"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   81
         Top             =   255
         Width           =   420
      End
   End
   Begin VB.Frame frm_contraseña 
      Height          =   1485
      Left            =   1800
      TabIndex        =   74
      Top             =   5535
      Width           =   3810
      Begin VB.TextBox txtusuario 
         Height          =   315
         Left            =   1815
         MaxLength       =   13
         TabIndex        =   77
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtpassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1815
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   76
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Clave:"
         Height          =   195
         Left            =   840
         TabIndex        =   79
         Top             =   660
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Contraseña:"
         Height          =   195
         Left            =   840
         TabIndex        =   78
         Top             =   1005
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   135
         Picture         =   "frmagentes.frx":08CA
         Top             =   600
         Width           =   480
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         Caption         =   " Contraseña de acceso"
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   45
         TabIndex        =   75
         Top             =   135
         Width           =   3705
      End
   End
   Begin VB.Frame frm_colonias 
      Height          =   2400
      Left            =   2055
      TabIndex        =   68
      Top             =   3765
      Width           =   5685
      Begin MSComctlLib.ListView lv_colonias 
         Height          =   1830
         Left            =   45
         TabIndex        =   69
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
         TabIndex        =   70
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame frm_lista 
      Height          =   2385
      Left            =   3540
      TabIndex        =   48
      Top             =   2595
      Width           =   5670
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1875
         Left            =   30
         TabIndex        =   49
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
         TabIndex        =   50
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11205
      Picture         =   "frmagentes.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1470
      Picture         =   "frmagentes.frx":17CE
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
      Picture         =   "frmagentes.frx":18D0
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
      Picture         =   "frmagentes.frx":19D2
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
      Picture         =   "frmagentes.frx":1AA4
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
      Picture         =   "frmagentes.frx":1BA6
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
      Left            =   -60
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   46
      Top             =   1230
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   " Agentes "
      Height          =   6780
      Left            =   150
      TabIndex        =   0
      Top             =   480
      Width           =   5655
      Begin VB.TextBox txt_oracle 
         Height          =   315
         Left            =   3855
         MaxLength       =   50
         TabIndex        =   31
         Top             =   6390
         Width           =   1290
      End
      Begin VB.TextBox txt_anterior 
         Height          =   315
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   30
         Top             =   6390
         Width           =   1290
      End
      Begin VB.TextBox txt_nombre_canal_venta 
         Height          =   315
         Left            =   2265
         MaxLength       =   50
         TabIndex        =   15
         Top             =   1905
         Width           =   3300
      End
      Begin VB.TextBox txt_canal_venta 
         Height          =   315
         Left            =   1350
         TabIndex        =   14
         Top             =   1905
         Width           =   900
      End
      Begin VB.ComboBox cmb_estatus 
         Height          =   315
         ItemData        =   "frmagentes.frx":1CA8
         Left            =   1350
         List            =   "frmagentes.frx":1CB2
         TabIndex        =   16
         Top             =   2250
         Width           =   1890
      End
      Begin VB.TextBox txt_email 
         Height          =   315
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   29
         Top             =   6060
         Width           =   4110
      End
      Begin VB.TextBox txt_codigo_postal 
         Height          =   315
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   23
         Top             =   3975
         Width           =   1005
      End
      Begin VB.TextBox txt_domicilio 
         Height          =   315
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   22
         Top             =   3630
         Width           =   4110
      End
      Begin VB.TextBox txt_nombre_pais 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   59
         Top             =   5715
         Width           =   4110
      End
      Begin VB.TextBox txt_nombre_estado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   58
         Top             =   5370
         Width           =   4110
      End
      Begin VB.TextBox txt_nombre_ciudad 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   57
         Top             =   4680
         Width           =   4110
      End
      Begin VB.TextBox txt_nombre_colonia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   56
         Top             =   4335
         Width           =   4110
      End
      Begin VB.TextBox txt_nombre_municipio 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   55
         Top             =   5025
         Width           =   4110
      End
      Begin VB.CommandButton cmdcomisiones 
         Height          =   285
         Left            =   5190
         Picture         =   "frmagentes.frx":1CC8
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Comisiones"
         Top             =   6420
         Width           =   315
      End
      Begin VB.CommandButton cmdfecha 
         Height          =   285
         Index           =   0
         Left            =   2655
         Picture         =   "frmagentes.frx":1DCA
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Seleccione la fecha"
         Top             =   3300
         Width           =   315
      End
      Begin VB.TextBox txt_fecha_alta 
         Height          =   315
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   20
         Top             =   3285
         Width           =   1245
      End
      Begin VB.TextBox txt_ruta_archivos 
         Height          =   315
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   19
         Top             =   2940
         Width           =   4200
      End
      Begin VB.TextBox txt_clave_almacen 
         Height          =   315
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   17
         Top             =   2595
         Width           =   900
      End
      Begin VB.TextBox txt_nombre_almacen 
         Height          =   315
         Left            =   2265
         MaxLength       =   50
         TabIndex        =   18
         Top             =   2595
         Width           =   3300
      End
      Begin VB.TextBox txt_nombre_tipo_agente 
         Height          =   315
         Left            =   2265
         MaxLength       =   50
         TabIndex        =   13
         Top             =   1560
         Width           =   3300
      End
      Begin VB.TextBox txt_nombre_empresa 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2265
         MaxLength       =   50
         TabIndex        =   8
         Top             =   180
         Width           =   3300
      End
      Begin VB.TextBox txt_clave_tipo_agente 
         Height          =   315
         Left            =   1350
         TabIndex        =   12
         Top             =   1560
         Width           =   900
      End
      Begin VB.TextBox txt_telefono 
         Height          =   315
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1215
         Width           =   1620
      End
      Begin VB.TextBox txt_nombre_agente 
         Height          =   315
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   10
         Top             =   870
         Width           =   4185
      End
      Begin VB.TextBox txt_clave_agente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1350
         TabIndex        =   9
         Top             =   525
         Width           =   900
      End
      Begin VB.TextBox txt_clave_empresa 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1350
         TabIndex        =   7
         Top             =   180
         Width           =   900
      End
      Begin VB.TextBox txt_municipio 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   26
         Top             =   5025
         Width           =   1005
      End
      Begin VB.TextBox txt_colonia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   24
         Top             =   4335
         Width           =   1005
      End
      Begin VB.TextBox txt_ciudad 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   25
         Top             =   4680
         Width           =   1005
      End
      Begin VB.TextBox txt_estado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   27
         Top             =   5370
         Width           =   1005
      End
      Begin VB.TextBox txt_pais 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   28
         Top             =   5715
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Clave ORACLE:"
         Height          =   195
         Left            =   2655
         TabIndex        =   73
         Top             =   6450
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clave Anterior:"
         Height          =   195
         Left            =   90
         TabIndex        =   72
         Top             =   6435
         Width           =   1035
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Canal de Venta:"
         Height          =   195
         Index           =   8
         Left            =   105
         TabIndex        =   71
         Top             =   1965
         Width           =   1140
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail"
         Height          =   195
         Index           =   16
         Left            =   120
         TabIndex        =   67
         Top             =   6120
         Width           =   435
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Pais:"
         Height          =   195
         Index           =   17
         Left            =   120
         TabIndex        =   66
         Top             =   5775
         Width           =   345
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Index           =   18
         Left            =   120
         TabIndex        =   65
         Top             =   5430
         Width           =   540
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad:"
         Height          =   195
         Index           =   19
         Left            =   120
         TabIndex        =   64
         Top             =   4740
         Width           =   540
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Colonia:"
         Height          =   195
         Index           =   20
         Left            =   120
         TabIndex        =   63
         Top             =   4395
         Width           =   570
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Domicilio:"
         Height          =   195
         Index           =   21
         Left            =   120
         TabIndex        =   62
         Top             =   3645
         Width           =   675
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "CP:"
         Height          =   195
         Index           =   22
         Left            =   120
         TabIndex        =   61
         Top             =   4020
         Width           =   255
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Municipio:"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   60
         Top             =   5085
         Width           =   720
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Alta:"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   53
         Top             =   3345
         Width           =   810
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Ruta Archivos:"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   52
         Top             =   3000
         Width           =   1050
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Almacen:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   51
         Top             =   2655
         Width           =   660
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Estatus:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   45
         Top             =   2310
         Width           =   570
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Tipo agentes:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   44
         Top             =   1620
         Width           =   975
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   43
         Top             =   1275
         Width           =   675
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   38
         Top             =   930
         Width           =   555
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   37
         Top             =   600
         Width           =   450
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Empresa:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   255
         Width           =   660
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   5880
      TabIndex        =   39
      Top             =   1065
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   1950
         TabIndex        =   35
         Top             =   150
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3510
         TabIndex        =   47
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
         Caption         =   "Busqueda de agente:"
         Height          =   195
         Left            =   195
         TabIndex        =   40
         Top             =   195
         Width           =   1530
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5685
      Left            =   5880
      TabIndex        =   41
      Top             =   1575
      Width           =   5655
      Begin MSComctlLib.ListView lv_agentes 
         Height          =   5430
         Left            =   45
         TabIndex        =   54
         Top             =   210
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   9578
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
         NumItems        =   26
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nombre"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "telefono"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "tipoagente"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "estatus"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Nombre empresa"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Nombre Tipo Agente"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Nombre Zona"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "clave almacen"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Nombre Almacen"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Ruta archivos"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Fecha Alta"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "domicilio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "cp"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "colonia"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "ciudad"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "municipio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "estado"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "pais"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "email"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "canal venta"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "anterior"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Text            =   "oracle"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Text            =   "Zona"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   25
            Text            =   "Nombre zona"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   -75
      Top             =   480
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
            Picture         =   "frmagentes.frx":1ECC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmagentes.frx":27A6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   42
      Top             =   285
      Width           =   11475
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   -60
      Top             =   870
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
            Picture         =   "frmagentes.frx":3080
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmagentes.frx":395A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmagentes.frx":4234
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmagentes.frx":47D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmagentes.frx":50AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmagentes.frx":5986
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmagentes.frx":6260
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmagentes.frx":6372
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmagentes.frx":6484
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmagentes.frx":6596
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmagentes.frx":66A8
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmagentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_hubo_cambios As Boolean
Dim numero_items_agentes As Integer
Dim var_tipo_lista As Integer

Private Sub cmb_estatus_Change()
   var_hubo_cambios = True
End Sub

Private Sub cmb_estatus_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
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
         rsaux5.Open "select vcha_emp_conexion from tb_empresas where len(vcha_emp_conexion) > 0", cnn, adOpenDynamic, adLockOptimistic
         While Not rsaux5.EOF
               var_conexion_importacion = IIf(IsNull(rsaux5(0).Value), "", rsaux5(0).Value)
               If Trim(var_conexion_importacion) <> "" Then
                  If cnn_importacion.State = 1 Then
                     cnn_importacion.Close
                  End If
                  cnn_importacion.Open var_conexion_importacion
                  Call pro_elimina_agentes
                End If
                rsaux5.MoveNext
         Wend
         rsaux5.Close
      
         MsgBox "Se Elimino Correctamente el Registro", vbInformation
         lv_agentes.ListItems.Remove (lv_agentes.selectedItem.Index)
         numero_items_agentes = numero_items_agentes - 1
         Call pro_limpiatextos(Me)
         txt_registros = lv_agentes.ListItems.Count
         lv_agentes.selectedItem.Selected = True
         pro_textos
      
      
      
         rs.Open "select * from tb_agentes", cnn, adOpenDynamic, adLockOptimistic
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
   Dim var_existe As Boolean
   If Me.txt_nombre_agente <> "" And Me.txt_clave_tipo_agente <> "" And Me.txt_canal_venta <> "" Then
      rs.Open "select * from tb_agentes where vcha_age_agente_id = '" + txt_clave_agente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_existe = True
      End If
      rs.Close
      If var_modifica_registro_agente = False And var_existe = True Then
         MsgBox "La clave del agente ya existe", vbOKOnly, "ATENCION"
      Else
         If IsDate(txt_fecha_alta) Then
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
               rsaux5.Open "select vcha_emp_conexion from tb_empresas where len(vcha_emp_conexion) > 0 and vcha_emp_empresa_id = '02' ORDER BY INTE_EMP_ORDEN_CONEXION", cnn, adOpenDynamic, adLockOptimistic
               While Not rsaux5.EOF
                     var_conexion_importacion = IIf(IsNull(rsaux5(0).Value), "", rsaux5(0).Value)
                     If Trim(var_conexion_importacion) <> "" Then
                        If cnn_importacion.State = 1 Then
                           cnn_importacion.Close
                        End If
                        cnn_importacion.Open var_conexion_importacion
                        Call pro_guardar_agentes
                     End If
                     rsaux5.MoveNext
               Wend
               rsaux5.Close
               pro_actualiza_ListView
               txt_clave_empresa.Enabled = False
               MsgBox "Informacion Guardada Correctamente!", vbOKOnly + vbInformation, "Aviso"
               txt_registros = lv_agentes.ListItems.Count
               var_modifica_registro_agente = True
               var_hubo_cambios = False
               
               rs.Open "select * from tb_agentes where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
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
               var_n = lv_agentes.ListItems.Count
               If var_n > 0 Then
                  Me.lv_agentes.SetFocus
               Else
                  Me.cmd_nuevo.SetFocus
               End If
            End If
         Else
            MsgBox "Fecha de alta incorrecta", vbOKOnly, ""
         End If
      End If
   Else
      MsgBox "Información incompleta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
   Set reporte = appl.OpenReport(App.Path + "\rep_catalogo_AGENTES.rpt")
   frmvistasprevias.cr.ReportSource = reporte
   For ntablas = 1 To reporte.Database.Tables.Count
       reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
   Next ntablas
   frmvistasprevias.cr.ViewReport
   frmvistasprevias.Caption = "Catálogo de Clientes"
   frmvistasprevias.Show
   Set reporte = Nothing
End Sub

Private Sub cmd_nuevo_Click()
   Call pro_limpiatextos(Me)
   cmb_estatus = "ACTIVO"
   rs.Open "SELECT * FROM TB_EMPRESAS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      txt_clave_empresa = var_empresa
      txt_nombre_empresa = IIf(IsNull(rs!VCHA_EMP_NOMBRE), "", rs!VCHA_EMP_NOMBRE)
   Else
      txt_clave_empresa = ""
      MsgBox "Clave de empresa incorrecta", vbOKOnly, "ATENCION"
      txt_nombre_empresa = ""
   End If
   rs.Close
   Me.txt_fecha_alta = Date
   txt_nombre_agente.Enabled = False
   txt_clave_agente.Enabled = False
   txt_nombre_agente.Enabled = True
   txt_nombre_agente.SetFocus: var_modifica_registro_agente = False
   txt_nombre_agente.Enabled = True
   cmd_guardar.Enabled = True
   cmd_deshacer.Enabled = True
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_agente = False Then
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
   var_modifica_registro_agente = False
   Unload Me
   Exit Sub
salir:
End Sub

Private Sub cmdcomisiones_Click()
   Me.txtusuario = ""
   Me.txtpassword = ""
   Me.frm_contraseña.Visible = True
   Me.txtusuario.SetFocus
End Sub

Private Sub cmdfecha_Click(Index As Integer)
   If Trim(txt_fecha_alta) <> "" Then
      If IsDate(txt_fecha_alta) Then
         frmcalendario.mes.Value = CDate(txt_fecha_alta)
      Else
         frmcalendario.mes.Value = Date
      End If
   Else
      frmcalendario.mes.Value = Date
   End If
   frmcalendario.Show 1
   txt_fecha_alta = var_fecha_general
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
   Me.frm_contraseña.Visible = False
   var_cadena_seguridad = ""
   frm_colonias.Visible = False
   var_tipo_lista = 0
   frm_lista.Visible = False
   frmagentes.Top = 0
   frmagentes.Left = 0
   numero_items_agentes = 0
   var_modifica_registro_agente = True
   lv_agentes.SmallIcons = ImageList1
   'Call pro_encabezadosView(Me, lv_agentes, False)
   Call pro_llena_listview1
   pro_textos
   rs.Open "select * from tb_agentes", cnn, adOpenDynamic, adLockOptimistic
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
   Call activa_forma(var_activa_forma_agentes)
End Sub

Private Sub lv_agentes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_agentes, ColumnHeader)
End Sub

Private Sub lv_agentes_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Set lv_agentes.selectedItem = Item
   pro_textos
   var_modifica_registro_agente = True
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
      txt_email.SetFocus
      frm_colonias.Visible = False
   End If
   If KeyAscii = 27 Then
      txt_codigo_postal.SetFocus
      frm_colonias.Visible = False
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
      If var_tipo_lista = 1 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_clave_empresa = lv_lista.selectedItem
            txt_nombre_empresa = lv_lista.selectedItem.SubItems(1)
         Else
            txt_clave_empresa = ""
            txt_nombre_empresa = ""
         End If
         txt_clave_empresa.SetFocus
      End If
      If var_tipo_lista = 2 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_clave_tipo_agente = lv_lista.selectedItem
            txt_nombre_tipo_agente = lv_lista.selectedItem.SubItems(1)
         Else
            txt_clave_tipo_agente = ""
            txt_nombre_tipo_agente = ""
         End If
         txt_clave_tipo_agente.SetFocus
      End If
      If var_tipo_lista = 3 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_zona = lv_lista.selectedItem
            txt_nombre_zona = lv_lista.selectedItem.SubItems(1)
         Else
            txt_zona = ""
            txt_nombre_zona = ""
         End If
         txt_zona.SetFocus
      End If
      If var_tipo_lista = 4 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_clave_almacen = lv_lista.selectedItem
            txt_nombre_almacen = lv_lista.selectedItem.SubItems(1)
         Else
            txt_clave_almacen = ""
            txt_nombre_almacen = ""
         End If
         txt_clave_almacen.SetFocus
      End If
      If var_tipo_lista = 5 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_canal_venta = lv_lista.selectedItem
            txt_nombre_canal_venta = lv_lista.selectedItem.SubItems(1)
         Else
            txt_canal_venta = ""
            txt_nombre_canal_venta = ""
         End If
         txt_canal_venta.SetFocus
      End If
      If var_tipo_lista = 16 Then
         If lv_lista.ListItems.Count > 0 Then
            rs.Open "select * from vw_colonias where vcha_col_cp = '" + Me.txt_codigo_postal + "' AND VCHA_PAI_PAIS_ID = '" + lv_lista.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
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
               'rs.MoveFirst
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
            txt_telefono.SetFocus
         End If
      End If
      If var_tipo_lista = 20 Then
         Me.txt_zona = Me.lv_lista.selectedItem
         Me.txt_nombre_zona = Me.lv_lista.selectedItem.SubItems(1)
         Me.txt_zona.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub Text2_Change()

End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_agentes.SetFocus
      Call pro_avanzar(Me, lv_agentes, Button)
      pro_textos
      lv_agentes.selectedItem.EnsureVisible
   End If
   If Button.Index = 1 Then
      lv_agentes.ListItems(1).Selected = True
      pro_textos
      lv_agentes.selectedItem.EnsureVisible
   End If
   If Button.Index = 4 Then
      numero_items_agentes = lv_agentes.ListItems.Count
      lv_agentes.ListItems(numero_items_agentes).Selected = True
      pro_textos
      lv_agentes.selectedItem.EnsureVisible
   End If
err0:
End Sub


Sub pro_guardar_agentes()
Dim ok As Boolean
Set TB_AGENTES = New TB_AGENTES
Set TB_BITACORA_AGENTES = New TB_BITACORA_AGENTES
Dim txt_estatus As String
Dim var_fecha_inicio As String
   If cmb_estatus = "ACTIVO" Then
      txt_estatus = "A"
   End If
   If cmb_estatus = "INACTIVO" Then
      txt_estatus = "I"
   End If
   If txt_clave_empresa <> "" Or txt_estatus <> "" Then
      If var_hubo_cambios Then
         rs.Open "select * from tb_agentes where vcha_age_agente_id = '" + txt_clave_agente + "'", cnn_importacion, adOpenDynamic, adLockOptimistic
         var_dia = CStr(Day(CDate(txt_fecha_alta)))
         var_mes = CStr(Month(CDate(txt_fecha_alta)))
         var_año = CStr(Year(CDate(txt_fecha_alta)))
         If Len(Trim(var_dia)) = 1 Then
            var_dia = "0" + var_dia
         End If
         If Len(Trim(var_mes)) = 1 Then
            var_mes = "0" + var_mes
         End If
         var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
         var_agente_regreso = txt_clave_agente
         ok = TB_AGENTES.Anadir(txt_clave_empresa, txt_clave_agente, txt_nombre_agente, txt_telefono, txt_clave_tipo_agente, txt_estatus, txt_clave_almacen, txt_ruta_archivos, txt_fecha_alta, txt_domicilio, txt_codigo_postal, txt_colonia, txt_ciudad, txt_municipio, txt_estado, txt_pais, txt_email, txt_canal_venta)
         If Trim(var_agente_regreso) <> "" Then
            txt_clave_agente = var_agente_regreso
         End If
         If ok Then
            bitacora = True
            rsaux4.Open "update tb_agentes set vcha_age_agente_anterior_id = '" + txt_anterior + "', VCHA_AGE_CLAVE_ORACLE = '" + Me.txt_oracle + "', vcha_zon_zona_id = '" + Me.txt_zona + "'  where vcha_age_agente_id = '" + txt_clave_agente + "'", cnn_importacion, adOpenDynamic, adLockOptimistic
            If var_modifica_registro_agente = False Then
               var_operacion_bitacora = "I"
               var_agente_regreso = txt_clave_agente
               bitacora = TB_BITACORA_AGENTES.Anadir(txt_clave_agente, "VCHA_AGE_NOMBRE", var_operacion_bitacora, "", txt_nombre_agente, var_clave_usuario_global, fun_NombrePc, Date)
               txt_clave_agente = var_agente_regreso
            Else
               var_operacion_bitacora = "M"
               If rs!VCHA_EMP_EMPRESA_ID <> txt_clave_empresa Then
                  bitacora = TB_BITACORA_AGENTES.Anadir(txt_clave_agente, "VCHA_EMP_EMPRESA_ID", var_operacion_bitacora, rs!VCHA_EMP_EMPRESA_ID, txt_clave_empresa, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!VCHA_AGE_AGENTE_ID <> txt_clave_agente Then
                  bitacora = TB_BITACORA_AGENTES.Anadir(txt_clave_agente, "VCHA_AGE_AGENTE_ID", var_operacion_bitacora, rs!VCHA_AGE_AGENTE_ID, txt_clave_agente, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!VCHA_AGE_NOMBRE <> txt_nombre_agente Then
                  bitacora = TB_BITACORA_AGENTES.Anadir(txt_clave_agente, "VCHA_AGE_NOMBRE", var_operacion_bitacora, rs!VCHA_AGE_NOMBRE, txt_nombre_agente, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!VCHA_AGE_TELEFONO <> txt_telefono Then
                  bitacora = TB_BITACORA_AGENTES.Anadir(txt_clave_agente, "VCHA_AGE_TELEFONO", var_operacion_bitacora, rs!VCHA_AGE_TELEFONO, txt_telefono, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!vcha_tag_tipoagente_id <> txt_clave_tipo_agente Then
                  bitacora = TB_BITACORA_AGENTES.Anadir(txt_clave_agente, "VCHA_TAG_TIPOAGENTE_ID ", var_operacion_bitacora, rs!vcha_tag_tipoagente_id, txt_clave_tipo_agente, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!VCHA_AGE_ESTATUS <> txt_estatus Then
                  bitacora = TB_BITACORA_AGENTES.Anadir(txt_clave_agente, "VCHA_AGE_ESTATUS", var_operacion_bitacora, rs!VCHA_AGE_ESTATUS, txt_estatus, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!VCHA_ALM_ALMACEN_ID <> txt_clave_almacen Then
                  bitacora = TB_BITACORA_AGENTES.Anadir(txt_clave_agente, "VCHA_ALM_ALMACEN_ID", var_operacion_bitacora, rs!VCHA_ALM_ALMACEN_ID, txt_clave_almacen, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!VCHA_AGE_RUTA_ARCHIVOS <> txt_ruta_archivos Then
                  bitacora = TB_BITACORA_AGENTES.Anadir(txt_clave_agente, "VCHA_AGE_RUTA_ARCHIVOS", var_operacion_bitacora, rs!VCHA_AGE_RUTA_ARCHIVOS, txt_ruta_archivos, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs!DTIM_AGE_FECHA_ALTA <> txt_fecha_alta Then
                  bitacora = TB_BITACORA_AGENTES.Anadir(txt_clave_agente, "DTIM_AGE_FECHA_ALTA", var_operacion_bitacora, rs!DTIM_AGE_FECHA_ALTA, txt_fecha_alta, var_clave_usuario_global, fun_NombrePc, Date)
               End If
            End If
            rs.Close
         Else
            MsgBox "No se puede grabar registro: " + TB_AGENTES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
         End If
      End If
   End If
 Set TB_AGENTES = Nothing

End Sub

Sub pro_elimina_agentes()
   Dim var_llave_usuarios As String
   Set TB_AGENTES = New TB_AGENTES
   Set TB_BITACORA_AGENTES = New TB_BITACORA_AGENTES
   On Error GoTo salir
   ok = True
   If txt_clave_empresa <> "" And txt_clave_agente <> "" And var_modifica_registro_agente = True Then
      'If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_AGENTES.Eliminar(txt_clave_agente)
      'Else
      '   GoTo salir:
      'End If
      If ok Then
         var_operacion_bitacora = "E"
         bitacora = TB_BITACORA_AGENTES.Anadir(txt_clave_agente, "VCHA_AGE_NOMBRE", var_operacion_bitacora, "", txt_nombre_agente, var_clave_usuario_global, fun_NombrePc, Date)
      Else
         MsgBox "No se puede grabar registro: " + TB_AGENTES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
Set TB_AGENTES = Nothing
End Sub


Sub pro_llena_listview1()
   Dim list_item As ListItem
   rs.Open "select * from VW_CATALOGO_AGENTES where vcha_Emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
      Set list_item = lv_agentes.ListItems.Add(, , rs!VCHA_EMP_EMPRESA_ID)
      list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
      list_item.SubItems(2) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
      list_item.SubItems(3) = IIf(IsNull(rs!VCHA_AGE_TELEFONO), "", rs!VCHA_AGE_TELEFONO)
      list_item.SubItems(4) = IIf(IsNull(rs!vcha_tag_tipoagente_id), "", rs!vcha_tag_tipoagente_id)
      list_item.SubItems(5) = IIf(IsNull(rs!VCHA_AGE_ESTATUS), "", rs!VCHA_AGE_ESTATUS)
      list_item.SubItems(6) = IIf(IsNull(rs!VCHA_EMP_NOMBRE), "", rs!VCHA_EMP_NOMBRE)
      list_item.SubItems(7) = IIf(IsNull(rs!VCHA_TAG_DESCRIPCION), "", rs!VCHA_TAG_DESCRIPCION)
      list_item.SubItems(8) = ""
      list_item.SubItems(9) = IIf(IsNull(rs!VCHA_ALM_ALMACEN_ID), "", rs!VCHA_ALM_ALMACEN_ID)
      list_item.SubItems(10) = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
      list_item.SubItems(11) = IIf(IsNull(rs!VCHA_AGE_RUTA_ARCHIVOS), "", rs!VCHA_AGE_RUTA_ARCHIVOS)
      list_item.SubItems(12) = IIf(IsNull(rs!DTIM_AGE_FECHA_ALTA), "", rs!DTIM_AGE_FECHA_ALTA)
      list_item.SubItems(13) = IIf(IsNull(rs!VCHA_AGE_DOMICILIO), "", rs!VCHA_AGE_DOMICILIO)
      list_item.SubItems(14) = IIf(IsNull(rs!VCHA_AGE_CP), "", rs!VCHA_AGE_CP)
      list_item.SubItems(15) = IIf(IsNull(rs!VCHA_COL_COLONIA_ID), "", rs!VCHA_COL_COLONIA_ID)
      list_item.SubItems(16) = IIf(IsNull(rs!VCHA_CIU_CIUDAD_ID), "", rs!VCHA_CIU_CIUDAD_ID)
      list_item.SubItems(17) = IIf(IsNull(rs!VCHA_MUN_MUNICIPIO_ID), "", rs!VCHA_MUN_MUNICIPIO_ID)
      list_item.SubItems(18) = IIf(IsNull(rs!VCHA_EST_ESTADO_ID), "", rs!VCHA_EST_ESTADO_ID)
      list_item.SubItems(19) = IIf(IsNull(rs!VCHA_PAI_PAIS_ID), "", rs!VCHA_PAI_PAIS_ID)
      list_item.SubItems(20) = IIf(IsNull(rs!VCHA_AGE_EMAIL), "", rs!VCHA_AGE_EMAIL)
      list_item.SubItems(21) = IIf(IsNull(rs!vcha_can_canal_venta_id), "", rs!vcha_can_canal_venta_id)
      list_item.SubItems(22) = IIf(IsNull(rs!VCHA_AGE_AGENTE_ANTERIOR_ID), "", rs!VCHA_AGE_AGENTE_ANTERIOR_ID)
      list_item.SubItems(23) = IIf(IsNull(rs!VCHA_AGE_CLAVE_ORACLE), "", rs!VCHA_AGE_CLAVE_ORACLE)
      list_item.SubItems(24) = IIf(IsNull(rs!vcha_zon_zona_id), "", rs!vcha_zon_zona_id)
      list_item.SubItems(25) = IIf(IsNull(rs!vcha_zon_descripcion), "", rs!vcha_zon_descripcion)
      rs.MoveNext:
      numero_items_agentes = numero_items_agentes + 1
   Wend
   rs.Close
End Sub


Sub pro_textos()
'On Error GoTo err0:
Dim var_n As Double
var_n = lv_agentes.ListItems.Count
   If var_n > 0 Then
      txt_clave_empresa = lv_agentes.selectedItem
      txt_clave_agente = lv_agentes.selectedItem.SubItems(1)
      txt_nombre_agente = lv_agentes.selectedItem.SubItems(2)
      txt_telefono = lv_agentes.selectedItem.SubItems(3)
      txt_clave_tipo_agente = lv_agentes.selectedItem.SubItems(4)
      txt_estatus = lv_agentes.selectedItem.SubItems(5)
      txt_nombre_empresa = lv_agentes.selectedItem.SubItems(6)
      txt_nombre_tipo_agente = lv_agentes.selectedItem.SubItems(7)
      txt_nombre_zona = lv_agentes.selectedItem.SubItems(8)
      txt_clave_almacen = lv_agentes.selectedItem.SubItems(9)
      txt_nombre_almacen = lv_agentes.selectedItem.SubItems(10)
      txt_ruta_archivos = lv_agentes.selectedItem.SubItems(11)
      txt_fecha_alta = lv_agentes.selectedItem.SubItems(12)
      txt_domicilio = lv_agentes.selectedItem.SubItems(13)
      txt_codigo_postal = lv_agentes.selectedItem.SubItems(14)
      txt_colonia = lv_agentes.selectedItem.SubItems(15)
      txt_ciudad = lv_agentes.selectedItem.SubItems(16)
      txt_municipio = lv_agentes.selectedItem.SubItems(17)
      txt_estado = lv_agentes.selectedItem.SubItems(18)
      txt_pais = lv_agentes.selectedItem.SubItems(19)
      txt_email = lv_agentes.selectedItem.SubItems(20)
      txt_canal_venta = lv_agentes.selectedItem.SubItems(21)
      txt_anterior = lv_agentes.selectedItem.SubItems(22)
      txt_oracle = lv_agentes.selectedItem.SubItems(23)
      Me.txt_zona = lv_agentes.selectedItem.SubItems(24)
      Me.txt_nombre_zona = lv_agentes.selectedItem.SubItems(25)
      txt_clave_agente.Enabled = False
      txt_clave_empresa.Enabled = False
      txt_nombre_empresa.Enabled = False
      If txt_estatus = "A" Then
         cmb_estatus = "ACTIVO"
      Else
         If txt_estatus = "I" Then
            cmb_estatus = "INACTIVO"
         Else
            cmb_estatus = ""
         End If
      End If
      rs.Open "SELECT * FROM TB_PAISES WHERE VCHA_PAI_PAIS_ID = '" + txt_pais + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_pais = IIf(IsNull(rs!vcha_pai_nombre), "", rs!vcha_pai_nombre)
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
         txt_nombre_municipio = IIf(IsNull(rs!vcha_mun_nombre), "", rs!vcha_mun_nombre)
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
      rs.Open "select * from tb_Canalesventas where vcha_can_canal_Venta_id = '" + txt_canal_venta + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_canal_venta = IIf(IsNull(rs!vcha_can_nombre), "", rs!vcha_can_nombre)
      Else
         txt_nombre_canal_venta = ""
      End If
      rs.Close
   End If
   var_numero_renglones = lv_agentes.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_agentes.ColumnHeaders(3).Width = 4200
   Else
      lv_agentes.ColumnHeaders(3).Width = 4499.71
   End If
   var_hubo_cambios = False
   var_modifica_registro_agente = True
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro_agente = False Then
        Set list_item = lv_agentes.ListItems.Add(, , txt_clave_empresa)
        list_item.SubItems(1) = txt_clave_agente
        list_item.SubItems(2) = txt_nombre_agente
        list_item.SubItems(3) = txt_telefono
        list_item.SubItems(4) = txt_clave_tipo_agente
        list_item.SubItems(5) = txt_estatus
        list_item.SubItems(6) = txt_nombre_empresa
        list_item.SubItems(7) = txt_nombre_tipo_agente
        list_item.SubItems(8) = txt_nombre_zona
        list_item.SubItems(9) = txt_clave_almacen
        list_item.SubItems(10) = txt_nombre_almacen
        list_item.SubItems(11) = txt_ruta_archivos
        list_item.SubItems(12) = txt_fecha_alta
        list_item.SubItems(13) = txt_domicilio
        list_item.SubItems(14) = txt_codigo_postal
        list_item.SubItems(15) = txt_colonia
        list_item.SubItems(16) = txt_ciudad
        list_item.SubItems(17) = txt_municipio
        list_item.SubItems(18) = txt_estado
        list_item.SubItems(19) = txt_pais
        list_item.SubItems(20) = txt_email
        list_item.SubItems(21) = txt_canal_venta
        list_item.SubItems(22) = txt_anterior
        list_item.SubItems(23) = txt_oracle
        list_item.SubItems(24) = Me.txt_zona
        list_item.SubItems(25) = Me.txt_nombre_zona
        list_item.EnsureVisible
        list_item.Selected = True
       numero_items_agentes = numero_items_agentes + 1
    Else
        lv_agentes.ListItems.Item(lv_agentes.selectedItem.Index).Checked = False
        lv_agentes.ListItems.Item(lv_agentes.selectedItem.Index) = txt_clave_empresa
        lv_agentes.ListItems.Item(lv_agentes.selectedItem.Index).ListSubItems(1) = txt_clave_agente
        lv_agentes.ListItems.Item(lv_agentes.selectedItem.Index).ListSubItems(2) = txt_nombre_agente
        lv_agentes.ListItems.Item(lv_agentes.selectedItem.Index).ListSubItems(3) = txt_telefono
        lv_agentes.ListItems.Item(lv_agentes.selectedItem.Index).ListSubItems(4) = txt_clave_tipo_agente
        lv_agentes.ListItems.Item(lv_agentes.selectedItem.Index).ListSubItems(5) = txt_estatus
        lv_agentes.ListItems.Item(lv_agentes.selectedItem.Index).ListSubItems(6) = txt_nombre_empresa
        lv_agentes.ListItems.Item(lv_agentes.selectedItem.Index).ListSubItems(7) = txt_nombre_tipo_agente
        lv_agentes.ListItems.Item(lv_agentes.selectedItem.Index).ListSubItems(8) = txt_nombre_zona
        lv_agentes.ListItems.Item(lv_agentes.selectedItem.Index).ListSubItems(9) = txt_clave_almacen
        lv_agentes.ListItems.Item(lv_agentes.selectedItem.Index).ListSubItems(10) = txt_nombre_almacen
        lv_agentes.ListItems.Item(lv_agentes.selectedItem.Index).ListSubItems(11) = txt_ruta_archivos
        lv_agentes.ListItems.Item(lv_agentes.selectedItem.Index).ListSubItems(12) = txt_fecha_alta
        lv_agentes.ListItems.Item(lv_agentes.selectedItem.Index).ListSubItems(13) = txt_domicilio
        lv_agentes.ListItems.Item(lv_agentes.selectedItem.Index).ListSubItems(14) = txt_cp
        lv_agentes.ListItems.Item(lv_agentes.selectedItem.Index).ListSubItems(15) = txt_colonia
        lv_agentes.ListItems.Item(lv_agentes.selectedItem.Index).ListSubItems(16) = txt_ciudad
        lv_agentes.ListItems.Item(lv_agentes.selectedItem.Index).ListSubItems(17) = txt_municipio
        lv_agentes.ListItems.Item(lv_agentes.selectedItem.Index).ListSubItems(18) = txt_estado
        lv_agentes.ListItems.Item(lv_agentes.selectedItem.Index).ListSubItems(19) = txt_pais
        lv_agentes.ListItems.Item(lv_agentes.selectedItem.Index).ListSubItems(20) = txt_email
        lv_agentes.ListItems.Item(lv_agentes.selectedItem.Index).ListSubItems(21) = txt_canal_venta
        lv_agentes.ListItems.Item(lv_agentes.selectedItem.Index).ListSubItems(22) = txt_anterior
        lv_agentes.ListItems.Item(lv_agentes.selectedItem.Index).ListSubItems(23) = txt_oracle
        lv_agentes.ListItems.Item(lv_agentes.selectedItem.Index).ListSubItems(24) = Me.txt_zona
        lv_agentes.ListItems.Item(lv_agentes.selectedItem.Index).ListSubItems(25) = Me.txt_nombre_zona
        lv_agentes.ListItems.Item(lv_agentes.selectedItem.Index).Selected = True
    End If
    lv_agentes.SetFocus
End Sub






Private Sub txt_anterior_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_anterior_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Me.cmd_guardar.Enabled = True Then
         Me.cmd_guardar.SetFocus
      End If
   End If
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(txt_buscar) <> "" Then
         Call pro_busca_registro(Me.lv_agentes, txt_buscar, True)
         txt_buscar = ""
         pro_textos
      End If
   End If
End Sub

Private Sub txt_canal_venta_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_canal_venta_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_canal_venta_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frm_lista.Left = 1500
      frm_lista.Top = 2715
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_CANALESVENTAS order by vcha_can_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_can_canal_venta_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_can_nombre), "", rs!vcha_can_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Canales de Venta"
      var_tipo_lista = 5
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 7 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
   If KeyCode = 117 Then
      Me.Enabled = False
      frmcanalesventas.Show
      var_activa_forma_canalesventas = Me.Name
   End If
End Sub

Private Sub txt_canal_venta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_canal_venta_LostFocus()
   If Trim(txt_canal_venta) <> "" Then
      rs.Open "select * from TB_CANALESVENTAS where vcha_Can_canal_venta_id = '" + txt_canal_venta + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_canal_venta = IIf(IsNull(rs!vcha_can_nombre), "", rs!vcha_can_nombre)
      Else
         txt_nombre_canal_venta = ""
         txt_canal_venta = ""
         MsgBox "Clave de canal de venta incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      txt_nombre_canal_venta = ""
   End If
End Sub

Private Sub txt_clave_agente_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_clave_agente_GotFocus()
   Me.frm_contraseña.Visible = False
End Sub

Private Sub txt_clave_agente_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_clave_almacen_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_clave_almacen_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_clave_almacen_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frm_lista.Left = 1500
      frm_lista.Top = 3405
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_almacenes order by vcha_alm_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ALM_ALMACEN_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      var_tipo_lista = 4
      lbl_lista = "Almacenes"
      var_n = lv_lista.ListItems.Count
      If var_n > 7 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
   If KeyCode = 117 Then
      Me.Enabled = False
      var_activa_forma_almacenes = Me.Name
      frmalmacenes.Show
   End If
End Sub

Private Sub txt_clave_almacen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_clave_almacen_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_clave_almacen) <> "" Then
      rs.Open "SELECT * FROM TB_ALMACENES WHERE VCHA_ALM_ALMACEN_ID = '" + txt_clave_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_almacen = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
      Else
         MsgBox "Clave de almacen incorrecta", vbOKOnly, "ATENCION"
         txt_nombre_almacen = ""
         txt_clave_almacen = ""
      End If
      rs.Close
   Else
      txt_nombre_almacen = ""
   End If
End Sub

Private Sub txt_clave_empresa_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_clave_empresa_GotFocus()
   Me.frm_contraseña.Visible = False
End Sub

Private Sub txt_clave_empresa_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_empresas order by vcha_emp_empresa_id", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_EMP_EMPRESA_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_EMP_NOMBRE), "", rs!VCHA_EMP_NOMBRE)
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
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_clave_tipo_agente_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_clave_tipo_agente_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_clave_tipo_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frm_lista.Left = 1500
      frm_lista.Top = 2385
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_tipoagentes order by vcha_tag_descripcion", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_tag_tipoagente_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_TAG_DESCRIPCION), "", rs!VCHA_TAG_DESCRIPCION)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Tipo Agentes"
      var_tipo_lista = 2
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 7 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
   If KeyCode = 117 Then
      Me.Enabled = False
      var_activa_forma_tipoagentes = Me.Name
      frmtipoagentes.Show
   End If
End Sub

Private Sub txt_clave_tipo_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_clave_tipo_agente_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_clave_tipo_agente) <> "" Then
      rs.Open "SELECT * FROM tb_tipoagentes where VCHA_TAG_TIPOAGENTE_ID = '" + txt_clave_tipo_agente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_tipo_agente = IIf(IsNull(rs!VCHA_TAG_DESCRIPCION), "", rs!VCHA_TAG_DESCRIPCION)
      Else
         MsgBox "Clave de tipo agente incorrecta", vbOKOnly, "ATENCION"
         txt_clave_tipo_agente = ""
         txt_nombre_tipo_agente = ""
      End If
      rs.Close
   Else
      txt_nombre_tipo_agente = ""
   End If
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
      frmagentes.Enabled = False
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
         rs.Open "select distinct vcha_pai_pais_id from tb_colonias where vcha_col_cp = '" + txt_codigo_postal + "'", cnn, adOpenDynamic, adLockOptimistic
         Dim var_ren As Integer
         var_ren = rs.RecordCount
         rs.Close
         If var_ren > 1 Then
            frm_lista.Left = 1500
            frm_lista.Top = 4800
            lv_lista.ListItems.Clear
            rsaux.Open "select DISTINCT VCHA_PAI_PAIS_ID, VCHA_PAI_NOMBRE from vw_colonias order by vcha_pai_nombre", cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux.EOF
                  Set list_item = lv_lista.ListItems.Add(, , rsaux!VCHA_PAI_PAIS_ID)
                  list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_pai_nombre), "", rsaux!vcha_pai_nombre)
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

Private Sub txt_email_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_email_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Me.cmd_guardar.Enabled = True Then
         Call pro_enfoque(KeyAscii)
      End If
   End If
End Sub



Private Sub txt_fecha_alta_Change()
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   var_hubo_cambios = True
End Sub

Private Sub txt_fecha_alta_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If Trim(txt_fecha_alta) <> "" Then
         If IsDate(txt_fecha_alta) Then
            frmcalendario.mes.Value = CDate(txt_fecha_alta)
         Else
            frmcalendario.mes.Value = Date
         End If
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_fecha_alta = var_fecha_general
   End If
End Sub

Private Sub txt_fecha_alta_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_agente_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_agente_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_almacen_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_almacen_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Then
      KeyCode = 0
   End If
   If KeyCode = 116 Then
      frm_lista.Left = 1500
      frm_lista.Top = 3405
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_almacenes order by vcha_alm_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ALM_ALMACEN_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      var_tipo_lista = 4
      lbl_lista = "Almacenes"
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
   If KeyCode = 117 Then
      Me.Enabled = False
      var_activa_forma_almacenes = Me.Name
      frmalmacenes.Show
   End If
End Sub

Private Sub txt_nombre_almacen_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 13
   Case Else
       KeyAscii = 0
   End Select
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_almacen_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_canal_venta_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_canal_venta_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_canal_venta_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Then
      KeyCode = 0
   End If
   If KeyCode = 116 Then
      frm_lista.Left = 1500
      frm_lista.Top = 2715
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_CANALESVENTAS order by vcha_can_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_can_canal_venta_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_can_nombre), "", rs!vcha_can_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Canales de Venta"
      var_tipo_lista = 5
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 7 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
   If KeyCode = 117 Then
      Me.Enabled = False
      var_activa_forma_canalesventas = Me.Name
      frmcanalesventas.Show
   End If
End Sub

Private Sub txt_nombre_canal_venta_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_empresa_GotFocus()
   Me.frm_contraseña.Visible = False
End Sub

Private Sub txt_nombre_empresa_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_tipo_agente_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_tipo_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Then
      KeyCode = 0
   End If
   If KeyCode = 116 Then
      frm_lista.Left = 1500
      frm_lista.Top = 2385
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_tipoagentes order by vcha_tag_descripcion", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_tag_tipoagente_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_TAG_DESCRIPCION), "", rs!VCHA_TAG_DESCRIPCION)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Tipo Agentes"
      var_tipo_lista = 2
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
   If KeyCode = 117 Then
      Me.Enabled = False
      frmtipoagentes.Show
      var_activa_forma_tipoagentes = Me.Name
   End If
End Sub

Private Sub txt_nombre_tipo_agente_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 13
   Case Else
       KeyAscii = 0
   End Select
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_tipo_agente_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_zona_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_oracle_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_ruta_archivos_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_ruta_archivos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub



Private Sub txt_telefono_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_telefono_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_zona_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_zona_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_zona_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frm_lista.Left = 1500
      frm_lista.Top = 2715
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_zonas order by vcha_zon_descripcion", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_zon_zona_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_zon_descripcion), "", rs!vcha_zon_descripcion)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Zonas"
      var_tipo_lista = 20
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 7 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
   If KeyCode = 117 Then
      Me.Enabled = False
      frmcanalesventas.Show
      var_activa_forma_canalesventas = Me.Name
   End If
End Sub

Private Sub txt_zona_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_zona_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(Me.txt_zona) <> "" Then
      rs.Open "SELECT * FROM tb_zonas where vcha_zon_zona_id = '" + Me.txt_zona + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_tipo_agente = IIf(IsNull(rs!vcha_zon_descripcion), "", rs!vcha_zon_descripcion)
      Else
         MsgBox "Clave de zona incorrecta", vbOKOnly, "ATENCION"
         Me.txt_zona = ""
         Me.txt_nombre_zona = ""
      End If
      rs.Close
   Else
      Me.txt_nombre_zona = ""
   End If
End Sub

Private Sub txtpassword_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If UCase(Me.txtusuario) = "FBI" Then
         If UCase(Me.txtpassword) = "FBI" Then
            var_agente_seleccionado = txt_clave_agente
            frmcomisiones.Caption = txt_nombre_agente
            frmcomisiones.Show
         Else
            MsgBox "Contraseña incorrecta", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Clave de usuario incorrecto", vbo
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_contraseña.Visible = False
   End If
End Sub

Private Sub txtusuario_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If UCase(Me.txtusuario) = "FBI" Then
         Me.txtpassword.SetFocus
      Else
         MsgBox "Clave dde usuario incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_contraseña.Visible = False
   End If
End Sub
