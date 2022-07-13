VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmreporte_movimientos_2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de movimientos"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   " Movimientos "
      Height          =   3855
      Left            =   75
      TabIndex        =   55
      Top             =   3360
      Width           =   5625
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1425
         Picture         =   "frmreporte_movimientos_2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   765
         Picture         =   "frmreporte_movimientos_2.frx":0216
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Marcar (Enter)"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1110
         Picture         =   "frmreporte_movimientos_2.frx":0460
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   105
         Picture         =   "frmreporte_movimientos_2.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   435
         Picture         =   "frmreporte_movimientos_2.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   210
         Width           =   330
      End
      Begin VB.Frame Frame7 
         Height          =   120
         Left            =   15
         TabIndex        =   56
         Top             =   525
         Width           =   5565
      End
      Begin MSComctlLib.ListView lv_movimientos 
         Height          =   3075
         Left            =   45
         TabIndex        =   14
         Top             =   675
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   5424
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
         NumItems        =   3
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
            Text            =   "marca"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmreporte_movimientos_2.frx":084A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frmreporte_movimientos_2.frx":094C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11175
      Picture         =   "frmreporte_movimientos_2.frx":0A4E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame cmb_filtro 
      Height          =   3045
      Left            =   5730
      TabIndex        =   23
      Top             =   1680
      Width           =   5790
      Begin VB.Frame Frame10 
         Height          =   120
         Left            =   30
         TabIndex        =   35
         Top             =   675
         Width           =   5730
      End
      Begin VB.CommandButton Command14 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   390
         Picture         =   "frmreporte_movimientos_2.frx":1088
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   405
         Width           =   330
      End
      Begin VB.CommandButton Command15 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   60
         Picture         =   "frmreporte_movimientos_2.frx":15BA
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Aceptar"
         Top             =   405
         Width           =   330
      End
      Begin VB.TextBox txt_clave_filtrar 
         Height          =   315
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   2475
         Width           =   1230
      End
      Begin VB.ComboBox cmb_filtrar 
         Height          =   315
         Left            =   2190
         TabIndex        =   31
         Top             =   2475
         Width           =   3480
      End
      Begin VB.ComboBox cmb_filtrar_por 
         Height          =   315
         ItemData        =   "frmreporte_movimientos_2.frx":1AEC
         Left            =   945
         List            =   "frmreporte_movimientos_2.frx":1B2C
         Sorted          =   -1  'True
         TabIndex        =   30
         Top             =   2115
         Width           =   4725
      End
      Begin VB.CommandButton cmd_y 
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   1935
         TabIndex        =   29
         Top             =   1395
         Width           =   1875
      End
      Begin VB.CommandButton cmd_o 
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   3810
         TabIndex        =   28
         Top             =   1395
         Width           =   1800
      End
      Begin VB.Frame Frame11 
         Height          =   120
         Left            =   15
         TabIndex        =   27
         Top             =   1260
         Width           =   5745
      End
      Begin VB.TextBox txt_filtro 
         Height          =   315
         Left            =   585
         TabIndex        =   26
         Top             =   945
         Width           =   5010
      End
      Begin VB.Frame Frame12 
         Height          =   120
         Left            =   15
         TabIndex        =   25
         Top             =   1905
         Width           =   5745
      End
      Begin VB.CommandButton cmd_agregar 
         Caption         =   "Agregar al filtro"
         Height          =   510
         Left            =   165
         TabIndex        =   24
         Top             =   1395
         Width           =   1770
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         Caption         =   "Filtrar Artículos "
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   30
         TabIndex        =   39
         Top             =   120
         Width           =   5730
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Left            =   105
         TabIndex        =   38
         Top             =   2535
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Filtrar por:"
         Height          =   195
         Left            =   90
         TabIndex        =   37
         Top             =   2175
         Width           =   690
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Filtro:"
         Height          =   195
         Left            =   135
         TabIndex        =   36
         Top             =   990
         Width           =   375
      End
   End
   Begin MSComCtl2.MonthView mon_mes2 
      Height          =   2370
      Left            =   8640
      TabIndex        =   41
      Top             =   4395
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   74514433
      CurrentDate     =   37761
   End
   Begin MSComCtl2.MonthView mon_mes1 
      Height          =   2370
      Left            =   5850
      TabIndex        =   42
      Top             =   4410
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   74514433
      CurrentDate     =   37761
   End
   Begin VB.Frame Frame1 
      Caption         =   " Almacenes "
      Height          =   2880
      Left            =   75
      TabIndex        =   57
      Top             =   405
      Width           =   5625
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         Picture         =   "frmreporte_movimientos_2.frx":1C41
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   795
         Picture         =   "frmreporte_movimientos_2.frx":1E57
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Marcar (Enter)"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1125
         Picture         =   "frmreporte_movimientos_2.frx":20A1
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   120
         Picture         =   "frmreporte_movimientos_2.frx":2173
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   450
         Picture         =   "frmreporte_movimientos_2.frx":2275
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   225
         Width           =   330
      End
      Begin VB.Frame Frame6 
         Height          =   120
         Left            =   30
         TabIndex        =   58
         Top             =   540
         Width           =   5565
      End
      Begin MSComctlLib.ListView lv_almacenes 
         Height          =   2025
         Left            =   45
         TabIndex        =   8
         Top             =   690
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   3572
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
         NumItems        =   3
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
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame5 
      Height          =   120
      Left            =   60
      TabIndex        =   40
      Top             =   270
      Width           =   11445
   End
   Begin VB.Frame Frame3 
      Caption         =   "  Artículos "
      Height          =   6075
      Left            =   5745
      TabIndex        =   46
      Top             =   405
      Width           =   5790
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1455
         Picture         =   "frmreporte_movimientos_2.frx":248B
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   225
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton Command7 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   795
         Picture         =   "frmreporte_movimientos_2.frx":26A1
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Marcar (Enter)"
         Top             =   225
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton Command8 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1125
         Picture         =   "frmreporte_movimientos_2.frx":28EB
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   225
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton Command9 
         Height          =   315
         Left            =   135
         Picture         =   "frmreporte_movimientos_2.frx":29BD
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   225
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   465
         Picture         =   "frmreporte_movimientos_2.frx":2ABF
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   225
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Frame Frame8 
         Height          =   120
         Left            =   15
         TabIndex        =   48
         Top             =   540
         Width           =   5745
      End
      Begin VB.CommandButton Command13 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1785
         Picture         =   "frmreporte_movimientos_2.frx":2CD5
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Filtrar Artículos"
         Top             =   225
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.OptionButton opt_seleccion 
         Caption         =   "Selección de artículos"
         Height          =   210
         Left            =   300
         TabIndex        =   15
         Top             =   240
         Width           =   1980
      End
      Begin VB.OptionButton opt_todos 
         Caption         =   "Todos los artículos"
         Height          =   210
         Left            =   3225
         TabIndex        =   16
         Top             =   255
         Width           =   1845
      End
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   2025
         TabIndex        =   17
         Top             =   780
         Width           =   1740
      End
      Begin MSComctlLib.ListView lv_articulos 
         Height          =   4785
         Left            =   75
         TabIndex        =   18
         Top             =   1215
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   8440
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "marca"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Busqueda de artículo:"
         Height          =   195
         Index           =   4
         Left            =   390
         TabIndex        =   54
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   720
      Left            =   5745
      TabIndex        =   43
      Top             =   6495
      Width           =   5775
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   3780
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   293
         Width           =   1305
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   293
         Width           =   1320
      End
      Begin VB.CommandButton Command11 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2355
         Picture         =   "frmreporte_movimientos_2.frx":2DCF
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Fecha Inicial"
         Top             =   300
         Width           =   330
      End
      Begin VB.CommandButton Command12 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5115
         Picture         =   "frmreporte_movimientos_2.frx":4041
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Fecha Final"
         Top             =   315
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   3465
         TabIndex        =   45
         Top             =   353
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Index           =   0
         Left            =   525
         TabIndex        =   44
         Top             =   353
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmreporte_movimientos_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim numero_items_ALMACENES As Integer
Dim numero_items_movimientos As Integer
Dim numero_items_articulos As Integer
Dim numero_control As Double
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim contador_filtro As Integer
Dim filtro_espanol As String
Dim filtro_ingles As String
Dim var_condicion As String
Dim var_condicion_2 As String
Dim var_clave As String


Private Sub cmb_sublineas_Click()
   txt_sublinea = Obtener_llave(cnn, rs, "TB_SUBLINEAS", "VCHA_SLI_NOMBRE", cmb_sublineas, 1, "T")
End Sub

Private Sub cmb_filtrar_Click()
   If Trim(cmb_filtrar_por) = "CATALOGO INICIO" Then
      txt_clave_filtrar = Obtener_llave(cnn, rs, "TB_CATALOGOS", "VCHA_CAT_NOMBRE", cmb_filtrar, 0, "T")
      var_clave = "VCHA_ART_CATALOGO_INICIO"
   End If
   If Trim(cmb_filtrar_por) = "CATALOGO VIGENTE" Then
      txt_clave_filtrar = Obtener_llave(cnn, rs, "TB_CATALOGOS", "VCHA_CAT_NOMBRE", cmb_filtrar, 0, "T")
      var_clave = "VCHA_ART_CATALOGO_VIGENTE"
   End If
   If Trim(cmb_filtrar_por) = "FAMILIA" Then
      txt_clave_filtrar = Obtener_llave(cnn, rs, "TB_DISEÑOS", "VCHA_DIS_NOMBRE", cmb_filtrar, 0, "T")
      var_clave = "VCHA_DIS_DISEÑO_ID"
   End If
   If Trim(cmb_filtrar_por) = "LINEA" Then
      txt_clave_filtrar = Obtener_llave(cnn, rs, "TB_LINEAS", "VCHA_LIN_NOMBRE", cmb_filtrar, 0, "T")
      var_clave = "VCHA_LIN_LINEA_ID"
   End If
   If Trim(cmb_filtrar_por) = "SUBLINEA" Then
      txt_clave_filtrar = Obtener_llave(cnn, rs, "TB_SUBLINEAS", "VCHA_SLI_NOMBRE", cmb_filtrar, 1, "T")
      var_clave = "VCHA_SLI_SUBLINEA_ID"
   End If
   If Trim(cmb_filtrar_por) = "PRODUCTO" Then
      txt_clave_filtrar = Obtener_llave(cnn, rs, "TB_PRODUCTOS", "VCHA_PRO_NOMBRE", cmb_filtrar, 0, "T")
      var_clave = "VCHA_PRO_PRODUCTO_ID"
   End If
   If Trim(cmb_filtrar_por) = "TIPO DE PRODUCTO" Then
      txt_clave_filtrar = Obtener_llave(cnn, rs, "TB_TIPOARTICULOS", "VCHA_TAR_NOMBRE", cmb_filtrar, 0, "T")
      var_clave = "VCHA_TAR_TIPO_ARTICULO_ID"
   End If
   If Trim(cmb_filtrar_por) = "CLASE" Then
      txt_clave_filtrar = Obtener_llave(cnn, rs, "TB_CLASEARTICULOS", "VCHA_CAR_NOMBRE", cmb_filtrar, 0, "T")
      var_clave = "VCHA_CAR_CLASE_ID"
   End If
   If Trim(cmb_filtrar_por) = "ESTAMPADO ANVERSO" Then
      txt_clave_filtrar = Obtener_llave(cnn, rs, "TB_ESTAMPADOS", "VCHA_EST_NOMBRE", cmb_filtrar, 0, "T")
      var_clave = "VCHA_ART_ESTAMPADO1"
   End If
   If Trim(cmb_filtrar_por) = "ESTAMPADO REVERSO" Then
      txt_clave_filtrar = Obtener_llave(cnn, rs, "TB_ESTAMPADOS", "VCHA_EST_NOMBRE", cmb_filtrar, 0, "T")
      var_clave = "VCHA_ART_ESTAMPADO2"
   End If
   If Trim(cmb_filtrar_por) = "TIPO ESTAMPADO ANVERSO" Then
      txt_clave_filtrar = Obtener_llave(cnn, rs, "TB_TIPOESTAMPADOS", "VCHA_TES_NOMBRE", cmb_filtrar, 0, "T")
      var_clave = "VCHA_ART_TIPO_ESTAMPADO1"
   End If
   If Trim(cmb_filtrar_por) = "TIPO ESTAMPADO REVERSO" Then
      txt_clave_filtrar = Obtener_llave(cnn, rs, "TB_TIPOESTAMPADOS", "VCHA_TES_NOMBRE", cmb_filtrar, 0, "T")
      var_clave = "VCHA_ART_TIPO_ESTAMPADO2"
   End If
   If Trim(cmb_filtrar_por) = "COLOR ANVERSO" Then
      txt_clave_filtrar = Obtener_llave(cnn, rs, "TB_COLORES", "VCHA_CLR_NOMBRE", cmb_filtrar, 0, "T")
      var_clave = "VCHA_ART_COLOR1"
   End If
   If Trim(cmb_filtrar_por) = "COLOR REVERSO" Then
      txt_clave_filtrar = Obtener_llave(cnn, rs, "TB_COLORES", "VCHA_CLR_NOMBRE", cmb_filtrar, 0, "T")
      var_clave = "VCHA_ART_COLOR2"
   End If
   If Trim(cmb_filtrar_por) = "TONO ANVERSO" Then
      txt_clave_filtrar = Obtener_llave(cnn, rs, "TB_TONOS", "VCHA_TON_NOMBRE", cmb_filtrar, 0, "T")
      var_clave = "VCHA_ART_TONO1"
   End If
   If Trim(cmb_filtrar_por) = "TONO REVERSO" Then
      txt_clave_filtrar = Obtener_llave(cnn, rs, "TB_TONOS", "VCHA_TON_NOMBRE", cmb_filtrar, 0, "T")
      var_clave = "VCHA_ART_TONO2"
   End If
   If Trim(cmb_filtrar_por) = "USO" Then
      txt_clave_filtrar = Obtener_llave(cnn, rs, "TB_USOS", "VCHA_USO_NOMBRE", cmb_filtrar, 0, "T")
      var_clave = "VCHA_USO_USO_ID"
   End If
   If Trim(cmb_filtrar_por) = "SUBTIPO USO" Then
      txt_clave_filtrar = Obtener_llave(cnn, rs, "TB_SUBTIPOSUSOS", "VCHA_SUS_NOMBRE", cmb_filtrar, 0, "T")
      var_clave = "VCHA_SUS_SUBTIPO_USO_ID"
   End If
   If Trim(cmb_filtrar_por) = "TALLA" Then
      txt_clave_filtrar = Obtener_llave(cnn, rs, "TB_TALLAS", "VCHA_TAL_NOMBRE", cmb_filtrar, 0, "T")
      var_clave = "VCHA_TAL_TALLA_ID"
   End If
   If Trim(cmb_filtrar_por) = "UNIDAD" Then
      txt_clave_filtrar = Obtener_llave(cnn, rs, "TB_UNIDADES", "VCHA_UNI_NOMBRE", cmb_filtrar, 0, "T")
      var_clave = "VCHA_UNI_UNIDAD_ID"
   End If
End Sub

Private Sub cmb_filtrar_por_Click()
   txt_clave_filtrar = ""
   cmb_filtrar.Clear
End Sub

Private Sub cmb_filtrar_por_LostFocus()
   If Trim(cmb_filtrar_por) <> "" Then
      If Trim(cmb_filtrar_por) = "CATALOGO INICIO" Then
         rs.Open "select * from tb_catalogos", cnn, adOpenDynamic, adLockBatchOptimistic
         Call RecsetToCombo(cmb_filtrar.hwnd, rs, 1)
         rs.Close
      End If
      If Trim(cmb_filtrar_por) = "CATALOGO VIGENTE" Then
         rs.Open "select * from tb_catalogos", cnn, adOpenDynamic, adLockBatchOptimistic
         Call RecsetToCombo(cmb_filtrar.hwnd, rs, 1)
         rs.Close
      End If
      If Trim(cmb_filtrar_por) = "FAMILIA" Then
         rs.Open "select * from tb_diseños", cnn, adOpenDynamic, adLockBatchOptimistic
         Call RecsetToCombo(cmb_filtrar.hwnd, rs, 1)
         rs.Close
      End If
      If Trim(cmb_filtrar_por) = "LINEA" Then
         rs.Open "select * from tb_lineas", cnn, adOpenDynamic, adLockBatchOptimistic
         Call RecsetToCombo(cmb_filtrar.hwnd, rs, 1)
         rs.Close
      End If
      If Trim(cmb_filtrar_por) = "SUBLINEA" Then
         rs.Open "select * from tb_sublineas", cnn, adOpenDynamic, adLockBatchOptimistic
         Call RecsetToCombo(cmb_filtrar.hwnd, rs, 2)
         rs.Close
      End If
      If Trim(cmb_filtrar_por) = "PRODUCTO" Then
         rs.Open "select * from tb_productos", cnn, adOpenDynamic, adLockBatchOptimistic
         Call RecsetToCombo(cmb_filtrar.hwnd, rs, 1)
         rs.Close
      End If
      If Trim(cmb_filtrar_por) = "TIPO DE PRODUCTO" Then
         rs.Open "select * from tb_tipoarticulos", cnn, adOpenDynamic, adLockBatchOptimistic
         Call RecsetToCombo(cmb_filtrar.hwnd, rs, 1)
         rs.Close
      End If
      If Trim(cmb_filtrar_por) = "CLASE" Then
         rs.Open "select * from tb_clasearticulos", cnn, adOpenDynamic, adLockBatchOptimistic
         Call RecsetToCombo(cmb_filtrar.hwnd, rs, 1)
         rs.Close
      End If
      If Trim(cmb_filtrar_por) = "ESTAMPADO ANVERSO" Then
         rs.Open "select * from tb_estampados", cnn, adOpenDynamic, adLockBatchOptimistic
         Call RecsetToCombo(cmb_filtrar.hwnd, rs, 1)
         rs.Close
      End If
      If Trim(cmb_filtrar_por) = "ESTAMPADO REVERSO" Then
         rs.Open "select * from tb_estampados", cnn, adOpenDynamic, adLockBatchOptimistic
         Call RecsetToCombo(cmb_filtrar.hwnd, rs, 1)
         rs.Close
      End If
      If Trim(cmb_filtrar_por) = "TIPO ESTAMPADO ANVERSO" Then
         rs.Open "select * from tb_tipoestampados", cnn, adOpenDynamic, adLockBatchOptimistic
         Call RecsetToCombo(cmb_filtrar.hwnd, rs, 1)
         rs.Close
      End If
      If Trim(cmb_filtrar_por) = "TIPO ESTAMPADO REVERSO" Then
         rs.Open "select * from tb_tipoestampados", cnn, adOpenDynamic, adLockBatchOptimistic
         Call RecsetToCombo(cmb_filtrar.hwnd, rs, 1)
         rs.Close
      End If
      If Trim(cmb_filtrar_por) = "COLOR ANVERSO" Then
         rs.Open "select * from tb_colores", cnn, adOpenDynamic, adLockBatchOptimistic
         Call RecsetToCombo(cmb_filtrar.hwnd, rs, 1)
         rs.Close
      End If
      If Trim(cmb_filtrar_por) = "COLOR REVERSO" Then
         rs.Open "select * from tb_colores", cnn, adOpenDynamic, adLockBatchOptimistic
         Call RecsetToCombo(cmb_filtrar.hwnd, rs, 1)
         rs.Close
      End If
      If Trim(cmb_filtrar_por) = "TONO ANVERSO" Then
         rs.Open "select * from tb_tonos", cnn, adOpenDynamic, adLockBatchOptimistic
         Call RecsetToCombo(cmb_filtrar.hwnd, rs, 1)
         rs.Close
      End If
      If Trim(cmb_filtrar_por) = "TONO REVERSO" Then
         rs.Open "select * from tb_tonos", cnn, adOpenDynamic, adLockBatchOptimistic
         Call RecsetToCombo(cmb_filtrar.hwnd, rs, 1)
         rs.Close
      End If
      If Trim(cmb_filtrar_por) = "USO" Then
         rs.Open "select * from tb_usos", cnn, adOpenDynamic, adLockBatchOptimistic
         Call RecsetToCombo(cmb_filtrar.hwnd, rs, 1)
         rs.Close
      End If
      If Trim(cmb_filtrar_por) = "SUBTIPO USO" Then
         rs.Open "select * from tb_subtiposusos", cnn, adOpenDynamic, adLockBatchOptimistic
         Call RecsetToCombo(cmb_filtrar.hwnd, rs, 2)
         rs.Close
      End If
      If Trim(cmb_filtrar_por) = "TALLA" Then
         rs.Open "select * from tb_tallas", cnn, adOpenDynamic, adLockBatchOptimistic
         Call RecsetToCombo(cmb_filtrar.hwnd, rs, 1)
         rs.Close
      End If
      If Trim(cmb_filtrar_por) = "UNIDAD" Then
         rs.Open "select * from tb_unidades", cnn, adOpenDynamic, adLockBatchOptimistic
         Call RecsetToCombo(cmb_filtrar.hwnd, rs, 1)
         rs.Close
      End If
   End If
End Sub

Private Sub cmd_agregar_Click()
   If contador_filtro = 0 Then
      If Trim(txt_clave_filtrar) <> "" Then
         filtro_espanol = Trim(cmb_filtrar_por) + " = " + cmb_filtrar
         filtro_ingles = var_clave + " = '" + txt_clave_filtrar + "'"
         txt_filtro = filtro_espanol
         contador_filtro = contador_filtro + 1
         cmd_y.Enabled = True
         cmd_o.Enabled = True
         txt_clave_filtrar.Enabled = False
         cmb_filtrar.Enabled = False
         cmb_filtrar_por.Enabled = False
      Else
      End If
   Else
       txt_filtro = txt_filtro + " " + var_condicion + " " + Trim(cmb_filtrar_por) + " = " + cmb_filtrar
       filtro_ingles = filtro_ingles + " " + var_condicion_2 + " " + var_clave + " = '" + txt_clave_filtrar + "'"
   End If
End Sub

Private Sub cmd_imprimir_Click()

   Dim var_fecha_fin As Date
   var_fecha_fin = txt_fin
   var_fecha_fin = var_fecha_fin
   var_contador_almacenes = 0
   var_contador_movimientos = 0
   For var_j = 1 To lv_almacenes.ListItems.Count
       lv_almacenes.ListItems.Item(var_j).Selected = True
       If lv_almacenes.selectedItem.SubItems(2) = "*" Then
          var_contador_almacenes = var_contador_almacenes + 1
       End If
   Next var_j
   
   For var_j = 1 To lv_movimientos.ListItems.Count
       lv_movimientos.ListItems.Item(var_j).Selected = True
       If lv_movimientos.selectedItem.SubItems(2) = "*" Then
          var_contador_movimientos = var_contador_movimientos + 1
       End If
   Next var_j
   If var_contador_almacenes > 0 Then
      If var_contador_movimientos > 0 Then
         If lv_articulos.ListItems.Count > 0 Or Me.opt_todos = True Then
            cnn.CommandTimeout = 3600
            cnn.BeginTrans
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open "SELECT MAX(inte_rmo_numero) FROM tb_temp_rep_mov_reporte WHERE vcha_rmo_maquina = '" + fun_NombrePc + "' and vcha_rmo_usuario = '" + var_clave_usuario_global + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
            Else
               var_consecutivo = 1
            End If
            rs.Close
            rs.Open "INSERT INTO  tb_temp_rep_mov_reporte (INTE_RMO_NUMERO, VCHA_RMO_MAQUINA, VCHA_RMO_USUARIO) VALUES (" + CStr(var_consecutivo) + ",'" + fun_NombrePc + "','" + var_clave_usuario_global + "')", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            For var_j = 1 To lv_almacenes.ListItems.Count
                lv_almacenes.ListItems.Item(var_j).Selected = True
                If lv_almacenes.selectedItem.SubItems(2) = "*" Then
                   rs.Open "insert into tb_temp_rep_mov_almacenes values (" + Str(var_consecutivo) + ", '" + fun_NombrePc + "', '" + var_clave_usuario_global + "', '" + Me.lv_almacenes.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
                End If
            Next var_j
            For var_j = 1 To lv_movimientos.ListItems.Count
                lv_movimientos.ListItems.Item(var_j).Selected = True
                If lv_movimientos.selectedItem.SubItems(2) = "*" Then
                   rs.Open "insert into tb_temp_rep_mov_movimientos (inte_rmo_numero, vcha_rmo_maquina, vcha_rmo_usuario, vcha_mov_movimiento_id) values (" + Str(var_consecutivo) + ", '" + fun_NombrePc + "', '" + var_clave_usuario_global + "', '" + lv_movimientos.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
                End If
            Next var_j
            If Me.opt_todos.Value = True Then
               rs.Open "insert into tb_temp_rep_mov_articulos select  " + Str(var_consecutivo) + " as inte_rmo_numero, '" + fun_NombrePc + "' as vcha_rmo_maquina, '" + var_clave_usuario_global + "' as vcha_rmo_usuario, vcha_art_articulo_id from tb_articulos", cnn, adOpenDynamic, adLockOptimistic
            Else
               For var_j = 1 To lv_articulos.ListItems.Count
                   lv_articulos.ListItems.Item(var_j).Selected = True
                   rs.Open "insert into tb_temp_rep_mov_articulos VALUES (" + Str(var_consecutivo) + ", '" + fun_NombrePc + "', '" + var_clave_usuario_global + "', '" + lv_articulos.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
               Next var_j
            End If
            rs.Open "execute reporte_movimientos " + Str(var_consecutivo) + ", '" + fun_NombrePc + "', '" + var_clave_usuario_global + "', '" + txt_inicio + "', '" + CStr(var_fecha_fin) + "'", cnn, adOpenDynamic, adLockOptimistic
            var_si = 6
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_movimientos_2.rpt")
               reporte.RecordSelectionFormula = "{VW_REPORTE_MOVIMIENTOS_2.VCHA_RMO_MAQUINA} = '" + fun_NombrePc + "' AND {VW_REPORTE_MOVIMIENTOS_2.VCHA_RMO_USUARIO} = '" + var_clave_usuario_global + "' AND {VW_REPORTE_MOVIMIENTOS_2.INTE_RMO_NUMERO} = " + Str(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\Reporte_movimientos_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
            rs.Open "delete tb_temp_rep_mov_reporte where inte_rmo_numero = " + Str(var_consecutivo) + " and vcha_rmo_maquina = '" + fun_NombrePc + "' and vcha_rmo_usuario = '" + var_clave_usuario_global + "'", cnn, adOpenDynamic, adLockOptimistic
            rs.Open "delete tb_temp_rep_mov_almacenes where inte_rmo_numero = " + Str(var_consecutivo) + " and vcha_rmo_maquina = '" + fun_NombrePc + "' and vcha_rmo_usuario = '" + var_clave_usuario_global + "'", cnn, adOpenDynamic, adLockOptimistic
            rs.Open "delete tb_temp_rep_mov_movimientos where inte_rmo_numero = " + Str(var_consecutivo) + " and vcha_rmo_maquina = '" + fun_NombrePc + "' and vcha_rmo_usuario = '" + var_clave_usuario_global + "'", cnn, adOpenDynamic, adLockOptimistic
            rs.Open "delete tb_temp_rep_mov_articulos where inte_rmo_numero = " + Str(var_consecutivo) + " and vcha_rmo_maquina = '" + fun_NombrePc + "' and vcha_rmo_usuario = '" + var_clave_usuario_global + "'", cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "No se a seleccionado un movimiento", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "No se a seleccionado un almacén", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a seleccionado un artículo", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_invertir_Click()
Dim n As Integer
Dim i As Integer
Dim j As Integer
   n = lv_almacenes.ListItems.Count
   For i = 1 To n
      lv_almacenes.ListItems.Item(i).Selected = True
      If lv_almacenes.selectedItem.SubItems(2) = "*" Then
         lv_almacenes.selectedItem.SubItems(2) = ""
         lv_almacenes.ListItems.Item(i).Bold = False
         lv_almacenes.ListItems.Item(i).ForeColor = &H80000012
         lv_almacenes.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_almacenes.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_almacenes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_almacenes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_almacenes.selectedItem.SubItems(2) = "*"
         lv_almacenes.ListItems.Item(i).Bold = True
         lv_almacenes.ListItems.Item(i).ForeColor = &HFF0000
         lv_almacenes.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_almacenes.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_almacenes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_almacenes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub cmd_marcar_Click()
Dim i As Integer
   i = lv_almacenes.selectedItem.Index
   If lv_almacenes.selectedItem.SubItems(2) = "*" Then
      lv_almacenes.selectedItem.SubItems(2) = ""
      lv_almacenes.ListItems.Item(i).Bold = False
      lv_almacenes.ListItems.Item(i).ForeColor = &H80000012
      lv_almacenes.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_almacenes.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_almacenes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_almacenes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_almacenes.Refresh
   Else
      lv_almacenes.selectedItem.SubItems(2) = "*"
      lv_almacenes.ListItems.Item(i).Bold = True
      lv_almacenes.ListItems.Item(i).ForeColor = &HFF0000
      lv_almacenes.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_almacenes.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_almacenes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_almacenes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_almacenes.Refresh
   End If
End Sub

Private Sub cmd_ninguno_Click()
Dim n As Integer
Dim i As Integer
Dim j As Integer
   n = lv_almacenes.ListItems.Count
   For i = 1 To n
      lv_almacenes.ListItems.Item(i).Selected = True
      lv_almacenes.selectedItem.SubItems(2) = ""
      lv_almacenes.ListItems.Item(i).Bold = False
      lv_almacenes.ListItems.Item(i).ForeColor = &H80000012
      lv_almacenes.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_almacenes.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_almacenes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_almacenes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   lv_almacenes.Refresh
End Sub

Private Sub cmd_nuevo_Click()
Dim n As Integer
Dim i As Integer
Dim j As Integer
   numero_control = 0
   txt_fin = Date
   txt_inicio = Date
   n = lv_almacenes.ListItems.Count
   For i = 1 To n
      lv_almacenes.ListItems.Item(i).Selected = True
      lv_almacenes.selectedItem.SubItems(2) = ""
      lv_almacenes.ListItems.Item(i).Bold = False
      lv_almacenes.ListItems.Item(i).ForeColor = &H80000012
      lv_almacenes.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_almacenes.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_almacenes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_almacenes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   n = lv_movimientos.ListItems.Count
   For i = 1 To n
      lv_movimientos.ListItems.Item(i).Selected = True
      lv_movimientos.selectedItem.SubItems(2) = ""
      lv_movimientos.ListItems.Item(i).Bold = False
      lv_movimientos.ListItems.Item(i).ForeColor = &H80000012
      lv_movimientos.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_movimientos.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_movimientos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_movimientos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   n = lv_articulos.ListItems.Count
   For i = 1 To n
      lv_articulos.ListItems.Item(i).Selected = True
      lv_articulos.selectedItem.SubItems(2) = ""
      lv_articulos.ListItems.Item(i).Bold = False
      lv_articulos.ListItems.Item(i).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   lv_articulos.Refresh
   lv_movimientos.Refresh
   lv_almacenes.Refresh
End Sub

Private Sub cmd_o_Click()
   var_condicion = "O"
   var_condicion_2 = "OR"
   txt_clave_filtrar.Enabled = True
   cmb_filtrar.Enabled = True
   cmb_filtrar_por.Enabled = True
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmd_seleccion_Click()
Dim n As Integer
Dim i As Integer
Dim j As Integer
Dim var_rellena As Boolean
Dim var_encontro As Boolean
   n = lv_almacenes.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_almacenes.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_almacenes.selectedItem.SubItems(2) = "" And var_rellena = True Then
         lv_almacenes.selectedItem.SubItems(2) = "*"
         lv_almacenes.ListItems.Item(i).Bold = True
         lv_almacenes.ListItems.Item(i).ForeColor = &HFF0000
         lv_almacenes.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_almacenes.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_almacenes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_almacenes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_almacenes.selectedItem.SubItems(2) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_almacenes.selectedItem.SubItems(2) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub cmd_todos_Click()
Dim n As Integer
Dim i As Integer
Dim j As Integer
   n = lv_almacenes.ListItems.Count
   For i = 1 To n
      lv_almacenes.ListItems.Item(i).Selected = True
      lv_almacenes.selectedItem.SubItems(2) = "*"
      lv_almacenes.ListItems.Item(i).Bold = True
      lv_almacenes.ListItems.Item(i).ForeColor = &HFF0000
      lv_almacenes.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_almacenes.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_almacenes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_almacenes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
   Next i
   lv_almacenes.Refresh
End Sub

Private Sub cmd_y_Click()
   var_condicion = "Y"
   var_condicion_2 = "AND"
   txt_clave_filtrar.Enabled = True
   cmb_filtrar.Enabled = True
   cmb_filtrar_por.Enabled = True
End Sub

Private Sub Command1_Click()
Dim n As Integer
Dim i As Integer
Dim j As Integer
Dim var_rellena As Boolean
Dim var_encontro As Boolean
   n = lv_movimientos.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_movimientos.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_movimientos.selectedItem.SubItems(2) = "" And var_rellena = True Then
         lv_movimientos.selectedItem.SubItems(2) = "*"
         lv_movimientos.ListItems.Item(i).Bold = True
         lv_movimientos.ListItems.Item(i).ForeColor = &HFF0000
         lv_movimientos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_movimientos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_movimientos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_movimientos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_movimientos.selectedItem.SubItems(2) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_movimientos.selectedItem.SubItems(2) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub Command10_Click()
Dim n As Integer
Dim i As Integer
Dim j As Integer
   If numero_control = 0 Then
      Dim cmd As New Command
      Dim var_numero_tabla As Double
      Set cmd.ActiveConnection = cnn
      cmd.CommandType = adCmdStoredProc
      cmd.CommandText = "NUMERO_CONTROL"
      cmd("@maquina") = fun_NombrePc
      cmd("@usuario") = var_clave_usuario_global
      cmd("@numero") = numero_control
      cmd.execute
      numero_control = cmd("@numero")
      Set cmd = Nothing
   End If
   n = lv_articulos.ListItems.Count
   For i = 1 To n
      lv_articulos.ListItems.Item(i).Selected = True
      lv_articulos.selectedItem.SubItems(2) = "*"
      lv_articulos.ListItems.Item(i).Bold = True
      lv_articulos.ListItems.Item(i).ForeColor = &HFF0000
      lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
   Next i
   lv_articulos.Refresh
   rs.Open "delete from tb_temp_rep_mov_articulos where inte_rmo_numero = " + Str(numero_control) + " and vcha_rmo_maquina = '" + fun_NombrePc + "' and vcha_rmo_usuario = '" + var_clave_usuario_global + "'", cnn, adOpenDynamic, adLockOptimistic
   rs.Open "insert into tb_temp_rep_mov_articulos select  " + Str(numero_control) + " as inte_rmo_numero, '" + fun_NombrePc + "' as vcha_rmo_maquina, '" + var_clave_usuario_global + "' as vcha_rmo_usuario, vcha_art_articulo_id from tb_articulos", cnn, adOpenDynamic, adLockOptimistic
End Sub

Private Sub Command11_Click()
   mon_mes1.Value = txt_fin
   mon_mes1.Visible = True
   mon_mes1.SetFocus
End Sub

Private Sub Command12_Click()
   mon_mes2.Value = txt_fin
   mon_mes2.Visible = True
   mon_mes2.SetFocus
End Sub

Private Sub Command13_Click()
   contador_filtro = 0
   filtro_espanol = ""
   filtro_ingles = ""
   cmb_filtro.Visible = True
   txt_filtro = ""
   cmd_y.Enabled = False
   cmd_o.Enabled = False
   txt_clave_filtrar = ""
   txt_clave_filtrar.Enabled = True
   cmb_filtrar.Enabled = True
   cmb_filtrar.Clear
   cmb_filtrar_por.Enabled = True
   cmb_filtrar_por = ""
End Sub

Private Sub Command14_Click()
   cmb_filtro.Visible = False
End Sub

Private Sub Command15_Click()
   If Trim(filtro_ingles) <> "" Then
      rs.Open "delete from tb_temp_rep_mov_articulos where inte_rmo_numero = " + Str(numero_control) + " and vcha_rmo_maquina = '" + fun_NombrePc + "' and vcha_rmo_usuario = '" + var_clave_usuario_global + "'", cnn, adOpenDynamic, adLockOptimistic
      rs.Open "select * from tb_articulos where " + filtro_ingles + " order by vcha_art_nombre_español", cnn, adOpenDynamic, adLockOptimistic
      numero_items_articulos = 0
      lv_articulos.ListItems.Clear
      While Not rs.EOF
         Set list_item = lv_articulos.ListItems.Add(, , rs!vcha_Art_articulo_id)
         list_item.SubItems(1) = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
         list_item.SubItems(2) = ""
         list_item.Bold = False
         list_item.ForeColor = &H80000012
         rs.MoveNext:
         numero_items_articulos = numero_items_articulos + 1
       Wend
      rs.Close
      If numero_items_articulos > 12 Then
         lv_articulos.ColumnHeaders(2).Width = 4000
      Else
         lv_articulos.ColumnHeaders(2).Width = 4200
      End If
      rs.Open "delete from tb_temp_rep_mov_articulos where inte_rmo_numero = " + Str(numero_control) + " and vcha_rmo_maquina = '" + fun_NombrePc + "' and vcha_rmo_usuario = '" + var_clave_usuario_global + "'", cnn, adOpenDynamic, adLockOptimistic
      cmb_filtro.Visible = False
   End If
End Sub

Private Sub Command2_Click()
Dim i As Integer
   i = lv_movimientos.selectedItem.Index
   If lv_movimientos.selectedItem.SubItems(2) = "*" Then
      lv_movimientos.selectedItem.SubItems(2) = ""
      lv_movimientos.ListItems.Item(i).Bold = False
      lv_movimientos.ListItems.Item(i).ForeColor = &H80000012
      lv_movimientos.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_movimientos.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_movimientos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_movimientos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_movimientos.Refresh
   Else
      lv_movimientos.selectedItem.SubItems(2) = "*"
      lv_movimientos.ListItems.Item(i).Bold = True
      lv_movimientos.ListItems.Item(i).ForeColor = &HFF0000
      lv_movimientos.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_movimientos.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_movimientos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_movimientos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_movimientos.Refresh
   End If
End Sub

Private Sub Command3_Click()
Dim n As Integer
Dim i As Integer
Dim j As Integer
   n = lv_movimientos.ListItems.Count
   For i = 1 To n
      lv_movimientos.ListItems.Item(i).Selected = True
      If lv_movimientos.selectedItem.SubItems(2) = "*" Then
         lv_movimientos.selectedItem.SubItems(2) = ""
         lv_movimientos.ListItems.Item(i).Bold = False
         lv_movimientos.ListItems.Item(i).ForeColor = &H80000012
         lv_movimientos.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_movimientos.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_movimientos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_movimientos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_movimientos.selectedItem.SubItems(2) = "*"
         lv_movimientos.ListItems.Item(i).Bold = True
         lv_movimientos.ListItems.Item(i).ForeColor = &HFF0000
         lv_movimientos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_movimientos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_movimientos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_movimientos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      End If
   Next i

End Sub

Private Sub Command4_Click()
Dim n As Integer
Dim i As Integer
Dim j As Integer
   n = lv_movimientos.ListItems.Count
   For i = 1 To n
      lv_movimientos.ListItems.Item(i).Selected = True
      lv_movimientos.selectedItem.SubItems(2) = ""
      lv_movimientos.ListItems.Item(i).Bold = False
      lv_movimientos.ListItems.Item(i).ForeColor = &H80000012
      lv_movimientos.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_movimientos.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_movimientos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_movimientos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   lv_movimientos.Refresh
End Sub

Private Sub Command5_Click()
Dim n As Integer
Dim i As Integer
Dim j As Integer
   n = lv_movimientos.ListItems.Count
   For i = 1 To n
      lv_movimientos.ListItems.Item(i).Selected = True
      lv_movimientos.selectedItem.SubItems(2) = "*"
      lv_movimientos.ListItems.Item(i).Bold = True
      lv_movimientos.ListItems.Item(i).ForeColor = &HFF0000
      lv_movimientos.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_movimientos.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_movimientos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_movimientos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
   Next i
   lv_movimientos.Refresh
End Sub

Private Sub Command6_Click()
Dim n As Integer
Dim i As Integer
Dim j As Integer
Dim var_rellena As Boolean
Dim var_encontro As Boolean
   If numero_control = 0 Then
      Dim cmd As New Command
      Dim var_numero_tabla As Double
      Set cmd.ActiveConnection = cnn
      cmd.CommandType = adCmdStoredProc
      cmd.CommandText = "NUMERO_CONTROL"
      cmd("@maquina") = fun_NombrePc
      cmd("@usuario") = var_clave_usuario_global
      cmd("@numero") = numero_control
      cmd.execute
      numero_control = cmd("@numero")
      Set cmd = Nothing
   End If
   n = lv_articulos.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_articulos.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_articulos.selectedItem.SubItems(2) = "" And var_rellena = True Then
         lv_articulos.selectedItem.SubItems(2) = "*"
         lv_articulos.ListItems.Item(i).Bold = True
         lv_articulos.ListItems.Item(i).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_articulos.selectedItem.SubItems(2) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_articulos.selectedItem.SubItems(2) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i

End Sub

Private Sub Command7_Click()
Dim i As Integer
   If numero_control = 0 Then
      Dim cmd As New Command
      Dim var_numero_tabla As Double
      Set cmd.ActiveConnection = cnn
      cmd.CommandType = adCmdStoredProc
      cmd.CommandText = "NUMERO_CONTROL"
      cmd("@maquina") = fun_NombrePc
      cmd("@usuario") = var_clave_usuario_global
      cmd("@numero") = numero_control
      cmd.execute
      numero_control = cmd("@numero")
      Set cmd = Nothing
   End If
   i = lv_articulos.selectedItem.Index
   If lv_articulos.selectedItem.SubItems(2) = "*" Then
      lv_articulos.selectedItem.SubItems(2) = ""
      lv_articulos.ListItems.Item(i).Bold = False
      lv_articulos.ListItems.Item(i).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      rs.Open "delete from tb_temp_rep_mov_articulos where inte_rmo_numero = " + Str(numero_control) + " and vcha_rmo_maquina = '" + fun_NombrePc + "' and vcha_rmo_usuario = '" + var_clave_usuario_global + "' and vcha_art_articulo_id = '" + lv_articulos.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
      lv_articulos.Refresh
   Else
      lv_articulos.selectedItem.SubItems(2) = "*"
      lv_articulos.ListItems.Item(i).Bold = True
      lv_articulos.ListItems.Item(i).ForeColor = &HFF0000
      lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      rs.Open "insert into tb_temp_rep_mov_articulos (inte_rmo_numero, vcha_rmo_maquina, vcha_rmo_usuario, vcha_art_articulo_id) values (" + Str(numero_control) + ", '" + fun_NombrePc + "', '" + var_clave_usuario_global + "', '" + lv_articulos.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
      lv_articulos.Refresh
   End If

End Sub

Private Sub Command8_Click()
Dim n As Integer
Dim i As Integer
Dim j As Integer
   If numero_control = 0 Then
      Dim cmd As New Command
      Dim var_numero_tabla As Double
      Set cmd.ActiveConnection = cnn
      cmd.CommandType = adCmdStoredProc
      cmd.CommandText = "NUMERO_CONTROL"
      cmd("@maquina") = fun_NombrePc
      cmd("@usuario") = var_clave_usuario_global
      cmd("@numero") = numero_control
      cmd.execute
      numero_control = cmd("@numero")
      Set cmd = Nothing
   End If
   n = lv_articulos.ListItems.Count
   For i = 1 To n
      lv_articulos.ListItems.Item(i).Selected = True
      If lv_articulos.selectedItem.SubItems(2) = "*" Then
         lv_articulos.selectedItem.SubItems(2) = ""
         lv_articulos.ListItems.Item(i).Bold = False
         lv_articulos.ListItems.Item(i).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_articulos.selectedItem.SubItems(2) = "*"
         lv_articulos.ListItems.Item(i).Bold = True
         lv_articulos.ListItems.Item(i).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      End If
   Next i


End Sub

Private Sub Command9_Click()
Dim n As Integer
Dim i As Integer
Dim j As Integer
   If numero_control = 0 Then
      Dim cmd As New Command
      Dim var_numero_tabla As Double
      Set cmd.ActiveConnection = cnn
      cmd.CommandType = adCmdStoredProc
      cmd.CommandText = "NUMERO_CONTROL"
      cmd("@maquina") = fun_NombrePc
      cmd("@usuario") = var_clave_usuario_global
      cmd("@numero") = numero_control
      cmd.execute
      numero_control = cmd("@numero")
      Set cmd = Nothing
   End If
   n = lv_articulos.ListItems.Count
   For i = 1 To n
      lv_articulos.ListItems.Item(i).Selected = True
      lv_articulos.selectedItem.SubItems(2) = ""
      lv_articulos.ListItems.Item(i).Bold = False
      lv_articulos.ListItems.Item(i).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   rs.Open "delete from tb_temp_rep_mov_articulos where inte_rmo_numero = " + Str(numero_control) + " and vcha_rmo_maquina = '" + fun_NombrePc + "' and vcha_rmo_usuario = '" + var_clave_usuario_global + "'", cnn, adOpenDynamic, adLockOptimistic
   lv_articulos.Refresh


End Sub

Private Sub Form_Load()
   Me.opt_seleccion.Value = True
   var_cadena_seguridad = ""
   cmb_filtro.Visible = False
   numero_control = 0
   mon_mes1.Visible = False
   mon_mes2.Visible = False
   txt_inicio = Date
   txt_fin = Date
   Top = 0
   Left = 0
   Dim list_item As ListItem
   rs.Open "select DISTINCT A.VCHA_ALM_ALMACEN_ID, A.VCHA_ALM_NOMBRE from tb_almacenes A, TB_ENCABEZADO_MOVIMIENTOS B WHERE A.VCHA_aLM_ALMACEN_ID = B.VCHA_EMO_ALMACEN_ORIGEN OR A.VCHA_ALM_ALMACEN_ID = B.VCHA_EMO_ALMACEN_DESTINO AND A.vcha_emp_empresa_id = '" + var_empresa + "' order by vcha_alm_nombre", cnn, adOpenDynamic, adLockOptimistic
   numero_items_ALMACENES = 0
   While Not rs.EOF
      Set list_item = lv_almacenes.ListItems.Add(, , rs!VCHA_ALM_ALMACEN_ID)
      list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
      list_item.SubItems(2) = ""
      rs.MoveNext:
      numero_items_ALMACENES = numero_items_ALMACENES + 1
   Wend
   rs.Close
   If numero_items_ALMACENES > 8 Then
      lv_almacenes.ColumnHeaders(2).Width = 4200.71
   Else
      lv_almacenes.ColumnHeaders(2).Width = 4499.71
   End If
   rs.Open "select distinct a.vcha_mov_movimiento_id, a.vcha_mov_nombre from tb_movimientos a, tb_encabezado_movimientos b where a.vcha_mov_movimiento_id = b.vcha_mov_movimiento_id order by vcha_mov_nombre", cnn, adOpenDynamic, adLockOptimistic
   numero_items_movimientos = 0
   While Not rs.EOF
      Set list_item = lv_movimientos.ListItems.Add(, , rs!VCHA_MOV_MOVIMIENTO_ID)
      list_item.SubItems(1) = IIf(IsNull(rs!vcha_mov_nombre), "", rs!vcha_mov_nombre)
      list_item.SubItems(2) = ""
      rs.MoveNext:
      numero_items_movimientos = numero_items_movimientos + 1
    Wend
   rs.Close
   If numero_items_movimientos > 12 Then
      lv_movimientos.ColumnHeaders(2).Width = 4200.71
   Else
      lv_movimientos.ColumnHeaders(2).Width = 4499.71
   End If
   
  ' rs.Open "select * from tb_articulos order by vcha_art_nombre_español", cnn, adOpenDynamic, adLockOptimistic
  ' numero_items_articulos = 0
  ' While Not rs.EOF
  '    Set list_item = lv_articulos.ListItems.Add(, , rs!vcha_art_articulo_id)
  '    list_item.SubItems(1) = IIf(IsNull(rs!vcha_art_nombre_Español), "", rs!vcha_art_nombre_Español)
  '    list_item.SubItems(2) = ""
  '    rs.MoveNext:
  '    numero_items_articulos = numero_items_articulos + 1
  '  Wend
  ' rs.Close
  ' If numero_items_articulos > 12 Then
  '    lv_articulos.ColumnHeaders(2).Width = 4000
  ' Else
  '    lv_articulos.ColumnHeaders(2).Width = 4200
  ' End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   rs.Open "delete from tb_temp_rep_mov_almacenes where inte_rmo_numero = " + Str(numero_control) + " and vcha_rmo_maquina = '" + fun_NombrePc + "' and vcha_rmo_usuario = '" + var_clave_usuario_global + "'", cnn, adOpenDynamic, adLockOptimistic
   rs.Open "delete from tb_temp_rep_mov_movimientos where inte_rmo_numero = " + Str(numero_control) + " and vcha_rmo_maquina = '" + fun_NombrePc + "' and vcha_rmo_usuario = '" + var_clave_usuario_global + "'", cnn, adOpenDynamic, adLockOptimistic
   rs.Open "delete from tb_temp_rep_mov_articulos where inte_rmo_numero = " + Str(numero_control) + " and vcha_rmo_maquina = '" + fun_NombrePc + "' and vcha_rmo_usuario = '" + var_clave_usuario_global + "'", cnn, adOpenDynamic, adLockOptimistic
   Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub Frame9_DragDrop(Source As Control, x As Single, Y As Single)

End Sub

Private Sub lv_almacenes_KeyPress(KeyAscii As Integer)
   Dim i As Integer
   If KeyAscii = 13 Then
      If numero_control = 0 Then
         Dim cmd As New Command
         Dim var_numero_tabla As Double
         Set cmd.ActiveConnection = cnn
         cmd.CommandType = adCmdStoredProc
         cmd.CommandText = "NUMERO_CONTROL"
         cmd("@maquina") = fun_NombrePc
         cmd("@usuario") = var_clave_usuario_global
         cmd("@numero") = numero_control
         cmd.execute
         numero_control = cmd("@numero")
         Set cmd = Nothing
      End If
      i = lv_almacenes.selectedItem.Index
      If lv_almacenes.selectedItem.SubItems(2) = "*" Then
         lv_almacenes.selectedItem.SubItems(2) = ""
         lv_almacenes.ListItems.Item(i).Bold = False
         lv_almacenes.ListItems.Item(i).ForeColor = &H80000012
         lv_almacenes.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_almacenes.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_almacenes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_almacenes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         rs.Open "delete from tb_temp_rep_mov_almacenes where inte_rmo_numero = " + Str(numero_control) + " and vcha_rmo_maquina = '" + fun_NombrePc + "' and vcha_rmo_usuario = '" + var_clave_usuario_global + "' and vcha_alm_almacen_id = '" + lv_almacenes.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
         lv_almacenes.Refresh
      Else
         lv_almacenes.selectedItem.SubItems(2) = "*"
         lv_almacenes.ListItems.Item(i).Bold = True
         lv_almacenes.ListItems.Item(i).ForeColor = &HFF0000
         lv_almacenes.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_almacenes.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_almacenes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_almacenes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         rs.Open "insert into tb_temp_rep_mov_almacenes (inte_rmo_numero, vcha_rmo_maquina, vcha_rmo_usuario, vcha_alm_almacen_id) values (" + Str(numero_control) + ", '" + fun_NombrePc + "', '" + var_clave_usuario_global + "', '" + lv_almacenes.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
         lv_almacenes.Refresh
      End If
   End If
End Sub

Private Sub lv_articulos_KeyPress(KeyAscii As Integer)
   Dim i As Integer
   If KeyAscii = 13 Then
      If numero_control = 0 Then
         Dim cmd As New Command
         Dim var_numero_tabla As Double
         Set cmd.ActiveConnection = cnn
         cmd.CommandType = adCmdStoredProc
         cmd.CommandText = "NUMERO_CONTROL"
         cmd("@maquina") = fun_NombrePc
         cmd("@usuario") = var_clave_usuario_global
         cmd("@numero") = numero_control
         cmd.execute
         numero_control = cmd("@numero")
         Set cmd = Nothing
      End If
      i = lv_articulos.selectedItem.Index
      If lv_articulos.selectedItem.SubItems(2) = "*" Then
         lv_articulos.selectedItem.SubItems(2) = ""
         lv_articulos.ListItems.Item(i).Bold = False
         lv_articulos.ListItems.Item(i).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         rs.Open "delete from tb_temp_rep_mov_articulos where inte_rmo_numero = " + Str(numero_control) + " and vcha_rmo_maquina = '" + fun_NombrePc + "' and vcha_rmo_usuario = '" + var_clave_usuario_global + "' and vcha_art_articulo_id = '" + lv_articulos.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
         lv_articulos.Refresh
      Else
         lv_articulos.selectedItem.SubItems(2) = "*"
         lv_articulos.ListItems.Item(i).Bold = True
         lv_articulos.ListItems.Item(i).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         rs.Open "insert into tb_temp_rep_mov_articulos (inte_rmo_numero, vcha_rmo_maquina, vcha_rmo_usuario, vcha_art_articulo_id) values (" + Str(numero_control) + ", '" + fun_NombrePc + "', '" + var_clave_usuario_global + "', '" + lv_articulos.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
         lv_articulos.Refresh
      End If
   End If
End Sub

Private Sub lv_movimientos_KeyPress(KeyAscii As Integer)
   Dim i As Integer
   If KeyAscii = 13 Then
      If numero_control = 0 Then
         Dim cmd As New Command
         Dim var_numero_tabla As Double
         Set cmd.ActiveConnection = cnn
         cmd.CommandType = adCmdStoredProc
         cmd.CommandText = "NUMERO_CONTROL"
         cmd("@maquina") = fun_NombrePc
         cmd("@usuario") = var_clave_usuario_global
         cmd("@numero") = numero_control
         cmd.execute
         numero_control = cmd("@numero")
         Set cmd = Nothing
      End If
      i = lv_movimientos.selectedItem.Index
      If lv_movimientos.selectedItem.SubItems(2) = "*" Then
         lv_movimientos.selectedItem.SubItems(2) = ""
         lv_movimientos.ListItems.Item(i).Bold = False
         lv_movimientos.ListItems.Item(i).ForeColor = &H80000012
         lv_movimientos.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_movimientos.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_movimientos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_movimientos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_movimientos.Refresh
         rs.Open "delete from tb_temp_rep_mov_movimientos where inte_rmo_numero = " + Str(numero_control) + " and vcha_rmo_maquina = '" + fun_NombrePc + "' and vcha_rmo_usuario = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + lv_movimientos.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
      Else
         lv_movimientos.selectedItem.SubItems(2) = "*"
         lv_movimientos.ListItems.Item(i).Bold = True
         lv_movimientos.ListItems.Item(i).ForeColor = &HFF0000
         lv_movimientos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_movimientos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_movimientos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_movimientos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         rs.Open "insert into tb_temp_rep_mov_movimientos (inte_rmo_numero, vcha_rmo_maquina, vcha_rmo_usuario, vcha_mov_movimiento_id) values (" + Str(numero_control) + ", '" + fun_NombrePc + "', '" + var_clave_usuario_global + "', '" + lv_movimientos.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
         lv_movimientos.Refresh
      End If
   End If
End Sub

Private Sub mon_mes1_DateDblClick(ByVal DateDblClicked As Date)
   txt_inicio = mon_mes1.Value
   mon_mes1.Visible = False
End Sub

Private Sub mon_mes2_DateDblClick(ByVal DateDblClicked As Date)
   txt_fin = mon_mes2.Value
   mon_mes2.Visible = False
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim itmfound As ListItem
      var_posible = False
      If rs.State = 1 Then
         rs.Close
      End If
      If var_empresa = "18" Then
         If Len(Trim(Me.txt_buscar)) = 11 Or Len(Trim(Me.txt_buscar)) = 5 Or Len(Trim(Me.txt_buscar)) = 10 Then
            rs.Open "select * from tb_articulos where vcha_art_articulo_id like '" + txt_buscar + "%'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               While Not rs.EOF
                     var_codigo = txt_buscar
                     valor = var_codigo
                     If var_codigo Then
                        Set itmfound = lv_articulos.findItem(valor, lvwSubItem, , lvwPartial)
                     Else
                        Set itmfound = lv_articulos.findItem(valor, lvwText, , lvwPartial)
                     End If
    
                     If itmfound Is Nothing Then
                        Set itmfound = lv_articulos.findItem(valor, lvwText, , lvwPartial)
                        If itmfound Is Nothing Then
                           Set itmfound = lv_articulos.findItem(valor, lvwSubItem, , lvwPartial)
                           If itmfound Is Nothing Then
                              Set list_item = lv_articulos.ListItems.Add(, , rs!vcha_Art_articulo_id)
                              list_item.SubItems(1) = Trim(IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español))
                              txt_buscar = ""
                           Else
                              itmfound.EnsureVisible
                              itmfound.Selected = True
                              lv_articulos.SetFocus
                           End If
                        Else
                           itmfound.EnsureVisible
                           itmfound.Selected = True
                        End If
                     Else
                        itmfound.EnsureVisible
                        itmfound.Selected = True
                     End If
                     rs.MoveNext
               Wend
            Else
               MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
            End If
            rs.Close
            txt_buscar = ""
         
         End If
         If Len(Trim(Me.txt_buscar)) = 12 Then
            rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_buscar + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_posible = True
               rs.Close
            Else
               rs.Close
               rs.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + txt_buscar + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  If rsaux.State = 1 Then
                     rsaux.Close
                  End If
                  rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     var_posible = True
                     txt_buscar = rs!vcha_Art_articulo_id
                     rsaux.Close
                     rs.Close
                  Else
                     var_posible = False
                     rsaux.Close
                     rs.Close
                  End If
               Else
                  rs.Close
               End If
            End If
            rs.Open "select * from tb_articulos where vcha_Art_articulo_id = '" + txt_buscar + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_codigo = txt_buscar
               valor = var_codigo
               If var_codigo Then
                  Set itmfound = lv_articulos.findItem(valor, lvwSubItem, , lvwPartial)
               Else
                  Set itmfound = lv_articulos.findItem(valor, lvwText, , lvwPartial)
               End If
       
               If itmfound Is Nothing Then
                  Set itmfound = lv_articulos.findItem(valor, lvwText, , lvwPartial)
                  If itmfound Is Nothing Then
                     Set itmfound = lv_articulos.findItem(valor, lvwSubItem, , lvwPartial)
                     If itmfound Is Nothing Then
                        Set list_item = lv_articulos.ListItems.Add(, , rs!vcha_Art_articulo_id)
                        list_item.SubItems(1) = Trim(IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español))
                        txt_buscar = ""
                        Exit Sub
                     Else
                        itmfound.EnsureVisible
                        itmfound.Selected = True
                        lv_articulos.SetFocus
                      End If
                  Else
                     itmfound.EnsureVisible
                     itmfound.Selected = True
                     lv_articulos.SetFocus
                  End If
               Else
                  itmfound.EnsureVisible
                  itmfound.Selected = True
                  lv_articulos.SetFocus
               End If
            Else
               MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
            End If
            rs.Close
            txt_buscar = ""
         End If
      Else
         rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_buscar + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_posible = True
            rs.Close
         Else
            rs.Close
            rs.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + txt_buscar + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               If rsaux.State = 1 Then
                  rsaux.Close
               End If
               rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  var_posible = True
                  txt_buscar = rs!vcha_Art_articulo_id
                  rsaux.Close
                  rs.Close
               Else
                  var_posible = False
                  rsaux.Close
                  rs.Close
               End If
            Else
               rs.Close
            End If
         End If
         rs.Open "select * from tb_articulos where vcha_Art_articulo_id = '" + txt_buscar + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_codigo = txt_buscar
            valor = var_codigo
            If var_codigo <> "" Then
               Set itmfound = lv_articulos.findItem(valor, lvwSubItem, , lvwPartial)
            Else
               Set itmfound = lv_articulos.findItem(valor, lvwText, , lvwPartial)
            End If
    
            If itmfound Is Nothing Then
               Set itmfound = lv_articulos.findItem(valor, lvwText, , lvwPartial)
               If itmfound Is Nothing Then
                  Set itmfound = lv_articulos.findItem(valor, lvwSubItem, , lvwPartial)
                  If itmfound Is Nothing Then
                     Set list_item = lv_articulos.ListItems.Add(, , rs!vcha_Art_articulo_id)
                     list_item.SubItems(1) = Trim(IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español))
                     txt_buscar = ""
                     Exit Sub
                  Else
                     itmfound.EnsureVisible
                     itmfound.Selected = True
                     lv_articulos.SetFocus
                   End If
               Else
                  itmfound.EnsureVisible
                  itmfound.Selected = True
                  lv_articulos.SetFocus
               End If
            Else
               itmfound.EnsureVisible
               itmfound.Selected = True
               lv_articulos.SetFocus
            End If
         Else
            MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
         txt_buscar = ""
      End If
   End If
End Sub

Private Sub txt_clave_filtrar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
End Sub

Private Sub txt_fin_Change()
   If Not IsDate(txt_fin) Then
      MsgBox "Fecha Incorrecta", vbOKOnly, "ATENCION"
      txt_fin = Date
   End If
End Sub

Private Sub txt_inicio_LostFocus()
   If Not IsDate(txt_inicio) Then
      MsgBox "Fecha Incorrecta", vbOKOnly, "ATENCION"
      txt_inicio = Date
   End If
End Sub

