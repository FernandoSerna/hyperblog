VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_traspasos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Traspasos"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_mensaje_2 
      Caption         =   "mensaje 2"
      Height          =   195
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8865
      Picture         =   "frmoracle_traspasos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   705
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   795
      Picture         =   "frmoracle_traspasos.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   705
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   465
      Picture         =   "frmoracle_traspasos.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Buscar Movimiento"
      Top             =   705
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      Picture         =   "frmoracle_traspasos.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   705
      Width           =   330
   End
   Begin VB.Frame frm_busqueda 
      Height          =   960
      Left            =   1170
      TabIndex        =   15
      Top             =   570
      Width           =   3135
      Begin VB.TextBox txt_busqueda_folio 
         Height          =   315
         Left            =   195
         TabIndex        =   16
         Top             =   495
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   2
         Left            =   30
         TabIndex        =   17
         Top             =   120
         Width           =   3075
      End
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   1290
      TabIndex        =   18
      Top             =   615
      Width           =   5970
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1950
         Left            =   45
         TabIndex        =   19
         Top             =   420
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   3440
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
         TabIndex        =   20
         Top             =   120
         Width           =   5895
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   60
      TabIndex        =   21
      Top             =   975
      Width           =   9210
   End
   Begin VB.Frame Frame3 
      Height          =   2010
      Index           =   0
      Left            =   6915
      TabIndex        =   32
      Top             =   1095
      Width           =   2355
      Begin VB.TextBox txt_folio 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   510
         Left            =   30
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   945
         Width           =   2265
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   1
         Left            =   30
         TabIndex        =   34
         Top             =   120
         Width           =   2265
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   0
      Left            =   60
      TabIndex        =   22
      Top             =   570
      Width           =   9210
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   60
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_traspasos.frx":0940
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_traspasos.frx":121A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_traspasos.frx":1AF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_traspasos.frx":2090
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_traspasos.frx":296C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_traspasos.frx":3246
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_traspasos.frx":3B20
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_traspasos.frx":3C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_traspasos.frx":3D44
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_traspasos.frx":3E56
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_traspasos.frx":3F68
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_traspasos.frx":407A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   2010
      Index           =   1
      Left            =   105
      TabIndex        =   35
      Top             =   1095
      Width           =   6765
      Begin VB.TextBox txt_folio_envio 
         Height          =   315
         Left            =   1545
         TabIndex        =   45
         Top             =   1545
         Width           =   1635
      End
      Begin VB.TextBox txt_almacen_final 
         Height          =   315
         Left            =   1545
         TabIndex        =   10
         Top             =   1545
         Width           =   825
      End
      Begin VB.TextBox txt_nombre_almacen_final 
         Height          =   315
         Left            =   2385
         TabIndex        =   11
         Top             =   1545
         Width           =   4260
      End
      Begin VB.TextBox txt_nombre_almacen_destino 
         Height          =   315
         Left            =   2385
         TabIndex        =   9
         Top             =   1200
         Width           =   4260
      End
      Begin VB.TextBox txt_almacen_destino 
         Height          =   315
         Left            =   1545
         TabIndex        =   8
         Top             =   1200
         Width           =   825
      End
      Begin VB.TextBox txt_almacen_origen 
         Height          =   315
         Left            =   1545
         TabIndex        =   4
         Top             =   510
         Width           =   825
      End
      Begin VB.TextBox txt_nombre_almacen_origen 
         Height          =   315
         Left            =   2385
         TabIndex        =   5
         Top             =   495
         Width           =   4260
      End
      Begin VB.TextBox txt_unidad_destino 
         Height          =   315
         Left            =   1545
         TabIndex        =   6
         Top             =   855
         Width           =   825
      End
      Begin VB.TextBox txt_nombre_nunidad_destino 
         Height          =   315
         Left            =   2385
         TabIndex        =   7
         Top             =   855
         Width           =   4260
      End
      Begin VB.Label lbl_folio_envio 
         AutoSize        =   -1  'True
         Caption         =   "Folio enviado:"
         Height          =   195
         Left            =   165
         TabIndex        =   46
         Top             =   1605
         Width           =   990
      End
      Begin VB.Label lbl_almacen_final 
         AutoSize        =   -1  'True
         Caption         =   "Almacen final:"
         Height          =   195
         Left            =   165
         TabIndex        =   43
         Top             =   1605
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Almacen destino:"
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   40
         Top             =   1260
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Unidad destino:"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   38
         Top             =   915
         Width           =   1110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Almacén origen:"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   37
         Top             =   555
         Width           =   1140
      End
      Begin VB.Label label 
         BackColor       =   &H8000000D&
         Caption         =   " Datos del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   36
         Top             =   120
         Width           =   6705
      End
   End
   Begin VB.Frame var_numero_nota_traspaso 
      Height          =   4110
      Left            =   105
      TabIndex        =   23
      Top             =   3105
      Width           =   9165
      Begin VB.TextBox txt_foco 
         Height          =   300
         Left            =   9570
         TabIndex        =   14
         Top             =   570
         Width           =   885
      End
      Begin VB.TextBox txt_cantidad 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
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
         Left            =   5115
         TabIndex        =   13
         Top             =   495
         Width           =   1890
      End
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
         Left            =   1545
         TabIndex        =   12
         Top             =   450
         Width           =   2640
      End
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   5445
         TabIndex        =   24
         Top             =   2010
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar 
            Height          =   330
            Left            =   60
            TabIndex        =   25
            Top             =   375
            Width           =   2745
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000D&
            Caption         =   "Cantidad a eliminar"
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   26
            Top             =   15
            Width           =   2895
         End
      End
      Begin MSComctlLib.ListView lv_traspasossalidas 
         Height          =   2580
         Left            =   45
         TabIndex        =   27
         Top             =   1035
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   4551
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   10407
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ESTATUS"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView lv_traspasosentradas 
         Height          =   2580
         Left            =   45
         TabIndex        =   44
         Top             =   1020
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   4551
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2734
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   7585
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Enviado"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Recibido"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Diferencia"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ESTATUS"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lbl_cantidad 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   4410
         TabIndex        =   31
         Top             =   615
         Width           =   675
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Lectura de Artículos"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   0
         Left            =   30
         TabIndex        =   30
         Top             =   120
         Width           =   9090
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código del Artículo:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   615
         Width           =   1395
      End
      Begin VB.Label lbl_cantidad_total 
         Alignment       =   1  'Right Justify
         Caption         =   "99999999999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7035
         TabIndex        =   28
         Top             =   3660
         Width           =   1965
      End
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp2 
      Height          =   135
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   30
      URL             =   "C:\sistemas\desarrollo\integral\CFFOUND.wav"
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
      _cx             =   53
      _cy             =   238
   End
   Begin VB.Label lblnombremovimiento 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   90
      TabIndex        =   39
      Top             =   90
      Width           =   9090
   End
End
Attribute VB_Name = "frmoracle_traspasos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_ventana As Integer
Dim var_primera_vez As Integer
Dim var_encontro As Integer
Dim var_origen_encabezado_id As Integer
Dim var_descripcion_articulo As String
Dim var_fecha_inicio As Date
Dim var_cantidad_leida As Double
Dim var_renglon As Double

Sub ilumina_grid()
   var_n = lv_traspasossalidas.ListItems.Count
   For var_i = 1 To var_n
       If var_i = var_renglon Then
          lv_traspasossalidas.ListItems.Item(var_i).Bold = True
          lv_traspasossalidas.ListItems.Item(var_i).ListSubItems(1).Bold = True
          lv_traspasossalidas.ListItems.Item(var_i).ListSubItems(2).Bold = True
          lv_traspasossalidas.ListItems.Item(var_i).ForeColor = &H8000&
          lv_traspasossalidas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H8000&
          lv_traspasossalidas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H8000&
       Else
          lv_traspasossalidas.ListItems.Item(var_i).Bold = False
          lv_traspasossalidas.ListItems.Item(var_i).ListSubItems(1).Bold = False
          lv_traspasossalidas.ListItems.Item(var_i).ListSubItems(2).Bold = False
          lv_traspasossalidas.ListItems.Item(var_i).ForeColor = &H80000012
          lv_traspasossalidas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H80000012
          lv_traspasossalidas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H80000012
       End If
   Next var_i
   If var_renglon > 0 Then
      lv_traspasossalidas.ListItems.Item(var_renglon).Selected = True
      lv_traspasossalidas.selectedItem.EnsureVisible
   End If
   lv_traspasossalidas.Refresh
End Sub


Sub ilumina_grid_2()
   var_n = lv_traspasosentradas.ListItems.Count
   For var_i = 1 To var_n
       If var_i = var_renglon Then
          lv_traspasosentradas.ListItems.Item(var_i).Bold = True
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(1).Bold = True
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(2).Bold = True
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(3).Bold = True
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(4).Bold = True
          lv_traspasosentradas.ListItems.Item(var_i).ForeColor = &H8000&
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H8000&
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H8000&
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H8000&
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H8000&
       Else
          lv_traspasosentradas.ListItems.Item(var_i).Bold = False
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(1).Bold = False
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(2).Bold = False
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(3).Bold = False
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(4).Bold = False
          lv_traspasosentradas.ListItems.Item(var_i).ForeColor = &H80000012
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H80000012
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H80000012
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H80000012
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H80000012
       End If
   Next var_i
   If var_renglon > 0 Then
      lv_traspasosentradas.ListItems.Item(var_renglon).Selected = True
      lv_traspasosentradas.selectedItem.EnsureVisible
   End If
   lv_traspasosentradas.Refresh
End Sub



Private Sub cmd_buscar_Click()
   Me.frm_busqueda.Visible = True
   Me.txt_busqueda_folio = ""
   Me.txt_busqueda_folio.SetFocus
End Sub

Private Sub cmd_imprimir_Click()
   Dim clnt As New SoapClient30
   Dim var_arreglo() As String
   Dim var_s As String
   'If Me.lv_entradas.ListItems.Count > 0 Then
   '   Me.lv_entradas.ListItems.Item(1).Selected = True
   '   VAR_ESTATUS = Me.lv_entradas.selectedItem.SubItems(13)
   'Else
    '  VAR_ESTATUS = "I"
   'End If
   If VAR_ESTATUS <> "I" Then
      var_si = MsgBox("Se va a imprimir la entrada y cerrar el movimiento", vbYesNo, "ATENCION")
      If var_si = 6 Then
         clnt.MSSoapInit "http://intranet/wsoracle/wsInterfaceINV.asmx?wsdl"
         var_s = clnt.registraDatosRecordSet(CDbl(Me.txt_folio), var_clave_movimiento)
         Set clnt = Nothing
         If var_s = "0" Then
            MsgBox "Se cerro el movimiento correctamente", vbOKOnly, "ATENCION"
         Else
            MsgBox "Error al cerrar el movimiento, vuelvalo a intentar"
         End If
      End If
   Else
   End If
End Sub

Private Sub cmd_mensaje_2_Click()
   Me.wmp2.Controls.Play
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_almacen_destino = ""
   Me.txt_almacen_origen = ""
   Me.txt_busqueda_folio = ""
   Me.txt_cantidad = ""
   Me.txt_nombre_almacen_destino = ""
   Me.txt_nombre_almacen_origen = ""
   Me.txt_unidad_destino = ""
   Me.txt_nombre_nunidad_destino = ""
   Me.lv_traspasossalidas.ListItems.Clear
   Me.lbl_cantidad = ""
   Me.lbl_cantidad_total = ""
   Me.txt_almacen_origen.Enabled = True
   Me.txt_nombre_almacen_origen.Enabled = True
   Me.txt_unidad_destino.Enabled = True
   Me.txt_nombre_nunidad_destino.Enabled = True
   Me.txt_almacen_destino.Enabled = True
   Me.txt_nombre_almacen_destino.Enabled = True
   Me.txt_codigo = ""
   Me.txt_folio = ""
   Me.txt_foco.Enabled = False
   Me.txt_codigo.Enabled = False
   var_primera_vez = 1
   Me.lbl_cantidad_total = "0"
   If var_clave_movimiento = "2" Then
      Me.txt_unidad_destino = var_unidad_organizacional
      rsaux.Open "SELECT * FROM HR_ALL_ORGANIZATION_UNITS WHERE ORGANIZATION_ID = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rsaux.EOF Then
         Me.txt_nombre_nunidad_destino = IIf(IsNull(rsaux!Name), "", rsaux!Name)
      End If
      rsaux.Close
      Me.txt_unidad_destino.Enabled = False
      Me.txt_nombre_nunidad_destino.Enabled = False
   End If
   Me.txt_almacen_origen.SetFocus
End Sub

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Me.lblnombremovimiento = var_nombre_movimiento_global
   Top = 0
   Left = 1000
   var_primera_vez = 1
   Me.frm_busqueda.Visible = False
   Me.frm_eliminar.Visible = False
   Me.frm_lista.Visible = False
   Me.lbl_cantidad.Visible = False
   Me.txt_cantidad.Visible = False
   Me.lbl_cantidad_total = ""
   Me.txt_foco.Enabled = False
   Me.lbl_cantidad_total = "0"
   If var_clave_movimiento = "2" Then
      Me.txt_unidad_destino = var_unidad_organizacional
      rsaux.Open "SELECT * FROM HR_ALL_ORGANIZATION_UNITS WHERE ORGANIZATION_ID = " + Me.txt_unidad_destino, cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rsaux.EOF Then
         Me.txt_nombre_nunidad_destino = IIf(IsNull(rsaux!Name), "", rsaux!Name)
      End If
      rsaux.Close
      Me.txt_unidad_destino.Enabled = False
      Me.txt_nombre_nunidad_destino.Enabled = False
      Me.lv_traspasosentradas.Visible = False
      Me.lbl_folio_envio.Visible = False
      Me.lbl_almacen_final.Visible = True
      Me.txt_almacen_final.Visible = True
      Me.txt_nombre_almacen_final.Visible = True
      Me.txt_folio_envio = ""
      Me.txt_folio_envio.Visible = False
      Me.lv_traspasossalidas.Visible = True
   End If
   If var_clave_movimiento = "51" Then
      Me.cmd_nuevo.Enabled = False
      Me.txt_unidad_destino = var_unidad_organizacional
      rsaux.Open "SELECT * FROM HR_ALL_ORGANIZATION_UNITS WHERE ORGANIZATION_ID = " + Me.txt_unidad_destino, cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rsaux.EOF Then
         Me.txt_nombre_nunidad_destino = IIf(IsNull(rsaux!Name), "", rsaux!Name)
      End If
      rsaux.Close
      
      Me.txt_almacen_destino = var_almacen_destino_traspaso
      rs.Open "select * from mtl_secondary_inventories where organization_id = " + Me.txt_unidad_destino + " AND secondary_inventory_name = '" + var_almacen_destino_traspaso + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_almacen_destino = IIf(IsNull(rs!Description), "", rs!Description)
         Me.txt_nombre_almacen_destino.Enabled = False
      Else
         Me.txt_almacen_destino = ""
         Me.txt_nombre_almacen_destino = ""
         MsgBox "El almacén destino no existe", vbOKOnly, "ATENCION"
      End If
      rs.Close
      Me.txt_almacen_origen = var_almacen_origen_traspaso
      rs.Open "select * from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional + " AND secondary_inventory_name = '" + var_almacen_origen_traspaso + "'  and (disable_date >= SYSDATE or disable_date is null) ", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_almacen_origen = IIf(IsNull(rs!Description), "", rs!Description)
         Me.txt_almacen_origen.Enabled = False
         Me.txt_nombre_almacen_origen.Enabled = False
      Else
         Me.txt_almacen_origen = ""
         Me.txt_nombre_almacen_origen = ""
         MsgBox "El almacén origen no existe", vbOKOnly, "ATENCION"
      End If
      
      rs.Close
      rs.Open "select * from xxvia_vw_transito_sub where organizacion = " + var_unidad_organizacional + " and almacenorigen = '" + var_almacen_origen_traspaso + "' and almacendestino = '" + var_almacen_destino_traspaso + "' and foliodocumento = '" + var_numero_nota_traspaso_n + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         While Not rs.EOF
               Set list_item = Me.lv_traspasosentradas.ListItems.Add(, , IIf(IsNull(rs!CODIGO), "", rs!CODIGO))
               var_cantidad_total = var_cantidad_total + Format(rs!cantidadpendiente)
               list_item.SubItems(1) = IIf(IsNull(rs!DESCRIPCION), "", rs!DESCRIPCION)
               list_item.SubItems(2) = Format(IIf(IsNull(rs!cantidadpendiente), 0, rs!cantidadpendiente), "###,###,##0.00")
               list_item.SubItems(3) = Format(0, "###,###,##0.00")
               list_item.SubItems(4) = Format(IIf(IsNull(rs!cantidadpendiente), 0, rs!cantidadpendiente), "###,###,##0.00")
               rs.MoveNext
         Wend
         rsaux.Open "select codigo_articulo, sum(cantidad) as cantidad from xxvia_tb_traspasos_sub where tipo_transaccion = 51 and subinventario = '" + Me.txt_almacen_origen + "' and almacen_Destino_final = '" + Me.txt_almacen_destino + "' and folio = " + var_numero_nota_traspaso_n + " group by codigo_articulo", cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               For var_j = 1 To lv_traspasosentradas.ListItems.Count
                   Me.lv_traspasosentradas.ListItems.Item(var_j).Selected = True
                   If rsaux!codigo_Articulo = Me.lv_traspasosentradas.selectedItem Then
                      Me.lv_traspasosentradas.selectedItem.SubItems(3) = Format(IIf(IsNull(rsaux!Cantidad), 0, rsaux!Cantidad), "###,###,##0.00")
                      Me.lv_traspasosentradas.selectedItem.SubItems(4) = Format(CDbl(Me.lv_traspasosentradas.selectedItem.SubItems(4)) - IIf(IsNull(rsaux!Cantidad), 0, rsaux!Cantidad), "###,###,##0.00")
                   End If
               Next var_j
               rsaux.MoveNext
         Wend
         rsaux.Close
         
      Else
         MsgBox "El entrada ya fue hecha con anterioridad", vbOKOnly, "ATENCION"
      End If
      rs.Close
      Me.txt_unidad_destino.Enabled = False
      Me.txt_nombre_nunidad_destino.Enabled = False
      Me.txt_almacen_final.Enabled = False
      Me.txt_nombre_almacen_final.Enabled = False
      Me.lv_traspasossalidas.Visible = False
      Me.lv_traspasosentradas.Visible = True
      Me.lbl_folio_envio.Visible = True
      Me.lbl_almacen_final.Visible = False
      Me.txt_almacen_final.Visible = False
      Me.txt_nombre_almacen_final.Visible = False
      Me.txt_folio_envio.Visible = True
      Me.txt_folio_envio.Enabled = False
      Me.txt_folio_envio = var_numero_nota_traspaso_n
      var_primera_vez = 1
   End If
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.lv_lista.ListItems.Count > 0 Then
         If var_ventana = 1 Then
            Me.txt_almacen_origen = Me.lv_lista.selectedItem
            Me.txt_nombre_almacen_origen = Me.lv_lista.selectedItem.SubItems(1)
            Me.txt_almacen_origen.Enabled = True
            Me.txt_almacen_origen.SetFocus
         End If
         If var_ventana = 2 Then
            Me.txt_unidad_destino = Me.lv_lista.selectedItem
            Me.txt_nombre_nunidad_destino = Me.lv_lista.selectedItem.SubItems(1)
            Me.txt_unidad_destino.Enabled = True
            Me.txt_unidad_destino.SetFocus
         End If
         If var_ventana = 3 Then
            Me.txt_almacen_destino = Me.lv_lista.selectedItem
            Me.txt_nombre_almacen_destino = Me.lv_lista.selectedItem.SubItems(1)
            Me.txt_almacen_destino.Enabled = True
            Me.txt_almacen_destino.SetFocus
         End If
         If var_ventana = 4 Then
            Me.txt_almacen_final = Me.lv_lista.selectedItem
            Me.txt_nombre_almacen_final = Me.lv_lista.selectedItem.SubItems(1)
            Me.txt_almacen_final.Enabled = True
            Me.txt_almacen_final.SetFocus
         End If
      
      End If
   End If
   If KeyAscii = 27 Then
         If var_ventana = 1 Then
            If Me.txt_almacen_origen.Enabled = True Then
               Me.txt_almacen_origen.SetFocus
            Else
               Me.frm_lista.Visible = False
            End If
         End If
         If var_ventana = 2 Then
            If Me.txt_unidad_destino.Enabled = True Then
               Me.txt_unidad_destino.SetFocus
            Else
               Me.frm_lista.Visible = False
            End If
         End If
         If var_ventana = 3 Then
            If Me.txt_almacen_destino.Enabled = True Then
               Me.txt_almacen_destino.SetFocus
            Else
               Me.frm_lista.Visible = False
            End If
         End If
         If var_ventana = 4 Then
            If Me.txt_almacen_final.Enabled = True Then
               Me.txt_almacen_final.SetFocus
            Else
               Me.frm_lista.Visible = False
            End If
         End If
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub lv_traspasossalidas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      If Me.lv_traspasossalidas.selectedItem.SubItems(3) <> "I" Then
         Me.frm_eliminar.Visible = True
         Me.txt_cantidad_eliminar = ""
         Me.txt_cantidad_eliminar.SetFocus
      Else
         MsgBox "El movimiento ya no puede ser modificado", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_almacen_destino_Change()
   Me.txt_almacen_final = ""
   Me.txt_nombre_almacen_final = ""
   If Me.txt_almacen_destino = "TRANS" Then
      Me.txt_almacen_final.Enabled = True
      Me.txt_nombre_almacen_final.Enabled = True
   Else
      Me.txt_almacen_destino.Enabled = False
      Me.txt_nombre_almacen_final = ""
   End If
End Sub

Private Sub txt_almacen_destino_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If Me.txt_unidad_destino <> "" Then
         var_ventana = 3
         Me.lv_lista.ListItems.Clear
         rs.Open "select secondary_inventory_name, description from mtl_secondary_inventories where organization_id = " + Me.txt_unidad_destino + " and (disable_date >= SYSDATE or disable_date is null) AND secondary_inventory_name <> '" + Me.txt_almacen_origen + "'order by description", cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs(0).Value)
               list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
               rs.MoveNext
         Wend
         rs.Close
         Me.frm_lista.Visible = True
         Me.lv_lista.SetFocus
      Else
         MsgBox "No se a seleccionado una unidad organizacional", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_almacen_destino_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.txt_almacen_destino <> "" Then
         If Me.txt_almacen_destino = "TRANS" Then
            Me.txt_almacen_final.Enabled = True
            Me.txt_nombre_almacen_final.Enabled = True
            Me.txt_nombre_almacen_destino.Enabled = True
            Me.txt_nombre_almacen_destino.SetFocus
         Else
            Me.txt_nombre_almacen_destino.Enabled = True
            Me.txt_nombre_almacen_destino.SetFocus
         End If
      Else
         Me.txt_nombre_almacen_destino = ""
      End If
   End If
End Sub

Private Sub txt_almacen_destino_LostFocus()
      If Me.txt_almacen_destino <> "" Then
         rs.Open "select * from mtl_secondary_inventories where organization_id = " + Me.txt_unidad_destino + " AND secondary_inventory_name = '" + Me.txt_almacen_destino + "' AND (disable_date >= SYSDATE or disable_date is null)  AND secondary_inventory_name <> '" + Me.txt_almacen_origen + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_nombre_almacen_destino = IIf(IsNull(rs!Description), "", rs!Description)
            Me.txt_nombre_almacen_destino.Enabled = False
            Me.txt_almacen_destino.Enabled = False
            If Me.txt_almacen_destino = "TRANS" Then
               Me.txt_almacen_final.Enabled = True
               Me.txt_nombre_almacen_final.Enabled = True
               Me.txt_almacen_final.SetFocus
            Else
               Me.txt_almacen_final = ""
               Me.txt_nombre_almacen_final = ""
               Me.txt_almacen_final.Enabled = False
               Me.txt_nombre_almacen_final.Enabled = False
               Me.txt_codigo.Enabled = True
               Me.txt_codigo.SetFocus
            End If
         Else
            Me.txt_almacen_destino = ""
            Me.txt_nombre_almacen_destino = ""
            Me.txt_codigo.Enabled = False
            Me.txt_foco.Enabled = False
            MsgBox "El almacén destino no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         Me.txt_nombre_almacen_destino = ""
      End If
End Sub

Private Sub txt_almacen_final_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If Me.txt_unidad_destino <> "" Then
         var_ventana = 4
         Me.lv_lista.ListItems.Clear
         rs.Open "select secondary_inventory_name, description from mtl_secondary_inventories where organization_id = " + Me.txt_unidad_destino + " and (disable_date >= SYSDATE or disable_date is null) AND secondary_inventory_name NOT IN ('" + Me.txt_almacen_origen + "','" + Me.txt_almacen_destino + "') order by description", cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs(0).Value)
               list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
               rs.MoveNext
         Wend
         rs.Close
         Me.frm_lista.Visible = True
         Me.lv_lista.SetFocus
      Else
         MsgBox "No se a seleccionado una unidad organizacional", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_almacen_final_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_almacen_final_LostFocus()
      If Me.txt_almacen_final <> "" Then
         rs.Open "select * from mtl_secondary_inventories where organization_id = " + Me.txt_unidad_destino + " AND secondary_inventory_name = '" + Me.txt_almacen_final + "' AND (disable_date >= SYSDATE or disable_date is null)  AND secondary_inventory_name not in ('" + Me.txt_almacen_origen + "','" + Me.txt_almacen_destino + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_nombre_almacen_final = IIf(IsNull(rs!Description), "", rs!Description)
            Me.txt_nombre_almacen_final.Enabled = False
            Me.txt_almacen_final.Enabled = False
            Me.txt_almacen_final.Enabled = False
            Me.txt_nombre_almacen_final.Enabled = False
            Me.txt_codigo.Enabled = True
            Me.txt_codigo.SetFocus
         Else
            Me.txt_almacen_final = ""
            Me.txt_nombre_almacen_final = ""
            Me.txt_codigo.Enabled = False
            Me.txt_foco.Enabled = False
            MsgBox "El almacén final no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         Me.txt_nombre_almacen_final = ""
      End If
   
End Sub

Private Sub txt_almacen_origen_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_ventana = 1
      Me.lv_lista.ListItems.Clear
      rs.Open "select secondary_inventory_name, description from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional + "  and (disable_date >= SYSDATE or disable_date is null) order by description", cnnoracle_4, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs(0).Value)
            list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
            rs.MoveNext
      Wend
      rs.Close
      Me.frm_lista.Visible = True
      Me.lv_lista.SetFocus
   End If
End Sub

Private Sub txt_almacen_origen_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_almacen_origen_LostFocus()
   If Me.txt_almacen_origen <> "" Then
      rs.Open "select * from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional + " AND secondary_inventory_name = '" + Me.txt_almacen_origen + "'  and (disable_date >= SYSDATE or disable_date is null) ", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_almacen_origen = IIf(IsNull(rs!Description), "", rs!Description)
         Me.txt_almacen_origen.Enabled = False
         Me.txt_nombre_almacen_origen.Enabled = False
      Else
         Me.txt_almacen_origen = ""
         Me.txt_nombre_almacen_origen = ""
         MsgBox "El almacén origen no existe", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      Me.txt_nombre_almacen_origen = ""
   End If
End Sub

Private Sub txt_busqueda_folio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_busqueda_folio) Then
         rs.Open "SELECT * FROM XXVIA_TB_TRASPASOS_SUB WHERE FOLIO = " + Me.txt_busqueda_folio, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            If IIf(IsNull(rs!ESTATUS), "", rs!ESTATUS) = "" Then
               Me.txt_codigo.Enabled = True
            Else
               Me.txt_codigo.Enabled = False
            End If
            Me.txt_almacen_origen.Enabled = False
            Me.txt_nombre_almacen_origen.Enabled = False
            Me.txt_unidad_destino.Enabled = False
            Me.txt_nombre_nunidad_destino.Enabled = False
            Me.txt_almacen_destino.Enabled = False
            Me.txt_nombre_almacen_destino.Enabled = False
            rsaux.Open "select secondary_inventory_name, description from mtl_secondary_inventories where secondary_inventory_name = '" + rs!SUBINVENTARIO + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            Me.txt_almacen_origen = rsaux!secondary_inventory_name
            Me.txt_nombre_almacen_origen = rsaux!Description
            rsaux.Close
            rsaux.Open "select organization_id, name from hr_all_organization_units where organization_id = " + CStr(rs!organizacion_destino), cnnoracle_4, adOpenDynamic, adLockOptimistic
            Me.txt_unidad_destino = rsaux!organization_id
            Me.txt_nombre_nunidad_destino = rsaux!Name
            rsaux.Close
            rsaux.Open "select secondary_inventory_name, description from mtl_secondary_inventories where secondary_inventory_name = '" + rs!SUBINVENTARIO_DESTINO + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            Me.txt_almacen_destino = rsaux!secondary_inventory_name
            Me.txt_nombre_almacen_destino = rsaux!Description
            rsaux.Close
            var_origen_encabezado_id = rs!origen_encabezado_id
            Me.txt_folio = Me.txt_busqueda_folio
            var_cantidad_total = 0
            Me.lv_traspasossalidas.ListItems.Clear
            While Not rs.EOF
                  Set list_item = Me.lv_traspasossalidas.ListItems.Add(, , IIf(IsNull(rs!codigo_Articulo), "", rs!codigo_Articulo))
                  rsaux.Open "select * from xxvia_system_items_b where organization_id = " + var_unidad_organizacional + " and segment1 = '" + rs!codigo_Articulo + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                  list_item.SubItems(1) = IIf(IsNull(rsaux!Description), "", rsaux!Description)
                  End If
                  rsaux.Close
                  var_cantidad_total = var_cantidad_total + Format(rs!Cantidad)
                  list_item.SubItems(2) = Format(rs!Cantidad)
                  list_item.SubItems(3) = IIf(IsNull(rs!ESTATUS), "", rs!ESTATUS)
                  rs.MoveNext
            Wend
            Me.lbl_cantidad_total = Format(var_cantidad_total, "###,###,##0.00")
            Me.frm_busqueda.Visible = False
         Else
            MsgBox "El movimiento no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         MsgBox "Número de movimiento incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_busqueda.Visible = False
   End If
End Sub

Private Sub txt_busqueda_folio_LostFocus()
   Me.frm_busqueda.Visible = False
End Sub

Private Sub txt_cantidad_eliminar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_cantidad_eliminar) Then
         If CDbl(Me.lv_traspasossalidas.selectedItem.SubItems(2)) - CDbl(Me.txt_cantidad_eliminar) >= 0 Then
            rs.Open "UPDATE XXVIA_TB_TRASPASOS_SUB SET CANTIDAD = CANTIDAD - " + Me.txt_cantidad_eliminar + " WHERE FOLIO = " + Me.txt_folio + " AND CODIGO_ARTICULO = '" + Me.lv_traspasossalidas.selectedItem + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            Me.lv_traspasossalidas.selectedItem.SubItems(2) = Format(CDbl(Me.lv_traspasossalidas.selectedItem.SubItems(2)) - CDbl(Me.txt_cantidad_eliminar), "###,###,##0.00")
            Me.lv_traspasossalidas.SetFocus
         Else
            MsgBox "La cantidad a eliminar excede a la leida", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Cantidad a eliminar incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
   If KeyAscii = 27 Then
      If Me.lv_traspasossalidas.ListItems.Count > 0 Then
         Me.lv_traspasossalidas.SetFocus
      Else
        Me.frm_eliminar.Visible = False
      End If
   End If
End Sub

Private Sub txt_cantidad_eliminar_LostFocus()
   Me.frm_eliminar.Visible = False
End Sub

Private Sub txt_cantidad_GotFocus()
   Me.txt_cantidad = ""
End Sub

Private Sub txt_Cantidad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_cantidad) Then
         var_cantidad_leida = CDbl(Me.txt_cantidad)
         Me.txt_foco.Enabled = True
         Me.txt_foco.SetFocus
      Else
         MsgBox "Cantidad_incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
   If KeyAscii = 27 Then
      If Me.txt_codigo.Enabled = True Then
         Me.txt_codigo.SetFocus
      End If
      Me.lbl_cantidad.Visible = False
      Me.txt_cantidad.Visible = False
   End If
End Sub

Private Sub txt_cantidad_LostFocus()
   Me.txt_cantidad.Visible = False
   Me.lbl_cantidad.Visible = False
End Sub

Private Sub txt_codigo_GotFocus()
   Me.txt_codigo = ""
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(Me.txt_codigo) <> "" Then
         If var_clave_movimiento <> "51" Then
            rs.Open "select * from xxvia_system_items_b where organization_id = " + var_unidad_organizacional + " and segment1 = '" + Me.txt_codigo + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_descripcion_articulo = IIf(IsNull(rs!Description), "", rs!Description)
               var_encontro = 0
               For var_j = 1 To Me.lv_traspasossalidas.ListItems.Count
                   Me.lv_traspasossalidas.ListItems.Item(var_j).Selected = True
                   If Me.lv_traspasossalidas.selectedItem = Me.txt_codigo Then
                      var_encontro = var_j
                   End If
               Next var_j
               var_salida_masiva = IIf(IsNull(rs!attribute10), "N", rs!attribute10)
               If var_salida_masiva = "Y" Then
                  Me.txt_foco.Enabled = False
                  Me.txt_cantidad.Visible = True
                  Me.lbl_cantidad.Visible = True
                  Me.txt_cantidad.SetFocus
               Else
                  Me.txt_cantidad.Visible = False
                  Me.lbl_cantidad.Visible = False
                  var_cantidad_leida = 1
                  Me.txt_foco.Enabled = True
                  Me.txt_foco.SetFocus
               End If
            Else
               Call cmd_mensaje_2_Click
               txt_codigo = ""
               frmmensaje.lbl_articulo = ""
               frmmensaje.lbl_mensaje = "El artículo no existe"
               frmmensaje.Show 1
            End If
            rs.Close
         Else
            
            var_encontro = 0
            For var_j = 1 To Me.lv_traspasosentradas.ListItems.Count
                If Me.txt_codigo = Me.lv_traspasosentradas.selectedItem Then
                   var_encontro = var_j
                End If
            Next var_j
            If var_encontro > 0 Then
               rs.Open "select * from xxvia_system_items_b where organization_id = " + var_unidad_organizacional + " and segment1 = '" + Me.txt_codigo + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_salida_masiva = IIf(IsNull(rs!attribute10), "N", rs!attribute10)
                  If var_salida_masiva = "Y" Then
                     Me.txt_foco.Enabled = False
                     Me.txt_cantidad.Visible = True
                     Me.lbl_cantidad.Visible = True
                     Me.txt_cantidad.SetFocus
                  Else
                     Me.txt_cantidad.Visible = False
                     Me.lbl_cantidad.Visible = False
                     var_cantidad_leida = 1
                     Me.txt_foco.Enabled = True
                     Me.txt_foco.SetFocus
                  End If
                  rs.Close
               Else
                  rs.Close
                  Call cmd_mensaje_2_Click
                  txt_codigo = ""
                  frmmensaje.lbl_articulo = ""
                  frmmensaje.lbl_mensaje = "El artículo no existe"
                  frmmensaje.Show 1
               End If
               
            Else
               Call cmd_mensaje_2_Click
               txt_codigo = ""
               frmmensaje.lbl_articulo = ""
               frmmensaje.lbl_mensaje = "El artículo no viene en la relación"
               frmmensaje.Show 1
            End If
         End If
      End If
   End If
End Sub

Private Sub txt_foco_GotFocus()
   If Me.txt_codigo <> "" Then
      If var_primera_vez = 1 Then
         var_primera_vez = 0
         cnnoracle_4.BeginTrans
         rs.Open "select * as folio from xxvia_tb_folios_tr", cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_folio = IIf(IsNull(rs!folio), 0, rs!folio) + 1
            rsaux.Open "update xxvia_tb_folios_tr SET FOLIO = " + CStr(var_folio), cnnoracle_4, adOpenDynamic, adLockOptimistic
         Else
            var_folio = 1
            rsaux.Open "insert INTO xxvia_tb_folios_tr (folio) values (" + CStr(var_folio) + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
         End If
         rs.Close
         Me.txt_folio = var_folio
         rs.Open "select xxvia.XXVIA_SQ_LINEA_TM.NEXTval from DUAL", cnnoracle_4, adOpenDynamic, adLockOptimistic
         var_origen_encabezado_id = rs(0).Value
         rs.Close
         cnnoracle_4.CommitTrans
      End If
      var_descuento = 0
      var_almacen_destino_final = ""
      
      rs.Open "select * from xxvia_tb_traspasos_sub where folio = " + Me.txt_folio + " and codigo_articulo = '" + Me.txt_codigo + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         rsaux.Open "update xxvia_tb_traspasos_sub set cantidad = cantidad + " + CStr(var_cantidad_leida) + " where folio = " + Me.txt_folio + " and codigo_Articulo = '" + Me.txt_codigo + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         If var_clave_movimiento <> "51" Then
            Set itmfound = Me.lv_traspasossalidas.findItem(Me.txt_codigo, lvwText, , lvwPartial)
            itmfound.EnsureVisible
            itmfound.Selected = True
            Me.lv_traspasossalidas.selectedItem.SubItems(2) = Me.lv_traspasossalidas.selectedItem.SubItems(2) + var_cantidad_leida
            Call ilumina_grid
         Else
            Set itmfound = Me.lv_traspasosentradas.findItem(Me.txt_codigo, lvwText, , lvwPartial)
            itmfound.EnsureVisible
            itmfound.Selected = True
            Me.lv_traspasosentradas.selectedItem.SubItems(3) = Me.lv_traspasosentradas.selectedItem.SubItems(3) + var_cantidad_leida
            Me.lv_traspasosentradas.selectedItem.SubItems(4) = Me.lv_traspasosentradas.selectedItem.SubItems(4) - var_cantidad_leida
            Call ilumina_grid
         End If
      Else
         var_cadena = "insert into xxvia_tb_traspasos_sub (folio, organizacion_id, organizacion_destino, subinventario, subinventario_destino, tipo_transaccion, codigo_articulo, cantidad, origen_transaccion, origen_encabezado_id, costo, alias_contable, referencia_transaccion, lote_descuento, almacen_destino_final, estatus, USUARIO, MAQUINA, TIPO_MOVIMIENTO, NUMERO_ORIGEN)"
         var_cadena = var_cadena + " values (" + Me.txt_folio + "," + var_unidad_organizacional + "," + Me.txt_unidad_destino + ", '" + Me.txt_almacen_origen + "','" + Me.txt_almacen_destino + "','" + var_clave_movimiento + "','" + Me.txt_codigo + "'," + CStr(var_cantidad_leida) + ",'VIADIS_INTERFACE'," + CStr(var_origen_encabezado_id) + ",0,NULL," + Me.txt_folio + "," + CStr(var_descuento) + ",'" + Me.txt_almacen_final + "','','" + VAR_CLAVE_USUARIO_FINAL + "','" + fun_NombrePc + "','" + Trim(Me.lblnombremovimiento) + "','" + Me.txt_folio_envio + "')"
         rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If var_clave_movimiento <> "51" Then
            Set list_item = lv_traspasossalidas.ListItems.Add(, , Trim(txt_codigo))
            list_item.SubItems(1) = var_descripcion_articulo
            list_item.SubItems(2) = var_cantidad_leida
            Call ilumina_grid_2
         Else
            Set itmfound = Me.lv_traspasosentradas.findItem(Me.txt_codigo, lvwText, , lvwPartial)
            itmfound.EnsureVisible
            itmfound.Selected = True
            Me.lv_traspasosentradas.selectedItem.SubItems(3) = Me.lv_traspasosentradas.selectedItem.SubItems(3) + var_cantidad_leida
            Me.lv_traspasosentradas.selectedItem.SubItems(4) = Me.lv_traspasosentradas.selectedItem.SubItems(4) - var_cantidad_leida
            Call ilumina_grid_2
         End If
      End If
      rs.Close
      Me.lbl_cantidad_total = Format(CDbl(Me.lbl_cantidad_total) + var_cantidad_leida, "###,###,##0.00")
      Me.txt_codigo.SetFocus
   End If
End Sub

Private Sub txt_nombre_almacen_destino_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If Me.txt_unidad_destino <> "" Then
         var_ventana = 3
         Me.lv_lista.ListItems.Clear
         rs.Open "select secondary_inventory_name, description from mtl_secondary_inventories where organization_id = " + Me.txt_unidad_destino + " AND (disable_date >= SYSDATE or disable_date is null)  AND secondary_inventory_name <> '" + Me.txt_almacen_origen + "' order by description", cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs(0).Value)
               list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
               rs.MoveNext
         Wend
         rs.Close
         Me.frm_lista.Visible = True
         Me.lv_lista.SetFocus
      Else
         MsgBox "No se a seleccionado una unidad organizacional", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_nombre_almacen_destino_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_almacen_final_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If Me.txt_unidad_destino <> "" Then
         var_ventana = 4
         Me.lv_lista.ListItems.Clear
         rs.Open "select secondary_inventory_name, description from mtl_secondary_inventories where organization_id = " + Me.txt_unidad_destino + " and (disable_date >= SYSDATE or disable_date is null) AND secondary_inventory_name NOT IN ('" + Me.txt_almacen_origen + "','" + Me.txt_almacen_destino + "') order by description", cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs(0).Value)
               list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
               rs.MoveNext
         Wend
         rs.Close
         Me.frm_lista.Visible = True
         Me.lv_lista.SetFocus
      Else
         MsgBox "No se a seleccionado una unidad organizacional", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_nombre_almacen_origen_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_ventana = 1
      Me.lv_lista.ListItems.Clear
      rs.Open "select secondary_inventory_name, description from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional + "  and (disable_date >= SYSDATE or disable_date is null) order by description", cnnoracle_4, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs(0).Value)
            list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
            rs.MoveNext
      Wend
      rs.Close
      Me.frm_lista.Visible = True
      Me.lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_almacen_origen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_nunidad_destino_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_ventana = 2
      Me.lv_lista.ListItems.Clear
      rs.Open "select TO_ORGANIZATION_ID, TO_ORGANIZATION_NAME from XXVIA_VW_REDES_ENVIOS WHERE FROM_ORGANIZATION_ID = '" + var_unidad_organizacional + "'  order by TO_ORGANIZATION_NAME", cnnoracle_4, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs(0).Value)
            list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
            rs.MoveNext
      Wend
      rs.Close
      Me.frm_lista.Visible = True
      Me.lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_nunidad_destino_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_unidad_destino_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_ventana = 2
      Me.lv_lista.ListItems.Clear
      rs.Open "select TO_ORGANIZATION_ID, TO_ORGANIZATION_NAME from XXVIA_VW_REDES_ENVIOS WHERE FROM_ORGANIZATION_ID = '" + var_unidad_organizacional + "'  order by TO_ORGANIZATION_NAME", cnnoracle_4, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs(0).Value)
            list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
            rs.MoveNext
      Wend
      rs.Close
      Me.frm_lista.Visible = True
      Me.lv_lista.SetFocus
   End If
End Sub

Private Sub txt_unidad_destino_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_unidad_destino_LostFocus()
   If Me.txt_unidad_destino <> "" Then
      rs.Open "select TO_ORGANIZATION_ID, TO_ORGANIZATION_NAME from XXVIA_VW_REDES_ENVIOS WHERE FROM_ORGANIZATION_ID = '" + var_unidad_organizacional + "' AND TO_ORGANIZATION_ID = " + Me.txt_unidad_destino, cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_nunidad_destino = IIf(IsNull(rs!TO_ORGANIZATION_NAME), "", rs!TO_ORGANIZATION_NAME)
         Me.txt_unidad_destino.Enabled = False
         Me.txt_nombre_nunidad_destino.Enabled = False
      Else
         Me.txt_unidad_destino = ""
         Me.txt_nombre_nunidad_destino = ""
         MsgBox "La unidad organizacional destino no existe", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      Me.txt_nombre_nunidad_destino = ""
   End If
End Sub
