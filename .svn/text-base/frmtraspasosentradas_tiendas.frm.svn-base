VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmtraspasosentradas_tiendas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Traspasos entradas tiendas"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   1890
      TabIndex        =   0
      Top             =   1005
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   30
         TabIndex        =   1
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
         TabIndex        =   2
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame frm_numero_traspaso 
      Height          =   1185
      Left            =   735
      TabIndex        =   11
      Top             =   1170
      Width           =   5985
      Begin VB.TextBox txt_numero_traspaso 
         Height          =   315
         Left            =   765
         TabIndex        =   14
         Top             =   780
         Width           =   1680
      End
      Begin VB.TextBox txt_movimiento 
         Height          =   315
         Left            =   765
         TabIndex        =   13
         Top             =   450
         Width           =   1035
      End
      Begin VB.TextBox txt_nombre_movimiento 
         Height          =   315
         Left            =   1815
         TabIndex        =   12
         Top             =   450
         Width           =   4110
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Número de traspaso"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   4
         Left            =   45
         TabIndex        =   17
         Top             =   120
         Width           =   5880
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   810
         Width           =   600
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tienda:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   465
         Width           =   540
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1335
      Index           =   0
      Left            =   5235
      TabIndex        =   39
      Top             =   1155
      Width           =   3210
      Begin VB.TextBox txt_folio 
         Alignment       =   1  'Right Justify
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
         Height          =   480
         Left            =   1605
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   420
         Width           =   1500
      End
      Begin VB.TextBox txt_folio_enviado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1425
         TabIndex        =   40
         Top             =   945
         Width           =   1680
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   1
         Left            =   30
         TabIndex        =   44
         Top             =   120
         Width           =   3135
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Número Movimiento:"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   43
         Top             =   570
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Número Enviado:"
         Height          =   195
         Index           =   5
         Left            =   150
         TabIndex        =   42
         Top             =   990
         Width           =   1230
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   0
      Left            =   75
      TabIndex        =   19
      Top             =   630
      Width           =   8430
   End
   Begin VB.TextBox txt_foco 
      Height          =   285
      Left            =   8865
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   3705
      Width           =   1125
   End
   Begin VB.Frame frm_busqueda 
      Height          =   960
      Left            =   345
      TabIndex        =   8
      Top             =   1155
      Width           =   3135
      Begin VB.TextBox txt_busqueda_folio 
         Height          =   315
         Left            =   180
         TabIndex        =   9
         Top             =   465
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   2
         Left            =   30
         TabIndex        =   10
         Top             =   120
         Width           =   3075
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1095
      Picture         =   "frmtraspasosentradas_tiendas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Cancelar Movimiento"
      Top             =   750
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frmtraspasosentradas_tiendas.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   750
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   435
      Picture         =   "frmtraspasosentradas_tiendas.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Buscar Movimiento"
      Top             =   750
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   765
      Picture         =   "frmtraspasosentradas_tiendas.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   750
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8085
      Picture         =   "frmtraspasosentradas_tiendas.frx":0408
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   780
      Width           =   330
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   495
      Top             =   60
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
            Picture         =   "frmtraspasosentradas_tiendas.frx":0A42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasosentradas_tiendas.frx":131C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasosentradas_tiendas.frx":1BF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasosentradas_tiendas.frx":2192
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasosentradas_tiendas.frx":2A6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasosentradas_tiendas.frx":3348
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasosentradas_tiendas.frx":3C22
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasosentradas_tiendas.frx":3D34
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasosentradas_tiendas.frx":3E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasosentradas_tiendas.frx":3F58
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasosentradas_tiendas.frx":406A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasosentradas_tiendas.frx":417C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   1335
      Index           =   1
      Left            =   120
      TabIndex        =   32
      Top             =   1155
      Width           =   5085
      Begin VB.TextBox txt_almacen_origen 
         Height          =   345
         Left            =   780
         TabIndex        =   34
         Top             =   480
         Width           =   4200
      End
      Begin VB.TextBox txt_almacen_destino 
         Height          =   345
         Left            =   780
         TabIndex        =   33
         Top             =   855
         Width           =   4200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   37
         Top             =   870
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Origen:"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   36
         Top             =   510
         Width           =   510
      End
      Begin VB.Label label 
         BackColor       =   &H8000000D&
         Caption         =   " Datos del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   35
         Top             =   120
         Width           =   5010
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4860
      Left            =   120
      TabIndex        =   20
      Top             =   2490
      Width           =   8340
      Begin VB.TextBox txt_cantidad 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
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
         TabIndex        =   25
         Top             =   495
         Width           =   1890
      End
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   2175
         TabIndex        =   22
         Top             =   2100
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar 
            Height          =   330
            Left            =   60
            TabIndex        =   23
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
            TabIndex        =   24
            Top             =   15
            Width           =   2895
         End
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
         TabIndex        =   21
         Top             =   465
         Width           =   2640
      End
      Begin MSComctlLib.ListView lv_traspasosentradas 
         Height          =   3375
         Left            =   60
         TabIndex        =   26
         Top             =   1035
         Width           =   8220
         _ExtentX        =   14499
         _ExtentY        =   5953
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   6350
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Enviaron"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Recibidos"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "COSTO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Diferencia"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "año"
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
         Left            =   45
         TabIndex        =   30
         Top             =   120
         Width           =   8250
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
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
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
         Height          =   375
         Left            =   3540
         TabIndex        =   28
         Top             =   4395
         Width           =   2715
      End
      Begin VB.Label lbl_cantidad_leida 
         Alignment       =   1  'Right Justify
         Caption         =   "9999999999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5535
         TabIndex        =   27
         Top             =   4410
         Width           =   2715
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   75
      TabIndex        =   38
      Top             =   1020
      Width           =   8430
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
      Left            =   105
      TabIndex        =   45
      Top             =   135
      Width           =   8325
   End
End
Attribute VB_Name = "frmtraspasosentradas_tiendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_unidad_origen As String
Dim var_almacen_Destino As String
Dim var_almacen_origen As String
Dim var_primera_vez As Boolean
Dim var_numero_folio As Double
Dim var_cantidad_leida As Double
Dim var_costo As Double
Dim var_precio As Double
Dim var_descripcion_articulo As String
Dim var_estatus_movimiento As String
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_movimiento_salida As String
Dim var_numero_salida As Double
Dim var_tabla As ADODB.Connection
Dim var_ruta As String
Dim var_clave_moneda As String
Dim var_año As Integer
Dim var_renglon As Double

Sub ilumina_grid()
   var_n = lv_traspasosentradas.ListItems.Count
   For var_i = 1 To var_n
       If var_i = var_renglon Then
          lv_traspasosentradas.ListItems.Item(var_i).Bold = True
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(1).Bold = True
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(2).Bold = True
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(3).Bold = True
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(4).Bold = True
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(5).Bold = True
          lv_traspasosentradas.ListItems.Item(var_i).ForeColor = &H8000&
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H8000&
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H8000&
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H8000&
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H8000&
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(5).ForeColor = &H8000&
       Else
          lv_traspasosentradas.ListItems.Item(var_i).Bold = False
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(1).Bold = False
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(2).Bold = False
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(3).Bold = False
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(4).Bold = False
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(5).Bold = False
          lv_traspasosentradas.ListItems.Item(var_i).ForeColor = &H80000012
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H80000012
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H80000012
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H80000012
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H80000012
          lv_traspasosentradas.ListItems.Item(var_i).ListSubItems(5).ForeColor = &H80000012
       End If
   Next var_i
   If var_renglon > 0 Then
      lv_traspasosentradas.ListItems.Item(var_renglon).Selected = True
      lv_traspasosentradas.selectedItem.EnsureVisible
   End If
   lv_traspasosentradas.Refresh
End Sub



Private Sub cmd_buscar_Click()
            frm_busqueda.Visible = True: var_ventana = 1
            txt_busqueda_folio.SetFocus
End Sub

Private Sub cmd_imprimir_Click()
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
   Set TB_EXISTENCIAS_INSERTA = New TB_EXISTENCIAS_INSERTA
   Set TB_ENTRADAS_INSERTA = New TB_ENTRADAS_INSERTA
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
   Set var_tabla = CreateObject("ADODB.connection")
            If var_numero_folio > 0 Then
               If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
                  Set reporte = appl.OpenReport(App.Path + "\rep_entradas_traspasos.rpt")
                  reporte.RecordSelectionFormula = "{VW_ENTRADAS_TRASPASOS.VCHA_EMO_MOVIMIENTO_ORIGEN} = '" + var_movimiento_salida + "' and {VW_ENTRADAS_TRASPASOS.INTE_EMO_NUMERO_ORIGEN} = " + Str(var_numero_salida) + " and {VW_ENTRADAS_TRASPASOS.INTE_EMO_NUMERO} = " + Str(var_numero_folio)
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Reporte de Movimientos"
                  frmvistasprevias.Show 1
                  Set reporte = Nothing
                  rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_alm_almacen_id = '" + var_almacen_Destino + "'", cnn, adOpenDynamic, adLockOptimistic
               Else
                  var_si = MsgBox("¿Se va a imprimir el movimiento?", vbOKCancel, "ATENCION")
                  If var_si = 1 Then
                     cnn.BeginTrans
                     Cadena = "select * from tb_TRASPASOS where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_movimiento_salida + "' and inte_TRA_numero = " + Str(var_numero_salida) + " and floa_tra_cantidad_recibida is not null"
                     rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                     cnntraspasos_tiendas.BeginTrans
                     While Not rs.EOF
                         'var_inserta = False
                         'var_inserta = TB_ENTRADAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_destino, var_clave_movimiento, var_numero_folio, rs!vcha_art_articulo_id, rs!FLOA_TRA_CANTIDAD_RECIBIDA, rs!FLOA_TRA_COSTO, rs!FLOA_TRA_PRECIO, IIf(IsNull(rs!FLOA_TRA_DESCUENTO), 0, rs!FLOA_TRA_DESCUENTO), var_almacen_origen, CInt(rs!inte_tra_año))
                         rsaux2.Open "insert into tb_entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id,INTE_ent_NUMERO,vcha_art_articulo_id,floa_ent_cantidad, floa_ent_costo, floa_ent_precio, floa_ent_descuento, VCHA_ENT_ALMACEN_ORIGEN, inte_ent_año) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", '" + rs!vcha_Art_articulo_id + "', " + CStr(rs!FLOA_TRA_CANTIDAD_RECIBIDA) + ", " + CStr(rs!floa_Tra_Costo) + ", " + CStr(rs!FLOA_TRA_PRECIO) + ", " + CStr(IIf(IsNull(rs!FLOA_TRA_DESCUENTO), 0, rs!FLOA_TRA_DESCUENTO)) + ", '" + var_almacen_origen + "', " + CStr(rs!INTE_TRA_AÑO) + ")", cnn, adOpenDynamic, adLockOptimistic
                         rsaux2.Open "UPDATE TB_TRASPASOS SET NUME_TRA_CANTIDAD_DESTINO = NUME_TRA_CANTIDAD_DESTINO + " + CStr(rs!FLOA_TRA_CANTIDAD_RECIBIDA) + ", DTIM_TRA_FECHA_DESTINO = SYSDATE WHERE vcha_emp_empresa_id = '" + rs!vcha_Tra_empresa_Externa + "' and nume_tra_folio = " + rs!inte_tra_numero_Externo + " and vcha_Can_canal_id = '01' and vcha_tra_empresa_id_destino = '0500' and vcha_Art_Articulo_id = '" + rs!VCHA_TRA_CODIGO_EXTERNO + "'", cnntraspasos_tiendas, adOpenDynamic, adLockOptimistic
                         rs.MoveNext
                     Wend
                     cnntraspasos_tiendas.CommitTrans
                     rs.Close
                     var_estatus_movimiento = "I"
                     var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "I", Now, 1)
                     rs.Open "update tb_encabezado_movimientos set DTIM_EMO_FECHA_FINALIZO = getdate() where vcha_Emp_empresa_id = '" + var_empresa + "' and  VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_EMO_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                     cnn.CommitTrans
                     Set reporte = appl.OpenReport(App.Path + "\rep_entradas_traspasos.rpt")
                     reporte.RecordSelectionFormula = "{VW_ENTRADAS_TRASPASOS.VCHA_EMO_MOVIMIENTO_ORIGEN} = '" + var_movimiento_salida + "' and {VW_ENTRADAS_TRASPASOS.INTE_EMO_NUMERO_ORIGEN} = " + Str(var_numero_salida) + " and {VW_ENTRADAS_TRASPASOS.INTE_EMO_NUMERO} = " + Str(var_numero_folio)
                     frmvistasprevias.cr.ReportSource = reporte
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     frmvistasprevias.cr.ViewReport
                     frmvistasprevias.Caption = "Reporte de Movimientos"
                     frmvistasprevias.Show 1
                     Set reporte = Nothing
                     rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_alm_almacen_id = '" + var_almacen_Destino + "'", cnn, adOpenDynamic, adLockOptimistic
                     txt_codigo.Enabled = False
                     txt_foco.Enabled = False
                  End If
               End If
            Else
               MsgBox "No se a seleccionado ningún movimiento", vbOKOnly, "ATENCION"
            End If
End Sub

Private Sub cmd_nuevo_Click()
   If var_numero_folio > 0 Then
      rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 where inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
   End If
   txt_codigo.Enabled = False
   var_primera_vez = True
   frm_busqueda.Visible = False: var_ventana = 0
   lv_traspasosentradas.ListItems.Clear
   var_numero_folio = 0
   txt_folio = ""
   txt_codigo = ""
   var_estatus_movimiento = ""
   var_movimiento_salida = ""
   frm_numero_traspaso.Visible = True
   txt_movimiento = ""
   txt_nombre_movimiento = ""
   txt_movimiento.Enabled = True
   Me.txt_nombre_movimiento.Enabled = True
   txt_movimiento.SetFocus
End Sub

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = 116 Then
      frmexisten_rapidas.Show 1
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
End Sub

Private Sub Form_Load()
   var_posible_kanban = 0
   var_cadena_seguridad = ""
   frm_lista.Visible = False
   Top = 0
   Left = 1600
   rs.Open "select * from tb_monedas where inte_mon_moneda_local = 1", cnn, adOpenDynamic
   var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
   rs.Close
   var_estatus_movimiento = ""
   frm_busqueda.Visible = False: var_ventana = 0
   frm_eliminar.Visible = False
   frm_numero_traspaso.Visible = False
   lbl_cantidad.Visible = False
   txt_cantidad.Visible = False
   txt_codigo.Enabled = False
   var_primera_vez = True
   var_cantidad_leida = 1#
   txt_almacen_origen = ""
   txt_almacen_destino = ""
   txt_almacen_origen.Enabled = False
   txt_almacen_destino.Enabled = False
   Me.lbl_cantidad_leida = "0"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_numero_folio > 0 Then
      rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 where inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "'", cnn, adOpenDynamic, adLockOptimistic
   End If
   Call activa_forma(var_activa_forma_traspasosentradas)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         txt_movimiento = lv_lista.selectedItem
         txt_nombre_movimiento = lv_lista.selectedItem.SubItems(1)
      Else
         txt_movimiento = ""
         txt_nombre_movimiento = ""
      End If
      txt_movimiento.SetFocus
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
    frm_lista.Visible = False
End Sub

Private Sub lv_traspasosentradas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
         MsgBox "Imporsible modificar el movimiento", vbOKOnly, "ATENCION"
      Else
         frm_eliminar.Visible = True
         txt_cantidad_eliminar.SetFocus
      End If
   End If
End Sub

Private Sub lv_traspasosentradas_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub




Private Sub txt_busqueda_folio_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 27 Then
       frm_busqueda.Visible = False: var_ventana = 0
   
   
   End If
   If KeyAscii = 13 Then
      If Trim(txt_busqueda_folio) <> "" Then
         If var_numero_folio = CDbl(txt_busqueda_folio) Then
            rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 where inte_emo_numero = " + txt_busqueda_folio + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
         End If
         rs.Open "select * from tb_encabezado_movimientos where inte_emo_numero = " + txt_busqueda_folio + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            If var_numero_folio > 0 Then
               rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 where inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
            End If
            var_movimiento_bloqueado = IIf(IsNull(rs!INTE_EMO_BLOQUEADO), 0, rs!INTE_EMO_BLOQUEADO)
            If var_movimiento_bloqueado = 0 Then
               var_almacen_destino_tem = rs!VCHA_EMO_ALMACEN_DESTINO
               var_almacen_origen_tem = rs!vcha_emo_almacen_origen
               var_posible = 1
               If var_tipo_permiso = 1 Then
                  rsaux.Open "select * from tb_permisos_movimientos where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_per_almacen_1 = '" + var_almacen_destino_tem + "'", cnn, adOpenDynamic, adLockBatchOptimistic
                  If rsaux.EOF Then
                     var_posible = 0
                  End If
                  rsaux.Close
                  rsaux.Open "select * from tb_permisos_movimientos where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_per_almacen_2 = '" + var_almacen_origen_tem + "'", cnn, adOpenDynamic, adLockBatchOptimistic
                  If rsaux.EOF Then
                     var_posible = 0
                  End If
                  rsaux.Close
               End If
               var_posible = 1
               If var_posible = 1 Then
                  var_numero_folio = rs!INTE_EMO_NUMERO
                  var_numero_salida = rs!INTE_EMO_NUMERO_ORIGEN
                  var_movimiento_salida = rs!VCHA_EMO_MOVIMIENTO_ORIGEN
                  var_almacen_Destino = rs!VCHA_EMO_ALMACEN_DESTINO
                  var_almacen_origen = rs!vcha_emo_almacen_origen
                  var_estatus_movimiento = rs!char_Emo_estatus
                  rs.Close
                  var_primera_vez = False
                  lv_traspasosentradas.ListItems.Clear
                  txt_folio_enviado = var_numero_salida
                  txt_folio = var_numero_folio
                  rsaux.Open "select * from tb_ALMACENES where vcha_alm_almacen_id = '" + var_almacen_Destino + "'", cnn, adOpenDynamic, adLockOptimistic
                  txt_almacen_destino = rsaux(2).Value
                  txt_almacen_destino.Text = rsaux(3).Value
                  txt_almacen_destino.Enabled = False
                  rsaux.Close
                  rsaux.Open "select * from tb_ALMACENES where vcha_alm_almacen_id = '" + var_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
                  var_unidad_origen = IIf(IsNull(rsaux!vcha_uor_unidad_id), "", rsaux!vcha_uor_unidad_id)
                  txt_almacen_origen = rsaux(2).Value
                  txt_almacen_origen.Text = rsaux(3).Value
                  txt_almacen_origen.Enabled = False
                  rsaux.Close
                  lbl_cantidad_leida = Format("0", "###,###,##0.0000")
                  rsaux.Open "select * from tb_TRASPASOS where inte_TRA_numero = " + Str(var_numero_salida) + " and vcha_mov_movimiento_id = '" + var_movimiento_salida + "' and vcha_emp_empresa_id = '" + var_empresa + "' order by vcha_art_articulo_id, inte_tra_año", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     While Not rsaux.EOF
                        rsaux2.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rsaux!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           Set list_item = lv_traspasosentradas.ListItems.Add(, , rsaux!vcha_Art_articulo_id)
                           list_item.SubItems(1) = IIf(IsNull(rsaux2(1).Value), "", rsaux2(1).Value)
                           list_item.SubItems(2) = Format(Round(IIf(IsNull(rsaux!floa_tra_Cantidad), 0, rsaux!floa_tra_Cantidad), 4), "###,###,##0.0000")
                           list_item.SubItems(3) = Format(Round(IIf(IsNull(rsaux!FLOA_TRA_CANTIDAD_RECIBIDA), 0, rsaux!FLOA_TRA_CANTIDAD_RECIBIDA), 4), "###,###,##0.0000")
                           list_item.SubItems(4) = IIf(IsNull(rsaux!floa_Tra_Costo), 0, rsaux!floa_Tra_Costo)
                           list_item.SubItems(5) = Format(Round(CDbl(list_item.SubItems(2)) - CDbl(list_item.SubItems(3)), 4), "###,###,##0.0000")
                           list_item.SubItems(6) = IIf(IsNull(rsaux!INTE_TRA_AÑO), 0, rsaux!INTE_TRA_AÑO)
                           lbl_cantidad_leida = Format(CDbl(lbl_cantidad_leida) + Round(IIf(IsNull(rsaux!FLOA_TRA_CANTIDAD_RECIBIDA), 0, rsaux!FLOA_TRA_CANTIDAD_RECIBIDA), 4), "###,###,##0.0000")
                           rsaux2.Close
                           rsaux.MoveNext:
                        End If
                     Wend
                     rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 1 where inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "'", cnn, adOpenDynamic, adLockOptimistic
                     frm_busqueda.Visible = False: var_ventana = 0
                     If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
                        txt_codigo.Enabled = False
                     Else
                        txt_codigo.Enabled = True
                        txt_codigo.SetFocus
                     End If
                  Else
                     MsgBox "El traspaso no a sido enviado o no a sido impreso", vbOKOnly, "ATENCION"
                  End If
                  rsaux.Close
               Else
                  MsgBox "No esta autorizado para afectar este movimiento"
                  rs.Close
               End If
            Else
               rs.Close
               MsgBox "El movimiento esta siendo utilizado por otro usuario ", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El número de movimiento no existe ", vbOKOnly, "ATENCION"
            rs.Close
         End If
      End If
      frm_numero_traspaso.Visible = False
   End If
End Sub

Private Sub txt_cantidad_eliminar_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If IsNumeric(txt_cantidad_eliminar) Then
         Set TB_TRASPASOS_MODIFICA = New TB_TRASPASOS_MODIFICA
         var_cantidad_eliminar = Round(CDbl(txt_cantidad_eliminar), 4)
         If var_cantidad_eliminar <= CDbl(lv_traspasosentradas.selectedItem.SubItems(3)) Then
            var_inserta = False
            var_año = CInt(lv_traspasosentradas.selectedItem.SubItems(6) * 1)
            var_inserta = TB_TRASPASOS_MODIFICA.Anadir(var_empresa, var_unidad_origen, var_almacen_origen, var_movimiento_salida, CInt(var_numero_salida), lv_traspasosentradas.selectedItem, 0 - CDbl(txt_cantidad_eliminar), var_almacen_origen, var_año)
            lbl_cantidad_leida = Format(CDbl(lbl_cantidad_leida) - var_cantidad_eliminar, "###,###,##0.0000")
            lv_traspasosentradas.selectedItem.SubItems(3) = Format(Round(lv_traspasosentradas.selectedItem.SubItems(3) - CDbl(txt_cantidad_eliminar), 4), "###,###,##0.0000")
            lv_traspasosentradas.selectedItem.SubItems(5) = Format(Round(lv_traspasosentradas.selectedItem.SubItems(5) + CDbl(txt_cantidad_eliminar), 4), "###,###,##0.0000")
            var_renglon = lv_traspasosentradas.selectedItem.Index
            Call ilumina_grid
            frm_eliminar.Visible = False
            txt_codigo.SetFocus
         Else
            MsgBox "La cantidad no debe de ser mayor a la enviada", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Cantidad incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
   If KeyAscii = 27 Then
      frm_eliminar.Visible = False
      txt_codigo.SetFocus
   End If
End Sub

Private Sub txt_cantidad_GotFocus()
   Me.txt_cantidad = ""
End Sub

Private Sub txt_cantidad_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If Trim(txt_cantidad) <> "" Then
         If IsNumeric(txt_cantidad) Then
            var_cantidad_leida = txt_cantidad
            txt_foco.Enabled = True
            txt_foco.SetFocus
            lbl_cantidad.Visible = False
            txt_cantidad.Visible = False
         Else
            MsgBox "Cantidad incorrecta", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub

Private Sub txt_codigo_GotFocus()
   txt_codigo = ""
End Sub

Private Sub txt_codigo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_codigo_seleccionado = ""
      frmbusqueda_articulo.Show 1
      Me.txt_codigo = var_codigo_seleccionado
   End If
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   Dim var_recontable As Integer
   Dim var_caja As String
   Dim var_cantidad_caja As Integer
   txt_codigo = Trim(txt_codigo)
   If KeyAscii = 13 Then
      var_costo = 0
      var_precio = 0
      var_caja = Left(txt_codigo, 6)
      var_cantidad_caja = 0
      If var_caja = "000005" Or var_caja = "000010" Or var_caja = "000015" Or var_caja = "000020" Then
         var_cantidad_caja = CInt(var_caja)
         txt_codigo = Mid(txt_codigo, 7, 5)
      End If
      If Trim(txt_codigo) <> "" Then
         rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            If IsNull(rs(43).Value) Then
               var_recontable = 0
            Else
               var_recontable = rs(43).Value
            End If
            var_descripcion_articulo = rs(1).Value
            var_costo = rs(3).Value
            var_precio = rs(2).Value
            rs.Close
            rs.Open "SELECT * FROM TB_EXISTENCIAS WHERE VCHA_ART_ARTICULO_ID = '" + txt_codigo + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_costo = IIf(IsNull(rs!floa_exi_costo_2005), 0, rs!floa_exi_costo_2005)
            End If
            If var_costo = 0 Then
               var_costo = rs!FLOA_eXI_COSTO
            End If
            rs.Close
            If var_recontable = 1 Then
               var_cantidad_leida = 1#
               lbl_cantidad.Visible = True
               txt_cantidad.Visible = True
               txt_cantidad.SetFocus
            Else
               var_cantidad_leida = 1#
               txt_foco.Enabled = True
               txt_foco.SetFocus
            End If
         Else
            rs.Close
            rs.Open "select * from tb_equivalencias where VCHA_EQU_CODIGO_EQUIVALENTE = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               txt_codigo = rs(0).Value
               rs.Close
               rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  If var_cantidad_caja = 0 Then
                     If IsNull(rs!inte_Art_salida_masiva) Then
                        var_recontable = 0
                     Else
                        var_recontable = rs!inte_Art_salida_masiva
                     End If
                  Else
                     var_recontable = 0
                  End If
                  var_descripcion_articulo = rs(1).Value
                  var_costo = IIf(IsNull(rs(3).Value), 0, rs(3).Value)
                  var_precio = IIf(IsNull(rs(2).Value), 0, rs(2).Value)
                  rs.Close
                  rs.Open "SELECT * FROM TB_EXISTENCIAS WHERE VCHA_ART_ARTICULO_ID = '" + txt_codigo + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_costo = rs!floa_exi_costo_2005
                  Else
                     var_costo = 0
                  End If
                  rs.Close
                  If var_recontable = 1 Then
                     var_cantidad_leida = 1#
                     lbl_cantidad.Visible = True
                     txt_cantidad.Visible = True
                     txt_cantidad.SetFocus
                  Else
                     If var_cantidad_caja = 0 Then
                        var_cantidad_leida = 1#
                     Else
                        var_cantidad_leida = var_cantidad_caja
                     End If
                     txt_foco.Enabled = True
                     txt_foco.SetFocus
                  End If
               Else
                  txt_codigo = ""
                  frmmensaje.lbl_mensaje = "El artículo no existe"
                  frmmensaje.Show 1
                  'MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
               End If
            Else
               txt_codigo = ""
               frmmensaje.lbl_mensaje = "El artículo no existe"
               frmmensaje.Show 1
               'MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
               rs.Close
            End If
         End If
      Else
      End If
   End If
End Sub

Private Sub txt_foco_GotFocus()
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TRASPASOS_INSERTA = New TB_TRASPASOS_INSERTA
   Set TB_TRASPASOS_MODIFICA = New TB_TRASPASOS_MODIFICA
   Dim var_inserta As Boolean
   Dim var_posible As Boolean
   If Trim(txt_codigo.Text) <> "" Then
      bandera_suma = False
      If var_primera_vez = True Then
         var_inserta = False
         var_insreta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, Now, var_numero_folio, Str(var_numero_salida), "", "", var_almacen_origen, var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, "", var_movimiento_salida, "", "", "B", "", "", 0, 0, 0, var_clave_moneda, 1)
         var_numero_folio = var_numero_folio_regreso
         txt_folio = var_numero_folio
         var_primera_vez = False
      End If
      'Cadena = "select * from TB_TRASPASOS where vcha_alm_almacen_id = " + var_almacen_origen + "and  VCHA_MOV_MOVIMIENTO_ID = '" + var_movimiento_salida + "' and inte_tra_numero = " + Str(var_numero_salida) + " and vcha_art_articulo_id = '" + txt_codigo + "' AND VCHA_TRA_ALMACEN_ORIGEN = '" + var_almacen_origen + "'"
      'rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
      var_suma_cantidad = 0
      var_cantidad_leida_llegar = Round(var_cantidad_leida, 4)
      var_cantidad_leida = Round(var_cantidad_leida, 4)
      var_n = Me.lv_traspasosentradas.ListItems.Count
      var_encontro = 0
      var_i = 1
      var_posible = False
      x = 1
      While (var_i <= var_n)
            lv_traspasosentradas.ListItems.Item(var_i).Selected = True
            valor = Trim(lv_traspasosentradas.selectedItem)
            If txt_codigo = valor Then
               var_posible = True
               var_año = 2005
               var_i = var_n
               
               If CDbl(lv_traspasosentradas.selectedItem.SubItems(3)) + Round(var_cantidad_leida, 4) > Me.lv_traspasosentradas.selectedItem.SubItems(2) Then
                  MsgBox "La cantidad supera a la enviada por la tienda", vbOKOnly, "ATENCION"
                  x = 0
                  var_posible = False
               Else
                  x = 1
                  var_posible = True
               End If
            End If
            var_i = var_i + 1
         Wend
         If var_posible Then
            var_inserta = False
            var_costo = lv_traspasosentradas.selectedItem.SubItems(4)
            'var_inserta = TB_TRASPASOS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_movimiento_salida, var_numero_salida, txt_codigo, var_cantidad_leida, var_almacen_origen, var_año)
            Cadena = "UPDATE [TB_TRASPASOS] SET  [FLOA_TRA_CANTIDAD_RECIBIDA]    = ISNULL([FLOA_TRA_CANTIDAD_RECIBIDA],0) + " + CStr(var_cantidad_leida) + " Where ([VCHA_EMP_EMPRESA_ID]  = '" + var_empresa + "' AND [VCHA_ALM_ALMACEN_ID] = '" + var_almacen_origen + "' and  [VCHA_MOV_MOVIMIENTO_ID] = '" + var_movimiento_salida + "' AND [INTE_TRA_NUMERO]   = " + CStr(var_numero_salida) + " AND [VCHA_ART_ARTICULO_ID]  = '" + txt_codigo + "' AND [VCHA_TRA_ALMACEN_ORIGEN]   = '" + var_almacen_origen + "' and [inte_tra_año] = " + CStr(var_año) + ")"
            lbl_cantidad_leida = Format(CDbl(lbl_cantidad_leida) + var_cantidad_leida, "###,###,##0.0000")
            rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
            lv_traspasosentradas.selectedItem.SubItems(3) = Format(Round(CDbl(lv_traspasosentradas.selectedItem.SubItems(3)) + var_cantidad_leida, 4), "###,###,##0.0000")
            lv_traspasosentradas.selectedItem.SubItems(5) = Format(Round(CDbl(lv_traspasosentradas.selectedItem.SubItems(2)) - CDbl(lv_traspasosentradas.selectedItem.SubItems(3)), 4), "###,###,##0.0000")
            var_renglon = lv_traspasosentradas.selectedItem.Index
            Call ilumina_grid
         Else
            If x = 1 Then
               MsgBox "El artículo no viene en la relación", vbOKOnly, "ATENCION"
            End If
            'var_i = var_n
            'var_cantidad_leida = var_cantidad_leida_llegar - var_suma_cantidad
            'var_suma_cantidad = var_cantidad_leida_llegar
            'var_año = 2005
            'var_inserta = False
            'lbl_cantidad_leida = Format(CDbl(lbl_cantidad_leida) + var_cantidad_leida, "###,###,##0.0000")
            'var_inserta = TB_TRASPASOS_INSERTA.Anadir(var_empresa, var_unidad_origen, var_almacen_origen, var_movimiento_salida, CInt(var_numero_salida), txt_codigo, 0, Round(var_cantidad_leida, 4), var_costo, var_precio, "0", var_almacen_origen, var_año)
            'Set list_item = lv_traspasosentradas.ListItems.Add(, , Trim(txt_codigo))
            'list_item.SubItems(1) = var_descripcion_articulo
            'list_item.SubItems(2) = 0
            'list_item.SubItems(3) = Format(Round(var_cantidad_leida, 4), "###,###,##0.0000")
            'list_item.SubItems(4) = var_costo
            'list_item.SubItems(5) = Format(Round(CDbl(list_item.SubItems(2)) - CDbl(list_item.SubItems(3)), 4))
            'list_item.SubItems(6) = 2005
            'var_renglon = lv_traspasosentradas.ListItems.Count
            'Call ilumina_grid
         End If

      txt_codigo.SetFocus
   End If
End Sub


Private Sub txt_movimiento_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_movimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "SELECT dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_AGENTES.VCHA_AGE_CLAVE_ORACLE FROM dbo.TB_AGENTES INNER JOIN dbo.TB_ALMACENES ON dbo.TB_AGENTES.VCHA_ALM_ALMACEN_ID = dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID WHERE (dbo.TB_AGENTES.VCHA_AGE_CLAVE_ORACLE IS NOT NULL) AND (LEN(dbo.TB_AGENTES.VCHA_AGE_CLAVE_ORACLE) > 0) order by VCHA_age_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
      var_tipo_lista = 1
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

Private Sub txt_movimiento_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      txt_nombre_movimiento.SetFocus
   End If
   If KeyAscii = 27 Then
      Me.frm_numero_traspaso.Visible = False
   End If
End Sub

Private Sub txt_movimiento_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_movimiento) <> "" Then
      rs.Open "SELECT dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_AGENTES.VCHA_AGE_CLAVE_ORACLE FROM dbo.TB_AGENTES INNER JOIN dbo.TB_ALMACENES ON dbo.TB_AGENTES.VCHA_ALM_ALMACEN_ID = dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID WHERE (dbo.TB_AGENTES.VCHA_AGE_CLAVE_ORACLE IS NOT NULL) AND (LEN(dbo.TB_AGENTES.VCHA_AGE_CLAVE_ORACLE) > 0) and dbo.TB_AGENTES.vcha_age_agente_id = '" + Me.txt_movimiento + "'", cnn, adOpenDynamic, adLockBatchOptimistic
      If Not rs.EOF Then
         txt_nombre_movimiento = rs!VCHA_AGE_NOMBRE
         txt_movimiento.Enabled = False
         txt_numero_traspaso = ""
         var_movimiento_salida = "TTS"
      Else
         MsgBox "Clave de agente incorrecta", vbOKOnly, "ATENCION"
         txt_movimiento = ""
         txt_nombre_movimiento = ""
         txt_numero_traspaso = ""
      End If
      rs.Close
   End If
End Sub

Private Sub txt_nombre_movimiento_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_nombre_movimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   If txt_movimiento.Enabled = True Then
      If KeyCode = 116 Then
         lv_lista.ListItems.Clear
         rs.Open "SELECT dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_AGENTES.VCHA_AGE_CLAVE_ORACLE FROM dbo.TB_AGENTES INNER JOIN dbo.TB_ALMACENES ON dbo.TB_AGENTES.VCHA_ALM_ALMACEN_ID = dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID WHERE (dbo.TB_AGENTES.VCHA_AGE_CLAVE_ORACLE IS NOT NULL) AND (LEN(dbo.TB_AGENTES.VCHA_AGE_CLAVE_ORACLE) > 0) order by VCHA_age_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
               list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
               rs.MoveNext
         Wend
         rs.Close
         lbl_lista = "AGENTES"
         var_tipo_lista = 1
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
   End If
End Sub

Private Sub txt_nombre_movimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txt_numero_traspaso.Enabled = True Then
         txt_numero_traspaso.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_numero_traspaso.Visible = False
   End If
End Sub

Private Sub txt_nombre_movimiento_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_numero_traspaso_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 27 Then
      frm_numero_traspaso.Visible = False
   End If
   If KeyAscii = 13 Then
      If Trim(txt_numero_traspaso) <> "" Then
         rs.Open "select vcha_alm_almacen_id from tb_agentes where vcha_age_agente_id = '" + Me.txt_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_clave_almacen_agente = IIf(IsNull(rs!VCHA_ALM_ALMACEN_ID), "", rs!VCHA_ALM_ALMACEN_ID)
         Else
            var_clave_almacen_agente = ""
         End If
         rs.Close
         rs.Open "select * from tb_encabezado_movimientos where VCHA_EMO_MOVIMIENTO_ORIGEN = '" + var_clave_movimiento + "' and INTE_EMO_NUMERO_ORIGEN = " + txt_numero_traspaso + " and vcha_emo_almacen_origen = '" + var_clave_almacen_agente + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            MsgBox "El traspaso ya fue cargado en el movimiento número " + Str(rs!INTE_EMO_NUMERO), vbOKOnly, "ATENCION"
            rs.Close
         Else
            rs.Close
            rs.Open "select * from tb_encabezado_movimientos where inte_emo_numero = " + txt_numero_traspaso + " and vcha_mov_movimiento_id = '" + var_movimiento_salida + "' and vcha_emo_almacen_origen = '" + var_clave_almacen_agente + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_almacen_destino_tem = rs!VCHA_EMO_ALMACEN_DESTINO
               var_almacen_origen_tem = rs!vcha_emo_almacen_origen
               var_posible = 1
               If var_tipo_permiso = 1 Then
                  rsaux.Open "select * from tb_permisos_movimientos where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_per_almacen_1 = '" + var_almacen_destino_tem + "'", cnn, adOpenDynamic, adLockBatchOptimistic
                  If rsaux.EOF Then
                     var_posible = 0
                  End If
                  rsaux.Close
                  rsaux.Open "select * from tb_permisos_movimientos where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_per_almacen_2 = '" + var_almacen_origen_tem + "'", cnn, adOpenDynamic, adLockBatchOptimistic
                  If rsaux.EOF Then
                     var_posible = 0
                  End If
                  rsaux.Close
               End If
               If var_posible = 1 Then
                  var_numero_salida = Val(txt_numero_traspaso)
                  txt_folio_enviado = var_numero_salida
                  var_almacen_Destino = rs!VCHA_EMO_ALMACEN_DESTINO
                  var_almacen_origen = rs!vcha_emo_almacen_origen
                  lv_traspasosentradas.ListItems.Clear
                  var_primera_vez = True
                  rsaux.Open "select * from tb_ALMACENES where vcha_alm_almacen_id = '" + var_almacen_Destino + "'", cnn, adOpenDynamic, adLockOptimistic
                  txt_almacen_destino = rsaux(2).Value
                  txt_almacen_destino.Text = rsaux(3).Value
                  txt_almacen_destino.Enabled = False
                  rsaux.Close
                  rsaux.Open "select * from tb_ALMACENES where vcha_alm_almacen_id = '" + var_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
                  txt_almacen_origen = rsaux(2).Value
                  txt_almacen_origen.Text = rsaux(3).Value
                  txt_almacen_origen.Enabled = False
                  var_unidad_origen = IIf(IsNull(rsaux!vcha_uor_unidad_id), "", rsaux!vcha_uor_unidad_id)
                  rsaux.Close
                  lbl_cantidad_leida = "0"
                  rsaux.Open "select * from tb_TRASPASOS where vcha_emp_empresa_id = '" + var_empresa + "' and inte_TRA_numero = " + txt_numero_traspaso + " and vcha_mov_movimiento_id = '" + var_movimiento_salida + "' order by vcha_art_articulo_id, inte_tra_año", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     While Not rsaux.EOF
                        rsaux2.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rsaux!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           Set list_item = lv_traspasosentradas.ListItems.Add(, , rsaux!vcha_Art_articulo_id)
                           list_item.SubItems(1) = IIf(IsNull(rsaux2(1).Value), "", rsaux2(1).Value)
                           list_item.SubItems(2) = IIf(IsNull(rsaux!floa_tra_Cantidad), 0, rsaux!floa_tra_Cantidad)
                           list_item.SubItems(3) = IIf(IsNull(rsaux!FLOA_TRA_CANTIDAD_RECIBIDA), 0, rsaux!FLOA_TRA_CANTIDAD_RECIBIDA)
                           list_item.SubItems(4) = IIf(IsNull(rsaux!floa_Tra_Costo), 0, rsaux!floa_Tra_Costo)
                           list_item.SubItems(6) = IIf(IsNull(rsaux!INTE_TRA_AÑO), 0, rsaux!INTE_TRA_AÑO)
                           lbl_cantidad_leida = Format(CDbl(lbl_cantidad_leida) + IIf(IsNull(rsaux!FLOA_TRA_CANTIDAD_RECIBIDA), 0, rsaux!FLOA_TRA_CANTIDAD_RECIBIDA), "###,###,##0.0000")
                           rsaux2.Close
                           rsaux.MoveNext:
                        End If
                     Wend
                     txt_codigo.Enabled = True
                     txt_codigo.SetFocus
                  Else
                     MsgBox "El traspaso no a sido enviado o no a sido impreso", vbOKOnly, "ATENCION"
                  End If
                  rsaux.Close
               Else
                  MsgBox "No esta autorizado para modificar este movimiento", vbOKOnly, "ATENCION"
               End If
            Else
               rsaux.Open "select vcha_alm_almacen_id, vcha_age_clave_oracle from tb_agentes where vcha_Age_Agente_id = '" + Me.txt_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
               rsaux2.Open "select * from tb_traspasos where vcha_emp_empresa_id = '" + rsaux!VCHA_AGE_CLAVE_ORACLE + "' and nume_tra_folio = " + Me.txt_numero_traspaso + " and vcha_Can_canal_id = '01' and vcha_tra_empresa_id_destino = '0500'", cnntraspasos_tiendas, adOpenDynamic, adLockOptimistic
               If Not rsaux2.EOF Then
                  rsaux3.Open "select * from tb_folios_movimientos where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_mov_movimiento_id = 'TTS'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux3.EOF Then
                     var_numero_folio_entrada = IIf(IsNull(rsaux3!INTE_EMO_NUMERO), 0, rsaux3!INTE_EMO_NUMERO) + 1
                     rsaux4.Open "update tb_folios_movimientos set inte_emo_numero = inte_emo_numero + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_mov_movimiento_id = 'TTS'", cnn, adOpenDynamic, adLockOptimistic
                  Else
                     var_numero_folio_entrada = 1
                     rsaux4.Open "insert into tb_folios_movimientos (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_mov_movimiento_id, inte_emo_numero) values ('" + var_empresa + "','12','TTS',1)", cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux3.Close
                  var_cadena = "insert into tb_encabezado_movimientos  (vcha_emp_Empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, dtim_emo_fecha, inte_emo_numero, inte_emo_numero_origen, vcha_emo_almacen_origen, vcha_emo_almacen_destino, char_emo_estatus, vcha_aud_usuario, vcha_aud_maquina, floa_emo_descuento_1, floa_emo_Descuento_2, floa_emo_descuento_3, vcha_mon_moneda_id, floa_emo_tipo_cambio, char_emo_tipo_cliente_proveedor, vcha_emo_afectacion, vcha_emo_movimiento_origen)"
                  var_cadena = var_cadena + " values ('" + var_empresa + "','12','" + rsaux!VCHA_ALM_ALMACEN_ID + "','TTS',getdate()," + CStr(var_numero_folio_entrada) + "," + Me.txt_numero_traspaso + ",'" + rsaux!VCHA_ALM_ALMACEN_ID + "','8','I','" + var_clave_usuario_global + "','" + fun_NombrePc + "',0,0,0,'1',1,'A','TS','TTS')"
                  rsaux3.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux2.EOF
                        rsaux4.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + rsaux2!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux4.EOF Then
                           var_codigo = IIf(IsNull(rsaux4!vcha_Art_articulo_id), "", rsaux4!vcha_Art_articulo_id)
                        Else
                           var_codigo = ""
                        End If
                        rsaux4.Close
                        rsaux4.Open "select mone_art_precio_base from tb_articulos where vcha_Art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux4.EOF Then
                           var_precio = IIf(IsNull(rsaux4(0).Value), 0, rsaux4(0).Value)
                        Else
                           var_precio = 0
                        End If
                        rsaux4.Close
                        var_cadena = "insert into tb_traspasos (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_tra_numero, vcha_Art_articulo_id, floa_tra_Cantidad, floa_tra_cantidad_recibida, floa_Tra_costo, floa_Tra_precio, floa_tra_Descuento, vcha_tra_almacen_origen, inte_tra_año,VCHA_TRA_CODIGO_EXTERNO, INTE_TRA_NUMERO_EXTERNO, VCHA_TRA_EMPRESA_EXTERNA)"
                        var_cadena = var_cadena + " values ('" + var_empresa + "',   '12',  '" + rsaux!VCHA_ALM_ALMACEN_ID + "','TTS'," + CStr(var_numero_folio_entrada) + ",'" + var_codigo + "'," + CStr(rsaux2!nume_tra_cantidad_origen) + ",0," + CStr(rsaux2!nume_Tra_costo) + "," + CStr(var_precio) + ",0,'" + rsaux!VCHA_ALM_ALMACEN_ID + "',2005,'" + rsaux2!vcha_Art_articulo_id + "'," + Me.txt_numero_traspaso + ",'" + rsaux!VCHA_AGE_CLAVE_ORACLE + "')"
                        rsaux4.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        rsaux2.MoveNext
                  Wend
               
                  var_numero_salida = Val(var_numero_folio_entrada)
                  txt_folio_enviado = var_numero_salida
                  var_almacen_Destino = 8
                  var_almacen_origen = rsaux!VCHA_ALM_ALMACEN_ID
                  lv_traspasosentradas.ListItems.Clear
                  var_primera_vez = True
                  rsaux5.Open "select * from tb_ALMACENES where vcha_alm_almacen_id = '" + var_almacen_Destino + "'", cnn, adOpenDynamic, adLockOptimistic
                  txt_almacen_destino = rsaux5(2).Value
                  txt_almacen_destino.Text = rsaux5(3).Value
                  txt_almacen_destino.Enabled = False
                  rsaux5.Close
                  rsaux5.Open "select * from tb_ALMACENES where vcha_alm_almacen_id = '" + var_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
                  txt_almacen_origen = rsaux5(2).Value
                  txt_almacen_origen.Text = rsaux5(3).Value
                  txt_almacen_origen.Enabled = False
                  var_unidad_origen = "12"
                  rsaux5.Close
                  lbl_cantidad_leida = "0"
                  rsaux5.Open "select * from tb_TRASPASOS where vcha_emp_empresa_id = '" + var_empresa + "' and inte_TRA_numero = " + CStr(var_numero_folio_entrada) + " and vcha_mov_movimiento_id = 'TTS' order by vcha_art_articulo_id, inte_tra_año", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux5.EOF Then
                     While Not rsaux5.EOF
                        rsaux7.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rsaux5!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux5.EOF Then
                           Set list_item = lv_traspasosentradas.ListItems.Add(, , rsaux5!vcha_Art_articulo_id)
                           list_item.SubItems(1) = IIf(IsNull(rsaux7(1).Value), "", rsaux7(1).Value)
                           list_item.SubItems(2) = IIf(IsNull(rsaux5!floa_tra_Cantidad), 0, rsaux5!floa_tra_Cantidad)
                           list_item.SubItems(3) = IIf(IsNull(rsaux5!FLOA_TRA_CANTIDAD_RECIBIDA), 0, rsaux5!FLOA_TRA_CANTIDAD_RECIBIDA)
                           list_item.SubItems(4) = IIf(IsNull(rsaux5!floa_Tra_Costo), 0, rsaux5!floa_Tra_Costo)
                           list_item.SubItems(5) = IIf(IsNull(rsaux5!floa_tra_Cantidad), 0, rsaux5!floa_tra_Cantidad) - IIf(IsNull(rsaux5!FLOA_TRA_CANTIDAD_RECIBIDA), 0, rsaux5!FLOA_TRA_CANTIDAD_RECIBIDA)
                           list_item.SubItems(6) = IIf(IsNull(rsaux5!INTE_TRA_AÑO), 0, rsaux5!INTE_TRA_AÑO)
                           lbl_cantidad_leida = Format(CDbl(lbl_cantidad_leida) + IIf(IsNull(rsaux5!FLOA_TRA_CANTIDAD_RECIBIDA), 0, rsaux5!FLOA_TRA_CANTIDAD_RECIBIDA), "###,###,##0.0000")
                           rsaux7.Close
                           rsaux5.MoveNext:
                        End If
                     Wend
                     txt_codigo.Enabled = True
                     txt_codigo.SetFocus
                  End If
                  rsaux5.Close
               
               
               
               
               
               Else
                 MsgBox "El número de movimiento no existe ", vbOKOnly, "ATENCION"
               End If
               rsaux2.Close
               rsaux.Close
            End If
            rs.Close
         End If
      End If
      frm_numero_traspaso.Visible = False
   End If
End Sub

