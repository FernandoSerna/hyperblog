VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsalidas_proveedores_intercompañias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_facturas 
      Height          =   2295
      Left            =   1755
      TabIndex        =   38
      Top             =   3555
      Width           =   4905
      Begin MSComctlLib.ListView lv_facturas 
         Height          =   1830
         Left            =   30
         TabIndex        =   39
         Top             =   390
         Width           =   4785
         _ExtentX        =   8440
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
      Begin VB.Label lbl_lista_facturas 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   40
         Top             =   120
         Width           =   4785
      End
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   975
      TabIndex        =   16
      Top             =   540
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   30
         TabIndex        =   17
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
         TabIndex        =   18
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.TextBox txt_foco 
      Height          =   285
      Left            =   8205
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2745
      Width           =   1125
   End
   Begin VB.Frame Frame3 
      Height          =   1230
      Index           =   0
      Left            =   5580
      TabIndex        =   32
      Top             =   1035
      Width           =   1965
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
         Left            =   45
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   540
         Width           =   1860
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
         Width           =   1890
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   0
      Left            =   90
      TabIndex        =   31
      Top             =   495
      Width           =   7455
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmsalidas_proveedores_intercompañias.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   645
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   420
      Picture         =   "frmsalidas_proveedores_intercompañias.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Buscar Movimiento"
      Top             =   645
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   750
      Picture         =   "frmsalidas_proveedores_intercompañias.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   645
      Width           =   330
   End
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1080
      Picture         =   "frmsalidas_proveedores_intercompañias.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cancelar Movimiento"
      Top             =   645
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7185
      Picture         =   "frmsalidas_proveedores_intercompañias.frx":0408
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   645
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   4890
      Left            =   135
      TabIndex        =   19
      Top             =   2280
      Width           =   7425
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   2940
         TabIndex        =   20
         Top             =   2025
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar 
            Height          =   330
            Left            =   60
            TabIndex        =   21
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
            TabIndex        =   22
            Top             =   15
            Width           =   2895
         End
      End
      Begin VB.Frame frm_presione_F5 
         Height          =   615
         Left            =   2445
         TabIndex        =   41
         Top             =   990
         Width           =   4905
         Begin VB.Label Label5 
            Caption         =   "Presione F5 para ver las facturas posibles"
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
            Height          =   300
            Left            =   120
            TabIndex        =   42
            Top             =   225
            Width           =   4665
         End
      End
      Begin VB.TextBox txt_factura 
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
         Left            =   4260
         TabIndex        =   11
         Top             =   555
         Visible         =   0   'False
         Width           =   1230
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
         Left            =   6360
         TabIndex        =   12
         Top             =   555
         Width           =   930
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
         Left            =   720
         TabIndex        =   10
         Top             =   495
         Width           =   2685
      End
      Begin MSComctlLib.ListView lv_entradas 
         Height          =   3720
         Left            =   45
         TabIndex        =   23
         Top             =   1110
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   6562
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
            Text            =   "Código"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   8617
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Factura:"
         Height          =   195
         Left            =   3630
         TabIndex        =   37
         Top             =   675
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lbl_cantidad 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   5655
         TabIndex        =   26
         Top             =   675
         Width           =   675
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Lectura de Artículos"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   0
         Left            =   30
         TabIndex        =   25
         Top             =   120
         Width           =   7350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   675
         Width           =   540
      End
   End
   Begin VB.Frame frm_busqueda 
      Height          =   960
      Left            =   600
      TabIndex        =   2
      Top             =   1035
      Width           =   3135
      Begin VB.TextBox txt_busqueda_folio 
         Height          =   315
         Left            =   195
         TabIndex        =   14
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
         TabIndex        =   15
         Top             =   120
         Width           =   3060
      End
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   585
      Top             =   750
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
            Picture         =   "frmsalidas_proveedores_intercompañias.frx":0A42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_proveedores_intercompañias.frx":131C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_proveedores_intercompañias.frx":1BF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_proveedores_intercompañias.frx":2192
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_proveedores_intercompañias.frx":2A6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_proveedores_intercompañias.frx":3348
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_proveedores_intercompañias.frx":3C22
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_proveedores_intercompañias.frx":3D34
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_proveedores_intercompañias.frx":3E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_proveedores_intercompañias.frx":3F58
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_proveedores_intercompañias.frx":406A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_proveedores_intercompañias.frx":417C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   1230
      Index           =   1
      Left            =   135
      TabIndex        =   27
      Top             =   1035
      Width           =   5415
      Begin VB.TextBox txt_almacen 
         Height          =   315
         Left            =   945
         TabIndex        =   6
         Top             =   450
         Width           =   1140
      End
      Begin VB.TextBox txt_nombre_almacen 
         Height          =   315
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   450
         Width           =   3255
      End
      Begin VB.TextBox txt_nombre_proveedor 
         Height          =   315
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   825
         Width           =   3255
      End
      Begin VB.TextBox txt_proveedor 
         Height          =   315
         Left            =   945
         TabIndex        =   8
         Top             =   810
         Width           =   1140
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Origen:"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   30
         Top             =   510
         Width           =   510
      End
      Begin VB.Label label 
         BackColor       =   &H8000000D&
         Caption         =   " Datos del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   29
         Top             =   120
         Width           =   5325
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Planta:"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   28
         Top             =   885
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   90
      TabIndex        =   35
      Top             =   900
      Width           =   7455
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
      Left            =   120
      TabIndex        =   36
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frmsalidas_proveedores_intercompañias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_kanban As String
Dim var_almacen_Destino As String
Dim var_primera_vez As Boolean
Dim var_numero_folio As Double
Dim var_cantidad_leida As Double
Dim var_costo As Double
Dim var_precio As Double
Dim var_descripcion_articulo As String
Dim var_estatus_movimiento As String
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_numero_causa As Integer
Dim var_elimina As Boolean
Dim var_ventana As Integer
Dim var_clave_moneda As String
Dim var_año As Integer
Dim var_suma_cantidad As Double
Dim var_cantidad_llegar As Double
Dim var_cantidad As Double
Dim var_renglon As Double
Dim var_tipo_lista As Integer

Sub ilumina_grid()
   var_n = lv_entradas.ListItems.Count
   For var_i = 1 To var_n
       If var_i = var_renglon Then
          lv_entradas.ListItems.Item(var_i).Bold = True
          lv_entradas.ListItems.Item(var_i).ListSubItems(1).Bold = True
          lv_entradas.ListItems.Item(var_i).ListSubItems(2).Bold = True
          lv_entradas.ListItems.Item(var_i).ForeColor = &H8000&
          lv_entradas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H8000&
          lv_entradas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H8000&
       Else
          lv_entradas.ListItems.Item(var_i).Bold = False
          lv_entradas.ListItems.Item(var_i).ListSubItems(1).Bold = False
          lv_entradas.ListItems.Item(var_i).ListSubItems(2).Bold = False
          lv_entradas.ListItems.Item(var_i).ForeColor = &H80000012
          lv_entradas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H80000012
          lv_entradas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H80000012
       End If
   Next var_i
   If var_renglon > 0 Then
      lv_entradas.ListItems.Item(var_renglon).Selected = True
      lv_entradas.selectedItem.EnsureVisible
   End If
   lv_entradas.Refresh
End Sub




Private Sub cmd_buscar_Click()
   var_ventana = 1
   frm_busqueda.Visible = True
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
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
   If var_numero_folio > 0 Then
      If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
         Set reporte = appl.OpenReport(App.Path + "\rep_MOVIMIENTOS_SALIDA_Proveedores_intercompañias.rpt")
         reporte.RecordSelectionFormula = "{VW_MOVIMIENTOS_SALIDA.VCHA_ALM_ALMACEN_ID} = '" + var_almacen_Destino + "' AND {VW_MOVIMIENTOS_SALIDA.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' AND {VW_MOVIMIENTOS_SALIDA.INTE_EMO_NUMERO} = " + Str(var_numero_folio)
         frmvistasprevias.cr.ReportSource = reporte
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Reporte de Movimientos"
         frmvistasprevias.Show 1
         Set reporte = Nothing
         rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
      Else
         var_posible_Cantidad = 1
         If var_empresa = "18" Or var_empresa = "31" Then
            Cadena = "select * from tb_temporal_salidas with (nolock) where vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and floa_Sal_cantidad > 0"
            rsaux10.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux10.EOF
                  rsaux9.Open "select * from tb_existencias where vcha_Alm_almacen_id = '" + var_almacen_Destino + "' and vcha_Art_Articulo_id = '" + rsaux10!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux9.EOF Then
                     var_cantidad = IIf(IsNull(rsaux9!floa_Exi_Cantidad_disponible), 0, rsaux9!floa_Exi_Cantidad_disponible)
                     If var_empresa = "18" Then
                        If rsaux10!vcha_Art_Articulo_id = "360010000002" Or rsaux10!vcha_Art_Articulo_id = "360020000009" Or rsaux10!vcha_Art_Articulo_id = "900000000003" Or rsaux10!vcha_Art_Articulo_id = "911110000005" Then
                           var_cantidad = Round(IIf(IsNull(rsaux10!floa_Sal_Cantidad), 0, rsaux10!floa_Sal_Cantidad), 4) + 1
                        End If
                     End If
                     
                     If Round(var_cantidad, 4) < Round(IIf(IsNull(rsaux10!floa_Sal_Cantidad), 0, rsaux10!floa_Sal_Cantidad), 4) Then
                        var_posible_Cantidad = 0
                        If var_cadena_articulos = "" Then
                           rsaux8.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + rsaux10!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux8.EOF Then
                              var_nombre_articulo = IIf(IsNull(rsaux8!vcha_Art_nombre_español), "", rsaux8!vcha_Art_nombre_español)
                           Else
                              var_nombre_articulo = ""
                           End If
                           rsaux8.Close
                           var_cadena_articulos = rsaux10!vcha_Art_Articulo_id + " " + var_nombre_articulo + " Existen [" + CStr(var_cantidad) + "] y salen [" + CStr(rsaux10!floa_Sal_Cantidad) + "]"
                        Else
                           rsaux8.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + rsaux10!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux8.EOF Then
                              var_nombre_articulo = IIf(IsNull(rsaux8!vcha_Art_nombre_español), "", rsaux8!vcha_Art_nombre_español)
                           Else
                              var_nombre_articulo = ""
                           End If
                           rsaux8.Close
                           var_cadena_articulos = var_cadena_articulos + ", " + rsaux10!vcha_Art_Articulo_id + " " + var_nombre_articulo + " Existen [" + CStr(var_cantidad) + "] y salen [" + CStr(rsaux10!floa_Sal_Cantidad) + "]"
                        End If
                     
                     
                     End If
                  Else
                     If var_cadena_articulos = "" Then
                        rsaux8.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + rsaux10!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux8.EOF Then
                           var_nombre_articulo = IIf(IsNull(rsaux8!vcha_Art_nombre_español), "", rsaux8!vcha_Art_nombre_español)
                        Else
                           var_nombre_articulo = ""
                        End If
                        rsaux8.Close
                        var_cadena_articulos = rsaux10!vcha_Art_Articulo_id + " " + var_nombre_articulo
                     Else
                        rsaux8.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + rsaux10!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux8.EOF Then
                           var_nombre_articulo = IIf(IsNull(rsaux8!vcha_Art_nombre_español), "", rsaux8!vcha_Art_nombre_español)
                        Else
                           var_nombre_articulo = ""
                        End If
                        rsaux8.Close
                        var_cadena_articulos = var_cadena_articulos + ", " + rsaux10!vcha_Art_Articulo_id + " " + var_nombre_articulo
                     End If
                     var_posible_Cantidad = 0
                  End If
                  rsaux9.Close
                  rsaux10.MoveNext
            Wend
            rsaux10.Close
         End If
         If var_empresa = "31" Then
            var_posible_Cantidad = 1
         End If
         If var_posible_Cantidad = 1 Then
            var_si = MsgBox("¿Se va a imprimir el movimiento?", vbOKCancel, "ATENCION")
            If var_si = 1 Then
               var_posible_cerrar_KANBAN = True
               If var_posible_kanban = 1 Then
                  Set TB_PROC_KANBANS_EN_MOVIMIENTO = New TB_PROC_KANBANS_EN_MOVIMIENTO
                  var_inserta = TB_PROC_KANBANS_EN_MOVIMIENTO.Anadir(Me.txt_almacen, var_clave_movimiento, CDbl(Me.txt_folio), "", "")
                  If var_kanban_exito = "N" Then
                     var_posible_cerrar_KANBAN = False
                  End If
               Else
                  var_posible_cerrar_KANBAN = True
               End If
               If var_posible_cerrar_KANBAN = True Then
                  cnn.BeginTrans
                  Cadena = "select * from tb_temporal_salidas with (nolock) where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio)
                  rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                  While Not rs.EOF
                        rsaux.Open "insert into tb_salidas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_sal_numero, vcha_art_articulo_id, floa_sal_cantidad, floa_sal_costo, floa_sal_precio, inte_sal_año) values ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!VCHA_UOR_UNIDAD_ID + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_SAL_NUMERO) + ", '" + rs!vcha_Art_Articulo_id + "', " + CStr(rs!floa_Sal_Cantidad) + ", " + CStr(rs!floa_Sal_costo) + " , " + CStr(rs!floa_Sal_precio) + ", 2005)", cnn, adOpenDynamic, adLockOptimistic
                        rs.MoveNext
                  Wend
                  rs.Close
                  var_estatus_movimiento = "I"
                  
   'inicio poliza
   
                  rsaux10.Open "select VCHA_eMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_sal_NUMERO,SUM(FLOA_sal_CANTIDAD), SUM(FLOA_sal_CANTIDAD * FLOA_sal_COSTO) AS COSTO, SUM(FLOA_sal_CANTIDAD * FLOA_sal_PRECIO) AS PRECIO from tb_salidas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_sal_numero = " + CStr(var_numero_folio) + " GROUP BY VCHA_eMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_sal_NUMERO", cnn, adOpenDynamic, adLockOptimistic
                  'rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  var_x = 0
               
                  If var_x = 1 Then
                  rsaux11.Open "select * from tb_generador_polizas where empresa_id = '" + var_empresa + "'", cnnoracle, adOpenDynamic, adLockOptimistic
                  While Not rsaux11.EOF
                        var_tipo_poliza = rsaux11!tipo
                        var_origen_poliza = rsaux11!Origen
                        var_categoria_poliza = rsaux11!categoria
                        var_moneda_poliza = rsaux11!moneda
                        var_segmento1_poliza = rsaux11!segmento1
                        var_segmento2_poliza = rsaux11!segmento2
                        var_segmento3_poliza = rsaux11!segmento3
                        var_segmento4_poliza = rsaux11!segmento4
                        var_segmento5_poliza = rsaux11!segmento5
                        var_segmento6_poliza = rsaux11!segmento6
                        var_segmento7_poliza = rsaux11!segmento7
                        var_juego_libros_poliza = rsaux11!juego_libros
                        var_descripcion_poliza = rsaux11!descripcion
                        var_cargo_poliza = rsaux11!cargo
                        var_abono_poliza = rsaux11!abono
                        var_precio = rsaux11!Precio
                        If var_precio = 1 Then
                           var_importe_precio = rsaux10!Precio
                        Else
                           var_importe_precio = rsaux10!Costo
                        End If
                        var_cadena = "InsERT INTO IN_TB_POLIZAS_INT (STATUS, SET_OF_BOOKS_ID, USER_JE_SOURCE_NAME, USER_JE_CATEGORY_NAME, ACCOUNTING_DATE, CURRENCY_CODE, DATE_CREATED, ACTUAL_FLAG,  SEGMENT1, SEGMENT2, SEGMENT3, SEGMENT4, SEGMENT5, SEGMENT6, SEGMENT7, ENTERED_DR, ENTERED_CR, ACCOUNTED_DR, ACCOUNTED_CR, GROUP_ID, REFERENCE4, REFERENCE5, REFERENCe10, REFERENCE1, REFERENCE2, CREATED_BY)"
                        If var_cargo_poliza = 1 Then
                        '   var_cadena = var_cadena + " VALUES ('NEW', " + CStr(var_juego_libros_poliza) + ",'" + var_origen_poliza + "','" + var_categoria_poliza + "',TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'" + var_moneda_poliza + "',TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'A','" + var_segmento1_poliza + "','" + var_segmento2_poliza + "','" + var_segmento3_poliza + "','" + var_segmento4_poliza + "','" + var_segmento5_poliza + "','" + var_segmento6_poliza + "','" + var_segmento7_poliza + "'," + CStr(var_importe_precio) + ",0," + CStr(var_importe_precio) + ",0,1,'FACTURAS INTERCOMPAÑIA " + Me.txt_archivo + "','FACTURA NUM: " + Me.txt_archivo + "','" + var_descripcion_poliza + "','POLIZA FACTURAS INTERCOMPAÑIA','POLIZA FACTURAS INTERCOMPAÑIA',1143)"
                        Else
                        '   var_cadena = var_cadena + " VALUES ('NEW', " + CStr(var_juego_libros_poliza) + ",'" + var_origen_poliza + "','" + var_categoria_poliza + "',TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'" + var_moneda_poliza + "',TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'A','" + var_segmento1_poliza + "','" + var_segmento2_poliza + "','" + var_segmento3_poliza + "','" + var_segmento4_poliza + "','" + var_segmento5_poliza + "','" + var_segmento6_poliza + "','" + var_segmento7_poliza + "',0," + CStr(var_importe_precio) + ",0," + CStr(var_importe_precio) + ",1,'FACTURAS INTERCOMPAÑIA " + Me.txt_archivo + "','FACTURA NUM: " + Me.txt_archivo + "','" + var_descripcion_poliza + "','POLIZA FACTURAS INTERCOMPAÑIA','POLIZA FACTURAS INTERCOMPAÑIA',1143)"
                        End If
                        'MsgBox var_cadena
                        rsaux9.Open var_cadena, cnnoracle, adOpenDynamic, adLockOptimistic
                        rsaux11.MoveNext
                  Wend
                  rsaux11.Close
                  End If
                  'rsaux11.Open "select sq_id_facturas.nextval from dual", cnnoracle, adOpenDynamic, adLockOptimistic
                  'var_consecutivo = rsaux11(0).Value
                  'rsaux11.Close
               
               
                  'var_proveedor_oracle = Me.txt_proveedor
                 
                  
                  'rsaux11.Open "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE VCHA_UOR_unidad_id = '" + var_proveedor_oracle + "'", cnn, adOpenDynamic, adLockOptimistic
                  'var_empresa_emite = rsaux11!VCHA_EMP_EMPRESA_ID
                  'var_proveedor_oracle_2 = rsaux11!vcha_uor_proveedor_oracle
                  'rsaux11.Close
               
                  'rsaux11.Open "select * from tb_empresas_cruzadas_oracle where vcha_emp_Empresa_emite = '" + var_empresa_emite + "' and vcha_emp_empresa_recibe = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                  'var_unidad_oracle = rsaux11!vcha_emp_organizacion
                  'rsaux11.Close
                  
                  'rsaux11.Open "SELECT vendor_site_id FROM po_vendor_sites_all@perpvia.vianney.com.mx Where vendor_id = '" + var_proveedor_oracle_2 + "'  AND vendor_site_id in (4070,2803,1125,1200,1202,1126,1519,1327,1520,3545,1127,1674,4383,2668,1529,1326,2669,2925,3755,1268,1324,3768,2392,9016,3332) AND ORG_ID = '" + var_unidad_oracle + "'", cnnoracle, adOpenDynamic, adLockOptimistic
                  'var_clave_proveedor_oracle = rsaux11!vendor_site_id
                  'rsaux11.Close
                  
                  'rsaux11.Open "select sum(FLOA_SAL_CANTIDAD * floa_SAL_costo) FROM  tb_temporal_SALIDAS where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_SAL_numero = " + Me.txt_folio, cnn, adOpenDynamic, adLockOptimistic
                  'rsaux11.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  'var_importe_total = rsaux11(0).Value
                  'rsaux11.Close
                  'var_importe_total = var_importe_total * 1.16
                  
                  'var_cadena = "insert into IN_TB_FACTURAS_INT (INVOICE_ID,INVOICE_NUM,INVOICE_TYPE_LOOKUP_CODE,VENDOR_ID,VENDOR_SITE_ID,INVOICE_AMOUNT,INVOICE_CURRENCY_CODE,EXCHANGE_RATE_TYPE,EXCHANGE_DATE,EXCHANGE_RATE,Description,Source,GL_DATE,INVOICE_DATE,ORG_ID) values (" + CStr(var_consecutivo) + ",'" + Me.txt_archivo + "','STANDARD'," + CStr(var_proveedor_oracle_2) + "," + CStr(var_clave_proveedor_oracle) + "," + CStr(var_importe_total) + ",'MXP',null,null,null,'FACTURA DE RECEPCION NUM: " + Me.txt_archivo + "','FACTURA INTERCOMPAÑIAS',TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),TO_DATE('" + CStr(Date) + "','DD/MM/YYYY')," + var_unidad_oracle + ")"
                  ''MsgBox var_cadena
                  'rsaux11.Open var_cadena, cnnoracle, adOpenDynamic, adLockOptimistic
                  
                  'rsaux11.Open "select sq_id_lineas_factura.nextval from dual", cnnoracle, adOpenDynamic, adLockOptimistic
                  'var_consecutivo_linea = rsaux11(0).Value
                  'rsaux11.Close
                  'var_subimporte = var_importe_total / 1.16
                  'var_importe_iva = var_importe_total - var_subimporte
                  'rsaux11.Open "select amount_includes_tax_flag, vat_code from po_vendor_sites_all@perpvia.vianney.com.mx Where vendor_id = " + CStr(var_proveedor_oracle_2) + " and vendor_site_id = " + CStr(var_clave_proveedor_oracle) + " and org_id = " + CStr(var_unidad_oracle), cnnoracle, adOpenDynamic, adLockOptimistic
                  'amount_includes_tax_flag = rsaux11!amount_includes_tax_flag
                  'TAX_CODE = IIf(IsNull(rsaux11!vat_code), 0, rsaux11!vat_code)
                  'rsaux11.Close
                  'rsaux.Open "select awt_group_id from po_vendors@perpvia.vianney.com.mx Where vendor_id = " + CStr(var_proveedor_oracle), cnnoracle, adOpenDynamic, adLockOptimistic
                  'AWT_GROUP_ID = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
                  'rsaux.Close
                  ''MsgBox CStr(AWT_GROUP_ID)
                  'If TAX_CODE = 0 Then
                  '   If AWT_GROUP_ID = 0 Then
                  '      var_cadena = "insert into IN_TB_LINEAS_FACT_INT (INVOICE_ID,LINE_NUMBER,LINE_TYPE_LOOKUP_CODE,AMOUNT,ACCOUNTING_DATE,Description, AMOUNT_INCLUDES_TAX_FLAG, TAX_CODE,DIST_CODE_COMBINATION_ID,AWT_GROUP_ID,RECEIPT_NUMBER,RECEIPT_LINE_NUMBER,PO_NUMBER,PO_LINE_NUMBER,MATCH_OPTION,ORG_ID) values (" + CStr(var_consecutivo) + "," + CStr(var_consecutivo_linea) + ",'ITEM'," + CStr(var_subimporte) + ", TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'RECEPCION','" + CStr(amount_includes_tax_flag) + "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL," + var_unidad_oracle + ")"
                  '   Else
                  '      var_cadena = "insert into IN_TB_LINEAS_FACT_INT (INVOICE_ID,LINE_NUMBER,LINE_TYPE_LOOKUP_CODE,AMOUNT,ACCOUNTING_DATE,Description, AMOUNT_INCLUDES_TAX_FLAG, TAX_CODE,DIST_CODE_COMBINATION_ID,AWT_GROUP_ID,RECEIPT_NUMBER,RECEIPT_LINE_NUMBER,PO_NUMBER,PO_LINE_NUMBER,MATCH_OPTION,ORG_ID) values (" + CStr(var_consecutivo) + "," + CStr(var_consecutivo_linea) + ",'ITEM'," + CStr(var_subimporte) + ", TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'RECEPCION','" + CStr(amount_includes_tax_flag) + "',NULL,NULL," + CStr(AWT_GROUP_ID) + ",NULL,NULL,NULL,NULL,NULL," + var_unidad_oracle + ")"
                  '   End If
                  'Else
                  '   If AWT_GROUP_ID = 0 Then
                  '      var_cadena = "insert into IN_TB_LINEAS_FACT_INT (INVOICE_ID,LINE_NUMBER,LINE_TYPE_LOOKUP_CODE,AMOUNT,ACCOUNTING_DATE,Description, AMOUNT_INCLUDES_TAX_FLAG, TAX_CODE,DIST_CODE_COMBINATION_ID,AWT_GROUP_ID,RECEIPT_NUMBER,RECEIPT_LINE_NUMBER,PO_NUMBER,PO_LINE_NUMBER,MATCH_OPTION,ORG_ID) values (" + CStr(var_consecutivo) + "," + CStr(var_consecutivo_linea) + ",'ITEM'," + CStr(var_subimporte) + ", TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'RECEPCION','" + CStr(amount_includes_tax_flag) + "'," + CStr(TAX_CODE) + ",NULL,NULL,NULL,NULL,NULL,NULL,NULL," + var_unidad_oracle + ")"
                  '   Else
                  '      var_cadena = "insert into IN_TB_LINEAS_FACT_INT (INVOICE_ID,LINE_NUMBER,LINE_TYPE_LOOKUP_CODE,AMOUNT,ACCOUNTING_DATE,Description, AMOUNT_INCLUDES_TAX_FLAG, TAX_CODE,DIST_CODE_COMBINATION_ID,AWT_GROUP_ID,RECEIPT_NUMBER,RECEIPT_LINE_NUMBER,PO_NUMBER,PO_LINE_NUMBER,MATCH_OPTION,ORG_ID) values (" + CStr(var_consecutivo) + "," + CStr(var_consecutivo_linea) + ",'ITEM'," + CStr(var_subimporte) + ", TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'RECEPCION','" + CStr(amount_includes_tax_flag) + "'," + CStr(TAX_CODE) + ",NULL," + CStr(AWT_GROUP_ID) + ",NULL,NULL,NULL,NULL,NULL," + var_unidad_oracle + ")"
                  '   End If
                  'End If
                  ''MsgBox var_cadena
                  'rsaux11.Open var_cadena, cnnoracle, adOpenDynamic, adLockOptimistic
                  'rsaux11.Open "select sq_id_lineas_factura.nextval from dual", cnnoracle, adOpenDynamic, adLockOptimistic
                  'var_consecutivo_linea = rsaux11(0).Value
                  'rsaux11.Close
                     
                  'var_cadena = "insert into IN_TB_LINEAS_FACT_INT (INVOICE_ID,LINE_NUMBER,LINE_TYPE_LOOKUP_CODE,AMOUNT,ACCOUNTING_DATE,Description,AMOUNT_INCLUDES_TAX_FLAG,TAX_CODE,DIST_CODE_COMBINATION_ID,AWT_GROUP_ID,RECEIPT_NUMBER,RECEIPT_LINE_NUMBER,PO_NUMBER,PO_LINE_NUMBER,MATCH_OPTION,ORG_ID) values (" + CStr(var_consecutivo) + "," + CStr(var_consecutivo_linea) + ",'TAX'," + CStr(var_importe_iva) + ", TO_DATE('" + CStr(Date) + "','DD/MM/YYYY'),'IMPUESTO',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL," + var_unidad_oracle + ")"
                  'rsaux11.Open var_cadena, cnnoracle, adOpenDynamic, adLockOptimistic
   'fin poliza
                  var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "", Now, 1)
                  var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "I", Now, 1)
                  cnn.CommitTrans
                  Set reporte = appl.OpenReport(App.Path + "\rep_MOVIMIENTOS_SALIDA_Proveedores_intercompañias.rpt")
                  reporte.RecordSelectionFormula = "{VW_MOVIMIENTOS_SALIDA.VCHA_ALM_ALMACEN_ID} = '" + var_almacen_Destino + "' AND {VW_MOVIMIENTOS_SALIDA.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' AND {VW_MOVIMIENTOS_SALIDA.INTE_EMO_NUMERO} = " + Str(var_numero_folio)
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Reporte de Movimientos"
                  frmvistasprevias.Show 1
                  Set reporte = Nothing
                  rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                  txt_codigo.Enabled = False
                  txt_foco.Enabled = False
               Else
                  MsgBox "No se pudo cerrar el movimiento Kanban"
               End If
            End If
         Else
            MsgBox "El movimiento no se puede imprimir ya que las existencias de los siguientes artículos exceden a la cantidad disponible en el almacen " + var_cadena_articulos
         End If
      End If
   Else
      MsgBox "No se a seleccionado ningún movimiento", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   If var_numero_folio > 0 Then
      rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
   End If
   txt_nombre_proveedor = ""
   txt_almacen = ""
   txt_nombre_almacen = ""
   var_ventana = 0
   txt_codigo.Enabled = False
   var_primera_vez = True
   frm_busqueda.Visible = False
   lv_entradas.ListItems.Clear
   var_numero_folio = 0
   txt_folio = ""
   txt_codigo = ""
   var_estatus_movimiento = ""
   txt_proveedor = ""
   txt_proveedor.Enabled = False
   txt_almacen.Enabled = True
   txt_almacen.SetFocus
End Sub

Private Sub cmd_salir_Click()
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

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 And var_ventana = 0 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   var_posible_kanban = 0
   var_cadena_seguridad = ""
   frm_lista.Visible = False
   Top = 0
   Left = 1500
   rs.Open "select * from tb_monedas where inte_mon_moneda_local = 1", cnn, adOpenDynamic, adLockOptimistic
   var_clave_moneda = ""
   If Not rs.EOF Then
      var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
   End If
   rs.Close
   var_ventana = 0
   var_estatus_movimiento = ""
   frm_busqueda.Visible = False
   frm_eliminar.Visible = False
   lbl_cantidad.Visible = False
   txt_cantidad.Visible = False
   txt_proveedor.Enabled = False
   txt_codigo.Enabled = False
   var_primera_vez = True
   var_cantidad_leida = 1#
   frm_facturas.Visible = False
   Me.frm_presione_F5.Visible = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
   If var_numero_folio > 0 Then
      rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
   End If
   Call activa_forma(var_activa_forma_salidas_proveedor)
End Sub

Private Sub lv_entradas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
         MsgBox "Imporsible modificar el movimiento", vbOKOnly, "ATENCION"
      Else
         If var_causa_devolucion = True Then
            rs.Open "select * from tb_causas_devolucion order by vcha_cde_nombre", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_elimina = True
               lv_causas_devolucion.ListItems.Clear
               While Not rs.EOF
                  Set list_item = lv_causas_devolucion.ListItems.Add(, , rs!INTE_CDE_CAUSA_ID)
                  list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
                  rs.MoveNext
               Wend
               rs.Close
               lv_causas_devolucion.SetFocus
            Else
               var_elimina = False
               var_ventana = 1
               frm_eliminar.Visible = True
               txt_cantidad_eliminar.SetFocus
            End If
         Else
            var_elimina = False
            var_ventana = 1
            frm_eliminar.Visible = True
            txt_cantidad_eliminar.SetFocus
         End If
      End If
   End If
End Sub

Private Sub Toolbar1_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)

End Sub

Private Sub lv_facturas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_facturas_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    If Me.lv_facturas.ListItems.Count > 0 Then
       Me.txt_factura = lv_facturas.selectedItem
       Me.txt_factura.SetFocus
    End If
 End If
 If KeyAscii = 27 Then
    Me.txt_factura.SetFocus
 End If
End Sub

Private Sub lv_facturas_LostFocus()
   Me.frm_facturas.Visible = False
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim var_n As Integer
      If var_tipo_lista = 1 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_almacen = lv_lista.selectedItem
            txt_nombre_almacen = lv_lista.selectedItem.SubItems(1)
         Else
            txt_almacen = ""
            txt_nombre_almacen = ""
         End If
         txt_almacen.SetFocus
         frm_lista.Visible = False
      End If
      If var_tipo_lista = 2 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_proveedor = lv_lista.selectedItem
            txt_nombre_proveedor = lv_lista.selectedItem.SubItems(1)
         Else
            txt_proveedor = ""
            txt_nombre_proveedor = ""
         End If
         txt_proveedor.SetFocus
         frm_lista.Visible = False
      End If
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub txt_almacen_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_almacen_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      If var_tipo_permiso = 1 Then
         rs.Open "select * from vw_almacen_permiso_1 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      Else
         rs.Open "select * from vw_movimientos_almacenes where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      End If
      'rs.Open "select distinct vcha_cli_nombre from vw_establecimientos where vcha_esb_establecimiento_id = '" + txt_establecimiento + "' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ALM_ALMACEN_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Almacenes"
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

Private Sub txt_almacen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      txt_nombre_almacen.SetFocus
   End If
End Sub

Private Sub txt_almacen_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_almacen) <> "" Then
      If var_tipo_permiso = 1 Then
         rs.Open "select * from vw_almacen_permiso_1 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_alm_almacen_id = '" + txt_almacen + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      Else
         rs.Open "select * from vw_movimientos_almacenes where vcha_mov_movimiento_id = '" + var_clave_movimiento + "'  AND VCHA_ALM_ALMACEN_ID = '" + txt_almacen + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      End If
      If Not rs.EOF Then
         txt_almacen.Enabled = False
         txt_nombre_almacen = rs!VCHA_ALM_NOMBRE
         var_almacen_Destino = txt_almacen
         txt_proveedor.Enabled = True
      Else
         MsgBox "Clave de almacen Incorrecta", vbOKOnly, "ATENCION"
         txt_almacen = ""
         txt_nombre_almacen = ""
         txt_proveedor.Enabled = False
      End If
      If rs.State = 1 Then
         rs.Close
      End If
   End If
End Sub

Private Sub txt_busqueda_folio_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If Trim(txt_busqueda_folio) <> "" Then
         If var_numero_folio = CDbl(txt_busqueda_folio) Then
            rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 where inte_emo_numero = " + txt_busqueda_folio + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
         End If
         rs.Open "select * from tb_encabezado_movimientos where inte_emo_numero = " + txt_busqueda_folio + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            If var_numero_folio > 0 Then
               rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
            End If
            var_movimiento_bloqueado = IIf(IsNull(rs!INTE_EMO_BLOQUEADO), 0, rs!INTE_EMO_BLOQUEADO)
            If var_movimiento_bloqueado = 0 Then
               var_almacen_destino_tem = rs!VCHA_ALM_ALMACEN_ID
               var_posible = 1
               If var_tipo_permiso = 1 Then
                  rsaux.Open "select * from tb_permisos_movimientos where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_per_almacen_1 = '" + var_almacen_destino_tem + "'", cnn, adOpenDynamic, adLockBatchOptimistic
                  If rsaux.EOF Then
                     var_posible = 0
                  End If
                  rsaux.Close
               End If
               If var_posible = 1 Then
                  var_estatus_movimiento = rs!char_Emo_estatus
                  var_almacen_Destino = rs!VCHA_ALM_ALMACEN_ID
                  txt_almacen = rs!VCHA_ALM_ALMACEN_ID
                  txt_almacen.Enabled = False
                  txt_proveedor = rs!VCHA_PRO_PROVEEDOR_ID
                  rsaux.Open "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE VCHA_UOR_UNIDAD_ID = '" + txt_proveedor + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     txt_nombre_proveedor = IIf(IsNull(rsaux!VCHA_UOR_NOMBRE), "", rsaux!VCHA_UOR_NOMBRE)
                  Else
                     txt_nombre_proveedor = ""
                  End If
                  rsaux.Close
                  txt_proveedor.Enabled = False
                  lv_entradas.ListItems.Clear
                  var_primera_vez = False
                  var_numero_folio = rs!INTE_EMO_NUMERO
                  txt_folio = var_numero_folio
                  rsaux.Open "select * from tb_ALMACENES where vcha_alm_almacen_id = '" + var_almacen_Destino + "'", cnn, adOpenDynamic, adLockOptimistic
                  txt_almacen_destino = rsaux(3).Value
                  txt_nombre_almacen.Text = rsaux(3).Value
                  rsaux.Close
                  rsaux.Open "select * from tb_temporal_salidas with (nolock) where inte_SAL_numero = " + txt_busqueda_folio + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     While Not rsaux.EOF
                        rsaux2.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rsaux!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           Set list_item = lv_entradas.ListItems.Add(, , rsaux!vcha_Art_Articulo_id)
                           list_item.SubItems(1) = IIf(IsNull(rsaux2(1).Value), "", rsaux2(1).Value)
                           list_item.SubItems(2) = IIf(IsNull(rsaux!floa_Sal_Cantidad), "", rsaux!floa_Sal_Cantidad)
                           rsaux2.Close
                           rsaux.MoveNext:
                        End If
                     Wend
                  End If
                  rsaux.Close
                  rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 1 WHERE inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
                  If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
                     txt_codigo.Enabled = False
                     txt_cantidad.Visible = False
                     lbl_cantidad.Visible = False
                     txt_foco.Enabled = False
                  Else
                     txt_foco.Enabled = False
                     txt_codigo.Enabled = True
                     txt_cantidad.Visible = False
                     lbl_cantidad.Visible = False
                  End If
               Else
                  MsgBox "No esta autorizado para modificar este movimiento", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "El movimiento esta siendo utilizado por otro usuario", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El número de movimiento no existe ", vbOKOnly, "ATENCION"
         End If
         rs.Close
      End If
      var_ventana = 0
      frm_busqueda.Visible = False
   End If
End Sub

Private Sub txt_cantidad_eliminar_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If var_posible_kanban = 1 Then
         If IsNumeric(Me.txt_cantidad_eliminar) Then
            Set TB_CANCELAR_RES_FUERA_DE_KANBAN = New TB_CANCELAR_RES_FUERA_DE_KANBAN
            var_inserta = TB_CANCELAR_RES_FUERA_DE_KANBAN.Anadir(Me.txt_almacen, var_clave_movimiento, var_numero_folio, Me.lv_entradas.selectedItem, CDbl(Me.txt_cantidad_eliminar), "", "")
            var_kanban_es_un_kanban = var_kanban_es_un_kanban
            var_kanban_almacen_id = var_kanban_almacen_id
            var_kanban_articulo_id = var_kanban_articulo_id
            var_kanban_exito = var_kanban_exito
            var_kanban_mensaje = var_kanban_mensaje
            If var_kanban_exito = "S" Then
               var_posible = True
            Else
               frmmensaje.lbl_mensaje = var_kanban_mensaje
               frmmensaje.Show 1
               var_posible = False
            End If
         Else
            Set TB_ES_UN_KANBAN = New TB_ES_UN_KANBAN
            var_kanban = Me.txt_codigo
            var_inserta = TB_ES_UN_KANBAN.Anadir(Me.txt_cantidad_eliminar, "", "", "", "", "")
            var_kanban_es_un_kanban = var_kanban_es_un_kanban
            var_kanban_almacen_id = var_kanban_almacen_id
            var_kanban_articulo_id = var_kanban_articulo_id
            var_kanban_exito = var_kanban_exito
            var_kanban_mensaje = var_kanban_mensaje
            If var_kanban_es_un_kanban = "S" Then
               If Me.lv_entradas.selectedItem = var_kanban_articulo_id Then
                  Set TB_CANCELAR_RESERVACION_KANBAN = New TB_CANCELAR_RESERVACION_KANBAN
                  var_kanban = Me.txt_codigo
                  var_inserta = TB_CANCELAR_RESERVACION_KANBAN.Anadir(Me.txt_almacen, var_clave_movimiento, var_numero_folio, Me.txt_cantidad_eliminar, "", "")
                  var_kanban_es_un_kanban = var_kanban_es_un_kanban
                  var_kanban_almacen_id = var_kanban_almacen_id
                  var_kanban_articulo_id = var_kanban_articulo_id
                  var_kanban_exito = var_kanban_exito
                  var_kanban_mensaje = var_kanban_mensaje
                  If var_kanban_exito = "S" Then
                     var_posible = True
                  Else
                     frmmensaje.lbl_mensaje = var_kanban_mensaje
                     frmmensaje.Show 1
                     var_posible = False
                  End If
               Else
                  frmmensaje.lbl_mensaje = "El codigo de kanban no corresponde al del artículo seleccionado"
                  frmmensaje.Show 1
                  var_posible = False
               End If
            Else
               frmmensaje.lbl_mensaje = var_kanban_mensaje
               frmmensaje.Show 1
               var_posible = False
            End If
         End If
      Else
         var_posible = True
      End If
         
      If var_posible = True Then
         If var_posible_kanban = 1 Then
            If Not IsNumeric(txt_cantidad_eliminar) Then
               Me.txt_cantidad_eliminar = 1
            End If
         End If
         If IsNumeric(txt_cantidad_eliminar) Then
            Dim var_posible_eliminar As Boolean
            var_cantidad_eliminar = Val(txt_cantidad_eliminar)
            var_posible_eliminar = True
            If CDbl(txt_cantidad_eliminar) <= CDbl(Me.lv_entradas.selectedItem.SubItems(2)) Then
               If var_posible_eliminar = True Then
                  var_inserta = False
                  rsaux.Open "UPDATE TB_TEMPORAL_SALIDAS SET FLOA_SAL_CANTIDAD = ISNULL(FLOA_SAL_CANTIDAD,0) - " + txt_cantidad_eliminar + " WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_SAL_NUMERO = " + CStr(var_numero_folio) + " AND VCHA_ART_ARTICULO_ID= '" + lv_entradas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                  'var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_destino, var_clave_movimiento, var_numero_folio, lv_entradas.SelectedItem, 0 - Val(txt_cantidad_eliminar))
                  lv_entradas.selectedItem.SubItems(2) = lv_entradas.selectedItem.SubItems(2) - Val(txt_cantidad_eliminar)
                  var_renglon = lv_entradas.selectedItem.Index
                  Call ilumina_grid
               Else
                  MsgBox "La cantidad a eliminar supera a la cantidad asignada a la causa de devolución seleccionada", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "La cantidad supera a la cantidad del movimiento", vbOKOnly, "ATENCION"
               Me.frm_eliminar.Visible = False
            End If
            var_ventana = 0
            frm_eliminar.Visible = False
            txt_codigo.SetFocus
         Else
            MsgBox "Cantidad incorrecta", vbOKOnly, "ATENCION"
         End If
      End If
   End If
   If KeyAscii = 27 Then
      var_ventana = 0
      frm_eliminar.Visible = False
      txt_codigo.SetFocus
   End If
End Sub

Private Sub txt_cantidad_GotFocus()
   txt_cantidad = 1#
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

Private Sub txt_codigo_Change()
   Me.txt_factura = ""
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
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   txt_codigo = Trim(txt_codigo)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If var_posible_kanban = 1 Then
         Set TB_ES_UN_KANBAN = New TB_ES_UN_KANBAN
         var_kanban = Me.txt_codigo
         var_inserta = TB_ES_UN_KANBAN.Anadir(Me.txt_codigo, "", "", "", "", "")
         var_kanban_es_un_kanban = var_kanban_es_un_kanban
         var_kanban_almacen_id = var_kanban_almacen_id
         var_kanban_articulo_id = var_kanban_articulo_id
         var_kanban_exito = var_kanban_exito
         var_kanban_mensaje = var_kanban_mensaje
         
         If var_kanban_es_un_kanban = "S" Then
            Me.txt_codigo = var_kanban_articulo_id
         Else
            var_kanban_almacen_id = Me.txt_almacen
         End If
         If var_kanban_almacen_id = Me.txt_almacen Then
            If var_empresa = 16 Then
               If Len(Me.txt_codigo) = 6 Then
                  Me.txt_codigo = Mid(Me.txt_codigo, 1, 3) + "-" + Mid(Me.txt_codigo, 4, 3) + "-"
               Else
                  If Len(Me.txt_codigo) = 7 Then
                     Me.txt_codigo = Mid(Me.txt_codigo, 1, 3) + "-" + Mid(Me.txt_codigo, 4, 3) + "-" + Mid(Me.txt_codigo, 7, 1)
                  End If
               End If
            End If
            
            var_verificador = True
            If Len(Trim(txt_codigo)) = 12 Then
               Call calcula_verificador(Trim(txt_codigo))
            End If
            If var_verificador = True Then
               var_es_caja = False
               If Trim(txt_codigo) <> "" Then
                  If Left(Trim(txt_codigo), 1) = "C" Then
                     x = Mid(txt_codigo, 2, 6)
                     var_embarque_caja = 0
                     If IsNumeric(x) Then
                        var_embarque_caja = CDbl(x)
                        If var_embarque_caja = var_numero_embarque Then
                           var_es_caja = True
                        Else
                           frmmensaje.lbl_mensaje = "La caja pertenece a otro embarque"
                           frmmensaje.Show 1
                           'MsgBox "La caja pertenece al embarque " + CStr(var_embarque_caja)
                           var_es_caja = False
                        End If
                     Else
                        frmmensaje.lbl_mensaje = "Caja incorrecta"
                        frmmensaje.Show 1
                        'MsgBox "Caja incorrecta", vbOKOnly, "ATENCION"
                        var_es_caja = False
                     End If
                  Else
                     var_es_caja = False
                  End If
                  If var_es_caja = True Then
                     txt_foco.Enabled = True
                     txt_foco.SetFocus
                  Else
                     var_caja = Left(txt_codigo, 6)
                     If var_caja = "000005" Or var_caja = "000010" Or var_caja = "000015" Or var_caja = "000020" Or var_caja = "000001" Or var_caja = "000002" Or var_caja = "000003" Or var_caja = "000004" Or var_caja = "000006" Or var_caja = "000007" Or var_caja = "000008" Or var_caja = "000009" Or var_caja = "000010" Or var_caja = "000011" Or var_caja = "000012" Or var_caja = "000013" Or var_caja = "000014" Or var_caja = "000015" Or var_caja = "000016" Or var_caja = "000017" Or var_caja = "000018" Or var_caja = "000019" Or var_caja = "000020" Or var_caja = "000021" Or var_caja = "000022" Or var_caja = "000023" Or var_caja = "000024" Or var_caja = "000025" Or var_caja = "000026" Or var_caja = "000027" Or var_caja = "000028" Or var_caja = "000029" Or var_caja = "000030" Then
                        var_cantidad_caja = CInt(var_caja)
                        txt_codigo = Mid(txt_codigo, 7, 5)
                     End If
                     rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        var_descripcion_articulo = IIf(IsNull(rs!vcha_Art_nombre_español), "", rs!vcha_Art_nombre_español)
                        If IsNull(rs(43).Value) Then
                           var_recontable = 0
                        Else
                           var_recontable = rs(43).Value
                        End If
                        rs.Close
                        If var_recontable = 1 Then
                           var_cantidad_leida = 1#
                           lbl_cantidad.Visible = True
                           txt_cantidad.Visible = True
                           'Me.txt_factura.SetFocus
                           Me.txt_cantidad.SetFocus
                        Else
                           var_cantidad_leida = 1#
                           txt_foco.Enabled = True
                           'Me.txt_factura.SetFocus
                           Me.txt_foco.SetFocus
                        End If
                     Else
                        rs.Close
                        rs.Open "select * from tb_equivalencias where VCHA_EQU_CODIGO_EQUIVALENTE = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           txt_codigo = rs(0).Value
                           rs.Close
                           rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rs.EOF Then
                              var_descripcion_articulo = IIf(IsNull(rs!vcha_Art_nombre_español), "", rs!vcha_Art_nombre_español)
                              If var_cantidad_caja = 0 Then
                                 If IsNull(rs(43).Value) Then
                                    var_recontable = 0
                                 Else
                                    var_recontable = rs(43).Value
                                 End If
                              Else
                                 var_recontable = 0
                              End If
                              rs.Close
                              If var_recontable = 1 Then
                                 var_cantidad_leida = 1#
                                 lbl_cantidad.Visible = True
                                 txt_cantidad.Visible = True
                                 'Me.txt_factura.SetFocus
                                 Me.txt_cantidad.SetFocus
                              Else
                                 If var_cantidad_caja = 0 Then
                                    var_cantidad_leida = 1#
                                 Else
                                    var_cantidad_leida = var_cantidad_caja
                                 End If
                                 txt_foco.Enabled = True
                                 'Me.txt_factura.SetFocus
                                 Me.txt_foco.SetFocus
                              End If
                           Else
                              frmmensaje.lbl_mensaje = "El artículo no existe"
                              frmmensaje.Show 1
                              'MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
                              txt_codigo = ""
                           End If
                        Else
                           frmmensaje.lbl_mensaje = "El artículo no existe"
                           frmmensaje.Show 1
                          'MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
                           txt_codigo = ""
                           rs.Close
                        End If
                     End If
                  End If
               End If
            Else
               txt_codigo = ""
               frmmensaje.lbl_mensaje = "Error en Código"
               frmmensaje.Show 1
               ' MsgBox "Error en Código", vbOKOnly, "ATENCION"
            End If
         Else
            txt_codigo = ""
            frmmensaje.lbl_mensaje = "El almacén del Kanban no pertenece al almacén del movimiento"
            frmmensaje.Show 1
         End If
      Else
''' FIN KANBAN
         var_verificador = True
         If Len(Trim(txt_codigo)) = 12 Then
            Call calcula_verificador(Trim(txt_codigo))
         End If
         If var_verificador = True Then
            var_caja = Left(txt_codigo, 6)
            If var_caja = "000005" Or var_caja = "000010" Or var_caja = "000015" Or var_caja = "000020" Then
               var_cantidad_caja = CInt(var_caja)
               txt_codigo = Mid(txt_codigo, 7, 5)
            End If
            var_costo = 0
            var_precio = 0
            If Trim(txt_codigo) <> "" Then
               rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  If var_clave_movimiento = "SA" Then
                     var_recontable = 1
                  Else
                     If IsNull(rs(43).Value) Then
                        var_recontable = 0
                     Else
                        var_recontable = rs(43).Value
                     End If
                  End If
                  var_descripcion_articulo = rs(1).Value
                  var_costo = rs(3).Value
                  var_precio = rs(2).Value
                  rs.Close
                  If var_recontable = 1 Then
                     var_cantidad_leida = 1#
                     lbl_cantidad.Visible = True
                     txt_cantidad.Visible = True
                     'Me.txt_factura.SetFocus
                     Me.txt_cantidad.SetFocus
                  Else
                     var_cantidad_leida = 1#
                     txt_foco.Enabled = True
                     'Me.txt_factura.SetFocus
                     Me.txt_foco.SetFocus
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
                           If var_clave_movimiento = "SA" Then
                              var_recontable = 1
                           Else
                              If IsNull(rs(43).Value) Then
                                 var_recontable = 0
                              Else
                                 var_recontable = rs(43).Value
                              End If
                           End If
                        Else
                           var_recontable = 0
                        End If
                        var_descripcion_articulo = rs(1).Value
                        var_costo = rs(3).Value
                        var_precio = rs(2).Value
                        rs.Close
                        If var_recontable = 1 Then
                           var_cantidad_leida = 1#
                           lbl_cantidad.Visible = True
                           txt_cantidad.Visible = True
                           'Me.txt_factura.SetFocus
                           Me.txt_cantidad.SetFocus
                        Else
                           If var_cantidad_caja = 0 Then
                              var_cantidad_leida = 1#
                           Else
                              var_cantidad_leida = var_cantidad_caja
                           End If
                           txt_foco.Enabled = True
                           'Me.txt_factura.SetFocus
                           Me.txt_foco.SetFocus
                        End If
                     Else
                        Me.txt_codigo = ""
                        frmmensaje.lbl_mensaje = "El artículo no existe"
                        frmmensaje.Show 1
                        'MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
                     End If
                  Else
                     Me.txt_codigo = ""
                     frmmensaje.lbl_mensaje = "El artículo no existe"
                     frmmensaje.Show 1
                     'MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
                     rs.Close
                  End If
               End If
            Else
            End If
         Else
            txt_codigo = ""
            frmmensaje.lbl_mensaje = "Error en Código"
            frmmensaje.Show 1
            'MsgBox "Error en Código", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub

Private Sub txt_factura_GotFocus()
   Me.frm_presione_F5.Visible = True
End Sub

Private Sub txt_factura_KeyDown(KeyCode As Integer, Shift As Integer)
   If Me.txt_factura.Enabled = True Then
      If KeyCode = 116 Then
         rsaux11.Open "select vcha_com_Referencia, max(dtim_com_fecha) as fecha from tb_archivo_comparacion where vcha_Art_Articulo_id = '" + Me.txt_codigo + "' and VCHA_COM_PROVEEDOR = '" + Me.txt_proveedor + "' and vcha_emp_empresa_id = '" + var_empresa + "' group by vcha_Com_referencia", cnn, adOpenDynamic, adLockOptimistic
         lv_facturas.ListItems.Clear
         While Not rsaux11.EOF
               Set list_item = lv_facturas.ListItems.Add(, , rsaux11!vcha_com_Referencia)
               list_item.SubItems(1) = IIf(IsNull(rsaux11!fecha), "", rsaux11!fecha)
               rsaux11.MoveNext
         Wend
         rsaux11.Close
         Me.lbl_lista_facturas = "FACTURAS"
         var_tipo_lista = 2
         Dim var_n As Integer
         var_n = lv_facturas.ListItems.Count
         If var_n > 6 Then
            lv_facturas.ColumnHeaders(2).Width = 4270.71
         Else
            lv_facturas.ColumnHeaders(2).Width = 4499.71
         End If
         frm_facturas.Visible = True
         lv_facturas.SetFocus
      End If
   End If
 
End Sub

Private Sub txt_factura_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_factura_LostFocus()
   Me.frm_presione_F5.Visible = False
End Sub

Private Sub txt_foco_GotFocus()
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_SALIDAS_INSERTA = New TB_TEMPORAL_SALIDAS_INSERTA
   Set TB_TEMPORAL_SALIDAS_MODIFICA = New TB_TEMPORAL_SALIDAS_MODIFICA
   Set TB_ENCABEZADO_MOVIMIENTOS_I = New TB_ENCABEZADO_MOVIMIENTOS_I
   Dim var_inserta As Boolean
   
   If Trim(txt_codigo.Text) <> "" Then
      var_pase_existencias = 1
      If var_empresa = "18" Or var_empresa = "31" Then
         If var_numero_folio = 0 Or Trim(Me.txt_folio) = "" Then
            var_cantidad_temporal = 0
         Else
            rsaux.Open "select isnull(floa_sal_cantidad,0) from tb_Temporal_salidas where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_sal_numero = " + CStr(var_numero_folio) + " and vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               var_cantidad_temporal = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
            Else
               var_cantidad_temporal = 0
            End If
            rsaux.Close
         End If
         'MsgBox CStr(var_cantidad_temporal)
         If rsaux.State = 1 Then
            rsaux.Close
         End If
         rsaux.Open "select floa_exi_Cantidad_disponible from tb_existencias where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            var_cantidad_Existencias = IIf(IsNull(rsaux!floa_Exi_Cantidad_disponible), 0, rsaux!floa_Exi_Cantidad_disponible)
         Else
            var_cantidad_Existencias = 0
         End If
         rsaux.Close
         var_cantidad_posible = var_cantidad_Existencias - (var_cantidad_temporal + var_cantidad_leida)
         If var_cantidad_posible < 0 Then
            var_pase_existencias = 0
         End If
      End If
      If var_empresa = "18" Then
         If Me.txt_codigo = "360010000002" Or Me.txt_codigo = "360020000009" Or Me.txt_codigo = "900000000003" Or Me.txt_codigo = "911110000005" Then
            var_pase_existencias = 1
         End If
      End If
      If var_empresa = "31" Then
         var_pase_existencias = 1
      End If
      If var_pase_existencias = 1 Then
         rsaux11.Open "select max(vcha_com_Referencia) as vcha_com_Referencia  from tb_archivo_comparacion where vcha_Art_Articulo_id = '" + Me.txt_codigo + "' and VCHA_COM_PROVEEDOR = '" + Me.txt_proveedor + "' and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux11.EOF Then
            Me.txt_factura = IIf(IsNull(rsaux11!vcha_com_Referencia), "", rsaux11!vcha_com_Referencia)
         End If
         rsaux11.Close
         If Me.txt_factura <> "" Then
            rsaux11.Open "select * from tb_archivo_comparacion where vcha_com_referencia = '" + Me.txt_factura + "' and vcha_Art_Articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux11.EOF Then
               var_costo = IIf(IsNull(rsaux11!FLOA_COM_COSTO), 0, rsaux11!FLOA_COM_COSTO)
               var_precio = IIf(IsNull(rsaux11!FLOA_COM_PRECIO), 0, rsaux11!FLOA_COM_PRECIO)
               bandera_suma = False
               If var_primera_vez = True Then
                  var_inserta = False
                  var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, Now, var_numero_folio, 0, "", txt_proveedor, "", var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, 0, "", "", "", "B", "", "", 0, 0, 0, var_clave_moneda, 0)
                  var_numero_folio = var_numero_folio_regreso
                   txt_folio = var_numero_folio
                  var_primera_vez = False
               End If
               If var_posible_kanban = 1 Then
                  Set TB_RESERVAR_FUERA_DE_KANBAN = New TB_RESERVAR_FUERA_DE_KANBAN
                  Set TB_RESERVAR_KANBAN = New TB_RESERVAR_KANBAN
                  If var_kanban_es_un_kanban = "S" Then
                     var_inserta = TB_RESERVAR_KANBAN.Anadir(var_kanban, var_clave_movimiento, var_numero_folio, Me.txt_almacen, Me.txt_codigo, "", "")
                     If var_kanban_exito = "S" Then
                        var_posible_leido = 1
                     Else
                        var_posible_leido = 0
                     End If
                  Else
                     var_inserta = TB_RESERVAR_FUERA_DE_KANBAN.Anadir(var_numero_folio, var_clave_movimiento, Me.txt_almacen, Me.txt_codigo, "", "")
                     If var_kanban_exito = "S" Then
                        var_posible_leido = 1
                     Else
                        var_posible_leido = 0
                     End If
                  End If
               Else
                  var_kanban_mensaje = ""
                  var_posible_leido = 1
               End If
               If var_posible_leido = 1 Then
                  Cadena = "select * from tb_temporal_salidas with (nolock) where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "'"
                  rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_inserta = False
                     var_inserta = TB_TEMPORAL_SALIDAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida)
                     rs.Close
                     valor = Trim(txt_codigo)
                     Set itmfound = lv_entradas.findItem(valor, lvwText, , lvwPartial)
                     itmfound.EnsureVisible
                     itmfound.Selected = True
                     lv_entradas.selectedItem.SubItems(2) = lv_entradas.selectedItem.SubItems(2) + var_cantidad_leida
                     var_renglon = lv_entradas.selectedItem.Index
                     Call ilumina_grid
                  Else
                     var_inserta = False
                     rsaux.Open "INSERT INTO TB_TEMPORAL_SALIDAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO) VALUES ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", '" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo) + ", " + CStr(var_precio) + ")", cnn, adOpenDynamic, adLockOptimistic
                     'var_inserta = TB_TEMPORAL_SALIDAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_destino, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida, var_costo, var_precio, "0", "")
                     rs.Close
                     Set list_item = lv_entradas.ListItems.Add(, , Trim(txt_codigo))
                     list_item.SubItems(1) = var_descripcion_articulo
                     list_item.SubItems(2) = var_cantidad_leida
                     var_renglon = lv_entradas.ListItems.Count
                     Call ilumina_grid
                  End If
                  Me.txt_factura = ""
               Else
                  frmmensaje.lbl_mensaje = var_kanban_mensaje
                  frmmensaje.Show 1
                  txt_codigo = ""
                  Me.txt_factura = ""
               End If
            Else
               'frmmensaje.lbl_mensaje = "El artículo no viene en la factura indicada"
               'frmmensaje.Show 1
               'txt_codigo = ""
               'Me.txt_factura = ""
               rsaux10.Open "select * from tb_articulos where vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               var_costo = IIf(IsNull(rsaux10!mone_Art_costo_estandar), 0, rsaux10!mone_Art_costo_estandar)
               var_precio = IIf(IsNull(rsaux10!mone_Art_precio_base), 0, rsaux10!mone_Art_precio_base)
               rsaux10.Close
               bandera_suma = False
               If var_primera_vez = True Then
                  var_inserta = False
                  var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, Now, var_numero_folio, 0, "", txt_proveedor, "", var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, 0, "", "", "", "B", "", "", 0, 0, 0, var_clave_moneda, 0)
                  var_numero_folio = var_numero_folio_regreso
                  txt_folio = var_numero_folio
                  var_primera_vez = False
               End If
               If var_posible_kanban = 1 Then
                  Set TB_RESERVAR_FUERA_DE_KANBAN = New TB_RESERVAR_FUERA_DE_KANBAN
                  Set TB_RESERVAR_KANBAN = New TB_RESERVAR_KANBAN
                  If var_kanban_es_un_kanban = "S" Then
                     var_inserta = TB_RESERVAR_KANBAN.Anadir(var_kanban, var_clave_movimiento, var_numero_folio, Me.txt_almacen, Me.txt_codigo, "", "")
                     If var_kanban_exito = "S" Then
                        var_posible_leido = 1
                     Else
                        var_posible_leido = 0
                     End If
                  Else
                     var_inserta = TB_RESERVAR_FUERA_DE_KANBAN.Anadir(var_numero_folio, var_clave_movimiento, Me.txt_almacen, Me.txt_codigo, "", "")
                     If var_kanban_exito = "S" Then
                        var_posible_leido = 1
                     Else
                        var_posible_leido = 0
                     End If
                  End If
               Else
                  var_kanban_mensaje = ""
                  var_posible_leido = 1
               End If
               If var_posible_leido = 1 Then
                  Cadena = "select * from tb_temporal_salidas with (nolock) where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "'"
                  rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_inserta = False
                     var_inserta = TB_TEMPORAL_SALIDAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida)
                     rs.Close
                     valor = Trim(txt_codigo)
                     Set itmfound = lv_entradas.findItem(valor, lvwText, , lvwPartial)
                     itmfound.EnsureVisible
                     itmfound.Selected = True
                     lv_entradas.selectedItem.SubItems(2) = lv_entradas.selectedItem.SubItems(2) + var_cantidad_leida
                     var_renglon = lv_entradas.selectedItem.Index
                     Call ilumina_grid
                  Else
                     var_inserta = False
                     rsaux.Open "INSERT INTO TB_TEMPORAL_SALIDAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO) VALUES ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", '" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo) + ", " + CStr(var_precio) + ")", cnn, adOpenDynamic, adLockOptimistic
                     'var_inserta = TB_TEMPORAL_SALIDAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_destino, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida, var_costo, var_precio, "0", "")
                     rs.Close
                     Set list_item = lv_entradas.ListItems.Add(, , Trim(txt_codigo))
                     list_item.SubItems(1) = var_descripcion_articulo
                     list_item.SubItems(2) = var_cantidad_leida
                     var_renglon = lv_entradas.ListItems.Count
                     Call ilumina_grid
                  End If
                  Me.txt_factura = ""
               Else
                  frmmensaje.lbl_mensaje = var_kanban_mensaje
                  frmmensaje.Show 1
                  txt_codigo = ""
                  Me.txt_factura = ""
               End If
         
            
            End If
            rsaux11.Close
         Else
            rsaux10.Open "select * from tb_articulos where vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            var_costo = IIf(IsNull(rsaux10!mone_Art_costo_estandar), 0, rsaux10!mone_Art_costo_estandar)
            var_precio = IIf(IsNull(rsaux10!mone_Art_precio_base), 0, rsaux10!mone_Art_precio_base)
            rsaux10.Close
            bandera_suma = False
            If var_primera_vez = True Then
               var_inserta = False
               var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, Now, var_numero_folio, 0, "", txt_proveedor, "", var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, 0, "", "", "", "B", "", "", 0, 0, 0, var_clave_moneda, 0)
               var_numero_folio = var_numero_folio_regreso
               txt_folio = var_numero_folio
               var_primera_vez = False
            End If
            If var_posible_kanban = 1 Then
               Set TB_RESERVAR_FUERA_DE_KANBAN = New TB_RESERVAR_FUERA_DE_KANBAN
               Set TB_RESERVAR_KANBAN = New TB_RESERVAR_KANBAN
               If var_kanban_es_un_kanban = "S" Then
                  var_inserta = TB_RESERVAR_KANBAN.Anadir(var_kanban, var_clave_movimiento, var_numero_folio, Me.txt_almacen, Me.txt_codigo, "", "")
                  If var_kanban_exito = "S" Then
                     var_posible_leido = 1
                  Else
                     var_posible_leido = 0
                  End If
               Else
                  var_inserta = TB_RESERVAR_FUERA_DE_KANBAN.Anadir(var_numero_folio, var_clave_movimiento, Me.txt_almacen, Me.txt_codigo, "", "")
                  If var_kanban_exito = "S" Then
                     var_posible_leido = 1
                  Else
                     var_posible_leido = 0
                  End If
               End If
            Else
               var_kanban_mensaje = ""
               var_posible_leido = 1
            End If
            If var_posible_leido = 1 Then
               Cadena = "select * from tb_temporal_salidas with (nolock) where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "'"
               rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_inserta = False
                  var_inserta = TB_TEMPORAL_SALIDAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida)
                  rs.Close
                  valor = Trim(txt_codigo)
                  Set itmfound = lv_entradas.findItem(valor, lvwText, , lvwPartial)
                  itmfound.EnsureVisible
                  itmfound.Selected = True
                  lv_entradas.selectedItem.SubItems(2) = lv_entradas.selectedItem.SubItems(2) + var_cantidad_leida
                  var_renglon = lv_entradas.selectedItem.Index
                  Call ilumina_grid
               Else
                  var_inserta = False
                  rsaux.Open "INSERT INTO TB_TEMPORAL_SALIDAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO) VALUES ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", '" + txt_codigo + "', " + CStr(var_cantidad_leida) + ", " + CStr(var_costo) + ", " + CStr(var_precio) + ")", cnn, adOpenDynamic, adLockOptimistic
                  'var_inserta = TB_TEMPORAL_SALIDAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_destino, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida, var_costo, var_precio, "0", "")
                  rs.Close
                  Set list_item = lv_entradas.ListItems.Add(, , Trim(txt_codigo))
                  list_item.SubItems(1) = var_descripcion_articulo
                  list_item.SubItems(2) = var_cantidad_leida
                  var_renglon = lv_entradas.ListItems.Count
                  Call ilumina_grid
               End If
               Me.txt_factura = ""
            Else
               frmmensaje.lbl_mensaje = var_kanban_mensaje
               frmmensaje.Show 1
               txt_codigo = ""
               Me.txt_factura = ""
            End If
         
         
            'frmmensaje.lbl_mensaje = "No se indico una factura"
            'frmmensaje.Show 1
            'txt_codigo = ""
            'Me.txt_factura = ""
         End If
      Else
         Me.txt_codigo = ""
         frmmensaje.lbl_mensaje = "La cantidad excede a la cantidad en existencias"
         frmmensaje.Show 1
      End If
      txt_codigo.SetFocus
   End If
End Sub

Private Sub txt_nombre_almacen_KeyDown(KeyCode As Integer, Shift As Integer)
   If txt_almacen.Enabled = True Then
      If KeyCode = 116 Then
         lv_lista.ListItems.Clear
         If var_tipo_permiso = 1 Then
            rs.Open "select * from vw_almacen_permiso_1 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
         Else
            rs.Open "select * from vw_movimientos_almacenes where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
         End If
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ALM_ALMACEN_ID)
               list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
               rs.MoveNext
         Wend
         rs.Close
         lbl_lista = "Almacenes"
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

Private Sub txt_nombre_almacen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(txt_almacen) <> "" Then
         If txt_proveedor.Enabled = True Then
            txt_proveedor.SetFocus
         End If
      End If
   End If
End Sub

Private Sub txt_proveedor_KeyDown(KeyCode As Integer, Shift As Integer)
   If txt_proveedor.Enabled = True Then
      If KeyCode = 116 Then
         rs.Open "SELECT DISTINCT dbo.TB_ARCHIVO_COMPARACION.VCHA_COM_PROVEEDOR, dbo.TB_UNIDADESORGANIZACIONALES.VCHA_UOR_NOMBRE FROM dbo.TB_ARCHIVO_COMPARACION INNER JOIN dbo.TB_UNIDADESORGANIZACIONALES ON dbo.TB_ARCHIVO_COMPARACION.VCHA_COM_PROVEEDOR = dbo.TB_UNIDADESORGANIZACIONALES.VCHA_UOR_UNIDAD_ID WHERE (dbo.TB_ARCHIVO_COMPARACION.VCHA_MOV_MOVIMIENTO_ID = 'EI') AND (dbo.TB_ARCHIVO_COMPARACION.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "')", cnn, adOpenDynamic, adLockOptimistic
         lv_lista.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_COM_PROVEEDOR)
               list_item.SubItems(1) = IIf(IsNull(rs!VCHA_UOR_NOMBRE), "", rs!VCHA_UOR_NOMBRE)
               rs.MoveNext
         Wend
         rs.Close
         lbl_lista = "PLANTAS"
         var_tipo_lista = 2
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

Private Sub txt_proveedor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Len(Trim(txt_proveedor)) > 0 Then
         'rs.Open "Select * from tb_proveedores where vcha_pro_proveedor_id = '" + txt_proveedor + "'", cnn, adOpenDynamic, adLockOptimistic
         rs.Open "SELECT DISTINCT dbo.TB_ARCHIVO_COMPARACION.VCHA_COM_PROVEEDOR, dbo.TB_UNIDADESORGANIZACIONALES.VCHA_UOR_NOMBRE FROM dbo.TB_ARCHIVO_COMPARACION INNER JOIN dbo.TB_UNIDADESORGANIZACIONALES ON dbo.TB_ARCHIVO_COMPARACION.VCHA_COM_PROVEEDOR = dbo.TB_UNIDADESORGANIZACIONALES.VCHA_UOR_UNIDAD_ID WHERE (dbo.TB_ARCHIVO_COMPARACION.VCHA_MOV_MOVIMIENTO_ID = 'EI') AND (dbo.TB_ARCHIVO_COMPARACION.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND dbo.TB_ARCHIVO_COMPARACION.VCHA_COM_PROVEEDOR = '" + Me.txt_proveedor + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            txt_nombre_proveedor = IIf(IsNull(rs!VCHA_UOR_NOMBRE), "", rs!VCHA_UOR_NOMBRE)
            txt_codigo.Enabled = True
            txt_codigo.SetFocus
            txt_proveedor.Enabled = False
         Else
            MsgBox "Clave de planta incorrecto", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         MsgBox "Debe de indicar una planta", vbOKOnly, "ATENCION"
      End If
   End If
End Sub


Private Sub txt_referencia_KeyDown(KeyCode As Integer, Shift As Integer)
   If txt_proveedor.Enabled = True Then
      If KeyCode = 116 Then
         rs.Open "SELECT * FROM TB_PROVEEDORES ORDER BY VCHA_PRO_NOMBRE", cnn, adOpenDynamic, adLockOptimistic
         lv_lista.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_PRO_PROVEEDOR_ID)
               list_item.SubItems(1) = IIf(IsNull(rs!VCHA_PRO_NOMBRE), "", rs!VCHA_PRO_NOMBRE)
               rs.MoveNext
         Wend
         rs.Close
         lbl_lista = "PROVEEDORES"
         var_tipo_lista = 2
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


