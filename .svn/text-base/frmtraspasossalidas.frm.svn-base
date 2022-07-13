VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmtraspasossalidas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Salida para traspasos"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   Icon            =   "frmtraspasossalidas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   7665
   Begin VB.CommandButton cmd_reclasificacion 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1110
      Picture         =   "frmtraspasossalidas.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Salidas de Reclasificación"
      Top             =   720
      Width           =   330
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   675
      TabIndex        =   29
      Top             =   990
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   30
         TabIndex        =   30
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
         TabIndex        =   31
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1110
      Picture         =   "frmtraspasossalidas.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Cancelar Movimiento"
      Top             =   720
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmtraspasossalidas.frx":0ACE
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   720
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   450
      Picture         =   "frmtraspasossalidas.frx":0BD0
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Buscar Movimiento"
      Top             =   720
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   780
      Picture         =   "frmtraspasossalidas.frx":0CD2
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   720
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7125
      Picture         =   "frmtraspasossalidas.frx":0DD4
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Salir"
      Top             =   720
      Width           =   330
   End
   Begin VB.Frame frm_busqueda 
      Height          =   960
      Left            =   390
      TabIndex        =   0
      Top             =   1110
      Width           =   3135
      Begin VB.TextBox txt_busqueda_folio 
         Height          =   315
         Left            =   195
         TabIndex        =   1
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
         TabIndex        =   2
         Top             =   120
         Width           =   3075
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1350
      Index           =   0
      Left            =   5400
      TabIndex        =   18
      Top             =   1095
      Width           =   2175
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
         Height          =   480
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   540
         Width           =   2040
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   1
         Left            =   30
         TabIndex        =   20
         Top             =   120
         Width           =   2085
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   0
      Left            =   75
      TabIndex        =   7
      Top             =   570
      Width           =   7485
   End
   Begin VB.Frame Frame3 
      Height          =   1350
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1095
      Width           =   5235
      Begin VB.TextBox txt_almacen_origen 
         Height          =   315
         Left            =   780
         TabIndex        =   35
         Top             =   525
         Width           =   765
      End
      Begin VB.TextBox txt_nombre_almacen_origen 
         Height          =   315
         Left            =   1560
         TabIndex        =   34
         Top             =   525
         Width           =   3600
      End
      Begin VB.TextBox txt_almacen_destino 
         Height          =   315
         Left            =   780
         TabIndex        =   33
         Top             =   855
         Width           =   765
      End
      Begin VB.TextBox txt_nombre_almacen_destino 
         Height          =   315
         Left            =   1560
         TabIndex        =   32
         Top             =   855
         Width           =   3600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   23
         Top             =   885
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Origen:"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   6
         Top             =   555
         Width           =   510
      End
      Begin VB.Label label 
         BackColor       =   &H8000000D&
         Caption         =   " Datos del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   5
         Top             =   120
         Width           =   5145
      End
   End
   Begin VB.TextBox txt_foco 
      Height          =   285
      Left            =   7845
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3615
      Width           =   1125
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   75
      Top             =   0
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
            Picture         =   "frmtraspasossalidas.frx":140E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasossalidas.frx":1CE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasossalidas.frx":25C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasossalidas.frx":2B5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasossalidas.frx":343A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasossalidas.frx":3D14
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasossalidas.frx":45EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasossalidas.frx":4700
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasossalidas.frx":4812
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasossalidas.frx":4924
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasossalidas.frx":4A36
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasossalidas.frx":4B48
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   75
      TabIndex        =   21
      Top             =   975
      Width           =   7485
   End
   Begin VB.Frame Frame2 
      Height          =   4860
      Left            =   120
      TabIndex        =   8
      Top             =   2430
      Width           =   7425
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
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   1785
         TabIndex        =   10
         Top             =   1755
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar 
            Height          =   330
            Left            =   60
            TabIndex        =   11
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
            TabIndex        =   12
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
         TabIndex        =   9
         Top             =   450
         Width           =   2640
      End
      Begin MSComctlLib.ListView lv_traspasossalidas 
         Height          =   3375
         Left            =   45
         TabIndex        =   14
         Top             =   1035
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   5953
         View            =   3
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
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   8441
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label lbl_total 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3795
         TabIndex        =   36
         Top             =   4455
         Width           =   3510
      End
      Begin VB.Label lbl_cantidad 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   4410
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   120
         Width           =   7350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código del Artículo:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   615
         Width           =   1395
      End
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
      TabIndex        =   22
      Top             =   60
      Width           =   7365
   End
End
Attribute VB_Name = "frmtraspasossalidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Dim var_tipo_almacen As String
Dim var_correo_electronico As String
Dim var_clave_moneda As String
Dim var_tipo_lista As Integer
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



Private Sub cmb_almacen_destino_Click()
   var_almacen_Destino = Obtener_llave(cnn, rsaux, "TB_almacenes", "VCHA_ALM_NOMBRE", cmb_almacen_destino, 2, "T")
   var_tipo_almacen = Obtener_llave(cnn, rsaux, "TB_almacenes", "VCHA_ALM_NOMBRE", cmb_almacen_destino, 10, "T")
   var_correo_electronico = Obtener_llave(cnn, rsaux, "TB_almacenes", "VCHA_ALM_NOMBRE", cmb_almacen_destino, 9, "T")
   txt_codigo.Enabled = True
End Sub

Private Sub cmb_almacen_destino_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      var_almacen_Destino = Obtener_llave(cnn, rsaux, "TB_almacenes", "VCHA_ALM_NOMBRE", cmb_almacen_destino, 2, "T")
      txt_codigo.Enabled = True
      txt_codigo.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub cmb_almacen_origen_Click()
   var_almacen_origen = Obtener_llave(cnn, rsaux, "TB_almacenes", "VCHA_ALM_NOMBRE", cmb_almacen_origen, 2, "T")
   cmb_almacen_destino.Enabled = True
End Sub

Private Sub cmb_almacen_origen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      var_almacen_origen = Obtener_llave(cnn, rsaux, "TB_almacenes", "VCHA_ALM_NOMBRE", cmb_almacen_origen, 2, "T")
      cmb_almacen_destino.Enabled = True
      cmb_almacen_destino.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub cmd_buscar_Click()
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
   Set TB_EXISTENCIAS_INSERTA = New TB_EXISTENCIAS_INSERTA
   Set TB_SALIDAS_INSERTA = New TB_SALIDAS_INSERTA
   Set TB_TRASPASOS_INSERTA = New TB_TRASPASOS_INSERTA
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
   Dim var_codigo As String
   If var_numero_folio > 0 Then
      If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
         rs.Open "select top 1 vcha_alm_almacen_id from tb_reclasificacion where vcha_alm_almacen_id = '" + var_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Set reporte = appl.OpenReport(App.Path + "\rep_salidas_traspasos_reclasificacion.rpt")
            reporte.RecordSelectionFormula = "{VW_SALIDAS_TRASPASOS_RECLASIFICACION.vcha_emp_empresa_id} = '" + var_empresa + "' and  {VW_SALIDAS_TRASPASOS_RECLASIFICACION.VCHA_EMO_ALMACEN_ORIGEN} = '" + var_almacen_origen + "' and {VW_SALIDAS_TRASPASOS_RECLASIFICACION.INTE_tra_NUMERO} = " + Str(var_numero_folio) + " and {VW_SALIDAS_TRASPASOS_RECLASIFICACION.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' and {VW_SALIDAS_TRASPASOS_RECLASIFICACION.VCHA_EMP_EMPRESA_ID} =  '" + var_empresa + "' AND {VW_SALIDAS_TRASPASOS_RECLASIFICACION.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "'"
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Movimientos"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            If var_unidad_organizacional = "21" Then
               rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '12' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
            Else
               rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
            End If
            If var_tipo_almacen = "T" Then
               Call pro_envio_correo_app(var_correo_electronico, "Nota de Envio " & var_numero_folio, "Se anexa nota de envio", App.Path & "\dev_tien.dbf")
            End If
            
            Set reporte = appl.OpenReport(App.Path + "\rep_salidas_traspasos.rpt")
            If var_unidad_organizacional = "21" Then
               reporte.RecordSelectionFormula = "{VW_SALIDAS_TRASPASOS.vcha_emp_empresa_id} = '" + var_empresa + "' and  {VW_SALIDAS_TRASPASOS.VCHA_EMO_ALMACEN_ORIGEN} = '" + var_almacen_origen + "' and {VW_SALIDAS_TRASPASOS.INTE_EMO_NUMERO} = " + Str(var_numero_folio) + " and {VW_SALIDAS_TRASPASOS.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' and {VW_SALIDAS_TRASPASOS.VCHA_EMP_EMPRESA_ID} =  '" + var_empresa + "' AND {VW_SALIDAS_TRASPASOS.VCHA_UOR_UNIDAD_ID} = '12'"
            Else
               reporte.RecordSelectionFormula = "{VW_SALIDAS_TRASPASOS.vcha_emp_empresa_id} = '" + var_empresa + "' and  {VW_SALIDAS_TRASPASOS.VCHA_EMO_ALMACEN_ORIGEN} = '" + var_almacen_origen + "' and {VW_SALIDAS_TRASPASOS.INTE_EMO_NUMERO} = " + Str(var_numero_folio) + " and {VW_SALIDAS_TRASPASOS.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' and {VW_SALIDAS_TRASPASOS.VCHA_EMP_EMPRESA_ID} =  '" + var_empresa + "' AND {VW_SALIDAS_TRASPASOS.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "'"
            End If
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Movimientos"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
            If var_tipo_almacen = "T" Then
               Call pro_envio_correo_app(var_correo_electronico, "Nota de Envio " & var_numero_folio, "Se anexa nota de envio", App.Path & "\dev_tien.dbf")
            End If
            
         Else
            Set reporte = appl.OpenReport(App.Path + "\rep_salidas_traspasos.rpt")
            If var_unidad_organizacional = "21" Then
               reporte.RecordSelectionFormula = "{VW_SALIDAS_TRASPASOS.vcha_emp_empresa_id} = '" + var_empresa + "' and  {VW_SALIDAS_TRASPASOS.VCHA_EMO_ALMACEN_ORIGEN} = '" + var_almacen_origen + "' and {VW_SALIDAS_TRASPASOS.INTE_EMO_NUMERO} = " + Str(var_numero_folio) + " and {VW_SALIDAS_TRASPASOS.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' and {VW_SALIDAS_TRASPASOS.VCHA_EMP_EMPRESA_ID} =  '" + var_empresa + "' AND {VW_SALIDAS_TRASPASOS.VCHA_UOR_UNIDAD_ID} = '12'"
            Else
               reporte.RecordSelectionFormula = "{VW_SALIDAS_TRASPASOS.vcha_emp_empresa_id} = '" + var_empresa + "' and  {VW_SALIDAS_TRASPASOS.VCHA_EMO_ALMACEN_ORIGEN} = '" + var_almacen_origen + "' and {VW_SALIDAS_TRASPASOS.INTE_EMO_NUMERO} = " + Str(var_numero_folio) + " and {VW_SALIDAS_TRASPASOS.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' and {VW_SALIDAS_TRASPASOS.VCHA_EMP_EMPRESA_ID} =  '" + var_empresa + "' AND {VW_SALIDAS_TRASPASOS.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "'"
            End If
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Movimientos"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            If var_unidad_organizacional = "21" Then
               rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '12' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
            Else
               rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
            End If
            If var_tipo_almacen = "T" Then
               Call pro_envio_correo_app(var_correo_electronico, "Nota de Envio " & var_numero_folio, "Se anexa nota de envio", App.Path & "\dev_tien.dbf")
            End If
         End If
         rs.Close
      Else
         var_posible_Cantidad = 1
         If (var_empresa = "18" And var_almacen_origen <> "RETEX") Or var_empresa = "31" Then
            Cadena = "select * from tb_temporal_salidas with (nolock) where vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and floa_Sal_cantidad > 0"
            rsaux10.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux10.EOF
                  rsaux9.Open "select * from tb_existencias where vcha_Alm_almacen_id = '" + var_almacen_origen + "' and vcha_Art_Articulo_id = '" + rsaux10!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
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
         If var_posible_Cantidad = 1 Then
            var_si = MsgBox("¿Se va a imprimir el movimiento?", vbOKCancel, "ATENCION")
            If var_si = 1 Then
               cnn.BeginTrans
               Cadena = "select * from tb_temporal_salidas with (nolock) where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio)
               If rs.State = 1 Then
                  rs.Close
               End If
               rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  While Not rs.EOF
                        
                        var_codigo = rs!vcha_Art_Articulo_id
                        rsaux.Open "select * from TB_RECLASIFICACION where vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_art_Articulo_id = '" + rs!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           var_codigo = IIf(IsNull(rsaux!VCHA_REC_CODIGO_GENERAL), "", rsaux!VCHA_REC_CODIGO_GENERAL)
                        End If
                        rsaux.Close
                        
                        
                        var_suma_cantidad = 0
                        var_cantidad_llegar = IIf(IsNull(rs!floa_Sal_Cantidad), 0, rs!floa_Sal_Cantidad)
                        var_cantidad = var_cantidad_llegar
                        var_precio = IIf(IsNull(rs!floa_Sal_precio), 0, rs!floa_Sal_precio)
                        rsaux5.Open "select * from tb_existencias where vcha_art_articulo_id =  '" + var_codigo + "' and vcha_alm_almacen_id = '" + rs!VCHA_ALM_ALMACEN_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux5.EOF Then
                           var_costo = IIf(IsNull(rsaux5!FLOA_eXI_COSTO), 0, rsaux5!FLOA_eXI_COSTO)
                        Else
                           rsaux4.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_ART_ARTICULO_ID = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux4.EOF Then
                              var_costo = IIf(IsNull(rsaux4!mone_Art_costo_estandar), 0, rsaux4!mone_Art_costo_estandar)
                           Else
                              var_costo = 0
                              var_precio = 0
                           End If
                           rsaux4.Close
                        End If
                        rsaux5.Close
                        'While var_suma_cantidad < var_cantidad_llegar
                        '      rsaux2.Open "select * from tb_existencias where vcha_art_articulo_id =  '" + var_codigo + "' and vcha_alm_almacen_id = '" + rs!VCHA_ALM_ALMACEN_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                        '      If Not rsaux2.EOF Then
                        '         If rsaux2!floa_exi_cantidad_2004 >= var_cantidad_llegar Then
                        '            var_año = 2004
                        '            var_suma_cantidad = var_cantidad_llegar
                        '            var_cantidad = var_cantidad_llegar
                        '            var_costo = rsaux2!FLOA_EXI_COSTO_2004
                        '         Else
                        '            var_cantidad_disponible = IIf(IsNull(rsaux2!floa_exi_cantidad_2004), 0, rsaux2!floa_exi_cantidad_2004)
                        '            If var_cantidad_disponible > 0 Then
                        '               var_año = 2004
                        '               var_suma_cantidad = IIf(IsNull(rsaux2!floa_exi_cantidad_2004), 0, rsaux2!floa_exi_cantidad_2004)
                        '               var_cantidad = IIf(IsNull(rsaux2!floa_exi_cantidad_2004), 0, rsaux2!floa_exi_cantidad_2004)
                        '               var_costo = rsaux2!FLOA_EXI_COSTO_2004
                        '            Else
                        '               var_año = 2005
                        '               var_cantidad = rs!floa_sal_cantidad - var_suma_cantidad
                        '               var_suma_cantidad = var_cantidad_llegar
                        '               var_costo = rsaux2!floa_exi_costo_2005
                        '            End If
                        '         End If
                        '      Else
                        '         var_año = 2005
                        '         var_suma_cantidad = var_cantidad_llegar
                        '         var_cantidad = var_cantidad_llegar
                        '         rsaux4.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_ART_ARTICULO_ID = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                        '         If Not rsaux4.EOF Then
                        '            var_costo = IIf(IsNull(rsaux4!mone_art_costo_estandar), 0, rsaux4!mone_art_costo_estandar)
                        '         Else
                        '            var_costo = 0
                        '         End If
                        '         rsaux4.Close
                        '     End If
                        '      rsaux2.Close
                        '      rsaux4.Open "select * from tb_articulos where vcha_art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                        '      If Not rsaux4.EOF Then
                        '         var_precio = IIf(IsNull(rsaux4!mone_Art_precio_base), 0, rsaux4!mone_Art_precio_base)
                        '      Else
                        '         var_precio = 0
                        '      End If
                        '      rsaux4.Close
                        '      rsaux4.Open "insert into tb_salidas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_sal_numero, vcha_art_articulo_id, floa_sal_cantidad, floa_sal_costo, floa_sal_precio, inte_sal_año) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_origen + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!inte_sal_numero) + ", '" + var_codigo + "', " + CStr(var_cantidad) + ", " + CStr(var_costo) + " , " + CStr(var_precio) + ", " + CStr(var_año) + ")", cnn, adOpenDynamic, adLockOptimistic
                        '      rsaux4.Open "insert into TB_TRASPASOS (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_tra_numero, vcha_art_articulo_id, floa_tra_cantidad, floa_tra_costo, floa_tra_precio, INTE_tra_AÑO, VCHA_TRA_ALMACEN_ORIGEN) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_origen + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!inte_sal_numero) + ", '" + rs!vcha_Art_articulo_id + "', " + CStr(var_cantidad) + ", " + CStr(var_costo) + " , " + CStr(rs!floa_sal_precio) + ", " + CStr(var_año) + ", '" + var_almacen_origen + "')", cnn, adOpenDynamic, adLockOptimistic
                        'Wend
                        var_año = 2005
                        rsaux4.Open "insert into tb_salidas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_sal_numero, vcha_art_articulo_id, floa_sal_cantidad, floa_sal_costo, floa_sal_precio, inte_sal_año) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_origen + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_SAL_NUMERO) + ", '" + var_codigo + "', " + CStr(var_cantidad) + ", " + CStr(var_costo) + " , " + CStr(var_precio) + ", " + CStr(var_año) + ")", cnn, adOpenDynamic, adLockOptimistic
                        rsaux4.Open "insert into TB_TRASPASOS (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_tra_numero, vcha_art_articulo_id, floa_tra_cantidad, floa_tra_costo, floa_tra_precio, INTE_tra_AÑO, VCHA_TRA_ALMACEN_ORIGEN) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_origen + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_SAL_NUMERO) + ", '" + rs!vcha_Art_Articulo_id + "', " + CStr(var_cantidad) + ", " + CStr(var_costo) + " , " + CStr(rs!floa_Sal_precio) + ", " + CStr(var_año) + ", '" + var_almacen_origen + "')", cnn, adOpenDynamic, adLockOptimistic
                        rs.MoveNext
                  Wend
                  rs.Close
               End If
               var_estatus_movimiento = "I"
               If var_unidad_organizacional = "21" Then
                  rsaux4.Open "update tb_encabezado_movimientos set char_emo_estatus = 'I' where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '12' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
               Else
                  rsaux4.Open "update tb_encabezado_movimientos set char_emo_estatus = 'I' where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
               End If
               cnn.CommitTrans
               rs.Open "select top 1 vcha_alm_almacen_id  from tb_reclasificacion where vcha_alm_almacen_id = '" + var_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
              
                  Set reporte = appl.OpenReport(App.Path + "\rep_salidas_traspasos_reclasificacion.rpt")
                  If var_unidad_organizacional = "21" Then
                     reporte.RecordSelectionFormula = "{VW_SALIDAS_TRASPASOS_RECLASIFICACION.vcha_emp_empresa_id} = '" + var_empresa + "' and  {VW_SALIDAS_TRASPASOS_RECLASIFICACION.VCHA_EMO_ALMACEN_ORIGEN} = '" + var_almacen_origen + "' and {VW_SALIDAS_TRASPASOS_RECLASIFICACION.INTE_tra_NUMERO} = " + Str(var_numero_folio) + " and {VW_SALIDAS_TRASPASOS_RECLASIFICACION.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' and {VW_SALIDAS_TRASPASOS_RECLASIFICACION.VCHA_EMP_EMPRESA_ID} =  '" + var_empresa + "' AND {VW_SALIDAS_TRASPASOS_RECLASIFICACION.VCHA_UOR_UNIDAD_ID} = '12'"
                  Else
                     reporte.RecordSelectionFormula = "{VW_SALIDAS_TRASPASOS_RECLASIFICACION.vcha_emp_empresa_id} = '" + var_empresa + "' and  {VW_SALIDAS_TRASPASOS_RECLASIFICACION.VCHA_EMO_ALMACEN_ORIGEN} = '" + var_almacen_origen + "' and {VW_SALIDAS_TRASPASOS_RECLASIFICACION.INTE_tra_NUMERO} = " + Str(var_numero_folio) + " and {VW_SALIDAS_TRASPASOS_RECLASIFICACION.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' and {VW_SALIDAS_TRASPASOS_RECLASIFICACION.VCHA_EMP_EMPRESA_ID} =  '" + var_empresa + "' AND {VW_SALIDAS_TRASPASOS_RECLASIFICACION.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "'"
                  End If
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Reporte de Movimientos"
                  frmvistasprevias.Show 1
                  Set reporte = Nothing
                  If var_unidad_organizacional = "21" Then
                     rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '12' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                  Else
                     rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                  End If
                  If var_tipo_almacen = "T" Then
                     Call pro_envio_correo_app(var_correo_electronico, "Nota de Envio " & var_numero_folio, "Se anexa nota de envio", App.Path & "\dev_tien.dbf")
                  End If
               
                  Set reporte = appl.OpenReport(App.Path + "\rep_salidas_traspasos.rpt")
                  reporte.RecordSelectionFormula = "{VW_SALIDAS_TRASPASOS.vcha_emp_empresa_id} = '" + var_empresa + "' and  {VW_SALIDAS_TRASPASOS.VCHA_EMO_ALMACEN_ORIGEN} = '" + var_almacen_origen + "' and {VW_SALIDAS_TRASPASOS.INTE_EMO_NUMERO} = " + Str(var_numero_folio) + " and {VW_SALIDAS_TRASPASOS.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' and {VW_SALIDAS_TRASPASOS.VCHA_EMP_EMPRESA_ID} =  '" + var_empresa + "' AND {VW_SALIDAS_TRASPASOS.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "'"
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Reporte de Movimientos"
                  frmvistasprevias.Show 1
                  Set reporte = Nothing
                  rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                  If var_tipo_almacen = "T" Then
                     Call pro_envio_correo_app(var_correo_electronico, "Nota de Envio " & var_numero_folio, "Se anexa nota de envio", App.Path & "\dev_tien.dbf")
                  End If
               Else
                  var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, var_numero_folio, "I", Now, 1)
                  Set reporte = appl.OpenReport(App.Path + "\rep_salidas_traspasos.rpt")
                  reporte.RecordSelectionFormula = "{VW_SALIDAS_TRASPASOS.vcha_emp_empresa_id} = '" + var_empresa + "' and {VW_SALIDAS_TRASPASOS.VCHA_EMO_ALMACEN_ORIGEN} = '" + var_almacen_origen + "' and {VW_SALIDAS_TRASPASOS.INTE_EMO_NUMERO} = " + Str(var_numero_folio) + " and {VW_SALIDAS_TRASPASOS.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' and {VW_SALIDAS_TRASPASOS.VCHA_EMP_EMPRESA_ID} =  '" + var_empresa + "' AND {VW_SALIDAS_TRASPASOS.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "'"
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Reporte de Movimientos"
                  frmvistasprevias.Show 1
                  rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                  Set reporte = Nothing
                  If var_tipo_almacen = "T" Then
                     Call pro_envio_correo_app(var_correo_electronico, "Nota de Envio " & var_numero_folio, "Se anexa nota de envio", App.Path & "\dev_tien.dbf")
                  End If
                  txt_codigo.Enabled = False
                  txt_foco.Enabled = False
               End If
            End If
         Else
            MsgBox "El movimiento no se puede imprimir ya que las existencias de los siguientes artículos exceden a la cantidad disponible en el almacen " + var_cadena_articulos
         End If
       End If
    Else
       MsgBox "No se a seleccionado ningún movimiento", vbOKOnly, "ATENCION"
    End If
    If rs.State = 1 Then
       rs.Close
    End If
    If rsaux.State = 1 Then
       rs.Close
    End If
    If rsaux2.State = 1 Then
       rsaux2.Close
    End If
    If rsaux3.State = 1 Then
       rsaux3.Close
    End If
    If rsaux4.State = 1 Then
       rsaux4.Close
    End If
End Sub

Private Sub cmd_nuevo_Click()
   txt_codigo.Enabled = False
   var_primera_vez = True
   frm_busqueda.Visible = False
   lv_traspasossalidas.ListItems.Clear
   var_numero_folio = 0
   txt_folio = ""
   txt_codigo = ""
   var_estatus_movimiento = ""
   txt_nombre_almacen_destino = ""
   txt_almacen_destino = ""
   txt_nombre_almacen_origen = ""
   txt_almacen_origen = ""
   txt_almacen_origen.Enabled = True
   Me.lbl_total = Format(0, "###,###,##0.00")
   txt_almacen_origen.SetFocus
End Sub

Private Sub cmd_reclasificacion_Click()
   cnn.BeginTrans
   rs.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_RECLASIFICACION", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
   Else
      var_consecutivo = 1
   End If
   rs.Close
   rs.Open "Insert Into TB_TEMP_RECLASIFICACION (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
   cnn.CommitTrans
   Cadena = "select * from tb_temporal_salidas with (nolock) where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio)
   If rs.State = 1 Then
      rs.Close
    End If
    rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
       While Not rs.EOF
             var_codigo = rs!vcha_Art_Articulo_id
             rsaux.Open "select * from TB_RECLASIFICACION where vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_art_Articulo_id = '" + rs!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
             If Not rsaux.EOF Then
                var_codigo = IIf(IsNull(rsaux!VCHA_REC_CODIGO_GENERAL), "", rsaux!VCHA_REC_CODIGO_GENERAL)
             End If
             rsaux.Close
             rsaux.Open "INSERT INTO TB_TEMP_RECLASIFICACION (INTE_TEM_CONSECUTIVO, VCHA_aRT_ARTICULO_ID, FLOA_TEM_cANTIDAD) VALUES (" + CStr(var_consecutivo) + ",'" + var_codigo + "'," + CStr(rs!floa_Sal_Cantidad) + ")", cnn, adOpenDynamic, adLockOptimistic
             rs.MoveNext
       Wend
    End If
    rs.Close
    
   Set reporte = appl.OpenReport(App.Path + "\rep_reclasificacion_agrupamiento.rpt")
   reporte.RecordSelectionFormula = "{VW_TEMPORAL_RECLASIFICACION.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
   frmvistasprevias.cr.ReportSource = reporte
   For ntablas = 1 To reporte.Database.Tables.Count
       reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
   Next ntablas
   frmvistasprevias.cr.ViewReport
   frmvistasprevias.Caption = "Reporte de Antigüedad de Saldos"
   frmvistasprevias.Show 1
   Set reporte = Nothing
   rs.Open "delete from TB_TEMP_RECLASIFICACION where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
    
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

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   var_posible_kanban = 0
   var_cadena_seguridad = ""
   frm_lista.Visible = False
   lbl_total = Format(0, "###,###,##0.00")
   Top = 0
   Left = 2000
   var_estatus_movimiento = ""
   frm_busqueda.Visible = False
   frm_eliminar.Visible = False
   lbl_cantidad.Visible = False
   txt_cantidad.Visible = False
   txt_codigo.Enabled = False
   var_primera_vez = True
   var_cantidad_leida = 1#
   rs.Open "select * from tb_monedas where inte_mon_moneda_local = 1", cnn, adOpenDynamic
   var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
   rs.Close
   txt_nombre_almacen_destino = ""
   txt_almacen_destino = ""
   txt_nombre_almacen_origen = ""
   txt_almacen_origen = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_traspasossalidas)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 0 Then
         If var_tipo_lista = 1 Then
            txt_almacen_origen = lv_lista.selectedItem
            txt_nombre_almacen_origen = lv_lista.selectedItem.SubItems(1)
            txt_almacen_origen.SetFocus
         End If
         If var_tipo_lista = 2 Then
            txt_almacen_destino = lv_lista.selectedItem
            txt_nombre_almacen_destino = lv_lista.selectedItem.SubItems(1)
            txt_almacen_destino.SetFocus
         End If
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

Private Sub lv_traspasossalidas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
         MsgBox "Imporsible modificar el movimiento", vbOKOnly, "ATENCION"
      Else
         frm_eliminar.Visible = True
         txt_cantidad_eliminar.SetFocus
      End If
   End If
End Sub

Private Sub Toolbar1_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
End Sub

Private Sub txt_almacen_destino_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_almacen_destino_KeyDown(KeyCode As Integer, Shift As Integer)
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
      lbl_lista = "Almacenes Destino"
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
End Sub

Private Sub txt_almacen_destino_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      txt_nombre_almacen_destino.SetFocus
   End If
End Sub

Private Sub txt_almacen_destino_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_almacen_destino) <> "" Then
      If var_tipo_permiso = 1 Then
         rs.Open "select * from vw_almacen_permiso_1 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_alm_almacen_id = '" + txt_almacen_destino + "'", cnn, adOpenDynamic, adLockBatchOptimistic
      Else
         rs.Open "select * from vw_movimientos_almacenes where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_alm_almacen_id = '" + txt_almacen_destino + "'", cnn, adOpenDynamic, adLockBatchOptimistic
      End If
      If Not rs.EOF Then
         txt_almacen_destino = rs!VCHA_ALM_ALMACEN_ID
         txt_nombre_almacen_destino = rs!VCHA_ALM_NOMBRE
         var_almacen_Destino = rs!VCHA_ALM_ALMACEN_ID
         txt_almacen_destino.Enabled = False
         var_tipo_almacen = IIf(IsNull(rs!char_alm_tipo), "", rs!char_alm_tipo)
         var_correo_electronico = IIf(IsNull(rs!vcha_alm_correo), "", rs!vcha_alm_correo)
         txt_codigo.Enabled = True
      Else
         txt_codigo.Enabled = False
         txt_almacen_destino = ""
         txt_nombre_almacen_destino = ""
         MsgBox "Clave de almacen incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   End If
End Sub

Private Sub txt_almacen_origen_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_almacen_origen_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      If var_tipo_permiso = 1 Then
         rs.Open "select distinct * from vw_almacen_permiso_2 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      Else
         rs.Open "select distinct * from vw_movimientos_almacenes where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
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
End Sub

Private Sub txt_almacen_origen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      txt_nombre_almacen_origen.SetFocus
   End If
End Sub

Private Sub txt_almacen_origen_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_almacen_origen) <> "" Then
      If var_tipo_permiso = 1 Then
         rs.Open "select * from vw_almacen_permiso_2 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_alm_almacen_id = '" + txt_almacen_origen + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      Else
         rs.Open "select * from vw_movimientos_almacenes where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_alm_almacen_id = '" + txt_almacen_origen + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      End If
      If Not rs.EOF Then
         var_almacen_origen = rs!VCHA_ALM_ALMACEN_ID
         txt_nombre_almacen_origen = rs!VCHA_ALM_NOMBRE
         txt_almacen_destino.Enabled = True
         txt_almacen_origen.Enabled = False
      Else
         var_almacen_origen = ""
         txt_almacen_origen = ""
         txt_nombre_almacen_origen = ""
         MsgBox "Clave de almacen incorrecto", vbOKOnly, "ATENCION"
      End If
      rs.Close
   End If
End Sub

Private Sub txt_busqueda_folio_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If Trim(txt_busqueda_folio) <> "" Then
         If var_unidad_organizacional = "21" Then
            rs.Open "select * from tb_encabezado_movimientos where inte_emo_numero = " + txt_busqueda_folio + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '12'", cnn, adOpenDynamic, adLockOptimistic
         Else
            rs.Open "select * from tb_encabezado_movimientos where inte_emo_numero = " + txt_busqueda_folio + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
         End If
         If Not rs.EOF Then
            var_almacen_destino_tem = rs!VCHA_EMO_ALMACEN_DESTINO
            var_almacen_origen_tem = rs!VCHA_ALM_ALMACEN_ID
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
               var_estatus_movimiento = rs!char_Emo_estatus
               var_almacen_Destino = rs!VCHA_EMO_ALMACEN_DESTINO
               var_almacen_origen = rs!VCHA_ALM_ALMACEN_ID
               lv_traspasossalidas.ListItems.Clear
               var_primera_vez = False
               var_numero_folio = rs!INTE_EMO_NUMERO
               txt_folio = var_numero_folio
               rsaux.Open "select * from tb_ALMACENES where vcha_alm_almacen_id = '" + var_almacen_Destino + "'", cnn, adOpenDynamic, adLockOptimistic
               txt_almacen_destino = rsaux(2).Value
               txt_nombre_almacen_destino = rsaux(3).Value
               rsaux.Close
               rsaux.Open "select * from tb_ALMACENES where vcha_alm_almacen_id = '" + var_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
               txt_almacen_origen = rsaux(2).Value
               txt_nombre_almacen_origen = rsaux(3).Value
               rsaux.Close
               rsaux.Open "SELECT * FROM TB_ALMACENES WHERE VCHA_ALM_ALMACEN_ID = '" + txt_almacen_destino + "'", cnn, adOpenDynamic, adLockOptimistic
               var_tipo_almacen = IIf(IsNull(rsaux!char_alm_tipo), "", rsaux!char_alm_tipo)
               var_correo_electronico = IIf(IsNull(rsaux!vcha_alm_correo), "", rsaux!vcha_alm_correo)
               rsaux.Close
               rsaux.Open "select * from tb_temporal_salidas with (nolock) where inte_sal_numero = " + txt_busqueda_folio + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
               Me.lbl_total = Format(0, "###,###,##0.0000")
               If Not rsaux.EOF Then
                  While Not rsaux.EOF
                     rsaux2.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rsaux!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        Set list_item = lv_traspasossalidas.ListItems.Add(, , rsaux!vcha_Art_Articulo_id)
                        list_item.SubItems(1) = IIf(IsNull(rsaux2(1).Value), "", rsaux2(1).Value)
                        list_item.SubItems(2) = Format(Round(IIf(IsNull(rsaux!floa_Sal_Cantidad), 0, rsaux!floa_Sal_Cantidad), 4), "###,###,##0.0000")
                        rsaux2.Close
                        Me.lbl_total = Format(CDbl(lbl_total) + IIf(IsNull(rsaux!floa_Sal_Cantidad), 0, rsaux!floa_Sal_Cantidad), "###,###,##0.0000")
                        rsaux.MoveNext:
                     End If
                  Wend
               End If
               rsaux.Close
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
            MsgBox "El número de movimiento no existe ", vbOKOnly, "ATENCION"
         End If
         rs.Close
      End If
      frm_busqueda.Visible = False
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
         If CDbl(txt_cantidad_eliminar) > CDbl(lv_traspasossalidas.selectedItem.SubItems(2)) Then
         Else
            Set TB_TEMPORAL_SALIDAS_MODIFICA = New TB_TEMPORAL_SALIDAS_MODIFICA
            var_cantidad_eliminar = Val(txt_cantidad_eliminar)
            Cadena = "select * from tb_temporal_salidas with (nolock) where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + lv_traspasossalidas.selectedItem + "' and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'"
            rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
            var_inserta = False
            var_inserta = TB_TEMPORAL_SALIDAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, var_numero_folio, lv_traspasossalidas.selectedItem, 0 - Val(txt_cantidad_eliminar))
            rs.Close
            lv_traspasossalidas.selectedItem.SubItems(2) = Format(Round(lv_traspasossalidas.selectedItem.SubItems(2) - CDbl(txt_cantidad_eliminar), 4), "###,###,##0.0000")
            var_renglon = lv_traspasossalidas.selectedItem.Index
            Call ilumina_grid
            frm_eliminar.Visible = False
            Me.lbl_total = Format(CDbl(lbl_total) - CDbl(var_cantidad_leida), "###,###,##0.0000")
            txt_codigo.SetFocus
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
   txt_cantidad = ""
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
         var_cantidad_leida = txt_cantidad
         txt_foco.Enabled = True
         txt_foco.SetFocus
         lbl_cantidad.Visible = False
         txt_cantidad.Visible = False
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
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      var_costo = 4
      var_precio = 0
      var_cantidad_caja = 0
      If var_caja = "000005" Or var_caja = "000010" Or var_caja = "000015" Or var_caja = "000020" Then
         var_cantidad_caja = CInt(var_caja)
         txt_codigo = Mid(txt_codigo, 7, 5)
      End If
      If Trim(txt_codigo) <> "" Then
         If rs.State = 1 Then
            rs.Close
         End If
         rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            If IsNull(rs(43).Value) Then
               var_recontable = 0
            Else
               var_recontable = rs!inte_Art_salida_masiva
            End If
            var_descripcion_articulo = rs(1).Value
            var_costo = rs(3).Value
            var_precio = rs(2).Value
            rs.Close
            rs.Open "SELECT * FROM TB_EXISTENCIAS WHERE VCHA_ART_ARTICULO_ID = '" + txt_codigo + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_costo = rs(4).Value
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
                  If IsNull(rs(43).Value) Then
                     var_recontable = 0
                  Else
                     var_recontable = rs(43).Value
                  End If
                  var_descripcion_articulo = rs(1).Value
                  var_costo = IIf(IsNull(rs(3).Value), 0, rs(3).Value)
                  var_precio = IIf(IsNull(rs(2).Value), 0, rs(2).Value)
                  rs.Close
                  rs.Open "SELECT * FROM TB_EXISTENCIAS WHERE VCHA_ART_ARTICULO_ID = '" + txt_codigo + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_origen + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_costo = rs(4).Value
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
   Set TB_TEMPORAL_SALIDAS_INSERTA = New TB_TEMPORAL_SALIDAS_INSERTA
   Set TB_TEMPORAL_SALIDAS_MODIFICA = New TB_TEMPORAL_SALIDAS_MODIFICA
   Dim var_inserta As Boolean
   If Trim(txt_codigo.Text) <> "" Then
      var_pase_existencias = 1
      If (var_empresa = "18" And var_almacen_origen <> "RETEX") Or var_empresa = "31" Then
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
         rsaux.Open "select floa_exi_Cantidad_disponible from tb_existencias where vcha_alm_almacen_id = '" + var_almacen_origen + "' and vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            var_cantidad_Existencias = IIf(IsNull(rsaux!floa_Exi_Cantidad_disponible), 0, rsaux!floa_Exi_Cantidad_disponible)
         Else
            var_cantidad_Existencias = 0
         End If
         rsaux.Close
         var_cantidad_posible = var_cantidad_Existencias - (var_cantidad_temporal + var_cantidad_leida)
         If Round(var_cantidad_posible, 4) < 0 Then
            var_pase_existencias = 0
         End If
      End If
      If var_empresa = "18" Then
         If Me.txt_codigo = "360010000002" Or Me.txt_codigo = "360020000009" Or Me.txt_codigo = "900000000003" Or Me.txt_codigo = "911110000005" Then
            var_pase_existencias = True
         End If
      End If
      If var_pase_existencias = 1 Then
         bandera_suma = False
         If var_primera_vez = True Then
            var_inserta = False
            If var_unidad_organizacional = "16" Then
               var_insreta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, Now, var_numero_folio, 0, "", "", var_almacen_origen, var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, "", "", "", "", "B", "", "", 0, 0, 0, var_clave_moneda, 1)
            Else
               var_insreta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, "12", var_almacen_origen, var_clave_movimiento, Now, var_numero_folio, 0, "", "", var_almacen_origen, var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, "", "", "", "", "B", "", "", 0, 0, 0, var_clave_moneda, 1)
            End If
            var_numero_folio = var_numero_folio_regreso
            txt_folio = var_numero_folio
            var_primera_vez = False
         End If
         Cadena = "select * from tb_temporal_salidas with (nolock) where vcha_alm_almacen_id = '" + var_almacen_origen + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "' and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'"
         rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_inserta = False
            var_inserta = TB_TEMPORAL_SALIDAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, var_numero_folio, txt_codigo, Round(var_cantidad_leida, 4))
            rs.Close
            valor = Trim(txt_codigo)
            Set itmfound = lv_traspasossalidas.findItem(valor, lvwText, , lvwPartial)
            itmfound.EnsureVisible
            itmfound.Selected = True
            lv_traspasossalidas.selectedItem.SubItems(2) = Format(Round(lv_traspasossalidas.selectedItem.SubItems(2) + var_cantidad_leida, 4), "###,###,##0.0000")
            var_renglon = lv_traspasossalidas.selectedItem.Index
            Me.lbl_total = Format(CDbl(lbl_total) + var_cantidad_leida, "###,###,##0.0000")
            Call ilumina_grid
         Else
            var_inserta = False
            var_inserta = TB_TEMPORAL_SALIDAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida, var_costo, var_precio, "0", 0, 0)
            rs.Close
            Set list_item = lv_traspasossalidas.ListItems.Add(, , Trim(txt_codigo))
            list_item.SubItems(1) = Format(var_descripcion_articulo, "###,###,##0.0000")
            list_item.SubItems(2) = Format(var_cantidad_leida, "###,###,##0.0000")
            var_renglon = lv_traspasossalidas.ListItems.Count
            Me.lbl_total = Format(CDbl(lbl_total) + var_cantidad_leida, "###,###,##0.0000")
            Call ilumina_grid
         End If
      Else
         'mensage exceden
      End If
      txt_codigo.SetFocus
   End If
End Sub

Private Sub txt_nombre_almacen_destino_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_nombre_almacen_destino_KeyDown(KeyCode As Integer, Shift As Integer)
   If txt_almacen_destino.Enabled = True Then
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
         lbl_lista = "Almacenes Destino"
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

Private Sub txt_nombre_almacen_destino_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txt_codigo.Enabled = True Then
         txt_codigo.SetFocus
      End If
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_almacen_destino_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_almacen_origen_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_nombre_almacen_origen_KeyDown(KeyCode As Integer, Shift As Integer)
   If txt_almacen_origen.Enabled = True Then
      If KeyCode = 116 Then
         lv_lista.ListItems.Clear
         If var_tipo_permiso = 1 Then
            rs.Open "select * from vw_almacen_permiso_2 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
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

Private Sub txt_nombre_almacen_origen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txt_almacen_destino.Enabled = True Then
         txt_almacen_destino.SetFocus
      End If
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_almacen_origen_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub
