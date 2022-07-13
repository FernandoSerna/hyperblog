VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmsalidas_intercompañias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Salidas Intercompañia"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   KeyPreview      =   -1  'True
   LinkTopic       =   "Salidas Intercomañias"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_archivo_envios 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1065
      Picture         =   "frmsalidas_intercompañias.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Verificar archivo envios"
      Top             =   720
      Width           =   330
   End
   Begin VB.TextBox txt_foco 
      Height          =   285
      Left            =   8790
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   3360
      Width           =   1125
   End
   Begin VB.Frame frm_busqueda 
      Height          =   960
      Left            =   1455
      TabIndex        =   0
      Top             =   945
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
         Width           =   3060
      End
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   1215
      TabIndex        =   3
      Top             =   885
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   30
         TabIndex        =   4
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
         TabIndex        =   5
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1230
      Index           =   0
      Left            =   5565
      TabIndex        =   30
      Top             =   1110
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
         TabIndex        =   31
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
         TabIndex        =   32
         Top             =   120
         Width           =   1890
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   0
      Left            =   75
      TabIndex        =   29
      Top             =   570
      Width           =   7455
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmsalidas_intercompañias.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   720
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frmsalidas_intercompañias.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Buscar Movimiento"
      Top             =   720
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   735
      Picture         =   "frmsalidas_intercompañias.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   720
      Width           =   330
   End
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1065
      Picture         =   "frmsalidas_intercompañias.frx":0408
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Cancelar Movimiento"
      Top             =   720
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7170
      Picture         =   "frmsalidas_intercompañias.frx":050A
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Salir"
      Top             =   720
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   1230
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   1110
      Width           =   5415
      Begin VB.TextBox txt_almacen 
         Height          =   315
         Left            =   945
         TabIndex        =   20
         Top             =   450
         Width           =   1140
      End
      Begin VB.TextBox txt_nombre_almacen 
         Height          =   315
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   450
         Width           =   3255
      End
      Begin VB.TextBox txt_nombre_proveedor 
         Height          =   315
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   825
         Width           =   3255
      End
      Begin VB.TextBox txt_proveedor 
         Height          =   315
         Left            =   945
         TabIndex        =   17
         Top             =   810
         Width           =   1140
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Origen:"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   23
         Top             =   510
         Width           =   510
      End
      Begin VB.Label label 
         BackColor       =   &H8000000D&
         Caption         =   " Datos del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   22
         Top             =   120
         Width           =   5325
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   21
         Top             =   885
         Width           =   585
      End
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   570
      Top             =   825
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
            Picture         =   "frmsalidas_intercompañias.frx":0B44
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_intercompañias.frx":141E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_intercompañias.frx":1CF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_intercompañias.frx":2294
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_intercompañias.frx":2B70
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_intercompañias.frx":344A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_intercompañias.frx":3D24
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_intercompañias.frx":3E36
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_intercompañias.frx":3F48
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_intercompañias.frx":405A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_intercompañias.frx":416C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_intercompañias.frx":427E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   75
      TabIndex        =   33
      Top             =   975
      Width           =   7455
   End
   Begin VB.Frame Frame2 
      Height          =   4890
      Left            =   120
      TabIndex        =   6
      Top             =   2355
      Width           =   7425
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
         Left            =   5115
         TabIndex        =   11
         Top             =   555
         Width           =   1890
      End
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   1785
         TabIndex        =   8
         Top             =   1755
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar 
            Height          =   330
            Left            =   60
            TabIndex        =   9
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
            TabIndex        =   10
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
         TabIndex        =   7
         Top             =   495
         Width           =   2640
      End
      Begin MSComctlLib.ListView lv_entradas 
         Height          =   3150
         Left            =   45
         TabIndex        =   12
         Top             =   1110
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   5556
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
      Begin VB.Label txt_total 
         Alignment       =   1  'Right Justify
         Caption         =   "9999999999999999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3690
         TabIndex        =   37
         Top             =   4283
         Width           =   3660
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   90
         TabIndex        =   36
         Top             =   4320
         Width           =   3435
      End
      Begin VB.Label lbl_cantidad 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   4410
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   120
         Width           =   7350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código del Artículo:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   675
         Width           =   1395
      End
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   615
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
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
      Left            =   105
      TabIndex        =   34
      Top             =   75
      Width           =   7335
   End
End
Attribute VB_Name = "frmsalidas_intercompañias"
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
Dim var_cadena_conexion As String
Dim cnn_traspaso_intecomañia As ADODB.Connection

Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long


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




Private Sub cmd_archivo_envios_Click()
   If Me.txt_folio <> "" Then
      var_conexion_intercompañia = ""
      rsaux.Open "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE VCHA_UOR_UNIDAD_ID = '" + Me.txt_proveedor + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux.EOF Then
         var_conexion_intercompañia = IIf(IsNull(rsaux!vcha_uor_conexion), "", rsaux!vcha_uor_conexion)
      End If
      rsaux.Close
      Set cnn_intercompañias = CreateObject("ADODB.connection")
      cnn_intercompañias.Open var_conexion_intercompañia
      cnn_intercompañias.CursorLocation = adUseClient
      Cadena = "select * from tb_salidas where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'"
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         While Not rs.EOF
               var_codigo = rs!vcha_Art_articulo_id
               rsaux3.Open "select * from tb_Articulos where vcha_art_Articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux3.EOF Then
                  var_DEscripcion = IIf(IsNull(rsaux3!vcha_art_nombre_español), "", rsaux3!vcha_art_nombre_español)
               End If
               rsaux3.Close
               If rsaux2.State = 1 Then
                  rsuax2.Close
               End If
               rsaux5.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_ART_ARTICULO_ID = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
               var_linea = IIf(IsNull(rsaux5!vcha_lin_linea_id), "", rsaux5!vcha_lin_linea_id)
               var_catalogo = IIf(IsNull(rsaux5!vcha_Art_catalogo_vigente), "0", rsaux5!vcha_Art_catalogo_vigente)
               rsaux5.Close
               If IsNumeric(var_almacen_Destino) Then
                  rsaux4.Open "SELECT * FROM TB_ARCHIVOS_ENVIOS WHERE VCHA_aCO_PROVEEDOR = '" + var_unidad_organizacional + "' AND INTE_aCO_NUMERO = '" + Trim(var_almacen_Destino) + Trim(CStr(var_numero_folio)) + "' AND VCHA_ACO_CODIGO_EXTERNO = '" + rs!vcha_Art_articulo_id + "'", cnn_intercompañias, adOpenDynamic, adLockOptimistic
                  If rsaux4.EOF Then
                     rsaux2.Open "INSERT INTO TB_ARCHIVOS_ENVIOS (VCHA_ACO_PROVEEDOR, INTE_ACO_NUMERO, VCHA_ACO_CODIGO_EXTERNO, vcha_aco_descripcion_externa, VCHA_ART_ARTICULO_ID, FLOA_ACO_CANTIDAD, FLOA_ACO_COSTO, INTE_ACO_AÑO, VCHA_LIN_LINEA_ID, VCHA_CAT_CATALOGO_ID, FLOA_ACO_PRECIO) Values ( '" + var_unidad_organizacional + "', '" + Trim(var_almacen_Destino) + Trim(CStr(var_numero_folio)) + "', '" + rs!vcha_Art_articulo_id + "','" + var_DEscripcion + "', ''," + CStr(rs!FLOA_sAL_cANTIDAD) + ", " + CStr(rs!floa_Sal_costo) + ", " + CStr(rs!INTE_sAL_AÑO) + ",'" + var_linea + "','" + var_catalogo + "'," + CStr(rs!floa_Sal_precio) + ")", cnn_intercompañias, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux4.Close
               Else
                  rsaux4.Open "SELECT * FROM TB_ARCHIVOS_ENVIOS WHERE VCHA_aCO_PROVEEDOR = '" + var_unidad_organizacional + "' AND INTE_aCO_NUMERO =  '" + Trim(CStr(var_numero_folio)) + "' AND VCHA_ACO_CODIGO_EXTERNO = '" + rs!vcha_Art_articulo_id + "'", cnn_intercompañias, adOpenDynamic, adLockOptimistic
                  If rsaux4.EOF Then
                     rsaux2.Open "INSERT INTO TB_ARCHIVOS_ENVIOS (VCHA_ACO_PROVEEDOR, INTE_ACO_NUMERO, VCHA_ACO_CODIGO_EXTERNO, vcha_aco_descripcion_externa, VCHA_ART_ARTICULO_ID, FLOA_ACO_CANTIDAD, FLOA_ACO_COSTO, INTE_ACO_AÑO, VCHA_LIN_LINEA_ID, VCHA_CAT_CATALOGO_ID, FLOA_ACO_PRECIO) Values ( '" + var_unidad_organizacional + "', '" + Trim(CStr(var_numero_folio)) + "', '" + rs!vcha_Art_articulo_id + "','" + var_DEscripcion + "', ''," + CStr(rs!FLOA_sAL_cANTIDAD) + ", " + CStr(rs!floa_Sal_costo) + ", " + CStr(rs!INTE_sAL_AÑO) + ",'" + var_linea + "','" + var_catalogo + "'," + CStr(rs!floa_Sal_precio) + ")", cnn_intercompañias, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux4.Close
               End If
               rs.MoveNext
         Wend
         MsgBox "Se a terminado de cargar el archivo", vbOKOnly, "ATENCION"
      Else
         MsgBox "El movimiento no a sido impreso o no contiene información", vbOKOnly, "ATENCION"
      End If
      rs.Close
      Set cnn_intercompañias = Nothing
   End If
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
   Dim var_conexion_intercompañia As String
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
   If var_numero_folio > 0 Then
      If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
         Set reporte = appl.OpenReport(App.Path + "\rep_MOVIMIENTOS_SALIDA.rpt")
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
         var_ruta = App.Path
         Set var_tabla = CreateObject("ADODB.connection")
         var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=DSN=dBASE Files;DBQ=" & var_ruta & ";DefaultDir=" & var_ruta & ";DriverId=533;MaxBufferSize=2048;PageTimeout=5;"
          
         var_Archivo = Trim(var_almacen_Destino) + CStr(var_numero_folio)
         var_l = Len(var_Archivo)
         If var_l = 1 Then
            var_Archivo = "000000" + Trim(var_Archivo)
         End If
         If var_l = 2 Then
            var_Archivo = "00000" + Trim(var_Archivo)
         End If
         If var_l = 3 Then
            var_Archivo = "0000" + Trim(var_Archivo)
         End If
         If var_l = 4 Then
            var_Archivo = "000" + Trim(var_Archivo)
         End If
         If var_l = 5 Then
            var_Archivo = "00" + Trim(var_Archivo)
         End If
         If var_l = 6 Then
            var_Archivo = "0" + Trim(var_Archivo)
         End If
         var_archivo_2 = "t" + var_Archivo
      
         If Dir(var_ruta & "\" + var_archivo_2 + ".dbf") <> "" Then
            Kill var_ruta & "\" + var_archivo_2 + ".dbf"
         End If
         If Dir(var_ruta & "\" + var_Archivo + ".dbf") <> "" Then
            Kill var_ruta & "\" + var_Archivo + ".dbf"
         End If
         
            
         var_copia = CopyFile(var_ruta & "\TARCHENV.dbf", var_ruta & "\" + var_archivo_2 + ".dbf", 1)
         If var_copia = 1 Then
            Cadena = "select * from tb_salidas where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'"
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  var_codigo = rs!vcha_Art_articulo_id
                  rsaux3.Open "select VCHA_ART_CODIGO_EXTERNO from TB_aRTICULOS WHERE VCHA_ART_ARTICULO_ID = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux3.EOF Then
                     var_codigo = rsaux3!VCHA_aRT_CODIGO_EXTERNO
                  End If
                  rsaux3.Close
                  'rsaux2.Open "insert into " + var_archivo_2 + " (numnota,planta,codigo,descripcio,tallas,talla1,talla2,talla3,talla4,talla5,talla6,costo,cant1,cant2,cant3,cant4,cant5,cant6,anocosto) values ('" + Trim(var_almacen_Destino) + Trim(Str(var_numero_folio)) + "', '" + var_almacen_Destino + "', '" + var_codigo + "',' ', 1,0,0,0,0,0,0," + CStr(Round(rs!FLOA_SAL_COSTO, 2)) + ", " + CStr(Round(rs!floa_sal_cantidad, 2)) + ",0,0,0,0,0,'" + CStr(rs!INTE_sAL_AÑO) + "')", var_tabla, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
            Wend
            var_tabla.Close
            Set var_tabla = Nothing
         
            Cadena = "select * from tb_entradas where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'"
            var_archivo_origen = var_ruta & "\" + var_archivo_2 + ".dbf"
            var_archivo_destino = var_ruta & "\" + var_Archivo + ".dbf"
            var_eliminar = DeleteFile(var_ruta & "\" + var_Archivo + ".dbf")
            var_copia = CopyFile(var_archivo_origen, var_archivo_destino, 1)
            var_Archivo = var_almacen_origen + var_Archivo
            
            rs.Close
            rs.Open "select vcha_uor_mail from tb_unidadesorganizacionales where  vcha_uor_unidad_id = '" + txt_proveedor + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_correo_electronico = IIf(IsNull(rs!vcha_uor_mail), "", rs!vcha_uor_mail)
            Else
               var_correo_electronico = ""
            End If
            rs.Close
            If var_correo_electronico <> "" Then
               If MAPISession1.SessionID = 0 Then
                  MAPISession1.SignOn
               End If
               MAPIMessages1.SessionID = MAPISession1.SessionID
               MAPIMessages1.Compose
               MAPIMessages1.RecipDisplayName = var_correo_electronico
               MAPIMessages1.RecipAddress = var_correo_electronico
               MAPIMessages1.AddressResolveUI = True
               MAPIMessages1.ResolveName
               MAPIMessages1.MsgSubject = "Nota de envio " + Trim(var_Archivo)
               MAPIMessages1.MsgNoteText = "Se adjunta archivo con mercancia enviada"
               MAPIMessages1.AttachmentPathName = var_ruta & "\" + var_Archivo + ".dbf"
               MAPIMessages1.Send True
               If MAPISession1.SessionID > 0 Then
                  MAPISession1.SignOff
              End If
            Else
               MsgBox "No se a indicado una dirección de correo electrónico", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "No se a podido crear el archivo", vbOKOnly, "ATENCION"
         End If
      Else
         var_si = MsgBox("¿Se va a imprimir el movimiento?", vbOKCancel, "ATENCION")
         If var_si = 1 Then
            cnn.BeginTrans
            rsaux.Open "select max(inte_tem_consecutivo) from TB_TEMP_CODIGO_SALIDAS_INTERCOMPAÑIAS", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
            Else
               var_consecutivo = 0
            End If
            rsaux.Close
            var_consecutivo = var_consecutivo + 1
            rsaux1.Open "insert into TB_TEMP_CODIGO_SALIDAS_INTERCOMPAÑIAS (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            rs.Open "Select * from TB_UNIDADESORGANIZACIONALES where vcha_uor_unidad_id = '" + txt_proveedor + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_conexion = IIf(IsNull(rs!vcha_uor_conexion), "", rs!vcha_uor_conexion)
               If var_conexion <> "" Then
                  txt_nombre_proveedor = IIf(IsNull(rs!VCHA_UOR_NOMBRE), "", rs!VCHA_UOR_NOMBRE)
                  txt_codigo.Enabled = True
                  txt_codigo.SetFocus
                  txt_proveedor.Enabled = False
                  'If cnn_traspaso_intecomañia.State = 1 Then
                  '   cnn_traspaso_intecomañia.Close
                  Set cnn_traspaso_intecomañia = Nothing
                  'End If
                  Set cnn_traspaso_intecomañia = CreateObject("ADODB.connection")
                  cnn_traspaso_intecomañia.Open var_conexion
                  cnn_traspaso_intecomañia.CursorLocation = adUseClient
               Else
                  MsgBox "No existe conexión entre este almacén y la planta a donde se enviara la mercancía, consulte a sistemas", vbYesNo, "ATENCION"
               End If
            Else
               MsgBox "Clave de planta incorrecto", vbOKOnly, "ATENCION"
            End If
            rs.Close
            rs.Open "select vcha_Art_Articulo_id from tb_temporal_salidas where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  If rsaux2.State = 1 Then
                     rsaux2.Close
                  End If
                  rsaux2.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + rs!vcha_Art_articulo_id + "'", cnn_traspaso_intecomañia, adOpenDynamic, adLockOptimistic
                  If Not rsaux2.EOF Then
                      rsaux.Open "insert into TB_TEMP_CODIGO_SALIDAS_INTERCOMPAÑIAS (inte_tem_consecutivo, vcha_art_articulo_id, vcha_tem_codigo_externo) values (" + CStr(var_consecutivo) + ",'" + rs!vcha_Art_articulo_id + "','" + Trim(rsaux2!vcha_equ_codigo_equivalente) + "')", cnn, adOpenDynamic, adLockOptimistic
                      'MsgBox "insert into TB_TEMP_CODIGO_SALIDAS_INTERCOMPAÑIAS (inte_tem_consecutivo, vcha_art_articulo_id, vcha_tem_codigo_externo) values (" + CStr(var_consecutivo) + ",'" + rs!vcha_Art_articulo_id + "','" + Trim(rsaux2!VCHA_EQU_CODIGO_EQUIVALENTE) + "')"
                  Else
                      rsaux.Open "insert into TB_TEMP_CODIGO_SALIDAS_INTERCOMPAÑIAS (inte_tem_consecutivo, vcha_art_articulo_id, vcha_tem_codigo_externo) values (" + CStr(var_consecutivo) + ",'" + rs!vcha_Art_articulo_id + "','')", cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux2.Close
                  rs.MoveNext
            Wend
            rs.Close
            If rsaux10.State = 1 Then
               rsaux10.Close
            End If
              
            rsaux10.Open "select * from TB_TEMP_CODIGO_SALIDAS_INTERCOMPAÑIAS where vcha_tem_codigo_externo = ''", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux10.EOF Then
               cnn.BeginTrans
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
                  Cadena = "select * from tb_temporal_salidas with (nolock) where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio)
                  rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                  While Not rs.EOF
                        var_inserta = False
                        var_suma_cantidad = 0
                        var_cantidad_llegar = IIf(IsNull(rs!FLOA_sAL_cANTIDAD), 0, rs!FLOA_sAL_cANTIDAD)
                        var_cantidad = 0
                        While var_suma_cantidad < var_cantidad_llegar
                              rsaux2.Open "select * from tb_existencias where vcha_art_articulo_id =  '" + rs!vcha_Art_articulo_id + "' and vcha_alm_almacen_id = '" + rs!VCHA_ALM_ALMACEN_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux2.EOF Then
                                 If rsaux2!floa_exi_cantidad_2004 >= var_cantidad_llegar Then
                                    var_año = 2004
                                    var_suma_cantidad = var_cantidad_llegar
                                    var_cantidad = var_cantidad_llegar
                                    var_costo = IIf(IsNull(rsaux2!FLOA_EXI_COSTO_2004), 0, rsaux2!FLOA_EXI_COSTO_2004)
                                 Else
                                    var_cantidad_disponible = IIf(IsNull(rsaux2!floa_exi_cantidad_2004), 0, rsaux2!floa_exi_cantidad_2004)
                                    If var_cantidad_disponible > 0 Then
                                       var_año = 2004
                                       var_suma_cantidad = IIf(IsNull(rsaux2!floa_exi_cantidad_2004), 0, rsaux2!floa_exi_cantidad_2004)
                                       var_cantidad = IIf(IsNull(rsaux2!floa_exi_cantidad_2004), 0, rsaux2!floa_exi_cantidad_2004)
                                       var_costo = rsaux2!FLOA_EXI_COSTO_2004
                                    Else
                                       var_año = 2005
                                       var_cantidad = rs!FLOA_sAL_cANTIDAD - var_suma_cantidad
                                       var_suma_cantidad = var_cantidad_llegar
                                       var_costo = IIf(IsNull(rsaux2!floa_exi_costo_2005), 0, rsaux2!floa_exi_costo_2005)
                                    End If
                                 End If
                              Else
                                 var_año = 2005
                                 var_suma_cantidad = var_cantidad_llegar
                                 var_cantidad = var_cantidad_llegar
                                 rsaux4.Open "select * from tb_articulos where vcha_art_articulo_id =  '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                 If Not rsaux4.EOF Then
                                    var_costo = IIf(IsNull(rsaux4!mone_Art_costo_estandar), 0, rsaux4!mone_Art_costo_estandar)
                                 Else
                                    var_costo = 0
                                 End If
                                 rsaux4.Close
                              End If
                              rsaux2.Close
                              If var_costo = 0 Then
                                 rsaux4.Open "select floa_exi_costo_2005 from tb_existencias where vcha_alm_almacen_id = '8' and vcha_Art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                 If Not rsaux4.EOF Then
                                    var_costo = IIf(IsNull(rsaux4!floa_exi_costo_2005), 0, rsaux4!floa_exi_costo_2005)
                                 Else
                                    var_costo = 0
                                 End If
                                 rsaux4.Close
                              End If
                              rsaux.Open "insert into tb_salidas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_sal_numero, vcha_art_articulo_id, floa_sal_cantidad, floa_sal_costo, floa_sal_precio, inte_sal_año) values ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!vcha_uor_unidad_id + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_SAL_NUMERO) + ", '" + rs!vcha_Art_articulo_id + "', " + CStr(var_cantidad) + ", " + CStr(var_costo) + " , " + CStr(rs!floa_Sal_precio) + ", " + CStr(var_año) + ")", cnn, adOpenDynamic, adLockOptimistic
                        Wend
                        rs.MoveNext
                  Wend
                  rs.Close
                  var_estatus_movimiento = "I"
                  var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "", Now, 1)
                  var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "I", Now, 1)
               End If
               cnn.CommitTrans
               If var_posible_cerrar_KANBAN Then
                  txt_codigo.Enabled = False
                  txt_foco.Enabled = False
                  var_conexion_intercompañia = ""
                  rsaux.Open "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE VCHA_UOR_UNIDAD_ID = '" + Me.txt_proveedor + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     var_conexion_intercompañia = IIf(IsNull(rsaux!vcha_uor_conexion), "", rsaux!vcha_uor_conexion)
                  End If
                  rsaux.Close
                  If Trim(var_conexion_intercompañia) = "" Then
               
                     var_ruta = App.Path
                     Set var_tabla = CreateObject("ADODB.connection")
                     var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=DSN=dBASE Files;DBQ=" & var_ruta & ";DefaultDir=" & var_ruta & ";DriverId=533;MaxBufferSize=2048;PageTimeout=5;"
               
                     var_Archivo = Trim(var_almacen_Destino) + CStr(var_numero_folio)
                     var_l = Len(var_Archivo)
                     If var_l = 1 Then
                        var_Archivo = "000000" + Trim(var_Archivo)
                     End If
                     If var_l = 2 Then
                        var_Archivo = "00000" + Trim(var_Archivo)
                     End If
                     If var_l = 3 Then
                        var_Archivo = "0000" + Trim(var_Archivo)
                     End If
                     If var_l = 4 Then
                        var_Archivo = "000" + Trim(var_Archivo)
                     End If
                     If var_l = 5 Then
                        var_Archivo = "00" + Trim(var_Archivo)
                     End If
                     If var_l = 6 Then
                        var_Archivo = "0" + Trim(var_Archivo)
                     End If
                     var_archivo_2 = "t" + var_Archivo
                     
                     If Dir(var_ruta & "\" + var_archivo_2 + ".dbf") <> "" Then
                        Kill var_ruta & "\" + var_archivo_2 + ".dbf"
                     End If
                     If Dir(var_ruta & "\" + var_Archivo + ".dbf") <> "" Then
                        Kill var_ruta & "\" + var_Archivo + ".dbf"
                     End If
                
                     var_copia = CopyFile(var_ruta & "\TARCHENV.dbf", var_ruta & "\" + var_archivo_2 + ".dbf", 1)
                     If var_copia = 1 Then
                        Cadena = "select * from tb_salidas where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'"
                        If rs.State = 1 Then
                           rs.Close
                        End If
                        rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                        While Not rs.EOF
                              var_codigo = rs!vcha_Art_articulo_id
                              rsaux3.Open "select VCHA_EQU_CODIGO_EQUIVALENTE from tb_equivalencias WHERE VCHA_ART_ARTICULO_ID = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux3.EOF Then
                                 var_codigo = rsaux3!vcha_equ_codigo_equivalente
                              End If
                              rsaux3.Close
                              'Text1 = "insert into " + var_archivo_2 + " (numnota,planta,codigo,descripcio,tallas,talla1,talla2,talla3,talla4,talla5,talla6,costo,cant1,cant2,cant3,cant4,cant5,cant6,anocosto) values ('" + Str(var_numero_folio) + "', '" + var_almacen_destino + "', '" + var_codigo + "',' ', 1,0,0,0,0,0,0," + Str(rs!floa_sal_costo) + ", " + Str(rs!floa_sal_cantidad) + ",0,0,0,0,0,'" + CStr(rs!INTE_sal_AÑO) + "')"
                              rsaux2.Open "insert into " + var_archivo_2 + " (numnota,planta,codigo,descripcio,tallas,talla1,talla2,talla3,talla4,talla5,talla6,costo,cant1,cant2,cant3,cant4,cant5,cant6,anocosto) values ('" + Trim(var_almacen_Destino) + Trim(Str(var_numero_folio)) + "', '" + var_almacen_Destino + "', '" + var_codigo + "',' ', 1,0,0,0,0,0,0," + Str(rs!floa_Sal_costo) + ", " + Str(rs!FLOA_sAL_cANTIDAD) + ",0,0,0,0,0,'" + CStr(rs!INTE_sAL_AÑO) + "')", var_tabla, adOpenDynamic, adLockOptimistic
                              rs.MoveNext
                        Wend
                        var_tabla.Close
                        Set var_tabla = Nothing
                        
                        Cadena = "select * from tb_entradas where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'"
                        var_archivo_origen = var_ruta & "\" + var_archivo_2 + ".dbf"
                        var_archivo_destino = var_ruta & "\" + var_Archivo + ".dbf"
                        var_eliminar = DeleteFile(var_ruta & "\" + var_Archivo + ".dbf")
                        var_copia = CopyFile(var_archivo_origen, var_archivo_destino, 1)
                        var_Archivo = var_almacen_origen + var_Archivo
                        
                        rs.Close
                        rs.Open "select vcha_uor_mail from tb_unidadesorganizacionales where  vcha_uor_unidad_id = '" + txt_proveedor + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           var_correo_electronico = IIf(IsNull(rs!vcha_uor_mail), "", rs!vcha_uor_mail)
                        Else
                           var_correo_electronico = ""
                        End If
                        rs.Close
                        If var_correo_electronico <> "" Then
                           If MAPISession1.SessionID = 0 Then
                              MAPISession1.SignOn
                           End If
                           MAPIMessages1.SessionID = MAPISession1.SessionID
                           MAPIMessages1.Compose
                           MAPIMessages1.RecipDisplayName = var_correo_electronico
                           MAPIMessages1.RecipAddress = var_correo_electronico
                           MAPIMessages1.MsgSubject = "Nota de Envio " + Trim(var_Archivo)
                           MAPIMessages1.MsgNoteText = "Se adjunta archivo con mercancia enviada"
                           MAPIMessages1.AttachmentPathName = var_ruta & "\" + var_Archivo + ".dbf"
                           MAPIMessages1.Send True
                           If MAPISession1.SessionID > 0 Then
                             MAPISession1.SignOff
                           End If
                        Else
                           MsgBox "No se a indicado una dirección de correo electrónico", vbOKOnly, "ATENCION"
                        End If
                     Else
                        MsgBox "No se a podido crear el archivo", vbOKOnly, "ATENCION"
                     End If
                  Else
                     Set cnn_intercompañias = CreateObject("ADODB.connection")
                     'MsgBox txt_proveedor
                     cnn_intercompañias.Open var_conexion_intercompañia
                     cnn_intercompañias.CursorLocation = adUseClient
                     Cadena = "select * from tb_salidas where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'"
                     If rs.State = 1 Then
                        rs.Close
                     End If
                     rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                     While Not rs.EOF
                           var_codigo = rs!vcha_Art_articulo_id
                           rsaux3.Open "select * from tb_Articulos where vcha_art_Articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux3.EOF Then
                              var_DEscripcion = IIf(IsNull(rsaux3!vcha_art_nombre_español), "", rsaux3!vcha_art_nombre_español)
                           End If
                           rsaux3.Close
                           If rsaux2.State = 1 Then
                              rsuax2.Close
                           End If
                           rsaux5.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_ART_ARTICULO_ID = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                           var_linea = IIf(IsNull(rsaux5!vcha_lin_linea_id), "", rsaux5!vcha_lin_linea_id)
                           var_catalogo = IIf(IsNull(rsaux5!vcha_Art_catalogo_vigente), "0", rsaux5!vcha_Art_catalogo_vigente)
                           rsaux5.Close
                           If IsNumeric(var_almacen_Destino) Then
                              'MsgBox var_conexion_intercompañia
                               rsaux2.Open "INSERT INTO TB_ARCHIVOS_ENVIOS (VCHA_ACO_PROVEEDOR, INTE_ACO_NUMERO, VCHA_ACO_CODIGO_EXTERNO, vcha_aco_descripcion_externa, VCHA_ART_ARTICULO_ID, FLOA_ACO_CANTIDAD, FLOA_ACO_COSTO, INTE_ACO_AÑO, VCHA_LIN_LINEA_ID, VCHA_CAT_CATALOGO_ID, FLOA_ACO_PRECIO) Values ( '" + var_unidad_organizacional + "', '" + Trim(var_almacen_Destino) + Trim(CStr(var_numero_folio)) + "', '" + rs!vcha_Art_articulo_id + "','" + var_DEscripcion + "', ''," + CStr(rs!FLOA_sAL_cANTIDAD) + ", " + CStr(rs!floa_Sal_costo) + ", " + CStr(rs!INTE_sAL_AÑO) + ",'" + var_linea + "','" + var_catalogo + "'," + CStr(rs!floa_Sal_precio) + ")", cnn_intercompañias, adOpenDynamic, adLockOptimistic
                           Else
                              rsaux2.Open "INSERT INTO TB_ARCHIVOS_ENVIOS (VCHA_ACO_PROVEEDOR, INTE_ACO_NUMERO, VCHA_ACO_CODIGO_EXTERNO, vcha_aco_descripcion_externa, VCHA_ART_ARTICULO_ID, FLOA_ACO_CANTIDAD, FLOA_ACO_COSTO, INTE_ACO_AÑO, VCHA_LIN_LINEA_ID, VCHA_CAT_CATALOGO_ID, FLOA_ACO_PRECIO) Values ( '" + var_unidad_organizacional + "', '" + Trim(CStr(var_numero_folio)) + "', '" + rs!vcha_Art_articulo_id + "','" + var_DEscripcion + "', ''," + CStr(rs!FLOA_sAL_cANTIDAD) + ", " + CStr(rs!floa_Sal_costo) + ", " + CStr(rs!INTE_sAL_AÑO) + ",'" + var_linea + "','" + var_catalogo + "'," + CStr(rs!floa_Sal_precio) + ")", cnn_intercompañias, adOpenDynamic, adLockOptimistic
                           End If
                           rs.MoveNext
                     Wend
                     rs.Close
                     Set cnn_intercompañias = Nothing
                  End If
               
               
               
               
               
                  Set reporte = appl.OpenReport(App.Path + "\rep_MOVIMIENTOS_SALIDA.rpt")
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
                  MsgBox "No se pudo cerrar el movimiento Kanban", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "Existen códigos que no estan dados de alta en la planta de " + Me.txt_nombre_proveedor, vbOKOnly, "ATENCION"
            End If
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
   txt_total = "0"
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Function fun_copia_archivo(Origen, Destino)
    Copy_File = CopyFile(Origen, Destino, 1)
End Function

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
      If Me.frm_busqueda.Visible = True Then
         Me.frm_busqueda.Visible = False
      Else
         If Me.frm_eliminar.Visible = True Then
            Me.frm_eliminar.Visible = False
         Else
            If Me.frm_lista.Visible = True Then
               Me.frm_lista.Visible = False
            Else
               Unload Me
            End If
         End If
      End If
   End If
End Sub

Private Sub Form_Load()
   If var_empresa <> "18" Then
      Me.cmd_archivo_envios.Visible = False
   End If
   var_posible_kanban = 0
   txt_total = "0"
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
      Dim var_cantidad_total As Double
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
                  cnn.CommandTimeout = 360
                  rsaux.Open "select * from tb_temporal_salidas with (nolock) where inte_SAL_numero = " + txt_busqueda_folio + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                  var_cantidad_total = 0
                  If Not rsaux.EOF Then
                     While Not rsaux.EOF
                        rsaux2.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rsaux!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           Set list_item = lv_entradas.ListItems.Add(, , rsaux!vcha_Art_articulo_id)
                           list_item.SubItems(1) = IIf(IsNull(rsaux2(1).Value), "", rsaux2(1).Value)
                           list_item.SubItems(2) = IIf(IsNull(rsaux!FLOA_sAL_cANTIDAD), "", rsaux!FLOA_sAL_cANTIDAD)
                           var_cantidad_total = var_cantidad_total + IIf(IsNull(rsaux!FLOA_sAL_cANTIDAD), 0, rsaux!FLOA_sAL_cANTIDAD)
                           rsaux2.Close
                           rsaux.MoveNext:
                        End If
                     Wend
                  End If
                  rsaux.Close
                  If Me.lv_entradas.ListItems.Count > 13 Then
                     lv_entradas.ColumnHeaders(2).Width = 4685.22
                  Else
                     lv_entradas.ColumnHeaders(2).Width = 4885.22
                  End If
                  txt_total = CStr(var_cantidad_total)
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
   'Select Case KeyAscii
   'Case 48 To 57, 52, 13, 8, 46, 27
   'Case Else
   '    KeyAscii = 0
   'End Select
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
         Dim var_posible_eliminar As Boolean
         var_cantidad_eliminar = Val(txt_cantidad_eliminar)
         var_posible_eliminar = True
         If var_posible_eliminar = True Then
            var_inserta = False
            txt_total = CStr(CDbl(txt_total) - var_cantidad_eliminar)
            rsaux.Open "UPDATE TB_TEMPORAL_SALIDAS SET FLOA_SAL_CANTIDAD = ISNULL(FLOA_SAL_CANTIDAD,0) - " + txt_cantidad_eliminar + " WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_SAL_NUMERO = " + CStr(var_numero_folio) + " AND VCHA_ART_ARTICULO_ID= '" + lv_entradas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
            'var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_destino, var_clave_movimiento, var_numero_folio, lv_entradas.SelectedItem, 0 - Val(txt_cantidad_eliminar))
            lv_entradas.selectedItem.SubItems(2) = lv_entradas.selectedItem.SubItems(2) - Val(txt_cantidad_eliminar)
            var_renglon = lv_entradas.selectedItem.Index
            Call ilumina_grid
         Else
            MsgBox "La cantidad a eliminar supera a la cantidad asignada a la causa de devolución seleccionada", vbOKOnly, "ATENCION"
         End If
         var_ventana = 0
         frm_eliminar.Visible = False
         txt_codigo.SetFocus
      Else
         MsgBox "Cantidad incorrecta", vbOKOnly, "ATENCION"
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

Private Sub txt_foco_GotFocus()
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_SALIDAS_INSERTA = New TB_TEMPORAL_SALIDAS_INSERTA
   Set TB_TEMPORAL_SALIDAS_MODIFICA = New TB_TEMPORAL_SALIDAS_MODIFICA
   Set TB_ENCABEZADO_MOVIMIENTOS_I = New TB_ENCABEZADO_MOVIMIENTOS_I
   Dim var_inserta As Boolean
   If Trim(txt_codigo.Text) <> "" Then
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
         If var_costo = 0 Then
            rs.Open "select * from tb_existencias where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and vcha_Art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_costo = IIf(IsNull(rs!FLOA_eXI_COSTO), 0, rs!FLOA_eXI_COSTO)
            Else
               var_costo = 0
            End If
            rs.Close
         End If
         Cadena = "select * from tb_temporal_salidas with (nolock) where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = " + txt_codigo
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
            txt_total = CStr(CDbl(txt_total) + var_cantidad_leida)
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
            txt_total = CStr(CDbl(txt_total) + var_cantidad_leida)
            Call ilumina_grid
         End If
      Else
         frmmensaje.lbl_mensaje = var_kanban_mensaje
         frmmensaje.Show 1
         txt_codigo = ""
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

Private Sub txt_nombre_proveedor_KeyDown(KeyCode As Integer, Shift As Integer)
   If txt_proveedor.Enabled = True Then
      If KeyCode = 116 Then
         rs.Open "SELECT * FROM tb_unidadesorganizacionales ORDER BY VCHA_UOR_NOMBRE", cnn, adOpenDynamic, adLockOptimistic
         lv_lista.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!vcha_uor_unidad_id)
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


Private Sub txt_proveedor_KeyDown(KeyCode As Integer, Shift As Integer)
   If txt_proveedor.Enabled = True Then
      If KeyCode = 116 Then
         rs.Open "SELECT * FROM tb_unidadesorganizacionales ORDER BY VCHA_UOR_NOMBRE", cnn, adOpenDynamic, adLockOptimistic
         lv_lista.ListItems.Clear
         While Not rs.EOF
                
               Set list_item = lv_lista.ListItems.Add(, , rs!vcha_uor_unidad_id)
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
         rs.Open "Select * from TB_UNIDADESORGANIZACIONALES where vcha_uor_unidad_id = '" + txt_proveedor + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_conexion = IIf(IsNull(rs!vcha_uor_conexion), "", rs!vcha_uor_conexion)
            If var_conexion <> "" Then
               txt_nombre_proveedor = IIf(IsNull(rs!VCHA_UOR_NOMBRE), "", rs!VCHA_UOR_NOMBRE)
               txt_codigo.Enabled = True
               txt_codigo.SetFocus
               txt_proveedor.Enabled = False
               'If cnn_traspaso_intecomañia.State = 1 Then
               '   cnn_traspaso_intecomañia.Close
               Set cnn_traspaso_intecomañia = Nothing
               'End If
               Set cnn_traspaso_intecomañia = CreateObject("ADODB.connection")
               cnn_traspaso_intecomañia.Open var_conexion
               cnn_traspaso_intecomañia.CursorLocation = adUseClient

            Else
               MsgBox "No existe conexión entre este almacén y la planta a donde se enviara la mercancía, consulte a sistemas", vbYesNo, "ATENCION"
            End If
         Else
            MsgBox "Clave de planta incorrecto", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         MsgBox "Debe de indicar un proveedor", vbOKOnly, "ATENCION"
      End If
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
                        var_descripcion_articulo = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
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
                              var_descripcion_articulo = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
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
                  var_costo = IIf(IsNull(rs(3).Value), 0, rs(3).Value)
                  var_precio = rs(2).Value
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
                        var_costo = IIf(IsNull(rs(3).Value), 0, rs(3).Value)
                        var_precio = rs(2).Value
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

