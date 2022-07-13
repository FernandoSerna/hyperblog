VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmsalidas_crossdocking_cantia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Salidas por crossdocking"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8850
      Picture         =   "frmcrossdocking_cantia.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   390
      Picture         =   "frmcrossdocking_cantia.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   2325
      TabIndex        =   12
      Top             =   1050
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   30
         TabIndex        =   10
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
         TabIndex        =   13
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.CommandButton Command1 
      Height          =   315
      Left            =   60
      Picture         =   "frmcrossdocking_cantia.frx":0784
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Desmarcar Todos Alt + D"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   120
      Left            =   15
      TabIndex        =   14
      Top             =   285
      Width           =   9195
   End
   Begin VB.Frame Frame1 
      Height          =   6885
      Left            =   15
      TabIndex        =   15
      Top             =   360
      Width           =   9225
      Begin VB.TextBox txt_nombre_almacen_destino 
         Height          =   345
         Left            =   4230
         TabIndex        =   22
         Top             =   195
         Width           =   4440
      End
      Begin VB.TextBox txt_almacen_destino 
         Height          =   345
         Left            =   3480
         TabIndex        =   21
         Top             =   195
         Width           =   720
      End
      Begin VB.Frame Frame2 
         Height          =   30
         Left            =   30
         TabIndex        =   18
         Top             =   570
         Width           =   9180
      End
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1380
         Picture         =   "frmcrossdocking_cantia.frx":0886
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   150
         Width           =   330
      End
      Begin VB.CommandButton Command7 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   720
         Picture         =   "frmcrossdocking_cantia.frx":0A9C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Marcar (Enter)"
         Top             =   150
         Width           =   330
      End
      Begin VB.CommandButton Command8 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1050
         Picture         =   "frmcrossdocking_cantia.frx":0CE6
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   150
         Width           =   330
      End
      Begin VB.CommandButton Command9 
         Height          =   315
         Left            =   60
         Picture         =   "frmcrossdocking_cantia.frx":0DB8
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   150
         Width           =   330
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   390
         Picture         =   "frmcrossdocking_cantia.frx":0EBA
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   150
         Width           =   330
      End
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   5040
         TabIndex        =   16
         Top             =   3375
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
            Caption         =   "Cantidad "
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   17
            Top             =   15
            Width           =   2895
         End
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1710
         Picture         =   "frmcrossdocking_cantia.frx":10D0
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Pasar todos F4"
         Top             =   150
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_articulos 
         Height          =   6135
         Left            =   45
         TabIndex        =   9
         Top             =   645
         Width           =   9105
         _ExtentX        =   16060
         _ExtentY        =   10821
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   5997
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Costo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Precio"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Pasar"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "movimiento"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "numero"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   2865
         TabIndex        =   19
         Top             =   255
         Width           =   585
      End
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   0
      Top             =   45
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
            Picture         =   "frmcrossdocking_cantia.frx":11D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcrossdocking_cantia.frx":1AAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcrossdocking_cantia.frx":2386
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcrossdocking_cantia.frx":2922
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcrossdocking_cantia.frx":31FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcrossdocking_cantia.frx":3AD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcrossdocking_cantia.frx":43B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcrossdocking_cantia.frx":44C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcrossdocking_cantia.frx":45D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcrossdocking_cantia.frx":46E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcrossdocking_cantia.frx":47FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcrossdocking_cantia.frx":490C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   870
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
            Picture         =   "frmcrossdocking_cantia.frx":4A1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcrossdocking_cantia.frx":52F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcrossdocking_cantia.frx":5BD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcrossdocking_cantia.frx":616E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcrossdocking_cantia.frx":6A4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcrossdocking_cantia.frx":7324
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcrossdocking_cantia.frx":7BFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcrossdocking_cantia.frx":7D10
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcrossdocking_cantia.frx":7E22
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcrossdocking_cantia.frx":7F34
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcrossdocking_cantia.frx":8046
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcrossdocking_cantia.frx":8158
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   1890
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
   Begin VB.Label lbl_movimiento_entrada 
      Height          =   285
      Left            =   2550
      TabIndex        =   23
      Top             =   45
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label lbl_numero_entrada 
      Height          =   285
      Left            =   945
      TabIndex        =   20
      Top             =   30
      Visible         =   0   'False
      Width           =   1065
   End
End
Attribute VB_Name = "frmsalidas_crossdocking_cantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_aceptar_pedidos_Click()
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_SALIDAS_INSERTA = New TB_TEMPORAL_SALIDAS_INSERTA
   Set TB_TEMPORAL_SALIDAS_MODIFICA = New TB_TEMPORAL_SALIDAS_MODIFICA
   Set TB_ENCABEZADO_MOVIMIENTOS_I = New TB_ENCABEZADO_MOVIMIENTOS_I
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
   Set TB_TEMPORAL_ENTRADAS_INSERTA = New TB_TEMPORAL_ENTRADAS_INSERTA
   Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
   
   
   Dim var_conexion_intercompañia As String
   Dim var_inserta As Boolean
   Dim var_posible_articulos As Boolean
   Dim var_primera_vez  As Boolean
   Dim var_cantidad_leida As Double
   Dim var_costo As Double
   Dim var_precio As Double
   Dim var_almacen_destino_cross As String
   Dim var_clave_movimiento_cross As String
   Dim var_clave_moneda As String
   Dim txt_proveedor As String
   Dim txt_nombre_proveedor As String
   Dim var_numero_folio_cross  As Double
   Dim var_almacen_origen_cross As String
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
   
   var_posible_articulos = False
   If Me.txt_almacen_destino <> "" Then
      If Me.lv_articulos.ListItems.Count > 0 Then
         For var_j = 1 To Me.lv_articulos.ListItems.Count
             Me.lv_articulos.ListItems.Item(var_j).Selected = True
             If Me.lv_articulos.selectedItem.SubItems(5) <> "" Then
                If CDbl(Me.lv_articulos.selectedItem.SubItems(5)) > 0 Then
                   var_posible_articulos = True
                End If
             End If
         Next var_j
         If var_posible_articulos = True Then
            var_acepta_traspaso = 1
            If var_empresa = "31" Then
               var_almacen_traspaso_cantia = txt_almacen_destino
               var_numero_traspaso_cantia = var_numero_folio_cross
               frmpassword_traspasos_cantia.lbl_movimiento = var_clave_movimiento
               frmpassword_traspasos_cantia.Show 1
               If var_acepta_traspaso_global = 0 Then
                  var_acepta_traspaso = 0
               Else
                  var_acepta_traspaso = 1
               End If
            End If
            If var_acepta_traspaso = 1 Then
                      
            var_si = MsgBox("¿Desea cerrar el movimiento?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               var_si = MsgBox("Confirmar el cerrado del movimiento", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  var_primera_vez = True
                  var_almacen_destino_cross = Me.txt_almacen_destino
                  var_almacen_origen_cross = "PTVH"
                  var_clave_movimiento_cross = "T"
                  var_clave_moneda = "1"
                  txt_proveedor = ""
                  txt_nombre_proveedor = ""
                  For var_j = 1 To Me.lv_articulos.ListItems.Count
                      Me.lv_articulos.ListItems.Item(var_j).Selected = True
                      If Me.lv_articulos.selectedItem.SubItems(5) = "" Then
                         Me.lv_articulos.selectedItem.SubItems(5) = "0"
                      End If
                      If CDbl(Me.lv_articulos.selectedItem.SubItems(5)) > 0 Then
                         txt_codigo = Me.lv_articulos.selectedItem
                         var_cantidad_leida = CDbl(Me.lv_articulos.selectedItem.SubItems(5))
                         var_costo = CDbl(Me.lv_articulos.selectedItem.SubItems(3))
                         var_precio = CDbl(Me.lv_articulos.selectedItem.SubItems(4))
                         
                         If Trim(txt_codigo) <> "" Then
                            If var_primera_vez = True Then
                               var_inserta = False
                               var_insreta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen_cross, "T", Now, 0, 0, "", "", var_almacen_origen_cross, var_almacen_destino_cross, "", var_clave_usuario_global, fun_NombrePc, "", "", "", "", "", "", "", 0, 0, 0, "1", 1)
                               var_numero_folio_cross = var_numero_folio_regreso
                               var_primera_vez = False
                            End If
                            Cadena = "select * from tb_temporal_salidas with (nolock) where vcha_alm_almacen_id = '" + var_almacen_origen_cross + "' and  VCHA_MOV_MOVIMIENTO_ID = 'T' and inte_sal_numero = " + Str(var_numero_folio_cross) + " and vcha_art_articulo_id = '" + txt_codigo + "'"
                            rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                            If Not rs.EOF Then
                               var_inserta = False
                               var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_destino_cross, "T", var_numero_folio_cross, txt_codigo, var_cantidad_leida, var_año)
                               var_inserta = False
                               var_inserta = TB_TEMPORAL_SALIDAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen_cross, "T", var_numero_folio_cross, txt_codigo, var_cantidad_leida)
                               rs.Close
                            Else
                               var_inserta = False
                               var_inserta = TB_TEMPORAL_ENTRADAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_destino_cross, "T", var_numero_folio_cross, CStr(txt_codigo), var_cantidad_leida, var_costo, var_precio, "0", var_almacen_origen_cross, 2005)
                               var_inserta = False
                               var_inserta = TB_TEMPORAL_SALIDAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen_cross, "T", var_numero_folio_cross, CStr(txt_codigo), var_cantidad_leida, var_costo, var_precio, "0", 0, 0)
                               rs.Close
                            End If
                            rs.Open "UPDATE TB_TEMPORAL_ENTRADAS SET FLOA_ENT_CANTIDAD_CROSSDOCKING =  ISNULL(FLOA_ENT_CANTIDAD_CROSSDOCKING,0) + " + CStr(var_cantidad_leida) + " WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + Me.lv_articulos.selectedItem.SubItems(7) + "' AND INTE_ENT_NUMERO = " + Me.lv_articulos.selectedItem.SubItems(8) + " AND VCHA_aRT_aRTICULO_ID = '" + Me.lv_articulos.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                         End If
                      End If
                  Next var_j
                  Dim var_posible_Cantidad As Integer
                  var_posible_Cantidad = 1
                  var_cadena_articulos = ""
               
                  Cadena = "select * from tb_temporal_salidas with (nolock) where vcha_alm_almacen_id = '" + var_almacen_origen_cross + "' and  VCHA_MOV_MOVIMIENTO_ID = 'T' and inte_sal_numero = " + Str(var_numero_folio_cross) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'"
                  cnn.BeginTrans
                  rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                  While Not rs.EOF
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
                                    var_costo = rsaux2!FLOA_EXI_COSTO_2004
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
                                 If Not rsaux2.EOF Then
                                    var_costo = rsaux2!floa_exi_costo_2005
                                 Else
                                    rsaux5.Open "select FLOA_EXI_COSTO_2005 from tb_existencias where vcha_art_articulo_id = '" + rs!vcha_Art_articulo_id + "' and vcha_alm_almacen_id = '8'", cnn, adOpenDynamic, adLockOptimistic
                                    If Not rsaux5.EOF Then
                                       var_costo = rsaux5(0).Value
                                    Else
                                       var_costo = 0
                                    End If
                                    rsaux5.Close
                                 End If
                              End If
                              rsaux2.Close
                              rsaux4.Open "insert into tb_salidas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_sal_numero, vcha_art_articulo_id, floa_sal_cantidad, floa_sal_costo, floa_sal_precio, inte_sal_año) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_origen_cross + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_SAL_NUMERO) + ", '" + rs!vcha_Art_articulo_id + "', " + CStr(var_cantidad) + ", " + CStr(var_costo) + " , " + CStr(rs!floa_Sal_precio) + ", " + CStr(var_año) + ")", cnn, adOpenDynamic, adLockOptimistic
                              If var_almacen_origen = "AG" Then
                                 rsaux4.Open "insert into tb_entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_art_articulo_id, floa_ent_cantidad, floa_ent_costo, floa_ent_precio, INTE_ENT_AÑO, VCHA_ENT_ALMACEN_ORIGEN) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_destino_cross + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_SAL_NUMERO) + ", '" + rs!vcha_Art_articulo_id + "', " + CStr(var_cantidad) + ", 0 , " + CStr(rs!floa_Sal_precio) + ", " + CStr(var_año) + ", '" + var_almacen_origen_cross + "')", cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 rsaux4.Open "insert into tb_entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_art_articulo_id, floa_ent_cantidad, floa_ent_costo, floa_ent_precio, INTE_ENT_AÑO, VCHA_ENT_ALMACEN_ORIGEN) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_destino_cross + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_SAL_NUMERO) + ", '" + rs!vcha_Art_articulo_id + "', " + CStr(var_cantidad) + ", " + CStr(var_costo) + " , " + CStr(rs!floa_Sal_precio) + ", " + CStr(var_año) + ", '" + var_almacen_origen_cross + "')", cnn, adOpenDynamic, adLockOptimistic
                              End If
                         Wend
                         rs.MoveNext
                  Wend
                  rs.Close
                  var_estatus_movimiento = "I"
                  var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen_cross, "T", var_numero_folio_cross, "I", Now, 1)
                  cnn.CommitTrans
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
                  
                  If var_empresa = "31" Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_salidas_traspasos_cantia.rpt")
                  Else
                     Set reporte = appl.OpenReport(App.Path + "\rep_salidas_traspasos.rpt")
                  End If
                  reporte.RecordSelectionFormula = "{VW_SALIDAS_TRASPASOS.VCHA_EMO_ALMACEN_ORIGEN} = '" + var_almacen_origen_cross + "' and {VW_SALIDAS_TRASPASOS.INTE_EMO_NUMERO} = " + Str(var_numero_folio_cross) + " and {VW_SALIDAS_TRASPASOS.VCHA_MOV_MOVIMIENTO_ID} = 'T' and {VW_SALIDAS_TRASPASOS.vcha_emp_empresa_id} = '" + var_empresa + "' and {VW_SALIDAS_TRASPASOS.vcha_uor_unidad_id} = '" + var_unidad_organizacional + "'"
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Reporte de Movimientos"
                  frmvistasprevias.Show 1
                  Set reporte = Nothing
                  rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen_cross + "' and vcha_mov_movimiento_id = 'T' and inte_emo_numero = " + CStr(var_numero_folio_cross), cnn, adOpenDynamic, adLockOptimistic
                  Me.lv_articulos.ListItems.Clear
                  rs.Open "select * from tb_TEMPORAL_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' AND vcha_mov_movimiento_id = '" + Me.lbl_movimiento_entrada + "' and inte_ent_numero = " + Me.lbl_numero_entrada + " AND FLOA_ENT_CANTIDAD - ISNULL(FLOA_ENT_cANTIDAD_CROSSDOCKING,0) > 0", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     While Not rs.EOF
                           Set list_item = lv_articulos.ListItems.Add(, , Trim(rs!vcha_Art_articulo_id))
                           rsaux.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux.EOF Then
                              list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_art_nombre_español), "", rsaux!vcha_art_nombre_español)
                           End If
                           rsaux.Close
                           list_item.SubItems(2) = Format(rs!floa_ent_Cantidad - IIf(IsNull(rs!floa_ent_Cantidad_crossdockinG), 0, rs!floa_ent_Cantidad_crossdockinG), "###,###,##0.00")
                           list_item.SubItems(3) = Format(IIf(IsNull(rs!floa_ent_costo), 0, rs!floa_ent_costo), "###,###,##0.00")
                           list_item.SubItems(4) = Format(IIf(IsNull(rs!floa_ent_precio), 0, rs!floa_ent_precio), "###,###,##0.00")
                           list_item.SubItems(7) = Me.lbl_movimiento_entrada
                           list_item.SubItems(8) = Me.lbl_numero_entrada
                           rs.MoveNext
                     Wend
                  End If
                  rs.Close
                  
                  
                  
               End If
            End If
            Else
            MsgBox "No se puede cerrar el traspaso hasta que haya una autorización del almacén " + Me.txt_nombre_almacen_destino, vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "No se a seleccionado ningún artículo para el traspaso", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El movimiento no contiene artículos", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a seleccionado un almacén", vbOKOnly, "ATENCION"
   End If
End Sub


Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   Me.txt_almacen_destino = ""
   Me.txt_nombre_almacen_destino = ""
   
   var_todos_lineas = 0
   n = lv_articulos.ListItems.Count
   For i = 1 To n
      lv_articulos.ListItems.Item(i).Selected = True
      lv_articulos.selectedItem.SubItems(6) = ""
      lv_articulos.selectedItem.SubItems(5) = ""
      lv_articulos.ListItems.Item(i).Bold = False
      lv_articulos.ListItems.Item(i).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(4).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(5).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(6).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
   Next i
   lv_articulos.Refresh
   Me.txt_almacen_destino.SetFocus
End Sub

Private Sub Command10_Click()
   var_todos_lineas = 1
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   n = lv_articulos.ListItems.Count
   For i = 1 To n
       lv_articulos.ListItems.Item(i).SubItems(6) = "*"
       lv_articulos.ListItems.Item(i).Bold = True
       lv_articulos.ListItems.Item(i).ForeColor = &HFF0000
       lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
       lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
       lv_articulos.ListItems.Item(i).ListSubItems(3).Bold = True
       lv_articulos.ListItems.Item(i).ListSubItems(4).Bold = True
       lv_articulos.ListItems.Item(i).ListSubItems(5).Bold = True
       lv_articulos.ListItems.Item(i).ListSubItems(6).Bold = True
       lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
       lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
       lv_articulos.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
       lv_articulos.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
       lv_articulos.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
       lv_articulos.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
   Next
   lv_articulos.Refresh
End Sub

Private Sub Command2_Click()
   For var_j = 1 To lv_articulos.ListItems.Count
       Me.lv_articulos.ListItems.Item(var_j).Selected = True
       If Me.lv_articulos.selectedItem.SubItems(6) = "*" Then
          Me.lv_articulos.selectedItem.SubItems(5) = Format(Me.lv_articulos.selectedItem.SubItems(2), "###,###,##0.00")
          lv_articulos.selectedItem.SubItems(6) = ""
          lv_articulos.ListItems.Item(var_j).Bold = False
          Me.lv_articulos.selectedItem.SubItems(6) = ""
          lv_articulos.ListItems.Item(var_j).ForeColor = &H80000012
          lv_articulos.ListItems.Item(var_j).ListSubItems(1).Bold = False
          lv_articulos.ListItems.Item(var_j).ListSubItems(2).Bold = False
          lv_articulos.ListItems.Item(var_j).ListSubItems(3).Bold = False
          lv_articulos.ListItems.Item(var_j).ListSubItems(4).Bold = False
          lv_articulos.ListItems.Item(var_j).ListSubItems(5).Bold = False
          lv_articulos.ListItems.Item(var_j).ListSubItems(6).Bold = False
          lv_articulos.ListItems.Item(var_j).ListSubItems(1).ForeColor = &H80000012
          lv_articulos.ListItems.Item(var_j).ListSubItems(2).ForeColor = &H80000012
          lv_articulos.ListItems.Item(var_j).ListSubItems(3).ForeColor = &H80000012
          lv_articulos.ListItems.Item(var_j).ListSubItems(4).ForeColor = &H80000012
          lv_articulos.ListItems.Item(var_j).ListSubItems(5).ForeColor = &H80000012
          lv_articulos.ListItems.Item(var_j).ListSubItems(6).ForeColor = &H80000012
          lv_articulos.Refresh
       End If
   Next var_j
End Sub

Private Sub Command6_Click()
   If var_todos_lineas = 1 Then
   Else
         var_todos_lineas = 0
   End If
   n = lv_articulos.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_articulos.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_articulos.selectedItem.SubItems(6) = "" And var_rellena = True Then
         lv_articulos.selectedItem.SubItems(6) = "*"
         lv_articulos.ListItems.Item(i).Bold = True
         lv_articulos.ListItems.Item(i).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(5).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(6).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_articulos.selectedItem.SubItems(6) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_articulos.selectedItem.SubItems(6) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub Command7_Click()
   var_todos_lineas = 0
   i = lv_articulos.selectedItem.Index
   If lv_articulos.selectedItem.SubItems(6) = "*" Then
      lv_articulos.selectedItem.SubItems(6) = ""
      lv_articulos.ListItems.Item(i).Bold = False
      lv_articulos.ListItems.Item(i).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(4).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(5).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(6).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
      lv_articulos.Refresh
   Else
      lv_articulos.selectedItem.SubItems(6) = "*"
      lv_articulos.ListItems.Item(i).Bold = True
      lv_articulos.ListItems.Item(i).ForeColor = &HFF0000
      lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_articulos.ListItems.Item(i).ListSubItems(3).Bold = True
      lv_articulos.ListItems.Item(i).ListSubItems(4).Bold = True
      lv_articulos.ListItems.Item(i).ListSubItems(5).Bold = True
      lv_articulos.ListItems.Item(i).ListSubItems(6).Bold = True
      lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_articulos.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
      lv_articulos.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
      lv_articulos.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
      lv_articulos.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
      lv_articulos.Refresh
   End If
End Sub

Private Sub Command8_Click()
   If var_todos_lineas = 1 Then
   Else
        var_todos_lineas = 0
   End If
   n = lv_articulos.ListItems.Count
   For i = 1 To n
      lv_articulos.ListItems.Item(i).Selected = True
      If lv_articulos.selectedItem.SubItems(6) = "*" Then
         lv_articulos.selectedItem.SubItems(6) = ""
         lv_articulos.ListItems.Item(i).Bold = False
         lv_articulos.ListItems.Item(i).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(4).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(5).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(6).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
      Else
         lv_articulos.selectedItem.SubItems(6) = "*"
         lv_articulos.ListItems.Item(i).Bold = True
         lv_articulos.ListItems.Item(i).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(5).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(6).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub Command9_Click()
   var_todos_lineas = 0
   n = lv_articulos.ListItems.Count
   For i = 1 To n
      lv_articulos.ListItems.Item(i).Selected = True
      lv_articulos.selectedItem.SubItems(5) = ""
      lv_articulos.ListItems.Item(i).Bold = False
      lv_articulos.ListItems.Item(i).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(4).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(5).Bold = False
      lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
      lv_articulos.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
   Next i
   lv_articulos.Refresh
End Sub

Private Sub Form_Load()
   Me.frm_lista.Visible = False
   Me.frm_eliminar.Visible = False
End Sub

Private Sub lv_articulos_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      Me.txt_cantidad_eliminar = ""
      frm_eliminar.Visible = True
      txt_cantidad_eliminar.SetFocus
   End If
End Sub

Private Sub lv_articulos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      var_todos_lineas = 0
      i = lv_articulos.selectedItem.Index
      If lv_articulos.selectedItem.SubItems(6) = "*" Then
         lv_articulos.selectedItem.SubItems(6) = ""
         lv_articulos.ListItems.Item(i).Bold = False
         lv_articulos.ListItems.Item(i).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(4).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(5).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(6).Bold = False
         lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
         lv_articulos.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
         lv_articulos.Refresh
      Else
         lv_articulos.selectedItem.SubItems(6) = "*"
         lv_articulos.ListItems.Item(i).Bold = True
         lv_articulos.ListItems.Item(i).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(5).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(6).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
         lv_articulos.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
         lv_articulos.Refresh
      End If
   End If
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.lv_lista.ListItems.Count > 0 Then
         Me.txt_almacen_destino = Me.lv_lista.selectedItem
         Me.txt_nombre_almacen_destino = Me.lv_lista.selectedItem.SubItems(1)
         Me.txt_almacen_destino.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      Me.txt_almacen_destino.SetFocus
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub txt_almacen_destino_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      If var_tipo_permiso = 1 Then
         rs.Open "select distinct * from vw_almacen_permiso_1 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = 'T' order by vcha_alm_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
      Else
         rs.Open "select distinct * from vw_movimientos_almacenes where vcha_mov_movimiento_id = 'T' order by vcha_alm_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
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
   If KeyAscii = 13 Then
      Me.txt_nombre_almacen_destino.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_almacen_destino_LostFocus()
   If Trim(Me.txt_almacen_destino) <> "" Then
      rs.Open "select * from tb_almacenes where vcha_alm_almacen_id = '" + Me.txt_almacen_destino + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_almacen_destino = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
      Else
         MsgBox "El almacén no existe", vbOKOnly, "ATENCION"
         Me.txt_almacen_destino = ""
         Me.txt_nombre_almacen_destino = ""
      End If
      rs.Close
   End If
End Sub

Private Sub txt_cantidad_eliminar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_cantidad_eliminar) Then
         If CDbl(Me.lv_articulos.selectedItem.SubItems(2)) >= CDbl(Me.txt_cantidad_eliminar) Then
            Me.lv_articulos.selectedItem.SubItems(5) = Format(CDbl(Me.txt_cantidad_eliminar), "###,###,##0.00")
            Me.lv_articulos.SetFocus
         Else
            MsgBox "La cantidad debe ser menor o igual que a la cantidad del movimiento", vbOKOnly, "ATENCION"
            Me.lv_articulos.SetFocus
         End If
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_eliminar.Visible = False
   End If
End Sub

Private Sub txt_cantidad_eliminar_LostFocus()
   Me.frm_eliminar.Visible = False
End Sub

Private Sub txt_planta_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      rs.Open "SELECT * FROM tb_unidadesorganizacionales where vcha_emp_empresa_id = '" + var_empresa + "' ORDER BY VCHA_UOR_NOMBRE", cnn, adOpenDynamic, adLockOptimistic
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
End Sub



Private Sub txt_referencia_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_tipo_lista = 2
      Me.lv_lista.ListItems.Clear
      Set list_item = lv_lista.ListItems.Add(, , "3")
      list_item.SubItems(1) = "SERVICIO AL CLIENTE"
      Set list_item = lv_lista.ListItems.Add(, , "1")
      list_item.SubItems(1) = "TIENDA"
      Set list_item = lv_lista.ListItems.Add(, , "5")
      list_item.SubItems(1) = "EXHIBICION"
      Set list_item = lv_lista.ListItems.Add(, , "2")
      list_item.SubItems(1) = "ALMACEN GENERAL"
      Me.frm_lista.Visible = True
      Me.lv_lista.SetFocus
   End If
End Sub

Private Sub txt_referencia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.lv_articulos.ListItems.Count > 0 Then
         Me.lv_articulos.SetFocus
      End If
   Else
      If KeyAscii = 27 Then
         Unload Me
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txt_nombre_almacen_destino_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      If var_tipo_permiso = 1 Then
         rs.Open "select distinct * from vw_almacen_permiso_1 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by vcha_alm_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
      Else
         rs.Open "select distinct * from vw_movimientos_almacenes where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by vcha_alm_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
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
