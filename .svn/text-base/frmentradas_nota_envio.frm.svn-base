VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmentradas_nota_envio 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_busqueda 
      Height          =   960
      Left            =   465
      TabIndex        =   4
      Top             =   870
      Width           =   3135
      Begin VB.TextBox txt_busqueda_folio 
         Height          =   315
         Left            =   195
         TabIndex        =   5
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
         TabIndex        =   6
         Top             =   120
         Width           =   3060
      End
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   1320
      TabIndex        =   1
      Top             =   450
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   2
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
         TabIndex        =   3
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1230
      Index           =   0
      Left            =   5160
      TabIndex        =   21
      Top             =   1110
      Width           =   2370
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
         TabIndex        =   22
         Top             =   540
         Width           =   2280
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   1
         Left            =   30
         TabIndex        =   23
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   0
      Left            =   75
      TabIndex        =   20
      Top             =   570
      Width           =   7455
   End
   Begin VB.Frame Frame3 
      Height          =   1245
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   1110
      Width           =   4995
      Begin VB.TextBox txt_referencia 
         Height          =   315
         Left            =   1005
         MaxLength       =   50
         TabIndex        =   16
         Top             =   825
         Width           =   3885
      End
      Begin VB.TextBox txt_almacen 
         Height          =   315
         Left            =   1005
         TabIndex        =   15
         Top             =   480
         Width           =   780
      End
      Begin VB.TextBox txt_nombre_almacen 
         Height          =   315
         Left            =   1800
         TabIndex        =   14
         Top             =   480
         Width           =   3105
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   19
         Top             =   510
         Width           =   585
      End
      Begin VB.Label label 
         BackColor       =   &H8000000D&
         Caption         =   " Datos del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   18
         Top             =   120
         Width           =   4920
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Referencia:"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   17
         Top             =   885
         Width           =   825
      End
   End
   Begin VB.TextBox txt_foco 
      Height          =   285
      Left            =   8205
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2910
      Width           =   1125
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmentradas_nota_envio.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   720
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      Picture         =   "frmentradas_nota_envio.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Buscar Movimiento"
      Top             =   720
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   810
      Picture         =   "frmentradas_nota_envio.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   720
      Width           =   330
   End
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      Picture         =   "frmentradas_nota_envio.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Cancelar Movimiento"
      Top             =   720
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7170
      Picture         =   "frmentradas_nota_envio.frx":0408
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   720
      Width           =   330
   End
   Begin VB.CommandButton cdm_sonido 
      Caption         =   "Command2"
      Height          =   915
      Left            =   7425
      TabIndex        =   0
      Top             =   900
      Visible         =   0   'False
      Width           =   270
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
            Picture         =   "frmentradas_nota_envio.frx":0A42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_nota_envio.frx":131C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_nota_envio.frx":1BF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_nota_envio.frx":2192
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_nota_envio.frx":2A6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_nota_envio.frx":3348
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_nota_envio.frx":3C22
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_nota_envio.frx":3D34
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_nota_envio.frx":3E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_nota_envio.frx":3F58
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_nota_envio.frx":406A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_nota_envio.frx":417C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   75
      TabIndex        =   24
      Top             =   975
      Width           =   7455
   End
   Begin VB.Frame Frame2 
      Height          =   4965
      Left            =   120
      TabIndex        =   25
      Top             =   2295
      Width           =   7425
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   1785
         TabIndex        =   28
         Top             =   1755
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar 
            Height          =   330
            Left            =   60
            TabIndex        =   29
            Top             =   390
            Width           =   2745
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000D&
            Caption         =   "Cantidad a eliminar"
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   30
            Top             =   15
            Width           =   2895
         End
      End
      Begin MSComctlLib.ListView lv_entradas 
         Height          =   4065
         Left            =   45
         TabIndex        =   32
         Top             =   405
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   7170
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
         NumItems        =   5
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
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Costo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Precio"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.CommandButton Command1 
         Height          =   360
         Left            =   6945
         Picture         =   "frmentradas_nota_envio.frx":428E
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   585
         Width           =   390
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
         TabIndex        =   27
         Top             =   495
         Width           =   2640
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
         Left            =   5010
         TabIndex        =   31
         Top             =   555
         Width           =   1890
      End
      Begin VB.Label lbl_cantidad_total 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5190
         TabIndex        =   39
         Top             =   4530
         Width           =   1995
      End
      Begin VB.Label lbl_cantidad 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   4305
         TabIndex        =   35
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
         TabIndex        =   34
         Top             =   120
         Width           =   7350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código del Artículo:"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   675
         Width           =   1395
      End
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
      TabIndex        =   38
      Top             =   75
      Width           =   7335
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   645
      Left            =   6555
      TabIndex        =   37
      Top             =   75
      Visible         =   0   'False
      Width           =   765
      URL             =   "C:\sistemas\desarrollo\integral\Articulo no existe.wav"
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
      enabled         =   0   'False
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   1349
      _cy             =   1138
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp2 
      Height          =   870
      Left            =   0
      TabIndex        =   36
      Top             =   1425
      Visible         =   0   'False
      Width           =   435
      URL             =   "C:\sistemas\desarrollo\integral\type.wma"
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
      _cx             =   767
      _cy             =   1535
   End
End
Attribute VB_Name = "frmentradas_nota_envio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn_cantia As ADODB.Connection
Dim var_año As Integer
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
Dim var_renglon As Double

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


Private Sub cmb_almacen_destino_Click()
   var_almacen_Destino = Obtener_llave(cnn, rsaux, "TB_almacenes", "VCHA_ALM_NOMBRE", cmb_almacen_destino, 2, "T")
   txt_referencia.Enabled = True
End Sub

Private Sub cmb_almacen_destino_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_referencia.SetFocus
      cmb_almacen_destino.Enabled = False
   End If
End Sub

Private Sub cdm_sonido_2_Click()
   Me.wmp2.Controls.Play
End Sub

Private Sub cdm_sonido_Click()
   If Trim(Me.txt_codigo) <> "" Then
      wmp1.Controls.Play
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
   Set TB_EXISTENCIAS_INSERTA = New TB_EXISTENCIAS_INSERTA
   Set TB_ENTRADAS_I = New TB_ENTRADAS_I
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
   'If var_numero_folio > 0 Then
      If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
         Set reporte = appl.OpenReport(App.Path + "\rep_MOVIMIENTOS_ENTRADAS.rpt")
         reporte.RecordSelectionFormula = "{VW_MOVIMIENTOS_ENTRADA.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_MOVIMIENTOS_ENTRADA.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' AND {VW_MOVIMIENTOS_ENTRADA.INTE_EMO_NUMERO} = " + Str(var_numero_folio) + " AND {VW_MOVIMIENTOS_ENTRADA.VCHA_ALM_ALMACEN_ID} = '" + var_almacen_Destino + "' and {VW_MOVIMIENTOS_ENTRADA.VCHA_eMP_EMPRESA_ID} = '" + var_empresa + "'"
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
         var_si = MsgBox("¿Se va a imprimir el movimiento?", vbOKCancel, "ATENCION")
         If var_si = 1 Then
            If Me.lv_entradas.ListItems.Count > 0 Then
            rs.Open "select * from tb_encabezado_movimientos where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_emo_referencia = '" + Me.txt_referencia + "'", cnn, adOpenDynamic, adLockOptimistic
            If rs.EOF Then
               rs.Close
               Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
               Set TB_TEMPORAL_ENTRADAS_INSERTA = New TB_TEMPORAL_ENTRADAS_INSERTA
               Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
               Dim var_inserta As Boolean
               var_primera_vez = True
               For var_j = 1 To lv_entradas.ListItems.Count
                   Me.lv_entradas.ListItems.Item(var_j).Selected = True
                   Me.txt_codigo = Me.lv_entradas.selectedItem
                   If Trim(txt_codigo.Text) <> "" Then
                       bandera_suma = False
                       If var_primera_vez = True Then
                          var_inserta = False
                          var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, Now, var_numero_folio, 0, "", "", "", var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, 0, "", txt_referencia, "", "B", "", "", 0, 0, 0, var_clave_moneda, 0)
                          var_numero_folio = var_numero_folio_regreso
                          txt_folio = var_numero_folio
                          var_primera_vez = False
                       End If
                       Cadena = "select * from TB_TEMPORAL_ENTRADAS where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "'"
                       rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                       var_cantidad_leida = lv_entradas.selectedItem.SubItems(2)
                       var_costo = Me.lv_entradas.selectedItem.SubItems(3)
                       var_precio = Me.lv_entradas.selectedItem.SubItems(4)
                       If Not rs.EOF Then
                          var_inserta = False
                          var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida, var_año)
                          rs.Close
                       Else
                          var_inserta = False
                          var_inserta = TB_TEMPORAL_ENTRADAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida, var_costo, var_precio, "0", "", var_año)
                          rs.Close
                       End If
                    End If
               Next var_j
              
               
               
               cnn.BeginTrans
               Cadena = "select * from tb_temporal_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio)
               rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
               While Not rs.EOF
                     var_inserta = False
                     If rsaux.State = 1 Then
                        rsaux.Close
                     End If
                        rsaux.Open "insert into tb_entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_art_articulo_id, floa_ent_cantidad, floa_ent_costo, floa_ent_precio, INTE_ENT_AÑO) values ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!VCHA_UOR_UNIDAD_ID + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!inte_ent_numero) + ", '" + rs!vcha_Art_Articulo_id + "', " + CStr(rs!floa_ent_Cantidad) + ", " + CStr(rs!floa_ent_costo) + " , " + CStr(rs!floa_ent_precio) + ", " + CStr(var_año) + ")", cnn, adOpenDynamic, adLockOptimistic
                        rs.MoveNext
                  Wend
                  rs.Close
                  var_estatus_movimiento = "I"
                  var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "", Now, 1)
                  var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "I", Now, 1)
                  cnn.CommitTrans
                  Set reporte = appl.OpenReport(App.Path + "\rep_MOVIMIENTOS_ENTRADAS.rpt")
                  reporte.RecordSelectionFormula = "{VW_MOVIMIENTOS_ENTRADA.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_MOVIMIENTOS_ENTRADA.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' AND {VW_MOVIMIENTOS_ENTRADA.INTE_EMO_NUMERO} = " + Str(var_numero_folio) + " AND {VW_MOVIMIENTOS_ENTRADA.VCHA_ALM_ALMACEN_ID} = '" + var_almacen_Destino + "'  and {VW_MOVIMIENTOS_ENTRADA.VCHA_eMP_EMPRESA_ID} = '" + var_empresa + "'"
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Reporte de Movimientos"
                  frmvistasprevias.Show 1
                  Set reporte = Nothing
                  txt_codigo.Enabled = False
                  txt_foco.Enabled = False
                  rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
               Else
                  MsgBox "La nota  " + Me.txt_referencia + " fue leida en el movimiento " + CStr(rs!INTE_EMO_NUMERO), vbOKOnly, "ATENCION"
                  rs.Close
               End If
            Else
               MsgBox "La nota esta vacia", vbOKOnly, "ATENCION"
            End If
         End If
      End If
   'Else
   '   MsgBox "No se a seleccionado ningún movimiento", vbOKOnly, "ATENCION"
   'End If
End Sub

Private Sub cmd_nuevo_Click()
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
   txt_referencia = ""
   txt_referencia.Enabled = False
   txt_almacen.Enabled = True
   txt_almacen.SetFocus
   Me.lbl_cantidad_total = "0"
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   'On Error GoTo salir:
   strConnectionString = "DRIVER=Microsoft Excel Driver (*.xls); DBQ=c:\inventario.xls"
   rsaux2.Open "SELECT * FROM [entradas$]", strConnectionString
   'rsaux2.Open "SELECT dbo.TB_EXISTENCIAS.VCHA_ART_ARTICULO_ID AS CODIGO, dbo.TB_EXISTENCIAS.FLOA_EXI_CANTIDAD * -1 AS CANTIDAD, dbo.TB_ARTICULOS.MONE_ART_PRECIO_BASE AS PRECIO, dbo.TB_ARTICULOS.MONE_ART_COSTO_ESTANDAR AS COSTO FROM         dbo.TB_EXISTENCIAS INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_EXISTENCIAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID WHERE     (dbo.TB_EXISTENCIAS.VCHA_ALM_ALMACEN_ID = 'THCAN') AND FLOA_eXI_CANTIDAD < 0", cnn, adOpenDynamic, adLockOptimistic
   rs.Open "delete from TB_TEMP_ENTRADAS_SALIDAS_AJUSTES", cnn, adOpenDynamic, adLockOptimistic
   While Not rsaux2.EOF
         If Not IsNull(rsaux2!codigo) Then
            If rsaux2!Cantidad > 0 Then
               rsaux4.Open "select * from tb_articulos where vcha_Art_articulo_id = '" + CStr(rsaux2!codigo) + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux4.EOF Then
                  var_codigo = rsaux4!vcha_Art_Articulo_id
                  var_DEscripcion = IIf(IsNull(rsaux4!vcha_Art_nombre_español), "", rsaux4!vcha_Art_nombre_español)
                  var_costo = rsaux2!Costo
                  var_precio = rsaux4!mone_art_precio_base
                  var_cantidad = rsaux2!Cantidad
               Else
                  rsaux5.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + CStr(IIf(IsNull(rsaux2!codigo), "", rsaux2!codigo)) + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux5.EOF Then
                     rsaux6.Open "select * from tb_articulos where vcha_Art_articulo_id = '" + CStr(rsaux5!vcha_Art_Articulo_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux6.EOF Then
                        var_codigo = rsaux6!vcha_Art_Articulo_id
                        var_DEscripcion = IIf(IsNull(rsaux6!vcha_Art_nombre_español), "", rsaux6!vcha_Art_nombre_español)
                        var_costo = rsaux2!Costo
                        var_precio = rsaux2!Precio
                        var_cantidad = rsaux2!Cantidad
                     Else
                        var_codigo = rsaux2!codigo
                        var_DEscripcion = "-no-"
                        var_costo = 0
                        var_precio = 0
                        var_cantidad = 0
                     End If
                     rsaux6.Close
                  Else
                     var_codigo = rsaux2!codigo
                     var_DEscripcion = "-no-"
                     var_costo = 0
                     var_precio = 0
                     var_cantidad = 0
                  End If
                  rsaux5.Close
               End If
               rsaux4.Close
               rsaux.Open "INSERT INTO TB_TEMP_ENTRADAS_SALIDAS_AJUSTES (vcha_Art_articulo_id, vcha_art_descripcion, floa_tem_cantidad, floa_tem_costo, floa_tem_precio) VALUES ('" + CStr(var_codigo) + "','" + var_DEscripcion + "'," + CStr(var_cantidad) + "," + CStr(var_costo) + "," + CStr(var_precio) + ")", cnn, adOpenDynamic, adLockOptimistic
            End If
         End If
         rsaux2.MoveNext
   Wend
   rsaux2.Close
   
   
   rsaux9.Open "select * from TB_TEMP_ENTRADAS_SALIDAS_AJUSTES where vcha_Art_descripcion = '-no-'", cnn, adOpenDynamic, adLockOptimistic
   var_cadena = ""
   If Not rsaux9.EOF Then
      While Not rsaux9.EOF
            If var_cadena = "" Then
               var_cadena = var_cadena + IIf(IsNull(rsaux9!vcha_Art_Articulo_id), "", rsaux9!vcha_Art_Articulo_id)
            Else
               var_cadena = var_cadena + "," + IIf(IsNull(rsaux9!vcha_Art_Articulo_id), "", rsaux9!vcha_Art_Articulo_id)
            End If
            rsaux9.MoveNext
      Wend
   End If
   rsaux9.Close
   If rsaux8.State = 1 Then
      rsaux8.Close
   End If
   If var_cadena <> "" Then
      MsgBox "No existen los siguientes artículos " + var_cadena, vbOKOnly, "ATENCION"
   Else
      If Me.txt_almacen <> "" Then
         If Me.txt_nombre_almacen <> "" Then
            If Me.txt_referencia <> "" Then
               rsaux8.Open "SELECT * FROM TB_TEMP_ENTRADAS_SALIDAS_AJUSTES", cnn, adOpenDynamic, adLockOptimistic
               var_cantidad = 0
               While Not rsaux8.EOF
                     txt_codigo = IIf(IsNull(rsaux8!vcha_Art_Articulo_id), "", rsaux8!vcha_Art_Articulo_id)
                     var_costo = IIf(IsNull(rsaux8!floa_tem_costo), 0, rsaux8!floa_tem_costo)
                     var_precio = IIf(IsNull(rsaux8!floa_tem_Precio), 0, rsaux8!floa_tem_Precio)
                     var_descripcion_articulo = IIf(IsNull(rsaux8!vcha_art_descripcion), "", rsaux8!vcha_art_descripcion)
                     var_cantidad_leida = IIf(IsNull(rsaux8!floa_tem_cantidad), 0, rsaux8!floa_tem_cantidad)
                     var_cantidad = var_cantidad + var_cantidad_leida
                     Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
                     Set TB_TEMPORAL_ENTRADAS_INSERTA = New TB_TEMPORAL_ENTRADAS_INSERTA
                     Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
                     Dim var_inserta As Boolean
                     If Trim(txt_codigo.Text) <> "" Then
                        bandera_suma = False
                        If var_primera_vez = True Then
                           var_inserta = False
                           var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, Now, var_numero_folio, 0, "", "", "", var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, 0, "", txt_referencia, "", "B", "", "", 0, 0, 0, var_clave_moneda, 0)
                           var_numero_folio = var_numero_folio_regreso
                           txt_folio = var_numero_folio
                           var_primera_vez = False
                        End If
                        Cadena = "select * from TB_TEMPORAL_ENTRADAS where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "'"
                        rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           var_inserta = False
                           var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida, var_año)
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
                           var_inserta = TB_TEMPORAL_ENTRADAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida, var_costo, var_precio, "0", "", var_año)
                           rs.Close
                           Set list_item = lv_entradas.ListItems.Add(, , Trim(txt_codigo))
                           list_item.SubItems(1) = var_descripcion_articulo
                           list_item.SubItems(2) = var_cantidad_leida
                           var_renglon = lv_entradas.ListItems.Count
                           Call ilumina_grid
                        End If
                     End If
                     rsaux8.MoveNext
               Wend
               rsaux8.Close
               MsgBox "Se a terminado de cargar " + CStr(var_cantidad) + " piezas", vbOKOnly, "ATENCION"
            Else
               MsgBox "Falta agregar una referencia", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Falta indicar el almacén", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Falta indicar el almacén", vbOKOnly, "ATENCION"
      End If
   End If
   Exit Sub
salir:
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   If rsaux1.State = 1 Then
      rsaux1.Close
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
   If rsaux5.State = 1 Then
      rsaux5.Close
   End If
   If rsaux6.State = 1 Then
      rsaux6.Close
   End If
   If rsaux7.State = 1 Then
      rsaux7.Close
   End If
   If rsaux8.State = 1 Then
      rsaux8.Close
   End If
   If rsaux9.State = 1 Then
      rsaux9.Close
   End If
   If rsaux10.State = 1 Then
      rsaux10.Close
   End If
   If rsaux11.State = 1 Then
      rsaux11.Close
   End If
   If Err.Number = -2147217900 Then
      MsgBox "DEBE DE CREAR EL ARCHIVO DE EXCEL INVENTARIO Y ESTE DEBE DE CONTAR CON LA HOJA LLAMADA ENTRADAS", vbOKOnly, "ATENCION"
   Else
      If Err.Number = 3265 Then
         MsgBox "LOS NOMBRES DE LAS COLUMNAS DEBEN DE SER CODIGO Y CANTIDAD", vbOKOnly, "ATENCION"
      Else
         MsgBox "A surgido un error al cargar el archivo de entradas", vbOKOnly, "ATENCION"
      End If
   End If
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
   Set cnn_cantia = CreateObject("ADODB.connection")
   rs.Open "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_conexion_cantia = IIf(IsNull(rs!vcha_uor_conexion), "", rs!vcha_uor_conexion)
   Else
      var_conexion_cantia = ""
   End If
   rs.Close
   If var_conexion_cantia <> "" Then
      cnn_cantia.Open var_conexion_cantia
   End If
   'If var_clave_movimiento = "EA" Then
   '   Me.Command1.Visible = True
   'Else
   '   Me.Command1.Visible = False
   'End If
   var_año = 2005
   var_numero_folio = 0
   var_cadena_seguridad = ""
   Top = 0
   Left = 1500
   frm_lista.Visible = False
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
   txt_referencia.Enabled = False
   txt_codigo.Enabled = False
   var_primera_vez = True
   var_cantidad_leida = 1#
End Sub


Private Sub Form_Unload(Cancel As Integer)
   If var_numero_folio > 0 Then
      rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_EMO_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
   End If
   Call activa_forma(var_activa_forma_entradas_sin_comparacion)
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

Private Sub Text1_Change()

End Sub


Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim var_n As Integer
      txt_almacen = lv_lista.selectedItem
      txt_nombre_almacen = lv_lista.selectedItem.SubItems(1)
      txt_almacen.SetFocus
      frm_lista.Visible = False
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
         rs.Open "select * from vw_movimientos_almacenes where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_Alm_almacen_id = '" + txt_almacen + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      End If
      If Not rs.EOF Then
         var_almacen_Destino = txt_almacen
         txt_nombre_almacen = rs!VCHA_ALM_NOMBRE
         txt_referencia.Enabled = True
         txt_almacen.Enabled = False
      Else
         var_almacen_Destino = ""
         txt_nombre_almacen = ""
         MsgBox "Clave de almacen incorrecto", vbOKOnly, "ATENCION"
      End If
      rs.Close
   End If
End Sub

Private Sub txt_busqueda_folio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(txt_busqueda_folio) <> "" Then
         If var_numero_folio = CDbl(txt_busqueda_folio) Then
            rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_EMO_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
         End If
         rs.Open "select * from tb_encabezado_movimientos where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emo_numero = " + txt_busqueda_folio + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            If var_numero_folio > 0 Then
               rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_EMO_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
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
                  txt_referencia = IIf(IsNull(rs!vcha_Emo_referencia), "", rs!vcha_Emo_referencia)
                  txt_referencia.Enabled = False
                  lv_entradas.ListItems.Clear
                  var_primera_vez = False
                  txt_almacen.Enabled = False
                  var_numero_folio = rs!INTE_EMO_NUMERO
                  txt_folio = var_numero_folio
                  rsaux.Open "select * from tb_ALMACENES where vcha_alm_almacen_id = '" + var_almacen_Destino + "'", cnn, adOpenDynamic, adLockOptimistic
                  txt_almacen_destino = rsaux(3).Value
                  txt_almacen = rsaux!VCHA_ALM_ALMACEN_ID
                  txt_nombre_almacen = rsaux(3).Value
                  rsaux.Close
                  rsaux.Open "select * from tb_temporal_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and inte_ent_numero = " + txt_busqueda_folio + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                  Me.lbl_cantidad_total = "0.00"
                  If Not rsaux.EOF Then
                     While Not rsaux.EOF
                        rsaux2.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rsaux!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           Set list_item = lv_entradas.ListItems.Add(, , rsaux!vcha_Art_Articulo_id)
                           list_item.SubItems(1) = IIf(IsNull(rsaux2(1).Value), "", rsaux2(1).Value)
                           list_item.SubItems(2) = IIf(IsNull(rsaux!floa_ent_Cantidad), "", rsaux!floa_ent_Cantidad)
                           Me.lbl_cantidad_total = Format(Me.lbl_cantidad_total, "###,###,##0.00")
                           rsaux2.Close
                           rsaux.MoveNext:
                        End If
                     Wend
                  End If
                  rsaux.Close
                  rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 1 WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_EMO_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
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
               MsgBox "El movimiento esta siendo usudo por otro usuario", vbOKOnly, "ATENCION"
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
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If IsNumeric(txt_cantidad_eliminar) Then
         Dim var_posible_eliminar As Boolean
         Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
         var_cantidad_eliminar = Val(txt_cantidad_eliminar)
         var_posible_eliminar = True
         If var_cantidad_eliminar > (lv_entradas.selectedItem.SubItems(2) * 1) Then
            var_posible_eliminar = False
         End If
         If var_posible_eliminar = True Then
            var_inserta = False
            var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, lv_entradas.selectedItem, 0 - Val(txt_cantidad_eliminar), var_año)
            lv_entradas.selectedItem.SubItems(2) = lv_entradas.selectedItem.SubItems(2) - Val(txt_cantidad_eliminar)
            var_renglon = lv_entradas.selectedItem.Index
            Call ilumina_grid
         Else
           MsgBox "La cantidad a eliminar supera a la cantidad asignada a la causa de devolución seleccionada", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Cantidad incorrecta", vbOKOnly, "ATENCION"
      End If
      var_ventana = 0
      frm_eliminar.Visible = False
      txt_codigo.SetFocus
   End If
   If KeyAscii = 27 Then
      var_ventana = 0
      frm_eliminar.Visible = False
      txt_codigo.SetFocus
   End If
End Sub

Private Sub txt_cantidad_GotFocus()
   txt_cantidad = ""
End Sub

Private Sub txt_cantidad_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
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
   Dim var_caja As String
   Dim var_cantidad_caja As Integer
   Dim var_recontable As Integer
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   txt_codigo = Trim(txt_codigo)
   If KeyAscii = 13 Then
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
      If var_empresa = "31" Then
         var_verificador = True
      End If
      If var_verificador = True Then
         var_caja = Left(txt_codigo, 6)
         'If var_caja = "000005" Or var_caja = "000010" Or var_caja = "000015" Or var_caja = "000020" Then
         If var_caja = "000005" Or var_caja = "000010" Or var_caja = "000015" Or var_caja = "000020" Or var_caja = "000001" Or var_caja = "000002" Or var_caja = "000003" Or var_caja = "000004" Or var_caja = "000006" Or var_caja = "000007" Or var_caja = "000008" Or var_caja = "000009" Or var_caja = "000011" Or var_caja = "0000012" Or var_caja = "0000013" Or var_caja = "0000014" Or var_caja = "000015" Or var_caja = "000016" Or var_caja = "000017" Or var_caja = "000018" Or var_caja = "000019" Or var_caja = "000021" Or var_caja = "000022" Or var_caja = "000023" Or var_caja = "000024" Or var_caja = "000025" Or var_caja = "000026" Or var_caja = "000027" Or var_caja = "000028" Or var_caja = "000029" Or var_caja = "000030" Then
            var_cantidad_caja = CInt(var_caja)
            txt_codigo = Mid(txt_codigo, 7, 5)
         End If
         var_costo = 0
         var_precio = 0
         If Trim(txt_codigo) <> "" Then
            rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               If var_clave_movimiento = "EA" Then
                  var_recontable = 1
               Else
                  If IsNull(rs(43).Value) Then
                     var_recontable = 0
                  Else
                     var_recontable = rs(43).Value
                  End If
               End If
               var_descripcion_articulo = rs(1).Value
               If rsaux4.State = 1 Then
                  rsaux4.Close
               End If
               rsaux4.Open "SELECT * FROM TB_EXISTENCIAS WHERE VCHA_ART_ARTICULO_ID = '" + txt_codigo + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux4.EOF Then
                  If var_clave_movimiento = "EPTM" Then
                     var_costo = IIf(IsNull(rs(3).Value), 0, rs(3).Value)
                  Else
                     var_costo = IIf(IsNull(rsaux4!FLOA_eXI_COSTO), 0, rsaux4!FLOA_eXI_COSTO)
                  End If
               Else
                  var_costo = IIf(IsNull(rs(3).Value), 0, rs(3).Value)
               End If
               rsaux4.Close
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
                        If var_clave_movimiento = "EA" Then
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
                     rsaux4.Open "SELECT * FROM TB_EXISTENCIAS WHERE VCHA_ART_ARTICULO_ID = '" + txt_codigo + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux4.EOF Then
                        If var_clave_movimiento = "EPTM" Then
                           var_costo = IIf(IsNull(rs(3).Value), 0, rs(3).Value)
                        Else
                           var_costo = IIf(IsNull(rsaux4!FLOA_eXI_COSTO), 0, rsaux4!FLOA_eXI_COSTO)
                        End If
                     Else
                        var_costo = IIf(IsNull(rs(3).Value), 0, rs(3).Value)
                     End If
                     rsaux4.Close
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
                     Call cdm_sonido_Click
                     txt_codigo = ""
                     frmmensaje.lbl_mensaje = "El artículo no existe"
                     frmmensaje.Show
                     'MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
                  End If
               Else
                  Call cdm_sonido_Click
                  txt_codigo = ""
                  frmmensaje.lbl_mensaje = "El artículo no existe"
                  frmmensaje.Show
                  'MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
                  rs.Close
               End If
            End If
         Else
         End If
      Else
         Call cdm_sonido_Click
         txt_codigo = ""
         frmmensaje.lbl_mensaje = "Error en Código"
         frmmensaje.Show
         'MsgBox "Error en Código", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_foco_GotFocus()
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_ENTRADAS_INSERTA = New TB_TEMPORAL_ENTRADAS_INSERTA
   Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
   Dim var_inserta As Boolean
   If Trim(txt_codigo.Text) <> "" Then
      bandera_suma = False
      If var_primera_vez = True Then
         var_inserta = False
         var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, Now, var_numero_folio, 0, "", "", "", var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, 0, "", txt_referencia, "", "B", "", "", 0, 0, 0, var_clave_moneda, 0)
         var_numero_folio = var_numero_folio_regreso
         txt_folio = var_numero_folio
         var_primera_vez = False
      End If
      Cadena = "select * from TB_TEMPORAL_ENTRADAS where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "'"
      rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Call cdm_sonido_2_Click
         var_inserta = False
         var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida, var_año)
         rs.Close
         valor = Trim(txt_codigo)
         Set itmfound = lv_entradas.findItem(valor, lvwText, , lvwPartial)
         itmfound.EnsureVisible
         itmfound.Selected = True
         lv_entradas.selectedItem.SubItems(2) = lv_entradas.selectedItem.SubItems(2) + var_cantidad_leida
         var_renglon = lv_entradas.selectedItem.Index
         Call ilumina_grid
      Else
         Call cdm_sonido_2_Click
         var_inserta = False
         var_inserta = TB_TEMPORAL_ENTRADAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida, var_costo, var_precio, "0", "", var_año)
         rs.Close
         Set list_item = lv_entradas.ListItems.Add(, , Trim(txt_codigo))
         list_item.SubItems(1) = var_descripcion_articulo
         list_item.SubItems(2) = var_cantidad_leida
         var_renglon = lv_entradas.ListItems.Count
         Call ilumina_grid
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
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Trim(txt_almacen) <> "" Then
         If txt_referencia.Enabled = True Then
            txt_referencia.SetFocus
         End If
      End If
   End If
End Sub

Private Sub txt_referencia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      'If Me.txt_establecimiento <> "" Then
         Me.lbl_cantidad_total = "0"
         Me.lv_entradas.ListItems.Clear
         Dim list_item As ListItem
         If IsNumeric(Me.txt_referencia) Then
            If rsaux11.State = 1 Then
               rsaux11.Close
            End If
            'rsaux11.Open "select VCHA_PRO_PRODUCTO_ID, SUM(BINT_DNO_CANTIDAD) AS BINT_DNO_CANTIDAD  from tb_dnotas where bint_not_nota_id = " + Me.txt_referencia + " GROUP BY VCHA_PRO_PRODUCTO_ID", cnn_cantia, adOpenDynamic, adLockOptimistic
            'var_cadena = "SELECT  d .VCHA_PRO_PRODUCTO_ID, CASE p.mon_pro_costorea WHEN 0 THEN p.mon_pro_costo ELSE p.mon_pro_costorea END floa_com_costo, SUM(d .bint_dno_cantidad) AS floa_com_cantidad, bint_not_nota_id FROM tb_producto p, tb_dnotas d WHERE p.vcha_pro_producto_id = d .vcha_pro_producto_id and bint_not_nota_id = " + Me.txt_referencia + " GROUP BY d .VCHA_PRO_PRODUCTO_ID, mon_pro_costorea, mon_pro_costo, bint_not_nota_id"
            If var_empresa = "15" Then
               var_cadena = " SELECT d.VCHA_PRO_PRODUCTO_ID, p.FLOA_CTO_UNIT AS floa_com_costo, SUM(d.BINT_DNO_CANTIDAD) AS floa_com_cantidad, d.BINT_NOT_NOTA_ID, C.VCHA_NOT_CLASIFICACION , p.VCHA_LOT_LOTE_ID FROM dbo.TB_LOTES p INNER JOIN dbo.TB_DNOTAS d ON p.VCHA_PRO_PRODUCTO_ID = d.VCHA_PRO_PRODUCTO_ID AND p.VCHA_LOT_LOTE_ID = d.VCHA_LOT_LOTE_ID INNER JOIN dbo.TB_NOTAS C ON d.BINT_NOT_NOTA_ID = C.BINT_NOT_NOTA_ID Where (D.BINT_NOT_NOTA_ID = " + Me.txt_referencia + ") GROUP BY d.VCHA_PRO_PRODUCTO_ID, p.FLOA_CTO_UNIT, d.BINT_NOT_NOTA_ID, C.VCHA_NOT_CLASIFICACION, p.VCHA_LOT_LOTE_ID"
            Else
               var_cadena = " SELECT d .VCHA_PRO_PRODUCTO_ID, CASE p.mon_pro_costorea WHEN 0 THEN p.mon_pro_costo ELSE p.mon_pro_costorea END floa_com_costo, SUM(d .bint_dno_cantidad) AS floa_com_cantidad, D .bint_not_nota_id, VCHA_NOT_CLASIFICACION FROM tb_producto p, tb_dnotas d, TB_NOTAS C WHERE     p.vcha_pro_producto_id = d .vcha_pro_producto_id AND D .bint_not_nota_id = " + Me.txt_referencia + " AND C.BINT_NOT_NOTA_ID = D .BINT_NOT_NOTA_ID GROUP BY d .VCHA_PRO_PRODUCTO_ID, mon_pro_costorea, mon_pro_costo, D .bint_not_nota_id, VCHA_NOT_CLASIFICACION"
            End If
            'MsgBox cnn_cantia.ConnectionString
            
            rsaux11.Open var_cadena, cnn_cantia, adOpenDynamic, adLockOptimistic
            If Not rsaux11.EOF Then
               VAR_SI_aRTICULOS = ""
               var_clasificacion = rsaux11!vcha_not_clasificacion
               While Not rsaux11.EOF
                     'rs.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + rsaux11!vcha_pro_producto_id + "'", cnn, adOpenDynamic, adLockOptimistic
                     If var_empresa = "15" Then
                        rs.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_ART_ARTICULO_ID = '" + rsaux11!vcha_pro_producto_id + "'", cnn, adOpenDynamic, adLockOptimistic
                     Else
                        rs.Open "SELECT * FROM TB_ARTICULOS WHERE SUBSTRING(VCHA_ART_ARTICULO_ID,7,5) = '" + rsaux11!vcha_pro_producto_id + "'", cnn, adOpenDynamic, adLockOptimistic
                     End If
                     If rs.EOF Then
                        If VAR_SI_aRTICULOS = "" Then
                           VAR_SI_aRTICULOS = rsaux11!vcha_pro_producto_id
                        Else
                           VAR_SI_aRTICULOS = VAR_SI_aRTICULOS + ", " + rsaux11!vcha_pro_producto_id
                        End If
                     End If
                     rs.Close
                     rsaux11.MoveNext
               Wend
               
               'rs.Open "select * from tb_clientes where vcha_cli_clave_id = '" + Me.txt_almacen_destino + "'", cnn, adOpenDynamic, adLockOptimistic
               'If Not rs.EOF Then
               '   var_lista_precios = IIf(IsNull(rs!vcha_lis_lista_id), "", rs!vcha_lis_lista_id)
               'End If
               'rs.Close
               
               VAR_SI_aRTICULOS_precios = ""
               'rsaux11.MoveFirst
               'While Not rsaux11.EOF
               '      'rs.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + rsaux11!vcha_pro_producto_id + "'", cnn, adOpenDynamic, adLockOptimistic
               '      If var_empresa = "15" Then
               '         rs.Open "SELECT * FROM tb_Detalle_lista_precios WHERE VCHA_ART_ARTICULO_ID = '" + rsaux11!vcha_pro_producto_id + "' and vcha_lis_lista_precios_id = '" + var_lista_precios + "'", cnn, adOpenDynamic, adLockOptimistic
               '      Else
               '         rs.Open "SELECT * FROM tb_Detalle_lista_precios WHERE SUBSTRING(VCHA_ART_ARTICULO_ID,7,5) = '" + rsaux11!vcha_pro_producto_id + "' and vcha_lis_lista_precios_id = '" + var_lista_precios + "'", cnn, adOpenDynamic, adLockOptimistic
               '      End If
               '      If rs.EOF Then
               '         If VAR_SI_aRTICULOS_precios = "" Then
               '            VAR_SI_aRTICULOS_precios = rsaux11!vcha_pro_producto_id
               '         Else
               '            VAR_SI_aRTICULOS_precios = VAR_SI_aRTICULOS_precios + ", " + rsaux11!vcha_pro_producto_id
               '         End If
               '      End If
               '      rs.Close
               '      rsaux11.MoveNext
               'Wend
               
               If VAR_SI_aRTICULOS_precios = "" Then
                  If VAR_SI_aRTICULOS = "" Then
                     rsaux11.MoveFirst
                     While Not rsaux11.EOF
                           'rs.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + rsaux11!vcha_pro_producto_id + "'", cnn, adOpenDynamic, adLockOptimistic
                           If var_empresa = "15" Then
                              rsaux2.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_aRT_ARTICULO_ID = '" + rsaux11!vcha_pro_producto_id + "'", cnn, adOpenDynamic, adLockOptimistic
                           Else
                              rsaux2.Open "SELECT * FROM TB_ARTICULOS WHERE SUBSTRING(VCHA_aRT_ARTICULO_ID,7,5) = '" + rsaux11!vcha_pro_producto_id + "'", cnn, adOpenDynamic, adLockOptimistic
                           End If
                           If Not rsaux2.EOF Then
                              var_nombre_articulo = IIf(IsNull(rsaux2!vcha_Art_nombre_español), "", rsaux2!vcha_Art_nombre_español)
                              'var_costo = IIf(IsNull(rsaux2!mone_art_costo_estandar), 0, rsaux2!mone_art_costo_estandar)
                           Else
                              var_nombre_articulo = ""
                              'var_costo = 0
                           End If
                           
                           var_costo = IIf(IsNull(rsaux11!FLOA_COM_COSTO), 0, rsaux11!FLOA_COM_COSTO)
                           rsaux10.Open "select mone_Art_precio_base from tb_Articulos where vcha_Art_articulo_id = '" + rsaux2!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux10.EOF Then
                              var_precio = IIf(IsNull(rsaux10!mone_art_precio_base), 0, rsaux10!mone_art_precio_base)
                           Else
                              var_precio = 0
                           End If
                           rsaux10.Close
                           Set list_item = Me.lv_entradas.ListItems.Add(, , rsaux2!vcha_Art_Articulo_id)
                           rsaux2.Close
                           list_item.SubItems(1) = var_nombre_articulo
                           list_item.SubItems(2) = rsaux11!floa_com_cantidad
                           list_item.SubItems(3) = var_costo
                           list_item.SubItems(4) = var_precio
                           Me.lbl_cantidad_total = Format(CDbl(Me.lbl_cantidad_total) + rsaux11!floa_com_cantidad, "###,###,##0.00")
                           'rs.Close
                           rsaux11.MoveNext
                     Wend
                     If Me.lv_entradas.ListItems.Count > 18 Then
                        Me.lv_entradas.ColumnHeaders.Item(3).Width = 800
                     Else
                        Me.lv_entradas.ColumnHeaders.Item(3).Width = 1000
                     End If
                     If Me.lv_entradas.ListItems.Count > 0 Then
                        Me.lv_entradas.SetFocus
                     End If
                  Else
                     MsgBox "Los siguientes codigos no tienen equivalencia en el S.I.D. " + VAR_SI_aRTICULOS, vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "Los siguientes codigos no estan en la lista de precios asignada al cliente " + VAR_SI_aRTICULOS_precios, vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "La nota no existe", vbOKOnly, "ATENCION"
            End If
            rsaux11.Close
         Else
            MsgBox "Nota de envio incorrecta", vbOKOnly, "ATENCION"
         End If
      'Else
      '   MsgBox "No se a seleccionado un establecimiento", vbOKOnly, "ATENCION"
      '   Me.txt_establecimiento.Enabled = True
      '   Me.txt_nombre_Establecimiento.Enabled = True
      '   Me.txt_establecimiento.SetFocus
      'End If
   End If
End Sub

