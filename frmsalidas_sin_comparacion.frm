VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmsalidas_sin_comparacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   7710
   Visible         =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   1440
      TabIndex        =   30
      Top             =   1920
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   30
         TabIndex        =   31
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
         TabIndex        =   32
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.TextBox txt_foco 
      Height          =   285
      Left            =   8265
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   2850
      Width           =   1125
   End
   Begin VB.Frame Frame3 
      Height          =   1230
      Index           =   0
      Left            =   5145
      TabIndex        =   14
      Top             =   1140
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
         TabIndex        =   15
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
         TabIndex        =   16
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   0
      Left            =   60
      TabIndex        =   13
      Top             =   600
      Width           =   7455
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      Picture         =   "frmsalidas_sin_comparacion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   750
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   390
      Picture         =   "frmsalidas_sin_comparacion.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Buscar Movimiento"
      Top             =   750
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   720
      Picture         =   "frmsalidas_sin_comparacion.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   750
      Width           =   330
   End
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1050
      Picture         =   "frmsalidas_sin_comparacion.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cancelar Movimiento"
      Top             =   750
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7155
      Picture         =   "frmsalidas_sin_comparacion.frx":0408
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   750
      Width           =   330
   End
   Begin VB.Frame frm_busqueda 
      Height          =   960
      Left            =   480
      TabIndex        =   0
      Top             =   3960
      Width           =   3135
      Begin VB.TextBox txt_busqueda_folio 
         Height          =   315
         Left            =   195
         TabIndex        =   6
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
         TabIndex        =   7
         Top             =   120
         Width           =   3060
      End
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   465
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
            Picture         =   "frmsalidas_sin_comparacion.frx":0A42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_sin_comparacion.frx":131C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_sin_comparacion.frx":1BF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_sin_comparacion.frx":2192
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_sin_comparacion.frx":2A6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_sin_comparacion.frx":3348
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_sin_comparacion.frx":3C22
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_sin_comparacion.frx":3D34
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_sin_comparacion.frx":3E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_sin_comparacion.frx":3F58
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_sin_comparacion.frx":406A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_sin_comparacion.frx":417C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   60
      TabIndex        =   17
      Top             =   1005
      Width           =   7455
   End
   Begin VB.Frame Frame3 
      Height          =   1245
      Index           =   1
      Left            =   105
      TabIndex        =   8
      Top             =   1140
      Width           =   4995
      Begin VB.TextBox txt_almacen 
         Height          =   315
         Left            =   825
         TabIndex        =   34
         Top             =   450
         Width           =   780
      End
      Begin VB.TextBox txt_nombre_almacen 
         Height          =   315
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   450
         Width           =   3255
      End
      Begin VB.TextBox txt_referencia 
         Height          =   315
         Left            =   1140
         MaxLength       =   50
         TabIndex        =   9
         Top             =   840
         Width           =   3750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Origen:"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   12
         Top             =   510
         Width           =   510
      End
      Begin VB.Label label 
         BackColor       =   &H8000000D&
         Caption         =   " Datos del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   11
         Top             =   120
         Width           =   4920
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Referencia:"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   10
         Top             =   885
         Width           =   825
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4965
      Left            =   105
      TabIndex        =   18
      Top             =   2325
      Width           =   7425
      Begin VB.CommandButton Command1 
         Height          =   390
         Left            =   6945
         Picture         =   "frmsalidas_sin_comparacion.frx":428E
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   570
         Width           =   375
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
         Left            =   5025
         TabIndex        =   23
         Top             =   555
         Width           =   1890
      End
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   1785
         TabIndex        =   20
         Top             =   1755
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
         TabIndex        =   19
         Top             =   495
         Width           =   2640
      End
      Begin MSComctlLib.ListView lv_entradas 
         Height          =   3780
         Left            =   45
         TabIndex        =   24
         Top             =   1110
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   6668
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
            Text            =   "C?digo"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripci?n"
            Object.Width           =   8617
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label lbl_cantidad 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   4320
         TabIndex        =   27
         Top             =   675
         Width           =   675
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Lectura de Art?culos"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   0
         Left            =   30
         TabIndex        =   26
         Top             =   120
         Width           =   7350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "C?digo del Art?culo:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   675
         Width           =   1395
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   930
      Top             =   -15
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
            Picture         =   "frmsalidas_sin_comparacion.frx":4390
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_sin_comparacion.frx":4C6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_sin_comparacion.frx":5544
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_sin_comparacion.frx":5AE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_sin_comparacion.frx":63BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_sin_comparacion.frx":6C96
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_sin_comparacion.frx":7570
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_sin_comparacion.frx":7682
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_sin_comparacion.frx":7794
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_sin_comparacion.frx":78A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_sin_comparacion.frx":79B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsalidas_sin_comparacion.frx":7ACA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   1950
      Top             =   -15
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
   Begin WMPLibCtl.WindowsMediaPlayer wmp2 
      Height          =   870
      Left            =   0
      TabIndex        =   36
      Top             =   1350
      Visible         =   0   'False
      Width           =   435
      URL             =   "C:\sistemas\desarrollo\integral\sound2.wav"
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
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   645
      Left            =   6555
      TabIndex        =   37
      Top             =   0
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
      Left            =   90
      TabIndex        =   28
      Top             =   105
      Width           =   7335
   End
End
Attribute VB_Name = "frmsalidas_sin_comparacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Dim var_a?o As Integer
Dim var_suma_cantidad As Double
Dim var_cantidad_llegar As Double
Dim var_cantidad As Double
Dim var_renglon As Double
Dim var_tipo_lista As Integer
Dim var_detenido As Integer
Dim cnnMultibondeados As New ADODB.Connection
Private Sub cdm_sonido_Click()
   If Trim(Me.txt_codigo) <> "" Then
      wmp1.Controls.Play
      
      
      
      
   End If
End Sub
Private Sub cdm_sonido_2_Click()
   If Trim(Me.txt_codigo) <> "" Then
      wmp2.Controls.Play
   End If
End Sub


Sub ilumina_grid()
   var_n = lv_entradas.ListItems.Count
   For var_i = 1 To var_n
       If var_i = var_renglon Then
          lv_entradas.ListItems.item(var_i).Bold = True
          lv_entradas.ListItems.item(var_i).ListSubItems(1).Bold = True
          lv_entradas.ListItems.item(var_i).ListSubItems(2).Bold = True
          lv_entradas.ListItems.item(var_i).ForeColor = &H8000&
          lv_entradas.ListItems.item(var_i).ListSubItems(1).ForeColor = &H8000&
          lv_entradas.ListItems.item(var_i).ListSubItems(2).ForeColor = &H8000&
       Else
          lv_entradas.ListItems.item(var_i).Bold = False
          lv_entradas.ListItems.item(var_i).ListSubItems(1).Bold = False
          lv_entradas.ListItems.item(var_i).ListSubItems(2).Bold = False
          lv_entradas.ListItems.item(var_i).ForeColor = &H80000012
          lv_entradas.ListItems.item(var_i).ListSubItems(1).ForeColor = &H80000012
          lv_entradas.ListItems.item(var_i).ListSubItems(2).ForeColor = &H80000012
       End If
   Next var_i
   If var_renglon > 0 Then
      lv_entradas.ListItems.item(var_renglon).Selected = True
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
   
   Dim dl As Long                                 ' Valor devuelto por la funci?n API
   Dim sAttributes As String                  ' Aributos
   Dim sDriver As String                       ' Nombre del controlador
   Dim sDescription As String                ' Descripci?n del DSN
   Dim sDsnName As String                  ' Nombre del DSN

   Const ODBC_ADD_SYS_DSN As Long = 4         ' Se crear? un DSN de sistema
   Const vbAPINull As Long = 0&                         ' Puntero NULL

   ' se elimina
   Const ODBC_REMOVE_SYS_DSN As Long = 6    ' Se eliminar? un DSN de sistema
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
         If var_empresa = "06" Then
            Set reporte = appl.OpenReport(App.Path + "\rep_MOVIMIENTOS_SALIDA_q0z.rpt")
         Else
            If var_clave_movimiento = "SACJ" Then
               Set reporte = appl.OpenReport(App.Path + "\rep_MOVIMIENTOS_SALIDA_COMPUCAJA.rpt")
            Else
               Set reporte = appl.OpenReport(App.Path + "\rep_MOVIMIENTOS_SALIDA.rpt")
            End If
         End If
         reporte.RecordSelectionFormula = "{VW_MOVIMIENTOS_SALIDA.vcha_emp_empresa_id} = '" + var_empresa + "' and {VW_MOVIMIENTOS_SALIDA.VCHA_ALM_ALMACEN_ID} = '" + var_almacen_Destino + "' AND {VW_MOVIMIENTOS_SALIDA.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' AND {VW_MOVIMIENTOS_SALIDA.INTE_EMO_NUMERO} = " + Str(var_numero_folio)
         frmvistasprevias.cr.ReportSource = reporte
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Reporte de Movimientos"
         frmvistasprevias.Show 1
         Set reporte = Nothing
         rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
         
         If var_empresa = "31" Then
            If var_clave_movimiento = "SACJ" Then
               If IsNumeric(Me.txt_folio) Then
                  rs.Open "select * from tb_Salidas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_sal_numero = " + CStr(CDbl(Me.txt_folio)), cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     VAR_FALTANTES = ""
                     While Not rs.EOF
                           rsaux.Open "select isnull(vcha_Art_articulo_id,''), VCHA_aRT_NOMBRE_ESPA?OL from TB_aRTICULOS where vcha_Art_articulo_id = '" + rs!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                           If rsaux(0).Value = "" Then
                              If VAR_FALTANTES = "" Then
                                 VAR_FALTANTES = rs!vcha_Art_Articulo_id + " " + rsaux!VCHA_ART_NOMBRE_ESPA?OL
                                 VAR_FALTANTES = rs!vcha_Art_Articulo_id + " " + rsaux!vcha_Art_nombre_espa?ol
                              Else
                                 VAR_FALTANTES = VAR_FALTANTES + ", " + rs!vcha_Art_Articulo_id + " " + rsaux!vcha_Art_nombre_espa?ol
                              End If
                           End If
                           rsaux.Close
                           
                           rs.MoveNext
                     Wend
                     rs.MoveFirst
                     If VAR_FALTANTES = "" Then
                        Open (App.Path & "\traspaso_" + Trim(Str(rs!INTE_SAL_NUMERO)) + ".txt") For Output As #1
                        While Not rs.EOF
                              rsaux.Open "select * from TB_aRTICULOS where vcha_Art_articulo_id = '" + rs!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                              Print #1, rsaux!vcha_Art_Articulo_id + "," + CStr(rs!floa_Sal_Cantidad) + "," + CStr(rs!floa_Sal_costo)
                              rsaux.Close
                              rs.MoveNext
                        Wend
                        Close #1
                        var_correo_electronico = "rafael.cortes@cantia.com.mx"
                        If Trim(var_correo_electronico) <> "" Then
                           If MAPISession1.SessionID = 0 Then
                              MAPISession1.SignOn
                           End If
                           MAPIMessages1.SessionID = MAPISession1.SessionID
                           MAPIMessages1.Compose
                           MAPIMessages1.RecipDisplayName = var_correo_electronico
                           MAPIMessages1.RecipAddress = var_correo_electronico
                           MAPIMessages1.AddressResolveUI = True
                           MAPIMessages1.ResolveName
                           MAPIMessages1.MsgSubject = "Traspaso " + Me.txt_folio
                           MAPIMessages1.MsgNoteText = "Se adjunta nota del traspaso " + Me.txt_folio
                           MAPIMessages1.AttachmentPathName = App.Path + "\traspaso_" + Me.txt_folio + ".txt"
                           MAPIMessages1.Send True
                           If MAPISession1.SessionID > 0 Then
                              MAPISession1.SignOff
                           End If
                        End If
                     Else
                        MsgBox "Los siguientes c?digos no tienen equivalencia " + VAR_FALTANTES, vbOKOnly, "ATENCION"
                     End If
                  End If
                  rs.Close
               End If
            End If
         End If '''AQUI
         
         
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
                           If rsaux8.State = 1 Then
                              rsaux8.Close
                           End If
                           rsaux8.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + rsaux10!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux8.EOF Then
                              var_nombre_articulo = IIf(IsNull(rsaux8!vcha_Art_nombre_espa?ol), "", rsaux8!vcha_Art_nombre_espa?ol)
                           Else
                              var_nombre_articulo = ""
                           End If
                           rsaux8.Close
                           var_cadena_articulos = rsaux10!vcha_Art_Articulo_id + " " + var_nombre_articulo + " Existen [" + CStr(var_cantidad) + "] y salen [" + CStr(rsaux10!floa_Sal_Cantidad) + "]"
                        Else
                           rsaux8.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + rsaux10!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux8.EOF Then
                              var_nombre_articulo = IIf(IsNull(rsaux8!vcha_Art_nombre_espa?ol), "", rsaux8!vcha_Art_nombre_espa?ol)
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
                           var_nombre_articulo = IIf(IsNull(rsaux8!vcha_Art_nombre_espa?ol), "", rsaux8!vcha_Art_nombre_espa?ol)
                        Else
                           var_nombre_articulo = ""
                        End If
                        rsaux8.Close
                        var_cadena_articulos = rsaux10!vcha_Art_Articulo_id + " " + var_nombre_articulo
                     Else
                        rsaux8.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + rsaux10!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux8.EOF Then
                           var_nombre_articulo = IIf(IsNull(rsaux8!vcha_Art_nombre_espa?ol), "", rsaux8!vcha_Art_nombre_espa?ol)
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
         
            var_si = MsgBox("?Se va a imprimir el movimiento?", vbOKCancel, "ATENCION")
            If var_si = 1 Then
               
               
               cnn.BeginTrans
               
               var_fecha_inicio = CStr(Now)
               rs.Open "EXEC SP_INSERTA_MOVIMIENTOS_SALIDA '" + var_empresa + "','" + var_unidad_organizacional + "', '" + var_almacen_Destino + "','" + var_clave_movimiento + "'," + Str(var_numero_folio) + ",1", cnn, adOpenDynamic, adLockOptimistic
               var_fecha_fin = CStr(Now)
               x = 0
               If x = 1 Then
               Cadena = "select * from tb_temporal_salidas with (nolock) where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_alm_almacen_id = " + var_almacen_Destino + " and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio)
               rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
               While Not rs.EOF
                     var_inserta = False
                     var_suma_cantidad = 0
                     var_cantidad_llegar = IIf(IsNull(rs!floa_Sal_Cantidad), 0, rs!floa_Sal_Cantidad)
                     var_cantidad = 0
                     While var_suma_cantidad < var_cantidad_llegar
                           rsaux2.Open "select * from tb_existencias where vcha_art_articulo_id =  '" + rs!vcha_Art_Articulo_id + "' and vcha_alm_almacen_id = '" + rs!VCHA_ALM_ALMACEN_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux2.EOF Then
                              If rsaux2!floa_exi_cantidad_2004 >= var_cantidad_llegar Then
                                 var_a?o = 2004
                                 var_suma_cantidad = var_cantidad_llegar
                                 var_cantidad = var_cantidad_llegar
                                 var_costo = rsaux2!FLOA_EXI_COSTO_2004
                              Else
                                 var_cantidad_disponible = IIf(IsNull(rsaux2!floa_exi_cantidad_2004), 0, rsaux2!floa_exi_cantidad_2004)
                                 If var_cantidad_disponible > 0 Then
                                    var_a?o = 2004
                                    var_suma_cantidad = IIf(IsNull(rsaux2!floa_exi_cantidad_2004), 0, rsaux2!floa_exi_cantidad_2004)
                                    var_cantidad = IIf(IsNull(rsaux2!floa_exi_cantidad_2004), 0, rsaux2!floa_exi_cantidad_2004)
                                    var_costo = rsaux2!FLOA_EXI_COSTO_2004
                                 Else
                                    var_a?o = 2005
                                    var_cantidad = rs!floa_Sal_Cantidad - var_suma_cantidad
                                    var_suma_cantidad = var_cantidad_llegar
                                    var_costo = rsaux2!floa_exi_costo_2005
                                 End If
                              End If
                           Else
                              var_a?o = 2005
                              var_suma_cantidad = var_cantidad_llegar
                              var_cantidad = var_cantidad_llegar
                              rsaux4.Open "select * from tb_articulos where vcha_art_articulo_id =  '" + rs!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux4.EOF Then
                                 var_costo = IIf(IsNull(rsaux4!mone_Art_costo_estandar), 0, rsaux4!mone_Art_costo_estandar)
                              Else
                                 var_costo = 0
                              End If
                              rsaux4.Close
                           End If
                           rsaux2.Close
                           rsaux.Open "insert into tb_salidas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_sal_numero, vcha_art_articulo_id, floa_sal_cantidad, floa_sal_costo, floa_sal_precio, inte_sal_a?o) values ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!VCHA_UOR_UNIDAD_ID + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_SAL_NUMERO) + ", '" + rs!vcha_Art_Articulo_id + "', " + CStr(var_cantidad) + ", " + CStr(var_costo) + " , " + CStr(rs!floa_Sal_precio) + ", " + CStr(var_a?o) + ")", cnn, adOpenDynamic, adLockOptimistic
                     Wend
                     rs.MoveNext
               Wend
               rs.Close
               var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "", Now, 1)
               var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "I", Now, 1)
               End If
               var_estatus_movimiento = "I"
               
               If var_clave_movimiento = "SPC" And var_empresa = "16" Then
                    
                    
                  rs.Open "select * from tb_Salidas with(nolock) where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_sal_numero = " + Me.txt_folio, cnn, adOpenDynamic, adLockOptimistic
                  var_consecutivo = 0
                  If Not Conectar_BD(cnnMultibondeados, "multibondeados", "admcdindustrial") Then
                    MsgBox "Error al conectar con la base de datos del SIP en Multibondeados", vbCritical, "SID"
                    cnn.RollbackTrans
                    Exit Sub
                  End If
                    Dim rsSubAlmacen As New ADODB.recordSet
                    If IsNumeric(Mid(txt_referencia.Text, 1, IIf(InStr(txt_referencia.Text, " ") > 0, InStr(txt_referencia.Text, " "), 1) - 1)) Then
                        rsSubAlmacen.Open "Select bint_sba_almacen_id " & _
                                            "from tb_sub_almacen with(nolock) " & _
                                            "where bint_sba_almacen_id =" & Mid(txt_referencia.Text, 1, InStr(txt_referencia.Text, " ") - 1), _
                                    cnnMultibondeados, _
                                    adOpenDynamic, _
                                    adLockOptimistic
                        If rsSubAlmacen.RecordCount = 0 Then
                            cnn.RollbackTrans
                            cnnMultibondeados.Close
                            rsSubAlmacen.Close
                            MsgBox "El almacen destino NO existe " & vbCrLf & _
                                    "En el campo referencia precione la tecla ''F5''" & vbCrLf & _
                                    "para seleccionar un almacen destino valido", vbExclamation, "SID"
                            Exit Sub
                        End If
                    Else
                        cnn.RollbackTrans
                        cnnMultibondeados.Close
                        MsgBox "El almacen destino NO existe " & vbCrLf & _
                                "En el campo referencia precione la tecla ''F5''" & vbCrLf & _
                                "para seleccionar un almacen destino valido", vbExclamation, "SID"
                        If rsSubAlmacen.State = 1 Then rsSubAlmacen.Close
                        Exit Sub
                    End If
                    rsSubAlmacen.Close
                  cnnMultibondeados.BeginTrans
                  
                  While Not rs.EOF
                        var_consecutivo = var_consecutivo + 1
                        
                        Call pro_guardarTraspasosSubAlmacenes(rs("vcha_Art_Articulo_id").Value, rs("floa_Sal_Cantidad").Value, 0, cnnMultibondeados)
                        
                        'var_cadena = "insert into mf_tb_movimientos_myg (vcha_art_myg,                    floa_mov_cantidad,                    mon_mov_costo, mon_mov_importe,                                  vcha_mov_afectacion, numb_epe_id, numb_epp_num,                    vcha_mov_entrega_produccion, vcha_mov_movimiento, dtim_aud_fecha, vcha_aud_usuario, bint_pla_planta_id, bint_mov_almacen)"
                        'var_cadena = var_cadena + " values ('" + rs!vcha_Art_Articulo_id + "', " + CStr(rs!floa_Sal_Cantidad) + ", " + CStr(rs!floa_Sal_costo) + "," + CStr(rs!floa_Sal_Cantidad * rs!floa_Sal_costo) + ",'SUMA'," + Me.txt_folio + "," + CStr(var_consecutivo) + "," + Me.txt_folio + ",'SPC', GETDATE(),'" + var_clave_usuario_global + "', 28,1)"
                        'rsaux.Open var_cadena, cnn_sip_multibondeados, adOpenDynamic, adLockOptimistic
                        'rsaux.Open "UPDATE MF_TB_ART_MYG SET FLOA_ART_EXISTENCIA = FLOA_ART_EXISTENCIA + " + CStr(rs!floa_Sal_Cantidad) + " WHERE VCHA_aRT_ID = '" + rs!vcha_Art_Articulo_id + "'", cnn_sip_multibondeados, adOpenDynamic, adLockOptimistic
                        rs.MoveNext
                  Wend
                  rs.Close
               End If
               
               If var_empresa = "06" Then
                  Set reporte = appl.OpenReport(App.Path + "\rep_MOVIMIENTOS_SALIDA_q0z.rpt")
               Else
                  If var_clave_movimiento = "SACJ" Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_MOVIMIENTOS_SALIDA_COMPUCAJA.rpt")
                  Else
                     Set reporte = appl.OpenReport(App.Path + "\rep_MOVIMIENTOS_SALIDA.rpt")
                  End If
               End If
               reporte.RecordSelectionFormula = "{VW_MOVIMIENTOS_SALIDA.vcha_emp_empresa_id} = '" + var_empresa + "' and {VW_MOVIMIENTOS_SALIDA.VCHA_ALM_ALMACEN_ID} = '" + var_almacen_Destino + "' AND {VW_MOVIMIENTOS_SALIDA.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' AND {VW_MOVIMIENTOS_SALIDA.INTE_EMO_NUMERO} = " + Str(var_numero_folio)
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Reporte de Movimientos"
               frmvistasprevias.Show 1
               Set reporte = Nothing
               rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
            
               If var_empresa = "31" Then
                  If var_clave_movimiento = "SACJ" Then
                     If IsNumeric(Me.txt_folio) Then
                        rs.Open "select * from tb_Salidas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_sal_numero = " + CStr(CDbl(Me.txt_folio)), cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           VAR_FALTANTES = ""
                           While Not rs.EOF
                                 rsaux.Open "select isnull(vcha_Art_articulo_id,''), VCHA_aRT_NOMBRE_ESPA?OL from TB_aRTICULOS where vcha_Art_articulo_id = '" + rs!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                 If rsaux(0).Value = "" Then
                                    If VAR_FALTANTES = "" Then
                                       VAR_FALTANTES = rs!vcha_Art_Articulo_id + " " + rsaux!vcha_Art_nombre_espa?ol
                                    Else
                                       VAR_FALTANTES = VAR_FALTANTES + ", " + rs!vcha_Art_Articulo_id + " " + rsaux!vcha_Art_nombre_espa?ol
                                    End If
                                 End If
                                 rsaux.Close
                                 
                                 rs.MoveNext
                           Wend
                           rs.MoveFirst
                           If VAR_FALTANTES = "" Then
                              Open (App.Path & "\traspaso_" + Trim(Str(rs!INTE_SAL_NUMERO)) + ".txt") For Output As #1
                              While Not rs.EOF
                                    rsaux.Open "select * from TB_aRTICULOS where vcha_Art_articulo_id = '" + rs!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                    Print #1, rsaux!vcha_Art_Articulo_id + "," + CStr(rs!floa_Sal_Cantidad) + "," + CStr(rs!floa_Sal_costo)
                                    rsaux.Close
                                    rs.MoveNext
                              Wend
                              Close #1
                              var_correo_electronico = "rafael.cortes@cantia.com.mx"
                              If Trim(var_correo_electronico) <> "" Then
                                 If MAPISession1.SessionID = 0 Then
                                    MAPISession1.SignOn
                                 End If
                                 MAPIMessages1.SessionID = MAPISession1.SessionID
                                 MAPIMessages1.Compose
                                 MAPIMessages1.RecipDisplayName = var_correo_electronico
                                 MAPIMessages1.RecipAddress = var_correo_electronico
                                 MAPIMessages1.AddressResolveUI = True
                                 MAPIMessages1.ResolveName
                                 MAPIMessages1.MsgSubject = "Traspaso " + Me.txt_folio
                                 MAPIMessages1.MsgNoteText = "Se adjunta nota del traspaso " + Me.txt_folio
                                 MAPIMessages1.AttachmentPathName = App.Path + "\traspaso_" + Me.txt_folio + ".txt"
                                 MAPIMessages1.Send True
                                 If MAPISession1.SessionID > 0 Then
                                    MAPISession1.SignOff
                                 End If
                              End If
                           Else
                              MsgBox "Los siguientes c?digos no tienen equivalencia " + VAR_FALTANTES, vbOKOnly, "ATENCION"
                           End If
                        End If
                        rs.Close
                     End If
                  End If
               End If ''
               If cnnMultibondeados.State = 1 Then
                  cnnMultibondeados.Close
               End If
               cnn.CommitTrans
               If cnnMultibondeados.State = 1 Then cnnMultibondeados.CommitTrans: cnnMultibondeados.Close
               
               txt_codigo.Enabled = False
               txt_foco.Enabled = False
            End If
         Else
            MsgBox "El movimiento no se puede imprimir ya que las existencias de los siguientes art?culos exceden a la cantidad disponible en el almac?n " + var_cadena_articulos
         End If
      End If
   Else
      MsgBox "No se a seleccionado ning?n movimiento", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   If var_numero_folio > 0 Then
      rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
   End If
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
   
   
   
   
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   'On Error GoTo salir:
   strConnectionString = "DRIVER=Microsoft Excel Driver (*.xls); DBQ=c:\inventario.xls"
   rsaux2.Open "SELECT * FROM [SALIDAS$]", strConnectionString
   'rsaux2.Open "SELECT dbo.TB_EXISTENCIAS.VCHA_ART_ARTICULO_ID AS CODIGO, dbo.TB_EXISTENCIAS.FLOA_EXI_CANTIDAD  AS CANTIDAD, dbo.TB_ARTICULOS.MONE_ART_PRECIO_BASE AS PRECIO, dbo.TB_ARTICULOS.MONE_ART_COSTO_ESTANDAR AS COSTO FROM         dbo.TB_EXISTENCIAS INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_EXISTENCIAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID WHERE     (dbo.TB_EXISTENCIAS.VCHA_ALM_ALMACEN_ID = 'PTVH') AND FLOA_eXI_CANTIDAD > 0", cnn, adOpenDynamic, adLockOptimistic
   'rsaux2.Open "select vcha_art_Articulo_id as codigo, floa_Exi_Cantidad cantidad from tb_Existencias where vcha_alm_almacen_id = 'AB' and floa_Exi_cantidad > 0", cnn, adOpenDynamic, adLockOptimistic
   rs.Open "delete from TB_TEMP_ENTRADAS_SALIDAS_AJUSTES", cnn, adOpenDynamic, adLockOptimistic
   While Not rsaux2.EOF
         If Not IsNull(rsaux2!codigo) Then
            If rsaux2!Cantidad > 0 Then
               rsaux4.Open "select * from tb_articulos where vcha_Art_articulo_id = '" + CStr(rsaux2!codigo) + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux4.EOF Then
                  var_codigo = rsaux4!vcha_Art_Articulo_id
                  var_DEscripcion = IIf(IsNull(rsaux4!vcha_Art_nombre_espa?ol), "", rsaux4!vcha_Art_nombre_espa?ol)
                  var_costo = IIf(IsNull(rsaux4!mone_Art_costo_estandar), 0, rsaux4!mone_Art_costo_estandar)
                  var_precio = IIf(IsNull(rsaux4!mone_Art_precio_base), 0, rsaux4!mone_Art_precio_base)
                  var_cantidad = rsaux2!Cantidad
               Else
                  rsaux5.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + CStr(IIf(IsNull(rsaux2!codigo), "", rsaux2!codigo)) + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux5.EOF Then
                     rsaux6.Open "select * from tb_articulos where vcha_Art_articulo_id = '" + CStr(rsaux5!vcha_Art_Articulo_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux6.EOF Then
                        var_codigo = rsaux6!vcha_Art_Articulo_id
                        var_DEscripcion = IIf(IsNull(rsaux6!vcha_Art_nombre_espa?ol), "", rsaux6!vcha_Art_nombre_espa?ol)
                        var_costo = IIf(IsNull(rsaux6!mone_Art_costo_estandar), 0, rsaux6!mone_Art_costo_estandar)
                        var_precio = IIf(IsNull(rsaux6!mone_Art_precio_base), 0, rsaux6!mone_Art_precio_base)
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
               rsaux.Open "INSERT INTO TB_TEMP_ENTRADAS_SALIDAS_AJUSTES (vcha_Art_articulo_id, vcha_art_descripcion, floa_tem_cantidad, floa_tem_costo, floa_tem_precio) VALUES ('" + Mid(CStr(var_codigo), 1, 50) + "','" + Mid(var_DEscripcion, 1, 50) + "'," + CStr(var_cantidad) + "," + CStr(var_costo) + "," + CStr(var_precio) + ")", cnn, adOpenDynamic, adLockOptimistic
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
      MsgBox "No existen los siguientes art?culos " + var_cadena, vbOKOnly, "ATENCION"
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
                     Set TB_TEMPORAL_SALIDAS_INSERTA = New TB_TEMPORAL_SALIDAS_INSERTA
                     Set TB_TEMPORAL_SALIDAS_MODIFICA = New TB_TEMPORAL_SALIDAS_MODIFICA
                     Set TB_ENCABEZADO_MOVIMIENTOS_I = New TB_ENCABEZADO_MOVIMIENTOS_I
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
                        Cadena = "select * from TB_TEMPORAL_SALIDAS with (nolock) where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "'"
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
                     End If
                     rsaux8.MoveNext
               Wend
               MsgBox "Se a terminado de cargar " + CStr(var_cantidad) + " piezas", vbOKOnly, "ATENCION"
            Else
               MsgBox "Falta agregar una referencia", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Falta indicar el almac?n", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Falta indicar el almac?n", vbOKOnly, "ATENCION"
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
      MsgBox "DEBE DE CREAR EL ARCHIVO DE EXCEL INVENTARIO Y ESTE DEBE DE CONTAR CON LA HOJA LLAMADA SALIDAS", vbOKOnly, "ATENCION"
   Else
      If Err.Number = 3265 Then
         MsgBox "LOS NOMBRES DE LAS COLUMNAS DEBEN DE SER CODIGO Y CANTIDAD", vbOKOnly, "ATENCION"
      Else
         MsgBox "A surgido un error al cargar el archivo de salidas", vbOKOnly, "ATENCION"
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
   lbl_Cantidad.Visible = False
   txt_Cantidad.Visible = False
   txt_referencia.Enabled = False
   txt_codigo.Enabled = False
   var_primera_vez = True
   var_cantidad_leida = 1#
   If var_clave_movimiento = "SA" Then
      Me.Command1.Visible = True
   Else
      Me.Command1.Visible = False
   End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
   If var_numero_folio > 0 Then
      rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
   End If
   Call activa_forma(var_activa_forma_salidas_sin_comparacion)
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
      Else
         If var_tipo_lista = 2 Then
            Me.txt_referencia = lv_lista.selectedItem + " " + lv_lista.selectedItem.SubItems(1)
            Me.txt_referencia.SetFocus
         End If
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
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la informaci?n disponible"
End Sub

Private Sub txt_almacen_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
        If rs.State = 1 Then rs.Close
      If var_tipo_permiso = 1 Then
         rs.Open "select * from vw_almacen_permiso_1 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      Else
         rs.Open "select * from vw_movimientos_almacenes where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_emp_Empresa_id = '" + var_empresa + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
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
         txt_referencia.Enabled = True
      Else
         MsgBox "Clave de almacen Incorrecta", vbOKOnly, "ATENCION"
         txt_almacen = ""
         txt_nombre_almacen = ""
         txt_referencia.Enabled = False
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
            rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 where inte_emo_numero = " + txt_busqueda_folio + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
         End If
         rs.Open "select * from tb_encabezado_movimientos where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emo_numero = " + txt_busqueda_folio + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            If var_numero_folio > 0 Then
               rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
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
                  txt_referencia = IIf(IsNull(rs!vcha_Emo_referencia), "", rs!vcha_Emo_referencia)
                  txt_referencia.Enabled = False
                  lv_entradas.ListItems.Clear
                  var_primera_vez = False
                  var_numero_folio = rs!INTE_EMO_NUMERO
                  txt_folio = var_numero_folio
                  rsaux.Open "select * from tb_ALMACENES where vcha_alm_almacen_id = '" + var_almacen_Destino + "'", cnn, adOpenDynamic, adLockOptimistic
                  txt_almacen_destino = rsaux(3).Value
                  txt_nombre_almacen.Text = rsaux(3).Value
                  rsaux.Close
                  rsaux.Open "select * from tb_temporal_SALIDAS with (nolock)  where vcha_emp_empresa_id = '" + var_empresa + "' and inte_SAL_numero = " + txt_busqueda_folio + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
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
                     txt_Cantidad.Visible = False
                     lbl_Cantidad.Visible = False
                     txt_foco.Enabled = False
                  Else
                     txt_foco.Enabled = False
                     txt_codigo.Enabled = True
                     txt_Cantidad.Visible = False
                     lbl_Cantidad.Visible = False
                  End If
               Else
                  MsgBox "No esta autorizado para modificar este movimiento", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "El movimiento esta siendo utilizado por otro usuario", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El n?mero de movimiento no existe ", vbOKOnly, "ATENCION"
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
      If IsNumeric(txt_cantidad_eliminar) Then
         Dim var_posible_eliminar As Boolean
         var_cantidad_eliminar = Val(txt_cantidad_eliminar)
         var_posible_eliminar = True
         If var_posible_eliminar = True Then
            var_inserta = False
            rsaux.Open "UPDATE TB_TEMPORAL_SALIDAS SET FLOA_SAL_CANTIDAD = ISNULL(FLOA_SAL_CANTIDAD,0) - " + txt_cantidad_eliminar + " WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_SAL_NUMERO = " + CStr(var_numero_folio) + " AND VCHA_ART_ARTICULO_ID= '" + lv_entradas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
            'var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_destino, var_clave_movimiento, var_numero_folio, lv_entradas.SelectedItem, 0 - Val(txt_cantidad_eliminar))
            lv_entradas.selectedItem.SubItems(2) = lv_entradas.selectedItem.SubItems(2) - Val(txt_cantidad_eliminar)
            var_renglon = lv_entradas.selectedItem.Index
            Call ilumina_grid
         Else
            MsgBox "La cantidad a eliminar supera a la cantidad asignada a la causa de devoluci?n seleccionada", vbOKOnly, "ATENCION"
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
   txt_Cantidad = ""
End Sub

Private Sub txt_cantidad_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If Trim(txt_Cantidad) <> "" Then
         If IsNumeric(Me.txt_Cantidad) Then
            If CDbl(Me.txt_Cantidad) < 10000 Then
               var_cantidad_leida = txt_Cantidad
               txt_foco.Enabled = True
               txt_foco.SetFocus
               lbl_Cantidad.Visible = False
               txt_Cantidad.Visible = False
            Else
               Call cdm_sonido_2_Click
               'Me.txt_codigo = ""
               frmmensaje.lbl_mensaje = "La cantidad no debe de ser mayor de 10000"
               frmmensaje.Show 1
            End If
         Else
            Call cdm_sonido_2_Click
            'Me.txt_codigo = ""
            frmmensaje.lbl_mensaje = "El art?culo no existe"
            frmmensaje.Show 1
         End If
      End If
   End If
End Sub

Private Sub txt_codigo_GotFocus()
   If Len(var_codigo_seleccionado) = 0 Then
      txt_codigo = ""
   End If
   var_codigo_seleccionado = ""
End Sub

Private Sub txt_Codigo_KeyDown(KeyCode As Integer, Shift As Integer)
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
   var_codigo_seleccionado = ""
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      var_detenido = 0
      var_verificador = True
      If Len(Trim(txt_codigo)) = 12 Then
         Call calcula_verificador(Trim(txt_codigo))
      End If
      If var_empresa = "31" Then
         var_verificador = True
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
               var_detenido = IIf(IsNull(rs!INTE_ART_detenido), 0, rs!INTE_ART_detenido)
               If var_clave_movimiento = "SA" Or (var_clave_movimiento = "SACJ" And Mid(Me.txt_codigo, 1, 1) = "A") Then
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
               var_precio = IIf(IsNull(rs(2).Value), 0, rs(2).Value)
               rs.Close
               If var_recontable = 1 Then
                  var_cantidad_leida = 1#
                  lbl_Cantidad.Visible = True
                  txt_Cantidad.Visible = True
                  txt_Cantidad.SetFocus
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
                     var_detenido = IIf(IsNull(rs!INTE_ART_detenido), 0, rs!INTE_ART_detenido)
                     If var_cantidad_caja = 0 Then
                        If var_clave_movimiento = "SA" Or (var_clave_movimiento = "SACJ" And Mid(Me.txt_codigo, 1, 1) = "A") Then
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
                     var_precio = IIf(IsNull(rs(2).Value), 0, rs(2).Value)
                     rs.Close
                     If var_recontable = 1 Then
                        var_cantidad_leida = 1#
                        lbl_Cantidad.Visible = True
                        txt_Cantidad.Visible = True
                        txt_Cantidad.SetFocus
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
                     Me.txt_codigo = ""
                     frmmensaje.lbl_mensaje = "El art?culo no existe"
                     frmmensaje.Show 1
                     'MsgBox "El art?culo no existe", vbOKOnly, "ATENCION"
                  End If
               Else
                  Call cdm_sonido_Click
                  Me.txt_codigo = ""
                  frmmensaje.lbl_mensaje = "El art?culo no existe"
                  frmmensaje.Show 1
                  'MsgBox "El art?culo no existe", vbOKOnly, "ATENCION"
                  rs.Close
               End If
            End If
         Else
         End If
      Else
         txt_codigo = ""
         frmmensaje.lbl_mensaje = "Error en C?digo"
         frmmensaje.Show 1
         'MsgBox "Error en C?digo", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_foco_GotFocus()
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_SALIDAS_INSERTA = New TB_TEMPORAL_SALIDAS_INSERTA
   Set TB_TEMPORAL_SALIDAS_MODIFICA = New TB_TEMPORAL_SALIDAS_MODIFICA
   Set TB_ENCABEZADO_MOVIMIENTOS_I = New TB_ENCABEZADO_MOVIMIENTOS_I
   Dim var_inserta As Boolean
   Dim var_pase_existencias As Double
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
            var_pase_existencias = True
         End If
      End If
      If var_pase_existencias = 1 Then
         var_pase = 0
         If var_clave_movimiento = "SACJ" Then
            If var_detenido = 0 Then
               var_pase = 0
            Else
               var_pase = 1
            End If
         Else
            var_pase = 0
         End If
         If var_pase = 0 Then
            bandera_suma = False
            If var_primera_vez = True Then
               var_inserta = False
               var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, Now, var_numero_folio, 0, "", "", "", var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, 0, "", txt_referencia, "", "B", "", "", 0, 0, 0, var_clave_moneda, 0)
               var_numero_folio = var_numero_folio_regreso
               txt_folio = var_numero_folio
               var_primera_vez = False
            End If
            var_pase_existencias = 1
            If var_empresa = "18" And var_almacen_Destino <> "RETEX" Then
               rsaux.Open "select isnull(floa_sal_cantidad,0) from tb_Temporal_salidas where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_sal_numero = " + CStr(var_numero_folio) + " and vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  var_cantidad_temporal = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
               Else
                  var_cantidad_temporal = 0
               End If
               rsaux.Close
               rsaux.Open "select floa_exi_Cantidad from tb_existencias where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  var_cantidad_Existencias = IIf(IsNull(rsaux!floa_Exi_Cantidad), 0, rsaux!floa_Exi_Cantidad)
               Else
                  var_cantidad_Existencias = 0
               End If
               rsaux.Close
               var_cantidad_posible = var_cantidad_Existencias - (var_cantidad_temporal + var_cantidad_leida)
               If var_cantidad_posible < 0 Then
                  var_pase_existencias = 0
               End If
            End If
         
            If var_pase_existencias = 1 Then
               Cadena = "select * from TB_TEMPORAL_SALIDAS with (nolock) where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_sal_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "'"
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
            Else
               Me.txt_codigo.SetFocus
               frmmensaje.lbl_mensaje = "La cantidad excede a la cantidad en existencias"
               frmmensaje.Show 1
            End If
         Else
            MsgBox "El art?culo esta bloqueado para su venta", vbOKOnly, "ATENCION"
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
         If txt_referencia.Enabled = True Then
            txt_referencia.SetFocus
         End If
      End If
   End If
End Sub

Private Sub txt_referencia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = "116" Then
        Select Case var_clave_movimiento
            Case "SACJ"
                pro_llena_lista_para_SACJ
            Case "SPC"
                pro_llena_lista_para_SPC
        End Select
        If var_clave_movimiento = "SACJ" Then
           
        End If
    End If
End Sub
Private Sub pro_llena_lista_para_SPC()
    Dim rsDestino As New ADODB.recordSet
    Dim strQry As String
    
    Dim fila As Integer
    
    If Conectar_BD(cnnMultibondeados, "multibondeados", "admcdindustrial") Then
        rsDestino.Open "Select bint_sba_almacen_id, vcha_sba_nombre " & _
                        "from tb_sub_almacen " & _
                        "where inte_sba_consignacion =0  " & _
                        "and  vcha_sba_activo ='1' " & _
                        "and vcha_sba_lineaId is not null ", _
                cnnMultibondeados, _
                adOpenDynamic, _
                adLockOptimistic
        var_tipo_lista = 2
        Me.lv_lista.ListItems.Clear
        
        If rsDestino.RecordCount > 0 Then
            
            For fila = 1 To rsDestino.RecordCount
                Set list_item = lv_lista.ListItems.Add(, , rsDestino("bint_sba_almacen_id").Value)
                list_item.SubItems(1) = rsDestino("vcha_sba_nombre").Value
                rsDestino.MoveNext
            Next
        Else
            MsgBox "No se encantraron destinos para este movimiento", vbCritical, "SID"
        End If
        Me.frm_lista.Visible = True
        Me.lv_lista.SetFocus
    Else
        MsgBox "Error al concectar con el servidor de multibondeados", vbCritical, "SID"
    End If
    'cnnMultibondeados.RollbackTrans
    If cnnMultibondeados.State = 1 Then cnnMultibondeados.Close
    
End Sub
Private Function Conectar_BD(ByRef cnnCBD As ADODB.Connection, ByVal bd As String, ByVal servidor As String) As Boolean
    'Variables de bloque
    Dim strConnection_String As String
    
On Error GoTo Error_Conectar_BDS
    Conectar_BD = True
    'Establecer connection strings para realizar las conexiones a las bases de
    'datos
    
    If servidor = "dbpruebas" Then
        MsgBox "Esta en modo de Pruebas", vbExclamation, "SID"
    End If
    
    strConnection_String_SID = "Provider=SQLOLEDB.1;Password=elia" & _
                                ";Persist Security Info=True;User ID=sa" & _
                                ";Initial Catalog=" & UCase(bd) & ";Data Source=" & UCase(servidor)
    
    'Configurar objetos Connection
    'cnnCBD.CursorLocation = adUseClient
    If cnnCBD.State = 1 Then
        cnnCBD.Close
    End If
    cnnCBD.ConnectionString = strConnection_String_SID
    cnnCBD.CommandTimeout = 60
    cnnCBD.CursorLocation = adUseClient
    
    'Abrir conexiones a las bases de datos
    cnnCBD.Open
    Exit Function
Error_Conectar_BDS:
    Conectar_BD = False
    MsgBox Err.Description, vbCritical, "SID"
End Function

Private Sub pro_llena_lista_para_SACJ()
    var_tipo_lista = 2
    Me.lv_lista.ListItems.Clear
    Set list_item = lv_lista.ListItems.Add(, , "3")
    list_item.SubItems(1) = "SERVICIO AL CLIENTE"
    Set list_item = lv_lista.ListItems.Add(, , "1")
    list_item.SubItems(1) = "TIENDA"
    Set list_item = lv_lista.ListItems.Add(, , "5")
    list_item.SubItems(1) = "EXHIBICION"
    Me.frm_lista.Visible = True
    Me.lv_lista.SetFocus
End Sub

Private Sub txt_referencia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If var_clave_movimiento = "SACJ" Then
      If KeyAscii <> 13 Then
         KeyAscii = 0
      End If
   End If
   If KeyAscii = 13 Then
      If Len(Trim(txt_referencia)) > 0 Then
         txt_codigo.Enabled = True
         txt_codigo.SetFocus
         txt_referencia.Enabled = False
      Else
         MsgBox "Debe introducir una referencia", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub pro_guardarTraspasosSubAlmacenes(strCodigo As String, numbCantidad As Double, i As Integer, cnnMulti As ADODB.Connection)
    Dim cmd As New ADODB.Command
    Dim rsSalida As New ADODB.recordSet
    
    rsSalida.Open "Select case mone_art_costo_estandar " & _
                            "when  0 then mone_art_precio_base " & _
                            "else mone_art_costo_estandar end  costoSTD " & _
                    "from tb_articulos with(nolock) " & _
                    "where vcha_art_articulo_id ='" & strCodigo & "' ", _
            cnn, _
            adOpenDynamic, _
            adLockOptimistic
                    
    cmd.CommandText = "[mf_sp_reg_movimientos_SIP_myg]"
    cmd.CommandType = adCmdStoredProc
    cmd.ActiveConnection = cnnMultibondeados
    cmd("@p_vcha_art_id").Value = strCodigo
    cmd("@p_floa_cantiad").Value = numbCantidad
    cmd("@P_floa_cto_std").Value = IIf(IsNull(rsSalida("costoSTD").Value), "0", rsSalida("costoSTD").Value)
    cmd("@p_bint_folioReferencia").Value = txt_folio.Text
    cmd("@p_vcha_movimientoId").Value = "ENTTSB"
    cmd("@p_vcha_afectacion").Value = "SUMA"
    cmd("@p_vcha_ope").Value = var_usuario_global
    cmd("@p_sba_subAlmacen").Value = Mid(txt_referencia.Text, 1, InStr(txt_referencia.Text, " ") - 1)
    cmd.execute
    Set cmd = Nothing

End Sub


