VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmentradas_devolucion_completa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_pasar_todo 
      Height          =   1500
      Left            =   1920
      TabIndex        =   0
      Top             =   1875
      Width           =   3675
      Begin VB.TextBox txt_serie 
         Height          =   315
         Left            =   735
         TabIndex        =   5
         Top             =   990
         Width           =   1000
      End
      Begin VB.TextBox txt_numero 
         Height          =   315
         Left            =   2445
         TabIndex        =   4
         Top             =   990
         Width           =   1000
      End
      Begin VB.Frame Frame5 
         Height          =   30
         Left            =   0
         TabIndex        =   3
         Top             =   765
         Width           =   3660
      End
      Begin VB.CommandButton cmd_aceptar_pedidos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   60
         Picture         =   "frmentradas_devolucion_completa.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Aceptar Alt + A"
         Top             =   420
         Width           =   330
      End
      Begin VB.CommandButton cmd_cancelar_pedidos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   390
         Picture         =   "frmentradas_devolucion_completa.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Cancelar Alt + C"
         Top             =   420
         Width           =   330
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         Caption         =   " Factura a Devolver"
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   30
         TabIndex        =   8
         Top             =   120
         Width           =   3600
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   210
         TabIndex        =   7
         Top             =   1050
         Width           =   405
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   1785
         TabIndex        =   6
         Top             =   1050
         Width           =   600
      End
   End
   Begin VB.Frame frm_busqueda 
      Height          =   960
      Left            =   1140
      TabIndex        =   12
      Top             =   585
      Width           =   3135
      Begin VB.TextBox txt_busqueda_folio 
         Height          =   315
         Left            =   195
         MaxLength       =   10
         TabIndex        =   13
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
         TabIndex        =   14
         Top             =   120
         Width           =   3060
      End
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   1545
      TabIndex        =   9
      Top             =   150
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
         TabIndex        =   11
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.TextBox txt_foco 
      Height          =   285
      Left            =   8985
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   3735
      Width           =   1125
   End
   Begin VB.Frame Frame4 
      Height          =   1035
      Left            =   6255
      TabIndex        =   53
      Top             =   2205
      Width           =   2040
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   60
         TabIndex        =   55
         Top             =   150
         Width           =   1830
      End
      Begin VB.Label lbl_total 
         Alignment       =   2  'Center
         Caption         =   "12345619999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   60
         TabIndex        =   54
         Top             =   525
         Width           =   1830
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4005
      Left            =   90
      TabIndex        =   41
      Top             =   3240
      Width           =   8235
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
         TabIndex        =   46
         Top             =   405
         Width           =   2640
      End
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   1785
         TabIndex        =   43
         Top             =   1755
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar 
            Height          =   330
            Left            =   60
            MaxLength       =   10
            TabIndex        =   44
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
            TabIndex        =   45
            Top             =   15
            Width           =   2895
         End
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
         Left            =   6240
         TabIndex        =   42
         Top             =   465
         Width           =   1890
      End
      Begin MSComctlLib.ListView lv_entradas 
         Height          =   2835
         Left            =   45
         TabIndex        =   47
         Top             =   1065
         Width           =   8115
         _ExtentX        =   14314
         _ExtentY        =   5001
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
            Object.Width           =   2478
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   9349
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   2328
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código del Artículo:"
         Height          =   195
         Left            =   120
         TabIndex        =   51
         Top             =   585
         Width           =   1395
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Lectura de Artículos"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   0
         Left            =   30
         TabIndex        =   50
         Top             =   120
         Width           =   8160
      End
      Begin VB.Label lbl_cantidad 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   5535
         TabIndex        =   49
         Top             =   585
         Width           =   675
      End
      Begin VB.Label lbl_cancelado 
         Alignment       =   2  'Center
         Caption         =   "MOVIMIENTO CANCELADO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   4290
         TabIndex        =   48
         Top             =   390
         Width           =   3765
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1080
      Index           =   0
      Left            =   6240
      TabIndex        =   20
      Top             =   1125
      Width           =   2055
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
         TabIndex        =   21
         Top             =   435
         Width           =   1950
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Folio del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   1
         Left            =   30
         TabIndex        =   22
         Top             =   120
         Width           =   1980
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7920
      Picture         =   "frmentradas_devolucion_completa.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Salir"
      Top             =   735
      Width           =   330
   End
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1065
      Picture         =   "frmentradas_devolucion_completa.frx":08CE
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Cancelar Movimiento"
      Top             =   735
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   735
      Picture         =   "frmentradas_devolucion_completa.frx":09D0
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   735
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frmentradas_devolucion_completa.frx":0AD2
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Buscar Movimiento Alt + B"
      Top             =   735
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmentradas_devolucion_completa.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Nuevo Movimiento Alt + N"
      Top             =   735
      Width           =   330
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   30
      Top             =   15
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
            Picture         =   "frmentradas_devolucion_completa.frx":0CD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_devolucion_completa.frx":15B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_devolucion_completa.frx":1E8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_devolucion_completa.frx":2426
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_devolucion_completa.frx":2D02
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_devolucion_completa.frx":35DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_devolucion_completa.frx":3EB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_devolucion_completa.frx":3FC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_devolucion_completa.frx":40DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_devolucion_completa.frx":41EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_devolucion_completa.frx":42FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_devolucion_completa.frx":4410
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   30
      TabIndex        =   52
      Top             =   990
      Width           =   8250
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   0
      Left            =   30
      TabIndex        =   40
      Top             =   585
      Width           =   8250
   End
   Begin VB.Frame Frame3 
      Height          =   2115
      Index           =   1
      Left            =   75
      TabIndex        =   23
      Top             =   1125
      Width           =   6150
      Begin VB.TextBox txt_cliente 
         Height          =   315
         Left            =   1290
         TabIndex        =   33
         Top             =   1080
         Width           =   1005
      End
      Begin VB.TextBox txt_almacen 
         Height          =   315
         Left            =   1290
         TabIndex        =   32
         Top             =   420
         Width           =   1005
      End
      Begin VB.TextBox txt_establecimiento 
         Height          =   315
         Left            =   1290
         TabIndex        =   31
         Top             =   1410
         Width           =   1005
      End
      Begin VB.TextBox txt_referencia 
         Height          =   315
         Left            =   1290
         TabIndex        =   30
         Top             =   1740
         Width           =   4380
      End
      Begin VB.TextBox txt_agente 
         Height          =   315
         Left            =   1290
         TabIndex        =   29
         Top             =   750
         Width           =   1005
      End
      Begin VB.TextBox txt_nombre_almacen 
         Height          =   315
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   420
         Width           =   3750
      End
      Begin VB.TextBox txt_nombre_agente 
         Height          =   315
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   750
         Width           =   3750
      End
      Begin VB.TextBox txt_nombre_establecimiento 
         Height          =   315
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1410
         Width           =   3750
      End
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   315
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1080
         Width           =   3750
      End
      Begin VB.CommandButton cmd_pasar_todo 
         Height          =   330
         Left            =   5700
         Picture         =   "frmentradas_devolucion_completa.frx":4522
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Pasar una factura"
         Top             =   1740
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   39
         Top             =   1155
         Width           =   525
      End
      Begin VB.Label label 
         BackColor       =   &H8000000D&
         Caption         =   " Datos del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   30
         TabIndex        =   38
         Top             =   120
         Width           =   6075
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   37
         Top             =   495
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Establecimiento:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   36
         Top             =   1485
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Referencia:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   35
         Top             =   1815
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   34
         Top             =   810
         Width           =   555
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
      Left            =   30
      TabIndex        =   56
      Top             =   120
      Width           =   8325
   End
End
Attribute VB_Name = "frmentradas_devolucion_completa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn_cantia As ADODB.Connection
Dim var_conexion_cantia As String
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
Dim var_clave_cliente As String
Dim var_clave_titular As String
Dim var_solo_lectura As Boolean
Dim var_clave_almacen_costo As String
Dim var_ventana As Integer
Dim var_clave_moneda As String
Dim var_tipo_lista As Integer
Dim var_renglon As Double
Dim var_prefijo As String

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
   If lv_entradas.ListItems.Count > 11 Then
      lv_entradas.ColumnHeaders(2).Width = 5050.22
   Else
      lv_entradas.ColumnHeaders(2).Width = 5300.22
   End If
   
   lv_entradas.Refresh
   
End Sub
















Private Sub cmd_aceptar_pedidos_Click()
   Dim var_agente_todo As String
   Dim var_cliente_todo As String
   If rsaux5.State = 1 Then
      rsaux5.Close
   End If
   rsaux5.Open "select * from tb_encabezado_Cartera where vcha_ser_serie_id = '" + txt_serie + "' and inte_car_numero = " + txt_numero + " and vcha_Car_documento = 'FA'", cnn, adOpenDynamic, adLockOptimistic
   If Not rsaux5.EOF Then
      var_agente_todo = IIf(IsNull(rsaux5!VCHA_AGE_AGENTE_ID), "", rsaux5!VCHA_AGE_AGENTE_ID)
      var_cliente_todo = IIf(IsNull(rsaux5!vcha_cli_clave_id), "", rsaux5!vcha_cli_clave_id)
      If var_agente_todo = Me.txt_agente Then
         If var_cliente_todo = Me.txt_cliente Then
            var_si = MsgBox("¿Desea devolver los articulos de la factura?" + txt_factura, vbYesNo, "ATENCION")
            If var_si = 6 Then
               rsaux.Open "SELECT * FROM TB_sALIDAS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_SER_SERIE_ID = '" + txt_serie + "' AND INTE_CAR_NUMERO = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
               While Not rsaux.EOF
                     txt_codigo = rsaux!VCHA_ART_ARTICULO_ID
                     var_cantidad_leida = rsaux!floa_Sal_Cantidad
                     
                     
                     Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
                     Set TB_TEMPORAL_ENTRADAS_INSERTA = New TB_TEMPORAL_ENTRADAS_INSERTA
                     Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
                     Set TB_BLOQUEOS = New TB_BLOQUEOS
                     Dim var_inserta As Boolean
                     Dim var_factura As Integer
                     If Trim(txt_codigo.Text) <> "" Then
                        bandera_suma = False
                        If var_primera_vez = True Then
                           var_inserta = False
                           var_insreta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, Now, var_numero_folio, 0, txt_cliente, "", "", var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, "", "", txt_referencia, txt_establecimiento, "B", var_clave_titular, txt_agente, 0, 0, 0, var_clave_moneda, 1)
                           var_numero_folio = var_numero_folio_regreso
                           var_global_bloqueado = 1
                           var_inserta = False
                           var_inserta = TB_BLOQUEOS.Anadir(var_empresa, var_unidad_organizacional, "DEVOLUCION" + Trim(var_clave_movimiento) + Trim(Str(var_numero_folio)), Now, var_clave_usuario_global, fun_NombrePc)
                           var_solo_lectura = False
                           txt_folio = var_numero_folio
                           var_primera_vez = False
                        End If
      
                        rs.Open "select * from tb_existencias where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           var_costo = IIf(IsNull(rs!floa_exi_costo_2005), 0, rs!floa_exi_costo_2005)
                           If var_costo = 0 Then
                              var_costo = IIf(IsNull(rs!FLOA_EXI_COSTO_2004), 0, rs!FLOA_EXI_COSTO_2004)
                           End If
                        End If
                        rs.Close
      
                        If var_costo = 0 Then
                           rs.Open "select * from tb_existencias where vcha_alm_almacen_id = '8' and vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rs.EOF Then
                              var_costo = IIf(IsNull(rs!floa_exi_costo_2005), 0, rs!floa_exi_costo_2005)
                              If var_costo = 0 Then
                                 var_costo = IIf(IsNull(rs!FLOA_EXI_COSTO_2004), 0, rs!FLOA_EXI_COSTO_2004)
                              End If
                           End If
                           rs.Close
                        End If
      
      
      
                        If var_costo = 0 Then
                           rs.Open "SELECT MONE_ART_COSTO_ESTANDAR FROM TB_ARTICULOS WHERE VCHA_ART_ARTICULO_ID = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rs.EOF Then
                              var_costo = IIf(IsNull(rs!mone_Art_costo_estandar), 0, rs!mone_Art_costo_estandar)
                           End If
                           rs.Close
                        End If
                        rs.Open "SELECT * FROM TB_aRTICULOS WHERE VCHA_ART_ARTICULO_ID = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           var_descripcion_articulo = IIf(IsNull(rs!vcha_Art_nombre_español), "", rs!vcha_Art_nombre_español)
                        End If
                        rs.Close
                        Cadena = "select * from TB_TEMPORAL_ENTRADAS where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "'"
                        rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           lbl_total = CStr(CDbl(lbl_total) + var_cantidad_leida)
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
                           lbl_total = CStr(CDbl(lbl_total) + var_cantidad_leida)
                           var_inserta = TB_TEMPORAL_ENTRADAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida, var_costo, var_precio, "0", "", var_año)
                           rs.Close
                           Set list_item = lv_entradas.ListItems.Add(, , Trim(txt_codigo))
                           list_item.SubItems(1) = var_descripcion_articulo
                           list_item.SubItems(2) = var_cantidad_leida
                           var_renglon = lv_entradas.ListItems.Count
                           Call ilumina_grid
                        End If
                        txt_codigo = ""
                     End If
                     
                     
                     
                     
                     rsaux.MoveNext
               Wend
               rsaux.Close
               txt_codigo = ""
            End If
         Else
            MsgBox "El cliente seleccionado en el movimiento no pertenece al de la factura", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El agente seleccionado en el movimiento no pertenece al de la factura", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "La factura no existe", vbOKOnly, "ATENCION"
   End If
   rsaux5.Close
   Me.frm_pasar_todo.Visible = False
End Sub

Private Sub cmd_buscar_Click()
   var_ventana = 1
   frm_busqueda.Visible = True
   txt_busqueda_folio.SetFocus
End Sub

Private Sub cmd_cancelar_Click()
   Set TB_EXISTENCIAS_INSERTA = New TB_EXISTENCIAS_INSERTA
   Set TB_ENTRADAS_I = New TB_ENTRADAS_I
   Set TB_SALIDAS_INSERTA = New TB_SALIDAS_INSERTA
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
   Set TB_ENTRADAS_VISTAS_I = New TB_ENTRADAS_VISTAS_I
   If var_numero_folio > 0 Then
      If var_estatus_movimiento = "C" Then
         MsgBox "El Movimiento ya fue cancelado", vbOKOnly, "ATENCION"
      Else
         If var_estatus_movimiento = "I" Then
            If var_fecha_movimiento <> Date Then
               var_posible_accion = False
               frmsupervisor1.Show 1
               If var_posible_accion = True Then
                  si = MsgBox("¿Desea cancelar el movimiento?", vbYesNo, "ATENCION")
                  If si = 6 Then
                     si = MsgBox("Confirmar la cancelación del movimiento", vbYesNo, "ATENCION")
                     If si = 6 Then
                        Set TB_ENC_MOV_CANCELACION = New TB_ENC_MOV_CANCELACION
                        var_actualizar = False
                        var_actualizar = TB_ENC_MOV_CANCELACION.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "C", var_global_supervisor_1, var_global_supervisor_2)
                        lbl_cancelado = "MOVIMIENTO CANCELADO"
                        Me.cmd_imprimir.Enabled = False
                        Me.cmd_cancelar.Enabled = False
                        MsgBox "El movimiento a sido cancelado", vbOKOnly, "ATENCION"
                        var_estatus_movimiento = "C"
                        Me.lbl_cancelado = "MOVIMIENTO CANCELADO"
                     End If
                  End If
               End If
            End If
         Else
            MsgBox "El Movimiento no a sido cerrado aun", vbOKOnly, "ATENCION"
         End If
      End If
   Else
      MsgBox "No se a seleccionado un movimiento", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_cancelar_pedidos_Click()
   Me.frm_pasar_todo.Visible = True
End Sub

Private Sub cmd_imprimir_Click()
   Set TB_EXISTENCIAS_INSERTA = Nothing
   Set TB_EXISTENCIAS_INSERTA = New TB_EXISTENCIAS_INSERTA
   Set TB_ENTRADAS_I = Nothing
   Set TB_ENTRADAS_I = New TB_ENTRADAS_I
   Set TB_ENCABEZADO_MOVIMIENTOS_M = Nothing
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
   Set TB_FOLIOS_MOVIMIENTOS = Nothing
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_DEVOLUCIONES_NUM_DESTINO = Nothing
   Set TB_DEVOLUCIONES_NUM_DESTINO = New TB_DEVOLUCIONES_NUM_DESTINO
   Dim var_primera_folio_detalle As Boolean
   Dim var_primera_vez_Devolucion As Boolean
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
   If var_numero_folio > 0 Then
      If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
         Set reporte = appl.OpenReport(App.Path + "\rep_MOVIMIENTOS_ENTRADAS_devoluciones.rpt")
         reporte.RecordSelectionFormula = "{VW_MOVIMIENTOS_ENTRADAs_devoluciones.vcha_uor_unidad_id} = '" + var_unidad_organizacional + "' and {VW_MOVIMIENTOS_ENTRADAs_devoluciones.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' AND {VW_MOVIMIENTOS_ENTRADAs_Devoluciones.INTE_EMO_NUMERO} = " + Str(var_numero_folio) + " and {VW_MOVIMIENTOS_ENTRADAs_Devoluciones.vcha_emp_empresa_id} = '" + var_empresa + "'"
         frmvistasprevias.cr.ReportSource = reporte
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Reporte de Movimientos"
         frmvistasprevias.Show
         Set reporte = Nothing
         rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_Destino + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
      Else
         var_primera_vez_Devolucion = True
         Set TB_DEVOLUCIONES_INSERTA = New TB_DEVOLUCIONES_INSERTA
         var_si = MsgBox("¿Se va a imprimir el movimiento?", vbOKCancel, "ATENCION")
         If var_si = 1 Then
            'cnn.BeginTrans
            
            
            cnn.CommandTimeout = 360
            rsaux2.Open "select * from vw_clientes where vcha_cli_clave_id ='" + Me.txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux2.EOF Then
               var_tipo_busqueda = IIf(IsNull(rsaux2!INTE_CAN_BUSQUEDA_FACTURA_GRUPO), 1, rsaux2!INTE_CAN_BUSQUEDA_FACTURA_GRUPO)
               var_grupo = IIf(IsNull(rsaux2!VCHA_GAC_GRUPO_aCTUAL_ID), "", rsaux2!VCHA_GAC_GRUPO_aCTUAL_ID)
               var_lista_precios = IIf(IsNull(rsaux2!vcha_LIS_LISTA_iD), "", rsaux2!vcha_LIS_LISTA_iD)
               var_canal_venta = IIf(IsNull(rsaux2!vcha_can_canal_venta_id), "", rsaux2!vcha_can_canal_venta_id)
               var_iva_canal = IIf(IsNull(rsaux2!FLOA_TPE_IVA), 0, rsaux2!FLOA_TPE_IVA)
               var_clave_titular = IIf(IsNull(rsaux2!vcha_tit_titular_id), "", rsaux2!vcha_tit_titular_id)
               var_clave_agente = IIf(IsNull(rsaux2!VCHA_AGE_AGENTE_ID), "", rsaux2!VCHA_AGE_AGENTE_ID)
            Else
               var_grupo = ""
               var_tipo_busqueda = 1
               var_lista_precios = ""
               var_canal_venta = ""
               var_iva_canal = 0
               var_clave_titular = ""
               var_clave_agente = ""
               
            End If
            rsaux2.Close
            
            
            Cadena = "select * from tb_temporal_entradas where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'"
            rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
            var_consecutivo = 0
            var_primera_vez = True
            While Not rs.EOF
                  var_inserta = False
                  var_consecutivo = var_consecutivo + 1
                                       
                  rsaux2.Open "select * from VW_ORDEN_FECHAS_FACTURAS where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_art_articulo_id = '" + rs!VCHA_ART_ARTICULO_ID + "' and vcha_cli_clave_id ='" + Me.txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux2.EOF Then
                     var_clave_moneda = IIf(IsNull(rsaux2!vcha_mon_moneda_id), "", rsaux2!vcha_mon_moneda_id)
                     var_moneda_local = IIf(IsNull(rsaux2!inte_mon_moneda_local), 0, rsaux2!inte_mon_moneda_local)
                     var_tipo_Cambio = 1
                     var_precio = IIf(IsNull(rsaux2!floa_Sal_precio), 0, rsaux2!floa_Sal_precio)
                     var_descuento_1 = IIf(IsNull(rsaux2!FLOA_SAL_DESCUENTO_1), 0, rsaux2!FLOA_SAL_DESCUENTO_1)
                     var_descuento_2 = IIf(IsNull(rsaux2!FLOA_SAL_DESCUENTO_2), 0, rsaux2!FLOA_SAL_DESCUENTO_2)
                     var_iva = IIf(IsNull(rsaux2!floa_car_porcentaje_iva), 0, rsaux2!floa_car_porcentaje_iva)
                     var_factura = IIf(IsNull(rsaux2!inte_Car_numero), 0, rsaux2!inte_Car_numero)
                     var_serie = IIf(IsNull(rsaux2!vcha_Ser_Serie_id), 0, rsaux2!vcha_Ser_Serie_id)
                     rsaux4.Open "SELECT max(FLOA_cAR_DESCUENTO_APLICADO) FROM TB_rELACION_COBRANZA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_CAR_NUMERO =  " + CStr(var_factura) + " AND VCHA_CAR_DOCUMENTO= 'FA'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux4.EOF Then
                        var_descuento_3 = IIf(IsNull(rsaux4(0).Value), 0, rsaux4(0).Value)
                     Else
                        var_descuento_3 = 0
                     End If
                     rsaux4.Close
                     var_serie_FACTURA = IIf(IsNull(rsaux2!vcha_Ser_Serie_id), "", rsaux2!vcha_Ser_Serie_id)
                  Else
                     rsaux3.Open "select * from tb_detalle_lista_precios where vcha_Art_Articulo_id = '" + rs!VCHA_ART_ARTICULO_ID + "' and vcha_lis_lista_precios_id = '" + var_lista_precios + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux3.EOF Then
                        var_precio = rsaux3!floa_dli_Precio
                     Else
                        var_precio = rs!floa_ent_precio
                     End If
                     rsaux3.Close
                     var_serie = ""
                     var_factura = 0
                     rsaux3.Open "select * from tb_gruposactuales where vcha_gac_grupo_Actual_id = '" + var_grupo + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux3.EOF Then
                        var_descuento_1 = IIf(IsNull(rsaux3!floa_gac_Descuento_1), 0, rsaux3!floa_gac_Descuento_1)
                        var_descuento_2 = IIf(IsNull(rsaux3!FLOA_GAC_DESCUENTO_2), 0, rsaux3!FLOA_GAC_DESCUENTO_2)
                     Else
                        var_descuento_1 = 0
                        var_descuento_2 = 0
                     End If
                     rsaux3.Close
                    
                  End If
                  rsaux2.Close
                  
               
               
                  rsaux10.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_aRT_aRTICULO_ID = '" + rs!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux10.EOF Then
                     var_nombre_articulo = IIf(IsNull(rsaux10!VCHA_ART_ARTICULO_ID), "", rsaux10!VCHA_ART_ARTICULO_ID)
                  Else
                     var_nombre_articulo = ""
                  End If
                  rsaux10.Close
                  Text1 = "INSERT INTO TB_ENTRADAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID,  INTE_ENT_NUMERO, VCHA_ART_ARTICULO_ID,  FLOA_ENT_CANTIDAD, FLOA_ENT_COSTO,  FLOA_ENT_PRECIO,  FLOA_ENT_DESCUENTO, VCHA_ENT_ALMACEN_ORIGEN, INTE_ENT_AÑO, inte_ent_consecutivo)   VALUES  ( '" + var_empresa + "',  '" + var_unidad_organizacional + "',  '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", '" + rs!VCHA_ART_ARTICULO_ID + "', " + CStr(rs!floa_ent_cantidaD) + ",  " + CStr(rs!floa_ent_costo) + ",  " + CStr(var_precio) + ", 0, '" + var_almacen_Destino + "', 2005, " + CStr(var_consecutivo) + ")"
                  rsaux.Open "INSERT INTO TB_ENTRADAS (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID,  INTE_ENT_NUMERO, VCHA_ART_ARTICULO_ID,  FLOA_ENT_CANTIDAD, FLOA_ENT_COSTO,  FLOA_ENT_PRECIO,  FLOA_ENT_DESCUENTO, VCHA_ENT_ALMACEN_ORIGEN, INTE_ENT_AÑO, inte_ent_consecutivo)   VALUES  ( '" + var_empresa + "',  '" + var_unidad_organizacional + "',  '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", '" + rs!VCHA_ART_ARTICULO_ID + "', " + CStr(rs!floa_ent_cantidaD) + ",  " + CStr(rs!floa_ent_costo) + ",  " + CStr(var_precio) + ", 0, '" + var_almacen_Destino + "', 2005, " + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                  var_cadena = " INSERT INTO TB_DEVOLUCIONES (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO, VCHA_ART_ARTICULO_ID, CHAR_CDE_ESTATUS, INTE_CDE_CONSECUTIVO, VCHA_CDE_DESTINO, FLOA_CDE_COSTO, FLOA_CDE_PRECIO, FLOA_CDE_DESCUENTO_1, FLOA_CDE_DESCUENTO_2, FLOA_CDE_DESCUENTO_3, FLOA_CDE_IVA, VCHA_SER_SERIE_ID, INTE_FAC_FACTURA, INTE_CDE_NUMERO_DESTINO, VCHA_CDE_MOVIMIENTO_DESTINO, VCHA_CDE_REFERENCIA, INTE_CDE_ASIGNADO, VCHA_MON_MONEDA_ID, FLOA_DEV_TIPO_CAMBIO, VCHA_FAG_FAMILIA_AGRUPADOR_ID, VCHA_AGR_AGRUPADOR_ID, VCHA_DEV_DESCRIPCION_AGRUPADOR, INTE_DEV_AÑO, INTE_DEV_LOTE, VCHA_DEV_SUPERVISOR, DTIM_DEV_FECHA_LOTE, VCHA_DEV_NOTA_ENVIO, FLOA_DEV_CANTIDAD, VCHA_DEV_TIPO_DEFECTO, VCHA_DEV_PROVEEDOR, VCHA_DEV_JUSTIFICA_DEVOLUCION, VCHA_DEV_ESTATUS, DTIM_INT_FECHA, INTE_INT_INTERFACE,INTE_DEV_RECHAZADO)"
                  var_cadena = var_cadena + " VALUES  ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!VCHA_UOR_UNIDAD_ID + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "'," + CStr(rs!inte_ent_numero) + ", '" + rs!VCHA_ART_ARTICULO_ID + "', 'I', " + CStr(var_consecutivo) + ", '" + rs!VCHA_ALM_ALMACEN_ID + "', " + CStr(rs!floa_ent_costo) + ", " + CStr(var_precio) + ", " + CStr(var_descuento_1) + ", " + CStr(var_descuento_2) + ", 0, 16, '" + var_serie + "', " + CStr(var_factura) + ", 0,'" + rs!VCHA_MOV_MOVIMIENTO_ID + "', '" + Me.txt_referencia + "', 1, '" + var_clave_moneda + "', 1, '', '" + rs!VCHA_ART_ARTICULO_ID + "', '" + var_nombre_articulo + "', 2005, "
                  var_cadena = var_cadena + " NULL, NULL, NULL, NULL, " + CStr(rs!floa_ent_cantidaD) + ", NULL, NULL, NULL, NULL, NULL, NULL, NULL)"
                  'Text1 = var_cadena
                  rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               
                  If var_clave_movimiento = "CA" Then
                     var_clave_movimiento_DESTINO = "DC"
                  Else
                     var_clave_movimiento_DESTINO = "TC"
                  End If
               
                  x_z = 1
                  If x_z = 1 Then
                  If var_primera_vez_Devolucion = True Then
                     rsaux10.Open "select * from tb_folios_movimientos where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento_DESTINO + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux10.EOF Then
                        rsaux11.Open "select max(inte_emo_numero) from tb_folios_movimientos where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento_DESTINO + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux11.EOF Then
                           var_numero_folio_regreso = IIf(IsNull(rsaux11(0).Value), 0, rsaux11(0).Value) + 1
                           rsaux9.Open "update tb_folios_movimientos set inte_emo_numero = isnull(inte_emo_numero,0) + 1 where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento_DESTINO + "'", cnn, adOpenDynamic, adLockOptimistic
                        End If
                        rsaux11.Close
                     Else
                        rsaux11.Open "insert into tb_folios_movimientos (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_mov_movimiento_id, inte_emo_numero) values ('" + var_empresa + "', '" + var_unidad_organizacional + "','" + var_clave_movimiento_DESTINO + "',1)", cnn, adOpenDynamic, adLockOptimistic
                        var_numero_folio_regreso = 1
                     End If
                     rsaux10.Close
                     var_cadena = "insert into tb_encabezado_movimientos (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_emo_numero, dtim_emo_fecha,vcha_Cli_clave_id, vcha_emo_almacen_destino, vcha_emo_almacen_origen, vcha_aud_usuario, vcha_aud_maquina, vcha_esb_establecimiento_id, vcha_tit_titular_id, vcha_Age_agente_id,floa_emo_Descuento_1, floa_emo_descuento_2, floa_emo_Descuento_3, vcha_mon_moneda_id, floa_emo_tipo_cambio)"
                     var_cadena = var_cadena + " values ('" + var_empresa + "', '" + var_unidad_organizacional + "','" + var_almacen_Destino + "','" + var_clave_movimiento_DESTINO + "'," + CStr(var_numero_folio_regreso) + ",getdate(),'" + Me.txt_cliente + "','" + var_almacen_Destino + "','" + var_almacen_Destino + "','" + var_clave_usuario_global + "','" + fun_NombrePc + "','" + Me.txt_establecimiento + "','" + var_clave_titular + "','" + Me.txt_agente + "',0,0,0,'1',1)"
                     rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     'var_insreta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, CStr(rs!VCHA_ALM_ALMACEN_ID), CStr(var_clave_movimiento_DESTINO), Now, 0, 0, Me.txt_cliente, "", rs!VCHA_ALM_ALMACEN_ID, rs!VCHA_ALM_ALMACEN_ID, "", var_clave_usuario_global, fun_NombrePc, "", "", "", CStr(Me.txt_establecimiento), "", CStr(var_clave_titular), Me.txt_agente, 0, 0, 0, CStr(var_clave_moneda), 1)
                     var_primera_vez_Devolucion = False
                     'var_numero_folio = var_numero_folio_regreso
                  End If
                  
                  rsaux4.Open "Update tb_encabezado_movimientos set DTIM_EMO_FECHA_FINALIZO = getdate(), char_emo_estatus = '' where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND vcha_alm_almacen_id = '" + rs!VCHA_ALM_ALMACEN_ID + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento_DESTINO + "' AND INTE_EMO_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                  rsaux11.Open "insert into tb_entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_art_articulo_id, floa_ent_cantidad, floa_ent_costo, floa_ent_precio, floa_ent_descuento, VCHA_ENT_ALMACEN_ORIGEN,inte_ent_año) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + var_clave_movimiento_DESTINO + "', " + CStr(var_numero_folio_regreso) + ", '" + rs!VCHA_ART_ARTICULO_ID + "', " + CStr(rs!floa_ent_cantidaD) + ", " + CStr(var_costo) + ", " + CStr(var_precio) + ", 0, '" + rs!VCHA_ALM_ALMACEN_ID + "', 2005)", cnn, adOpenDynamic, adLockOptimistic
                  rsaux11.Open "insert into tb_salidas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_sal_numero, vcha_art_articulo_id, floa_sal_cantidad, floa_sal_costo, floa_sal_precio, floa_sal_descuento, inte_sal_año) values('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + var_clave_movimiento_DESTINO + "', " + CStr(var_numero_folio_regreso) + ", '" + rs!VCHA_ART_ARTICULO_ID + "', " + CStr(rs!floa_ent_cantidaD) + ", " + CStr(var_costo) + ", " + CStr(var_precio) + ", 0,2005)", cnn, adOpenDynamic, adLockOptimistic
                  rsaux4.Open "Update tb_encabezado_movimientos set DTIM_EMO_FECHA_FINALIZO = getdate(), char_emo_estatus = 'I' where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND vcha_alm_almacen_id = '" + rs!VCHA_ALM_ALMACEN_ID + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento_DESTINO + "' AND INTE_EMO_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                  End If
               
               rs.MoveNext
            Wend
            rsaux10.Open "update tb_encabezado_movimientos set char_emo_estatus = 'I' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento_DESTINO + "' and inte_emo_numero = " + CStr(var_numero_folio_regreso), cnn, adOpenDynamic, adLockOptimistic
            var_modifica = TB_DEVOLUCIONES_NUM_DESTINO.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, var_almacen_Destino, CDbl(var_numero_folio), var_clave_movimiento)
            rs.Close
            
            var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "I", Now, 1)
            ' inicio de la afectacion del almacen de materia prima de la planta
            If var_unidad_organizacional = "27" Or var_unidad_organizacional = "28" Or var_unidad_organizacional = "29" Then
               var_numero_planta = CDbl(var_unidad_organizacional)
               If var_unidad_organizacional = 27 Then
                  VAR_PLANTA_CORRECTA = "03"
               End If
               If var_unidad_organizacional = 28 Then
                  VAR_PLANTA_CORRECTA = "04"
               End If
               If var_unidad_organizacional = 29 Then
                  VAR_PLANTA_CORRECTA = "1"
               End If
               If rsaux.State = 1 Then
                  rsaux.Close
               End If
               rsaux.Open "SELECT VCHA_ALM_ALMACEN_ID FROM TB_ALMACENES WHERE BINT_PLA_PLANTA_ID = " + CStr(VAR_PLANTA_CORRECTA), cnn_cantia, adOpenDynamic, adLockOptimistic
               VAR_ALMACEN_MP = rsaux!VCHA_ALM_ALMACEN_ID
               rsaux.Close
               cnn_cantia.BeginTrans
               rsaux.Open "SELECT * FROM TB_TFOLIOS", cnn_cantia, adOpenDynamic, adLockOptimistic
               var_folio_tfolios = rsaux(0).Value + 1
               rsaux.Close
               rsaux.Open "UPDATE TB_TFOLIOS SET BINT_FTR_TFOLIOS = BINT_FTR_TFOLIOS + 1", cnn_cantia, adOpenDynamic, adLockOptimistic
               cnn_cantia.CommitTrans
               var_cadena = "INSERT INTO TB_TRANSACCIONES (BINT_TRA_TRANSACCIONES_ID, VCHA_MOV_MOVIMIENTO_ID, VCHA_TRA_MOVIMIENTO1, VCHA_TRA_MOVIMIENTO2, VCHA_TRA_ALAMACEN, VCHA_TRA_MOVIMIENTO3, VCHA_TRA_STATUS, DTIM_AUD_FECHA, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA, BINT_PLA_PLANTA_ID, BINT_TRA_REIMPRESION )"
               var_cadena = var_cadena + " VALUES (" + CStr(var_folio_tfolios) + ",'DVDMP','DVDMP','DVDMP','" + VAR_ALMACEN_MP + "','DVDMP','A',GETDATE(),'" + var_clave_usuario_global + "','" + fun_NombrePc + "'," + CStr(VAR_PLANTA_CORRECTA) + ",1)"
               var_referencia_mp = CStr(var_numero_planta) + "_" + CStr(var_folio_tfolios)
               rsaux.Open var_cadena, cnn_cantia, adOpenDynamic, adLockOptimistic
               var_primera_folio = False
               'Cadena = "select * from tb_temporal_salidas with (nolock) where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_SAL_numero = " + Me.txt_folio + " AND VCHA_aRT_ARTICULO_ID = '" + rs!vcha_Art_articulo_id + "'"
               Cadena = "select * from TB_tEMPORAL_entradas where VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Me.txt_folio
               If rs.State = 1 Then
                  rs.Close
               End If
               rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
               var_primera_folio_detalle = True
               While Not rs.EOF
                     If Mid(rs!VCHA_ART_ARTICULO_ID, 1, 6) = "646244" Then
                        var_codigo_sip = Mid(rs!VCHA_ART_ARTICULO_ID, 7, 5)
                     Else
                        var_codigo_sip = rs!VCHA_ART_ARTICULO_ID
                     End If
                     If var_primera_folio_detalle = True Then
                        cnn_cantia.BeginTrans
                        rsaux10.Open "select * from tb_folios where vcha_mov_movimiento_id = 'DVDMP'", cnn_cantia, adOpenDynamic, adLockOptimistic
                        If Not rsaux10.EOF Then
                           VAR_FOLIO_DETALLE = rsaux10!BINT_FOL_FOLIO
                        Else
                           rsaux2.Open "INSERT INTO TB_FOLIOS (VCHA_FOL_FOLIO_ID, BINT_FOL_FOLIO, VCHA_MOV_MOVIMIENTO_ID, VCHA_FOL_STATUS, DTIM_AUD_FECHA, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA, BINT_PLA_PLANTA_ID) VALUES ('DVDMP', 1, 'DVDMP','A', GETDATE(),'SID_" + var_clave_usuario_global + "','" + fun_NombrePc + "'," + CStr(VAR_PLANTA_CORRECTA) + ")", cnn_cantia, adOpenDynamic, adLockOptimistic
                           VAR_FOLIO_DETALLE = 1
                        End If
                        rsaux10.Close
                        rsaux10.Open "UPDATE TB_FOLIOS SET BINT_FOL_FOLIO = BINT_FOL_FOLIO + 1 WHERE VCHA_MOV_MOVIMIENTO_ID = 'DVDMP'", cnn_cantia, adOpenDynamic, adLockOptimistic
                        cnn_cantia.CommitTrans
                        var_primera_folio_detalle = False
                     End If
                     
                     
                     
                     var_inserta = False
                     cnn_cantia.BeginTrans
                     rsaux.Open "SELECT * FROM TB_DFOLIOS", cnn_cantia, adOpenDynamic, adLockOptimistic
                     VAR_FOLIO_DFOLIOS = rsaux!BINT_TDE_TFOLIOS + 1
                     rsaux.Close
                     rsaux.Open "UPDATE TB_DFOLIOS SET BINT_TDE_TFOLIOS =  BINT_TDE_TFOLIOS + 1", cnn_cantia, adOpenDynamic, adLockOptimistic
                     cnn_cantia.CommitTrans
                     rsaux.Open "SELECT * FROM TB_aRTICULOS WHERE VCHA_aRT_ARTICULO_ID = '" + var_prefijo + var_codigo_sip + "'", cnn_cantia, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        var_costo = IIf(IsNull(rsaux!FLOA_ART_COSPROMEDIO), 0, rsaux!FLOA_ART_COSPROMEDIO)
                     End If
                     rsaux.Close
                     var_cadena = "INSERT INTO TB_DETALLE (BINT_DET_DETALLE_ID, VCHA_ART_ARTICULO_ID, FLOA_DET_CANTIDAD, MON_DET_PRECIO, MON_DET_IMPORTE, BINT_TRA_TRANSACCIONES_ID, VCHA_DET_AFECTACION, BINT_DET_FOLIO, VCHA_DET_MOVIMIENTO, VCHA_DET_STATUS, DTIM_AUD_FECHA, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA, BINT_PLA_PLANTA_ID, FLOA_DET_CANTIDADSURTIDA, FLOA_ART_EXISTENCIAANT, FLOA_ART_IMPORTEULT, VCHA_DET_ALMACEN, VCHA_DET_AÑOINVENTARIO)"
                     var_cadena = var_cadena + " Values  ( " + CStr(VAR_FOLIO_DFOLIOS) + ",'" + var_prefijo + var_codigo_sip + "', " + CStr(rs!floa_ent_cantidaD) + ", " + CStr(rs!floa_ent_costo) + ", " + CStr(rs!floa_ent_cantidaD * rs!floa_ent_costo) + ", " + CStr(var_folio_tfolios) + ", 'SUMA', " + CStr(VAR_FOLIO_DETALLE) + ", 'DVDMP', '', GETDATE(), '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', " + CStr(VAR_PLANTA_CORRECTA) + ", 0, 0, 0, '1', 2005)"
                     rsaux.Open var_cadena, cnn_cantia, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               rs.Close
            End If 'fin de afectacion al almacen de materia prima de la planta
            
            
            'cnn.CommitTrans
            var_estatus_movimiento = "I"
            Set reporte = appl.OpenReport(App.Path + "\rep_MOVIMIENTOS_ENTRADAS_devoluciones.rpt")
            reporte.RecordSelectionFormula = "{VW_MOVIMIENTOS_ENTRADAs_devoluciones.vcha_uor_unidad_id} = '" + var_unidad_organizacional + "' and {VW_MOVIMIENTOS_ENTRADAs_devoluciones.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' AND {VW_MOVIMIENTOS_ENTRADAs_Devoluciones.INTE_EMO_NUMERO} = " + Str(var_numero_folio) + " and {VW_MOVIMIENTOS_ENTRADAs_Devoluciones.vcha_emp_empresa_id} = '" + var_empresa + "'"
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
         End If
      End If
   Else
      MsgBox "No se a seleccionado ningún movimiento", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   lbl_total = "0"
   lbl_cancelado = ""
   If var_numero_folio > 0 Then
     rs.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_EMO_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
   End If
   txt_codigo.Enabled = False
   var_primera_vez = True
   var_ventana = 0
   frm_busqueda.Visible = False
   lv_entradas.ListItems.Clear
   var_numero_folio = 0
   txt_folio = ""
   txt_codigo = ""
   var_estatus_movimiento = ""
   txt_cliente = ""
   txt_establecimiento = ""
   txt_agente = ""
   txt_almacen = ""
   txt_referencia = ""
   txt_almacen.Enabled = True
   txt_cliente.Enabled = False
   txt_agente.Enabled = False
   txt_establecimiento.Enabled = False
   txt_referencia.Enabled = False
   txt_codigo.Enabled = False
   txt_cliente.Enabled = False
   txt_almacen.SetFocus
   txt_nombre_almacen = ""
   txt_nombre_agente = ""
   txt_nombre_establecimiento = ""
   txt_nombre_cliente = ""
End Sub

Private Sub cmd_pasar_todo_Click()
   Me.frm_pasar_todo.Visible = True
   Me.txt_serie.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = 116 Then
      frmexisten_rapidas.Show
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
      If var_ventana = 0 Then
         Unload Me
      Else
         Me.frm_busqueda.Visible = False
         Me.frm_eliminar.Visible = False
         Me.frm_lista.Visible = False
         var_ventana = 0
      End If
   End If
End Sub

Private Sub Form_Load()
   Set cnn_cantia = CreateObject("ADODB.connection")
   rs.Open "select VCHA_PRI_RUTA_NOTAS_ENVIO from tb_principal where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_ruta = IIf(IsNull(rs!VCHA_PRI_RUTA_NOTAS_ENVIO), "", rs!VCHA_PRI_RUTA_NOTAS_ENVIO)
   End If
   rs.Close
   var_conexion_cantia = ""
   rs.Open "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_conexion_cantia = IIf(IsNull(rs!vcha_uor_conexion), "", rs!vcha_uor_conexion)
   Else
      var_conexion_cantia = ""
   End If
   rs.Close
   If var_conexion_cantia <> "" Then
      cnn_cantia.Open var_conexion_cantia
      'MsgBox var_conexion_cantia
   End If
   
   Me.frm_pasar_todo.Visible = False
   If var_clave_usuario_global = "11" Or var_clave_usuario_global = "8" Or var_clave_usuario_global = "U0000000068" Then
      Me.cmd_pasar_todo.Visible = True
   Else
      Me.cmd_pasar_todo.Visible = False
   End If
   lbl_total = "0"
   lbl_cancelado = ""
   var_año = 2005
   var_numero_folio = 0
   var_cadena_seguridad = ""
   Top = 0
   Left = 1500
   frm_lista.Visible = False
   var_estatus_movimiento = ""
   var_ventana = 0
   frm_busqueda.Visible = False
   frm_eliminar.Visible = False
   lbl_Cantidad.Visible = False
   txt_Cantidad.Visible = False
   txt_cliente.Enabled = False
   txt_codigo.Enabled = False
   txt_agente.Enabled = False
   txt_establecimiento.Enabled = False
   var_primera_vez = True
   var_cantidad_leida = 1#
   rs.Open "select * from tb_almacenes where inte_alm_costeo = 1 AND VCHA_eMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
   var_clave_almacen_costo = IIf(IsNull(rs!VCHA_ALM_ALMACEN_ID), "", rs!VCHA_ALM_ALMACEN_ID)
   rs.Close
   rs.Open "select * from tb_monedas where inte_mon_moneda_local = 1", cnn, adOpenDynamic, adLockOptimistic
   var_clave_moneda = ""
   If Not rs.EOF Then
      var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
   End If
   rs.Close
   var_ventana = 0
   If var_unidad_organizacional = "27" Then
      var_prefijo = "3_"
   End If
   If var_unidad_organizacional = "28" Then
      var_prefijo = "4_"
   End If
   If var_unidad_organizacional = "29" Then
      var_prefijo = "1_"
   End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
   If var_solo_lectura = False Then
   End If
   Call activa_forma(var_activa_forma_entradas_devoluciones)
   If var_numero_folio > 0 Then
     rsaux.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_EMO_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
   End If
End Sub

Private Sub lv_entradas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
         MsgBox "Imporsible modificar el movimiento", vbOKOnly, "ATENCION"
      Else
         var_elimina = False
         var_ventana = 1
         frm_eliminar.Visible = True
         txt_cantidad_eliminar.SetFocus
      End If
   End If
End Sub

Private Sub Toolbar1_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_GotFocus()
   var_ventana = 1
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 0 Then
         If var_tipo_lista = 1 Then
            txt_almacen = lv_lista.selectedItem
            txt_nombre_almacen = lv_lista.selectedItem.SubItems(1)
            txt_almacen.Enabled = True
            txt_almacen.SetFocus
         End If
         If var_tipo_lista = 2 Then
            txt_agente = lv_lista.selectedItem
            txt_nombre_agente = lv_lista.selectedItem.SubItems(1)
            txt_agente.Enabled = True
            txt_agente.SetFocus
         End If
         If var_tipo_lista = 3 Then
            txt_establecimiento = lv_lista.selectedItem
            txt_nombre_establecimiento = lv_lista.selectedItem.SubItems(1)
            txt_establecimiento.Enabled = True
            txt_establecimiento.SetFocus
         End If
         If var_tipo_lista = 4 Then
            txt_cliente = lv_lista.selectedItem
            txt_nombre_cliente = lv_lista.selectedItem.SubItems(1)
            txt_cliente.Enabled = True
            txt_cliente.SetFocus
         End If
      End If
      frm_lista.Visible = False
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
      
   End If
End Sub

Private Sub lv_lista_LostFocus()
   var_ventana = 0
   frm_lista.Visible = False
End Sub

Private Sub txt_agente_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct vcha_age_agente_id, vcha_age_nombre from tb_agentes where vcha_emp_empresa_id = '" + var_empresa + "' or vcha_age_Agente_id = '00100' order by vcha_age_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Agentes"
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

Private Sub txt_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      txt_nombre_cliente.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txt_agente_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_agente) <> "" Then
      rs.Open "select * from vw_establecimientos where vcha_age_agente_id = '" + txt_agente + "'"
      If Not rs.EOF Then
         txt_nombre_agente = rs!VCHA_AGE_NOMBRE
         rs.Close
         txt_agente.Enabled = False
         txt_cliente.Enabled = True
         txt_cliente.SetFocus
      Else
         rs.Close
         MsgBox "Clave de agente incorrecta", vbOKOnly, "ATENCION"
         txt_agente = ""
         txt_nombre_agente = ""
      End If
   End If
End Sub

Private Sub txt_almacen_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_almacen_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      If var_tipo_permiso = 1 Then
         rs.Open "select * from vw_almacen_permiso_1 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' AND VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      Else
         rs.Open "select * from vw_movimientos_almacenes where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' AND VCHA_EMP_eMPRESA_ID = '" + var_empresa + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
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

Private Sub txt_almacen_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      txt_agente.Enabled = True
      txt_agente.SetFocus
   End If
End Sub

Private Sub txt_almacen_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_almacen) <> "" Then
      If var_tipo_permiso = 1 Then
         rs.Open "select * from vw_almacen_permiso_1 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_alm_almacen_id = '" + txt_almacen + "'", cnn, adOpenDynamic, adLockBatchOptimistic
         If Not rs.EOF Then
            txt_nombre_almacen = rs!VCHA_ALM_NOMBRE
            var_almacen_Destino = txt_almacen
            txt_almacen.Enabled = False
            txt_nombre_almacen.Enabled = False
            txt_agente.Enabled = True
         Else
            MsgBox "Clave de almacen incorrecto", vbOKOnly, "ATENCION"
            txt_almacen = ""
            txt_nombre_almacen = ""
         End If
         rs.Close
      Else
         rs.Open "select * from vw_movimientos_almacenes where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' AND VCHA_ALM_ALMACEN_ID = '" + txt_almacen + "' AND VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockBatchOptimistic
         If Not rs.EOF Then
            txt_nombre_almacen = rs!VCHA_ALM_NOMBRE
            var_almacen_Destino = txt_almacen
            txt_almacen.Enabled = False
            txt_nombre_almacen.Enabled = False
            txt_agente.Enabled = True
         Else
            MsgBox "Clave de almacen incorrecto", vbOKOnly, "ATENCION"
            txt_almacen = ""
         End If
         rs.Close
      End If
   End If
End Sub

Private Sub txt_busqueda_folio_GotFocus()
   var_ventana = 1
End Sub

Private Sub txt_busqueda_folio_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If Trim(txt_busqueda_folio) <> "" Then
         If var_numero_folio = CDbl(txt_busqueda_folio) Then
            rsaux2.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_EMO_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
         End If
         rs.Open "select * from tb_encabezado_movimientos where inte_emo_numero = " + txt_busqueda_folio + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' AND VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
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
                  txt_almacen = var_almacen_Destino
                  txt_referencia = rs!vcha_Emo_referencia
                  rsaux2.Open "select * from tb_clientes where vcha_cli_clave_id = '" + rs!vcha_cli_clave_id + "'", cnn, adOpenDynamic, adLockOptimistic
                  txt_cliente = rs!vcha_cli_clave_id
                  txt_nombre_cliente = rsaux2!VCHA_CLI_NOMBRE
                  rsaux2.Close
                  rsaux2.Open "select * from tb_agentes where vcha_age_agente_id = '" + rs!VCHA_AGE_AGENTE_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                  txt_agente = rs!VCHA_AGE_AGENTE_ID
                  txt_nombre_agente = rsaux2!VCHA_AGE_NOMBRE
                  rsaux2.Close
                  rsaux2.Open "select * from tb_establecimientos where vcha_esb_establecimiento_id = '" + rs!vcha_ESB_ESTABLECIMIENTO_id + "'", cnn, adOpenDynamic, adLockOptimistic
                  txt_establecimiento = rs!vcha_ESB_ESTABLECIMIENTO_id
                  txt_nombre_establecimiento = rsaux2!VCHA_ESB_NOMBRE
                  rsaux2.Close
                  txt_cliente.Enabled = False
                  txt_agente.Enabled = False
                  txt_establecimiento.Enabled = False
                  txt_cliente.Enabled = False
                  txt_almacen.Enabled = False
                  txt_referencia.Enabled = False
                  lv_entradas.ListItems.Clear
                  var_primera_vez = False
                  var_numero_folio = rs!INTE_EMO_NUMERO
                  txt_folio = var_numero_folio
                  rsaux.Open "select * from tb_ALMACENES where vcha_alm_almacen_id = '" + var_almacen_Destino + "'", cnn, adOpenDynamic, adLockOptimistic
                  txt_almacen_destino = rsaux(3).Value
                  txt_nombre_almacen = rsaux(3).Value
                  rsaux.Close
                  lbl_total = "0"
                  rsaux.Open "select * from tb_temporal_entradas where inte_ent_numero = " + txt_busqueda_folio + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     While Not rsaux.EOF
                           rsaux2.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rsaux!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux.EOF Then
                              Set list_item = lv_entradas.ListItems.Add(, , rsaux!VCHA_ART_ARTICULO_ID)
                              list_item.SubItems(1) = IIf(IsNull(rsaux2(1).Value), "", rsaux2(1).Value)
                              list_item.SubItems(2) = IIf(IsNull(rsaux!floa_ent_cantidaD), "", rsaux!floa_ent_cantidaD)
                              lbl_total = CStr(CDbl(lbl_total) + IIf(IsNull(rsaux!floa_ent_cantidaD), "", rsaux!floa_ent_cantidaD))
                              rsaux2.Close
                              rsaux.MoveNext:
                           End If
                     Wend
                  End If
                  rsaux.Close
                  
                  If lv_entradas.ListItems.Count > 11 Then
                     lv_entradas.ColumnHeaders(2).Width = 5050.22
                  Else
                     lv_entradas.ColumnHeaders(2).Width = 5300.22
                  End If
                  
                  
                  
                  rsaux4.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 1 WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_EMO_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                  If var_estatus_movimiento = "C" Or var_estatus_movimiento = "I" Then
                     If var_estatus_movimiento = "C" Then
                        Me.cmd_cancelar.Enabled = False
                        Me.cmd_imprimir.Enabled = False
                        lbl_cancelado = "MOVIMIENTO CANCELADO"
                     End If
                     Me.txt_codigo.Enabled = False
                  Else
                     Me.cmd_cancelar.Enabled = True
                     Me.cmd_imprimir.Enabled = True
                     Me.txt_codigo.Enabled = True
                     lbl_cancelado = ""
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

Private Sub txt_busqueda_folio_LostFocus()
   var_ventana = 0
   frm_busqueda.Visible = False
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
         Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
         var_cantidad_eliminar = Val(txt_cantidad_eliminar)
         var_posible_eliminar = True
         If var_cantidad_eliminar <= lv_entradas.selectedItem.SubItems(2) * 1 = True Then
            var_inserta = False
            var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, lv_entradas.selectedItem, 0 - Val(txt_cantidad_eliminar), 2005)
            lv_entradas.selectedItem.SubItems(2) = lv_entradas.selectedItem.SubItems(2) - Val(txt_cantidad_eliminar)
            lbl_total = CStr(CDbl(lbl_total) - Val(txt_cantidad_eliminar))
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
   txt_Cantidad = ""
End Sub

Private Sub txt_cantidad_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If Trim(txt_Cantidad) <> "" Then
         var_cantidad_leida = txt_Cantidad
         txt_foco.Enabled = True
         txt_foco.SetFocus
         lbl_Cantidad.Visible = False
         txt_Cantidad.Visible = False
      End If
   End If
End Sub

Private Sub txt_cliente_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      If Me.txt_agente = "00100" Then
         rs.Open "select distinct vcha_cli_clave_id, vcha_cli_nombre from vw_clientes where vcha_age_agente_id = '" + Me.txt_agente + "' and vcha_emp_empresa_id = '" + var_empresa + "' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
      Else
         rs.Open "select distinct vcha_cli_clave_id, vcha_cli_nombre from vw_clientes where vcha_age_agente_id = '" + Me.txt_agente + "' and vcha_emp_empresa_id = '" + var_empresa + "' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
      End If
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Clientes"
      var_tipo_lista = 4
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

Private Sub txt_cliente_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_cliente) <> "" Then
      If Me.txt_agente = "00100" Then
         rs.Open "select * from vw_clientes where vcha_cli_clave_id = '" + txt_cliente + "' and vcha_age_agente_id = '" + txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
      Else
         rs.Open "select * from vw_clientes where vcha_cli_clave_id = '" + txt_cliente + "' and vcha_age_agente_id = '" + txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
      End If
      If Not rs.EOF Then
         txt_nombre_cliente = rs!VCHA_CLI_NOMBRE
         var_clave_titular = rs!vcha_tit_titular_id
         rs.Close
         txt_establecimiento.Enabled = True
         txt_establecimiento.SetFocus
      Else
         rs.Close
         MsgBox "Clave de Cliente Incorrecta", vbOKOnly, "ATENCION"
         txt_cliente = ""
         txt_nombre_cliente = ""
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
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Dim var_recontable As Integer
   Dim var_cantidad_caja As Integer
   Dim var_caja As String
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   txt_codigo = Trim(txt_codigo)
   If KeyAscii = 13 Then
      If var_empresa = 16 Then
        'If Len(Me.txt_codigo) = 6 Then
        '   Me.txt_codigo = Mid(Me.txt_codigo, 1, 3) + "-" + Mid(Me.txt_codigo, 4, 3) + "-"
        'Else
        '   If Len(Me.txt_codigo) = 7 Then
        '      Me.txt_codigo = Mid(Me.txt_codigo, 1, 3) + "-" + Mid(Me.txt_codigo, 4, 3) + "-" + Mid(Me.txt_codigo, 7, 1)
        '   End If
        'End If
      End If
      var_verificador = True
      If Len(Trim(txt_codigo)) = 12 Then
         Call calcula_verificador(Trim(txt_codigo))
      End If
      If var_empresa <> "02" Then
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
               If IsNull(rs(43).Value) Then
                  var_recontable = 0
               Else
                  var_recontable = rs(43).Value
               End If
               If var_empresa = 31 Then
                  var_recontable = 1
               End If
               
               var_descripcion_articulo = rs(1).Value
               var_costo = IIf(IsNull(rs!mone_Art_costo_estandar), 0, rs!mone_Art_costo_estandar)
               var_precio = rs!mone_Art_precio_base
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
                     If var_cantidad_caja = 0 Then
                        If IsNull(rs(43).Value) Then
                           var_recontable = 0
                        Else
                           var_recontable = rs(43).Value
                        End If
                     Else
                        var_recontable = 0
                     End If
                     If var_empresa = 31 Then
                        var_recontable = 1
                     End If
                     
                     var_descripcion_articulo = rs(1).Value
                     var_costo = IIf(IsNull(rs!mone_Art_costo_estandar), 0, rs!mone_Art_costo_estandar)
                     var_precio = rs!mone_Art_precio_base
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
                      txt_codigo = ""
                      frmmensaje.lbl_mensaje = "El artículo no existe"
                      frmmensaje.Show
                      'MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
                  End If
               Else
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
         txt_codigo = ""
         frmmensaje.lbl_mensaje = "Error en código"
         frmmensaje.Show
         MsgBox "Error en código", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_establecimiento_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_establecimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select DISTINCT vcha_esb_establecimiento_id, vcha_esb_nombre from vw_establecimientos where vcha_cli_clave_id = '" + Me.txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_ESB_ESTABLECIMIENTO_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Establecimientos"
      var_tipo_lista = 3
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

Private Sub txt_establecimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      txt_nombre_establecimiento.SetFocus
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txt_establecimiento_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_establecimiento) <> "" Then
      rs.Open "select * from vw_establecimientos where vcha_age_agente_id = '" + txt_agente + "' and vcha_esb_establecimiento_id = '" + txt_establecimiento + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_establecimiento = rs!VCHA_ESB_NOMBRE
         rs.Close
         txt_establecimiento.Enabled = False
         txt_referencia.Enabled = True
         txt_referencia.SetFocus
      Else
         rs.Close
         txt_establecimiento.Enabled = False
         txt_referencia.Enabled = True
         txt_referencia.SetFocus
      End If
   End If
End Sub

Private Sub txt_foco_GotFocus()
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_ENTRADAS_INSERTA = New TB_TEMPORAL_ENTRADAS_INSERTA
   Set TB_TEMPORAL_ENTRADAS_MODIFICA = New TB_TEMPORAL_ENTRADAS_MODIFICA
   Set TB_BLOQUEOS = New TB_BLOQUEOS
   Dim var_inserta As Boolean
   Dim var_factura As Integer
   Dim var_posible_sip As Boolean
   If Trim(Me.txt_codigo) <> "" Then
   If var_unidad_organizacional = "27" Or var_unidad_organizacional = "28" Or var_unidad_organizacional = "29" Then
      If Mid(Me.txt_codigo, 1, 6) = "646244" Then
         var_codigo_sip = Mid(Me.txt_codigo, 7, 5)
      Else
         var_codigo_sip = Me.txt_codigo
      End If
      
      var_numero_planta = CDbl(var_unidad_organizacional)
      If var_unidad_organizacional = 27 Then
         VAR_PLANTA_CORRECTA = "03"
      End If
      If var_unidad_organizacional = 28 Then
         VAR_PLANTA_CORRECTA = "04"
      End If
      If var_unidad_organizacional = 29 Then
         VAR_PLANTA_CORRECTA = "1"
      End If
      rs.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + var_prefijo + var_codigo_sip + "'", cnn_cantia, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_posible_sip = True
      Else
         var_posible_sip = False
      End If
      rs.Close
   Else
      var_posible_sip = True
   End If
   
   If var_posible_sip = True Then
      If Trim(txt_codigo.Text) <> "" Then
      If var_empresa = "18000" Then
         var_cadena = "SELECT dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO, dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID FROM dbo.TB_CLIENTES INNER JOIN dbo.TB_ENCABEZADO_CARTERA ON dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID INNER JOIN dbo.TB_SALIDAS ON dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO = dbo.TB_SALIDAS.INTE_CAR_NUMERO AND dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO = dbo.TB_SALIDAS.VCHA_CAR_DOCUMENTO AND dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID = dbo.TB_SALIDAS.VCHA_SER_SERIE_ID WHERE (dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID = '" + Me.txt_codigo + "') AND (dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO = 'FA') AND (dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID = '" + Me.txt_cliente + "')"
         rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_posible_cliente = True
         Else
            var_posible_cliente = False
         End If
         rs.Close
      Else
         var_posible_cliente = True
      End If
      If var_posible_cliente = True Then
         
         bandera_suma = False
         If var_primera_vez = True Then
            var_inserta = False
            var_insreta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, Now, var_numero_folio, 0, txt_cliente, "", "", var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, "", "", txt_referencia, txt_establecimiento, "B", var_clave_titular, txt_agente, 0, 0, 0, var_clave_moneda, 1)
            var_numero_folio = var_numero_folio_regreso
            var_global_bloqueado = 1
            var_inserta = False
            var_inserta = TB_BLOQUEOS.Anadir(var_empresa, var_unidad_organizacional, "DEVOLUCION" + Trim(var_clave_movimiento) + Trim(Str(var_numero_folio)), Now, var_clave_usuario_global, fun_NombrePc)
            var_solo_lectura = False
            txt_folio = var_numero_folio
            var_primera_vez = False
         End If
         If var_empresa = "28" Then
            rs.Open "SELECT * FROM TB_DETALLE_LISTA_PRECIOS WHERE VCHA_LIS_LISTA_PRECIOS_ID = '02' AND VCHA_ART_ARTICULO_ID = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_costo = IIf(IsNull(rs!floa_dli_Precio), 0, rs!floa_dli_Precio)
            Else
               var_costo = 0
            End If
            rs.Close
            
            rs.Open "SELECT * FROM TB_DETALLE_LISTA_PRECIOS WHERE VCHA_LIS_LISTA_PRECIOS_ID = '50' AND VCHA_ART_ARTICULO_ID = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_precio = IIf(IsNull(rs!floa_dli_Precio), 0, rs!floa_dli_Precio)
            Else
               var_precio = 0
            End If
            rs.Close
         Else
            rs.Open "select * from tb_existencias where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_costo = IIf(IsNull(rs!floa_exi_costo_2005), 0, rs!floa_exi_costo_2005)
               If var_costo = 0 Then
                  var_costo = IIf(IsNull(rs!FLOA_EXI_COSTO_2004), 0, rs!FLOA_EXI_COSTO_2004)
               End If
            End If
            rs.Close
            
            If var_costo = 0 Then
               rs.Open "select * from tb_existencias where vcha_alm_almacen_id = '8' and vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_costo = IIf(IsNull(rs!floa_exi_costo_2005), 0, rs!floa_exi_costo_2005)
                  If var_costo = 0 Then
                     var_costo = IIf(IsNull(rs!FLOA_EXI_COSTO_2004), 0, rs!FLOA_EXI_COSTO_2004)
                  End If
               End If
               rs.Close
            End If
         
         
         
            If var_costo = 0 Then
               rs.Open "SELECT MONE_ART_COSTO_ESTANDAR FROM TB_ARTICULOS WHERE VCHA_ART_ARTICULO_ID = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_costo = IIf(IsNull(rs!mone_Art_costo_estandar), 0, rs!mone_Art_costo_estandar)
               End If
               rs.Close
            End If
         End If
         Cadena = "select * from TB_TEMPORAL_ENTRADAS where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "' and vcha_emp_empresa_id = '" + var_empresa + "'"
         rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            lbl_total = CStr(CDbl(lbl_total) + var_cantidad_leida)
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
            lbl_total = CStr(CDbl(lbl_total) + var_cantidad_leida)
            var_inserta = TB_TEMPORAL_ENTRADAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida, var_costo, var_precio, "0", "", var_año)
            rs.Close
            Set list_item = lv_entradas.ListItems.Add(, , Trim(txt_codigo))
            list_item.SubItems(1) = var_descripcion_articulo
            list_item.SubItems(2) = var_cantidad_leida
            var_renglon = lv_entradas.ListItems.Count
            Call ilumina_grid
         End If
         txt_codigo.SetFocus
         
      Else
         rsaux11.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux11.EOF Then
            var_DEscripcion = IIf(IsNull(rsaux11!vcha_Art_nombre_español), "", rsaux11!vcha_Art_nombre_español)
         Else
            var_DEscripcion = ""
         End If
         rsaux11.Close
         txt_codigo = ""
         frmmensaje.lbl_mensaje = "El artículo " + var_DEscripcion + " no se le a vendido al cliente"
         frmmensaje.Show 1
      End If
      Else
         rsaux11.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux11.EOF Then
            var_DEscripcion = IIf(IsNull(rsaux11!vcha_Art_nombre_español), "", rsaux11!vcha_Art_nombre_español)
         Else
            var_DEscripcion = ""
         End If
         rsaux11.Close
         txt_codigo = ""
         frmmensaje.lbl_mensaje = "El artículo " + var_DEscripcion + " no se le a vendido al cliente"
         frmmensaje.Show 1
      End If
   Else
      txt_codigo = ""
      Me.txt_foco.Enabled = False
      frmmensaje.lbl_mensaje = "El artículo no existe en el S.I.P."
      frmmensaje.Show 1
      Me.txt_codigo.SetFocus
   End If
   End If
End Sub

Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Me.txt_establecimiento.Enabled = True
      Me.txt_establecimiento.SetFocus
      Me.txt_cliente.Enabled = False
      
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub txt_nombre_agente_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_nombre_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   KeyCode = 0
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct vcha_age_agente_id, vcha_age_nombre from tb_agentes where vcha_emp_empresa_id = '" + var_empresa + "' order by vcha_age_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Agentes"
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

Private Sub txt_nombre_agente_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_almacen_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_nombre_almacen_KeyDown(KeyCode As Integer, Shift As Integer)
   KeyCode = 0
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      If var_tipo_permiso = 1 Then
         rs.Open "select * from vw_almacen_permiso_1 where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' AND VCHA_eMP_EMRPESA_ID = '" + var_empresa + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      Else
         rs.Open "select * from vw_movimientos_almacenes where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' AND VCHA_eMP_EMPRESA_ID = '" + var_empresa + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
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

Private Sub txt_nombre_almacen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txt_agente.Enabled = True Then
         txt_agente.SetFocus
      End If
   End If
End Sub

Private Sub txt_nombre_almacen_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_cliente_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_nombre_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   KeyCode = 0
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct vcha_cli_clave_id, vcha_cli_nombre from vw_establecimientos where vcha_age_agente_id= '" + txt_agente + "' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Clientes"
      var_tipo_lista = 4
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

Private Sub txt_nombre_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If txt_referencia.Enabled = True Then
         txt_referencia.SetFocus
      End If
   End If
End Sub

Private Sub txt_nombre_cliente_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_establecimiento_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_nombre_establecimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   KeyCode = 0
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select DISTINCT vcha_esb_establecimiento_id, vcha_esb_nombre from vw_establecimientos where vcha_age_agente_id = '" + txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_ESB_ESTABLECIMIENTO_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Establecimientos"
      var_tipo_lista = 3
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

Private Sub txt_nombre_establecimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_referencia.Enabled = True
      Me.txt_referencia.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_establecimiento_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   Me.txt_establecimiento.Enabled = False
   Me.txt_nombre_establecimiento.Enabled = False
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Me.cmd_aceptar_pedidos.SetFocus
   End If
   If KeyAscii = 27 Then
      Me.frm_pasar_todo.Visible = False
   End If
End Sub

Private Sub txt_referencia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Trim(txt_referencia) <> "" Then
         txt_codigo.Enabled = True
         txt_codigo.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_pasar_todo.Visible = False
   End If
End Sub

Private Sub txt_referencia_LostFocus()
      If Trim(txt_referencia) <> "" Then
         txt_codigo.Enabled = True
         txt_referencia.Enabled = False
         txt_codigo.SetFocus
      Else
         MsgBox "Debe de introducir una referencia", vbOKOnly, "ATENCION"
      End If
End Sub

Private Sub txt_serie_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_numero.SetFocus
   End If
End Sub

