VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmentradas_devoluciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entradas Devoluciones"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8475
   Icon            =   "frmentradas_devoluciones.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   8475
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1080
      Picture         =   "frmentradas_devoluciones.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   61
      ToolTipText     =   "Aplicar nota de cr?dito"
      Top             =   720
      Width           =   330
   End
   Begin VB.CommandButton cmd_nota_credito 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1410
      Picture         =   "frmentradas_devoluciones.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   60
      ToolTipText     =   "Generar nota de cr?dito"
      Top             =   720
      Width           =   330
   End
   Begin VB.TextBox txt_movimiento 
      Height          =   345
      Left            =   7650
      TabIndex        =   59
      Top             =   30
      Width           =   480
   End
   Begin VB.CheckBox chk_factura 
      Caption         =   "Check1"
      Height          =   255
      Left            =   7290
      TabIndex        =   58
      Top             =   765
      Width           =   420
   End
   Begin VB.Frame frm_pasar_todo 
      Height          =   1500
      Left            =   1935
      TabIndex        =   49
      Top             =   1860
      Width           =   3675
      Begin VB.CommandButton cmd_cancelar_pedidos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   390
         Picture         =   "frmentradas_devoluciones.frx":0ACE
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Cancelar Alt + C"
         Top             =   420
         Width           =   330
      End
      Begin VB.CommandButton cmd_aceptar_pedidos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   60
         Picture         =   "frmentradas_devoluciones.frx":0C18
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Aceptar Alt + A"
         Top             =   420
         Width           =   330
      End
      Begin VB.Frame Frame5 
         Height          =   30
         Left            =   0
         TabIndex        =   55
         Top             =   765
         Width           =   3660
      End
      Begin VB.TextBox txt_numero 
         Height          =   315
         Left            =   2445
         TabIndex        =   52
         Top             =   990
         Width           =   1000
      End
      Begin VB.TextBox txt_serie 
         Height          =   315
         Left            =   735
         TabIndex        =   51
         Top             =   990
         Width           =   1000
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "N?mero:"
         Height          =   195
         Left            =   1785
         TabIndex        =   54
         Top             =   1050
         Width           =   600
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   210
         TabIndex        =   53
         Top             =   1050
         Width           =   405
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         Caption         =   " Factura a Devolver"
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   30
         TabIndex        =   50
         Top             =   120
         Width           =   3600
      End
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   1575
      TabIndex        =   41
      Top             =   1665
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   30
         TabIndex        =   42
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
         TabIndex        =   43
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame frm_busqueda 
      Height          =   960
      Left            =   1275
      TabIndex        =   0
      Top             =   1395
      Width           =   3135
      Begin VB.TextBox txt_busqueda_folio 
         Height          =   315
         Left            =   195
         MaxLength       =   10
         TabIndex        =   18
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
         TabIndex        =   19
         Top             =   120
         Width           =   3060
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmentradas_devoluciones.frx":0D62
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Movimiento Alt + N"
      Top             =   720
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   420
      Picture         =   "frmentradas_devoluciones.frx":0E64
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Buscar Movimiento Alt + B"
      Top             =   720
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   750
      Picture         =   "frmentradas_devoluciones.frx":0F66
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   720
      Width           =   330
   End
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1080
      Picture         =   "frmentradas_devoluciones.frx":1068
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cancelar Movimiento"
      Top             =   720
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7935
      Picture         =   "frmentradas_devoluciones.frx":116A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   720
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   1080
      Index           =   0
      Left            =   6255
      TabIndex        =   25
      Top             =   1110
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
         TabIndex        =   26
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
         TabIndex        =   27
         Top             =   120
         Width           =   1980
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2115
      Index           =   1
      Left            =   90
      TabIndex        =   20
      Top             =   1110
      Width           =   6150
      Begin VB.CommandButton cmd_pasar_todo 
         Height          =   330
         Left            =   5700
         Picture         =   "frmentradas_devoluciones.frx":17A4
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Pasar una factura"
         Top             =   1740
         Width           =   360
      End
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   315
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1080
         Width           =   3750
      End
      Begin VB.TextBox txt_nombre_establecimiento 
         Height          =   315
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1410
         Width           =   3750
      End
      Begin VB.TextBox txt_nombre_agente 
         Height          =   315
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   750
         Width           =   3750
      End
      Begin VB.TextBox txt_nombre_almacen 
         Height          =   315
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   420
         Width           =   3750
      End
      Begin VB.TextBox txt_agente 
         Height          =   315
         Left            =   1290
         TabIndex        =   8
         Top             =   750
         Width           =   1005
      End
      Begin VB.TextBox txt_referencia 
         Height          =   315
         Left            =   1290
         TabIndex        =   14
         Top             =   1740
         Width           =   4380
      End
      Begin VB.TextBox txt_establecimiento 
         Height          =   315
         Left            =   1290
         TabIndex        =   12
         Top             =   1410
         Width           =   1005
      End
      Begin VB.TextBox txt_almacen 
         Height          =   315
         Left            =   1290
         TabIndex        =   6
         Top             =   420
         Width           =   1005
      End
      Begin VB.TextBox txt_cliente 
         Height          =   315
         Left            =   1290
         TabIndex        =   10
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   40
         Top             =   810
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Referencia:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   39
         Top             =   1815
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Establecimiento:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   38
         Top             =   1485
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   495
         Width           =   585
      End
      Begin VB.Label label 
         BackColor       =   &H8000000D&
         Caption         =   " Datos del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   30
         TabIndex        =   22
         Top             =   120
         Width           =   6075
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   21
         Top             =   1155
         Width           =   525
      End
   End
   Begin VB.TextBox txt_foco 
      Height          =   285
      Left            =   9030
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2910
      Width           =   1125
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   45
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
            Picture         =   "frmentradas_devoluciones.frx":18A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_devoluciones.frx":2180
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_devoluciones.frx":2A5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_devoluciones.frx":2FF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_devoluciones.frx":38D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_devoluciones.frx":41AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_devoluciones.frx":4A86
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_devoluciones.frx":4B98
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_devoluciones.frx":4CAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_devoluciones.frx":4DBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_devoluciones.frx":4ECE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmentradas_devoluciones.frx":4FE0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   0
      Left            =   45
      TabIndex        =   24
      Top             =   570
      Width           =   8250
   End
   Begin VB.Frame Frame2 
      Height          =   4005
      Left            =   105
      TabIndex        =   29
      Top             =   3225
      Width           =   8235
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
         TabIndex        =   16
         Top             =   465
         Width           =   1890
      End
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   1785
         TabIndex        =   30
         Top             =   1755
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar 
            Height          =   330
            Left            =   60
            MaxLength       =   10
            TabIndex        =   31
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
            TabIndex        =   32
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
         TabIndex        =   15
         Top             =   405
         Width           =   2640
      End
      Begin MSComctlLib.ListView lv_entradas 
         Height          =   2835
         Left            =   45
         TabIndex        =   33
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
            Text            =   "C?digo"
            Object.Width           =   2478
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripci?n"
            Object.Width           =   9349
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   2328
         EndProperty
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
         TabIndex        =   44
         Top             =   390
         Width           =   3765
      End
      Begin VB.Label lbl_cantidad 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   5535
         TabIndex        =   36
         Top             =   585
         Width           =   675
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Lectura de Art?culos"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   0
         Left            =   30
         TabIndex        =   35
         Top             =   120
         Width           =   8160
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "C?digo del Art?culo:"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   585
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   45
      TabIndex        =   28
      Top             =   975
      Width           =   8250
   End
   Begin VB.Frame Frame4 
      Height          =   1035
      Left            =   6270
      TabIndex        =   45
      Top             =   2190
      Width           =   2040
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
         TabIndex        =   47
         Top             =   525
         Width           =   1830
      End
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
         TabIndex        =   46
         Top             =   150
         Width           =   1830
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
      Left            =   45
      TabIndex        =   37
      Top             =   105
      Width           =   8325
   End
End
Attribute VB_Name = "frmentradas_devoluciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_a?o As Integer
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



Private Sub txt_numero_2_KeyPress()
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   VAR_ZZ = 13
   If VAR_ZZ = 13 Then
      cnn.CommandTimeout = 10000
      Dim var_posible_factura As Boolean
      var_posible_factura = False
      If chk_factura.Value = 0 Then
         var_posible_factura = True
      Else
         If Not IsNumeric(txt_factura) Then
            MsgBox "N?mero de factura incorrecto", vbOKOnly, "ATENCION"
            var_posible_factura = False
         Else
            rs.Open "select * from tb_encabezado_Cartera where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_Car_documento = 'FA' and inte_car_numero =  " + txt_factura, cnn, adOpenDynamic, adLockOptimistic
            If rs.EOF Then
               MsgBox "La factura no existe", vbOKOnly, "ATENCION"
               var_posible_factura = False
            Else
               var_serie_FACTURA = rs!vcha_Ser_Serie_id
               var_cliente_factura = rs!vcha_cli_clave_id
            End If
            rs.Close
            rs.Open "select * from tb_encabezado_movimientos where vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero + " and char_emo_estatus = 'I' and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' ", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_referencia = rs!vcha_emo_referencia
               var_cliente = rs!vcha_cli_clave_id
               var_establecimiento = rs!vcha_ESB_ESTABLECIMIENTO_id
               var_titular = rs!vcha_tit_titular_id
            End If
            rs.Close
            If var_cliente_factura = var_cliente Then
               var_posible_factura = True
            Else
               MsgBox "La factura seleccionada no corresponde al cliente de la devoluci?n", vbOKOnly, "ATENCION"
               var_posible_factura = False
            End If
         End If
      End If
      If var_posible_factura = True Then
         If Trim(txt_numero) <> "" Then
            Dim var_cantidad_pasar As Double
            Dim var_factura As Double
            Dim var_posible As Boolean
            Dim var_consecutivo As Integer
            Dim var_contador_articulos As Double
            Dim var_tipo_busqueda As Integer
            Dim var_grupo As String
            Dim list_item As ListItem
            Dim var_contador As Double
            Dim var_cantidad As Double
            Dim var_nombre_articulo As String
            Dim var_nombre_causa_1 As String
            Dim var_nombre_causa_2 As String
            Dim var_descuento_1 As Double
            Dim var_descuento_2 As Double
            Dim var_descuento_3 As Double
            Dim var_iva As Double
            Dim var_precio As Double
            Dim var_precio_anterior As Double
            Dim var_costo As Double
            Dim var_clave_moneda As String
            Dim var_tipo_cambio_anterior As Double
            Dim var_tipo_Cambio As Double
            Dim var_moneda_local As Integer
            Dim var_numero_factura As Double
            Dim var_lista_precios As String
            Dim var_descuento_volumen As Double
            Dim var_descuento_financiero As Double
            Dim var_descuento_pago As Double
            Dim var_canal_venta As String
            Dim var_iva_canal As Double
            Dim var_almacen_costeo As String
            'Dim var_serie_factura As String
            var_posible = False
            rs.Open "select * from tb_movimientos where vcha_mov_movimiento_id = '" + txt_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
            var_requiere_factura = 0
            If Not rs.EOF Then
               var_requiere_factura = IIf(IsNull(rs!INTE_MOV_DEVOLUCION_FACTURA), 0, rs!INTE_MOV_DEVOLUCION_FACTURA)
            End If
            rs.Close
            rs.Open "select * from tb_almacenes where vcha_emp_empresa_id = '" + var_empresa + "' and inte_alm_costeo = 1", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_almacen_costeo = rs!VCHA_ALM_ALMACEN_ID
            End If
            rs.Close
            rs.Open "select * from tb_encabezado_movimientos where vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero + " and char_emo_estatus = 'I' and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_referencia = rs!vcha_emo_referencia
               var_cliente = rs!vcha_cli_clave_id
               var_establecimiento = rs!vcha_ESB_ESTABLECIMIENTO_id
               var_titular = rs!vcha_tit_titular_id
               var_tipo_busqueda = 1
               If rsaux2.State = 1 Then
                  rsaux2.Close
               End If
               rsaux2.Open "select * from vw_clientes where vcha_cli_clave_id ='" + var_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux2.EOF Then
                  var_tipo_busqueda = IIf(IsNull(rsaux2!INTE_CAN_BUSQUEDA_FACTURA_GRUPO), 1, rsaux2!INTE_CAN_BUSQUEDA_FACTURA_GRUPO)
                  var_grupo = IIf(IsNull(rsaux2!VCHA_GAC_GRUPO_aCTUAL_ID), "", rsaux2!VCHA_GAC_GRUPO_aCTUAL_ID)
                  var_lista_precios = IIf(IsNull(rsaux2!vcha_LIS_LISTA_iD), "", rsaux2!vcha_LIS_LISTA_iD)
                  var_canal_venta = IIf(IsNull(rsaux2!vcha_can_canal_venta_id), "", rsaux2!vcha_can_canal_venta_id)
                  var_iva_canal = IIf(IsNull(rsaux2!FLOA_TPE_IVA), 0, rsaux2!FLOA_TPE_IVA)
               Else
                  var_grupo = ""
                  var_tipo_busqueda = 1
                  var_lista_precios = ""
                  var_canal_venta = ""
                  var_iva_canal = 0
               End If
               rsaux2.Close
               
               If var_canal_venta <> "" Then
                  rsaux2.Open "select floa_gac_descuento_1, floa_gac_descuento_2 from tb_gruposactuales where vcha_gac_grupo_Actual_id = '" + var_grupo + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux2.EOF Then
                     var_descuento_volumen = IIf(IsNull(rsaux2!floa_gac_Descuento_1), 0, rsaux2!floa_gac_Descuento_1)
                     var_descuento_pago = IIf(IsNull(rsaux2!FLOA_GAC_DESCUENTO_2), 0, rsaux2!FLOA_GAC_DESCUENTO_2)
                  Else
                     var_descuento_volumen = 0
                     var_descuento_pago = 0
                  End If
                  rsaux2.Close
               Else
                  var_descuento_volumen = 0
                  var_descuento_financiero = 0
                  var_descuento_pago = 0
               End If
               rs.Close
               
               rs.Open "select * from tb_devoluciones where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
               If rs.EOF Then
                  Set TB_DEVOLUCIONES_INSERTA = New TB_DEVOLUCIONES_INSERTA
                  If rs.State = 1 Then
                     rs.Close
                  End If
                  rs.Open "select * from tb_entradas where vcha_emp_empresa_id = '" + var_empresa + "' and  vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_ent_numero = " + txt_numero + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_almacen = rs!VCHA_ALM_ALMACEN_ID
                     var_consecutivo = 0
                      While Not rs.EOF
                           var_descuento_1 = 0
                           var_descuento_2 = 0
                           var_descuento_3 = 0
                           var_precio = 0
                           var_costo = 0
                           var_iva = 0
                           var_factura = 0
                           var_costo = rs!floa_ent_costo
                           var_a?o = rs!inte_ent_a?o
                           If var_requiere_factura = 1 Then
                              
                              'var_serie_factura = ""
                              If var_empresa = "02" Or var_empresa = "18" Or var_empresa = "31" Then
                                 If Me.chk_factura = 1 Then
                                    var_numero_factura = CDbl(txt_factura)
                                    If var_numero_factura > 0 Then
                                       rsaux2.Open "select * from VW_ORDEN_FECHAS_FACTURAS where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_art_articulo_id = '" + rs!VCHA_ART_ARTICULO_ID + "' and inte_car_numero = " + Str(var_numero_factura) + " and vcha_Ser_Serie_id = '" + var_serie_FACTURA + "' order by dtim_car_fecha desc", cnn, adOpenDynamic, adLockOptimistic
                                       If Not rsaux2.EOF Then
                                          var_clave_moneda = IIf(IsNull(rsaux2!vcha_mon_moneda_id), "", rsaux2!vcha_mon_moneda_id)
                                          var_moneda_local = IIf(IsNull(rsaux2!inte_mon_moneda_local), 0, rsaux2!inte_mon_moneda_local)
                                          If var_moneda_local = 1 Then
                                             var_tipo_Cambio = 1
                                             var_precio = IIf(IsNull(rsaux2!floa_Sal_precio), 0, rsaux2!floa_Sal_precio)
                                             var_descuento_1 = IIf(IsNull(rsaux2!FLOA_SAL_DESCUENTO_1), 0, rsaux2!FLOA_SAL_DESCUENTO_1)
                                             var_descuento_2 = IIf(IsNull(rsaux2!FLOA_SAL_DESCUENTO_2), 0, rsaux2!FLOA_SAL_DESCUENTO_2)
                                             var_iva = IIf(IsNull(rsaux2!floa_car_porcentaje_iva), 0, rsaux2!floa_car_porcentaje_iva)
                                             var_factura = IIf(IsNull(rsaux2!inte_Car_numero), 0, rsaux2!inte_Car_numero)
                                             rsaux4.Open "SELECT max(FLOA_RCO_DESCUENTO_APLICAR) FROM TB_rELACION_COBRANZA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_CAR_NUMERO =  " + CStr(var_factura) + " AND VCHA_CAR_DOCUMENTO= 'FA' and vcha_Ser_serie_id = '" + rsaux2!vcha_Ser_Serie_id + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
                                             If Not rsaux4.EOF Then
                                                var_descuento_3 = IIf(IsNull(rsaux4(0).Value), 0, rsaux4(0).Value)
                                             Else
                                                var_descuento_3 = 0
                                             End If
                                             rsaux4.Close
                                             var_serie_FACTURA = IIf(IsNull(rsaux2!vcha_Ser_Serie_id), "", rsaux2!vcha_Ser_Serie_id)
                                          Else
                                             rsaux3.Open "select * from vw_tipocambio_fecha where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
                                             If Not rsaux3.EOF Then
                                                var_tipo_Cambio = IIf(IsNull(rsaux3!mone_tca_importe), 1, rsaux3!mone_tca_importe)
                                                var_tipo_cambio_anterior = IIf(IsNull(rsaux2!floa_car_tipo_cambio), 1, rsaux2!floa_car_tipo_cambio)
                                                var_precio = IIf(IsNull(rsaux2!floa_Sal_precio), 0, rsaux2!floa_Sal_precio)
                                                var_precio_anterior = var_precio / var_tipo_cambio_anterior
                                                var_precio = var_precio_anterior * var_tipo_Cambio
                                                var_descuento_1 = IIf(IsNull(rsaux2!FLOA_SAL_DESCUENTO_1), 0, rsaux2!FLOA_SAL_DESCUENTO_1)
                                                var_descuento_2 = IIf(IsNull(rsaux2!FLOA_SAL_DESCUENTO_2), 0, rsaux2!FLOA_SAL_DESCUENTO_2)
                                                rsaux4.Open "SELECT max(FLOA_RCO_DESCUENTO_APLICAR) FROM TB_rELACION_COBRANZA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_CAR_NUMERO =  " + CStr(var_factura) + " AND VCHA_CAR_DOCUMENTO= 'FA' and vcha_Ser_serie_id = '" + rsaux2!vcha_Ser_Serie_id + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
                                                If Not rsaux4.EOF Then
                                                   var_descuento_3 = IIf(IsNull(rsaux4(0).Value), 0, rsaux4(0).Value)
                                                Else
                                                   var_descuento_3 = 0
                                                End If
                                                rsaux4.Close
                                                var_iva = IIf(IsNull(rsaux2!floa_car_porcentaje_iva), 0, rsaux2!floa_car_porcentaje_iva)
                                                var_factura = IIf(IsNull(rsaux2!inte_Car_numero), 0, rsaux2!inte_Car_numero)
                                                var_serie_FACTURA = IIf(IsNull(rsaux2!vcha_Ser_Serie_id), "", rsaux2!vcha_Ser_Serie_id)
                                             Else
                                                GoTo salir:
                                             End If
                                             rsaux3.Close
                                          End If
                                       End If
                                       rsaux2.Close
                                    End If
                                 Else
                                    ''' proceso normal
                                    var_serie_FACTURA = ""
                                    
                                    'If var_tipo_busqueda = 1 And Trim(var_grupo) <> "" Then
                                    '   rsaux3.Open "select * from VW_ORDEN_FACTURAS_FECHA_GRUPO where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_gac_grupo_actual_id ='" + var_grupo + "' and vcha_art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                    '   If Not rsaux3.EOF Then
                                    '      var_numero_factura = rsaux3!inte_car_numero
                                    '      var_serie_factura = rsaux3!vcha_ser_serie_id
                                    '   End If
                                    '   rsaux3.Close
                                    'Else
                                    '   rsaux3.Open "select * from VW_ORDEN_FECHA_FACTURAS_TITULAR where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_tit_titular_id ='" + var_titular + "' and vcha_art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                    '   If Not rsaux3.EOF Then
                                    '      var_numero_factura = rsuax3!inte_car_numero
                                    '      var_serie_factura = rsaux3!vcha_ser_serie_id
                                    '   End If
                                    '   rsaux3.Close
                                    'End If
                                    var_numero_factura = 1
                                    If var_numero_factura > 0 Then
                                       var_numero_factura = 0
                                       'rsaux2.Open "select * from VW_ORDEN_FECHAS_FACTURAS where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_art_articulo_id = '" + rs!vcha_Art_articulo_id + "' and inte_car_numero = " + Str(var_numero_factura) + " and vcha_Ser_Serie_id = '" + var_serie_factura + "'", cnn, adOpenDynamic, adLockOptimistic
                                       '1
                                       If rsaux2.State = 1 Then
                                          rsaux2.Close
                                       End If
                                       var_cadena = "SELECT TOP 1 dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_PORCENTAJE_IVA, "
                                       var_cadena = var_cadena + " dbo.TB_SALIDAS.vcha_ser_serie_id,  dbo.TB_SALIDAS.INTE_CAR_NUMERO, dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA, dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_1, dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_2, dbo.TB_SALIDAS.FLOA_SAL_PRECIO, dbo.TB_MONEDAS.VCHA_MON_MONEDA_ID, dbo.TB_MONEDAS.INTE_MON_MONEDA_LOCAL FROM dbo.TB_SALIDAS INNER JOIN dbo.TB_ENCABEZADO_CARTERA ON dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_SALIDAS.VCHA_CAR_DOCUMENTO = dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO AND dbo.TB_SALIDAS.INTE_CAR_NUMERO = dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO AND dbo.TB_SALIDAS.VCHA_SER_SERIE_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID INNER JOIN dbo.TB_MONEDAS ON dbo.TB_ENCABEZADO_CARTERA.VCHA_MON_MONEDA_ID = dbo.TB_MONEDAS.VCHA_MON_MONEDA_ID WHERE (dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO = 'FA') AND (dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND "
                                       var_cadena = var_cadena + " (dbo.TB_ENCABEZADO_CARTERA.VCHA_TIT_TITULAR_ID = '" + var_titular + "') AND (dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID = '" + rs!VCHA_ART_ARTICULO_ID + "') ORDER BY dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA DESC"
                                       'MsgBox " (dbo.TB_ENCABEZADO_CARTERA.VCHA_TIT_TITULAR_ID = '" + var_titular + "') AND (dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID = '" + rs!vcha_Art_Articulo_id + "') ORDER BY dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA DESC"
                                        rsaux2.Open var_cadena
                                       'rsaux2.Open "select * from VW_ORDEN_FECHAS_FACTURAS where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_art_articulo_id = '" + rs!vcha_Art_Articulo_id + "'  AND VCHA_TIT_TITULAR_ID = '" + var_titular + "'  order by dtim_car_fecha desc", cnn, adOpenDynamic, adLockOptimistic
                                       
                                       
                                       
                                       If Not rsaux2.EOF Then
                                          
                                          var_numero_factura = IIf(IsNull(rsaux2!inte_Car_numero), 0, rsaux2!inte_Car_numero)
                                          var_serie_FACTURA = IIf(IsNull(rsaux2!vcha_Ser_Serie_id), "", rsaux2!vcha_Ser_Serie_id)
                                          var_clave_moneda = IIf(IsNull(rsaux2!vcha_mon_moneda_id), "", rsaux2!vcha_mon_moneda_id)
                                          var_moneda_local = IIf(IsNull(rsaux2!inte_mon_moneda_local), 0, rsaux2!inte_mon_moneda_local)
                                          If var_moneda_local = 1 Then
                                             var_tipo_Cambio = 1
                                             var_precio = IIf(IsNull(rsaux2!floa_Sal_precio), 0, rsaux2!floa_Sal_precio)
                                             var_descuento_1 = IIf(IsNull(rsaux2!FLOA_SAL_DESCUENTO_1), 0, rsaux2!FLOA_SAL_DESCUENTO_1)
                                             var_descuento_2 = IIf(IsNull(rsaux2!FLOA_SAL_DESCUENTO_2), 0, rsaux2!FLOA_SAL_DESCUENTO_2)
                                             var_iva = IIf(IsNull(rsaux2!floa_car_porcentaje_iva), 0, rsaux2!floa_car_porcentaje_iva)
                                             var_factura = IIf(IsNull(rsaux2!inte_Car_numero), 0, rsaux2!inte_Car_numero)
                                             cnn.CommandTimeout = 360
                                             If rsaux4.State = 1 Then
                                                rsaux4.Close
                                             End If
                                             rsaux4.Open "SELECT max(FLOA_RCO_DESCUENTO_APLICAR) FROM TB_rELACION_COBRANZA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_CAR_NUMERO =  " + CStr(var_factura) + " AND VCHA_CAR_DOCUMENTO= 'FA' and vcha_Ser_serie_id = '" + rsaux2!vcha_Ser_Serie_id + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
                                             If Not rsaux4.EOF Then
                                                var_descuento_3 = IIf(IsNull(rsaux4(0).Value), 0, rsaux4(0).Value)
                                             Else
                                                var_descuento_3 = 0
                                             End If
                                             rsaux4.Close
                                             var_serie_FACTURA = IIf(IsNull(rsaux2!vcha_Ser_Serie_id), "", rsaux2!vcha_Ser_Serie_id)
                                           Else
                                             rsaux3.Open "select * from vw_tipocambio_fecha where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
                                             If Not rsaux3.EOF Then
                                                var_tipo_Cambio = IIf(IsNull(rsaux3!mone_tca_importe), 1, rsaux3!mone_tca_importe)
                                                var_tipo_cambio_anterior = IIf(IsNull(rsaux2!floa_car_tipo_cambio), 1, rsaux2!floa_car_tipo_cambio)
                                                var_precio = IIf(IsNull(rsaux2!floa_Sal_precio), 0, rsaux2!floa_Sal_precio)
                                                var_precio_anterior = var_precio / var_tipo_cambio_anterior
                                                var_precio = var_precio_anterior * var_tipo_Cambio
                                                var_descuento_1 = IIf(IsNull(rsaux2!FLOA_SAL_DESCUENTO_1), 0, rsaux2!FLOA_SAL_DESCUENTO_1)
                                                var_descuento_2 = IIf(IsNull(rsaux2!FLOA_SAL_DESCUENTO_2), 0, rsaux2!FLOA_SAL_DESCUENTO_2)
                                                rsaux4.Open "SELECT max(FLOA_RCO_DESCUENTO_APLICAR) FROM TB_rELACION_COBRANZA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_CAR_NUMERO =  " + CStr(var_factura) + " AND VCHA_CAR_DOCUMENTO= 'FA' and vcha_Ser_serie_id = '" + rsaux2!vcha_Ser_Serie_id + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
                                                If Not rsaux4.EOF Then
                                                   var_descuento_3 = IIf(IsNull(rsaux4(0).Value), 0, rsaux4(0).Value)
                                                Else
                                                   var_descuento_3 = 0
                                                End If
                                                rsaux4.Close
                                                var_iva = IIf(IsNull(rsaux2!floa_car_porcentaje_iva), 0, rsaux2!floa_car_porcentaje_iva)
                                                var_factura = IIf(IsNull(rsaux2!inte_Car_numero), 0, rsaux2!inte_Car_numero)
                                                var_serie_FACTURA = IIf(IsNull(rsaux2!vcha_Ser_Serie_id), "", rsaux2!vcha_Ser_Serie_id)
                                             Else
                                                GoTo salir:
                                             End If
                                             rsaux3.Close
                                          End If
                                       Else
                                          rsaux9.Open "select isnull(floa_Gac_descuento_1,0), isnull(floa_gac_descuento_2,0) from tb_gruposactuales where vcha_gac_grupo_actual_id = '" + var_grupo + "'", cnn, adOpenDynamic, adLockOptimistic
                                          If Not rsaux9.EOF Then
                                             var_descuento_1 = rsaux9(0).Value
                                             var_descuento_2 = rsaux9(1).Value
                                             
                                          End If
                                          rsaux9.Close
                                       
                                       End If
                                       rsaux2.Close
                                    Else
                                       If Trim(var_lista_precios) <> "" Then
                                          rsaux3.Open "select * from vw_detalle_lista_precios where vcha_lis_lista_precios_id = '" + var_lista_precios + "' and vcha_Art_Articulo_id = '" + rs!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                                          If Not rsaux3.EOF Then
                                             var_clave_moneda = IIf(IsNull(rsaux3!vcha_mon_moneda), "", rsaux3!vcha_mon_moneda)
                                             var_moneda_local = IIf(IsNull(rsaux3!inte_mon_moneda_local), 0, rsaux3!inte_mon_moneda_local)
                                             If Not rsaux3.EOF Then
                                                If var_moneda_local = 1 Then
                                                   var_tipo_Cambio = 1
                                                   var_precio = IIf(IsNull(rsaux3!floa_dli_Precio), 0, rsaux3!floa_dli_Precio)
                                                   var_descuento_1 = var_descuento_volumen
                                                   var_descuento_2 = var_descuento_pago
                                                   var_descuento_3 = var_descuento_financiero
                                                   var_iva = var_iva_canal
                                                Else
                                                   rsaux.Open "select * from vw_tipocambio_fecha where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
                                                   If Not rsaux.EOF Then
                                                      var_tipo_Cambio = IIf(IsNull(rsaux!mone_tca_importe), 1, rsaux!mone_tca_importe)
                                                      var_precio = IIf(IsNull(rsaux3!floa_dli_Precio), 0, rsaux3!floa_dli_Precio * var_tipo_Cambio)
                                                      var_descuento_1 = var_descuento_volumen
                                                      var_descuento_2 = var_descuento_pago
                                                      var_descuento_3 = var_descuento_financiero
                                                      var_iva = var_iva_canal
                                                   Else
                                                      GoTo salir:
                                                   End If
                                                   rsaux.Close
                                                End If
                                             Else
                                                GoTo salir_3:
                                             End If
                                             rsaux3.Close
                                          Else
                                             rsaux3.Close
                                             Cadena = "SELECT dbo.TB_DETALLE_LISTA_PRECIOS.VCHA_LIS_LISTA_PRECIOS_ID, dbo.TB_DETALLE_LISTA_PRECIOS.VCHA_ART_ARTICULO_ID,   dbo.TB_DETALLE_LISTA_PRECIOS.FLOA_DLI_PRECIO, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPA?OL, "
                                             Cadena = Cadena + " dbo.TB_LISTADEPRECIOS.DTIM_LIS_FECHA_INICIO , dbo.TB_LISTADEPRECIOS.DTIM_LIS_FECHA_FIN, dbo.TB_LISTADEPRECIOS.VCHA_MON_MONEDA, dbo.TB_MONEDAS.INTE_MON_MONEDA_LOCAL"
                                             Cadena = Cadena + " FROM dbo.TB_DETALLE_LISTA_PRECIOS INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_DETALLE_LISTA_PRECIOS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID INNER JOIN dbo.TB_LISTADEPRECIOS ON dbo.TB_DETALLE_LISTA_PRECIOS.VCHA_LIS_LISTA_PRECIOS_ID = dbo.TB_LISTADEPRECIOS.VCHA_LIS_LISTA_ID INNER JOIN dbo.TB_MONEDAS ON dbo.TB_LISTADEPRECIOS.VCHA_MON_MONEDA = dbo.TB_MONEDAS.VCHA_MON_MONEDA_ID  where dbo.TB_DETALLE_LISTA_PRECIOS.vcha_lis_lista_precios_id = '" + var_lista_precios + "' and dbo.TB_DETALLE_LISTA_PRECIOS.vcha_Art_Articulo_id = '" + rs!VCHA_ART_ARTICULO_ID + "'"
                                             rsaux3.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                             If Not rsaux3.EOF Then
                                                var_clave_moneda = IIf(IsNull(rsaux3!vcha_mon_moneda), "", rsaux3!vcha_mon_moneda)
                                                var_moneda_local = IIf(IsNull(rsaux3!inte_mon_moneda_local), 0, rsaux3!inte_mon_moneda_local)
                                                If Not rsaux3.EOF Then
                                                   If var_moneda_local = 1 Then
                                                      var_tipo_Cambio = 1
                                                      var_precio = IIf(IsNull(rsaux3!floa_dli_Precio), 0, rsaux3!floa_dli_Precio)
                                                      var_descuento_1 = var_descuento_volumen
                                                      var_descuento_2 = var_descuento_pago
                                                      var_descuento_3 = var_descuento_financiero
                                                      var_iva = var_iva_canal
                                                   Else
                                                      rsaux.Open "select * from vw_tipocambio_fecha where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
                                                      If Not rsaux.EOF Then
                                                         var_tipo_Cambio = IIf(IsNull(rsaux!mone_tca_importe), 1, rsaux!mone_tca_importe)
                                                         var_precio = IIf(IsNull(rsaux3!floa_dli_Precio), 0, rsaux3!floa_dli_Precio * var_tipo_Cambio)
                                                         var_descuento_1 = var_descuento_volumen
                                                         var_descuento_2 = var_descuento_pago
                                                         var_descuento_3 = var_descuento_financiero
                                                         var_iva = var_iva_canal
                                                      Else
                                                         GoTo salir:
                                                      End If
                                                      rsaux.Close
                                                   End If
                                                Else
                                                   GoTo salir_3:
                                                End If
                                             End If
                                             rsaux3.Close
                                          End If
                                       Else
                                          GoTo salir_2:
                                       End If
                                    End If
                                 End If
                              End If
                              If var_empresa = "03" Then
                                 If Me.chk_factura = 1 Then
                                    var_numero_factura = CDbl(txt_factura)
                                    If var_numero_factura > 0 Then
                                       rsaux2.Open "select * from VW_ORDEN_FECHAS_FACTURAS where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_art_articulo_id = '" + rs!VCHA_ART_ARTICULO_ID + "' and inte_car_numero = " + Str(var_numero_factura) + " and vcha_Ser_Serie_id = '" + var_serie_FACTURA + "'  order by dtim_car_fecha desc", cnn, adOpenDynamic, adLockOptimistic
                                       If Not rsaux2.EOF Then
                                          var_clave_moneda = IIf(IsNull(rsaux2!vcha_mon_moneda_id), "", rsaux2!vcha_mon_moneda_id)
                                          var_moneda_local = IIf(IsNull(rsaux2!inte_mon_moneda_local), 0, rsaux2!inte_mon_moneda_local)
                                          If var_moneda_local = 1 Then
                                             var_tipo_Cambio = 1
                                             var_precio = IIf(IsNull(rsaux2!floa_Sal_precio), 0, rsaux2!floa_Sal_precio)
                                             var_descuento_1 = IIf(IsNull(rsaux2!FLOA_SAL_DESCUENTO_1), 0, rsaux2!FLOA_SAL_DESCUENTO_1)
                                             var_descuento_2 = IIf(IsNull(rsaux2!FLOA_SAL_DESCUENTO_2), 0, rsaux2!FLOA_SAL_DESCUENTO_2)
                                             var_iva = IIf(IsNull(rsaux2!floa_car_porcentaje_iva), 0, rsaux2!floa_car_porcentaje_iva)
                                             var_factura = IIf(IsNull(rsaux2!inte_Car_numero), 0, rsaux2!inte_Car_numero)
                                             rsaux4.Open "SELECT max(FLOA_RCO_DESCUENTO_APLICAR) FROM TB_rELACION_COBRANZA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_CAR_NUMERO =  " + CStr(var_factura) + " AND VCHA_CAR_DOCUMENTO= 'FA' and vcha_Ser_serie_id = '" + rsaux2!vcha_Ser_Serie_id + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
                                             If Not rsaux4.EOF Then
                                                var_descuento_3 = IIf(IsNull(rsaux4(0).Value), 0, rsaux4(0).Value)
                                             Else
                                                var_descuento_3 = 0
                                             End If
                                             rsaux4.Close
                                             var_serie_FACTURA = IIf(IsNull(rsaux2!vcha_Ser_Serie_id), "", rsaux2!vcha_Ser_Serie_id)
                                          Else
                                             rsaux3.Open "select * from vw_tipocambio_fecha where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
                                             If Not rsaux3.EOF Then
                                                var_tipo_Cambio = IIf(IsNull(rsaux3!mone_tca_importe), 1, rsaux3!mone_tca_importe)
                                                var_tipo_cambio_anterior = IIf(IsNull(rsaux2!floa_car_tipo_cambio), 1, rsaux2!floa_car_tipo_cambio)
                                                var_precio = IIf(IsNull(rsaux2!floa_Sal_precio), 0, rsaux2!floa_Sal_precio)
                                                var_precio_anterior = var_precio / var_tipo_cambio_anterior
                                                var_precio = var_precio_anterior * var_tipo_Cambio
                                                var_descuento_1 = IIf(IsNull(rsaux2!FLOA_SAL_DESCUENTO_1), 0, rsaux2!FLOA_SAL_DESCUENTO_1)
                                                var_descuento_2 = IIf(IsNull(rsaux2!FLOA_SAL_DESCUENTO_2), 0, rsaux2!FLOA_SAL_DESCUENTO_2)
                                                rsaux4.Open "SELECT max(FLOA_RCO_DESCUENTO_APLICAR) FROM TB_rELACION_COBRANZA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_CAR_NUMERO =  " + CStr(var_factura) + " AND VCHA_CAR_DOCUMENTO= 'FA' and vcha_Ser_serie_id = '" + rsaux2!vcha_Ser_Serie_id + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
                                                'empieza descuento 3
                                                If Not rsaux4.EOF Then
                                                   var_descuento_3 = IIf(IsNull(rsaux4(0).Value), 0, rsaux4(0).Value)
                                                Else
                                                   var_descuento_3 = 0
                                                End If
                                                rsaux4.Close
                                                var_iva = IIf(IsNull(rsaux2!floa_car_porcentaje_iva), 0, rsaux2!floa_car_porcentaje_iva)
                                                var_factura = IIf(IsNull(rsaux2!inte_Car_numero), 0, rsaux2!inte_Car_numero)
                                                var_serie_FACTURA = IIf(IsNull(rsaux2!vcha_Ser_Serie_id), "", rsaux2!vcha_Ser_Serie_id)
                                             Else
                                                GoTo salir:
                                             End If
                                             rsaux3.Close
                                          End If
                                       End If
                                       rsaux2.Close
                                    End If
                                 Else
                                    If Trim(var_lista_precios) <> "" Then
                                        rsaux3.Open "select * from vw_detalle_lista_precios where vcha_lis_lista_precios_id = '" + var_lista_precios + "' and vcha_Art_Articulo_id = '" + rs!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                                        If Not rsaux3.EOF Then
                                           var_clave_moneda = IIf(IsNull(rsaux3!vcha_mon_moneda), "", rsaux3!vcha_mon_moneda)
                                           var_moneda_local = IIf(IsNull(rsaux3!inte_mon_moneda_local), 0, rsaux3!inte_mon_moneda_local)
                                           If Not rsaux3.EOF Then
                                              If var_moneda_local = 1 Then
                                                 var_tipo_Cambio = 1
                                                 var_precio = IIf(IsNull(rsaux3!floa_dli_Precio), 0, rsaux3!floa_dli_Precio)
                                                 var_descuento_1 = var_descuento_volumen
                                                 var_descuento_2 = var_descuento_pago
                                                 var_descuento_3 = var_descuento_financiero
                                                 var_iva = var_iva_canal
                                              Else
                                                 rsaux.Open "select * from vw_tipocambio_fecha where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
                                                 If Not rsaux.EOF Then
                                                    var_tipo_Cambio = IIf(IsNull(rsaux!mone_tca_importe), 1, rsaux!mone_tca_importe)
                                                    var_precio = IIf(IsNull(rsaux3!floa_dli_Precio), 0, rsaux3!floa_dli_Precio * var_tipo_Cambio)
                                                    var_descuento_1 = var_descuento_volumen
                                                    var_descuento_2 = var_descuento_pago
                                                    var_descuento_3 = var_descuento_financiero
                                                    var_iva = var_iva_canal
                                                 Else
                                                    GoTo salir:
                                                 End If
                                                 rsaux.Close
                                              End If
                                           Else
                                              GoTo salir_3:
                                           End If
                                           rsaux3.Close
                                        Else
                                           rsaux3.Close
                                           Cadena = "SELECT dbo.TB_DETALLE_LISTA_PRECIOS.VCHA_LIS_LISTA_PRECIOS_ID, dbo.TB_DETALLE_LISTA_PRECIOS.VCHA_ART_ARTICULO_ID,   dbo.TB_DETALLE_LISTA_PRECIOS.FLOA_DLI_PRECIO, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPA?OL, "
                                           Cadena = Cadena + " dbo.TB_LISTADEPRECIOS.DTIM_LIS_FECHA_INICIO , dbo.TB_LISTADEPRECIOS.DTIM_LIS_FECHA_FIN, dbo.TB_LISTADEPRECIOS.VCHA_MON_MONEDA, dbo.TB_MONEDAS.INTE_MON_MONEDA_LOCAL"
                                           Cadena = Cadena + " FROM dbo.TB_DETALLE_LISTA_PRECIOS INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_DETALLE_LISTA_PRECIOS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID INNER JOIN dbo.TB_LISTADEPRECIOS ON dbo.TB_DETALLE_LISTA_PRECIOS.VCHA_LIS_LISTA_PRECIOS_ID = dbo.TB_LISTADEPRECIOS.VCHA_LIS_LISTA_ID INNER JOIN dbo.TB_MONEDAS ON dbo.TB_LISTADEPRECIOS.VCHA_MON_MONEDA = dbo.TB_MONEDAS.VCHA_MON_MONEDA_ID  where dbo.TB_DETALLE_LISTA_PRECIOS.vcha_lis_lista_precios_id = '" + var_lista_precios + "' and dbo.TB_DETALLE_LISTA_PRECIOS.vcha_Art_Articulo_id = '" + rs!VCHA_ART_ARTICULO_ID + "'"
                                           rsaux3.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                           If Not rsaux3.EOF Then
                                              var_clave_moneda = IIf(IsNull(rsaux3!vcha_mon_moneda), "", rsaux3!vcha_mon_moneda)
                                              var_moneda_local = IIf(IsNull(rsaux3!inte_mon_moneda_local), 0, rsaux3!inte_mon_moneda_local)
                                              If Not rsaux3.EOF Then
                                                 If var_moneda_local = 1 Then
                                                    var_tipo_Cambio = 1
                                                    var_precio = IIf(IsNull(rsaux3!floa_dli_Precio), 0, rsaux3!floa_dli_Precio)
                                                    var_descuento_1 = var_descuento_volumen
                                                    var_descuento_2 = var_descuento_pago
                                                    var_descuento_3 = var_descuento_financiero
                                                    var_iva = var_iva_canal
                                                 Else
                                                    rsaux.Open "select * from vw_tipocambio_fecha where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
                                                    If Not rsaux.EOF Then
                                                       var_tipo_Cambio = IIf(IsNull(rsaux!mone_tca_importe), 1, rsaux!mone_tca_importe)
                                                       var_precio = IIf(IsNull(rsaux3!floa_dli_Precio), 0, rsaux3!floa_dli_Precio * var_tipo_Cambio)
                                                       var_descuento_1 = var_descuento_volumen
                                                       var_descuento_2 = var_descuento_pago
                                                       var_descuento_3 = var_descuento_financiero
                                                       var_iva = var_iva_canal
                                                    Else
                                                       GoTo salir:
                                                    End If
                                                    rsaux.Close
                                                 End If
                                              Else
                                                 GoTo salir_3:
                                              End If
                                           End If
                                           rsaux3.Close
                                        End If
                                     Else
                                        GoTo salir_2:
                                     End If
                                  End If
                               End If
                            End If
                            If Trim(var_clave_moneda) = "" Then
                               rsaux2.Open "select * from tb_monedas where inte_mon_moneda_local = 1", cnn, adOpenDynamic, adLockOptimistic
                               var_clave_moneda = rsaux2!vcha_mon_moneda_id
                               var_tipo_Cambio = 1
                               rsaux2.Close
                            End If
                            If var_precio = 0 Then
                               If var_empresa = "06" Then
                                  rsaux2.Open "SELECT * FROM TB_dETALLE_LISTA_PRECIOS WHERE VCHA_LIS_LISTA_PRECIOS_ID = '" + var_lista_precios + "' AND VCHA_aRT_ARTICULO_ID = '" + rs!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                                  If Not rsaux2.EOF Then
                                     var_precio = IIf(IsNull(rsaux2!floa_dli_Precio), 0, rsaux2!floa_dli_Precio)
                                  Else
                                     var_precio = 0
                                  End If
                                  rsaux2.Close
                                  If var_precio = 0 Then
                                     rsaux2.Open "select * from tb_Articulos where vcha_art_articulo_id = '" + rs!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                                     If Not rsaux2.EOF Then
                                        var_precio = IIf(IsNull(rsaux2!mone_Art_precio_base), 0, rsaux2!mone_Art_precio_base)
                                     End If
                                     rsaux2.Close
                                  End If
                               Else
                                  rsaux2.Open "select * from tb_Articulos where vcha_art_articulo_id = '" + rs!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                                  If Not rsaux2.EOF Then
                                     var_precio = IIf(IsNull(rsaux2!mone_Art_precio_base), 0, rsaux2!mone_Art_precio_base)
                                  End If
                                  rsaux2.Close
                               End If
                            End If
                            var_contador = 0
                            var_cantidad = rs!floa_ent_cantidaD
                            'For var_contador = 1 To var_cantidad
                            var_contador = var_cantidad
                            var_cantidad_pasar = 0
                             While var_contador > 0
                                   If var_empresa <> "16" Then
                                      If var_empresa <> "06" Then
                                         If var_contador >= 1 Then
                                            var_cantidad_pasar = 1
                                            var_contador = var_contador - 1
                                         Else
                                            var_cantidad_pasar = var_contador
                                            var_contador = 0
                                         End If
                                      End If
                                   End If
                                   var_consecutivo = var_consecutivo + 1
                                   If var_requiere_factura = 1 Then
                                      If var_iva = 0 Then
                                         rsaux5.Open "select * from vw_clientes where vcha_cli_clave_id = '" + var_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                                         If Not rsaux5.EOF Then
                                            var_iva = IIf(IsNull(rsaux5!FLOA_TPE_IVA), 0, rsaux5!FLOA_TPE_IVA)
                                         End If
                                         'If var_iva = 0 Then
                                         '   If var_empresa = "02" Or var_empresa = "18" Or var_empresa = "16" Or var_empresa = "06" Or var_empresa = "31" Or var_empresa = "15" Then
                                         '      var_iva = 15
                                         '   End If
                                         'End If
                                         rsaux6.Open "SELECT top 1 FLOA_IVA_PORCENTAJE FROM TB_IVA WHERE VCHA_eMP_EMPRESA_ID = " + var_empresa + " AND GETDATE() BETWEEN DTIM_IVA_LIMITE_INFERIOR AND DTIM_IVA_LIMITE_SUPERIOR", cnn, adOpenDynamic, adLockOptimistic
                                         If Not rsaux6.EOF Then
                                            var_iva = IIf(IsNull(rsaux6(0).Value), 16, rsaux6(0).Value)
                                         Else
                                            var_iva = 16
                                         End If
                                         rsaux6.Close
                                         rsaux5.Close
                                      End If
                                   End If
                                   
                                   
                                   If var_empresa = "16" Or var_empresa = "06" Then
                                      var_cantidad_pasar = var_contador
                                      var_contador = 0
                                      If var_iva = "15" Then
                                         var_iva = "16"
                                      End If
                                      var_inserta = TB_DEVOLUCIONES_INSERTA.Anadir(CStr(var_empresa), CStr(var_unidad_organizacional), CStr(var_almacen), CStr(txt_movimiento), CDbl(txt_numero), CStr(rs!VCHA_ART_ARTICULO_ID), 0, 0, "", CInt(var_consecutivo), "", CDbl(var_costo), CDbl(var_precio), CDbl(var_descuento_1), CDbl(var_descuento_2), CDbl(var_descuento_3), CDbl(var_iva), CDbl(var_factura), CStr(var_referencia), CStr(var_clave_moneda), CDbl(var_tipo_Cambio), CStr(var_serie_FACTURA), CInt(var_a?o))
                                      rsaux5.Open "update tb_Devoluciones set floa_dev_cantidad = " + CStr(var_cantidad_pasar) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen + "' and vcha_mov_movimiento_id = '" + Me.txt_movimiento + "' and inte_emo_numero = " + Me.txt_numero + " and vcha_art_articulo_id = '" + rs!VCHA_ART_ARTICULO_ID + "' and inte_cde_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                                   Else
                                      If var_iva = "15" Then
                                         var_iva = "16"
                                      End If
                                      var_inserta = TB_DEVOLUCIONES_INSERTA.Anadir(CStr(var_empresa), CStr(var_unidad_organizacional), CStr(var_almacen), CStr(txt_movimiento), CDbl(txt_numero), CStr(rs!VCHA_ART_ARTICULO_ID), 0, 0, "", CInt(var_consecutivo), "", CDbl(var_costo), CDbl(var_precio), CDbl(var_descuento_1), CDbl(var_descuento_2), CDbl(var_descuento_3), CDbl(var_iva), CDbl(var_factura), CStr(var_referencia), CStr(var_clave_moneda), CDbl(var_tipo_Cambio), CStr(var_serie_FACTURA), CInt(var_a?o))
                                      rsaux5.Open "update tb_Devoluciones set floa_dev_cantidad = " + CStr(var_cantidad_pasar) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen + "' and vcha_mov_movimiento_id = '" + Me.txt_movimiento + "' and inte_emo_numero = " + Me.txt_numero + " and vcha_art_articulo_id = '" + rs!VCHA_ART_ARTICULO_ID + "' and inte_cde_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                                   End If
                             Wend
                            'Next var_contador
                            rs.MoveNext
                      Wend
                      var_posible = True
                   Else
                      rs.Close
                      MsgBox "El movimiento no a sido cerrado", vbOKOnly, "ATENCION"
                      var_posible = False
                   End If
                Else
                   var_posible = True
                End If
                If rs.State = 1 Then
                   rs.Close
                End If
                If var_posible = True Then
                   
                   rs.Open "select * from tb_devoluciones where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic
                   var_almacen = rs!VCHA_ALM_ALMACEN_ID
                   If Not rs.EOF Then
                      var_contador_articulos = 0
                      While Not rs.EOF
                            var_estatus = IIf(IsNull(rs!CHAR_CDE_ESTATUS), "", rs!CHAR_CDE_ESTATUS)
                            var_nombre_articulo = ""
                            rsaux2.Open "select * from tb_articulos where vcha_Art_articulo_id ='" + rs!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                            If Not rsaux2.EOF Then
                               var_nombre_articulo = rsaux2!vcha_Art_nombre_espa?ol
                            End If
                            rsaux2.Close
                            
                            var_contador_articulos = var_contador_articulos + 1
                            rs.MoveNext
                      Wend
                      txt_numero.Enabled = False
                      txt_movimiento.Enabled = False
                   End If
                   rs.Close
                End If
             Else
                MsgBox "El Movimiento no existe o no a sido terminado aun", vbOKOnly, "ATENCION"
                rs.Close
             End If
          Else
             MsgBox "N?mero de Movimiento Incorrecto", vbOKOnly, "ATENCION"
          End If
       End If
       
    End If
    
Exit Sub
salir:
   MsgBox "No es posible asignar este movimiento ya que no se a indicado el tipo de cambio del dia de hoy", vbOKOnly, "ATENCION"
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
   Exit Sub
salir_2:
   MsgBox "El cliente no tiene una lista de precios asignada", vbOKOnly, "ATENCION"
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
   Exit Sub
salir_3:
   MsgBox "El cliente no tiene asignado todos los articulos en la lista de precios", vbOKOnly, "ATENCION"
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
   Exit Sub
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
            var_si = MsgBox("?Desea devolver los articulos de la factura?" + txt_factura, vbYesNo, "ATENCION")
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
                           var_descripcion_articulo = IIf(IsNull(rs!vcha_Art_nombre_espa?ol), "", rs!vcha_Art_nombre_espa?ol)
                        End If
                        rs.Close
                        Cadena = "select * from TB_TEMPORAL_ENTRADAS where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "'"
                        rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           lbl_total = CStr(CDbl(lbl_total) + var_cantidad_leida)
                           var_inserta = False
                           var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida, var_a?o)
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
                           var_inserta = TB_TEMPORAL_ENTRADAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida, var_costo, var_precio, "0", "", var_a?o)
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
                  si = MsgBox("?Desea cancelar el movimiento?", vbYesNo, "ATENCION")
                  If si = 6 Then
                     si = MsgBox("Confirmar la cancelaci?n del movimiento", vbYesNo, "ATENCION")
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
   Set TB_EXISTENCIAS_INSERTA = New TB_EXISTENCIAS_INSERTA
   Set TB_ENTRADAS_I = New TB_ENTRADAS_I
   Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
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
                  Set TB_DEVOLUCIONES_INSERTA = New TB_DEVOLUCIONES_INSERTA
                  var_si = MsgBox("?Se va a imprimir el movimiento?", vbOKCancel, "ATENCION")
                  If var_si = 1 Then
                     cnn.BeginTrans
                     Cadena = "select * from tb_temporal_entradas where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_emp_empresa_id = '" + var_empresa + "'"
                     rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        var_inserta = False
                        var_inserta = TB_ENTRADAS_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "", "", 0)
                     End If
                     rs.Close
                     var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, "I", Now, 1)
                     cnn.CommitTrans
                      
                     Me.txt_numero = Me.txt_folio
                     Call txt_numero_2_KeyPress
                     
                     var_si = MsgBox("?El agente especifico a donde aplicar la nota de cr?dito?", vbYesNo, "ATENCION")
                     If var_si = 6 Then
                        Call Command3_Click
                     Else
                        Call cmd_nota_credito_Click
                     End If
                     
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
               MsgBox "No se a seleccionado ning?n movimiento", vbOKOnly, "ATENCION"
            End If
End Sub

Private Sub cmd_nota_credito_Click()
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
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
   Dim var_importe_neto_1 As Double
   Dim var_importe_total_1 As Double
   Dim var_subimporte_1 As Double
   Dim var_importe_iva_1 As Double
   Dim var_tipo_Cambio As Double
   Dim var_importe_factura As Double
   Dim var_importe_pago As Double
   Dim var_importe_saldo_pago As Double
   Dim var_importe_total As Double
   Dim var_fecha_pago As Date
   Dim var_fecha_factura As Date
   Dim var_contador_pagos As Double
   Dim var_contador_facturas As Double
   Dim var_descuento_agente As Double
   Dim var_descuento_sistema As Double
   Dim var_saldo As Double
   Dim si As Integer
   Dim i, n As Integer
   Dim var_importe As Double
   Dim var_descuento As Double
   Dim var_importe_descuento As Double
   Dim var_moneda_local As Integer
   Dim var_posible_tipo_cambio As Boolean
   Dim var_serie_cargo As String
   Dim var_importe_neto As Double
   Dim var_subimporte As Double
   Dim var_importe_iva As Double
   Dim var_k As Double
   Dim var_l As Double
   Dim var_numero_nota As Double
   Dim var_contador_notas As Double
   Dim var_saldo_real As Double
   Dim var_iva_pasado As Double
   var_posible_iva = 1
   var_iva_pasado = 0
   Me.txt_movimiento = var_clave_movimiento
   Me.txt_numero = Me.txt_folio
   If Trim(Me.txt_movimiento) <> "" Then
      If IsNumeric(Me.txt_numero) Then
         rs.Open "SELECT * FROM TB_NOTAS_CREDITO_DEVOLUCIONES WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + Me.txt_movimiento + "' and inte_emo_numero = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
         If rs.EOF Then
            var_si = MsgBox("?Desea aplicar la nota de cr?dito?", vbYesNo, "ATENCION")
            If var_si = 6 Then
                var_si = MsgBox("Confirmar la aplicaci?n de la nota de cr?dito", vbYesNo, "ATENCION")
                If var_si = 6 Then
                   Set TB_ESTADO_CUENTA_INSERTA = New TB_ESTADO_CUENTA_INSERTA
                   Set TB_ENCABEZADO_CARTERA_I = New TB_ENCABEZADO_CARTERA_I
                   If rsaux.State = 1 Then
                      rsaux.Close
                   End If
                   rsaux.Open "select * from tb_encabezado_movimientos where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + Me.txt_movimiento + "' and inte_emo_numero = '" + Me.txt_numero + "'", cnn, adOpenDynamic, adLockOptimistic
                   If Not rsaux.EOF Then
                      var_cliente = rsaux!vcha_cli_clave_id
                   End If
                   rsaux.Close
                   rsaux.Open "select * from vw_clientes where vcha_cli_clave_id = '" + var_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                   var_agente = rsaux!VCHA_AGE_AGENTE_ID
                   var_grupo_actual = rsaux!VCHA_GAC_GRUPO_aCTUAL_ID
                   var_grupo_real = rsaux!vcha_gre_grupo_real_id
                   var_titular = rsaux!vcha_tit_titular_id
                   txt_clave_cliente = rsaux!vcha_cli_clave_id
                   var_clave_moneda = rsaux!vcha_mon_moneda_id
                   txt_clase = "DV"
                   If Trim(txt_clase) <> "" Then
                      If var_empresa = "02" Then
                         If var_unidad_organizacional = "23" Then
                            var_serie = "NCEFT"
                         Else
                            var_serie = "NCEMX"
                         End If
                      End If
                      If var_empresa = "03" Then
                         var_serie = "NCEVII"
                      End If
                      If var_empresa = "18" Then
                         var_serie = "NCEVXX"
                      End If
                      If var_empresa = "30" Then
                         var_serie = "NCETR"
                      End If
                      rs.Close
                      txt_serie = "DV"
                      var_plazo = 0
                      
                      
                      rs.Open "select * from vw_devolucion_nota_credito where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
                      If Not rs.EOF Then
                         txt_falta_aplicar = 0
                         var_total_neto = 0
                         var_importe_total = 0
                         var_importe_total_iva = 0
                         var_importe_total_descuento_1 = 0
                         var_importe_total_descuento_2 = 0
                         var_importe_total_descuento_3 = 0
                         var_importe_total_subimporte = 0
                         While Not rs.EOF
                               var_text_descuento = ""
                               var_cantidad = 0
                               var_precio = 0
                               var_precio_sin_descuentos = 0
                               var_subimporte = 0
                               var_imp_descuento_1 = 0
                               var_imp_descuento_2 = 0
                               var_imp_descuento_3 = 0
                               var_iva = 0
                               var_cantidad = Format(IIf(IsNull(rs!Cantidad_leida), 0, rs!Cantidad_leida), "###,###,##0.00")
                               If var_unidad_organizacional = "23" Then
                                  var_precio = IIf(IsNull(rs!Precio), 0, rs!Precio) / IIf(IsNull(rs!Cantidad_leida), 0, rs!Cantidad_leida)
                               Else
                                  'cambio realizado por carlos aleman ya que el importe unitario se estaba dividiento entre la cantidad de piezas para ALM GRAL
                                  'var_precio = IIf(IsNull(rs!floa_cde_precio), 0, rs!floa_cde_precio) / IIf(IsNull(rs!cantidad_leida), 0, rs!cantidad_leida)
                                  var_precio = IIf(IsNull(rs!floa_cde_precio), 0, rs!floa_cde_precio) / IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)
                               End If
                               'var_precio_sin_descuentos = IIf(IsNull(rs!floa_cde_precio), 0, rs!floa_cde_precio)
                               var_precio_sin_descuentos = IIf(IsNull(rs!Precio), 0, rs!Precio)
                               var_descuento_1 = IIf(IsNull(rs!floa_cde_descuento_1), 0, rs!floa_cde_descuento_1)
                               var_descuento_2 = IIf(IsNull(rs!floa_cde_descuento_2), 0, rs!floa_cde_descuento_2)
                               var_descuento_3 = IIf(IsNull(rs!floa_cde_descuento_3), 0, rs!floa_cde_descuento_3)
                               var_tipo_Cambio = IIf(IsNull(rs!floa_dev_tipo_cambio), 1, rs!floa_dev_tipo_cambio)
                               var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
                               var_precio = var_precio * (1 - (var_descuento_1 / 100))
                               var_imp_descuento_1 = ((var_precio_sin_descuentos - var_precio) * var_cantidad) / var_tipo_Cambio
                               var_precio_sin_descuentos = var_precio
                               var_precio = var_precio * (1 - (var_descuento_2 / 100))
                               var_imp_descuento_2 = ((var_precio_sin_descuentos - var_precio) * var_cantidad) / var_tipo_Cambio
                               var_precio_sin_descuentos = var_precio
                               var_precio = var_precio * (1 - (var_descuento_3 / 100))
                               var_imp_descuento_3 = ((var_precio_sin_descuentos - var_precio) * var_cantidad) / var_tipo_Cambio
                               var_precio = var_precio / var_tipo_Cambio
                               If var_empresa = "03" Or var_empresa = "28" Then
                                  var_iva = 0
                               Else
                                  '  var_iva = IIf(IsNull(rs!floa_cde_iva), 0, rs!floa_cde_iva)
                                  var_iva = 16
                               End If
                               var_subimporte = var_precio * var_cantidad
                               var_importe_total = var_importe_total + var_subimporte
                               var_importe_total_descuento_1 = var_importe_total_descuento_1 + var_imp_descuento_1
                               var_importe_total_descuento_2 = var_importe_total_descuento_2 + var_imp_descuento_2
                               var_importe_total_descuento_3 = var_importe_total_descuento_3 + var_imp_descuento_3
                               var_total = var_subimporte
                               var_imp_iva = var_total * (var_iva / 100)
                               var_importe_total_iva = var_importe_total_iva + var_imp_iva
                               var_importe_total_subimporte = var_importe_total_subimporte + var_total
                               var_total = var_total + var_imp_iva
                               var_total_neto = var_total_neto + var_total
                               lbl_moneda = rs!vcha_mon_nombre_plural
                               rs.MoveNext
                         Wend
                         var_total_neto = var_total_neto
                         txt_total_neto = Format(Round(var_total_neto, 2), "###,###,##0.00")
                         txt_falta_aplicar = Format(Round(var_total_neto, 2), "###,###,##0.00")
                         txt_importe = Format(var_total_neto, "###,###,##0.00")
                      End If
                      rs.Close
                      
                      
                      rsaux3.Open "select inte_ser_nota_credito from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                      var_numero_nota = (IIf(IsNull(rsaux3!inte_ser_nota_credito), 1, rsaux3!inte_ser_nota_credito))
                      rsaux3.Close
                      
                      var_insertar = TB_ENCABEZADO_CARTERA_I.Anadir(var_empresa, var_unidad_organizacional, "NC", CStr(txt_clase), CStr(txt_clase), CDbl(var_numero_nota), "-", "", "", 0, CStr(Date), CStr(var_agente), CStr(var_grupo_actual), CStr(var_grupo_real), CStr(var_titular), CStr(txt_clave_cliente), "", CDbl(var_plazo), CDbl(var_iva), 0, 0, 0, 0, 0, CDbl(var_importe_total), CDbl(var_importe_iva), 0, 0, 0, 0, 0, CDbl(var_subimporte), CDbl(var_total_neto), "", var_clave_usuario_global, fun_NombrePc, Date, 0, Date, Date, CStr(var_clave_moneda), CDbl(var_tipo_Cambio), CStr(var_serie), "")
                      'MsgBox "update tb_encabezado_cartera set inte_car_nota_credito_aplicada = 0, vcha_alm_almacen_id = '" + var_almacen_Destino + "', vcha_mov_movimiento_id = '" + var_clave_movimiento + "', inte_emo_numero = " + Me.txt_numero + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'DV' and vcha_Ser_Serie_id = '" + var_serie + "' and inte_car_numero = " + CStr(var_numero_nota)
                      rs.Open "update tb_encabezado_cartera set inte_car_nota_credito_aplicada = 0, vcha_alm_almacen_id = '" + var_almacen_Destino + "', vcha_mov_movimiento_id = '" + var_clave_movimiento + "', inte_emo_numero = " + Me.txt_numero + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'DV' and vcha_Ser_Serie_id = '" + var_serie + "' and inte_car_numero = " + CStr(var_numero_nota), cnn, adOpenDynamic, adLockOptimistic
                      rs.Open "update tb_series set inte_ser_nota_Credito = inte_ser_nota_credito + 1 where vcha_Ser_Serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                      rs.Open "insert into TB_NOTAS_CREDITO_DEVOLUCIONES (vcha_Emp_Empresa_id, vcha_uor_unidad_id, vcha_mov_movimiento_id, inte_emo_numero) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + Me.txt_movimiento + "', '" + Me.txt_numero + "')", cnn, adOpenDynamic, adLockOptimistic
                      rs.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
                      rsaux1.Open "SELECT * FROM TB_CLIENTES WHERE VCHA_CLI_CLAVE_ID = '" + txt_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                      var_referencia_cliente_tienda = IIf(IsNull(rsaux1!VCHA_CLI_REFERENCIA), "", rsaux1!VCHA_CLI_REFERENCIA)
                      rsaux1.Close
                      If cnn_clientes_tiendas.State = 1 Then
                         cnn_clientes_tiendas.Close
                      End If
                      cnn_clientes_tiendas.Open var_conexion_pedidos_tiendas
                      rsaux8.Open "CALL SP_AGREGA_ABONO('" + Trim(var_referencia_cliente_tienda) + "'," + CStr(CDbl(var_total_neto)) + "," + CStr(CDbl(var_total_neto)) + ",SYSDATE,SYSDATE,'" + CStr(var_numero_nota) + "','','NCT','')", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
                      
                      'var_numero_nota = 6799
                      If Not rs.EOF Then
                         var_moneda_local = IIf(IsNull(rs!inte_mon_moneda_local), 0, rs!inte_mon_moneda_local)
                      End If
                      rs.Close
                      'var_numero_nota = 6799
                      Open (App.Path & "\renombra" + Trim(var_serie) + Trim(Str(var_numero_nota)) + ".bat") For Output As #2
                      Print #2, "ren " + var_ruta_documentos_electronicos + "\" + Trim(var_serie) + Trim(Str(var_numero_nota)) + ".fi " + Trim(var_serie) + Trim(Str(var_numero_nota)) + ".ff"
                      Close #2
                        
                      Open (var_ruta_documentos_electronicos & "\" + Trim(var_serie) + Trim(Str(var_numero_nota)) + ".fi") For Output As #1
                      'Open ("c:\NC_" + Trim(var_serie) + Trim(Str(rs!inte_car_numero)) + ".fi") For Output As #1
                      var_cadena = "Outputmode=" + Chr(13) + "<Factura>" + Chr(13) + "<Comprobante>" + Chr(13) + "Version=2.0" + Chr(13) + "Serie=" + var_serie + Chr(13) + "folio=" + CStr(var_numero_nota) + Chr(13)
                      var_a?o_str = CStr(Year(Now))
                      var_mes = CStr(Month(Now))
                      var_dia = CStr(Day(Now))
                      var_hora = CStr(Hour(Now))
                      var_minuto = CStr(Minute(Now))
                      var_segundo = CStr(Second(Now))
                      If Len(var_a?o_str) = 2 Then
                         var_a?o_str = "20" + var_a?o_str
                      End If
                      If Len(var_mes) = 1 Then
                          var_mes = "0" + var_mes
                      End If
                      If Len(var_dia) = 1 Then
                         var_dia = "0" + var_dia
                      End If
                      If Len(var_hora) = 1 Then
                         var_hora = "0" + var_hora
                      End If
                      If Len(var_minuto) = 1 Then
                         var_minuto = "0" + var_minuto
                      End If
                      If Len(var_segundo) = 1 Then
                         var_segundo = "0" + var_segundo
                      End If
                      rs.Open "SELECT * FROM VW_CLIENTES WHERE VCHA_cLI_CLAVE_ID = '" + txt_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                      var_rfc_cliente_1 = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                      var_rfc_cliente = ""
                      If var_rfc_cliente_1 = "" Then
                         var_rfc_cliente = "XAXX010101000"
                      Else
                         For var_j = 1 To Len(var_rfc_cliente_1)
                             If Mid(var_rfc_cliente_1, var_j, 1) <> "-" Then
                                If Mid(var_rfc_cliente_1, var_j, 1) <> "" Then
                                   If Mid(var_rfc_cliente_1, var_j, 1) <> " " Then
                                      var_rfc_cliente = var_rfc_cliente + Mid(var_rfc_cliente_1, var_j, 1)
                                   End If
                                End If
                             End If
                         Next var_j
                      End If
                           
                           
                      var_cadena_fecha = CStr(var_a?o_str) + "-" + CStr(var_mes) + "-" + CStr(var_dia) + "T" + CStr(var_hora) + ":" + CStr(var_minuto) + ":" + CStr(var_segundo)
                      var_cadena = var_cadena + "fecha=" + var_cadena_fecha + Chr(13)
                      var_cadena = var_cadena + "noAprobacion=" + Chr(13)
                      var_cadena = var_cadena + "anoAprobacion=" + Chr(13)
                      var_cadena = var_cadena + "tipoDeComprobante=NOTA DE CREDITO" + Chr(13)
                      var_cadena = var_cadena + "formaDePago=PAGO HECHO EN UNA SOLA EXHIBICION" + Chr(13)
                      var_cadena = var_cadena + "condicionesDePago=" + Chr(13)
                      If var_rfc_cliente = "XAXX010101000" Then
                         var_cadena = var_cadena + "subtotal=" + Format(CStr(var_importe_total * 1.16), "###,###,##0.000000") + Chr(13)
                      Else
                         var_cadena = var_cadena + "subtotal=" + Format(CStr(var_importe_total_subimporte / 1), "###,###,##0.000000") + Chr(13)
                      End If
                      var_cadena = var_cadena + "descuento=" + Chr(13)
                      var_cadena = var_cadena + "descuento1=" + Chr(13)
                      var_cadena = var_cadena + "descuento2=" + Chr(13)
                      var_cadena = var_cadena + "conceptodescuento1=" + Chr(13)
                      var_cadena = var_cadena + "conceptodescuento2=" + Chr(13)
                      var_cadena = var_cadena + "tasadescuento1=" + Chr(13)
                      var_cadena = var_cadena + "tasadescuento2=" + Chr(13)
                      If rsaux2.State = 1 Then
                         rsaux2.Close
                      End If
                      rsaux2.Open "select * from tb_empresa_FACTURA_ELECTRONICA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                      var_certificado = rsaux2!vcha_emp_certificado
                      var_expedido = rsaux2!vcha_emp_expedido
                      If var_rfc_cliente = "XAXX010101000" Then
                         var_cadena = var_cadena + "iva=" + Format(CStr(0), "###,###,##0.000000") + Chr(13)
                      Else
                         var_cadena = var_cadena + "iva=" + Format(CStr(var_importe_total_iva / 1), "###,###,##0.000000") + Chr(13)
                      End If
                      var_cadena = var_cadena + "total=" + Format(CStr(var_total_neto / 1), "###,###,##0.000000") + Chr(13)
                      var_cadena = var_cadena + "retencion=" + Chr(13)
                      var_cadena = var_cadena + "factorretencioniva=" + Chr(13)
                      var_cadena = var_cadena + "</Comprobante>" + Chr(13) + Chr(13)
                      var_cadena = var_cadena + "<Emisor>" + Chr(13)
                      var_cadena = var_cadena + "erfc=" + rsaux2!VCHA_eMP_RFC + Chr(13)
                      var_cadena = var_cadena + "enombre=" + rsaux2!VCHA_EMP_NOMBRE + Chr(13)
                      var_cadena = var_cadena + "</Emisor>" + Chr(13) + Chr(13)
                      var_cadena = var_cadena + "<DomicilioFiscal>" + Chr(13)
                      var_cadena = var_cadena + "ecalle=" + rsaux2!VCHA_eMP_CALLE + Chr(13)
                      var_cadena = var_cadena + "enoExterior=" + rsaux2!VCHA_eMP_exterior + Chr(13)
                      var_cadena = var_cadena + "enoInterior=" + Chr(13)
                      var_cadena = var_cadena + "ecolonia=" + rsaux2!VCHA_eMP_COLONIA + Chr(13)
                      var_cadena = var_cadena + "elocalidad=" + rsaux2!VCHA_EMP_LOCALIDAD + Chr(13)
                      var_cadena = var_cadena + "ereferencia=" + Chr(13)
                      var_cadena = var_cadena + "emunicipio=" + rsaux2!VCHA_EMP_MUNICIPIO + Chr(13)
                      var_cadena = var_cadena + "eestado=" + rsaux2!VCHA_EMP_ESTADO + Chr(13)
                      var_cadena = var_cadena + "epais=" + rsaux2!VCHA_eMP_PAIS + Chr(13)
                      var_cadena = var_cadena + "ecodigoPostal=" + rsaux2!VCHA_EMP_CODIGO_POSTAL + Chr(13)
                      var_cadena = var_cadena + "etel=" + IIf(IsNull(rsaux2!VCHA_EMP_TELEFONO), "", rsaux2!VCHA_EMP_TELEFONO) + Chr(13)
                      var_cadena = var_cadena + "eemail=" + IIf(IsNull(rsaux2!VCHA_EMP_EMAIL), "", rsaux2!VCHA_EMP_EMAIL) + Chr(13)
                      var_cadena = var_cadena + "</DomicilioFiscal>" + Chr(13) + Chr(13)
                      
                      var_cadena = var_cadena + "<ExpedidoEn>" + Chr(13) + Chr(13)
                      var_cadena = var_cadena + "ex_calle=" + rsaux2!VCHA_eMP_CALLE + Chr(13)
                      var_cadena = var_cadena + "ex_noExterior=" + rsaux2!VCHA_eMP_exterior + Chr(13)
                      var_cadena = var_cadena + "ex_noInterior=" + Chr(13)
                      var_cadena = var_cadena + "ex_colonia=" + rsaux2!VCHA_eMP_COLONIA + Chr(13)
                      var_cadena = var_cadena + "ex_localidad=" + rsaux2!VCHA_EMP_LOCALIDAD + Chr(13)
                      var_cadena = var_cadena + "ex_referencia=" + Chr(13)
                      var_cadena = var_cadena + "ex_municipio=" + rsaux2!VCHA_EMP_MUNICIPIO + Chr(13)
                      var_cadena = var_cadena + "ex_estado=" + rsaux2!VCHA_EMP_ESTADO + Chr(13)
                      var_cadena = var_cadena + "ex_pais=" + rsaux2!VCHA_eMP_PAIS + Chr(13)
                      var_cadena = var_cadena + "ex_codigoPostal=" + rsaux2!VCHA_EMP_CODIGO_POSTAL + Chr(13)
                      var_cadena = var_cadena + "</ExpedidoEn>"
                      
                      
                      
                      
                      var_cadena = var_cadena + "<Receptor>" + Chr(13)
                      var_cadena = var_cadena + "noCliente=" + rs!vcha_cli_clave_id + Chr(13)
                      rsaux2.Close
                                         
                      var_rfc_cliente_1 = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                      var_rfc_cliente = ""
                      If var_rfc_cliente_1 = "" Then
                         var_rfc_cliente = "XAXX010101000"
                      Else
                         For var_j = 1 To Len(var_rfc_cliente_1)
                             If Mid(var_rfc_cliente_1, var_j, 1) <> "-" Then
                                If Mid(var_rfc_cliente_1, var_j, 1) <> "" Then
                                   If Mid(var_rfc_cliente_1, var_j, 1) <> " " Then
                                      var_rfc_cliente = var_rfc_cliente + Mid(var_rfc_cliente_1, var_j, 1)
                                   End If
                                End If
                             End If
                         Next var_j
                      End If
                      If var_empresa = "03" Or var_empresa = "28" Then
                         var_rfc_cliente = "XEXX010101000"
                      End If
                      var_cadena = var_cadena + "rfc=" + var_rfc_cliente + Chr(13)
                      var_cadena = var_cadena + "nombre=" + rs!VCHA_CLI_NOMBRE + Chr(13)
                      var_cadena = var_cadena + "</Receptor>" + Chr(13) + Chr(13)
                      var_cadena = var_cadena + "<Cliente>" + Chr(13)
                      var_cadena = var_cadena + "domicilio=" + IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + Chr(13)
                      var_cadena = var_cadena + "calle=" + Chr(13)
                      var_cadena = var_cadena + "noExterior=" + Chr(13)
                      var_cadena = var_cadena + "noInterior=" + Chr(13)
                      var_cadena = var_cadena + "colonia=" + IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre) + Chr(13)
                      var_cadena = var_cadena + "localidad=" + IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre) + Chr(13)
                      rsaux2.Open "select * from vw_clientes where vcha_Cli_clave_id = '" + rs!vcha_cli_clave_id + "'"
                      var_cadena = var_cadena + "referencia=" + Chr(13)
                      var_cadena = var_cadena + "municipio=" + IIf(IsNull(rsaux2!vcha_mun_nombre), "", rsaux2!vcha_mun_nombre) + Chr(13)
                      var_cadena = var_cadena + "estado=" + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre) + Chr(13)
                      VAR_NOMBRE_PAIS = IIf(IsNull(rs!vcha_pai_nombre), "MEXICO", rs!vcha_pai_nombre)
                      If Trim(VAR_NOMBRE_PAIS) = "" Then
                         VAR_NOMBRE_PAIS = "MEXICO"
                      End If
                      var_cadena = var_cadena + "pais=" + VAR_NOMBRE_PAIS + Chr(13)
                      var_cadena = var_cadena + Chr(13)
                      var_cadena = var_cadena + "codigoPostal=" + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP) + Chr(13)
                      var_cadena = var_cadena + "tel=" + Chr(13)
                      var_cadena = var_cadena + "email=" + IIf(IsNull(rs!vcha_cli_email), "", rs!vcha_cli_email) + Chr(13)
                      var_cadena = var_cadena + "</Cliente>" + Chr(13) + Chr(13)
                                    
                      var_cadena = var_cadena + "<EntregarEn>" + Chr(13)
                      var_cadena = var_cadena + "endomicilio=" + Chr(13)
                      var_cadena = var_cadena + "encalle=" + Chr(13)
                      var_cadena = var_cadena + "ennoExterior=" + Chr(13)
                      var_cadena = var_cadena + "ennoInterior=" + Chr(13)
                      var_cadena = var_cadena + "encolonia=" + Chr(13)
                      var_cadena = var_cadena + "enlocalidad=" + Chr(13)
                      var_cadena = var_cadena + "enreferencia=" + Chr(13)
                      var_cadena = var_cadena + "enmunicipio=" + Chr(13)
                      var_cadena = var_cadena + "enestado=" + Chr(13)
                      var_cadena = var_cadena + "enpais=" + Chr(13)
                      var_cadena = var_cadena + "encodigoPostal=" + Chr(13)
                      var_cadena = var_cadena + "entel=" + Chr(13)
                      var_cadena = var_cadena + "enemail=" + Chr(13)
                      var_cadena = var_cadena + "</EntregarEn>" + Chr(13) + Chr(13)
                      var_cadena = var_cadena + "<Concepto>" + Chr(13)
                       
                      pxx = CStr(var_i)
                      var_i = var_i + 1
                      
                      
                      
                      
                      rsaux11.Open "select * from vw_devolucion_nota_credito where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
                      If Not rsaux11.EOF Then
                         txt_falta_aplicar = 0
                         var_total_neto = 0
                         var_importe_total = 0
                         var_importe_total_iva = 0
                         var_importe_total_descuento_1 = 0
                         var_importe_total_descuento_2 = 0
                         var_importe_total_descuento_3 = 0
                         var_importe_total_subimporte = 0
                         var_i = 0
                         var_piezas_totales = 0
                         var_referencia_agente = "NA " + IIf(IsNull(rsaux11!vcha_emo_referencia), "", rsaux11!vcha_emo_referencia)
                         While Not rsaux11.EOF
                               var_text_descuento = ""
                               var_cantidad = 0
                               var_precio = 0
                               var_precio_sin_descuentos = 0
                               var_subimporte = 0
                               var_imp_descuento_1 = 0
                               var_imp_descuento_2 = 0
                               var_imp_descuento_3 = 0
                               var_iva = 0
                               var_cantidad = Format(IIf(IsNull(rsaux11!Cantidad_leida), 0, rsaux11!Cantidad_leida), "###,###,##0.00")
                               var_piezas_totales = var_piezas_totales + var_cantidad
                               If var_unidad_organizacional = "23" Then
                                  var_precio = IIf(IsNull(rsaux11!Precio), 0, rsaux11!Precio) / IIf(IsNull(rsaux11!Cantidad_leida), 0, rsaux11!Cantidad_leida)
                               Else
                                  'cambio realizado por carlos aleman ya que el importe unitario se estaba dividiento entre la cantidad de piezas para ALM GRAL
                                  'var_precio = IIf(IsNull(rs!floa_cde_precio), 0, rs!floa_cde_precio) / IIf(IsNull(rs!cantidad_leida), 0, rs!cantidad_leida)
                                  var_precio = IIf(IsNull(rsaux11!floa_cde_precio), 0, rsaux11!floa_cde_precio) / IIf(IsNull(rsaux11!Cantidad), 0, rsaux11!Cantidad)
                               End If
                               'var_precio_sin_descuentos = IIf(IsNull(rs!floa_cde_precio), 0, rs!floa_cde_precio)
                               var_precio_sin_descuentos = IIf(IsNull(rsaux11!Precio), 0, rsaux11!Precio)
                               var_descuento_1 = IIf(IsNull(rsaux11!floa_cde_descuento_1), 0, rsaux11!floa_cde_descuento_1)
                               var_descuento_2 = IIf(IsNull(rsaux11!floa_cde_descuento_2), 0, rsaux11!floa_cde_descuento_2)
                               var_descuento_3 = IIf(IsNull(rsaux11!floa_cde_descuento_3), 0, rsaux11!floa_cde_descuento_3)
                               var_tipo_Cambio = IIf(IsNull(rsaux11!floa_dev_tipo_cambio), 1, rsaux11!floa_dev_tipo_cambio)
                               var_clave_moneda = IIf(IsNull(rsaux11!vcha_mon_moneda_id), "", rsaux11!vcha_mon_moneda_id)
                               var_precio = var_precio * (1 - (var_descuento_1 / 100))
                               var_imp_descuento_1 = ((var_precio_sin_descuentos - var_precio) * var_cantidad) / var_tipo_Cambio
                               var_precio_sin_descuentos = var_precio
                               var_precio = var_precio * (1 - (var_descuento_2 / 100))
                               var_imp_descuento_2 = ((var_precio_sin_descuentos - var_precio) * var_cantidad) / var_tipo_Cambio
                               var_precio_sin_descuentos = var_precio
                               var_precio = var_precio * (1 - (var_descuento_3 / 100))
                               var_imp_descuento_3 = ((var_precio_sin_descuentos - var_precio) * var_cantidad) / var_tipo_Cambio
                               var_precio = var_precio / var_tipo_Cambio
                               If var_empresa = "03" Or var_empresa = "28" Then
                                  var_iva = 0
                               Else
                                  '  var_iva = IIf(IsNull(rs!floa_cde_iva), 0, rs!floa_cde_iva)
                                  var_iva = 16
                               End If
                               var_subimporte = var_precio * var_cantidad
                               var_importe_total = var_importe_total + var_subimporte
                               var_importe_total_descuento_1 = var_importe_total_descuento_1 + var_imp_descuento_1
                               var_importe_total_descuento_2 = var_importe_total_descuento_2 + var_imp_descuento_2
                               var_importe_total_descuento_3 = var_importe_total_descuento_3 + var_imp_descuento_3
                               var_total = var_subimporte
                               var_imp_iva = var_total * (var_iva / 100)
                               var_importe_total_iva = var_importe_total_iva + var_imp_iva
                               var_importe_total_subimporte = var_importe_total_subimporte + var_total
                               var_total = var_total + var_imp_iva
                               var_total_neto = var_total_neto + var_total
                               lbl_moneda = rs!vcha_mon_nombre_plural
                               var_i = var_i + 1
                               pxx = CStr(var_i)
                               If Len(pxx) = 1 Then
                                  pxx = "0" + pxx
                               End If

                               var_cadena = var_cadena + "p" + pxx + "_cantidad=" + CStr(IIf(IsNull(rsaux11!Cantidad_leida), 0, rsaux11!Cantidad_leida)) + Chr(13)
                               var_cadena = var_cadena + "p" + pxx + "_unidad=" + "PIEZA" + Chr(13)
                               var_cadena = var_cadena + "p" + pxx + "_noIdentificacion=" + rsaux11!VCHA_ART_ARTICULO_ID + Chr(13)
                               rsaux9.Open "SELECT * FROM TB_aRTICULOS WHERE VCHA_aRT_aRTICULO_ID = '" + rsaux11!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                               If Not rsaux.EOF Then
                                  var_nombre_articulo = IIf(IsNull(rsaux9!vcha_Art_nombre_espa?ol), "", rsaux9!vcha_Art_nombre_espa?ol)
                               End If
                               rsaux9.Close
                               var_factura = IIf(IsNull(rsaux11!vcha_Ser_Serie_id), "", rsaux11!vcha_Ser_Serie_id) + CStr(IIf(IsNull(rsaux11!inte_fac_factura), "", rsaux11!inte_fac_factura))
                               var_linea = var_factura + " " + rsaux11!VCHA_ART_ARTICULO_ID + " " + var_nombre_articulo
                               var_cadena = var_cadena + "p" + pxx + "_descripcion=" + var_linea + Chr(13)
                               If var_rfc_cliente = "XAXX010101000" Then
                                  var_total = var_total
                               Else
                                  var_total = var_total / 1.16
                               End If
                               var_cadena = var_cadena + "p" + pxx + "_valorUnitario=" + Format(CStr(var_total / IIf(IsNull(rsaux11!Cantidad_leida), 1, rsaux11!Cantidad_leida)), "###,###,##0.000000") + Chr(13)
                               var_cadena = var_cadena + "p" + pxx + "_importe=" + Format(CStr(var_total), "###,###,##0.000000") + Chr(13)
                               
                               
                               rsaux11.MoveNext
                         Wend
                         var_total_neto = var_total_neto
                         txt_total_neto = Format(Round(var_total_neto, 2), "###,###,##0.00")
                         txt_falta_aplicar = Format(Round(var_total_neto, 2), "###,###,##0.00")
                         txt_importe = Format(var_total_neto, "###,###,##0.00")
                      End If
                      rsaux11.Close
                      
                      
                      
                      
                      var_cadena = var_cadena + "</Concepto>" + Chr(13) + Chr(13)
                      var_cadena = var_cadena + "<Otros>" + Chr(13)
                      var_cadena = var_cadena + "certificado=" + IIf(IsNull(var_certificado), "", var_certificado) + Chr(13)
                      rs.MoveFirst
                      Call numero_letras(var_total_neto, "1")
                      var_cadena = var_cadena + "cant_letra=" + canstr + Chr(13)
                      var_cadena = var_cadena + "factoriva=" + CStr(16) + "%" + Chr(13)
                      rsaux1.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + IIf(IsNull(var_clave_moneda), "", var_clave_moneda) + "'", cnn, adOpenDynamic, adLockOptimistic
                      var_cadena = var_cadena + "moneda=" + IIf(IsNull(rsaux1!vcha_mon_nombre_plural), "", rsaux1!vcha_mon_nombre_plural) + Chr(13)
                      rsaux1.Close
                      var_cadena = var_cadena + "tipodeCambio=" + CStr(1) + Chr(13)
                      var_cadena = var_cadena + "pedido=" + var_referencia_agente + Chr(13)
                      var_cadena = var_cadena + "Embarque=CA " + Me.txt_folio + Chr(13)
                      var_referencia_Bancaria = ""
                      var_cadena = var_cadena + "referenciabancaria=" + Chr(13)
                      var_cadena = var_cadena + "fechaPedido=" + Chr(13)
                      var_cadena = var_cadena + "expedicion=" + Chr(13)
                      var_cadena = var_cadena + "observaciones=" + Chr(13)
                      var_cadena = var_cadena + "conceptoExtra1=" + Chr(13)
                      var_cadena = var_cadena + "montoconceptoExtra1=" + Chr(13)
                      var_cadena = var_cadena + "conceptoExtra2=" + Chr(13)
                      var_cadena = var_cadena + "montoconceptoExtra2=" + Chr(13)
                      var_cadena = var_cadena + "tipoimpresion=2" + Chr(13)
                        
                      var_cadena = var_cadena + "agente=" + rs!VCHA_AGE_AGENTE_ID + " " + rs!VCHA_AGE_NOMBRE + Chr(13)
                         
                      If var_empresa = "02" Or var_empresa = "15" Or var_empresa = "16" Or var_empresa = "03" Or var_empresa = "18" Or var_empresa = "17" Or var_empresa = "06" Then
                         var_cadena = var_cadena + "formato=MHNCVTH_V01.dat" + Chr(13)
                      End If
                      If var_empresa = "07" Then
                         var_cadena = var_cadena + "formato=MHNCARE_V01.dat" + Chr(13)
                      End If
                      If var_empresa = "31" Then
                         var_cadena = var_cadena + "formato=MHNCCAN_V01.dat" + Chr(13)
                      End If
                      If var_empresa = "42" Then
                         var_cadena = var_cadena + "formato=MHNCCMA_V01.dat" + Chr(13)
                      End If
                      If var_empresa = "41" Then
                         var_cadena = var_cadena + "formato=MHNCCOP_V01.dat" + Chr(13)
                      End If
                      If var_empresa = "150000" Then
                         var_cadena = var_cadena + "formato=MHNCERE_V01.dat" + Chr(13)
                      End If
                      If var_empresa = "33" Then
                         var_cadena = var_cadena + "formato=MHNCMPU_V01.dat" + Chr(13)
                      End If
                      If var_empresa = "34" Then
                         var_cadena = var_cadena + "formato=MHNCMYG_V01.dat" + Chr(13)
                      End If
                      If var_empresa = "160000" Then
                         var_cadena = var_cadena + "formato=MHNCMYG_V01.dat" + Chr(13)
                      End If
                      If var_empresa = "36" Then
                         var_cadena = var_cadena + "formato=MHNCSME_V01.dat" + Chr(13)
                      End If
                      If var_empresa = "30" Then
                         var_cadena = var_cadena + "formato=MHNCTUR_V01.dat" + Chr(13)
                      End If
                      If var_empresa = "44" Then
                         var_cadena = var_cadena + "formato=MHNCUTV_V01.dat" + Chr(13)
                      End If
                      If var_empresa = "38" Then
                         var_cadena = var_cadena + "formato=MHNCVIA_V01.dat" + Chr(13)
                      End If
                      If var_empresa = "40" Then
                         var_cadena = var_cadena + "formato=MHNCVIN_V01.dat" + Chr(13)
                      End If
                      If var_empresa = "43" Then
                          var_cadena = var_cadena + "formato=MHNCVOP_V01.dat" + Chr(13)
                      End If
                       
                        
                        
                        
                 
                      var_cadena = var_cadena + "</Otros>" + Chr(13) + Chr(13)
                      var_cadena = var_cadena + "piezas_totales=" + CStr(var_piezas_totales) + Chr(13)
                      var_cadena = var_cadena + "<addenda>" + Chr(13)
                      var_cadena = var_cadena + "</addenda>" + Chr(13) + Chr(13)
                      var_cadena = var_cadena + "</Factura>"
                      Print #1, var_cadena
                      Close #1
                       
                      var_Archivo = App.Path & "\renombra" + Trim(var_serie) + Trim(Str(var_numero_nota)) + ".bat"
                      x = Shell(var_Archivo, vbHide)
                      MsgBox "Se a terminado el proceso, nota de cr?dito " + CStr(var_numero_nota), vbOKOnly, "ATENCION"
'''''' fin de impresion de notas de credito externas
                   End If
                End If
            End If
         Else
            MsgBox "La nota de cr?dito ya fue impresa", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         MsgBox "N?mero de movimiento incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Movimiento incorrecto", vbOKOnly, "ATENCION"
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

Private Sub Command3_Click()
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
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
   Me.txt_movimiento = var_clave_movimiento
   Me.txt_numero = Me.txt_folio
   If Me.txt_movimiento <> "" Then
      If IsNumeric(Me.txt_numero) Then
         rs.Open "SELECT * FROM TB_NOTAS_CREDITO_DEVOLUCIONES WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + Me.txt_movimiento + "' and inte_emo_numero = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
         If rs.EOF Then
            var_clave_movimiento_nc = Me.txt_movimiento
            var_numero_nc = Me.txt_numero
            var_aplicar_nota_credito = 1
            rs.Close
            frmnotas_credito.Show
          Else
             rs.Close
             MsgBox "La nota de cr?dito ya fue impresa", vbOKOnly, "ATENCION"
         End If
          
      Else
         MsgBox "N?mero de devoluci?n incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Movimiento incorrecto", vbOKOnly, "ATENCION"
   End If
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
   Me.chk_factura = 0
   Me.chk_factura.Visible = False
   Me.txt_movimiento = var_clave_movimiento
   Me.txt_movimiento.Visible = False
   Me.frm_pasar_todo.Visible = False
   If var_clave_usuario_global = "11" Or var_clave_usuario_global = "8" Then
      Me.cmd_pasar_todo.Visible = True
   Else
      Me.cmd_pasar_todo.Visible = False
   End If
   lbl_total = "0"
   lbl_cancelado = ""
   var_a?o = 2005
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
   rs.Open "select * from tb_almacenes where inte_alm_costeo = 1", cnn, adOpenDynamic, adLockOptimistic
   var_clave_almacen_costo = IIf(IsNull(rs!VCHA_ALM_ALMACEN_ID), "", rs!VCHA_ALM_ALMACEN_ID)
   rs.Close
   rs.Open "select * from tb_monedas where inte_mon_moneda_local = 1", cnn, adOpenDynamic, adLockOptimistic
   var_clave_moneda = ""
   If Not rs.EOF Then
      var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
   End If
   rs.Close
   var_ventana = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
   If var_solo_lectura = False Then
   End If
   Call activa_forma(var_activa_forma_entradas_devoluciones)
   If var_numero_folio > 0 Then
     rs.Open "UPDATE TB_encabezado_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_ALM_ALMACEN_ID = '" + var_almacen_Destino + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_EMO_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
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
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la informaci?n disponible"
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
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la informaci?n disponible"
End Sub

Private Sub txt_almacen_KeyDown(KeyCode As Integer, Shift As Integer)
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
         rs.Open "select * from vw_movimientos_almacenes where vcha_mov_movimiento_id = '" + var_clave_movimiento + "' AND VCHA_ALM_ALMACEN_ID = '" + txt_almacen + "'", cnn, adOpenDynamic, adLockBatchOptimistic
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
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
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
                  txt_referencia = rs!vcha_emo_referencia
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
            MsgBox "El n?mero de movimiento no existe ", vbOKOnly, "ATENCION"
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
            'var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, lv_entradas.selectedItem, 0 - Val(txt_cantidad_eliminar), 2005)
            rs.Open "update tb_Temporal_entradas set floa_ent_cantidad = floa_ent_Cantidad - " + CStr(CDbl(Me.txt_cantidad_eliminar)) + " where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_ent_numero = " + Me.txt_folio + " and vcha_art_Articulo_id = '" + Me.lv_entradas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
            lv_entradas.selectedItem.SubItems(2) = lv_entradas.selectedItem.SubItems(2) - Val(txt_cantidad_eliminar)
            lbl_total = CStr(CDbl(lbl_total) - Val(txt_cantidad_eliminar))
            var_renglon = lv_entradas.selectedItem.Index
            Call ilumina_grid
         Else
            MsgBox "La cantidad a eliminar supera a la cantidad asignada a la causa de devoluci?n seleccionada", vbOKOnly, "ATENCION"
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
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la informaci?n disponible"
End Sub

Private Sub txt_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      If Me.txt_agente = "00100" Then
         rs.Open "select distinct vcha_cli_clave_id, vcha_cli_nombre from vw_clientes where vcha_age_agente_id = '" + Me.txt_agente + "' and vcha_emp_empresa_id = '" + var_empresa + "' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
      Else
         rs.Open "select distinct vcha_cli_clave_id, vcha_cli_nombre from vw_establecimientos where vcha_age_agente_id = '" + Me.txt_agente + "' and vcha_emp_empresa_id = '" + var_empresa + "' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
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
         rs.Open "select * from vw_establecimientos where vcha_cli_clave_id = '" + txt_cliente + "' and vcha_age_agente_id = '" + txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
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
                      frmmensaje.lbl_mensaje = "El art?culo no existe"
                      frmmensaje.Show
                      'MsgBox "El art?culo no existe", vbOKOnly, "ATENCION"
                  End If
               Else
                   txt_codigo = ""
                   frmmensaje.lbl_mensaje = "El art?culo no existe"
                   frmmensaje.Show
                  'MsgBox "El art?culo no existe", vbOKOnly, "ATENCION"
                  rs.Close
               End If
            End If
         Else
         End If
      Else
         txt_codigo = ""
         frmmensaje.lbl_mensaje = "Error en c?digo"
         frmmensaje.Show
         MsgBox "Error en c?digo", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_establecimiento_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la informaci?n disponible"
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
         MsgBox "Clave de Establecimiento Incorrecta", vbOKOnly, "ATENCION"
         txt_establecimiento = ""
         txt_nombre_establecimiento = ""
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
   Dim var_posible_cliente As Boolean
   
   If Trim(txt_codigo.Text) <> "" Then
      If var_empresa = "18" Then
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
         
         Cadena = "select * from TB_TEMPORAL_ENTRADAS where vcha_alm_almacen_id = '" + var_almacen_Destino + "' and  VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' and inte_ent_numero = " + Str(var_numero_folio) + " and vcha_art_articulo_id = '" + txt_codigo + "' and vcha_emp_empresa_id = '" + var_empresa + "'"
         rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            lbl_total = CStr(CDbl(lbl_total) + var_cantidad_leida)
            var_inserta = False
            var_inserta = TB_TEMPORAL_ENTRADAS_MODIFICA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida, var_a?o)
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
            var_inserta = TB_TEMPORAL_ENTRADAS_INSERTA.Anadir(var_empresa, var_unidad_organizacional, var_almacen_Destino, var_clave_movimiento, var_numero_folio, txt_codigo, var_cantidad_leida, var_costo, var_precio, "0", "", var_a?o)
            rs.Close
            Set list_item = lv_entradas.ListItems.Add(, , Trim(txt_codigo))
            list_item.SubItems(1) = var_descripcion_articulo
            list_item.SubItems(2) = var_cantidad_leida
            var_renglon = lv_entradas.ListItems.Count
            Call ilumina_grid
         End If
      Else
         rsaux11.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux11.EOF Then
            var_DEscripcion = IIf(IsNull(rsaux11!vcha_Art_nombre_espa?ol), "", rsaux11!vcha_Art_nombre_espa?ol)
         Else
            var_DEscripcion = ""
         End If
         rsaux11.Close
         txt_codigo = ""
         frmmensaje.lbl_mensaje = "El art?culo " + var_DEscripcion + " no se le a vendido al cliente"
         frmmensaje.Show 1
      End If
      txt_codigo.SetFocus
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
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la informaci?n disponible"
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
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la informaci?n disponible"
End Sub

Private Sub txt_nombre_almacen_KeyDown(KeyCode As Integer, Shift As Integer)
   KeyCode = 0
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
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la informaci?n disponible"
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
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la informaci?n disponible"
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

Private Sub txt_nombre_establecimiento_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
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
