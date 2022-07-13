VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmremision_vistas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturacón de mercancía a vistas"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2430
      Picture         =   "frmremision_vistas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   66
      ToolTipText     =   "Cerrar remisión"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_aplicacion_anticipos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2100
      Picture         =   "frmremision_vistas.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Aplicación de anticipo"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_remisiones_sin_facturar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1770
      Picture         =   "frmremision_vistas.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Remisiones sin facturar"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame frm_resumen 
      Height          =   3510
      Left            =   3240
      TabIndex        =   61
      Top             =   2565
      Width           =   5445
      Begin VB.TextBox txt_cantidad_total_linea 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   330
         Left            =   3765
         TabIndex        =   64
         Top             =   3060
         Width           =   1605
      End
      Begin MSComctlLib.ListView lv_resumen 
         Height          =   2910
         Left            =   30
         TabIndex        =   62
         Top             =   120
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   5133
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
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   2293
         EndProperty
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   3270
         TabIndex        =   63
         Top             =   3135
         Width           =   405
      End
   End
   Begin VB.CommandButton cmd_resumen 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      Picture         =   "frmremision_vistas.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Resumen de Remisión"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame frm_numero_nota 
      Height          =   840
      Left            =   2955
      TabIndex        =   53
      Top             =   15
      Width           =   2910
      Begin VB.TextBox txt_numero_nota 
         Height          =   330
         Left            =   75
         TabIndex        =   54
         Top             =   375
         Width           =   2745
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Número de Nota"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   55
         Top             =   15
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmd_buscar_remision 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   450
      Picture         =   "frmremision_vistas.frx":0408
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Buscar Alt + B"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame frm_factura_cancelar 
      Height          =   2250
      Left            =   765
      TabIndex        =   36
      Top             =   495
      Width           =   5820
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   60
         Picture         =   "frmremision_vistas.frx":050A
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Cancelar Factura"
         Top             =   375
         Width           =   330
      End
      Begin VB.Frame Frame3 
         Height          =   30
         Left            =   0
         TabIndex        =   50
         Top             =   705
         Width           =   5805
      End
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2850
         TabIndex        =   44
         Top             =   1830
         Width           =   1440
      End
      Begin VB.TextBox txt_fecha 
         Height          =   315
         Left            =   915
         TabIndex        =   43
         Top             =   1830
         Width           =   1245
      End
      Begin VB.TextBox txt_nombre_cliente_cancelar 
         Height          =   315
         Left            =   915
         TabIndex        =   42
         Top             =   1485
         Width           =   4620
      End
      Begin VB.TextBox txt_nombre_agente_cancelar 
         Height          =   315
         Left            =   915
         TabIndex        =   41
         Top             =   1140
         Width           =   4620
      End
      Begin VB.TextBox txt_numero 
         Height          =   315
         Left            =   3390
         TabIndex        =   40
         Top             =   795
         Width           =   1575
      End
      Begin VB.TextBox txt_serie 
         Height          =   315
         Left            =   915
         TabIndex        =   39
         Top             =   795
         Width           =   825
      End
      Begin VB.Label lbl_moneda 
         Height          =   195
         Left            =   4335
         TabIndex        =   52
         Top             =   1905
         Width           =   1320
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
         Height          =   195
         Left            =   2220
         TabIndex        =   49
         Top             =   1890
         Width           =   570
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   1890
         Width           =   495
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   47
         Top             =   1545
         Width           =   525
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   1200
         Width           =   555
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   2655
         TabIndex        =   45
         Top             =   855
         Width           =   600
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   855
         Width           =   405
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         Caption         =   " Factura a cancelar"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   15
         TabIndex        =   37
         Top             =   120
         Width           =   5760
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1110
      Picture         =   "frmremision_vistas.frx":0654
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cancelar Factura"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame frm_lista 
      Height          =   2385
      Left            =   1530
      TabIndex        =   27
      Top             =   885
      Width           =   5670
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1875
         Left            =   30
         TabIndex        =   28
         Top             =   435
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   3307
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
         Left            =   45
         TabIndex        =   29
         Top             =   135
         Width           =   5580
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11070
      Picture         =   "frmremision_vistas.frx":079E
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   780
      Picture         =   "frmremision_vistas.frx":0DD8
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmremision_vistas.frx":0EDA
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   45
      TabIndex        =   19
      Top             =   345
      Width           =   11445
   End
   Begin VB.Frame Frame1 
      Height          =   6810
      Left            =   60
      TabIndex        =   0
      Top             =   390
      Width           =   11415
      Begin VB.TextBox txt_titular 
         Height          =   345
         Left            =   1200
         TabIndex        =   11
         Top             =   555
         Width           =   1125
      End
      Begin VB.TextBox txt_nombre_titular 
         Height          =   345
         Left            =   2370
         TabIndex        =   12
         Top             =   555
         Width           =   4410
      End
      Begin VB.TextBox txt_remision_agente 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1200
         TabIndex        =   15
         Top             =   1365
         Width           =   1125
      End
      Begin VB.TextBox txt_importe_neto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   9150
         TabIndex        =   60
         Top             =   6285
         Width           =   2130
      End
      Begin VB.TextBox txt_nombre_agente 
         Height          =   345
         Left            =   2370
         TabIndex        =   10
         Top             =   165
         Width           =   4410
      End
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   345
         Left            =   2370
         TabIndex        =   14
         Top             =   960
         Width           =   4410
      End
      Begin VB.TextBox txt_remision 
         Enabled         =   0   'False
         Height          =   345
         Left            =   10005
         TabIndex        =   57
         Top             =   165
         Width           =   1290
      End
      Begin VB.TextBox txt_almacen 
         Enabled         =   0   'False
         Height          =   345
         Left            =   7755
         TabIndex        =   56
         Top             =   165
         Width           =   1125
      End
      Begin VB.Frame frm_cantidad 
         Height          =   840
         Left            =   7515
         TabIndex        =   33
         Top             =   1725
         Width           =   2910
         Begin VB.TextBox txt_cantidad 
            Height          =   330
            Left            =   75
            TabIndex        =   34
            Top             =   375
            Width           =   2745
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000D&
            Caption         =   " Cantidad"
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   35
            Top             =   15
            Width           =   2895
         End
      End
      Begin VB.Frame frm_eliminar 
         Height          =   840
         Left            =   7515
         TabIndex        =   30
         Top             =   2025
         Width           =   2910
         Begin VB.TextBox txt_cantidad_eliminar 
            Height          =   330
            Left            =   75
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
      Begin VB.TextBox txt_importe_total 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   9165
         TabIndex        =   26
         Top             =   6270
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.TextBox txt_cantidad_total 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7410
         TabIndex        =   25
         Top             =   6285
         Width           =   1725
      End
      Begin VB.TextBox txt_folio 
         Enabled         =   0   'False
         Height          =   345
         Left            =   10005
         TabIndex        =   17
         Top             =   585
         Width           =   1290
      End
      Begin VB.TextBox txt_descuentos 
         Enabled         =   0   'False
         Height          =   345
         Left            =   7755
         TabIndex        =   16
         Top             =   585
         Width           =   1125
      End
      Begin VB.TextBox txt_cliente 
         Height          =   345
         Left            =   1200
         TabIndex        =   13
         Top             =   960
         Width           =   1125
      End
      Begin VB.TextBox txt_agente 
         Height          =   345
         Left            =   1200
         TabIndex        =   9
         Top             =   165
         Width           =   1125
      End
      Begin MSComctlLib.ListView lv_detalle 
         Height          =   4485
         Left            =   90
         TabIndex        =   18
         Top             =   1785
         Width           =   11220
         _ExtentX        =   19791
         _ExtentY        =   7911
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Precio"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Oferta"
            Object.Width           =   1640
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Precio Oferta"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Disponible"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Canitdad"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Importe"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Titular:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   67
         Top             =   630
         Width           =   480
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Remision Agte:"
         Height          =   195
         Left            =   120
         TabIndex        =   65
         Top             =   1440
         Width           =   1065
      End
      Begin VB.Label Label13 
         Caption         =   "Almacen:"
         Height          =   225
         Left            =   6825
         TabIndex        =   59
         Top             =   225
         Width           =   795
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Remision Sist.:"
         Height          =   195
         Left            =   8940
         TabIndex        =   58
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6555
         TabIndex        =   24
         Top             =   6345
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Descuentos:"
         Height          =   195
         Left            =   6825
         TabIndex        =   23
         Top             =   660
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Folio:"
         Height          =   195
         Left            =   8955
         TabIndex        =   22
         Top             =   660
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   1035
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmremision_vistas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tipo_lista  As Integer
Dim var_descuento_1 As Double
Dim var_descuento_2 As Double
Dim var_clave_almacen As String
Dim var_primera_vez_remision As Integer
Dim var_numero_remision As Double
Dim var_estatus_remision As String
Dim var_nombre_articulo_mensaje As String
Dim var_nombre_tabla As String
Dim var_consecutivo As Double
Dim var_numero_pedido_cliente As Double
Dim var_origen As String
Dim var_transporto As String
Dim var_tipo_proveedor As String
Dim var_primera_vez As Boolean
Dim var_numero_folio As Double
Dim VAR_TABLA_NOMBRE_ORIGEN As String
Dim VAR_RUTA_TABLA_ORIGEN As String
Dim VAR_CAMPO_CODIGO_ORIGEN As String
Dim VAR_CAMPO_DESCRIPCION_ORIGEN As String
Dim VAR_CAMPO_COSTO_ORIGEN As String
Dim VAR_CAMPO_CANTIDAD_ORIGEN As String
Dim VAR_CAMPO_CANTIDAD_ENTRADA As String
Dim VAR_TABLA_DESTINO As String
Dim VAR_CAMPO_CODIGO_DESTINO As String
Dim VAR_CAMPO_DESCRIPCION_DESTINO As String
Dim VAR_CAMPO_COSTO_DESTINO As String
Dim VAR_CAMPO_CANTIDAD_DESTINO  As String
Dim VAR_CAMPO_NUMERO  As String
Dim var_cantidad_enviada As Double
Dim var_cantidad_recibida As Double
Dim var_articulo_enviado As String
Dim var_costo_enviado As Double
Dim var_almacen_Destino As String
Dim var_almacen_origen As String
Dim var_proveedor As String
Dim var_estatus_movimiento As String
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_modifica As Boolean
Dim var_factura As String
Dim var_cantidad_leida As Double
Dim var_tabla As ADODB.Connection
Dim var_ruta As String
Dim var_folio_enviado As Double
Dim var_referencia As String
Dim var_suma_cantidad_enviada As Double
Dim var_suma_cantidad_recibida As Double
Dim var_orden_surtido As Double
Dim var_clave_agente As String
Dim var_clave_establecimiento As String
Dim var_clave_titular As String
Dim var_clave_cliente As String
Dim var_clave_ruta As String
Dim var_plazo As Integer
Dim var_iva As Variant
Dim var_agrupador As String
Dim var_correo_electronico As String
Dim var_autorizo_embarque As Boolean
Dim var_es_caja As Boolean
Dim var_cajas As Boolean
Dim var_almacen_OS As String
Dim var_nota As recordSet
Dim var_movimiento_dependencia As String
Dim var_clave_moneda As String
Dim var_factura_ceros As Integer
Dim var_renglon As Double

Private Sub cmd_aplicacion_anticipos_Click()
   If var_empresa = "31" Then
      If var_estatus_remision <> "F" Then
         cnn.CommandTimeout = 360
         var_n = lv_detalle.ListItems.Count
         var_posible = False
         For var_i = 1 To var_n
             lv_detalle.ListItems.Item(var_i).Selected = True
             If lv_detalle.selectedItem.SubItems(6) * 1 > 0 Then
                var_posible = True
             End If
         Next var_i
         VAR_POSIBLE_COdIGO = 1
         For var_i = 1 To var_n
             lv_detalle.ListItems.Item(var_i).Selected = True
             If lv_detalle.selectedItem = "S1005" Then
                VAR_POSIBLE_COdIGO = 0
             End If
         Next var_i
         If VAR_POSIBLE_COdIGO = 1 Then
            var_codigo_seleccionado = "S1005"
            var_codigo_anticipo = "S1005"
            var_consecutivo_anticipo = 0
            var_cliente_anticipo = Me.txt_cliente
            frmsaldo_anticipos.Show 1
            var_cantidad_leida = CDbl(var_importe_anticipo)
            rsaux.Open "SELECT * FROM TB_dETALLE_LISTA_PRECIOS WHERE VCHA_aRT_ARTICULO_ID = 'S1005'", cnn, adOpenDynamic, adLockOptimistic
            Set list_item = lv_detalle.ListItems.Add(, , rsaux!vcha_Art_Articulo_id)
            rsaux2.Open "SELECT * FROM TB_aRTICULOS WHERE VCHA_aRT_ARTICULO_ID = 'S1005'", cnn, adOpenDynamic, adLockOptimistic
            list_item.SubItems(1) = IIf(IsNull(rsaux2!vcha_art_nombre_Español), "", rsaux2!vcha_art_nombre_Español)
            rsaux2.Close
            var_precio = IIf(IsNull(rsaux!floa_dli_precio), 0, rsaux!floa_dli_precio) * 1.16
            list_item.SubItems(2) = Format(var_precio, "###,###,##0.00")
            list_item.SubItems(3) = Format(0, "###,###,##0.00")
            list_item.SubItems(4) = Format(var_precio, "###,###,##0.00")
            list_item.SubItems(5) = Format(0, "###,###,##0.00")
            list_item.SubItems(6) = Format(var_cantidad_leida, "###,###,##0.00")
            list_item.SubItems(7) = Format(var_cantidad_leida * -1, "###,###,##0.00")
            rsaux2.Open "insert into tb_Remisiones (vcha_emp_empresa_id, vcha_alm_almacen_id,vcha_age_agente_id,vcha_cli_clave_id, inte_rem_numero,floa_rem_cantidad, vcha_Art_articulo_id, VCHA_REM_REMISION_AGENTE) values ('" + var_empresa + "', '" + Me.txt_almacen + "','" + Me.txt_agente + "', '" + Me.txt_cliente + "'," + CStr(var_numero_remision) + "," + CStr(CDbl(var_cantidad_leida)) + ",'S1005','" + Me.txt_remision_agente + "' )", cnn, adOpenDynamic, adLockOptimistic
            rsaux.Close
            If Me.txt_cantidad_total = "" Then
               Me.txt_cantidad_total = 0
            End If
            Me.txt_cantidad_total = Me.txt_cantidad_total + var_cantidad_leida
            
            x = CDbl(1) / (1 - (var_descuento_1 / 100))
            x = CDbl(x) / (1 - (var_descuento_2 / 100))
            z = (x * 1) * CDbl(var_cantidad_leida * 1)

            Me.txt_importe_total = Me.txt_importe_total - z
            
            
            
            
            If var_importe_anticipo > 0 Then
               'rs.Open "select * from tb_existencias where vcha_alm_almacen_id = '" + Me.txt_almacen_origen + "' and vcha_art_articulo_id = 'S1005'", cnn, adOpenDynamic, adLockOptimistic
               'If rs.EOF Then
               '   rsaux.Open "insert into tb_existencias (vcha_Alm_almacen_id, vcha_Art_Articulo_id, floa_exi_cantidad, floa_Exi_costo, floa_Exi_cantidad_2004, floa_Exi_costo_2004, floa_exi_cantidad_2005, floa_exi_costo_2005) values ('" + Me.txt_almacen_origen + "','S1005',0,0,0,0,0,0)", cnn, adOpenDynamic, adLockOptimistic
               'End If
               'rs.Close
          
               rsaux.Open "select * from tb_anticipos where vcha_cli_clave_id = '" + Me.txt_cliente + "' and floa_sal_cantidad - floa_ant_aplicado > 0 order by dtim_sal_Fecha", cnn, adOpenDynamic, adLockOptimistic
               While Not rsaux.EOF
                     var_cantidad_disponible = IIf(IsNull(rsaux!floa_Sal_Cantidad), 0, rsaux!floa_Sal_Cantidad) - IIf(IsNull(rsaux!floa_ant_aplicado), 0, rsaux!floa_ant_aplicado)
                     If var_importe_anticipo > 0 Then
                        If var_cantidad_disponible >= var_importe_anticipo Then
                           rsaux1.Open "update tb_anticipos set floa_ant_aplicado = isnull(floa_ant_aplicado,0) + " + CStr(var_importe_anticipo) + " where inte_ant_consecutivo = " + CStr(rsaux!INTE_ANT_CONSECUTIVO), cnn, adOpenDynamic, adLockOptimistic
                           var_cadena = "insert into tb_aplicacion_anticipos (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_mov_movimiento_id, inte_emo_numero, inte_ant_consecutivo, floa_ant_importe) "
                           var_cadena = var_cadena + " values ('" + var_empresa + "','" + var_unidad_organizacional + "','FV'," + CStr(Me.txt_remision) + ", " + CStr(rsaux!INTE_ANT_CONSECUTIVO) + "," + CStr(var_importe_anticipo) + ")"
                           rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                           var_importe_anticipo = 0
                        Else
                           var_importe_aplicar = var_cantidad_disponible
                           rsaux1.Open "update tb_anticipos set floa_ant_aplicado = isnull(floa_ant_aplicado,0) + " + CStr(var_importe_aplicar) + " where inte_ant_consecutivo = " + CStr(rsaux!INTE_ANT_CONSECUTIVO), cnn, adOpenDynamic, adLockOptimistic
                           var_importe_anticipo = var_importe_anticipo - var_importe_aplicar
                           var_cadena = "insert into tb_aplicacion_anticipos (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_mov_movimiento_id, inte_emo_numero, inte_ant_consecutivo, floa_ant_importe) "
                           var_cadena = var_cadena + " values ('" + var_empresa + "','" + var_unidad_organizacional + "','FV'," + CStr(Me.txt_remision) + ", " + CStr(rsaux!INTE_ANT_CONSECUTIVO) + "," + CStr(var_importe_aplicar) + ")"
                           rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        End If
                     End If
                     rsaux.MoveNext
               Wend
               rsaux.Close
            End If
         Else
            MsgBox "Ya existe un anticipo en el movimiento", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El movimiento ya fue cerrado", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No es posible aplicación de anticipos", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_buscar_remision_Click()
   Me.frm_numero_nota.Visible = True
   Me.txt_numero_nota = ""
   Me.txt_numero_nota.SetFocus
End Sub

Private Sub cmd_cancelar_Click()
   rs.Open "select * from tb_Series where vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      Me.txt_serie = IIf(IsNull(rs!vcha_ser_Serie_id), "", rs!vcha_ser_Serie_id)
   End If
   rs.Close
   Me.txt_numero = ""
   Me.txt_nombre_agente_cancelar = ""
   Me.txt_nombre_cliente_cancelar = ""
   Me.txt_fecha = ""
   Me.txt_importe = ""
   lbl_moneda = ""
   Me.frm_factura_cancelar.Visible = True
   Me.txt_numero.SetFocus
End Sub

Private Sub cmd_cancelar_GotFocus()
   Me.frm_factura_cancelar.Visible = False
End Sub

Private Sub cmd_imprimir_Click()
   Dim pError As ADODB.Error
   Dim var_actualiza As Boolean
   Dim var_inserta As Boolean
   Dim bandera_suma As Boolean
   Dim var_cantidad_1 As Variant
   Dim var_costo As Variant
   Dim var_precio_1 As Variant
   Dim var_posible_caja As Boolean
   Dim var_cantidad_posible As Variant
   Dim var_embarque_paquete As Integer
   Dim var_embarque_caja As Integer
   Dim var_estatus_caja As String
   Dim var_orden_surtido_caja As Integer
   Dim var_posible_empaque As Boolean
   Dim var_promocion_1 As Double
   Dim var_promocion_2 As Double
   Dim var_encontrado As Integer
   Dim var_canal_venta As String
   Dim var_i As Integer
   Dim var_n As Integer
   Dim var_j As Integer
   Dim var_tipo_pedido As String
   Dim var_posible As Boolean
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_SALIDAS_INSERTA = New TB_TEMPORAL_SALIDAS_INSERTA
   Set TB_TEMPORAL_SALIDAS_MODIFICA = New TB_TEMPORAL_SALIDAS_MODIFICA
   Set TB_DET_ORDEN_SURTIDO_M = New TB_DET_ORDEN_SURTIDO_M
   Set TB_DET_EMBARQUE_I = New TB_DET_EMBARQUE_I
   Set TB_DETALLE_CAJAS_M = New TB_DETALLE_CAJAS_M
   z = 0
   
   Dim var_numero_movimientos As Double
   Dim var_numero_factura_inicio As Double
   Dim var_k As Double
   Dim var_cliente As String
   Dim var_expedicion As String
   Dim var_domicilio As String
   Dim var_ciudad As String
   Dim var_agente As String
   Dim var_linea As String
   Dim var_cantidad As String
   
   Dim var_precio As Double
   Dim var_precio_str As String
   Dim var_importe As String
   Dim var_subimporte As String
   Dim var_cantidad_letra As String
   Dim var_iva As String
   Dim var_rfc As String
   Dim var_dia As String
   Dim var_mes As String
   Dim var_año As String
   Dim var_porcentaje As Double
   Dim var_Archivo As String
   Dim var_importe_descuento_1 As Double
   Dim var_importe_descuento_2 As Double
   Dim var_importe_descuento_3 As Double
   Dim var_importe_descuento_1_2 As Double
   Dim var_importe_descuento_2_2 As Double
   Dim var_importe_descuento_3_2 As Double
   Dim var_importe_descuento_1_str As String
   Dim var_importe_descuento_2_str As String
   Dim var_importe_descuento_3_str As String
   Dim var_marca_promocion As String
   Dim var_contador_promociones As Double
   Dim var_cantidad_total As Double
   Dim var_cantidad_total_str As String
   Dim var_factura_envio As Double
   Dim var_consecutivo As Double
   Dim var_x As Double
   If var_estatus_remision <> "F" Then
   cnn.CommandTimeout = 360
   var_n = lv_detalle.ListItems.Count
   var_posible = False
   For var_i = 1 To var_n
       lv_detalle.ListItems.Item(var_i).Selected = True
       If lv_detalle.selectedItem.SubItems(6) * 1 > 0 Then
          var_posible = True
       End If
   Next var_i
   If lv_detalle.ListItems.Count > 0 Then
      If var_posible = True Then
         var_si = MsgBox("¿Desea imprimir y cerrar el movimiento?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_si = MsgBox("Confirmar la impresion del documento", vbYesNo, "ATENCION")
            If var_si = 6 Then
               rs.Open "select * from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
               var_factura_inicio = rs!inte_ser_factura
               var_serie = rs!vcha_ser_Serie_id
               rs.Close
               MsgBox "Se va a imprimir la factura " + Trim(Str(var_factura_inicio)), vbOKOnly, "ATENCION"
               si = MsgBox("¿La impresora esta lista?", vbYesNo, "ATENCION")
               If si = 6 Then
                  var_inserta = False
                  cnn.BeginTrans
                  
                  var_estatus_remision = "F"
                  rsaux.Open "select vcha_can_canal_venta_id from tb_agentes where vcha_age_agente_id = '" + txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
                  var_canal_venta = IIf(IsNull(rsaux!vcha_can_canal_venta_id), "", rsaux!vcha_can_canal_venta_id)
                  rsaux.Close
                  var_almacen_origen = var_clave_almacen
                  rs.Open "select * from vw_clientes where vcha_cli_clave_id = '" + txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_clave_moneda = rs!vcha_mon_moneda_id
                     var_clave_titular = rs!vcha_tit_titular_id
                     var_canal_venta = rs!vcha_can_canal_venta_id
                  End If
                  rs.Close
                  rs.Open "select * from tb_Detalle_Establecimientos where vcha_cli_clave_id = '" + txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_clave_establecimiento = rs!vcha_ESB_ESTABLECIMIENTO_id
                  End If
                  rs.Close
                  var_clave_movimiento = "FV"
                  var_numero_folio = 0
                  var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, Now, CDbl(var_numero_folio), 0, txt_cliente, "", var_almacen_origen, "", "", var_clave_usuario_global, fun_NombrePc, 0, "", "", var_clave_establecimiento, "", var_clave_titular, txt_agente, var_descuento_1, var_descuento_2, 0, var_clave_moneda, 0)
                  var_numero_folio = var_numero_folio_regreso
                  rsaux.Open "update tb_encabezado_movimientos set  VCHA_CAN_CANAL_VENTA_ID = '" + var_canal_venta + "' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
                  var_inserta = False
                  txt_folio = var_numero_folio
                  var_primera_vez = False
                  var_consecutivo = 0
                  For var_i = 1 To var_n
                      lv_detalle.ListItems.Item(var_i).Selected = True
                      If lv_detalle.selectedItem.SubItems(6) * 1 > 0 Then
                         var_consecutivo = var_consecutivo + 1
                         var_promocion_1 = 0
                         
                         If lv_detalle.selectedItem.SubItems(3) * 1 > 0 Then
                            var_promocion_1 = lv_detalle.selectedItem.SubItems(3) * 1
                         End If
                         var_precio_1 = 0
                         
                         If var_empresa = "31" And (Me.lv_detalle.selectedItem = "S1005" Or Me.lv_detalle.selectedItem = "S1003") Then
                            var_precio_1 = (lv_detalle.selectedItem.SubItems(4) * 1) / 1.16
                         Else
                            If lv_detalle.selectedItem.SubItems(4) * 1 > 0 Then
                               var_precio_1 = (lv_detalle.selectedItem.SubItems(4) * 1) / 1.16
                            End If
                         End If
                         
                         rsaux2.Open "INSERT INTO tb_temporal_salidas (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, CHAR_PED_TIPO, INTE_SAL_CONSECUTIVO) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", '" + Trim(lv_detalle.selectedItem) + "', " + CStr(CDbl(lv_detalle.selectedItem.SubItems(6))) + ", 0, " + CStr(var_precio_1) + ", 0, " + CStr(var_promocion_1) + ", 0, 'M', " + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                         If Me.lv_detalle.selectedItem = "S1005" Then
                            rsaux2.Open "UPDATE TB_REMISIONES SET INTE_EMO_NUMERO = " + CStr(var_numero_folio) + ", FLOA_DLI_PRECIO = " + CStr(CDbl(Me.lv_detalle.selectedItem.SubItems(2))) + ", CHAR_REM_ESTATUS = 'F', FLOA_REM_PRECIO = " + CStr(var_precio_1) + ", FLOA_REM_PROMOCION_1 = " + CStr(CDbl(lv_detalle.selectedItem.SubItems(3))) + ", FLOA_REM_DESCUENTO_1 = 0, FLOA_REM_DESCUENTO_2 = 0 WHERE INTE_REM_NUMERO = " + CStr(var_numero_remision) + " AND VCHA_aRT_ARTICULO_ID = '" + lv_detalle.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                         Else
                            rsaux2.Open "UPDATE TB_REMISIONES SET INTE_EMO_NUMERO = " + CStr(var_numero_folio) + ", FLOA_DLI_PRECIO = " + CStr(CDbl(Me.lv_detalle.selectedItem.SubItems(2))) + ", CHAR_REM_ESTATUS = 'F', FLOA_REM_PRECIO = " + CStr(var_precio_1) + ", FLOA_REM_PROMOCION_1 = " + CStr(CDbl(lv_detalle.selectedItem.SubItems(3))) + ", FLOA_REM_DESCUENTO_1 = " + CStr(var_descuento_1) + ", FLOA_REM_DESCUENTO_2 = " + CStr(var_descuento_2) + " WHERE INTE_REM_NUMERO = " + CStr(var_numero_remision) + " AND VCHA_aRT_ARTICULO_ID = '" + lv_detalle.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                         End If
                         var_tipo_Cambio = 1
                         var_catalogo_1 = ""
                         var_catalogo_2 = ""
                         var_año_catalogo = 0
                         var_mes_catalogo = 0
                         var_si_surtir_catalogo = 0
                      End If
                  Next var_i
                  If rsaux4.State = 1 Then
                     rsaux4.Close
                  End If
                  Text1 = "exec SP_INSERTA_MOVIMIENTOS_SALIDA_EMBARQUES '" + var_empresa + "','" + var_unidad_organizacional + "','" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", " + CStr(var_tipo_Cambio) + ",'" + var_catalogo_1 + "','" + var_catalogo_2 + "','" + var_clave_titular + "','" + txt_cliente + "'," + CStr(var_año_catalogo) + "," + CStr(var_mes_catalogo) + "," + CStr(var_si_surtir_catalogo)
                  rsaux4.Open "exec SP_INSERTA_MOVIMIENTOS_SALIDA_EMBARQUES '" + var_empresa + "','" + var_unidad_organizacional + "','" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", " + CStr(var_tipo_Cambio) + ",'" + var_catalogo_1 + "','" + var_catalogo_2 + "','" + var_clave_titular + "','" + txt_cliente + "'," + CStr(var_año_catalogo) + "," + CStr(var_mes_catalogo) + "," + CStr(var_si_surtir_catalogo), cnn, adOpenDynamic, adLockOptimistic
                  'rsaux5.Open "select sum(floa_Sal_Cantidad) from tb_salidas where vcha_mov_movimiento_id = 'fv' and inte_sal_numero = 105", cnn, adOpenDynamic, adLockOptimistic
                  'MsgBox CStr(rsaux5(0).Value)
                  'rsaux5.Close
                  
                  rs.Open "select * from vw_maximo_embarque where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                  If rs.EOF Then
                     var_numero_embarque = 1
                  Else
                     var_numero_embarque = rs!maximo_embarque + 1
                  End If
                  rs.Close
                  Set TB_ENC_EMBARQUE_I = New TB_ENC_EMBARQUE_I
                  ok = False
                  rs.Open "insert into tb_encabezado_embarques (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, INTE_EMB_EMBARQUE, INTE_JAU_JAULA_ID, VCHA_VEH_VEHICULO_ID, VCHA_AGE_AGENTE_ID, DTIM_EMB_FECHA_INICIO, DTIM_EMB_FECHA_FINAL, CHAR_EMB_ESTATUS, VCHA_CHO_CHOFER_ID, FLOA_EMB_CUBICAJE, CHAR_EMB_TIPO, INTE_EMB_BLOQUEADO, VCHA_EMB_BLOQUEADO_POR, VCHA_AUD_MAQUINA, VCHA_AUD_USUARIO) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', " + CStr(var_numero_embarque) + ", 0, '', '" + txt_agente + "', getdate(), '', '', '', 0,'',0, '','" + fun_NombrePc + "','" + var_clave_usuario_global + "')", cnn, adOpenDynamic, adLockOptimistic
                  var_inserta = TB_DET_EMBARQUE_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_numero_embarque, var_clave_movimiento, var_numero_folio, "")
                  txt_numero_embarque = CStr(var_numero_embarque)
                  var_estatus_embarque = "I"
               
                  If Trim(txt_numero_embarque) <> "" Then
                     'Sirve para validar que no vaya mercancia con cantidad en NULL
                     Cadena = "SELECT     dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID, "
                     Cadena = Cadena + " dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO, dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID,"
                     Cadena = Cadena + " dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID , dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD"
                     Cadena = Cadena + " FROM         dbo.TB_DETALLE_EMBARQUES INNER JOIN"
                     Cadena = Cadena + " dbo.TB_SALIDAS ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND"
                     Cadena = Cadena + " dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID AND"
                     Cadena = Cadena + " dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO = dbo.TB_SALIDAS.INTE_SAL_NUMERO AND"
                     Cadena = Cadena + " dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID"
                     Cadena = Cadena + " WHERE     (dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD IS NULL) AND (dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = " + CStr(var_numero_embarque) + ") AND"
                     Cadena = Cadena + " (dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "')"
                     If rsaux4.State = 1 Then
                        rsaux4.Close
                     End If
                     rsaux4.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                     'MsgBox CStr(rsaux4.RecordCount)
                     If Not rsaux4.EOF Then
                        rsaux4.Close
                        MsgBox "El movimiento tiene cantidad en NULL", vbOKOnly, "ATENCION"
                     Else
                        fecha_inicio = CStr(Now)
                        Set TB_ENC_EMBARQUE_M = New TB_ENC_EMBARQUE_M
                        'MsgBox txt_numero_embarque
                        
                        rs.Open "execute factura_embarques_vistas '" + var_empresa + "', '" + var_unidad_organizacional + "', " + txt_numero_embarque + ", '', '','" + var_serie + "', 'FA'", cnn, adOpenDynamic, adLockOptimistic
                        ok = TB_ENC_EMBARQUE_M.Anadir(var_empresa, var_unidad_organizacional, CDbl(txt_numero_embarque), "")
                        fecha_fin = CStr(Now)
                        var_estatus_embarque = "F"
                       'aqui se imprime la factura
                        If rs.State = 1 Then
                           rs.Close
                        End If
                        rs.Open "select isnull(max(inte_tem_consecutivo),0) from tb_temp_factura_embarques", cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           var_consecutivo = rs(0).Value
                        Else
                           var_consecutivo = 0
                        End If
                        rs.Close
                        var_consecutivo = var_consecutivo + 1
                        rs.Open "insert into tb_temp_factura_embarques (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                     
                        Cadena = "EXEC SP_CREA_TABLA_FACTURAS_VISTAS " + CStr(var_consecutivo) + ",'" + var_empresa + "'," + txt_numero_embarque
                        If rsaux3.State = 1 Then
                           rsaux3.Close
                        End If
                        rsaux3.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                        rsaux3.Open "select distinct inte_car_numero from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " ORDER BY INTE_CAR_NUMERO", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux3.EOF Then
                           var_Archivo = App.Path & "\factura" + Trim(Str(rsaux3!inte_car_numero)) + ".bat"
                           Open (App.Path & "\factura" + Trim(Str(rsaux3!inte_car_numero)) + ".bat") For Output As #2
                           While Not rsaux3.EOF
                                 If rs.State = 1 Then
                                    rs.Close
                                 End If
                                 If var_empresa <> "03" Then
                                    rs.Open "select * from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_Car_numero = " + Str(rsaux3!inte_car_numero) + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " ORDER BY INTE_SAL_CONSECUTIVO_FACTURA", cnn, adOpenDynamic, adLockOptimistic
                                 Else
                                    rs.Open "select * from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_Car_numero = " + Str(rsaux3!inte_car_numero) + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " ORDER BY vcha_sal_descripcion_factura", cnn, adOpenDynamic, adLockOptimistic
                                 End If
                                 If Not rs.EOF Then
                                   'AQUI EMPIEZA LA FACTURA
                                    Open (App.Path & "\factura" + Trim(Str(rsaux3!inte_car_numero)) + ".txt") For Output As #1
                                    'Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
                                    'Print #1, Chr(27) + Chr(67) + Chr(44) + Chr(13)
                                    'Print #1, ""
                                    Print #1, Chr(15) + Chr(27) + Chr(64)
                                    If var_empresa = "18" Then
                                       Print #1, ""
                                    End If
                                    Print #1, Spc(105); Str(rsaux3!inte_car_numero)
                                    Print #1, ""
                                    Print #1, Spc(105); Str(rs!INTE_CAR_PLAZO) + " DIAS DE VENCIMIENTO" + "                  " + Format(rs!dtim_Car_fecha, "Short Date")
                                    Print #1, ""
                                    Print #1, ""
                                    'Print #1, Spc(92); Str(rs!inte_car_PLAZO) + " DIAS DE VENCIMIENTO"
                                    var_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                                    For var_j = 1 + Len(Trim(var_cliente)) To 83
                                        var_cliente = var_cliente + " "
                                    Next var_j
                                    If var_unidad_organizacional = "21" Then
                                       var_cliente = var_cliente + "               MEXICO, D.F."
                                    Else
                                       var_cliente = var_cliente + "               AGUASCALIENTES, AGS."
                                    End If
                                    Print #1, Spc(10); var_cliente
                                    var_domicilio = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " C.P. " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
                                    'For var_j = 1 + Len(Trim(var_domicilio)) To 83
                                    '    var_domicilio = var_domicilio + " "
                                    'Next var_j
                                    
                                    rsaux11.Open "select vcha_cli_referencia from tb_Clientes where vcha_Cli_clave_id = '" + IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
                                    If Not rsaux11.EOF Then
                                       var_referencia_Bancaria = IIf(IsNull(rsaux11!VCHA_CLI_REFERENCIA), "", rsaux11!VCHA_CLI_REFERENCIA)
                                       If var_referencia_Bancaria <> "" Then
                                          For var_j = 1 + Len(Trim(var_domicilio)) To 105
                                              var_domicilio = var_domicilio + " "
                                          Next var_j
                                          var_domicilio = var_domicilio + " REF. BANCARIA: " + var_referencia_Bancaria
                                       End If
                                    End If
                                    rsaux11.Close
                                    
                                    
                                    
                                    
                                    var_agente = ""
                                    var_agente = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
                                    For var_j = 1 + Len(Trim(var_agente)) To 8
                                        var_agente = var_agente + " "
                                    Next var_j
                                    If rsaux4.State = 1 Then
                                       rsaux4.Close
                                    End If
                                    rsaux4.Open "SELECT * FROM TB_AGENTES WHERE VCHA_AGE_AGENTE_ID = '" + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
                                    If Not rsaux4.EOF Then
                                       var_agente = var_agente + IIf(IsNull(rsaux4!VCHA_AGE_NOMBRE), "", rsaux4!VCHA_AGE_NOMBRE)
                                    Else
                                       var_agente = var_agente + ""
                                    End If
                                    rsaux4.Close
                                    var_domicilio = var_domicilio
                                    'Print #1, Spc(111); var_agente
                                    Print #1, Spc(10); var_domicilio
                                    var_ciudad = ""
                                    var_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                                    For var_j = 1 + Len(Trim(var_ciudad)) To 37
                                        var_ciudad = var_ciudad + " "
                                    Next var_j
                                 
                                    var_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                                    var_ciudad = var_ciudad
                                    var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                                    var_ciudad = var_ciudad + var_rfc
                                             
                                    For var_j = 1 + Len(Trim(var_estado)) To 46
                                        var_estado = var_estado + " "
                                    Next var_j
                                 
   
                                    For var_j = 1 + Len(Trim(var_ciudad)) To 14
                                        var_ciudad = var_ciudad + " "
                                    Next var_j
                               
                                    var_ciudad = var_ciudad + "                                                      " + var_agente
                                
                                    VAR_EMBARQUE = "EMB.: " + txt_numero_embarque
                                    var_ordern_surtido = x
                                    Print #1, Spc(10); var_ciudad
                                    var_rfc = "RFC:  " + var_rfc
                                    var_rfc = IIf(IsNull(rs!vcha_ESB_ESTABLECIMIENTO_id), "", rs!vcha_ESB_ESTABLECIMIENTO_id) + " " + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                                    For var_j = 1 + Len(Trim(var_rfc)) To 89
                                        var_rfc = var_rfc + " "
                                    Next var_j
                                    If var_empresa = "18" Then
                                       If rs!VCHA_MOV_MOVIMIENTO_ID = "FV" Then
                                          rsaux5.Open "SELECT * FROM TB_REMISIONES WHERE INTE_EMO_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                                          If Not rsaux5.EOF Then
                                             If IsNumeric(IIf(IsNull(rsaux5!VCHA_REM_REMISION_AGENTE), "0", rsaux5!VCHA_REM_REMISION_AGENTE)) Then
                                                var_rfc = var_rfc + "               R.A.: " + Trim(Str(IIf(IsNull(rsaux5!VCHA_REM_REMISION_AGENTE), "0", rsaux5!VCHA_REM_REMISION_AGENTE))) + " "
                                             Else
                                                var_rfc = var_rfc + "               R.A.: " + Trim((IIf(IsNull(rsaux5!VCHA_REM_REMISION_AGENTE), "0", rsaux5!VCHA_REM_REMISION_AGENTE))) + " "
                                             End If
                                             var_rfc = var_rfc + " R.S.: " + Trim(Str(IIf(IsNull(rsaux5!INTE_REM_NUMERO), 0, rsaux5!INTE_REM_NUMERO))) + " " + VAR_EMBARQUE
                                          Else
                                             var_rfc = var_rfc + "               PED.: " + Trim(Str(IIf(IsNull(rs!inte_ped_numero), 0, rs!inte_ped_numero))) + " "
                                             var_rfc = var_rfc + " O.S.: " + Trim(Str(IIf(IsNull(rs!INTE_ORS_ORDEN_SURTIDO), 0, rs!INTE_ORS_ORDEN_SURTIDO))) + " " + VAR_EMBARQUE
                                          End If
                                          rsaux5.Close
                                       Else
                                          var_rfc = var_rfc + "               PED.: " + Trim(Str(IIf(IsNull(rs!inte_ped_numero), 0, rs!inte_ped_numero))) + " "
                                          var_rfc = var_rfc + " O.S.: " + Trim(Str(IIf(IsNull(rs!INTE_ORS_ORDEN_SURTIDO), 0, rs!INTE_ORS_ORDEN_SURTIDO))) + " " + VAR_EMBARQUE
                                       End If
                                    Else
                                    var_rfc = var_rfc + "               PED.: " + Trim(Str(IIf(IsNull(rs!inte_ped_numero), 0, rs!inte_ped_numero))) + " "
                                    var_rfc = var_rfc + " O.S.: " + Trim(Str(IIf(IsNull(rs!INTE_ORS_ORDEN_SURTIDO), 0, rs!INTE_ORS_ORDEN_SURTIDO))) + " " + VAR_EMBARQUE
                                    End If
                                    Print #1, var_rfc
                                    'Print #1, Spc(10); IIf(IsNull(rs!vcha_esb_establecimiento_id), "", rs!vcha_esb_establecimiento_id)
                                    Print #1, ""
                                    Print #1, ""
                                    var_importe_descuento_1 = 0
                                    var_importe_descuento_2 = 0
                                    var_importe_descuento_3 = 0
                                    var_contador_promociones = 0
                                    var_cantidad_total = 0
                                    For var_k = 1 To var_renglones_factura
                                        If Not rs.EOF Then
                                           var_linea = ""
                                           var_marca_promocion = " "
                                           var_promocion_1 = IIf(IsNull(rs!floa_sal_promocion_1), 0, rs!floa_sal_promocion_1)
                                           var_promocion_2 = IIf(IsNull(rs!FLOA_SAL_PROMOCION_2), 0, rs!FLOA_SAL_PROMOCION_2)
                                           If var_promocion_1 > 0 Then
                                              var_marca_promocion = "*"
                                              var_contador_promociones = var_contador_promociones + 1
                                           End If
                                           If var_promocion_2 > 0 Then
                                              var_marca_promocion = "*"
                                              var_contador_promociones = var_contador_promociones + 1
                                           End If
                                           var_linea = IIf(IsNull(rs!vcha_agr_agrupador_id), "", rs!vcha_agr_agrupador_id)
                                           If var_empresa = "31" Then
                                              For var_j = 1 + Len(Trim(var_linea)) To 18
                                                  var_linea = var_linea + " "
                                              Next var_j
                                           Else
                                              For var_j = 1 + Len(Trim(var_linea)) To 15
                                                  var_linea = var_linea + " "
                                              Next var_j
                                           End If
                                           var_linea = var_linea + IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura)
                                           var_i = 0
                                           While Len((var_linea)) < 115
                                                 var_linea = var_linea + " "
                                           Wend
                                           var_linea = var_linea + " "
                                           var_linea = var_linea + var_marca_promocion
                                           var_cantidad = Format(IIf(IsNull(rs!Cantidad), 0, rs!Cantidad), "###,###,##0.00")
                                           var_cantidad_total = var_cantidad_total + IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)
                                           If Len(Trim(var_cantidad)) < 14 Then
                                              For var_j = 1 + Len(Trim(var_cantidad)) To 14
                                                  var_cantidad = " " + var_cantidad
                                              Next var_j
                                           End If
                                           var_precio = IIf(IsNull(rs!Importe), 0, rs!Importe)
                                           var_descuento_1 = IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1)
                                           var_descuento_2 = IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2)
                                           var_descuento_3 = IIf(IsNull(rs!floa_car_porcentaje_descuento_3), 0, rs!floa_car_porcentaje_descuento_3)
                                           var_porcentaje = (100 - var_descuento_1) / 100
                                           var_precio = var_precio * var_porcentaje
                                           var_importe_descuento_1_2 = (IIf(IsNull(rs!Importe), 0, rs!Importe) - var_precio)
                                           var_importe_descuento_1 = var_importe_descuento_1 + (IIf(IsNull(rs!Importe), 0, rs!Importe) - var_precio)
                                           var_precio = var_precio * ((100 - var_descuento_2) / 100)
                                           var_importe_descuento_2 = var_importe_descuento_2 + (IIf(IsNull(rs!Importe), 0, rs!Importe) - (var_importe_descuento_1_2 + var_precio))
                                           var_precio = var_precio * ((100 - var_descuento_3) / 100)
                                           var_precio = var_precio / IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)
                                           'var_precio_str = Format(var_precio / IIf(IsNull(rs!cantidad), 0, rs!cantidad), "###,###,##0.00")
                                           var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                                           If Len(Trim(var_rfc)) > 0 Then
                                              var_precio_str = Format(IIf(IsNull(rs!Importe), 0, rs!Importe) / IIf(IsNull(rs!Cantidad), 0, rs!Cantidad), "###,###,##0.00")
                                           Else
                                              var_precio_str = Format((IIf(IsNull(rs!Importe), 0, rs!Importe) / IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)) * (1 + (rs!floa_car_porcentaje_iva / 100)), "###,###,##0.00")
                                           End If
                                           If Len(Trim(var_precio_str)) < 14 Then
                                              For var_j = 1 + Len(Trim(var_precio_str)) To 14
                                                  var_precio_str = " " + var_precio_str
                                              Next var_j
                                           End If
                                           var_linea = var_linea + var_cantidad + var_precio_str
                                           If Len(Trim(var_rfc)) > 0 Then
                                              var_importe = Format((IIf(IsNull(rs!Importe), 0, rs!Importe)), "###,###,##0.00")
                                              If Len(Trim(var_importe)) < 14 Then
                                                 For var_j = 1 + Len(Trim(var_importe)) To 14
                                                     var_importe = " " + var_importe
                                                 Next var_j
                                              End If
                                           Else
                                              var_importe = Format((IIf(IsNull(rs!Importe), 0, rs!Importe) * (1 + (rs!floa_car_porcentaje_iva / 100))), "###,###,##0.00")
                                              If Len(Trim(var_importe)) < 14 Then
                                                 For var_j = 1 + Len(Trim(var_importe)) To 14
                                                     var_importe = " " + var_importe
                                                 Next var_j
                                              End If
                                           End If
                                           var_linea = var_linea + var_importe
                                     
                                           Print #1, var_linea
                                           rs.MoveNext
                                        Else
                                           Print #1, ""
                                        End If
                                    Next var_k
                                    Print #1, ""
                                    'Print #1, ""
                                    rs.MoveFirst
                                
                                    var_cantidad_total_str = Format(var_cantidad_total, "###,###,##0.00")
                                    var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                                    If Len(Trim(var_rfc)) > 0 Then
                                       var_cantidad_letra = rs!vcha_car_importe_letra
                                       var_importe_descuento_1_str = Format(IIf(IsNull(rs!FLOA_CAR_IMPORTE_DESCUENTO_1), 0, rs!FLOA_CAR_IMPORTE_DESCUENTO_1) / IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio), "###,###,##0.00")
                                       If Len(Trim(var_importe_descuento_1_str)) < 14 Then
                                          For var_j = 1 + Len(Trim(var_importe_descuento_1_str)) To 14
                                              var_importe_descuento_1_str = " " + var_importe_descuento_1_str
                                          Next var_j
                                       End If
                                       var_importe_descuento_2_str = Format(IIf(IsNull(rs!FLOA_CAR_IMPORTE_DESCUENTO_2), 0, rs!FLOA_CAR_IMPORTE_DESCUENTO_2) / IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio), "###,###,##0.00")
                                       If Len(Trim(var_importe_descuento_2_str)) < 14 Then
                                          For var_j = 1 + Len(Trim(var_importe_descuento_2_str)) To 14
                                              var_importe_descuento_2_str = " " + var_importe_descuento_2_str
                                          Next var_j
                                       End If
                                    Else
                                       var_cantidad_letra = rs!vcha_car_importe_letra
                                       var_importe_descuento_1_str = Format((IIf(IsNull(rs!FLOA_CAR_IMPORTE_DESCUENTO_1), 0, rs!FLOA_CAR_IMPORTE_DESCUENTO_1)) * (1 + (rs!floa_car_porcentaje_iva / 100) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio))), "###,###,##0.00")
                                       If Len(Trim(var_importe_descuento_1_str)) < 14 Then
                                          For var_j = 1 + Len(Trim(var_importe_descuento_1_str)) To 14
                                              var_importe_descuento_1_str = " " + var_importe_descuento_1_str
                                          Next var_j
                                       End If
                                       var_importe_descuento_2_str = Format((IIf(IsNull(rs!FLOA_CAR_IMPORTE_DESCUENTO_2), 0, rs!FLOA_CAR_IMPORTE_DESCUENTO_2)) * (1 + (rs!floa_car_porcentaje_iva / 100) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio))), "###,###,##0.00")
                                       If Len(Trim(var_importe_descuento_2_str)) < 14 Then
                                          For var_j = 1 + Len(Trim(var_importe_descuento_2_str)) To 14
                                              var_importe_descuento_2_str = " " + var_importe_descuento_2_str
                                          Next var_j
                                       End If
                                    End If
                                    var_linea = "- DESCUENTO DEL " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1))) + "%"
                                    If Len(Trim(var_linea)) < 145 Then
                                       For var_j = 1 + Len(Trim(var_linea)) To 145
                                           var_linea = var_linea + " "
                                       Next var_j
                                    End If
                                    Print #1, var_linea + var_importe_descuento_1_str
                                    var_linea = "- DESCUENTO POR PAGO OPORTUNO " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2))) + "%"
                                    If Len(Trim(var_linea)) < 145 Then
                                       For var_j = 1 + Len(Trim(var_linea)) To 145
                                           var_linea = var_linea + " "
                                       Next var_j
                                    End If
                                    var_linea = var_linea + var_importe_descuento_2_str
                                    Print #1, var_linea
                                    If var_contador_promociones > 0 Then
                                       Print #1, "PROMOCION EN ARTICULOS MARCADOS CON *"
                                    Else
                                       Print #1, ""
                                    End If
                                             
                                    var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                                    var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
                                            
                                    If Len(Trim(var_linea)) < 117 Then
                                       For var_j = 1 + Len(Trim(var_linea)) To 117
                                           var_x = var_j Mod 2
                                           If var_x >= 1 Then
                                              var_linea = " " + var_linea
                                           Else
                                              var_linea = var_linea + " "
                                           End If
                                       Next var_j
                                    End If
                               
                                    If Len(Trim(var_rfc)) = 0 Then
                                       var_subimporte = Format(Round((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                                       If Len(Trim(var_subimporte)) < 14 Then
                                          For var_j = 1 + Len(Trim(var_subimporte)) To 14
                                              var_subimporte = " " + var_subimporte
                                          Next var_j
                                       End If
                                       var_iva = "-"
                                       For var_j = 1 + Len(Trim(var_iva)) To 11
                                           var_iva = " " + var_iva
                                       Next var_j
                                    Else
                                       var_subimporte = Format(Round(((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                                       If Len(Trim(var_subimporte)) < 14 Then
                                          For var_j = 1 + Len(Trim(var_subimporte)) To 14
                                              var_subimporte = " " + var_subimporte
                                          Next var_j
                                       End If
                                       var_iva = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                       If Len(Trim(var_iva)) < 14 Then
                                          For var_j = 1 + Len(Trim(var_iva)) To 14
                                              var_iva = " " + var_iva
                                          Next var_j
                                       End If
                                    End If
                                
                                    If Len(Trim(var_subimporte)) < 14 Then
                                       For var_j = 1 + Len(Trim(var_subimporte)) To 14
                                           var_subimporte = " " + var_subimporte
                                       Next var_j
                                    End If
                                    var_espacios = 131 - Len(var_cantidad_total_str)
                                    var_cantidad_total_str = Trim(var_cantidad_total_str)
                                    If Len(Trim(var_cantidad_total_str)) < 14 Then
                                       For var_j = 1 + Len(Trim(var_cantidad_total_str)) To 14
                                           var_cantidad_total_str = " " + var_cantidad_total_str
                                       Next var_j
                                    End If
                                    var_subimporte = Trim(var_subimporte)
                                    If Len(Trim(var_subimporte)) < 24 Then
                                       For var_j = 1 + Len(Trim(var_subimporte)) To 24
                                           var_subimporte = " " + var_subimporte
                                       Next var_j
                                    End If
                                         
                                    var_cantidad_total_str = var_linea + var_cantidad_total_str + "    " + var_subimporte
                                    'Print #1, Spc(var_espacios); var_cantidad_total_str; Spc(8); var_subimporte
                                    Print #1, var_cantidad_total_str
                                    var_linea = "                                                                          ESTA FACTURA SERA PAGADA EN UNA SOLA EXHIBICION                        " + var_iva
                                    Print #1, var_linea
                                    var_dia = Day(rs!dtim_Car_fecha)
                                    var_mes = Month(rs!dtim_Car_fecha)
                                    var_año = Year(rs!dtim_Car_fecha)
                                    If var_empresa = "31" Then
                                       var_linea = "                                                       " + CStr(var_dia) + "     " + CStr(var_mes)
                                    Else
                                       var_linea = "                                                             " + CStr(var_dia) + "     " + CStr(var_mes)
                                    End If
                                    If Len(var_linea) < 145 Then
                                       For var_j = 1 + Len(var_linea) To 145
                                           var_linea = var_linea + " "
                                       Next var_j
                                    End If
                                
                                    var_importe = Format(Round((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                                            
                                    If Len(Trim(var_importe)) < 14 Then
                                       For var_j = 1 + Len(Trim(var_importe)) To 14
                                           var_importe = " " + var_importe
                                       Next var_j
                                    End If
                                       
                                    'var_linea = "                                                                   ESTA FACTURA SERA PAGADA EN UNA SOLA EXHIBICION                               " + var_iva
                                    'var_linea = "                                                                                                                                                 " + var_importe
                                            
                                    var_linea = var_linea + var_importe
                                    Print #1, var_linea
                                            
                                    var_linea = var_importe
                                    If Len(Trim(var_linea)) < 20 Then
                                       For var_j = 1 + Len(Trim(var_linea)) To 20
                                           var_linea = " " + var_linea
                                       Next var_j
                                    End If
                                    var_linea = var_linea + " " + var_cantidad_letra
                                    Print #1, Spc(2); CStr(var_año); var_linea
                                
                                    var_linea = ""
                                    Print #1, ""
                                    Print #1, ""
                                    If var_empresa = "31" Then
                                       Print #1, Spc(10); Trim(IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE))
                                       Print #1, Spc(10); Trim(IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " " + IIf(IsNull(rs!VCHA_CLI_COLONIA), "", rs!VCHA_CLI_COLONIA))
                                       Print #1, Spc(10); Trim(IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre) + " " + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre))
                                    Else
                                       Print #1, Spc(5); Trim(IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE))
                                       Print #1, Spc(5); Trim(IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " " + IIf(IsNull(rs!VCHA_CLI_COLONIA), "", rs!VCHA_CLI_COLONIA))
                                       Print #1, Spc(5); Trim(IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre) + " " + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre))
                                    End If
                                    If var_empresa <> "03" Then
                                       Print #1, ""
                                       Print #1, ""
                                    Else
                                       Print #1, ""
                                       Print #1, ""
                                    End If
                                    Print #1, ""
                                    Print #1, ""
                                    Close #1
                                    If Trim(var_empresa) <> "03" Then
                                       Print #2, "copy " + App.Path + "\factura" + Trim(Str(rsaux3!inte_car_numero)) + ".txt lpt1"
                                    Else
                                       Print #2, "copy " + App.Path + "\factura" + Trim(Str(rsaux3!inte_car_numero)) + ".txt lpt1"
                                    End If
                                    'AQUI TERMINA LA FACTURA
                                  End If
                                  rs.Close
                                  rsaux3.MoveNext
                           Wend
                           Close #2
                           x = Shell(var_Archivo, vbHide)
                           rsaux3.Close
                           'Aqui se termina de imprimir la factura
                        End If
                        If rsaux3.State = 1 Then
                           rsaux3.Close
                        End If
                        rsaux3.Open "delete from TB_TEMP_FACTURA_EMBARQUES where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                        lv_detalle.ListItems.Clear
                        rsaux.Open "select * from VW_MERCANCIA_DISPONIBLE_AGENTES WHERE VCHA_AGE_AGENTE_ID = '" + txt_agente + "' and floa_Exi_cantidad > 0", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           While Not rsaux.EOF
                                 Set list_item = lv_detalle.ListItems.Add(, , rsaux!vcha_Art_Articulo_id)
                                 list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_art_nombre_Español), "", rsaux!vcha_art_nombre_Español)
                                 var_precio_1 = IIf(IsNull(rsaux!floa_dli_precio), 0, rsaux!floa_dli_precio) * 1.16
                                 list_item.SubItems(2) = Format(var_precio_1, "###,###,##0.00")
                                 list_item.SubItems(3) = Format(IIf(IsNull(rsaux!floa_dpr_desCuento), 0, rsaux!floa_dpr_desCuento), "###,###,##0.00")
                                 list_item.SubItems(4) = Format(var_precio_1 - (var_precio_1 * (IIf(IsNull(rsaux!floa_dpr_desCuento), 0, rsaux!floa_dpr_desCuento) / 100)), "###,###,##0.00")
                                 list_item.SubItems(5) = Format(IIf(IsNull(rsaux!floa_Exi_Cantidad), 0, rsaux!floa_Exi_Cantidad), "###,###,##0.00")
                                 list_item.SubItems(6) = 0
                                 list_item.SubItems(7) = 0
                                 rsaux.MoveNext
                           Wend
                        End If
                        rsaux.Close
                     End If
                  End If
               Else
                  MsgBox "Se a cancelado la impresion de las facturas", vbOKOnly, "ATENCION"
               End If
               cnn.CommitTrans
               MsgBox "Se a terminado el proceso de facturación", vbOKOnly, "ATENCION"
            Else
               MsgBox "Se a cancelado la facturación", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Se a cancelado la facturación", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "No se han seleccionado articulos para facturar", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "El agente no tiene mercancia disponible", vbo
   End If
   Else
      MsgBox "La remisión ya fue facturada", vbOKOnly, "ATENCION"
   End If
   
End Sub

Private Sub cmd_imprimir_GotFocus()
   Me.frm_factura_cancelar.Visible = False
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_agente.Enabled = True
   Me.txt_nombre_agente.Enabled = True
   Me.txt_cliente.Enabled = True
   Me.txt_nombre_cliente.Enabled = True
   Me.txt_titular.Enabled = False
   Me.txt_nombre_titular.Enabled = True
   Me.txt_remision_agente.Enabled = True
   Me.txt_remision_agente = ""
   Me.txt_importe_total = ""
   Me.lv_detalle.ListItems.Clear
   Me.txt_cliente = ""
   Me.txt_nombre_cliente = ""
   Me.txt_descuentos = ""
   Me.txt_remision = ""
   Me.txt_agente = ""
   Me.txt_nombre_agente = ""
   Me.txt_almacen = ""
   Me.txt_importe = ""
   Me.txt_importe_total = ""
   Me.txt_importe_neto = ""
   Me.txt_cantidad = ""
   Me.txt_titular = ""
   Me.txt_nombre_titular = ""
   Me.txt_cantidad_total = ""
   If var_estatus_remision = "F" Then
      var_estatus_remision = ""
      txt_agente.SetFocus
      Me.txt_almacen = ""
      Me.txt_agente = ""
      Me.txt_nombre_agente = ""
      Me.txt_almacen = ""
   Else
      var_estatus_remision = ""
      If Trim(Me.txt_agente) = "" Then
         txt_agente.SetFocus
         Me.txt_almacen = ""
      Else
         Me.txt_cliente.SetFocus
      End If
   End If
   var_primera_vez_remision = 1
   var_numero_remision = 0
End Sub

Private Sub cmd_nuevo_GotFocus()
   Me.frm_factura_cancelar.Visible = False
End Sub

Private Sub cmd_remisiones_sin_facturar_Click()
   Set reporte = appl.OpenReport(App.Path + "\rep_remisiones_sin_facturar.rpt")
   frmvistasprevias.cr.ReportSource = reporte
   For ntablas = 1 To reporte.Database.Tables.Count
       reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
   Next ntablas
   frmvistasprevias.cr.ViewReport
   frmvistasprevias.Caption = "Reporte de remisiones sin facturar"
   frmvistasprevias.Show 1
   Set reporte = Nothing
End Sub

Private Sub cmd_resumen_Click()
   
   rs.Open "select * from VW_RESUMEN_REMISION_LINEAS where inte_rem_numero = " + Me.txt_remision, cnn, adOpenDynamic, adLockOptimistic
   lv_resumen.ListItems.Clear
   txt_cantidad_total_linea = 0
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = lv_resumen.ListItems.Add(, , Trim(rs!VCHA_SUB_SUBDIVISION_ID))
            list_item.SubItems(1) = Trim(IIf(IsNull(rs!vcha_sub_nombre), "", rs!vcha_sub_nombre))
            list_item.SubItems(2) = IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)
            txt_cantidad_total_linea = Format(CDbl(txt_cantidad_total_linea) + IIf(IsNull(rs!Cantidad), 0, rs!Cantidad), "###,###,##0.00")
            rs.MoveNext
      Wend
   End If
   rs.Close
   If lv_resumen.ListItems.Count > 0 Then
      frm_resumen.Visible = True
      Me.lv_resumen.SetFocus
   Else
      MsgBox "La remisión esta vacia", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Dim var_si As Integer
   Dim var_posible As Boolean
   rsaux2.Open "SELECT * FROM VW_DOCUMENTOS_DEL_DIA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_CAR_TIPO_DOCUMENTO = 'FA' AND INTE_CAR_NUMERO = " + txt_numero + " and vcha_Ser_serie_id = '" + txt_serie + "'", cnn, adOpenDynamic, adLockOptimistic
   var_posible = False
   If Not rsaux2.EOF Then
      var_posible = True
   End If
   rsaux2.Close
   If var_posible = True Then
      var_si = MsgBox("¿Desea cancelar la factura?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         var_si = MsgBox("Confirmar la cancelacion de la factura", vbYesNo, "ATENCION")
         If var_si = 6 Then
            cnn.BeginTrans
            rs.Open "select * from tb_salidas where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_MOV_MOVIMIENTO_ID = 'FV' AND inte_Car_numero = " + txt_numero + " and vcha_ser_serie_id = '" + txt_serie + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               rsaux2.Open "update tb_encabezado_cartera set char_car_estatus = 'C' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_Car_documento =  'FA' and inte_car_numero = " + txt_numero + " and vcha_Ser_serie_id = '" + txt_serie + "'", cnn, adOpenDynamic, adLockOptimistic
               rsaux.Open "SELECT * FROM TB_ENCABEZADO_MOVIMIENTOS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_MOV_MOVIMIENTO_ID = 'FV' AND inte_emo_numero = " + CStr(rs!INTE_SAL_NUMERO), cnn, adOpenDynamic, adLockOptimistic
               var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, rs!VCHA_ALM_ALMACEN_ID, "CAFV", Now, CDbl(var_numero_folio), 0, rsaux!vcha_cli_clave_id, "", "", rs!VCHA_ALM_ALMACEN_ID, "", var_clave_usuario_global, fun_NombrePc, 0, "", "", rsaux!vcha_ESB_ESTABLECIMIENTO_id, "", rsaux!vcha_tit_titular_id, rsaux!VCHA_AGE_AGENTE_ID, IIf(IsNull(rsaux!floa_emo_descuento_1), 0, rsaux!floa_emo_descuento_1), IIf(IsNull(rsaux!floa_emo_descuento_2), 0, rsaux!floa_emo_descuento_2), 0, IIf(IsNull(rsaux!vcha_mon_moneda_id), "", rsaux!vcha_mon_moneda_id), 0)
               var_agente = rsaux!VCHA_AGE_AGENTE_ID
               rsaux.Close
               var_numero_folio = var_numero_folio_regreso
               While Not rs.EOF
                     var_cadena = "insert into tb_entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_Ent_numero, vcha_Art_articulo_id, floa_ent_cantidad, floa_ent_costo, floa_ent_precio) values "
                     var_cadena = var_cadena + "('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!VCHA_UOR_UNIDAD_ID + "','" + rs!VCHA_ALM_ALMACEN_ID + "', 'CAFV', " + CStr(var_numero_folio) + ", '" + rs!vcha_Art_Articulo_id + "', " + CStr(rs!floa_Sal_Cantidad) + ", " + CStr(rs!floa_Sal_costo) + ", " + CStr(rs!floa_Sal_precio) + ")"
                     rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     rsaux2.Open "UPDATE TB_SALIDAS SET INTE_CAR_NUMERO = 0, VCHA_SER_SERIE_ID = '', VCHA_CAR_DOCUMENTO = 'CA' WHERE VCHA_EMP_EMPRESA_ID = '" + rs!VCHA_EMP_EMPRESA_ID + "' AND VCHA_UOR_UNIDAD_ID = '" + rs!VCHA_UOR_UNIDAD_ID + "' AND VCHA_CAR_DOCUMENTO = 'FA' AND VCHA_SER_SERIE_ID = '" + rs!vcha_ser_Serie_id + "' AND INTE_CAR_NUMERO = " + txt_numero + " and vcha_art_articulo_id = '" + rs!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               rs.MoveFirst
               Set TB_ENCABEZADO_MOVIMIENTOS_M = New TB_ENCABEZADO_MOVIMIENTOS_M
               var_inserta = TB_ENCABEZADO_MOVIMIENTOS_M.Anadir(rs!VCHA_EMP_EMPRESA_ID, rs!VCHA_UOR_UNIDAD_ID, rs!VCHA_ALM_ALMACEN_ID, "CAFV", var_numero_folio, "I", Now, 1)
               rs.Close
               cnn.CommitTrans
               lv_detalle.ListItems.Clear
               rsaux.Open "select * from VW_MERCANCIA_DISPONIBLE_AGENTES WHERE VCHA_AGE_AGENTE_ID = '" + var_agente + "' and floa_Exi_cantidad > 0", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  While Not rsaux.EOF
                        Set list_item = lv_detalle.ListItems.Add(, , rsaux!vcha_Art_Articulo_id)
                        list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_art_nombre_Español), "", rsaux!vcha_art_nombre_Español)
                        var_precio = IIf(IsNull(rsaux!floa_dli_precio), 0, rsaux!floa_dli_precio) * 1.16
                        list_item.SubItems(2) = Format(var_precio, "###,###,##0.00")
                        list_item.SubItems(3) = Format(IIf(IsNull(rsaux!floa_dpr_desCuento), 0, rsaux!floa_dpr_desCuento), "###,###,##0.00")
                        list_item.SubItems(4) = Format(var_precio - (var_precio * (IIf(IsNull(rsaux!floa_dpr_desCuento), 0, rsaux!floa_dpr_desCuento) / 100)), "###,###,##0.00")
                        list_item.SubItems(5) = Format(IIf(IsNull(rsaux!floa_Exi_Cantidad), 0, rsaux!floa_Exi_Cantidad), "###,###,##0.00")
                        list_item.SubItems(6) = 0
                        list_item.SubItems(7) = 0
                        rsaux.MoveNext
                  Wend
               End If
               rsaux.Close
               
            Else
               MsgBox "No existe la factura o ya fue cancelada", vbOKOnly, "ATENCION"
            End If
         End If
      End If
   Else
      MsgBox "La factura no puede ser cancelada ya que pertenece a otro dia", vbOKOnly, "ATENCION"
   End If
   Me.frm_factura_cancelar.Visible = False
End Sub

Private Sub Command2_Click()
   Dim pError As ADODB.Error
   Dim var_actualiza As Boolean
   Dim var_inserta As Boolean
   Dim bandera_suma As Boolean
   Dim var_cantidad_1 As Variant
   Dim var_costo As Variant
   Dim var_precio_1 As Variant
   Dim var_posible_caja As Boolean
   Dim var_cantidad_posible As Variant
   Dim var_embarque_paquete As Integer
   Dim var_embarque_caja As Integer
   Dim var_estatus_caja As String
   Dim var_orden_surtido_caja As Integer
   Dim var_posible_empaque As Boolean
   Dim var_promocion_1 As Double
   Dim var_promocion_2 As Double
   Dim var_encontrado As Integer
   Dim var_canal_venta As String
   Dim var_i As Integer
   Dim var_n As Integer
   Dim var_j As Integer
   Dim var_tipo_pedido As String
   Dim var_posible As Boolean
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Set TB_TEMPORAL_SALIDAS_INSERTA = New TB_TEMPORAL_SALIDAS_INSERTA
   Set TB_TEMPORAL_SALIDAS_MODIFICA = New TB_TEMPORAL_SALIDAS_MODIFICA
   Set TB_DET_ORDEN_SURTIDO_M = New TB_DET_ORDEN_SURTIDO_M
   Set TB_DET_EMBARQUE_I = New TB_DET_EMBARQUE_I
   Set TB_DETALLE_CAJAS_M = New TB_DETALLE_CAJAS_M
   z = 0
   
   Dim var_numero_movimientos As Double
   Dim var_numero_factura_inicio As Double
   Dim var_k As Double
   Dim var_cliente As String
   Dim var_expedicion As String
   Dim var_domicilio As String
   Dim var_ciudad As String
   Dim var_agente As String
   Dim var_linea As String
   Dim var_cantidad As String
   
   Dim var_precio As Double
   Dim var_precio_str As String
   Dim var_importe As String
   Dim var_subimporte As String
   Dim var_cantidad_letra As String
   Dim var_iva As String
   Dim var_rfc As String
   Dim var_dia As String
   Dim var_mes As String
   Dim var_año As String
   Dim var_porcentaje As Double
   Dim var_Archivo As String
   Dim var_importe_descuento_1 As Double
   Dim var_importe_descuento_2 As Double
   Dim var_importe_descuento_3 As Double
   Dim var_importe_descuento_1_2 As Double
   Dim var_importe_descuento_2_2 As Double
   Dim var_importe_descuento_3_2 As Double
   Dim var_importe_descuento_1_str As String
   Dim var_importe_descuento_2_str As String
   Dim var_importe_descuento_3_str As String
   Dim var_marca_promocion As String
   Dim var_contador_promociones As Double
   Dim var_cantidad_total As Double
   Dim var_cantidad_total_str As String
   Dim var_factura_envio As Double
   Dim var_consecutivo As Double
   Dim var_x As Double
   If var_estatus_remision <> "F" Then
      cnn.CommandTimeout = 360
      var_n = lv_detalle.ListItems.Count
      var_posible = False
      For var_i = 1 To var_n
          lv_detalle.ListItems.Item(var_i).Selected = True
          If lv_detalle.selectedItem.SubItems(6) * 1 > 0 Then
             var_posible = True
          End If
      Next var_i
      If lv_detalle.ListItems.Count > 0 Then
         If var_posible = True Then
            var_si = MsgBox("¿Desea cerrar el movimiento?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               var_si = MsgBox("Confirmar el cerrado del documento", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  rs.Open "select * from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                  var_factura_inicio = rs!inte_ser_factura
                  var_serie = rs!vcha_ser_Serie_id
                  rs.Close
                  'MsgBox "Se va a imprimir la factura " + Trim(Str(var_factura_inicio)), vbOKOnly, "ATENCION"
                  'si = MsgBox("¿La impresora esta lista?", vbYesNo, "ATENCION")
                  si = 6
                  If si = 6 Then
                     var_inserta = False
                     cnn.BeginTrans
                  
                     var_estatus_remision = "F"
                     rsaux.Open "select vcha_can_canal_venta_id from tb_agentes where vcha_age_agente_id = '" + txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
                     var_canal_venta = IIf(IsNull(rsaux!vcha_can_canal_venta_id), "", rsaux!vcha_can_canal_venta_id)
                     rsaux.Close
                     var_almacen_origen = var_clave_almacen
                     rs.Open "select * from vw_clientes where vcha_cli_clave_id = '" + txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        var_clave_moneda = rs!vcha_mon_moneda_id
                        var_clave_titular = rs!vcha_tit_titular_id
                        var_canal_venta = rs!vcha_can_canal_venta_id
                     End If
                     rs.Close
                     rs.Open "select * from tb_Detalle_Establecimientos where vcha_cli_clave_id = '" + txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        var_clave_establecimiento = rs!vcha_ESB_ESTABLECIMIENTO_id
                     End If
                     rs.Close
                     var_clave_movimiento = "FV"
                     var_numero_folio = 0
                     var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, Now, CDbl(var_numero_folio), 0, txt_cliente, "", var_almacen_origen, "", "", var_clave_usuario_global, fun_NombrePc, 0, "", "", var_clave_establecimiento, "", var_clave_titular, txt_agente, var_descuento_1, var_descuento_2, 0, var_clave_moneda, 0)
                     var_numero_folio = var_numero_folio_regreso
                     rsaux.Open "update tb_encabezado_movimientos set  VCHA_CAN_CANAL_VENTA_ID = '" + var_canal_venta + "' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen_origen + "' and inte_emo_numero = " + CStr(var_numero_folio) + " and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
                     var_inserta = False
                     txt_folio = var_numero_folio
                     var_primera_vez = False
                     var_consecutivo = 0
                     For var_i = 1 To var_n
                         lv_detalle.ListItems.Item(var_i).Selected = True
                         If lv_detalle.selectedItem.SubItems(6) * 1 > 0 Then
                            var_consecutivo = var_consecutivo + 1
                            var_promocion_1 = 0
                            
                            If lv_detalle.selectedItem.SubItems(3) * 1 > 0 Then
                               var_promocion_1 = lv_detalle.selectedItem.SubItems(3) * 1
                            End If
                            var_precio_1 = 0
                         
                            If var_empresa = "31" And (Me.lv_detalle.selectedItem = "S1005" Or Me.lv_detalle.selectedItem = "S1003") Then
                               var_precio_1 = (lv_detalle.selectedItem.SubItems(4) * 1) / 1.16
                            Else
                               If lv_detalle.selectedItem.SubItems(4) * 1 > 0 Then
                                  var_precio_1 = (lv_detalle.selectedItem.SubItems(4) * 1) / 1.16
                               End If
                            End If
                         
                            rsaux2.Open "INSERT INTO tb_temporal_salidas (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_SAL_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_SAL_CANTIDAD, FLOA_SAL_COSTO, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO, FLOA_SAL_PROMOCION_1, FLOA_SAL_PROMOCION_2, CHAR_PED_TIPO, INTE_SAL_CONSECUTIVO) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", '" + Trim(lv_detalle.selectedItem) + "', " + CStr(CDbl(lv_detalle.selectedItem.SubItems(6))) + ", 0, " + CStr(var_precio_1) + ", 0, " + CStr(var_promocion_1) + ", 0, 'M', " + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                            If Me.lv_detalle.selectedItem = "S1005" Then
                               rsaux2.Open "UPDATE TB_REMISIONES SET INTE_EMO_NUMERO = " + CStr(var_numero_folio) + ", FLOA_DLI_PRECIO = " + CStr(CDbl(Me.lv_detalle.selectedItem.SubItems(2))) + ", CHAR_REM_ESTATUS = 'F', FLOA_REM_PRECIO = " + CStr(var_precio_1) + ", FLOA_REM_PROMOCION_1 = " + CStr(CDbl(lv_detalle.selectedItem.SubItems(3))) + ", FLOA_REM_DESCUENTO_1 = 0, FLOA_REM_DESCUENTO_2 = 0 WHERE INTE_REM_NUMERO = " + CStr(var_numero_remision) + " AND VCHA_aRT_ARTICULO_ID = '" + lv_detalle.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                            Else
                               rsaux2.Open "UPDATE TB_REMISIONES SET INTE_EMO_NUMERO = " + CStr(var_numero_folio) + ", FLOA_DLI_PRECIO = " + CStr(CDbl(Me.lv_detalle.selectedItem.SubItems(2))) + ", CHAR_REM_ESTATUS = 'F', FLOA_REM_PRECIO = " + CStr(var_precio_1) + ", FLOA_REM_PROMOCION_1 = " + CStr(CDbl(lv_detalle.selectedItem.SubItems(3))) + ", FLOA_REM_DESCUENTO_1 = " + CStr(var_descuento_1) + ", FLOA_REM_DESCUENTO_2 = " + CStr(var_descuento_2) + " WHERE INTE_REM_NUMERO = " + CStr(var_numero_remision) + " AND VCHA_aRT_ARTICULO_ID = '" + lv_detalle.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                            End If
                            var_tipo_Cambio = 1
                            var_catalogo_1 = ""
                            var_catalogo_2 = ""
                            var_año_catalogo = 0
                            var_mes_catalogo = 0
                            var_si_surtir_catalogo = 0
                         End If
                     Next var_i
                     If rsaux4.State = 1 Then
                        rsaux4.Close
                     End If
                     Text1 = "exec SP_INSERTA_MOVIMIENTOS_SALIDA_EMBARQUES '" + var_empresa + "','" + var_unidad_organizacional + "','" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", " + CStr(var_tipo_Cambio) + ",'" + var_catalogo_1 + "','" + var_catalogo_2 + "','" + var_clave_titular + "','" + txt_cliente + "'," + CStr(var_año_catalogo) + "," + CStr(var_mes_catalogo) + "," + CStr(var_si_surtir_catalogo)
                     rsaux4.Open "exec SP_INSERTA_MOVIMIENTOS_SALIDA_EMBARQUES '" + var_empresa + "','" + var_unidad_organizacional + "','" + var_almacen_origen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", " + CStr(var_tipo_Cambio) + ",'" + var_catalogo_1 + "','" + var_catalogo_2 + "','" + var_clave_titular + "','" + txt_cliente + "'," + CStr(var_año_catalogo) + "," + CStr(var_mes_catalogo) + "," + CStr(var_si_surtir_catalogo), cnn, adOpenDynamic, adLockOptimistic
                     'rsaux5.Open "select sum(floa_Sal_Cantidad) from tb_salidas where vcha_mov_movimiento_id = 'fv' and inte_sal_numero = 105", cnn, adOpenDynamic, adLockOptimistic
                     'MsgBox CStr(rsaux5(0).Value)
                     'rsaux5.Close
                     
                     rs.Open "select * from vw_maximo_embarque where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                     If rs.EOF Then
                        var_numero_embarque = 1
                     Else
                        var_numero_embarque = rs!maximo_embarque + 1
                     End If
                     rs.Close
                     Set TB_ENC_EMBARQUE_I = New TB_ENC_EMBARQUE_I
                     ok = False
                     rs.Open "insert into tb_encabezado_embarques (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, INTE_EMB_EMBARQUE, INTE_JAU_JAULA_ID, VCHA_VEH_VEHICULO_ID, VCHA_AGE_AGENTE_ID, DTIM_EMB_FECHA_INICIO, DTIM_EMB_FECHA_FINAL, CHAR_EMB_ESTATUS, VCHA_CHO_CHOFER_ID, FLOA_EMB_CUBICAJE, CHAR_EMB_TIPO, INTE_EMB_BLOQUEADO, VCHA_EMB_BLOQUEADO_POR, VCHA_AUD_MAQUINA, VCHA_AUD_USUARIO) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', " + CStr(var_numero_embarque) + ", 0, '', '" + txt_agente + "', getdate(), '', '', '', 0,'',0, '','" + fun_NombrePc + "','" + var_clave_usuario_global + "')", cnn, adOpenDynamic, adLockOptimistic
                     var_inserta = TB_DET_EMBARQUE_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_numero_embarque, var_clave_movimiento, var_numero_folio, "")
                     txt_numero_embarque = CStr(var_numero_embarque)
                     var_estatus_embarque = "I"
               
                     If Trim(txt_numero_embarque) <> "" Then
                        'Sirve para validar que no vaya mercancia con cantidad en NULL
                        Cadena = "SELECT     dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID, "
                        Cadena = Cadena + " dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO, dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID,"
                        Cadena = Cadena + " dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID , dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD"
                        Cadena = Cadena + " FROM         dbo.TB_DETALLE_EMBARQUES INNER JOIN"
                        Cadena = Cadena + " dbo.TB_SALIDAS ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND"
                        Cadena = Cadena + " dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID AND"
                        Cadena = Cadena + " dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO = dbo.TB_SALIDAS.INTE_SAL_NUMERO AND"
                        Cadena = Cadena + " dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID"
                        Cadena = Cadena + " WHERE     (dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD IS NULL) AND (dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = " + CStr(var_numero_embarque) + ") AND"
                        Cadena = Cadena + " (dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "')"
                        If rsaux4.State = 1 Then
                           rsaux4.Close
                        End If
                        rsaux4.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux4.EOF Then
                           rsaux4.Close
                           MsgBox "El movimiento tiene cantidad en NULL", vbOKOnly, "ATENCION"
                        Else
                           fecha_inicio = CStr(Now)
                           Set TB_ENC_EMBARQUE_M = New TB_ENC_EMBARQUE_M
                        ok = TB_ENC_EMBARQUE_M.Anadir(var_empresa, var_unidad_organizacional, CDbl(txt_numero_embarque), "I")
                        rsaux.Open "update tb_encabezado_embarques set dtim_emb_fecha_final =  getdate() where vcha_Emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque, cnn, adOpenDynamic, adLockOptimistic
                        rsaux.Open "select * from VW_MERCANCIA_DISPONIBLE_AGENTES WHERE VCHA_AGE_AGENTE_ID = '" + txt_agente + "' and floa_Exi_cantidad > 0", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           While Not rsaux.EOF
                                 Set list_item = lv_detalle.ListItems.Add(, , rsaux!vcha_Art_Articulo_id)
                                 list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_art_nombre_Español), "", rsaux!vcha_art_nombre_Español)
                                 var_precio_1 = IIf(IsNull(rsaux!floa_dli_precio), 0, rsaux!floa_dli_precio) * 1.16
                                 list_item.SubItems(2) = Format(var_precio_1, "###,###,##0.00")
                                 list_item.SubItems(3) = Format(IIf(IsNull(rsaux!floa_dpr_desCuento), 0, rsaux!floa_dpr_desCuento), "###,###,##0.00")
                                 list_item.SubItems(4) = Format(var_precio_1 - (var_precio_1 * (IIf(IsNull(rsaux!floa_dpr_desCuento), 0, rsaux!floa_dpr_desCuento) / 100)), "###,###,##0.00")
                                 list_item.SubItems(5) = Format(IIf(IsNull(rsaux!floa_Exi_Cantidad), 0, rsaux!floa_Exi_Cantidad), "###,###,##0.00")
                                 list_item.SubItems(6) = 0
                                 list_item.SubItems(7) = 0
                                 rsaux.MoveNext
                           Wend
                        End If
                        rsaux.Close
                     End If
                  End If
               Else
                  MsgBox "Se a cancelado la impresion de las facturas", vbOKOnly, "ATENCION"
               End If
               cnn.CommitTrans
               MsgBox "Se a cerrado el movimiento", vbOKOnly, "ATENCION"
            Else
               MsgBox "Se a cancelado el cerrado del movimiento", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Se a cancelado el cerrado del movimiento", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "No se han seleccionado articulos", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "El agente no tiene mercancia disponible", vbOKOnly, "ATENCION"
   End If
   Else
      MsgBox "La remisión ya fue facturada", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 78 Then
      cmd_nuevo_Click
   End If
   If Shift = 4 And KeyCode = 66 Then
      cmd_buscar_remision_Click
   End If
   If Shift = 4 And KeyCode = 73 Then
      cmd_imprimir_Click
   End If
   If Shift = 4 And KeyCode = 67 Then
      cmd_cancelar_Click
   End If
End Sub

Private Sub Form_Load()
   If var_empresa = "31" Then
      Me.cmd_aplicacion_anticipos.Enabled = True
   Else
      Me.cmd_aplicacion_anticipos.Enabled = False
   End If
   If var_empresa = "18" Or var_empresa = "31" Then
      Me.cmd_imprimir.Enabled = False
      Me.Command2.Enabled = True
   Else
      Me.cmd_imprimir.Enabled = True
      Me.Command2.Enabled = False
   End If
   Top = 0
   Left = 0
   frm_lista.Visible = False
   frm_cantidad.Visible = False
   frm_eliminar.Visible = False
   txt_cantidad_total = Format(0, "###,###,##0.00")
   txt_importe_total = Format(0, "###,###,##0.00")
   Me.frm_factura_cancelar.Visible = False
   Me.frm_numero_nota.Visible = False
   var_numero_remision = 0
   var_primera_vez_remision = 1
   frm_resumen.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_reporte_valuacion_devoluciones)
End Sub

Private Sub lv_detalle_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_detalle, ColumnHeader)
End Sub

Private Sub lv_detalle_GotFocus()
   Me.frm_factura_cancelar.Visible = False
End Sub

Private Sub lv_detalle_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 115 Then
      If var_estatus_remision <> "F" Then
         If Trim(txt_cliente) <> "" Then
            If Trim(Me.txt_remision_agente) <> "" Then
               If Me.lv_detalle.selectedItem = "S1005" Then
                  MsgBox "No se puede manipular la cantidad del anticipo", vbOKOnly, "ATENCION"
               Else
                  If lv_detalle.ListItems.Count > 0 Then
                     frm_cantidad.Visible = True
                     txt_cantidad = ""
                     txt_cantidad.SetFocus
                  End If
               End If
            Else
               MsgBox "Debe de indicar la remisión del agente", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Debe de seleccionar un cliente", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "La remisión ya no puede ser modificada", vbOKOnly, "ATENCION"
      End If
   End If
   If KeyCode = 114 Then
      If var_estatus_remision <> "F" Then
         If Trim(txt_cliente) <> "" Then
            If lv_detalle.ListItems.Count > 0 Then
               If Me.lv_detalle.selectedItem = "S1005" Then
                  var_si = MsgBox("¿Se va a eliminar el anticipo", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     rs.Open "SELECT * FROM TB_REMISIONES WHERE  inte_rem_numero = " + CStr(var_numero_remision) + " and vcha_art_articulo_id = 'S1005'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        var_consecutivo = IIf(IsNull(rs!INTE_ANT_CONSECUTIVO), 0, rs!INTE_ANT_CONSECUTIVO)
                        txt_cantidad_total = Format(CDbl(txt_cantidad_total) - CDbl(Me.lv_detalle.selectedItem.SubItems(6)), "###,###,##0.00")
                        If Me.lv_detalle.selectedItem = "S1005" Then
                           x = CDbl(1) / (1 - (var_descuento_1 / 100))
                           x = CDbl(x) / (1 - (var_descuento_2 / 100))
                           z = (x * 1) * CDbl(lv_detalle.selectedItem.SubItems(6) * 1)
                           z = Format(CDbl(txt_importe_total) + CDbl(z), "###,###,##0.00")
                           txt_importe_total = Format(z, "###,###,##0.00")
                        Else
                           txt_importe_total = Format(CDbl(txt_importe_total) - (CDbl(lv_detalle.selectedItem.SubItems(7) * 1)), "###,###,##0.00")
                        End If
                        'rsaux.Open "update tb_anticipos set inte_ant_Cargado = 0 where inte_Ant_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                        rsaux.Open "delete from tb_remisiones WHERE  inte_rem_numero = " + CStr(var_numero_remision) + " and vcha_art_articulo_id = 'S1005'", cnn, adOpenDynamic, adLockOptimistic
                        If Trim(Me.txt_remision) <> "" Then
                           rsaux1.Open "select * from tb_aplicacion_Anticipos where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'FV' and inte_emo_numero = " + CStr(Me.txt_remision), cnn, adOpenDynamic, adLockOptimistic
                           While Not rsaux1.EOF
                                 rsaux.Open "update tb_anticipos set floa_ant_aplicado = floa_ant_aplicado - " + CStr(rsaux1!floa_ant_importe) + " where inte_ant_consecutivo = " + CStr(rsaux1!INTE_ANT_CONSECUTIVO), cnn, adOpenDynamic, adLockOptimistic
                                 rsaux1.MoveNext
                           Wend
                           rsaux1.Close
                           rsaux1.Open "delete from tb_aplicacion_Anticipos where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = 'FV' and inte_emo_numero = " + CStr(Me.txt_remision), cnn, adOpenDynamic, adLockOptimistic
                        End If
                        
                        lv_detalle.ListItems.Remove (lv_detalle.selectedItem.Index)
                     End If
                     
                     rs.Close
                  End If
               Else
                  frm_eliminar.Visible = True
                  txt_cantidad_eliminar = ""
                  txt_cantidad_eliminar.SetFocus
               End If
            End If
         Else
            MsgBox "Debe de seleccionar un cliente", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "La remisión ya no puede ser modificada", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_tipo_lista = 1 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_agente = lv_lista.selectedItem
            txt_nombre_agente = lv_lista.selectedItem.SubItems(1)
         Else
            txt_agente = ""
            txt_nombre_agente = ""
         End If
         Me.txt_agente.Enabled = True
         txt_agente.SetFocus
      End If
      If var_tipo_lista = 2 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_cliente = lv_lista.selectedItem
            txt_nombre_cliente = lv_lista.selectedItem.SubItems(1)
         Else
            txt_cliente = ""
            txt_nombre_cliente = ""
         End If
         Me.txt_cliente.Enabled = True
         txt_cliente.SetFocus
      End If
      If var_tipo_lista = 3 Then
         If lv_lista.ListItems.Count > 0 Then
            Me.txt_titular = lv_lista.selectedItem
            Me.txt_nombre_titular = lv_lista.selectedItem.SubItems(1)
         Else
            Me.txt_titular = ""
            Me.txt_nombre_titular = ""
         End If
         Me.txt_titular.Enabled = True
         txt_titular.SetFocus
      End If
      
   End If
   If KeyAscii = 27 Then
      If var_tipo_lista = 1 Then
         txt_agente.SetFocus
      End If
      If var_tipo_lista = 2 Then
         txt_cliente.SetFocus
      End If
      If var_tipo_lista = 3 Then
         Me.txt_titular.SetFocus
      End If
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub lv_resumen_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_resumen, ColumnHeader)
End Sub

Private Sub lv_resumen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frm_resumen.Visible = False
   End If
End Sub

Private Sub lv_resumen_LostFocus()
   frm_resumen.Visible = False
End Sub

Private Sub txt_agente_GotFocus()
   Me.frm_factura_cancelar.Visible = False
End Sub

Private Sub txt_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_agentes where vcha_emp_empresa_id = '" + var_empresa + "' and (vcha_alm_almacen_id <> '') order by vcha_age_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Agentes"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 7 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_agente_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_agente_LostFocus()
   Dim list_item As ListItem
   If Trim(Me.txt_agente) = "" Then
      Me.txt_titular = ""
      Me.txt_nombre_titular = ""
      Me.txt_nombre_agente = ""
      Me.txt_cliente = ""
      Me.txt_nombre_agente = ""
      Me.txt_descuentos = ""
      Me.txt_remision = ""
      Me.txt_folio = ""
      Me.lv_detalle.ListItems.Clear
   Else
      Me.txt_titular = ""
      Me.txt_nombre_titular = ""
      Me.txt_nombre_agente = ""
      Me.txt_cliente = ""
      Me.txt_nombre_agente = ""
      Me.txt_descuentos = ""
      Me.txt_remision = ""
      Me.txt_folio = ""
      Me.lv_detalle.ListItems.Clear
      rs.Open "select * from tb_agentes where vcha_age_agente_id = '" + txt_agente + "' and VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_agente = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
         var_clave_almacen = IIf(IsNull(rs!VCHA_ALM_ALMACEN_ID), "", rs!VCHA_ALM_ALMACEN_ID)
         Me.txt_almacen = var_clave_almacen
         If Trim(var_clave_almacen) <> "" Then
            rsaux.Open "select * from VW_MERCANCIA_DISPONIBLE_AGENTES WHERE VCHA_AGE_AGENTE_ID = '" + txt_agente + "' and floa_Exi_cantidad > 0 and vcha_Art_Articulo_id <> 'S1005' order by vcha_Art_nombre_Español", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               While Not rsaux.EOF
                     Set list_item = lv_detalle.ListItems.Add(, , rsaux!vcha_Art_Articulo_id)
                     list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_art_nombre_Español), "", rsaux!vcha_art_nombre_Español)
                     var_precio = Round(IIf(IsNull(rsaux!floa_dli_precio), 0, rsaux!floa_dli_precio) * 1.16, 2)
                     list_item.SubItems(2) = Format(var_precio, "###,###,##0.00")
                     list_item.SubItems(3) = Format(IIf(IsNull(rsaux!floa_dpr_desCuento), 0, rsaux!floa_dpr_desCuento), "###,###,##0.00")
                     list_item.SubItems(4) = Format(var_precio - (var_precio * (IIf(IsNull(rsaux!floa_dpr_desCuento), 0, rsaux!floa_dpr_desCuento) / 100)), "###,###,##0.00")
                     list_item.SubItems(5) = Format(IIf(IsNull(rsaux!floa_Exi_Cantidad), 0, rsaux!floa_Exi_Cantidad), "###,###,##0.00")
                     list_item.SubItems(6) = 0
                     list_item.SubItems(7) = 0
                     rsaux.MoveNext
               Wend
            End If
            rsaux.Close
            Me.txt_agente.Enabled = False
            Me.txt_nombre_agente.Enabled = False
            If Me.lv_detalle.ListItems.Count > 22 Then
               lv_detalle.ColumnHeaders(3).Width = 1099.96
            Else
               lv_detalle.ColumnHeaders(3).Width = 1299.96
            End If
            Me.txt_titular.Enabled = True
         End If
      Else
         MsgBox "Clave de agente incorrecta", vbOKOnly, "ATENCION"
         Me.txt_agente = ""
         Me.txt_nombre_agente = ""
         Me.txt_cliente = ""
         Me.txt_nombre_agente = ""
         Me.txt_descuentos = ""
         Me.txt_almacen = ""
         Me.txt_remision = ""
      End If
      rs.Close
      Me.txt_cliente = ""
      Me.txt_nombre_cliente = ""
   End If
End Sub

Private Sub txt_cantidad_eliminar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(txt_cantidad_eliminar) Then
         If CDbl(txt_cantidad_eliminar) <= lv_detalle.selectedItem.SubItems(6) * 1 Then
            rsaux.Open "update tb_Remisiones set floa_rem_cantidad = isnull(floa_rem_cantidad,0) - " + CStr(CDbl(txt_cantidad_eliminar)) + " where inte_rem_numero = " + CStr(var_numero_remision) + " and vcha_art_articulo_id = '" + Me.lv_detalle.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
            lv_detalle.selectedItem.SubItems(6) = Format(CDbl(lv_detalle.selectedItem.SubItems(6) * 1) - CDbl(txt_cantidad_eliminar), "###,###,##0.00")
            lv_detalle.selectedItem.SubItems(7) = Format((lv_detalle.selectedItem.SubItems(6) * 1) * (lv_detalle.selectedItem.SubItems(4) * 1), "###,###,##0.00")
            txt_cantidad_total = Format(CDbl(txt_cantidad_total) - CDbl(txt_cantidad_eliminar), "###,###,##0.00")
            txt_importe_total = Format(CDbl(txt_importe_total) - (CDbl(lv_detalle.selectedItem.SubItems(4) * 1) * CDbl(txt_cantidad_eliminar)), "###,###,##0.00")
            lv_detalle.SetFocus
         Else
            MsgBox "Cantidad a eliminar incorrecta", vbOKOnly, "ATENCION"
         End If
      End If
   End If
   If KeyAscii = 27 Then
      If lv_detalle.ListItems.Count > 0 Then
         lv_detalle.SetFocus
      End If
      frm_eliminar.Visible = False
   End If
End Sub

Private Sub txt_cantidad_eliminar_LostFocus()
   Me.frm_eliminar.Visible = False
End Sub

Private Sub txt_cantidad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(txt_cantidad) Then
         If (lv_detalle.selectedItem.SubItems(6) * 1) + CDbl(txt_cantidad) > lv_detalle.selectedItem.SubItems(5) * 1 Then
            MsgBox "Cantidad supera a la disponible", vbOKOnly, "ATENCION"
         Else
            If var_primera_vez_remision = 1 Then
               cnn.BeginTrans
               rs.Open "select max(inte_rem_numero) from tb_remisiones", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_numero_remision = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
               Else
                  var_numero_remision = 1
               End If
               rs.Close
               Me.txt_remision = var_numero_remision
               rs.Open "insert into tb_Remisiones (vcha_emp_empresa_id, vcha_alm_almacen_id,vcha_age_agente_id,vcha_cli_clave_id, inte_rem_numero,floa_rem_cantidad) values ('" + var_empresa + "', '" + Me.txt_almacen + "','" + Me.txt_agente + "', '" + Me.txt_cliente + "'," + CStr(var_numero_remision) + ",0 )", cnn, adOpenDynamic, adLockOptimistic
               cnn.CommitTrans
               var_primera_vez_remision = 0
            End If
            rs.Open "select * from tb_remisiones where inte_rem_numero = " + CStr(var_numero_remision) + " and vcha_art_articulo_id = '" + Me.lv_detalle.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               rsaux.Open "update tb_Remisiones set floa_rem_cantidad = isnull(floa_rem_cantidad,0) + " + CStr(CDbl(txt_cantidad)) + " where inte_rem_numero = " + CStr(var_numero_remision) + " and vcha_art_articulo_id = '" + Me.lv_detalle.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
            Else
               rsaux.Open "insert into tb_Remisiones (vcha_emp_empresa_id, vcha_alm_almacen_id,vcha_age_agente_id,vcha_cli_clave_id, inte_rem_numero,floa_rem_cantidad, vcha_Art_articulo_id, VCHA_REM_REMISION_AGENTE) values ('" + var_empresa + "', '" + Me.txt_almacen + "','" + Me.txt_agente + "', '" + Me.txt_cliente + "'," + CStr(var_numero_remision) + "," + CStr(CDbl(txt_cantidad)) + ",'" + Me.lv_detalle.selectedItem + "','" + Me.txt_remision_agente + "' )", cnn, adOpenDynamic, adLockOptimistic
            End If
            rs.Close
            lv_detalle.selectedItem.SubItems(6) = Format((lv_detalle.selectedItem.SubItems(6) * 1) + CDbl(txt_cantidad), "###,###,##0.00")
            lv_detalle.selectedItem.SubItems(7) = Format((lv_detalle.selectedItem.SubItems(6) * 1) * (lv_detalle.selectedItem.SubItems(4) * 1), "###,###,##0.00")
            If Trim(Me.txt_cantidad_total) = "" Then
               Me.txt_cantidad_total = "0"
            End If
            If Trim(Me.txt_importe_total) = "" Then
               Me.txt_importe_total = "0"
            End If
            txt_cantidad_total = Format(CDbl(txt_cantidad_total) + CDbl(txt_cantidad), "###,###,##0.00")
            rsaux.Open "select * from tb_descuentos_promocion_clientes where vcha_art_articulo_id = '" + Me.lv_detalle.selectedItem + "' and vcha_cli_clave_id = '" + Me.txt_cliente + "' and dtim_dpr_fecha_inicio <= getdate() and dtim_dpr_fecha_fin >= getdate()", cnn, adOpenDynamic, adLockOptimistic
            var_marca = ""
            If Not rsaux.EOF Then
               var_marca = IIf(IsNull(rsaux!char_dpr_marca), "", rsaux!char_dpr_marca)
            End If
            rsaux.Close
            If var_marca = "" Then
               If lv_detalle.selectedItem.SubItems(3) * 1 > var_descuento_1 Then
                  x = CDbl(lv_detalle.selectedItem.SubItems(4) * 1) / (1 - (var_descuento_1 / 100))
                  txt_importe_total = Format(CDbl(txt_importe_total) + ((x * 1) * CDbl(txt_cantidad)), "###,###,##0.00")
               Else
                  txt_importe_total = Format(CDbl(txt_importe_total) + (CDbl(lv_detalle.selectedItem.SubItems(4) * 1) * CDbl(txt_cantidad)), "###,###,##0.00")
               End If
            Else
               txt_importe_total = Format(CDbl(txt_importe_total) + (CDbl(lv_detalle.selectedItem.SubItems(4) * 1) * CDbl(txt_cantidad)), "###,###,##0.00")
            End If
         End If
         lv_detalle.SetFocus
      Else
         MsgBox "Cantidad incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
   If KeyAscii = 27 Then
      If lv_detalle.ListItems.Count > 0 Then
         lv_detalle.SetFocus
      End If
      frm_cantidad.Visible = False
   End If
End Sub

Private Sub txt_cantidad_LostFocus()
   frm_cantidad.Visible = False
End Sub

Private Sub txt_cantidad_total_GotFocus()
   Me.frm_factura_cancelar.Visible = False
End Sub

Private Sub txt_cliente_GotFocus()
   Me.frm_factura_cancelar.Visible = False
End Sub

Private Sub txt_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_clientes where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_age_agente_id = '" + Me.txt_agente + "' and vcha_tit_titular_id = '" + Me.txt_titular + "' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Clientes"
      var_tipo_lista = 2
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 7 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_cliente_LostFocus()
   If Trim(txt_cliente) <> "" Then
      rs.Open "SELECT * FROM vw_clientes WHERE VCHA_CLI_CLAVE_ID = '" + txt_cliente + "' AND VCHA_AGE_AGENTE_ID = '" + txt_agente + "' and vcha_tit_titular_id = '" + Me.txt_titular + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         
         Me.txt_cliente.Enabled = False
         Me.txt_nombre_cliente.Enabled = False
         txt_nombre_cliente = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
         var_descuento_1 = IIf(IsNull(rs!floa_gac_Descuento_1), 0, rs!floa_gac_Descuento_1)
         var_descuento_2 = IIf(IsNull(rs!FLOA_GAC_DESCUENTO_2), 0, rs!FLOA_GAC_DESCUENTO_2)
         txt_descuentos = CStr(var_descuento_1) + "% + " + CStr(var_descuento_2) + "%"
         txt_importe_total = 0
         If Me.lv_detalle.ListItems.Count > 0 Then
            For var_j = 1 To lv_detalle.ListItems.Count
                lv_detalle.ListItems.Item(var_j).Selected = True
                rsaux.Open "select * from TB_DESCUENTOS_PROMOCION_CLIENTES where vcha_cli_clave_id = '" + Me.txt_cliente + "' and vcha_Art_articulo_id = '" + Me.lv_detalle.selectedItem + "' and dtim_dpr_fecha_inicio <= getdate() and dtim_dpr_fecha_fin >= getdate()", cnn, adOpenDynamic, adLockOptimistic
                If Not rsaux.EOF Then
                   var_oferta = IIf(IsNull(rsaux!floa_dpr_desCuento), 0, rsaux!floa_dpr_desCuento)
                   var_marca = IIf(IsNull(rsaux!char_dpr_marca), "", rsaux!char_dpr_marca)
                   Me.lv_detalle.selectedItem.SubItems(3) = var_oferta
                   If var_marca <> "*" Then
                      If var_descuento_1 = 100 Then
                         Me.lv_detalle.selectedItem.SubItems(2) = 0
                      Else
                         Me.lv_detalle.selectedItem.SubItems(2) = Format(lv_detalle.selectedItem.SubItems(2) / (1 - (var_descuento_1 / 100)), "###,###,##0.00")
                      End If
                   End If
                   Me.lv_detalle.selectedItem.SubItems(4) = Round(CDbl(lv_detalle.selectedItem.SubItems(2)) - (CDbl(lv_detalle.selectedItem.SubItems(2)) * (var_oferta / 100)), 2)
                   Me.lv_detalle.selectedItem.SubItems(7) = Round(CDbl(lv_detalle.selectedItem.SubItems(4)) * CDbl(lv_detalle.selectedItem.SubItems(6)), 2)
                   txt_importe_total = Format(CDbl(txt_importe_total) + (CDbl(lv_detalle.selectedItem.SubItems(4) * 1) * CDbl(lv_detalle.selectedItem.SubItems(6))), "###,###,##0.00")
                End If
                rsaux.Close
            Next var_j
        End If
      Else
         MsgBox "Clave de cliente incorrecta", vbOKOnly, "ATENCION"
         Me.txt_cliente = ""
         Me.txt_descuentos = ""
         Me.txt_nombre_cliente = ""
      End If
      rs.Close
   Else
      txt_nombre_cliente = ""
   End If
End Sub

Private Sub txt_descuentos_GotFocus()
   Me.frm_factura_cancelar.Visible = False
End Sub

Private Sub txt_fecha_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      If KeyAscii <> 27 Then
         KeyAscii = 0
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_factura_cancelar.Visible = False
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_folio_GotFocus()
   Me.frm_factura_cancelar.Visible = False
End Sub

Private Sub txt_importe_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      If KeyAscii <> 27 Then
         KeyAscii = 0
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_factura_cancelar.Visible = False
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_importe_total_Change()
   If IsNumeric(Me.txt_importe_total) Then
      If Me.lv_detalle.ListItems.Count > 0 Then
         var_descuento_total_1 = (100 - var_descuento_1) / 100
         Me.txt_importe_neto = Me.txt_importe_total * var_descuento_total_1
         var_descuento_total_2 = (100 - var_descuento_2) / 100
         Me.txt_importe_neto = Format(Me.txt_importe_neto * var_descuento_total_2, "###,###,##0.00")
      Else
         txt_importe_neto = 0
         txt_importe_total = 0
      End If
   End If
End Sub

Private Sub txt_importe_total_GotFocus()
   Me.frm_factura_cancelar.Visible = False
End Sub

Private Sub txt_nombre_agente_cancelar_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      If KeyAscii <> 27 Then
         KeyAscii = 0
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_factura_cancelar.Visible = False
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_agente_GotFocus()
   Me.frm_factura_cancelar.Visible = False
End Sub

Private Sub txt_nombre_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_agentes where vcha_emp_empresa_id = '" + var_empresa + "' order by vcha_age_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Agentes"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 7 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_agente_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_cliente_cancelar_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      If KeyAscii <> 27 Then
         KeyAscii = 0
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_factura_cancelar.Visible = False
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_cliente_GotFocus()
   Me.frm_factura_cancelar.Visible = False
End Sub

Private Sub txt_nombre_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_clientes where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_age_agente_id = '" + txt_agente + "' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Clientes"
      var_tipo_lista = 2
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 7 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_cliente_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_titular_GotFocus()
   Me.frm_factura_cancelar.Visible = False
End Sub

Private Sub txt_nombre_titular_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct vcha_tit_titular_id, vcha_tit_nombre from vw_clientes where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_age_agente_id = '" + txt_agente + "' order by vcha_tit_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_tit_titular_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Titulares"
      var_tipo_lista = 3
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 7 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_titular_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frm_factura_cancelar.Visible = False
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_numero_LostFocus()
   If IsNumeric(txt_numero) Then
      If Trim(Me.txt_serie) <> "" Then
         rs.Open "select * from tb_encabezado_cartera where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_Car_documento = 'FA' and vcha_ser_serie_id = '" + txt_serie + "' and inte_car_numero = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            rsaux.Open "SELECT * from TB_AGENTES WHERE VCHA_AGE_AGENTE_ID = '" + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               Me.txt_nombre_agente_cancelar = IIf(IsNull(rsaux!VCHA_AGE_NOMBRE), "", rsaux!VCHA_AGE_NOMBRE)
            End If
            rsaux.Close
            rsaux.Open "SELECT * FROM TB_CLIENTES WHERE VCHA_CLI_CLAVE_ID = '" + IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               Me.txt_nombre_cliente_cancelar = IIf(IsNull(rsaux!VCHA_CLI_NOMBRE), "", rsaux!VCHA_CLI_NOMBRE)
            End If
            rsaux.Close
            rsaux.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + rs!vcha_mon_moneda_id + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               lbl_moneda = IIf(IsNull(rsaux!vcha_mon_nombre_plural), "", rsaux!vcha_mon_nombre_plural)
            End If
            rsaux.Close
            Me.txt_fecha = Format(rs!dtim_Car_fecha, "Short Date")
            Me.txt_importe = Format((rs!floa_Car_importe_neto / rs!floa_car_tipo_cambio), "###,###,##0.00")
         Else
            MsgBox "La factura no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         MsgBox "Serie incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Número de factura incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub txt_numero_nota_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 27 Then
      Me.frm_numero_nota.Visible = False
   End If
   If KeyAscii = 13 Then
      If IsNumeric(txt_numero_nota) Then
         cnn.CommandTimeout = 360
         rsaux5.Open "SELECT * FROM TB_REMISIONES WHERE INTE_REM_NUMERO = " + Me.txt_numero_nota + " AND VCHA_ART_ARTICULO_ID IS NOT NULL", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux5.EOF Then
            txt_importe_neto = 0
            Me.txt_importe_total = 0
            var_primera_vez_remision = 0
            lv_detalle.ListItems.Clear
            Me.txt_remision_agente = IIf(IsNull(rsaux5!VCHA_REM_REMISION_AGENTE), "", rsaux5!VCHA_REM_REMISION_AGENTE)
            Me.txt_remision_agente.Enabled = False
            Me.txt_cliente = IIf(IsNull(rsaux5!vcha_cli_clave_id), "", rsaux5!vcha_cli_clave_id)
            Me.txt_cliente.Enabled = False
            Me.txt_nombre_cliente.Enabled = False
            Me.txt_titular.Enabled = False
            Me.txt_nombre_titular.Enabled = False
            Me.txt_agente.Enabled = False
            Me.txt_nombre_agente.Enabled = False
            var_numero_remision = rsaux5!INTE_REM_NUMERO
            var_estatus_remision = IIf(IsNull(rsaux5!char_rem_estatus), "", rsaux5!char_rem_estatus)
            Me.txt_agente = IIf(IsNull(rsaux5!VCHA_AGE_AGENTE_ID), "", rsaux5!VCHA_AGE_AGENTE_ID)
            rsaux.Open "SELECT * FROM TB_AGENTES WHERE VCHA_AGE_AGENTE_ID = '" + Me.txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               Me.txt_nombre_agente = IIf(IsNull(rsaux!VCHA_AGE_NOMBRE), "", rsaux!VCHA_AGE_NOMBRE)
               Me.txt_almacen = IIf(IsNull(rsaux!VCHA_ALM_ALMACEN_ID), "", rsaux!VCHA_ALM_ALMACEN_ID)
            Else
               Me.txt_nombre_cliente = ""
               Me.txt_almacen = ""
            End If
            rsaux.Close
            rsaux.Open "SELECT * FROM vw_clientes WHERE VCHA_CLI_CLAVE_ID = '" + Me.txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               Me.txt_titular = IIf(IsNull(rsaux!vcha_tit_titular_id), "", rsaux!vcha_tit_titular_id)
               Me.txt_nombre_titular = IIf(IsNull(rsaux!VCHA_TIT_NOMBRE), "", rsaux!VCHA_TIT_NOMBRE)
               Me.txt_nombre_cliente = IIf(IsNull(rsaux!VCHA_CLI_NOMBRE), "", rsaux!VCHA_CLI_NOMBRE)
               var_descuento_1 = IIf(IsNull(rsaux!floa_gac_Descuento_1), 0, rsaux!floa_gac_Descuento_1)
               var_descuento_2 = IIf(IsNull(rsaux!FLOA_GAC_DESCUENTO_2), 0, rsaux!FLOA_GAC_DESCUENTO_2)
            Else
               Me.txt_nombre_cliente = ""
               Me.txt_titular = ""
               Me.txt_nombre_titular = ""
            End If
            rsaux.Close
            If var_estatus_remision = "F" Then
               var_descuento_1 = IIf(IsNull(rsaux5!floa_REM_descuento_1), 0, rsaux5!floa_REM_descuento_1)
               var_descuento_2 = IIf(IsNull(rsaux5!floa_REM_descuento_2), 0, rsaux5!floa_REM_descuento_2)
            End If
                        
            
            txt_descuentos = CStr(var_descuento_1) + "% + " + CStr(var_descuento_2) + "%"
            Me.txt_remision = rsaux5!INTE_REM_NUMERO
            Me.txt_cliente = IIf(IsNull(rsaux5!vcha_cli_clave_id), "", rsaux5!vcha_cli_clave_id)
            
            
            If var_estatus_remision <> "F" Then
               Me.lv_detalle.ListItems.Clear
               If rs.State = 1 Then
                  rs.Close
               End If
               rs.Open "select * from tb_agentes where vcha_age_agente_id = '" + txt_agente + "' and VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  txt_nombre_agente = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
                  var_clave_almacen = IIf(IsNull(rs!VCHA_ALM_ALMACEN_ID), "", rs!VCHA_ALM_ALMACEN_ID)
                  Me.txt_almacen = var_clave_almacen
                  If Trim(var_clave_almacen) <> "" Then
                     
                     
                     If var_empresa = "31" Then
                        While Not rsaux5.EOF
                              If rsaux5!vcha_Art_Articulo_id = "S1005" Then
                                 Set list_item = lv_detalle.ListItems.Add(, , rsaux5!vcha_Art_Articulo_id)
                                 rsaux2.Open "select * from tb_Articulos where vcha_Art_articulo_id = 'S1005'", cnn, adOpenDynamic, adLockOptimistic
                                 list_item.SubItems(1) = IIf(IsNull(rsaux2!vcha_art_nombre_Español), "", rsaux2!vcha_art_nombre_Español)
                                 rsaux2.Close
                                 rsaux.Open "select * from tb_Detalle_lista_precios where vcha_Art_Articulo_id = 'S1005'", cnn, adOpenDynamic, adLockOptimistic
                                 var_precio = IIf(IsNull(rsaux!floa_dli_precio), 0, rsaux!floa_dli_precio) * 1.16
                                 rsaux.Close
                                 list_item.SubItems(2) = Format(var_precio, "###,###,##0.00")
                                 list_item.SubItems(3) = Format(0, "###,###,##0.00")
                                 list_item.SubItems(4) = Format(var_precio, "###,###,##0.00")
                                 If rsaux5!vcha_Art_Articulo_id <> "S1005" Then
                                    list_item.SubItems(5) = Format(IIf(IsNull(rsaux5!floa_Exi_Cantidad), 0, rsaux5!floa_Exi_Cantidad), "###,###,##0.00")
                                 Else
                                    list_item.SubItems(5) = Format(0, "###,###,##0.00")
                                 End If
                                 list_item.SubItems(6) = 0
                                 list_item.SubItems(7) = 0
                              End If
                              rsaux5.MoveNext
                        Wend
                        rsaux5.MoveFirst
                     End If
                     
                     
                     rsaux.Open "select * from VW_MERCANCIA_DISPONIBLE_AGENTES WHERE VCHA_AGE_AGENTE_ID = '" + txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        While Not rsaux.EOF
                              Set list_item = lv_detalle.ListItems.Add(, , rsaux!vcha_Art_Articulo_id)
                              list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_art_nombre_Español), "", rsaux!vcha_art_nombre_Español)
                              var_precio = IIf(IsNull(rsaux!floa_dli_precio), 0, rsaux!floa_dli_precio) * 1.16
                              list_item.SubItems(2) = Format(var_precio, "###,###,##0.00")
                              list_item.SubItems(3) = Format(IIf(IsNull(rsaux!floa_dpr_desCuento), 0, rsaux!floa_dpr_desCuento), "###,###,##0.00")
                              list_item.SubItems(4) = Format(var_precio - (var_precio * (IIf(IsNull(rsaux!floa_dpr_desCuento), 0, rsaux!floa_dpr_desCuento) / 100)), "###,###,##0.00")
                              If rsaux!vcha_Art_Articulo_id <> "S1005" Then
                                 list_item.SubItems(5) = Format(IIf(IsNull(rsaux!floa_Exi_Cantidad), 0, rsaux!floa_Exi_Cantidad), "###,###,##0.00")
                              Else
                                 list_item.SubItems(5) = Format(0, "###,###,##0.00")
                              End If
                              list_item.SubItems(6) = 0
                              list_item.SubItems(7) = 0
                              rsaux.MoveNext
                        Wend
                     End If
                     rsaux.Close
                     txt_cantidad_total = "0"
                     txt_importe_total = "0"
                     While Not rsaux5.EOF
                           If Not IsNull(rsaux5!vcha_Art_Articulo_id) Then
                              Call pro_busca_registro(lv_detalle, rsaux5!vcha_Art_Articulo_id, False)
                              If rsaux5!vcha_Art_Articulo_id = "S1005" Then
                                 lv_detalle.selectedItem.SubItems(5) = Format(0, "###,###,##0.00")
                              Else
                                 lv_detalle.selectedItem.SubItems(5) = Format((lv_detalle.selectedItem.SubItems(5) * 1) + rsaux5!FLOA_REM_CANTIDAD, "###,###,##0.00")
                              End If
                              lv_detalle.selectedItem.SubItems(6) = Format(rsaux5!FLOA_REM_CANTIDAD, "###,###,##0.00")
                              lv_detalle.selectedItem.SubItems(7) = Format(rsaux5!FLOA_REM_CANTIDAD * CDbl(lv_detalle.selectedItem.SubItems(4)), "###,###,##0.00")
                              txt_cantidad_total = Format(CDbl(txt_cantidad_total) + rsaux5!FLOA_REM_CANTIDAD, "###,###,##0.00")
                              If lv_detalle.selectedItem = "S1005" Then
                                 x = CDbl(1) / (1 - (var_descuento_1 / 100))
                                 x = CDbl(x) / (1 - (var_descuento_2 / 100))
                                 'x = x * (-1)
                                 z = (x * 1) * CDbl(lv_detalle.selectedItem.SubItems(6) * 1)
                                 'z = Format(CDbl(txt_importe_total) + CDbl(z), "###,###,##0.00")
                                 txt_importe_total = txt_importe_total - CDbl(z)
                              Else
                                 If lv_detalle.selectedItem.SubItems(3) * 1 > var_descuento_1 Then
                                    x = CDbl(lv_detalle.selectedItem.SubItems(4) * 1) / (1 - (var_descuento_1 / 100))
                                    txt_importe_total = Format(CDbl(txt_importe_total) + ((x * 1) * CDbl(lv_detalle.selectedItem.SubItems(6) * 1)), "###,###,##0.00")
                                 Else
                                    txt_importe_total = Format(CDbl(txt_importe_total) + (CDbl(lv_detalle.selectedItem.SubItems(4) * 1) * rsaux5!FLOA_REM_CANTIDAD), "###,###,##0.00")
                                 End If
                              End If
                           End If
                           rsaux5.MoveNext
                     Wend
                     var_j = lv_detalle.ListItems.Count
                     While var_j > 0
                         lv_detalle.ListItems.Item(var_j).Selected = True
                         If CDbl(lv_detalle.selectedItem.SubItems(5)) = 0 Then
                            If Me.lv_detalle.selectedItem <> "S1005" Then
                               lv_detalle.ListItems.Remove (lv_detalle.selectedItem.Index)
                            End If
                         End If
                         var_j = var_j - 1
                     Wend
                     
                     var_j = lv_detalle.ListItems.Count
                     While var_j > 0
                         lv_detalle.ListItems.Item(var_j).Selected = True
                         If CDbl(lv_detalle.selectedItem.SubItems(6)) = 0 Then
                            If Me.lv_detalle.selectedItem = "S1005" Then
                               lv_detalle.ListItems.Remove (lv_detalle.selectedItem.Index)
                            End If
                         End If
                         var_j = var_j - 1
                     Wend
                     
                     txt_importe_total = 0
                     If Me.lv_detalle.ListItems.Count > 0 Then
                        For var_j = 1 To lv_detalle.ListItems.Count
                            lv_detalle.ListItems.Item(var_j).Selected = True
                            rsaux.Open "select * from TB_DESCUENTOS_PROMOCION_CLIENTES where vcha_cli_clave_id = '" + Me.txt_cliente + "' and vcha_Art_articulo_id = '" + Me.lv_detalle.selectedItem + "' and dtim_dpr_fecha_inicio <= getdate() and dtim_dpr_fecha_fin >= getdate()", cnn, adOpenDynamic, adLockOptimistic
                            If Not rsaux.EOF Then
                               var_oferta = IIf(IsNull(rsaux!floa_dpr_desCuento), 0, rsaux!floa_dpr_desCuento)
                               var_marca = IIf(IsNull(rsaux!char_dpr_marca), "", rsaux!char_dpr_marca)
                               Me.lv_detalle.selectedItem.SubItems(3) = var_oferta
                               If var_marca <> "*" Then
                                  If CDbl(lv_detalle.selectedItem.SubItems(2)) = 0 Then
                                     Me.lv_detalle.selectedItem.SubItems(2) = Format(0, "###,###,##0.00")
                                  Else
                                     If Trim(Me.lv_detalle.selectedItem) = "S1005" Then
                                        Me.lv_detalle.selectedItem.SubItems(2) = Format(0, "###,###,##0.00")
                                     Else
                                        If var_descuento_1 = 100 Then
                                           Me.lv_detalle.selectedItem.SubItems(2) = Format(0, "###,###,##0.00")
                                        Else
                                           Me.lv_detalle.selectedItem.SubItems(2) = Format(lv_detalle.selectedItem.SubItems(2) / (1 - (var_descuento_1 / 100)), "###,###,##0.00")
                                        End If
                                     End If
                                  End If
                               End If
                               Me.lv_detalle.selectedItem.SubItems(4) = Round(CDbl(lv_detalle.selectedItem.SubItems(2)) - (CDbl(lv_detalle.selectedItem.SubItems(2)) * (var_oferta / 100)), 2)
                               Me.lv_detalle.selectedItem.SubItems(7) = Round(CDbl(lv_detalle.selectedItem.SubItems(4)) * CDbl(lv_detalle.selectedItem.SubItems(6)), 2)
                            End If
                            If Me.lv_detalle.selectedItem = "S1005" Then
                               x = CDbl(1) / (1 - (var_descuento_1 / 100))
                               x = CDbl(x) / (1 - (var_descuento_2 / 100))
                               z = (x * 1) * CDbl(lv_detalle.selectedItem.SubItems(6) * 1)
                               txt_importe_total = Format(CDbl(txt_importe_total) - (z * 1), "###,###,##0.00")
                            Else
                               txt_importe_total = Format(CDbl(txt_importe_total) + (CDbl(lv_detalle.selectedItem.SubItems(4) * 1) * CDbl(lv_detalle.selectedItem.SubItems(6))), "###,###,##0.00")
                            End If
                            rsaux.Close
                        Next var_j
                     End If

                  End If
               End If
               rs.Close
            Else
               MsgBox "La remisión ya fue facturada", vbOKOnly, "ATENCION"
               txt_cantidad_total = "0"
               txt_importe_total = "0"
               While Not rsaux5.EOF
                     If Not IsNull(rsaux5!vcha_Art_Articulo_id) Then
                        Set list_item = lv_detalle.ListItems.Add(, , rsaux5!vcha_Art_Articulo_id)
                        If rsaux4.State = 1 Then
                           rsaux4.Close
                        End If
                        rsaux4.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_ART_ARTICULO_ID = '" + IIf(IsNull(rsaux5!vcha_Art_Articulo_id), "", rsaux5!vcha_Art_Articulo_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                        list_item.SubItems(1) = IIf(IsNull(rsaux4!vcha_art_nombre_Español), "", rsaux4!vcha_art_nombre_Español)
                        rsaux4.Close
                        var_precio = IIf(IsNull(rsaux5!floa_dli_precio), 0, rsaux5!floa_dli_precio)
                        
                        list_item.SubItems(2) = Format(var_precio, "###,###,##0.00")
                        list_item.SubItems(3) = Format(IIf(IsNull(rsaux5!FLOA_REM_PROMOCION_1), 0, rsaux5!FLOA_REM_PROMOCION_1), "###,###,##0.00")
                        list_item.SubItems(5) = Format("0", "###,###,##0.00")
                        list_item.SubItems(6) = Format(rsaux5!FLOA_REM_CANTIDAD, "###,###,##0.00")
                        list_item.SubItems(4) = Format(IIf(IsNull(rsaux5!FLOA_REM_PRECIO), 0, rsaux5!FLOA_REM_PRECIO) * 1.16, "###,###,##0.00")
                        list_item.SubItems(7) = Format(rsaux5!FLOA_REM_CANTIDAD * CDbl(list_item.SubItems(4)), "###,###,##0.00")
                        txt_cantidad_total = Format(CDbl(txt_cantidad_total) + rsaux5!FLOA_REM_CANTIDAD, "###,###,##0.00")
                        txt_importe_total = Format(CDbl(txt_importe_total) + (CDbl(list_item.SubItems(4) * 1) * CDbl(rsaux5!FLOA_REM_CANTIDAD)), "###,###,##0.00")
                     End If
                     rsaux5.MoveNext
               Wend
            End If
         Else
            MsgBox "Número de remisión no existe", vbOKOnly, "ATENCION"
         End If
         rsaux5.Close
      Else
         MsgBox "Número de nota incorrecto", vbOKOnly, "ATENCION"
      End If
      frm_numero_nota.Visible = False
   End If
End Sub

Private Sub txt_numero_nota_LostFocus()
   Me.frm_numero_nota.Visible = False
End Sub

Private Sub txt_remision_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(Me.txt_remision_agente) <> "" Then
         If Me.lv_detalle.ListItems.Count > 0 Then
            Me.lv_detalle.ListItems.Item(1).Selected = True
            Me.lv_detalle.ListItems.Item(1).EnsureVisible
         End If
         Call pro_enfoque(KeyAscii)
      Else
         MsgBox "Debe de indicar la remisión del agente", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_remision_agente_LostFocus()
   If Trim(Me.txt_remision_agente) <> "" Then
      Me.txt_remision_agente.Enabled = False
   Else
      MsgBox "Debe de indicar la remisión del agente", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub txt_serie_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frm_factura_cancelar.Visible = False
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub factura_vistas()

                        If rs.State = 1 Then
                           rs.Close
                        End If
                        rs.Open "select isnull(max(inte_tem_consecutivo),0) from tb_temp_factura_embarques", cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           var_consecutivo = rs(0).Value
                        Else
                           var_consecutivo = 0
                        End If
                        rs.Close
                        var_consecutivo = var_consecutivo + 1
                        rs.Open "insert into tb_temp_factura_embarques (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                     
                        Cadena = "EXEC SP_CREA_TABLA_FACTURAS_VISTAS " + CStr(var_consecutivo) + ",'" + var_empresa + "'," + txt_numero_embarque
                        If rsaux3.State = 1 Then
                           rsaux3.Close
                        End If
                        rsaux3.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                        rsaux3.Open "select distinct inte_car_numero from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " ORDER BY INTE_CAR_NUMERO", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux3.EOF Then
                           var_Archivo = App.Path & "\factura" + Trim(Str(rsaux3!inte_car_numero)) + ".bat"
                           Open (App.Path & "\factura" + Trim(Str(rsaux3!inte_car_numero)) + ".bat") For Output As #2
                           While Not rsaux3.EOF
                                 If rs.State = 1 Then
                                    rs.Close
                                 End If
                                 If var_empresa <> "03" Then
                                    rs.Open "select * from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_Car_numero = " + Str(rsaux3!inte_car_numero) + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " ORDER BY INTE_SAL_CONSECUTIVO_FACTURA", cnn, adOpenDynamic, adLockOptimistic
                                 Else
                                    rs.Open "select * from TB_TEMP_FACTURA_EMBARQUES where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + txt_numero_embarque + " and inte_Car_numero = " + Str(rsaux3!inte_car_numero) + " and inte_tem_consecutivo = " + CStr(var_consecutivo) + " ORDER BY vcha_sal_descripcion_factura", cnn, adOpenDynamic, adLockOptimistic
                                 End If
                                 If Not rs.EOF Then
                                   'AQUI EMPIEZA LA FACTURA
                                    Open (App.Path & "\factura" + Trim(Str(rsaux3!inte_car_numero)) + ".txt") For Output As #1
                                    'Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
                                    'Print #1, Chr(27) + Chr(67) + Chr(44) + Chr(13)
                                    'Print #1, ""
                                    Print #1, Chr(15) + Chr(27) + Chr(64)
                                    If var_empresa = "18" Then
                                       Print #1, ""
                                    End If
                                    Print #1, Spc(105); Str(rsaux3!inte_car_numero)
                                    Print #1, ""
                                    Print #1, Spc(105); Str(rs!INTE_CAR_PLAZO) + " DIAS DE VENCIMIENTO" + "                  " + Format(rs!dtim_Car_fecha, "Short Date")
                                    Print #1, ""
                                    Print #1, ""
                                    'Print #1, Spc(92); Str(rs!inte_car_PLAZO) + " DIAS DE VENCIMIENTO"
                                    var_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                                    For var_j = 1 + Len(Trim(var_cliente)) To 83
                                        var_cliente = var_cliente + " "
                                    Next var_j
                                    If var_unidad_organizacional = "21" Then
                                       var_cliente = var_cliente + "               MEXICO, D.F."
                                    Else
                                       var_cliente = var_cliente + "               AGUASCALIENTES, AGS."
                                    End If
                                    Print #1, Spc(10); var_cliente
                                    var_domicilio = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " C.P. " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
                                    'For var_j = 1 + Len(Trim(var_domicilio)) To 83
                                    '    var_domicilio = var_domicilio + " "
                                    'Next var_j
                                    
                                    rsaux11.Open "select vcha_cli_referencia from tb_Clientes where vcha_Cli_clave_id = '" + IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
                                    If Not rsaux11.EOF Then
                                       var_referencia_Bancaria = IIf(IsNull(rsaux11!VCHA_CLI_REFERENCIA), "", rsaux11!VCHA_CLI_REFERENCIA)
                                       If var_referencia_Bancaria <> "" Then
                                          For var_j = 1 + Len(Trim(var_domicilio)) To 105
                                              var_domicilio = var_domicilio + " "
                                          Next var_j
                                          var_domicilio = var_domicilio + " REF. BANCARIA: " + var_referencia_Bancaria
                                       End If
                                    End If
                                    rsaux11.Close
                                    
                                    
                                    
                                    
                                    var_agente = ""
                                    var_agente = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
                                    For var_j = 1 + Len(Trim(var_agente)) To 8
                                        var_agente = var_agente + " "
                                    Next var_j
                                    If rsaux4.State = 1 Then
                                       rsaux4.Close
                                    End If
                                    rsaux4.Open "SELECT * FROM TB_AGENTES WHERE VCHA_AGE_AGENTE_ID = '" + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
                                    If Not rsaux4.EOF Then
                                       var_agente = var_agente + IIf(IsNull(rsaux4!VCHA_AGE_NOMBRE), "", rsaux4!VCHA_AGE_NOMBRE)
                                    Else
                                       var_agente = var_agente + ""
                                    End If
                                    rsaux4.Close
                                    var_domicilio = var_domicilio
                                    'Print #1, Spc(111); var_agente
                                    Print #1, Spc(10); var_domicilio
                                    var_ciudad = ""
                                    var_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                                    For var_j = 1 + Len(Trim(var_ciudad)) To 37
                                        var_ciudad = var_ciudad + " "
                                    Next var_j
                                 
                                    var_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                                    var_ciudad = var_ciudad
                                    var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                                    var_ciudad = var_ciudad + var_rfc
                                             
                                    For var_j = 1 + Len(Trim(var_estado)) To 46
                                        var_estado = var_estado + " "
                                    Next var_j
                                 
   
                                    For var_j = 1 + Len(Trim(var_ciudad)) To 14
                                        var_ciudad = var_ciudad + " "
                                    Next var_j
                               
                                    var_ciudad = var_ciudad + "                                                      " + var_agente
                                
                                    VAR_EMBARQUE = "EMB.: " + txt_numero_embarque
                                    var_ordern_surtido = x
                                    Print #1, Spc(10); var_ciudad
                                    var_rfc = "RFC:  " + var_rfc
                                    var_rfc = IIf(IsNull(rs!vcha_ESB_ESTABLECIMIENTO_id), "", rs!vcha_ESB_ESTABLECIMIENTO_id) + " " + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                                    For var_j = 1 + Len(Trim(var_rfc)) To 89
                                        var_rfc = var_rfc + " "
                                    Next var_j
                                    If var_empresa = "18" Then
                                       If rs!VCHA_MOV_MOVIMIENTO_ID = "FV" Then
                                          rsaux5.Open "SELECT * FROM TB_REMISIONES WHERE INTE_EMO_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                                          If Not rsaux5.EOF Then
                                             If IsNumeric(IIf(IsNull(rsaux5!VCHA_REM_REMISION_AGENTE), "0", rsaux5!VCHA_REM_REMISION_AGENTE)) Then
                                                var_rfc = var_rfc + "               R.A.: " + Trim(Str(IIf(IsNull(rsaux5!VCHA_REM_REMISION_AGENTE), "0", rsaux5!VCHA_REM_REMISION_AGENTE))) + " "
                                             Else
                                                var_rfc = var_rfc + "               R.A.: " + Trim((IIf(IsNull(rsaux5!VCHA_REM_REMISION_AGENTE), "0", rsaux5!VCHA_REM_REMISION_AGENTE))) + " "
                                             End If
                                             var_rfc = var_rfc + " R.S.: " + Trim(Str(IIf(IsNull(rsaux5!INTE_REM_NUMERO), 0, rsaux5!INTE_REM_NUMERO))) + " " + VAR_EMBARQUE
                                          Else
                                             var_rfc = var_rfc + "               PED.: " + Trim(Str(IIf(IsNull(rs!inte_ped_numero), 0, rs!inte_ped_numero))) + " "
                                             var_rfc = var_rfc + " O.S.: " + Trim(Str(IIf(IsNull(rs!INTE_ORS_ORDEN_SURTIDO), 0, rs!INTE_ORS_ORDEN_SURTIDO))) + " " + VAR_EMBARQUE
                                          End If
                                          rsaux5.Close
                                       Else
                                          var_rfc = var_rfc + "               PED.: " + Trim(Str(IIf(IsNull(rs!inte_ped_numero), 0, rs!inte_ped_numero))) + " "
                                          var_rfc = var_rfc + " O.S.: " + Trim(Str(IIf(IsNull(rs!INTE_ORS_ORDEN_SURTIDO), 0, rs!INTE_ORS_ORDEN_SURTIDO))) + " " + VAR_EMBARQUE
                                       End If
                                    Else
                                    var_rfc = var_rfc + "               PED.: " + Trim(Str(IIf(IsNull(rs!inte_ped_numero), 0, rs!inte_ped_numero))) + " "
                                    var_rfc = var_rfc + " O.S.: " + Trim(Str(IIf(IsNull(rs!INTE_ORS_ORDEN_SURTIDO), 0, rs!INTE_ORS_ORDEN_SURTIDO))) + " " + VAR_EMBARQUE
                                    End If
                                    Print #1, var_rfc
                                    'Print #1, Spc(10); IIf(IsNull(rs!vcha_esb_establecimiento_id), "", rs!vcha_esb_establecimiento_id)
                                    Print #1, ""
                                    Print #1, ""
                                    var_importe_descuento_1 = 0
                                    var_importe_descuento_2 = 0
                                    var_importe_descuento_3 = 0
                                    var_contador_promociones = 0
                                    var_cantidad_total = 0
                                    For var_k = 1 To var_renglones_factura
                                        If Not rs.EOF Then
                                           var_linea = ""
                                           var_marca_promocion = " "
                                           var_promocion_1 = IIf(IsNull(rs!floa_sal_promocion_1), 0, rs!floa_sal_promocion_1)
                                           var_promocion_2 = IIf(IsNull(rs!FLOA_SAL_PROMOCION_2), 0, rs!FLOA_SAL_PROMOCION_2)
                                           If var_promocion_1 > 0 Then
                                              var_marca_promocion = "*"
                                              var_contador_promociones = var_contador_promociones + 1
                                           End If
                                           If var_promocion_2 > 0 Then
                                              var_marca_promocion = "*"
                                              var_contador_promociones = var_contador_promociones + 1
                                           End If
                                           var_linea = IIf(IsNull(rs!vcha_agr_agrupador_id), "", rs!vcha_agr_agrupador_id)
                                           If var_empresa = "31" Then
                                              For var_j = 1 + Len(Trim(var_linea)) To 18
                                                  var_linea = var_linea + " "
                                              Next var_j
                                           Else
                                              For var_j = 1 + Len(Trim(var_linea)) To 15
                                                  var_linea = var_linea + " "
                                              Next var_j
                                           End If
                                           var_linea = var_linea + IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura)
                                           var_i = 0
                                           While Len((var_linea)) < 115
                                                 var_linea = var_linea + " "
                                           Wend
                                           var_linea = var_linea + " "
                                           var_linea = var_linea + var_marca_promocion
                                           var_cantidad = Format(IIf(IsNull(rs!Cantidad), 0, rs!Cantidad), "###,###,##0.00")
                                           var_cantidad_total = var_cantidad_total + IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)
                                           If Len(Trim(var_cantidad)) < 14 Then
                                              For var_j = 1 + Len(Trim(var_cantidad)) To 14
                                                  var_cantidad = " " + var_cantidad
                                              Next var_j
                                           End If
                                           var_precio = IIf(IsNull(rs!Importe), 0, rs!Importe)
                                           var_descuento_1 = IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1)
                                           var_descuento_2 = IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2)
                                           var_descuento_3 = IIf(IsNull(rs!floa_car_porcentaje_descuento_3), 0, rs!floa_car_porcentaje_descuento_3)
                                           var_porcentaje = (100 - var_descuento_1) / 100
                                           var_precio = var_precio * var_porcentaje
                                           var_importe_descuento_1_2 = (IIf(IsNull(rs!Importe), 0, rs!Importe) - var_precio)
                                           var_importe_descuento_1 = var_importe_descuento_1 + (IIf(IsNull(rs!Importe), 0, rs!Importe) - var_precio)
                                           var_precio = var_precio * ((100 - var_descuento_2) / 100)
                                           var_importe_descuento_2 = var_importe_descuento_2 + (IIf(IsNull(rs!Importe), 0, rs!Importe) - (var_importe_descuento_1_2 + var_precio))
                                           var_precio = var_precio * ((100 - var_descuento_3) / 100)
                                           var_precio = var_precio / IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)
                                           'var_precio_str = Format(var_precio / IIf(IsNull(rs!cantidad), 0, rs!cantidad), "###,###,##0.00")
                                           var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                                           If Len(Trim(var_rfc)) > 0 Then
                                              var_precio_str = Format(IIf(IsNull(rs!Importe), 0, rs!Importe) / IIf(IsNull(rs!Cantidad), 0, rs!Cantidad), "###,###,##0.00")
                                           Else
                                              var_precio_str = Format((IIf(IsNull(rs!Importe), 0, rs!Importe) / IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)) * (1 + (rs!floa_car_porcentaje_iva / 100)), "###,###,##0.00")
                                           End If
                                           If Len(Trim(var_precio_str)) < 14 Then
                                              For var_j = 1 + Len(Trim(var_precio_str)) To 14
                                                  var_precio_str = " " + var_precio_str
                                              Next var_j
                                           End If
                                           var_linea = var_linea + var_cantidad + var_precio_str
                                           If Len(Trim(var_rfc)) > 0 Then
                                              var_importe = Format((IIf(IsNull(rs!Importe), 0, rs!Importe)), "###,###,##0.00")
                                              If Len(Trim(var_importe)) < 14 Then
                                                 For var_j = 1 + Len(Trim(var_importe)) To 14
                                                     var_importe = " " + var_importe
                                                 Next var_j
                                              End If
                                           Else
                                              var_importe = Format((IIf(IsNull(rs!Importe), 0, rs!Importe) * (1 + (rs!floa_car_porcentaje_iva / 100))), "###,###,##0.00")
                                              If Len(Trim(var_importe)) < 14 Then
                                                 For var_j = 1 + Len(Trim(var_importe)) To 14
                                                     var_importe = " " + var_importe
                                                 Next var_j
                                              End If
                                           End If
                                           var_linea = var_linea + var_importe
                                     
                                           Print #1, var_linea
                                           rs.MoveNext
                                        Else
                                           Print #1, ""
                                        End If
                                    Next var_k
                                    Print #1, ""
                                    'Print #1, ""
                                    rs.MoveFirst
                                
                                    var_cantidad_total_str = Format(var_cantidad_total, "###,###,##0.00")
                                    var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                                    If Len(Trim(var_rfc)) > 0 Then
                                       var_cantidad_letra = rs!vcha_car_importe_letra
                                       var_importe_descuento_1_str = Format(IIf(IsNull(rs!FLOA_CAR_IMPORTE_DESCUENTO_1), 0, rs!FLOA_CAR_IMPORTE_DESCUENTO_1) / IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio), "###,###,##0.00")
                                       If Len(Trim(var_importe_descuento_1_str)) < 14 Then
                                          For var_j = 1 + Len(Trim(var_importe_descuento_1_str)) To 14
                                              var_importe_descuento_1_str = " " + var_importe_descuento_1_str
                                          Next var_j
                                       End If
                                       var_importe_descuento_2_str = Format(IIf(IsNull(rs!FLOA_CAR_IMPORTE_DESCUENTO_2), 0, rs!FLOA_CAR_IMPORTE_DESCUENTO_2) / IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio), "###,###,##0.00")
                                       If Len(Trim(var_importe_descuento_2_str)) < 14 Then
                                          For var_j = 1 + Len(Trim(var_importe_descuento_2_str)) To 14
                                              var_importe_descuento_2_str = " " + var_importe_descuento_2_str
                                          Next var_j
                                       End If
                                    Else
                                       var_cantidad_letra = rs!vcha_car_importe_letra
                                       var_importe_descuento_1_str = Format((IIf(IsNull(rs!FLOA_CAR_IMPORTE_DESCUENTO_1), 0, rs!FLOA_CAR_IMPORTE_DESCUENTO_1)) * (1 + (rs!floa_car_porcentaje_iva / 100) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio))), "###,###,##0.00")
                                       If Len(Trim(var_importe_descuento_1_str)) < 14 Then
                                          For var_j = 1 + Len(Trim(var_importe_descuento_1_str)) To 14
                                              var_importe_descuento_1_str = " " + var_importe_descuento_1_str
                                          Next var_j
                                       End If
                                       var_importe_descuento_2_str = Format((IIf(IsNull(rs!FLOA_CAR_IMPORTE_DESCUENTO_2), 0, rs!FLOA_CAR_IMPORTE_DESCUENTO_2)) * (1 + (rs!floa_car_porcentaje_iva / 100) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio))), "###,###,##0.00")
                                       If Len(Trim(var_importe_descuento_2_str)) < 14 Then
                                          For var_j = 1 + Len(Trim(var_importe_descuento_2_str)) To 14
                                              var_importe_descuento_2_str = " " + var_importe_descuento_2_str
                                          Next var_j
                                       End If
                                    End If
                                    var_linea = "- DESCUENTO DEL " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1))) + "%"
                                    If Len(Trim(var_linea)) < 145 Then
                                       For var_j = 1 + Len(Trim(var_linea)) To 145
                                           var_linea = var_linea + " "
                                       Next var_j
                                    End If
                                    Print #1, var_linea + var_importe_descuento_1_str
                                    var_linea = "- DESCUENTO POR PAGO OPORTUNO " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2))) + "%"
                                    If Len(Trim(var_linea)) < 145 Then
                                       For var_j = 1 + Len(Trim(var_linea)) To 145
                                           var_linea = var_linea + " "
                                       Next var_j
                                    End If
                                    var_linea = var_linea + var_importe_descuento_2_str
                                    Print #1, var_linea
                                    If var_contador_promociones > 0 Then
                                       Print #1, "PROMOCION EN ARTICULOS MARCADOS CON *"
                                    Else
                                       Print #1, ""
                                    End If
                                             
                                    var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                                    var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
                                            
                                    If Len(Trim(var_linea)) < 117 Then
                                       For var_j = 1 + Len(Trim(var_linea)) To 117
                                           var_x = var_j Mod 2
                                           If var_x >= 1 Then
                                              var_linea = " " + var_linea
                                           Else
                                              var_linea = var_linea + " "
                                           End If
                                       Next var_j
                                    End If
                               
                                    If Len(Trim(var_rfc)) = 0 Then
                                       var_subimporte = Format(Round((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                                       If Len(Trim(var_subimporte)) < 14 Then
                                          For var_j = 1 + Len(Trim(var_subimporte)) To 14
                                              var_subimporte = " " + var_subimporte
                                          Next var_j
                                       End If
                                       var_iva = "-"
                                       For var_j = 1 + Len(Trim(var_iva)) To 11
                                           var_iva = " " + var_iva
                                       Next var_j
                                    Else
                                       var_subimporte = Format(Round(((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                                       If Len(Trim(var_subimporte)) < 14 Then
                                          For var_j = 1 + Len(Trim(var_subimporte)) To 14
                                              var_subimporte = " " + var_subimporte
                                          Next var_j
                                       End If
                                       var_iva = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                       If Len(Trim(var_iva)) < 14 Then
                                          For var_j = 1 + Len(Trim(var_iva)) To 14
                                              var_iva = " " + var_iva
                                          Next var_j
                                       End If
                                    End If
                                
                                    If Len(Trim(var_subimporte)) < 14 Then
                                       For var_j = 1 + Len(Trim(var_subimporte)) To 14
                                           var_subimporte = " " + var_subimporte
                                       Next var_j
                                    End If
                                    var_espacios = 131 - Len(var_cantidad_total_str)
                                    var_cantidad_total_str = Trim(var_cantidad_total_str)
                                    If Len(Trim(var_cantidad_total_str)) < 14 Then
                                       For var_j = 1 + Len(Trim(var_cantidad_total_str)) To 14
                                           var_cantidad_total_str = " " + var_cantidad_total_str
                                       Next var_j
                                    End If
                                    var_subimporte = Trim(var_subimporte)
                                    If Len(Trim(var_subimporte)) < 24 Then
                                       For var_j = 1 + Len(Trim(var_subimporte)) To 24
                                           var_subimporte = " " + var_subimporte
                                       Next var_j
                                    End If
                                         
                                    var_cantidad_total_str = var_linea + var_cantidad_total_str + "    " + var_subimporte
                                    'Print #1, Spc(var_espacios); var_cantidad_total_str; Spc(8); var_subimporte
                                    Print #1, var_cantidad_total_str
                                    var_linea = "                                                                          ESTA FACTURA SERA PAGADA EN UNA SOLA EXHIBICION                        " + var_iva
                                    Print #1, var_linea
                                    var_dia = Day(rs!dtim_Car_fecha)
                                    var_mes = Month(rs!dtim_Car_fecha)
                                    var_año = Year(rs!dtim_Car_fecha)
                                    If var_empresa = "31" Then
                                       var_linea = "                                                       " + CStr(var_dia) + "     " + CStr(var_mes)
                                    Else
                                       var_linea = "                                                             " + CStr(var_dia) + "     " + CStr(var_mes)
                                    End If
                                    If Len(var_linea) < 145 Then
                                       For var_j = 1 + Len(var_linea) To 145
                                           var_linea = var_linea + " "
                                       Next var_j
                                    End If
                                
                                    var_importe = Format(Round((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                                            
                                    If Len(Trim(var_importe)) < 14 Then
                                       For var_j = 1 + Len(Trim(var_importe)) To 14
                                           var_importe = " " + var_importe
                                       Next var_j
                                    End If
                                       
                                    'var_linea = "                                                                   ESTA FACTURA SERA PAGADA EN UNA SOLA EXHIBICION                               " + var_iva
                                    'var_linea = "                                                                                                                                                 " + var_importe
                                            
                                    var_linea = var_linea + var_importe
                                    Print #1, var_linea
                                            
                                    var_linea = var_importe
                                    If Len(Trim(var_linea)) < 20 Then
                                       For var_j = 1 + Len(Trim(var_linea)) To 20
                                           var_linea = " " + var_linea
                                       Next var_j
                                    End If
                                    var_linea = var_linea + " " + var_cantidad_letra
                                    Print #1, Spc(2); CStr(var_año); var_linea
                                
                                    var_linea = ""
                                    Print #1, ""
                                    Print #1, ""
                                    If var_empresa = "31" Then
                                       Print #1, Spc(10); Trim(IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE))
                                       Print #1, Spc(10); Trim(IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " " + IIf(IsNull(rs!VCHA_CLI_COLONIA), "", rs!VCHA_CLI_COLONIA))
                                       Print #1, Spc(10); Trim(IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre) + " " + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre))
                                    Else
                                       Print #1, Spc(5); Trim(IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE))
                                       Print #1, Spc(5); Trim(IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " " + IIf(IsNull(rs!VCHA_CLI_COLONIA), "", rs!VCHA_CLI_COLONIA))
                                       Print #1, Spc(5); Trim(IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre) + " " + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre))
                                    End If
                                    If var_empresa <> "03" Then
                                       Print #1, ""
                                       Print #1, ""
                                    Else
                                       Print #1, ""
                                       Print #1, ""
                                    End If
                                    Print #1, ""
                                    Print #1, ""
                                    Close #1
                                    If Trim(var_empresa) <> "03" Then
                                       Print #2, "copy " + App.Path + "\factura" + Trim(Str(rsaux3!inte_car_numero)) + ".txt lpt1"
                                    Else
                                       Print #2, "copy " + App.Path + "\factura" + Trim(Str(rsaux3!inte_car_numero)) + ".txt lpt1"
                                    End If
                                    'AQUI TERMINA LA FACTURA
                                  End If
                                  rs.Close
                                  rsaux3.MoveNext
                           Wend
                           Close #2
                           x = Shell(var_Archivo, vbHide)
                           rsaux3.Close
                           'Aqui se termina de imprimir la factura
                        End If
                        If rsaux3.State = 1 Then
                           rsaux3.Close
                        End If
                        rsaux3.Open "delete from TB_TEMP_FACTURA_EMBARQUES where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic


End Sub

Private Sub txt_titular_GotFocus()
   Me.frm_factura_cancelar.Visible = False
End Sub

Private Sub txt_titular_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct vcha_tit_titular_id, vcha_tit_nombre from vw_clientes where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_age_agente_id = '" + txt_agente + "' order by vcha_tit_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_tit_titular_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Titulares"
      var_tipo_lista = 3
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 7 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_titular_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_titular_LostFocus()
   If Trim(txt_titular) <> "" Then
      rs.Open "SELECT * FROM vw_clientes WHERE vcha_tit_titular_id = '" + Me.txt_titular + "' AND VCHA_AGE_AGENTE_ID = '" + txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_titular.Enabled = False
         Me.txt_nombre_titular.Enabled = False
         Me.txt_nombre_titular = IIf(IsNull(rs!VCHA_TIT_NOMBRE), "", rs!VCHA_TIT_NOMBRE)
         Me.txt_cliente = ""
         Me.txt_nombre_cliente = ""
         Me.txt_descuentos = ""
         Me.txt_titular.Enabled = False
         Me.txt_nombre_titular.Enabled = False
      Else
         MsgBox "Clave de titular incorrecta", vbOKOnly, "ATENCION"
         Me.txt_titular = ""
         Me.txt_nombre_titular = ""
         Me.txt_cliente = ""
         Me.txt_nombre_cliente = ""
         Me.txt_descuentos = ""
         Me.txt_nombre_cliente = ""
      End If
      rs.Close
   Else
      txt_nombre_cliente = ""
   End If
End Sub
