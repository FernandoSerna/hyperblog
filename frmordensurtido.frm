VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmordensurtido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generacion de Ordenes de Surtido"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   Icon            =   "frmordensurtido.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11670
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   795
      Picture         =   "frmordensurtido.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame frm_factura_ceros 
      Height          =   1110
      Left            =   3345
      TabIndex        =   21
      Top             =   315
      Width           =   2310
      Begin VB.TextBox txt_factura_ceros 
         Height          =   345
         Left            =   195
         TabIndex        =   22
         Top             =   555
         Width           =   1905
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         Caption         =   " Factura en ceros"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   30
         TabIndex        =   23
         Top             =   120
         Width           =   2235
      End
   End
   Begin VB.CommandButton cmd_factura_ceros 
      Appearance      =   0  'Flat
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1125
      Picture         =   "frmordensurtido.frx":09CC
      TabIndex        =   3
      ToolTipText     =   "Factura en ceros"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   285
      Left            =   4860
      TabIndex        =   20
      Top             =   60
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      Picture         =   "frmordensurtido.frx":0ACE
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Generar Ordenes de Surtido Alt + G"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame frm_pedido_resurtir 
      Height          =   1110
      Left            =   1950
      TabIndex        =   17
      Top             =   285
      Width           =   2310
      Begin VB.TextBox txt_pedido_resurtir 
         Height          =   345
         Left            =   195
         TabIndex        =   19
         Top             =   555
         Width           =   1905
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         Caption         =   " Número de Pedido a Resurtir"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   30
         TabIndex        =   18
         Top             =   120
         Width           =   2235
      End
   End
   Begin VB.CommandButton cmd_resurtir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   465
      Picture         =   "frmordensurtido.frx":0C18
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Resurtir Pedido"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11205
      Picture         =   "frmordensurtido.frx":0D1A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_generar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      Picture         =   "frmordensurtido.frx":1354
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Generar Ordenes de Surtido Alt + G"
      Top             =   30
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Frame frm_imprimir 
      Height          =   1200
      Left            =   3060
      TabIndex        =   11
      Top             =   270
      Width           =   3255
      Begin VB.TextBox txt_numero 
         Height          =   345
         Left            =   1065
         TabIndex        =   14
         Top             =   555
         Width           =   1905
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   360
         TabIndex        =   15
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Imprimir Orden de Surtido"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   30
         TabIndex        =   13
         Top             =   120
         Width           =   3180
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   120
      TabIndex        =   10
      Top             =   330
      Width           =   11475
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordensurtido.frx":149E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordensurtido.frx":1D78
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordensurtido.frx":2652
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordensurtido.frx":2BEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordensurtido.frx":34CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordensurtido.frx":3DA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordensurtido.frx":467E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordensurtido.frx":4790
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordensurtido.frx":48A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordensurtido.frx":49B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordensurtido.frx":4AC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordensurtido.frx":4BD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordensurtido.frx":4CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordensurtido.frx":522C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordensurtido.frx":576E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordensurtido.frx":5880
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordensurtido.frx":5992
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordensurtido.frx":5AA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmordensurtido.frx":5BAE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   6735
      Left            =   60
      TabIndex        =   24
      Top             =   525
      Width           =   11535
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   375
         Picture         =   "frmordensurtido.frx":5CC0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   150
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   45
         Picture         =   "frmordensurtido.frx":5ED6
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   150
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1035
         Picture         =   "frmordensurtido.frx":5FD8
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   150
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   705
         Picture         =   "frmordensurtido.frx":60AA
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Marcar (Enter)"
         Top             =   150
         Width           =   330
      End
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1365
         Picture         =   "frmordensurtido.frx":62F4
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   150
         Width           =   330
      End
      Begin VB.Frame Frame3 
         Height          =   90
         Left            =   30
         TabIndex        =   26
         Top             =   465
         Width           =   11475
      End
      Begin MSComctlLib.ListView lv_pedidos 
         Height          =   6045
         Left            =   45
         TabIndex        =   25
         Top             =   585
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   10663
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   20
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Pedido"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Orden"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Agente"
            Object.Width           =   3263
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Titular"
            Object.Width           =   3263
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Establecimiento"
            Object.Width           =   3263
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Cliente     "
            Object.Width           =   3263
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Piezas    "
            Object.Width           =   1589
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Importe    "
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Autorizado"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "autorizo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Fecha autorizo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Fecha Pedido"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Estatus"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Descuento 1"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Descuento 2"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Nombre Autorizo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "almacen"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "M"
            Object.Width           =   441
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "pedido ceros"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "M"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   5325
      TabIndex        =   12
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmordensurtido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_almacen As String
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_pedido_factura_ceros As Integer

Private Sub cmd_factura_ceros_Click()
   Me.txt_factura_ceros.Text = ""
   Me.frm_factura_ceros.Visible = True
   Me.txt_factura_ceros.SetFocus
End Sub

Private Sub cmd_generar_Click()
   Set TB_ENC_ORDEN_SURTIDO = New TB_ENC_ORDEN_SURTIDO
   Set TB_DET_ORDEN_SURTIDO_I = New TB_DET_ORDEN_SURTIDO_I
   Set TB_ENC_PEDIDOS_M = New TB_ENC_PEDIDOS_M
   Dim var_maximo_orden As Double
   Dim var_existen As Double
   Dim var_apartados As Double
   Dim var_disponible As Double
   Dim var_cantidad_pedidia As Double
   Dim var_surtir As Double
   Dim var_contador As Double
   Dim var_costo As Double
   Dim var_factura_ceros As Double
   Dim var_clave_moneda As String
   Dim var_promocion_1 As Double
   Dim var_promocion_2 As Double
   Dim var_hora As String
   Dim var_tipo_pedido As String
   Dim var_posibles_catalogos As Boolean
   Dim var_articulos_pedidos As Boolean
   Dim var_posible As Boolean
   Dim i As Double
   Dim j As Double
   si = MsgBox("¿Deseas ejecutar el proceso de ordenes de surtido", vbYesNo, "ATENCION")
   If si = 6 Then
      var_contador = 0
      rs.Open "select * from vw_suma_pedidos where char_ped_estatus <> 'C' or char_ped_estatus <> 'S'  and vcha_emp_empresa_id ='" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' order by char_pri_prioridad_id,inte_ped_numero", cnn, adOpenDynamic, adLockOptimistic
      var_hora = CStr(Time)
      While Not rs.EOF
         var_clave_moneda = rs!VCHA_MON_MONEDA_ID
         j = lv_pedidos.ListItems.Count
         var_factura_ceros = 0
         For i = 1 To j
            lv_pedidos.ListItems.item(i).Selected = True
            If lv_pedidos.selectedItem = rs!inte_ped_numero Then
               If lv_pedidos.selectedItem.SubItems(17) = "*" Then
                  var_factura_ceros = 1
               End If
            End If
         Next i
         var_almacen = IIf(IsNull(rs!VCHA_ALM_ALMACEN_ID), "", rs!VCHA_ALM_ALMACEN_ID)
         If Trim(var_almacen) <> "" Then
            If IsNull(rs!inte_ped_autorizo) Then
               var_a = 0
            Else
               var_a = rs!inte_ped_autorizo
            End If
            If var_a = 1 And (Len(Trim(rs!char_ped_estatus)) = 0 Or rs!char_ped_estatus = "I") Then
               var_contador = var_contador + 1
               rsaux.Open "select * from vw_maximo_orden_surtido where vcha_emp_empresa_id ='" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' ", cnn, adOpenDynamic, adLockOptimistic
               If rsaux.EOF Then
                  var_maximo_orden = 1
               Else
                  If IsNull(rsaux!maximo) Then
                     var_maximo_orden = 1
                  Else
                     var_maximo_orden = rsaux!maximo + 1
                  End If
               End If
               rsaux.Close
'Se verifica si el pedido tiene catalogos para venta obligatoria u obsequio
               var_posibles_catalogos = True
               rsaux2.Open "select * from tb_detalle_pedidos where inte_ped_numero = " + Str(rs!inte_ped_numero), cnn, adOpenDynamic, adLockOptimistic
               var_articulos_pedidos = False
               While Not rsaux2.EOF
                     var_tipo_pedido = IIf(IsNull(rsaux2!char_ped_tipo), "", rsaux2!char_ped_tipo)
                     If var_tipo_pedido = "V" Or var_tipo_pedido = "O" Then
                        rsaux.Open "select * from tb_existencias where vcha_alm_almacen_id = '" + var_almacen + "' and vcha_art_articulo_id = '" + rsaux2!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If rsaux.EOF Then
                           var_existen = 0
                           var_apartados = 0
                           var_disponible = 0
                        Else
                           var_existen = IIf(IsNull(rsaux!floa_Exi_Cantidad), 0, rsaux!floa_Exi_Cantidad)
                           var_apartados = IIf(IsNull(rsaux!FLOA_EXI_cANTIDAD_APARTADA), 0, rsaux!FLOA_EXI_cANTIDAD_APARTADA)
                           var_disponible = IIf(IsNull(rsaux!floa_Exi_Cantidad_disponible), 0, rsaux!floa_Exi_Cantidad_disponible)
                        End If
                        rsaux.Close
                        var_cantidad_pedida = IIf(IsNull(rsaux2!FLOA_PED_CANTIDAD), 0, rsaux2!FLOA_PED_CANTIDAD)
                        var_surtir = 0
                        If var_cantidad_pedida > var_disponible Then
                           var_posibles_catalogos = False
                           var_articulos_pedidos = False
                        End If
                     Else
                        var_articulos_pedidos = True
                     End If
                     rsaux2.MoveNext
                Wend
               rsaux2.Close
               If var_posibles_catalogos = True Then
                  var_articulos_pedidos = True
               End If
'termina la verificacion de existencia de catalogos para venta obligatora y de obsequio
               If var_articulos_pedidos = True Then
                  ok = TB_ENC_ORDEN_SURTIDO.Anadir(var_empresa, var_unidad_organizacional, rs!char_tpe_tipo_pedido_id, rs!inte_ped_numero, var_almacen, var_maximo_orden, Date, Date + rs!INTE_PED_DIAS_CADUCIDAD, "", rs!vcha_tit_titular_id, rs!vcha_cli_clave_id, rs!vcha_ESB_ESTABLECIMIENTO_id, rs!floa_ped_descuento_1, rs!floa_ped_descuento_2, rs!floa_ped_Descuento_3, "", "", Date, var_factura_ceros, var_clave_moneda, var_hora)
                  rsaux2.Open "select * from tb_detalle_pedidos where inte_ped_numero = " + Str(rs!inte_ped_numero), cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux2.EOF
                     var_tipo_pedido = IIf(IsNull(rsaux2!char_ped_tipo), "", rsaux2!char_ped_tipo)
                     var_posible = True
                     
                     If var_tipo_pedido = "V" Then
                        If var_posibles_catalogos = False Then
                           var_posible = False
                        End If
                     End If
                     If var_tipo_pedido = "O" Then
                        If var_posibles_catalogos = False Then
                           var_posible = False
                        End If
                     End If
                     If var_tipo_pedido = "P" Then
                        var_posible = True
                     End If
                     If var_posible = True Then
                        var_promocion_1 = 0
                        var_promocion_2 = 0
                        var_promocion_1 = IIf(IsNull(rsaux2!floa_ped_promocion_1), 0, rsaux2!floa_ped_promocion_1)
                        var_promocion_2 = IIf(IsNull(rsaux2!floa_ped_promocion_2), 0, rsaux2!floa_ped_promocion_2)
                        rsaux.Open "select * from tb_existencias where vcha_alm_almacen_id = '" + var_almacen + "' and vcha_art_articulo_id = '" + rsaux2!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If rsaux.EOF Then
                           var_existen = 0
                           var_apartados = 0
                           var_disponible = 0
                        Else
                           var_existen = IIf(IsNull(rsaux!floa_Exi_Cantidad), 0, rsaux!floa_Exi_Cantidad)
                           var_apartados = IIf(IsNull(rsaux!FLOA_EXI_cANTIDAD_APARTADA), 0, rsaux!FLOA_EXI_cANTIDAD_APARTADA)
                           var_disponible = IIf(IsNull(rsaux!floa_Exi_Cantidad_disponible), 0, rsaux!floa_Exi_Cantidad_disponible)
                        End If
                        rsaux.Close
                        var_cantidad_pedida = IIf(IsNull(rsaux2!FLOA_PED_CANTIDAD), 0, rsaux2!FLOA_PED_CANTIDAD)
                        var_surtir = 0
                        If var_cantidad_pedida <= var_disponible Then
                           var_surtir = var_cantidad_pedida
                        Else
                           If var_disponible <= 0 Then
                              var_surtir = 0
                           Else
                              var_surtir = var_disponible
                           End If
                        End If
                        rsaux3.Open "select * from tb_EXISTENCIAS where vcha_alm_almacen_id  = '" + var_almacen + "' and vcha_art_articulo_id = '" + rsaux2!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux3.EOF Then
                           var_costo = IIf(IsNull(rsaux3!FLOA_eXI_COSTO), 0, rsaux3!FLOA_eXI_COSTO)
                        End If
                        rsaux3.Close
                        ok = TB_DET_ORDEN_SURTIDO_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen, var_maximo_orden, rsaux2!vcha_Art_Articulo_id, var_costo, rsaux2!FLOA_PED_PRECIO, var_cantidad_pedida, var_existen, var_apartados, var_disponible, var_surtir, 0, 0, var_promocion_1, var_promocion_2, var_tipo_pedido)
                     End If
                     rsaux2.MoveNext
                  Wend
                  rsaux2.Close
                  ok = TB_ENC_PEDIDOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen, rs!inte_ped_numero, "S")
                  valor = rs!inte_ped_numero
                  Set itmfound = lv_pedidos.findItem(valor, lvwText, , lvwPartial)
                  itmfound.EnsureVisible
                  itmfound.Selected = True
                  bandera_suma = True
                  lv_pedidos.selectedItem.SubItems(1) = var_maximo_orden
                  Set reporte = appl.OpenReport(App.Path + "\rep_orden_surtido.rpt")
                  reporte.RecordSelectionFormula = "{VW_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO} = " + Str(var_maximo_orden)
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  reporte.PrintOut False
                  'frmvistasprevias.cr.ViewReport
                  'frmvistasprevias.Caption = "Orden de Surtido"
                  'frmvistasprevias.Show 1
                  Set reporte = Nothing
               End If
               If var_posibles_catalogos = False Then
                  var_mes = Month(rs!dtim_ped_fecha)
                  var_año = Year(rs!dtim_ped_fecha)
                  If var_mes = 1 Then
                     var_mes = 12
                     var_año = var_año - 1
                  Else
                    var_mes = var_mes - 1
                  End If
                  rsaux.Open "UPDATE TB_CATALOGOS_ASIGNADOS_FACTURACION SET FLOA_FCA_CATALOGOS_VENTA_CLIENTE_APARTADOS = 0 WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_CLI_CLAVE_ID = '" + rs!vcha_cli_clave_id + "' AND INTE_FCA_AÑO = " + CStr(var_año) + " AND INTE_FCA_MES = " + CStr(var_mes), cnn, adOpenDynamic, adLockOptimistic
                  rsaux.Open "UPDATE TB_CATALOGOS_ASIGNADOS_FACTURACION SET FLOA_FCA_CATALOGOS_OBSEQUIO_CLIENTE_APARTADOS = 0 WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_CLI_CLAVE_ID = '" + rs!vcha_cli_clave_id + "' AND INTE_FCA_AÑO = " + CStr(var_año) + " AND INTE_FCA_MES = " + CStr(var_mes), cnn, adOpenDynamic, adLockOptimistic
               End If
            End If
         End If
         rs.MoveNext
      Wend
      rs.Close
      If var_contador > 0 Then
         'MsgBox "Se a terminado la generación de las ordenes de surtido", vbOKOnly, "ATENCION"
      Else
         'MsgBox "No existen pedidos a cargar", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "El proceso de generación de ordenes de surtido a sido cancelado", vbOKOnly, "ATENCION"
   End If
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
   Set TB_ENC_ORDEN_SURTIDO = New TB_ENC_ORDEN_SURTIDO
   Set TB_DET_ORDEN_SURTIDO_I = New TB_DET_ORDEN_SURTIDO_I
   Set TB_ENC_PEDIDOS_M = New TB_ENC_PEDIDOS_M
   Dim var_maximo_orden As Double
   Dim var_existen As Double
   Dim var_apartados As Double
   Dim var_disponible As Double
   Dim var_cantidad_pedidia As Double
   Dim var_surtir As Double
   Dim var_contador As Double
   Dim var_costo As Double
   Dim var_factura_ceros As Double
   Dim i As Double
   Dim j As Double
   frm_imprimir.Visible = True
   txt_numero.SetFocus
End Sub

Private Sub cmd_invertir_Click()
   n = lv_pedidos.ListItems.Count
   For i = 1 To n
      lv_pedidos.ListItems.item(i).Selected = True
      If Me.lv_pedidos.selectedItem.SubItems(17) <> "*" Then
         If Me.lv_pedidos.selectedItem.SubItems(18) <> 1 Then
            If lv_pedidos.selectedItem.SubItems(19) = "*" Then
               lv_pedidos.selectedItem.SubItems(19) = ""
               lv_pedidos.ListItems.item(i).Bold = False
               lv_pedidos.ListItems.item(i).ForeColor = &H80000012
               lv_pedidos.ListItems.item(i).ListSubItems(1).Bold = False
               lv_pedidos.ListItems.item(i).ListSubItems(2).Bold = False
               lv_pedidos.ListItems.item(i).ListSubItems(3).Bold = False
               lv_pedidos.ListItems.item(i).ListSubItems(4).Bold = False
               lv_pedidos.ListItems.item(i).ListSubItems(5).Bold = False
               lv_pedidos.ListItems.item(i).ListSubItems(6).Bold = False
               lv_pedidos.ListItems.item(i).ListSubItems(7).Bold = False
               lv_pedidos.ListItems.item(i).ListSubItems(8).Bold = False
               lv_pedidos.ListItems.item(i).ListSubItems(1).ForeColor = &H80000012
               lv_pedidos.ListItems.item(i).ListSubItems(2).ForeColor = &H80000012
               lv_pedidos.ListItems.item(i).ListSubItems(3).ForeColor = &H80000012
               lv_pedidos.ListItems.item(i).ListSubItems(4).ForeColor = &H80000012
               lv_pedidos.ListItems.item(i).ListSubItems(5).ForeColor = &H80000012
               lv_pedidos.ListItems.item(i).ListSubItems(6).ForeColor = &H80000012
               lv_pedidos.ListItems.item(i).ListSubItems(7).ForeColor = &H80000012
               lv_pedidos.ListItems.item(i).ListSubItems(8).ForeColor = &H80000012
            Else
               lv_pedidos.selectedItem.SubItems(19) = "*"
               lv_pedidos.ListItems.item(i).Bold = True
               lv_pedidos.ListItems.item(i).ForeColor = &HC0&
               lv_pedidos.ListItems.item(i).ListSubItems(1).Bold = True
               lv_pedidos.ListItems.item(i).ListSubItems(2).Bold = True
               lv_pedidos.ListItems.item(i).ListSubItems(3).Bold = True
               lv_pedidos.ListItems.item(i).ListSubItems(4).Bold = True
               lv_pedidos.ListItems.item(i).ListSubItems(5).Bold = True
               lv_pedidos.ListItems.item(i).ListSubItems(6).Bold = True
               lv_pedidos.ListItems.item(i).ListSubItems(7).Bold = True
               lv_pedidos.ListItems.item(i).ListSubItems(8).Bold = True
               lv_pedidos.ListItems.item(i).ListSubItems(1).ForeColor = &HC0&
               lv_pedidos.ListItems.item(i).ListSubItems(2).ForeColor = &HC0&
               lv_pedidos.ListItems.item(i).ListSubItems(3).ForeColor = &HC0&
               lv_pedidos.ListItems.item(i).ListSubItems(4).ForeColor = &HC0&
               lv_pedidos.ListItems.item(i).ListSubItems(5).ForeColor = &HC0&
               lv_pedidos.ListItems.item(i).ListSubItems(6).ForeColor = &HC0&
               lv_pedidos.ListItems.item(i).ListSubItems(7).ForeColor = &HC0&
               lv_pedidos.ListItems.item(i).ListSubItems(8).ForeColor = &HC0&
            End If
         End If
      End If
   Next i
End Sub

Private Sub cmd_marcar_Click()
   i = lv_pedidos.selectedItem.Index
   If Me.lv_pedidos.selectedItem.SubItems(17) <> "*" Then
      If Me.lv_pedidos.selectedItem.SubItems(18) <> 1 Then
         If lv_pedidos.selectedItem.SubItems(19) = "*" Then
            lv_pedidos.selectedItem.SubItems(19) = ""
            lv_pedidos.ListItems.item(i).Bold = False
            lv_pedidos.ListItems.item(i).ForeColor = &H80000012
            lv_pedidos.ListItems.item(i).ListSubItems(1).Bold = False
            lv_pedidos.ListItems.item(i).ListSubItems(2).Bold = False
            lv_pedidos.ListItems.item(i).ListSubItems(3).Bold = False
            lv_pedidos.ListItems.item(i).ListSubItems(4).Bold = False
            lv_pedidos.ListItems.item(i).ListSubItems(5).Bold = False
            lv_pedidos.ListItems.item(i).ListSubItems(6).Bold = False
            lv_pedidos.ListItems.item(i).ListSubItems(7).Bold = False
            lv_pedidos.ListItems.item(i).ListSubItems(8).Bold = False
            lv_pedidos.ListItems.item(i).ListSubItems(1).ForeColor = &H80000012
            lv_pedidos.ListItems.item(i).ListSubItems(2).ForeColor = &H80000012
            lv_pedidos.ListItems.item(i).ListSubItems(3).ForeColor = &H80000012
            lv_pedidos.ListItems.item(i).ListSubItems(4).ForeColor = &H80000012
            lv_pedidos.ListItems.item(i).ListSubItems(5).ForeColor = &H80000012
            lv_pedidos.ListItems.item(i).ListSubItems(6).ForeColor = &H80000012
            lv_pedidos.ListItems.item(i).ListSubItems(7).ForeColor = &H80000012
            lv_pedidos.ListItems.item(i).ListSubItems(8).ForeColor = &H80000012
            lv_pedidos.Refresh
         Else
            lv_pedidos.selectedItem.SubItems(19) = "*"
            lv_pedidos.ListItems.item(i).Bold = True
            lv_pedidos.ListItems.item(i).ForeColor = &HC0&
            lv_pedidos.ListItems.item(i).ListSubItems(1).Bold = True
            lv_pedidos.ListItems.item(i).ListSubItems(2).Bold = True
            lv_pedidos.ListItems.item(i).ListSubItems(3).Bold = True
            lv_pedidos.ListItems.item(i).ListSubItems(4).Bold = True
            lv_pedidos.ListItems.item(i).ListSubItems(5).Bold = True
            lv_pedidos.ListItems.item(i).ListSubItems(6).Bold = True
            lv_pedidos.ListItems.item(i).ListSubItems(7).Bold = True
            lv_pedidos.ListItems.item(i).ListSubItems(8).Bold = True
            lv_pedidos.ListItems.item(i).ListSubItems(1).ForeColor = &HC0&
            lv_pedidos.ListItems.item(i).ListSubItems(2).ForeColor = &HC0&
            lv_pedidos.ListItems.item(i).ListSubItems(3).ForeColor = &HC0&
            lv_pedidos.ListItems.item(i).ListSubItems(4).ForeColor = &HC0&
            lv_pedidos.ListItems.item(i).ListSubItems(5).ForeColor = &HC0&
            lv_pedidos.ListItems.item(i).ListSubItems(6).ForeColor = &HC0&
            lv_pedidos.ListItems.item(i).ListSubItems(7).ForeColor = &HC0&
            lv_pedidos.ListItems.item(i).ListSubItems(8).ForeColor = &HC0&
            lv_pedidos.Refresh
         End If
      End If
   End If
End Sub

Private Sub cmd_ninguno_Click()
   n = lv_pedidos.ListItems.Count
   For i = 1 To n
      lv_pedidos.ListItems.item(i).Selected = True
      If Me.lv_pedidos.selectedItem.SubItems(17) <> "*" Then
         If lv_pedidos.selectedItem.SubItems(18) <> 1 Then
            lv_pedidos.selectedItem.SubItems(19) = ""
            lv_pedidos.ListItems.item(i).Bold = False
            lv_pedidos.ListItems.item(i).ForeColor = &H80000012
            lv_pedidos.ListItems.item(i).ListSubItems(1).Bold = False
            lv_pedidos.ListItems.item(i).ListSubItems(2).Bold = False
            lv_pedidos.ListItems.item(i).ListSubItems(3).Bold = False
            lv_pedidos.ListItems.item(i).ListSubItems(4).Bold = False
            lv_pedidos.ListItems.item(i).ListSubItems(5).Bold = False
            lv_pedidos.ListItems.item(i).ListSubItems(6).Bold = False
            lv_pedidos.ListItems.item(i).ListSubItems(7).Bold = False
            lv_pedidos.ListItems.item(i).ListSubItems(8).Bold = False
            lv_pedidos.ListItems.item(i).ListSubItems(1).ForeColor = &H80000012
            lv_pedidos.ListItems.item(i).ListSubItems(2).ForeColor = &H80000012
            lv_pedidos.ListItems.item(i).ListSubItems(3).ForeColor = &H80000012
            lv_pedidos.ListItems.item(i).ListSubItems(4).ForeColor = &H80000012
            lv_pedidos.ListItems.item(i).ListSubItems(5).ForeColor = &H80000012
            lv_pedidos.ListItems.item(i).ListSubItems(6).ForeColor = &H80000012
            lv_pedidos.ListItems.item(i).ListSubItems(7).ForeColor = &H80000012
            lv_pedidos.ListItems.item(i).ListSubItems(8).ForeColor = &H80000012
         End If
      End If
   Next i
   lv_pedidos.Refresh
End Sub

Private Sub cmd_resurtir_Click()
   frm_pedido_resurtir.Visible = True
   txt_pedido_resurtir.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmd_seleccion_Click()
   n = lv_pedidos.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_pedidos.ListItems.item(i).Selected = True
      If var_encontro = True And lv_pedidos.selectedItem.SubItems(19) = "" And var_rellena = True Then
         If Me.lv_pedidos.selectedItem.SubItems(17) <> "*" Then
            If lv_pedidos.selectedItem.SubItems(18) <> 1 Then
               lv_pedidos.selectedItem.SubItems(19) = "*"
               lv_pedidos.ListItems.item(i).Bold = True
               lv_pedidos.ListItems.item(i).ForeColor = &HC0&
               lv_pedidos.ListItems.item(i).ListSubItems(1).Bold = True
               lv_pedidos.ListItems.item(i).ListSubItems(2).Bold = True
               lv_pedidos.ListItems.item(i).ListSubItems(3).Bold = True
               lv_pedidos.ListItems.item(i).ListSubItems(4).Bold = True
               lv_pedidos.ListItems.item(i).ListSubItems(5).Bold = True
               lv_pedidos.ListItems.item(i).ListSubItems(6).Bold = True
               lv_pedidos.ListItems.item(i).ListSubItems(7).Bold = True
               lv_pedidos.ListItems.item(i).ListSubItems(8).Bold = True
               lv_pedidos.ListItems.item(i).ListSubItems(1).ForeColor = &HC0&
               lv_pedidos.ListItems.item(i).ListSubItems(2).ForeColor = &HC0&
               lv_pedidos.ListItems.item(i).ListSubItems(3).ForeColor = &HC0&
               lv_pedidos.ListItems.item(i).ListSubItems(4).ForeColor = &HC0&
               lv_pedidos.ListItems.item(i).ListSubItems(5).ForeColor = &HC0&
               lv_pedidos.ListItems.item(i).ListSubItems(6).ForeColor = &HC0&
               lv_pedidos.ListItems.item(i).ListSubItems(7).ForeColor = &HC0&
               lv_pedidos.ListItems.item(i).ListSubItems(8).ForeColor = &HC0&
            End If
         End If
      Else
         If var_encontro = True And lv_pedidos.selectedItem.SubItems(19) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_pedidos.selectedItem.SubItems(19) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub cmd_todos_Click()
   n = lv_pedidos.ListItems.Count
   For i = 1 To n
      lv_pedidos.ListItems.item(i).Selected = True
      If Me.lv_pedidos.selectedItem.SubItems(17) <> "*" Then
         If Me.lv_pedidos.selectedItem.SubItems(18) <> 1 Then
            lv_pedidos.selectedItem.SubItems(19) = "*"
            lv_pedidos.ListItems.item(i).Bold = True
            lv_pedidos.ListItems.item(i).ForeColor = &HC0&
            lv_pedidos.ListItems.item(i).ListSubItems(1).Bold = True
            lv_pedidos.ListItems.item(i).ListSubItems(2).Bold = True
            lv_pedidos.ListItems.item(i).ListSubItems(3).Bold = True
            lv_pedidos.ListItems.item(i).ListSubItems(4).Bold = True
            lv_pedidos.ListItems.item(i).ListSubItems(5).Bold = True
            lv_pedidos.ListItems.item(i).ListSubItems(6).Bold = True
            lv_pedidos.ListItems.item(i).ListSubItems(7).Bold = True
            lv_pedidos.ListItems.item(i).ListSubItems(8).Bold = True
            lv_pedidos.ListItems.item(i).ListSubItems(1).ForeColor = &HC0&
            lv_pedidos.ListItems.item(i).ListSubItems(2).ForeColor = &HC0&
            lv_pedidos.ListItems.item(i).ListSubItems(3).ForeColor = &HC0&
            lv_pedidos.ListItems.item(i).ListSubItems(4).ForeColor = &HC0&
            lv_pedidos.ListItems.item(i).ListSubItems(5).ForeColor = &HC0&
            lv_pedidos.ListItems.item(i).ListSubItems(6).ForeColor = &HC0&
            lv_pedidos.ListItems.item(i).ListSubItems(7).ForeColor = &HC0&
            lv_pedidos.ListItems.item(i).ListSubItems(8).ForeColor = &HC0&
         End If
      End If
   Next i
   lv_pedidos.Refresh
End Sub

Private Sub Command1_Click()
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
   Set TB_ENC_ORDEN_SURTIDO = New TB_ENC_ORDEN_SURTIDO
   Set TB_DET_ORDEN_SURTIDO_I = New TB_DET_ORDEN_SURTIDO_I
   Set TB_ENC_PEDIDOS_M = New TB_ENC_PEDIDOS_M
   Dim var_maximo_orden As Double
   Dim var_existen As Double
   Dim var_apartados As Double
   Dim var_disponible As Double
   Dim var_cantidad_pedidia As Double
   Dim var_surtir As Double
   Dim var_contador As Double
   Dim var_costo As Double
   Dim var_factura_ceros As Double
   Dim var_clave_moneda As String
   Dim var_promocion_1 As Double
   Dim var_promocion_2 As Double
   Dim var_hora As String
   Dim var_tipo_pedido As String
   Dim var_posibles_catalogos As Boolean
   Dim var_articulos_pedidos As Boolean
   Dim var_posible As Boolean
   Dim var_numero_pedido As Double
   Dim i As Double
   Dim j As Double
   If Me.lv_pedidos.ListItems.Count > 0 Then
      var_cadena = ""
      For var_j = 1 To lv_pedidos.ListItems.Count
          lv_pedidos.ListItems.item(var_j).Selected = True
          If lv_pedidos.selectedItem.SubItems(17) = "*" Then
             If Me.lv_pedidos.selectedItem.SubItems(18) = 1 Or Me.lv_pedidos.selectedItem.SubItems(19) = "*" Then
             Else
                If var_cadena = "" Then
                   var_cadena = var_cadena + Me.lv_pedidos.selectedItem
                Else
                   var_cadena = var_cadena + ", " + Me.lv_pedidos.selectedItem
                End If
             End If
          End If
      Next var_j
      
      var_cadena = ""
      If var_cadena = "" Then
         si = MsgBox("¿Deseas ejecutar el proceso de ordenes de surtido", vbYesNo, "ATENCION")
         If si = 6 Then
            For var_i = 1 To lv_pedidos.ListItems.Count
                lv_pedidos.ListItems.item(var_i).Selected = True
                If Trim(lv_pedidos.selectedItem.SubItems(1)) = "" Then
                   var_factura_ceros = 0
                   If lv_pedidos.selectedItem.SubItems(17) = "*" Then
                      var_factura_ceros = 1
                   End If
                   Dim cmd As New Command
                   var_numero_pedido = lv_pedidos.selectedItem
                   Set cmd.ActiveConnection = cnn
                   cmd.CommandType = adCmdStoredProc
                   If var_empresa = "30" Then
                      cmd.CommandText = "SP_ORDENES_SURTIDO_TURBINA"
                   Else
                      cmd.CommandText = "SP_ORDENES_SURTIDO"
                   End If
                
                   cmd("@VAR_EMPRESA") = var_empresa
                   cmd("@VAR_UNIDAD_ORGANIZACIONAL") = var_unidad_organizacional
                   cmd("@SERIE") = ""
                   cmd("@USUARIO") = var_clave_usuario_global
                   cmd("@MAQUINA") = ""
                   cmd("@VAR_FACTURA_CEROS") = var_factura_ceros
                   cmd("@VAR_NUMERO_PEDIDO") = var_numero_pedido
                   cmd("@VAR_MAXIMO_ORDEN") = 0
                   cmd.execute
                   var_numero_orden = cmd("@VAR_MAXIMO_ORDEN")
                   lv_pedidos.selectedItem.SubItems(1) = var_numero_orden
                   Set cmd = Nothing
                   If var_empresa <> "31" Then
                      If var_numero_orden > 0 Then
                         cnn.BeginTrans
                         rs.Open "select max(INTE_TEM_CONSECUTIVO) as consecutivo from tb_temp_orden_surtido", cnn, adOpenDynamic, adLockOptimistic
                         If Not rs.EOF Then
                            var_consecutivo = IIf(IsNull(rs!consecutivo), 0, rs!consecutivo) + 1
                         Else
                            var_consecutivo = 1
                         End If
                         rs.Close
                         rs.Open "insert into tb_temp_orden_surtido (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                         cnn.CommitTrans
                         Cadena = "INSERT INTO [TB_TEMP_ORDEN_SURTIDO] ([INTE_TEM_CONSECUTIVO], [CHAR_TPE_TIPO_PEDIDO], [INTE_PED_NUMERO], [VCHA_ALM_ALMACEN_ID], [INTE_ORS_ORDEN_SURTIDO], [DTIM_ORS_FECHA_CARGA], [DTIM_ORS_FECHA_CADUCA], [VCHA_ESB_NOMBRE], [VCHA_ESB_ESTABLECIMIENTO_ID], [VCHA_CLI_CLAVE_ID], [VCHA_CLI_NOMBRE], "
                         Cadena = Cadena + "[CHAR_PRI_PRIORIDAD_ID], [FLOA_ORS_DESCUENTO_1], [FLOA_ORS_DESCUENTO_2], [VCHA_ART_ARTICULO_ID], [VCHA_ART_NOMBRE_ESPAÑOL], [FLOA_ORS_PRECIO], [FLOA_ORS_CANTIDAD_SURTIR], [VCHA_AGE_NOMBRE], [VCHA_RUT_NOMBRE], [INTE_PED_DIAS_CONDICIONES], [INTE_PED_DIAS_CADUCIDAD], [MONE_ART_COSTO_ESTANDAR], [FLOA_ORS_CANTIDAD_SURTIDA], "
                         Cadena = Cadena + "[VCHA_TIT_TITULAR_ID], [VCHA_TIT_NOMBRE], [VCHA_RUT_RUTA_ID], [VCHA_AGE_AGENTE_ID], [CHAR_ORS_ESTATUS], [VCHA_MOV_MOVIMIENTO_ID], [VCHA_CLI_EMAIL], [FLOA_TPE_IVA], [ALMACEN_AGENTE], [INTE_TCL_EMPACADO], [VCHA_TCL_ALMACEN_EMPAQUE], [FLOA_ORS_CANTIDAD_EMPACADA], [VCHA_PLA_PAZO_ID], [INTE_CLI_AGRUPADOR], [VCHA_MON_MONEDA_ID],"
                         Cadena = Cadena + "[VCHA_ALM_NOMBRE], [FLOA_ORS_PROMOCION_1], [FLOA_ORS_PROMOCION_2], [INTE_ORS_FACTURA_CEROS], [CHAR_PED_TIPO], [VCHA_UOR_UNIDAD_ID], [INTE_PED_REFERENCIA],  [VCHA_EXI_UBICACION], [FLOA_ORS_CANTIDAD_NEGADA], [VCHA_TEM_CODIGO_BARRAS]) "
                         Cadena = Cadena + " select " + CStr(var_consecutivo) + ", CHAR_TPE_TIPO_PEDIDO_ID, INTE_PED_NUMERO, VCHA_ALM_ALMACEN_ID, INTE_ORS_ORDEN_SURTIDO, DTIM_ORS_FECHA_CARGA, DTIM_ORS_FECHA_CADUCA, VCHA_ESB_NOMBRE, VCHA_ESB_ESTABLECIMIENTO_ID, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, CHAR_PRI_PRIORIDAD_ID, FLOA_ORS_DESCUENTO_1, FLOA_ORS_DESCUENTO_2, VCHA_ART_ARTICULO_ID, "
                         Cadena = Cadena + " substring(VCHA_ART_NOMBRE_ESPAÑOL,1,50), FLOA_ORS_PRECIO, FLOA_ORS_CANTIDAD_SURTIR, VCHA_AGE_NOMBRE, VCHA_RUT_NOMBRE, INTE_PED_DIAS_CONDICIONES, INTE_PED_DIAS_CADUCIDAD, MONE_ART_COSTO_ESTANDAR, FLOA_ORS_CANTIDAD_SURTIDA, VCHA_TIT_TITULAR_ID, VCHA_TIT_NOMBRE, VCHA_RUT_RUTA_ID, VCHA_AGE_AGENTE_ID, CHAR_ORS_ESTATUS, VCHA_MOV_MOVIMIENTO_ID, VCHA_CLI_EMAIL, FLOA_TPE_IVA,"
                         Cadena = Cadena + " ALMACEN_AGENTE, INTE_TCL_EMPACADO, VCHA_TCL_ALMACEN_EMPAQUE, FLOA_ORS_CANTIDAD_EMPACADA, VCHA_PLA_PlAZO_ID, INTE_CLI_AGRUPADOR, VCHA_MON_MONEDA_ID, VCHA_ALM_NOMBRE, FLOA_ORS_PROMOCION_1, FLOA_ORS_PROMOCION_2, INTE_ORS_FACTURA_CEROS, CHAR_PED_TIPO, VCHA_UOR_UNIDAD_ID, INTE_PED_REFERENCIA, VCHA_EXI_UBICACION, FLOA_ORS_CANTIDAD_NEGADA, " + Format(var_numero_orden, "##########") + " from vw_orden_surtido where inte_ors_orden_surtido = " + CStr(var_numero_orden)
                         rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                         If var_empresa = "18" Then
                            Set reporte = appl.OpenReport(App.Path + "\rep_orden_surtido_nueva_textilera.rpt")
                         Else
                            Set reporte = appl.OpenReport(App.Path + "\rep_orden_surtido_nueva.rpt")
                         End If
                         reporte.RecordSelectionFormula = "{TB_TEMP_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO} = " + Str(var_numero_orden) + " AND {TB_TEMP_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND {TB_TEMP_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIR} > 0"
                         frmvistasprevias.cr.ReportSource = reporte
                         For ntablas = 1 To reporte.Database.Tables.Count
                             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                         Next ntablas
                         reporte.PrintOut False
                         Set reporte = Nothing
                         rs.Open "delete from tb_temp_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                      Else
                         MsgBox "No se pudo generar la orden de surtido del pedido " + Trim(CStr(var_numero_pedido)), vbOKOnly, "ATENCION"
                      End If
                   Else
                   
                      var_numero_pedido = var_numero_orden
                      rs.Open "delete from TB_PEDIDOS_CANTIA where inte_ped_numero = " + CStr(var_numero_orden), cnn, adOpenDynamic, adLockOptimistic
                      rsaux2.Open "select vcha_Art_articulo_id as codigo, floa_ors_Cantidad_surtir as pedido, floa_ors_existen from tb_Det_orden_surtido where inte_ors_orden_surtido = " + CStr(var_numero_orden) + " and floa_ors_Cantidad_surtir > 0", cnn, adOpenDynamic, adLockOptimistic
                      While Not rsaux2.EOF
                            If Not IsNull(rsaux2!codigo) Then
                               If rsaux2!pedido > 0 Then
                                  rsaux3.Open "insert into TB_PEDIDOS_CANTIA (INTE_PED_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_PED_CANTIDAD_PEDIDA, DTIM_PED_FECHA, floa_Exi_cantidad) values (" + CStr(var_numero_pedido) + ",'" + rsaux2!codigo + "'," + CStr(rsaux2!pedido) + ",GETDATE()," + CStr(rsaux2!floa_ors_existen) + ")", cnn, adOpenDynamic, adLockOptimistic
                               End If
                            End If
                            rsaux2.MoveNext
                      Wend
                      rsaux2.Close
         
                      rsaux2.Open "SELECT * FROM TB_PEDIDOS_CANTIA WHERE INTE_PED_NUMERO = " + CStr(var_numero_pedido), cnn, adOpenDynamic, adLockOptimistic
                      While Not rsaux2.EOF
                            If rsaux3.State = 1 Then
                               rsaux3.Close
                            End If
                            'rsaux3.Open "SELECT * FROM TB_EXISTENCIAS WHERE VCHA_ALM_ALMACEN_ID = 'PTVH' AND VCHA_aRT_ARTICULO_ID = '" + rsaux2!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                            'If Not rsaux3.EOF Then
                            '   rsaux4.Open "UPDATE TB_PEDIDOS_CANTIA SET FLOA_EXI_CANTIDAD = " + CStr(IIf(IsNull(rsaux3!floa_Exi_Cantidad), 0, rsaux3!floa_Exi_Cantidad)) + " WHERE INTE_PED_NUMERO = " + CStr(var_numero_pedido) + " AND VCHA_ART_ARTICULO_ID = '" + rsaux3!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                            'End If
                            'rsaux3.Close
                             rsaux3.Open "SELECT * FROM TB_CODIGOS_PROVEEDOR_CANTIA WHERE VCHA_ART_ARTICULO_ID = '" + rsaux2!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                             If Not rsaux3.EOF Then
                                rsaux4.Open "UPDATE TB_PEDIDOS_CANTIA SET VCHA_COD_CODIGO_BARRAS = '" + IIf(IsNull(rsaux3!VCHA_COD_CODIGO_BARRAS), "", rsaux3!VCHA_COD_CODIGO_BARRAS) + "', VCHA_COD_CODIGO_PROVEEDOR = '" + IIf(IsNull(rsaux3!VCHA_COD_CODIGO_PROVEEDOR), "", rsaux3!VCHA_COD_CODIGO_PROVEEDOR) + "', VCHA_COD_NOMBRE_PROVEEDOR = '" + IIf(IsNull(rsaux3!VCHA_COD_NOMBRE_PROVEEDOR), "", rsaux3!VCHA_COD_NOMBRE_PROVEEDOR) + "' WHERE INTE_PED_NUMERO = " + CStr(var_numero_pedido) + " AND VCHA_ART_ARTICULO_ID = '" + rsaux3!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                             End If
                             rsaux3.Close
                             If Mid(rsaux2!vcha_Art_Articulo_id, 1, 1) = "T" Then
                                rsaux3.Open "select * from tb_equivalencias where vcha_Art_articulo_id = '" + rsaux2!vcha_Art_Articulo_id + "' and substring(vcha_equ_codigo_equivalente,1,5) = '64624'", cnn, adOpenDynamic, adLockOptimistic
                                If Not rsaux3.EOF Then
                                   rsaux4.Open "UPDATE TB_PEDIDOS_CANTIA SET VCHA_COD_CODIGO_BARRAS = '" + IIf(IsNull(rsaux3!vcha_equ_codigo_equivalente), "", rsaux3!vcha_equ_codigo_equivalente) + "' WHERE INTE_PED_NUMERO = " + CStr(var_numero_pedido) + " AND VCHA_ART_ARTICULO_ID = '" + rsaux3!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                End If
                                rsaux3.Close
                             End If
                             rsaux2.MoveNext
                      Wend
                      rsaux2.Close
                      rsaux2.Open "SELECT INTE_PED_NUMERO FROM TB_ENC_ORDEN_SURTIDO WHERE INTE_ORS_ORDEN_SURTIDO = " + CStr(var_numero_orden), cnn, adOpenDynamic, adLockOptimistic
                      If Not rsaux2.EOF Then
                         rsaux3.Open "SELECT * FROM TB_ENCABEZADO_PEDIDOS WHERE INTE_PED_NUMERO = " + CStr(IIf(IsNull(rsaux2!inte_ped_numero), 0, rsaux2!inte_ped_numero)), cnn, adOpenDynamic, adLockOptimistic
                         If Not rsaux3.EOF Then
                            rsaux4.Open "UPDATE TB_PEDIDOS_CANTIA SET VCHA_PED_LOCALIZACION = '" + IIf(IsNull(rsaux3!VCHA_PED_PEDIDO_EXTERNO), "", rsaux3!VCHA_PED_PEDIDO_EXTERNO) + "' WHERE INTE_PED_NUMERO = " + CStr(var_numero_orden), cnn, adOpenDynamic, adLockOptimistic
                         End If
                         rsaux3.Close
                      End If
                      rsaux2.Close
         
                      Set reporte = appl.OpenReport(App.Path + "\rep_pedido_Cantia.rpt")
                      frmvistasprevias.cr.ReportSource = reporte
                      reporte.RecordSelectionFormula = "{VW_PEDIDO_UBICACIONES.INTE_PED_NUMERO} = " + CStr(var_numero_orden)
                      frmvistasprevias.cr.ReportSource = reporte
                      For ntablas = 1 To reporte.Database.Tables.Count
                          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                      Next ntablas
                      reporte.PrintOut False
                      Set reporte = Nothing
                   
                   End If
                End If
           Next var_i
         Else
            MsgBox "El proceso de generación de ordenes de surtido a sido cancelado", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Los pedidos " + var_cadena + " estan marcados en facturación en ceros", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No existen pedidos", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Command2_Click()
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If IsNumeric(txt_pedido_resurtir) Then
         rs.Open "select * from vw_ordenes_surtido_vivas where inte_ped_numero = " + txt_pedido_resurtir, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            rs.Close
            MsgBox "El pedido número " + Trim(txt_pedido_resurtir) + " tiene todavia ordenes de surtido vivas por lo que no podra ser resurtido aun", vbOKOnly, "ATENCION"
         Else
            rs.Close
            Set TB_ENC_ORDEN_SURTIDO = New TB_ENC_ORDEN_SURTIDO
            Set TB_DET_ORDEN_SURTIDO_I = New TB_DET_ORDEN_SURTIDO_I
            Set TB_ENC_PEDIDOS_M = New TB_ENC_PEDIDOS_M
            Dim var_maximo_orden As Double
            Dim var_existen As Double
            Dim var_apartados As Double
            Dim var_disponible As Double
            Dim var_cantidad_pedidia As Double
            Dim var_surtir As Double
            Dim var_contador As Double
            Dim var_costo As Double
            Dim var_factura_ceros As Double
            Dim var_clave_moneda As String
            Dim var_promocion_1 As Double
            Dim var_promocion_2 As Double
            Dim var_diferencia As Double
            Dim var_hora  As String
            Dim i As Double
            Dim j As Double
            Dim var_tipo_pedido As String
            Dim var_orden_surtido_resurtir As Double
            si = MsgBox("¿Deseas ejecutar el proceso de ordenes de surtido", vbYesNo, "ATENCION")
            If si = 6 Then
               var_contador = 0
               rs.Open "select max(inte_ors_orden_surtido) from tb_enc_orden_surtido where (char_ped_estatus <> 'C' or char_ped_estatus <> 'S') and  inte_ped_numero = " + txt_pedido_resurtir, cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_orden_surtido_resurtir = IIf(IsNull(rs!INTE_ORS_ORDEN_SURTIDO), 0, rs!INTE_ORS_ORDEN_SURTIDO)
               Else
                  var_orden_surtido_resurtir = 0
               End If
               'rs.Open "select * from vw_resurtido where (char_ped_estatus <> 'C' or char_ped_estatus <> 'S') and inte_ped_numero = " + txt_pedido_resurtir + " and diferencia > 0", cnn, adOpenDynamic, adLockOptimistic
               var_hora = CStr(Time)
               If Not rs.EOF Then
                  rs.Close
                  If var_orden_surtido_resurtir > 0 Then
                     rs.Open "select * from tb_encabezado_pedidos where inte_ped_numero = " + txt_pedido_resurtir, cnn, adOpenDynamic, adLockOptimistic
                     var_tipo_pedido = IIf(IsNull(rs!char_tpe_tipo_pedido_id), "", rs!char_tpe_tipo_pedido_id)
                     var_resurtible = IIf(IsNull(rs!inte_ped_resurtible), 0, rs!inte_ped_resurtible)
                     If var_resurtible = 1 Then
                        var_clave_moneda = rs!VCHA_MON_MONEDA_ID
                        var_factura_ceros = 0
                        var_almacen = IIf(IsNull(rs!VCHA_ALM_ALMACEN_ID), "", rs!VCHA_ALM_ALMACEN_ID)
                        If Trim(var_almacen) <> "" Then
                           If IsNull(rs!inte_ped_autorizo) Then
                              var_a = 0
                           Else
                              var_a = rs!inte_ped_autorizo
                           End If
                           If var_a = 1 And Trim(rs!char_ped_estatus) = "S" Then
                              rsaux.Open "update tb_detalle_pedidos set floa_ped_cantidad_depurada = 0 where inte_ped_numero = " + txt_pedido_resurtir, cnn, adOpenDynamic, adLockOptimistic
                              var_contador = var_contador + 1
                              rsaux.Open "select * from vw_maximo_orden_surtido ", cnn, adOpenDynamic, adLockOptimistic
                              If rsaux.EOF Then
                                 var_maximo_orden = 1
                              Else
                                 If IsNull(rsaux!maximo) Then
                                    var_maximo_orden = 1
                                 Else
                                    var_maximo_orden = rsaux!maximo + 1
                                 End If
                              End If
                              rsaux.Close
                              ok = TB_ENC_ORDEN_SURTIDO.Anadir(var_empresa, var_unidad_organizacional, rs!char_tpe_tipo_pedido_id, rs!inte_ped_numero, var_almacen, var_maximo_orden, Date, Date + rs!INTE_PED_DIAS_CADUCIDAD, "", rs!vcha_tit_titular_id, rs!vcha_cli_clave_id, rs!vcha_ESB_ESTABLECIMIENTO_id, rs!floa_ped_descuento_1, rs!floa_ped_descuento_2, rs!floa_ped_Descuento_3, "", "", Date, var_factura_ceros, var_clave_moneda, var_hora)
                              rsaux2.Open "select * from VW_RESURTIDO_NEGADO_PRODUCCION where inte_ped_numero = " + Str(rs!inte_ped_numero) + " AND INTE_ORS_ORDEN_SURTIDO = " + CStr(var_orden_surtido_resurtir), cnn, adOpenDynamic, adLockOptimistic
                              While Not rsaux2.EOF
                                    If (rs!FLOA_PED_CANTIDAD - rs!cantidad_empacada) > 0 Then
                                       var_promocion_1 = 0
                                       var_promocion_2 = 0
                                       var_promocion_1 = IIf(IsNull(rsaux2!floa_ped_promocion_1), 0, rsaux2!floa_ped_promocion_1)
                                       var_promocion_2 = IIf(IsNull(rsaux2!floa_ped_promocion_2), 0, rsaux2!floa_ped_promocion_2)
                                       rsaux.Open "select * from tb_existencias where vcha_alm_almacen_id = '8' and vcha_art_articulo_id = '" + rsaux2!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                       If rsaux.EOF Then
                                          var_existen = 0
                                          var_apartados = 0
                                          var_disponible = 0
                                       Else
                                          If IsNull(rsaux!floa_Exi_Cantidad) Then
                                             var_existen = 0
                                          Else
                                             var_existen = rsaux!floa_Exi_Cantidad
                                          End If
                                          If IsNull(rsaux!FLOA_EXI_cANTIDAD_APARTADA) Then
                                             var_apartados = 0
                                          Else
                                             var_apartados = rsaux!FLOA_EXI_cANTIDAD_APARTADA
                                          End If
                                          If IsNull(rsaux!floa_Exi_Cantidad_disponible) Then
                                             var_disponible = 0
                                          Else
                                             var_disponible = rsaux!floa_Exi_Cantidad_disponible
                                          End If
                                       End If
                                       rsaux.Close
                                       var_cantidad_pedida = IIf(IsNull(rsaux2!floa_ors_cantidad_pedida), 0, rsaux2!floa_ors_cantidad_pedida) - IIf(IsNull(rsaux2!floa_ped_cantidad_surtir), 0, rsaux2!floa_ped_cantidad_surtir)
                                       var_surtir = 0
                                       If var_cantidad_pedida > 0 Then
                                          If var_cantidad_pedida <= var_disponible Then
                                             var_surtir = var_cantidad_pedida
                                          Else
                                             If var_disponible <= 0 Then
                                                var_surtir = 0
                                             Else
                                                var_surtir = var_disponible
                                             End If
                                          End If
                                          rsaux3.Open "select * from tb_EXISTENCIAS where vcha_alm_almacen_id  = '" + var_almacen + "' and vcha_art_articulo_id = '" + rsaux2!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                          If Not rsaux3.EOF Then
                                             var_costo = IIf(IsNull(rsaux3!FLOA_eXI_COSTO), 0, rsaux3!FLOA_eXI_COSTO)
                                          End If
                                          rsaux3.Close
                                          ok = TB_DET_ORDEN_SURTIDO_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen, var_maximo_orden, rsaux2!vcha_Art_Articulo_id, var_costo, rsaux2!FLOA_PED_PRECIO, var_cantidad_pedida, var_existen, var_apartados, var_disponible, var_surtir, 0, 0, var_promocion_1, var_promocion_2, var_tipo_pedido)
                                       End If
                                    End If
                                    rsaux2.MoveNext
                               Wend
                               rsaux2.Close
                               ok = TB_ENC_PEDIDOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen, rs!inte_ped_numero, "S")
                               valor = rs!inte_ped_numero
                               Set itmfound = lv_pedidos.findItem(valor, lvwText, , lvwPartial)
                               Set reporte = appl.OpenReport(App.Path + "\rep_orden_surtido.rpt")
                               reporte.RecordSelectionFormula = "{VW_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO} = " + Str(var_maximo_orden)
                               frmvistasprevias.cr.ReportSource = reporte
                               For ntablas = 1 To reporte.Database.Tables.Count
                                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                               Next ntablas
                               frmvistasprevias.cr.ViewReport
                               frmvistasprevias.Caption = "Orden de Surtido"
                               frmvistasprevias.Show 1
                               Set reporte = Nothing
                            Else
                               If rs!char_ped_estatus = "E" Then
                                  MsgBox "El pedido ya no puede ser resurtido ya que ya fue cerrado", vbOKOnly, "ATENCION"
                               End If
                               If var_a = 0 Then
                                  MsgBox "El pedido no puede ser resurtido ya que no fue autorizado", vbOKOnly, "ATENCION"
                               End If
                            End If
                         End If
                     Else
                        MsgBox "El pedido no es resurtible", vbOKOnly, "ATENCION"
                     End If
                  Else
                  End If
               Else
                  rs.Close
                  MsgBox "El pedido ya no puede ser resurtido", vbOKOnly, "ATENCION"
               End If
               If var_contador > 0 Then
                  'MsgBox "Se a terminado la generación de las ordenes de surtido", vbOKOnly, "ATENCION"
               Else
                  'MsgBox "No existen pedidos a cargar", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "El proceso de generación de ordenes de surtido a sido cancelado", vbOKOnly, "ATENCION"
            End If
         End If
      End If
      frm_pedido_resurtir.Visible = False
   End If

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 71 Then
      cmd_generar_Click
   End If
   If Shift = 4 And KeyCode = 73 Then
      cmd_imprimir_Click
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 0
   Left = 0
   frm_pedido_resurtir.Visible = False
   frm_imprimir.Visible = False
   Me.frm_factura_ceros.Visible = False
   var_fecha_inicio = Date
   var_decha_fin = Date
   cnn.CommandTimeout = 360
   If var_pedido_internet = 0 Then
      rsaux10.Open "select * from vw_suma_pedidos where (char_ped_estatus <> 'C' or char_ped_estatus <> 'S') and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'  AND VCHA_ALM_ALMACEN_ID <> 'AVIN' order by CHAR_PRI_PRIORIDAD_ID,inte_ped_numero", cnn, adOpenDynamic, adLockOptimistic
   Else
      rsaux10.Open "select * from vw_suma_pedidos where (char_ped_estatus <> 'C' or char_ped_estatus <> 'S') and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'  AND VCHA_ALM_ALMACEN_ID = 'AVIN' and inte_ped_numero  = " + CStr(var_pedido_internet) + " order by CHAR_PRI_PRIORIDAD_ID,inte_ped_numero", cnn, adOpenDynamic, adLockOptimistic
   End If
   lv_pedidos.SmallIcons = ImageList1
   lv_pedidos.ListItems.Clear
   Dim list_item As ListItem
   While Not rsaux10.EOF
      If IsNull(rsaux10!inte_ped_autorizo) Then
         var_a = 0
      Else
         var_a = rsaux10!inte_ped_autorizo
      End If
      If var_a = 1 And (Len(Trim(rsaux10!char_ped_estatus)) = 0 Or rsaux10!char_ped_estatus = "I") Then
         Set list_item = lv_pedidos.ListItems.Add(, , rsaux10!inte_ped_numero)
         list_item.SubItems(1) = ""
         list_item.SubItems(2) = IIf(IsNull(rsaux10!VCHA_AGE_NOMBRE), "", rsaux10!VCHA_AGE_NOMBRE)
         list_item.SubItems(3) = IIf(IsNull(rsaux10!VCHA_TIT_NOMBRE), "", rsaux10!VCHA_TIT_NOMBRE)
         list_item.SubItems(4) = IIf(IsNull(rsaux10!VCHA_ESB_NOMBRE), "", rsaux10!VCHA_ESB_NOMBRE)
         list_item.SubItems(5) = IIf(IsNull(rsaux10!VCHA_CLI_NOMBRE), "", rsaux10!VCHA_CLI_NOMBRE)
         If IsNull(rsaux10!Cantidad) Then
            list_item.SubItems(6) = Format(0, "###,###,##0.00")
         Else
            list_item.SubItems(6) = Format(rsaux10!Cantidad, "###,###,##0.00")
         End If
         If IsNull(rsaux10!Importe) Then
            list_item.SubItems(7) = Format(0, "###,###,##0.00")
         Else
            list_item.SubItems(7) = Format(rsaux10!Importe, "###,###,##0.00")
         End If
         If IsNull(rsaux10!inte_ped_autorizo) Then
            list_item.SubItems(8) = 0
         Else
            list_item.SubItems(8) = rsaux10!inte_ped_autorizo
            If rsaux10!inte_ped_autorizo = 2 Then
               list_item.SmallIcon = 13
               list_item.Bold = True
               list_item.ForeColor = &H8000&
               list_item.ListSubItems.item(1).Bold = True
               list_item.ListSubItems.item(2).Bold = True
               list_item.ListSubItems.item(3).Bold = True
               list_item.ListSubItems.item(4).Bold = True
               list_item.ListSubItems.item(5).Bold = True
               list_item.ListSubItems.item(6).Bold = True
               list_item.ListSubItems.item(1).ForeColor = &H8000&
               list_item.ListSubItems.item(2).ForeColor = &H8000&
               list_item.ListSubItems.item(3).ForeColor = &H8000&
               list_item.ListSubItems.item(4).ForeColor = &H8000&
               list_item.ListSubItems.item(5).ForeColor = &H8000&
               list_item.ListSubItems.item(6).ForeColor = &H8000&
            End If
         End If
         If IsNull(rsaux10!VCHA_PED_AUTORIZO) Then
            list_item.SubItems(9) = ""
         Else
            list_item.SubItems(9) = rsaux10!VCHA_PED_AUTORIZO
         End If
         If IsNull(rsaux10!DTIM_PED_AUTORIZO) Then
            list_item.SubItems(10) = ""
         Else
            list_item.SubItems(10) = rsaux10!DTIM_PED_AUTORIZO
         End If
         If IsNull(rsaux10!dtim_ped_fecha) Then
            list_item.SubItems(11) = ""
         Else
            list_item.SubItems(11) = rsaux10!dtim_ped_fecha
         End If
         If IsNull(rsaux10!char_ped_estatus) Then
            list_item.SubItems(12) = ""
         Else
            list_item.SubItems(12) = rsaux10!char_ped_estatus
         End If
         If IsNull(rsaux10!floa_ped_descuento_1) Then
            list_item.SubItems(13) = 0
         Else
            list_item.SubItems(13) = rsaux10!floa_ped_descuento_1
         End If
         If IsNull(rsaux10!floa_ped_descuento_2) Then
            list_item.SubItems(14) = 0
         Else
            list_item.SubItems(14) = rsaux10!floa_ped_descuento_2
         End If
         If IsNull(rsaux10!VCHA_USU_NOMBRE) Then
            list_item.SubItems(15) = ""
         Else
            If IsNull(rsaux10!vcha_usu_apellidos) Then
               list_item.SubItems(15) = rsaux10!VCHA_USU_NOMBRE
            Else
               list_item.SubItems(15) = Trim(rsaux10!VCHA_USU_NOMBRE) + " " + Trim(rsaux10!vcha_usu_apellidos)
            End If
         End If
         list_item.SubItems(16) = IIf(IsNull(rsaux10!VCHA_ALM_ALMACEN_ID), "", rsaux10!VCHA_ALM_ALMACEN_ID)
         list_item.SubItems(18) = IIf(IsNull(rsaux10!inte_ped_factura_ceros), 0, rsaux10!inte_ped_factura_ceros)
         var_pedido_factura_ceros = IIf(IsNull(rsaux10!inte_ped_factura_ceros), 0, rsaux10!inte_ped_factura_ceros)
         If var_pedido_factura_ceros = 1 Then
            list_item.SubItems(17) = "*"
            list_item.Bold = True
            list_item.ListSubItems.item(1).Bold = True
            list_item.ListSubItems.item(2).Bold = True
            list_item.ListSubItems.item(3).Bold = True
            list_item.ListSubItems.item(4).Bold = True
            list_item.ListSubItems.item(5).Bold = True
            list_item.ListSubItems.item(6).Bold = True
            list_item.ListSubItems.item(7).Bold = True
            list_item.ListSubItems.item(8).Bold = True
            list_item.ForeColor = &HFF0000
            list_item.ListSubItems.item(1).ForeColor = &HFF0000
            list_item.ListSubItems.item(2).ForeColor = &HFF0000
            list_item.ListSubItems.item(3).ForeColor = &HFF0000
            list_item.ListSubItems.item(4).ForeColor = &HFF0000
            list_item.ListSubItems.item(5).ForeColor = &HFF0000
            list_item.ListSubItems.item(6).ForeColor = &HFF0000
            list_item.ListSubItems.item(7).ForeColor = &HFF0000
            list_item.ListSubItems.item(8).ForeColor = &HFF0000
          
         
         End If
      End If
      rsaux10.MoveNext:
   Wend
   rsaux10.Close

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_ordensurtido)
End Sub

Private Sub lv_pedidos_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = 116 Then
      If lv_pedidos.selectedItem.SubItems(18) = 1 Then
         MsgBox "El pedido no se puede modificar", vbOKOnly, "ATENCION"
      Else
         If lv_pedidos.selectedItem.SubItems(17) = "*" Then
            lv_pedidos.selectedItem.SubItems(17) = ""
            lv_pedidos.selectedItem.Bold = False
            lv_pedidos.selectedItem.ListSubItems.item(1).Bold = False
            lv_pedidos.selectedItem.ListSubItems.item(2).Bold = False
            lv_pedidos.selectedItem.ListSubItems.item(3).Bold = False
            lv_pedidos.selectedItem.ListSubItems.item(4).Bold = False
            lv_pedidos.selectedItem.ListSubItems.item(5).Bold = False
            lv_pedidos.selectedItem.ListSubItems.item(6).Bold = False
            lv_pedidos.selectedItem.ListSubItems.item(7).Bold = False
            lv_pedidos.selectedItem.ListSubItems.item(8).Bold = False
            lv_pedidos.selectedItem.ForeColor = &H0&
            lv_pedidos.selectedItem.ListSubItems.item(1).ForeColor = &H0&
            lv_pedidos.selectedItem.ListSubItems.item(2).ForeColor = &H0&
            lv_pedidos.selectedItem.ListSubItems.item(3).ForeColor = &H0&
            lv_pedidos.selectedItem.ListSubItems.item(4).ForeColor = &H0&
            lv_pedidos.selectedItem.ListSubItems.item(5).ForeColor = &H0&
            lv_pedidos.selectedItem.ListSubItems.item(6).ForeColor = &H0&
            lv_pedidos.selectedItem.ListSubItems.item(7).ForeColor = &H0&
            lv_pedidos.selectedItem.ListSubItems.item(8).ForeColor = &H0&
         Else
            var_si = MsgBox("¿La orden de surtido se facturara en ceros?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               lv_pedidos.selectedItem.SubItems(17) = "*"
               lv_pedidos.selectedItem.Bold = True
               lv_pedidos.selectedItem.ListSubItems.item(1).Bold = True
               lv_pedidos.selectedItem.ListSubItems.item(2).Bold = True
               lv_pedidos.selectedItem.ListSubItems.item(3).Bold = True
               lv_pedidos.selectedItem.ListSubItems.item(4).Bold = True
               lv_pedidos.selectedItem.ListSubItems.item(5).Bold = True
               lv_pedidos.selectedItem.ListSubItems.item(6).Bold = True
               lv_pedidos.selectedItem.ListSubItems.item(7).Bold = True
               lv_pedidos.selectedItem.ListSubItems.item(8).Bold = True
               lv_pedidos.selectedItem.ForeColor = &HFF0000
               lv_pedidos.selectedItem.ListSubItems.item(1).ForeColor = &HFF0000
               lv_pedidos.selectedItem.ListSubItems.item(2).ForeColor = &HFF0000
               lv_pedidos.selectedItem.ListSubItems.item(3).ForeColor = &HFF0000
               lv_pedidos.selectedItem.ListSubItems.item(4).ForeColor = &HFF0000
               lv_pedidos.selectedItem.ListSubItems.item(5).ForeColor = &HFF0000
               lv_pedidos.selectedItem.ListSubItems.item(6).ForeColor = &HFF0000
               lv_pedidos.selectedItem.ListSubItems.item(7).ForeColor = &HFF0000
               lv_pedidos.selectedItem.ListSubItems.item(8).ForeColor = &HFF0000
            End If
         End If
      End If
   End If
      
   If Shift = 1 And KeyCode = 117 Then
      For var_j = 1 To lv_pedidos.ListItems.Count
         lv_pedidos.ListItems.item(var_j).Selected = True
         If lv_pedidos.selectedItem.SubItems(19) = "*" Then
            If lv_pedidos.selectedItem.SubItems(18) <> 1 Then
               If lv_pedidos.selectedItem.SubItems(17) = "*" Then
                  lv_pedidos.selectedItem.SubItems(17) = ""
                  lv_pedidos.selectedItem.Bold = False
                  lv_pedidos.selectedItem.ListSubItems.item(1).Bold = False
                  lv_pedidos.selectedItem.ListSubItems.item(2).Bold = False
                  lv_pedidos.selectedItem.ListSubItems.item(3).Bold = False
                  lv_pedidos.selectedItem.ListSubItems.item(4).Bold = False
                  lv_pedidos.selectedItem.ListSubItems.item(5).Bold = False
                  lv_pedidos.selectedItem.ListSubItems.item(6).Bold = False
                  lv_pedidos.selectedItem.ListSubItems.item(7).Bold = False
                  lv_pedidos.selectedItem.ListSubItems.item(8).Bold = False
                  lv_pedidos.selectedItem.ForeColor = &H0&
                  lv_pedidos.selectedItem.ListSubItems.item(1).ForeColor = &H0&
                  lv_pedidos.selectedItem.ListSubItems.item(2).ForeColor = &H0&
                  lv_pedidos.selectedItem.ListSubItems.item(3).ForeColor = &H0&
                  lv_pedidos.selectedItem.ListSubItems.item(4).ForeColor = &H0&
                  lv_pedidos.selectedItem.ListSubItems.item(5).ForeColor = &H0&
                  lv_pedidos.selectedItem.ListSubItems.item(6).ForeColor = &H0&
                  lv_pedidos.selectedItem.ListSubItems.item(7).ForeColor = &H0&
                  lv_pedidos.selectedItem.ListSubItems.item(8).ForeColor = &H0&
               Else
                  'var_si = MsgBox("¿La orden de surtido se facturara en ceros?", vbYesNo, "ATENCION")
                  var_si = 6
                  If var_si = 6 Then
                     lv_pedidos.selectedItem.SubItems(17) = "*"
                     lv_pedidos.selectedItem.Bold = True
                     lv_pedidos.selectedItem.ListSubItems.item(1).Bold = True
                     lv_pedidos.selectedItem.ListSubItems.item(2).Bold = True
                     lv_pedidos.selectedItem.ListSubItems.item(3).Bold = True
                     lv_pedidos.selectedItem.ListSubItems.item(4).Bold = True
                     lv_pedidos.selectedItem.ListSubItems.item(5).Bold = True
                     lv_pedidos.selectedItem.ListSubItems.item(6).Bold = True
                     lv_pedidos.selectedItem.ListSubItems.item(7).Bold = True
                     lv_pedidos.selectedItem.ListSubItems.item(8).Bold = True
                     lv_pedidos.selectedItem.ForeColor = &HFF0000
                     lv_pedidos.selectedItem.ListSubItems.item(1).ForeColor = &HFF0000
                     lv_pedidos.selectedItem.ListSubItems.item(2).ForeColor = &HFF0000
                     lv_pedidos.selectedItem.ListSubItems.item(3).ForeColor = &HFF0000
                     lv_pedidos.selectedItem.ListSubItems.item(4).ForeColor = &HFF0000
                     lv_pedidos.selectedItem.ListSubItems.item(5).ForeColor = &HFF0000
                     lv_pedidos.selectedItem.ListSubItems.item(6).ForeColor = &HFF0000
                     lv_pedidos.selectedItem.ListSubItems.item(7).ForeColor = &HFF0000
                     lv_pedidos.selectedItem.ListSubItems.item(8).ForeColor = &HFF0000
                  End If
               End If
            End If
         Else
            If Me.lv_pedidos.selectedItem.SubItems(19) = "*" Then
               If lv_pedidos.selectedItem.SubItems(18) <> 1 Then
                  If lv_pedidos.selectedItem.SubItems(17) = "*" Then
                     lv_pedidos.selectedItem.SubItems(17) = ""
                     lv_pedidos.selectedItem.Bold = False
                     lv_pedidos.selectedItem.ListSubItems.item(1).Bold = False
                     lv_pedidos.selectedItem.ListSubItems.item(2).Bold = False
                     lv_pedidos.selectedItem.ListSubItems.item(3).Bold = False
                     lv_pedidos.selectedItem.ListSubItems.item(4).Bold = False
                     lv_pedidos.selectedItem.ListSubItems.item(5).Bold = False
                     lv_pedidos.selectedItem.ListSubItems.item(6).Bold = False
                     lv_pedidos.selectedItem.ListSubItems.item(7).Bold = False
                     lv_pedidos.selectedItem.ListSubItems.item(8).Bold = False
                     lv_pedidos.selectedItem.ForeColor = &H0&
                     lv_pedidos.selectedItem.ListSubItems.item(1).ForeColor = &H0&
                     lv_pedidos.selectedItem.ListSubItems.item(2).ForeColor = &H0&
                     lv_pedidos.selectedItem.ListSubItems.item(3).ForeColor = &H0&
                     lv_pedidos.selectedItem.ListSubItems.item(4).ForeColor = &H0&
                     lv_pedidos.selectedItem.ListSubItems.item(5).ForeColor = &H0&
                     lv_pedidos.selectedItem.ListSubItems.item(6).ForeColor = &H0&
                     lv_pedidos.selectedItem.ListSubItems.item(7).ForeColor = &H0&
                     lv_pedidos.selectedItem.ListSubItems.item(8).ForeColor = &H0&
                  Else
                     'var_si = MsgBox("¿La orden de surtido se facturara en ceros?", vbYesNo, "ATENCION")
                     var_si = 6
                     If var_si = 6 Then
                        lv_pedidos.selectedItem.SubItems(17) = "*"
                        lv_pedidos.selectedItem.Bold = True
                        lv_pedidos.selectedItem.ListSubItems.item(1).Bold = True
                        lv_pedidos.selectedItem.ListSubItems.item(2).Bold = True
                        lv_pedidos.selectedItem.ListSubItems.item(3).Bold = True
                        lv_pedidos.selectedItem.ListSubItems.item(4).Bold = True
                        lv_pedidos.selectedItem.ListSubItems.item(5).Bold = True
                        lv_pedidos.selectedItem.ListSubItems.item(6).Bold = True
                        lv_pedidos.selectedItem.ListSubItems.item(7).Bold = True
                        lv_pedidos.selectedItem.ListSubItems.item(8).Bold = True
                        lv_pedidos.selectedItem.ForeColor = &HFF0000
                        lv_pedidos.selectedItem.ListSubItems.item(1).ForeColor = &HFF0000
                        lv_pedidos.selectedItem.ListSubItems.item(2).ForeColor = &HFF0000
                        lv_pedidos.selectedItem.ListSubItems.item(3).ForeColor = &HFF0000
                        lv_pedidos.selectedItem.ListSubItems.item(4).ForeColor = &HFF0000
                        lv_pedidos.selectedItem.ListSubItems.item(5).ForeColor = &HFF0000
                        lv_pedidos.selectedItem.ListSubItems.item(6).ForeColor = &HFF0000
                        lv_pedidos.selectedItem.ListSubItems.item(7).ForeColor = &HFF0000
                        lv_pedidos.selectedItem.ListSubItems.item(8).ForeColor = &HFF0000
                     End If
                  End If
               End If
            End If
         End If
      Next var_j
   End If
   
   
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Unload Me
End Sub


Private Sub lv_pedidos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      i = lv_pedidos.selectedItem.Index
      If Me.lv_pedidos.selectedItem.SubItems(17) <> "*" Then
         If Me.lv_pedidos.selectedItem.SubItems(18) <> 1 Then
            If lv_pedidos.selectedItem.SubItems(19) = "*" Then
               lv_pedidos.selectedItem.SubItems(19) = ""
               lv_pedidos.ListItems.item(i).Bold = False
               lv_pedidos.ListItems.item(i).ForeColor = &H80000012
               lv_pedidos.ListItems.item(i).ListSubItems(1).Bold = False
               lv_pedidos.ListItems.item(i).ListSubItems(2).Bold = False
               lv_pedidos.ListItems.item(i).ListSubItems(3).Bold = False
               lv_pedidos.ListItems.item(i).ListSubItems(4).Bold = False
               lv_pedidos.ListItems.item(i).ListSubItems(5).Bold = False
               lv_pedidos.ListItems.item(i).ListSubItems(6).Bold = False
               lv_pedidos.ListItems.item(i).ListSubItems(7).Bold = False
               lv_pedidos.ListItems.item(i).ListSubItems(8).Bold = False
               lv_pedidos.ListItems.item(i).ListSubItems(1).ForeColor = &H80000012
               lv_pedidos.ListItems.item(i).ListSubItems(2).ForeColor = &H80000012
               lv_pedidos.ListItems.item(i).ListSubItems(3).ForeColor = &H80000012
               lv_pedidos.ListItems.item(i).ListSubItems(4).ForeColor = &H80000012
               lv_pedidos.ListItems.item(i).ListSubItems(5).ForeColor = &H80000012
               lv_pedidos.ListItems.item(i).ListSubItems(6).ForeColor = &H80000012
               lv_pedidos.ListItems.item(i).ListSubItems(7).ForeColor = &H80000012
               lv_pedidos.ListItems.item(i).ListSubItems(8).ForeColor = &H80000012
               lv_pedidos.Refresh
            Else
               lv_pedidos.selectedItem.SubItems(19) = "*"
               lv_pedidos.ListItems.item(i).Bold = True
               lv_pedidos.ListItems.item(i).ForeColor = &HC0&
               lv_pedidos.ListItems.item(i).ListSubItems(1).Bold = True
               lv_pedidos.ListItems.item(i).ListSubItems(2).Bold = True
               lv_pedidos.ListItems.item(i).ListSubItems(3).Bold = True
               lv_pedidos.ListItems.item(i).ListSubItems(4).Bold = True
               lv_pedidos.ListItems.item(i).ListSubItems(5).Bold = True
               lv_pedidos.ListItems.item(i).ListSubItems(6).Bold = True
               lv_pedidos.ListItems.item(i).ListSubItems(7).Bold = True
               lv_pedidos.ListItems.item(i).ListSubItems(8).Bold = True
               lv_pedidos.ListItems.item(i).ListSubItems(1).ForeColor = &HC0&
               lv_pedidos.ListItems.item(i).ListSubItems(2).ForeColor = &HC0&
               lv_pedidos.ListItems.item(i).ListSubItems(3).ForeColor = &HC0&
               lv_pedidos.ListItems.item(i).ListSubItems(4).ForeColor = &HC0&
               lv_pedidos.ListItems.item(i).ListSubItems(5).ForeColor = &HC0&
               lv_pedidos.ListItems.item(i).ListSubItems(6).ForeColor = &HC0&
               lv_pedidos.ListItems.item(i).ListSubItems(7).ForeColor = &HC0&
               lv_pedidos.ListItems.item(i).ListSubItems(8).ForeColor = &HC0&
               lv_pedidos.Refresh
            End If
         End If
      End If
   End If
End Sub

Private Sub txt_factura_ceros_KeyPress(KeyAscii As Integer)
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
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Dim var_estatus_factura_ceros As Integer
      If IsNumeric(txt_factura_ceros) Then
         rs.Open "select * from tb_enc_orden_surtido where inte_ors_orden_surtido = " + Me.txt_factura_ceros, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_estatus_factura_ceros = IIf(IsNull(rs!inte_ors_factura_ceros), 0, rs!inte_ors_factura_ceros)
            rsaux.Open "select * from tb_det_orden_surtido where inte_ors_orden_surtido = " + Me.txt_factura_ceros + " and (floa_ors_cantidad_surtida > 0 or floa_ors_cantidad_negada > 0 or floa_ors_cantidad_empacada > 0)", cnn, adOpenDynamic, adLockOptimistic
            If rsaux.EOF Then
               If var_estatus_factura_ceros = 0 Then
                  var_si = MsgBox("¿Desea facturar en ceros la orden de surtido?", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     rsaux3.Open "UPDATE TB_ENC_ORDEN_SURTIDO SET INTE_ORS_FACTURA_CEROS = 1 WHERE INTE_ORS_ORDEN_SURTIDO = " + Me.txt_factura_ceros, cnn, adOpenDynamic, adLockOptimistic
                     MsgBox "La orden de surtido se facturara en ceros", vbOKOnly, "ATENCION"
                  End If
               Else
                  var_si = MsgBox("La orden de surtido ya tiene estatus de factura en ceros, ¿Desea quitar este estatus?", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     rsaux3.Open "UPDATE TB_ENC_ORDEN_SURTIDO SET INTE_ORS_FACTURA_CEROS = 0 WHERE INTE_ORS_ORDEN_SURTIDO = " + Me.txt_factura_ceros, cnn, adOpenDynamic, adLockOptimistic
                     MsgBox "La orden de surtido no se facturara en ceros", vbOKOnly, "ATENCION"
                  End If
               End If
            Else
               MsgBox "La orden de surtido ya no puede ser modificada ya que fue surtida", vbOKOnly, "ATENCION"
            End If
            rsaux.Close
         Else
            MsgBox "La orden de surtido no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
      End If
      Me.frm_factura_ceros.Visible = False
   End If
   If KeyAscii = 27 Then
      Me.frm_factura_ceros.Visible = False
   End If
End Sub

Private Sub txt_factura_ceros_LostFocus()
   Me.frm_factura_ceros.Visible = False
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
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
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      cnn.BeginTrans
      rs.Open "select max(INTE_TEM_CONSECUTIVO) as consecutivo from tb_temp_orden_surtido", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_consecutivo = IIf(IsNull(rs!consecutivo), 0, rs!consecutivo) + 1
      Else
         var_consecutivo = 1
      End If
      rs.Close
      rs.Open "insert into tb_temp_orden_surtido (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
      cnn.CommitTrans
      Cadena = "INSERT INTO [TB_TEMP_ORDEN_SURTIDO] ([INTE_TEM_CONSECUTIVO], [CHAR_TPE_TIPO_PEDIDO], [INTE_PED_NUMERO], [VCHA_ALM_ALMACEN_ID], [INTE_ORS_ORDEN_SURTIDO], [DTIM_ORS_FECHA_CARGA], [DTIM_ORS_FECHA_CADUCA], [VCHA_ESB_NOMBRE], [VCHA_ESB_ESTABLECIMIENTO_ID], [VCHA_CLI_CLAVE_ID], [VCHA_CLI_NOMBRE], "
      Cadena = Cadena + "[CHAR_PRI_PRIORIDAD_ID], [FLOA_ORS_DESCUENTO_1], [FLOA_ORS_DESCUENTO_2], [VCHA_ART_ARTICULO_ID], [VCHA_ART_NOMBRE_ESPAÑOL], [FLOA_ORS_PRECIO], [FLOA_ORS_CANTIDAD_SURTIR], [VCHA_AGE_NOMBRE], [VCHA_RUT_NOMBRE], [INTE_PED_DIAS_CONDICIONES], [INTE_PED_DIAS_CADUCIDAD], [MONE_ART_COSTO_ESTANDAR], [FLOA_ORS_CANTIDAD_SURTIDA], "
      Cadena = Cadena + "[VCHA_TIT_TITULAR_ID], [VCHA_TIT_NOMBRE], [VCHA_RUT_RUTA_ID], [VCHA_AGE_AGENTE_ID], [CHAR_ORS_ESTATUS], [VCHA_MOV_MOVIMIENTO_ID], [VCHA_CLI_EMAIL], [FLOA_TPE_IVA], [ALMACEN_AGENTE], [INTE_TCL_EMPACADO], [VCHA_TCL_ALMACEN_EMPAQUE], [FLOA_ORS_CANTIDAD_EMPACADA], [VCHA_PLA_PAZO_ID], [INTE_CLI_AGRUPADOR], [VCHA_MON_MONEDA_ID],"
      Cadena = Cadena + "[VCHA_ALM_NOMBRE], [FLOA_ORS_PROMOCION_1], [FLOA_ORS_PROMOCION_2], [INTE_ORS_FACTURA_CEROS], [CHAR_PED_TIPO], [VCHA_UOR_UNIDAD_ID], [INTE_PED_REFERENCIA],  [VCHA_EXI_UBICACION], [FLOA_ORS_CANTIDAD_NEGADA], [VCHA_TEM_CODIGO_BARRAS]) "
      Cadena = Cadena + " select " + CStr(var_consecutivo) + ", CHAR_TPE_TIPO_PEDIDO_ID, INTE_PED_NUMERO, VCHA_ALM_ALMACEN_ID, INTE_ORS_ORDEN_SURTIDO, DTIM_ORS_FECHA_CARGA, DTIM_ORS_FECHA_CADUCA, VCHA_ESB_NOMBRE, VCHA_ESB_ESTABLECIMIENTO_ID, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, CHAR_PRI_PRIORIDAD_ID, FLOA_ORS_DESCUENTO_1, FLOA_ORS_DESCUENTO_2, VCHA_ART_ARTICULO_ID, "
      Cadena = Cadena + " substring(VCHA_ART_NOMBRE_ESPAÑOL,1,50), FLOA_ORS_PRECIO, FLOA_ORS_CANTIDAD_SURTIR, VCHA_AGE_NOMBRE, VCHA_RUT_NOMBRE, INTE_PED_DIAS_CONDICIONES, INTE_PED_DIAS_CADUCIDAD, MONE_ART_COSTO_ESTANDAR, FLOA_ORS_CANTIDAD_SURTIDA, VCHA_TIT_TITULAR_ID, VCHA_TIT_NOMBRE, VCHA_RUT_RUTA_ID, VCHA_AGE_AGENTE_ID, CHAR_ORS_ESTATUS, VCHA_MOV_MOVIMIENTO_ID, VCHA_CLI_EMAIL, FLOA_TPE_IVA,"
      Cadena = Cadena + " ALMACEN_AGENTE, INTE_TCL_EMPACADO, VCHA_TCL_ALMACEN_EMPAQUE, FLOA_ORS_CANTIDAD_EMPACADA, VCHA_PLA_PlAZO_ID, INTE_CLI_AGRUPADOR, VCHA_MON_MONEDA_ID, VCHA_ALM_NOMBRE, FLOA_ORS_PROMOCION_1, FLOA_ORS_PROMOCION_2, INTE_ORS_FACTURA_CEROS, CHAR_PED_TIPO, VCHA_UOR_UNIDAD_ID, INTE_PED_REFERENCIA, VCHA_EXI_UBICACION, FLOA_ORS_CANTIDAD_NEGADA, " + Format(txt_numero, "##########") + " from vw_orden_surtido where inte_ors_orden_surtido = " + txt_numero
      rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
      If var_empresa = "18" Then
         Set reporte = appl.OpenReport(App.Path + "\rep_orden_surtido_nueva_textilera.rpt")
      Else
         Set reporte = appl.OpenReport(App.Path + "\rep_orden_surtido_nueva.rpt")
      End If
      reporte.RecordSelectionFormula = "{TB_TEMP_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO} = " + txt_numero + " AND {TB_TEMP_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND {TB_TEMP_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIR} > 0"
      frmvistasprevias.cr.ReportSource = reporte
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      frmvistasprevias.cr.ViewReport
      frmvistasprevias.Caption = "Orden de Surtido"
      frmvistasprevias.Show 1
      Set reporte = Nothing
      rs.Open "delete from tb_temp_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
   End If
   If KeyAscii = 27 Then
      frm_imprimir.Visible = False
   End If
End Sub

Private Sub txt_numero_LostFocus()
   frm_imprimir.Visible = False
End Sub

Private Sub txt_pedido_resurtir_KeyPress(KeyAscii As Integer)
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
   
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If IsNumeric(txt_pedido_resurtir) Then
         Set TB_ENC_ORDEN_SURTIDO = New TB_ENC_ORDEN_SURTIDO
         Set TB_DET_ORDEN_SURTIDO_I = New TB_DET_ORDEN_SURTIDO_I
         Set TB_ENC_PEDIDOS_M = New TB_ENC_PEDIDOS_M
         Dim var_maximo_orden As Double
         Dim var_existen As Double
         Dim var_apartados As Double
         Dim var_disponible As Double
         Dim var_cantidad_pedidia As Double
         Dim var_surtir As Double
         Dim var_contador As Double
         Dim var_costo As Double
         Dim var_factura_ceros As Double
         Dim var_clave_moneda As String
         Dim var_promocion_1 As Double
         Dim var_promocion_2 As Double
         Dim var_diferencia As Double
         Dim var_cliente_referencia As String
         Dim var_hora  As String
         Dim i As Double
         Dim j As Double
         Dim var_tipo_pedido As String
         Dim var_orden_surtido_resurtir As Double
         si = MsgBox("¿Deseas ejecutar el proceso de ordenes de surtido", vbYesNo, "ATENCION")
         If si = 6 Then
            var_contador = 0
            rs.Open "select max(inte_ors_orden_surtido) from tb_enc_orden_surtido where inte_ped_numero = " + txt_pedido_resurtir, cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_orden_surtido_resurtir = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
            Else
               var_orden_surtido_resurtir = 0
            End If
            'rs.Open "select * from vw_resurtido where (char_ped_estatus <> 'C' or char_ped_estatus <> 'S') and inte_ped_numero = " + txt_pedido_resurtir + " and diferencia > 0", cnn, adOpenDynamic, adLockOptimistic
            var_hora = CStr(Time)
            If Not rs.EOF Then
               rs.Close
               If var_orden_surtido_resurtir > 0 Then
                  rs.Open "select * from tb_encabezado_pedidos where inte_ped_numero = " + txt_pedido_resurtir, cnn, adOpenDynamic, adLockOptimistic
                  var_tipo_pedido = IIf(IsNull(rs!char_tpe_tipo_pedido_id), "", rs!char_tpe_tipo_pedido_id)
                  var_resurtible = IIf(IsNull(rs!inte_ped_resurtible), 0, rs!inte_ped_resurtible)
                  If var_empresa = "28" Then
                     var_resurtible = 1
                  End If
                  var_cliente_referencia = IIf(IsNull(rs!VCHA_PED_CLIENTE_REFERENCIA), "", rs!VCHA_PED_CLIENTE_REFERENCIA)
                  If var_resurtible = 1 Then
                     var_clave_moneda = rs!VCHA_MON_MONEDA_ID
                     var_factura_ceros = 0
                     var_almacen = IIf(IsNull(rs!VCHA_ALM_ALMACEN_ID), "", rs!VCHA_ALM_ALMACEN_ID)
                     If Trim(var_almacen) <> "" Then
                        If IsNull(rs!inte_ped_autorizo) Then
                           var_a = 0
                        Else
                           var_a = rs!inte_ped_autorizo
                        End If
                        If var_a = 1 And Trim(rs!char_ped_estatus) = "S" Then
                           'rsaux.Open "update tb_detalle_pedidos set floa_ped_cantidad_depurada = 0 where inte_ped_numero = " + txt_pedido_resurtir, cnn, adOpenDynamic, adLockOptimistic
                           var_contador = var_contador + 1
                           rsaux.Open "select * from vw_maximo_orden_surtido ", cnn, adOpenDynamic, adLockOptimistic
                           If rsaux.EOF Then
                              var_maximo_orden = 1
                           Else
                              If IsNull(rsaux!maximo) Then
                                 var_maximo_orden = 1
                              Else
                                 var_maximo_orden = rsaux!maximo + 1
                              End If
                           End If
                           rsaux.Close
                           ok = TB_ENC_ORDEN_SURTIDO.Anadir(var_empresa, var_unidad_organizacional, rs!char_tpe_tipo_pedido_id, rs!inte_ped_numero, var_almacen, var_maximo_orden, Date, Date + IIf(IsNull(rs!INTE_PED_DIAS_CADUCIDAD), 0, rs!INTE_PED_DIAS_CADUCIDAD), "", rs!vcha_tit_titular_id, rs!vcha_cli_clave_id, rs!vcha_ESB_ESTABLECIMIENTO_id, rs!floa_ped_descuento_1, rs!floa_ped_descuento_2, rs!floa_ped_Descuento_3, "", "", Date, var_factura_ceros, var_clave_moneda, var_hora)
                           rsaux.Open "update tb_enc_orden_surtido set VCHA_PED_CLIENTE_REFERENCIA = '" + var_cliente_referencia + "', INTE_ORS_LIBERADA =  1 where INTE_ORS_ORDEN_SURTIDO = " + CStr(var_maximo_orden), cnn, adOpenDynamic, adLockOptimistic
                           rsaux2.Open "select * from VW_RESURTIDO_NEGADO_PRODUCCION where inte_ped_numero = " + Str(rs!inte_ped_numero) + " AND INTE_ORS_ORDEN_SURTIDO = " + CStr(var_orden_surtido_resurtir), cnn, adOpenDynamic, adLockOptimistic
                           While Not rsaux2.EOF
                                 If (rsaux2!floa_ors_cantidad_pedida - rsaux2!FLOA_ORS_CANTIDAD_SURTIR) > 0 Then
                                    var_promocion_1 = 0
                                    var_promocion_2 = 0
                                    var_promocion_1 = IIf(IsNull(rsaux2!floa_ped_promocion_1), 0, rsaux2!floa_ped_promocion_1)
                                    var_promocion_2 = IIf(IsNull(rsaux2!floa_ped_promocion_2), 0, rsaux2!floa_ped_promocion_2)
                                    If var_empresa = "28" Then
                                       rsaux.Open "select * from tb_existencias where vcha_alm_almacen_id = 'CDH' and vcha_art_articulo_id = '" + rsaux2!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                    Else
                                       rsaux.Open "select * from tb_existencias where vcha_alm_almacen_id = '" + var_almacen + "' and vcha_art_articulo_id = '" + rsaux2!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                    End If
                                    If rsaux.EOF Then
                                       var_existen = 0
                                       var_apartados = 0
                                       var_disponible = 0
                                    Else
                                       If IsNull(rsaux!floa_Exi_Cantidad) Then
                                          var_existen = 0
                                       Else
                                          var_existen = rsaux!floa_Exi_Cantidad
                                       End If
                                       If IsNull(rsaux!FLOA_EXI_cANTIDAD_APARTADA) Then
                                          var_apartados = 0
                                       Else
                                          var_apartados = rsaux!FLOA_EXI_cANTIDAD_APARTADA
                                       End If
                                       If IsNull(rsaux!floa_Exi_Cantidad_disponible) Then
                                          var_disponible = 0
                                       Else
                                          var_disponible = rsaux!floa_Exi_Cantidad_disponible
                                       End If
                                    End If
                                    rsaux.Close
                                    var_cantidad_pedida = IIf(IsNull(rsaux2!floa_ors_cantidad_pedida), 0, rsaux2!floa_ors_cantidad_pedida) - IIf(IsNull(rsaux2!FLOA_ORS_CANTIDAD_SURTIR), 0, rsaux2!FLOA_ORS_CANTIDAD_SURTIR)
                                    var_surtir = 0
                                    If var_cantidad_pedida > 0 Then
                                       If var_cantidad_pedida <= var_disponible Then
                                          var_surtir = var_cantidad_pedida
                                       Else
                                          If var_disponible <= 0 Then
                                             var_surtir = 0
                                          Else
                                             var_surtir = var_disponible
                                          End If
                                       End If
                                       rsaux3.Open "select * from tb_EXISTENCIAS where vcha_alm_almacen_id  = '" + var_almacen + "' and vcha_art_articulo_id = '" + rsaux2!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                       If Not rsaux3.EOF Then
                                          var_costo = IIf(IsNull(rsaux3!FLOA_eXI_COSTO), 0, rsaux3!FLOA_eXI_COSTO)
                                       End If
                                       rsaux3.Close
                                       ok = TB_DET_ORDEN_SURTIDO_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen, var_maximo_orden, rsaux2!vcha_Art_Articulo_id, var_costo, rsaux2!FLOA_PED_PRECIO, var_cantidad_pedida, var_existen, var_apartados, var_disponible, var_surtir, 0, 0, var_promocion_1, var_promocion_2, var_tipo_pedido)
                                    End If
                                 End If
                                 rsaux2.MoveNext
                            Wend
                            rsaux2.Close
                            ok = TB_ENC_PEDIDOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen, rs!inte_ped_numero, "S")
                            valor = rs!inte_ped_numero
                            Set itmfound = lv_pedidos.findItem(valor, lvwText, , lvwPartial)
                            
                            
                            
                            'Set reporte = appl.OpenReport(App.Path + "\rep_orden_surtido.rpt")
                            'reporte.RecordSelectionFormula = "{VW_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO} = " + Str(var_maximo_orden)
                            'frmvistasprevias.cr.ReportSource = reporte
                            'For ntablas = 1 To reporte.Database.Tables.Count
                            '    reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                            'Next ntablas
                            'frmvistasprevias.cr.ViewReport
                            'frmvistasprevias.Caption = "Orden de Surtido"
                            'frmvistasprevias.Show 1
                            'Set reporte = Nothing
                            MsgBox "Orden de surtido número " + CStr(var_maximo_orden), vbOKOnly, "ATENCION"
                            var_numero_orden = var_maximo_orden
                            If var_numero_orden > 0 Then
                            
                               cnn.BeginTrans
                               If rs.State = 1 Then
                                  rs.Close
                               End If
                               rs.Open "select max(INTE_TEM_CONSECUTIVO) as consecutivo from tb_temp_orden_surtido", cnn, adOpenDynamic, adLockOptimistic
                               If Not rs.EOF Then
                                  var_consecutivo = IIf(IsNull(rs!consecutivo), 0, rs!consecutivo) + 1
                               Else
                                  var_consecutivo = 1
                               End If
                               rs.Close
                               rs.Open "insert into tb_temp_orden_surtido (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                               cnn.CommitTrans
                               
                               Cadena = "INSERT INTO [TB_TEMP_ORDEN_SURTIDO] ([INTE_TEM_CONSECUTIVO], [CHAR_TPE_TIPO_PEDIDO], [INTE_PED_NUMERO], [VCHA_ALM_ALMACEN_ID], [INTE_ORS_ORDEN_SURTIDO], [DTIM_ORS_FECHA_CARGA], [DTIM_ORS_FECHA_CADUCA], [VCHA_ESB_NOMBRE], [VCHA_ESB_ESTABLECIMIENTO_ID], [VCHA_CLI_CLAVE_ID], [VCHA_CLI_NOMBRE], "
                               Cadena = Cadena + "[CHAR_PRI_PRIORIDAD_ID], [FLOA_ORS_DESCUENTO_1], [FLOA_ORS_DESCUENTO_2], [VCHA_ART_ARTICULO_ID], [VCHA_ART_NOMBRE_ESPAÑOL], [FLOA_ORS_PRECIO], [FLOA_ORS_CANTIDAD_SURTIR], [VCHA_AGE_NOMBRE], [VCHA_RUT_NOMBRE], [INTE_PED_DIAS_CONDICIONES], [INTE_PED_DIAS_CADUCIDAD], [MONE_ART_COSTO_ESTANDAR], [FLOA_ORS_CANTIDAD_SURTIDA], "
                               Cadena = Cadena + "[VCHA_TIT_TITULAR_ID], [VCHA_TIT_NOMBRE], [VCHA_RUT_RUTA_ID], [VCHA_AGE_AGENTE_ID], [CHAR_ORS_ESTATUS], [VCHA_MOV_MOVIMIENTO_ID], [VCHA_CLI_EMAIL], [FLOA_TPE_IVA], [ALMACEN_AGENTE], [INTE_TCL_EMPACADO], [VCHA_TCL_ALMACEN_EMPAQUE], [FLOA_ORS_CANTIDAD_EMPACADA], [VCHA_PLA_PAZO_ID], [INTE_CLI_AGRUPADOR], [VCHA_MON_MONEDA_ID],"
                               Cadena = Cadena + "[VCHA_ALM_NOMBRE], [FLOA_ORS_PROMOCION_1], [FLOA_ORS_PROMOCION_2], [INTE_ORS_FACTURA_CEROS], [CHAR_PED_TIPO], [VCHA_UOR_UNIDAD_ID], [INTE_PED_REFERENCIA],  [VCHA_EXI_UBICACION], [FLOA_ORS_CANTIDAD_NEGADA], [VCHA_TEM_CODIGO_BARRAS]) "
                               Cadena = Cadena + " select " + CStr(var_consecutivo) + ", CHAR_TPE_TIPO_PEDIDO_ID, INTE_PED_NUMERO, VCHA_ALM_ALMACEN_ID, INTE_ORS_ORDEN_SURTIDO, DTIM_ORS_FECHA_CARGA, DTIM_ORS_FECHA_CADUCA, VCHA_ESB_NOMBRE, VCHA_ESB_ESTABLECIMIENTO_ID, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, CHAR_PRI_PRIORIDAD_ID, FLOA_ORS_DESCUENTO_1, FLOA_ORS_DESCUENTO_2, VCHA_ART_ARTICULO_ID, "
                               Cadena = Cadena + " VCHA_ART_NOMBRE_ESPAÑOL, FLOA_ORS_PRECIO, FLOA_ORS_CANTIDAD_SURTIR, VCHA_AGE_NOMBRE, VCHA_RUT_NOMBRE, INTE_PED_DIAS_CONDICIONES, INTE_PED_DIAS_CADUCIDAD, MONE_ART_COSTO_ESTANDAR, FLOA_ORS_CANTIDAD_SURTIDA, VCHA_TIT_TITULAR_ID, VCHA_TIT_NOMBRE, VCHA_RUT_RUTA_ID, VCHA_AGE_AGENTE_ID, CHAR_ORS_ESTATUS, VCHA_MOV_MOVIMIENTO_ID, VCHA_CLI_EMAIL, FLOA_TPE_IVA,"
                               Cadena = Cadena + " ALMACEN_AGENTE, INTE_TCL_EMPACADO, VCHA_TCL_ALMACEN_EMPAQUE, FLOA_ORS_CANTIDAD_EMPACADA, VCHA_PLA_PlAZO_ID, INTE_CLI_AGRUPADOR, VCHA_MON_MONEDA_ID, VCHA_ALM_NOMBRE, FLOA_ORS_PROMOCION_1, FLOA_ORS_PROMOCION_2, INTE_ORS_FACTURA_CEROS, CHAR_PED_TIPO, VCHA_UOR_UNIDAD_ID, INTE_PED_REFERENCIA, VCHA_EXI_UBICACION, FLOA_ORS_CANTIDAD_NEGADA, " + Format(var_numero_orden, "##########") + " from vw_orden_surtido where inte_ors_orden_surtido = " + CStr(var_numero_orden)
                               
                               rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                               If var_empresa = "18" Then
                                  Set reporte = appl.OpenReport(App.Path + "\rep_orden_surtido_nueva_textilera.rpt")
                               Else
                                  Set reporte = appl.OpenReport(App.Path + "\rep_orden_surtido_nueva.rpt")
                               End If
                               reporte.RecordSelectionFormula = "{TB_TEMP_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO} = " + Str(var_numero_orden) + " AND {TB_TEMP_ORDEN_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND {TB_TEMP_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIR} > 0"
                               frmvistasprevias.cr.ReportSource = reporte
                               For ntablas = 1 To reporte.Database.Tables.Count
                                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                               Next ntablas
                               
                               frmvistasprevias.cr.ViewReport
                               frmvistasprevias.Caption = "Resurtido de Pedidos"
                               frmvistasprevias.Show
                               Set reporte = Nothing
                               rs.Open "delete from tb_temp_orden_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                            Else
                               MsgBox "No se pudo generar la orden de surtido del pedido " + Trim(CStr(var_numero_pedido)), vbOKOnly, "ATENCION"
                            End If
                         Else
                            If rs!char_ped_estatus = "E" Then
                               MsgBox "El pedido ya no puede ser resurtido ya que ya fue cerrado", vbOKOnly, "ATENCION"
                            End If
                            If var_a = 0 Then
                               MsgBox "El pedido no puede ser resurtido ya que no fue autorizado", vbOKOnly, "ATENCION"
                            End If
                         End If
                      End If
                  Else
                     MsgBox "El pedido no es resurtible", vbOKOnly, "ATENCION"
                  End If
               Else
               End If
            Else
               rs.Close
               MsgBox "El pedido ya no puede ser resurtido", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El proceso de generación de ordenes de surtido a sido cancelado", vbOKOnly, "ATENCION"
         End If
      End If
      frm_pedido_resurtir.Visible = False
      If rs.State = 1 Then
         rs.Close
      End If
   End If
   
   
   
   
   'Select Case KeyAscii
   'Case 48 To 57, 52, 13, 8, 46, 27
   'Case Else
   '    KeyAscii = 0
   'End Select
   'If KeyAscii = 13 Then
   '   If IsNumeric(txt_pedido_resurtir) Then
   '      Set TB_ENC_ORDEN_SURTIDO = New TB_ENC_ORDEN_SURTIDO
   '      Set TB_DET_ORDEN_SURTIDO_I = New TB_DET_ORDEN_SURTIDO_I
   '      Set TB_ENC_PEDIDOS_M = New TB_ENC_PEDIDOS_M
   '      Dim var_maximo_orden As Double
   '      Dim var_existen As Double
   '      Dim var_apartados As Double
   '      Dim var_disponible As Double
   '      Dim var_cantidad_pedidia As Double
   '      Dim var_surtir As Double
   '      Dim var_contador As Double
   '      Dim var_costo As Double
   '      Dim var_factura_ceros As Double
   '      Dim var_clave_moneda As String
   '      Dim var_promocion_1 As Double
   '      Dim var_promocion_2 As Double
   '      Dim var_diferencia As Double
   '      Dim var_hora  As String
   '      Dim i As Double
   '      Dim j As Double
   '      Dim var_tipo_pedido As String
   '      si = MsgBox("¿Deseas ejecutar el proceso de ordenes de surtido", vbYesNo, "ATENCION")
   '      If si = 6 Then
   '         var_contador = 0
   '         rs.Open "select * from vw_resurtido where (char_ped_estatus <> 'C' or char_ped_estatus <> 'S') and inte_ped_numero = " + txt_pedido_resurtir + " and diferencia > 0", cnn, adOpenDynamic, adLockOptimistic
   '         var_hora = CStr(Time)
   '         If Not rs.EOF Then
   '            var_tipo_pedido = IIf(IsNull(rs!char_tpe_tipo_pedido_id), "", rs!char_tpe_tipo_pedido_id)
   '            var_resurtible = IIf(IsNull(rs!inte_ped_resurtible), 0, rs!inte_ped_resurtible)
   '            If var_resurtible = 1 Then
   '               var_clave_moneda = rs!VCHA_MON_MONEDA_ID
   '               var_factura_ceros = 0
   '               var_almacen = IIf(IsNull(rs!VCHA_ALM_ALMACEN_ID), "", rs!VCHA_ALM_ALMACEN_ID)
   '               If Trim(var_almacen) <> "" Then
   '                  If IsNull(rs!inte_ped_autorizo) Then
   '                     var_a = 0
   '                  Else
   '                     var_a = rs!inte_ped_autorizo
   '                  End If
   '                  If var_a = 1 And Trim(rs!CHAR_PED_ESTATUS) = "S" Then
   '                     rsaux.Open "update tb_detalle_pedidos set floa_ped_cantidad_depurada = 0 where inte_ped_numero = " + txt_pedido_resurtir, cnn, adOpenDynamic, adLockOptimistic
   '                     var_contador = var_contador + 1
   '                     rsaux.Open "select * from vw_maximo_orden_surtido ", cnn, adOpenDynamic, adLockOptimistic
   '                     If rsaux.EOF Then
   '                        var_maximo_orden = 1
   '                     Else
   '                        If IsNull(rsaux!maximo) Then
   '                           var_maximo_orden = 1
   '                        Else
   '                           var_maximo_orden = rsaux!maximo + 1
   '                        End If
   '                     End If
   '                     rsaux.Close
   '                     ok = TB_ENC_ORDEN_SURTIDO.Anadir(var_empresa, var_unidad_organizacional, rs!char_tpe_tipo_pedido_id, rs!inte_ped_numero, var_almacen, var_maximo_orden, Date, Date + rs!inte_ped_dias_caducidad, "", rs!VCHA_TIT_TITULAR_ID, rs!VCHA_CLI_CLave_ID, rs!vcha_esb_establecimiento_id, rs!floa_ped_descuento_1, rs!floa_ped_Descuento_2, rs!floa_ped_Descuento_3, "", "", Date, var_factura_ceros, var_clave_moneda, var_hora)
   '                     rsaux2.Open "select * from tb_resurtido where inte_ped_numero = " + Str(rs!inte_ped_numero), cnn, adOpenDynamic, adLockOptimistic
   '                     While Not rsaux2.EOF
   '                           If (rs!floa_ped_cantidad - rs!cantidad_empacada) > 0 Then
   '                              var_promocion_1 = 0
   '                              var_promocion_2 = 0
   '                              var_promocion_1 = IIf(IsNull(rsaux2!floa_ped_promocion_1), 0, rsaux2!floa_ped_promocion_1)
   '                              var_promocion_2 = IIf(IsNull(rsaux2!floa_ped_promocion_2), 0, rsaux2!floa_ped_promocion_2)
   '                              rsaux.Open "select * from tb_existencias where vcha_alm_almacen_id = '" + var_almacen + "' and vcha_art_articulo_id = '" + rsaux2!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
   '                              If rsaux.EOF Then
   '                                 var_existen = 0
   '                                 var_apartados = 0
   '                                 var_disponible = 0
   '                              Else
   '                                 If IsNull(rsaux!floa_exi_cantidad) Then
   '                                    var_existen = 0
   '                                 Else
   '                                    var_existen = rsaux!floa_exi_cantidad
   '                                 End If
   '                                 If IsNull(rsaux!floa_exi_cantidad_apartada) Then
   '                                    var_apartados = 0
   '                                 Else
   '                                    var_apartados = rsaux!floa_exi_cantidad_apartada
   '                                 End If
   '                                 If IsNull(rsaux!floa_exi_cantidad_disponible) Then
   '                                    var_disponible = 0
   '                                 Else
   '                                    var_disponible = rsaux!floa_exi_cantidad_disponible
   '                                 End If
   '                              End If
   '                              rsaux.Close
   '                              var_cantidad_pedida = IIf(IsNull(rsaux2!floa_ped_cantidad), 0, rsaux2!floa_ped_cantidad) - IIf(IsNull(rsaux2!floa_ped_cantidad_surtida), 0, rsaux2!floa_ped_cantidad_surtida)
   '                              var_surtir = 0
   '                              If var_cantidad_pedida > 0 Then
   '                                 If var_cantidad_pedida <= var_disponible Then
   '                                    var_surtir = var_cantidad_pedida
   '                                 Else
   '                                    If var_disponible <= 0 Then
   '                                       var_surtir = 0
   '                                    Else
   '                                       var_surtir = var_disponible
   '                                    End If
   '                                 End If
   '                                 rsaux3.Open "select * from tb_EXISTENCIAS where vcha_alm_almacen_id  = '" + var_almacen + "' and vcha_art_articulo_id = '" + rsaux2!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
   '                                 If Not rsaux3.EOF Then
   '                                    var_costo = IIf(IsNull(rsaux3!floa_exi_costo), 0, rsaux3!floa_exi_costo)
   '                                 End If
   '                                 rsaux3.Close
   '                                 ok = TB_DET_ORDEN_SURTIDO_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen, var_maximo_orden, rsaux2!VCHA_ART_ARTICULO_ID, var_costo, rsaux2!FLOA_PED_PRECIO, var_cantidad_pedida, var_existen, var_apartados, var_disponible, var_surtir, 0, 0, var_promocion_1, var_promocion_2, var_tipo_pedido)
   '                              End If
   '                           End If
   '                           rsaux2.MoveNext
   '                      Wend
   '                      rsaux2.Close
   '                      ok = TB_ENC_PEDIDOS_M.Anadir(var_empresa, var_unidad_organizacional, var_almacen, rs!inte_ped_numero, "S")
   '                      valor = rs!inte_ped_numero
   '                      Set itmfound = lv_pedidos.findItem(valor, lvwText, , lvwPartial)
   '                      Set reporte = appl.OpenReport(App.Path + "\rep_orden_surtido.rpt")
   '                      reporte.RecordSelectionFormula = "{VW_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO} = " + Str(var_maximo_orden)
   '                      frmvistasprevias.cr.ReportSource = reporte
   '                      For ntablas = 1 To reporte.Database.Tables.Count
   '                          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
   '                      Next ntablas
   '                      frmvistasprevias.cr.ViewReport
   '                      frmvistasprevias.Caption = "Orden de Surtido"
   '                      frmvistasprevias.Show 1
   '                      Set reporte = Nothing
   '                   Else
   '                      If rs!CHAR_PED_ESTATUS = "E" Then
   '                         MsgBox "El pedido ya no puede ser resurtido ya que ya fue cerrado", vbOKOnly, "ATENCION"
   '                      End If
   '                      If var_a = 0 Then
   '                         MsgBox "El pedido no puede ser resurtido ya que no fue autorizado", vbOKOnly, "ATENCION"
   '                      End If
   '                   End If
   '                End If
   '            Else
   '               MsgBox "El pedido no es resurtible", vbOKOnly, "ATENCION"
   '            End If
   '         Else
   '            MsgBox "El pedido ya no puede ser resurtido", vbOKOnly, "ATENCION"
   '         End If
   '         rs.Close
   '         If var_contador > 0 Then
   '            'MsgBox "Se a terminado la generación de las ordenes de surtido", vbOKOnly, "ATENCION"
   '         Else
   '            'MsgBox "No existen pedidos a cargar", vbOKOnly, "ATENCION"
   '         End If
   '      Else
   '         MsgBox "El proceso de generación de ordenes de surtido a sido cancelado", vbOKOnly, "ATENCION"
   '      End If
   '   End If
   '   frm_pedido_resurtir.Visible = False
   'End If
   'If KeyAscii = 27 Then
   '   frm_pedido_resurtir.Visible = False
   'End If
End Sub
