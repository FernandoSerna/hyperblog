VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmasignacion_negado_orden_surtido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Depuración de ordenes de surtido"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11640
   Begin VB.Frame Frame1 
      Caption         =   " Orden de Surtido "
      Height          =   1860
      Left            =   60
      TabIndex        =   21
      Top             =   555
      Width           =   5400
      Begin VB.TextBox txt_nombre_cliente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1995
         TabIndex        =   5
         Top             =   1050
         Width           =   3345
      End
      Begin VB.TextBox txt_cliente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   720
         TabIndex        =   4
         Top             =   1050
         Width           =   1260
      End
      Begin VB.TextBox txt_nombre_agente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1995
         TabIndex        =   3
         Top             =   705
         Width           =   3345
      End
      Begin VB.TextBox txt_agente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   720
         TabIndex        =   2
         Top             =   705
         Width           =   1260
      End
      Begin VB.TextBox txt_numero_pedido 
         Enabled         =   0   'False
         Height          =   315
         Left            =   720
         TabIndex        =   6
         Top             =   1395
         Width           =   1260
      End
      Begin VB.TextBox txt_numero_orden 
         Height          =   315
         Left            =   720
         TabIndex        =   1
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   105
         TabIndex        =   25
         Top             =   1110
         Width           =   525
      End
      Begin VB.Label lbl_agente 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   105
         TabIndex        =   24
         Top             =   765
         Width           =   555
      End
      Begin VB.Label lbl_numero_pedido 
         AutoSize        =   -1  'True
         Caption         =   "Pedido:"
         Height          =   195
         Left            =   105
         TabIndex        =   23
         Top             =   1455
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   105
         TabIndex        =   22
         Top             =   420
         Width           =   600
      End
   End
   Begin VB.Frame frm_causas 
      Height          =   1860
      Left            =   855
      TabIndex        =   11
      Top             =   3435
      Width           =   3870
      Begin VB.TextBox txt_cantidad 
         Height          =   315
         Left            =   825
         TabIndex        =   12
         Top             =   1470
         Width           =   1155
      End
      Begin MSComctlLib.ListView lv_causas 
         Height          =   1185
         Left            =   15
         TabIndex        =   13
         Top             =   240
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   2090
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
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Causa"
            Object.Width           =   6174
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Causa de Negado"
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   3855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   90
         TabIndex        =   14
         Top             =   1515
         Width           =   675
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11130
      Picture         =   "frmasignacion_negado_orden_surtido.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Salir"
      Top             =   75
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frmasignacion_negado_orden_surtido.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   75
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Frame frm_cantidad_eliminar 
      Height          =   765
      Left            =   7965
      TabIndex        =   0
      Top             =   2760
      Width           =   1770
      Begin VB.TextBox txt_cantidad_eliminar 
         Height          =   315
         Left            =   105
         TabIndex        =   7
         Top             =   315
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         Caption         =   " Cantidad a Eliminar"
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   0
         TabIndex        =   8
         Top             =   15
         Width           =   1755
      End
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   2940
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   12
      ImageHeight     =   12
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasignacion_negado_orden_surtido.frx":073C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasignacion_negado_orden_surtido.frx":1016
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   930
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasignacion_negado_orden_surtido.frx":18F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasignacion_negado_orden_surtido.frx":21CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasignacion_negado_orden_surtido.frx":2AA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasignacion_negado_orden_surtido.frx":3040
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasignacion_negado_orden_surtido.frx":391C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasignacion_negado_orden_surtido.frx":41F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasignacion_negado_orden_surtido.frx":4AD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasignacion_negado_orden_surtido.frx":4BE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasignacion_negado_orden_surtido.frx":4CF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasignacion_negado_orden_surtido.frx":4E06
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasignacion_negado_orden_surtido.frx":4F18
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasignacion_negado_orden_surtido.frx":502A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasignacion_negado_orden_surtido.frx":51AC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3855
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasignacion_negado_orden_surtido.frx":52BE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   " Detalle de Orden de Surtido"
      Height          =   4770
      Left            =   60
      TabIndex        =   18
      Top             =   2460
      Width           =   5400
      Begin VB.Frame frm_causas_negado_2 
         Height          =   1485
         Left            =   1005
         TabIndex        =   27
         Top             =   2310
         Width           =   3870
         Begin MSComctlLib.ListView lv_causas_negado_2 
            Height          =   1185
            Left            =   60
            TabIndex        =   28
            Top             =   240
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   2090
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
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Causa"
               Object.Width           =   6174
            EndProperty
         End
         Begin VB.Label Label4 
            BackColor       =   &H8000000D&
            Caption         =   " Causa de Negado"
            ForeColor       =   &H8000000E&
            Height          =   210
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   3855
         End
      End
      Begin VB.CommandButton cmd_pasar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4995
         Picture         =   "frmasignacion_negado_orden_surtido.frx":53D0
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Pasar todos los articulos"
         Top             =   4395
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_detalle 
         Height          =   4140
         Left            =   45
         TabIndex        =   19
         Top             =   240
         Width           =   5280
         _ExtentX        =   9313
         _ExtentY        =   7303
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
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   5433
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "marca"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Precio"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Detalle de Negado"
      Height          =   6675
      Left            =   5550
      TabIndex        =   16
      Top             =   540
      Width           =   6000
      Begin MSComctlLib.ListView lv_negado 
         Height          =   6330
         Left            =   60
         TabIndex        =   17
         Top             =   240
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   11165
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   5203
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Causa"
            Object.Width           =   1676
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "clave causa"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Precio"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   0
      TabIndex        =   20
      Top             =   345
      Width           =   11535
   End
End
Attribute VB_Name = "frmasignacion_negado_orden_surtido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_almacen As String

Private Sub detalle()
   Dim list_item As ListItem
   Dim var_cantidad_surtida As Double
   Dim var_cantidad_surtir As Double
   Dim var_cantidad_asignada As Double
   lv_detalle.ListItems.Clear
   lv_negado.ListItems.Clear
         rs.Open "select vcha_alm_almacen_id from tb_Det_orden_surtido where  inte_ors_orden_surtido = " + Me.txt_numero_orden, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_almacen = rs!VCHA_ALM_ALMACEN_ID
         End If
         rs.Close
         
         lv_detalle.ListItems.Clear
         rs.Open "select distinct vcha_cli_clave_id, vcha_cli_nombre, inte_ped_numero, vcha_age_agente_id, vcha_age_nombre, vcha_art_articulo_id,vcha_art_nombre_español,floa_ors_cantidad_surtida,floa_ors_cantidad_surtir, floa_ors_precio from vw_negado_depuracion where  inte_ors_orden_surtido = " + Me.txt_numero_orden + " and floa_ors_Cantidad_surtida < floa_ors_Cantidad_surtir", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_agente = IIf(IsNull(rs!vcha_age_agente_id), "", rs!vcha_age_agente_id)
            Me.txt_nombre_agente = IIf(IsNull(rs!vcha_age_nombre), "", rs!vcha_age_nombre)
            Me.txt_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
            Me.txt_nombre_cliente = IIf(IsNull(rs!vcha_cli_nombre), "", rs!vcha_cli_nombre)
            Me.txt_numero_pedido = IIf(IsNull(rs!inte_ped_numero), 0, rs!inte_ped_numero)
            While Not rs.EOF
               Set list_item = lv_detalle.ListItems.Add(, , rs!VCHA_aRT_ARTICULO_ID)
               list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ART_NOMBRE_ESPAÑOL), "", Trim(rs!VCHA_ART_NOMBRE_ESPAÑOL))
               var_cantidad_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
               var_cantidad_surtir = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), 0, rs!FLOA_ORS_CANTIDAD_SURTIR)
               rsaux2.Open "select * from vw_negado_embarque_suma_cantidades where inte_ors_orden_surtido = " + Me.txt_numero_orden + " and vcha_Art_articulo_id = '" + rs!VCHA_aRT_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux2.EOF Then
                  var_cantidad_asignada = rsaux2!cantidad_negado
               Else
                  var_cantidad_asignada = 0
               End If
               rsaux2.Close
               list_item.SubItems(2) = var_cantidad_surtir - var_cantidad_surtida - var_cantidad_asignada
               list_item.SubItems(4) = IIf(IsNull(rs!floa_ors_precio), 0, rs!floa_ors_precio)
               rs.MoveNext
            Wend
         End If
         rs.Close
         rs.Open "select distinct vcha_art_articulo_id,vcha_art_nombre_español, floa_neg_cantidad,vcha_cne_causa_id,vcha_cne_nombre, FLOA_NEG_PRECIO  from vw_negado_embarque_detalle where inte_ors_orden_surtido = " + Me.txt_numero_orden, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            While Not rs.EOF
               Set list_item = lv_negado.ListItems.Add(, , rs!VCHA_aRT_ARTICULO_ID)
               list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ART_NOMBRE_ESPAÑOL), "", Trim(rs!VCHA_ART_NOMBRE_ESPAÑOL))
               list_item.SubItems(2) = IIf(IsNull(rs!floa_neg_cantidad), 0, rs!floa_neg_cantidad)
               list_item.SubItems(3) = IIf(IsNull(rs!vcha_cne_nombre), 0, rs!vcha_cne_nombre)
               list_item.SubItems(4) = IIf(IsNull(rs!vcha_cne_causa_id), "", rs!vcha_cne_causa_id)
               list_item.SubItems(5) = IIf(IsNull(rs!floa_neg_precio), 0, rs!floa_neg_precio)
               rs.MoveNext
            Wend
         End If
         rs.Close
   If Me.lv_detalle.ListItems.Count > 18 Then
      lv_detalle.ColumnHeaders(2).Width = 2830.12
   Else
      lv_detalle.ColumnHeaders(2).Width = 3080.12
   End If
         
End Sub

Private Sub cmd_imprimir_Click()
   MsgBox "Modulo en proceso", vbOKOnly, "ATENCION"
End Sub

Private Sub cmd_pasar_Click()
      Dim list_item As ListItem
      rs.Open "select * from tb_causas_negado where char_cne_tipo <> 'P' and char_cne_tipo <> 'A' order by vcha_cne_nombre", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         lv_causas_negado_2.ListItems.Clear
         While Not rs.EOF
            Set list_item = lv_causas_negado_2.ListItems.Add(, , rs!vcha_cne_causa_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_cne_nombre), "", rs!vcha_cne_nombre)
            rs.MoveNext
         Wend
         frm_causas_negado_2.Visible = True
         lv_causas_negado_2.SetFocus
      End If
      rs.Close
End Sub

Private Sub cmd_salir_Click()
   Unload Me
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
   frm_causas.Visible = False
   frm_cantidad_eliminar.Visible = False
   Me.frm_causas_negado_2.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_asignacion_negado)
End Sub

Private Sub lv_causas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_causas, ColumnHeader)
End Sub

Private Sub lv_causas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      frm_causas.Visible = False
      lv_detalle.SetFocus
   End If
   If KeyAscii = 13 Then
      txt_cantidad.SetFocus
   End If
End Sub

Private Sub lv_causas_LostFocus()
   'txt_cantidad.SetFocus
End Sub

Private Sub lv_causas_negado_2_KeyPress(KeyAscii As Integer)
   Dim var_n As Double
   var_n = Me.lv_detalle.ListItems.Count
   If KeyAscii = 13 Then
      For var_i = 1 To var_n
         lv_detalle.ListItems.Item(var_i).Selected = True
         
         If CDbl(Me.lv_detalle.selectedItem.SubItems(2) * 1) > 0 Then
            txt_cantidad = Me.lv_detalle.selectedItem.SubItems(2) * 1
            If CDbl(txt_cantidad) > lv_detalle.selectedItem.SubItems(2) Then
               MsgBox "Cantidad Incorrecta", vbOKOnly, "ATENCION"
               lv_detalle.SetFocus
            Else
               Set TB_NEGADO_I = New TB_NEGADO_I
               Dim list_item As ListItem
               Dim var_precio As Double
               var_anadir = False
               var_precio = (lv_detalle.selectedItem.SubItems(4) * 1)
               rs.Open "select * from tb_negado where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen + "' and inte_ors_orden_surtido = " + Me.txt_numero_orden + " and VCHA_CNE_CAUSA_ID = '" + lv_causas_negado_2.selectedItem + "' and vcha_art_articulo_id = '" + lv_detalle.selectedItem + "' and floa_neg_precio = " + CStr(var_precio), cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  rsaux.Open "update tb_negado set floa_neg_cantidad = floa_neg_cantidad + " + txt_cantidad + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen + "' and inte_ors_orden_surtido = " + Me.txt_numero_orden + " and VCHA_CNE_CAUSA_ID = '" + lv_causas_negado_2.selectedItem + "' and vcha_art_articulo_id = '" + lv_detalle.selectedItem + "' and floa_neg_precio = " + CStr(var_precio), cnn, adOpenDynamic, adLockOptimistic
                  var_n = lv_negado.ListItems.Count
                  lv_negado.selectedItem.SubItems(2) = (lv_negado.selectedItem.SubItems(2) * 1) + CDbl(txt_cantidad)
                  lv_detalle.selectedItem.SubItems(2) = (lv_detalle.selectedItem.SubItems(2) * 1) - CDbl(txt_cantidad)
               Else
                  var_precio = (lv_detalle.selectedItem.SubItems(4) * 1)
                  var_anadir = TB_NEGADO_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen, CDbl(Me.txt_numero_orden), lv_causas_negado_2.selectedItem, lv_detalle.selectedItem, CDbl(txt_cantidad), "", "", CDbl(txt_numero_pedido), var_precio)
                  Set list_item = lv_negado.ListItems.Add(, , lv_detalle.selectedItem)
                  list_item.SubItems(1) = Trim(lv_detalle.selectedItem.SubItems(1))
                  list_item.SubItems(2) = txt_cantidad
                  list_item.SubItems(3) = Trim(lv_causas_negado_2.selectedItem.SubItems(1))
                  list_item.SubItems(4) = lv_causas_negado_2.selectedItem
                  list_item.SubItems(5) = var_precio
                  lv_detalle.selectedItem.SubItems(2) = lv_detalle.selectedItem.SubItems(2) - CDbl(txt_cantidad)
               End If
               rs.Close
            End If
            lv_detalle.SetFocus
            frm_causas.Visible = False
         End If
      Next var_i
      Me.frm_causas_negado_2.Visible = False
      rs.Open "select inte_ped_numero from tb_enc_orden_surtido where inte_ors_orden_surtido  = " + Me.txt_numero_orden, cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         rsaux2.Open "update tb_encabezado_pedidos set char_ped_estatus = 'E' where inte_ped_numero = " + CStr(IIf(IsNull(rs(0).Value), 0, rs(0).Value)), cnn, adOpenDynamic, adLockOptimistic
      End If
      rs.Close
   End If
   If KeyAscii = 27 Then
      Me.frm_causas_negado_2.Visible = False
   End If
End Sub

Private Sub lv_causas_negado_2_LostFocus()
   Me.frm_causas_negado_2.Visible = False
End Sub

Private Sub lv_detalle_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F4 para seleccionar la causa de negado"
End Sub

Private Sub lv_detalle_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 115 Then
      Dim list_item As ListItem
      rs.Open "select * from tb_causas_negado where char_cne_tipo <> 'P' and char_cne_tipo <> 'A' order by vcha_cne_nombre", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         lv_causas.ListItems.Clear
         While Not rs.EOF
            Set list_item = lv_causas.ListItems.Add(, , rs!vcha_cne_causa_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_cne_nombre), "", rs!vcha_cne_nombre)
            rs.MoveNext
         Wend
         txt_cantidad = lv_detalle.selectedItem.SubItems(2)
         frm_causas.Visible = True
         lv_causas.SetFocus
      End If
      rs.Close
   End If
End Sub

Private Sub lv_detalle_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_detalle.selectedItem.SubItems(3) = "*" Then
         lv_detalle.selectedItem.SubItems(3) = ""
      Else
         lv_detalle.selectedItem.SubItems(3) = "*"
      End If
   End If
End Sub

Private Sub lv_detalle_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub lv_negado_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      frm_cantidad_eliminar.Visible = True
      txt_cantidad_eliminar.SetFocus
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
         If CDbl(txt_cantidad_eliminar) > 0 Then
            If CDbl(txt_cantidad_eliminar) <= (lv_negado.selectedItem.SubItems(2) * 1) Then
               Dim var_n As Integer
               Dim var_i As Integer
               var_i = 1
               var_n = lv_detalle.ListItems.Count
               While var_i <= var_n
                     lv_detalle.ListItems.Item(var_i).Selected = True
                     If lv_negado.selectedItem = lv_detalle.selectedItem And (lv_negado.selectedItem.SubItems(5) * 1) = (lv_detalle.selectedItem.SubItems(4) * 1) Then
                        var_i = var_n + 1
                     Else
                        var_i = var_i + 1
                     End If
               Wend
               'Set itmfound = lv_detalle.FindItem(lv_negado.SelectedItem, lvwText, , lvwPartial)
               'itmfound.EnsureVisible
               'itmfound.Selected = True
               
               lv_negado.selectedItem.SubItems(2) = (lv_negado.selectedItem.SubItems(2) * 1) - CDbl(txt_cantidad_eliminar)
               lv_detalle.selectedItem.SubItems(2) = (lv_detalle.selectedItem.SubItems(2) * 1) + CDbl(txt_cantidad_eliminar)
               rsaux.Open "update tb_negado set floa_neg_cantidad = floa_neg_cantidad - " + txt_cantidad_eliminar + " where  inte_ors_orden_surtido = " + Me.txt_numero_orden + " and VCHA_CNE_CAUSA_ID = '" + lv_negado.selectedItem.SubItems(4) + "' and vcha_art_articulo_id = '" + lv_detalle.selectedItem + "' and floa_neg_precio = " + CStr(lv_negado.selectedItem.SubItems(5) * 1), cnn, adOpenDynamic, adLockOptimistic
               
            End If
         Else
            MsgBox "Cantidad Incorrecta", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Cantidad Incorrecta", vbOKOnly, "ATENCION"
      End If
      frm_cantidad_eliminar.Visible = False
   End If
   If KeyAscii = 27 Then
      Me.frm_cantidad_eliminar.Visible = False
   End If
End Sub

Private Sub txt_Cantidad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      frm_causas.Visible = False
   End If
   If KeyAscii = 13 Then
      If CDbl(txt_cantidad) > 0 Then
         If CDbl(txt_cantidad) > lv_detalle.selectedItem.SubItems(2) Then
            MsgBox "Cantidad Incorrecta", vbOKOnly, "ATENCION"
            lv_detalle.SetFocus
         Else
            cnn.BeginTrans
            Set TB_NEGADO_I = New TB_NEGADO_I
            Dim list_item As ListItem
            Dim var_precio As Double
            Dim var_n As Integer
            Dim var_i As Integer
            var_anadir = False
            var_precio = (lv_detalle.selectedItem.SubItems(4) * 1)
            rs.Open "select * from tb_negado where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen + "' and inte_ors_orden_surtido = " + Me.txt_numero_orden + " and VCHA_CNE_CAUSA_ID = '" + lv_causas.selectedItem + "' and vcha_art_articulo_id = '" + lv_detalle.selectedItem + "' and floa_neg_precio = " + CStr(var_precio), cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               If rsaux.State = 1 Then
                  rsaux.Close
               End If
               rsaux.Open "update tb_negado set floa_neg_cantidad = floa_neg_cantidad + " + txt_cantidad + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen + "' and inte_ors_orden_surtido = " + Me.txt_numero_orden + " and VCHA_CNE_CAUSA_ID = '" + lv_causas.selectedItem + "' and vcha_art_articulo_id = '" + lv_detalle.selectedItem + "' and floa_neg_precio = " + CStr(var_precio), cnn, adOpenDynamic, adLockOptimistic
               var_n = lv_negado.ListItems.Count
               var_i = 1
               While var_i <= var_n
                   lv_negado.ListItems.Item(var_i).Selected = True
                   If lv_negado.selectedItem = lv_detalle.selectedItem And (lv_negado.selectedItem.SubItems(5) * 1) = (lv_detalle.selectedItem.SubItems(4) * 1) And lv_causas.selectedItem = lv_negado.selectedItem.SubItems(4) Then
                      var_i = var_n + 1
                   Else
                      var_i = var_i + 1
                   End If
               Wend
               lv_negado.selectedItem.SubItems(2) = (lv_negado.selectedItem.SubItems(2) * 1) + CDbl(txt_cantidad)
               lv_detalle.selectedItem.SubItems(2) = (lv_detalle.selectedItem.SubItems(2) * 1) - CDbl(txt_cantidad)
            Else
               var_precio = (lv_detalle.selectedItem.SubItems(4) * 1)
               var_anadir = TB_NEGADO_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen, CDbl(Me.txt_numero_orden), lv_causas.selectedItem, lv_detalle.selectedItem, CDbl(txt_cantidad), "", "", CDbl(txt_numero_pedido), var_precio)
               Set list_item = lv_negado.ListItems.Add(, , lv_detalle.selectedItem)
               list_item.SubItems(1) = Trim(lv_detalle.selectedItem.SubItems(1))
               list_item.SubItems(2) = txt_cantidad
               list_item.SubItems(3) = Trim(lv_causas.selectedItem.SubItems(1))
               list_item.SubItems(4) = lv_causas.selectedItem
               list_item.SubItems(5) = var_precio
               lv_detalle.selectedItem.SubItems(2) = lv_detalle.selectedItem.SubItems(2) - CDbl(txt_cantidad)
            End If
            rs.Close
            cnn.CommitTrans
         End If
         lv_detalle.SetFocus
         frm_causas.Visible = False
      Else
         MsgBox "Canitdad Incorrecta", vbOKOnly, "ATENCION"
         lv_detalle.SetFocus
         frm_causas.Visible = False
      End If
   End If
End Sub

Private Sub txt_numero_orden_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If Trim(Me.txt_numero_orden) <> "" Then
         Me.txt_agente = ""
         Me.txt_cliente = ""
         Me.txt_nombre_agente = ""
         Me.txt_nombre_cliente = ""
         Me.txt_numero_pedido = ""
         Me.lv_detalle.ListItems.Clear
         Me.lv_negado.ListItems.Clear
         Call detalle
      Else
         Me.txt_agente = ""
         Me.txt_cliente = ""
         Me.txt_nombre_agente = ""
         Me.txt_nombre_cliente = ""
         Me.txt_numero_pedido = ""
         Me.lv_detalle.ListItems.Clear
         Me.lv_negado.ListItems.Clear
      End If
   End If
End Sub
