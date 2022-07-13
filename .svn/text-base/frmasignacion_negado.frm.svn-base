VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmasignacion_negado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de Causa a la Mercancía Negada"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11670
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm_cantidad_eliminar 
      Height          =   765
      Left            =   7965
      TabIndex        =   21
      Top             =   2760
      Width           =   1770
      Begin VB.TextBox txt_cantidad_eliminar 
         Height          =   315
         Left            =   105
         TabIndex        =   23
         Top             =   315
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         Caption         =   " Cantidad a Eliminar"
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   0
         TabIndex        =   22
         Top             =   15
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frmasignacion_negado.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   75
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11130
      Picture         =   "frmasignacion_negado.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   75
      Width           =   330
   End
   Begin VB.Frame frm_causas 
      Height          =   1860
      Left            =   1710
      TabIndex        =   18
      Top             =   3915
      Width           =   3870
      Begin VB.TextBox txt_cantidad 
         Height          =   315
         Left            =   825
         TabIndex        =   10
         Top             =   1470
         Width           =   1155
      End
      Begin MSComctlLib.ListView lv_causas 
         Height          =   1185
         Left            =   30
         TabIndex        =   9
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   90
         TabIndex        =   20
         Top             =   1515
         Width           =   675
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Causa de Negado"
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   3855
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Detalle de Negado"
      Height          =   6675
      Left            =   5550
      TabIndex        =   17
      Top             =   540
      Width           =   6000
      Begin MSComctlLib.ListView lv_negado 
         Height          =   6330
         Left            =   60
         TabIndex        =   8
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
   Begin VB.Frame Frame3 
      Caption         =   "Ordenes de Surtido"
      Height          =   1920
      Left            =   105
      TabIndex        =   16
      Top             =   1680
      Width           =   5400
      Begin MSComctlLib.ListView lv_ordenes 
         Height          =   1605
         Left            =   45
         TabIndex        =   0
         Top             =   210
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   2831
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
            Text            =   "Número"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cliente"
            Object.Width           =   7382
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Pedido"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Embarque "
      Height          =   1095
      Left            =   105
      TabIndex        =   12
      Top             =   540
      Width           =   5415
      Begin VB.TextBox txt_numero_embarque 
         Enabled         =   0   'False
         Height          =   315
         Left            =   885
         TabIndex        =   5
         Top             =   315
         Width           =   1215
      End
      Begin VB.TextBox txt_agente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   885
         TabIndex        =   7
         Top             =   645
         Width           =   4470
      End
      Begin VB.TextBox txt_jaula 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2610
         TabIndex        =   6
         Top             =   315
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   195
         TabIndex        =   15
         Top             =   375
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   195
         TabIndex        =   14
         Top             =   705
         Width           =   555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Jaula:"
         Height          =   195
         Left            =   2175
         TabIndex        =   13
         Top             =   375
         Width           =   420
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Detalle de Orden de Surtido"
      Height          =   3615
      Left            =   105
      TabIndex        =   11
      Top             =   3615
      Width           =   5400
      Begin MSComctlLib.ListView lv_detalle 
         Height          =   3270
         Left            =   45
         TabIndex        =   1
         Top             =   240
         Width           =   5280
         _ExtentX        =   9313
         _ExtentY        =   5768
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
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   0
      TabIndex        =   2
      Top             =   345
      Width           =   11535
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
            Picture         =   "frmasignacion_negado.frx":073C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasignacion_negado.frx":1016
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
            Picture         =   "frmasignacion_negado.frx":18F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasignacion_negado.frx":21CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasignacion_negado.frx":2AA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasignacion_negado.frx":3040
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasignacion_negado.frx":391C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasignacion_negado.frx":41F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasignacion_negado.frx":4AD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasignacion_negado.frx":4BE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasignacion_negado.frx":4CF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasignacion_negado.frx":4E06
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasignacion_negado.frx":4F18
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasignacion_negado.frx":502A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmasignacion_negado.frx":51AC
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
            Picture         =   "frmasignacion_negado.frx":52BE
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmasignacion_negado"
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
   If lv_ordenes.ListItems.Count > 0 Then
      If lv_ordenes.selectedItem > 0 Then
         lv_detalle.ListItems.Clear
         rs.Open "select vcha_alm_almacen_id from tb_Det_orden_surtido where  inte_ors_orden_surtido = " + Str(lv_ordenes.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_almacen = rs!VCHA_ALM_ALMACEN_ID
         End If
         rs.Close
         If var_empresa = "03" Then
            'MsgBox lv_ordenes.selectedItem
            var_cadena = " SELECT dbo.VW_ORDENES_SURTIDO_EMBARQUE.floa_ors_precio, dbo.VW_ORDENES_SURTIDO_EMBARQUE.INTE_EMB_EMBARQUE, dbo.VW_ORDENES_SURTIDO_EMBARQUE.VCHA_EMP_EMPRESA_ID, dbo.VW_ORDENES_SURTIDO_EMBARQUE.INTE_ORS_ORDEN_SURTIDO, dbo.VW_ORDENES_SURTIDO_EMBARQUE.VCHA_ART_ARTICULO_ID, dbo.VW_ORDENES_SURTIDO_EMBARQUE.FLOA_ORS_CANTIDAD_SURTIR, dbo.VW_ORDENES_SURTIDO_EMBARQUE.FLOA_ORS_CANTIDAD_SURTIDA, dbo.VW_ORDENES_SURTIDO_EMBARQUE.FLOA_ORS_CANTIDAD_NEGADA, dbo.VW_ORDENES_SURTIDO_EMBARQUE.FLOA_ORS_CANTIDAD_EMPACADA, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_Articulos.VCHA_ART_NOMBRE_ESPAÑOL FROM dbo.VW_ORDENES_SURTIDO_EMBARQUE INNER JOIN dbo.TB_ARTICULOS ON dbo.VW_ORDENES_SURTIDO_EMBARQUE.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID INNER JOIN dbo.TB_CLIENTES ON dbo.VW_ORDENES_SURTIDO_EMBARQUE.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID "
            var_cadena = var_cadena + " WHERE     (dbo.VW_ORDENES_SURTIDO_EMBARQUE.FLOA_ORS_CANTIDAD_SURTIR > dbo.VW_ORDENES_SURTIDO_EMBARQUE.FLOA_ORS_CANTIDAD_SURTIDA + ISNULL(dbo.VW_ORDENES_SURTIDO_EMBARQUE.FLOA_ORS_CANTIDAD_NEGADA, 0)) AND (dbo.VW_ORDENES_SURTIDO_EMBARQUE.INTE_ORS_ORDEN_SURTIDO = " + Str(lv_ordenes.selectedItem) + ")"
            'MsgBox var_cadena
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         Else
            rs.Open "select distinct vcha_art_articulo_id,vcha_art_nombre_español,floa_ors_cantidad_surtida,floa_ors_cantidad_surtir, floa_ors_precio from vw_negado_embarque with (nolock) where  inte_ors_orden_surtido = " + Str(lv_ordenes.selectedItem) + " and floa_ors_Cantidad_surtida < floa_ors_Cantidad_surtir", cnn, adOpenDynamic, adLockOptimistic
         End If
         If Not rs.EOF Then
            While Not rs.EOF
               Set list_item = lv_detalle.ListItems.Add(, , rs!vcha_Art_Articulo_id)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_Art_nombre_español), "", Trim(rs!vcha_Art_nombre_español))
               var_cantidad_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
               var_cantidad_surtir = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), 0, rs!FLOA_ORS_CANTIDAD_SURTIR)
               rsaux2.Open "select * from vw_negado_embarque_suma_cantidades where  inte_ors_orden_surtido = " + Str(lv_ordenes.selectedItem) + " and vcha_Art_articulo_id = '" + rs!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
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
         rs.Open "select distinct vcha_art_articulo_id,vcha_art_nombre_español, floa_neg_cantidad,vcha_cne_causa_id,vcha_cne_nombre, FLOA_NEG_PRECIO  from vw_negado_embarque_detalle where  inte_ors_orden_surtido = " + Str(lv_ordenes.selectedItem), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            While Not rs.EOF
               Set list_item = lv_negado.ListItems.Add(, , rs!vcha_Art_Articulo_id)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_Art_nombre_español), "", Trim(rs!vcha_Art_nombre_español))
               list_item.SubItems(2) = IIf(IsNull(rs!floa_neg_cantidad), 0, rs!floa_neg_cantidad)
               list_item.SubItems(3) = IIf(IsNull(rs!vcha_cne_nombre), 0, rs!vcha_cne_nombre)
               list_item.SubItems(4) = IIf(IsNull(rs!vcha_cne_causa_id), "", rs!vcha_cne_causa_id)
               list_item.SubItems(5) = IIf(IsNull(rs!floa_neg_precio), 0, rs!floa_neg_precio)
               rs.MoveNext
            Wend
         End If
         rs.Close
      End If
   End If
End Sub

Private Sub cmd_imprimir_Click()
   MsgBox "Modulo en proceso", vbOKOnly, "ATENCION"
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
   If var_negado_desde = 1 Then
      Me.txt_numero_embarque = frmsalidas.txt_embarque
   Else
      Me.txt_numero_embarque = frmsalidas_cajas.txt_embarque
   End If
   If IsNumeric(txt_numero_embarque) Then
      If Trim(txt_numero_embarque) <> "" Then
         rs.Open "select * from tb_encabezado_embarques where inte_emb_embarque = " + txt_numero_embarque, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            rsaux2.Open "Select * from tb_agentes where vcha_age_agente_id ='" + rs!VCHA_AGE_AGENTE_ID + "'", cnn, adOpenDynamic, adLockOptimistic
            txt_agente = rsaux2!VCHA_AGE_NOMBRE
            txt_jaula = rs!inte_jau_jaula_id
            rsaux2.Close
            rs.Close
            lv_ordenes.ListItems.Clear
            If var_empresa = "03" Then
               var_cadena = "SELECT DISTINCT dbo.VW_ORDENES_SURTIDO_EMBARQUE.INTE_EMB_EMBARQUE, dbo.VW_ORDENES_SURTIDO_EMBARQUE.VCHA_EMP_EMPRESA_ID, dbo.VW_ORDENES_SURTIDO_EMBARQUE.INTE_ORS_ORDEN_SURTIDO, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID , dbo.VW_ORDENES_SURTIDO_EMBARQUE.INTE_PED_NUMERO FROM  dbo.VW_ORDENES_SURTIDO_EMBARQUE INNER JOIN dbo.TB_CLIENTES ON dbo.VW_ORDENES_SURTIDO_EMBARQUE.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID WHERE (dbo.VW_ORDENES_SURTIDO_EMBARQUE.FLOA_ORS_CANTIDAD_SURTIR > dbo.VW_ORDENES_SURTIDO_EMBARQUE.FLOA_ORS_CANTIDAD_SURTIDA + ISNULL(dbo.VW_ORDENES_SURTIDO_EMBARQUE.FLOA_ORS_CANTIDAD_NEGADA, 0)) AND (dbo.VW_ORDENES_SURTIDO_EMBARQUE.INTE_EMB_EMBARQUE = " + Me.txt_numero_embarque + ") AND (dbo.VW_ORDENES_SURTIDO_EMBARQUE.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "')"
               rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            Else
               rs.Open "SELECT distinct inte_ors_orden_surtido,vcha_cli_nombre, inte_ped_numero FROM vw_negado_embarque WHERE INTE_EMB_EMBARQUE = " + txt_numero_embarque + " and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
            End If
            If Not rs.EOF Then
               While Not rs.EOF
                     Set list_item = lv_ordenes.ListItems.Add(, , rs!INTE_ORS_ORDEN_SURTIDO)
                     list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                     list_item.SubItems(2) = IIf(IsNull(rs!inte_ped_numero), 0, rs!inte_ped_numero)
                     rs.MoveNext
               Wend
            Else
               lv_ordenes.ListItems.Clear
               lv_detalle.ListItems.Clear
               lv_negado.ListItems.Clear
               MsgBox "El embarque no tiene movimientos asignados", vbOKOnly, "ATENCION"
            End If
            rs.Close
         Else
            rs.Close
            MsgBox "El número de embarque no existe", vbOKOnly, "ATENCION"
            txt_agente = ""
            lv_ordenes.ListItems.Clear
         End If
      End If
   Else
      MsgBox "Número incorrecto", vbOKOnly, "ATENCION"
      txt_numero_embarque = ""
   End If
   frm_causas.Visible = False
   frm_cantidad_eliminar.Visible = False
   
   If Me.lv_ordenes.ListItems.Count > 6 Then
      lv_ordenes.ColumnHeaders(2).Width = 3950.07
   Else
      lv_ordenes.ColumnHeaders(2).Width = 4185.07
   End If

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
      txt_Cantidad.SetFocus
   End If
End Sub

Private Sub lv_causas_LostFocus()
   'txt_cantidad.SetFocus
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
         txt_Cantidad = lv_detalle.selectedItem.SubItems(2)
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

Private Sub lv_ordenes_GotFocus()
   Call detalle
End Sub

Private Sub lv_ordenes_ItemClick(ByVal item As MSComctlLib.ListItem)
   Call detalle
End Sub


Private Sub lv_ordenes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      lv_detalle.SetFocus
   End If
End Sub

Private Sub txt_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
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
                     lv_detalle.ListItems.item(var_i).Selected = True
                     If lv_negado.selectedItem = lv_detalle.selectedItem And (lv_negado.selectedItem.SubItems(5) * 1) = (lv_detalle.selectedItem.SubItems(3) * 1) Then
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
               rsaux.Open "update tb_negado set floa_neg_cantidad = floa_neg_cantidad - " + txt_cantidad_eliminar + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen + "' and inte_ors_orden_surtido = " + lv_ordenes.selectedItem + " and VCHA_CNE_CAUSA_ID = '" + lv_negado.selectedItem.SubItems(4) + "' and vcha_art_articulo_id = '" + lv_detalle.selectedItem + "' and floa_neg_precio = " + CStr(lv_negado.selectedItem.SubItems(5) * 1), cnn, adOpenDynamic, adLockOptimistic
               
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

Private Sub txt_cantidad_KeyPress(KeyAscii As Integer)
'On Error GoTo salir:
   If KeyAscii = 27 Then
      frm_causas.Visible = False
   End If
   If KeyAscii = 13 Then
      If CDbl(txt_Cantidad) > 0 Then
         If CDbl(txt_Cantidad) > lv_detalle.selectedItem.SubItems(2) Then
            MsgBox "Cantidad Incorrecta", vbOKOnly, "ATENCION"
            lv_detalle.SetFocus
         Else
            Set TB_NEGADO_I = New TB_NEGADO_I
            Dim list_item As ListItem
            Dim var_precio As Double
            Dim var_n As Integer
            Dim var_i As Integer
            var_anadir = False
            var_precio = (lv_detalle.selectedItem.SubItems(4) * 1)
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open "select * from tb_negado where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen + "' and inte_ors_orden_surtido = " + lv_ordenes.selectedItem + " and VCHA_CNE_CAUSA_ID = '" + lv_causas.selectedItem + "' and vcha_art_articulo_id = '" + lv_detalle.selectedItem + "' and floa_neg_precio = " + CStr(var_precio), cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               rsaux.Open "update tb_negado set floa_neg_cantidad = floa_neg_cantidad + " + txt_Cantidad + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen + "' and inte_ors_orden_surtido = " + lv_ordenes.selectedItem + " and VCHA_CNE_CAUSA_ID = '" + lv_causas.selectedItem + "' and vcha_art_articulo_id = '" + lv_detalle.selectedItem + "' and floa_neg_precio = " + CStr(var_precio), cnn, adOpenDynamic, adLockOptimistic
               var_n = lv_negado.ListItems.Count
               var_i = 1
               While var_i <= var_n
                   lv_negado.ListItems.item(var_i).Selected = True
                   If lv_negado.selectedItem = lv_detalle.selectedItem And (lv_negado.selectedItem.SubItems(5) * 1) = (lv_detalle.selectedItem.SubItems(3) * 1) And lv_causas.selectedItem = lv_negado.selectedItem.SubItems(4) Then
                      var_i = var_n + 1
                   Else
                      var_i = var_i + 1
                   End If
               Wend
               lv_negado.selectedItem.SubItems(2) = (lv_negado.selectedItem.SubItems(2) * 1) + CDbl(txt_Cantidad)
               lv_detalle.selectedItem.SubItems(2) = (lv_detalle.selectedItem.SubItems(2) * 1) - CDbl(txt_Cantidad)
            Else
               var_precio = (lv_detalle.selectedItem.SubItems(4) * 1)
               var_anadir = TB_NEGADO_I.Anadir(var_empresa, var_unidad_organizacional, var_almacen, CDbl(lv_ordenes.selectedItem), lv_causas.selectedItem, lv_detalle.selectedItem, CDbl(txt_Cantidad), "", "", lv_ordenes.selectedItem.SubItems(2), var_precio)
               Set list_item = lv_negado.ListItems.Add(, , lv_detalle.selectedItem)
               list_item.SubItems(1) = Trim(lv_detalle.selectedItem.SubItems(1))
               list_item.SubItems(2) = txt_Cantidad
               list_item.SubItems(3) = Trim(lv_causas.selectedItem.SubItems(1))
               list_item.SubItems(4) = lv_causas.selectedItem
               list_item.SubItems(5) = var_precio
               lv_detalle.selectedItem.SubItems(2) = lv_detalle.selectedItem.SubItems(2) - CDbl(txt_Cantidad)
            End If
            rs.Close
         End If
         lv_detalle.SetFocus
         frm_causas.Visible = False
      Else
         MsgBox "Cantidad Incorrecta", vbOKOnly, "ATENCION"
         lv_detalle.SetFocus
         frm_causas.Visible = False
      End If
   End If
   Exit Sub
salir:
'Resume
End Sub

Private Sub txt_jaula_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
End Sub

Private Sub txt_numero_embarque_KeyPress(KeyAscii As Integer)
   Dim list_item As ListItem
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If IsNumeric(txt_numero_embarque) Then
         If Trim(txt_numero_embarque) <> "" Then
            rs.Open "select * from tb_encabezado_embarques where inte_emb_embarque = " + txt_numero_embarque, cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               rsaux2.Open "Select * from tb_agentes where vcha_age_agente_id ='" + rs!VCHA_AGE_AGENTE_ID + "'", cnn, adOpenDynamic, adLockOptimistic
               txt_agente = rsaux2!VCHA_AGE_NOMBRE
               txt_jaula = rs!inte_jau_jaula_id
               rsaux2.Close
               rs.Close
               lv_ordenes.ListItems.Clear
               rs.Open "SELECT distinct inte_ors_orden_surtido,vcha_cli_nombre, inte_ped_numero FROM vw_negado_embarque WHERE INTE_EMB_EMBARQUE = " + txt_numero_embarque, cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  While Not rs.EOF
                     Set list_item = lv_ordenes.ListItems.Add(, , rs!INTE_ORS_ORDEN_SURTIDO)
                     list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                     list_item.SubItems(2) = IIf(IsNull(rs!inte_ped_numero), 0, rs!inte_ped_numero)
                     rs.MoveNext
                  Wend
                  lv_ordenes.SetFocus
               Else
                  lv_ordenes.ListItems.Clear
                  lv_detalle.ListItems.Clear
                  lv_negado.ListItems.Clear
                  MsgBox "El embarque no tiene movimientos asignados", vbOKOnly, "ATENCION"
               End If
               rs.Close
            Else
               rs.Close
               MsgBox "El número de embarque no existe", vbOKOnly, "ATENCION"
               txt_agente = ""
               lv_ordenes.ListItems.Clear
            End If
         End If
      Else
         MsgBox "Número incorrecto", vbOKOnly, "ATENCION"
         txt_numero_embarque = ""
      End If
   End If

End Sub
