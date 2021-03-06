VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmtraspasos_calidad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Devoluciones de Calidad"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11670
   Begin VB.CommandButton cmd_nota_credito 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1740
      Picture         =   "frmtraspasos_calidad.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Generar nota de cr?dito"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1410
      Picture         =   "frmtraspasos_calidad.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Aplicar nota de cr?dito"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1080
      Picture         =   "frmtraspasos_calidad.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Enviar correo de devolucion de tiendas"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   1425
      TabIndex        =   26
      Top             =   660
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   30
         TabIndex        =   27
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
         TabIndex        =   28
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   750
      Picture         =   "frmtraspasos_calidad.frx":0486
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Imprimir Movimiento Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1080
      Picture         =   "frmtraspasos_calidad.frx":0588
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Cancelar Movimiento"
      Top             =   15
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11175
      Picture         =   "frmtraspasos_calidad.frx":068A
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_cerrar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   420
      Picture         =   "frmtraspasos_calidad.frx":0CC4
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Cerrar Movimiento"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmtraspasos_calidad.frx":0DC6
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Height          =   90
      Left            =   90
      TabIndex        =   18
      Top             =   990
      Width           =   11475
   End
   Begin VB.Frame frm_almacenes 
      Height          =   3225
      Left            =   4095
      TabIndex        =   14
      Top             =   2100
      Width           =   5625
      Begin MSComctlLib.ListView lv_almacenes 
         Height          =   2865
         Left            =   45
         TabIndex        =   15
         Top             =   300
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   5054
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripci?n"
            Object.Width           =   8290
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "marca"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "rechazo"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         Caption         =   "  Almacenes"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   5610
      End
   End
   Begin VB.Frame Frame2 
      Height          =   90
      Left            =   60
      TabIndex        =   11
      Top             =   300
      Width           =   11490
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos del Movimiento Origen "
      Height          =   1080
      Left            =   105
      TabIndex        =   8
      Top             =   1080
      Width           =   11460
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   11025
         Picture         =   "frmtraspasos_calidad.frx":0EC8
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   705
         Width           =   330
      End
      Begin VB.TextBox txt_nombre_movimiento 
         Height          =   315
         Left            =   1710
         TabIndex        =   29
         Top             =   345
         Width           =   4050
      End
      Begin VB.TextBox txt_referencia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7170
         TabIndex        =   19
         Top             =   675
         Width           =   3660
      End
      Begin VB.TextBox txt_nombre_almacen 
         Enabled         =   0   'False
         Height          =   300
         Left            =   7170
         TabIndex        =   13
         Top             =   360
         Width           =   4200
      End
      Begin VB.TextBox txt_numero 
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   675
         Width           =   1755
      End
      Begin VB.TextBox txt_movimiento 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   345
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "N?mero Origen:"
         Height          =   195
         Left            =   5970
         TabIndex        =   20
         Top             =   735
         Width           =   1110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Almacen Origen:"
         Height          =   195
         Left            =   5970
         TabIndex        =   12
         Top             =   405
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "N?mero:"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   735
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Movimiento:"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   405
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Art?culos"
      Height          =   3615
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   11445
      Begin MSComctlLib.ListView lv_articulos 
         Height          =   3285
         Left            =   45
         TabIndex        =   7
         Top             =   210
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   5794
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "C?digo"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripci?n"
            Object.Width           =   6050
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Almacen"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nombre Almacen"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "N?mero Movimiento"
            Object.Width           =   3237
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "consecutivo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "destino"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "marca"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "rechazado"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Moneda"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Tipo Cambio"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Causas de Devoluci?n"
      Height          =   1485
      Left            =   120
      TabIndex        =   4
      Top             =   5790
      Width           =   5805
      Begin MSComctlLib.ListView lv_detalle_real 
         Height          =   1200
         Left            =   45
         TabIndex        =   5
         Top             =   240
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   2117
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
            Text            =   "C?digo"
            Object.Width           =   2249
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripci?n"
            Object.Width           =   7735
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "consecutivo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "destino"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "marca"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   " Causas de Rechazo "
      Height          =   1485
      Left            =   5955
      TabIndex        =   0
      Top             =   5790
      Width           =   5595
      Begin MSComctlLib.ListView lv_detalle_rechazo 
         Height          =   1200
         Left            =   45
         TabIndex        =   3
         Top             =   240
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   2117
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
            Text            =   "C?digo"
            Object.Width           =   2214
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripci?n"
            Object.Width           =   7320
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "consecutivo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "destino"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "marca"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   4500
      Top             =   420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasos_calidad.frx":10DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasos_calidad.frx":19B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasos_calidad.frx":2292
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasos_calidad.frx":282E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasos_calidad.frx":310A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasos_calidad.frx":39E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasos_calidad.frx":42BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasos_calidad.frx":43D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasos_calidad.frx":44E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasos_calidad.frx":45F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtraspasos_calidad.frx":4706
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   0
      Top             =   15
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
      Left            =   810
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
      Left            =   90
      TabIndex        =   17
      Top             =   480
      Width           =   11445
   End
End
Attribute VB_Name = "frmtraspasos_calidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_almacen As String
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_ventana As Integer
Dim var_requiere_factura As Integer
Dim var_clave_moneda As String
Dim var_clave_cliente As String
Dim var_clave_agente As String
Dim var_clave_establecimiento As String
Dim var_clave_titular  As String
Dim var_a?o As Integer


Private Sub detalle_causas()
   Set TB_DEVOLUCIONES_ALMACEN = New TB_DEVOLUCIONES_ALMACEN
   Dim var_numero_rechazo As Integer
   Dim list_item As ListItem
   lv_detalle_real.ListItems.Clear
   lv_detalle_rechazo.ListItems.Clear
   
   'rsaux3.Open "select * from vw_detalle_devolucion_real where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen + "' and vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero + " and vcha_art_articulo_id = '" + lv_articulos.SelectedItem + "' and inte_cde_consecutivo  = " + lv_articulos.SelectedItem.SubItems(2), cnn, adOpenDynamic, adLockOptimistic
   
   rsaux3.Open "select * from vw_detalle_devolucion_real where  vcha_alm_almacen_id = '" + var_almacen + "' and vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero + " and vcha_art_articulo_id = '" + lv_articulos.selectedItem + "' and inte_cde_consecutivo  = " + lv_articulos.selectedItem.SubItems(5), cnn, adOpenDynamic, adLockOptimistic
   If Not rsaux3.EOF Then
      While Not rsaux3.EOF
         Set list_item = lv_detalle_real.ListItems.Add(, , rsaux3!INTE_CDE_CAUSA_ID)
         list_item.SubItems(1) = IIf(IsNull(rsaux3!vcha_cde_nombre), "", rsaux3!vcha_cde_nombre)
         rsaux3.MoveNext
      Wend
   End If
   rsaux3.Close
   'rsaux3.Open "select * from vw_detalle_devolucion_RECHAZO where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen + "' and vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero + " and vcha_art_articulo_id = '" + lv_articulos.SelectedItem + "' and inte_cde_consecutivo  = " + lv_articulos.SelectedItem.SubItems(2), cnn, adOpenDynamic, adLockOptimistic
   
   rsaux3.Open "select * from vw_detalle_devolucion_RECHAZO where vcha_alm_almacen_id = '" + var_almacen + "' and vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero + " and vcha_art_articulo_id = '" + lv_articulos.selectedItem + "' and inte_cde_consecutivo  = " + lv_articulos.selectedItem.SubItems(5), cnn, adOpenDynamic, adLockOptimistic
   If Not rsaux3.EOF Then
      If lv_articulos.selectedItem.SubItems(8) <> "*" Then
         var_modifica = False
         rs.Open "SELECT * FROM TB_ALMACENES WHERE INTE_ALM_RECHAZO = 1", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_numero_rechazo = lv_articulos.selectedItem.SubItems(4)
            var_modifica = TB_DEVOLUCIONES_ALMACEN.Anadir(var_empresa, var_unidad_organizacional, var_almacen, txt_movimiento, CDbl(txt_numero), lv_articulos.selectedItem, lv_articulos.selectedItem.SubItems(5), rs!VCHA_ALM_ALMACEN_ID, CDbl(var_numero_rechazo))
            lv_articulos.selectedItem.SubItems(2) = rs!VCHA_ALM_ALMACEN_ID
            lv_articulos.selectedItem.SubItems(3) = rs!VCHA_ALM_NOMBRE
         Else
            MsgBox "No se a definido un almac?n de rechazo", vbOKOnly, "ATENCION"
         End If
         rs.Close
         lv_articulos.selectedItem.SubItems(8) = "*"
         lv_articulos.selectedItem.Bold = True
         lv_articulos.selectedItem.ListSubItems.item(1).Bold = True
         lv_articulos.selectedItem.ListSubItems.item(2).Bold = True
         lv_articulos.selectedItem.ListSubItems.item(3).Bold = True
         lv_articulos.selectedItem.ListSubItems.item(4).Bold = True
         lv_articulos.selectedItem.ForeColor = &HC0&
         lv_articulos.selectedItem.ListSubItems.item(1).ForeColor = &HC0&
         lv_articulos.selectedItem.ListSubItems.item(2).ForeColor = &HC0&
         lv_articulos.selectedItem.ListSubItems.item(3).ForeColor = &HC0&
         lv_articulos.selectedItem.ListSubItems.item(4).ForeColor = &HC0&
      End If
      While Not rsaux3.EOF
         Set list_item = lv_detalle_rechazo.ListItems.Add(, , rsaux3!inte_cRe_causa_id)
         list_item.SubItems(1) = rsaux3!vcha_cRe_nombre
         rsaux3.MoveNext
      Wend
   End If
   rsaux3.Close
   End Sub



Private Sub cmd_cerrar_Click()
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
   Dim var_posible As Boolean
   Dim i As Integer
   Dim j As Integer
   Dim var_almacen_Destino As String
   Dim var_numero_folio As Double
   Dim var_costo As Double
   Dim var_precio As Double
   Dim n As Integer
   Dim var_contador_tipo_cambio As Double
   Dim var_veces_tipo_cambio As Double
   Dim var_tipo_cambio_promedio As Double
   var_contador_tipo_cambio = 0
   var_veces_tipo_cambio = 0
   n = lv_articulos.ListItems.Count
   For i = 1 To n
      var_clave_moneda = lv_articulos.selectedItem.SubItems(9)
      lv_articulos.ListItems.item(i).Selected = True
      If Trim(lv_articulos.selectedItem.SubItems(8)) <> "*" Or Trim(lv_articulos.selectedItem.SubItems(2)) <> "" Then
         var_contador_tipo_cambio = var_contador_tipo_cambio + lv_articulos.selectedItem.SubItems(10)
         var_veces_tipo_cambio = var_veces_tipo_cambio + 1
      End If
   Next i
   If var_contador_tipo_cambio > 0 Then
      var_tipo_cambio_promedio = var_contador_tipo_cambio / var_veces_tipo_cambio
   End If
   var_posible = True
   j = lv_articulos.ListItems.Count
   For i = 1 To j
      lv_articulos.ListItems.item(i).Selected = True
      If lv_articulos.selectedItem.SubItems(3) = "" Then
         var_posible = False
      End If
   Next i
   If var_posible = True Then
      var_posible = True
      j = lv_articulos.ListItems.Count
      For i = 1 To j
         lv_articulos.ListItems.item(i).Selected = True
         If lv_articulos.selectedItem.SubItems(4) <> 0 Then
            var_posible = False
         End If
      Next i
      If var_posible = True Then
         Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
         Set TB_SALIDAS_INSERTA = New TB_SALIDAS_INSERTA
         Set TB_ENTRADAS_INSERTA = New TB_ENTRADAS_INSERTA
         Set TB_DET_DEV_REAL_IMPORTE = New TB_DET_DEV_REAL_IMPORTE
         Set TB_DEVOLUCIONES_NUM_DESTINO = New TB_DEVOLUCIONES_NUM_DESTINO
         Set TB_ENCABEZADO_CARTERA_I = New TB_ENCABEZADO_CARTERA_I
         Dim var_rechazado As Integer
         Dim var_descuento_1 As Double
         Dim var_descuento_2 As Double
         Dim var_descuento_3 As Double
         Dim var_imp_desc_1 As Double
         Dim var_imp_desc_2 As Double
         Dim var_imp_desc_3 As Double
         Dim var_tot_desc_1 As Double
         Dim var_tot_desc_2 As Double
         Dim var_tot_desc_3 As Double
         Dim var_iva As Double
         Dim var_imp_iva As Double
         Dim var_importe As Double
         Dim var_importe_neto As Double
         Dim var_importe_total As Double
         Dim var_subimporte As Double
         Dim var_cantidad As Double
         Dim var_tipo_Cambio As Double
         Dim var_si_correo As Integer
         rsaux3.Open "select distinct vcha_cde_destino from tb_devoluciones where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen + "' and vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
         var_si_correo = 0
         If Not rsaux3.EOF Then
            var_importe_neto = 0
            
            cnn.BeginTrans
            While Not rsaux3.EOF
               If Trim(rsaux3!VCHA_CDE_DESTINO) <> "15" Then
                  var_si_correo = 1
               End If
               var_almacen_Destino = rsaux3!VCHA_CDE_DESTINO
               
               rs.Open "select * from tb_almacenes where vcha_alm_almacen_id = '" + var_almacen_Destino + "'", cnn, adOpenDynamic, adLockOptimistic
               var_rechazado = IIf(IsNull(rs!inte_alm_rechazo), 0, rs!inte_alm_rechazo)
               rs.Close
               var_inserta = False
  '            var_inserta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen_origen, var_clave_movimiento, Now, var_numero_folio, var_orden_surtido, var_clave_cliente, "", var_almacen_origen, var_almacen_destino, "", var_clave_usuario_global, fun_NombrePc, 0, "", txt_archivo, var_clave_establecimiento, "B", var_clave_titular, var_clave_agente, var_descuento_1, var_descuento_2, var_descuento_3, var_clave_moneda, 0)
               var_insreta = TB_FOLIOS_MOVIMIENTOS.Anadir(var_empresa, var_unidad_organizacional, var_almacen, var_clave_movimiento, Now, var_numero_folio, 0, var_clave_cliente, "", var_almacen, var_almacen_Destino, "", var_clave_usuario_global, fun_NombrePc, "", "", "", var_clave_establecimiento, "B", var_clave_titular, var_clave_agente, 0, 0, 0, var_clave_moneda, var_tipo_cambio_promedio)
               var_numero_folio = var_numero_folio_regreso
               If var_almacen_Destino = "140000" Then
                  VAR_NUMERO_FOLIO_STR = CStr(var_numero_folio)
                  If Len(VAR_NUMERO_FOLIO_STR) = 1 Then
                     VAR_NUMERO_FOLIO_STR = "0000000" + VAR_NUMERO_FOLIO_STR
                  Else
                     If Len(VAR_NUMERO_FOLIO_STR) = 2 Then
                        VAR_NUMERO_FOLIO_STR = "000000" + VAR_NUMERO_FOLIO_STR
                     Else
                        If Len(VAR_NUMERO_FOLIO_STR) = 3 Then
                           VAR_NUMERO_FOLIO_STR = "00000" + VAR_NUMERO_FOLIO_STR
                        Else
                           If Len(VAR_NUMERO_FOLIO_STR) = 4 Then
                              VAR_NUMERO_FOLIO_STR = "0000" + VAR_NUMERO_FOLIO_STR
                           Else
                              If Len(VAR_NUMERO_FOLIO_STR) = 5 Then
                                 VAR_NUMERO_FOLIO_STR = "000" + VAR_NUMERO_FOLIO_STR
                              Else
                                 If Len(VAR_NUMERO_FOLIO_STR) = 6 Then
                                    VAR_NUMERO_FOLIO_STR = "00" + VAR_NUMERO_FOLIO_STR
                                 Else
                                    If Len(VAR_NUMERO_FOLIO_STR) = 7 Then
                                       VAR_NUMERO_FOLIO_STR = "0" + VAR_NUMERO_FOLIO_STR
                                    Else
                                       If Len(VAR_NUMERO_FOLIO_STR) = 8 Then
                                          VAR_NUMERO_FOLIO_STR = "0" + VAR_NUMERO_FOLIO_STR
                                       End If
                                    End If
                                 End If
                              End If
                           End If
                        End If
                     End If
                  End If
                  VAR_NUMERO_FOLIO_STR = "TC" + txt_movimiento + VAR_NUMERO_FOLIO_STR
                  rs.Open "UPDATE TB_DEVOLUCIONES SET VCHA_EMO_REFERENCIA ='" + VAR_NUMERO_FOLIO_STR + "' WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + Me.txt_movimiento + "' AND INTE_EMO_NUMERO = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
                  rs.Open "update tb_encabezado_movimientos set vcha_emo_referencia = '" + VAR_NUMERO_FOLIO_STR + "' where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organziacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
               End If
               rs.Open "select * from vw_acumulado_devoluciones where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen + "' and vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero + " and vcha_cde_destino = '" + var_almacen_Destino + "'", cnn, adOpenDynamic, adLockOptimistic
               var_importe_total = 0
               var_importe_neto = 0
               var_subimporte = 0
               If Not rs.EOF Then
                  While Not rs.EOF
                     var_importe = 0
                     var_costo = rs!floa_cde_costo
                     var_precio = rs!floa_cde_precio
                     var_a?o = rs!inte_dev_a?o
                     If var_rechazado = 0 Then
                        var_cantidad = IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)
                        var_descuento_1 = IIf(IsNull(rs!floa_cde_descuento_1), 0, rs!floa_cde_descuento_1)
                        var_descuento_2 = IIf(IsNull(rs!floa_cde_descuento_2), 0, rs!floa_cde_descuento_2)
                        var_descuento_3 = IIf(IsNull(rs!floa_cde_descuento_3), 0, rs!floa_cde_descuento_3)
                        var_iva = IIf(IsNull(rs!floa_cde_iva), 0, rs!floa_cde_iva)
                        var_importe_total = var_importe_total + (var_precio * Cantidad)
                        var_imp_descuento_1 = var_precio * (var_descuento_1 / 100)
                        var_imp_descuento_2 = ((var_precio - var_imp_descuento_1) * var_descuento_2 / 100)
                        var_imp_descuento_3 = (((var_precio - var_imp_descuento_1) - var_imp_descuento_2) * var_descuento_3 / 100)
                        var_importe = (var_precio - var_imp_descuento_1 - var_imp_descuento_2 - var_imp_descuento_3) * rs!Cantidad
                        var_importe_iva = var_importe_iva + (var_importe * (var_iva / 100))
                        var_subimporte = var_subimporte + var_importe
                        var_importe = var_importe * (1 + (var_iva / 100))
                        var_importe_neto = var_importe_neto + var_importe
                     End If
                     var_inserta = False
                     If var_almacen_Destino <> "140000" Then
                        rsaux.Open "insert into tb_entradas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_ent_numero, vcha_art_articulo_id, floa_ent_cantidad, floa_ent_costo, floa_ent_precio, floa_ent_descuento, VCHA_ENT_ALMACEN_ORIGEN,inte_ent_a?o) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen_Destino + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", '" + rs!VCHA_ART_ARTICULO_ID + "', " + CStr(rs!Cantidad_leida) + ", " + CStr(var_costo) + ", " + CStr(var_precio) + ", 0, '" + var_almacen + "', " + CStr(var_a?o) + ")", cnn, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux.Open "insert into tb_salidas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_alm_almacen_id, vcha_mov_movimiento_id, inte_sal_numero, vcha_art_articulo_id, floa_sal_cantidad, floa_sal_costo, floa_sal_precio, floa_sal_descuento, inte_sal_a?o) values('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + var_almacen + "', '" + var_clave_movimiento + "', " + CStr(var_numero_folio) + ", '" + rs!VCHA_ART_ARTICULO_ID + "', " + CStr(rs!Cantidad_leida) + ", " + CStr(var_costo) + ", " + CStr(var_precio) + ", 0," + CStr(var_a?o) + ")", cnn, adOpenDynamic, adLockOptimistic
                     If var_almacen_Destino = "140000" Then
                        rsaux.Open "select max(inte_com_consecutivo) from tb_Archivo_comparacion where vcha_com_referencia = '" + VAR_NUMERO_FOLIO_STR + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
                        Else
                           var_consecutivo = 1
                        End If
                        rsaux.Close
                        var_cadena = "insert into tb_archivo_comparacion (vcha_emp_empresa_id, vcha_uor_unidad_id, VCHA_aLM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_COM_NUMERO, DTIM_COM_FECHA, CHAR_COM_TIPO_PROVEEDOR, VCHA_COM_PROVEEDOR, VCHA_ART_ARTICULO_ID, FLOA_COM_COSTO, FLOA_COM_cANTIDAD_ENVIADA, FLOA_COM_CANTIDAD_RECIBIDA, VCHA_COM_REFERENCIA, INTE_COM_CONSECUTIVO, FLOA_COM_PESO, INTE_COM_A?O)"
                        var_cadena = var_cadena + "              VALUES ('" + var_empresa + "','" + var_unidad_organizacional + "', '8','TA',       " + CStr(var_numero_folio) + ",GETDATE(),   'T',                  '" + var_clave_cliente + "','" + rs!VCHA_ART_ARTICULO_ID + "'," + CStr(var_costo) + "," + CStr(rs!Cantidad_leida) + ",0,'" + VAR_NUMERO_FOLIO_STR + "'," + CStr(var_consecutivo) + ",0,2005)"
                        'MsgBox var_cadena
                        rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     End If
                     rs.MoveNext
                  Wend
                  var_modifica = False
                  var_modifica = TB_DEVOLUCIONES_NUM_DESTINO.Anadir(var_empresa, var_unidad_organizacional, var_almacen, txt_movimiento, txt_numero, var_almacen_Destino, CDbl(var_numero_folio), var_clave_movimiento)
                  rsaux4.Open "Update tb_encabezado_movimientos set DTIM_EMO_FECHA_FINALIZO = getdate(), char_emo_estatus = 'I' where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND vcha_alm_almacen_id = '" + var_almacen + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND INTE_EMO_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                  
                  j = lv_articulos.ListItems.Count
                  For i = 1 To j
                     lv_articulos.ListItems.item(i).Selected = True
                     If lv_articulos.selectedItem.SubItems(2) = var_almacen_Destino Then
                        lv_articulos.selectedItem.SubItems(4) = var_numero_folio
                     End If
                  Next i
               End If
               rs.Close
               rsaux3.MoveNext
            Wend
            rsaux3.MoveFirst
            If var_requiere_factura = 1 Then
               Call numero_letras(Round(var_importe_neto / var_tipo_cambio_promedio, 2), var_clave_moneda)
               var_modifica = False
               var_modifica = TB_DET_DEV_REAL_IMPORTE.Anadir(var_empresa, var_unidad_organizacional, var_almacen, txt_movimiento, txt_numero, canstr)
            End If
            cnn.CommitTrans
         End If
         rsaux3.Close
         If var_si_correo = 1 Then
            rsaux3.Open "select vcha_tcl_tipo_cliente_id, vcha_age_Agente_id, vcha_cli_nombre from tb_clientes where vcha_cli_clave_id = '" + var_clave_cliente + "'"
            If Not rsaux3.EOF Then
               var_tipo_cliente = IIf(IsNull(rsaux3!VCHA_TCL_TIPO_CLIENTE_ID), "", rsaux3!VCHA_TCL_TIPO_CLIENTE_ID)
               var_nombre_cliente = IIf(IsNull(rsaux3!VCHA_CLI_NOMBRE), "", rsaux3!VCHA_CLI_NOMBRE)
               If var_tipo_cliente = "FT" Then
                  rsaux4.Open "select vcha_age_email from tb_agentes where vcha_age_Agente_id = '" + var_clave_agente + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux4.EOF Then
                     var_correo_ft = IIf(IsNull(rsaux4!VCHA_AGE_EMAIL), "", rsaux4!VCHA_AGE_EMAIL)
                     If Trim(var_correo_ft) <> "" Then
                     
                        If MAPISession1.SessionID = 0 Then
                           MAPISession1.SignOn
                        End If
                        MAPIMessages1.SessionID = MAPISession1.SessionID
                        MAPIMessages1.Compose
                        MAPIMessages1.RecipDisplayName = var_correo_ft
                        MAPIMessages1.RecipAddress = var_correo_ft
                        MAPIMessages1.AddressResolveUI = True
                        MAPIMessages1.ResolveName
                        MAPIMessages1.MsgSubject = "Informaci?n de devoluci?n n?mero " + Trim(Me.txt_numero)
                        MAPIMessages1.MsgNoteText = "Se genero la devoluci?n n?mero " + Trim(Me.txt_numero) + " del cliente " + var_clave_cliente + "   " + Trim(var_nombre_cliente)
                        MAPIMessages1.Send True
                        If MAPISession1.SessionID > 0 Then
                           MAPISession1.SignOff
                        End If
                        
                     End If
                  End If
                  rsaux4.Close
               End If
            End If
            rsaux3.Close
         End If
         
         MsgBox "Se a terminado de cerrar el movimiento", vbOKOnly, "ATENCION"
      Else
         MsgBox "El movimiento ya fue cerrado", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se han asignado todos los art?culos", vbOKOnly, "ATENCION"
   End If
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
   Dim var_posible As Boolean
   Dim i As Integer
   Dim j As Integer
   Dim var_almacen_Destino As String
   Dim var_numero_folio As Double
   Dim var_costo As Double
   Dim var_precio As Double
   Dim n As Integer
   Dim var_contador_tipo_cambio As Double
   Dim var_veces_tipo_cambio As Double
   Dim var_tipo_cambio_promedio As Double
   var_contador_tipo_cambio = 0
   var_veces_tipo_cambio = 0
   n = lv_articulos.ListItems.Count
   If n > 0 Then
      For i = 1 To n
         var_clave_moneda = lv_articulos.selectedItem.SubItems(9)
         lv_articulos.ListItems.item(i).Selected = True
         If Trim(lv_articulos.selectedItem.SubItems(8)) <> "*" Then
            var_contador_tipo_cambio = var_contador_tipo_cambio + lv_articulos.selectedItem.SubItems(10)
            var_veces_tipo_cambio = var_veces_tipo_cambio + 1
         End If
      Next i
      If var_contador_tipo_cambio > 0 Then
         var_tipo_cambio_promedio = var_contador_tipo_cambio / var_veces_tipo_cambio
      End If
      var_posible = True
      j = lv_articulos.ListItems.Count
      For i = 1 To j
         lv_articulos.ListItems.item(i).Selected = True
         If lv_articulos.selectedItem.SubItems(3) = "" Then
            var_posible = False
         End If
      Next i
      If var_posible = True Then
         var_posible = True
         j = lv_articulos.ListItems.Count
         For i = 1 To j
            lv_articulos.ListItems.item(i).Selected = True
            If lv_articulos.selectedItem.SubItems(4) <> 0 Then
               var_posible = False
            End If
         Next i
         If var_posible = True Then
            MsgBox "El movimiento aun no a sido cerrado", vbOKOnly, "ATENCION"
         Else
         
            cnn.BeginTrans
            rs.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_REPORTE_DEVOLUCIONES_MOVIMIENTO", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
            Else
               var_consecutivo = 1
            End If
            rs.Close
            rs.Open "INSERT INTO TB_TEMP_REPORTE_DEVOLUCIONES_MOVIMIENTO (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            
            If Me.txt_movimiento = "DT" Then
               var_cadena = "INSERT INTO TB_TEMP_REPORTE_DEVOLUCIONES_MOVIMIENTO (INTE_TEM_CONSECUTIVO,FLOA_TEM_CANTIDAD, VCHA_EMP_EMPRESA_ID,                                                          VCHA_UOR_UNIDAD_ID,                             VCHA_ALM_ALMACEN_ID,          VCHA_MOV_MOVIMIENTO_ID,                      INTE_EMO_NUMERO,                                   VCHA_ART_ARTICULO_ID, VCHA_CLI_CLAVE_ID, VCHA_AGE_NOMBRE, VCHA_CLI_NOMBRE, VCHA_ALM_NOMBRE,VCHA_EMP_NOMBRE, VCHA_ART_NOMBRE_ESPA?OL, VCHA_MOV_NOMBRE, VCHA_AGE_AGENTE_ID,VCHA_COM_REFERENCIA) "
               var_cadena = var_cadena + " SELECT " + CStr(var_consecutivo) + " AS CONSECUTIVO, SUM(dbo.TB_DEVOLUCIONES.FLOA_DEV_CANTIDAD) AS CANTIDAD, dbo.TB_DEVOLUCIONES.VCHA_EMP_EMPRESA_ID, dbo.TB_DEVOLUCIONES.VCHA_UOR_UNIDAD_ID, dbo.TB_ALMACENES.vcha_alm_almacen_id, dbo.TB_DEVOLUCIONES.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_DEVOLUCIONES.INTE_EMO_NUMERO, dbo.TB_DEVOLUCIONES.VCHA_ART_ARTICULO_ID,  dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_PRO_PROVEEDOR_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, dbo.TB_EMPRESAS.VCHA_EMP_NOMBRE, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPA?OL, dbo.TB_MOVIMIENTOS.VCHA_MOV_NOMBRE, dbo.tb_agentes.vcha_age_agente_id, dbo.TB_DEVOLUCIONES.VCHA_EMO_REFERENCIA FROM         dbo.TB_ENCABEZADO_MOVIMIENTOS INNER JOIN dbo.TB_DEVOLUCIONES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_DEVOLUCIONES.VCHA_EMP_EMPRESA_ID AND "
               var_cadena = var_cadena + " dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_DEVOLUCIONES.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_DEVOLUCIONES.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_DEVOLUCIONES.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_DEVOLUCIONES.INTE_EMO_NUMERO INNER JOIN dbo.TB_CLIENTES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_PRO_PROVEEDOR_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID INNER JOIN Dbo.TB_AGENTES ON dbo.TB_CLIENTES.VCHA_AGE_AGENTE_ID = dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID INNER JOIN dbo.TB_EMPRESAS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_EMPRESAS.VCHA_EMP_EMPRESA_ID INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_DEVOLUCIONES.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID INNER JOIN dbo.TB_ALMACENES ON dbo.TB_DEVOLUCIONES.VCHA_CDE_DESTINO = dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID INNER JOIN "
               var_cadena = var_cadena + " dbo.TB_MOVIMIENTOS ON dbo.TB_DEVOLUCIONES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID GROUP BY dbo.TB_DEVOLUCIONES.VCHA_EMP_EMPRESA_ID, dbo.TB_DEVOLUCIONES.VCHA_UOR_UNIDAD_ID, dbo.tb_almacenes.VCHA_ALM_ALMACEN_ID, dbo.TB_DEVOLUCIONES.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_DEVOLUCIONES.INTE_EMO_NUMERO, dbo.TB_DEVOLUCIONES.VCHA_ART_ARTICULO_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.vcha_pro_proveedor_id, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_PRO_PROVEEDOR_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, dbo.TB_EMPRESAS.VCHA_EMP_NOMBRE, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPA?OL, dbo.TB_DEVOLUCIONES.VCHA_CDE_DESTINO, dbo.TB_MOVIMIENTOS.VCHA_MOV_NOMBRE, dbo.tb_agentes.vcha_age_agente_id, dbo.TB_DEVOLUCIONES.VCHA_EMO_REFERENCIA  HAVING      (dbo.TB_DEVOLUCIONES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND "
               var_cadena = var_cadena + " (dbo.TB_DEVOLUCIONES.VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "') AND (dbo.TB_DEVOLUCIONES.INTE_EMO_NUMERO = " + Me.txt_numero + ") AND (dbo.TB_DEVOLUCIONES.VCHA_MOV_MOVIMIENTO_ID = '" + Me.txt_movimiento + "') ORDER BY dbo.TB_DEVOLUCIONES.VCHA_EMP_EMPRESA_ID, dbo.TB_DEVOLUCIONES.VCHA_UOR_UNIDAD_ID, dbo.TB_DEVOLUCIONES.INTE_EMO_NUMERO, dbo.TB_DEVOLUCIONES.VCHA_CDE_DESTINO, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, dbo.TB_Articulos.VCHA_ART_NOMBRE_ESPA?OL "
               rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            Else
               var_cadena = "INSERT INTO TB_TEMP_REPORTE_DEVOLUCIONES_MOVIMIENTO (INTE_TEM_CONSECUTIVO,FLOA_TEM_CANTIDAD, VCHA_EMP_EMPRESA_ID,                                                          VCHA_UOR_UNIDAD_ID,                             VCHA_ALM_ALMACEN_ID,          VCHA_MOV_MOVIMIENTO_ID,                      INTE_EMO_NUMERO,                                   VCHA_ART_ARTICULO_ID, VCHA_CLI_CLAVE_ID, VCHA_AGE_NOMBRE, VCHA_CLI_NOMBRE, VCHA_ALM_NOMBRE,VCHA_EMP_NOMBRE, VCHA_ART_NOMBRE_ESPA?OL, VCHA_MOV_NOMBRE, VCHA_AGE_AGENTE_ID,VCHA_COM_REFERENCIA) "
               var_cadena = var_cadena + " SELECT " + CStr(var_consecutivo) + " AS CONSECUTIVO, SUM(dbo.TB_DEVOLUCIONES.FLOA_DEV_CANTIDAD) AS CANTIDAD, dbo.TB_DEVOLUCIONES.VCHA_EMP_EMPRESA_ID, dbo.TB_DEVOLUCIONES.VCHA_UOR_UNIDAD_ID, dbo.tb_almacenes.VCHA_ALM_ALMACEN_ID, dbo.TB_DEVOLUCIONES.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_DEVOLUCIONES.INTE_EMO_NUMERO, dbo.TB_DEVOLUCIONES.VCHA_ART_ARTICULO_ID,  dbo.TB_ENCABEZADO_MOVIMIENTOS.vcha_cli_clave_id, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, dbo.TB_EMPRESAS.VCHA_EMP_NOMBRE, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPA?OL, dbo.TB_MOVIMIENTOS.VCHA_MOV_NOMBRE, dbo.tb_agentes.vcha_age_agente_id, dbo.TB_DEVOLUCIONES.VCHA_EMO_REFERENCIA FROM         dbo.TB_ENCABEZADO_MOVIMIENTOS INNER JOIN dbo.TB_DEVOLUCIONES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_DEVOLUCIONES.VCHA_EMP_EMPRESA_ID AND "
               var_cadena = var_cadena + " dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_DEVOLUCIONES.VCHA_UOR_UNIDAD_ID AND  dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_DEVOLUCIONES.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_DEVOLUCIONES.INTE_EMO_NUMERO INNER JOIN dbo.TB_CLIENTES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.vcha_cli_clave_id = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID INNER JOIN Dbo.TB_AGENTES ON dbo.TB_CLIENTES.VCHA_AGE_AGENTE_ID = dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID INNER JOIN dbo.TB_EMPRESAS ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_EMPRESAS.VCHA_EMP_EMPRESA_ID INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_DEVOLUCIONES.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID INNER JOIN dbo.TB_ALMACENES ON dbo.TB_DEVOLUCIONES.VCHA_CDE_DESTINO = dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID INNER JOIN "
               var_cadena = var_cadena + " dbo.TB_MOVIMIENTOS ON dbo.TB_DEVOLUCIONES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID GROUP BY dbo.TB_DEVOLUCIONES.VCHA_EMP_EMPRESA_ID, dbo.TB_DEVOLUCIONES.VCHA_UOR_UNIDAD_ID, dbo.tb_almacenes.VCHA_ALM_ALMACEN_ID, dbo.TB_DEVOLUCIONES.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_DEVOLUCIONES.INTE_EMO_NUMERO, dbo.TB_DEVOLUCIONES.VCHA_ART_ARTICULO_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.vcha_cli_clave_id, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_PRO_PROVEEDOR_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, dbo.TB_EMPRESAS.VCHA_EMP_NOMBRE, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPA?OL, dbo.TB_DEVOLUCIONES.VCHA_CDE_DESTINO, dbo.TB_MOVIMIENTOS.VCHA_MOV_NOMBRE, dbo.tb_agentes.vcha_age_agente_id, dbo.TB_DEVOLUCIONES.VCHA_EMO_REFERENCIA  HAVING      (dbo.TB_DEVOLUCIONES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND "
               var_cadena = var_cadena + " (dbo.TB_DEVOLUCIONES.VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "') AND (dbo.TB_DEVOLUCIONES.INTE_EMO_NUMERO = " + Me.txt_numero + ") AND (dbo.TB_DEVOLUCIONES.VCHA_MOV_MOVIMIENTO_ID = '" + Me.txt_movimiento + "') ORDER BY dbo.TB_DEVOLUCIONES.VCHA_EMP_EMPRESA_ID, dbo.TB_DEVOLUCIONES.VCHA_UOR_UNIDAD_ID, dbo.TB_DEVOLUCIONES.INTE_EMO_NUMERO, dbo.TB_DEVOLUCIONES.VCHA_CDE_DESTINO, dbo.TB_ALMACENES.VCHA_ALM_NOMBRE, dbo.TB_Articulos.VCHA_ART_NOMBRE_ESPA?OL "
               rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            End If
            
            rs.Open "delete from TB_TEMP_REPORTE_DEVOLUCIONES_MOVIMIENTO where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_emp_empresa_id is null", cnn, adOpenDynamic, adLockOptimistic
            Set reporte = appl.OpenReport(App.Path + "\rep_devoluciones_movimiento.rpt")
            reporte.RecordSelectionFormula = "{TB_TEMP_REPORTE_DEVOLUCIONES_MOVIMIENTO.inte_tem_consecutivo} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Movimientos"
            frmvistasprevias.Show 1
            Set reporte = Nothing
         
         
         
         
            rs.Open "DELETE FROM TB_TEMP_REPORTE_DEVOLUCIONES_MOVIMIENTO WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            'rsaux3.Open "select distinct inte_cde_numero_destino from tb_devoluciones where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen + "' and vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
            'While Not rsaux3.EOF
            '      var_numero_folio = rsaux3!inte_cde_numero_destino
            '      Set reporte = appl.OpenReport(App.Path + "\rep_salidas_traspasos.rpt")
            '      reporte.RecordSelectionFormula = "{VW_SALIDAS_TRASPASOS.vcha_uor_unidad_id} = '" + var_unidad_organizacional + "' and {VW_SALIDAS_TRASPASOS.VCHA_EMO_ALMACEN_ORIGEN} = '" + var_almacen + "' and {VW_SALIDAS_TRASPASOS.INTE_EMO_NUMERO} = " + Str(var_numero_folio) + " and {VW_SALIDAS_TRASPASOS.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "' and {VW_SALIDAS_TRASPASOS.vcha_emp_empresa_id} = '" + var_empresa + "' and {VW_SALIDAS_TRASPASOS.vcha_uor_unidad_id} = '" + var_unidad_organizacional + "'"
            '      frmvistasprevias.cr.ReportSource = reporte
            '      For ntablas = 1 To reporte.Database.Tables.Count
            '          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            '      Next ntablas
            '      frmvistasprevias.cr.ViewReport
            '      frmvistasprevias.Caption = "Reporte de Movimientos"
            '      frmvistasprevias.Show 1
            '      Set reporte = Nothing
            '      rsaux4.Open "update tb_encabezado_movimientos set inte_emo_impresiones = inte_emo_impresiones + 1 where VCHA_EMO_ALMACEN_ORIGEN = '" + var_almacen + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
            '      rsaux3.MoveNext
            'Wend
            'rsaux3.Close
         End If
      Else
         MsgBox "No se han asignado todos los art?culos", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a seleccionado ningun movimiento", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nota_credito_Click()
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
                      rs.Open "update tb_encabezado_cartera set inte_car_nota_credito_aplicada = 0 where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'DV' and vcha_Ser_Serie_id = '" + txt_serie + "' and inte_car_numero = " + CStr(var_numero_nota), cnn, adOpenDynamic, adLockOptimistic
                      rs.Open "update tb_series set inte_ser_nota_Credito = inte_ser_nota_credito + 1 where vcha_Ser_Serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                      rs.Open "insert into TB_NOTAS_CREDITO_DEVOLUCIONES (vcha_Emp_Empresa_id, vcha_uor_unidad_id, vcha_mov_movimiento_id, inte_emo_numero) values ('" + var_empresa + "', '" + var_unidad_organizacional + "', '" + Me.txt_movimiento + "', '" + Me.txt_numero + "')", cnn, adOpenDynamic, adLockOptimistic
                      rs.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
                      If Not rs.EOF Then
                         var_moneda_local = IIf(IsNull(rs!inte_mon_moneda_local), 0, rs!inte_mon_moneda_local)
                      End If
                      rs.Close

                      Open (App.Path & "\renombra" + Trim(var_serie) + Trim(Str(var_numero_nota)) + ".bat") For Output As #2
                      Print #2, "ren " + var_ruta_documentos_electronicos + "\" + Trim(var_serie) + Trim(Str(var_numero_nota)) + ".fi " + Trim(var_serie) + Trim(Str(var_numero_nota)) + ".ff"
                      Close #2
                        
                      Open (var_ruta_documentos_electronicos & "\" + Trim(var_serie) + Trim(Str(var_numero_nota)) + ".fi") For Output As #1
                      'Open ("c:\NC_" + Trim(var_serie) + Trim(Str(rs!inte_car_numero)) + ".fi") For Output As #1
                      var_cadena = "Outputmode=" + Chr(13) + "<Factura>" + Chr(13) + "<Comprobante>" + Chr(13) + "Version=2.0" + Chr(13) + "Serie=" + var_serie + Chr(13) + "folio=" + CStr(var_numero_nota) + Chr(13)
                      var_a?o = CStr(Year(Now))
                      var_mes = CStr(Month(Now))
                      var_dia = CStr(Day(Now))
                      var_hora = CStr(Hour(Now))
                      var_minuto = CStr(Minute(Now))
                      var_segundo = CStr(Second(Now))
                      If Len(var_a?o) = 2 Then
                         var_a?o = "20" + var_a?o
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
                           
                           
                      var_cadena_fecha = CStr(var_a?o) + "-" + CStr(var_mes) + "-" + CStr(var_dia) + "T" + CStr(var_hora) + ":" + CStr(var_minuto) + ":" + CStr(var_segundo)
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
                               var_linea = var_nombre_articulo
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
                      Call numero_letras(var_importe_total, var_clave_moneda)
                      var_cadena = var_cadena + "cant_letra=" + canstr + Chr(13)
                      var_cadena = var_cadena + "factoriva=" + CStr(16) + "%" + Chr(13)
                      rsaux1.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + IIf(IsNull(var_clave_moneda), "", var_clave_moneda) + "'", cnn, adOpenDynamic, adLockOptimistic
                      var_cadena = var_cadena + "moneda=" + IIf(IsNull(rsaux1!vcha_mon_nombre_plural), "", rsaux1!vcha_mon_nombre_plural) + Chr(13)
                      rsaux1.Close
                      var_cadena = var_cadena + "tipodeCambio=" + CStr(1) + Chr(13)
                      var_cadena = var_cadena + "pedido=" + Chr(13)
                      var_cadena = var_cadena + "Embarque=" + Chr(13)
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
                      var_cadena = var_cadena + "<addenda>" + Chr(13)
                      var_cadena = var_cadena + "</addenda>" + Chr(13) + Chr(13)
                      var_cadena = var_cadena + "</Factura>"
                      Print #1, var_cadena
                      Close #1
                       
                      var_Archivo = App.Path & "\renombra" + Trim(var_serie) + Trim(Str(var_numero_nota)) + ".bat"
                      x = Shell(var_Archivo, vbHide)
                      MsgBox "Se a terminado el proceso", vbOKOnly, "ATENCION"
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
   txt_movimiento = ""
   txt_numero = ""
   txt_nombre_almacen = ""
   txt_referencia = ""
   lv_articulos.ListItems.Clear
   txt_numero.Enabled = True
   txt_movimiento.Enabled = True
   txt_nombre_movimiento.Enabled = True
   lv_detalle_real.ListItems.Clear
   lv_detalle_rechazo.ListItems.Clear
   txt_nombre_movimiento = ""
   txt_movimiento.SetFocus
End Sub

Private Sub cmd_todos_Click()
   For var_i = 1 To lv_articulos.ListItems.Count
      lv_articulos.ListItems.item(var_i).Selected = True
      Dim i As Integer
      i = lv_articulos.selectedItem.Index
      If lv_articulos.selectedItem.SubItems(8) = "*" Then
         MsgBox "No puede modificar el almacen", vbOKOnly, "ATENCION"
      Else
         If lv_articulos.selectedItem.SubItems(7) = "*" Then
            lv_articulos.selectedItem.SubItems(7) = ""
            lv_articulos.selectedItem.Bold = False
            lv_articulos.selectedItem.ListSubItems.item(1).Bold = False
            lv_articulos.selectedItem.ListSubItems.item(2).Bold = False
            lv_articulos.selectedItem.ListSubItems.item(3).Bold = False
            lv_articulos.selectedItem.ListSubItems.item(4).Bold = False
            lv_articulos.selectedItem.ForeColor = &H0&
            lv_articulos.selectedItem.ListSubItems.item(1).ForeColor = &H0&
            lv_articulos.selectedItem.ListSubItems.item(2).ForeColor = &H0&
            lv_articulos.selectedItem.ListSubItems.item(3).ForeColor = &H0&
            lv_articulos.selectedItem.ListSubItems.item(4).ForeColor = &H0&
         Else
            lv_articulos.selectedItem.SubItems(7) = "*"
            lv_articulos.selectedItem.Bold = True
            lv_articulos.selectedItem.ListSubItems.item(1).Bold = True
            lv_articulos.selectedItem.ListSubItems.item(2).Bold = True
            lv_articulos.selectedItem.ListSubItems.item(3).Bold = True
            lv_articulos.selectedItem.ListSubItems.item(4).Bold = True
            lv_articulos.selectedItem.ForeColor = &HFF0000
            lv_articulos.selectedItem.ListSubItems.item(1).ForeColor = &HFF0000
            lv_articulos.selectedItem.ListSubItems.item(2).ForeColor = &HFF0000
            lv_articulos.selectedItem.ListSubItems.item(3).ForeColor = &HFF0000
            lv_articulos.selectedItem.ListSubItems.item(4).ForeColor = &HFF0000
         End If
      End If
   Next var_i
End Sub

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Command2_Click()
   Dim var_posible As Boolean
   Dim i As Integer
   Dim j As Integer
   Dim var_almacen_Destino As String
   Dim var_numero_folio As Double
   Dim var_costo As Double
   Dim var_precio As Double
   Dim n As Integer
   Dim var_contador_tipo_cambio As Double
   Dim var_veces_tipo_cambio As Double
   Dim var_tipo_cambio_promedio As Double
   var_contador_tipo_cambio = 0
   var_veces_tipo_cambio = 0
   n = lv_articulos.ListItems.Count
   If Trim(Me.txt_movimiento) <> "" Then
   If Trim(Me.txt_numero) <> "" Then
   For i = 1 To n
      var_clave_moneda = lv_articulos.selectedItem.SubItems(9)
      lv_articulos.ListItems.item(i).Selected = True
      If Trim(lv_articulos.selectedItem.SubItems(8)) <> "*" Or Trim(lv_articulos.selectedItem.SubItems(2)) <> "" Then
         var_contador_tipo_cambio = var_contador_tipo_cambio + lv_articulos.selectedItem.SubItems(10)
         var_veces_tipo_cambio = var_veces_tipo_cambio + 1
      End If
   Next i
         If var_contador_tipo_cambio > 0 Then
            var_tipo_cambio_promedio = var_contador_tipo_cambio / var_veces_tipo_cambio
         End If
         var_posible = True
         j = lv_articulos.ListItems.Count
         For i = 1 To j
            lv_articulos.ListItems.item(i).Selected = True
            If lv_articulos.selectedItem.SubItems(3) = "" Then
               var_posible = False
            End If
         Next i
         If var_posible = True Then
            var_posible = True
            j = lv_articulos.ListItems.Count
            For i = 1 To j
               lv_articulos.ListItems.item(i).Selected = True
               If lv_articulos.selectedItem.SubItems(4) <> 0 Then
                  var_posible = False
               End If
            Next i
            var_posible = True
            If var_posible = True Then
               Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
               Set TB_SALIDAS_INSERTA = New TB_SALIDAS_INSERTA
               Set TB_ENTRADAS_INSERTA = New TB_ENTRADAS_INSERTA
               Set TB_DET_DEV_REAL_IMPORTE = New TB_DET_DEV_REAL_IMPORTE
               Set TB_DEVOLUCIONES_NUM_DESTINO = New TB_DEVOLUCIONES_NUM_DESTINO
               Set TB_ENCABEZADO_CARTERA_I = New TB_ENCABEZADO_CARTERA_I
               Dim var_rechazado As Integer
               Dim var_descuento_1 As Double
               Dim var_descuento_2 As Double
               Dim var_descuento_3 As Double
               Dim var_imp_desc_1 As Double
               Dim var_imp_desc_2 As Double
               Dim var_imp_desc_3 As Double
               Dim var_tot_desc_1 As Double
               Dim var_tot_desc_2 As Double
               Dim var_tot_desc_3 As Double
               Dim var_iva As Double
               Dim var_imp_iva As Double
               Dim var_importe As Double
               Dim var_importe_neto As Double
               Dim var_importe_total As Double
               Dim var_subimporte As Double
               Dim var_cantidad As Double
               Dim var_tipo_Cambio As Double
               Dim var_si_correo As Integer
               rsaux3.Open "select distinct vcha_cde_destino from tb_devoluciones where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_alm_almacen_id = '" + var_almacen + "' and vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
               var_si_correo = 0
               If Not rsaux3.EOF Then
                  var_importe_neto = 0
                  While Not rsaux3.EOF
                     If Trim(rsaux3!VCHA_CDE_DESTINO) <> "15" Then
                        var_si_correo = 1
                     End If
                     rsaux3.MoveNext
                  Wend
               End If
               rsaux3.Close
               If var_si_correo = 1 Then
                  rsaux3.Open "select vcha_cli_clave_id, vcha_tcl_tipo_cliente_id, vcha_age_Agente_id, vcha_cli_nombre from tb_clientes where vcha_cli_clave_id = '" + var_clave_cliente + "'"
                  If Not rsaux3.EOF Then
                     var_tipo_cliente = IIf(IsNull(rsaux3!VCHA_TCL_TIPO_CLIENTE_ID), "", rsaux3!VCHA_TCL_TIPO_CLIENTE_ID)
                     var_nombre_cliente = IIf(IsNull(rsaux3!VCHA_CLI_NOMBRE), "", rsaux3!VCHA_CLI_NOMBRE)
                     If var_tipo_cliente = "FT" Or var_tipo_cliente = "VT" Then
                        rsaux4.Open "select vcha_age_email from tb_agentes where vcha_age_Agente_id = '" + var_clave_agente + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux4.EOF Then
                           var_correo_ft = IIf(IsNull(rsaux4!VCHA_AGE_EMAIL), "", rsaux4!VCHA_AGE_EMAIL)
                           If Trim(var_correo_ft) <> "" Then
                           
                              If MAPISession1.SessionID = 0 Then
                                 MAPISession1.SignOn
                              End If
                              MAPIMessages1.SessionID = MAPISession1.SessionID
                              MAPIMessages1.Compose
                              MAPIMessages1.RecipDisplayName = var_correo_ft
                              MAPIMessages1.RecipAddress = var_correo_ft
                              MAPIMessages1.AddressResolveUI = True
                              MAPIMessages1.ResolveName
                              MAPIMessages1.MsgSubject = "Informaci?n de devoluci?n n?mero " + Trim(Me.txt_numero)
                              MAPIMessages1.MsgNoteText = "Se genero la devoluci?n n?mero " + Trim(Me.txt_numero) + " del cliente " + var_clave_cliente + "   " + Trim(var_nombre_cliente)
                              MAPIMessages1.Send True
                              If MAPISession1.SessionID > 0 Then
                                 MAPISession1.SignOff
                              End If
                              
                           End If
                        End If
                        rsaux4.Close
                     End If
                  End If
                  rsaux3.Close
               End If
         
            Else
               MsgBox "El movimiento ya fue cerrado", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "No se han asignado todos los art?culos", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "No se a indicado un n?mero de movimiento"
      End If
   Else
      MsgBox "No se a seleccionado un movimiento", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Command3_Click()
   If Me.txt_movimiento <> "" Then
      If IsNumeric(Me.txt_numero) Then
         var_clave_movimiento_nc = Me.txt_movimiento
         var_numero_nc = Me.txt_numero
         var_aplicar_nota_credito = 1
         frmnotas_credito.Show
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
   End If
   If Shift = 4 And KeyCode = 73 Then
      cmd_imprimir_Click
   End If
   If Shift = 4 And KeyCode = 67 Then
   End If
End Sub

Private Sub Form_Load()
   var_posible_kanban = 0
   var_cadena_seguridad = ""
   frm_lista.Visible = False
   Top = 0
   Left = 0
   var_ventana = 0
   frm_almacenes.Visible = False: var_ventana = 0
   txt_referencia = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_traspasos_calidad)
End Sub

Private Sub lv_almacenes_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione Enter para indicar el almacen al que se efectuara el traspaso"
End Sub

Private Sub lv_almacenes_KeyPress(KeyAscii As Integer)
   Dim i As Integer
   Dim j As Integer
   If KeyAscii = 13 Then
      si = MsgBox("Confirmar el traspaso de mercancia a: " + lv_almacenes.selectedItem.SubItems(1), vbYesNo, "ATENCION")
      If si = 6 Then
         j = lv_articulos.ListItems.Count
         Set TB_DEVOLUCIONES_ALMACEN = New TB_DEVOLUCIONES_ALMACEN
         For i = 1 To j
            lv_articulos.ListItems.item(i).Selected = True
            If lv_articulos.selectedItem.SubItems(7) = "*" Then
               var_modifica = False
               If lv_almacenes.selectedItem = "8" Then
                  var_modifica = TB_DEVOLUCIONES_ALMACEN.Anadir(var_empresa, var_unidad_organizacional, var_almacen, txt_movimiento, txt_numero, lv_articulos.selectedItem, lv_articulos.selectedItem.SubItems(5), "ATC", 0)
                  rsaux3.Open "select * from tb_almacenes where vcha_alm_almacen_id = '" + lv_almacenes.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux3.EOF Then
                     rs.Open "SELECT * FROM TB_ALMACENES WHERE VCHA_ALM_ALMACEN_ID = 'ATC'", cnn, adOpenDynamic, adLockOptimistic
                     VAR_NOMBRE_ALMACEN = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
                     rs.Close
                     lv_articulos.selectedItem.SubItems(2) = "ATC"
                     lv_articulos.selectedItem.SubItems(3) = VAR_NOMBRE_ALMACEN
                     lv_articulos.selectedItem.SubItems(8) = 0
                  End If
                  rsaux3.Close
                  lv_articulos.selectedItem.SubItems(7) = ""
                  lv_articulos.selectedItem.Bold = False
                  lv_articulos.selectedItem.ListSubItems.item(1).Bold = False
                  lv_articulos.selectedItem.ListSubItems.item(2).Bold = False
                  lv_articulos.selectedItem.ListSubItems.item(3).Bold = False
                  lv_articulos.selectedItem.ListSubItems.item(4).Bold = False
                  lv_articulos.selectedItem.ForeColor = &H0&
                  lv_articulos.selectedItem.ListSubItems.item(1).ForeColor = &H0&
                  lv_articulos.selectedItem.ListSubItems.item(2).ForeColor = &H0&
                  lv_articulos.selectedItem.ListSubItems.item(3).ForeColor = &H0&
                  lv_articulos.selectedItem.ListSubItems.item(4).ForeColor = &H0&
               Else
                  var_modifica = TB_DEVOLUCIONES_ALMACEN.Anadir(var_empresa, var_unidad_organizacional, var_almacen, txt_movimiento, txt_numero, lv_articulos.selectedItem, lv_articulos.selectedItem.SubItems(5), lv_almacenes.selectedItem, 0)
                  rsaux3.Open "select * from tb_almacenes where vcha_alm_almacen_id = '" + lv_almacenes.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux3.EOF Then
                     lv_articulos.selectedItem.SubItems(2) = lv_almacenes.selectedItem
                     lv_articulos.selectedItem.SubItems(3) = rsaux3!VCHA_ALM_NOMBRE
                     lv_articulos.selectedItem.SubItems(8) = IIf(IsNull(rsaux3!inte_alm_rechazo), 0, rsaux3!inte_alm_rechazo)
                  End If
                  rsaux3.Close
                  lv_articulos.selectedItem.SubItems(7) = ""
                  lv_articulos.selectedItem.Bold = False
                  lv_articulos.selectedItem.ListSubItems.item(1).Bold = False
                  lv_articulos.selectedItem.ListSubItems.item(2).Bold = False
                  lv_articulos.selectedItem.ListSubItems.item(3).Bold = False
                  lv_articulos.selectedItem.ListSubItems.item(4).Bold = False
                  lv_articulos.selectedItem.ForeColor = &H0&
                  lv_articulos.selectedItem.ListSubItems.item(1).ForeColor = &H0&
                  lv_articulos.selectedItem.ListSubItems.item(2).ForeColor = &H0&
                  lv_articulos.selectedItem.ListSubItems.item(3).ForeColor = &H0&
                  lv_articulos.selectedItem.ListSubItems.item(4).ForeColor = &H0&
               End If
            End If
         Next i
      End If
      frm_almacenes.Visible = False: var_ventana = 0
      lv_articulos.SetFocus
   End If
   If KeyAscii = 27 Then
      frm_almacenes.Visible = False: var_ventana = 0
      lv_articulos.SetFocus
   End If
End Sub

Private Sub lv_almacenes_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   frm_almacenes.Visible = False: var_ventana = 0
End Sub

Private Sub lv_articulos_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione Enter para seleccionar y F4 para indicar los almacenes a los que va a hacerse el traspaso"
End Sub

Private Sub lv_articulos_ItemClick(ByVal item As MSComctlLib.ListItem)
   Frmmenu2.StatusBar1.Panels(1) = "Presione Enter para seleccionar y F4 para indicar los almacenes a los que va a hacerse el traspaso"
   Call detalle_causas
End Sub

Private Sub lv_articulos_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 115 Then
      If (lv_articulos.selectedItem.SubItems(4) * 1) > 0 Then
         MsgBox "Imposible modificar el movimiento", vbOKOnly, "ATENCION"
      Else
         Dim list_item As ListItem
         If var_empresa = "06" Then
            rsaux3.Open "select distinct vcha_alm_almacen_id, vcha_alm_nombre, inte_alm_rechazo from vw_movimientos_almacenes WHERE vcha_mov_movimiento_id = '" + var_clave_movimiento + "'  order by vcha_alm_nombre", cnn, adOpenDynamic
            If Not rsaux3.EOF Then
               lv_almacenes.ListItems.Clear
               While Not rsaux3.EOF
                  Set list_item = lv_almacenes.ListItems.Add(, , rsaux3!VCHA_ALM_ALMACEN_ID)
                  list_item.SubItems(1) = IIf(IsNull(rsaux3!VCHA_ALM_NOMBRE), "", rsaux3!VCHA_ALM_NOMBRE)
                  list_item.SubItems(3) = IIf(IsNull(rsaux3!inte_alm_rechazo), 0, rsaux3!inte_alm_rechazo)
                  rsaux3.MoveNext
               Wend
               frm_almacenes.Visible = True: var_ventana = 1
               lv_almacenes.SetFocus
            Else
               MsgBox "El movimiento no tiene almacenes relacionados", vbOKOnly, "ATENCION"
            End If
            rsaux3.Close
         Else
            If var_empresa = "03" Then
               rsaux3.Open "select distinct vcha_alm_almacen_id, vcha_alm_nombre, inte_alm_rechazo from vw_movimientos_almacenes WHERE vcha_mov_movimiento_id = '" + var_clave_movimiento + "' order by vcha_alm_nombre", cnn, adOpenDynamic
            Else
               rsaux3.Open "select distinct vcha_alm_almacen_id, vcha_alm_nombre, inte_alm_rechazo from vw_movimientos_almacenes WHERE vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and vcha_emp_empresa_id = '" + var_empresa + "' order by vcha_alm_nombre", cnn, adOpenDynamic
            End If
            If Not rsaux3.EOF Then
               lv_almacenes.ListItems.Clear
               While Not rsaux3.EOF
                  Set list_item = lv_almacenes.ListItems.Add(, , rsaux3!VCHA_ALM_ALMACEN_ID)
                  list_item.SubItems(1) = IIf(IsNull(rsaux3!VCHA_ALM_NOMBRE), "", rsaux3!VCHA_ALM_NOMBRE)
                  list_item.SubItems(3) = IIf(IsNull(rsaux3!inte_alm_rechazo), 0, rsaux3!inte_alm_rechazo)
                  rsaux3.MoveNext
               Wend
               frm_almacenes.Visible = True: var_ventana = 1
               lv_almacenes.SetFocus
            Else
               MsgBox "El movimiento no tiene almacenes relacionados", vbOKOnly, "ATENCION"
            End If
            rsaux3.Close
         End If
      End If
   End If
End Sub

Private Sub lv_articulos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         Dim i As Integer
         i = lv_articulos.selectedItem.Index
         
         If lv_articulos.selectedItem.SubItems(8) = "*" Then
            MsgBox "No puede modificar el almacen", vbOKOnly, "ATENCION"
         Else
            If lv_articulos.selectedItem.SubItems(7) = "*" Then
               lv_articulos.selectedItem.SubItems(7) = ""
               lv_articulos.selectedItem.Bold = False
               lv_articulos.selectedItem.ListSubItems.item(1).Bold = False
               lv_articulos.selectedItem.ListSubItems.item(2).Bold = False
               lv_articulos.selectedItem.ListSubItems.item(3).Bold = False
               lv_articulos.selectedItem.ListSubItems.item(4).Bold = False
               lv_articulos.selectedItem.ForeColor = &H0&
               lv_articulos.selectedItem.ListSubItems.item(1).ForeColor = &H0&
               lv_articulos.selectedItem.ListSubItems.item(2).ForeColor = &H0&
               lv_articulos.selectedItem.ListSubItems.item(3).ForeColor = &H0&
               lv_articulos.selectedItem.ListSubItems.item(4).ForeColor = &H0&
            Else
               lv_articulos.selectedItem.SubItems(7) = "*"
               lv_articulos.selectedItem.Bold = True
               lv_articulos.selectedItem.ListSubItems.item(1).Bold = True
               lv_articulos.selectedItem.ListSubItems.item(2).Bold = True
               lv_articulos.selectedItem.ListSubItems.item(3).Bold = True
               lv_articulos.selectedItem.ListSubItems.item(4).Bold = True
               lv_articulos.selectedItem.ForeColor = &HFF0000
               lv_articulos.selectedItem.ListSubItems.item(1).ForeColor = &HFF0000
               lv_articulos.selectedItem.ListSubItems.item(2).ForeColor = &HFF0000
               lv_articulos.selectedItem.ListSubItems.item(3).ForeColor = &HFF0000
               lv_articulos.selectedItem.ListSubItems.item(4).ForeColor = &HFF0000
            End If
         End If
   End If
End Sub


Private Sub lv_articulos_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         txt_movimiento = lv_lista.selectedItem
         txt_nombre_movimiento = lv_lista.selectedItem.SubItems(1)
      Else
         txt_movimiento = ""
         txt_nombre_movimiento = ""
      End If
      Me.txt_nombre_movimiento.Enabled = True
      Me.txt_movimiento.Enabled = True
      txt_movimiento.SetFocus
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub txt_movimiento_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la informaci?n disponible"
End Sub

Private Sub txt_movimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      If var_clave_movimiento = "DC" Then
         rs.Open "select * from tb_movimientos where (char_mov_afectacion <> 'T' and inte_mov_causa_devolucion = 1 AND VCHA_MOV_MOVIMIENTO_ID = 'CA' OR VCHA_MOV_MOVIMIENTO_ID = 'CAPT')", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_MOV_MOVIMIENTO_ID)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_mov_nombre), "", rs!vcha_mov_nombre)
               rs.MoveNext
         Wend
         rs.Close
      Else
         rs.Open "select * from tb_movimientos where (char_mov_afectacion <> 'T' and inte_mov_causa_devolucion = 1 AND VCHA_MOV_MOVIMIENTO_ID <> 'CA') OR vcha_mov_movimiento_id = 'TEC'", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_MOV_MOVIMIENTO_ID)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_mov_nombre), "", rs!vcha_mov_nombre)
               rs.MoveNext
         Wend
         rs.Close
      End If
      lbl_lista = "Movimientos"
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

Private Sub txt_movimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      txt_nombre_movimiento.SetFocus
   End If
End Sub

Private Sub txt_movimiento_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_movimiento) <> "" Then
      If var_clave_movimiento = "DC" Then
         rs.Open "select * from tb_movimientos where (vcha_mov_movimiento_id ='" + txt_movimiento + "' and char_mov_afectacion <> 'T' AND (VCHA_MOV_MOVIMIENTO_ID = 'CA' OR VCHA_MOV_MOVIMIENTO_ID = 'CAPT') and inte_mov_causa_devolucion = 1)", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            txt_nombre_movimiento = rs!vcha_mov_nombre
            rs.Close
            txt_numero.Enabled = True
            txt_movimiento.Enabled = False
         Else
            rs.Close
            txt_movimiento = ""
            txt_nombre_movimiento = ""
            MsgBox "Clave de movimiento incorrecta", vbOKOnly, "ATENCION"
         End If
      Else
         rs.Open "select * from tb_movimientos where (vcha_mov_movimiento_id ='" + txt_movimiento + "' and char_mov_afectacion <> 'T' and inte_mov_causa_devolucion = 1 AND VCHA_MOV_MOVIMIENTO_ID <> 'CA') or vcha_mov_movimiento_id = 'TEC'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            If Me.txt_movimiento <> "CA" Then
            
               rs.Close
               rs.Open "SELECT * FROM TB_MOVIMIENTOS WHERE VCHA_MOV_MOVIMIENTO_ID = '" + Me.txt_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  txt_nombre_movimiento = rs!vcha_mov_nombre
               Else
                  txt_nombre_movimiento = ""
               End If
               rs.Close
               txt_numero.Enabled = True
               txt_movimiento.Enabled = False
            Else
               rs.Close
               txt_movimiento = ""
               txt_nombre_movimiento = ""
               MsgBox "Clave de movimiento incorrecta", vbOKOnly, "ATENCION"
            End If
         Else
            rs.Close
            txt_movimiento = ""
            txt_nombre_movimiento = ""
            MsgBox "Clave de movimiento incorrecta", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub

Private Sub txt_nombre_movimiento_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la informaci?n disponible"
End Sub

Private Sub txt_nombre_movimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      If var_clave_movimiento = "DC" Then
         rs.Open "select * from tb_movimientos where char_mov_afectacion <> 'T' and inte_mov_causa_devolucion = 1 AND VCHA_MOV_MOVIMIENTO_ID = 'CA'", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_MOV_MOVIMIENTO_ID)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_mov_nombre), "", rs!vcha_mov_nombre)
               rs.MoveNext
         Wend
         rs.Close
      Else
         rs.Open "select * from tb_movimientos where (char_mov_afectacion <> 'T' and inte_mov_causa_devolucion = 1 AND VCHA_MOV_MOVIMIENTO_ID <> 'CA') OR VCHA_MOV_MOVIMIENTO_ID = 'TEC'", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_MOV_MOVIMIENTO_ID)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_mov_nombre), "", rs!vcha_mov_nombre)
               rs.MoveNext
         Wend
         rs.Close
      End If
      lbl_lista = "Movimientos"
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

Private Sub txt_nombre_movimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txt_numero.Enabled = True Then
         txt_numero.SetFocus
      End If
   End If
End Sub

Private Sub txt_nombre_movimiento_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If Trim(txt_numero) <> "" Then
         lv_articulos.ListItems.Clear
         Dim var_posible As Boolean
         Dim var_consecutivo As Integer
         Dim var_contador_articulos As Integer
         var_posible = False
         rs.Open "select * from tb_movimientos where vcha_mov_movimiento_id = '" + txt_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
         var_requiere_factura = 0
         If Not rs.EOF Then
            var_requiere_factura = IIf(IsNull(rs!INTE_MOV_DEVOLUCION_FACTURA), 0, rs!INTE_MOV_DEVOLUCION_FACTURA)
         End If
         rs.Close
         rs.Open "select * from tb_devoluciones where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
         If rs.EOF Then
            rs.Close
            MsgBox "El movimiento no existe o no a sido clasificado", vbOKOnly, "ATENCION"
         Else
            var_posible = True
         End If
         If rs.State = 1 Then
            rs.Close
         End If
         If var_posible = True Then
            rs.Open "select * from tb_devoluciones where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + txt_numero + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic
            If Not rs.EOF Then
               If rs!CHAR_CDE_ESTATUS = "I" Or rs!CHAR_CDE_ESTATUS = "N" Then
                  var_clave_cliente = ""
                  var_calve_establecimiento = ""
                  var_clave_agente = ""
                  If rsaux3.State = 1 Then
                     rsaux3.Close
                  End If
                  rsaux3.Open "select * from tb_encabezado_movimientos where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_MOV_MOVIMIENTO_ID = '" + rs!VCHA_MOV_MOVIMIENTO_ID + "' and inte_emo_numero = " + Str(rs!INTE_EMO_NUMERO) + " and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux3.EOF Then
                     var_clave_cliente = IIf(IsNull(rsaux3!vcha_cli_clave_id), "", rsaux3!vcha_cli_clave_id)
                     var_clave_establecimiento = IIf(IsNull(rsaux3!vcha_ESB_ESTABLECIMIENTO_id), "", rsaux3!vcha_ESB_ESTABLECIMIENTO_id)
                     var_clave_agente = IIf(IsNull(rsaux3!VCHA_AGE_AGENTE_ID), "", rsaux3!VCHA_AGE_AGENTE_ID)
                     var_clave_titular = IIf(IsNull(rsaux3!vcha_tit_titular_id), "", rsaux3!vcha_tit_titular_id)
                  End If
                  rsaux3.Close
                  var_almacen = rs!VCHA_ALM_ALMACEN_ID
                  rsaux3.Open "select * from tb_almacenes where vcha_alm_almacen_id = '" + var_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
                  txt_nombre_almacen = rsaux3!VCHA_ALM_NOMBRE
                  txt_referencia = rs!vcha_cde_referencia
                  rsaux3.Close
                  var_contador_articulos = 0
                  While Not rs.EOF
                     var_estatus = IIf(IsNull(rs!CHAR_CDE_ESTATUS), "", rs!CHAR_CDE_ESTATUS)
                     var_nombre_articulo = ""
                     If rsaux2.State = 1 Then
                        rsaux2.Close
                     End If
                     rsaux2.Open "select * from tb_articulos where vcha_Art_articulo_id ='" + rs!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux2.EOF Then
                        var_nombre_articulo = rsaux2!vcha_Art_nombre_espa?ol
                     End If
                     rsaux2.Close
                     Set list_item = lv_articulos.ListItems.Add(, , rs!VCHA_ART_ARTICULO_ID)
                     list_item.SubItems(1) = Trim(var_nombre_articulo)
                     list_item.SubItems(2) = rs!VCHA_CDE_DESTINO
                     list_item.SubItems(9) = rs!vcha_mon_moneda_id
                     list_item.SubItems(10) = rs!floa_dev_tipo_cambio
                     If rs!VCHA_CDE_DESTINO <> "" Then
                        rsaux3.Open "select * from tb_almacenes where vcha_alm_almacen_id = '" + rs!VCHA_CDE_DESTINO + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux3.EOF Then
                           list_item.SubItems(3) = rsaux3!VCHA_ALM_NOMBRE
                        End If
                        rsaux3.Close
                     End If
                     list_item.SubItems(4) = IIf(IsNull(rs!inte_cde_numero_destino), 0, rs!inte_cde_numero_destino)
                     list_item.SubItems(5) = IIf(IsNull(rs!INTE_CDE_CONSECUTIVO), "", rs!INTE_CDE_CONSECUTIVO)
                     var_contador_articulos = var_contador_articulos + 1
                     rs.MoveNext
                  Wend
                  If var_contador_articulos > 12 Then
                     lv_articulos.ColumnHeaders(2).Width = 3200
                  Else
                     lv_articulos.ColumnHeaders(2).Width = 3429.92
                  End If
                  txt_numero.Enabled = False
                  txt_movimiento.Enabled = False
                  rs.Close
                  lv_articulos.SetFocus
                  Call detalle_causas
                  j = lv_articulos.ListItems.Count
                  For i = 1 To j
                      lv_articulos.ListItems.item(i).Selected = True
                      Call detalle_causas
                  Next i
                  lv_articulos.ListItems.item(1).Selected = True
               Else
                  rs.Close
                  MsgBox "El movimiento aun no a sido cerrado", vbOKOnly, "ATENCION"
               End If
            End If
         End If
      Else
         MsgBox "N?mero Incorrecto de Movimiento", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

