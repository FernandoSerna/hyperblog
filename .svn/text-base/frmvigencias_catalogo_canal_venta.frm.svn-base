VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmvigencias_catalogo_canal_venta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.MonthView mes 
      Height          =   2370
      Left            =   1620
      TabIndex        =   21
      Top             =   1860
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   51249153
      CurrentDate     =   37581
   End
   Begin VB.TextBox txt_catalogo 
      Height          =   330
      Left            =   6480
      TabIndex        =   20
      Top             =   1350
      Width           =   975
   End
   Begin VB.Frame Frame5 
      Caption         =   " Vigencias "
      Height          =   825
      Left            =   120
      TabIndex        =   12
      Top             =   1035
      Width           =   5670
      Begin VB.TextBox txt_vigencia_fin 
         Height          =   315
         Left            =   3405
         TabIndex        =   17
         Top             =   300
         Width           =   1065
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   4485
         Picture         =   "frmvigencias_catalogo_canal_venta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Vigencias por canal de venta"
         Top             =   300
         Width           =   330
      End
      Begin VB.TextBox txt_vigencia_inicio 
         Height          =   315
         Left            =   1005
         TabIndex        =   14
         Top             =   300
         Width           =   1065
      End
      Begin VB.CommandButton cmd_vigencias 
         Height          =   315
         Left            =   2085
         Picture         =   "frmvigencias_catalogo_canal_venta.frx":1272
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Vigencias por canal de venta"
         Top             =   300
         Width           =   330
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Index           =   0
         Left            =   3090
         TabIndex        =   18
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Index           =   2
         Left            =   555
         TabIndex        =   15
         Top             =   360
         Width           =   420
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2490
      Left            =   135
      TabIndex        =   10
      Top             =   1845
      Width           =   5670
      Begin MSComctlLib.ListView lv_catalogos 
         Height          =   2235
         Left            =   45
         TabIndex        =   11
         Top             =   150
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   3942
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "fecha inicio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "fecha fin"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   45
         Top             =   2505
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
               Picture         =   "frmvigencias_catalogo_canal_venta.frx":24E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmvigencias_catalogo_canal_venta.frx":2DBE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmvigencias_catalogo_canal_venta.frx":3698
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmvigencias_catalogo_canal_venta.frx":3C34
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmvigencias_catalogo_canal_venta.frx":4510
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmvigencias_catalogo_canal_venta.frx":4DEA
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmvigencias_catalogo_canal_venta.frx":56C4
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmvigencias_catalogo_canal_venta.frx":57D6
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmvigencias_catalogo_canal_venta.frx":58E8
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmvigencias_catalogo_canal_venta.frx":59FA
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmvigencias_catalogo_canal_venta.frx":5B0C
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "  Canal de Venta "
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   405
      Width           =   5655
      Begin VB.ComboBox cmb_canales_venta 
         Height          =   315
         Left            =   1260
         TabIndex        =   19
         Top             =   210
         Width           =   4320
      End
      Begin VB.TextBox txt_canal_venta 
         Height          =   315
         Left            =   225
         TabIndex        =   8
         Top             =   210
         Width           =   1005
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2895
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   6
      Top             =   75
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5415
      Picture         =   "frmvigencias_catalogo_canal_venta.frx":5C1E
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1425
      Picture         =   "frmvigencias_catalogo_canal_venta.frx":6258
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1095
      Picture         =   "frmvigencias_catalogo_canal_venta.frx":635A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   765
      Picture         =   "frmvigencias_catalogo_canal_venta.frx":645C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   435
      Picture         =   "frmvigencias_catalogo_canal_venta.frx":652E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frmvigencias_catalogo_canal_venta.frx":6630
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   90
      TabIndex        =   9
      Top             =   240
      Width           =   5685
   End
End
Attribute VB_Name = "frmvigencias_catalogo_canal_venta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tipo_mes As Integer

Private Sub cmb_canales_venta_Click()
 txt_canal_venta = Obtener_llave(cnn, rs, "TB_CANALESVENTAS", "VCHA_CAN_NOMBRE", cmb_canales_venta, 0, "T")
End Sub

Private Sub cmb_canales_venta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_vigencia_inicio.SetFocus
   End If
End Sub

Private Sub cmd_eliminar_Click()
   Dim si As Integer
   si = MsgBox("¿Deseas eliminar el registro?", vbYesNo, "ATENCION")
   If si = 6 Then
      rs.Open "delete from tb_Catalogos_vigencias where vcha_cat_catalogo_id = '" + txt_catalogo + "' and vcha_can_canal_venta_id = '" + txt_canal_venta + "'", cnn, adOpenDynamic, adLockOptimistic
      lv_catalogos.ListItems.Remove (lv_catalogos.selectedItem.Index)
      lv_catalogos.selectedItem.Selected = True
   End If
End Sub

Private Sub cmd_guardar_Click()
   If IsDate(txt_vigencia_inicio) And IsDate(txt_vigencia_fin) Then
      rs.Open "select * from tb_catalogos_vigencias where vcha_emp_empresa_id = '" + txt_empresa + "' and vcha_cat_catalogo_id = '" + txt_catalogo + "' and vcha_can_canal_venta_id = '" + txt_canal_venta + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         rsaux2.Open "update tb_catalogos_vigencias set dtim_vig_vigencia_inicio = '" + txt_fecha_inicio + "' and dtim_vig_vigencia_fin = '" + txt_vigencia_fin + "' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_cat_catalogo_id = '" + txt_catalogo + "' and vcha_can_canal_venta_id = '" + txt_canal_venta + "'", cnn, adOpenDynamic, adLockOptimistic
      Else
         rsaux2.Open "insert into tb_catalogos_vigencias (vcha_emp_empresa_id, vcha_cat_catalogo_id, vcha_can_canal_venta_id, dtim_vig_fecha_inicio, dtim_vig_fecha_fin) values ('" + var_empresa + "','" + txt_catalogo + "', '" + txt_canal_venta + "', '" + txt_vigencia_inicio + "','" + txt_vigencia_fin + "')", cnn, adOpenDynamic, adLockOptimistic
         Dim list_item As ListItem
         Set list_item = lv_catalogos.ListItems.Add(, , txt_canal_venta)
         list_item.SubItems(1) = cmb_canales_venta
         list_item.SubItems(2) = txt_vigencia_inicio
         list_item.SubItems(3) = txt_vigencia_fin
         list_item.EnsureVisible
         list_item.Selected = True
      End If
      rs.Close
   End If
   txt_canal_venta.Enabled = False
   cmb_canales_venta.Enabled = False
End Sub

Private Sub cmd_nuevo_Click()
   txt_canal_venta.Enabled = True
   txt_canal_venta = ""
   cmb_canales_venta = ""
   cmb_canales_venta.Enabled = True
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmd_vigencias_Click()
   If IsDate(txt_vigencia_inicio) Then
      mes.Value = txt_vigencia_inicio
   Else
      mes.Value = Date
   End If
   var_tipo_mes = 1
   mes.Visible = True
End Sub

Private Sub Command1_Click()
   If IsDate(txt_vigencia_fin) Then
      mes.Value = txt_vigencia_fin
   Else
      mes.Value = Date
   End If
   var_tipo_mes = 2
   mes.Visible = True
End Sub

Sub pro_llena_listview1()
   Dim list_item As ListItem
   numero_items_rutas = 0
    rs.Open "select * from tb_catalogos_vigencias where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_Cat_catalogo_id = '" + txt_catalogo + "'", cnn, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
       While Not rs.EOF
           Set list_item = lv_catalogos.ListItems.Add(, , rs!vcha_can_canal_venta_id)
           rsaux2.Open "select * from tb_canalesventas where vcha_can_canal_venta_id = '" + rs!vcha_can_canal_venta_id + "'", cnn, adOpenDynamic, adLockOptimistic
           list_item.SubItems(1) = IIf(IsNull(rsaux2!vcha_can_nombre), "", rsaux2!vcha_can_nombre)
           rsaux2.Close
           list_item.SubItems(2) = IIf(IsNull(rs!dtim_vig_fecha_inicio), "", rs!dtim_vig_fecha_inicio)
           list_item.SubItems(3) = IIf(IsNull(rs!dtim_vig_fecha_fin), "", rs!dtim_vig_fecha_fin)
           rs.MoveNext
           numero_items_rutas = numero_items_rutas + 1
       Wend
       txt_canal_venta = lv_catalogos.selectedItem
       cmb_canales_venta = lv_catalogos.selectedItem.SubItems(1)
       txt_vigencia_inicio = lv_catalogos.selectedItem.SubItems(2)
       txt_vigencia_fin = lv_catalogos.selectedItem.SubItems(3)
    End If
    rs.Close
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   txt_catalogo = frmcatalogos.txt_catalogos(0)
   txt_canal_venta.Enabled = False
   cmb_canales_venta.Enabled = False
   mes.Visible = False
   rs.Open "select * from tb_canalesventas order by vcha_Can_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
   Call RecsetToCombo(cmb_canales_venta.hwnd, rs, 1)
   rs.Close
   Call pro_llena_listview1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_vigencias_catalogo_canal_Venta)
End Sub

Private Sub lv_catalogos_ItemClick(ByVal Item As MSComctlLib.ListItem)
   txt_canal_venta = lv_catalogos.selectedItem
   cmb_canales_venta = lv_catalogos.selectedItem.SubItems(1)
   txt_vigencia_inicio = lv_catalogos.selectedItem.SubItems(2)
   txt_vigencia_fin = lv_catalogos.selectedItem.SubItems(3)
End Sub

Private Sub mes_DateDblClick(ByVal DateDblClicked As Date)
   If var_tipo_mes = 1 Then
      txt_vigencia_inicio = mes.Value
   End If
   If var_tipo_mes = 2 Then
      txt_vigencia_fin = mes.Value
   End If
   mes.Visible = False
End Sub

Private Sub mes_LostFocus()
   mes.Visible = False
End Sub

Private Sub txt_canal_venta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      cmb_canales_venta.SetFocus
   End If
End Sub

Private Sub txt_canal_venta_LostFocus()
   rs.Open "select * from tb_canalesventas where vcha_can_canal_venta_id = '" + txt_canal_venta + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      cmb_canales_venta = rs!vcha_can_nombre
   Else
      MsgBox "Clave de canal de venta incorrecta", vbOKOnly, "ATENCION"
      txt_canal_venta = ""
      cmb_canales_venta = ""
   End If
   rs.Close
End Sub

Private Sub txt_vigencia_fin_LostFocus()
   If Not IsDate(txt_vigencia_fin) Then
      MsgBox "Fecha Invalida", vbOKOnly, "ATENCION"
      txt_vigencia_fin = ""
   End If
End Sub

Private Sub txt_vigencia_inicio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_vigencia_fin.SetFocus
   End If
End Sub

Private Sub txt_vigencia_inicio_LostFocus()
   If Not IsDate(txt_vigencia_inicio) Then
      MsgBox "Fecha Invalida", vbOKOnly, "ATENCION"
      txt_vigencia_inicio = ""
   End If
End Sub
