VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmpromociones_inicio_catalogo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Promociones por inicio de catálogo"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   5805
   Begin MSComCtl2.MonthView mes 
      Height          =   2370
      Left            =   1680
      TabIndex        =   0
      Top             =   1605
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   19660801
      CurrentDate     =   37581
   End
   Begin VB.Frame Frame2 
      Caption         =   " Artículo "
      Height          =   615
      Left            =   60
      TabIndex        =   24
      Top             =   1080
      Width           =   5655
      Begin VB.ComboBox cmb_articulos 
         Height          =   315
         Left            =   1545
         TabIndex        =   10
         Top             =   210
         Width           =   4020
      End
      Begin VB.TextBox txt_articulo 
         Height          =   315
         Left            =   90
         TabIndex        =   9
         Top             =   210
         Width           =   1440
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   15
      Picture         =   "frmpromociones_inicio_catalogo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   345
      Picture         =   "frmpromociones_inicio_catalogo.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   675
      Picture         =   "frmpromociones_inicio_catalogo.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1005
      Picture         =   "frmpromociones_inicio_catalogo.frx":02D6
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1335
      Picture         =   "frmpromociones_inicio_catalogo.frx":03D8
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5325
      Picture         =   "frmpromociones_inicio_catalogo.frx":04DA
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2805
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   22
      Top             =   75
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   "  Canal de Venta "
      Height          =   615
      Left            =   30
      TabIndex        =   21
      Top             =   405
      Width           =   5655
      Begin VB.TextBox txt_canal_venta 
         Height          =   315
         Left            =   90
         TabIndex        =   7
         Top             =   210
         Width           =   1440
      End
      Begin VB.ComboBox cmb_canales_venta 
         Height          =   315
         Left            =   1545
         TabIndex        =   8
         Top             =   210
         Width           =   4035
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2490
      Left            =   45
      TabIndex        =   20
      Top             =   3075
      Width           =   5670
      Begin MSComctlLib.ListView lv_articulos 
         Height          =   2235
         Left            =   60
         TabIndex        =   16
         Top             =   150
         Width           =   5520
         _ExtentX        =   9737
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Inicio"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fin"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descuento"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Nombre artículo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Canal de venta"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Nombre canal venta"
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
               Picture         =   "frmpromociones_inicio_catalogo.frx":0B14
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmpromociones_inicio_catalogo.frx":13EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmpromociones_inicio_catalogo.frx":1CC8
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmpromociones_inicio_catalogo.frx":2264
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmpromociones_inicio_catalogo.frx":2B40
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmpromociones_inicio_catalogo.frx":341A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmpromociones_inicio_catalogo.frx":3CF4
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmpromociones_inicio_catalogo.frx":3E06
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmpromociones_inicio_catalogo.frx":3F18
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmpromociones_inicio_catalogo.frx":402A
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmpromociones_inicio_catalogo.frx":413C
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   " Vigencias "
      Height          =   735
      Left            =   30
      TabIndex        =   17
      Top             =   1695
      Width           =   5670
      Begin VB.CommandButton cmd_vigencias 
         Height          =   315
         Left            =   2280
         Picture         =   "frmpromociones_inicio_catalogo.frx":424E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Vigencias por canal de venta"
         Top             =   300
         Width           =   330
      End
      Begin VB.TextBox txt_vigencia_inicio 
         Height          =   315
         Left            =   1140
         TabIndex        =   11
         Top             =   300
         Width           =   1125
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   4920
         Picture         =   "frmpromociones_inicio_catalogo.frx":54C0
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Vigencias por canal de venta"
         Top             =   300
         Width           =   330
      End
      Begin VB.TextBox txt_vigencia_fin 
         Height          =   315
         Left            =   3780
         TabIndex        =   13
         Top             =   300
         Width           =   1125
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Index           =   2
         Left            =   690
         TabIndex        =   19
         Top             =   360
         Width           =   420
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Index           =   0
         Left            =   3465
         TabIndex        =   18
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   0
      TabIndex        =   23
      Top             =   255
      Width           =   5685
   End
   Begin VB.Frame Frame6 
      Height          =   570
      Left            =   60
      TabIndex        =   25
      Top             =   2505
      Width           =   5670
      Begin VB.TextBox txt_descuento 
         Height          =   315
         Left            =   2550
         TabIndex        =   15
         Top             =   180
         Width           =   1065
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Descuento:"
         Height          =   195
         Index           =   1
         Left            =   1650
         TabIndex        =   26
         Top             =   210
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmpromociones_inicio_catalogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tipo_mes As Integer

Private Sub cmb_articulos_Click()
   txt_Articulo = Obtener_llave(cnn, rs, "TB_ARTICULOS", "VCHA_ART_NOMBRE_ESPAÑOL", cmb_articulos, 0, "T")
End Sub

Private Sub cmb_articulos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_vigencia_inicio.SetFocus
   End If
End Sub

Private Sub cmb_canales_venta_Click()
   txt_canal_venta = Obtener_llave(cnn, rs, "TB_CANALESVENTAS", "VCHA_CAN_NOMBRE", cmb_canales_venta, 0, "T")
End Sub

Private Sub cmb_canales_venta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_Articulo.SetFocus
   End If
End Sub

Private Sub cmd_deshacer_Click()
   
   x = 1
End Sub

Private Sub cmd_eliminar_Click()
   x = 1
End Sub

Private Sub cmd_guardar_Click()
   If Trim(txt_descuento) = "" Then
      txt_descuento = 0
   End If
   If Not IsNumeric(txt_descuento) Then
      txt_descuento = 0
   End If
   var_opcion_seguridad = 2
   var_acepta_seguridad = 1
   If var_global_permiso3 = 1 Then
      var_acepta_seguridad = 2
      If var_global_permiso4 = 1 Then
         frmpasswords2.Show 1
      Else
         frmpasswords.Show 1
      End If
   End If
   If var_acepta_seguridad = 1 Then
      If Trim(txt_canal_venta) <> "" And Trim(txt_Articulo) <> "" And Trim(txt_vigencia_inicio) <> "" And Trim(txt_vigencia_fin) <> "" Then
         rs.Open "select * from TB_PROMOCIONES WHERE vcha_emp_empresa_id = '" + var_empresa + "' and vcha_can_canal_venta_id = '" + txt_canal_venta + "' and vcha_art_articulo_id = '" + txt_Articulo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            rsaux2.Open "update tb_promociones set dtim_pro_fecha_inicio = '" + txt_vigencia_inicio + "', dtim_pro_fecha_fin = '" + txt_vigencia_fin + "', floa_pro_Descuento = " + txt_descuento, cnn, adOpenDynamic, adLockOptimistic
            MsgBox "La información a actualizado correctamente", vbOKOnly, "ATENCION"
         Else
            rsaux2.Open "insert into tb_promociones (vcha_emp_empresa_id, vcha_can_canal_venta_id, vcha_art_articulo_id, dtim_pro_fecha_inicio, dtim_pro_fecha_fin, floa_pro_descuento) values ('" + var_empresa + "', '" + txt_canal_venta + "', '" + txt_Articulo + "', '" + txt_vigencia_inicio + "', '" + txt_vigencia_fin + "', " + txt_descuento + ")", cnn, adOpenDynamic, adLockOptimistic
            MsgBox "La información se a agregado correctamente", vbOKOnly, "ATENCION"
         End If
         rs.Close
         txt_canal_venta.Enabled = False
         cmb_canales_venta.Enabled = False
         txt_Articulo.Enabled = False
         cmb_articulos.Enabled = False
      Else
         MsgBox "Información Incompleta", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub cmd_imprimir_Click()
   x = 1
End Sub

Private Sub cmd_nuevo_Click()
   txt_canal_venta.Enabled = True
   cmb_canales_venta.Enabled = True
   txt_Articulo.Enabled = True
   cmb_articulos.Enabled = True
   txt_canal_venta = ""
   cmb_canales_venta = ""
   txt_Articulo = ""
   cmb_articulos = ""
   txt_canal_venta.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmd_vigencias_Click()
   var_tipo_mes = 1
   If IsDate(txt_vigencia_inicio) Then
      mes.Value = txt_vigencia_inicio
   Else
      mes = Date
   End If
   mes.Visible = True
   mes.SetFocus
End Sub

Private Sub Command1_Click()
   var_tipo_mes = 2
   If IsDate(txt_vigencia_fin) Then
      mes = txt_vigencia_fin
   Else
      mes = Date
   End If
   mes.Visible = True
   mes.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 78 Then
      cmd_nuevo_Click
   End If
   If Shift = 4 And KeyCode = 71 Then
      cmd_guardar_Click
   End If
   If Shift = 4 And KeyCode = 68 Then
      cmd_deshacer_Click
   End If
   If Shift = 4 And KeyCode = 69 Then
      cmd_eliminar_Click
   End If
   If Shift = 4 And KeyCode = 73 Then
      cmd_imprimir_Click
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 1000
   Left = 2900
   mes.Visible = False
   rs.Open "select * from tb_canalesventas order by vcha_Can_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
   Call RecsetToCombo(cmb_canales_venta.hwnd, rs, 1)
   rs.Close
   rs.Open "select * from tb_articulos order by vcha_art_nombre_español", cnn, adOpenDynamic, adLockBatchOptimistic
   Call RecsetToCombo(cmb_articulos.hwnd, rs, 1)
   rs.Close
   txt_canal_venta.Enabled = False
   cmb_canales_venta.Enabled = False
   txt_Articulo.Enabled = False
   cmb_articulos.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_promociones_inicio_catalogo)
   var_swpassword = False
   var_modifica_registro = False
End Sub

Private Sub mes_DateDblClick(ByVal DateDblClicked As Date)
   If var_tipo_mes = 1 Then
      txt_vigencia_inicio = mes.Value
      mes.Visible = False
   End If
   If var_tipo_mes = 2 Then
      txt_vigencia_fin = mes.Value
      mes.Visible = False
   End If
End Sub

Private Sub mes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      mes.Visible = False
   End If
End Sub

Private Sub mes_LostFocus()
   mes.Visible = False
End Sub

Private Sub txt_articulo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      cmb_articulos.SetFocus
   End If
End Sub

Private Sub txt_articulo_LostFocus()
   If Trim(txt_Articulo) <> "" Then
      rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_Articulo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         cmb_articulos = rs!vcha_art_nombre_español
      Else
         cmb_articulos = ""
         MsgBox "Clave de artículo incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   End If
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
   If Trim(txt_canal_venta) <> "" Then
      rs.Open "select * from tb_canalesventas where vcha_can_canal_venta_id = '" + txt_canal_venta + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         cmb_canales_venta = rs!vcha_can_nombre
         rsaux2.Open "select * from VW_PROMOCIONES WHERE VCHA_CAN_CANAL_VENTA_ID = '" + txt_canal_venta + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux2.EOF Then
            While Not rsaux2.EOF
                  Set list_item = lv_rangos.ListItems.Add(, , rsaux2!vcha_art_articulo_id)
                  list_item.SubItems(1) = IIf(IsNull(rsaux2!dtim_pro_fecha_inicio), "", rsaux2!dtim_pro_fecha_inicio)
                  list_item.SubItems(2) = IIf(IsNull(rsaux2!dtim_pro_fecha_fin), "", rsaux2!dtim_pro_fecha_fin)
                  list_item.SubItems(3) = IIf(IsNull(rsaux2!floa_pro_Descuento), 0, rsaux2!floa_pro_Descuento)
                  list_item.SubItems(4) = IIf(IsNull(rsaux2!vcha_art_nombre_español), "", rsaux2!vcha_art_nombre_español)
                  list_item.SubItems(5) = IIf(IsNull(rsaux2!vcha_can_canal_venta_id), "", rsaux2!vcha_can_canal_venta_id)
                  list_item.SubItems(6) = IIf(IsNull(rsaux2!vcha_can_nombre), "", rsaux2!vcha_can_nombre)
                  rsaux2.MoveNext:
                  numero_items_rangos = numero_items_rangos + 1
            Wend
         End If
         rsaux2.Close
      Else
         cmb_canales_venta = ""
         MsgBox "Clave de canal de venta incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   End If
End Sub

Private Sub txt_descuento_LostFocus()
   If Trim(txt_descuento) = "" Then
      txt_descuento = 0
   End If
   If Not IsNumeric(txt_descuento) Then
      MsgBox "Descuento Incorrecto", vbOKOnly, "ATENCION"
      txt_descuento = 0
   End If
End Sub

Private Sub txt_vigencia_fin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_descuento.SetFocus
   End If
End Sub

Private Sub txt_vigencia_fin_LostFocus()
   If Trim(txt_vigencia_fin) <> "" Then
      If Not IsDate(txt_vigencia_fin) Then
         txt_vigencia_fin = ""
      End If
   End If
End Sub

Private Sub txt_vigencia_inicio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_vigencia_fin.SetFocus
   End If
End Sub

Private Sub txt_vigencia_inicio_LostFocus()
   If Trim(txt_vigencia_inicio) <> "" Then
      If Not IsDate(txt_vigencia_inicio) Then
         txt_vigencia_inicio = ""
      End If
   End If
End Sub
