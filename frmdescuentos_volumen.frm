VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmdescuentos_volumen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rangos de Descuento por Volumen"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5955
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   120
      TabIndex        =   21
      Top             =   390
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   22
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
         TabIndex        =   23
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1455
      Picture         =   "frmdescuentos_volumen.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   795
      Picture         =   "frmdescuentos_volumen.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   4830
      Left            =   135
      TabIndex        =   17
      Top             =   2385
      Width           =   5670
      Begin MSComctlLib.ListView lv_descuentos 
         Height          =   4635
         Left            =   45
         TabIndex        =   18
         Top             =   150
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   8176
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
            Text            =   "Tipo asignación"
            Object.Width           =   2452
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Importe Inferior"
            Object.Width           =   2417
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Importe Superior"
            Object.Width           =   2417
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descuento"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Canal de Venta"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Colonias"
      Height          =   1950
      Left            =   150
      TabIndex        =   12
      Top             =   420
      Width           =   5655
      Begin VB.ComboBox cmb_tipo_asignacion 
         Height          =   315
         ItemData        =   "frmdescuentos_volumen.frx":01D4
         Left            =   1590
         List            =   "frmdescuentos_volumen.frx":01E4
         TabIndex        =   8
         Top             =   555
         Width           =   3975
      End
      Begin VB.TextBox txt_nombre_canal_venta 
         Height          =   315
         Left            =   2355
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   225
         Width           =   3210
      End
      Begin VB.TextBox txt_importe_superior 
         Height          =   315
         Left            =   1590
         TabIndex        =   10
         Top             =   1215
         Width           =   1380
      End
      Begin VB.TextBox txt_canal_venta 
         Height          =   315
         Left            =   1590
         TabIndex        =   6
         Top             =   225
         Width           =   750
      End
      Begin VB.TextBox txt_importe_inferior 
         Height          =   315
         Left            =   1590
         TabIndex        =   9
         Top             =   885
         Width           =   1380
      End
      Begin VB.TextBox txt_descuento 
         Height          =   315
         Left            =   1590
         TabIndex        =   11
         Top             =   1545
         Width           =   750
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Asignación:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   24
         Top             =   630
         Width           =   1410
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Index           =   3
         Left            =   2385
         TabIndex        =   20
         Top             =   1605
         Width           =   120
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Canal de Venta:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   285
         Width           =   1140
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Importe Inferior:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   945
         Width           =   1095
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Importe Superior:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   1275
         Width           =   1200
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Descuento:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   1605
         Width           =   825
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5460
      Picture         =   "frmdescuentos_volumen.frx":0214
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1125
      Picture         =   "frmdescuentos_volumen.frx":084E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   465
      Picture         =   "frmdescuentos_volumen.frx":0950
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Guardar Alt + G"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      Picture         =   "frmdescuentos_volumen.frx":0A52
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   15
      Width           =   330
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   2430
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdescuentos_volumen.frx":0B54
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdescuentos_volumen.frx":142E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdescuentos_volumen.frx":1D08
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdescuentos_volumen.frx":22A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdescuentos_volumen.frx":2B80
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdescuentos_volumen.frx":345A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdescuentos_volumen.frx":3D34
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdescuentos_volumen.frx":3E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdescuentos_volumen.frx":3F58
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdescuentos_volumen.frx":406A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   60
      TabIndex        =   19
      Top             =   285
      Width           =   5850
   End
End
Attribute VB_Name = "frmdescuentos_volumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tipo_Asignacion As String
Private Sub llena_listview()
   Dim list_item As ListItem
   Dim var_tipo_asignacion_nombre As String
   lv_descuentos.ListItems.Clear
   If Trim(txt_canal_venta) <> "" Then
      rs.Open "select * from tb_descuentos_volumen where vcha_can_canal_venta_id =  '" + txt_canal_venta + "'", cnn, adOpenDynamic, adLockOptimistic
      lv_descuentos.ListItems.Clear
      While Not rs.EOF
            var_tipo_Asignacion = IIf(IsNull(rs!CHAR_PRI_TIPO_ASIGNACION), "", rs!CHAR_PRI_TIPO_ASIGNACION)
            var_tipo_asignacion_nombre = ""
            If var_tipo_Asignacion = "C" Then
               var_tipo_asignacion_nombre = "CLIENTE"
            End If
            If var_tipo_Asignacion = "T" Then
               var_tipo_asignacion_nombre = "TITULAR"
            End If
            If var_tipo_Asignacion = "A" Then
               var_tipo_asignacion_nombre = "GRUPO ACTUAL"
            End If
            If var_tipo_Asignacion = "R" Then
               var_tipo_asignacion_nombre = "GRUPO REAL"
            End If
            Set list_item = lv_descuentos.ListItems.Add(, , var_tipo_asignacion_nombre)
            list_item.SubItems(1) = IIf(IsNull(rs!FLOA_DVO_IMPORTE_INFERIOR), "", rs!FLOA_DVO_IMPORTE_INFERIOR)
            list_item.SubItems(2) = IIf(IsNull(rs!FLOA_DVO_IMPORTE_SUPERIOR), "", rs!FLOA_DVO_IMPORTE_SUPERIOR)
            list_item.SubItems(3) = IIf(IsNull(rs!floa_dvo_descuento), "", rs!floa_dvo_descuento)
            list_item.SubItems(4) = IIf(IsNull(rs!vcha_can_canal_venta_id), "", rs!vcha_can_canal_venta_id)
            rs.MoveNext:
            numero_items_descuentos = numero_items_descuentos + 1
      Wend
      rs.Close
      var_n = lv_descuentos.ListItems.Count
      If var_n > 0 Then
         lv_descuentos.ListItems.Item(1).Selected = True
         Call pro_textos
         cmd_guardar.Enabled = True
         cmd_deshacer.Enabled = True
      Else
         cmd_guardar.Enabled = False
         cmd_deshacer.Enabled = False
      End If
   End If
End Sub

Private Sub pro_textos()
   Dim var_n As Integer
   var_n = lv_descuentos.ListItems.Count
   If var_n > 0 Then
      cmb_tipo_asignacion = lv_descuentos.selectedItem
      If cmb_tipo_asignacion = "CLIENTE" Then
         var_tipo_Asignacion = "C"
      End If
      If cmb_tipo_asignacion = "TITULAR" Then
         var_tipo_Asignacion = "T"
      End If
      If cmb_tipo_asignacion = "GRUPO ACTUAL" Then
         var_tipo_Asignacion = "A"
      End If
      If cmb_tipo_asignacion = "GRUPO REAL" Then
         var_tipo_Asignacion = "R"
      End If
      txt_importe_inferior = lv_descuentos.selectedItem.SubItems(1)
      txt_importe_superior = lv_descuentos.selectedItem.SubItems(2)
      txt_descuento = lv_descuentos.selectedItem.SubItems(3)
      txt_canal_venta = lv_descuentos.selectedItem.SubItems(4)
      rs.Open "select * from tb_canalesventas where vcha_Can_canal_venta_id = '" + txt_canal_venta + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_canal_venta = IIf(IsNull(rs!vcha_can_nombre), "", rs!vcha_can_nombre)
      Else
         txt_nombre_canal_venta = ""
      End If
      rs.Close
   End If
   var_numero_renglones = lv_descuentos.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_descuentos.ColumnHeaders(4).Width = 1000.18
   Else
      lv_descuentos.ColumnHeaders(4).Width = 1200.18
   End If
   txt_canal_venta.Enabled = False
   txt_nombre_canal_venta.Enabled = False
   txt_importe_inferior.Enabled = False
   txt_importe_superior.Enabled = False
   txt_descuento.Enabled = False
   cmb_tipo_asignacion.Enabled = False
End Sub

Private Sub cmb_tipo_asignacion_Click()
      If cmb_tipo_asignacion = "CLIENTE" Then
         var_tipo_Asignacion = "C"
      End If
      If cmb_tipo_asignacion = "TITULAR" Then
         var_tipo_Asignacion = "T"
      End If
      If cmb_tipo_asignacion = "GRUPO ACTUAL" Then
         var_tipo_Asignacion = "A"
      End If
      If cmb_tipo_asignacion = "GRUPO REAL" Then
         var_tipo_Asignacion = "R"
      End If
End Sub

Private Sub cmb_tipo_asignacion_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub cmd_deshacer_Click()
   Call pro_textos
End Sub

Private Sub cmd_eliminar_Click()
   Dim var_si As Integer
   var_si = MsgBox("¿Deseas eliminar el regsitro?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      If cmb_tipo_asignacion = "CLIENTE" Then
         var_tipo_Asignacion = "C"
      End If
      If cmb_tipo_asignacion = "TITULAR" Then
         var_tipo_Asignacion = "T"
      End If
      If cmb_tipo_asignacion = "GRUPO ACTUAL" Then
         var_tipo_Asignacion = "A"
      End If
      If cmb_tipo_asignacion = "GRUPO REAL" Then
         var_tipo_Asignacion = "R"
      End If
      rs.Open "delete from TB_DESCUENTOS_VOLUMEN where VCHA_CAN_CANAL_VENTA_ID = '" + txt_canal_venta + "' and CHAR_PRI_TIPO_ASIGNACION = '" + var_tipo_Asignacion + "' and FLOA_DVO_IMPORTE_INFERIOR = " + txt_importe_inferior + " and FLOA_DVO_IMPORTE_SUPERIOR  = " + txt_importe_superior + " and FLOA_DVO_DESCUENTO = " + txt_descuento, cnn, adOpenDynamic, adLockOptimistic
      MsgBox "Se a eliminado el registro", vbOKOnly, "ATENCION"
      Call llena_listview
   End If
End Sub

Private Sub cmd_guardar_Click()
   If txt_canal_venta <> "" And txt_importe_inferior <> "" And txt_importe_superior <> "" And txt_descuento <> "" Then
      If Trim(cmb_tipo_asignacion) <> "" Then
         If cmb_tipo_asignacion = "CLIENTE" Then
            var_tipo_Asignacion = "C"
         End If
         If cmb_tipo_asignacion = "TITULAR" Then
            var_tipo_Asignacion = "T"
         End If
         If cmb_tipo_asignacion = "GRUPO ACTUAL" Then
            var_tipo_Asignacion = "A"
         End If
         If cmb_tipo_asignacion = "GRUPO REAL" Then
            var_tipo_Asignacion = "R"
         End If
         rsaux.Open "SELECT * from TB_DESCUENTOS_VOLUMEN where VCHA_CAN_CANAL_VENTA_ID = '" + txt_canal_venta + "' and CHAR_PRI_TIPO_ASIGNACION = '" + var_tipo_Asignacion + "' and FLOA_DVO_IMPORTE_INFERIOR = " + txt_importe_inferior + " and FLOA_DVO_IMPORTE_SUPERIOR  = " + txt_importe_superior + " and FLOA_DVO_DESCUENTO = " + txt_descuento, cnn, adOpenDynamic, adLockOptimistic
         If rsaux.EOF Then
            rs.Open "insert into tb_descuentos_volumen (VCHA_CAN_CANAL_VENTA_ID, CHAR_PRI_TIPO_ASIGNACION, FLOA_DVO_IMPORTE_INFERIOR, FLOA_DVO_IMPORTE_SUPERIOR, FLOA_DVO_DESCUENTO) values ('" + txt_canal_venta + "', '" + var_tipo_Asignacion + "'," + txt_importe_inferior + ", " + txt_importe_superior + ", " + txt_descuento + ")"
            MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
            Call llena_listview
         Else
            MsgBox "El registro ya existe", vbOKOnly, "ATENCION"
         End If
         rsaux.Close
      Else
         MsgBox "No se a seleccionado un tipo de asignación", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Información Incompleta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
  i = 0
End Sub

Private Sub cmd_nuevo_Click()
   txt_canal_venta = ""
   txt_nombre_canal_venta = ""
   txt_importe_inferior = ""
   txt_importe_superior = ""
   txt_descuento = ""
   cmb_tipo_asignacion = ""
   txt_canal_venta.Enabled = True
   txt_nombre_canal_venta.Enabled = True
   txt_importe_inferior.Enabled = True
   txt_importe_superior.Enabled = True
   txt_descuento.Enabled = True
   cmb_tipo_asignacion.Enabled = True
   txt_canal_venta.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
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
   Top = 0
   Left = 2900
   frm_lista.Visible = False
   rs.Open "select * from tb_descuentos_volumen ", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      txt_canal_venta = IIf(IsNull(rs!vcha_can_canal_venta_id), "", rs!vcha_can_canal_venta_id)
   Else
      txt_canal_venta = ""
   End If
   rs.Close
   Call llena_listview
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_descuentos_volumen)
End Sub

Private Sub lv_descuentos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_descuentos, ColumnHeader)
End Sub

Private Sub lv_descuentos_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Call pro_textos
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         txt_canal_venta = lv_lista.selectedItem
         txt_nombre_canal_venta = lv_lista.selectedItem.SubItems(1)
      Else
         txt_canal_venta = ""
         txt_nombre_canal_venta = ""
      End If
      txt_canal_venta.SetFocus
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub txt_canal_venta_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_canal_venta_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_canal_venta_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_canalesventas order by vcha_can_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_can_canal_venta_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_can_nombre), "", rs!vcha_can_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CANALES DE VENTA"
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
   If KeyCode = 117 Then
      var_catalogo_articulos = True
      frmcanalesventas.Show
   End If
End Sub

Private Sub txt_canal_venta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_canal_venta_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_canal_venta) <> "" Then
      rs.Open "select * from tb_canalesventas where VCHA_CAN_CANAL_VENTA_ID = '" + txt_canal_venta + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_canal_venta = IIf(IsNull(rs!vcha_can_nombre), "", rs!vcha_can_nombre)
         rs.Close
      Else
         rs.Close
         MsgBox "Clave de Canal de Venta Incorrecta", vbOKOnly, "ATENCION"
         txt_canal_venta = ""
         txt_nombre_canal_venta = ""
      End If
   Else
      txt_nombre_canal_venta = ""
   End If
End Sub

Private Sub txt_descuento_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_descuento_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If Me.cmd_guardar.Enabled = True Then
         Me.cmd_guardar.SetFocus
      End If
   End If
End Sub

Private Sub txt_descuento_LostFocus()
   If txt_descuento <> "" Then
      If Not IsNumeric(txt_descuento) Then
         MsgBox "Porcentaje Incorrecto", vbOKOnly, "ATENCION"
         txt_descuento = ""
      End If
   End If
End Sub
Private Sub txt_importe_inferior_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_importe_inferior_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_importe_inferior_LostFocus()
   If txt_importe_inferior <> "" Then
      If Not IsNumeric(txt_importe_inferior) Then
         MsgBox "Importe Incorrecto", vbOKOnly, "ATENCION"
         txt_importe_inferior = ""
      End If
   End If
End Sub

Private Sub txt_importe_superior_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_importe_superior_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_importe_superior_LostFocus()
   If txt_importe_superior <> "" Then
      If Not IsNumeric(txt_importe_superior) Then
         MsgBox "Importe Incorrecto", vbOKOnly, "ATENCION"
         txt_importe_superior = ""
      End If
   End If
End Sub

Private Sub txt_nombre_canal_venta_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_canal_venta_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_nombre_canal_venta_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_canalesventas order by vcha_can_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_CAN_CANAL_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_can_nombre), "", rs!vcha_can_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CANALES DE VENTA"
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
   If KeyCode = 117 Then
      var_catalogo_articulos = True
      frmcanalesventas.Show
   End If
End Sub

Private Sub txt_nombre_canal_venta_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_canal_venta_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub
