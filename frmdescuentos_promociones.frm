VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmdescuentos_promociones 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descuentos por Promoción"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   8910
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   2040
      TabIndex        =   26
      Top             =   450
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
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
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7407
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
   Begin MSComCtl2.MonthView mes_inicio 
      Height          =   2370
      Left            =   3135
      TabIndex        =   22
      Top             =   1335
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   51642369
      CurrentDate     =   37581
   End
   Begin MSComCtl2.MonthView mes_fin 
      Height          =   2370
      Left            =   3150
      TabIndex        =   23
      Top             =   1335
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   51642369
      CurrentDate     =   37581
   End
   Begin VB.Frame Frame1 
      Caption         =   " Descuentos por Promoción"
      Height          =   1995
      Left            =   150
      TabIndex        =   16
      Top             =   495
      Width           =   8610
      Begin VB.TextBox txt_nombre_canal_venta 
         Height          =   315
         Left            =   3090
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   255
         Width           =   5370
      End
      Begin VB.TextBox txt_nombre_articulo 
         Height          =   315
         Left            =   3090
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   585
         Width           =   5370
      End
      Begin VB.CommandButton cmd_calendario_2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2655
         Picture         =   "frmdescuentos_promociones.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Ejecutar Calendario Alt + E"
         Top             =   1260
         Width           =   330
      End
      Begin VB.CommandButton cmd_calendario_1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2655
         Picture         =   "frmdescuentos_promociones.frx":1272
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Ejecutar Calendario"
         Top             =   915
         Width           =   330
      End
      Begin VB.TextBox txt_articulo 
         Height          =   315
         Left            =   1455
         TabIndex        =   8
         Top             =   585
         Width           =   1620
      End
      Begin VB.TextBox txt_porcentaje 
         Height          =   315
         Left            =   1455
         MaxLength       =   50
         TabIndex        =   12
         Top             =   1575
         Width           =   1200
      End
      Begin VB.TextBox txt_fecha_fin 
         Height          =   315
         Left            =   1455
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1245
         Width           =   1185
      End
      Begin VB.TextBox txt_fecha_inicio 
         Height          =   315
         Left            =   1455
         MaxLength       =   10
         TabIndex        =   10
         Top             =   915
         Width           =   1185
      End
      Begin VB.TextBox txt_canal_venta 
         Height          =   315
         Left            =   1455
         TabIndex        =   6
         Top             =   255
         Width           =   1620
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Artículo:"
         Height          =   195
         Index           =   3
         Left            =   270
         TabIndex        =   21
         Top             =   645
         Width           =   600
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Porcentaje:"
         Height          =   195
         Index           =   4
         Left            =   270
         TabIndex        =   20
         Top             =   1635
         Width           =   810
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Fin:"
         Height          =   195
         Index           =   2
         Left            =   270
         TabIndex        =   19
         Top             =   1305
         Width           =   750
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio:"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   18
         Top             =   975
         Width           =   915
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Canal de Venta:"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   17
         Top             =   315
         Width           =   1140
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4740
      Left            =   150
      TabIndex        =   14
      Top             =   2475
      Width           =   8625
      Begin MSComctlLib.ListView lv_descuentos_promociones 
         Height          =   4545
         Left            =   45
         TabIndex        =   15
         Top             =   150
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   8017
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   6526
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha Inicio"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fecha Fin"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "%"
            Object.Width           =   1305
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Canal de Venta"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Nombre Catálogo"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   210
      Picture         =   "frmdescuentos_promociones.frx":24E4
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   540
      Picture         =   "frmdescuentos_promociones.frx":25E6
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   855
      Picture         =   "frmdescuentos_promociones.frx":26E8
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1185
      Picture         =   "frmdescuentos_promociones.frx":27BA
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1515
      Picture         =   "frmdescuentos_promociones.frx":28BC
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8325
      Picture         =   "frmdescuentos_promociones.frx":29BE
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   13
      Top             =   255
      Width           =   8610
   End
End
Attribute VB_Name = "frmdescuentos_promociones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim numero_items_descuentos_promociones As Integer
Dim var_tipo_lista As Integer








Private Sub cmd_calendario_1_Click()
   If Trim(txt_fecha_fin) = "" Then
      mes_inicio = Date
   Else
      mes_inicio = txt_fecha_inicio
   End If
   mes_inicio.Visible = True
   mes_inicio.SetFocus
End Sub

Private Sub cmd_calendario_2_Click()
   If Trim(txt_fecha_fin) = "" Then
      mes_fin = Date
   Else
      mes_fin = txt_fecha_fin
   End If
   mes_fin.Visible = True
   mes_fin.SetFocus
End Sub

Private Sub cmd_deshacer_Click()
        Call pro_textos
End Sub

Private Sub cmd_eliminar_Click()
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
      Call pro_elimina_descuentos_promociones
      rs.Open "select * from tb_descuentos_promociones", cnn, adOpenDynamic, adLockOptimistic
      If rs.BOF Then
         cmd_guardar.Enabled = False
         cmd_deshacer.Enabled = False
         cmd_eliminar.Enabled = False
      Else
         cmd_guardar.Enabled = True
         cmd_deshacer.Enabled = True
         cmd_eliminar.Enabled = True
      End If
      rs.Close
   End If
End Sub

Private Sub cmd_guardar_Click()
   If Not IsNumeric(txt_porcentaje) Then
      txt_porcentaje = 0
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
      Call pro_guardar_descuentos_promociones
   End If
End Sub

Private Sub cmd_imprimir_Click()
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call gPrintListView(lv_descuentos_promociones, "LISTADO DE descuentos_promociones")
        End If
End Sub

Private Sub cmd_nuevo_Click()
   Call pro_limpiatextos(Me)
   txt_canal_venta.Enabled = True
   txt_canal_venta.SetFocus: var_modifica_registro = False
   cmd_guardar.Enabled = True
   cmd_deshacer.Enabled = True
   txt_canal_venta.Enabled = True
   txt_Articulo.Enabled = True
   txt_fecha_inicio.Enabled = True
   txt_fecha_fin.Enabled = True
   txt_porcentaje.Enabled = True
   txt_fecha_inicio = Date
   txt_fecha_fin = Date
   Me.txt_nombre_articulo.Enabled = True
   Me.txt_nombre_canal_venta.Enabled = False
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro = False Then
      var_si = MsgBox("No se han guardado los cambios, ¿Desea salir?", vbYesNo, "ATENCION")
      If var_si <> 6 Then
         GoTo salir:
      End If
   Else
      If var_hubo_cambios = True Then
         var_si = MsgBox("No se han guardado los cambios, ¿Desea salir?", vbYesNo, "ATENCION")
         If var_si <> 6 Then
            GoTo salir:
         End If
      End If
   End If
   Unload Me
   Exit Sub
salir:
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
   Left = 1500
   frm_lista.Visible = False
   numero_items_descuentos_promociones = 0
   mes_inicio.Visible = False
   mes_fin.Visible = False
   var_modifica_registro = True
   'Call pro_encabezadosView(Me, lv_descuentos_promociones, False)
   Call pro_llena_listview1
   pro_textos
   txt_canal_venta.Enabled = False
   Me.txt_nombre_articulo.Enabled = False
   Me.txt_nombre_canal_venta.Enabled = False
   txt_Articulo.Enabled = False
   txt_fecha_inicio.Enabled = False
   txt_fecha_fin.Enabled = False
   txt_porcentaje.Enabled = False
 End Sub

Private Sub Form_Unload(Cancel As Integer)
    var_swpassword = False
    var_modifica_registro = False
    Call activa_forma(var_activa_forma_descuentos_promociones)
End Sub

Private Sub lv_descuentos_promociones_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_descuentos_promociones, ColumnHeader)
End Sub

Private Sub lv_descuentos_promociones_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_descuentos_promociones.selectedItem = Item
        pro_textos
        var_modifica_registro = True
        txt_canal_venta.Enabled = True
End Sub



Sub pro_guardar_descuentos_promociones()
   Dim ok As Boolean
   Set TB_DESCUENTOS_PROMOCIONES = New TB_DESCUENTOS_PROMOCIONES
   If txt_canal_venta <> "" And txt_fecha_inicio <> "" And txt_fecha_fin <> "" And txt_porcentaje <> "" Then
      ok = TB_DESCUENTOS_PROMOCIONES.Anadir(txt_canal_venta, txt_Articulo, txt_fecha_inicio, txt_fecha_fin, txt_porcentaje)
      If ok Then
         pro_actualiza_ListView
         txt_canal_venta.Enabled = False
         MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
         txt_registros = lv_descuentos_promociones.ListItems.Count
         var_modifica_registro = True
      Else
         MsgBox "No se puede grabar registro: " + TB_DESCUENTOS_PROMOCIONES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
End Sub

Sub pro_elimina_descuentos_promociones()
   Dim var_llave_usuarios As String
   Set TB_DESCUENTOS_PROMOCIONES = New TB_DESCUENTOS_PROMOCIONES
   On Error GoTo salir
   ok = True
   If txt_canal_venta <> "" And txt_fecha_inicio <> "" And var_modifica_registro = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_DESCUENTOS_PROMOCIONES.Eliminar(txt_canal_venta, txt_Articulo, txt_fecha_inicio, txt_fecha_fin, txt_porcentaje)
      Else
         GoTo salir:
      End If
      If ok Then
         MsgBox "Se Elimino Correctamente el Registro", vbInformation
         lv_descuentos_promociones.ListItems.Remove (lv_descuentos_promociones.selectedItem.Index)
         numero_items_descuentos_promociones = numero_items_descuentos_promociones - 1
         Call pro_limpiatextos(Me)
         txt_registros = lv_descuentos_promociones.ListItems.Count
         lv_descuentos_promociones.selectedItem.Selected = True
         pro_textos
      Else
         MsgBox "No se puede grabar registro: " + TB_DESCUENTOS_PROMOCIONES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
   Set TB_DESCUENTOS_PROMOCIONES = Nothing
End Sub


Sub pro_llena_listview1()
   Dim list_item As ListItem
   rs.Open "select * from VW_descuentos_promociones where vcha_can_canal_venta_id = '" + txt_canal_venta + "'", cnn, adOpenDynamic, adLockOptimistic
   lv_descuentos_promociones.ListItems.Clear
   While Not rs.EOF
      Set list_item = lv_descuentos_promociones.ListItems.Add(, , rs!vcha_art_articulo_id)
      list_item.SubItems(1) = IIf(IsNull(rs!vcha_art_nombre_ESPAÑOL), "", rs!vcha_art_nombre_ESPAÑOL)
      list_item.SubItems(2) = IIf(IsNull(rs!DTIM_DPR_FECHA_INICIO), "", rs!DTIM_DPR_FECHA_INICIO)
      list_item.SubItems(3) = IIf(IsNull(rs!DTIM_DPR_FECHA_FIN), "", rs!DTIM_DPR_FECHA_FIN)
      list_item.SubItems(4) = IIf(IsNull(rs!FLOA_DPR_DESCUENTO), "", rs!FLOA_DPR_DESCUENTO)
      list_item.SubItems(5) = IIf(IsNull(rs!vcha_can_canal_venta_id), "", rs!vcha_can_canal_venta_id)
      list_item.SubItems(6) = IIf(IsNull(rs!vcha_can_nombre), "", rs!vcha_can_nombre)
      rs.MoveNext:
      numero_items_descuentos_promociones = numero_items_descuentos_promociones + 1
   Wend
   rs.Close
   If numero_items_descuentos_promociones > 11 Then
      lv_descuentos_promociones.ColumnHeaders(2).Width = 3800.25
   Else
      lv_descuentos_promociones.ColumnHeaders(2).Width = 4000.25
   End If
End Sub


Sub pro_textos()
'On Error GoTo err0:
   txt_Articulo = ""
   txt_fecha_inicio = ""
   txt_fecha_fin = ""
   txt_porcentaje = 0
   txt_canal_venta = ""
   txt_nombre_canal_venta = ""
   txt_nombre_articulo = ""
   var_n = lv_descuentos_promociones.ListItems.Count
   If var_n > 0 Then
      txt_Articulo = lv_descuentos_promociones.selectedItem
      txt_fecha_inicio = lv_descuentos_promociones.selectedItem.SubItems(2)
      txt_fecha_fin = lv_descuentos_promociones.selectedItem.SubItems(3)
      txt_porcentaje = lv_descuentos_promociones.selectedItem.SubItems(4)
      txt_canal_venta = lv_descuentos_promociones.selectedItem.SubItems(5)
   End If
   rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_Articulo + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      txt_nombre_articulo = IIf(IsNull(rs!vcha_art_nombre_ESPAÑOL), "", rs!vcha_art_nombre_ESPAÑOL)
   Else
      txt_nombre_articulo = ""
   End If
   rs.Close
   rs.Open "select * from tb_canalesventas where vcha_can_canal_venta_id = '" + txt_canal_venta + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      txt_nombre_canal_venta = IIf(IsNull(rs!vcha_can_nombre), "", rs!vcha_can_nombre)
   Else
      txt_nombre_canal_venta = ""
   End If
   rs.Close
   var_modifica_registro = True
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro = False Then
        Set list_item = lv_descuentos_promociones.ListItems.Add(, , txt_Articulo)
        list_item.SubItems(1) = Me.txt_nombre_articulo
        list_item.SubItems(2) = txt_fecha_inicio
        list_item.SubItems(3) = txt_fecha_fin
        list_item.SubItems(4) = txt_porcentaje
        list_item.SubItems(5) = txt_canal_venta
        list_item.SubItems(6) = Me.txt_nombre_canal_venta
        list_item.EnsureVisible
        list_item.Selected = True
       numero_items_descuentos_promociones = numero_items_descuentos_promociones + 1
    Else
        lv_descuentos_promociones.ListItems.Item(lv_descuentos_promociones.selectedItem.Index).Checked = False
        lv_descuentos_promociones.ListItems.Item(lv_descuentos_promociones.selectedItem.Index) = txt_Articulo
        lv_descuentos_promociones.ListItems.Item(lv_descuentos_promociones.selectedItem.Index).ListSubItems(1) = Me.txt_nombre_articulo
        lv_descuentos_promociones.ListItems.Item(lv_descuentos_promociones.selectedItem.Index).ListSubItems(2) = txt_fecha_fin
        lv_descuentos_promociones.ListItems.Item(lv_descuentos_promociones.selectedItem.Index).ListSubItems(3) = txt_fecha_fin
        lv_descuentos_promociones.ListItems.Item(lv_descuentos_promociones.selectedItem.Index).ListSubItems(4) = txt_porcentaje
        lv_descuentos_promociones.ListItems.Item(lv_descuentos_promociones.selectedItem.Index).ListSubItems(5) = txt_canal_venta
        lv_descuentos_promociones.ListItems.Item(lv_descuentos_promociones.selectedItem.Index).ListSubItems(6) = Me.txt_nombre_canal_venta
        lv_descuentos_promociones.ListItems.Item(lv_descuentos_promociones.selectedItem.Index).Selected = True
    End If
    If numero_items_descuentos_promociones > 11 Then
      lv_descuentos_promociones.ColumnHeaders(2).Width = 3800.25
   Else
      lv_descuentos_promociones.ColumnHeaders(2).Width = 4000.25
    End If
    lv_descuentos_promociones.SetFocus
End Sub








Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_tipo_lista = 1 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_canal_venta = lv_lista.selectedItem
            txt_nombre_canal_venta = lv_lista.selectedItem.SubItems(1)
         Else
            txt_canal_venta = ""
            txt_nombre_canal_venta = ""
         End If
         txt_canal_venta.SetFocus
      End If
      If var_tipo_lista = 2 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_Articulo = lv_lista.selectedItem
            txt_nombre_articulo = lv_lista.selectedItem.SubItems(1)
         Else
            txt_Articulo = ""
            txt_nombre_articulo = ""
         End If
         txt_Articulo.SetFocus
      End If
      frm_lista.Visible = False
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub mes_fin_DateDblClick(ByVal DateDblClicked As Date)
   txt_fecha_fin = mes_fin
   mes_fin.Visible = False
End Sub

Private Sub mes_fin_LostFocus()
   mes_fin.Visible = False
End Sub

Private Sub mes_inicio_DateDblClick(ByVal DateDblClicked As Date)
   txt_fecha_inicio = mes_inicio
   mes_inicio.Visible = False
End Sub

Private Sub mes_inicio_LostFocus()
   mes_inicio.Visible = False
End Sub

Private Sub txt_articulo_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_articulo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_articulos order by vcha_art_nombre_español", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_art_articulo_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_art_nombre_ESPAÑOL), "", rs!vcha_art_nombre_ESPAÑOL)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ARTICULOS"
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
   If KeyCode = 117 Then
      frmarticulos2.Show
   End If
End Sub

Private Sub txt_articulo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_articulo_LostFocus()
   Dim var_posible As Boolean
   If Trim(txt_Articulo) <> "" Then
      var_posible = False
      rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_Articulo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_posible = True
         txt_nombre_articulo = rs!vcha_art_nombre_ESPAÑOL
         rs.Close
      Else
         rs.Close
         rs.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + txt_Articulo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rs!vcha_art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               var_posible = True
               txt_Articulo = rs!vcha_art_articulo_id
               txt_nombre_articulo = rsaux!vcha_art_nombre_ESPAÑOL
               rsaux.Close
               rs.Close
            Else
               var_posible = False
               rsaux.Close
               rs.Close
            End If
         Else
            rs.Close
         End If
      End If
      If var_posible = True Then
         If var_origen_codigo = 0 Then
            txt_cantidad = Format(0, "###0.00")
         Else
            var_origen_codigo = 0
            txt_cantidad = Format(0, "###0.00")
         End If
      Else
         MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
         txt_Articulo.SetFocus
      End If
   End If
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
      frmcanalesventas.Show
   End If
End Sub

Private Sub txt_canal_venta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_canal_venta_LostFocus()
   If Trim(txt_canal_venta) <> "" Then
      rsaux.Open "select * from TB_CANALESVENTAS where vcha_can_canal_venta_id = '" + txt_canal_venta + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux.EOF Then
         txt_nombre_canal_venta = IIf(IsNull(rsaux!vcha_can_nombre), "", rsaux!vcha_can_nombre)
         Call pro_llena_listview1
      Else
         txt_canal_venta = ""
         txt_nombre_canal_venta = ""
      End If
      rsaux.Close
   Else
      txt_nombre_canal_venta = ""
   End If
End Sub

Private Sub txt_fecha_fin_LostFocus()
   If Not IsDate(txt_fecha_fin) Then
      MsgBox "Fecha Invalida", vbOKOnly, "ATENCION"
      txt_fecha_fin = ""
   End If
End Sub

Private Sub txt_fecha_inicio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_fecha_fin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_fecha_inicio_LostFocus()
   If Not IsDate(txt_fecha_inicio) Then
      MsgBox "Fecha Invalida", vbOKOnly, "ATENCION"
      txt_fecha_inicio = ""
   End If
End Sub

Private Sub txt_nombre_articulo_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_nombre_articulo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_articulos order by vcha_art_nombre_español", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_art_articulo_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_art_nombre_ESPAÑOL), "", rs!vcha_art_nombre_ESPAÑOL)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ARTICULOS"
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
   If KeyCode = 117 Then
      frmarticulos2.Show
   End If
End Sub

Private Sub txt_nombre_articulo_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
       KeyAscii = 0
    Else
       Call pro_enfoque(KeyAscii)
    End If
End Sub

Private Sub txt_nombre_canal_venta_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_nombre_canal_venta_KeyDown(KeyCode As Integer, Shift As Integer)
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
      frmcanalesventas.Show
   End If
End Sub

Private Sub txt_nombre_canal_venta_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
       KeyAscii = 0
    Else
       Call pro_enfoque(KeyAscii)
    End If
End Sub

Private Sub txt_porcentaje_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
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

Private Sub txt_porcentaje_LostFocus()
   If Not IsNumeric(txt_porcentaje) Then
      MsgBox "Porcentaje incorrecto", vbOKOnly, "ATENCION"
      txt_porcentaje = 0
   End If
End Sub
