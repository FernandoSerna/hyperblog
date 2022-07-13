VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmlistamovimientos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "frmlistamovimientos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   6735
   Begin MSComctlLib.ListView lv_movimientos 
      Height          =   3345
      Left            =   -15
      TabIndex        =   0
      Top             =   75
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   5900
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Clave"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "                               Nombre del Movimiento"
         Object.Width           =   11642
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Afectacion"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "hace referencia"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "origen"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Factura"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "folio"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "tipo documento"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Causa Devolución"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "intercompañia"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "relectura"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "tipo proveedor"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Reporte"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "reempaque"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmlistamovimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_habilita_forma As Boolean
Private Sub Form_Load()
   Dim contador As Integer
   Dim var_ref_mov As Integer
   Dim var_ref_afe As String
   Dim var_ref_emp As String
   Dim var_ref_inte As Integer
   var_cadena_seguridad = ""
   Top = 2000
   Left = 2500
   contador = 0
   var_habilita_forma = True
   If var_nec_emb = False Then
      If var_tipo_permiso = 1 Then
         rs.Open "select * from vw_movimientos_permisos where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' order by vcha_mov_nombre", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
            var_ref_mov = IIf(IsNull(rs!INTE_MOV_REFEREANCIA), 0, rs!INTE_MOV_REFEREANCIA)
            var_ref_afe = IIf(IsNull(rs!CHAR_MOV_AFECTACION), "", rs!CHAR_MOV_AFECTACION)
            var_ref_emp = rs!VCHA_MOV_MOVIMIENTO_ID
            var_ref_inte = IIf(IsNull(rs!inte_mov_intercompañia), 0, rs!inte_mov_intercompañia)
            If (var_ref_mov = 1 And (Trim(var_ref_afe) = "-" Or Trim(var_ref_afe) = "T") And var_ref_inte = 0) Then
            Else
               If var_ref_emp <> "EM" And var_ref_emp <> "CJ" Then
                  Set list_item = lv_movimientos.ListItems.Add(, , rs!VCHA_MOV_MOVIMIENTO_ID)
                  list_item.SubItems(1) = IIf(IsNull(rs!vcha_mov_nombre), "", rs!vcha_mov_nombre)
                  list_item.SubItems(2) = IIf(IsNull(rs!CHAR_MOV_AFECTACION), "", rs!CHAR_MOV_AFECTACION)
                  list_item.SubItems(3) = IIf(IsNull(rs!INTE_MOV_REFEREANCIA), 0, rs!INTE_MOV_REFEREANCIA)
                  list_item.SubItems(4) = IIf(IsNull(rs!vcha_mov_titulo_origen), "", rs!vcha_mov_titulo_origen)
                  list_item.SubItems(5) = IIf(IsNull(rs!INTE_MOV_FACTURA), 0, rs!INTE_MOV_FACTURA)
                  list_item.SubItems(6) = IIf(IsNull(rs!INTE_MOV_FOLIO), 0, rs!INTE_MOV_FOLIO)
                  list_item.SubItems(7) = IIf(IsNull(rs!char_mov_documento), "", rs!char_mov_documento)
                  list_item.SubItems(8) = IIf(IsNull(rs!INTE_MOV_CAUSA_DEVOLUCION), 0, rs!INTE_MOV_CAUSA_DEVOLUCION)
                  list_item.SubItems(9) = IIf(IsNull(rs!inte_mov_intercompañia), 0, rs!inte_mov_intercompañia)
                  list_item.SubItems(10) = IIf(IsNull(rs!INTE_MOV_RELECTURA), 0, rs!INTE_MOV_RELECTURA)
                  list_item.SubItems(11) = IIf(IsNull(rs!char_mov_tipo_proveedor), "", rs!char_mov_tipo_proveedor)
                  list_item.SubItems(12) = IIf(IsNull(rs!vcha_mov_reporte_imprimir), "", rs!vcha_mov_reporte_imprimir)
                  list_item.SubItems(13) = IIf(IsNull(rs!INTE_MOV_REEMPAQUE), 0, rs!INTE_MOV_REEMPAQUE)
                  contador = contador + 1
               End If
            End If
            rs.MoveNext:
         Wend
         rs.Close
      Else
         rs.Open "select * from TB_movimientos order by vcha_mov_nombre", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
            var_ref_mov = IIf(IsNull(rs!INTE_MOV_REFEREANCIA), 0, rs!INTE_MOV_REFEREANCIA)
            var_ref_afe = IIf(IsNull(rs!CHAR_MOV_AFECTACION), "", rs!CHAR_MOV_AFECTACION)
            var_ref_emp = rs!VCHA_MOV_MOVIMIENTO_ID
            var_ref_inte = IIf(IsNull(rs!inte_mov_intercompañia), 0, rs!inte_mov_intercompañia)
            If (var_ref_mov = 1 And (Trim(var_ref_afe) = "-" Or Trim(var_ref_afe) = "T") And var_ref_inte = 0) Then
            Else
               If var_ref_emp <> "EM" And var_ref_emp <> "CJ" Then
                  Set list_item = lv_movimientos.ListItems.Add(, , rs!VCHA_MOV_MOVIMIENTO_ID)
                  list_item.SubItems(1) = IIf(IsNull(rs!vcha_mov_nombre), "", rs!vcha_mov_nombre)
                  list_item.SubItems(2) = IIf(IsNull(rs!CHAR_MOV_AFECTACION), "", rs!CHAR_MOV_AFECTACION)
                  list_item.SubItems(3) = IIf(IsNull(rs!INTE_MOV_REFEREANCIA), 0, rs!INTE_MOV_REFEREANCIA)
                  list_item.SubItems(4) = IIf(IsNull(rs!vcha_mov_titulo_origen), "", rs!vcha_mov_titulo_origen)
                  list_item.SubItems(5) = IIf(IsNull(rs!INTE_MOV_FACTURA), 0, rs!INTE_MOV_FACTURA)
                  list_item.SubItems(6) = IIf(IsNull(rs!INTE_MOV_FOLIO), 0, rs!INTE_MOV_FOLIO)
                  list_item.SubItems(7) = IIf(IsNull(rs!char_mov_documento), "", rs!char_mov_documento)
                  list_item.SubItems(8) = IIf(IsNull(rs!INTE_MOV_CAUSA_DEVOLUCION), 0, rs!INTE_MOV_CAUSA_DEVOLUCION)
                  list_item.SubItems(9) = IIf(IsNull(rs!inte_mov_intercompañia), 0, rs!inte_mov_intercompañia)
                  list_item.SubItems(10) = IIf(IsNull(rs!INTE_MOV_RELECTURA), 0, rs!INTE_MOV_RELECTURA)
                  list_item.SubItems(11) = IIf(IsNull(rs!char_mov_tipo_proveedor), "", rs!char_mov_tipo_proveedor)
                  list_item.SubItems(12) = IIf(IsNull(rs!vcha_mov_reporte_imprimir), "", rs!vcha_mov_reporte_imprimir)
                  list_item.SubItems(13) = IIf(IsNull(rs!INTE_MOV_REEMPAQUE), 0, rs!INTE_MOV_REEMPAQUE)
                  contador = contador + 1
               End If
            End If
            rs.MoveNext:
         Wend
         rs.Close
      End If
   Else
      If var_tipo_permiso = 1 Then
         rs.Open "select * from vw_movimientos_permisos where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' order by vcha_mov_nombre", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
            var_ref_mov = IIf(IsNull(rs!INTE_MOV_REFEREANCIA), 0, rs!INTE_MOV_REFEREANCIA)
            var_ref_afe = IIf(IsNull(rs!CHAR_MOV_AFECTACION), "", rs!CHAR_MOV_AFECTACION)
            var_ref_emp = rs!VCHA_MOV_MOVIMIENTO_ID
            var_ref_inte = IIf(IsNull(rs!inte_mov_intercompañia), 0, rs!inte_mov_intercompañia)
            If (var_ref_mov = 1 And (Trim(var_ref_afe) = "-" Or Trim(var_ref_afe) = "T") And var_ref_inte = 0) Then
               If var_ref_emp <> "EM" And var_ref_emp <> "CJ" Then
                  Set list_item = lv_movimientos.ListItems.Add(, , rs!VCHA_MOV_MOVIMIENTO_ID)
                  list_item.SubItems(1) = IIf(IsNull(rs!vcha_mov_nombre), "", rs!vcha_mov_nombre)
                  list_item.SubItems(2) = IIf(IsNull(rs!CHAR_MOV_AFECTACION), "", rs!CHAR_MOV_AFECTACION)
                  list_item.SubItems(3) = IIf(IsNull(rs!INTE_MOV_REFEREANCIA), 0, rs!INTE_MOV_REFEREANCIA)
                  list_item.SubItems(4) = IIf(IsNull(rs!vcha_mov_titulo_origen), "", rs!vcha_mov_titulo_origen)
                  list_item.SubItems(5) = IIf(IsNull(rs!INTE_MOV_FACTURA), 0, rs!INTE_MOV_FACTURA)
                  list_item.SubItems(6) = IIf(IsNull(rs!INTE_MOV_FOLIO), 0, rs!INTE_MOV_FOLIO)
                  list_item.SubItems(7) = IIf(IsNull(rs!char_mov_documento), "", rs!char_mov_documento)
                  list_item.SubItems(8) = IIf(IsNull(rs!INTE_MOV_CAUSA_DEVOLUCION), 0, rs!INTE_MOV_CAUSA_DEVOLUCION)
                  list_item.SubItems(9) = IIf(IsNull(rs!inte_mov_intercompañia), 0, rs!inte_mov_intercompañia)
                  list_item.SubItems(10) = IIf(IsNull(rs!INTE_MOV_RELECTURA), 0, rs!INTE_MOV_RELECTURA)
                  list_item.SubItems(11) = IIf(IsNull(rs!char_mov_tipo_proveedor), "", rs!char_mov_tipo_proveedor)
                  list_item.SubItems(12) = IIf(IsNull(rs!vcha_mov_reporte_imprimir), "", rs!vcha_mov_reporte_imprimir)
                  list_item.SubItems(13) = IIf(IsNull(rs!INTE_MOV_REEMPAQUE), 0, rs!INTE_MOV_REEMPAQUE)
                  contador = contador + 1
               End If
            End If
            rs.MoveNext:
         Wend
         rs.Close
      Else
         rs.Open "select * from TB_movimientos order by vcha_mov_nombre", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
            var_ref_mov = IIf(IsNull(rs!INTE_MOV_REFEREANCIA), 0, rs!INTE_MOV_REFEREANCIA)
            var_ref_afe = IIf(IsNull(rs!CHAR_MOV_AFECTACION), "", rs!CHAR_MOV_AFECTACION)
            var_ref_emp = rs!VCHA_MOV_MOVIMIENTO_ID
            var_ref_inte = IIf(IsNull(rs!inte_mov_intercompañia), 0, rs!inte_mov_intercompañia)
            If (var_ref_mov = 1 And (Trim(var_ref_afe) = "-" Or Trim(var_ref_afe) = "T") And var_ref_inte = 0) Then
               If var_ref_emp <> "EM" And var_ref_emp <> "CJ" Then
                  Set list_item = lv_movimientos.ListItems.Add(, , rs!VCHA_MOV_MOVIMIENTO_ID)
                  list_item.SubItems(1) = IIf(IsNull(rs!vcha_mov_nombre), "", rs!vcha_mov_nombre)
                  list_item.SubItems(2) = IIf(IsNull(rs!CHAR_MOV_AFECTACION), "", rs!CHAR_MOV_AFECTACION)
                  list_item.SubItems(3) = IIf(IsNull(rs!INTE_MOV_REFEREANCIA), 0, rs!INTE_MOV_REFEREANCIA)
                  list_item.SubItems(4) = IIf(IsNull(rs!vcha_mov_titulo_origen), "", rs!vcha_mov_titulo_origen)
                  list_item.SubItems(5) = IIf(IsNull(rs!INTE_MOV_FACTURA), 0, rs!INTE_MOV_FACTURA)
                  list_item.SubItems(6) = IIf(IsNull(rs!INTE_MOV_FOLIO), 0, rs!INTE_MOV_FOLIO)
                  list_item.SubItems(7) = IIf(IsNull(rs!char_mov_documento), "", rs!char_mov_documento)
                  list_item.SubItems(8) = IIf(IsNull(rs!INTE_MOV_CAUSA_DEVOLUCION), 0, rs!INTE_MOV_CAUSA_DEVOLUCION)
                  list_item.SubItems(9) = IIf(IsNull(rs!inte_mov_intercompañia), 0, rs!inte_mov_intercompañia)
                  list_item.SubItems(10) = IIf(IsNull(rs!INTE_MOV_RELECTURA), 0, rs!INTE_MOV_RELECTURA)
                  list_item.SubItems(11) = IIf(IsNull(rs!char_mov_tipo_proveedor), "", rs!char_mov_tipo_proveedor)
                  list_item.SubItems(12) = IIf(IsNull(rs!vcha_mov_reporte_imprimir), "", rs!vcha_mov_reporte_imprimir)
                  list_item.SubItems(13) = IIf(IsNull(rs!INTE_MOV_REEMPAQUE), 0, rs!INTE_MOV_REEMPAQUE)
                  contador = contador + 1
               End If
            End If
            rs.MoveNext:
         Wend
         rs.Close
      End If
   End If
   If contador > 9 Then
      lv_movimientos.ColumnHeaders(2).Width = 6380
   Else
   End If
End Sub

Private Sub lst_listamovimientos_DblClick()
End Sub

Private Sub lst_listamovimientos_KeyPress(KeyAscii As Integer)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If var_activa_menu = True And var_habilita_forma = True And var_es_embarque = False Then
       Frmmenu2.Enabled = True
    End If
End Sub

Private Sub lv_movimientos_DblClick()
Dim var_movimiento As String
Dim var_ajuste As Integer
Dim var_reempaque As Integer
'On Error GoTo SALIR:
   var_subir_directo = False
   var_reempaque = lv_movimientos.selectedItem.SubItems(13) * 1
   var_movimiento = Trim(lv_movimientos.selectedItem.SubItems(2))
   var_tipo_documento = Trim(lv_movimientos.selectedItem.SubItems(7))
   var_tipo_proveedor_movimiento = Trim(lv_movimientos.selectedItem.SubItems(11))
   var_reporte_imprimir = lv_movimientos.selectedItem.SubItems(12)
   If lv_movimientos.selectedItem.SubItems(8) = 1 Then
      var_causa_devolucion = True
   Else
      var_causa_devolucion = False
   End If
   If var_movimiento = "T" Then
      var_clave_movimiento = lv_movimientos.selectedItem
      If var_tipo_documento = "D" Then
         var_clave_movimiento = lv_movimientos.selectedItem
         frmtraspasos_calidad.Caption = lv_movimientos.selectedItem.SubItems(1)
         frmtraspasos_calidad.lblnombremovimiento.Caption = lv_movimientos.selectedItem.SubItems(1)
         If Len(lv_movimientos.selectedItem.SubItems(1)) > 27 Then
            frmtraspasos_calidad.lblnombremovimiento.Font.Size = 16
         Else
            frmtraspasos_calidad.lblnombremovimiento.Font.Size = 20
         End If
         var_habilita_forma = False
         var_activa_forma_traspasos_calidad = "MENU"
         frmtraspasos_calidad.Show
         Unload Me
      Else
         If var_clave_movimiento = "SV" Then
            var_z = 1
            If var_z = 1 Then
               var_habilita_forma = False
               var_activa_forma_salidas_proveedor = "MENU"
               frmsalidas_vistas.Caption = lv_movimientos.selectedItem.SubItems(1)
               frmsalidas_vistas.lblnombremovimiento = lv_movimientos.selectedItem.SubItems(1)
               frmsalidas_vistas.Show
               Unload Me
            Else
               frmsalidas.txt_clave_movimiento = var_clave_movimiento
               frmsalidas.lblnombremovimiento.Caption = lv_movimientos.selectedItem.SubItems(1)
               frmsalidas.Caption = lv_movimientos.selectedItem.SubItems(1)
               If Len(lv_movimientos.selectedItem.SubItems(1)) > 30 Then
                  frmsalidas.lblnombremovimiento.Font.Size = 18
               Else
                  frmsalidas.lblnombremovimiento.Font.Size = 24
               End If
               var_habilita_forma = False
               var_activa_forma_salidas = "MENU"
               frmsalidas.Show
               Unload Me
            End If
            
         Else
            If var_tipo_documento = "V" Then
               var_clave_movimiento = lv_movimientos.selectedItem
               frmentradas_devolucion_vistas.Caption = lv_movimientos.selectedItem.SubItems(1)
               frmentradas_devolucion_vistas.lblnombremovimiento.Caption = lv_movimientos.selectedItem.SubItems(1)
               If Len(lv_movimientos.selectedItem.SubItems(1)) > 27 Then
                  frmentradas_devolucion_vistas.lblnombremovimiento.Font.Size = 16
               Else
                  frmentradas_devolucion_vistas.lblnombremovimiento.Font.Size = 20
               End If
               var_habilita_forma = False
               var_activa_forma_salidas_proveedor = "MENU"
               frmentradas_devolucion_vistas.Show
               Unload Me
            Else
               If var_reempaque = 0 Then
                  If Trim(lv_movimientos.selectedItem) = "TEC" Then
                     var_clave_movimiento = lv_movimientos.selectedItem
                     frmtraspasos_entradas_calidad.Caption = lv_movimientos.selectedItem.SubItems(1)
                     frmtraspasos_entradas_calidad.lblnombremovimiento.Caption = lv_movimientos.selectedItem.SubItems(1)
                     If Len(lv_movimientos.selectedItem.SubItems(1)) > 27 Then
                        frmtraspasos_entradas_calidad.lblnombremovimiento.Font.Size = 16
                     Else
                        frmtraspasos_entradas_calidad.lblnombremovimiento.Font.Size = 20
                     End If
                     var_habilita_forma = False
                     var_activa_forma_traspasos = "MENU"
                     frmtraspasos_entradas_calidad.Show
                     Unload Me
                  Else
                     If var_clave_movimiento = "EST" Then
                        var_clave_movimiento = lv_movimientos.selectedItem
                        frmentradas_salidas_transformacion.Caption = lv_movimientos.selectedItem.SubItems(1)
                        frmentradas_salidas_transformacion.lblnombremovimiento.Caption = lv_movimientos.selectedItem.SubItems(1)
                        If Len(lv_movimientos.selectedItem.SubItems(1)) > 27 Then
                           frmentradas_salidas_transformacion.lblnombremovimiento.Font.Size = 16
                        Else
                           frmentradas_salidas_transformacion.lblnombremovimiento.Font.Size = 20
                        End If
                        var_habilita_forma = False
                        var_activa_forma_traspasos = "MENU"
                        frmentradas_salidas_transformacion.Show
                        Unload Me
                     Else
                        var_clave_movimiento = lv_movimientos.selectedItem
                        frmtraspasos.Caption = lv_movimientos.selectedItem.SubItems(1)
                        frmtraspasos.lblnombremovimiento.Caption = lv_movimientos.selectedItem.SubItems(1)
                        If Len(lv_movimientos.selectedItem.SubItems(1)) > 27 Then
                           frmtraspasos.lblnombremovimiento.Font.Size = 16
                        Else
                           frmtraspasos.lblnombremovimiento.Font.Size = 20
                        End If
                        var_habilita_forma = False
                        var_activa_forma_traspasos = "MENU"
                        frmtraspasos.Show
                        Unload Me
                     End If
                  End If
               Else
                  If var_reempaque = 1 Then
                     var_clave_movimiento = lv_movimientos.selectedItem
                     frmsalidas_reempaque.Caption = lv_movimientos.selectedItem.SubItems(1)
                     frmsalidas_reempaque.lblnombremovimiento.Caption = lv_movimientos.selectedItem.SubItems(1)
                     If Len(lv_movimientos.selectedItem.SubItems(1)) > 27 Then
                        frmsalidas_reempaque.lblnombremovimiento.Font.Size = 16
                     Else
                        frmsalidas_reempaque.lblnombremovimiento.Font.Size = 20
                     End If
                     var_habilita_forma = False
                     var_activa_forma_salidas_reempaque = "MENU"
                     frmsalidas_reempaque.Show
                     Unload Me
                  End If
                  If var_reempaque = 2 Then
                     var_clave_movimiento = lv_movimientos.selectedItem
                     frmentradas_reempaque.Caption = lv_movimientos.selectedItem.SubItems(1)
                     frmentradas_reempaque.lblnombremovimiento.Caption = lv_movimientos.selectedItem.SubItems(1)
                     If Len(lv_movimientos.selectedItem.SubItems(1)) > 27 Then
                        frmentradas_reempaque.lblnombremovimiento.Font.Size = 16
                     Else
                        frmentradas_reempaque.lblnombremovimiento.Font.Size = 20
                     End If
                     var_habilita_forma = False
                     var_activa_forma_entradas_reempaque = "MENU"
                     frmentradas_reempaque.Show
                     Unload Me
                  End If
               End If
            End If
         End If
      End If
   End If
   If var_movimiento = "TE" Then
      var_clave_movimiento = lv_movimientos.selectedItem
      If var_clave_movimiento = "TT" Then
         frmtraspasosentradas_tiendas.Caption = lv_movimientos.selectedItem.SubItems(1)
         frmtraspasosentradas_tiendas.lblnombremovimiento.Caption = lv_movimientos.selectedItem.SubItems(1)
         If Len(lv_movimientos.selectedItem.SubItems(1)) > 20 Then
            frmtraspasosentradas_tiendas.lblnombremovimiento.Font.Size = 12
         Else
            frmtraspasosentradas_tiendas.lblnombremovimiento.Font.Size = 20
         End If
         var_habilita_forma = False
         var_activa_forma_traspasosentradas = "MENU"
         frmtraspasosentradas_tiendas.Show
         Unload Me
      Else
         frmtraspasosentradas.Caption = lv_movimientos.selectedItem.SubItems(1)
         frmtraspasosentradas.lblnombremovimiento.Caption = lv_movimientos.selectedItem.SubItems(1)
         If Len(lv_movimientos.selectedItem.SubItems(1)) > 20 Then
            frmtraspasosentradas.lblnombremovimiento.Font.Size = 12
         Else
            frmtraspasosentradas.lblnombremovimiento.Font.Size = 20
         End If
         var_habilita_forma = False
         var_activa_forma_traspasosentradas = "MENU"
         frmtraspasosentradas.Show
         Unload Me
      End If
   End If
   If var_movimiento = "TS" Then
      var_clave_movimiento = lv_movimientos.selectedItem
      frmtraspasossalidas.Caption = lv_movimientos.selectedItem.SubItems(1)
      frmtraspasossalidas.lblnombremovimiento.Caption = lv_movimientos.selectedItem.SubItems(1)
      If Len(lv_movimientos.selectedItem.SubItems(1)) > 27 Then
         frmtraspasossalidas.lblnombremovimiento.Font.Size = 16
      Else
         frmtraspasossalidas.lblnombremovimiento.Font.Size = 20
      End If
      var_habilita_forma = False
      var_activa_forma_traspasossalidas = "MENU"
      frmtraspasossalidas.Show
      Unload Me
   End If
   If var_movimiento = "+" Then
      If var_reempaque = 0 Then
         If Trim(lv_movimientos.selectedItem.SubItems(3)) = 1 Then
            var_clave_movimiento = lv_movimientos.selectedItem
            frmentradas.txt_clave_movimiento = var_clave_movimiento
            frmentradas.txt_tipo_documento = lv_movimientos.selectedItem.SubItems(7)
            frmentradas.lblnombremovimiento.Caption = lv_movimientos.selectedItem.SubItems(1)
            frmentradas.Caption = lv_movimientos.selectedItem.SubItems(1)
            frmentradas.lbl_origen = lv_movimientos.selectedItem.SubItems(4)
            var_global_relectura = lv_movimientos.selectedItem.SubItems(10)
            If lv_movimientos.selectedItem.SubItems(5) = 0 Then
               frmentradas.txt_factura.Enabled = False
            Else
               frmentradas.txt_factura.Enabled = True
            End If
            If Len(lv_movimientos.selectedItem.SubItems(1)) > 30 Then
               frmentradas.lblnombremovimiento.Font.Size = 18
            Else
               frmentradas.lblnombremovimiento.Font.Size = 24
            End If
            var_habilita_forma = False
            var_activa_forma_entradas = "MENU"
            frmentradas.Show
            Unload Me
         Else
            If var_tipo_documento = "D" Then
               If var_empresa = "06" Or var_empresa = "16" Or var_empresa = "15" Or var_empresa = "31" Or var_empresa = "18" Or var_empresa = "17" Then
                  var_cambio = 0
                  If var_cambio = 1 Then
                  Else
                     var_clave_movimiento = lv_movimientos.selectedItem
                     If var_empresa = "16" Or var_empresa = "06" Or var_empresa = "18" Or var_empresa = "15" Or var_empresa = "31" Or var_empresa = "17" Or var_empresa = "28" Then
                        If var_clave_movimiento = "CAPT" Then
                           frmentradas_devoluciones.Caption = lv_movimientos.selectedItem.SubItems(1)
                           frmentradas_devoluciones.lblnombremovimiento.Caption = lv_movimientos.selectedItem.SubItems(1)
                           If Len(lv_movimientos.selectedItem.SubItems(1)) > 27 Then
                              frmentradas_devoluciones.lblnombremovimiento.Font.Size = 16
                           Else
                             frmentradas_devoluciones.lblnombremovimiento.Font.Size = 20
                           End If
                           var_habilita_forma = False
                           var_activa_forma_entradas_devoluciones = "MENU"
                           frmentradas_devoluciones.Show
                        Else
                           frmentradas_devolucion_completa.Caption = lv_movimientos.selectedItem.SubItems(1)
                           frmentradas_devolucion_completa.lblnombremovimiento.Caption = lv_movimientos.selectedItem.SubItems(1)
                           If Len(lv_movimientos.selectedItem.SubItems(1)) > 27 Then
                              frmentradas_devolucion_completa.lblnombremovimiento.Font.Size = 16
                           Else
                             frmentradas_devolucion_completa.lblnombremovimiento.Font.Size = 20
                           End If
                           var_habilita_forma = False
                           var_activa_forma_entradas_devoluciones = "MENU"
                           frmentradas_devolucion_completa.Show
                        End If
                     Else
                        frmentradas_devoluciones.Caption = lv_movimientos.selectedItem.SubItems(1)
                        frmentradas_devoluciones.lblnombremovimiento.Caption = lv_movimientos.selectedItem.SubItems(1)
                        If Len(lv_movimientos.selectedItem.SubItems(1)) > 27 Then
                           frmentradas_devoluciones.lblnombremovimiento.Font.Size = 16
                        Else
                          frmentradas_devoluciones.lblnombremovimiento.Font.Size = 20
                        End If
                        var_habilita_forma = False
                        var_activa_forma_entradas_devoluciones = "MENU"
                        frmentradas_devoluciones.Show
                     End If
                     Unload Me
                  End If
               Else
                  var_clave_movimiento = lv_movimientos.selectedItem
                  frmentradas_devoluciones.Caption = lv_movimientos.selectedItem.SubItems(1)
                  frmentradas_devoluciones.lblnombremovimiento.Caption = lv_movimientos.selectedItem.SubItems(1)
                  If Len(lv_movimientos.selectedItem.SubItems(1)) > 27 Then
                     frmentradas_devoluciones.lblnombremovimiento.Font.Size = 16
                  Else
                     frmentradas_devoluciones.lblnombremovimiento.Font.Size = 20
                  End If
                  var_habilita_forma = False
                  var_activa_forma_entradas_devoluciones = "MENU"
                  frmentradas_devoluciones.Show
                  Unload Me
               End If
            Else
               var_clave_movimiento = lv_movimientos.selectedItem
               If lv_movimientos.selectedItem.SubItems(5) = 1 Then
                  frmentradas_compras.Caption = lv_movimientos.selectedItem.SubItems(1)
                  frmentradas_compras.lblnombremovimiento.Caption = lv_movimientos.selectedItem.SubItems(1)
                  If Len(lv_movimientos.selectedItem.SubItems(1)) > 27 Then
                     frmentradas_compras.lblnombremovimiento.Font.Size = 12
                  Else
                     frmentradas_compras.lblnombremovimiento.Font.Size = 14
                  End If
                  var_habilita_forma = False
                  var_activa_forma_entradas_compras = "MENU"
                  frmentradas_compras.Show
                  Unload Me
               Else
                  If var_clave_movimiento = "DVI" Then
                     var_habilita_forma = False
                     var_activa_forma_salidas_proveedor = "MENU"
                     frmentradas_vistas.lblnombremovimiento = lv_movimientos.selectedItem.SubItems(1)
                     frmentradas_vistas.Show
                     Unload Me
                  Else
                     If var_clave_movimiento = "ETP" Then
                        frmentradas_compras_plantas.Caption = lv_movimientos.selectedItem.SubItems(1)
                        frmentradas_compras_plantas.lblnombremovimiento.Caption = lv_movimientos.selectedItem.SubItems(1)
                        If Len(lv_movimientos.selectedItem.SubItems(1)) > 27 Then
                           frmentradas_compras_plantas.lblnombremovimiento.Font.Size = 16
                        Else
                           frmentradas_compras_plantas.lblnombremovimiento.Font.Size = 20
                        End If
                        var_habilita_forma = False
                        var_activa_forma_entradas_sin_comparacion = "MENU"
                        frmentradas_compras_plantas.Show
                        Unload Me
                     Else
                        If var_clave_movimiento = "ETMP" Then
                           frmentradas_compras_plantas.Caption = lv_movimientos.selectedItem.SubItems(1)
                           frmentradas_compras_plantas.lblnombremovimiento.Caption = lv_movimientos.selectedItem.SubItems(1)
                           If Len(lv_movimientos.selectedItem.SubItems(1)) > 27 Then
                              frmentradas_compras_plantas.lblnombremovimiento.Font.Size = 16
                           Else
                              frmentradas_compras_plantas.lblnombremovimiento.Font.Size = 20
                           End If
                           var_habilita_forma = False
                           var_activa_forma_entradas_sin_comparacion = "MENU"
                           frmentradas_compras_plantas.Show
                           Unload Me
                        Else
                           If var_clave_movimiento = "DS" Then
                              frmentradas_devolucion_vistas.Caption = lv_movimientos.selectedItem.SubItems(1)
                              frmentradas_devolucion_vistas.lblnombremovimiento.Caption = lv_movimientos.selectedItem.SubItems(1)
                              If Len(lv_movimientos.selectedItem.SubItems(1)) > 27 Then
                                 frmentradas_devolucion_vistas.lblnombremovimiento.Font.Size = 16
                              Else
                                 frmentradas_devolucion_vistas.lblnombremovimiento.Font.Size = 20
                              End If
                              var_habilita_forma = False
                               var_activa_forma_salidas_proveedor = "MENU"
                              frmentradas_devolucion_vistas.Show
                              Unload Me
                           Else
                              If var_clave_movimiento = "DTA" Then
                                 frmentradas_devolucions_tiendas.Caption = lv_movimientos.selectedItem.SubItems(1)
                                 frmentradas_devolucions_tiendas.lblnombremovimiento.Caption = lv_movimientos.selectedItem.SubItems(1)
                                 If Len(lv_movimientos.selectedItem.SubItems(1)) > 27 Then
                                    frmentradas_devolucions_tiendas.lblnombremovimiento.Font.Size = 16
                                 Else
                                    frmentradas_devolucions_tiendas.lblnombremovimiento.Font.Size = 20
                                 End If
                                 var_habilita_forma = False
                                 var_activa_forma_entradas_sin_comparacion = "MENU"
                                 frmentradas_devolucions_tiendas.Show
                                 Unload Me
                              Else
                                 If var_clave_movimiento = "EVD" Then
                                    frmentradas_compras_plantas.Caption = lv_movimientos.selectedItem.SubItems(1)
                                    frmentradas_compras_plantas.lblnombremovimiento.Caption = lv_movimientos.selectedItem.SubItems(1)
                                    If Len(lv_movimientos.selectedItem.SubItems(1)) > 27 Then
                                       frmentradas_compras_plantas.lblnombremovimiento.Font.Size = 16
                                    Else
                                       frmentradas_compras_plantas.lblnombremovimiento.Font.Size = 20
                                    End If
                                    var_habilita_forma = False
                                    var_activa_forma_entradas_sin_comparacion = "MENU"
                                    frmentradas_compras_plantas.Show
                                    Unload Me
                                 Else
                                    If var_clave_movimiento = "ENVSIP" Then
                                       frmentradas_nota_envio.Caption = lv_movimientos.selectedItem.SubItems(1)
                                       frmentradas_nota_envio.lblnombremovimiento.Caption = lv_movimientos.selectedItem.SubItems(1)
                                       If Len(lv_movimientos.selectedItem.SubItems(1)) > 27 Then
                                          frmentradas_nota_envio.lblnombremovimiento.Font.Size = 16
                                       Else
                                          frmentradas_nota_envio.lblnombremovimiento.Font.Size = 20
                                       End If
                                       var_habilita_forma = False
                                       var_activa_forma_entradas_sin_comparacion = "MENU"
                                       frmentradas_nota_envio.Show
                                       Unload Me
                                    Else
                                       If var_clave_movimiento = "SAC" Then
                                          frmsalidas_almacen_calidad.Caption = lv_movimientos.selectedItem.SubItems(1)
                                          'frmsalidas_almacen_calidad.lblnombremovimiento.Caption = lv_movimientos.selectedItem.SubItems(1)
                                          'If Len(lv_movimientos.selectedItem.SubItems(1)) > 27 Then
                                          '   frmsalidas_almacen_calidad.lblnombremovimiento.Font.Size = 16
                                          'Else
                                          '   frmsalidas_almacen_calidad.lblnombremovimiento.Font.Size = 20
                                          'End If
                                          var_habilita_forma = False
                                          var_activa_forma_entradas_sin_comparacion = "MENU"
                                          frmsalidas_almacen_calidad.Show
                                          Unload Me
                                       Else
                                          If var_clave_movimiento = "EBI" Then
                                             frmentradas_bultos_intercompañias.Caption = lv_movimientos.selectedItem.SubItems(1)
                                             var_habilita_forma = False
                                             var_activa_forma_entradas_sin_comparacion = "MENU"
                                             frmentradas_bultos_intercompañias.Show
                                             Unload Me
                                          Else
                                             If var_clave_movimiento = "ECF" Then
                                                frmentradas_bultos_facturacion.Caption = lv_movimientos.selectedItem.SubItems(1)
                                                var_habilita_forma = False
                                                var_activa_forma_entradas_sin_comparacion = "MENU"
                                                frmentradas_bultos_facturacion.Show
                                                Unload Me
                                             Else
                                                frmentradas_sin_comparacion.Caption = lv_movimientos.selectedItem.SubItems(1)
                                                frmentradas_sin_comparacion.lblnombremovimiento.Caption = lv_movimientos.selectedItem.SubItems(1)
                                                If Len(lv_movimientos.selectedItem.SubItems(1)) > 27 Then
                                                   frmentradas_sin_comparacion.lblnombremovimiento.Font.Size = 16
                                                Else
                                                   frmentradas_sin_comparacion.lblnombremovimiento.Font.Size = 20
                                                End If
                                                var_habilita_forma = False
                                                var_activa_forma_entradas_sin_comparacion = "MENU"
                                                frmentradas_sin_comparacion.Show
                                                Unload Me
                                             End If
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                           End If
                        End If
                     End If
                  End If
               End If
            End If
         End If
      Else
         var_clave_movimiento = lv_movimientos.selectedItem
         frmentradas_reempaque.txt_clave_movimiento = var_clave_movimiento
         frmentradas_reempaque.txt_tipo_documento = lv_movimientos.selectedItem.SubItems(7)
         frmentradas_reempaque.lblnombremovimiento.Caption = lv_movimientos.selectedItem.SubItems(1)
         frmentradas_reempaque.Caption = lv_movimientos.selectedItem.SubItems(1)
         'frmentradas_reempaque.lbl_origen = lv_movimientos.SelectedItem.SubItems(4)
         If Len(lv_movimientos.selectedItem.SubItems(1)) > 30 Then
            frmentradas_reempaque.lblnombremovimiento.Font.Size = 18
         Else
            frmentradas_reempaque.lblnombremovimiento.Font.Size = 24
         End If
         var_habilita_forma = False
         var_activa_forma_entradas_reempaque = "MENU"
         frmentradas_reempaque.Show
         Unload Me
      End If
   End If
   If var_movimiento = "-" Then
      If lv_movimientos.selectedItem = "ST" Then
         var_clave_movimiento = lv_movimientos.selectedItem
         frmsalidas_intercompañias.lblnombremovimiento = lv_movimientos.selectedItem.SubItems(1)
         frmsalidas_intercompañias.Caption = lv_movimientos.selectedItem.SubItems(1)
         If Len(lv_movimientos.selectedItem.SubItems(1)) > 26 Then
            frmsalidas_intercompañias.lblnombremovimiento.Font.Size = 18
         Else
            frmsalidas_intercompañias.lblnombremovimiento.Font.Size = 24
         End If
         var_habilita_forma = False
         var_activa_forma_salidas_proveedor = "MENU"
         frmsalidas_intercompañias.Show
         Unload Me
      Else
         If Trim(lv_movimientos.selectedItem.SubItems(9)) = "1" Then
            var_clave_movimiento = lv_movimientos.selectedItem
            frmfactura_empresas.lblnombremovimiento.Caption = lv_movimientos.selectedItem.SubItems(1)
            frmfactura_empresas.Caption = lv_movimientos.selectedItem.SubItems(1)
            var_habilita_forma = False
            var_activa_forma_factura_empresas = "MENU"
            frmfactura_empresas.Show
            Unload Me
         Else
            If Trim(lv_movimientos.selectedItem.SubItems(3)) = 1 Then
               var_clave_movimiento = lv_movimientos.selectedItem
               If var_paquete = False Then
                  rs.Open "select * from tb_movimientos where vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
                  var_ajuste = 0
                  var_ajuste = IIf(IsNull(rs!INTE_MOV_AJUSTE), 0, rs!INTE_MOV_AJUSTE)
                  rs.Close
                  If var_ajuste = 1 Then
                     If var_clave_movimiento = "ENP" Then
                        frmsalidas_clientes.Caption = lv_movimientos.selectedItem.SubItems(1)
                        frmsalidas_clientes.lblnombremovimiento = lv_movimientos.selectedItem.SubItems(1)
                        frmsalidas_clientes.Caption = lv_movimientos.selectedItem.SubItems(1)
                        If Len(lv_movimientos.selectedItem.SubItems(1)) > 30 Then
                           frmsalidas_clientes.lblnombremovimiento.Font.Size = 18
                        Else
                           frmsalidas_clientes.lblnombremovimiento.Font.Size = 24
                        End If
                        var_habilita_forma = False
                        var_activa_forma_salidas_sin_comparacion = "MENU"
                        frmsalidas_clientes.Show
                        Unload Me
                     Else
                        If var_clave_movimiento = "NEM" Then
                           frmsalidas_numero_serie.Caption = lv_movimientos.selectedItem.SubItems(1)
                           frmsalidas_numero_serie.lblnombremovimiento = lv_movimientos.selectedItem.SubItems(1)
                           frmsalidas_numero_serie.Caption = lv_movimientos.selectedItem.SubItems(1)
                           If Len(lv_movimientos.selectedItem.SubItems(1)) > 30 Then
                              frmsalidas_numero_serie.lblnombremovimiento.Font.Size = 18
                           Else
                              frmsalidas_numero_serie.lblnombremovimiento.Font.Size = 24
                           End If
                           var_habilita_forma = False
                           var_activa_forma_salidas_sin_comparacion = "MENU"
                           frmsalidas_numero_serie.Show
                           Unload Me
                        Else
                           If var_clave_movimiento = "ENA" Then
                              frmsalidas_clientes.Caption = lv_movimientos.selectedItem.SubItems(1)
                              frmsalidas_clientes.lblnombremovimiento = lv_movimientos.selectedItem.SubItems(1)
                              frmsalidas_clientes.Caption = lv_movimientos.selectedItem.SubItems(1)
                              If Len(lv_movimientos.selectedItem.SubItems(1)) > 30 Then
                                 frmsalidas_clientes.lblnombremovimiento.Font.Size = 18
                              Else
                                 frmsalidas_clientes.lblnombremovimiento.Font.Size = 24
                              End If
                              var_habilita_forma = False
                              var_activa_forma_salidas_sin_comparacion = "MENU"
                              frmsalidas_clientes.Show
                              Unload Me
                           Else
                              frmsalidas_sin_comparacion.Caption = lv_movimientos.selectedItem.SubItems(1)
                              frmsalidas_sin_comparacion.lblnombremovimiento = lv_movimientos.selectedItem.SubItems(1)
                              frmsalidas_sin_comparacion.Caption = lv_movimientos.selectedItem.SubItems(1)
                              If Len(lv_movimientos.selectedItem.SubItems(1)) > 30 Then
                                 frmsalidas_sin_comparacion.lblnombremovimiento.Font.Size = 18
                              Else
                                 frmsalidas_sin_comparacion.lblnombremovimiento.Font.Size = 24
                              End If
                              var_habilita_forma = False
                              var_activa_forma_salidas_sin_comparacion = "MENU"
                              frmsalidas_sin_comparacion.Show
                              Unload Me
                           End If
                        End If
                     End If
                  Else
                     frmsalidas.txt_clave_movimiento = var_clave_movimiento
                     frmsalidas.lblnombremovimiento.Caption = lv_movimientos.selectedItem.SubItems(1)
                     frmsalidas.Caption = lv_movimientos.selectedItem.SubItems(1)
                     If Len(lv_movimientos.selectedItem.SubItems(1)) > 30 Then
                         frmsalidas.lblnombremovimiento.Font.Size = 18
                     Else
                        frmsalidas.lblnombremovimiento.Font.Size = 24
                     End If
                     var_habilita_forma = False
                     var_activa_forma_salidas = "MENU"
                     frmsalidas.Show
                     Unload Me
                  End If
               Else
                  If var_empresa = "03" Then
                     If var_trazabilidad = 1 Then
                        frmsalidas_empaques_trazabilidad.txt_clave_movimiento = var_clave_movimiento
                        frmsalidas_empaques_trazabilidad.txt_archivo = ""
                        frmsalidas_empaques_trazabilidad.txt_archivo.Enabled = True
                        frmsalidas_empaques_trazabilidad.lblnombremovimiento.Caption = "EMPACADO DE MERCANCIA"
                        frmsalidas_empaques_trazabilidad.Caption = "EMPACADO DE MERCANCIA"
                        var_habilita_forma = False
                        var_activa_forma_salidas_empaques = "MENU"
                        frmsalidas_empaques_trazabilidad.Show
                        Unload Me
                     Else
                        frmsalidas_empaques.txt_clave_movimiento = var_clave_movimiento
                        frmsalidas_empaques.txt_archivo = ""
                        frmsalidas_empaques.txt_archivo.Enabled = True
                        frmsalidas_empaques.lblnombremovimiento.Caption = "EMPACADO DE MERCANCIA"
                        frmsalidas_empaques.Caption = "EMPACADO DE MERCANCIA"
                        var_habilita_forma = False
                        var_activa_forma_salidas_empaques = "MENU"
                        frmsalidas_empaques.Show
                        Unload Me
                     End If
                  Else
                     frmsalidas_empaques.txt_clave_movimiento = var_clave_movimiento
                     frmsalidas_empaques.txt_archivo = ""
                     frmsalidas_empaques.txt_archivo.Enabled = True
                     frmsalidas_empaques.lblnombremovimiento.Caption = "EMPACADO DE MERCANCIA"
                     frmsalidas_empaques.Caption = "EMPACADO DE MERCANCIA"
                     var_habilita_forma = False
                     var_activa_forma_salidas_empaques = "MENU"
                     frmsalidas_empaques.Show
                     Unload Me
                  End If
               End If
            Else
               rs.Open "select * from tb_movimientos where vcha_mov_movimiento_id = '" + lv_movimientos.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
               var_clave_movimiento = lv_movimientos.selectedItem
               var_ajuste = 0
               If Not rs.EOF Then
                  var_ajuste = IIf(IsNull(rs!INTE_MOV_AJUSTE), 0, rs!INTE_MOV_AJUSTE)
               End If
               rs.Close
               If var_ajuste = 1 Then
                  If var_clave_movimiento = "DP" Then
                     frmsalidas_proveedor.lblnombremovimiento = lv_movimientos.selectedItem.SubItems(1)
                     frmsalidas_proveedor.Caption = lv_movimientos.selectedItem.SubItems(1)
                     If Len(lv_movimientos.selectedItem.SubItems(1)) > 26 Then
                         frmsalidas_proveedor.lblnombremovimiento.Font.Size = 18
                     Else
                        frmsalidas_proveedor.lblnombremovimiento.Font.Size = 24
                     End If
                     var_habilita_forma = False
                     var_activa_forma_salidas_proveedor = "MENU"
                     frmsalidas_proveedor.Show
                     Unload Me
                  Else
                     If var_clave_movimiento = "ENP" Then
                        frmsalidas_clientes.Caption = lv_movimientos.selectedItem.SubItems(1)
                        frmsalidas_clientes.lblnombremovimiento = lv_movimientos.selectedItem.SubItems(1)
                        frmsalidas_clientes.Caption = lv_movimientos.selectedItem.SubItems(1)
                        If Len(lv_movimientos.selectedItem.SubItems(1)) > 30 Then
                           frmsalidas_clientes.lblnombremovimiento.Font.Size = 18
                        Else
                           frmsalidas_clientes.lblnombremovimiento.Font.Size = 24
                        End If
                        var_habilita_forma = False
                        var_activa_forma_salidas_sin_comparacion = "MENU"
                        frmsalidas_clientes.Show
                        Unload Me
                     Else
                        If var_clave_movimiento = "NEM" Then
                           frmsalidas_numero_serie.Caption = lv_movimientos.selectedItem.SubItems(1)
                           frmsalidas_numero_serie.lblnombremovimiento = lv_movimientos.selectedItem.SubItems(1)
                           frmsalidas_numero_serie.Caption = lv_movimientos.selectedItem.SubItems(1)
                           If Len(lv_movimientos.selectedItem.SubItems(1)) > 10 Then
                              frmsalidas_numero_serie.lblnombremovimiento.Font.Size = 18
                           Else
                              frmsalidas_numero_serie.lblnombremovimiento.Font.Size = 24
                           End If
                           var_habilita_forma = False
                           var_activa_forma_salidas_sin_comparacion = "MENU"
                           frmsalidas_numero_serie.Show
                           Unload Me
                        Else
                           If var_clave_movimiento = "NESP" Then
                              frmsalidas_tiendas_sin_pedido.Caption = lv_movimientos.selectedItem.SubItems(1)
                              frmsalidas_tiendas_sin_pedido.lblnombremovimiento = lv_movimientos.selectedItem.SubItems(1)
                              frmsalidas_tiendas_sin_pedido.Caption = lv_movimientos.selectedItem.SubItems(1)
                              If Len(lv_movimientos.selectedItem.SubItems(1)) > 30 Then
                                 frmsalidas_tiendas_sin_pedido.lblnombremovimiento.Font.Size = 18
                              Else
                                 frmsalidas_tiendas_sin_pedido.lblnombremovimiento.Font.Size = 24
                              End If
                              var_habilita_forma = False
                              var_activa_forma_salidas_sin_comparacion = "MENU"
                              frmsalidas_tiendas_sin_pedido.Show
                              Unload Me
                           Else
                              'AQUI DEBE DE IR LA SALIDA A TEXTILERA'
                              If var_clave_movimiento = "SX" Then
                                 frmsalidas_textilera_almacen.Caption = lv_movimientos.selectedItem.SubItems(1)
                                 frmsalidas_textilera_almacen.lblnombremovimiento = lv_movimientos.selectedItem.SubItems(1)
                                 frmsalidas_textilera_almacen.Caption = lv_movimientos.selectedItem.SubItems(1)
                                 If Len(lv_movimientos.selectedItem.SubItems(1)) > 30 Then
                                    frmsalidas_textilera_almacen.lblnombremovimiento.Font.Size = 18
                                 Else
                                    frmsalidas_textilera_almacen.lblnombremovimiento.Font.Size = 24
                                 End If
                                 var_habilita_forma = False
                                 var_activa_forma_salidas_sin_comparacion = "MENU"
                                 frmsalidas_textilera_almacen.Show
                                 Unload Me
                              Else
                                 If var_clave_movimiento = "ENA" Then
                                    frmsalidas_clientes.Caption = lv_movimientos.selectedItem.SubItems(1)
                                    frmsalidas_clientes.lblnombremovimiento = lv_movimientos.selectedItem.SubItems(1)
                                    frmsalidas_clientes.Caption = lv_movimientos.selectedItem.SubItems(1)
                                    If Len(lv_movimientos.selectedItem.SubItems(1)) > 20 Then
                                       frmsalidas_clientes.lblnombremovimiento.Font.Size = 18
                                    Else
                                       frmsalidas_clientes.lblnombremovimiento.Font.Size = 24
                                    End If
                                    var_habilita_forma = False
                                    var_activa_forma_salidas_sin_comparacion = "MENU"
                                    frmsalidas_clientes.Show
                                    Unload Me
                                 Else
                                    If var_clave_movimiento = "DPL" Then
                                    
                                       
                                       
                                       frmsalidas_traspasos_plantas.Caption = lv_movimientos.selectedItem.SubItems(1)
                                       frmsalidas_traspasos_plantas.lblnombremovimiento = lv_movimientos.selectedItem.SubItems(1)
                                       frmsalidas_traspasos_plantas.Caption = lv_movimientos.selectedItem.SubItems(1)
                                       frmsalidas_traspasos_plantas.lblnombremovimiento.Font.Size = 18
                                       var_habilita_forma = False
                                       var_activa_forma_salidas_sin_comparacion = "MENU"
                                       frmsalidas_traspasos_plantas.Show
                                       Unload Me
                                    Else
                                       If var_clave_movimiento = "DPI" Then
                                          frmsalidas_proveedores_intercompañias.Caption = lv_movimientos.selectedItem.SubItems(1)
                                          frmsalidas_proveedores_intercompañias.lblnombremovimiento = lv_movimientos.selectedItem.SubItems(1)
                                          frmsalidas_proveedores_intercompañias.Caption = lv_movimientos.selectedItem.SubItems(1)
                                          frmsalidas_proveedores_intercompañias.lblnombremovimiento.Font.Size = 10
                                          var_habilita_forma = False
                                          var_activa_forma_salidas_proveedor = "MENU"
                                          frmsalidas_proveedores_intercompañias.Show
                                          Unload Me
                                       Else
                                          frmsalidas_sin_comparacion.Caption = lv_movimientos.selectedItem.SubItems(1)
                                          frmsalidas_sin_comparacion.lblnombremovimiento = lv_movimientos.selectedItem.SubItems(1)
                                          frmsalidas_sin_comparacion.Caption = lv_movimientos.selectedItem.SubItems(1)
                                          If Len(lv_movimientos.selectedItem.SubItems(1)) > 30 Then
                                             frmsalidas_sin_comparacion.lblnombremovimiento.Font.Size = 18
                                          Else
                                             frmsalidas_sin_comparacion.lblnombremovimiento.Font.Size = 24
                                          End If
                                          var_habilita_forma = False
                                          var_activa_forma_salidas_sin_comparacion = "MENU"
                                          frmsalidas_sin_comparacion.Show
                                          Unload Me
                                       End If
                                    End If
                                 End If
                              End If
                           End If
                        End If
                     End If
                  End If
               Else
                  var_clave_movimiento = lv_movimientos.selectedItem
                  If var_clave_movimiento = "FV" Then
                     var_habilita_forma = False
                     var_activa_forma_fact_merc_vistas = "MENU"
                     frmfact_merc_vistas.Show
                     Unload Me
                  End If
                  If var_clave_movimiento = "SV" Then
                     var_z = 1
                     If var_z = 1 Then
                        var_habilita_forma = False
                        var_activa_forma_salidas_proveedor = "MENU"
                        frmsalidas_vistas.Caption = lv_movimientos.selectedItem.SubItems(1)
                        frmsalidas_vistas.lblnombremovimiento = lv_movimientos.selectedItem.SubItems(1)
                        frmsalidas_vistas.Show
                        Unload Me
                     Else
                        frmsalidas.txt_clave_movimiento = var_clave_movimiento
                        frmsalidas.lblnombremovimiento.Caption = lv_movimientos.selectedItem.SubItems(1)
                        frmsalidas.Caption = lv_movimientos.selectedItem.SubItems(1)
                        If Len(lv_movimientos.selectedItem.SubItems(1)) > 30 Then
                            frmsalidas.lblnombremovimiento.Font.Size = 18
                        Else
                           frmsalidas.lblnombremovimiento.Font.Size = 24
                        End If
                        var_habilita_forma = False
                        var_activa_forma_salidas = "MENU"
                        frmsalidas.Show
                        Unload Me
                     End If
                     
                     
                  End If
                  If var_clave_movimiento = "VDI" Then
                     var_habilita_forma = False
                     var_activa_forma_salidas_proveedor = "MENU"
                     frmsalidas_ventas_directas.Caption = lv_movimientos.selectedItem.SubItems(1)
                     frmsalidas_ventas_directas.lblnombremovimiento = lv_movimientos.selectedItem.SubItems(1)
                     frmsalidas_ventas_directas.Show
                     Unload Me
                  Else
                     If var_clave_movimiento = "VDIP" Then
                        var_habilita_forma = False
                        var_activa_forma_salidas_proveedor = "MENU"
                        frmsalidas_ventas_directas_plantas.Caption = lv_movimientos.selectedItem.SubItems(1)
                        frmsalidas_ventas_directas_plantas.lblnombremovimiento = lv_movimientos.selectedItem.SubItems(1)
                        frmsalidas_ventas_directas_plantas.Show
                        Unload Me
                     Else
                        If var_clave_movimiento = "VMPSIP" Or var_clave_movimiento = "VDESP" Then
                           var_habilita_forma = False
                           var_activa_forma_salidas_proveedor = "MENU"
                           frmsalidas_ventas_materia_prima_SIP.Caption = lv_movimientos.selectedItem.SubItems(1)
                           frmsalidas_ventas_materia_prima_SIP.lblnombremovimiento = lv_movimientos.selectedItem.SubItems(1)
                           frmsalidas_ventas_materia_prima_SIP.Show
                           Unload Me
                        Else
                           If var_clave_movimiento = "VDIL" Then
                              var_habilita_forma = False
                              var_activa_forma_salidas_proveedor = "MENU"
                              frmsalidas_ventas_directas.Caption = lv_movimientos.selectedItem.SubItems(1)
                              frmsalidas_ventas_directas.lblnombremovimiento = lv_movimientos.selectedItem.SubItems(1)
                              frmsalidas_ventas_directas.Show
                              Unload Me
                           Else
                              If var_clave_movimiento = "STR" Then
                                 var_habilita_forma = False
                                 var_activa_forma_salidas_proveedor = "MENU"
                                 frmsalidas_transformacion.Caption = lv_movimientos.selectedItem.SubItems(1)
                                 frmsalidas_transformacion.lblnombremovimiento = lv_movimientos.selectedItem.SubItems(1)
                                 frmsalidas_transformacion.Show
                                 Unload Me
                              Else
                                 If var_clave_movimiento = "VDIPM" Then
                                    var_habilita_forma = False
                                    var_activa_forma_salidas_proveedor = "MENU"
                                    frmsalidas_ventas_cantia_textilera.Caption = lv_movimientos.selectedItem.SubItems(1)
                                    frmsalidas_ventas_cantia_textilera.lblnombremovimiento = lv_movimientos.selectedItem.SubItems(1)
                                    frmsalidas_ventas_cantia_textilera.Show
                                 Else
                                    If var_clave_movimiento = "DPOC" Then
                                       var_habilita_forma = False
                                       var_activa_forma_salidas_proveedor = "MENU"
                                       frmsalidas_proveedor_orden_compra.Caption = lv_movimientos.selectedItem.SubItems(1)
                                       frmsalidas_proveedor_orden_compra.lblnombremovimiento = lv_movimientos.selectedItem.SubItems(1)
                                       frmsalidas_proveedor_orden_compra.Show
                                    End If
                                 End If
                              End If
                           End If
                        End If
                     End If
                  End If
               
               End If
            End If
         End If
      End If
   End If

Exit Sub
salir:
   MsgBox "Existen errores en la configuración del movimiento, favor de reportarlo al departamento de sistemas", vbOKOnly, "ATENCION"
End Sub

Private Sub lv_movimientos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
   If KeyAscii = 13 Then
      Call lv_movimientos_DblClick
   End If
End Sub
