VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_listado_movimientos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Movimientos"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lv_movimientos 
      Height          =   3420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   6033
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
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
Attribute VB_Name = "frmoracle_listado_movimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Top = 2000
    Left = 2500
    Dim list_item As ListItem
    var_contador = 0
    If var_clave_usuario_global = "U0000000763" Then
       var_tipo_recepcion = "VDI"
       var_descripcion_recepcion = "VENTAS DIRECTAS"
       Set list_item = Me.lv_movimientos.ListItems.Add(, , var_tipo_recepcion)
       list_item.SubItems(1) = var_descripcion_recepcion
    Else
       For var_j = 1 To 4
           If var_clave_usuario_global = "U0000000680" Then
              If var_j = 1 Then
                 var_tipo_recepcion = "SP"
                 var_descripcion_recepcion = "SALIDAS PRIVALIA"
                 Set list_item = Me.lv_movimientos.ListItems.Add(, , var_tipo_recepcion)
                 list_item.SubItems(1) = var_descripcion_recepcion
                 var_tipo_recepcion = "EP"
                 var_descripcion_recepcion = "ENTRADAS PRIVALIA"
                 Set list_item = Me.lv_movimientos.ListItems.Add(, , var_tipo_recepcion)
                 list_item.SubItems(1) = var_descripcion_recepcion
                 var_contador = var_contador + 1
                 'rs.Open "select * from tb_movimientos WHERE VCHA_MOV_MOVIMIENTO_ID = '51'", cnn, adOpenDynamic, adLockOptimistic
                 'While Not rs.EOF
                 '      Set list_item = Me.lv_movimientos.ListItems.Add(, , rs!VCHA_MOV_MOVIMIENTO_ID)
                 '      list_item.SubItems(1) = UCase(rs!vcha_mov_nombre)
                 '      var_contador = var_contador + 1
                 '      rs.MoveNext
                 'Wend
                 'rs.Close
              
              End If
           Else
              If var_clave_usuario_global = "U0000000763" Then
                 var_tipo_recepcion = "VDI"
                 var_descripcion_recepcion = "VENTAS DIRECTAS"
              Else
                 
                 If var_clave_usuario_global = "U0000001250" Then
                    If var_j = 1 Then
                       var_tipo_recepcion = "SML"
                       var_descripcion_recepcion = "SALIDAS MERCADO LIBRE"
                       Set list_item = Me.lv_movimientos.ListItems.Add(, , var_tipo_recepcion)
                       list_item.SubItems(1) = var_descripcion_recepcion
                       'var_tipo_recepcion = "EML"
                       'var_descripcion_recepcion = "ENTRADAS MERCADO LIBRE"
                       'Set list_item = Me.lv_movimientos.ListItems.Add(, , var_tipo_recepcion)
                       'list_item.SubItems(1) = var_descripcion_recepcion
                       var_contador = var_contador + 1
           
                    End If
                 Else
                    If var_j = 1 Then
                       If var_clave_usuario_global <> "U0000000430" Then
                          var_tipo_recepcion = "VDI"
                          var_descripcion_recepcion = "VENTAS DIRECTAS"
                       End If
                    End If
                    If var_j = 2 Then
                       var_tipo_recepcion = "DC"
                       var_descripcion_recepcion = "DEVOLUCION DE CLIENTES"
                    End If
                    If var_j = 3 Then
                       If var_unidad_organizacional = 93 Then
                          var_tipo_recepcion = "DCS"
                          var_descripcion_recepcion = "DEVOLUCION DE CLIENTES CON REFERENCIA"
                       End If
                    End If
                    If var_j = 4 Then
                       If var_unidad_organizacional = 93 Or var_unidad_organizacional = 90 Then
                          var_tipo_recepcion = "SNC"
                          var_descripcion_recepcion = "D.C. SOLO NOTA DE CREDITO"
                       End If
                    End If
                    Set list_item = Me.lv_movimientos.ListItems.Add(, , var_tipo_recepcion)
                    list_item.SubItems(1) = var_descripcion_recepcion
                    var_contador = var_contador + 1
                 End If

              End If
           End If
       Next var_j
    End If
    'rs.Open "select * from mtl_transaction_types where transaction_type_id in (21,2,51) order by transaction_type_name ", cnnoracle_4, adOpenDynamic, adLockOptimistic
    'While Not rs.EOF
    '      Set list_item = Me.lv_movimientos.ListItems.Add(, , rs!transaction_type_id)
    '      list_item.SubItems(1) = UCase(rs!Description)
    '      var_contador = var_contador + 1
    '      rs.MoveNext
    'Wend
    'rs.Close
    'MsgBox cnnoracle_4.ConnectionString
    
    'rs.Open "select * from xxvia_tb_tipo_movimientos where numb_tmo_movimiento_id in (21,51,2,3,0, 22, 23,35,17,40,13, 49, 50, 54) order by vcha_tmo_descripcion", cnnoracle_4, adOpenDynamic, adLockOptimistic
    'rs.Open "select * from vw_permisos_movimientos where vcha_usu_usuario_id = '" + var_clave_usuario_global + "'", cnn, adOpenDynamic, adLockOptimistic
    If var_clave_usuario_global = "U0000001250" Then
       If var_j = 1 Then
          var_tipo_recepcion = "SML"
          var_descripcion_recepcion = "SALIDAS MERCADO LIBRE"
          Set list_item = Me.lv_movimientos.ListItems.Add(, , var_tipo_recepcion)
          list_item.SubItems(1) = var_descripcion_recepcion
          var_tipo_recepcion = "EML"
          var_descripcion_recepcion = "ENTRADAS MERCADO LIBRE"
          Set list_item = Me.lv_movimientos.ListItems.Add(, , var_tipo_recepcion)
          list_item.SubItems(1) = var_descripcion_recepcion
          var_contador = var_contador + 1
          'rs.Open "select * from tb_movimientos WHERE VCHA_MOV_MOVIMIENTO_ID = '51'", cnn, adOpenDynamic, adLockOptimistic
          'While Not rs.EOF
          '      Set list_item = Me.lv_movimientos.ListItems.Add(, , rs!VCHA_MOV_MOVIMIENTO_ID)
          '      list_item.SubItems(1) = UCase(rs!vcha_mov_nombre)
          '      var_contador = var_contador + 1
          '      rs.MoveNext
          'Wend
          'rs.Close
           
        End If
    Else
    
       If var_clave_usuario_global <> "U0000000430" Then
          If var_clave_usuario_global = "U0000000680" Then
          Else
             rs.Open "select * from tb_movimientos", cnn, adOpenDynamic, adLockOptimistic
             While Not rs.EOF
                   Set list_item = Me.lv_movimientos.ListItems.Add(, , rs!VCHA_MOV_MOVIMIENTO_ID)
                   list_item.SubItems(1) = UCase(rs!vcha_mov_nombre)
                   var_contador = var_contador + 1
                   rs.MoveNext
             Wend
             rs.Close
          End If
       End If
    End If
    If var_contador > 9 Then
       lv_movimientos.ColumnHeaders(2).Width = 6380
    Else
    End If
    Me.lv_movimientos.ListItems.Item(1).Selected = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_movimientos_DblClick()
   var_nombre_movimiento_global = Me.lv_movimientos.selectedItem.SubItems(1)
   var_clave_movimiento = Me.lv_movimientos.selectedItem
   var_descripcion_recepcion = Me.lv_movimientos.selectedItem.SubItems(1)
   If var_clave_movimiento = "VENDOR" Or var_clave_movimiento = "INTERNAL ORDER" Or var_clave_movimiento = "INVENTORY" Then
      frmoracle_entradas.Show
   Else
      If var_clave_movimiento = "DC" Or var_clave_movimiento = "VDI" Or var_clave_movimiento = "SNC" Or var_clave_movimiento = "SP" Or var_clave_movimiento = "SML" Then
         frmoracle_devoluciones_clientes.lblnombremovimiento = Me.lv_movimientos.selectedItem.SubItems(1)
         frmoracle_devoluciones_clientes.Show
      Else
         If var_clave_movimiento = "DCS" Then
            frmoracle_devoluciones_clientes_sello.lblnombremovimiento = Me.lv_movimientos.selectedItem.SubItems(1)
            frmoracle_devoluciones_clientes_sello.Show
         Else
            If var_clave_movimiento = "EP" Then
               frmoracle_entradas_comparacion.lblnombremovimiento = Me.lv_movimientos.selectedItem.SubItems(1)
               frmoracle_entradas_comparacion.Show 1
            Else
               If var_clave_movimiento = "ACD" Then
                  frmoracle_asignacion_causas_devolucion_2.Show 1
               Else
                  frmoracle_subinventarios.Show 1
                  x = Shell(App.Path + "/MovimietosInventarios.exe " + var_unidad_organizacional + "|" + var_almacen_global + "|" + Me.lv_movimientos.selectedItem + "|#|" + var_clave_usuario_global + "-" + var_nombre_usuario_global)
               End If
            End If
         End If
      End If
   End If
End Sub

Private Sub lv_movimientos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call lv_movimientos_DblClick
   End If
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub
