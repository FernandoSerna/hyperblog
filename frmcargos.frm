VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcargos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cargos distintos a Facturas y Notas de Cargo"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   1965
      Left            =   1560
      TabIndex        =   7
      Top             =   300
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1410
         Left            =   30
         TabIndex        =   8
         Top             =   450
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   2487
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
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7584
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   9
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frmcargos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   435
      Picture         =   "frmcargos.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7635
      Picture         =   "frmcargos.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Caption         =   " Datos Generales "
      Height          =   2010
      Left            =   90
      TabIndex        =   10
      Top             =   420
      Width           =   7920
      Begin VB.TextBox txt_clase_cartera 
         Height          =   315
         Left            =   885
         TabIndex        =   0
         Top             =   225
         Width           =   1725
      End
      Begin VB.TextBox txt_nombre_clase_cartera 
         Height          =   315
         Left            =   2640
         TabIndex        =   1
         Top             =   225
         Width           =   5190
      End
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   885
         TabIndex        =   5
         Top             =   1275
         Width           =   1725
      End
      Begin VB.TextBox txt_plazo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   885
         TabIndex        =   4
         Top             =   930
         Width           =   1725
      End
      Begin VB.TextBox txt_clave_cliente 
         Height          =   315
         Left            =   885
         TabIndex        =   2
         Top             =   585
         Width           =   1725
      End
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   315
         Left            =   2640
         TabIndex        =   3
         Top             =   585
         Width           =   5190
      End
      Begin VB.TextBox txt_serie 
         Height          =   315
         Left            =   885
         TabIndex        =   6
         Top             =   1605
         Width           =   795
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   165
         TabIndex        =   20
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lbl_moneda 
         Height          =   285
         Left            =   2640
         TabIndex        =   15
         Top             =   1290
         Width           =   2985
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Monto:"
         Height          =   195
         Left            =   165
         TabIndex        =   14
         Top             =   1335
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Plazo:"
         Height          =   195
         Left            =   165
         TabIndex        =   13
         Top             =   990
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   165
         TabIndex        =   12
         Top             =   645
         Width           =   525
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   165
         TabIndex        =   11
         Top             =   1665
         Width           =   405
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   75
      TabIndex        =   19
      Top             =   285
      Width           =   7920
   End
End
Attribute VB_Name = "frmcargos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Dim var_almacen As String
   Dim var_grupo_actual As String
   Dim var_grupo_real As String
   Dim var_cliente As String
   Dim var_titular As String
   Dim var_establecimiento As String
   Dim var_clave_moneda As String
   Dim var_agente As String
   Dim var_imprimir As Boolean
   Dim var_contador As Integer
   Dim var_contador_notas As Integer
   Dim var_tipo_Cambio As Double
   Dim var_iva As Double
   Dim var_importe_total As Double
   Dim var_importe_iva As Double
   Dim var_subimporte As Double
   Dim var_importe As Double
   Dim var_plazo As Integer
   Dim si, i, n As Integer
   Dim var_serie As String
   Dim var_tipo_lista As Integer
   




Private Sub cmd_imprimir_Click()
   Set TB_ENCABEZADO_CARTERA_I = New TB_ENCABEZADO_CARTERA_I
   Set TB_ESTADO_CUENTA_INSERTA = New TB_ESTADO_CUENTA_INSERTA
   Dim si As Integer
   Dim var_importe_iva As Double
   Dim var_importe_total As Double
   Dim var_importe_neto As Double
   Dim var_subimporte As Double
   Dim var_tipo_Cambio As Double
   Dim var_numero_folio As Double
   Dim var_moneda_local As Integer
   Dim var_posible_tipo_cambio As Boolean
   Dim var_mes_str As String
   Dim var_anio_str As String
   Dim var_dia_str As String
   Dim var_linea As String
   Dim var_posible_comision As Boolean
   Dim var_tipo_comision As Integer
   var_moneda_local = 0
   If rsaux4.State = 1 Then
      rsaux4.Close
   End If
   
   If Trim(Me.txt_clase_cartera) <> "" Then
      rsaux4.Open "select * from tb_Series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_ser_serie_id = '" + txt_serie + "'", cnn, adOpenDynamic, adLockOptimistic
      If rsaux4.EOF Then
         If Me.txt_serie <> "" Then
            If Trim(Me.txt_clave_cliente) <> "" Then
               If IsNumeric(Me.txt_importe) Then
                  If Trim(var_clave_moneda) <> "" Then
                     rs.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        var_moneda_local = IIf(IsNull(rs!inte_mon_moneda_local), 0, rs!inte_mon_moneda_local)
                     End If
                     rs.Close
                     var_tipo_Cambio = 1
                     If var_moneda_local = 0 Then
                        rs.Open "select * from vw_tipocambio_fecha where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           var_tipo_Cambio = IIf(IsNull(rs!mone_tca_importe), 1, rs!mone_tca_importe)
                           var_posible_tipo_cambio = True
                        Else
                           var_posible_tipo_cambio = False
                        End If
                        rs.Close
                     Else
                        var_posible_tipo_cambio = True
                     End If
                     If var_posible_tipo_cambio = True Then
                        si = MsgBox("¿Deseas hacer el cargo en cartera?", vbYesNo, "ATENCION")
                        If si = 6 Then
                           si = MsgBox("Confirmar el cargo en cartera", vbYesNo, "ATENCION")
                           If si = 6 Then
                              cnn.BeginTrans
                              rs.Open "select max(inte_Car_numero) from tb_encabezado_Cartera where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_tipo_documento = 'CR' and vcha_ser_Serie_id = '" + Me.txt_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rs.EOF Then
                                 var_numero_folio = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
                              Else
                                 var_numero_folio = 1
                              End If
                              rs.Close
                              var_importe_total = (txt_importe / (1 + (var_iva / 100))) * var_tipo_Cambio
                              var_importe_neto = txt_importe * var_tipo_Cambio
                              var_importe_iva = var_importe_neto - var_importe_total
                              var_subimporte = var_importe_total
                              var_insertar = False
                              var_serie = Me.txt_serie
                              var_tipo_comision = 0
                              var_tipo_comision = 0
                              Cadena = "INSERT INTO TB_ENCABEZADO_CARTERA (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_CAR_CLASE_ID, INTE_CAR_NUMERO, CHAR_CAR_AFECTACION, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO, DTIM_CAR_FECHA, VCHA_AGE_AGENTE_ID, VCHA_GAC_GRUPO_ACTUAL_ID, VCHA_GRE_GRUPO_REAL_ID, VCHA_TIT_TITULAR_ID, VCHA_CLI_CLAVE_ID, VCHA_ESB_ESTABLECIMIENTO_ID, INTE_CAR_PLAZO, FLOA_CAR_PORCENTAJE_IVA, FLOA_CAR_PORCENTAJE_IMPUESTO_1, FLOA_CAR_PORCENTAJE_IMPUESTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_1, "
                              Cadena = Cadena + " FLOA_CAR_PORCENTAJE_DESCUENTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_3, FLOA_CAR_IMPORTE_TOTAL, FLOA_CAR_IMPORTE_IVA, FLOA_CAR_IMPORTE_IMPUESTO_1, FLOA_CAR_IMPORTE_IMPUESTO_2, FLOA_CAR_IMPORTE_DESCUENTO_1, FLOA_CAR_IMPORTE_DESCUENTO_2, FLOA_CAR_IMPORTE_DESCUENTO_3, FLOA_CAR_SUBIMPORTE, FLOA_CAR_IMPORTE_NETO, VCHA_CAR_IMPORTE_LETRA, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA, VCHA_AUD_FECHA, FLOA_CAR_SALDO, DTIM_CAR_FECHA_VENCIMIENTO, DTIM_CAR_FECHA_ENTREGA, VCHA_MON_MONEDA_ID, FLOA_CAR_TIPO_CAMBIO, VCHA_SER_SERIE_ID, CHAR_CAR_ESTATUS, INTE_cAR_TIPO_COMISION) Values "
                              Cadena = Cadena + " ('" + var_empresa + "', '" + var_unidad_organizacional + "', 'CR', 'CR', '" + Me.txt_clase_cartera + "', " + CStr(var_numero_folio) + ", '+', '', '', 0, GETDATE(), '" + var_agente + "', '" + var_grupo_actual + "', '" + var_grupo_real + "', '" + var_titular + "', '" + txt_clave_cliente + "', '', " + CStr(var_plazo) + ", " + CStr(var_iva) + ", 0, 0, 0, 0, 0, " + CStr(var_importe_total) + ", " + CStr(var_importe_iva) + ", 0, 0, 0, 0, 0, " + CStr(var_subimporte) + ", " + CStr(var_importe_neto) + ", '', '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', GETDate(), 0, GETDate(), GETDate(), '" + var_clave_moneda + "', " + CStr(var_tipo_Cambio) + ", '" + var_serie + "', ''," + CStr(var_tipo_comision) + ")"
                              rsaux5.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                              var_inserta = TB_ESTADO_CUENTA_INSERTA.Anadir(var_empresa, var_serie, "CR", var_numero_folio, "", "", 0, var_importe_neto, 0)
                              cnn.CommitTrans
                           End If
                        End If
                     Else
                        MsgBox "No se a asignado el tipo de cambio del dia de hoy", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "El cliente no tiene una moneda asociada", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "Importe incorrecto", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "Se debe de indicar un cliente", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "No se a indicado una serie", vbOKOnly, "ATENCION"
         End If
      Else
          MsgBox "La serie no debe de ser " + txt_serie, vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Se debe de indicar un tipo de cargo", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   txt_plazo.Enabled = True
   txt_importe.Enabled = True
   txt_clave_cliente = ""
   txt_plazo = ""
   txt_importe = ""
   txt_clave_empresa = ""
   txt_serie = ""
   lbl_moneda = ""
   Me.txt_nombre_cliente = ""
   txt_clave_cliente.Enabled = True
   txt_clave_cliente.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 78 Then
      cmd_nuevo_Click
   End If
   If Shift = 4 And KeyCode = 73 Then
      cmd_imprimir_Click
   End If
End Sub

Private Sub Form_Load()
   frm_lista.Visible = False
   var_cadena_seguridad = ""
   Top = 2500
   Left = 1800
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_reporte_valuacion_devoluciones)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         If var_tipo_lista = 1 Then
            txt_clave_cliente = lv_lista.selectedItem
            txt_nombre_cliente = lv_lista.selectedItem.SubItems(1)
            If txt_clave_cliente.Enabled = True Then
               txt_clave_cliente.SetFocus
            End If
         End If
         If var_tipo_lista = 2 Then
            Me.txt_clase_cartera = lv_lista.selectedItem
            Me.txt_nombre_clase_cartera = lv_lista.selectedItem.SubItems(1)
            If txt_clase_cartera.Enabled = True Then
               txt_clase_cartera.SetFocus
            End If
         End If
      End If
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub











Private Sub txt_clase_cartera_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_tipo_lista = 1
      lv_lista.ListItems.Clear
      Dim list_item As ListItem
      Dim var_contador_lista As Integer
      rs.Open "select vcha_car_clase_id, vcha_car_nombre from tb_clases_cartera where vcha_car_documento= 'CR' order by vcha_Car_nombre", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         While Not rs.EOF
            var_contador_lista = var_contador_lista + 1
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_CAR_CLASE_ID)
            list_item.SubItems(1) = Trim(IIf(IsNull(rs!VCHA_CAR_NOMBRE), "", rs!VCHA_CAR_NOMBRE))
            rs.MoveNext
         Wend
      End If
      rs.Close
      lbl_lista = "TIPOS DE CARGO"
      var_tipo_lista = 2
      var_n = lv_lista.ListItems.Count
      If var_n > 5 Then
         lv_lista.ColumnHeaders(2).Width = 4070.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4299.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_clase_cartera_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_clase_cartera_LostFocus()
   If Trim(Me.txt_clase_cartera) <> "" Then
      rs.Open "select * from tb_clases_cartera where vcha_Car_clase_id = '" + Me.txt_clase_cartera + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_clase_cartera = IIf(IsNull(rs!VCHA_CAR_NOMBRE), "", rs!VCHA_CAR_NOMBRE)
      Else
         MsgBox "Clase de tipo de cargo no existe", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      Me.txt_nombre_clase_cartera = ""
   End If
End Sub

Private Sub txt_clave_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select DISTINCT VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE from VW_clientes WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_Cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_cli_nombre), "", rs!vcha_cli_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CLIENTES"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 5 Then
         lv_lista.ColumnHeaders(2).Width = 4070.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4299.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_clave_cliente_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_clave_cliente_LostFocus()
   If Trim(txt_clave_cliente) <> "" Then
      rs.Open "SELECT * FROM VW_CLIENTES WHERE VCHA_CLI_CLAVE_ID =  '" + txt_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         
         txt_nombre_cliente = IIf(IsNull(rs!vcha_cli_nombre), "", rs!vcha_cli_nombre)
         txt_plazo = IIf(IsNull(rs!inte_pla_dias), 0, rs!inte_pla_dias)
         var_grupo_actual = IIf(IsNull(rs!VCHA_GAC_GRUPO_ACTUAL_ID), "", rs!VCHA_GAC_GRUPO_ACTUAL_ID)
         var_grupo_real = IIf(IsNull(rs!vcha_gre_grupo_real_id), "", rs!vcha_gre_grupo_real_id)
         var_cliente = txt_clave_cliente
         var_titular = IIf(IsNull(rs!VCHA_TIT_TITULAR_ID), "", rs!VCHA_TIT_TITULAR_ID)
         var_clave_moneda = IIf(IsNull(rs!VCHA_MON_MONEDA_ID), "", rs!VCHA_MON_MONEDA_ID)
         var_agente = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
         var_plazo = IIf(IsNull(rs!inte_pla_dias), 0, rs!inte_pla_dias)
         txt_plazo = var_plazo
         var_iva = IIf(IsNull(rs!FLOA_TPE_IVA), 0, rs!FLOA_TPE_IVA)
         var_clave_moneda = IIf(IsNull(rs!VCHA_MON_MONEDA_ID), "", rs!VCHA_MON_MONEDA_ID)
         lbl_moneda = IIf(IsNull(rs!vcha_mon_nombre_plural), "", rs!vcha_mon_nombre_plural)
      Else
         MsgBox "Clave de cliente incorrecto", vbOKOnly, "ATENCION"
         txt_clave_cliente = ""
         txt_nombre_cliente = ""
         Me.txt_plazo = ""
         var_grupo_actual = ""
         var_grupo_real = ""
         var_cliente = ""
         var_titular = ""
         var_clave_moneda = ""
         var_agente = ""
         var_plazo = 0
         txt_plazo = ""
         var_iva = 0
         var_clave_moneda = ""
         lbl_moneda = ""
      End If
      rs.Close
   End If
End Sub


Private Sub txt_importe_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_importe_LostFocus()
   If Not IsNumeric(txt_importe) Then
      MsgBox "Importe Incorrecto", vbOKOnly, "ATENCION"
      txt_importe = 0
   End If
End Sub

Private Sub txt_nombre_clase_cartera_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_tipo_lista = 1
      lv_lista.ListItems.Clear
      Dim list_item As ListItem
      Dim var_contador_lista As Integer
      rs.Open "select vcha_car_clase_id, vcha_car_nombre from tb_clases_cartera where vcha_car_documento= 'CR' order by vcha_Car_nombre", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         While Not rs.EOF
            var_contador_lista = var_contador_lista + 1
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_CAR_CLASE_ID)
            list_item.SubItems(1) = Trim(IIf(IsNull(rs!VCHA_CAR_NOMBRE), "", rs!VCHA_CAR_NOMBRE))
            rs.MoveNext
         Wend
      End If
      rs.Close
      lbl_lista = "TIPOS DE CARGO"
      var_tipo_lista = 2
      var_n = lv_lista.ListItems.Count
      If var_n > 5 Then
         lv_lista.ColumnHeaders(2).Width = 4070.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4299.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_clase_cartera_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frm_lista.Left = 1530
      frm_lista.Top = 2280
      lv_lista.ListItems.Clear
      rs.Open "select DISTINCT VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE from VW_clientes VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_Cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_cli_nombre), "", rs!vcha_cli_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CLIENTES"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 5 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_tipo_comision_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_tipo_comision where  VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' order by vcha_TCO_nombre", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!inte_tco_tipo_comision_id)
               list_item.SubItems(1) = IIf(IsNull(rs!VCHA_tCO_NOMBRE), "", rs!VCHA_tCO_NOMBRE)
               rs.MoveNext
         Wend
         rs.Close
         lbl_lista = "COMISIONES"
         var_tipo_lista = 2
         Dim var_n As Integer
         var_n = lv_lista.ListItems.Count
         If var_n > 5 Then
            lv_lista.ColumnHeaders(2).Width = 4270.71
         Else
            lv_lista.ColumnHeaders(2).Width = 4499.71
         End If
         frm_lista.Visible = True
         lv_lista.SetFocus
      Else
         MsgBox "La empresa no tiene comisiones asociadas", vbOKOnly, "ATENCION"
         rs.Close
      End If
   End If
End Sub

Private Sub txt_nombre_tipo_comision_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_plazo_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_plazo_LostFocus()
   If IsNumeric(txt_plazo) Then
      var_plazo = txt_plazo
   Else
      MsgBox "Plazo incorrecto", vbOKOnly, "ATENCION"
      txt_plazo = var_plazo
   End If
End Sub

Private Sub txt_serie_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   End If
End Sub



