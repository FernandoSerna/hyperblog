VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcheques_devueltos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de cheques devueltos"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   2190
      Left            =   1560
      TabIndex        =   19
      Top             =   330
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1695
         Left            =   30
         TabIndex        =   20
         Top             =   450
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   2990
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
         TabIndex        =   21
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      Picture         =   "frmcheques_devueltos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   465
      Picture         =   "frmcheques_devueltos.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7665
      Picture         =   "frmcheques_devueltos.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Caption         =   " Datos Generales "
      Height          =   2355
      Left            =   120
      TabIndex        =   11
      Top             =   435
      Width           =   7920
      Begin VB.TextBox txt_nombre_banco 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2625
         MaxLength       =   4
         TabIndex        =   8
         Top             =   1260
         Width           =   5190
      End
      Begin VB.TextBox txt_banco 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   885
         MaxLength       =   4
         TabIndex        =   7
         Top             =   1260
         Width           =   1725
      End
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   315
         Left            =   2625
         TabIndex        =   4
         Top             =   255
         Width           =   5190
      End
      Begin VB.ComboBox cmb_series 
         Height          =   315
         Left            =   885
         TabIndex        =   10
         Top             =   1920
         Width           =   795
      End
      Begin VB.TextBox txt_cheque 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   885
         MaxLength       =   4
         TabIndex        =   6
         Top             =   930
         Width           =   1725
      End
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   885
         TabIndex        =   9
         Top             =   1590
         Width           =   1725
      End
      Begin VB.TextBox txt_plazo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   885
         TabIndex        =   5
         Top             =   585
         Width           =   1725
      End
      Begin VB.TextBox txt_clave_cliente 
         Height          =   315
         Left            =   885
         TabIndex        =   3
         Top             =   255
         Width           =   1725
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
         Height          =   195
         Left            =   165
         TabIndex        =   22
         Top             =   1320
         Width           =   510
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   165
         TabIndex        =   18
         Top             =   1980
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cheque:"
         Height          =   195
         Left            =   165
         TabIndex        =   17
         Top             =   990
         Width           =   600
      End
      Begin VB.Label lbl_moneda 
         Height          =   285
         Left            =   2670
         TabIndex        =   15
         Top             =   1575
         Width           =   2985
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Monto:"
         Height          =   195
         Left            =   165
         TabIndex        =   14
         Top             =   1650
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Plazo:"
         Height          =   195
         Left            =   165
         TabIndex        =   13
         Top             =   660
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   165
         TabIndex        =   12
         Top             =   315
         Width           =   525
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   105
      TabIndex        =   16
      Top             =   300
      Width           =   7920
   End
End
Attribute VB_Name = "frmcheques_devueltos"
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



Private Sub cmb_clientes_LostFocus()
   If Trim(txt_clave_cliente) <> "" Then
      rs.Open "select * from vw_clientes where vcha_cli_clave_id = '" + txt_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_grupo_actual = IIf(IsNull(rs!vcha_gac_grupo_Actual_id), "", rs!vcha_gac_grupo_Actual_id)
         var_grupo_real = IIf(IsNull(rs!vcha_gre_grupo_real_id), "", rs!vcha_gre_grupo_real_id)
         var_cliente = txt_clave_cliente
         var_titular = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
         var_clave_moneda = IIf(IsNull(rs!VCHA_MON_MONEDA_ID), "", rs!VCHA_MON_MONEDA_ID)
         var_agente = IIf(IsNull(rs!vcha_age_agente_id), "", rs!vcha_age_agente_id)
         var_plazo = 0
         txt_plazo = var_plazo
         var_iva = IIf(IsNull(rs!FLOA_TPE_IVA), 0, rs!FLOA_TPE_IVA)
         var_clave_moneda = IIf(IsNull(rs!VCHA_MON_MONEDA_ID), "", rs!VCHA_MON_MONEDA_ID)
         lbl_moneda = IIf(IsNull(rs!vcha_mon_nombre_plural), "", rs!vcha_mon_nombre_plural)
      End If
      rs.Close
   End If
End Sub

Private Sub cmb_series_Click()
  var_serie = cmb_series
End Sub

Private Sub cmb_series_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   End If
End Sub

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
   var_moneda_local = 0
   If Trim(txt_clave_cliente) <> "" Then
      If Trim(txt_cheque) <> "" Then
         If IsNumeric(txt_importe) Then
            If CDbl(txt_importe) > 0 Then
               If Me.txt_banco <> "" Then
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
                     txt_folio = 1
                     If Trim(txt_folio) <> "" Then
                        si = MsgBox("¿Deseas Aplicar el cargo por cheque devuelto?", vbYesNo, "ATENCION")
                        If si = 6 Then
                           si = MsgBox("Confirmar el cargo por cheque devuelto", vbYesNo, "ATENCION")
                           If si = 6 Then
                              cnn.BeginTrans
                              rs.Open "select max(inte_car_numero) from tb_encabezado_cartera where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_SER_SERIE_ID = '" + var_serie + "' and vcha_car_tipo_documento = 'CH'", cnn, adOpenDynamic, adLockOptimistic
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
                              var_insertar = TB_ENCABEZADO_CARTERA_I.Anadir(var_empresa, var_unidad_organizacional, "CH", "CH", "CH", var_numero_folio, "+", "", "", 0, CStr(Date), var_agente, var_grupo_actual, var_grupo_real, var_titular, txt_clave_cliente, "", var_plazo, var_iva, 0, 0, 0, 0, 0, var_importe_total, var_importe_iva, 0, 0, 0, 0, 0, var_subimporte, var_importe_neto, "", var_clave_usuario_global, fun_NombrePc, Date, 0, Date, Date, var_clave_moneda, var_tipo_Cambio, var_serie, "")
                              rsaux4.Open "update tb_encabezado_cartera set VCHA_CAR_CHEQUE = '" + txt_cheque + "',VCHA_CAR_CHEQUE_DEPOSITO = '" + Me.txt_banco + "', VCHA_CAR_banco_DEPOSITO = '" + Me.txt_banco + "'  where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_tipo_documento = 'CH' and inte_car_numero = " + CStr(var_numero_folio) + " and VCHA_SER_SERIE_ID = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                              var_inserta = TB_ESTADO_CUENTA_INSERTA.Anadir(var_empresa, var_serie, "CH", var_numero_folio, "", "", 0, var_importe_neto, 0)
                              cnn.CommitTrans
                              MsgBox "Se a aplicado el cheque devuelto", vbOKOnly, "ATENCION"
                              txt_plazo.Enabled = True
                              txt_importe.Enabled = True
                              txt_nombre_cliente.Enabled = True
                              txt_nombre_cliente = ""
                              txt_clave_cliente = ""
                              txt_plazo = ""
                              txt_importe = ""
                              txt_clave_empresa = ""
                              lbl_moneda = ""
                           End If
                        End If
                     Else
                        MsgBox "No se a indicado un folio", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "No se a asignado el tipo de cambio del dia de hoy", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "No se a indicado un banco", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "Importe incorrecto", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Importe incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "No se a indicado un chueque", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a seleccionado un cliente", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   txt_plazo.Enabled = True
   txt_importe.Enabled = True
   txt_nombre_cliente.Enabled = True
   txt_nombre_cliente = ""
   txt_clave_cliente = ""
   txt_plazo = ""
   txt_importe = ""
   txt_clave_empresa = ""
   lbl_moneda = ""
   txt_clave_cliente.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   frm_lista.Visible = False
   var_cadena_seguridad = ""
   Top = 2500
   Left = 1800
   rs.Open "select vcha_ser_serie_id from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_contador_serie = 0
      While Not rs.EOF
         var_contador_serie = var_contador_serie + 1
         rs.MoveNext
      Wend
      rs.MoveFirst
      txt_clave_cliente.Enabled = True
      Call RecsetToCombo(cmb_series.hwnd, rs, 0)
      If var_contador_serie > 1 Then
         cmb_series.Enabled = True
      Else
         cmb_series.Enabled = False
      End If
      rs.MoveFirst
      cmb_series = rs!VCHA_SER_SERIE_ID
      var_serie = rs!VCHA_SER_SERIE_ID
   Else
      MsgBox "No se a indicado una serie para esta Unidad organizacional", vbOKOnly, "ATENCION"
      txt_clave_cliente.Enabled = False
      txt_nombre_cliente.Enabled = False
   End If
   rs.Close
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_cheques_devueltos)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         If var_tipo_lista = 1 Then
            txt_clave_cliente = lv_lista.selectedItem
            txt_nombre_cliente = lv_lista.selectedItem.SubItems(1)
            txt_clave_cliente.SetFocus
         End If
         If var_tipo_lista = 2 Then
            Me.txt_banco = lv_lista.selectedItem
            Me.txt_nombre_banco = lv_lista.selectedItem.SubItems(1)
            Me.txt_banco.SetFocus
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

Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_bancos where vcha_ban_banco_id = '20' or vcha_ban_banco_id = '11' or vcha_ban_banco_id = '10' or vcha_ban_banco_id = '22' order by vcha_ban_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_ban_banco_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_BAN_NOMBRE), "", rs!VCHA_BAN_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "BANCO DEL DEPOSITO"
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
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_banco_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_bancos where vcha_ban_banco_id = '20' or vcha_ban_banco_id = '11' or vcha_ban_banco_id = '10' or vcha_ban_banco_id = '22' or vcha_ban_banco_id = '23'  order by vcha_ban_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_ban_banco_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_BAN_NOMBRE), "", rs!VCHA_BAN_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "BANCO DEL DEPOSITO"
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
End Sub

Private Sub txt_banco_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_cheque_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_clave_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select DISTINCT VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE from vw_clientes where vcha_emp_empresa_id = '" + var_empresa + "' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
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
      rs.Open "select * from vw_clientes where vcha_cli_clave_id = '" + txt_clave_cliente + "' and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_cliente = IIf(IsNull(rs!vcha_cli_nombre), "", rs!vcha_cli_nombre)
         var_grupo_actual = IIf(IsNull(rs!vcha_gac_grupo_Actual_id), "", rs!vcha_gac_grupo_Actual_id)
         var_grupo_real = IIf(IsNull(rs!vcha_gre_grupo_real_id), "", rs!vcha_gre_grupo_real_id)
         var_cliente = txt_clave_cliente
         var_titular = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
         var_clave_moneda = IIf(IsNull(rs!VCHA_MON_MONEDA_ID), "", rs!VCHA_MON_MONEDA_ID)
         var_agente = IIf(IsNull(rs!vcha_age_agente_id), "", rs!vcha_age_agente_id)
         var_plazo = 0
         txt_plazo = var_plazo
         var_iva = IIf(IsNull(rs!FLOA_TPE_IVA), 0, rs!FLOA_TPE_IVA)
         var_clave_moneda = IIf(IsNull(rs!VCHA_MON_MONEDA_ID), "", rs!VCHA_MON_MONEDA_ID)
         lbl_moneda = IIf(IsNull(rs!vcha_mon_nombre_plural), "", rs!vcha_mon_nombre_plural)
      Else
         MsgBox "Clave de cliente incorrecta", vbOKOnly, "ATENCION"
         txt_clave_cliente = ""
         txt_nombre_cliente = ""
      End If
      rs.Close
   Else
      txt_nombre_cliente = ""
   End If
End Sub

Private Sub txt_folio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   End If
End Sub

Private Sub txt_folio_LostFocus()
   If Not IsNumeric(txt_folio) Then
      MsgBox "Folio Incorrecto", vbOKOnly, "ATENCION"
      txt_folio = ""
   End If
End Sub

Private Sub txt_importe_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If Me.cmb_series.Enabled = False Then
         Me.cmd_imprimir.SetFocus
      Else
         Me.cmb_series.SetFocus
      End If
   End If
End Sub

Private Sub txt_importe_LostFocus()
   If Not IsNumeric(txt_importe) Then
      MsgBox "Importe Incorrecto", vbOKOnly, "ATENCION"
      txt_importe = 0
   End If
End Sub

Private Sub txt_nombre_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select DISTINCT VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE from vw_clientes where vcha_emp_empresa_id = '" + var_empresa + "' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
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

Private Sub txt_nombre_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_plazo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
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


