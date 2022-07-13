VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmfacturas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturación"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8115
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   1965
      Left            =   1335
      TabIndex        =   19
      Top             =   45
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1410
         Left            =   30
         TabIndex        =   20
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
         TabIndex        =   21
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Datos Generales "
      Height          =   1650
      Left            =   105
      TabIndex        =   11
      Top             =   450
      Width           =   7920
      Begin VB.TextBox txt_nombre_tipo_comision 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5595
         TabIndex        =   10
         Top             =   1260
         Width           =   2220
      End
      Begin VB.TextBox txt_tipo_comision 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5055
         TabIndex        =   9
         Top             =   1260
         Width           =   525
      End
      Begin VB.TextBox txt_serie 
         Height          =   315
         Left            =   885
         TabIndex        =   7
         Top             =   1260
         Width           =   795
      End
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   315
         Left            =   2640
         TabIndex        =   4
         Top             =   240
         Width           =   5190
      End
      Begin VB.TextBox txt_folio 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2115
         TabIndex        =   8
         Top             =   1260
         Width           =   1725
      End
      Begin VB.TextBox txt_clave_cliente 
         Height          =   315
         Left            =   885
         TabIndex        =   3
         Top             =   240
         Width           =   1725
      End
      Begin VB.TextBox txt_plazo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   885
         TabIndex        =   5
         Top             =   585
         Width           =   1725
      End
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   885
         TabIndex        =   6
         Top             =   930
         Width           =   1725
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Comision:"
         Height          =   195
         Left            =   3960
         TabIndex        =   22
         Top             =   1320
         Width           =   1035
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   165
         TabIndex        =   18
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Folio:"
         Height          =   195
         Left            =   1725
         TabIndex        =   17
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   165
         TabIndex        =   15
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Plazo:"
         Height          =   195
         Left            =   165
         TabIndex        =   14
         Top             =   645
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Monto:"
         Height          =   195
         Left            =   165
         TabIndex        =   13
         Top             =   990
         Width           =   495
      End
      Begin VB.Label lbl_moneda 
         Height          =   285
         Left            =   2670
         TabIndex        =   12
         Top             =   945
         Width           =   2985
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7650
      Picture         =   "frmfacturas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   450
      Picture         =   "frmfacturas.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmfacturas.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   45
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   90
      TabIndex        =   16
      Top             =   315
      Width           =   7920
   End
End
Attribute VB_Name = "frmfacturas"
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
   



Private Sub cmb_series_Click()
   var_serie = cmb_series
End Sub

Private Sub cmb_series_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
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
   var_posible_comision = True
   rsaux5.Open "SELECT * FROM TB_TIPO_COMISION WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rsaux5.EOF Then
      If Trim(txt_tipo_comision) = "" Then
         If Not IsNumeric(Me.txt_tipo_comision) Then
            var_posible_comision = False
         End If
      End If
   End If
   rsaux5.Close
   If var_posible_comision = True Then
      'rsaux4.Open "select * from tb_Series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_ser_serie_id = '" + txt_serie + "'", cnn, adOpenDynamic, adLockOptimistic
      'If rsaux4.EOF Then
         If Me.txt_serie <> "" Then
            If Trim(Me.txt_clave_cliente) <> "" Then
               If IsNumeric(Me.txt_importe) Then
                  If IsNumeric(Me.txt_folio) Then
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
                           rs.Open "select * from tb_encabezado_cartera where vcha_emp_empresa_id =  '" + var_empresa + "' and vcha_ser_serie_id = '" + txt_serie + "' and vcha_car_documento = 'FA' and inte_Car_numero = " + txt_folio, cnn, adOpenDynamic, adLockOptimistic
                           If rs.EOF Then
                              rs.Close
                              si = MsgBox("¿Deseas cargar la Factura?", vbYesNo, "ATENCION")
                              If si = 6 Then
                                 si = MsgBox("Confirmar la carga de la Factura", vbYesNo, "ATENCION")
                                 If si = 6 Then
                                    'rs.Open "select inte_ser_factura from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                                    'If Not rs.EOF Then
                                    var_numero_folio = txt_folio
                                    var_importe_total = (txt_importe / (1 + (var_iva / 100))) * var_tipo_Cambio
                                    var_importe_neto = txt_importe * var_tipo_Cambio
                                    var_importe_iva = var_importe_neto - var_importe_total
                                    var_subimporte = var_importe_total
                                    var_insertar = False
                                    var_serie = Me.txt_serie
                                    'var_insertar = TB_ENCABEZADO_CARTERA_I.Anadir(var_empresa, var_unidad_organizacional, "FA", "FA", "FA", var_numero_folio, "+", "", "", 0, CStr(Date), var_agente, var_grupo_actual, var_grupo_real, var_titular, txt_clave_cliente, "", var_plazo, var_iva, 0, 0, 0, 0, 0, var_importe_total, var_importe_iva, 0, 0, 0, 0, 0, var_subimporte, var_importe_neto, "", var_clave_usuario_global, fun_NombrePc, Date, 0, Date, Date, var_clave_moneda, var_tipo_Cambio, var_serie, "")
                                    var_tipo_comision = 0
                                    If IsNumeric(txt_tipo_comision) Then
                                       var_tipo_comision = Me.txt_tipo_comision
                                    End If
                                    If var_empresa = "02" Then
                                       Cadena = "INSERT INTO TB_ENCABEZADO_CARTERA (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_CAR_CLASE_ID, INTE_CAR_NUMERO, CHAR_CAR_AFECTACION, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO, DTIM_CAR_FECHA, VCHA_AGE_AGENTE_ID, VCHA_GAC_GRUPO_ACTUAL_ID, VCHA_GRE_GRUPO_REAL_ID, VCHA_TIT_TITULAR_ID, VCHA_CLI_CLAVE_ID, VCHA_ESB_ESTABLECIMIENTO_ID, INTE_CAR_PLAZO, FLOA_CAR_PORCENTAJE_IVA, FLOA_CAR_PORCENTAJE_IMPUESTO_1, FLOA_CAR_PORCENTAJE_IMPUESTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_1, "
                                       Cadena = Cadena + " FLOA_CAR_PORCENTAJE_DESCUENTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_3, FLOA_CAR_IMPORTE_TOTAL, FLOA_CAR_IMPORTE_IVA, FLOA_CAR_IMPORTE_IMPUESTO_1, FLOA_CAR_IMPORTE_IMPUESTO_2, FLOA_CAR_IMPORTE_DESCUENTO_1, FLOA_CAR_IMPORTE_DESCUENTO_2, FLOA_CAR_IMPORTE_DESCUENTO_3, FLOA_CAR_SUBIMPORTE, FLOA_CAR_IMPORTE_NETO, VCHA_CAR_IMPORTE_LETRA, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA, VCHA_AUD_FECHA, FLOA_CAR_SALDO, DTIM_CAR_FECHA_VENCIMIENTO, DTIM_CAR_FECHA_ENTREGA, VCHA_MON_MONEDA_ID, FLOA_CAR_TIPO_CAMBIO, VCHA_SER_SERIE_ID, CHAR_CAR_ESTATUS, INTE_cAR_TIPO_COMISION, INTE_CAR_PEDIDO_CREDITO) Values "
                                       Cadena = Cadena + " ('" + var_empresa + "', '" + var_unidad_organizacional + "', 'FA', 'FA', 'FA', " + CStr(var_numero_folio) + ", '+', '', '', 0, GETDATE(), '" + var_agente + "', '" + var_grupo_actual + "', '" + var_grupo_real + "', '" + var_titular + "', '" + txt_clave_cliente + "', '', " + CStr(var_plazo) + ", " + CStr(var_iva) + ", 0, 0, 0, 0, 0, " + CStr(var_importe_total) + ", " + CStr(var_importe_iva) + ", 0, 0, 0, 0, 0, " + CStr(var_subimporte) + ", " + CStr(var_importe_neto) + ", '', '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', GETDate(), 0, GETDate(), GETDate(), '" + var_clave_moneda + "', " + CStr(var_tipo_Cambio) + ", '" + var_serie + "', ''," + CStr(var_tipo_comision) + ",3)"
                                    Else
                                       Cadena = "INSERT INTO TB_ENCABEZADO_CARTERA (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_CAR_CLASE_ID, INTE_CAR_NUMERO, CHAR_CAR_AFECTACION, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO, DTIM_CAR_FECHA, VCHA_AGE_AGENTE_ID, VCHA_GAC_GRUPO_ACTUAL_ID, VCHA_GRE_GRUPO_REAL_ID, VCHA_TIT_TITULAR_ID, VCHA_CLI_CLAVE_ID, VCHA_ESB_ESTABLECIMIENTO_ID, INTE_CAR_PLAZO, FLOA_CAR_PORCENTAJE_IVA, FLOA_CAR_PORCENTAJE_IMPUESTO_1, FLOA_CAR_PORCENTAJE_IMPUESTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_1, "
                                       Cadena = Cadena + " FLOA_CAR_PORCENTAJE_DESCUENTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_3, FLOA_CAR_IMPORTE_TOTAL, FLOA_CAR_IMPORTE_IVA, FLOA_CAR_IMPORTE_IMPUESTO_1, FLOA_CAR_IMPORTE_IMPUESTO_2, FLOA_CAR_IMPORTE_DESCUENTO_1, FLOA_CAR_IMPORTE_DESCUENTO_2, FLOA_CAR_IMPORTE_DESCUENTO_3, FLOA_CAR_SUBIMPORTE, FLOA_CAR_IMPORTE_NETO, VCHA_CAR_IMPORTE_LETRA, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA, VCHA_AUD_FECHA, FLOA_CAR_SALDO, DTIM_CAR_FECHA_VENCIMIENTO, DTIM_CAR_FECHA_ENTREGA, VCHA_MON_MONEDA_ID, FLOA_CAR_TIPO_CAMBIO, VCHA_SER_SERIE_ID, CHAR_CAR_ESTATUS, INTE_cAR_TIPO_COMISION, INTE_CAR_PEDIDO_CREDITO) Values "
                                       Cadena = Cadena + " ('" + var_empresa + "', '" + var_unidad_organizacional + "', 'FA', 'FA', 'FA', " + CStr(var_numero_folio) + ", '+', '', '', 0, GETDATE(), '" + var_agente + "', '" + var_grupo_actual + "', '" + var_grupo_real + "', '" + var_titular + "', '" + txt_clave_cliente + "', '', " + CStr(var_plazo) + ", " + CStr(var_iva) + ", 0, 0, 0, 0, 0, " + CStr(var_importe_total) + ", " + CStr(var_importe_iva) + ", 0, 0, 0, 0, 0, " + CStr(var_subimporte) + ", " + CStr(var_importe_neto) + ", '', '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', GETDate(), 0, GETDate(), GETDate(), '" + var_clave_moneda + "', " + CStr(var_tipo_Cambio) + ", '" + var_serie + "', ''," + CStr(var_tipo_comision) + ",0)"
                                    End If
                                    rsaux5.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                    var_inserta = TB_ESTADO_CUENTA_INSERTA.Anadir(var_empresa, var_serie, "FA", var_numero_folio, "", "", 0, var_importe_neto, 0)
                                 End If
                              End If
                           Else
                              rs.Close
                              MsgBox "Ya existe la factura " + txt_folio, vbOKOnly, "ATENCION"
                           End If
                        Else
                           MsgBox "No se a asignado el tipo de cambio del dia de hoy", vbOKOnly, "ATENCION"
                        End If
                     Else
                        MsgBox "El cliente no tiene una moneda asociada", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "Folio incorrecto", vbOKOnly, "ATENCION"
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
      'Else
      '    MsgBox "La serie no debe de ser " + txt_serie, vbOKOnly, "ATENCION"
      'End If
   Else
      MsgBox "Se debe de indicar una clave de comisión", vbOKOnly, "ATENCION"
      Me.txt_tipo_comision = ""
      Me.txt_nombre_tipo_comision = ""
      Me.txt_tipo_comision.SetFocus
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
   Call activa_forma(var_activa_forma_facturas)
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
            Me.txt_tipo_comision = lv_lista.selectedItem
            Me.txt_nombre_tipo_comision = lv_lista.selectedItem.SubItems(1)
            If Me.txt_tipo_comision.Enabled = True Then
               Me.txt_tipo_comision.SetFocus
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

Private Sub txt_clave_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select DISTINCT VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE from VW_clientes WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_cli_nombre), "", rs!VCHA_cli_nombre)
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
      rs.Open "SELECT * FROM VW_CLIENTES WHERE VCHA_CLI_CLAVE_ID =  '" + txt_clave_cliente + "' and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         
         txt_nombre_cliente = IIf(IsNull(rs!VCHA_cli_nombre), "", rs!VCHA_cli_nombre)
         txt_plazo = IIf(IsNull(rs!inte_pla_dias), 0, rs!inte_pla_dias)
         var_grupo_actual = IIf(IsNull(rs!vcha_gac_grupo_Actual_id), "", rs!vcha_gac_grupo_Actual_id)
         var_grupo_real = IIf(IsNull(rs!vcha_gre_grupo_real_id), "", rs!vcha_gre_grupo_real_id)
         var_cliente = txt_clave_cliente
         var_titular = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
         var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
         var_agente = IIf(IsNull(rs!vcha_age_agente_id), "", rs!vcha_age_agente_id)
         var_plazo = IIf(IsNull(rs!inte_pla_dias), 0, rs!inte_pla_dias)
         txt_plazo = var_plazo
         var_iva = IIf(IsNull(rs!FLOA_TPE_IVA), 0, rs!FLOA_TPE_IVA)
         var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
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

Private Sub txt_folio_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
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
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_importe_LostFocus()
   If Not IsNumeric(txt_importe) Then
      MsgBox "Importe Incorrecto", vbOKOnly, "ATENCION"
      txt_importe = 0
   End If
End Sub

Private Sub txt_nombre_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frm_lista.Left = 1530
      frm_lista.Top = 2280
      lv_lista.ListItems.Clear
      rs.Open "select DISTINCT VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE from VW_clientes VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_cli_nombre), "", rs!VCHA_cli_nombre)
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
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_tipo_comision_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub txt_tipo_comision_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_tipo_comision_LostFocus()
   If Trim(txt_tipo_comision) <> "" Then
      rs.Open "SELECT * FROM TB_TIPO_COMISION WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_TCO_TIPO_COMISION_ID = " + txt_tipo_comision, cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_tipo_comision = rs!VCHA_tCO_NOMBRE
      Else
         MsgBox "Clave de comisión incorrecta", vbOKOnly, "ATENCION"
         Me.txt_tipo_comision = ""
         Me.txt_nombre_tipo_comision = ""
      End If
      rs.Close
   End If
End Sub
