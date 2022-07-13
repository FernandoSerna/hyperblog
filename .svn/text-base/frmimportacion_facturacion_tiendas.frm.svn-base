VERSION 5.00
Begin VB.Form frmimportacion_facturacion_tiendas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importación de facturación de tiendas"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   " Fecha a Importar "
      Height          =   825
      Left            =   75
      TabIndex        =   3
      Top             =   465
      Width           =   3720
      Begin VB.TextBox txt_fecha 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1230
         TabIndex        =   4
         Top             =   315
         Width           =   1380
      End
   End
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   0
      TabIndex        =   2
      Top             =   315
      Width           =   3840
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3465
      Picture         =   "frmimportacion_facturacion_tiendas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      Picture         =   "frmimportacion_facturacion_tiendas.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Aceptar"
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "frmimportacion_facturacion_tiendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_Click()
   var_dia = CStr(Day(CDate(Me.txt_fecha)))
   var_mes = CStr(Month(CDate(Me.txt_fecha)))
   var_año = CStr(Year(CDate(Me.txt_fecha)))
   If Len(Trim(var_dia)) = 1 Then
      var_dia = "0" + var_dia
   End If
   If Len(Trim(var_mes)) = 1 Then
      var_mes = "0" + var_mes
   End If
   var_fecha = var_dia + "/" + var_mes + "/" + var_año
   Text1 = "select vcha_fac_cliente_sid, date_fac_fecha, vcha_fac_serie, vcha_fac_importe, numb_fac_factura_id  from tb_facturas where numb_fac_credito = 1 and date_fac_fecha = '" + var_fecha + "'"
   rs.Open "select vcha_fac_cliente_sid, date_fac_fecha, vcha_fac_serie, vcha_fac_importe, numb_fac_factura_id,vcha_fac_status  from tb_facturas where numb_fac_credito = 1 and date_fac_fecha = to_date('" + var_fecha + "','DD/MM/YYYY')", cnnsorteo, adOpenDynamic, adLockOptimistic
   var_contador_facturas = 0
   While Not rs.EOF
         rsaux.Open "SELECT * FROM TB_ENCABEZADO_CARTERA WHERE VCHA_EMP_EMPRESA_ID = '02' AND VCHA_cAR_DOCUMENTO = 'FA' AND VCHA_SER_SERIE_ID = '" + rs!vcha_fac_serie + "' AND INTE_CAR_NUMERO = " + CStr(rs!numb_fac_factura_id), cnn, adOpenDynamic, adLockOptimistic
         If rsaux.EOF Then
            If rsaux2.State = 1 Then
               rsaux2.Close
            End If
            var_cliente = Trim(rs!vcha_fac_cliente_sid)
            If rsaux2.State = 1 Then
               rsaux2.Close
            End If
            rsaux2.Open "select vcha_cli_clave_id from vw_clientes where vcha_Cli_clave_id = '" + var_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
            var_cliente = IIf(IsNull(rsaux2!vcha_cli_clave_id), "", rsaux2!vcha_cli_clave_id)
            rsaux2.Close
            rsaux2.Open "SELECT * FROM VW_CLIENTES WHERE VCHA_CLI_CLAVE_ID = '" + Trim(var_cliente) + "'", cnn, adOpenDynamic, adLockOptimistic
                       
            txt_nombre_cliente = IIf(IsNull(rsaux2!VCHA_CLI_NOMBRE), "", rsaux2!VCHA_CLI_NOMBRE)
            txt_plazo = IIf(IsNull(rsaux2!inte_pla_dias), 0, rsaux2!inte_pla_dias)
            var_grupo_actual = IIf(IsNull(rsaux2!vcha_gac_grupo_Actual_id), "", rsaux2!vcha_gac_grupo_Actual_id)
            var_grupo_real = IIf(IsNull(rsaux2!vcha_gre_grupo_real_id), "", rsaux2!vcha_gre_grupo_real_id)
            var_titular = IIf(IsNull(rsaux2!vcha_tit_titular_id), "", rsaux2!vcha_tit_titular_id)
            var_clave_moneda = IIf(IsNull(rsaux2!vcha_mon_moneda_id), "", rsaux2!vcha_mon_moneda_id)
            var_agente = IIf(IsNull(rsaux2!vcha_age_agente_id), "", rsaux2!vcha_age_agente_id)
            var_plazo = IIf(IsNull(rsaux2!inte_pla_dias), 0, rsaux2!inte_pla_dias)
            txt_plazo = var_plazo
            var_iva = IIf(IsNull(rsaux2!FLOA_TPE_IVA), 0, rsaux2!FLOA_TPE_IVA)
            var_clave_moneda = IIf(IsNull(rsaux2!vcha_mon_moneda_id), "", rsaux2!vcha_mon_moneda_id)
            var_tipo_Cambio = 1
                        
            var_numero_folio = rs!numb_fac_factura_id
            var_estatus = IIf(IsNull(rs!vcha_fac_status), "", rs!vcha_fac_status)
            var_importe_total = (rs!vcha_fac_importe / (1 + (var_iva / 100))) * var_tipo_Cambio
            var_importe_neto = rs!vcha_fac_importe * var_tipo_Cambio
            var_importe_iva = var_importe_neto - var_importe_total
            var_subimporte = var_importe_total
            var_insertar = False
            var_serie = rs!vcha_fac_serie
            var_tipo_comision = 0
            
            cnn.BeginTrans
            var_cadena = "INSERT INTO TB_ENCABEZADO_CARTERA (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_CAR_CLASE_ID, INTE_CAR_NUMERO, CHAR_CAR_AFECTACION, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO, DTIM_CAR_FECHA, VCHA_AGE_AGENTE_ID, VCHA_GAC_GRUPO_ACTUAL_ID, VCHA_GRE_GRUPO_REAL_ID, VCHA_TIT_TITULAR_ID, VCHA_CLI_CLAVE_ID, VCHA_ESB_ESTABLECIMIENTO_ID, INTE_CAR_PLAZO, FLOA_CAR_PORCENTAJE_IVA, FLOA_CAR_PORCENTAJE_IMPUESTO_1, FLOA_CAR_PORCENTAJE_IMPUESTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_1, "
            var_cadena = var_cadena + " FLOA_CAR_PORCENTAJE_DESCUENTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_3, FLOA_CAR_IMPORTE_TOTAL, FLOA_CAR_IMPORTE_IVA, FLOA_CAR_IMPORTE_IMPUESTO_1, FLOA_CAR_IMPORTE_IMPUESTO_2, FLOA_CAR_IMPORTE_DESCUENTO_1, FLOA_CAR_IMPORTE_DESCUENTO_2, FLOA_CAR_IMPORTE_DESCUENTO_3, FLOA_CAR_SUBIMPORTE, FLOA_CAR_IMPORTE_NETO, VCHA_CAR_IMPORTE_LETRA, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA, VCHA_AUD_FECHA, FLOA_CAR_SALDO, DTIM_CAR_FECHA_VENCIMIENTO, DTIM_CAR_FECHA_ENTREGA, VCHA_MON_MONEDA_ID, FLOA_CAR_TIPO_CAMBIO, VCHA_SER_SERIE_ID, CHAR_CAR_ESTATUS, INTE_cAR_TIPO_COMISION) Values "
            var_cadena = var_cadena + " ('02', '12', 'FA', 'FA', 'FA', " + CStr(var_numero_folio) + ", '+', '', '', 0, '" + Me.txt_fecha + "', '" + var_agente + "', '" + var_grupo_actual + "', '" + var_grupo_real + "', '" + var_titular + "', '" + Trim(var_cliente) + "', '', " + CStr(var_plazo) + ", " + CStr(var_iva) + ", 0, 0, 0, 0, 0, " + CStr(var_importe_total) + ", " + CStr(var_importe_iva) + ", 0, 0, 0, 0, 0, " + CStr(var_subimporte) + ", " + CStr(var_importe_neto) + ", '', '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', GETDate(), 0, GETDate(), GETDate(), '" + var_clave_moneda + "', " + CStr(var_tipo_Cambio) + ", '" + var_serie + "', '" + var_estatus + "'," + CStr(var_tipo_comision) + ")"
            
            rsaux5.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            Set TB_ESTADO_CUENTA_INSERTA = New TB_ESTADO_CUENTA_INSERTA
            var_inserta = TB_ESTADO_CUENTA_INSERTA.Anadir("02", CStr(Trim(var_serie)), "FA", CDbl(var_numero_folio), "", "", 0, CDbl(var_importe_neto), 0)
            cnn.CommitTrans
            
            var_contador_facturas = var_contador_facturas + 1
         Else
            If var_estatus = "C" Then
               'rsaux10.Open "update tb_encabezado_Cartera set char_car_estatus = '" + var_estatus + "' where vcha_emp_empresa_id = '12' and vcha_car_documento = 'FA' and vcha_ser_serie_id = '" + var_serie + "' and inte_Car_numero  = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
            End If
         End If
         rsaux.Close
         rs.MoveNext
   Wend
   rs.Close
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   MsgBox "Se insertaron " + CStr(var_contador_facturas) + " facturas", vbOKOnly, "ATENCION"
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 3850
   Me.txt_fecha = Date
   If cnnsorteo.State = 0 Then
      cnnsorteo.Open var_conexion_sorteo
      cnnsorteo.CursorLocation = adUseClient
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub txt_fecha_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fecha) Then
         frmcalendario.mes.Value = CDate(Me.txt_fecha)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_fecha = var_fecha_general
   End If
End Sub

Private Sub txt_fecha_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_aceptar.SetFocus
   Else
      If KeyAscii = 27 Then
         Unload Me
      Else
         KeyAscii = 0
      End If
   End If
End Sub
