VERSION 5.00
Begin VB.Form frmimportacion_informacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importación y Exportación de Información"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   720
      Left            =   90
      TabIndex        =   2
      Top             =   105
      Width           =   4335
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   2700
         TabIndex        =   4
         Top             =   255
         Width           =   1095
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   690
         TabIndex        =   3
         Top             =   255
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2385
         TabIndex        =   6
         Top             =   315
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   315
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   75
      TabIndex        =   0
      Top             =   810
      Width           =   4350
      Begin VB.CommandButton cmd_importacion_cartera 
         Caption         =   "Importación y Exportación de Cartera"
         Height          =   570
         Left            =   60
         TabIndex        =   1
         Top             =   165
         Width           =   4200
      End
   End
End
Attribute VB_Name = "frmimportacion_informacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_importacion_cartera_Click()
   If IsDate(Me.txt_inicio) Then
      If IsDate(Me.txt_fin) Then
         If CDate(Me.txt_inicio) <= CDate(Me.txt_fin) Then
            cnn.CommandTimeout = 360000
            inicio = CStr(Now)
            var_fecha_fin_1 = CDate(txt_fin) + 1
            var_dia = CStr(Day(CDate(txt_inicio)))
            var_mes = CStr(Month(CDate(txt_inicio)))
            var_año = CStr(Year(CDate(txt_inicio)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
             
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            'var_empresa = "18"
            
            
            If var_empresa = "16" Then
               rsaux10.Open "select * from tb_encabezado_Cartera where vcha_emp_empresa_id = '" + var_empresa + "'  and vcha_car_documento = 'FA' and dtim_CAR_FECHA >= " + var_fecha_inicio + " AND DTIM_CAR_FECHA <= " + var_fecha_fin + "-.000000001", cnn, adOpenDynamic, adLockOptimistic
               While Not rsaux10.EOF
         
                     'MsgBox "select * from distribucion.vianney.dbo.tb_encabezado_cartera where vcha_Emp_Empresa_id = '" + rsaux10!vcha_Emp_Empresa_id + "' and vcha_car_documento = '" + rsaux10!vcha_car_documento + "' and vcha_ser_serie_id = '" + rsaux10!vcha_ser_serie_id + "' and inte_car_numero = " + CStr(rsaux10!inte_car_numero)
                     rsaux9.Open "select * from tb_encabezado_cartera where vcha_Emp_Empresa_id = '" + rsaux10!VCHA_EMP_EMPRESA_ID + "' and vcha_car_documento = '" + rsaux10!vcha_Car_documento + "' and vcha_ser_serie_id = '" + rsaux10!vcha_ser_Serie_id + "' and inte_car_numero = " + CStr(rsaux10!inte_car_numero), cnn_distribucion, adOpenDynamic, adLockOptimistic
                     If rsaux9.EOF Then
                        var_dia = CStr(Day(rsaux10!dtim_Car_fecha))
                        var_mes = CStr(Month(rsaux10!dtim_Car_fecha))
                        var_año = CStr(Year(rsaux10!dtim_Car_fecha))
                        If Len(Trim(var_dia)) = 1 Then
                           var_dia = "0" + var_dia
                        End If
                        If Len(Trim(var_mes)) = 1 Then
                           var_mes = "0" + var_mes
                        End If
                        var_fecha_factura = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                        
                        
                        var_cadena = "insert into tb_encabezado_Cartera (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_SER_SERIE_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_CAR_CLASE_ID, INTE_CAR_NUMERO, CHAR_CAR_AFECTACION, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID,"
                        var_cadena = var_cadena + "INTE_EMO_NUMERO, DTIM_CAR_FECHA, VCHA_AGE_AGENTE_ID, VCHA_GAC_GRUPO_ACTUAL_ID, VCHA_GRE_GRUPO_REAL_ID, VCHA_TIT_TITULAR_ID, VCHA_CLI_CLAVE_ID, VCHA_ESB_ESTABLECIMIENTO_ID, INTE_CAR_PLAZO, FLOA_CAR_PORCENTAJE_IVA, FLOA_CAR_PORCENTAJE_IMPUESTO_1, FLOA_CAR_PORCENTAJE_IMPUESTO_2,"
                        var_cadena = var_cadena + " FLOA_CAR_PORCENTAJE_DESCUENTO_1, FLOA_CAR_PORCENTAJE_DESCUENTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_3, FLOA_CAR_IMPORTE_TOTAL, FLOA_CAR_IMPORTE_IVA, FLOA_CAR_IMPORTE_IMPUESTO_1, FLOA_CAR_IMPORTE_IMPUESTO_2, FLOA_CAR_IMPORTE_DESCUENTO_1, FLOA_CAR_IMPORTE_DESCUENTO_2, FLOA_CAR_IMPORTE_DESCUENTO_3,"
                        var_cadena = var_cadena + "FLOA_CAR_SUBIMPORTE, FLOA_CAR_IMPORTE_NETO, VCHA_CAR_IMPORTE_LETRA, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA, VCHA_AUD_FECHA, FLOA_CAR_SALDO, DTIM_CAR_FECHA_VENCIMIENTO, DTIM_CAR_FECHA_ENTREGA, VCHA_MON_MONEDA_ID, FLOA_CAR_TIPO_CAMBIO, CHAR_CAR_ESTATUS, DTIM_CAR_FECHA_CANCELACION, VCHA_CAR_USUARIO_CANCELACION,"
                        var_cadena = var_cadena + "VCHA_CAR_MAQUINA_CANCELACION, CHAR_CAR_TIPO_FACTURACION, INTE_CAR_FACTURA_CEROS, FLOA_CAR_COSTO, VCHA_CAR_CHEQUE, VCHA_CAR_REFERENCIA, INTE_CAR_CAPTURADO, NUM_INTER_TRANC_TYPE, NUM_INTER_UPLOADED, DATE_INTER_DATE, VCHA_CAR_CHEQUE_DEPOSITO, VCHA_CAR_BANCO_CHEQUE, VCHA_CAR_DEPOSITO, VCHA_CAR_BANCO_DEPOSITO,"
                        var_cadena = var_cadena + "DTIM_CAR_FECHA_DEPOSITO, INTE_CAR_TIPO_COMISION) values "
                        
                        var_cadena = var_cadena + "('" + rsaux10!VCHA_EMP_EMPRESA_ID + "' , '" + rsaux10!VCHA_UOR_UNIDAD_ID + "', '" + rsaux10!vcha_ser_Serie_id + "', '" + rsaux10!vcha_Car_tipo_documento + "', '" + rsaux10!vcha_Car_documento + "', '" + rsaux10!vcha_Car_clase_id + "', " + CStr(rsaux10!inte_car_numero) + ", '" + rsaux10!char_car_afectacion + "', '" + rsaux10!VCHA_ALM_ALMACEN_ID + "', '" + rsaux10!VCHA_MOV_MOVIMIENTO_ID + "', "
                        var_cadena = var_cadena + CStr(rsaux10!INTE_EMO_NUMERO) + ", " + var_fecha_factura + ", '" + rsaux10!VCHA_AGE_AGENTE_ID + "', '" + rsaux10!VCHA_GAC_GRUPO_aCTUAL_ID + "', '" + rsaux10!vcha_gre_grupo_real_id + "', '" + rsaux10!vcha_tit_titular_id + "', '" + rsaux10!vcha_cli_clave_id + "', '" + rsaux10!vcha_ESB_ESTABLECIMIENTO_id + "', " + CStr(rsaux10!INTE_CAR_PLAZO) + ", " + CStr(rsaux10!floa_car_porcentaje_iva) + ", " + CStr(rsaux10!FLOA_CAR_PORCENTAJE_IMPUESTO_1) + ", " + CStr(rsaux10!FLOA_CAR_PORCENTAJE_IMPUESTO_2) + ","
                        var_cadena = var_cadena + CStr(rsaux10!FLOA_CAR_PORCENTAJE_DESCUENTO_1) + ", " + CStr(rsaux10!FLOA_CAR_PORCENTAJE_DESCUENTO_2) + ", " + CStr(rsaux10!floa_car_porcentaje_descuento_3) + ", " + CStr(rsaux10!FLOA_CAR_IMPORTE_TOTAL) + ", " + CStr(rsaux10!floa_car_importe_iva) + "," + CStr(rsaux10!FLOA_CAR_IMPORTE_IMPUESTO_1) + ", " + CStr(rsaux10!FLOA_CAR_IMPORTE_IMPUESTO_2) + ", " + CStr(rsaux10!FLOA_CAR_IMPORTE_DESCUENTO_1) + ", " + CStr(rsaux10!FLOA_CAR_IMPORTE_DESCUENTO_2) + ", " + CStr(rsaux10!FLOA_CAR_IMPORTE_DESCUENTO_3) + ","
                        var_cadena = var_cadena + CStr(rsaux10!floa_car_subimporte) + ", " + CStr(rsaux10!floa_Car_importe_neto) + ", '" + CStr(rsaux10!vcha_car_importe_letra) + "', '" + CStr(rsaux10!vcha_aud_usuario) + "', '" + CStr(rsaux10!vcha_aud_maquina) + "', getdate(), 0, getdate(), getdate(), '" + (rsaux10!vcha_mon_moneda_id) + "', " + CStr(rsaux10!floa_car_tipo_cambio) + ", '" + IIf(IsNull(rsaux10!CHAR_CAR_ESTATUS), "", rsaux10!CHAR_CAR_ESTATUS) + "', null, '',"
                        var_cadena = var_cadena + "'', '" + rsaux10!char_Car_tipo_facturacion + "', 0, 0, '', '', 0, 0, 0, getdate(), '', '', '', '',"
                        var_cadena = var_cadena + "null, 1)"
                        
                        'var_cadena2 = "('" + rsaux10!VCHA_EMP_EMPRESA_ID + "' , '" + rsaux10!VCHA_UOR_UNIDAD_ID + "', '" + rsaux10!vcha_ser_serie_id + "', '" + rsaux10!VCHA_CAR_TIPO_DOCUMENTO + "', '" + rsaux10!vcha_Car_documento + "', '" + rsaux10!vcha_Car_clase_id + "', " + CStr(rsaux10!inte_car_numero) + ", '" + rsaux10!char_car_afectacion + "', '" + rsaux10!VCHA_ALM_ALMACEN_ID + "', '" + rsaux10!VCHA_MOV_MOVIMIENTO_ID + "', "
                        'var_cadena2 = var_cadena2 + CStr(rsaux10!INTE_EMO_NUMERO) + ", " + CStr(rsaux10!DTIM_car_FECHA) + ", '" + rsaux10!vcha_age_agente_id + "', '" + rsaux10!vcha_gac_grupo_Actual_id + "', '" + rsaux10!vcha_gre_grupo_real_id + "', '" + rsaux10!vcha_tit_titular_id + "', '" + rsaux10!vcha_cli_clave_id + "', '" + rsaux10!vcha_esb_establecimiento_id + "', " + CStr(rsaux10!INTE_CAR_PLAZO) + ", " + CStr(rsaux10!floa_car_porcentaje_iva) + ", " + CStr(rsaux10!FLOA_CAR_PORCENTAJE_IMPUESTO_1) + ", " + CStr(rsaux10!FLOA_CAR_PORCENTAJE_IMPUESTO_2) + ","
                        'var_cadena2 = var_cadena2 + CStr(rsaux10!FLOA_CAR_PORCENTAJE_DESCUENTO_1) + ", " + CStr(rsaux10!FLOA_CAR_PORCENTAJE_DESCUENTO_2) + ", " + CStr(rsaux10!floa_car_porcentaje_descuento_3) + ", " + CStr(rsaux10!FLOA_CAR_IMPORTE_TOTAL) + ", " + CStr(rsaux10!floa_car_importe_iva) + "," + CStr(rsaux10!FLOA_CAR_IMPORTE_IMPUESTO_1) + ", " + CStr(rsaux10!FLOA_CAR_IMPORTE_IMPUESTO_2) + ", " + CStr(rsaux10!FLOA_CAR_IMPORTE_DESCUENTO_1) + ", " + CStr(rsaux10!FLOA_CAR_IMPORTE_DESCUENTO_2) + ", " + CStr(rsaux10!FLOA_CAR_IMPORTE_DESCUENTO_3) + ","
                        'var_cadena2 = var_cadena2 + CStr(rsaux10!floa_car_subimporte) + ", " + CStr(rsaux10!floa_car_importe_neto) + ", '" + CStr(rsaux10!vcha_car_importe_letra) + "', '" + CStr(rsaux10!vcha_aud_usuario) + "', '" + CStr(rsaux10!vcha_aud_maquina) + "', getdate(), 0, getdate(), getdate(), '" + (rsaux10!vcha_mon_moneda_id) + "', " + CStr(rsaux10!FLOA_cAR_TIPO_cAMBIO) + ", '" + IIf(IsNull(rsaux10!CHAR_CAR_ESTATUS), "", rsaux10!CHAR_CAR_ESTATUS) + "', null, '',"
                        'var_cadena2 = var_cadena2 + "'', '" + rsaux10!char_Car_tipo_facturacion + "', 0, 0, '', '', 0, 0, 0, getdate(), '', '', '', '',"
                        'var_cadena2 = var_cadena2 + "null, 1)"
                        
                        'MsgBox var_cadena2
                        
                        rsaux8.Open var_cadena, cnn_distribucion, adOpenDynamic, adLockOptimistic
                        rsaux8.Open " insert tb_Estado_cuenta (vcha_emp_empresa_id, vcha_ecu_movimiento_cargo, vcha_ecu_serie_Cargo, inte_ecu_numero_cargo, floa_ecu_importe_cargo) values ('" + rsaux10!VCHA_EMP_EMPRESA_ID + "','" + rsaux10!vcha_Car_documento + "','" + rsaux10!vcha_ser_Serie_id + "'," + CStr(rsaux10!inte_car_numero) + "," + CStr(rsaux10!floa_Car_importe_neto) + ")", cnn_distribucion, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux9.Close
                     rsaux10.MoveNext
               Wend
               rsaux10.Close
               MsgBox "Se a terminado de exportar la información", vbOKOnly, "ATENCION"
            End If
            
            
            
            
            If var_empresa = "06" Then
               rsaux10.Open "select * from tb_encabezado_Cartera where vcha_emp_empresa_id = '" + var_empresa + "'  and vcha_car_documento = 'FA' and dtim_CAR_FECHA >= " + var_fecha_inicio + " AND DTIM_CAR_FECHA <= " + var_fecha_fin + "-.000000001", cnn, adOpenDynamic, adLockOptimistic
               While Not rsaux10.EOF
                     'MsgBox "select * from distribucion.vianney.dbo.tb_encabezado_cartera where vcha_Emp_Empresa_id = '" + rsaux10!vcha_Emp_Empresa_id + "' and vcha_car_documento = '" + rsaux10!vcha_car_documento + "' and vcha_ser_serie_id = '" + rsaux10!vcha_ser_serie_id + "' and inte_car_numero = " + CStr(rsaux10!inte_car_numero)
                     rsaux9.Open "select * from tb_encabezado_cartera where vcha_Emp_Empresa_id = '" + rsaux10!VCHA_EMP_EMPRESA_ID + "' and vcha_car_documento = '" + rsaux10!vcha_Car_documento + "' and vcha_ser_serie_id = '" + rsaux10!vcha_ser_Serie_id + "' and inte_car_numero = " + CStr(rsaux10!inte_car_numero), cnn_distribucion, adOpenDynamic, adLockOptimistic
                     If rsaux9.EOF Then
                        var_dia = CStr(Day(rsaux10!dtim_Car_fecha))
                        var_mes = CStr(Month(rsaux10!dtim_Car_fecha))
                        var_año = CStr(Year(rsaux10!dtim_Car_fecha))
                        If Len(Trim(var_dia)) = 1 Then
                           var_dia = "0" + var_dia
                        End If
                        If Len(Trim(var_mes)) = 1 Then
                           var_mes = "0" + var_mes
                        End If
                        var_fecha_factura = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                        
                        
                        var_cadena = "insert into distribucion.vianney.dbo.tb_encabezado_Cartera (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_SER_SERIE_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_CAR_CLASE_ID, INTE_CAR_NUMERO, CHAR_CAR_AFECTACION, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID,"
                        var_cadena = var_cadena + "INTE_EMO_NUMERO, DTIM_CAR_FECHA, VCHA_AGE_AGENTE_ID, VCHA_GAC_GRUPO_ACTUAL_ID, VCHA_GRE_GRUPO_REAL_ID, VCHA_TIT_TITULAR_ID, VCHA_CLI_CLAVE_ID, VCHA_ESB_ESTABLECIMIENTO_ID, INTE_CAR_PLAZO, FLOA_CAR_PORCENTAJE_IVA, FLOA_CAR_PORCENTAJE_IMPUESTO_1, FLOA_CAR_PORCENTAJE_IMPUESTO_2,"
                        var_cadena = var_cadena + " FLOA_CAR_PORCENTAJE_DESCUENTO_1, FLOA_CAR_PORCENTAJE_DESCUENTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_3, FLOA_CAR_IMPORTE_TOTAL, FLOA_CAR_IMPORTE_IVA, FLOA_CAR_IMPORTE_IMPUESTO_1, FLOA_CAR_IMPORTE_IMPUESTO_2, FLOA_CAR_IMPORTE_DESCUENTO_1, FLOA_CAR_IMPORTE_DESCUENTO_2, FLOA_CAR_IMPORTE_DESCUENTO_3,"
                        var_cadena = var_cadena + "FLOA_CAR_SUBIMPORTE, FLOA_CAR_IMPORTE_NETO, VCHA_CAR_IMPORTE_LETRA, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA, VCHA_AUD_FECHA, FLOA_CAR_SALDO, DTIM_CAR_FECHA_VENCIMIENTO, DTIM_CAR_FECHA_ENTREGA, VCHA_MON_MONEDA_ID, FLOA_CAR_TIPO_CAMBIO, CHAR_CAR_ESTATUS, DTIM_CAR_FECHA_CANCELACION, VCHA_CAR_USUARIO_CANCELACION,"
                        var_cadena = var_cadena + "VCHA_CAR_MAQUINA_CANCELACION, CHAR_CAR_TIPO_FACTURACION, INTE_CAR_FACTURA_CEROS, FLOA_CAR_COSTO, VCHA_CAR_CHEQUE, VCHA_CAR_REFERENCIA, INTE_CAR_CAPTURADO, NUM_INTER_TRANC_TYPE, NUM_INTER_UPLOADED, DATE_INTER_DATE, VCHA_CAR_CHEQUE_DEPOSITO, VCHA_CAR_BANCO_CHEQUE, VCHA_CAR_DEPOSITO, VCHA_CAR_BANCO_DEPOSITO,"
                        var_cadena = var_cadena + "DTIM_CAR_FECHA_DEPOSITO, INTE_CAR_TIPO_COMISION) values "
                        
                        var_cadena = var_cadena + "('" + rsaux10!VCHA_EMP_EMPRESA_ID + "' , '" + rsaux10!VCHA_UOR_UNIDAD_ID + "', '" + rsaux10!vcha_ser_Serie_id + "', '" + rsaux10!vcha_Car_tipo_documento + "', '" + rsaux10!vcha_Car_documento + "', '" + rsaux10!vcha_Car_clase_id + "', " + CStr(rsaux10!inte_car_numero) + ", '" + rsaux10!char_car_afectacion + "', '" + rsaux10!VCHA_ALM_ALMACEN_ID + "', '" + rsaux10!VCHA_MOV_MOVIMIENTO_ID + "', "
                        var_cadena = var_cadena + CStr(rsaux10!INTE_EMO_NUMERO) + ", " + var_fecha_factura + ", '" + rsaux10!VCHA_AGE_AGENTE_ID + "', '" + IIf(IsNull(rsaux10!VCHA_GAC_GRUPO_aCTUAL_ID), "", rsaux10!VCHA_GAC_GRUPO_aCTUAL_ID) + "', '" + IIf(IsNull(rsaux10!vcha_gre_grupo_real_id), "", rsaux10!vcha_gre_grupo_real_id) + "', '" + rsaux10!vcha_tit_titular_id + "', '" + rsaux10!vcha_cli_clave_id + "', '" + rsaux10!vcha_ESB_ESTABLECIMIENTO_id + "', " + CStr(rsaux10!INTE_CAR_PLAZO) + ", " + CStr(rsaux10!floa_car_porcentaje_iva) + ", " + CStr(rsaux10!FLOA_CAR_PORCENTAJE_IMPUESTO_1) + ", " + CStr(rsaux10!FLOA_CAR_PORCENTAJE_IMPUESTO_2) + ","
                        var_cadena = var_cadena + CStr(rsaux10!FLOA_CAR_PORCENTAJE_DESCUENTO_1) + ", " + CStr(rsaux10!FLOA_CAR_PORCENTAJE_DESCUENTO_2) + ", " + CStr(rsaux10!floa_car_porcentaje_descuento_3) + ", " + CStr(rsaux10!FLOA_CAR_IMPORTE_TOTAL) + ", " + CStr(rsaux10!floa_car_importe_iva) + "," + CStr(rsaux10!FLOA_CAR_IMPORTE_IMPUESTO_1) + ", " + CStr(rsaux10!FLOA_CAR_IMPORTE_IMPUESTO_2) + ", " + CStr(rsaux10!FLOA_CAR_IMPORTE_DESCUENTO_1) + ", " + CStr(rsaux10!FLOA_CAR_IMPORTE_DESCUENTO_2) + ", " + CStr(rsaux10!FLOA_CAR_IMPORTE_DESCUENTO_3) + ","
                        var_cadena = var_cadena + CStr(rsaux10!floa_car_subimporte) + ", " + CStr(rsaux10!floa_Car_importe_neto) + ", '" + CStr(rsaux10!vcha_car_importe_letra) + "', '" + CStr(rsaux10!vcha_aud_usuario) + "', '" + CStr(rsaux10!vcha_aud_maquina) + "', getdate(), 0, getdate(), getdate(), '" + (rsaux10!vcha_mon_moneda_id) + "', " + CStr(rsaux10!floa_car_tipo_cambio) + ", '" + IIf(IsNull(rsaux10!CHAR_CAR_ESTATUS), "", rsaux10!CHAR_CAR_ESTATUS) + "', null, '',"
                        var_cadena = var_cadena + "'', '" + rsaux10!char_Car_tipo_facturacion + "', 0, 0, '', '', 0, 0, 0, getdate(), '', '', '', '',"
                        var_cadena = var_cadena + "null, 1)"
                        
                        'var_cadena2 = "('" + rsaux10!VCHA_EMP_EMPRESA_ID + "' , '" + rsaux10!VCHA_UOR_UNIDAD_ID + "', '" + rsaux10!vcha_ser_serie_id + "', '" + rsaux10!VCHA_CAR_TIPO_DOCUMENTO + "', '" + rsaux10!vcha_Car_documento + "', '" + rsaux10!vcha_Car_clase_id + "', " + CStr(rsaux10!inte_car_numero) + ", '" + rsaux10!char_car_afectacion + "', '" + rsaux10!VCHA_ALM_ALMACEN_ID + "', '" + rsaux10!VCHA_MOV_MOVIMIENTO_ID + "', "
                        'var_cadena2 = var_cadena2 + CStr(rsaux10!INTE_EMO_NUMERO) + ", " + CStr(rsaux10!DTIM_car_FECHA) + ", '" + rsaux10!vcha_age_agente_id + "', '" + rsaux10!vcha_gac_grupo_Actual_id + "', '" + rsaux10!vcha_gre_grupo_real_id + "', '" + rsaux10!vcha_tit_titular_id + "', '" + rsaux10!vcha_cli_clave_id + "', '" + rsaux10!vcha_esb_establecimiento_id + "', " + CStr(rsaux10!INTE_CAR_PLAZO) + ", " + CStr(rsaux10!floa_car_porcentaje_iva) + ", " + CStr(rsaux10!FLOA_CAR_PORCENTAJE_IMPUESTO_1) + ", " + CStr(rsaux10!FLOA_CAR_PORCENTAJE_IMPUESTO_2) + ","
                        'var_cadena2 = var_cadena2 + CStr(rsaux10!FLOA_CAR_PORCENTAJE_DESCUENTO_1) + ", " + CStr(rsaux10!FLOA_CAR_PORCENTAJE_DESCUENTO_2) + ", " + CStr(rsaux10!floa_car_porcentaje_descuento_3) + ", " + CStr(rsaux10!FLOA_CAR_IMPORTE_TOTAL) + ", " + CStr(rsaux10!floa_car_importe_iva) + "," + CStr(rsaux10!FLOA_CAR_IMPORTE_IMPUESTO_1) + ", " + CStr(rsaux10!FLOA_CAR_IMPORTE_IMPUESTO_2) + ", " + CStr(rsaux10!FLOA_CAR_IMPORTE_DESCUENTO_1) + ", " + CStr(rsaux10!FLOA_CAR_IMPORTE_DESCUENTO_2) + ", " + CStr(rsaux10!FLOA_CAR_IMPORTE_DESCUENTO_3) + ","
                        'var_cadena2 = var_cadena2 + CStr(rsaux10!floa_car_subimporte) + ", " + CStr(rsaux10!floa_car_importe_neto) + ", '" + CStr(rsaux10!vcha_car_importe_letra) + "', '" + CStr(rsaux10!vcha_aud_usuario) + "', '" + CStr(rsaux10!vcha_aud_maquina) + "', getdate(), 0, getdate(), getdate(), '" + (rsaux10!vcha_mon_moneda_id) + "', " + CStr(rsaux10!FLOA_cAR_TIPO_cAMBIO) + ", '" + IIf(IsNull(rsaux10!CHAR_CAR_ESTATUS), "", rsaux10!CHAR_CAR_ESTATUS) + "', null, '',"
                        'var_cadena2 = var_cadena2 + "'', '" + rsaux10!char_Car_tipo_facturacion + "', 0, 0, '', '', 0, 0, 0, getdate(), '', '', '', '',"
                        'var_cadena2 = var_cadena2 + "null, 1)"
                        
                        'MsgBox var_cadena2
                        
                        rsaux8.Open var_cadena, cnn_distribucion, adOpenDynamic, adLockOptimistic
                        rsaux8.Open " insert into tb_Estado_cuenta (vcha_emp_empresa_id, vcha_ecu_movimiento_cargo, vcha_ecu_serie_Cargo, inte_ecu_numero_cargo, floa_ecu_importe_cargo) values ('" + rsaux10!VCHA_EMP_EMPRESA_ID + "','" + rsaux10!vcha_Car_documento + "','" + rsaux10!vcha_ser_Serie_id + "'," + CStr(rsaux10!inte_car_numero) + "," + CStr(rsaux10!floa_Car_importe_neto) + ")", cnn_distribucion, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux9.Close
                     rsaux10.MoveNext
               Wend
               rsaux10.Close
               MsgBox "Se a terminado de exportar la información", vbOKOnly, "ATENCION"
            End If
            If var_empresa = "18" Or var_empresa = "31" Then
               If UCase(parametros(0)) = "DISTRIBUCION" Then
               End If
               If UCase(parametros(0)) = "sqlquezada2" Or UCase(parametros(0)) = "ADMCDINDUSTRIAL" Then
                  x = 0
                  If x = 0 Then
                  rsaux9.Open "select * from VW_FACTURA_TELAS_NO_COMISION", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux9.EOF
                        rsaux8.Open "update tb_encabezado_Cartera set inte_car_tipo_comision = 1 where vcha_emp_empresa_id = '" + rsaux9!VCHA_EMP_EMPRESA_ID + "' and vcha_car_documento = 'FA' and vcha_ser_serie_id = '" + rsaux9!vcha_ser_Serie_id + "' and inte_Car_numero = " + CStr(rsaux9!inte_car_numero), cnn, adOpenDynamic, adLockOptimistic
                        rsaux9.MoveNext
                  Wend
                  rsaux9.Close
                  var_cadena = "insert into distribucion.vianney.dbo.tb_temporal_encabezado_Cartera (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_SER_SERIE_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_CAR_CLASE_ID, INTE_CAR_NUMERO, CHAR_CAR_AFECTACION, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID,"
                  var_cadena = var_cadena + "INTE_EMO_NUMERO, DTIM_CAR_FECHA, VCHA_AGE_AGENTE_ID, VCHA_GAC_GRUPO_ACTUAL_ID, VCHA_GRE_GRUPO_REAL_ID, VCHA_TIT_TITULAR_ID, VCHA_CLI_CLAVE_ID, VCHA_ESB_ESTABLECIMIENTO_ID, INTE_CAR_PLAZO, FLOA_CAR_PORCENTAJE_IVA, FLOA_CAR_PORCENTAJE_IMPUESTO_1, FLOA_CAR_PORCENTAJE_IMPUESTO_2,"
                  var_cadena = var_cadena + " FLOA_CAR_PORCENTAJE_DESCUENTO_1, FLOA_CAR_PORCENTAJE_DESCUENTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_3, FLOA_CAR_IMPORTE_TOTAL, FLOA_CAR_IMPORTE_IVA, FLOA_CAR_IMPORTE_IMPUESTO_1, FLOA_CAR_IMPORTE_IMPUESTO_2, FLOA_CAR_IMPORTE_DESCUENTO_1, FLOA_CAR_IMPORTE_DESCUENTO_2, FLOA_CAR_IMPORTE_DESCUENTO_3,"
                  var_cadena = var_cadena + "FLOA_CAR_SUBIMPORTE, FLOA_CAR_IMPORTE_NETO, VCHA_CAR_IMPORTE_LETRA, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA, VCHA_AUD_FECHA, FLOA_CAR_SALDO, DTIM_CAR_FECHA_VENCIMIENTO, DTIM_CAR_FECHA_ENTREGA, VCHA_MON_MONEDA_ID, FLOA_CAR_TIPO_CAMBIO, CHAR_CAR_ESTATUS, DTIM_CAR_FECHA_CANCELACION, VCHA_CAR_USUARIO_CANCELACION,"
                  var_cadena = var_cadena + "VCHA_CAR_MAQUINA_CANCELACION, CHAR_CAR_TIPO_FACTURACION, INTE_CAR_FACTURA_CEROS, FLOA_CAR_COSTO, VCHA_CAR_CHEQUE, VCHA_CAR_REFERENCIA, INTE_CAR_CAPTURADO, NUM_INTER_TRANC_TYPE, NUM_INTER_UPLOADED, DATE_INTER_DATE, VCHA_CAR_CHEQUE_DEPOSITO, VCHA_CAR_BANCO_CHEQUE, VCHA_CAR_DEPOSITO, VCHA_CAR_BANCO_DEPOSITO,"
                  var_cadena = var_cadena + "DTIM_CAR_FECHA_DEPOSITO, INTE_CAR_TIPO_COMISION) select "
                  var_cadena = var_cadena + "VCHA_EMP_EMPRESA_ID , VCHA_UOR_UNIDAD_ID, VCHA_SER_SERIE_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_CAR_CLASE_ID, INTE_CAR_NUMERO, CHAR_CAR_AFECTACION, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, "
                  var_cadena = var_cadena + "INTE_EMO_NUMERO, DTIM_CAR_FECHA, VCHA_AGE_AGENTE_ID, VCHA_GAC_GRUPO_ACTUAL_ID, VCHA_GRE_GRUPO_REAL_ID, VCHA_TIT_TITULAR_ID, VCHA_CLI_CLAVE_ID, VCHA_ESB_ESTABLECIMIENTO_ID, INTE_CAR_PLAZO, FLOA_CAR_PORCENTAJE_IVA, FLOA_CAR_PORCENTAJE_IMPUESTO_1, FLOA_CAR_PORCENTAJE_IMPUESTO_2,"
                  var_cadena = var_cadena + "FLOA_CAR_PORCENTAJE_DESCUENTO_1, FLOA_CAR_PORCENTAJE_DESCUENTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_3, FLOA_CAR_IMPORTE_TOTAL, FLOA_CAR_IMPORTE_IVA, FLOA_CAR_IMPORTE_IMPUESTO_1, FLOA_CAR_IMPORTE_IMPUESTO_2, FLOA_CAR_IMPORTE_DESCUENTO_1, FLOA_CAR_IMPORTE_DESCUENTO_2, FLOA_CAR_IMPORTE_DESCUENTO_3,"
                  var_cadena = var_cadena + "FLOA_CAR_SUBIMPORTE, FLOA_CAR_IMPORTE_NETO, VCHA_CAR_IMPORTE_LETRA, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA, VCHA_AUD_FECHA, FLOA_CAR_SALDO, DTIM_CAR_FECHA_VENCIMIENTO, DTIM_CAR_FECHA_ENTREGA, VCHA_MON_MONEDA_ID, FLOA_CAR_TIPO_CAMBIO, CHAR_CAR_ESTATUS, DTIM_CAR_FECHA_CANCELACION, VCHA_CAR_USUARIO_CANCELACION,"
                  var_cadena = var_cadena + "VCHA_CAR_MAQUINA_CANCELACION, CHAR_CAR_TIPO_FACTURACION, INTE_CAR_FACTURA_CEROS, FLOA_CAR_COSTO, VCHA_CAR_CHEQUE, VCHA_CAR_REFERENCIA, INTE_CAR_CAPTURADO, NUM_INTER_TRANC_TYPE, NUM_INTER_UPLOADED, DATE_INTER_DATE, VCHA_CAR_CHEQUE_DEPOSITO, VCHA_CAR_BANCO_CHEQUE, VCHA_CAR_DEPOSITO, VCHA_CAR_BANCO_DEPOSITO,"
                  var_cadena = var_cadena + "DTIM_CAR_FECHA_DEPOSITO, INTE_CAR_TIPO_COMISION from tb_encabezado_cartera where (vcha_emp_empresa_ID = '" + var_empresa + "' or vcha_emp_empresa_id = '16') and dtim_CAR_FECHA >= " + var_fecha_inicio + " AND DTIM_CAR_FECHA <= " + var_fecha_fin + "-.000000001"
                  cnn.CommandTimeout = 3000
                  rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  
                  rs.Open "EXEC DISTRIBUCION.VIANNEY.DBO.SP_IMPORTACION_EXPORTACION_CARTERA", cnn, adOpenDynamic, adLockOptimistic
                  rs.Open "DELETE FROM DISTRIBUCION.VIANNEY.DBO.TB_TEMPORAL_ENCABEZADO_CARTERA", cnn, adOpenDynamic, adLockOptimistic
                  rs.Open "DELETE FROM DISTRIBUCION.VIANNEY.DBO.TB_TEMPORAL_ESTADO_CUENTA", cnn, adOpenDynamic, adLockOptimistic
                  
                  
                  End If
                  var_cadena = "insert into tb_temporal_encabezado_Cartera (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_SER_SERIE_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_CAR_CLASE_ID, INTE_CAR_NUMERO, CHAR_CAR_AFECTACION, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID,"
                  var_cadena = var_cadena + "INTE_EMO_NUMERO, DTIM_CAR_FECHA, VCHA_AGE_AGENTE_ID, VCHA_GAC_GRUPO_ACTUAL_ID, VCHA_GRE_GRUPO_REAL_ID, VCHA_TIT_TITULAR_ID, VCHA_CLI_CLAVE_ID, VCHA_ESB_ESTABLECIMIENTO_ID, INTE_CAR_PLAZO, FLOA_CAR_PORCENTAJE_IVA, FLOA_CAR_PORCENTAJE_IMPUESTO_1, FLOA_CAR_PORCENTAJE_IMPUESTO_2,"
                  var_cadena = var_cadena + " FLOA_CAR_PORCENTAJE_DESCUENTO_1, FLOA_CAR_PORCENTAJE_DESCUENTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_3, FLOA_CAR_IMPORTE_TOTAL, FLOA_CAR_IMPORTE_IVA, FLOA_CAR_IMPORTE_IMPUESTO_1, FLOA_CAR_IMPORTE_IMPUESTO_2, FLOA_CAR_IMPORTE_DESCUENTO_1, FLOA_CAR_IMPORTE_DESCUENTO_2, FLOA_CAR_IMPORTE_DESCUENTO_3,"
                  var_cadena = var_cadena + "FLOA_CAR_SUBIMPORTE, FLOA_CAR_IMPORTE_NETO, VCHA_CAR_IMPORTE_LETRA, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA, VCHA_AUD_FECHA, FLOA_CAR_SALDO, DTIM_CAR_FECHA_VENCIMIENTO, DTIM_CAR_FECHA_ENTREGA, VCHA_MON_MONEDA_ID, FLOA_CAR_TIPO_CAMBIO, CHAR_CAR_ESTATUS, DTIM_CAR_FECHA_CANCELACION, VCHA_CAR_USUARIO_CANCELACION,"
                  var_cadena = var_cadena + "VCHA_CAR_MAQUINA_CANCELACION, CHAR_CAR_TIPO_FACTURACION, INTE_CAR_FACTURA_CEROS, FLOA_CAR_COSTO, VCHA_CAR_CHEQUE, VCHA_CAR_REFERENCIA, INTE_CAR_CAPTURADO, NUM_INTER_TRANC_TYPE, NUM_INTER_UPLOADED, DATE_INTER_DATE, VCHA_CAR_CHEQUE_DEPOSITO, VCHA_CAR_BANCO_CHEQUE, VCHA_CAR_DEPOSITO, VCHA_CAR_BANCO_DEPOSITO,"
                  var_cadena = var_cadena + "DTIM_CAR_FECHA_DEPOSITO, INTE_CAR_TIPO_COMISION) select "
                  var_cadena = var_cadena + "VCHA_EMP_EMPRESA_ID , VCHA_UOR_UNIDAD_ID, VCHA_SER_SERIE_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_CAR_CLASE_ID, INTE_CAR_NUMERO, CHAR_CAR_AFECTACION, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, "
                  var_cadena = var_cadena + "INTE_EMO_NUMERO, DTIM_CAR_FECHA, VCHA_AGE_AGENTE_ID, VCHA_GAC_GRUPO_ACTUAL_ID, VCHA_GRE_GRUPO_REAL_ID, VCHA_TIT_TITULAR_ID, VCHA_CLI_CLAVE_ID, VCHA_ESB_ESTABLECIMIENTO_ID, INTE_CAR_PLAZO, FLOA_CAR_PORCENTAJE_IVA, FLOA_CAR_PORCENTAJE_IMPUESTO_1, FLOA_CAR_PORCENTAJE_IMPUESTO_2,"
                  var_cadena = var_cadena + "FLOA_CAR_PORCENTAJE_DESCUENTO_1, FLOA_CAR_PORCENTAJE_DESCUENTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_3, FLOA_CAR_IMPORTE_TOTAL, FLOA_CAR_IMPORTE_IVA, FLOA_CAR_IMPORTE_IMPUESTO_1, FLOA_CAR_IMPORTE_IMPUESTO_2, FLOA_CAR_IMPORTE_DESCUENTO_1, FLOA_CAR_IMPORTE_DESCUENTO_2, FLOA_CAR_IMPORTE_DESCUENTO_3,"
                  var_cadena = var_cadena + "FLOA_CAR_SUBIMPORTE, FLOA_CAR_IMPORTE_NETO, VCHA_CAR_IMPORTE_LETRA, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA, VCHA_AUD_FECHA, FLOA_CAR_SALDO, DTIM_CAR_FECHA_VENCIMIENTO, DTIM_CAR_FECHA_ENTREGA, VCHA_MON_MONEDA_ID, FLOA_CAR_TIPO_CAMBIO, CHAR_CAR_ESTATUS, DTIM_CAR_FECHA_CANCELACION, VCHA_CAR_USUARIO_CANCELACION,"
                  var_cadena = var_cadena + "VCHA_CAR_MAQUINA_CANCELACION, CHAR_CAR_TIPO_FACTURACION, INTE_CAR_FACTURA_CEROS, FLOA_CAR_COSTO, VCHA_CAR_CHEQUE, VCHA_CAR_REFERENCIA, INTE_CAR_CAPTURADO, NUM_INTER_TRANC_TYPE, NUM_INTER_UPLOADED, DATE_INTER_DATE, VCHA_CAR_CHEQUE_DEPOSITO, VCHA_CAR_BANCO_CHEQUE, VCHA_CAR_DEPOSITO, VCHA_CAR_BANCO_DEPOSITO,"
                  var_cadena = var_cadena + "DTIM_CAR_FECHA_DEPOSITO, INTE_CAR_TIPO_COMISION from distribucion.vianney.dbo.tb_encabezado_cartera where (vcha_emp_empresa_ID = '" + var_empresa + "' or vcha_emp_empresa_id = '16') and dtim_CAR_FECHA >= getdate()-1 AND DTIM_CAR_FECHA <= getdate() + 1 -.000000001 and  (char_car_estatus <> 'C')"
                  'MsgBox var_fecha_inicio
                  'MsgBox var_fecha_fin
                  cnn.CommandTimeout = 3000
                  rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  
                  rs.Open "INSERT INTO TB_TEMPORAL_ESTADO_CUENTA  (vcha_emp_empresa_id, vcha_uor_unidad_id, VCHA_ECU_SERIE_CARGO, VCHA_ECU_MOVIMIENTO_CARGO, INTE_ECU_NUMERO_CARGO, FLOA_ECU_IMPORTE_CARGO, CHAR_ECU_ESTATUS, FLOA_ECU_TIPO_CAMBIO_CARGO) SELECT vcha_emp_empresa_id, vcha_uor_unidad_id,  VCHA_ECU_SERIE_CARGO, VCHA_ECU_MOVIMIENTO_CARGO, INTE_ECU_NUMERO_CARGO, FLOA_ECU_IMPORTE_CARGO, CHAR_ECU_ESTATUS, FLOA_ECU_TIPO_CAMBIO_CARGO  FROM DISTRIBUCION.VIANNEY.DBO.VW_ESTADO_CUENTA_CARGOS WHERE vcha_emp_empresa_ID = '" + var_empresa + "' and dtim_CAR_FECHA >= getdate() -1 AND DTIM_CAR_FECHA <= getdate() + 1-.000000001", cnn, adOpenDynamic, adLockOptimistic
                  var_cadena = "INSERT INTO TB_TEMPORAL_ESTADO_CUENTA  (vcha_emp_empresa_id, vcha_uor_unidad_id, VCHA_ECU_SERIE_CARGO, VCHA_ECU_MOVIMIENTO_CARGO, INTE_ECU_NUMERO_CARGO, VCHA_ECU_SERIE_ABONO, VCHA_ECU_MOVIMIENTO_ABONO, INTE_ECU_NUMERO_ABONO, FLOA_ECU_IMPORTE_ABONO, CHAR_ECU_ESTATUS, FLOA_ECU_TIPO_CAMBIO_ABONO) "
                  var_cadena = var_cadena + "                    SELECT vcha_emp_empresa_id, vcha_uor_unidad_id, VCHA_ECU_SERIE_CARGO, VCHA_ECU_MOVIMIENTO_CARGO, INTE_ECU_NUMERO_CARGO, VCHA_ECU_SERIE_ABONO, VCHA_ECU_MOVIMIENTO_ABONO, INTE_ECU_NUMERO_ABONO, FLOA_ECU_IMPORTE_ABONO, CHAR_ECU_ESTATUS, FLOA_ECU_TIPO_CAMBIO_ABONO FROM DISTRIBUCION.VIANNEY.DBO.VW_ESTADO_CUENTA_ABONOS WHERE vcha_emp_empresa_ID = '" + var_empresa + "' and dtim_CAR_FECHA >= getdate() - 1 AND DTIM_CAR_FECHA <= getdate()+1 -.000000001"
                  rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  
                  rs.Open "EXEC SP_IMPORTACION_EXPORTACION_CARTERA", cnn, adOpenDynamic, adLockOptimistic
                  rs.Open "DELETE FROM TB_TEMPORAL_ENCABEZADO_CARTERA", cnn, adOpenDynamic, adLockOptimistic
                  rs.Open "DELETE FROM TB_TEMPORAL_ESTADO_CUENTA", cnn, adOpenDynamic, adLockOptimistic
                  fin = CStr(Now)
                  MsgBox "Termino la importación y exportacion de cartera " + inicio + " " + fin, vbOKOnly, "ATENCION"
              End If
            End If
         Else
            MsgBox "La fecha final debe de ser mayor a la fecha de inicio", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de inicio incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 3500
   Me.txt_inicio = Date - 2
   Me.txt_fin = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fin) Then
         frmcalendario.mes.Value = CDate(Me.txt_fin)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_fin = var_fecha_general
   End If
End Sub

Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_inicio) Then
         frmcalendario.mes.Value = CDate(Me.txt_inicio)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_inicio = var_fecha_general
   End If
End Sub
