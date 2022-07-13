VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_ENCABEZADO_CARTERA_I"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
   
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clVCHA_CAR_TIPO_DOCUMENTO As String, clvcha_car_documento As String, clVcha_car_clase_id As String, clinte_car_numero As Double, clchar_car_afectacion As String, _
   clvcha_alm_almacen_id As String, clvcha_mov_movimiento_id As String, clinte_emo_numero As Double, cldtim_Car_fecha As String, clVCHA_AGE_AGENTE_ID As String, _
   clVcha_gac_grupo_actual_id As String, clVcha_gre_grupo_real_id As String, clVCHA_TIT_TITULAR_ID As String, clVCHA_CLI_CLAVE_ID As String, clvcha_esb_establecimiento_ID As String, _
   clINTE_CAR_PLAZO As Integer, clfloa_car_porcentaje_iva As Double, clfloa_Car_porcentaje_impuesto_1 As Double, clfloa_car_porcentaje_impuesto_2 As Double, clfloa_car_porcentaje_descuento_1 As Double, _
   clfloa_car_porcentaje_descuento_2 As Double, clfloa_car_porcentaje_Descuento_3 As Double, clfloa_car_importe_total As Double, clfloa_car_importe_iva As Double, clfloa_car_importe_impuesto_1 As Double, _
   clfloa_car_importe_impuesto_2 As Double, clfloa_car_importe_descuento_1 As Double, clfloa_car_importe_descuento_2 As Double, clfloa_car_importe_descuento_3 As Double, clfloa_car_subimporte As Double, _
   clfloa_car_importe_neto As Double, clvcha_car_importe_letra As String, clVcha_aud_usuario As String, clVcha_aud_maquina As String, clvcha_aud_fecha As Date, _
   clfloa_Car_saldo As Double, cldtim_car_fecha_vencimiento As Date, cldtim_car_fecha_entrega As Date, clVCHA_MON_MONEDA_ID As String, clFLOA_CAR_TIPO_CAMBIO As Double, clVCHA_SER_SERIE_ID As String, clCHAR_CAR_ESTATUS As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL

Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "ENCABEZADO_CARTERA_I"
   CMD("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
   CMD("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
   CMD("@VCHA_CAR_TIPO_DOCUMENTO") = clVCHA_CAR_TIPO_DOCUMENTO
   CMD("@VCHA_CAR_DOCUMENTO") = clvcha_car_documento
   CMD("@VCHA_CAR_CLASE_ID") = clVcha_car_clase_id
   CMD("@INTE_CAR_NUMERO") = clinte_car_numero
   CMD("@CHAR_CAR_AFECTACION") = clchar_car_afectacion
   CMD("@VCHA_ALM_ALMACEN_ID") = clvcha_alm_almacen_id
   CMD("@VCHA_MOV_MOVIMIENTO_ID") = clvcha_mov_movimiento_id
   CMD("@INTE_EMO_NUMERO") = clinte_emo_numero
   CMD("@DTIM_CAR_FECHA") = cldtim_Car_fecha
   CMD("@VCHA_AGE_AGENTE_ID") = clVCHA_AGE_AGENTE_ID
   CMD("@VCHA_GAC_GRUPO_ACTUAL_ID") = clVcha_gac_grupo_actual_id
   CMD("@VCHA_GRE_GRUPO_REAL_ID") = clVcha_gre_grupo_real_id
   CMD("@VCHA_TIT_TITULAR_ID") = clVCHA_TIT_TITULAR_ID
   CMD("@VCHA_CLI_CLAVE_ID") = clVCHA_CLI_CLAVE_ID
   CMD("@VCHA_ESB_ESTABLECIMIENTO_ID") = clvcha_esb_establecimiento_ID
   CMD("@INTE_CAR_PLAZO") = clINTE_CAR_PLAZO
   CMD("@FLOA_CAR_PORCENTAJE_IVA") = clfloa_car_porcentaje_iva
   CMD("@FLOA_CAR_PORCENTAJE_IMPUESTO_1") = clfloa_Car_porcentaje_impuesto_1
   CMD("@FLOA_CAR_PORCENTAJE_IMPUESTO_2") = clfloa_car_porcentaje_impuesto_2
   CMD("@FLOA_CAR_PORCENTAJE_DESCUENTO_1") = clfloa_car_porcentaje_descuento_1
   CMD("@FLOA_CAR_PORCENTAJE_DESCUENTO_2") = clfloa_car_porcentaje_descuento_2
   CMD("@FLOA_CAR_PORCENTAJE_DESCUENTO_3") = clfloa_car_porcentaje_Descuento_3
   CMD("@FLOA_CAR_IMPORTE_TOTAL") = clfloa_car_importe_total
   CMD("@FLOA_CAR_IMPORTE_IVA") = clfloa_car_importe_iva
   CMD("@FLOA_CAR_IMPORTE_IMPUESTO_1") = clfloa_car_importe_impuesto_1
   CMD("@FLOA_CAR_IMPORTE_IMPUESTO_2") = clfloa_car_importe_impuesto_2
   CMD("@FLOA_CAR_IMPORTE_DESCUENTO_1") = clfloa_car_importe_descuento_1
   CMD("@FLOA_CAR_IMPORTE_DESCUENTO_2") = clfloa_car_importe_descuento_2
   CMD("@FLOA_CAR_IMPORTE_DESCUENTO_3") = clfloa_car_importe_descuento_3
   CMD("@FLOA_CAR_SUBIMPORTE") = clfloa_car_subimporte
   CMD("@FLOA_CAR_IMPORTE_NETO") = clfloa_car_importe_neto
   CMD("@VCHA_CAR_IMPORTE_LETRA") = clvcha_car_importe_letra
   CMD("@VCHA_AUD_USUARIO") = clVcha_aud_usuario
   CMD("@VCHA_AUD_MAQUINA") = clVcha_aud_maquina
   CMD("@VCHA_AUD_FECHA") = clvcha_aud_fecha
   CMD("@FLOA_CAR_SALDO") = clfloa_Car_saldo
   CMD("@DTIM_CAR_FECHA_VENCIMIENTO") = cldtim_car_fecha_vencimiento
   CMD("@DTIM_CAR_FECHA_ENTREGA") = cldtim_car_fecha_entrega
   CMD("@VCHA_MON_MONEDA_ID") = clVCHA_MON_MONEDA_ID
   CMD("@FLOA_CAR_TIPO_CAMBIO") = clFLOA_CAR_TIPO_CAMBIO
   CMD("@VCHA_SER_SERIE_ID") = clVCHA_SER_SERIE_ID
   CMD("@CHAR_CAR_ESTATUS") = clCHAR_CAR_ESTATUS
   CMD.execute
Set CMD = Nothing
SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Anadir = False
    GoTo SIGUE
End Function















