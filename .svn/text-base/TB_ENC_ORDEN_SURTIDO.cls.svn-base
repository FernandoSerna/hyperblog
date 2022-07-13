VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_ENC_ORDEN_SURTIDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clCHAR_TPE_TIPO_PEDIDO_ID As String, clINTE_PED_NUMERO As Double, clvcha_alm_almacen_id As String, clinte_ors_orden_surtido As Double, clDTIM_ORS_FECHA_CARGA As String, clDTIM_ORS_FECHA_CADUCA As String, _
clCHAR_ORS_ESTATUS As String, clVCHA_TIT_TITULAR_ID As String, clVCHA_CLI_CLAVE_ID As String, _
clvcha_esb_establecimiento_ID As String, clFLOA_ORS_DESCUENTO_1 As Variant, clFLOA_ORS_DESCUENTO_2 As Variant, clFLOA_ORS_DESCUENTO_3 As Variant, clVcha_aud_usuario As String, clVcha_aud_maquina As String, clDtim_aud_fecha As String, clINTE_ORS_FACTURA_CEROS As Double, clVCHA_MON_MONEDA_ID As String, clTIME_ORS_HORA As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL


Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "ENC_ORDEN_SURTIDO_I"
   CMD("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
   CMD("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
   CMD("@CHAR_TPE_TIPO_PEDIDO_ID") = clCHAR_TPE_TIPO_PEDIDO_ID
   CMD("@INTE_PED_NUMERO") = clINTE_PED_NUMERO
   CMD("@VCHA_ALM_ALMACEN_ID") = clvcha_alm_almacen_id
   CMD("@INTE_ORS_ORDEN_SURTIDO") = clinte_ors_orden_surtido
   CMD("@DTIM_ORS_FECHA_CARGA") = clDTIM_ORS_FECHA_CARGA
   CMD("@DTIM_ORS_FECHA_CADUCA") = clDTIM_ORS_FECHA_CADUCA
   CMD("@CHAR_ORS_ESTATUS") = clCHAR_ORS_ESTATUS
   CMD("@VCHA_TIT_TITULAR_ID") = clVCHA_TIT_TITULAR_ID
   CMD("@VCHA_CLI_CLAVE_ID") = clVCHA_CLI_CLAVE_ID
   CMD("@VCHA_ESB_ESTABLECIMIENTO_ID") = clvcha_esb_establecimiento_ID
   CMD("@FLOA_ORS_DESCUENTO_1") = clFLOA_ORS_DESCUENTO_1
   CMD("@FLOA_ORS_DESCUENTO_2") = clFLOA_ORS_DESCUENTO_2
   CMD("@FLOA_ORS_DESCUENTO_3") = clFLOA_ORS_DESCUENTO_3
   CMD("@VCHA_AUD_USAURIO") = clVcha_aud_usuario
   CMD("@VCHA_AUD_MAQUINA") = clVcha_aud_maquina
   CMD("@DTIM_AUD_FECHA") = clDtim_aud_fecha
   CMD("@INTE_ORS_FACTURA_CEROS") = clINTE_ORS_FACTURA_CEROS
   CMD("@VCHA_MON_MONEDA_ID") = clVCHA_MON_MONEDA_ID
   CMD("@TIME_ORS_HORA") = clTIME_ORS_HORA
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







