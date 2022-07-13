VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_ENC_PEDIDOS_I"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim CMD As New Command



'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clvcha_alm_almacen_id As String, clCHAR_TPE_TIPO_PEDIDO_ID As String, clINTE_PED_NUMERO As Variant, clINTE_PED_REFERENCIA As Variant, _
clDTIM_PED_FECHA As String, clDTIM_PED_REFERENCIA As String, clVCHA_AGE_AGENTE_ID As String, clVCHA_TIT_TITULAR_ID As String, clVCHA_CLI_CLAVE_ID As String, clvcha_esb_establecimiento_ID As String, _
clINTE_PED_RESURTIBLE As Integer, clINTE_PED_ESPECIALES As Integer, clCHAR_PED_ESTATUS As String, clfloa_ped_descuento_1 As Variant, clfloa_ped_descuento_2 As Variant, clfloa_ped_descuento_3 As Variant, _
clINTE_PED_DIAS_CONDICIONES As Integer, clINTE_PED_DIAS_CADUCIDAD As Integer, clVcha_aud_usuario As String, clVcha_aud_maquina As String, clDtim_aud_fecha As String, clVCHA_MON_MONEDA_ID As String, clINTE_PED_SUGERIDO As Integer) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL
Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    CMD.CommandText = "ENCABEZADO_PEDIDOS_I"
    CMD("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
    CMD("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
    CMD("@VCHA_ALM_ALMACEN_ID") = clvcha_alm_almacen_id
    CMD("@CHAR_TPE_TIPO_PEDIDO_ID") = clCHAR_TPE_TIPO_PEDIDO_ID
    CMD("@INTE_PED_NUMERO") = clINTE_PED_NUMERO
    CMD("@INTE_PED_REFERENCIA") = clINTE_PED_REFERENCIA
    CMD("@DTIM_PED_FECHA") = clDTIM_PED_FECHA
    CMD("@DTIM_PED_REFERENCIA") = clDTIM_PED_REFERENCIA
    CMD("@VCHA_AGE_AGENTE_ID") = clVCHA_AGE_AGENTE_ID
    CMD("@VCHA_TIT_TITULAR_ID") = clVCHA_TIT_TITULAR_ID
    CMD("@VCHA_CLI_CLAVE_ID") = clVCHA_CLI_CLAVE_ID
    CMD("@VCHA_ESB_ESTABLECIMIENTO_ID") = clvcha_esb_establecimiento_ID
    CMD("@INTE_PED_RESURTIBLE") = clINTE_PED_RESURTIBLE
    CMD("@INTE_PED_ESPECIALES") = clINTE_PED_ESPECIALES
    CMD("@CHAR_PED_ESTATUS") = clCHAR_PED_ESTATUS
    CMD("@FLOA_PED_DESCUENTO_1") = clfloa_ped_descuento_1
    CMD("@FLOA_PED_DESCUENTO_2") = clfloa_ped_descuento_2
    CMD("@FLOA_PED_DESCUENTO_3") = clfloa_ped_descuento_3
    CMD("@INTE_PED_DIAS_CONDICIONES") = clINTE_PED_DIAS_CONDICIONES
    CMD("@INTE_PED_DIAS_CADUCIDAD") = clINTE_PED_DIAS_CADUCIDAD
    CMD("@VCHA_AUD_USUARIO") = clVcha_aud_usuario
    CMD("@VCHA_AUD_MAQUINA") = clVcha_aud_maquina
    CMD("@DTIM_AUD_FECHA") = clDtim_aud_fecha
    CMD("@VCHA_MON_MONEDA_ID") = clVCHA_MON_MONEDA_ID
    CMD("@INTE_PED_SUGERIDO") = clINTE_PED_SUGERIDO
CMD.execute
    maximo_pedido = CMD("@INTE_PED_NUMERO")
Set CMD = Nothing
SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Anadir = False
    GoTo SIGUE
End Function
