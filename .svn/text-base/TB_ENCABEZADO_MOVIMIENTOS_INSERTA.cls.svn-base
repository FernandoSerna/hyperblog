VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_ENCABEZADO_MOVIMIENTOS_I"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clvcha_alm_almacen_id As String, clvcha_mov_movimiento_id As String, clDTIM_EMO_FECHA As String, clinte_emo_numero As Double, clINTE_EMO_NUMERO_ORIGEN As Integer, clVCHA_CLI_CLAVE_ID As String, clvcha_pro_proveedor_id As String, clVCHA_EMO_ALMACEN_ORIGEN As String, clVCHA_EMO_ALMACEN_DESTINO As String, clCHAR_EMO_ESTATUS As String, clVcha_aud_usuario As String, clVcha_aud_maquina As String, clvcha_emo_factura As String, clVCHA_EMO_MOVIMIENTO_ORIGEN As String, clVCHA_EMO_REFERENCIA As String, clvcha_esb_establecimiento_ID As String, clCHAR_EMO_BLOQUEADO As String, clVCHA_TIT_TITULAR_ID As String, clVCHA_AGE_AGENTE_ID As String, clFLOA_EMO_DESCUENTO_1 As Double, clFLOA_EMO_DESCUENTO_2 As Double, clFLOA_EMO_DESCUENTO_3, clVCHA_MON_MONEDA_ID As String, clFLOA_EMO_TIPO_CAMBIO As Double) As Boolean

Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL


Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "ENCABEZADO_MOVIMIENTOS_I"
   CMD("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
   CMD("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
   CMD("@VCHA_ALM_ALMACEN_ID") = clvcha_alm_almacen_id
   CMD("@VCHA_MOV_MOVIMIENTO_ID") = clvcha_mov_movimiento_id
   CMD("@DTIM_EMO_FECHA") = clDTIM_EMO_FECHA
   CMD("@INTE_EMO_NUMERO") = clinte_emo_numero
   CMD("@INTE_EMO_NUMERO_ORIGEN") = clINTE_EMO_NUMERO_ORIGEN
   CMD("@VCHA_CLI_CLAVE_ID") = clVCHA_CLI_CLAVE_ID
   CMD("@VCHA_PRO_PROVEEDOR_ID") = clvcha_pro_proveedor_id
   CMD("@VCHA_EMO_ALMACEN_ORIGEN") = clVCHA_EMO_ALMACEN_ORIGEN
   CMD("@VCHA_EMO_ALMACEN_DESTINO") = clVCHA_EMO_ALMACEN_DESTINO
   CMD("@CHAR_EMO_ESTATUS") = clCHAR_EMO_ESTATUS
   CMD("@VCHA_AUD_USUARIO") = clVcha_aud_usuario
   CMD("@VCHA_AUD_MAQUINA") = clVcha_aud_maquina
   CMD("@VCHA_EMO_FACTURA") = clvcha_emo_factura
   CMD("@VCHA_EMO_MOVIMIENTO_ORIGEN") = clVCHA_EMO_MOVIMIENTO_ORIGEN
   CMD("@VCHA_EMO_REFERENCIA") = clVCHA_EMO_REFERENCIA
   CMD("@VCHA_ESB_ESTABLECIMIENTO_ID") = clvcha_esb_establecimiento_ID
   CMD("@CHAR_EMO_BLOQUEADO") = clCHAR_EMO_BLOQUEADO
   CMD("@VCHA_TIT_TITULAR_ID") = clVCHA_TIT_TITULAR_ID
   CMD("@VCHA_AGE_AGENTE_ID") = clVCHA_AGE_AGENTE_ID
   CMD("@FLOA_EMO_DESCUENTO_1") = clFLOA_EMO_DESCUENTO_1
   CMD("@FLOA_EMO_DESCUENTO_2") = clFLOA_EMO_DESCUENTO_2
   CMD("@FLOA_EMO_DESCUENTO_3") = clFLOA_EMO_DESCUENTO_3
   CMD("@VCHA_MON_MONEDA_ID") = clVCHA_MON_MONEDA_ID
   CMD("@FLOA_EMO_TIPO_CAMBIO") = clFLOA_EMO_TIPO_CAMBIO
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











