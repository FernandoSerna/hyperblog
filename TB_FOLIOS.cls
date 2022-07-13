VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_FOLIOS_MOVIMIENTOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim cmd As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clvcha_alm_almacen_id As String, clvcha_mov_movimiento_id As String, clDTIM_EMO_FECHA As String, clinte_emo_numero As Double, clINTE_EMO_NUMERO_ORIGEN As Double, clVCHA_CLI_CLAVE_ID As String, clvcha_pro_proveedor_id As String, clVCHA_EMO_ALMACEN_ORIGEN As String, clVCHA_EMO_ALMACEN_DESTINO As String, clCHAR_EMO_ESTATUS As String, clVcha_aud_usuario As String, clVcha_aud_maquina As String, clvcha_emo_factura As String, clVCHA_EMO_MOVIMIENTO_ORIGEN As String, clVCHA_EMO_REFERENCIA As String, clvcha_esb_establecimiento_ID As String, clCHAR_EMO_BLOQUEADO As String, clVCHA_TIT_TITULAR_ID As String, clVCHA_AGE_AGENTE_ID As String, clFLOA_EMO_DESCUENTO_1 As Double, clFLOA_EMO_DESCUENTO_2 As Double, clFLOA_EMO_DESCUENTO_3 As Double, clVCHA_MON_MONEDA_ID As String, clFLOA_EMO_TIPO_CAMBIO As Double) As Boolean

Dim rs As ADODB.recordSet, Cadena As String
Anadir = True
'On Error GoTo HELL


Set cmd.ActiveConnection = cnn
   cmd.CommandType = adCmdStoredProc
   cmd.CommandText = "folio_movimiento"
   cmd("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
   cmd("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
   cmd("@VCHA_ALM_ALMACEN_ID") = clvcha_alm_almacen_id
   cmd("@VCHA_MOV_MOVIMIENTO_ID") = clvcha_mov_movimiento_id
   cmd("@DTIM_EMO_FECHA") = clDTIM_EMO_FECHA
   cmd("@INTE_EMO_NUMERO") = clinte_emo_numero
   cmd("@INTE_EMO_NUMERO_ORIGEN") = clINTE_EMO_NUMERO_ORIGEN
   cmd("@VCHA_CLI_CLAVE_ID") = clVCHA_CLI_CLAVE_ID
   cmd("@VCHA_PRO_PROVEEDOR_ID") = clvcha_pro_proveedor_id
   cmd("@VCHA_EMO_ALMACEN_ORIGEN") = clVCHA_EMO_ALMACEN_ORIGEN
   cmd("@VCHA_EMO_ALMACEN_DESTINO") = clVCHA_EMO_ALMACEN_DESTINO
   cmd("@CHAR_EMO_ESTATUS") = clCHAR_EMO_ESTATUS
   cmd("@VCHA_AUD_USUARIO") = clVcha_aud_usuario
   cmd("@VCHA_AUD_MAQUINA") = clVcha_aud_maquina
   cmd("@VCHA_EMO_FACTURA") = clvcha_emo_factura
   cmd("@VCHA_EMO_MOVIMIENTO_ORIGEN") = clVCHA_EMO_MOVIMIENTO_ORIGEN
   cmd("@VCHA_EMO_REFERENCIA") = clVCHA_EMO_REFERENCIA
   cmd("@VCHA_ESB_ESTABLECIMIENTO_ID") = clvcha_esb_establecimiento_ID
   cmd("@CHAR_EMO_BLOQUEADO") = clCHAR_EMO_BLOQUEADO
   cmd("@VCHA_TIT_TITULAR_ID") = clVCHA_TIT_TITULAR_ID
   cmd("@VCHA_AGE_AGENTE_ID") = clVCHA_AGE_AGENTE_ID
   cmd("@FLOA_EMO_DESCUENTO_1") = clFLOA_EMO_DESCUENTO_1
   cmd("@FLOA_EMO_DESCUENTO_2") = clFLOA_EMO_DESCUENTO_2
   cmd("@FLOA_EMO_DESCUENTO_3") = clFLOA_EMO_DESCUENTO_3
   cmd("@VCHA_MON_MONEDA_ID") = clVCHA_MON_MONEDA_ID
   cmd("@FLOA_EMO_TIPO_CAMBIO") = clFLOA_EMO_TIPO_CAMBIO
   cnn.CommandTimeout = 360
   cmd.execute
   var_numero_folio_regreso = IIf(IsNull(cmd("@numero_folio")), 0, cmd("@numero_folio"))
Set cmd = Nothing
SIGUE:
On Error GoTo 0
Exit Function
HELL:
   x = Err.Number
   If x = -2147467259 Then
      On Error GoTo HELL:
      Sleep 30000
      cmd.execute
   Else
      MensajeError = Err.Description
      Anadir = False
      MsgBox MensajeError
   End If
   GoTo SIGUE
End Function












