VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_ENC_FACTURAS_I"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clvcha_alm_almacen_id As String, clvcha_mov_movimiento_id As String, clinte_fac_numero As Integer, cldtim_fac_fecha As String, clinte_emo_numero As Integer, clVCHA_ORS_ORDEN_SURTIDO As Integer, _
clVCHA_AGE_AGENTE_ID As String, clVCHA_TIT_TITULAR_ID As String, clVCHA_ESB_ESTABLECIMIENTO_ID As String, clVCHA_CLI_CLAVE_ID As String, _
CLVCHA_RUT_RUTA_ID As String, CLINTE_FAC_PLAZO As Integer, clfloa_fac_iva As Variant, CLFLOA_FAC_DESCUENTO_1 As Variant, CLFLOA_FAC_DESCUENTO_2 As Variant, CLFLOA_FAC_DESCUENTO_3 As Variant, VCHA_FAC_AGRUPADOR As String, _
clfloa_fac_subimporte As Variant, clfloa_fac_importe_descuento_1 As Variant, clfloa_fac_importe_descuento_2 As Variant, clfloa_fac_importe_descuento_3 As Variant, clfloa_fac_importe_iva As Variant, clfloa_fac_total As Variant, clchar_fac_estatus As String, clFLOA_FAC_PIEZAS As Variant, clVCHA_FAC_IMPORTE_LETRA As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL




Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "ENCABEZADO_FACTURAS_I"
   CMD("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
   CMD("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
   CMD("@VCHA_ALM_ALMACEN_ID") = clvcha_alm_almacen_id
   CMD("@VCHA_MOV_MOVIMIENTO_ID") = clvcha_mov_movimiento_id
   CMD("@INTE_FAC_NUMERO") = clinte_fac_numero
   CMD("@DTIM_FAC_FECHA") = cldtim_fac_fecha
   CMD("@INTE_EMO_NUMERO") = clinte_emo_numero
   CMD("@INTE_ORS_ORDEN_SURTIDO") = clVCHA_ORS_ORDEN_SURTIDO
   CMD("@VCHA_AGE_AGENTE_ID") = clVCHA_AGE_AGENTE_ID
   CMD("@VCHA_TIT_TITULAR_ID") = clVCHA_TIT_TITULAR_ID
   CMD("@VCHA_ESB_ESTABLECIMIENTO_ID") = clVCHA_ESB_ESTABLECIMIENTO_ID
   CMD("@VCHA_CLI_CLAVE_ID") = clVCHA_CLI_CLAVE_ID
   CMD("@VCHA_RUT_RUTA_ID") = CLVCHA_RUT_RUTA_ID
   CMD("@INTE_FAC_PLAZO") = CLINTE_FAC_PLAZO
   CMD("@FLOA_FAC_IVA") = clfloa_fac_iva
   CMD("@FLOA_FAC_DESCUENTO_1") = CLFLOA_FAC_DESCUENTO_1
   CMD("@FLOA_FAC_DESCUENTO_2") = CLFLOA_FAC_DESCUENTO_2
   CMD("@FLOA_FAC_DESCUENTO_3") = CLFLOA_FAC_DESCUENTO_3
   CMD("@VCHA_FAC_AGRUPADOR") = VCHA_FAC_AGRUPADOR
   CMD("@FLOA_FAC_SUBIMPORTE") = clfloa_fac_subimporte
   CMD("@FLOA_FAC_IMPORTE_DESCUENTO_1") = clfloa_fac_importe_descuento_1
   CMD("@FLOA_FAC_IMPORTE_DESCUENTO_2") = clfloa_fac_importe_descuento_2
   CMD("@FLOA_FAC_IMPORTE_DESCUENTO_3") = clfloa_fac_importe_descuento_3
   CMD("@FLOA_FAC_IMPORTE_IVA") = clfloa_fac_importe_iva
   CMD("@FLOA_FAC_TOTAL") = clfloa_fac_total
   CMD("@CHAR_FAC_ESTATUS") = clchar_fac_estatus
   CMD("@FLOA_FAC_PIEZAS") = clFLOA_FAC_PIEZAS
   CMD("@VCHA_FAC_IMPORTE_LETRA") = clVCHA_FAC_IMPORTE_LETRA
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






