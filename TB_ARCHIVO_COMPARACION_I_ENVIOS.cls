VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_ARCHIVO_COMPARACION_I_ENVIOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clvcha_alm_almacen_id As String, clvcha_mov_movimiento_id As String, clINTE_COM_NUMERO As Double, clDTIM_COM_FECHA As Date, _
clCHAR_COM_TIPO_PROVEEDOR As String, clVCHA_COM_PROVEEDOR As String, clVCHA_ART_ARTICULO_ID As String, clFLOA_COM_COSTO As Variant, _
clFLOA_COM_CANTIDAD_ENVIADA As Variant, clFLOA_COM_CANTIDAD_RECIBIDA As Variant, clVCHA_COM_TRANSPORTO As String, clVCHA_COM_REFERENCIA As String, clINTE_COM_LOTE As Double, clINTE_COM_CONSECUTIVO As Integer, clINTE_COM_AÑO As Integer, clVCHA_COM_CAJA As String, clFLOA_COM_PESO As Double) As Boolean
Dim rs As ADODB.recordSet, Cadena As String
Anadir = True
'On Error GoTo HELL

Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "ARCHIVO_COMPARACION_I_ENVIOS"
   CMD("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
   CMD("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
   CMD("@VCHA_ALM_ALMACEN_ID") = clvcha_alm_almacen_id
   CMD("@VCHA_MOV_MOVIMIENTO_ID") = clvcha_mov_movimiento_id
   CMD("@INTE_COM_NUMERO") = clINTE_COM_NUMERO
   CMD("@DTIM_COM_FECHA") = clDTIM_COM_FECHA
   CMD("@CHAR_COM_TIPO_PROVEEDOR") = clCHAR_COM_TIPO_PROVEEDOR
   CMD("@VCHA_COM_PROVEEDOR") = clVCHA_COM_PROVEEDOR
   CMD("@VCHA_ART_ARTICULO_ID") = clVCHA_ART_ARTICULO_ID
   CMD("@FLOA_COM_COSTO") = clFLOA_COM_COSTO
   CMD("@FLOA_COM_CANTIDAD_ENVIADA") = clFLOA_COM_CANTIDAD_ENVIADA
   CMD("@FLOA_COM_CANTIDAD_RECIBIDA") = clFLOA_COM_CANTIDAD_RECIBIDA
   CMD("@VCHA_COM_TRANSPORTO") = clVCHA_COM_TRANSPORTO
   CMD("@VCHA_COM_REFERENCIA") = clVCHA_COM_REFERENCIA
   CMD("@INTE_COM_LOTE") = clINTE_COM_LOTE
   CMD("@INTE_COM_CONSECUTIVO") = clINTE_COM_CONSECUTIVO
   CMD("@INTE_COM_AÑO") = clINTE_COM_AÑO
   CMD("@VCHA_COM_CAJA") = clVCHA_COM_CAJA
   CMD("@FLOA_COM_PESO") = clFLOA_COM_PESO
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













