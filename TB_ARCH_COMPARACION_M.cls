VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_ARCH_COMPARACION_M"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clvcha_alm_almacen_id As String, clvcha_mov_movimiento_id As String, clINTE_COM_NUMERO As Double, _
clCHAR_COM_TIPO_PROVEEDOR As String, clVCHA_COM_PROVEEDOR As String, clVCHA_ART_ARTICULO_ID As String, _
clFLOA_COM_CANTIDAD_RECIBIDA As Double, clVCHA_COM_REFERENCIA As String, clINTE_COM_CONSECUTIVO As Integer) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL

Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "ARCHIVO_COMPARACION_M"
   CMD("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
   CMD("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
   CMD("@VCHA_ALM_ALMACEN_ID") = clvcha_alm_almacen_id
   CMD("@VCHA_MOV_MOVIMIENTO_ID") = clvcha_mov_movimiento_id
   CMD("@INTE_COM_NUMERO") = clINTE_COM_NUMERO
   CMD("@CHAR_COM_TIPO_PROVEEDOR") = clCHAR_COM_TIPO_PROVEEDOR
   CMD("@VCHA_COM_PROVEEDOR") = clVCHA_COM_PROVEEDOR
   CMD("@VCHA_ART_ARTICULO_ID") = clVCHA_ART_ARTICULO_ID
   CMD("@FLOA_COM_CANTIDAD_RECIBIDA") = clFLOA_COM_CANTIDAD_RECIBIDA
   CMD("@VCHA_COM_REFERENCIA") = clVCHA_COM_REFERENCIA
   CMD("@INTE_COM_CONSECUTIVO") = clINTE_COM_CONSECUTIVO
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













