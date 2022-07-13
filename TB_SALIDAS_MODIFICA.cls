VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_SALIDAS_MODIFICA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clvcha_alm_almacen_id As String, clvcha_mov_movimiento_id As String, clINTE_SAL_NUMERO As Double, clvcha_art_articulo_ID As String, _
clFLOA_SAL_CANTIDAD As Variant, clFLOA_SAL_COSTO As Variant, clFLOA_SAL_PRECIO As Variant, clFLOA_SAL_DESCUENTO As Variant) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL




Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "SALIDAS_M"
   CMD("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
   CMD("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
   CMD("@VCHA_ALM_ALMACEN_ID") = clvcha_alm_almacen_id
   CMD("@VCHA_MOV_MOVIMIENTO_ID") = clvcha_mov_movimiento_id
   CMD("@INTE_SAL_NUMERO") = clINTE_SAL_NUMERO
   CMD("@VCHA_ART_ARTICULO_ID") = clvcha_art_articulo_ID
   CMD("@FLOA_SAL_CANTIDAD") = clFLOA_SAL_CANTIDAD
   CMD("@FLOA_SAL_COSTO") = clFLOA_SAL_COSTO
   CMD("@FLOA_SAL_PRECIO") = clFLOA_SAL_PRECIO
   CMD("@FLOA_SAL_DESCUENTO") = clFLOA_SAL_DESCUENTO
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

















