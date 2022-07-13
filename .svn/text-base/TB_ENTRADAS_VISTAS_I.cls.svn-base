VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_ENTRADAS_VISTAS_I"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clvcha_alm_almacen_id As String, clVcha_evi_movimiento_salida As String, clinte_evi_numero_salida As Double, clvcha_mov_movimiento_id As String, clINTE_evi_NUMERO As Double, clVCHA_ART_ARTICULO_ID As String, _
clFLOA_evi_CANTIDAD As Variant, clFLOA_evi_COSTO As Variant, clFLOA_evi_PRECIO As Variant) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL




Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "ENTRADA_VISTAS_I"
   CMD("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
   CMD("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
   CMD("@VCHA_ALM_ALMACEN_ID") = clvcha_alm_almacen_id
   CMD("@VCHA_EVI_MOVIMIENTO_SALIDA") = clVcha_evi_movimiento_salida
   CMD("@INTE_EVI_NUMERO_SALIDA") = clinte_evi_numero_salida
   CMD("@VCHA_MOV_MOVIMIENTO_ID") = clvcha_mov_movimiento_id
   CMD("@INTE_EVI_NUMERO") = clINTE_evi_NUMERO
   CMD("@VCHA_ART_ARTICULO_ID") = clVCHA_ART_ARTICULO_ID
   CMD("@FLOA_EVI_CANTIDAD") = clFLOA_evi_CANTIDAD
   CMD("@FLOA_EVI_COSTO") = clFLOA_evi_COSTO
   CMD("@FLOA_EVI_PRECIO") = clFLOA_evi_PRECIO
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

















