VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_MERC_VISTAS_I"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clVCHA_ALM_ALMACEN_ID As String, clVcha_mov_movimiento_id As String, clINTE_SAL_NUMERO As Integer, clVCHA_ART_ARTICULO_ID As String, _
clFLOA_SVI_CANTIDAD As Variant, clFLOA_SVI_COSTO As Variant, clFLOA_SVI_PRECIO As Variant) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL

Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "SALIDA_VISTAS_I"
   CMD("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
   CMD("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
   CMD("@VCHA_ALM_ALMACEN_ID") = clVCHA_ALM_ALMACEN_ID
   CMD("@VCHA_MOV_MOVIMIENTO_ID") = clVcha_mov_movimiento_id
   CMD("@INTE_SAL_NUMERO") = clINTE_SAL_NUMERO
   CMD("@VCHA_ART_ARTICULO_ID") = clVCHA_ART_ARTICULO_ID
   CMD("@FLOA_SVI_CANTIDAD") = clFLOA_SVI_CANTIDAD
   CMD("@FLOA_SVI_COSTO") = clFLOA_SVI_COSTO
   CMD("@FLOA_SVI_PRECIO") = clFLOA_SVI_PRECIO
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















