VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_MERC_VISTAS_M"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clVCHA_ALM_ALMACEN_ID As String, clVcha_mov_movimiento_id As String, clINTE_EMO_NUMERO As Integer, clVCHA_ART_ARTICULO_ID As String, _
clFLOA_MVI_CANTIDAD As Variant) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL




Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "MERCANCIA_VISTAS_M"
   CMD("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
   CMD("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
   CMD("@VCHA_ALM_ALMACEN_ID") = clVCHA_ALM_ALMACEN_ID
   CMD("@VCHA_MOV_MOVIMIENTO_ID") = clVcha_mov_movimiento_id
   CMD("@INTE_EMO_NUMERO") = clINTE_EMO_NUMERO
   CMD("@VCHA_ART_ARTICULO_ID") = clVCHA_ART_ARTICULO_ID
   CMD("@FLOA_MVI_CANTIDAD") = clFLOA_MVI_CANTIDAD
   CMD("@FLOA_MVI_COSTO") = clFLOA_MVI_COSTO
   CMD("@FLOA_MVI_PRECIO") = clFLOA_MVI_PRECIO
   CMD("@FLOA_MVI_CANTIDAD_DEVUELTA") = clFLOA_MVI_CANTIDAD_DEVUELTA
   CMD("@FLOA_MVI_CANTIDAD_FACTURADA") = clFLOA_MVI_CANTIDAD_FACTURADA
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















