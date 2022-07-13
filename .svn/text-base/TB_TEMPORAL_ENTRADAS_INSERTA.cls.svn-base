VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_TEMPORAL_ENTRADAS_INSERTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clvcha_alm_almacen_id As String, clvcha_mov_movimiento_id As String, clINTE_ENT_NUMERO As Double, clVCHA_ART_ARTICULO_ID As String, _
clFLOA_ENT_CANTIDAD As Variant, clFLOA_ENT_COSTO As Variant, clFLOA_ENT_PRECIO As Variant, clFLOA_ENT_DESCUENTO As Variant, _
clVCHA_ENT_ALMACEN_ORIGEN As String, clINTE_ENT_AÑO As Integer) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL




Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "TEMPORAL_ENTRADAS_I"
   CMD("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
   CMD("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
   CMD("@VCHA_ALM_ALMACEN_ID") = clvcha_alm_almacen_id
   CMD("@VCHA_MOV_MOVIMIENTO_ID") = clvcha_mov_movimiento_id
   CMD("@INTE_ENT_NUMERO") = clINTE_ENT_NUMERO
   CMD("@VCHA_ART_ARTICULO_ID") = clVCHA_ART_ARTICULO_ID
   CMD("@FLOA_ENT_CANTIDAD") = clFLOA_ENT_CANTIDAD
   CMD("@FLOA_ENT_COSTO") = clFLOA_ENT_COSTO
   CMD("@FLOA_ENT_PRECIO") = clFLOA_ENT_PRECIO
   CMD("@FLOA_ENT_DESCUENTO") = clFLOA_ENT_DESCUENTO
   CMD("@VCHA_ENT_ALMACEN_ORIGEN") = clVCHA_ENT_ALMACEN_ORIGEN
   CMD("@INTE_ENT_AÑO") = clINTE_ENT_AÑO
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















