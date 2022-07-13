VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_TEMPORAL_SALIDAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim cmd As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVcha_alm_almacen_id As String, clVcha_mov_movimiento_id As String, clINTE_SAL_NUMERO As Integer, clVCHA_ART_ARTICULO_ID As String, _
clFLOA_SAL_CANTIDAD As Variant, clFLOA_SAL_COSTO As Variant, clFLOA_SAL_PRECIO As Variant, clFLOA_SAL_DESCUENTO As Variant) As Boolean
Dim rs As adodb.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL




Set cmd.ActiveConnection = cnn
   cmd.CommandType = adCmdStoredProc
   cmd.CommandText = "TEMPORAL_SALIDAS_I"
   cmd("@VCHA_ALM_ALMACEN_ID") = clVcha_alm_almacen_id
   cmd("@VCHA_MOV_MOVIMIENTO_ID") = clVcha_mov_movimiento_id
   cmd("@INTE_SAL_NUMERO") = clINTE_SAL_NUMERO
   cmd("@VCHA_ART_ARTICULO_ID") = clVCHA_ART_ARTICULO_ID
   cmd("@FLOA_SAL_CANTIDAD") = clFLOA_SAL_CANTIDAD
   cmd("@FLOA_SAL_COSTO") = clFLOA_SAL_COSTO
   cmd("@FLOA_SAL_PRECIO") = clFLOA_SAL_PRECIO
   cmd("@FLOA_SAL_DESCUENTO") = clFLOA_SAL_DESCUENTO
   cmd.execute
Set cmd = Nothing
SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Anadir = False
    GoTo SIGUE
End Function

















