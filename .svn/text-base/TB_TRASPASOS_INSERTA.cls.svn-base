VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_TRASPASOS_INSERTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clvcha_alm_almacen_id As String, clvcha_mov_movimiento_id As String, clINTE_TRA_NUMERO As Integer, clVCHA_ART_ARTICULO_ID As String, _
clFLOA_TRA_CANTIDAD As Variant, clFLOA_TRA_CANTIDAD_recibida As Variant, clFLOA_TRA_COSTO As Variant, clFLOA_TRA_PRECIO As Variant, clFLOA_TRA_DESCUENTO As Variant, _
clVCHA_TRA_ALMACEN_ORIGEN As String, clINTE_TRA_AÑO As Integer) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL




Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "TRASPASOS_I"
   CMD("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
   CMD("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
   CMD("@VCHA_ALM_ALMACEN_ID") = clvcha_alm_almacen_id
   CMD("@VCHA_MOV_MOVIMIENTO_ID") = clvcha_mov_movimiento_id
   CMD("@INTE_TRA_NUMERO") = clINTE_TRA_NUMERO
   CMD("@VCHA_ART_ARTICULO_ID") = clVCHA_ART_ARTICULO_ID
   CMD("@FLOA_TRA_CANTIDAD") = clFLOA_TRA_CANTIDAD
   CMD("@FLOA_TRA_CANTIDAD_RECIBIDA") = clFLOA_TRA_CANTIDAD_recibida
   CMD("@FLOA_TRA_COSTO") = clFLOA_TRA_COSTO
   CMD("@FLOA_TRA_PRECIO") = clFLOA_TRA_PRECIO
   CMD("@FLOA_TRA_DESCUENTO") = clFLOA_TRA_DESCUENTO
   CMD("@VCHA_TRA_ALMACEN_ORIGEN") = clVCHA_TRA_ALMACEN_ORIGEN
   CMD("@INTE_TRA_AÑO") = clINTE_TRA_AÑO
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















