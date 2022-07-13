VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_DET_ORDEN_SURTIDO_M"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clvcha_alm_almacen_id As String, clinte_ors_orden_surtido As Double, clVCHA_ART_ARTICULO_ID As String, clFLOA_ORS_CANTIDAD_SURTIDA As Variant, clFLOA_ORS_CANTIDAD_EMPACADA As Variant, clFLOA_ORS_PRECIO As Variant, clCHAR_PED_TIPO As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL




Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "DET_ORDEN_SURTIDO_M"
   CMD("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
   CMD("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
   CMD("@VCHA_ALM_ALMACEN_ID") = clvcha_alm_almacen_id
   CMD("@VCHA_ART_ARTICULO_ID") = clVCHA_ART_ARTICULO_ID
   CMD("@INTE_ORS_ORDEN_SURTIDO") = clinte_ors_orden_surtido
   CMD("@FLOA_ORS_CANTIDAD_SURTIDA") = clFLOA_ORS_CANTIDAD_SURTIDA
   CMD("@FLOA_ORS_CANTIDAD_EMPACADA") = clFLOA_ORS_CANTIDAD_EMPACADA
   CMD("@FLOA_ORS_PRECIO") = clFLOA_ORS_PRECIO
   CMD("@CHAR_PED_TIPO") = clCHAR_PED_TIPO
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








