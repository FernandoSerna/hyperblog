VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_DET_FACTURAS_I"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clvcha_alm_almacen_id As String, clinte_fac_numero As Integer, _
clvcha_art_articulo_ID As String, clfloa_fac_cantidad As Double, clfloa_Fac_costo As Double, clfloa_fac_precio As Double, _
clfloa_fac_cantidad_devuelta As Double) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL


Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "DETALLE_FACTURAS_I"
   CMD("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
   CMD("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
   CMD("@VCHA_ALM_ALMACEN_ID") = clvcha_alm_almacen_id
   CMD("@INTE_FAC_NUMERO") = clinte_fac_numero
   CMD("@VCHA_ART_ARTICULO_ID") = clvcha_art_articulo_ID
   CMD("@FLOA_FAC_CANTIDAD") = clfloa_fac_cantidad
   CMD("@FLOA_FAC_COSTO") = clfloa_Fac_costo
   CMD("@FLOA_FAC_PRECIO") = clfloa_fac_precio
   CMD("@FLOA_FAC_CANTIDAD_DEVUELTA") = clfloa_fac_cantidad_devuelta
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








