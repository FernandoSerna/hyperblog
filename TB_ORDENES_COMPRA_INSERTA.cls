VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_ORDENES_COMPRA_INSERTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clINTE_OCO_NUMERO As Integer, clDTIM_OCO_FECHA As Date, clvcha_alm_almacen_id As String, clVcha_pro_proveedor_id As String, _
clvcha_art_articulo_ID As String, clFLOA_OCO_CANTIDAD As Variant, clFLOA_OCO_COSTO As Variant, clFLOA_OCO_CANTIDAD_RECIBIDA As Variant) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL




Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "ORDENES_COMPRA_I"
   CMD("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
   CMD("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
   CMD("@VCHA_ALM_ALMACEN_ID") = clvcha_alm_almacen_id
   CMD("@INTE_OCO_NUMERO") = clINTE_OCO_NUMERO
   CMD("@DTIM_OCO_FECHA") = clDTIM_OCO_FECHA
   CMD("@VCHA_PRO_PROVEEDOR_ID") = clVcha_pro_proveedor_id
   CMD("@VCHA_ART_ARTICULO_ID") = clvcha_art_articulo_ID
   CMD("@FLOA_OCO_CANTIDAD") = clFLOA_OCO_CANTIDAD
   CMD("@FLOA_OCO_COSTO") = clFLOA_OCO_COSTO
   CMD("@FLOA_OCO_CANTIDAD_RECIBIDA") = clFLOA_OCO_CANTIDAD_RECIBIDA
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















