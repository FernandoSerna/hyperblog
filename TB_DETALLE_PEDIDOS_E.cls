VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_DETALLE_PEDIDOS_E"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clvcha_emp_empresa_id As String, clvcha_uor_unidad_id As String, clvcha_alm_almacen_id As String, clINTE_PED_NUMERO As Variant, clvcha_art_articulo_ID As String, clFLOA_PED_CANTIDAD As Variant, clFLOA_PED_CANTIDAD_SURTIDA As Variant) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL

Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "DETALLE_PEDIDOS_E"
   CMD("@VCHA_EMP_EMPRESA_ID") = clvcha_emp_empresa_id
   CMD("@VCHA_UOR_UNIDAD_ID") = clvcha_uor_unidad_id
   CMD("@VCHA_ALM_ALMACEN_ID") = clvcha_alm_almacen_id
   CMD("@INTE_PED_NUMERO") = clINTE_PED_NUMERO
   CMD("@VCHA_ART_ARTICULO_ID") = clvcha_art_articulo_ID
   CMD("@FLOA_PED_CANTIDAD") = clFLOA_PED_CANTIDAD
   CMD("@FLOA_PED_CANTIDAD_SURTIDA") = clFLOA_PED_CANTIDAD_SURTIDA
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









