VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_NEGADO_I"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clvcha_alm_almacen_id As String, clinte_ors_orden_surtido As Double, clVCHA_CNE_CAUSA_ID As String, clVCHA_ART_ARTICULO_ID As String, clFLOA_NEG_CANTIDAD As String, _
clVcha_aud_usuario As String, clVcha_aud_maquina As String, clINTE_PED_NUMERO As Double, clFLOA_NEG_PRECIO As Double) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL



Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "NEGADO_I"
   CMD("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
   CMD("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
   CMD("@VCHA_ALM_ALMACEN_ID") = clvcha_alm_almacen_id
   CMD("@INTE_ORS_ORDEN_SURTIDO") = clinte_ors_orden_surtido
   CMD("@VCHA_CNE_CAUSA_ID") = clVCHA_CNE_CAUSA_ID
   CMD("@VCHA_ART_ARTICULO_ID") = clVCHA_ART_ARTICULO_ID
   CMD("@FLOA_NEG_CANTIDAD") = clFLOA_NEG_CANTIDAD
   CMD("@VCHA_AUD_USUARIO") = clVcha_aud_usuario
   CMD("@VCHA_AUD_MAQUINA") = clVcha_aud_maquina
   CMD("@DTIM_AUD_FECHA") = Date
   CMD("@INTE_PED_NUMERO") = clINTE_PED_NUMERO
   CMD("@FLOA_NEG_PRECIO") = clFLOA_NEG_PRECIO
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



















