VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_DEVOLUCIONES_MODIFICA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim CMD As New Command

Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clvcha_alm_almacen_id As String, clvcha_mov_movimiento_id As String, clinte_emo_numero As Integer, clvcha_art_articulo_id As String, _
clINTE_CDE_CAUSA_DEVOLUCION As Integer, clINTE_CDE_CAUSA_REAL As Integer, clCHAR_CDE_ESTATUS As String, clINTE_CDE_CONSECUTIVO As Integer, clVCHA_CDE_DESTINO As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL


Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "DEVOLUCIONES_M"
   CMD("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
   CMD("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
   CMD("@VCHA_ALM_ALMACEN_ID") = clvcha_alm_almacen_id
   CMD("@VCHA_MOV_MOVIMIENTO_ID") = clvcha_mov_movimiento_id
   CMD("@INTE_EMO_NUMERO") = clinte_emo_numero
   CMD("@VCHA_ART_ARTICULO_ID") = clvcha_art_articulo_id
   CMD("@INTE_CDE_CAUSA_DEVOLUCION") = clINTE_CDE_CAUSA_DEVOLUCION
   CMD("@INTE_CDE_CAUSA_REAL") = clINTE_CDE_CAUSA_REAL
   CMD("@CHAR_CDE_ESTATUS") = clCHAR_CDE_ESTATUS
   CMD("@INTE_CDE_CONSECUTIVO") = clINTE_CDE_CONSECUTIVO
   CMD("@VCHA_CDE_DESTINO") = clVCHA_CDE_DESTINO
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




