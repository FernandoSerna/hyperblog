VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_DETALLE_DEVOLUCION_RECHAZO_I"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clvcha_alm_almacen_id As String, clvcha_mov_movimiento_id As String, clinte_emo_numero As Double, clVCHA_ART_ARTICULO_ID As String, _
 clINTE_CDE_CONSECUTIVO As Integer, clINTE_CRE_CAUSA_ID As Integer) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL

Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   If var_tipo_detalle_devolucion = 0 Then
      CMD.CommandText = "DETALLE_DEVOLUCION_RECHAZO_I"
   End If
   If var_tipo_detalle_devolucion = 1 Then
      CMD.CommandText = "DETALLE_DEVOLUCION_RECHAZO_E"
   End If
   CMD("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
   CMD("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
   CMD("@VCHA_ALM_ALMACEN_ID") = clvcha_alm_almacen_id
   CMD("@VCHA_MOV_MOVIMIENTO_ID") = clvcha_mov_movimiento_id
   CMD("@INTE_EMO_NUMERO") = clinte_emo_numero
   CMD("@VCHA_ART_ARTICULO_ID") = clVCHA_ART_ARTICULO_ID
   CMD("@INTE_CDE_CONSECUTIVO") = clINTE_CDE_CONSECUTIVO
   CMD("@INTE_CRE_CAUSA_ID") = clINTE_CRE_CAUSA_ID
   
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












