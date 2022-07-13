VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_DETALLE_PAQUETES_M"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clINTE_PAQ_NUMERO As Integer, clinte_paq_caja As Integer, clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clvcha_alm_almacen_id As String, clCHAR_PAQ_ESTATUS As String, clVCHA_PAQ_MOVIMIENTO_DESTINO As String, clINTE_PAQ_NUMERO_DESTINO As Integer) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL

Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "DETALLE_PAQUETES_M"
   CMD("@INTE_PAQ_NUMERO") = clINTE_PAQ_NUMERO
   CMD("@INTE_PAQ_CAJA") = clinte_paq_caja
   CMD("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
   CMD("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
   CMD("@VCHA_ALM_ALMACEN_ID") = clvcha_alm_almacen_id
   CMD("@CHAR_PAQ_ESTATUS") = clCHAR_PAQ_ESTATUS
   CMD("@VCHA_PAQ_MOVIMIENTO_DESTINO") = clVCHA_PAQ_MOVIMIENTO_DESTINO
   CMD("@INTE_PAQ_NUMERO_DESTINO") = clINTE_PAQ_NUMERO_DESTINO
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










