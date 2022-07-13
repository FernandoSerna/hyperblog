VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_BLOQUEOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clVCHA_BLO_BLOQUEADO_POR As String, clDTIM_BLO_FECHA As String, clVCHA_USU_USUARIO_ID As String, clVCHA_BLO_MAQUINA As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL
If var_global_bloqueado = 1 Then
   Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "BLOQUEOS_I"
   CMD("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
   CMD("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
   CMD("@VCHA_BLO_BLOQUEADO_POR") = clVCHA_BLO_BLOQUEADO_POR
   CMD("@DTIM_BLO_FECHA") = clDTIM_BLO_FECHA
   CMD("@VCHA_USU_USUARIO_ID") = clVCHA_USU_USUARIO_ID
   CMD("@VCHA_BLO_MAQUINA") = clVCHA_BLO_MAQUINA
   CMD.execute
   Set CMD = Nothing
End If

If var_global_bloqueado = 0 Then
   Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "BLOQUEOS_E"
   CMD("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
   CMD("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
   CMD("@VCHA_BLO_BLOQUEADO_POR") = clVCHA_BLO_BLOQUEADO_POR
   CMD.execute
   Set CMD = Nothing
End If

SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Anadir = False
    GoTo SIGUE
End Function













