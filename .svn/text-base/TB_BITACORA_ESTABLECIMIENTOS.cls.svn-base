VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_BITACORA_ESTABLECIMIENTOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_TIT_TITULAR_ID As String, clvcha_esb_establecimiento_ID As String, clVCHA_BIT_CAMPO As String, clvar_operacion_bitacora As String, clVCHA_BIT_VALOR_ANTERIOR As String, clVCHA_BIT_VALOR_ACTUAL As String, clVcha_aud_usuario As String, clVcha_aud_maquina As String, clDtim_aud_fecha As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL

Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "BITESTABLECIMIENTOS_I"
   CMD("@VCHA_TIT_TITULAR_ID") = clVCHA_TIT_TITULAR_ID
   CMD("@VCHA_ESB_ESTABLECIMIENTO_ID") = clvcha_esb_establecimiento_ID
   CMD("@VCHA_BIT_CAMPO") = clVCHA_BIT_CAMPO
   CMD("@CHAR_BIT_OPERACION") = clvar_operacion_bitacora
   CMD("@VCHA_BIT_VALOR_ANTERIOR") = clVCHA_BIT_VALOR_ANTERIOR
   CMD("@VCHA_BIT_VALOR_ACTUAL") = Trim(clVCHA_BIT_VALOR_ACTUAL)
   CMD("@VCHA_AUD_USUARIO") = clVcha_aud_usuario
   CMD("@VCHA_AUD_MAQUINA") = clVcha_aud_maquina
   CMD("@DTIM_AUD_FECHA") = clDtim_aud_fecha
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




