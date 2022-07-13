VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_BITACORA_CLIENTES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVcha_cli_cliente_id As String, clVCHA_BIT_CAMPO As String, clVCHA_BIT_VALOR_ANTERIOR As String, clVCHA_BIT_VALOR_ACTUAL As String, clVcha_aud_usuario As String, clVcha_aud_maquina As String, clDtim_aud_fecha As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL

Set CMD.ActiveConnection = cnn_importacion
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "BITCLIENTES_I"
   CMD("@VCHA_CLI_CLAVE_ID") = clVcha_cli_cliente_id
   CMD("@VCHA_BIT_CAMPO") = clVCHA_BIT_CAMPO
   CMD("@CHAR_BIT_OPERACION") = var_operacion_bitacora
   CMD("@VCHA_BIT_VALOR_ANTERIOR") = clVCHA_BIT_VALOR_ANTERIOR
   CMD("@VCHA_BIT_VALOR_ACTUAL") = clVCHA_BIT_VALOR_ACTUAL
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

