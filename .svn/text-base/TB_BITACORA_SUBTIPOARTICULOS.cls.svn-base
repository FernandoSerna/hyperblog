VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_BITACORA_SUBTIPOARTICULOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim cmd As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVcha_sta_subtipoarticulo_id As String, clVCHA_BIT_CAMPO As String, clvar_operacion_bitacora As String, clVCHA_BIT_VALOR_ANTERIOR As String, clVCHA_BIT_VALOR_ACTUAL As String, clVcha_aud_usuario As String, clVcha_aud_maquina As String, clDtim_aud_fecha As String) As Boolean
Dim rs As adodb.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL

Set cmd.ActiveConnection = cnn
   cmd.CommandType = adCmdStoredProc
   cmd.CommandText = "BITSUBTIPOARTICULOS_I"
   cmd("@VCHA_STA_SUBTIPOARTICULO_ID") = clVcha_sta_subtipoarticulo_id
   cmd("@VCHA_BIT_CAMPO") = clVCHA_BIT_CAMPO
   cmd("@CHAR_BIT_OPERACION") = clvar_operacion_bitacora
   cmd("@VCHA_BIT_VALOR_ANTERIOR") = clVCHA_BIT_VALOR_ANTERIOR
   cmd("@VCHA_BIT_VALOR_ACTUAL") = clVCHA_BIT_VALOR_ACTUAL
   cmd("@VCHA_AUD_USUARIO") = clVcha_aud_usuario
   cmd("@VCHA_AUD_MAQUINA") = clVcha_aud_maquina
   cmd("@DTIM_AUD_FECHA") = clDtim_aud_fecha
   cmd.execute
Set cmd = Nothing
SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Anadir = False
    GoTo SIGUE
End Function







