VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_NIVEL_SURTIDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clFECHA_INICIO As Date, clFECHA_FIN As Date, clUSUARIO As String, clMAQUINA As String) As Boolean

Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL


Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "TB_TEMP_NIVEL_SURTIDO"
   CMD("@VCHA_EMP_EMPRESA_ID") = var_empresa
   CMD("@VCHA_UOR_UNIDAD_ID") = var_unidad_organizacional
   CMD("@FECHA_INICIO") = clFECHA_INICIO
   CMD("@FECHA_FIN") = clFECHA_FIN
   CMD("@USUARIO") = usuario
   CMD("@MAQUINA") = maquina
   CMD.execute
   var_numero_nivel_surtido = CMD("@numero")
Set CMD = Nothing
SIGUE:
On Error GoTo 0
Exit Function
HELL:
   x = Err.Number
   If x = -2147467259 Then
      On Error GoTo HELL:
      Sleep 30000
      CMD.execute
   Else
      MensajeError = Err.Description
      Anadir = False
      MsgBox MensajeError
   End If
   GoTo SIGUE
End Function










