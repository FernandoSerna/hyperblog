VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_CANCELA_PEDIDOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim cmd As New Command



'ANADIR UN REGISTRO
Public Function Anadir(clINTE_PED_NUMERO As Variant, clCHAR_PED_ESTATUS As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL
Set cmd.ActiveConnection = cnn
cmd.CommandType = adCmdStoredProc
   cmd.CommandText = "cancela_PEDIDOS"
    cmd("@INTE_PED_NUMERO") = clINTE_PED_NUMERO
    cmd("@CHAR_PED_ESTATUS") = clCHAR_PED_ESTATUS
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




