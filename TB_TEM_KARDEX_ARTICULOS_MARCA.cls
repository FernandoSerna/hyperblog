VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_TEM_KARDEX_ARTICULOS_MARCA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clMAXIMO_KARDEX As Integer, clUSUARIO As String, clMAQUINA As String, clARTICULO As String, clMARCA As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL

Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "TEM_KARDEX_ARTICULOS_MARCA"
   CMD("@MAXIMO_KARDEX") = clMAXIMO_KARDEX
   CMD("@USUARIO") = clUSUARIO
   CMD("@MAQUINA") = clMAQUINA
   CMD("@ARTICULO") = clARTICULO
   CMD("@MARCA") = clMARCA
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




