VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_TEM_KARDEX_INV_INI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clMAXIMO_KARDEX As Integer, clALMACEN As String, clUSUARIO As String, clMAQUINA As String, clFECHA_INICIO As String, clFECHA_FINAL As String) As Boolean
Dim rs As ADODB.recordSet, Cadena As String
Anadir = True
'On Error GoTo HELL
cnn.CommandTimeout = 360
Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "TEM_KARDEX_INVENTARIO_INICIAL"
   CMD("@MAXIMO_KARDEX") = clMAXIMO_KARDEX
   CMD("@ALMACEN") = clALMACEN
   CMD("@ARTICULO") = ""
   CMD("@USUARIO") = clUSUARIO
   CMD("@MAQUINA") = clMAQUINA
   CMD("@FECHA_INICIO") = clFECHA_INICIO
   CMD("@FECHA_FINAL") = clFECHA_FINAL
   cnn.CommandTimeout = 360
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





