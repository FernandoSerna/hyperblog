VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SP_CANCELA_FACTURAS_DEVOLUCION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clEMPRESA As String, clUNIDAD_ORGANIZACIONAL As String, clSERIE As String, clNUMERO As Double, clCLIENTE_NUEVO As String, clESTABLECIMIENTO_NUEVO As String, clUSUARIO As String, clMAQUINA As String) As Boolean

Dim rs As ADODB.recordSet, Cadena As String
Anadir = True
'On Error GoTo HELL


Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "SP_CANCELA_FACTURAS_DEVOLUCION"
   CMD("@EMPRESA") = clEMPRESA
   CMD("@UNIDAD_ORGANIZACIONAL") = clUNIDAD_ORGANIZACIONAL
   CMD("@SERIE") = clSERIE
   CMD("@NUMERO") = clNUMERO
   CMD("@CLIENTE_NUEVO") = clCLIENTE_NUEVO
   CMD("@ESTABLECIMIENTO_NUEVO") = clESTABLECIMIENTO_NUEVO
   CMD("@USUARIO") = clUSUARIO
   CMD("@MAQUINA") = clMAQUINA
   CMD.execute
   var_numero_embarque_regreso = CMD("@embarque")
   var_numero_movimiento_embarque_regreso = CMD("@folio")
   var_nombre_movimiento_embarque_regreso = CMD("@movimiento_calidad")
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













