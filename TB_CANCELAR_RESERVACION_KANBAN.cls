VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_CANCELAR_RESERVACION_KANBAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clvar_almacen As String, clp_tipo_movimiento_id As String, clp_numero_movimiento As Double, clvar_kanban_id As String, clvar_exito As String, clvar_mensaje As String) As Boolean

Dim rs As ADODB.recordSet, Cadena As String
Anadir = True
'On Error GoTo HELL



Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "sp_cancelar_reservacion_kanban"
  

   CMD("@p_almacen_id") = clvar_almacen
   CMD("@p_tipo_movimiento_id") = clp_tipo_movimiento_id
   CMD("@p_numero_movimiento") = clp_numero_movimiento
   CMD("@p_kanban_id") = clvar_kanban_id
   CMD("@x_exito") = ""
   CMD("@x_mensaje") = ""
   CMD.execute
   
   var_kanban_exito = CMD("@x_exito")
   var_kanban_mensaje = CMD("@x_mensaje")
   
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















