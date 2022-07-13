VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_ES_UN_KANBAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clkanban_id As String, clvar_es_un_kanban As String, clvar_almacen_id As String, clvar_Articulo_id As String, clvar_exito As String, clvar_mensaje As String) As Boolean

Dim rs As ADODB.recordSet, Cadena As String
Anadir = True
'On Error GoTo HELL



Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "SP_ES_UN_KANBAN"
   
    
   CMD("@p_kanban_id") = clkanban_id
   CMD("@x_es_un_kanban") = clvar_es_un_kanban
   CMD("@x_almacen_id") = clvar_almacen_id
   CMD("@x_articulo_id") = clvar_Articulo_id
   CMD("@x_exito") = clvar_exito
   CMD("@x_mensaje") = clvar_mensaje
   CMD.execute
   
   var_kanban_es_un_kanban = CMD("@x_es_un_kanban")
   var_kanban_almacen_id = CMD("@x_almacen_id")
   var_kanban_articulo_id = CMD("@x_articulo_id")
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













