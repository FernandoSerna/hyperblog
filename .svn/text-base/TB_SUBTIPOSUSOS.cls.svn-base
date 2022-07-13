VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_SUBTIPOSUSOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command

'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_USO_USO_ID As String, clVcha_sus_subtipo_uso_id As String, clVcha_sus_nombre As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    
If var_modifica_registro_subtipouso Then
    CMD.CommandText = "SUBTIPOSUSOS_M"
Else
    CMD.CommandText = "SUBTIPOSUSOS_I"
End If
    CMD("@VCHA_USO_USO_ID") = clVCHA_USO_USO_ID
    CMD("@VCHA_SUS_SUBTIPO_USO_ID") = clVcha_sus_subtipo_uso_id
    CMD("@VCHA_SUS_NOMBRE") = clVcha_sus_nombre
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
'ELIMINAR
Public Function Eliminar(clVcha_sus_subtipo_uso_id As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL

Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "SUBTIPOSUSOS_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@Vcha_sus_subtipo_uso_id") = clVcha_sus_subtipo_uso_id
CMD.execute                                         'Ejecutar el PA


SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function

