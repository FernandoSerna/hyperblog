VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_FAMILIASARTICULOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command

'ANADIR UN REGISTRO
Public Function Anadir(clVcha_far_familia_id As String, clVcha_far_descripcion As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL

Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    
If var_modifica_registro_familia_ariculos Then
    CMD.CommandText = "FAMILIASARTICULOS_M"
Else
    CMD.CommandText = "FAMILIASARTICULOS_I"
End If
    CMD("@VCHA_FAR_FAMILIA_ID") = clVcha_far_familia_id
    CMD("@VCHA_FAR_DESCRIPCION") = clVcha_far_descripcion
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
Public Function Eliminar(clVcha_far_familia_id As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "FAMILIASARTICULOS_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@Vcha_far_familia_id") = clVcha_far_familia_id
CMD.execute                                         'Ejecutar el PA
SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function
