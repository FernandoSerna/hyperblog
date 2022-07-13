VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_ZONAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_ZON_ZONA_ID As String, clVcha_zon_descripcion As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL

Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    
If var_modifica_registro_zona Then
    CMD.CommandText = "ZONAS_M"
Else
    CMD.CommandText = "ZONAS_I"
End If
    CMD("@Vcha_zon_zona_id") = clVCHA_ZON_ZONA_ID
    CMD("@Vcha_zon_descripcion") = clVcha_zon_descripcion
    
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
Public Function Eliminar(clVCHA_ZON_ZONA_ID As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "ZONAS_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@Vcha_zon_zona_id") = clVCHA_ZON_ZONA_ID
CMD.execute                                         'Ejecutar el PA





SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function

