VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_UNIDADESORGANIZACIONALES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 
Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clVcha_UOR_NOMBRE As String, clVcha_UOR_MAIL As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL

Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    
If var_modifica_registro_unidadorganizacional Then
    CMD.CommandText = "unidadesorganizacionales_M"
Else
    CMD.CommandText = "unidadesorganizacionales_I"
End If
    CMD("@Vcha_EMP_EMPRESA_id") = clVCHA_EMP_EMPRESA_ID
    CMD("@Vcha_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
    CMD("@Vcha_UOR_NOMBRE") = clVcha_UOR_NOMBRE
    CMD("@Vcha_UOR_MAIL") = clVcha_UOR_MAIL
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
Public Function Eliminar(clVCHA_UOR_UNIDAD_ID As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
'On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "UNIDADESORGANIZACIONALES_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
CMD("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
CMD.execute                                         'Ejecutar el PA

SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function


