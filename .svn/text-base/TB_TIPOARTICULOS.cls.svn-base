VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_TIPOARTICULOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command



'ANADIR UN REGISTRO
Public Function Anadir(clVcha_tar_tipoarticulo_id As String, clVcha_tar_descripcion As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL

Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    
If var_modifica_registro_tipoarticulo Then
    CMD.CommandText = "TIPOARTICULOS_M"
Else
    CMD.CommandText = "TIPOARTICULOS_I"
End If
    CMD("@VCHA_TAR_TIPO_ARTICULO_ID") = clVcha_tar_tipoarticulo_id
    CMD("@Vcha_tar_nombre") = clVcha_tar_descripcion
    
    
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
Public Function Eliminar(clVcha_tar_tipoarticulo_id As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "TIPOARTICULOS_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@Vcha_tar_tipo_articulo_id") = clVcha_tar_tipoarticulo_id
CMD.execute                                         'Ejecutar el PA
SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function
