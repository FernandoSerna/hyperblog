VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tb_sublineas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVcha_lin_linea_id As String, clVcha_sli_sublinea_id As String, VCHA_SLI_NOMBRE As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL

Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    
If var_modifica_registro_sublinea Then
    CMD.CommandText = "SUBLINEAS_M"
Else
    CMD.CommandText = "SUBLINEAS_I"
End If
    CMD("@VCHA_LIN_LINEA_ID") = clVcha_lin_linea_id
    CMD("@VCHA_SLI_SUBLINEA_ID") = clVcha_sli_sublinea_id
    CMD("@VCHA_SLI_nombre") = VCHA_SLI_NOMBRE
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
Public Function Eliminar(clVcha_sli_sublinea_id As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "SUBLINEAS_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@VCHA_SLI_SUBLINEA_ID") = clVcha_sli_sublinea_id
CMD.execute                                         'Ejecutar el PA





SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function


