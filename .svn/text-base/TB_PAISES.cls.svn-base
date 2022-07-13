VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_PAISES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command


'ANADIR UN REGISTRO
Public Function Anadir(clVcha_pai_pais As String, clVcha_pai_nombre As String, clVcha_pai_cve_tel_pais As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    
If var_modifica_registro_pais Then
    CMD.CommandText = "PAISES_M"
Else
    CMD.CommandText = "PAISES_I"
End If
    CMD("@Vcha_pai_pais_id") = clVcha_pai_pais
    CMD("@Vcha_pai_nombre") = clVcha_pai_nombre
    CMD("@VCHA_PAI_CLAVE_TELEFONO") = clVcha_pai_cve_tel_pais
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
Public Function Eliminar(clVcha_pai_pais As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "PAISES_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@Vcha_pai_pais_id") = clVcha_pai_pais
CMD.execute                                         'Ejecutar el PA

Set CMD = Nothing                                   'Liberar Memoria
SIGUE:
On Error GoTo 0

Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function
