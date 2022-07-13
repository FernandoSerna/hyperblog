VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_ESTADOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command

'ANADIR UN REGISTRO
Public Function Anadir(clVcha_pai_pais As String, clVcha_est_estado As String, clVCHA_EST_NOMBRE As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
If var_modifica_registro_estado Then
    CMD.CommandText = "ESTADOS_M"
Else
    CMD.CommandText = "ESTADOS_I"
End If
    CMD("@Vcha_pai_pais_id") = clVcha_pai_pais
    CMD("@Vcha_est_estado_id") = clVcha_est_estado
    CMD("@Vcha_est_nombre") = clVCHA_EST_NOMBRE
    If Not var_modifica_registro_estado Then
       CMD("@clave_estado") = clVcha_est_estado
       CMD.execute
       var_estado_regreso = CMD("@clave_estado")
    Else
       CMD.execute
    End If
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
Public Function Eliminar(clVcha_est_estado As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "ESTADOS_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@Vcha_est_estado_id") = clVcha_est_estado
CMD.execute                                         'Ejecutar el PA
SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function
