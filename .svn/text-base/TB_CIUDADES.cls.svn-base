VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_CIUDADES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command



'ANADIR UN REGISTRO
Public Function Anadir(clVcha_pai_pais As String, clVcha_est_estado As String, clVcha_ciu_ciudad As String, clVcha_ciu_cve_tel_ciudad As String, clVcha_ciu_nombre As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL
Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    
If var_modifica_registro_ciudad Then
    CMD.CommandText = "CIUDADES_M"
Else
    CMD.CommandText = "CIUDADES_I"
End If
    CMD("@Vcha_pai_pais_id") = clVcha_pai_pais
    CMD("@VCHA_EST_ESTADO_id") = clVcha_est_estado
    CMD("@Vcha_ciu_ciudad_id") = clVcha_ciu_ciudad
    CMD("@Vcha_ciu_nombre") = clVcha_ciu_nombre
    CMD("@VCHA_CIU_CLAVE_TELEFONO") = clVcha_ciu_cve_tel_ciudad
    If Not var_modifica_registro_ciudad Then
       CMD("@clave_CIUDAD") = clVcha_ciu_ciudad
       CMD.execute
       var_ciudad_regreso = CMD("@clave_CIUDAD")
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
Public Function Eliminar(clVcha_ciu_ciudad_id As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
'On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "CIUDADES_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@Vcha_ciu_ciudad_id") = clVcha_ciu_ciudad_id
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

