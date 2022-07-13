VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_MUNICIPIOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command



'ANADIR UN REGISTRO
Public Function Anadir(clVcha_pai_pais As String, clVcha_est_estado As String, clVcha_mun_municipio_id As String, clvcha_mun_nombre As String, clVcha_mun_cve_tel_ciudad As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL
Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    
If var_modifica_registro_municipio Then
    CMD.CommandText = "MUNICIPIOS_M"
Else
    CMD.CommandText = "MUNICIPIOS_I"
End If
    CMD("@Vcha_pai_pais_id") = clVcha_pai_pais
    CMD("@VCHA_EST_ESTADO_id") = clVcha_est_estado
    CMD("@Vcha_MUN_MUNICIPIO_id") = clVcha_mun_municipio_id
    CMD("@Vcha_MUN_nombre") = clvcha_mun_nombre
    CMD("@VCHA_MUN_CLAVE_TELEFONO") = clVcha_mun_cve_tel_ciudad
    If Not var_modifica_registro_municipio Then
       CMD("@clave_MUNICIPIO") = clVcha_mun_municipio_id
       CMD.execute
       var_municipio_regreso = CMD("@clave_MUNICIPIO")
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
    
CMD.CommandText = "MUNICIPIOS_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@vcha_mun_municipio_id") = clVcha_ciu_ciudad_id
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


