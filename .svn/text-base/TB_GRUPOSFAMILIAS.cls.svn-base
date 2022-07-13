VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_GRUPOSFAMILIAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim cmd As New Command


'ANADIR UN REGISTRO
Public Function Anadir(clVcha_gru_cve_grupo As String, clVcha_gru_nombre As String, clMone_gru_limite_credito As String, clMone_gru_cobranza As String, clChar_gru_clasificacion As String, clVcha_aud_usuario As String, clVcha_aud_maquina As String, clDate_gru_fecha_mod As String) As Boolean
Dim rs As adodb.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL
Set cmd.ActiveConnection = cnn                      'Esta es la conexión activa
cmd.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
If var_modifica_registro Then
    cmd.CommandText = "GRUPOSFAMILIAS_M"                         'Abrir Procedimiento Almacenado para Actualizar Cambios
Else
    cmd.CommandText = "GRUPOSFAMILIAS_I"                         'Abrir Procedimiento Almacenado y Agregar Banco
End If
    cmd("@Vcha_gru_cve_grupo") = clVcha_gru_cve_grupo
    cmd("@Vcha_gru_nombre") = clVcha_gru_nombre
    cmd("@Mone_gru_limite_credito") = clMone_gru_limite_credito
    cmd("@Mone_gru_cobranza") = clMone_gru_cobranza
    cmd("@Char_gru_clasificacion") = clChar_gru_clasificacion
    cmd("@Vcha_aud_usuario") = clVcha_aud_usuario
    cmd("@Vcha_aud_maquina") = clVcha_aud_maquina
    cmd("@date_gru_fecha_mod") = clDate_gru_fecha_mod
cmd.execute                                         'Ejecutar el PA
Set cmd = Nothing                                   'Liberar Memoria




SIGUE:
On Error GoTo 0

Exit Function
HELL:
    MensajeError = Err.Description
    Anadir = False
    GoTo SIGUE
End Function
'ELIMINAR
Public Function Eliminar(clVcha_gru_cve_grupo As String) As Boolean
Dim rs As adodb.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set cmd.ActiveConnection = cnn                      'Esta es la conexión activa
cmd.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
cmd.CommandText = "GRUPOSFAMILIAS_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
cmd("@Vcha_gru_cve_grupo") = clVcha_gru_cve_grupo
cmd.execute                                         'Ejecutar el PA
SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
    
End Function
