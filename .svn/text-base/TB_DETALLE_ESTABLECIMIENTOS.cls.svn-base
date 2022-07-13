VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_DETALLE_ESTABLECIMIENTOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim cmd As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVcha_clp_clave_id As String, clvcha_esb_establecimiento_ID As String) As Boolean
Dim rs As ADODB.recordSet, Cadena As String
Anadir = True
On Error GoTo HELL

Set cmd.ActiveConnection = cnn_distribucion
cmd.CommandType = adCmdStoredProc
    
If var_modifica_registro_detalle_establecimiento Then
    cmd.CommandText = "detalle_establecimientos_M"
Else
    cmd.CommandText = "detalle_establecimientos_I"
End If
    cmd("@Vcha_esb_establecimiento_id") = clvcha_esb_establecimiento_ID
    cmd("@Vcha_cli_clave_id") = clVcha_clp_clave_id
cmd.execute
Set cmd = Nothing
SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Anadir = False
    GoTo SIGUE
End Function
'ELIMINAR
Public Function Eliminar(clVcha_clp_clave_id As String, clvcha_esb_establecimiento_ID As String) As Boolean
Dim rs As ADODB.recordSet, Cadena As String
Eliminar = True
On Error GoTo HELL
Set cmd.ActiveConnection = cnn                      'Esta es la conexión activa
cmd.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
cmd.CommandText = "detalle_establecimientos_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    cmd("@Vcha_esb_establecimiento_id") = clvcha_esb_establecimiento_ID
    cmd("@Vcha_cli_clave_id") = clVcha_clp_clave_id
cmd.execute                                         'Ejecutar el PA

SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function



