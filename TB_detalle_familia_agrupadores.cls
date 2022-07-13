VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_detalle_familia_agrupadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVcha_fag_familia_agrupador_id As String, clVCHA_AGR_AGRUPADOR_ID As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL

Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    
If var_modifica_registro_familia_agrupador Then
    CMD.CommandText = "detalle_familia_agrupadores_M"
Else
    CMD.CommandText = "detalle_familia_agrupadores_I"
End If
    CMD("@VCHA_FAG_FAMILIA_AGRUPADOR_ID") = clVcha_fag_familia_agrupador_id
    CMD("@Vcha_agr_agrupador_id") = clVCHA_AGR_AGRUPADOR_ID
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
Public Function Eliminar(clVcha_fag_familia_agrupador_id As String, clVCHA_AGR_AGRUPADOR_ID As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "detalle_agrupadores_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@VCHA_FAG_FAMILIA_AGRUPADOR_ID") = clVcha_fag_familia_agrupador_id
    CMD("@VCHA_AGR_AGRUPADOR_ID") = clVCHA_AGR_AGRUPADOR_ID
CMD.execute                                         'Ejecutar el PA

SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function



