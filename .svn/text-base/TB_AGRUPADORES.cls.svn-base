VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_AGRUPADORES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVcha_fag_familia_agrupador_id As String, clVCHA_AGR_AGRUPADOR_ID As String, clVcha_agr_nombre As String, clVcha_agr_tipo As String, clfloa_agr_fraccion_arancelaria As String, clvcha_agr_pais As String, clvcha_agr_texto As String) As Boolean
Dim rs As ADODB.recordSet, Cadena As String
Anadir = True
On Error GoTo HELL

Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    
If var_modifica_registro_agrupadores Then
    CMD.CommandText = "AGRUPADORES_M"
Else
    CMD.CommandText = "AGRUPADORES_I"
End If
    CMD("@VCHA_FAG_FAMILIA_AGRUPADOR_ID") = clVcha_fag_familia_agrupador_id
    CMD("@Vcha_agr_agrupador_id") = clVCHA_AGR_AGRUPADOR_ID
    CMD("@Vcha_agr_NOMBRE") = clVcha_agr_nombre
    CMD("@Vcha_agr_tipo") = clVcha_agr_tipo
    CMD("@FLOA_AGR_FRACCION_aRANCELARIA") = clfloa_agr_fraccion_arancelaria
    CMD("@vcha_agr_pais") = clvcha_agr_pais
    CMD("@vcha_agr_texto") = clvcha_agr_texto
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
Public Function Eliminar(clVCHA_AGR_AGRUPADOR_ID As String) As Boolean
Dim rs As ADODB.recordSet, Cadena As String
Eliminar = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "AGRUPADORES_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@Vcha_agr_agrupador_id") = clVCHA_AGR_AGRUPADOR_ID
CMD.execute                                         'Ejecutar el PA

SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function

