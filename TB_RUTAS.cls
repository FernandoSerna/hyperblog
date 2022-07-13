VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_RUTAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command


'ANADIR UN REGISTRO
Public Function Anadir(CLVCHA_RUT_RUTA_ID As String, clVcha_rut_descripcion As String, _
                       clVCHA_AGE_AGENTE_ID As String, clVCHA_ZON_ZONA_ID As String, clINTE_RUT_TOLERANCIA As Integer) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL


Set CMD.ActiveConnection = cnn_importacion
CMD.CommandType = adCmdStoredProc
    
If var_modifica_registro_ruta Then
    CMD.CommandText = "RUTAS_M"
Else
    CMD.CommandText = "RUTAS_I"
End If
    CMD("@Vcha_rut_ruta_id") = CLVCHA_RUT_RUTA_ID
    CMD("@Vcha_rut_nombre") = clVcha_rut_descripcion
    CMD("@Vcha_age_agente_id") = clVCHA_AGE_AGENTE_ID
    CMD("@Vcha_zon_zona_id") = clVCHA_ZON_ZONA_ID
    CMD("@INTE_RUT_TOLERANCIA") = clINTE_RUT_TOLERANCIA
    
    
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
Public Function Eliminar(CLVCHA_RUT_RUTA_ID As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
'On Error GoTo HELL
Set CMD.ActiveConnection = cnn_importacion
CMD.CommandType = adCmdStoredProc
    
CMD.CommandText = "RUTAS_E"
    CMD("@Vcha_rut_ruta_id") = CLVCHA_RUT_RUTA_ID
CMD.execute

SIGUE:
On Error GoTo 0



Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function
