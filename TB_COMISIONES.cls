VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_COMISIONES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command


'ANADIR UN REGISTRO
Public Function Anadir(clVcha_age_egente_id As String, clVcha_lin_linea_id As String, clVcha_com_limite_inferior As Variant, clVcha_com_limite_superior As Variant, clVcha_com_porcentaje As Variant) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL


Set CMD.ActiveConnection = cnn_importacion
CMD.CommandType = adCmdStoredProc
    
If var_modifica_registro_comision Then
    CMD.CommandText = "COMISIONES_M"
Else
    CMD.CommandText = "COMISIONES_I"
End If
    CMD("@Vcha_age_agente_id") = clVcha_age_egente_id
    CMD("@Vcha_lin_linea_id") = clVcha_lin_linea_id
    CMD("@floa_com_limite_inferior") = clVcha_com_limite_inferior
    CMD("@floa_com_limite_superior") = clVcha_com_limite_superior
    CMD("@floa_com_porcentaje") = clVcha_com_porcentaje
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
Public Function Eliminar(clVCHA_AGE_AGENTE_ID As String, clVcha_lin_linea_id As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn_importacion
CMD.CommandType = adCmdStoredProc
    
CMD.CommandText = "COMISIONES_E"
    CMD("@Vcha_age_agente_id") = clVCHA_AGE_AGENTE_ID
    CMD("@Vcha_lin_linea_id") = clVcha_lin_linea_id
CMD.execute



SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function

