VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_PLAZOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command

'ANADIR UN REGISTRO
Public Function Anadir(clVcha_pla_plazo_id As String, clVCHA_PLA_NOMBRE As String, clINTE_PLA_DIAS As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
If var_modifica_registro_plazo Then
    CMD.CommandText = "PLAZOS_M"
Else
    CMD.CommandText = "PLAZOS_I"
End If
    CMD("@VCHA_PLA_PLAZO_ID") = clVcha_pla_plazo_id
    CMD("@VCHA_PLA_NOMBRE") = clVCHA_PLA_NOMBRE
    CMD("@INTE_PLA_DIAS") = clINTE_PLA_DIAS
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
Public Function Eliminar(clVcha_pla_plazo_id As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "PLAZOS_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@VCHA_PLA_PLAZO_ID") = clVcha_pla_plazo_id
CMD.execute                                         'Ejecutar el PA
SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function

