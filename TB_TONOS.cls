VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_TONOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_TON_TONO_ID As String, clVcha_ton_NOMBRE As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL

Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    
If var_modifica_registro_tono Then
    CMD.CommandText = "tonos_M"
Else
    CMD.CommandText = "tonos_I"
End If
    CMD("@Vcha_ton_tono_id") = clVCHA_TON_TONO_ID
    CMD("@Vcha_ton_nombre") = clVcha_ton_NOMBRE
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
Public Function Eliminar(clVCHA_TON_TONO_ID As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexi�n activa
CMD.CommandType = adCmdStoredProc                   'Aqu� le indico a ADO que se trata de un PA
    
CMD.CommandText = "tonos_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@Vcha_ton_tono_id") = clVCHA_TON_TONO_ID
CMD.execute                                         'Ejecutar el PA

SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function

