VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_CAUSAS_DEVOLUCION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clINTE_CDE_CAUSA_ID As Integer, clVcha_CDE_NOMBRE As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL

Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    
If var_modifica_registro_causas_devolucion Then
    CMD.CommandText = "causas_devolucion_M"
Else
    CMD.CommandText = "causas_devolucion_I"
End If
    CMD("@INTE_CDE_CAUSA_id") = clINTE_CDE_CAUSA_ID
    CMD("@Vcha_CDE_nombre") = clVcha_CDE_NOMBRE
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
Public Function Eliminar(clINTE_CDE_CAUSA_ID As Integer) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexi�n activa
CMD.CommandType = adCmdStoredProc                   'Aqu� le indico a ADO que se trata de un PA
    
CMD.CommandText = "causas_devolucion_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@INTE_CDE_CAUSA_id") = clINTE_CDE_CAUSA_ID
CMD.execute                                         'Ejecutar el PA

SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function
