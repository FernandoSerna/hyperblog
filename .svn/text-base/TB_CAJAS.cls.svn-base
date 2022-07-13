VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_CAJAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVcha_caj_caja_id As String, clVcha_caj_NOMBRE As String, clFLOA_CAJ_ANCHO As String, clFLOA_CAJ_ALTO As String, clFLOA_CAJ_LARGO As String) As Boolean
Dim rs As ADODB.recordSet, Cadena As String
Anadir = True
On Error GoTo HELL

Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    
If var_modifica_registro_caja Then
    CMD.CommandText = "CAJAS_M"
Else
    CMD.CommandText = "CAJAS_I"
End If
    CMD("@VCHA_CAJ_CAJA_ID") = clVcha_caj_caja_id
    CMD("@VCHA_CAJ_NOMBRE") = clVcha_caj_NOMBRE
    CMD("@FLOA_CAJ_ANCHO") = clFLOA_CAJ_ANCHO
    CMD("@FLOA_CAJ_ALTO") = clFLOA_CAJ_ALTO
    CMD("@FLOA_CAJ_LARGO") = clFLOA_CAJ_LARGO
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
Public Function Eliminar(clVcha_caj_caja_id As String) As Boolean
Dim rs As ADODB.recordSet, Cadena As String
Eliminar = True
'On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "CAJAS_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@VCHA_CAJ_CAJA_ID") = clVcha_caj_caja_id
CMD.execute                                         'Ejecutar el PA

SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function

