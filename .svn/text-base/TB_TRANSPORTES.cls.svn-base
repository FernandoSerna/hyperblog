VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_TRANSPORTES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_TRN_TRANSPORTE_ID As String, clVCHA_TRN_NOMBRE As String, clCHAR_TRN_TIPO As String, clFLOA_TRN_ANCHO As String, clFLOA_TRN_ALTO As String, clFLOA_TRN_LARGO As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL

Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    
If var_modifica_registro_transporte Then
    CMD.CommandText = "TRANSPORTES_M"
Else
    CMD.CommandText = "TRANSPORTES_I"
End If
    CMD("@VCHA_TRN_TRANSPORTE_ID") = clVCHA_TRN_TRANSPORTE_ID
    CMD("@VCHA_TRN_NOMBRE") = clVCHA_TRN_NOMBRE
    CMD("@CHAR_TRN_TIPO") = clCHAR_TRN_TIPO
    CMD("@FLOA_TRN_ANCHO") = clFLOA_TRN_ANCHO
    CMD("@FLOA_TRN_ALTO") = clFLOA_TRN_ALTO
    CMD("@FLOA_TRN_LARGO") = clFLOA_TRN_LARGO
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
Public Function Eliminar(clVCHA_TRN_TRANSPORTE_ID As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "TRANSPORTES_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@VCHA_TRN_TRANSPORTE_ID") = clVCHA_TRN_TRANSPORTE_ID
CMD.execute                                         'Ejecutar el PA

SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function


