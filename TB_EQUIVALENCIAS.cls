VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_EQUIVALENCIAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_ART_ARTICULO_ID As String, clVCHA_EQU_ARTICULO_PROVEEDOR_ID As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL

Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    
If var_modifica_registro_equivalencia Then
    CMD.CommandText = "EQUIVALENCIAS_M"
Else
    CMD.CommandText = "EQUIVALENCIAS_I"
End If
    CMD("@VCHA_ART_ARTICULO_ID") = clVCHA_ART_ARTICULO_ID
    CMD("@VCHA_EQU_CODIGO_EQUIVALENTE") = clVCHA_EQU_ARTICULO_PROVEEDOR_ID
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
Public Function Eliminar(clVCHA_ART_ARTICULO_ID As String, clVCHA_EQU_ARTICULO_PROVEEDOR_ID As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
'On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "EQUIVALENCIAS_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@VCHA_ART_ARTICULO_ID") = clVCHA_ART_ARTICULO_ID
    CMD("@VCHA_EQU_CODIGO_EQUIVALENTE") = clVCHA_EQU_ARTICULO_PROVEEDOR_ID
CMD.execute                                         'Ejecutar el PA

SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function


