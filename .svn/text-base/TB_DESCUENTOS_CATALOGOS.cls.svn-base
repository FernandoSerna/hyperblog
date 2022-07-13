VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_DESCUENTOS_CATALOGOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command


'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_CAN_CANAL_VENTA_ID As String, clINTE_DES_LIMITE_INFERIOR As Double, clINTE_DES_LIMITE_SUPERIOR As Double, clFLOA_DES_DESCUENTO As Double) As Boolean

Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    CMD.CommandText = "DESCUENTOS_CATALOGOS_I"
    CMD("@VCHA_CAN_CANAL_VENTA_ID") = clVCHA_CAN_CANAL_VENTA_ID
    CMD("@INTE_DES_LIMITE_INFERIOR") = clINTE_DES_LIMITE_INFERIOR
    CMD("@INTE_DES_LIMITE_SUPERIOR") = clINTE_DES_LIMITE_SUPERIOR
    CMD("@FLOA_DES_DESCUENTO") = clFLOA_DES_DESCUENTO
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
Public Function Eliminar(clVCHA_CAN_CANAL_VENTA_ID As String, clINTE_DES_LIMITE_INFERIOR As Double, clINTE_DES_LIMITE_SUPERIOR As Double, clFLOA_DES_DESCUENTO As Double) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
'On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "DESCUENTOS_CATALOGOS_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@VCHA_CAN_CANAL_VENTA_ID") = clVCHA_CAN_CANAL_VENTA_ID
    CMD("@INTE_DES_LIMITE_INFERIOR") = clINTE_DES_LIMITE_INFERIOR
    CMD("@INTE_DES_LIMITE_SUPERIOR") = clINTE_DES_LIMITE_SUPERIOR
    CMD("@FLOA_DES_DESCUENTO") = clFLOA_DES_DESCUENTO
CMD.execute                                         'Ejecutar el PA

Set CMD = Nothing                                   'Liberar Memoria
SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function


