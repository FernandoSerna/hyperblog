VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_DESCUENTOS_PROMOCIONES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command


'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_CAN_CANAL_VENTA_ID As String, clVCHA_ART_ARTICULO_ID As String, clDTIM_DPR_FECHA_INICIO As String, clDTIM_DPR_FECHA_FIN As String, clFLOA_DPR_DESCUENTO As Double) As Boolean

Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    CMD.CommandText = "DESCUENTOS_PROMOCIONES_I"
    CMD("@VCHA_CAN_CANAL_VENTA_ID") = clVCHA_CAN_CANAL_VENTA_ID
    CMD("@VCHA_ART_ARTICULO_ID") = clVCHA_ART_ARTICULO_ID
    CMD("@DTIM_DPR_FECHA_INICIO") = clDTIM_DPR_FECHA_INICIO
    CMD("@DTIM_DPR_FECHA_FIN") = clDTIM_DPR_FECHA_FIN
    CMD("@FLOA_DPR_DESCUENTO") = clFLOA_DPR_DESCUENTO
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
Public Function Eliminar(clVCHA_CAN_CANAL_VENTA_ID As String, clVCHA_ART_ARTICULO_ID As String, clDTIM_DPR_FECHA_INICIO As String, clDTIM_DPR_FECHA_FIN As String, clFLOA_DPR_DESCUENTO As Double) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
'On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "DESCUENTOS_promociones_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@VCHA_CAN_CANAL_VENTA_ID") = clVCHA_CAN_CANAL_VENTA_ID
    CMD("@VCHA_ART_ARTICULO_ID") = clVCHA_ART_ARTICULO_ID
    CMD("@DTIM_DPR_FECHA_INICIO") = clDTIM_DPR_FECHA_INICIO
    CMD("@DTIM_DPR_FECHA_FIN") = clDTIM_DPR_FECHA_FIN
    CMD("@FLOA_DPR_DESCUENTO") = clFLOA_DPR_DESCUENTO
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



