VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_ESTADO_CUENTA_INSERTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Public MensajeError As String
Dim CMD As New Command


'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_ECU_SERIE_CARGO As String, clVCHA_ECU_MOVIMIENTO_CARGO As String, clINTE_ECU_NUMERO_CARGO As Double, clVCHA_ECU_SERIE_ABONO As String, clVCHA_ECU_MOVIMIENTO_ABONO As String, clINTE_ECU_NUMERO_ABONO As Double, clFLOA_ECU_IMPORTE_CARGO As Double, clFLOA_ECU_IMPORTE_ABONO As Double) As Boolean

Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    CMD.CommandText = "ESTADO_CUENTA_I"
    CMD("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
    CMD("@VCHA_UOR_UNIDAD_ID") = var_unidad_organizacional
    CMD("@VCHA_ECU_SERIE_CARGO") = clVCHA_ECU_SERIE_CARGO
    CMD("@VCHA_ECU_MOVIMIENTO_CARGO") = clVCHA_ECU_MOVIMIENTO_CARGO
    CMD("@INTE_ECU_NUMERO_CARGO") = clINTE_ECU_NUMERO_CARGO
    CMD("@VCHA_ECU_SERIE_ABONO") = clVCHA_ECU_SERIE_ABONO
    CMD("@VCHA_ECU_MOVIMIENTO_ABONO") = clVCHA_ECU_MOVIMIENTO_ABONO
    CMD("@INTE_ECU_NUMERO_ABONO") = clINTE_ECU_NUMERO_ABONO
    CMD("@FLOA_ECU_IMPORTE_CARGO") = clFLOA_ECU_IMPORTE_CARGO
    CMD("@FLOA_ECU_IMPORTE_ABONO") = clFLOA_ECU_IMPORTE_ABONO
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
