VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_ENC_FACTURA_M"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clinte_fac_numero As Integer, clchar_fac_estatus As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL




Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "ENCABEZADO_FACTURAS_M"
   CMD("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
   CMD("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
   CMD("@INTE_FAC_NUMERO") = clinte_fac_numero
   CMD("@CHAR_FAC_ESTATUS") = clchar_fac_estatus
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








