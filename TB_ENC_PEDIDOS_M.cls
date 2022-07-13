VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_ENC_PEDIDOS_M"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim CMD As New Command



'ANADIR UN REGISTRO
Public Function Anadir(clvcha_emp_empresa_id As String, clvcha_uor_unidad_id As String, clvcha_alm_almacen_id As String, clINTE_PED_NUMERO As Variant, clCHAR_PED_ESTATUS As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL
Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "ENCABEZADO_PEDIDOS_M"
   CMD("@VCHA_EMP_EMPRESA_ID") = clvcha_emp_empresa_id
   CMD("@VCHA_UOR_UNIDAD_ID") = clvcha_uor_unidad_id
   CMD("@VCHA_ALM_ALMACEN_ID") = clvcha_alm_almacen_id
   CMD("@INTE_PED_NUMERO") = clINTE_PED_NUMERO
   CMD("@CHAR_PED_ESTATUS") = clCHAR_PED_ESTATUS
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


