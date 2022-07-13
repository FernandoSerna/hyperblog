VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_NOTAS_ENVIO_ACTUALIZA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim cmd As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clVCHA_ALM_ALMACEN_ID As String, clINTE_NEN_NUMERO As Integer, clVCHA_ART_ARTICULO_ID As String, clFLOA_NEN_CANTIDAD_RECIBIDA As Variant, clCHAR_NEN_ESTATUS As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL

Set cmd.ActiveConnection = cnn
   cmd.CommandType = adCmdStoredProc
   cmd.CommandText = "NOTAS_ENVIO_M"
   cmd("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
   cmd("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
   cmd("@VCHA_ALM_ALMACEN_ID") = clVCHA_ALM_ALMACEN_ID
   cmd("@INTE_NEN_NUMERO") = clINTE_NEN_NUMERO
   cmd("@VCHA_ART_ARTICULO_ID") = clVCHA_ART_ARTICULO_ID
   cmd("@FLOA_NEN_CANTIDAD_RECIBIDA") = clFLOA_NEN_CANTIDAD_RECIBIDA
   cmd("@CHAR_NEN_ESTATUS") = clCHAR_NEN_ESTATUS
   cmd.execute
Set cmd = Nothing
SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Anadir = False
    GoTo SIGUE
End Function








