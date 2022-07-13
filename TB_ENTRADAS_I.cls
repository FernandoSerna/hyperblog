VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_ENTRADAS_I"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim cmd As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clvcha_alm_almacen_id As String, clvcha_mov_movimiento_id As String, _
clINTE_ENT_NUMERO As Double, clTIPO_DOCUMENTO As String, clALMACEN_ORIGEN As String, clFOLIO_ENVIADO As Double) As Boolean
Dim rs As ADODB.recordSet, Cadena As String
Anadir = True
'On Error GoTo HELL




Set cmd.ActiveConnection = cnn
   cmd.CommandType = adCmdStoredProc
   cmd.CommandText = "entradas_I"
   'MsgBox clVCHA_EMP_EMPRESA_ID + clVCHA_UOR_UNIDAD_ID + clvcha_alm_almacen_id + clvcha_mov_movimiento_id + CStr(clINTE_ENT_NUMERO) + clTIPO_DOCUMENTO + clALMACEN_ORIGEN + CStr(clFOLIO_ENVIADO)
   cmd("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
   cmd("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
   cmd("@VCHA_ALM_ALMACEN_ID") = clvcha_alm_almacen_id
   cmd("@VCHA_MOV_MOVIMIENTO_ID") = clvcha_mov_movimiento_id
   cmd("@INTE_ENT_NUMERO") = clINTE_ENT_NUMERO
   cmd("@TIPO_DOCUMENTO") = clTIPO_DOCUMENTO
   cmd("@ALMACEN_ORIGEN") = clALMACEN_ORIGEN
   cmd("@FOLIO_ENVIADO") = clFOLIO_ENVIADO
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














