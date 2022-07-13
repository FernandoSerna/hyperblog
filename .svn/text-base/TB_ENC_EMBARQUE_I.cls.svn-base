VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_ENC_EMBARQUE_I"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clINTE_EMB_EMBARQUE As Integer, clINTE_JAU_JAULA_ID As Integer, clVCHA_VEH_VEHICULO_ID As String, clVCHA_AGE_AGENTE_ID As String, _
clDTIM_EMB_FECHA_INICIO As Date, clDTIM_EMB_FECHA_FINAL As Date, clCHAR_EMB_ESTATUS As String, clVCHA_CHO_CHOFER_ID As String, clfloa_emb_cubicaje As Double) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL


Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "ENCABEZADO_EMBARQUES_I"
   CMD("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
   CMD("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
   CMD("@INTE_EMB_EMBARQUE") = clINTE_EMB_EMBARQUE
   CMD("@INTE_JAU_JAULA_ID") = clINTE_JAU_JAULA_ID
   CMD("@VCHA_VEH_VEHICULO_ID") = clVCHA_VEH_VEHICULO_ID
   CMD("@VCHA_AGE_AGENTE_ID") = clVCHA_AGE_AGENTE_ID
   CMD("@DTIM_EMB_FECHA_INICIO") = clDTIM_EMB_FECHA_INICIO
   CMD("@DTIM_EMB_FECHA_FINAL") = clDTIM_EMB_FECHA_FINAL
   CMD("@CHAR_EMB_ESTATUS") = clCHAR_EMB_ESTATUS
   CMD("@VCHA_CHO_CHOFER_ID") = clVCHA_CHO_CHOFER_ID
   CMD("@FLOA_EMB_CUBICAJE") = clfloa_emb_cubicaje
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













