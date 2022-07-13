VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_DEVOLUCIONES_ALMACEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim cmd As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clvcha_alm_almacen_id As String, clvcha_mov_movimiento_id As String, clinte_emo_numero As Double, clVcha_art_articulo_id As String, clINTE_CDE_CONSECUTIVO As Integer, clVCHA_CDE_DESTINO As String, clINTE_CDE_NUMERO_DESTINO As Double) As Boolean
Dim rs As ADODB.recordSet, Cadena As String
Anadir = True
'On Error GoTo HELL


Set cmd.ActiveConnection = cnn
   cmd.CommandType = adCmdStoredProc
   cmd.CommandText = "DEVOLUCIONES_ALMACEN"
   cmd("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
   cmd("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
   cmd("@VCHA_ALM_ALMACEN_ID") = clvcha_alm_almacen_id
   cmd("@VCHA_MOV_MOVIMIENTO_ID") = clvcha_mov_movimiento_id
   cmd("@INTE_EMO_NUMERO") = clinte_emo_numero
   cmd("@VCHA_ART_ARTICULO_ID") = clVcha_art_articulo_id
   cmd("@INTE_CDE_CONSECUTIVO") = clINTE_CDE_CONSECUTIVO
   cmd("@VCHA_CDE_DESTINO") = clVCHA_CDE_DESTINO
   cmd("@INTE_CDE_NUMERO_DESTINO") = clINTE_CDE_NUMERO_DESTINO
   
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



















