VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_DEVOLUCIONES_INSERTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clvcha_alm_almacen_id As String, clvcha_mov_movimiento_id As String, clinte_emo_numero As Double, clVCHA_ART_ARTICULO_ID As String, _
clINTE_CDE_CAUSA_DEVOLUCION As Integer, clINTE_CDE_CAUSA_REAL As Integer, clCHAR_CDE_ESTATUS As String, clINTE_CDE_CONSECUTIVO As Integer, clVCHA_CDE_DESTINO As String, clFLOA_SAL_COSTO As Double, clFLOA_SAL_PRECIO As Double, _
clFLOA_CDE_DESCUENTO_1 As Double, clFLOA_CDE_DESCUENTO_2 As Double, clFLOA_CDE_DESCUENTO_3 As Double, clFLOA_CDE_IVA As Double, clINTE_FAC_FACTURA As Double, clVCHA_CDE_REFERENCIA As String, clVCHA_MON_MONEDA_ID As String, clFLOA_DEV_TIPO_CAMBIO As Double, clVCHA_SER_SERIE_ID As String, clINTE_DEV_AÑO As Integer) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL


Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "DEVOLUCIONES_I"
   CMD("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
   CMD("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
   CMD("@VCHA_ALM_ALMACEN_ID") = clvcha_alm_almacen_id
   CMD("@VCHA_MOV_MOVIMIENTO_ID") = clvcha_mov_movimiento_id
   CMD("@INTE_EMO_NUMERO") = clinte_emo_numero
   CMD("@VCHA_ART_ARTICULO_ID") = clVCHA_ART_ARTICULO_ID
   CMD("@CHAR_CDE_ESTATUS") = clCHAR_CDE_ESTATUS
   CMD("@INTE_CDE_CONSECUTIVO") = clINTE_CDE_CONSECUTIVO
   CMD("@VCHA_CDE_DESTINO") = clVCHA_CDE_DESTINO
   CMD("@FLOA_CDE_COSTO") = clFLOA_SAL_COSTO
   CMD("@FLOA_CDE_PRECIO") = clFLOA_SAL_PRECIO
   CMD("@FLOA_CDE_DESCUENTO_1") = clFLOA_CDE_DESCUENTO_1
   CMD("@FLOA_CDE_DESCUENTO_2") = clFLOA_CDE_DESCUENTO_2
   CMD("@FLOA_CDE_DESCUENTO_3") = clFLOA_CDE_DESCUENTO_3
   CMD("@FLOA_CDE_IVA") = clFLOA_CDE_IVA
   CMD("@INTE_FAC_FACTURA") = clINTE_FAC_FACTURA
   CMD("@VCHA_CDE_REFERENCIA") = clVCHA_CDE_REFERENCIA
   CMD("@VCHA_MON_MONEDA_ID") = clVCHA_MON_MONEDA_ID
   CMD("@FLOA_DEV_TIPO_CAMBIO") = clFLOA_DEV_TIPO_CAMBIO
   CMD("@VCHA_SER_SERIE_ID") = clVCHA_SER_SERIE_ID
   CMD("@INTE_DEV_AÑO") = clINTE_DEV_AÑO
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

















