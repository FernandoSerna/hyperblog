VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_DETALLE_CAJAS_I"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim CMD As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clinte_ors_orden_surtido As Double, clinte_paq_caja As Integer, clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clvcha_alm_almacen_id As String, clVCHA_ART_ARTICULO_ID As String, clFLOA_PAQ_CANTIDAD As Double, clCHAR_PAQ_ESTATUS As String, clVCHA_PAQ_MOVIMIENTO_DESTINO As String, clINTE_PAQ_NUMERO_DESTINO As Integer, clFLOA_PAQ_COSTO As Double, clFLOA_PAQ_PRECIO As Double, clCHAR_PED_TIPO As String, clINTE_EMB_EMBARQUE As Double) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL

Set CMD.ActiveConnection = cnn
   CMD.CommandType = adCmdStoredProc
   CMD.CommandText = "DETALLE_CAJAS_I"
   CMD("@INTE_ORS_ORDEN_SURTIDO") = clinte_ors_orden_surtido
   CMD("@INTE_PAQ_CAJA") = clinte_paq_caja
   CMD("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
   CMD("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
   CMD("@VCHA_ALM_ALMACEN_ID") = clvcha_alm_almacen_id
   CMD("@VCHA_ART_ARTICULO_ID") = clVCHA_ART_ARTICULO_ID
   CMD("@FLOA_PAQ_CANTIDAD") = clFLOA_PAQ_CANTIDAD
   CMD("@CHAR_PAQ_ESTATUS") = clCHAR_PAQ_ESTATUS
   CMD("@VCHA_PAQ_MOVIMIENTO_DESTINO") = clVCHA_PAQ_MOVIMIENTO_DESTINO
   CMD("@INTE_PAQ_NUMERO_DESTINO") = clINTE_PAQ_NUMERO_DESTINO
   CMD("@FLOA_PAQ_COSTO") = clFLOA_PAQ_COSTO
   CMD("@FLOA_PAQ_PRECIO") = clFLOA_PAQ_PRECIO
   CMD("@CHAR_PED_TIPO") = clCHAR_PED_TIPO
   CMD("@INTE_EMB_EMBARQUE") = clINTE_EMB_EMBARQUE
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










