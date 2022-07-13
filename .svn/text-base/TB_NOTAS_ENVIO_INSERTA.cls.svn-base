VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_NOTAS_ENVIO_INSERTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim cmd As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clVCHA_ALM_ALMACEN_ID As String, clVCHA_NEN_TIPO As String, clINTE_NEN_NUMERO As Variant, clINTE_NEN_LOTE As Integer, _
clDTIM_NEN_FECHA As String, clVCHA_ART_ARTICULO_ID As String, clFLOA_NEN_COSTO As Variant, clFLOA_NEN_CANTIDAD_ENVIADA As String, _
clFLOA_NEN_CANTIDAD_RECIBIDA As Variant, clVCHA_NEN_CHOFER As String, clCHAR_NEN_ESTATUS As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL
Set cmd.ActiveConnection = cnn
   cmd.CommandType = adCmdStoredProc
   cmd.CommandText = "NOTAS_ENVIO_I"
   cmd("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
   cmd("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
   cmd("@VCHA_ALM_ALMACEN_ID") = clVCHA_ALM_ALMACEN_ID
   cmd("@VCHA_NEN_TIPO") = clVCHA_NEN_TIPO
   cmd("@INTE_NEN_NUMERO") = clINTE_NEN_NUMERO
   cmd("@INTE_NEN_LOTE") = clINTE_NEN_LOTE
   cmd("@DTIM_NEN_FECHA") = clDTIM_NEN_FECHA
   cmd("@VCHA_ART_ARTICULO_ID") = clVCHA_ART_ARTICULO_ID
   cmd("@FLOA_NEN_COSTO") = clFLOA_NEN_COSTO
   cmd("@FLOA_NEN_CANTIDAD_ENVIADA") = clFLOA_NEN_CANTIDAD_ENVIADA
   cmd("@FLOA_NEN_CANTIDAD_RECIBIDA") = clFLOA_NEN_CANTIDAD_RECIBIDA
   cmd("@VCHA_NEN_CHOFER") = clVCHA_NEN_CHOFER
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









