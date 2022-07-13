VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_EXISTENCIAS_INSERTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim cmd As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVcha_alm_almacen_id As String, clVcha_art_articulo_id As String, _
clFLOA_EXI_CANTIDAD As Variant, clFLOA_EXI_COSTO As Variant, clVCHA_EXI_UBICACION As Variant) As Boolean
Dim rs As adodb.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL

Set cmd.ActiveConnection = cnn
   cmd.CommandType = adCmdStoredProc
   cmd.CommandText = "EXISTENCIAS_I"
   cmd("@VCHA_ALM_ALMACEN_ID") = clVcha_alm_almacen_id
   cmd("@VCHA_ART_ARTICULO_ID") = clVcha_art_articulo_id
   cmd("@FLOA_EXI_CANTIDAD") = clFLOA_EXI_CANTIDAD
   cmd("@FLOA_EXI_COSTO") = clFLOA_EXI_COSTO
   cmd("@VCHA_EXI_UBICACION") = clVCHA_EXI_UBICACION
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















