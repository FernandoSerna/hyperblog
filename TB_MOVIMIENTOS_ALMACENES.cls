VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_MOVIMIENTOS_ALMACENES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim cmd As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVcha_mov_movimiento_id As String, clVCHA_ALM_ALMACEN_ID As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL

Set cmd.ActiveConnection = cnn
   cmd.CommandType = adCmdStoredProc
   cmd.CommandText = "MOVIMIENTOS_ALMACENES_I"
   cmd("@VCHA_MOV_MOVIMIENTO_ID") = clVcha_mov_movimiento_id
   cmd("@VCHA_ALM_ALMACEN_ID") = clVCHA_ALM_ALMACEN_ID
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




Public Function Eliminar(clVcha_mov_movimiento_id As String, clVCHA_ALM_ALMACEN_ID As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
'On Error GoTo HELL

Set cmd.ActiveConnection = cnn
   cmd.CommandType = adCmdStoredProc
   cmd.CommandText = "MOVIMIENTOS_ALMACENES_E"
   cmd("@VCHA_MOV_MOVIMIENTO_ID") = clVcha_mov_movimiento_id
   cmd("@VCHA_ALM_ALMACEN_ID") = clVCHA_ALM_ALMACEN_ID
   cmd.execute
Set cmd = Nothing
SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function





