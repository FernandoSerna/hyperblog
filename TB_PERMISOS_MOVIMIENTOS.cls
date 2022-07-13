VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_PERMISOS_MOVIMIENTOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command


'ANADIR UN REGISTRO
Public Function Anadir(clVcha_usu_usuario_id As String, clVcha_mov_movimiento_id As String, clVCHA_PER_ALMACEN_1 As String, clVCHA_PER_ALMACEN_2 As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL

Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    
If var_modifica_registro Then
    CMD.CommandText = "PERMISOS_MOVIMIENTOS_M"
Else
    CMD.CommandText = "PERMISOS_MOVIMIENTOS_I"
End If
    CMD("@Vcha_usu_usuario_id") = clVcha_usu_usuario_id
    CMD("@Vcha_mov_movimiento_id") = clVcha_mov_movimiento_id
    CMD("@VCHA_PER_ALMACEN_1") = clVCHA_PER_ALMACEN_1
    CMD("@VCHA_PER_ALMACEN_2") = clVCHA_PER_ALMACEN_2
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
'ELIMINAR
Public Function Eliminar(clVcha_usu_usuario_id As String, clVcha_mov_movimiento_id As String, clVCHA_PER_ALMACEN_1 As String, clVCHA_PER_ALMACEN_2 As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "PERMISOS_MOVIMIENTOS_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@Vcha_usu_usuario_id") = clVcha_usu_usuario_id
    CMD("@Vcha_mov_movimiento_id") = clVcha_mov_movimiento_id
    CMD("@VCHA_PER_ALMACEN_1") = clVCHA_PER_ALMACEN_1
    CMD("@VCHA_PER_ALMACEN_2") = clVCHA_PER_ALMACEN_2
CMD.execute                                         'Ejecutar el PA

SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function


