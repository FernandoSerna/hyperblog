VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_TRANSACCIONES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String

Dim CMD As New Command


'ANADIR UN REGISTRO
Public Function Anadir(clBint_tra_transacciones_id As String, clvcha_mov_movimiento_id As String _
, clVcha_tra_movimiento1 As String, clVcha_tra_movimiento2 As String, clVcha_tra_alamacen As String _
, clVcha_tra_movimiento3 As String, clVcha_tra_status As String, clDtim_aud_fecha As String _
, clVcha_aud_usuario As String, clVcha_aud_maquina As String, clBint_pla_planta_id As String) As Boolean

Anadir = True
'On Error GoTo HELL


Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
If var_modifica_registro_transaccion Then
    CMD.CommandText = "TRANSACCIONES_M"                         'Abrir Procedimiento Almacenado para Actualizar Cambios
Else
    CMD.CommandText = "TRANSACCIONES_I"                         'Abrir Procedimiento Almacenado y Agregar Banco
End If

    CMD("@Bint_tra_transacciones_id") = clBint_tra_transacciones_id
    CMD("@Vcha_mov_movimiento_id") = clvcha_mov_movimiento_id
    CMD("@Vcha_tra_movimiento1") = IIf(clVcha_tra_movimiento1 = "", 0, clVcha_tra_movimiento1)
    CMD("@Vcha_tra_movimiento2") = IIf(clVcha_tra_movimiento2 = "", 0, clVcha_tra_movimiento2)
    CMD("@Vcha_tra_alamacen") = IIf(clVcha_tra_alamacen = "", 0, clVcha_tra_alamacen)
    CMD("@Vcha_tra_status") = clVcha_tra_status
    CMD("@Vcha_tra_movimiento3") = clVcha_tra_movimiento3
    
    CMD("@Dtim_aud_fecha") = clDtim_aud_fecha
    CMD("@Vcha_aud_usuario") = clVcha_aud_usuario
    CMD("@Vcha_aud_maquina") = clVcha_aud_maquina
    CMD("@Bint_pla_planta_id") = clBint_pla_planta_id
    
    
CMD.execute                                         'Ejecutar el PA
Set CMD = Nothing                                   'Liberar Memoria

SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Anadir = False
    GoTo SIGUE
End Function

'ELIMINAR
Public Function Eliminar(clBint_tra_transacciones_id As String) As Boolean

Eliminar = True
On Error GoTo HELL

Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "TRANSACCIONES_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@Bint_tra_transacciones_id") = clBint_tra_transacciones_id
CMD.execute                                         'Ejecutar el PA

Set CMD = Nothing                                   'Liberar Memoria


SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function

