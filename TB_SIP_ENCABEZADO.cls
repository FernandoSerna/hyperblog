VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_SIP_ENCABEZADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String

Dim cmd As New Command


'ANADIR UN REGISTRO
Public Function Anadir(clBint_tra_transacciones_id As String, clvcha_mov_movimiento_id As String _
, clVcha_tra_movimiento1 As String, clVcha_tra_movimiento2 As String, clVcha_tra_alamacen As String _
, clVcha_tra_movimiento3 As String, clVcha_tra_status As String, clDtim_aud_fecha As String _
, clVcha_aud_usuario As String, clVcha_aud_maquina As String, clBint_pla_planta_id As String _
, clBint_tra_reimpresion As String) As Boolean

Anadir = True
On Error GoTo HELL


Set cmd.ActiveConnection = cnn                      'Esta es la conexión activa
cmd.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
If var_modifica_registro Then
    cmd.CommandText = "TRANSACCIONES_M"                         'Abrir Procedimiento Almacenado para Actualizar Cambios
Else
    cmd.CommandText = "TRANSACCIONES_I"                         'Abrir Procedimiento Almacenado y Agregar Banco
End If

    cmd("@Bint_tra_transacciones_id") = TB_TFOLIOS 'numero
    cmd("@Vcha_mov_movimiento_id") = clvcha_mov_movimiento_id 'VMPSIP
    cmd("@Vcha_tra_movimiento1") = "SALVEDO"
    cmd("@Vcha_tra_movimiento2") = "SALVEDO"
    cmd("@Vcha_tra_alamacen") = IIf(clVcha_tra_alamacen = "", 0, clVcha_tra_alamacen) 'NUMERO DE PLANTAS QUE SE HACE LA VENTA
    cmd("@Vcha_tra_status") = "A"
    cmd("@Vcha_tra_movimiento3") = "SALVEDO"
    cmd("@Dtim_aud_fecha") = clDtim_aud_fecha 'FECHA
    cmd("@Vcha_aud_usuario") = clVcha_aud_usuario
    cmd("@Vcha_aud_maquina") = clVcha_aud_maquina
    cmd("@Bint_pla_planta_id") = clBint_pla_planta_id 'NUMERO DE PLANTAS QUE SE HACE LA VENTA
    cmd("@Bint_tra_reimpresion") = 1
    cmd.execute                                         'Ejecutar el PA
Set cmd = Nothing                                   'Liberar Memoria

Exit Function
HELL:
    MsgBox Err.Number & " " & Err.Description & " " & Err.Source
    Anadir = False

End Function

'ELIMINAR
Public Function Eliminar(clBint_tra_transacciones_id As String) As Boolean

Eliminar = True
On Error GoTo HELL

Set cmd.ActiveConnection = cnn                      'Esta es la conexión activa
cmd.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
cmd.CommandText = "TRANSACCIONES_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    cmd("@Bint_tra_transacciones_id") = clBint_tra_transacciones_id
cmd.execute                                         'Ejecutar el PA

Set cmd = Nothing                                   'Liberar Memoria


SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function



