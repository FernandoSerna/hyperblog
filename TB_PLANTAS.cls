VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_PLANTAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String

Dim cmd As New Command

'ANADIR UN REGISTRO
Public Function Anadir(clVcha_pla_planta_id As String, clVcha_pla_descripc As String, clVcha_pla_uniorgaz As String, clVcha_pla_responsa As String _
, clVcha_pla_domicilio As String, clVcha_pla_cp As String, clVcha_pla_telefono As String _
, clVcha_pla_rfc As String, clVcha_aud_fechaora As String, clVcha_aud_usuario As String, clVcha_aud_maquina As String, clBint_pla_planta_id As String) As Boolean

Anadir = True
On Error GoTo HELL

Set cmd.ActiveConnection = cnn                      'Esta es la conexión activa
cmd.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
If var_modifica_registro Then
    cmd.CommandText = "PLANTAS_M"                         'Abrir Procedimiento Almacenado para Actualizar Cambios
Else
    cmd.CommandText = "PLANTAS_I"                         'Abrir Procedimiento Almacenado y Agregar Banco
End If
    cmd("@VCHA_pla_planta_id") = clVcha_pla_planta_id
    cmd("@VCHA_pla_descripc") = clVcha_pla_descripc
    cmd("@VCHA_pla_uniorgaz") = clVcha_pla_uniorgaz
    cmd("@VCHA_pla_responsa") = clVcha_pla_responsa
    
    cmd("@VCHA_pla_domicilio") = clVcha_pla_domicilio
    cmd("@VCHA_pla_cp") = clVcha_pla_cp
    cmd("@VCHA_pla_telefono") = clVcha_pla_telefono
    cmd("@VCHA_pla_rfc") = clVcha_pla_rfc
    
    
    cmd("@VCHA_aud_fechaora") = clVcha_aud_fechaora
    cmd("@VCHA_aud_usuario") = clVcha_aud_usuario
    cmd("@VCHA_aud_maquina") = clVcha_aud_maquina
    cmd("@Bint_pla_planta_id") = clBint_pla_planta_id
    
cmd.execute                                         'Ejecutar el PA
Set cmd = Nothing                                   'Liberar Memoria

SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Anadir = False
    GoTo SIGUE
End Function



'ELIMINAR
Public Function Eliminar(clVcha_pla_planta_id As String) As Boolean

Eliminar = True
On Error GoTo HELL

Set cmd.ActiveConnection = cnn                      'Esta es la conexión activa
cmd.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
cmd.CommandText = "PLANTAS_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    cmd("@VCHA_pla_planta_id") = clVcha_pla_planta_id
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



'UBICAR UN REGISTRO
Public Function Obtener_llave(cn As ADODB.Connection, ByVal rs As Recordset, ByVal var_campo As String, ByVal var_comparar As String) As String

Ubicar = True
On Error GoTo HELL

rs.Open "select * from TB_PLANTAS where " & var_campo & " =  '" + Trim(var_comparar) + "'", cn, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
  Obtener_llave = rs(0).Value
End If

rs.Close


SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Ubicar = False
    GoTo SIGUE
End Function



'++++++++++++++++++++++ SIGUIENTE REGISTRO LLAVE

Public Function Siguiente(ByVal cn As ADODB.Connection, ByVal rs As Recordset) As String
Dim Cadena As String
Siguiente = ""
On Error GoTo HELL
rs.Open "select * from TB_plantas order by Vcha_pla_planta_id", cn, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
  rs.MoveLast
  Siguiente = rs(0) + 1
End If
rs.Close
SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    GoTo SIGUE
End Function

