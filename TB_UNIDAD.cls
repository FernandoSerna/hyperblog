VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_UNIDAD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String

Dim CMD As New Command


'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_UNI_UNIDAD_ID As String, clVcha_uni_descripcion As String _
, clVcha_uni_status As String, clDtim_aud_fecha As String, clVcha_aud_usuario As String _
, clVcha_aud_maquina As String, clBint_pla_planta_id As String, clVcha_pla_descripc As String) As Boolean

Anadir = True
On Error GoTo HELL

Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
If var_modifica_registro_unidad Then
    CMD.CommandText = "UNIDAD_M"                         'Abrir Procedimiento Almacenado para Actualizar Cambios
Else
    CMD.CommandText = "UNIDAD_I"                         'Abrir Procedimiento Almacenado y Agregar Banco
End If

    CMD("@Vcha_uni_unidad_id") = clVCHA_UNI_UNIDAD_ID
    CMD("@Vcha_uni_descripcion") = clVcha_uni_descripcion
    CMD("@Vcha_uni_status") = clVcha_uni_status
    
    CMD("@Dtim_aud_fecha") = clDtim_aud_fecha
    CMD("@Vcha_aud_usuario") = clVcha_aud_usuario
    CMD("@Vcha_aud_maquina") = clVcha_aud_maquina
    CMD("@Bint_pla_planta_id") = clBint_pla_planta_id
    CMD("@Vcha_pla_descripc") = clVcha_pla_descripc
    
    
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
Public Function Eliminar(clBint_uni_unidad_id As String) As Boolean

Eliminar = True
On Error GoTo HELL

Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "UNIDAD_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@Bint_uni_unidad_id") = clBint_uni_unidad_id
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




'UBICAR UN REGISTRO
Public Function Obtener_llave(cn As ADODB.Connection, ByVal rs As Recordset, ByVal var_campo As String, ByVal var_comparar As String) As String

Ubicar = True
On Error GoTo HELL

rs.Open "select * from TB_LINEAS_VIEW where " & var_campo & " =  '" + Trim(var_comparar) + "'", cn, adOpenDynamic, adLockOptimistic
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



'SIGUIENTE REGISTRO LLAVE
Public Function Siguiente(ByVal cn As ADODB.Connection, ByVal rs As Recordset) As String
Dim Cadena As String
Siguiente = ""
On Error GoTo HELL
rs.Open "select * from TB_UNIDAD_VIEW order by vcha_uni_unidad_id", cn, adOpenDynamic, adLockOptimistic
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

