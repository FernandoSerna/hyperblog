VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_USUARIOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String

Dim CMD As New Command

'ANADIR UN REGISTRO
Public Function Anadir(clVcha_usu_usuario_id As String, clVcha_usu_nombre As String, clVcha_usu_apellidos As String, clVcha_usu_usuario As String, clVcha_usu_password As String, clVcha_usu_sistema As String, clInte_usu_permiso As Integer) As Boolean

                                                    'Este es el objeto Command que declaramos
Anadir = True
'On Error GoTo HELL

Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA

If var_modifica_registro_usuario Then
    CMD.CommandText = "USUARIOS_M"                         'Abrir Procedimiento Almacenado para Actualizar Cambios
Else
    CMD.CommandText = "USUARIOS_I"                         'Abrir Procedimiento Almacenado y Agregar Banco
End If
    CMD("@VCHA_USU_USUARIO_ID") = clVcha_usu_usuario_id
    CMD("@VCHA_USU_NOMBRE") = clVcha_usu_nombre
    CMD("@VCHA_USU_APELLIDOS") = clVcha_usu_apellidos
    CMD("@VCHA_USU_USUARIO") = clVcha_usu_usuario
    CMD("@VCHA_USU_PASSWORD") = clVcha_usu_password
    CMD("@VCHA_USU_SISTEMA") = clVcha_usu_sistema
    CMD("@INTE_USU_PERMISO") = clInte_usu_permiso
   If Not var_modifica_registro_usuario Then
      CMD("@CLAVE_STRING") = ""
      CMD.execute
      var_usuario_regreso = CMD("@clave_string")
   Else
      CMD.execute
   End If
    
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
Public Function Eliminar(clVcha_usu_usuario_id As String) As Boolean
Eliminar = True
On Error GoTo HELL

Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "USUARIOS_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@VCHA_USU_USUARIO_ID") = clVcha_usu_usuario_id
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
Public Function Obtener_llave(cn As adodb.Connection, ByVal rs As Recordset, ByVal var_campo As String, ByVal var_comparar As String) As String

Ubicar = True
On Error GoTo HELL

'rs.Open "select * from TB_USUARIOS where " & var_campo & " =  '" + Trim(var_comparar) + "'", cn, adOpenDynamic, adLockOptimistic
'If Not rs.EOF Then
'  Obtener_llave = rs(0).Value
'End If

'rs.Close


SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Ubicar = False
    GoTo SIGUE
End Function



'SIGUIENTE REGISTRO LLAVE
Public Function Siguiente(ByVal cn As adodb.Connection, ByVal rs As Recordset) As String
Dim Cadena As String
Siguiente = ""
On Error GoTo HELL
'rs.Open "select * from TB_USUARIOS order by Vcha_usu_usuario_id", cn, adOpenDynamic, adLockOptimistic
'If Not rs.EOF Then
'  rs.MoveLast
'  Siguiente = rs(0) + 1
'End If
'rs.Close
SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    GoTo SIGUE
End Function

