VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_PROVEEDOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command



'ANADIR UN REGISTRO
Public Function Anadir(clvcha_pro_proveedor_id As String, clVcha_pro_nombre As String _
, clVcha_pro_domicilio As String, clvcha_pro_colonia As String, clBint_pro_cp As String, clVCHA_PRO_RFC As String, clvcha_pro_telefono As String _
, clVcha_pro_estado As String, clVcha_pro_status As String, clDtim_aud_fecha As String _
, clVcha_aud_usuario As String, clVcha_aud_maquina As String _
, clBint_pla_planta_id As String) As Boolean

Anadir = True
On Error GoTo HELL

Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
If var_modifica_registro_proveedor Then
    CMD.CommandText = "PROVEEDOR_M"                         'Abrir Procedimiento Almacenado para Actualizar Cambios
Else
    CMD.CommandText = "PROVEEDOR_I"                         'Abrir Procedimiento Almacenado y Agregar Banco
End If

    CMD("@Vcha_pro_proveedor_id") = clvcha_pro_proveedor_id
    CMD("@Vcha_pro_nombre") = clVcha_pro_nombre
    CMD("@Vcha_pro_domicilio") = clVcha_pro_domicilio
    CMD("@Vcha_pro_colonia") = clvcha_pro_colonia
    CMD("@Bint_pro_cp") = clBint_pro_cp
    CMD("@Vcha_pro_rfc") = clVCHA_PRO_RFC
    CMD("@Vcha_pro_telefono") = clvcha_pro_telefono
    CMD("@Vcha_pro_estado") = clVcha_pro_estado
    CMD("@Vcha_pro_status") = clVcha_pro_status
    
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
Public Function Eliminar(clvcha_pro_proveedor_id As String) As Boolean

Eliminar = True
On Error GoTo HELL

Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    

CMD.CommandText = "PROVEEDOR_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@Vcha_pro_proveedor_id") = clvcha_pro_proveedor_id  'Ejecutar el PA
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

rs.Open "select * from TB_PROVEEDOR_VIEW where " & var_campo & " =  '" + Trim(var_comparar) + "'", cn, adOpenDynamic, adLockOptimistic
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

Siguiente = ""
On Error GoTo HELL

rs.Open "select * from TB_PROVEEDOR_VIEW order by Vcha_pro_proveedor_id", cn, adOpenDynamic, adLockOptimistic
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


