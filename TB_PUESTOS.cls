VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_PUESTOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String

Dim CMD As New Command



'ANADIR UN REGISTRO
Public Function Anadir(clVcha_pue_puesto_id As String, clVcha_pue_descripcion As String, clVcha_men_menu_id As String _
, clINTE_PUE_SUPERVISOR As String) As Boolean

Anadir = True
On Error GoTo HELL

Set CMD.ActiveConnection = cnn                      'Esta es la conexi�n activa
CMD.CommandType = adCmdStoredProc                   'Aqu� le indico a ADO que se trata de un PA
    
If var_modifica_registro_puesto Then
    CMD.CommandText = "PUESTOS_M"                         'Abrir Procedimiento Almacenado para Actualizar Cambios
Else
    CMD.CommandText = "PUESTOS_I"                         'Abrir Procedimiento Almacenado y Agregar Banco
End If

    CMD("@VCHA_PUE_PUESTO_ID") = clVcha_pue_puesto_id
    CMD("@VCHA_PUE_DESCRIPCION") = clVcha_pue_descripcion
    CMD("@VCHA_MEN_MENU_ID") = clVcha_men_menu_id
    CMD("@INTE_PUE_SUPERVISOR") = clINTE_PUE_SUPERVISOR
    
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
Public Function Eliminar(clVcha_pue_puesto_id As String) As Boolean

Eliminar = True
'On Error GoTo HELL

Set CMD.ActiveConnection = cnn                      'Esta es la conexi�n activa
CMD.CommandType = adCmdStoredProc                   'Aqu� le indico a ADO que se trata de un PA
    
CMD.CommandText = "PUESTOS_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@VCHA_PUE_PUESTO_ID") = clVcha_pue_puesto_id
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

'rs.Open "select * from TB_LINEAS_VIEW where " & var_campo & " =  '" + Trim(var_comparar) + "'", cn, adOpenDynamic, adLockOptimistic
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
'rs.Open "select * from TB_LINEAS_VIEW order by Bint_lin_lineas_id", cn, adOpenDynamic, adLockOptimistic
'If rs.RecordCount <> 0 Then
'  rs.MoveLast
'  Siguiente = rs(0) + 1
'Else
'  Siguiente = 1
'End If
'rs.Close
SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    GoTo SIGUE
End Function

