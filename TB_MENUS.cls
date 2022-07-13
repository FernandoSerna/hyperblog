VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_MENUS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim cmd As New Command

'ANADIR UN REGISTRO
Public Function Anadir(clVcha_men_menu_id As String, clVcha_men_descripcion As String) As Boolean
Dim rs As adodb.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL

Set cmd.ActiveConnection = cnn
cmd.CommandType = adCmdStoredProc
    
If var_modifica_registro_menu Then
    cmd.CommandText = "menus_M"
Else
    cmd.CommandText = "menus_I"
End If
    cmd("@Vcha_MEN_MENU_id") = clVcha_men_menu_id
    cmd("@Vcha_MEN_descripcion") = clVcha_men_descripcion
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
'ELIMINAR
Public Function Eliminar(clVcha_men_menu_id As String) As Boolean
Dim rs As adodb.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set cmd.ActiveConnection = cnn                      'Esta es la conexión activa
cmd.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
cmd.CommandText = "menus_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    cmd("@Vcha_men_menu_id") = clVcha_men_menu_id
cmd.execute                                         'Ejecutar el PA

SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function


