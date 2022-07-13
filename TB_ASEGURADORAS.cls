VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_ASEGURADORAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim cmd As New Command


'ANADIR UN REGISTRO
Public Function Anadir(clVcha_ase_aseguradora_id As String, clVcha_ase_nombre As String, clVcha_ase_direcion As String, clVcha_ase_telefono As String, clVcha_ase_responsable As String) As Boolean
Dim rs As adodb.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL
Set cmd.ActiveConnection = cnn
cmd.CommandType = adCmdStoredProc
    
If var_modifica_registro Then
    cmd.CommandText = "ASEGURADORAS_M"
Else
    cmd.CommandText = "ASEGURADORAS_I"
End If
    cmd("@Vcha_ase_aseguradora_id") = clVcha_ase_aseguradora_id
    cmd("@Vcha_ase_nombre") = clVcha_ase_nombre
    cmd("@Vcha_ase_direcion") = clVcha_ase_direcion
    cmd("@Vcha_ase_telefono") = clVcha_ase_telefono
    cmd("@Vcha_ase_responsable") = clVcha_ase_responsable
   
    
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
Public Function Eliminar(clVcha_ase_aseguradora_id As String) As Boolean
Dim rs As adodb.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set cmd.ActiveConnection = cnn                      'Esta es la conexión activa
cmd.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
cmd.CommandText = "ASEGURADORAS_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    cmd("@Vcha_ase_aseguradora_id") = clVcha_ase_aseguradora_id
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
