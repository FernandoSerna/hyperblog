VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_FAMILIASCLIENTES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim cmd As New Command

'ANADIR UN REGISTRO
Public Function Anadir(clVcha_fcl_familia_id As String, clVcha_fcl_nombre As String, clVcha_emp_empresa_id As String) As Boolean
Dim rs As adodb.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL
SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Anadir = False
    GoTo SIGUE
End Function
'ELIMINAR
Public Function Eliminar(clVcha_fcl_familia_id As String) As Boolean
Dim rs As adodb.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function
