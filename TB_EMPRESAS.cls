VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_EMPRESAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command


'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVcha_emp_nombre As String, clVcha_emp_direccion As String, clVcha_emp_rfc As String, clVcha_emp_giro As String, clVcha_emp_gerente As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    
    
If var_modifica_registro_empresa Then
    CMD.CommandText = "EMPRESAS_M"
Else
    CMD.CommandText = "EMPRESAS_I"
End If
    CMD("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
    CMD("@VCHA_EMP_NOMBRE") = clVcha_emp_nombre
    CMD("@VCHA_EMP_DIRECCION") = clVcha_emp_direccion
    CMD("@VCHA_EMP_RFC") = clVcha_emp_rfc
    CMD("@VCHA_EMP_GIRO") = clVcha_emp_giro
    CMD("@VCHA_EMP_GERENTE") = clVcha_emp_gerente
    
    
CMD.execute
Set CMD = Nothing

SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Anadir = False
    GoTo SIGUE
End Function
'ELIMINAR
Public Function Eliminar(clVCHA_EMP_EMPRESA_ID As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
SIGUE:
On Error GoTo 0
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "EMPRESAS_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@Vcha_emp_empresa_id") = clVCHA_EMP_EMPRESA_ID
CMD.execute                                         'Ejecutar el PA

Set CMD = Nothing                                   'Liberar Memoria

Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function

