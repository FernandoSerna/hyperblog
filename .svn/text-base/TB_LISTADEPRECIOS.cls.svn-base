VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_LISTADEPRECIOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command


'ANADIR UN REGISTRO
Public Function Anadir(clVcha_lis_lista_id As String, clVcha_lis_nombre As String, clVcha_lis_fecha_inicio As Date, clVcha_lis_fecha_fin As Date, clVcha_mon_moneda As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL

Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    
If var_modifica_registro_lista_precios Then
    CMD.CommandText = "LISTADEPRECIOS_M"
Else
    CMD.CommandText = "LISTADEPRECIOS_I"
End If
    CMD("@Vcha_lis_lista_id") = clVcha_lis_lista_id
    CMD("@Vcha_lis_nombre") = clVcha_lis_nombre
    CMD("@DTIM_lis_fecha_inicio") = clVcha_lis_fecha_inicio
    CMD("@dtim_lis_fecha_fin") = clVcha_lis_fecha_fin
    CMD("@Vcha_mon_moneda") = clVcha_mon_moneda
     
    
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
Public Function Eliminar(clVcha_lis_lista_id As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "LISTADEPRECIOS_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@Vcha_lis_lista_id") = clVcha_lis_lista_id
CMD.execute                                         'Ejecutar el PA

SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function

