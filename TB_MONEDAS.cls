VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_MONEDAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command

'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_MON_MONEDA_ID As String, clVcha_mon_descripcion As String, clInte_mon_moneda_local As Integer, clVCHA_MON_NOMBRE_PLURAL As String, clVCHA_MON_NOMBRE_REGIONAL As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL

Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    
If var_modifica_registro_moneda Then
    CMD.CommandText = "MONEDAS_M"
Else
    CMD.CommandText = "MONEDAS_I"
End If
    CMD("@Vcha_mon_moneda_id") = clVCHA_MON_MONEDA_ID
    CMD("@Vcha_mon_NOMBRE") = clVcha_mon_descripcion
    CMD("@inte_mon_moneda_local") = clInte_mon_moneda_local
    CMD("@VCHA_MON_NOMBRE_PLURAL") = clVCHA_MON_NOMBRE_PLURAL
    CMD("@VCHA_MON_NOMBRE_REGIONAL") = clVCHA_MON_NOMBRE_REGIONAL
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
Public Function Eliminar(clVCHA_MON_MONEDA_ID As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "MONEDAS_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@Vcha_mon_moneda_id") = clVCHA_MON_MONEDA_ID
CMD.execute                                         'Ejecutar el PA

SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function

