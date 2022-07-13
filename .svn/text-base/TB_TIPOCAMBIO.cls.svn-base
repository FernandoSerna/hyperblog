VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_TIPOCAMBIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command


'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_MON_MONEDA_ID As String, clDtim_tca_fecha As String, clMone_tca_importe As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL
Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    
If var_modifica_registro_tipocambio Then
    CMD.CommandText = "TIPOCAMBIO_M"
Else
    CMD.CommandText = "TIPOCAMBIO_I"
End If
    CMD("@Vcha_mon_moneda_id") = clVCHA_MON_MONEDA_ID
    CMD("@Dtim_tca_fecha") = clDtim_tca_fecha
    CMD("@Mone_tca_importe") = clMone_tca_importe
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
Public Function Eliminar(clVCHA_MON_MONEDA_ID As String, clDtim_tca_fecha As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "TIPOCAMBIO_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@Vcha_mon_moneda_id") = clVCHA_MON_MONEDA_ID
    CMD("@Dtim_tca_fecha") = clDtim_tca_fecha
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
