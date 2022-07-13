VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_VENDEDORES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command



'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_VEN_VENDEDOR_ID As String, clVcha_ven_nombre As String, _
clVCHA_CAN_CANAL_VENTA_ID As String, clVcha_ven_telefono As String, clinte_ven_comision As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    
If var_modifica_registro_vendedor Then
    CMD.CommandText = "VENDEDORES_M"
Else
    CMD.CommandText = "VENDEDORES_I"
End If
    CMD("@Vcha_ven_vendedor_id") = clVCHA_VEN_VENDEDOR_ID
    CMD("@Vcha_ven_nombre") = clVcha_ven_nombre
    CMD("@Vcha_can_canal_venta_id") = clVCHA_CAN_CANAL_VENTA_ID
    CMD("@Vcha_ven_telefono") = clVcha_ven_telefono
    CMD("@inte_ven_comision") = clinte_ven_comision
   
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
Public Function Eliminar(clVCHA_VEN_VENDEDOR_ID As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "VENDEDORES_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@Vcha_ven_vendedor_id") = clVCHA_VEN_VENDEDOR_ID
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


