VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_TIPOPEDIDOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim CMD As New Command


'ANADIR UN REGISTRO
Public Function Anadir(clCHAR_TPE_TIPO_PEDIDO_ID As String, clVCHA_TPE_NOMBRE As String, clinte_tpe_resurtible As Integer, clinte_tpe_autorizacion As Integer, clVCHA_TCL_TIPO_CLIENTE_ID As String, clINTE_TPE_CARGA_ARCHIVO As Integer, clINTE_TPE_DIAS_CADUCIDAD As Integer, clVCHA_MOV_MOVIMIENTO_ID, clFLOA_TPE_IVA As Double) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL

Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    
    
    
If var_modifica_registro_tipopedido Then
    CMD.CommandText = "tipopedidos_M"
Else
    CMD.CommandText = "tipopedidos_I"
End If
    CMD("@CHAR_TPE_TIPO_PEDIDO_ID") = clCHAR_TPE_TIPO_PEDIDO_ID
    CMD("@VCHA_TPE_NOMBRE") = clVCHA_TPE_NOMBRE
    CMD("@INTE_TPE_RESURTIBLE") = clinte_tpe_resurtible
    CMD("@INTE_TPE_AUTORIZACION") = clinte_tpe_autorizacion
    CMD("@VCHA_TCL_TIPO_CLIENTE_ID") = clVCHA_TCL_TIPO_CLIENTE_ID
    CMD("@INTE_TPE_CARGA_ARCHIVO") = clINTE_TPE_CARGA_ARCHIVO
    CMD("@INTE_TPE_DIAS_CADUCIDAD") = clINTE_TPE_DIAS_CADUCIDAD
    CMD("@VCHA_MOV_MOVIMIENTO_ID") = clVCHA_MOV_MOVIMIENTO_ID
    CMD("@FLOA_TPE_IVA") = clFLOA_TPE_IVA

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
Public Function Eliminar(clCHAR_TPE_TIPO_PEDIDO_ID As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "tipopedidos_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@CHAR_TPE_TIPO_PEDIDO_ID") = clCHAR_TPE_TIPO_PEDIDO_ID
CMD.execute                                         'Ejecutar el PA

SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function


