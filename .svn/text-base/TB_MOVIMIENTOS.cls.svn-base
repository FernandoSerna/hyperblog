VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_MOVIMIENTOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command


'ANADIR UN REGISTRO
Public Function Anadir(clvcha_mov_movimiento_id As String, clVcha_mov_nombre As String, clchar_mov_afectacion As String, clVcha_mov_referencia As String, clvcha_mov_titulo_origen As String, clinte_mov_folio As Integer, clInte_mov_factura As Integer, _
ClCHAR_MOV_DOCUMENTO As String, clINTE_MOV_CAUSA_DEVOLUCION As Integer, clVCHA_MOV_MOVIMIENTO_DEPENDENCIA As String, clVCHA_MOV_CLASE As String, clINTE_MOV_INTERCOMPAÑIA As Integer, clinte_mov_relectura As Integer, clINTE_MOV_ACEPTAR_MAS As Integer, clCHAR_MOV_TIPO_PROVEEDOR As String, clCHAR_MOV_TIPO_CLIENTE As String, _
ClINTE_MOV_PROMEDIA_COSTO As Integer, clINTE_MOV_DEVOLUCION_FACTURA As Integer, clVCHA_MOV_REPORTE_IMPRIMIR As String, clINTE_MOV_AJUSTE As Integer, clINTE_MOV_ULTIMO_COSTO As Integer, clINTE_MOV_REEMPAQUE As Integer, clINTE_MOV_AJUSTE_REEMPAQUE As Integer, clINTE_MOV_SOBRANTE As Integer, clVCHA_MOV_AGRUPADOR_CONCENTRADO As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL
Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
   
If var_modifica_registro_movimiento Then
    CMD.CommandText = "movimientos_M"
Else
    CMD.CommandText = "movimientos_I"
End If
    CMD("@VCHA_MOV_MOVIMIENTO_ID") = clvcha_mov_movimiento_id
    CMD("@VCHA_MOV_NOMBRE") = clVcha_mov_nombre
    CMD("@CHAR_MOV_AFECTACION") = clchar_mov_afectacion
    CMD("@INTE_MOV_REFEREANCIA") = clVcha_mov_referencia
    CMD("@VCHA_MOV_TITULO_ORIGEN") = clvcha_mov_titulo_origen
    CMD("@INTE_MOV_FACTURA") = clInte_mov_factura
    CMD("@INTE_MOV_FOLIO") = clinte_mov_folio
    
    CMD("@CHAR_MOV_DOCUMENTO") = ClCHAR_MOV_DOCUMENTO
    CMD("@INTE_MOV_CAUSA_DEVOLUCION") = clINTE_MOV_CAUSA_DEVOLUCION
    CMD("@VCHA_MOV_MOVIMIENTO_DEPENDENCIA") = clVCHA_MOV_MOVIMIENTO_DEPENDENCIA
    CMD("@VCHA_MOV_CLASE") = clVCHA_MOV_CLASE
    CMD("@INTE_MOV_INTERCOMPAÑIA") = clINTE_MOV_INTERCOMPAÑIA
    CMD("@INTE_MOV_RELECTURA") = clinte_mov_relectura
    CMD("@INTE_MOV_ACEPTAR_MAS") = clINTE_MOV_ACEPTAR_MAS
    CMD("@CHAR_MOV_TIPO_PROVEEDOR") = clCHAR_MOV_TIPO_PROVEEDOR
    CMD("@CHAR_MOV_TIPO_CLIENTE") = clCHAR_MOV_TIPO_CLIENTE
    CMD("@INTE_MOV_PROMEDIA_COSTO") = ClINTE_MOV_PROMEDIA_COSTO
    CMD("@INTE_MOV_DEVOLUCION_FACTURA") = clINTE_MOV_DEVOLUCION_FACTURA
    CMD("@VCHA_MOV_REPORTE_IMPRIMIR") = clVCHA_MOV_REPORTE_IMPRIMIR
    CMD("@INTE_MOV_AJUSTE") = clINTE_MOV_AJUSTE
    CMD("@INTE_MOV_ULTIMO_COSTO") = clINTE_MOV_ULTIMO_COSTO
    CMD("@INTE_MOV_REEMPAQUE") = clINTE_MOV_REEMPAQUE
    CMD("@INTE_MOV_AJUSTE_REEMPAQUE") = clINTE_MOV_AJUSTE_REEMPAQUE
    CMD("@INTE_MOV_SOBRANTE") = clINTE_MOV_SOBRANTE
    CMD("@VCHA_MOV_AGRUPADOR_CONCENTRADO") = clVCHA_MOV_AGRUPADOR_CONCENTRADO
    
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
Public Function Eliminar(clvcha_mov_movimiento_id As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "movimientos_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@VCHA_MOV_MOVIMIENTO_ID") = clvcha_mov_movimiento_id
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
