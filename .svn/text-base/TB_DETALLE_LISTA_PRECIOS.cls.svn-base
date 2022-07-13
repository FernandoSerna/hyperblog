VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_DETALLE_LISTA_PRECIOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command

'ANADIR UN REGISTRO
Public Function Anadir(clVcha_lis_lista_precios_id As String, clVcha_art_articulo_id As String, clFLOA_DLI_PRECIO As Double) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
    CMD.CommandText = "DETALLE_LISTA_PRECIOS_I"
    CMD("@VCHA_LIS_LISTA_PRECIOS_ID") = clVcha_lis_lista_precios_id
    CMD("@VCHA_ART_ARTICULO_ID") = clVcha_art_articulo_id
    CMD("@FLOA_DLI_PRECIO") = clFLOA_DLI_PRECIO
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
Public Function Modificar(clVcha_lis_lista_precios_id As String, clVcha_art_articulo_id As String, clFLOA_DLI_PRECIO As Double) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Modificar = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
    CMD.CommandText = "DETALLE_LISTA_PRECIOS_M"
    CMD("@VCHA_LIS_LISTA_PRECIOS_ID") = clVcha_lis_lista_precios_id
    CMD("@VCHA_ART_ARTICULO_ID") = clVcha_art_articulo_id
    CMD("@FLOA_DLI_PRECIO") = clFLOA_DLI_PRECIO
CMD.execute
Set CMD = Nothing


SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Modificar = False
    GoTo SIGUE
End Function
'ELIMINAR
Public Function Eliminar(clVcha_lis_lista_precios_id As String, clVcha_art_articulo_id As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    
CMD.CommandText = "DETALLE_LISTA_PRECIOS_E"
    CMD("@VCHA_LIS_LISTA_PRECIOS_ID") = clVcha_lis_lista_precios_id
    CMD("@VCHA_ART_ARTICULO_ID") = clVcha_art_articulo_id
CMD.execute
SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function
