VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_CATALOGOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command


'ANADIR UN REGISTRO
Public Function Anadir(clVcha_cat_catalogo_id As String, clVcha_cat_nombre As String, clVcha_agr_agrupador_catalogo_id As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL

Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    
If var_modifica_registro_catalogo Then
    CMD.CommandText = "CATALOGOS_M"
Else
    CMD.CommandText = "CATALOGOS_I"
End If
    CMD("@Vcha_cat_catalogo_id") = clVcha_cat_catalogo_id
    CMD("@Vcha_cat_nombre") = clVcha_cat_nombre
    CMD("@vcha_agr_agrupador_catalogo_id") = clVcha_agr_agrupador_catalogo_id
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
Public Function Eliminar(clVcha_cat_catalogo_id As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexi�n activa
CMD.CommandType = adCmdStoredProc                   'Aqu� le indico a ADO que se trata de un PA
    
CMD.CommandText = "CATALOGOS_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@Vcha_cat_catalogo_id") = clVcha_cat_catalogo_id
CMD.execute                                         'Ejecutar el PA

SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function
