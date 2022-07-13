VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_REFERENCIAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command


'ANADIR UN REGISTRO
Public Function Anadir(clVcha_ref_referencia_id As String, clVCHA_REF_NOMBRE As String, clVCHA_REF_TITULO As String, _
clVCHA_REF_TABLA_NOMBRE_ORIGEN As String, clVCHA_REF_RUTA_TABLA_ORIGEN As String, clVCHA_REF_CAMPO_CODIGO_ORIGEN As String, _
clVCHA_REF_CAMPO_DESCRIPCION_ORIGEN As String, clVCHA_REF_CAMPO_COSTO_ORIGEN As String, clVCHA_REF_CAMPO_CANTIDAD_ORIGEN As String, clVCHA_REF_CAMPO_CANTIDAD_ENTRADA As String, _
clVCHA_REF_TABLA_DESTINO As String, clVCHA_REF_CAMPO_CODIGO_DESTINO As String, clVCHA_REF_CAMPO_DESCRIPCION_DESTINO As String, clVCHA_REF_CAMPO_COSTO_DESTINO As String, clVCHA_REF_CAMPO_CANTIDAD_DESTINO As String, clvcha_ref_campo_numero) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
If var_modifica_registro_referencia Then
    CMD.CommandText = "referencias_M"
Else
    CMD.CommandText = "referencias_I"
End If
    CMD("@VCHA_REF_REFERENCIA_ID") = clVcha_ref_referencia_id
    CMD("@VCHA_REF_NOMBRE") = clVCHA_REF_NOMBRE
    CMD("@VCHA_REF_TITULO") = clVCHA_REF_TITULO
    CMD("@VCHA_REF_TABLA_NOMBRE_ORIGEN") = clVCHA_REF_TABLA_NOMBRE_ORIGEN
    CMD("@VCHA_REF_RUTA_TABLA_ORIGEN") = clVCHA_REF_RUTA_TABLA_ORIGEN
    CMD("@VCHA_REF_CAMPO_CODIGO_ORIGEN") = clVCHA_REF_CAMPO_CODIGO_ORIGEN
    CMD("@VCHA_REF_CAMPO_DESCRIPCION_ORIGEN") = clVCHA_REF_CAMPO_DESCRIPCION_ORIGEN
    CMD("@VCHA_REF_CAMPO_COSTO_ORIGEN") = clVCHA_REF_CAMPO_COSTO_ORIGEN
    CMD("@VCHA_REF_CAMPO_CANTIDAD_ORIGEN") = clVCHA_REF_CAMPO_CANTIDAD_ORIGEN
    CMD("@VCHA_REF_CAMPO_CANTIDAD_ENTRADA") = clVCHA_REF_CAMPO_CANTIDAD_ENTRADA
    CMD("@VCHA_REF_TABLA_DESTINO") = clVCHA_REF_TABLA_DESTINO
    CMD("@VCHA_REF_CAMPO_CODIGO_DESTINO") = clVCHA_REF_CAMPO_CODIGO_DESTINO
    CMD("@VCHA_REF_CAMPO_DESCRIPCION_DESTINO") = clVCHA_REF_CAMPO_DESCRIPCION_DESTINO
    CMD("@VCHA_REF_CAMPO_COSTO_DESTINO") = clVCHA_REF_CAMPO_COSTO_DESTINO
    CMD("@VCHA_REF_CAMPO_CANTIDAD_DESTINO") = clVCHA_REF_CAMPO_CANTIDAD_DESTINO
    CMD("@VCHA_REF_CAMPO_NUMERO") = clvcha_ref_campo_numero
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
Public Function Eliminar(clVcha_ref_referencia_id As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "referencias_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@VCHA_REF_REFERENCIA_ID") = clVcha_ref_referencia_id
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

