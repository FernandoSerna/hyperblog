VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_GRUPOSACTUALES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command



'ANADIR UN REGISTRO
Public Function Anadir(clVcha_gac_grupo_actual_id As String, clVCHA_GAC_NOMBRE As String, clFLOA_GAC_DESCUENTO1 As String, clFLOA_GAC_DESCUENTO2 As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL
Set CMD.ActiveConnection = cnn_importacion
CMD.CommandType = adCmdStoredProc
    
If var_modifica_registro_ga Then
    CMD.CommandText = "gruposactuales_M"
Else
    CMD.CommandText = "gruposactuales_I"
End If
    CMD("@VCHA_GAC_GRUPO_ACTUAL_ID") = clVcha_gac_grupo_actual_id
    CMD("@VCHA_GAC_NOMBRE") = clVCHA_GAC_NOMBRE
    CMD("@FLOA_GAC_DESCUENTO_1") = clFLOA_GAC_DESCUENTO1
    CMD("@FLOA_GAC_DESCUENTO_2") = clFLOA_GAC_DESCUENTO2
    If var_modifica_registro_ga Then
       CMD.execute
    Else
       CMD("@CLAVE_STRING") = ""
       CMD.execute
       var_grupo_actual_regreso = CMD("@CLAVE_STRING")
    End If
   
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
Public Function Eliminar(clVcha_gac_grupo_actual_id As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn_importacion                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "gruposactuales_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@VCHA_GAC_GRUPO_ACTUAL_ID") = clVcha_gac_grupo_actual_id
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


