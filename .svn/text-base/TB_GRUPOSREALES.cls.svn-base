VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_GRUPOSREALES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command

'ANADIR UN REGISTRO
Public Function Anadir(clVcha_gac_grupo_actual_id As String, clVcha_gre_grupo_real_id As String, clVCHA_GRE_NOMBRE As String, _
clfloa_gre_descuento_1 As Variant, clfloa_gre_descuento_2 As Variant, clFLOA_GRE_DESCUENTO_3 As Variant, clchar_pri_prioridad_id As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL
Set CMD.ActiveConnection = cnn_importacion                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
If var_modifica_registro_gr Then
    CMD.CommandText = "gruposreales_M"
Else
    CMD.CommandText = "gruposreales_I"
End If
    CMD("@VCHA_GAC_GRUPO_ACTUAL_ID") = clVcha_gac_grupo_actual_id
    CMD("@VCHA_GRE_GRUPO_REAL_ID") = clVcha_gre_grupo_real_id
    CMD("@VCHA_GRE_NOMBRE") = clVCHA_GRE_NOMBRE
    CMD("@FLOA_GRE_DESCUENTO_1") = clfloa_gre_descuento_1
    CMD("@FLOA_GRE_DESCUENTO_2") = clfloa_gre_descuento_2
    CMD("@CHAR_PRI_PRIORIDAD_ID") = clchar_pri_prioridad_id
    CMD("@FLOA_GRE_DESCUENTO_3") = clFLOA_GRE_DESCUENTO_3
    If var_modifica_registro_gr Then
       CMD.execute
    Else
       CMD("@clave_string") = ""
       CMD.execute
       var_grupo_real_regreso = CMD("@clave_string")
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
Public Function Eliminar(clVcha_gre_grupo_real_id As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn_importacion                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "gruposreales_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@VCHA_GRE_GRUPO_REAL_ID") = clVcha_gre_grupo_real_id
CMD.execute                                         'Ejecutar el PA
SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function

