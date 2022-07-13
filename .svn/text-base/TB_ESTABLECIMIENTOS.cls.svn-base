VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_ESTABLECIMIENTOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command

'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_TIT_TITULAR_ID As String, clvcha_esb_establecimiento_ID As String, clVCHA_ESB_NOMBRE As String, clVcha_pai_pais_id As String, clVcha_est_estado_id As String, _
clVcha_ciu_ciudad_id As String, clVCHA_COL_COLONIA_ID As String, clVCHA_ESB_DOMICILIO As String, clVCHA_ESB_TELEFONO As String, clCHAR_ESB_FACTURA_CATALOGOS As String, clVcha_mun_municipio_id As String, clVCHA_ESB_CP As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL
Set CMD.ActiveConnection = cnn_importacion
CMD.CommandType = adCmdStoredProc
    
If var_modifica_regsitro_establecimientos Then
    CMD.CommandText = "ESTABLECIMIENTOS_M"
Else
    CMD.CommandText = "ESTABLECIMIENTOS_I"
End If
    CMD("@VCHA_TIT_TITULAR_ID") = clVCHA_TIT_TITULAR_ID
    CMD("@VCHA_ESB_ESTABLECIMIENTO_ID") = clvcha_esb_establecimiento_ID
    CMD("@VCHA_ESB_NOMBRE") = clVCHA_ESB_NOMBRE
    CMD("@VCHA_PAI_PAIS_ID") = clVcha_pai_pais_id
    CMD("@VCHA_EST_ESTADO_ID") = clVcha_est_estado_id
    CMD("@VCHA_CIU_CIUDAD_ID") = clVcha_ciu_ciudad_id
    CMD("@VCHA_COL_COLONIA_ID") = clVCHA_COL_COLONIA_ID
    CMD("@VCHA_ESB_DOMICILIO") = clVCHA_ESB_DOMICILIO
    CMD("@VCHA_ESB_TELEFONO") = clVCHA_ESB_TELEFONO
    CMD("@CHAR_ESB_FACTURA_CATALOGOS") = clCHAR_ESB_FACTURA_CATALOGOS
    CMD("@VCHA_MUN_MUNICIPIO_ID") = clVcha_mun_municipio_id
    CMD("@VCHA_ESB_CP") = clVCHA_ESB_CP
    If Not var_modifica_regsitro_establecimientos Then
       CMD("@clave_establecimiento") = clvcha_esb_establecimiento_ID
       CMD.execute
       var_establecimiento_regreso = CMD("@clave_establecimiento")
    Else
       CMD.execute
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
Public Function Eliminar(clVCHA_TIT_TITULAR_ID As String, clvcha_esb_establecimiento_ID As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn_importacion                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "ESTABLECIMIENTOS_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@VCHA_TIT_TITULAR_ID") = clVCHA_TIT_TITULAR_ID
    CMD("@VCHA_ESB_ESTABLECIMIENTO_ID") = clvcha_esb_establecimiento_ID
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



