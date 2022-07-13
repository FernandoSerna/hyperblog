VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_AGENTES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim cmd As New Command


'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_AGE_AGENTE_ID As String, clVcha_age_nombre As String, clVcha_age_telefono As String, clVcha_tag_tipoagente_id As String, clVcha_age_estatus As String, clvcha_alm_almacen_id As String, clVCHA_AGE_RUTA_ARCHIVOS As String, clDTIM_AGE_FECHA_ALTA As String, _
clVCHA_AGE_DOMICILIO As String, clVCHA_AGE_CP As String, clVCHA_COL_COLONIA_ID As String, clVcha_ciu_ciudad_id As String, clVcha_mun_municipio_id As String, clVcha_est_estado_id As String, clVcha_pai_pais_id As String, clVCHA_AGE_EMAIL As String, clVCHA_CAN_CANAL_VENTA_ID As String) As Boolean

Dim rs As ADODB.recordSet, Cadena As String
Anadir = True
'On Error GoTo HELL
Set cmd.ActiveConnection = cnn_importacion

cmd.CommandType = adCmdStoredProc
    
If var_modifica_registro_agente Then
    cmd.CommandText = "AGENTES_M"
Else
    cmd.CommandText = "AGENTES_I"
End If
    cmd("@Vcha_emp_empresa_id") = clVCHA_EMP_EMPRESA_ID
    cmd("@Vcha_age_agente_id") = clVCHA_AGE_AGENTE_ID
    cmd("@Vcha_age_nombre") = clVcha_age_nombre
    cmd("@Vcha_age_telefono") = clVcha_age_telefono
    cmd("@Vcha_tag_tipoagente_id") = clVcha_tag_tipoagente_id
    cmd("@VCHA_AGE_ESTATUS") = clVcha_age_estatus
    cmd("@Vcha_tag_tipoagente_id") = clVcha_tag_tipoagente_id
    cmd("@VCHA_ALM_ALMACEN_ID") = clvcha_alm_almacen_id
    cmd("@VCHA_AGE_RUTA_ARCHIVOS") = clVCHA_AGE_RUTA_ARCHIVOS
    cmd("@dtim_age_fecha_alta") = clDTIM_AGE_FECHA_ALTA
    cmd("@VCHA_PAI_PAIS_ID") = clVcha_pai_pais_id
    cmd("@VCHA_EST_ESTADO_ID") = clVcha_est_estado_id
    cmd("@VCHA_CIU_CIUDAD_ID") = clVcha_ciu_ciudad_id
    cmd("@VCHA_COL_COLONIA_ID") = clVCHA_COL_COLONIA_ID
    cmd("@VCHA_AGE_DOMICILIO") = clVCHA_AGE_DOMICILIO
    cmd("@VCHA_MUN_MUNICIPIO_ID") = clVcha_mun_municipio_id
    cmd("@VCHA_AGE_EMAIL") = clVCHA_AGE_EMAIL
    cmd("@VCHA_AGE_CP") = clVCHA_AGE_CP
    cmd("@VCHA_CAN_CANAL_VENTA_ID") = clVCHA_CAN_CANAL_VENTA_ID
    MsgBox cnn_importacion.ConnectionString
    
    If Not var_modifica_registro_agente Then
       cmd("@CLAVE_STRING") = ""
       cmd.execute
       var_agente_regreso = cmd("@clave_string")
    Else
       cmd.execute
    End If

Set cmd = Nothing
    
    
    
    Set cmd = Nothing
SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Anadir = False
    GoTo SIGUE
End Function
'ELIMINAR
Public Function Eliminar(clVCHA_AGE_AGENTE_ID As String) As Boolean
Dim rs As ADODB.recordSet, Cadena As String
Eliminar = True
On Error GoTo HELL
Set cmd.ActiveConnection = cnn_importacion                      'Esta es la conexión activa
cmd.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
cmd.CommandText = "AGENTES_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    cmd("@Vcha_age_agente_id") = clVCHA_AGE_AGENTE_ID
cmd.execute                                         'Ejecutar el PA

Set cmd = Nothing                                   'Liberar Memoria
SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function

