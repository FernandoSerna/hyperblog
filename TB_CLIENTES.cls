VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_CLIENTES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim cmd As New Command


'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_CLI_CLAVE_ID As String, clVcha_cli_nombre As String, clVcha_cli_representante As String, _
cldtim_cli_fecha_captura As String, clVCHA_VEN_VENDEDOR_ID As String, CLVCHA_RUT_RUTA_ID As String, clVcha_cli_curp As String, _
clVcha_cli_rfc As String, clVCHA_MON_MONEDA_ID As String, clVcha_pla_plazo_id As String, clVCHA_TCL_TIPO_CLIENTE_ID As String, _
clVcha_lis_lista_id As String, clVcha_tra_transporte_id As String, clVcha_des_descripcion_id As String, _
clinte_cli_agrupador As String, clinte_cli_estatus As String, clVCHA_TIT_TITULAR_ID As String, clchar_pri_prioridad_id As String, clvcha_cli_email As String, _
clVcha_pai_pais_id As String, clVcha_est_estado_id As String, clVcha_ciu_ciudad_id As String, clvcha_cli_colonia As String, clvcha_cli_direccion As String, clvcha_cli_cp As String, clVcha_mun_municipio_id As String, clINTE_CLI_ENVIO_FACTURA As Integer, clINTE_CLI_ASIGNACION_CATALOGOS As Integer, clVcha_cli_clave_anterior_id As String) As Boolean
Dim rs As ADODB.recordSet, Cadena As String
Anadir = True
'On Error GoTo HELL
Set cmd.ActiveConnection = cnn_importacion
cmd.CommandType = adCmdStoredProc
    
If var_modifica_registro_cliente Then
    cmd.CommandText = "clientes_M"
Else
    cmd.CommandText = "clientes_I"
    If UCase(parametros(0)) = "ADMCDINDUSTRIAL" Then
       clVcha_pla_plazo_id = "4"
    End If
End If
    cmd("@Vcha_cli_clave_id") = clVCHA_CLI_CLAVE_ID
    cmd("@Vcha_cli_nombre") = clVcha_cli_nombre
    cmd("@Vcha_cli_representante") = clVcha_cli_representante
    cmd("@dtim_cli_fecha_captura") = cldtim_cli_fecha_captura
    cmd("@VCHA_AGE_AGENTE_ID") = clVCHA_VEN_VENDEDOR_ID
    cmd("@Vcha_rut_ruta_id") = CLVCHA_RUT_RUTA_ID
    cmd("@Vcha_cli_curp") = clVcha_cli_curp
    cmd("@Vcha_cli_rfc") = clVcha_cli_rfc
    cmd("@Vcha_mon_moneda_id") = clVCHA_MON_MONEDA_ID
    cmd("@Vcha_pla_plazo_id") = clVcha_pla_plazo_id
    cmd("@Vcha_tcl_tipo_cliente_id") = clVCHA_TCL_TIPO_CLIENTE_ID
    cmd("@Vcha_lis_lista_id") = clVcha_lis_lista_id
    cmd("@Vcha_tra_transporte_id") = clVcha_tra_transporte_id
    cmd("@VCHA_FAG_FAMILIA_AGRUPADOR_ID") = clVcha_des_descripcion_id
    If clinte_cli_agrupador = "" Then
       clinte_cli_agrupador = "0"
    End If
    cmd("@inte_cli_agrupador") = clinte_cli_agrupador
    cmd("@Inte_cli_estatus") = clinte_cli_estatus
    cmd("@Vcha_tit_titular_id") = clVCHA_TIT_TITULAR_ID
    cmd("@char_pri_prioridad_id") = clchar_pri_prioridad_id
    cmd("@vcha_cli_email") = CStr(Trim(clvcha_cli_email))
    cmd("@VCHA_PAI_PAIS_ID") = clVcha_pai_pais_id
    cmd("@VCHA_EST_ESTADO_ID") = clVcha_est_estado_id
    cmd("@VCHA_CIU_CIUDAD_ID") = clVcha_ciu_ciudad_id
    cmd("@VCHA_CLI_COLONIA") = clvcha_cli_colonia
    cmd("@VCHA_CLI_DIRECCION") = clvcha_cli_direccion
    cmd("@VCHA_CLI_CP") = clvcha_cli_cp
    cmd("@VCHA_MUN_MUNICIPIO_ID") = clVcha_mun_municipio_id
    cmd("@INTE_CLI_ENVIO_FACTURA") = clINTE_CLI_ENVIO_FACTURA
    cmd("@INTE_CLI_ASIGNACION_CATALOGOS") = clINTE_CLI_ASIGNACION_CATALOGOS
    cmd("@Vcha_cli_clave_anterior_id") = clVcha_cli_clave_anterior_id
   If Not var_modifica_registro_cliente Then
       cmd("@clave_string") = ""
       cmd.execute
       var_cliente_regreso = cmd("@clave_string")
    Else
       cmd.execute
    End If

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
Public Function Eliminar(clVCHA_CLI_CLAVE_ID As String) As Boolean
Dim rs As ADODB.recordSet, Cadena As String
Eliminar = True
On Error GoTo HELL
Set cmd.ActiveConnection = cnn_importacion                      'Esta es la conexión activa
cmd.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
cmd.CommandText = "clientes_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    cmd("@Vcha_cli_clave_id") = clVCHA_CLI_CLAVE_ID
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


