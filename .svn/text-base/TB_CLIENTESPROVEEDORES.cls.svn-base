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
Public Function Anadir(clVcha_clp_clave_id As String, clVcha_clp_nombre As String, clVcha_clp_representante As String, _
clinte_clp_tipo As String, cldtim_clp_fecha_captura As String, clVcha_ven_vendedor_id As String, clVcha_rut_ruta_id As String, clVcha_clp_curp As String, _
clVcha_clp_rfc As String, clVcha_mon_moneda_id As String, clVcha_pla_plazo_id As String, clVcha_tcl_tipo_cliente_id As String, _
clVcha_lis_lista_id As String, clVcha_can_canal_venta_id As String, clVcha_des_descripcion_id As String, _
clVcha_tra_transporte_id As String, clVcha_clp_estatus As String, clVCHA_TIT_TITULAR_ID As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL
Set cmd.ActiveConnection = cnn
cmd.CommandType = adCmdStoredProc
    
If var_modifica_registro Then
    cmd.CommandText = "clientesproveedores_M"
Else
    cmd.CommandText = "clientesproveedores_I"
End If
    cmd("@Vcha_clp_clave_id") = clVcha_clp_clave_id
    cmd("@Vcha_clp_nombre") = clVcha_clp_nombre
    cmd("@Vcha_clp_representante") = clVcha_clp_representante
    cmd("@inte_clp_tipo") = clinte_clp_tipo
    cmd("@dtim_clp_fecha_captura") = cldtim_clp_fecha_captura
    cmd("@Vcha_ven_vendedor_id") = clVcha_ven_vendedor_id
    cmd("@Vcha_rut_ruta_id") = clVcha_rut_ruta_id
    cmd("@Vcha_clp_curp") = clVcha_clp_curp
    cmd("@Vcha_clp_rfc") = clVcha_clp_rfc
    cmd("@Vcha_mon_moneda_id") = clVcha_mon_moneda_id
    cmd("@Vcha_pla_plazo_id") = clVcha_pla_plazo_id
    cmd("@Vcha_tcl_tipo_cliente_id") = clVcha_tcl_tipo_cliente_id
    cmd("@Vcha_lis_lista_id") = clVcha_lis_lista_id
    cmd("@Vcha_can_canal_venta_id") = clVcha_can_canal_venta_id
    cmd("@Vcha_des_descripcion_id") = clVcha_des_descripcion_id
    cmd("@Vcha_tra_transporte_id") = clVcha_tra_transporte_id
    cmd("@Inte_clp_estatus") = clVcha_clp_estatus
    cmd("@Vcha_tit_titular_id") = clVCHA_TIT_TITULAR_ID

cmd.execute
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
Public Function Eliminar(clVcha_clp_clave_id As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set cmd.ActiveConnection = cnn                      'Esta es la conexión activa
cmd.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
cmd.CommandText = "clientesproveedores_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    cmd("@Vcha_clp_clave_id") = clVcha_clp_clave_id
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


