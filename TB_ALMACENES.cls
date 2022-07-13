VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_ALMACENES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MensajeError As String
Dim CMD As New Command

'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_EMP_EMPRESA_ID As String, clVCHA_UOR_UNIDAD_ID As String, clvcha_alm_almacen_id As String, clVcha_alm_nombre As String, _
clVcha_pai_pais_id As String, clVcha_est_estado_id As String, clVcha_ciu_ciudad_id As String, clVcha_alm_direccion As String, clvcha_alm_cp As String, _
clVCHA_ALM_AFECTACION As String, clINTE_ALM_SURTIR As Variant, clVCHA_ALM_NETEABLE As String, clINTE_ALM_PRIORIDAD As Variant, clVCHA_ALM_CORREO As String, _
clCHAR_ALM_TIPO As String, clINTE_ALM_RECHAZO As Integer, clinte_alm_calida As Integer, clinte_alm_costeo As Integer, clinte_alm_reempaque As Integer, clinte_alm_tipo_entrada_reempaque As Integer, clinte_alm_sobrantes As Integer, clVCHA_COL_COLONIA_ID As String, clVCHA_MUN_MUNICIPIO_ID As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL
Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    
If var_modifica_registro_almacen Then
    CMD.CommandText = "ALMACENES_M"
Else
    CMD.CommandText = "ALMACENES_I"
End If
    CMD("@VCHA_EMP_EMPRESA_ID") = clVCHA_EMP_EMPRESA_ID
    CMD("@VCHA_UOR_UNIDAD_ID") = clVCHA_UOR_UNIDAD_ID
    CMD("@VCHA_ALM_ALMACEN_ID") = clvcha_alm_almacen_id
    CMD("@VCHA_ALM_NOMBRE") = clVcha_alm_nombre
    CMD("@VCHA_PAI_PAIS_ID") = clVcha_pai_pais_id
    CMD("@VCHA_EST_ESTADO_ID") = clVcha_est_estado_id
    CMD("@VCHA_CIU_CIUDAD_ID") = clVcha_ciu_ciudad_id
    CMD("@VCHA_ALM_DIRECCION") = clVcha_alm_direccion
    CMD("@VCHA_ALM_CP") = clvcha_alm_cp
    CMD("@VCHA_ALM_AFECTACION") = clVCHA_ALM_AFECTACION
    CMD("@INTE_ALM_SURTIR") = clINTE_ALM_SURTIR
    CMD("@VCHA_ALM_NETEABLE") = clVCHA_ALM_NETEABLE
    If Trim(clINTE_ALM_PRIORIDAD) = "" Then
       clINTE_ALM_PRIORIDAD = 0
    End If
    
    CMD("@INTE_ALM_PRIORIDAD") = clINTE_ALM_PRIORIDAD
    CMD("@VCHA_ALM_CORREO") = clVCHA_ALM_CORREO
    CMD("@CHAR_ALM_TIPO") = clCHAR_ALM_TIPO
    CMD("@INTE_ALM_RECHAZO") = clINTE_ALM_RECHAZO
    
    CMD("@INTE_ALM_CALIDAD") = clinte_alm_calida
    CMD("@INTE_ALM_COSTEO") = clinte_alm_costeo
    CMD("@INTE_ALM_REEMPAQUE") = clinte_alm_reempaque
    CMD("@INTE_ALM_TIPO_ENTRADA_REEMPAQUE") = clinte_alm_tipo_entrada_reempaque
    CMD("@INTE_ALM_SOBRANTES") = clinte_alm_sobrantes
    CMD("@vcha_col_colonia_id") = clVCHA_COL_COLONIA_ID
    CMD("@vcha_mun_municipio_id") = clVCHA_MUN_MUNICIPIO_ID
    
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
Public Function Eliminar(clvcha_alm_almacen_id As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "ALMACENES_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@VCHA_ALM_ALMACEN_ID") = clvcha_alm_almacen_id
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







