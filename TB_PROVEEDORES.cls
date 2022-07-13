VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_PROVEEDORES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command


'ANADIR UN REGISTRO
Public Function Anadir(clvcha_pro_proveedor_id As String, clVcha_pro_nombre As String, clVcha_pro_representante As String, _
cldtim_pro_fecha_alta As String, clVCHA_PRO_RFC As String, clVcha_pai_pais_id As String, clVcha_est_estado_id As String, clVCHA_MUN_MUNICIPIO_ID As String, clVcha_ciu_ciudad_id As String, clvcha_pro_colonia As String, clvcha_pro_direccion As String, clvcha_pro_cp As String, clvcha_pro_telefono As String) As Boolean

Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    
If var_modifica_registro_proveedor Then
    CMD.CommandText = "proveedores_M"
Else
    CMD.CommandText = "proveedores_I"
End If
    
    
    CMD("@VCHA_PRO_PROVEEDOR_ID") = clvcha_pro_proveedor_id
    CMD("@VCHA_PRO_NOMBRE") = clVcha_pro_nombre
    CMD("@VCHA_PRO_REPRESENTANTE") = clVcha_pro_representante
    CMD("@DTIM_PRO_FECHA_ALTA") = cldtim_pro_fecha_alta
    CMD("@VCHA_PRO_RFC") = clVCHA_PRO_RFC
    CMD("@VCHA_PAI_PAIS_ID") = clVcha_pai_pais_id
    CMD("@VCHA_EST_ESTADO_ID") = clVcha_est_estado_id
    CMD("@VCHA_MUN_MUNICIPIO_ID") = clVCHA_MUN_MUNICIPIO_ID
    CMD("@VCHA_CIU_CIUDAD_ID") = clVcha_ciu_ciudad_id
    CMD("@VCHA_PRO_COLONIA") = clvcha_pro_colonia
    CMD("@VCHA_PRO_DIRECCION") = clvcha_pro_direccion
    CMD("@VCHA_PRO_CP") = clvcha_pro_cp
    CMD("@VCHA_PRO_telefono") = clvcha_pro_telefono
    


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
Public Function Eliminar(clvcha_pro_proveedor_id As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "proveedores_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@Vcha_pro_proveedor_id") = clvcha_pro_proveedor_id
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




