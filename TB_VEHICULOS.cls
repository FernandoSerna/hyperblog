VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_VEHICULOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command


'ANADIR UN REGISTRO
Public Function Anadir(clVcha_veh_placas_id As String, clVcha_veh_modelo As String, clVcha_veh_ano As String, clVcha_ase_aseguradora_id As String, clVcha_ase_poliza As String, clDtim_ase_vigencia_inicio As String, clDtim_ase_vigencia_final As String, clDtim_aud_fecha As String, clVcha_aud_usuario As String, clVcha_aud_maquina As String, clBint_pla_planta_id As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    
If var_modifica_registro_vehiculo Then
    CMD.CommandText = "VEHICULOS_M"
Else
    CMD.CommandText = "VEHICULOS_I"
End If
    CMD("@Vcha_veh_placas_id") = clVcha_veh_placas_id
    CMD("@Vcha_veh_modelo") = clVcha_veh_modelo
    CMD("@Vcha_veh_ano") = clVcha_veh_ano
    CMD("@Vcha_ase_aseguradora_id") = clVcha_ase_aseguradora_id
    CMD("@Vcha_ase_poliza") = clVcha_ase_poliza
    CMD("@Dtim_ase_vigencia_inicio") = clDtim_ase_vigencia_inicio
    CMD("@Dtim_ase_vigencia_final") = clDtim_ase_vigencia_final
    CMD("@Dtim_aud_fecha") = Date
    CMD("@Vcha_aud_usuario") = clVcha_aud_usuario
    CMD("@Vcha_aud_maquina") = clVcha_aud_maquina
    'cmd("@Bint_pla_planta_id") = clBint_pla_planta_id
    CMD("@Bint_pla_planta_id") = 4
    
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
Public Function Eliminar(clVcha_veh_placas_id As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = "VEHICULOS_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    CMD("@Vcha_veh_placas_id") = clVcha_veh_placas_id
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

