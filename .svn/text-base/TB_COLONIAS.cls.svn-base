VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_COLONIAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim cmd As New Command



'ANADIR UN REGISTRO
Public Function Anadir(clVcha_pai_pais_id As String, clVcha_est_estado_id As String, clVcha_mun_municipio_id As String, clVcha_ciu_ciudad_id As String, clVCHA_COL_COLONIA_ID As String, clVcha_col_nombre As String, clVcha_col_cp As String) As Boolean
Dim rs As ADODB.recordSet, Cadena As String
Anadir = True
'On Error GoTo HELL
Set cmd.ActiveConnection = cnn
cmd.CommandType = adCmdStoredProc
    
If var_modifica_registro_colonia Then
    cmd.CommandText = "COLONIAS_M"
Else
    cmd.CommandText = "COLONIAS_I"
End If
    cmd("@Vcha_pai_pais_id") = clVcha_pai_pais_id
    cmd("@VCHA_EST_ESTADO_id") = clVcha_est_estado_id
    cmd("@vcha_mun_municipio_id") = clVcha_mun_municipio_id
    cmd("@Vcha_ciu_ciudad_id") = clVcha_ciu_ciudad_id
    cmd("@Vcha_col_colonia_id") = clVCHA_COL_COLONIA_ID
    cmd("@Vcha_col_nombre") = clVcha_col_nombre
    cmd("@Vcha_col_cp") = clVcha_col_cp
    If Not var_modifica_registro_colonia Then
       cmd("@clave_COLONIA") = clVCHA_COL_COLONIA_ID
       cmd.execute
       var_colonia_regreso = cmd("@clave_COLONIA")
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
Public Function Eliminar(clVCHA_COL_COLONIA_ID As String) As Boolean
Dim rs As ADODB.recordSet, Cadena As String
Eliminar = True
On Error GoTo HELL
Set cmd.ActiveConnection = cnn                      'Esta es la conexión activa
cmd.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
cmd.CommandText = "COLONIAS_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    cmd("@Vcha_col_colonia_id") = clVCHA_COL_COLONIA_ID
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


