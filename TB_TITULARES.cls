VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_TITULARES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim cmd As New Command

'ANADIR UN REGISTRO
Public Function Anadir(clVcha_gre_grupo_real_id As String, clVCHA_TIT_TITULAR_ID As String, clVCHA_TIT_NOMBRE As String, clVcha_pai_pais_id As String, clVcha_est_estado_id As String, clVcha_mun_municipio_id As String, _
clVcha_ciu_ciudad_id As String, clVCHA_COL_COLONIA_ID As String, clVCHA_TIT_DOMICILIO As String, clVCHA_TIT_TELEFONO As String, clFLOA_TIT_LIMITE_CREDITO As Variant, clVcha_tit_cp As String) As Boolean
Dim rs As ADODB.recordSet, Cadena As String
Anadir = True
'On Error GoTo HELL
'MsgBox cnn_importacion.ConnectionString
Set cmd.ActiveConnection = cnn_importacion
cmd.CommandType = adCmdStoredProc
    
If var_modifica_registro_titular Then
    cmd.CommandText = "TITULARES_M"
Else
    cmd.CommandText = "TITULARES_I"
End If
    cmd("@VCHA_GRE_GRUPO_REAL_ID") = clVcha_gre_grupo_real_id
    cmd("@VCHA_TIT_TITULAR_ID") = clVCHA_TIT_TITULAR_ID
    cmd("@VCHA_TIT_NOMBRE") = clVCHA_TIT_NOMBRE
    cmd("@VCHA_PAI_PAIS_ID") = clVcha_pai_pais_id
    cmd("@VCHA_EST_ESTADO_ID") = clVcha_est_estado_id
    cmd("@VCHA_MUN_MUNICIPIO_ID") = clVcha_mun_municipio_id
    cmd("@VCHA_CIU_CIUDAD_ID") = clVcha_ciu_ciudad_id
    cmd("@VCHA_COL_COLONIA_ID") = clVCHA_COL_COLONIA_ID
    cmd("@VCHA_TIT_DOMICILIO") = clVCHA_TIT_DOMICILIO
    cmd("@VCHA_TIT_TELEFONO") = clVCHA_TIT_TELEFONO
    cmd("@FLOA_TIT_LIMITE_CREDITO") = clFLOA_TIT_LIMITE_CREDITO
    cmd("@VCHA_TIT_CP") = clVcha_tit_cp
   If Not var_modifica_registro_titular Then
      cmd("@CLAVE_STRING") = ""
      cmd.execute
      var_titular_regreso = cmd("@clave_string")
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
Public Function Eliminar(clVCHA_TIT_TITULAR_ID As String) As Boolean
Dim rs As ADODB.recordSet, Cadena As String
Eliminar = True
On Error GoTo HELL
Set cmd.ActiveConnection = cnn_importacion 'Esta es la conexión activa
cmd.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
cmd.CommandText = "TITULARES_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    cmd("@VCHA_TIT_TITULAR_ID") = clVCHA_TIT_TITULAR_ID
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

