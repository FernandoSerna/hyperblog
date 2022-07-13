VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_DETALLE_AGRUPADORES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim cmd As New Command
'ANADIR UN REGISTRO
Public Function Anadir(clVcha_agr_agrupador_id As String, clINTE_DEA_TIPO As Integer, clVcha_art_articulo_id As String, clVCHA_LIN_LINEA_ID As String, clVCHA_SLI_SUBLINEA_ID As String, clVcha_pro_producto_id As String, clVcha_tar_tipo_articulo_id As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
On Error GoTo HELL

Set cmd.ActiveConnection = cnn
cmd.CommandType = adCmdStoredProc
    
If var_modifica_registro Then
    cmd.CommandText = "detalle_agrupadores_M"
Else
    cmd.CommandText = "detalle_agrupadores_I"
End If
    cmd("@VCHA_AGR_AGRUPADOR_ID") = clVcha_agr_agrupador_id
    cmd("@INTE_DEA_TIPO") = clINTE_DEA_TIPO
    cmd("@VCHA_ART_ARTICULO_ID") = clVcha_art_articulo_id
    cmd("@VCHA_LIN_LINEA_ID") = clVCHA_LIN_LINEA_ID
    cmd("@VCHA_SLI_SUBLINEA_ID") = clVCHA_SLI_SUBLINEA_ID
    cmd("@VCHA_PRO_PRODUCTO_ID") = clVcha_pro_producto_id
    cmd("@VCHA_TAR_TIPO_ARTICULO_ID") = clVcha_tar_tipo_articulo_id
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
Public Function Eliminar(clVcha_agr_agrupador_id As String, clINTE_DEA_TIPO As String, _
clVcha_art_articulo_id As String, clVCHA_LIN_LINEA_ID As String, clVCHA_SLI_SUBLINEA_ID As String, _
clVcha_pro_producto_id As String, clVcha_tar_tipo_articulo_id As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
On Error GoTo HELL
Set cmd.ActiveConnection = cnn                      'Esta es la conexión activa
cmd.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
cmd.CommandText = "detalle_agrupadores_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    cmd("@VCHA_AGR_AGRUPADOR_ID") = clVcha_agr_agrupador_id
    cmd("@INTE_DEA_TIPO") = clINTE_DEA_TIPO
    cmd("@VCHA_ART_ARTICULO_ID") = clVcha_art_articulo_id
    cmd("@VCHA_LIN_LINEA_ID") = clVCHA_LIN_LINEA_ID
    cmd("@VCHA_SLI_SUBLINEA_ID") = clVCHA_SLI_SUBLINEA_ID
    cmd("@VCHA_PRO_PRODUCTO_ID") = clVcha_pro_producto_id
    cmd("@VCHA_TAR_TIPO_ARTICULO_ID") = clVcha_tar_tipo_articulo_id
cmd.execute                                         'Ejecutar el PA

SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function


