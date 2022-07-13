VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_SIP_DETALLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String

Dim cmd As New Command



'ANADIR UN REGISTRO
Public Function Anadir(clBint_det_detalle_id As String, clVcha_art_articulo_id As String _
, clBint_det_cantidad As String, clMon_det_precio As String, clMon_det_importe As String _
, clBint_tra_transacciones_id As String, clVcha_det_afectacion As String _
, clBint_tra_folio As String, clVcha_det_mov As String, clVcha_det_status As String, clDtim_aud_fecha As String _
, clVcha_aud_usuario As String, clVcha_aud_maquina As String, clBint_pla_planta_id As String, clFloa_det_cantidadsurtida As String _
, clFloa_art_existenciaant As String, clFloa_importeult As String, clVcha_det_almacen As String, clVcha_det_añoinventario As String) As Boolean

Anadir = True
'On Error GoTo HELL

Set cmd.ActiveConnection = cnn                      'Esta es la conexión activa
cmd.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA

If var_modifica_registro Then
    cmd.CommandText = "DETALLE_M"                         'Abrir Procedimiento Almacenado para Actualizar Cambios
Else
    cmd.CommandText = "DETALLE_I"                         'Abrir Procedimiento Almacenado y Agregar Banco
End If
    cmd("@Bint_det_detalle_id") = clBint_det_detalle_id 'TB_DFOLIOS uno por cada articulo
    cmd("@Vcha_art_articulo_id") = clVcha_art_articulo_id 'CODIGO
    cmd("@Floa_det_cantidad") = IIf(clBint_det_cantidad = "", 0, clBint_det_cantidad)
    cmd("@Mon_det_precio") = IIf(clMon_det_precio = "", 0, clMon_det_precio)
    cmd("@Mon_det_importe") = Val(clMon_det_precio) * Val(clBint_det_cantidad) 'clMon_det_importe 'Replace(clMon_det_importe = "", " ", "")
    cmd("@Bint_tra_transacciones_id") = clBint_tra_transacciones_id 'TB_TFOLIOS
    cmd("@Vcha_det_afectacion") = clVcha_det_afectacion 'RESTA
    cmd("@Bint_DET_folio") = clBint_tra_folio 'TB_FOLIOS
    cmd("@Vcha_DET_movimiento") = clVcha_det_mov 'SALVEDO
    cmd("@Vcha_det_status") = clVcha_det_status ''
        
    cmd("@Dtim_aud_fecha") = clDtim_aud_fecha
    cmd("@Vcha_aud_usuario") = clVcha_aud_usuario
    cmd("@Vcha_aud_maquina") = clVcha_aud_maquina
    cmd("@Bint_pla_planta_id") = clBint_pla_planta_id '4, 1, 6, ETC
    cmd("@Floa_det_cantidadsurtida") = IIf(clFloa_det_cantidadsurtida = "", 0, clFloa_det_cantidadsurtida) '0
    cmd("@Floa_art_existenciaant") = IIf(clFloa_art_existenciaant = "", 0, clFloa_art_existenciaant) '0
    cmd("@Floa_art_importeult") = IIf(clFloa_importeult = "", 0, clFloa_importeult) '0
    cmd("@Vcha_det_almacen") = clVcha_det_almacen
    cmd("@VCHA_DET_AÑOINVENTARIO") = clVcha_det_añoinventario '2005

cmd.execute                                         'Ejecutar el PA

Set cmd = Nothing                                   'Liberar Memoria

Exit Function
HELL:
    Anadir = False
End Function



'ELIMINAR
Public Function Eliminar(clBint_det_detalle_id As String) As Boolean
Dim rs As ADODB.recordSet, Cadena As String
Eliminar = True
On Error GoTo HELL

Set cmd.ActiveConnection = cnn                      'Esta es la conexión activa
cmd.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
cmd.CommandText = "DETALLE_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
    cmd("@Bint_det_detalle_id") = clBint_det_detalle_id
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





