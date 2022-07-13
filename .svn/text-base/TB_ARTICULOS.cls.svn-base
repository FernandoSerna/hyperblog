VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_ARTICULOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MensajeError As String
Dim CMD As New Command


'ANADIR UN REGISTRO
Public Function Anadir(clVCHA_ART_ARTICULO_ID As String, clVcha_art_nombre_español As String, clMone_art_precio_base As String, clMone_art_costo_estandar As String, clDtim_art_fecha_baja As String, clDtim_art_fecha_alta As String, clVcha_art_catalogo_inicio As String, clVcha_art_catalogo_vigente As String, clVcha_lic_licencia_id As String, clVcha_art_numero_lic As String, _
clVcha_dis_diseño_id As String, clVcha_lin_linea_id As String, clVcha_sli_sublinea_id As String, clVcha_pro_producto_id As String, clVcha_tar_tipo_articulo_id As String, clVcha_car_clase_id As String, clVcha_art_estampado1 As String, clVcha_art_tipo_estampado1 As String, clVcha_art_estampado2 As String, clVcha_art_tipo_estampado2 As String, _
clVcha_art_color1 As String, clVcha_art_color2 As String, clVcha_art_tono1 As String, clVcha_art_tono2 As String, clinte_art_numero_decorativos As String, clinte_art_fundas As String, clVCHA_USO_USO_ID As String, clVcha_sus_subtipo_uso_id As String, clVcha_tal_talla_id As String, clVCHA_UNI_UNIDAD_ID As String, _
clFloa_art_volumen As String, clfloa_art_tela As String, clvcha_art_composicion As String, clFloa_art_peso As String, clFloa_art_tara As String, clVcha_caj_caja_id As String, clFloa_art_piezas_caja As String, clFloa_art_maximo As String, _
clFloa_art_minimo As String, clFloa_art_punto_reorden As String, clFloa_art_dias_inventario As String, clVcha_ubi_unicacion_id As String, clFloa_art_bulto As String, clInte_art_salida_masiva As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Anadir = True
'On Error GoTo HELL
Set CMD.ActiveConnection = cnn
CMD.CommandType = adCmdStoredProc
    
If var_modifica_registro_articulo Then
    CMD.CommandText = "ARTICULOS_M"
Else
    CMD.CommandText = "ARTICULOS_I"
End If
    CMD("@Vcha_art_articulo_id") = clVCHA_ART_ARTICULO_ID
    CMD("@Vcha_art_nombre_español") = clVcha_art_nombre_español
    CMD("@Mone_art_precio_base") = clMone_art_precio_base
    CMD("@Mone_art_costo_estandar") = clMone_art_costo_estandar
    If clDtim_art_fecha_alta <> "" Then
       CMD("@Dtim_art_fecha_alta") = clDtim_art_fecha_alta
    Else
       CMD("@Dtim_art_fecha_alta") = Null
    End If
    If clDtim_art_fecha_baja <> "" Then
       CMD("@Dtim_art_fecha_baja") = clDtim_art_fecha_baja
    Else
       CMD("@Dtim_art_fecha_baja") = Null
    End If
    CMD("@Vcha_art_catalogo_inicio") = clVcha_art_catalogo_inicio
    CMD("@Vcha_art_catalogo_vigente") = clVcha_art_catalogo_vigente
    CMD("@Vcha_lic_licencia_id") = clVcha_lic_licencia_id
    CMD("@Vcha_art_numero_lic") = clVcha_art_numero_lic
    CMD("@Vcha_dis_diseño_id") = clVcha_dis_diseño_id
    CMD("@Vcha_lin_linea_id") = clVcha_lin_linea_id
    CMD("@Vcha_sli_sublinea_id") = clVcha_sli_sublinea_id
    CMD("@Vcha_pro_producto_id") = clVcha_pro_producto_id
    CMD("@Vcha_tar_tipo_articulo_id") = clVcha_tar_tipo_articulo_id
    CMD("@Vcha_car_clase_id") = clVcha_car_clase_id
    CMD("@VCHA_ART_ESTAMPADO1") = clVcha_art_estampado1
    CMD("@VCHA_ART_ESTAMPADO2") = clVcha_art_estampado2
    CMD("@Vcha_art_tipo_estampado1") = clVcha_art_tipo_estampado1
    CMD("@Vcha_art_tipo_estampado2") = clVcha_art_tipo_estampado2
    CMD("@Vcha_art_color1") = clVcha_art_color1
    CMD("@Vcha_art_color2") = clVcha_art_color2
    CMD("@Vcha_art_tono1") = clVcha_art_tono1
    CMD("@Vcha_art_tono2") = clVcha_art_tono2
    If clinte_art_numero_decorativos = "" Then
       CMD("@INTE_ART_NUMERO_DECORATIVOS") = 0
    Else
       CMD("@INTE_ART_NUMERO_DECORATIVOS") = clinte_art_numero_decorativos
    End If
    If clinte_art_fundas = "" Then
       CMD("@INTE_ART_FUNDAS") = 0
    Else
       CMD("@INTE_ART_FUNDAS") = clinte_art_fundas
    End If
    CMD("@Vcha_uso_uso_id") = clVCHA_USO_USO_ID
    CMD("@Vcha_sus_subtipo_uso_id") = clVcha_sus_subtipo_uso_id
    CMD("@Vcha_tal_talla_id") = clVcha_tal_talla_id
    CMD("@Vcha_uni_unidad_id") = clVCHA_UNI_UNIDAD_ID
    If clFloa_art_volumen = "" Then
       clFloa_art_volumen = "0"
    End If
    If clfloa_art_tela = "" Then
       clfloa_art_tela = "0"
    End If
    CMD("@Floa_art_volumen") = clFloa_art_volumen
    CMD("@FLOA_ART_TELA") = clfloa_art_tela
    CMD("@VCHA_ART_COMPOSICION") = clvcha_art_composicion
    If clFloa_art_peso = "" Then
       CMD("@Floa_art_peso") = 0
    Else
       CMD("@Floa_art_peso") = clFloa_art_peso
    End If
    If clFloa_art_tara = "" Then
       CMD("@Floa_art_tara") = 0
    Else
       CMD("@Floa_art_tara") = clFloa_art_tara
    End If
    CMD("@Vcha_caj_caja_id") = clVcha_caj_caja_id
    If clFloa_art_piezas_caja = "" Then
       CMD("@Floa_art_piezas_caja") = 0
    Else
       CMD("@Floa_art_piezas_caja") = clFloa_art_piezas_caja
    End If
    If clFloa_art_maximo = "" Then
       CMD("@Floa_art_maximo") = 0
    Else
       CMD("@Floa_art_maximo") = clFloa_art_maximo
    End If
    If clFloa_art_minimo = "" Then
       CMD("@Floa_art_minimo") = 0
    Else
       CMD("@Floa_art_minimo") = clFloa_art_minimo
    End If
    If clFloa_art_punto_reorden = "" Then
       CMD("@Floa_art_punto_reorden") = 0
    Else
       CMD("@Floa_art_punto_reorden") = clFloa_art_punto_reorden
    End If
    If clFloa_art_dias_inventario = "" Then
       CMD("@Floa_art_dias_inventario") = 0
    Else
       CMD("@Floa_art_dias_inventario") = clFloa_art_dias_inventario
    End If
    CMD("@Vcha_ubi_unicacion_id") = clVcha_ubi_unicacion_id
    If clFloa_art_bulto = "" Then
       CMD("@Floa_art_bulto") = 0
    Else
       CMD("@Floa_art_bulto") = clFloa_art_bulto
    End If
    CMD("@Inte_art_salida_masiva") = clInte_art_salida_masiva
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
Public Function Eliminar(clVCHA_ART_ARTICULO_ID As String) As Boolean
Dim rs As ADODB.Recordset, Cadena As String
Eliminar = True
'On Error GoTo HELL
    Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
    CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
        
    CMD.CommandText = "ARTICULOS_E"                         'Abrir Procedimiento Almacenado y Agregar Banco
        CMD("@VCHA_ART_ARTICULO_ID") = clVCHA_ART_ARTICULO_ID
    CMD.execute                                         'Ejecutar el PA
    
    Set CMD = Nothing



SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    Eliminar = False
    GoTo SIGUE
End Function

