Attribute VB_Name = "md"



Option Explicit
Dim x As String
Global var_aplica_PL As Integer
Global var_archivo_lote As String
Global var_cambio_embarque As Integer
Global var_pedido_cambio_embarque As String
Global var_cliente_costales As String
Global var_caja_pedido_padre As Integer
Global var_tipo_caja_padre As String
Global var_lote_anterior As Integer
Global var_lote_padre As Integer
Global var_caja_padre As Integer
Global var_embarque_unir As Double
Global var_pedido_unir As Double
Global var_global_unidad As String
Global var_global_nombre_unidad As String
Global var_global_volumen_unidad As String
Global var_permiso_cambiar_transporte As Integer
Global var_guia_aduana As String
Global var_contingencia As Integer
Global var_metodo_fraccionado As Integer
Global var_cn_frontera As String
Global var_transporte_global As String

Global var_usuario_reimpresion As String
Global var_usuario_h_x_h As String
Global var_leyenda_reimpresion As String
Global var_tipo_caja_sello As String
Global var_sello As String
Global indice As Double
Global Linea As String
Global var_codigo_kanban As String
Global var_subinventario_kanban As String
Global var_tipo_asigna_ruta As Integer
Global var_embarque_ruta As Double
Global var_ruta_distribucion As String
Global var_nombre_ruta_distribucion As String
Global var_ruta_cliente As String
Global var_nombre_ruta_cliente As String
Global var_embarque_costales As Double
Global var_codigo_complemento As String
Global var_descripcion_complemento As String
Global var_embarque_lotes As Double
Global var_ejecutar_programa As Integer
Global var_volumen_transporte As Double
Global var_usuario_reimprimir_etiqueta As String
Global var_contraseña_reimprimir_etiqueta As String
Global var_autoriza_REIMPRESION As Integer
Global var_usuario_cerrar_pantalla As String
Global var_contraseña_cerrar_pantalla As String

Global var_usuario_cambios_distribucion As String
Global var_contraseña_cambios_distribucion As String



Global var_orden_depurar As Double
Global var_lote_depurar As Integer
Global var_si_permiso As Integer
Global var_password_permiso As String
Global var_nombre_usuario As String
Global var_dvr_texto As String
Global var_dvr_texto_ip As String
Global var_puerto_texto As String
Global var_modo_texto_ip As Integer
Global var_baudios As String
Global var_observaciones_auditoria As String
Global var_prueba_2 As Integer
Global var_prueba As Integer
Global var_embarque_global As Double
Global var_pedido_global As Double
Global var_salida_cajas As Integer
Global var_anden_global As Integer
Global var_embarque_auditar As Double
Global var_caja_auditar As Double
Global var_nombre_agente_asignar As String
Global var_anden_asignar As Integer
Global var_embarque_asignar As Double
Global var_agente_asignar As String
Global var_bandera_asignacion As Integer
Global var_codigo_busqueda As String
Global var_descripcion_busqueda As String
Global var_nombre_caja As String
Global var_consecutivo_tiendas As Integer
Global var_webservice As String
Global var_webservice_correo As String
Global var_webservice_texto As String
Global var_webservice_uuid As String
Global var_sello_caja As String
Global var_pedido_tienda As Double
Global var_responsabilidad_facturacion As String
Global var_tipo_pedido_ventas_directas As Double
Global var_clave_lista_precios As Double
Global var_referencia_global_dev As String
Global var_clave_almacen_devolucion As String
Global var_numero_folio_devoluciones As Double
Global var_tipo_depurado As Double
Global var_cadena_pedidos_global As String
Global var_almacen_global As String
Global var_nombre_almacen_global As String
Global var_encontro_global As Double
Global var_codigo_global As String
Global var_cantidad_global As Double
Global var_nombre_movimiento_global As String
Global var_numero_nota_traspaso_n As String
Global var_numero_nota_traspaso As String
Global var_almacen_destino_traspaso As String
Global var_almacen_origen_traspaso As String
Global var_numero_jaula As Double
Global var_descripcion_recepcion As String
Global var_tipo_embarque As Integer
Global var_conexion_oracle As String
Global var_oracle_tipo_movimiento As String
Global var_clave_establecimiento_global As String
Global var_clave_titular_global As String
Global var_tipo_datos_adicionales As Integer
Global var_nombre_cliente_ad As String
Global var_paterno_cliente_ad As String
Global var_materno_cliente_ad As String
Global var_calle_cliente_ad  As String
Global var_numero_cliente_ad As String
Global var_clave_tel_pais_ad As String
Global var_clave_tel_estado_ad As String
Global var_numero_interno_cliente_ad As String
Global var_aplicar_nota_credito As Integer
Global var_clave_movimiento_nc As String
Global var_numero_nc  As Integer
Global var_planta_transito_global As String
Global var_ruta_documentos_electronicos_pdf As String
Global var_lista_precios_global As String
Global var_descripcion_global As String
Global var_precio_global As Double
Global var_ruta_factura_pdf As String
Global var_pedido_internet As Double
Global var_cliente_pedido_internet As String
Global var_ruta_documentos_electronicos As String
Global var_si_anticipo As Boolean
Global var_archivo_buscar As String
Global var_codigo_anticipo As String
Global var_importe_anticipo As Double
Global var_consecutivo_anticipo As Integer
Global var_cliente_anticipo As String
Global var_numero_traspaso_cantia As Double
Global var_almacen_traspaso_cantia As String
Global var_acepta_traspaso_global As Integer
Global var_clave_movimiento_apartados_Cantia As String
Global var_consecutivo_apartados_Cantia As Integer
Global var_cadena_promocion_171209 As String
Global var_tipo_cambio_global As Double
Global var_nota_traspasos As String
Global var_nota_traspasos_transito As String
Global var_servidor As String
Global var_numero_embarque_global As Double
Global var_nombre_empresa As String
Global var_posible_limite_credito As Integer
Global var_kanban_es_un_kanban As String
Global var_kanban_almacen_id As String
Global var_kanban_articulo_id As String
Global var_kanban_exito As String
Global var_kanban_mensaje As String
Global var_kanban_numero_linea As Double
Global var_codigo_seleccionado As String
'Global var_sello_caja As String
Global var_posible_kanban As Integer


Global var_conexion_traspasOS_Tiendas As String
Global var_tarjeta_kanban As String
Global var_bd_reportes As String
Global var_conexion_reportes As String
Global var_sr_movimientos As String
Global var_bd_movimientos As String

Global var_sr_reportes As String

Global var_numero_embarque As Double
Global var_embarque_packing_list As Double
Global var_consecutivo_packing_list As Double
Global var_agente_packing_list As String
Global var_nombre_reporte As String
Global var_nombre_paqueteria As String
Global var_clave_caja As String
'Global var_nombre_caja As String
Global var_paqueteria As String
Global var_posible_paqueteria As Integer
Global var_tamaño_caja As String
Global var_guia As String
Global var_si_asignacion_paqueteria As Integer
Global var_correo_estado_cuenta As Integer
Global var_tipo_reporte_estado_cuenta As Integer
Global var_clave_estado_cuenta As String
Global var_consecutivo_estado_cuenta As Integer
Global var_trazabilidad As Integer
Global var_posible_entrada As Boolean
Global var_si_elimino As Integer
Global organizacion_OC As String
Global var_unidad_OC As String
Global var_cadena_reporte_articulos As String
Global var_cadena_reporte_articulos_catalogos As String
Global var_cadena_reporte_articulos_familias As String
Global var_cadena_reporte_articulos_tallas As String
Global var_cadena_reporte_articulos_lineas As String
Global var_costo_tela As Double
Global var_codigo_tela As String
Global var_conexion_pedidos_tiendas As String
Global var_conexion_string As String
Global var_conexion_string_sqlquezada2 As String
Global var_conexion_string_distribucion As String
Global VAR_GLOBAL_ACCESO_SORTEO As Integer
Global strRecepcion_ID As String
Global var_tipo_lectura As Integer
Global var_activa_forma_reporte_catalogo_articulos_almacen  As String
Global var_cadena_seguridad As String
Global var_fecha_general As Date
Global fecha_sistema As Date
Global var_fecha_sistema_string As String
Global parametros(10) As String
Global cnn As ADODB.Connection
Global cnn_minegocio As ADODB.Connection
Global cnn_devolucion_anes As ADODB.Connection
Global cnn_lead_time As ADODB.Connection
Global cnn_eflow As ADODB.Connection
Global cnnicg_sql As ADODB.Connection
Global cnn_ver_factura_electronica As ADODB.Connection
Global cnn_pedido_cantia_textilera As ADODB.Connection
Global cnn_estampados As ADODB.Connection
Global cnn_sqlquezada2 As ADODB.Connection
Global cnn_distribucion As ADODB.Connection
Global cnn_admcdindustrial As ADODB.Connection
Global cnn_reportes As ADODB.Connection
Global cnn_excel As ADODB.Connection
Global cnn_clientes_tiendas As ADODB.Connection
Global cnn_etiquetas_textilera As ADODB.Connection
Global cnn_trazabilidad As ADODB.Connection
Global cnn_facturas_ei As ADODB.Connection
Global cnn_sid_estampados As ADODB.Connection
Global cnn_sid_quezada As ADODB.Connection
Global cnn_sip_multibondeados As ADODB.Connection
Global cnn_puntos_monedero As ADODB.Connection
Global x_v As textinsdk.txtIn

Global cnn_compucaja As ADODB.Connection
Global cnn_compucaja_f As ADODB.Connection
Global cnn_compucaja_T As ADODB.Connection
Global cnn_muebles As ADODB.Connection




Global cnn_importacion As ADODB.Connection
Global cnn_icg_posprod As ADODB.Connection
Global cnn_intercompañias As ADODB.Connection
Global cnnoracle As ADODB.Connection
Global cnnoracle_5 As ADODB.Connection
Global cnnoracle_4 As ADODB.Connection
Global cnnoracle_2 As ADODB.Connection
Global cnnoracle_3 As ADODB.Connection
Global cnnicg As ADODB.Connection
'Global cnn_minegocio As ADODB.Connection

Global cnntraspasos_tiendas As ADODB.Connection
Global cnnsorteo As ADODB.Connection

Global cnnaccess As ADODB.Connection
Global cnn_icg_usa As ADODB.Connection

Global rs As ADODB.Recordset
Global rs_bascula As ADODB.Recordset
Global rs_bascula_2 As ADODB.Recordset
Global rs_bascula_3 As ADODB.Recordset
Global rsaux As ADODB.Recordset
Global rsres As ADODB.Recordset
Global rsdet As ADODB.Recordset
Global rsaux1 As ADODB.Recordset
Global rsaux2 As ADODB.Recordset
Global rsaux3 As ADODB.Recordset
Global rsaux4 As ADODB.Recordset
Global rsaux5 As ADODB.Recordset
Global rsaux6 As ADODB.Recordset
Global rsaux7 As ADODB.Recordset
Global rsaux8 As ADODB.Recordset
Global rsaux9 As ADODB.Recordset
Global rsaux10 As ADODB.Recordset
Global rsaux11 As ADODB.Recordset
Global rsaux12 As ADODB.Recordset
Global rsaux13 As ADODB.Recordset
Global rsaux14 As ADODB.Recordset
Global rsaux15 As ADODB.Recordset
Global rsaux16 As ADODB.Recordset
Global rsaux17 As ADODB.Recordset

Global var_negado_desde As Integer
Global var_modifica_registro_agente As Boolean
Global var_modifica_registro_agrupadores As Boolean
Global var_modifica_registro_almacen As Boolean
Global var_modifica_registro_articulo As Boolean
Global var_modifica_registro_bloque As Boolean
Global var_modifica_registro_caja As Boolean
Global var_modifica_registro_canal_venta As Boolean
Global var_modifica_registro_catalogo As Boolean
Global var_modifica_registro_causas_devolucion As Boolean
Global var_modifica_registro_ciudad As Boolean
Global var_modifica_registro_clase_articulo As Boolean
Global var_modifica_registro_clase As Boolean
Global var_modifica_registro_cliente As Boolean
Global var_modifica_registro_colonia As Boolean
Global var_modifica_registro_color As Boolean
Global var_modifica_registro_comision As Boolean
Global var_modifica_registro_detalle_establecimiento As Boolean
Global var_modifica_registro_familia_agrupador As Boolean
Global var_modifica_registro_diseño As Boolean
Global var_modifica_registro_empresa As Boolean
Global var_modifica_registro_equivalencia As Boolean
Global var_modifica_regsitro_establecimientos As Boolean
Global var_modifica_registro_estado As Boolean
Global var_modifica_registro_estampado As Boolean
Global var_modifica_registro_familia_agrupadores As Boolean
Global var_modifica_registro_familia_ariculos As Boolean
Global var_modifica_registro_licencia As Boolean
Global var_modifica_registro_linea As Boolean
Global var_modifica_registro_lista_precios As Boolean
Global var_modifica_registro_menu As Boolean
Global var_modifica_registro_moneda As Boolean
Global var_modifica_registro_movimiento As Boolean
Global var_modifica_registro_municipio As Boolean
Global var_modifica_registro_pais As Boolean
Global var_modifica_registro_plazo As Boolean
Global var_modifica_registro_prioridad As Boolean
Global var_modifica_registro_producto As Boolean
Global var_modifica_registro_proveedor As Boolean
Global var_modifica_registro_puesto As Boolean
Global var_modifica_registro_referencia As Boolean
Global var_modifica_registro_ruta As Boolean
Global var_modifica_registro_sublinea As Boolean
Global var_modifica_registro_subtipoarticulo As Boolean
Global var_modifica_registro_subtipouso As Boolean
Global var_modifica_registro_talla As Boolean
Global var_modifica_registro_tipoagente As Boolean
Global var_modifica_registro_tipoarticulo As Boolean
Global var_modifica_registro_tipocambio As Boolean
Global var_modifica_registro_tipoestampado As Boolean
Global var_modifica_registro_tipopedido As Boolean
Global var_modifica_registro_tipo_cliente As Boolean
Global var_modifica_registro_titular As Boolean
Global var_modifica_registro_tono As Boolean
Global var_modifica_registro_transaccion As Boolean
Global var_modifica_registro_transporte As Boolean
Global var_modifica_registro_unidad As Boolean
Global var_modifica_registro_unidadorganizacional As Boolean
Global var_modifica_registro_uso As Boolean
Global var_modifica_registro_usuario As Boolean
Global var_modifica_registro_vehiculo As Boolean
Global var_modifica_registro_vendedor As Boolean
Global var_modifica_registro_zona As Boolean
Global var_modifica_registro_agrupador_catalogos As Boolean
Global var_modifica_registro_descuentos_pago_correcto As Boolean
Global var_modifica_registro_ubicacion_almacen As Boolean
Global var_modifica_registro_personas As Boolean

Global var_tipo_filtrado_cliente As Integer
Global var_establecimiento_regreso As String
Global var_usuario_regreso As String
Global var_titular_regreso As String
Global var_cliente_regreso As String
Global var_grupo_actual_regreso As String
Global var_grupo_real_regreso As String
Global var_agente_regreso As String
Global var_estado_regreso As String
Global var_municipio_regreso As String
Global var_ciudad_regreso As String
Global var_colonia_regreso As String
Global maximo_pedido As Double
Global var_lista_transportes As Double
Global var_clave_lista_global As String
Global var_nombre_lista_global As String

Global var_aceptar_direccion As Boolean

Global var_dir_codigo_postal As String
Global var_dir_pais As String
Global var_dir_nombre_pais As String
Global var_dir_estado As String
Global var_dir_nombre_estado As String
Global var_dir_municipio As String
Global var_dir_nombre_municipio As String
Global var_dir_ciudad As String
Global var_dir_nombre_ciudad As String
Global var_dir_colonia As String
Global var_dir_nombre_colonia As String

Global var_catalogo_articulos As Boolean
Global var_llamado_comisiones As Boolean
Global var_parametros_empresa As String
Global var_parametros_bloque As String
Global var_parametros_menus As String
Global ban_tipo_articulo As String
Global var_nombre_planta As String
Global var_usuario_global As String
Global var_passwor_global As String
Global sw_primera_validacion As Boolean
Global sw_mostrar_forma As Boolean
Global var_nombre_movimiento_embarque_regreso As String
Global var_numero_movimiento_embarque_regreso As Double
Global var_numero_nivel_surtido As Double
Global Tb_usuarios As Tb_usuarios
Global Tb_Bloques As Tb_Bloques
Global TB_MENUS As TB_MENUS
Global clsdate As clsdate
Global TB_PUESTOS As TB_PUESTOS

Global var_modifica_registro As Boolean
Global var_modifica_registro_gr As Boolean
Global var_modifica_registro_ga As Boolean
Global var_numero_planta As String
Global vector_valida_passwords(50) As String
Global var_indice_menu As Byte
Global var_valida_passwords As Boolean
Global var_menus As String
Global var_forma As Form
Global var_swpassword As Boolean
Global var_supervisor As String
Global var_sw_menus As Boolean
Global var_puesto As String
Global var_movimiento As Byte
Global varagrupador As String
Global vardetallearticulo As String
Global vardetallelinea As String
Global vardetallesublinea As String
Global vardetalleproducto As String
Global vardetalletipoarticulo As String
Global varfamiliaagrupador As String
Global vardetalleagrupador As String
Global varpais As String
Global varestado As String
Global varnombrepais As String
Global varnombreestado As String
Global vardetallecliente As String
Global vardetalleclavecliente As String
Global varestablecimiento  As String
Global vartipotitular As Integer
Global vartitular As String
Global varautomaticogrupoactual As Integer
Global varautomaticogruporeal As Integer
Global varfamiliaagrupadores As String

Global var_activa_forma_agentes As String
Global var_activa_forma_agrupadores As String
Global var_activa_forma_agrupadores2 As String
Global var_activa_forma_almacenes As String
Global var_activa_forma_articulos2 As String
Global var_activa_forma_aseguradoras As String
Global var_activa_forma_asigna_causa_devolucion As String
Global var_activa_forma_asigna_pagos_no_aplicados As String
Global var_activa_forma_asignacion_negado As String
Global var_activa_forma_autorizapedidos As String
Global var_activa_forma_bonificaciones As String
Global var_activa_forma_bonificaciones_financieras As String
Global var_activa_forma_cajas As String
Global var_activa_forma_calendario As String
Global var_activa_forma_canalesventas As String
Global var_activa_forma_cancela_cajas As String
Global var_activa_forma_cancela_documentos_existentes As String
Global var_activa_forma_cancela_facturas As String
Global var_activa_forma_cancela_facturas_devolucion As String
Global var_activa_forma_cargapedidos As String
Global var_activa_forma_catalogos As String
Global var_activa_forma_catalogos_canales As String
Global var_activa_forma_catempresas As String
Global var_activa_forma_causas_devolucion As String
Global var_activa_forma_causas_no_otorgamiento As String
Global var_activa_forma_cerrado_pedidos As String
Global var_activa_forma_cheques_devueltos As String
Global var_activa_forma_cheques_devueltos_inicio As String
Global var_activa_forma_ciudades As String
Global var_activa_forma_clasearticulos As String
Global var_activa_forma_clases As String
Global var_activa_forma_clases_cartera As String
Global var_activa_forma_clasificacion_clientes As String
Global var_activa_forma_clientes As String
Global var_activa_forma_clientes2 As String
Global var_activa_forma_clonacionagrupadores As String
Global var_activa_forma_codigo_acceso As String
Global var_activa_forma_colonias As String
Global var_activa_forma_colores As String
Global var_activa_forma_comisiones As String
Global var_activa_forma_concentrado_orden_surtido As String
Global var_activa_forma_costos_predeterminados As String
Global var_activa_forma_descuentos_catalogos As String
Global var_activa_forma_descuentos_promociones As String
Global var_activa_forma_descuentos_pronto_pago As String
Global var_activa_forma_descuentos_volumen As String
Global var_activa_forma_descuentos_volumen_cliente As String
Global var_activa_forma_descuentos_volumen_grupo_actual As String
Global var_activa_forma_descuentos_volumen_grupo_real As String
Global var_activa_forma_descuentos_volumen_titular As String
Global var_activa_forma_descuentos_pronto_pago_cambios As String
Global var_activa_forma_detalle_cajas As String
Global var_activa_forma_detalle_documentos_fiscales As String
Global var_activa_forma_detalle_establecimientos As String
Global var_activa_forma_detalle_lista_precios As String
Global var_activa_forma_detalleagrupadores As String
Global var_activa_forma_direcciones As String
Global var_activa_forma_diseños As String
Global var_activa_forma_ejecuta_sistema As String
Global var_activa_forma_embarques As String
Global var_activa_forma_embarques_activos As String
Global var_activa_forma_embarques_paquetes As String
Global var_activa_forma_embarques_paquetes_2 As String
Global var_activa_forma_empresas As String
Global var_activa_forma_entradas As String
Global var_activa_forma_entradas_compras As String
Global var_activa_forma_entradas_devoluciones As String
Global var_activa_forma_entradas_reempaque As String
Global var_activa_forma_entradas_sin_comparacion As String
Global var_activa_forma_equipos As String
Global var_activa_forma_equivalencias As String
Global var_activa_forma_establecimientos As String
Global var_activa_forma_estados As String
Global var_activa_forma_estados_cuenta As String
Global var_activa_forma_estampados As String
Global var_activa_forma_existen_rapidas As String
Global var_activa_forma_existencias_generales As String
Global var_activa_forma_fact_merc_vistas As String
Global var_activa_forma_factura_embarques As String
Global var_activa_forma_factura_empresas As String
Global var_activa_forma_facturas As String
Global var_activa_forma_familia_agrupadores As String
Global var_activa_forma_generapedido As String
Global var_activa_forma_gruposactuales As String
Global var_activa_forma_gruposreales As String
Global var_activa_forma_informacion_articulos_enviar As String
Global var_activa_forma_informacion_pedido_sugerido As String
Global var_activa_forma_inicio As String
Global var_activa_forma_inventario_documentos As String
Global var_activa_forma_kardex As String
Global var_activa_forma_licencias As String
Global var_activa_forma_lineas As String
Global var_activa_forma_listadeprecios As String
Global var_activa_forma_listamovimientos As String
Global var_activa_forma_listatitulares As String
Global var_activa_forma_listaunidades As String
Global var_activa_forma_menu1 As String
Global var_activa_forma_menu2 As String
Global var_activa_forma_menus As String
Global var_activa_forma_migracion As String
Global var_activa_forma_migrar_informacion_paises As String
Global var_activa_forma_monedas As String
Global var_activa_forma_mov_almacenes As String
Global var_activa_forma_movimientos As String
Global var_activa_forma_municipios As String
Global var_activa_forma_nota_credito_saldos_descuento_financiero As String
Global var_activa_forma_nota_credito_descuento_financiero As String
Global var_activa_forma_notas_cargo As String
Global var_activa_forma_notas_credito As String
Global var_activa_forma_numero_embarque As String
Global var_activa_forma_ordenescompra As String
Global var_activa_forma_ordensurtido As String
Global var_activa_forma_packing_list As String
Global var_activa_forma_paises As String
Global var_activa_forma_passwords As String
Global var_activa_forma_passwords2 As String
Global var_activa_forma_permisos As String
Global var_activa_forma_plazos As String
Global var_activa_forma_portada As String
Global var_activa_forma_prioridades As String
Global var_activa_forma_productos As String
Global var_activa_forma_promociones_inicio_catalogo As String
Global var_activa_forma_proveedores As String
Global var_activa_forma_puestos As String
Global var_activa_forma_rangos_promociones_catalogos As String
Global var_activa_forma_relacion_cobranza As String
Global var_activa_forma_relacion_cobranza_listado As String
Global var_activa_forma_reporte_acumulado_ventas As String
Global var_activa_forma_reporte_ajustes_reempaque As String
Global var_activa_forma_reporte_catalogo_articulos As String
Global var_activa_forma_reporte_comisiones As String
Global var_activa_forma_reporte_concentrado_entradas_salidas As String
Global var_activa_forma_reporte_entradas_produccion As String
Global var_activa_forma_reporte_entradas_salidas As String
Global var_activa_forma_reporte_totalizador_movimientos As String
Global var_activa_forma_reporte_envios_tiendas As String
Global var_activa_forma_reporte_mercancia_empacada_transito As String
Global var_activa_forma_reporte_movimientos As String
Global var_activa_forma_reporte_nivel_surtido As String
Global var_activa_forma_reporte_ordenes_surtido As String
Global var_activa_forma_reporte_ordenes_surtido_pendientes As String
Global var_activa_forma_reporte_valuacion_devoluciones As String
Global var_activa_forma_reporte_valuacion_facturas As String
Global var_activa_forma_rutas As String
Global var_activa_forma_salidas As String
Global var_activa_forma_salidas_empaques As String
Global var_activa_forma_salidas_reempaque As String
Global var_activa_forma_salidas_sin_comparacion As String
Global var_activa_forma_series As String
Global var_activa_forma_sublineas As String
Global var_activa_forma_subtipoarticulos As String
Global var_activa_forma_subtiposusos As String
Global var_activa_forma_supervisor1 As String
Global var_activa_forma_tallas As String
Global var_activa_forma_tipoagentes As String
Global var_activa_forma_tipoarticulos As String
Global var_activa_forma_tipocambio As String
Global var_activa_forma_tipoestampados As String
Global var_activa_forma_tipopedidos As String
Global var_activa_forma_tiposclientes As String
Global var_activa_forma_titulares As String
Global var_activa_forma_tonos As String
Global var_activa_forma_transportes As String
Global var_activa_forma_traspasos As String
Global var_activa_forma_traspasos_calidad As String
Global var_activa_forma_traspasosentradas As String
Global var_activa_forma_traspasossalidas As String
Global var_activa_forma_unidades As String
Global var_activa_forma_unidadesorganizacionales As String
Global var_activa_forma_usos As String
Global var_activa_forma_usuarios As String
Global var_activa_forma_vehiculos As String
Global var_activa_forma_vigencias_catalogo_canal_Venta As String
Global var_activa_forma_vistasprevias As String
Global var_activa_forma_zonas As String
Global var_activa_forma_principal_configuracion As String
Global var_activa_forma_aplicacion_descuentos_canal_venta As String
Global var_activa_forma_agrupador_catalogos As String
Global var_activa_forma_reporte_valuacion_facturacion_catalogos As String
Global var_activa_forma_descuentos_pago_correcto As String
Global var_activa_forma_asignacion_catalogo_lista_precios As String
Global var_activa_forma_listado_almacenes As String
Global var_activa_forma_bloqueos As String
Global var_activa_forma_movimientos_bloqueados As String
Global var_activa_forma_personas As String
Global var_activa_forma_embarques_bloqueados As String
Global var_activa_forma_salidas_proveedor As String
Global var_lote_global As String


Global var_activa_forma_salidas_cajas As String
Global vartipocliente As Integer
Global var_operacion_bitacora As String
Global var_operacion_bitacora_articulos As String
Global var_clave_referencia As String
Global var_clave_movimiento As String
Global var_despliega_menu As Boolean
Global var_modifica_cliente As Boolean
Global var_clave_usuario_global As String
Global var_nombre_usuario_global As String
Global var_apellidos_usuario_global As String
Global var_empresa_global As String
Global var_bloque_global As String
Global var_global_permiso1 As Integer
Global var_global_permiso2 As Integer
Global var_global_permiso3 As Integer
Global var_global_permiso4 As Integer
Global var_global_menu As String
Global var_accion_submenu As String
Global var_opcion_seguridad As Integer
Global var_accion_forma As Integer
Global var_acepta_seguridad As Integer
Global var_usuario_permiso As String
Global var_tipo_permiso As Integer
Global var_movimiento_almacen As String
Global var_agente_seleccionado As String
Global var_empresa As String
Global var_unidad_organizacional As String
Global var_requier_factura As Integer
Global var_tipo_documento As String
Global var_renglones_factura As Integer
Global var_numero_folio_regreso As Double
Global var_paquete As Boolean
Global var_numero_embarque_paquete As Integer
Public MensajeError As String
Public SQL As String
Public execute As Boolean
Global numero_devuelto As String
Global var_tipo_acceso As Integer
Global var_nec_emb As Boolean
Global var_autoriza_mov As Boolean
Global var_subir_directo As Boolean
Global var_causa_devolucion As Boolean
Global var_tipo_detalle_devolucion As Integer
Global canstr As String
Global var_posible_accion As Boolean
Global var_global_supervisor_1 As String
Global var_global_supervisor_2 As String
Global var_global_relectura As Integer
Global var_global_aceptar_demas As Integer
Global var_global_bloqueado As Integer
Global var_tipo_proveedor_movimiento As String
Global var_devolucion_factura As Integer
Global var_reporte_imprimir As String
Dim CMD As New Command
Global var_verificador As Boolean
Global var_clave_lista As String
Dim var_top As Double
Dim var_left As Double
Global var_activa_menu As Boolean
Global var_es_embarque As Boolean
Global var_conexion_sorteo As String
' The column selected for sorting.
Public m_SortColumn As Integer

' The current sort order.
Public m_SortOrder As SortSettings
Global var_numero_embarque_regreso As Double


Global Const ODBC_ADD_DSN = 1 ' Add data source
Global Const ODBC_CONFIG_DSN = 2 ' Configure (edit) data source
Global Const ODBC_REMOVE_DSN = 3 ' Remove data source
Global Const vbAPINull As Long = 0& ' NULL Pointer
'delcaracion de funciones


Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" ( _
            ByVal hwndParent As Long, ByVal fRequest As Long, ByVal _
             lpszDriver As String, ByVal lpszAttributes As String) As Long


'___________________________ Type para Clasificar nivel nodo__________________

Public Enum ObjectType
    otNone = 0
    otnivel2 = 1
    otnivel3 = 2
    otnivel4 = 3
    otnivel22 = 4
    otNivel32 = 5
    otNivel42 = 6
End Enum

Public SourceNode As Object
Public SourceType As ObjectType
Public TargetNode As Object

' constante para FORMATO DE LIST VIEW
Private Const GWL_STYLE = (-16)
Private Const LVM_FIRST = &H1000
Private Const LVM_GETHEADER = (LVM_FIRST + 31)

Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55
Private Const LVS_EX_FULLROWSELECT = &H20
Private Const HDS_BUTTONS = &H2
Private Const LVM_DELETEITEM = (LVM_FIRST + 8)

Public Declare Function Sleep Lib "kernel32" (ByVal A As Long) As Long
 
 
 





' constante para cajas de mensajeX
Public Const NV_CLOSEMSGBOX As Long = &H5000&
  
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'retarda los mensajes de textos
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long

'para buscar en una lista alfabeticamente

' obtiene el nombre de la maquina
Public Declare Function GetComputerName Lib "kernel32.dll" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

' obtiene el nombre del usuario
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
' despliega automaticos los combos boxes

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
        
Private Const CB_ERR = (-1)                 'combo error code
Private Const CB_FINDSTRING = &H14C         'hex value to invoke find
Private Const CB_FINDSTRINGEXACT = &H158    'hex value to invoke find with exact matching
Private Const CB_SHOWDROPDOWN = &H14F       'hex value to drop down combo

Public Function ListView_DeleteItem(hwnd As Long, ByVal Item As MSComctlLib.ListItem) As Boolean
  ListView_DeleteItem = SendMessage(hwnd, LVM_DELETEITEM, Item, 0)
End Function

'==== Procedimiento para definir que conexion que vamos a usar en toda la aplicacion ====

Public Sub Main()
    var_lista_transportes = 0
    On Error Resume Next
    Set cnn_compucaja = CreateObject("ADODB.connection")
    Set cnn_compucaja_f = CreateObject("ADODB.connection")
    Set cnn_compucaja_T = CreateObject("ADODB.connection")
    Dim x As Integer
    Dim i As Integer, Línea As String
    Dim var_archivo_local As String
    Dim var_audita As Integer
    Set cnn_devolucion_anes = CreateObject("ADODB.connection")
    Set cnn = CreateObject("ADODB.connection")
    Set cnn_lead_time = CreateObject("ADODB.connection")
    Set cnn_eflow = CreateObject("ADODB.connection")
    Set cnnicg_sql = CreateObject("ADODB.connection")
    Set cnn_ver_factura_electronica = CreateObject("ADODB.connection")
    Set cnn_pedido_cantia_textilera = CreateObject("ADODB.connection")
    Set cnn_estampados = CreateObject("ADODB.connection")
    Set cnn_etiquetas_textilera = CreateObject("ADODB.connection")
    Set cnn_sqlquezada2 = CreateObject("ADODB.connection")
    Set cnn_distribucion = CreateObject("ADODB.connection")
    Set cnn_admcdindustrial = CreateObject("ADODB.connection")
    Set cnn_reportes = CreateObject("ADODB.connection")
    Set cnn_importacion = CreateObject("ADODB.connection")
    Set cnnoracle = CreateObject("ADODB.connection")
    Set cnnoracle_5 = CreateObject("ADODB.connection")
    Set cnnoracle_4 = CreateObject("ADODB.connection")
    Set cnnoracle_2 = CreateObject("ADODB.connection")
    Set cnnoracle_3 = CreateObject("ADODB.connection")
    Set cnn_icg_posprod = CreateObject("ADODB.connection")
    Set cnnicg = CreateObject("ADODB.connection")
    Set cnntraspasos_tiendas = CreateObject("ADODB.connection")
    Set cnnaccess = CreateObject("ADODB.connection")
    Set cnn_clientes_tiendas = CreateObject("ADODB.connection")
    Set cnn_trazabilidad = CreateObject("ADODB.connection")
    Set cnnsorteo = CreateObject("ADODB.connection")
    Set cnn_facturas_ei = CreateObject("ADODB.connection")
    Set cnn_sid_estampados = CreateObject("ADODB.connection")
    Set cnn_sid_quezada = CreateObject("ADODB.connection")
    Set cnn_sip_multibondeados = CreateObject("ADODB.connection")
    Set cnn_compucaja = CreateObject("ADODB.connection")
    Set cnn_puntos_monedero = CreateObject("ADODB.connection")
    Set cnn_icg_usa = CreateObject("ADODB.connection")
    Set cnn_minegocio = CreateObject("ADODB.CONNECTION")
    Set cnn_muebles = CreateObject("ADODB.CONNECTION")
    Set rs = CreateObject("ADODB.recordset")
    Set rs_bascula = CreateObject("ADODB.recordset")
    Set rs_bascula_2 = CreateObject("ADODB.recordset")
    Set rs_bascula_3 = CreateObject("ADODB.recordset")
    Set rsaux = CreateObject("ADODB.recordset")
    Set rsaux1 = CreateObject("ADODB.recordset")
    Set rsaux2 = CreateObject("ADODB.recordset")
    Set rsaux3 = CreateObject("ADODB.recordset")
    Set rsaux4 = CreateObject("ADODB.recordset")
    Set rsaux5 = CreateObject("ADODB.recordset")
    Set rsaux6 = CreateObject("ADODB.recordset")
    Set rsaux7 = CreateObject("ADODB.recordset")
    Set rsaux8 = CreateObject("ADODB.recordset")
    Set rsaux9 = CreateObject("ADODB.recordset")
    Set rsaux10 = CreateObject("ADODB.recordset")
    Set rsaux11 = CreateObject("ADODB.recordset")
    Set rsaux12 = CreateObject("ADODB.recordset")
    Set rsaux13 = CreateObject("ADODB.recordset")
    Set rsaux14 = CreateObject("ADODB.recordset")
    Set rsaux15 = CreateObject("ADODB.recordset")
    Set rsaux16 = CreateObject("ADODB.recordset")
    Set rsaux17 = CreateObject("ADODB.recordset")
    
    
    Dim var_si_oracle As Integer
    var_cambio_embarque = 0
    var_pedido_tienda = 0
    var_tipo_lectura = 1
    Open (App.Path + "\SID.SID") For Input As #1
    i = 0
    'Do While Not EOF(1)
    For i = 0 To 9
       Line Input #1, Linea
       On Error GoTo sigue31
       parametros(i) = Linea
       'i = i + 1
    Next
    'Loop
sigue31:
    Close #1
    
    
    Open (App.Path + "\SID.SID") For Input As #1
    i = 0
    var_si_oracle = 0
    var_baudios = ""
    Do While Not EOF(1)
       Line Input #1, Linea
       If i = 8 Then
          If Trim(Linea) = "" Then
             var_si_oracle = 0
          Else
             If Not IsNumeric(Linea) Then
                var_si_oracle = 0
             Else
                If CDbl(Linea) > 0 Then
                   var_si_oracle = 1
                Else
                   var_si_oracle = 0
                End If
             End If
          End If
       End If
       If i = 9 Then
          If Trim(Linea) = "" Then
             var_baudios = "9600,N,8,1"
          Else
             var_baudios = Linea
          End If
       End If
       i = i + 1
    Loop
    Close #1
    
    If var_baudios = "" Then
       var_baudios = "9600,N,8,1"
    End If
    
    If Dir("c:\reportessid", vbDirectory) = "" Then
       MkDir ("c:\reportessid")
    End If
    If Dir("c:\notas_franquicias", vbDirectory) = "" Then
       MkDir ("c:\notas_franquicias")
    End If
    
    If Dir(App.Path + "\xml", vbDirectory) = "" Then
       MkDir (App.Path + "\xml")
    End If
    If Dir(App.Path + "\xml\epcer", vbDirectory) = "" Then
       MkDir (App.Path + "\xml\epcer")
    End If

On Error GoTo SIGUE:
                                                                          
    'FileCopy "\\fscdindustrial\fscdind\update_sip\MovimientosInventario\RespaldoCatalogos.exe", App.Path + "\RespaldoCatalogos.exe"
SIGUE:
    
   
    
    If Dir(App.Path + "\actualizacion.txt", vbArchive) = "" Then
       Open (App.Path + "\actualizacion.txt") For Output As #1
       Print #1, CStr(Date - 1)
       Close #1
    End If
    Dim var_fecha As String
    On Error GoTo sigue2
    Open (App.Path + "\actualizacion.txt") For Input As #1
    Do While Not EOF(1)
       Line Input #1, Linea
       var_fecha = Linea
    Loop
    Close #1
 
         
    var_ejecutar_programa = 0
    If var_fecha <> "" Then
       If CDate(var_fecha) <> Date Then
          Open (App.Path + "\actualizacion.txt") For Output As #1
          If EOF(1) Then
             Print #1, CStr(Date)
          End If
          Close #1
          x = 1
          If x = 1 Then
          'x = Shell(App.Path + "\actualizacion_informacion.exe")
          'x = Shell(App.Path + "\RespaldoCatalogos.exe 93|pvia")
          var_ejecutar_programa = 1
          
          End If
       End If
    Else
       Open (App.Path + "\actualizacion.txt") For Output As #1
       Print #1, CStr(Date)
       Close #1
    End If
    
sigue2:
    If Err.Number = 9 Then
       On Error GoTo sigue7:
       Close #1
       GoTo sigue3
    Else
       On Error GoTo sigue7:
       Close #1
    End If
sigue7:
    'MsgBox Err.Description
sigue3:
On Error GoTo sigue_5:
    Close #1

sigue_5:
    x = 0
    If x = 1 Then
       cnn_pedido_cantia_textilera.Open "Provider=SQLOLEDB.1;Password=" & parametros(3) & ";Persist Security Info=True;User ID=sa;Initial Catalog=SIDTEXTILERA;Data Source=sqlquezada2"
    End If
    'parametros(0) = "admcdindustrial"
    parametros(0) = "ADMCDINDUSTRIAL"
    parametros(1) = "SIDAlmacenbkp"
    'rs.Open "select * from tb_servidores where vcha_Ser_base_datos = 'SIDAlmacenbkp'", cnn_distribucion, adOpenDynamic, adLockOptimistic
    'If Not rs.EOF Then
    '   var_bd_movimientos = rs!vcha_ser_base_Datos_movimientos
    '   var_sr_movimientos = rs!vcha_ser_servidor_movimientos
    '   var_bd_reportes = rs!vcha_ser_base_datos_reportes
    '   var_sr_reportes = rs!vcha_ser_servidor_reportes
    '   var_conexion_reportes = rs!VCHA_SER_CONEXION_REPORTE
    'End If
    'rs.Close
    'cnn_reportes.Open var_conexion_reportes
    var_webservice_texto = "http://intranet:9100/wssid/service1.asmx?wsdl"
    var_webservice_correo = "http://serviciowebcedisdesa.vianney.com.mx/EnviarCorreos.asmx"
    var_webservice_uuid = ""
    var_prueba = 0
    

    'commit 1
    If var_prueba = 0 Then
       'MsgBox "1"
       var_conexion_string = "Provider=SQLOLEDB.1;Password=" & parametros(3) & ";Persist Security Info=True;User ID=sa;Initial Catalog=" & parametros(1) & ";Data Source=" & parametros(0)
       'cnn_muebles.Open "Provider=SQLOLEDB.1;Password=" & "sid2" & ";Persist Security Info=True;User ID=sid2;Initial Catalog=" & "SID2MBL" & ";Data Source=" & "10.4.250.32"
       var_bd_reportes = "SIDAlmacenbkp"
       cnn.Open var_conexion_string
       'MsgBox "2"
       var_webservice = "http://intranet/WsEBS12Prod/wsInterfaceOM.asmx?wsdl"
       cnnoracle_4.Open "Provider=OraOLEDB.Oracle.1;User ID=apps;Data Source=pvia;Extended Properties=;Persist Security Info=True;Password=apps"
       'cnnoracle_4.Open "Provider=OraOLEDB.Oracle.1;User ID=apps;Data Source=pviasy;Extended Properties=;Persist Security Info=True;Password=apps"
       'cnnoracle_4.Open "Provider=OraOLEDB.Oracle.1;User ID=apps;Data Source=tvia;Extended Properties=;Persist Security Info=True;Password=apps"
       'MsgBox "3"
    Else
       parametros(1) = "SIDEbs12bkp"
       var_bd_reportes = "SIDEbs12bkp"
       If var_prueba = 2 Then
          var_webservice = "http://intranet/WsEBS12Test/wsInterfaceOM.asmx?wsdl"
          parametros(0) = "sqlquezada2"
          cnnoracle_4.Open "Provider=OraOLEDB.Oracle.1;User ID=apps;Data Source=tvia;Extended Properties=;Persist Security Info=True;Password=apps"
          var_conexion_string = "Provider=SQLOLEDB.1;Password=" & parametros(3) & ";Persist Security Info=True;User ID=sa;Initial Catalog=SIDEbs12bkp;Data Source=" & parametros(0)
          'cnnoracle_4.Open "Provider=OraOLEDB.Oracle.1;User ID=apps;Data Source=pviasy;Extended Properties=;Persist Security Info=True;Password=apps"
          
          'var_conexion_string = "Provider=SQLOLEDB.1;Password=" & parametros(3) & ";Persist Security Info=True;User ID=sa;Initial Catalog=" & parametros(1) & ";Data Source=" & parametros(0)
          'var_conexion_string = "Provider=SQLOLEDB.1;Password=elia;Persist Security Info=True;User ID=sa;Initial Catalog=SIDAlmacenBkp_USA;Data Source=sqlposusa.VIANNEY.COM.mx"
          cnn.Open var_conexion_string
       Else
          cnnoracle_4.Open "Provider=OraOLEDB.Oracle.1;User ID=apps;Data Source=dvia;Extended Properties=;Persist Security Info=True;Password=apps"
          var_webservice = "http://intranet/WSOracle/wsInterfaceOM.asmx?wsdl"
          var_conexion_string = "Provider=SQLOLEDB.1;Password=" & parametros(3) & ";Persist Security Info=True;User ID=sa;Initial Catalog=" & parametros(1) & ";Data Source=" & parametros(0)
          cnn.Open var_conexion_string
       End If
    End If
    
    
    'cnn_icg_usa.Open "Provider=SQLOLEDB.1;Password=ICGUsa2014;Persist Security Info=True;User ID=ICGAdmin;Initial Catalog=bd1;Data Source=sqlcedishou.VIANNEYcatalog.COM"
                            
    'rs.Open "select * from IT_PEDIDOCOMPRA", cnn_icg_usa, adOpenDynamic, adLockOptimistic
    'rs.Close
                       
    
    cnn.CursorLocation = adUseClient
    
    var_contingencia = 1
    
    rs.Open "select * from tb_oracle_maquinas where maquina = '" + UCase(fun_NombrePc) + "'", cnn, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
       var_audita = IIf(IsNull(rs!METODO_ADUANA), 0, rs!METODO_ADUANA)
       var_dvr_texto = IIf(IsNull(rs!DVR), 0, rs!DVR)
       var_puerto_texto = IIf(IsNull(rs!PUERTO), 0, rs!PUERTO)
       var_metodo_fraccionado = IIf(IsNull(rs!metodo_fraccionado), 0, rs!metodo_fraccionado)
       If var_audita = 1 Then
          var_bandera_asignacion = 0
          var_prueba_2 = 1
       Else
          var_bandera_asignacion = 1
          var_prueba_2 = 0
       End If
    Else
       var_bandera_asignacion = 1
       var_prueba_2 = 0
       var_dvr_texto = 0
       var_puerto_texto = 0
    End If
    rs.Close
    If CDbl(var_dvr_texto) > 0 And CDbl(var_puerto_texto) > 0 Then
       var_modo_texto_ip = 1
       If var_dvr_texto = "1" Then
          var_dvr_texto_ip = "10.6.200.70"
       End If
       If var_dvr_texto = "2" Then
          var_dvr_texto_ip = "10.6.200.71"
       End If
       If var_dvr_texto = "3" Then
          var_dvr_texto_ip = "10.6.200.72"
       End If
       If var_dvr_texto = "4" Then
          var_dvr_texto_ip = "10.60.200.73"
       End If
       If var_dvr_texto = "5" Then
          var_dvr_texto_ip = "10.60.200.74"
       End If
       If var_dvr_texto = "6" Then
          var_dvr_texto_ip = "10.6.200.75"
       End If
       If var_dvr_texto = "7" Then
          var_dvr_texto_ip = "10.4.200.70"
       End If
       If var_dvr_texto = "8" Then
          var_dvr_texto_ip = "10.20.90.5"
       End If
    Else
       var_modo_texto_ip = 0
    End If
    
    cnnicg.Open "Provider=OraOLEDB.Oracle.1;User ID=xxpos; Data Source=pvia;Extended Properties=;Persist Security Info=True;Password=xxpos"
    var_conexion_oracle = cnnoracle_4.ConnectionString
    'rs.Open "SELECT * FROM AR_COLLECTORS", cnnoracle_4, adOpenDynamic, adLockOptimistic
    'rs.Close
    rs.Open "select * from tb_principal", cnn, adOpenDynamic, adLockOptimistic
    var_renglones_factura = IIf(IsNull(rs!INTE_PRI_RENGLONES_FACTURA), 0, rs!INTE_PRI_RENGLONES_FACTURA)
    var_trazabilidad = IIf(IsNull(rs!inte_pri_trazabilidad), 0, rs!inte_pri_trazabilidad)
    rs.Close
    var_archivo_local = App.Path + "\sistema.exe"
    var_fecha_sistema_string = "Sistema Integral de Distribución S.I.D.   [Ultima actualización " + Trim(CStr(FileDateTime(var_archivo_local))) + "]"
    If UCase(parametros(1)) = "SIDALMACENBKP" Then
       var_conexion_pedidos_tiendas = "Provider=OraOLEDB.Oracle.1;Password=mvtosbanca;Persist Security Info=True;User ID=mvtosbanca;Data Source=dbtest"
       var_conexion_sorteo = "Provider=OraOLEDB.Oracle.1;Password=tiendas;Persist Security Info=True;User ID=tiendas;Data Source=dbtest"
       var_conexion_traspasOS_Tiendas = "Provider=OraOLEDB.Oracle.1;Password=tiendas;Persist Security Info=True;User ID=tiendas;Data Source=dbtest"
    Else
       var_conexion_pedidos_tiendas = "Provider=OraOLEDB.Oracle.1;Password=mvtosbanca;Persist Security Info=True;User ID=mvtosbanca;Data Source=oradborc"
       var_conexion_traspasOS_Tiendas = "Provider=OraOLEDB.Oracle.1;Password=tiendas;Persist Security Info=True;User ID=tiendas;Data Source=ap"
    End If
    Call copiar_facturar_exe
    Call copiar_sonidos
    frmsistema_integral.Caption = var_fecha_sistema_string
    var_posible_paqueteria = 0
    If UCase(fun_NombrePc) = "JFSERNA" And UCase(Trim(parametros(0))) = "DBPRUEBAS" Then
       var_posible_kanban = 1
    Else
       var_posible_kanban = 0
    End If
    var_numero_embarque_global = 0
    If var_empresa = "02" Then
       var_cadena_promocion_171209 = "Artículos marcados con * incluyen promoción del 5%   "
    Else
       var_cadena_promocion_171209 = "Artículos marcados con * incluyen promoción  "
    End If
    var_cliente_pedido_internet = ""
    Frmacceso.Show
End Sub


'************************************************************************************

'                   PROCEDIMIENTOS GUARDA TODOS LOS DATOS

'************************************************************************************
 
'==================== Procedimiento para guardar Los menus ==========================
Public Function convierte_numero(ByVal x As String) As String
   Dim var_contador As Integer
   Dim Z As String
   numero_devuelto = ""
   Z = x
   x = ""
   For var_contador = 1 To Len(Trim(Z))
      If Mid(Z, var_contador, 1) <> "," Then
         x = x + Mid(Z, var_contador, 1)
      End If
   Next
   numero_devuelto = x
End Function

Public Function Obtener_Identificador_Recepcion(cnnoracle As ADODB.Connection, ByRef strError) As Boolean
    'Objetos command
    Dim cmdCommand As New ADODB.Command
    
On Error GoTo Error_Insertar_Recepcion
    Obtener_Identificador_Recepcion = True
    strError = ""
    
    Set cmdCommand.ActiveConnection = cnnoracle
    cmdCommand.CommandType = adCmdStoredProc
    cmdCommand.CommandText = "SP_RECEPCION_ID"
    cmdCommand.Parameters.Refresh
    
    cmdCommand.execute
    
    strRecepcion_ID = cmdCommand.Parameters("V_ID").Value
    
    Set cmdCommand = Nothing
    Exit Function
Error_Insertar_Recepcion:
    Obtener_Identificador_Recepcion = False
    strError = Err.Description
    var_posible_entrada = False
End Function


Public Function Insertar_Recepcion(var_orden_compra As Double, var_numero_linea As Double, var_cantidad As Double, var_linea As Double, var_numero_movimiento As Double, var_p_rc_ord_id As Double, ByRef strError As String, var_factura_oracle As String) As Boolean
    'Objetos command
    Dim cmdCommand As New ADODB.Command
    Dim var_si_paso As Boolean
    Dim var_factura_oracle_2 As Double
    Dim var_factura_x As String
    Dim var_j As Integer
'On Error GoTo Error_Insertar_Recepcion
    Insertar_Recepcion = True
    strError = ""
    'MsgBox cnnoracle
    Set cmdCommand.ActiveConnection = cnnoracle
    cmdCommand.CommandType = adCmdStoredProc
    'var_factura_oracle_2 = Cdbl(var_factura_oracle)
    var_factura_x = ""
    For var_j = 1 To Len(var_factura_oracle)
        If IsNumeric(Mid(var_factura_oracle, var_j, 1)) Then
           var_factura_x = var_factura_x + Mid(var_factura_oracle, var_j, 1)
        End If
    Next var_j
    
    var_factura_oracle_2 = CDbl(var_factura_x)
    cmdCommand.CommandText = "SP_RECEPCIONES_INS2"
    cmdCommand.Parameters.Refresh
    cmdCommand.Parameters("P_RC_NUMERO_RECEPCION").Size = 30
    cmdCommand.Parameters("P_RC_RECEPCION_ID").Value = CDbl(strRecepcion_ID)
    cmdCommand.Parameters("P_RC_ORDEN_COMPRA_ID").Value = var_orden_compra 'orden de compra
    cmdCommand.Parameters("P_RC_NUMERO_LINEA").Value = var_numero_linea 'numero de lineam po.line_num
    cmdCommand.Parameters("P_RC_CANTIDAD").Value = var_cantidad
    cmdCommand.Parameters("P_RC_LINEA_ID").Value = var_linea 'po.line_id
    cmdCommand.Parameters("P_RC_ORG_ID").Value = var_p_rc_ord_id 'po.line_id
    cmdCommand.Parameters("P_RC_FACTURA").Value = CDbl(var_factura_oracle_2)
    'cmdCommand.Parameters(8).Value = var_factura_oracle
    
    'cmdCommand.Parameters("P_RC_USUARIO_ID").Value = 0
     
    If var_empresa = "18" Then
       cmdCommand.Parameters("P_RC_NUMERO_RECEPCION").Value = "SIDT_" + Trim(CStr(var_numero_movimiento)) 'mi numero de movimiento
    Else
       If var_empresa = "31" Then
          cmdCommand.Parameters("P_RC_NUMERO_RECEPCION").Value = "CANT_" + Trim(CStr(var_numero_movimiento)) 'mi numero de movimiento
       Else
          cmdCommand.Parameters("P_RC_NUMERO_RECEPCION").Value = "SID_" + Trim(CStr(var_numero_movimiento)) 'mi numero de movimiento
       End If
    End If
    'MsgBox CDbl(rs!P_RC_LINEA_ID)
    cmdCommand.execute
    
    Set cmdCommand = Nothing
    Exit Function
Error_Insertar_Recepcion:
    Insertar_Recepcion = False
    strError = Err.Description
    MsgBox strError, vbOKOnly
    
End Function



Public Function calcula_verificador(mcodigo As String)
   Dim sum1 As Integer
   Dim sum2 As Integer
   Dim icont As Integer
   Dim VERIFICADOR As Integer
   Dim verificador2 As Integer
   Dim var_codigo As String
   Dim longitud As Integer
   Dim msuma As Integer
   sum1 = 0
   sum2 = 0
   longitud = Len(mcodigo)
   For icont = 1 To longitud - 1
      If ((icont / 2) - Int((icont / 2))) = 0 Then
         sum2 = sum2 + Val(Mid(mcodigo, icont, 1))
      Else
         sum1 = sum1 + Val(Mid(mcodigo, icont, 1))
      End If
   Next icont
   msuma = sum1 * 13 + sum2
   VERIFICADOR = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
   If VERIFICADOR = 10 Then
      VERIFICADOR = 0
   End If
   verificador2 = Val(Mid(Trim(mcodigo), longitud, 1))
   var_verificador = False
   If VERIFICADOR = verificador2 Then
      var_verificador = True
   End If
End Function

Public Sub SALIR()
   End
End Sub

Public Sub pro_guardar_menus(ByVal forma As Form)
Dim CMD As New Command                                  'Este es el objeto Command que declaramos

Set CMD.ActiveConnection = cnn                          'Esta es la conexión activa
    CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
    CMD.CommandText = "MENUS_I"                     'Abrir Procedimiento Almacenado y Agregar Banco
    With forma
    
        CMD("@VCHA_MEN_MENU_ID") = "MNU_" & UCase(.Text1(0))
        CMD("@VCHA_MEN_NIVEL") = .Text1(2)
        CMD("@VCHA_MEN_DESCRIPC") = .Text1(0)
        CMD("@VCHA_MEN_TOOLTIP") = .Text1(1)
        CMD("@VCHA_MEN_COMPONEN") = ""
        CMD("@VCHA_MEN_MODULO") = ""
        CMD("@VCHA_MEN_STATUS") = "A"

    CMD.execute                                         'Ejecutar el PA
    Set CMD = Nothing                                   'Liberar Memoria
    End With

End Sub


Public Sub inserta_codigo_barras(i_organizacion As Double, i_almacen As String, i_movimiento As String, i_numero As Double, i_codigo_barras As String, i_codigo As String, i_cantidad As Double)
   Dim comandoORA As New ADODB.Command
   Dim parametro As ADODB.Parameter
   Dim objConn As New ADODB.Connection
   Dim objCmd As New ADODB.Command
   Dim objParm As ADODB.Parameter
   Dim strconsulta As String
   strconsulta = "select * from xxvia_tb_transacciones where organizacion = ? and almacen = ? and movimiento = ? and numero = ? and codigo_barras = ?"
   With comandoORA
        .ActiveConnection = cnnoracle_4
        .CommandType = adCmdText
        .CommandText = strconsulta
        Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, i_organizacion)
        .Parameters.Append parametro
        Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, i_almacen)
        .Parameters.Append parametro
        Set parametro = .CreateParameter(, adNumeric, adParamInput, 50, i_movimiento)
        .Parameters.Append parametro
        Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, i_numero)
        .Parameters.Append parametro
        Set parametro = .CreateParameter(, adNumeric, adParamInput, 50, i_codigo_barras)
        .Parameters.Append parametro
   End With
   Set rsaux16 = comandoORA.execute
   Set comandoORA = Nothing
   Set parametro = Nothing
   If rsaux16.EOF Then
      strconsulta = "insert into xxvia_tb_transacciones where (organizacion, almacen, movimiento, numero, codigo_barras, codigo, cantidad) values (?,?,?,?,?,?,?) "
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, i_organizacion)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, i_almacen)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 50, i_movimiento)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, i_numero)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 50, i_codigo_barras)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 50, i_codigo)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, i_cantidad)
           .Parameters.Append parametro
      End With
      Set rsaux17 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
   Else
      strconsulta = "update xxvia_tb_transacciones set cantidad = cantidad + ? where organizacion = ? and almacen = ? and movimiento = ? and numero = ? and codigo_barras = ?"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, i_cantidad)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, i_organizacion)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, i_almacen)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 50, i_movimiento)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, i_numero)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 50, i_codigo_barras)
           .Parameters.Append parametro
      End With
      Set rsaux17 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
   End If
   rsaux16.Close
End Sub


Public Sub ejecuta_forma()
   Dim var_tipo_ejecucion As String
   Dim appl As New CRAXDRT.Application
   Dim reporte As New CRAXDRT.Report
   Dim ntablas As Double
   Dim archivo As String
   Dim var_cadena As String
   Dim var_si As Integer
   Dim x As Integer
   If Trim(var_accion_submenu) > 0 Then
      Frmmenu2.Enabled = False
   End If
   var_activa_menu = True
   var_numero_embarque = 0
   'formas
   Select Case var_accion_submenu
       Case "01"
          var_opcion_seguridad = 2
          var_nec_emb = False
          var_activa_forma_codigo_acceso = "MENU"
          var_tipo_embarque = 1
          frmcodigo_acceso.Show
       Case "02"
          Frmmenu2.Enabled = False
          var_opcion_seguridad = 2
          var_activa_forma_usuarios = "MENU"
          frmusuarios.Show
       Case "03"
          Frmmenu2.Enabled = False
          var_opcion_seguridad = 2
          var_activa_forma_menus = "MENU"
          frmmenus.Show
       Case "04"
          var_opcion_seguridad = 2
          var_activa_forma_puestos = "MENU"
          frmpuestos.Show
       Case "05"
          var_opcion_seguridad = 2
          var_activa_forma_unidadesorganizacionales = "MENU"
          frmunidadesorganizacionales.Show
       Case "06"
          var_opcion_seguridad = 2
          var_activa_forma_ejecuta_sistema = "MENU"
          frmejectuta_sistema.Show
        Case "07"
           var_opcion_seguridad = 4
           var_activa_forma_personas = "MENU"
           frmpersonal.Show
       Case "08"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmcreacion_equipos.Show
       Case "09"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmcomportamiento_equipos_embarque_2.Show
       Case "10"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_imprime_ordenes_surtido.Show
       Case "11"
          var_opcion_seguridad = 2
          var_tipo_embarque = 1
          var_activa_forma_existencias_generales = "MENU"
          var_salida_cajas = 0
          frmnumero_embarque.Show
       Case "12"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          If var_bandera_asignacion = 0 Then
             var_tipo_embarque = 1
          Else
             var_tipo_embarque = 2
          End If
          var_salida_cajas = 0
          frmnumero_embarque.Show
       Case "13"
          'MsgBox var_unidad_organizacional
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          var_tipo_embarque = 2
          frmoracle_listado_movimientos.Show
       Case "14"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          var_tipo_embarque = 2
          frmoracle_packing_list.Show
       Case "15"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          var_tipo_embarque = 2
          frmoracle_libera_pedidos_vxt.Show
       Case "16"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          var_tipo_embarque = 2
          frmoracle_facturas.Show
       Case "17"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          var_tipo_embarque = 2
          frmoracle_orden_a_depurar.Show
       Case "18"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          var_tipo_embarque = 2
          frmconcentrado_orden_surtido.Show
       Case "19"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          var_tipo_embarque = 2
          frmcreacion_equipos.Show
       Case "20"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          var_tipo_embarque = 2
          frmcomportamiento_equipos_embarque_2.Show
       Case "21"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          var_tipo_embarque = 2
          frmreporte_ordenes_surtido_pendientes.Show
       Case "22"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          var_tipo_embarque = 2
          frmreporte_entradas_produccion.Show
       Case "23"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          var_tipo_embarque = 2
          frmreporte_logistica_negado.Show
       Case "24"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          var_tipo_embarque = 2
          frmreporte_concecutivo_tipo_documento.Show
       Case "25"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          var_tipo_embarque = 2
          frmreporte_antiguedad_saldos.Show
       Case "26"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          var_tipo_embarque = 2
          frmoracle_cerrar_embarque.Show
       Case "27"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          var_tipo_embarque = 2
        
          frmoracle_concentrado_ordenes_surtido.Show
        Case "28"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          var_tipo_embarque = 2
          frmoracle_comparacion_pedido_afectado.Show
        Case "29"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          var_tipo_embarque = 2
          frmoracle_existencias_rapidas.Show
        Case "30"
          Frmmenu2.Enabled = True
          'var_opcion_seguridad = 2
          'var_activa_forma_existencias_generales = "MENU"
          'var_tipo_embarque = 2
          'frmimpresion_etiquetas_textilera.Show
          If var_unidad_organizacional = "93" Then
             x = Shell(App.Path + "/MovimietosInventarios.exe " + var_unidad_organizacional + "|CDI_ALMPT|8|#")
          End If
          If var_unidad_organizacional = "85" Then
             x = Shell(App.Path + "/MovimietosInventarios.exe " + var_unidad_organizacional + "|PRODTER|8|#")
          End If
          If var_unidad_organizacional = "90" Then
             x = Shell(App.Path + "/MovimietosInventarios.exe " + var_unidad_organizacional + "|CDISTEX_PT|8|#")
          End If
          
        Case "31"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          var_tipo_embarque = 2
          frmoracle_reporte_pedidos_cargados.Show
        Case "32"
          Frmmenu2.Enabled = True
          'var_opcion_seguridad = 2
          'var_activa_forma_existencias_generales = "MENU"
          'var_tipo_embarque = 2
          'frmimpresion_etiquetas_textilera.Show
          x = Shell(App.Path + "/MovimietosInventarios.exe " + var_unidad_organizacional + "|CDI_ALMPT|1|#")
        Case "33"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          var_tipo_embarque = 2
          frmoracle_ubicaciones.Show
        Case "34"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          var_tipo_embarque = 2
          frmoracle_grafica_surtido_OS.Show
        Case "35"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          var_tipo_embarque = 2
          frmoracle_reporte_ubicaciones.Show
        Case "36"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          var_tipo_embarque = 2
          frmoracle_valuacion_devoluciones.Show
        Case "37"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          var_tipo_embarque = 2
          frmoracle_cambiar_maquina_embarque.Show
        Case "38"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          var_tipo_embarque = 2
          frmoracle_grafica_surtido_avance.Show
        Case "39"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          var_tipo_embarque = 2
          frmoracle_etiquetas_kanbans.Show
        Case "40"
          var_opcion_seguridad = 2
          var_activa_forma_lineas = "MENU"
          var_tipo_embarque = 2
          frmoracle_tipos_cajas.Show
        Case "41"
          var_opcion_seguridad = 2
          var_activa_forma_lineas = "MENU"
          var_tipo_embarque = 2
          frmoracle_correcion_entradas_compra.Show
        Case "42"
          rs.Open "SELECT * FROM TB_USUARIOS WHERE VCHA_USU_USUARIO_ID = '" + var_clave_usuario_global + "' AND INTE_PERMISO_HR_X_HR = 1", cnn, adOpenDynamic, adLockOptimistic
          If Not rs.EOF Then
             rs.Close
             var_opcion_seguridad = 2
             var_activa_forma_lineas = "MENU"
             var_tipo_embarque = 2
             frmoracle_comportamiento_hora_hora.Show
          Else
             rs.Close
             var_opcion_seguridad = 2
             var_activa_forma_lineas = "MENU"
             var_tipo_embarque = 2
             MsgBox "No tiene permisos para la opción seleccionada."
             Call activa_forma(var_activa_forma_lineas)
          End If
        Case "43"
          var_opcion_seguridad = 2
          var_activa_forma_lineas = "MENU"
          var_tipo_embarque = 2
          frmoracle_catalogo_transportes.Show
        Case "44"
          var_opcion_seguridad = 2
          var_activa_forma_lineas = "MENU"
          var_tipo_embarque = 2
          frmoracle_impresion_documentos_fiscales.Show
        Case "45"
          var_opcion_seguridad = 2
          var_activa_forma_lineas = "MENU"
          var_tipo_embarque = 2
          frmoracle_bitacora_lectura.Show
        Case "46"
          var_opcion_seguridad = 2
          var_activa_forma_lineas = "MENU"
          var_tipo_embarque = 2
          frmoracle_reporte_control_relacion_mayoreo.Show
        Case "47"
          var_opcion_seguridad = 2
          var_activa_forma_lineas = "MENU"
          var_tipo_embarque = 2
          froracle_asignacion_embarques.Show 1
        Case "48"
          var_opcion_seguridad = 2
          var_activa_forma_lineas = "MENU"
          var_tipo_embarque = 2
          frmoracle_jaulas.Show
        Case "49"
          var_opcion_seguridad = 2
          var_tipo_embarque = 2
          var_activa_forma_existencias_generales = "MENU"
          var_salida_cajas = 1
          frmnumero_embarque.Show
        Case "50"
          var_opcion_seguridad = 2
          var_tipo_embarque = 1
          var_activa_forma_existencias_generales = "MENU"
          var_salida_cajas = 1
          frmoracle_reporte_sellos.Show
        Case "51"
          frmexistencias_rapidas.Show 1
        Case "52"
          var_opcion_seguridad = 2
          var_tipo_embarque = 1
          var_activa_forma_existencias_generales = "MENU"
          var_salida_cajas = 1
          frmoracle_asignar_prioridad_rutas.Show
        Case "53"
          var_opcion_seguridad = 2
          var_tipo_embarque = 1
          var_activa_forma_existencias_generales = "MENU"
          var_salida_cajas = 1
          frmpruebas.Show
        Case "54"
          var_opcion_seguridad = 2
          var_tipo_embarque = 1
          var_activa_forma_existencias_generales = "MENU"
          var_salida_cajas = 1
          frmoracle_reporte_bultos.Show
        Case "55"
          var_opcion_seguridad = 2
          var_tipo_embarque = 1
          var_activa_forma_existencias_generales = "MENU"
          var_salida_cajas = 1
          frmoracle_maquinas_aduana.Show
        Case "56"
          var_opcion_seguridad = 2
          var_tipo_embarque = 1
          var_activa_forma_existencias_generales = "MENU"
          var_salida_cajas = 1
          frmoracle_video.Show 1
        Case "57"
          var_opcion_seguridad = 2
          var_tipo_embarque = 1
          var_activa_forma_existencias_generales = "MENU"
          var_salida_cajas = 1
          frmoracle_reporte_relacion_documentos.Show
        Case "58"
          var_opcion_seguridad = 2
          var_tipo_embarque = 1
          var_activa_forma_existencias_generales = "MENU"
          var_salida_cajas = 1
          frmoracle_usuarios_permiso_cerrar_pedidos.Show
        Case "59"
          var_opcion_seguridad = 2
          var_tipo_embarque = 1
          var_activa_forma_existencias_generales = "MENU"
          var_salida_cajas = 1
          frmoracle_reporte_rendicion_cuentas.Show
        Case "60"
          var_opcion_seguridad = 2
          var_tipo_embarque = 1
          var_activa_forma_existencias_generales = "MENU"
          var_salida_cajas = 1
          frmoracle_desbloquear_lotes.Show
        Case "61"
          var_opcion_seguridad = 2
          var_tipo_embarque = 1
          var_activa_forma_existencias_generales = "MENU"
          var_salida_cajas = 1
          frmoracle_negado_distribucion.Show
        Case "62"
          var_opcion_seguridad = 2
          var_tipo_embarque = 1
          var_activa_forma_existencias_generales = "MENU"
          var_salida_cajas = 1
          frmoracle_complementos_articulos_packing_list.Show
        Case "63"
          var_opcion_seguridad = 2
          var_tipo_embarque = 1
          var_activa_forma_existencias_generales = "MENU"
          var_salida_cajas = 1
          frmoracle_subir_pedidos_contingencia.Show
        Case "64"
          MsgBox "Forma de pantalla de pedidos activos en oracle web", vbOKOnly, "ATENCION"
          Frmmenu2.Enabled = True
        Case "66"
          var_opcion_seguridad = 2
          var_tipo_embarque = 1
          var_activa_forma_existencias_generales = "MENU"
          var_salida_cajas = 1
          frmoracle_eliminar_documento_eflow.Show
        Case "67"
          var_opcion_seguridad = 2
          var_tipo_embarque = 1
          var_activa_forma_existencias_generales = "MENU"
          var_salida_cajas = 1
          frmoracle_desbloquear_cajas.Show
        Case "68"
          var_opcion_seguridad = 2
          var_tipo_embarque = 1
          var_activa_forma_existencias_generales = "MENU"
          var_salida_cajas = 1
          frmoracle_regenerar_pedido.Show
        Case "69"
          var_opcion_seguridad = 2
          var_tipo_embarque = 1
          var_activa_forma_existencias_generales = "MENU"
          var_salida_cajas = 1
          frmoracle_facturar_embarques_eflow.Show
        Case "70"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_reporte_existencias_costales.Show
        Case "71"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_reporte_reservas.Show
        Case "72"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_cargar_pedido_usa.Show
        Case "73"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_entrada_bultos.Show
        Case "74"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_validador_cubicaje.Show 1
        Case "75"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_embarques_cerrados_pedidos_abiertos.Show
        Case "76"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_reporte_validacion_doc_fiscales_vs_eflow.Show
        Case "77"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_multiplo_articulos.Show
        Case "78"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_imprimir_os_historica.Show
        Case "79"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_lead_time.Show
        Case "80"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_crear_pedidos_costales.Show
        Case "81"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_reporte_comportamiento_semanal.Show
        Case "82"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_carta_juramentada.Show
        Case "83"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_eliminar_bultos.Show
        Case "84"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_negado_distribucion_hora_x_hora.Show
        Case "85"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_peso_volumen_embarques.Show
        Case "86"
          var_opcion_seguridad = 2
          If var_clave_usuario_global = "U0000000528" Or var_clave_usuario_global = "U0000000529" Or var_clave_usuario_global = "U0000001098" Or var_clave_usuario_global = "U0000001109" Or var_clave_usuario_global = "U0000001027" Or var_clave_usuario_global = "U0000001027" Or var_clave_usuario_global = "U0000000932" Then
             var_activa_forma_existencias_generales = "MENU"
             frmoracle_rutas_distribucion.Show
          Else
             MsgBox "no tiene acceso a esta opción", vbOKOnly, "ATENCION"
             var_activa_forma_existencias_generales = "MENU"
             Call activa_forma(var_activa_forma_existencias_generales)
          End If
        Case "87"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_cargar_pedido_guatemala.Show
        Case "88"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_reporte_ruta_lechera.Show
        Case "89"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_reporte_volumen_importe.Show
        Case "90"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_ubicaciones_motor_logistico.Show
        Case "91"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_validador_codigos_barras.Show
        Case "92"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_semaforo_pedidos.Show 1
        Case "93"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_control_bultos.Show
        Case "94"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_tax_id.Show
        Case "95"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_factura_complementaria_exportaciones.Show
        Case "96"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_existencias_bultos_titulares.Show
        Case "97"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmprueba_puerto.Show 1
        Case "98"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_impresion_recepciones.Show
        Case "99"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmchoferes.Show
        Case "100"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_trajinantes.Show
        Case "101"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_reporte_bultos_por_embarque.Show
        Case "102"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_establecimientos_sin_dias.Show
        Case "103"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_activar_rutas_distribucion.Show 1
        Case "104"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_eliminar_pedidos.Show
        Case "105"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_rendimiento.Show
        Case "106"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_embarque_concentrado.Show
        Case "107"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_embarque_a_surtir.Show
        Case "108"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_embarques_a_surtir_lista.Show
        Case "109"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_reporte_entrega_planta_cedis.Show
        Case "110"
          Dim i, iFila, ifila2, iCol, icol2 As Integer
          Dim oexcel As Excel.Application
          Dim owbook As Excel.Workbook
          Dim osheet As Excel.Worksheet
          Set oexcel = CreateObject("Excel.Application")
          Set owbook = oexcel.Workbooks.Add
          Set osheet = owbook.Worksheets(1)
          osheet.Name = "PESO Y VOLUMEN"
          Screen.MousePointer = vbHourglass
          iFila = 1
          ifila2 = 1
          icol2 = 1
          iCol = 1
          'var_cadena = "select distinct b.segment1 as CODIGO, description AS DESCRIPCION, unit_weight PESO, unit_volume VOLUMEN from xxvia_system_items_b a,  XXVIA_tB_sALIDAS_cAJAS b WHERE SOURCE_HEADER_NUMBER >= 240292 AND SUBINVENTORY = 'CDI_ALMPT' and a.organization_id = 93 and a.segment1 = b.segment1"
          var_cadena = "select distinct a.segment1 as CODIGO, description AS DESCRIPCION, unit_weight PESO, unit_volume VOLUMEN from xxvia_system_items_b a WHERE a.organization_id = 93 "
          rsaux10.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
          For i = 0 To rsaux10.Fields.Count - 1
              osheet.Cells(iFila, i + 1) = rsaux10.Fields(i).Name
          Next
          iFila = iFila + 1
          With osheet
               ' carga los registros del recordset
               .Cells(iFila, iCol).CopyFromRecordset rsaux10
               'oExcel.Columns(1).Select
               'oExcel.Selection.NumberFormat = "#,##0.00"
               'oExcel.Columns(1).Select
               'oExcel.Selection.Font.Color = vbRed
               .Columns.AutoFit ' ajusta el ancho de las columnas
          End With
          archivo = "c:\reportessid\PESO_VOLUMEN_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
          owbook.SaveAs archivo
          oexcel.Visible = True
          Set oexcel = Nothing
          Screen.MousePointer = vbDefault
          rsaux10.Close
          MsgBox "Se a terminado de guardar el archivo " + archivo
          Frmmenu2.Enabled = True
        Case "111"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_salidas_privalia.Show
        Case "112"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_lead_time_embarques.Show
        Case "113"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmetiquetas_ubicaciones_contingencia.Show 1
          Call activa_forma(var_activa_forma_existencias_generales)
        Case "114"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_semaforo_bultos.Show 1
          Call activa_forma(var_activa_forma_existencias_generales)
        Case "115"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_factura_pedido.Show 1
          Call activa_forma(var_activa_forma_existencias_generales)
        Case "116"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_creacion_palets.Show 1
          Call activa_forma(var_activa_forma_existencias_generales)
        Case "117"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_CN_frontera.Show 1
          Call activa_forma(var_activa_forma_existencias_generales)
        Case "118"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmnegado_distribucion_hora_x_hora.Show 1
          Call activa_forma(var_activa_forma_existencias_generales)
        Case "119"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_reporte_bultos_periodo.Show
        Case "120"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_datos_embarque_exportacion.Show
        Case "121"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_estatus_embarques_exportaciones.Show
        Case "122"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_volumen_embarque.Show
        Case "123"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_clientes_costales.Show
        Case "124"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_agrupacion_bultos.Show
        Case "125"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmcambiar_transporte_embarque_exportaciones.Show 1
        Case "126"
           
           x = Shell(App.Path + "/sid_2.exe")
            Frmmenu2.Enabled = True
        Case "398"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_dividir_pedido.Show 1
        Case "399"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_eliminar_guia.Show
        Case "400"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_bitacora_aduana_cajas.Show
        Case "401"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_carta_porte.Show
        Case "402"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_asignar_chofer_unidad.Show 1
        Case "403"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmcarta_porte_CDMX.Show 1
        Case "404"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmcarta_porte_MTY.Show 1
        Case "405"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_configurar_direcciones_anes_cn.Show 1
        Case "406"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_carta_porte_pedido.Show 1
        Case "407"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmcarta_porte_QRO.Show 1
        Case "408"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_reporte_carta_porte_paqueterias.Show 1
        Case "409"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmcarta_porte_traspasos.Show 1
        Case "410"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_carta_porte_devoluciones_CN.Show 1
          
        Case "411"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_carta_porte_devolucion_ANES.Show 1
        Case "412"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_reporte_devoluciones_anes.Show 1
        Case "413"
          var_opcion_seguridad = 2
          var_activa_forma_existencias_generales = "MENU"
          frmoracle_reporte_pedidos_enviados_CN.Show 1
          
         

         End Select
End Sub

Public Sub copiar_sonidos()
   On Error GoTo SALIR:
   FileCopy "\\vianney.com.mx\srvarchivos\updatesid\Cerrar el lote.mp3", App.Path + "\Cerrar el lote.mp3"
   Exit Sub
SALIR:
   'MsgBox "Existe una nueva actualizacion en el servidor, salga de todas las instancias del sistema para poder actualizar.", vbOKOnly, "ATENCION"
End Sub



Public Sub copiar_facturar_exe()
   On Error GoTo SALIR:
   FileCopy "\\vianney.com.mx\srvarchivos\sidsinoracle\facturar.exe", "c:\sistemas\facturar.exe"
   Exit Sub
SALIR:
   'MsgBox "Existe una nueva actualizacion en el servidor, salga de todas las instancias del sistema para poder actualizar.", vbOKOnly, "ATENCION"
End Sub

Public Sub suma_lotes(var_pedido As Double, var_lote As Double, var_cantidad As Double, var_signo As String)
   rsaux10.Open "select * from tb_oracle_suma_lotes where pedido = " + CStr(var_pedido) + " and lote = " + CStr(var_lote), cnn, adOpenDynamic, adLockOptimistic
   If Not rsaux10.EOF Then
      If var_signo = "+" Then
         rsaux9.Open "update tb_oracle_suma_lotes set cantidad = cantidad + " + CStr(var_cantidad) + " where pedido = " + CStr(var_pedido) + " and lote = " + CStr(var_lote), cnn, adOpenDynamic, adLockOptimistic
      End If
      If var_signo = "-" Then
         rsaux9.Open "update tb_oracle_suma_lotes set cantidad = cantidad - " + CStr(var_cantidad) + " where pedido = " + CStr(var_pedido) + " and lote = " + CStr(var_lote), cnn, adOpenDynamic, adLockOptimistic
      End If
   Else
      rsaux9.Open "insert into tb_oracle_suma_lotes (pedido, lote, cantidad) values (" + CStr(var_pedido) + "," + CStr(var_lote) + ", " + CStr(var_cantidad) + ")", cnn, adOpenDynamic, adLockOptimistic
   End If
   rsaux10.Close

End Sub



Public Sub cantidad_leida_por_persona(var_cantidad As Double, var_signo As String)
Dim var_fecha_s As String
Dim var_fecha As Date
Dim var_hora As Integer
Dim var_dia_str, var_mes_str, var_año_str, var_hora_str, var_cadena As String
   'If var_cantidad = 1 Then
       var_fecha = Now
       var_dia_str = CStr(Day(var_fecha))
       If Len(var_dia_str) = 1 Then
          var_dia_str = "0" + var_dia_str
       End If
       var_mes_str = CStr(Month(var_fecha))
       If Len(var_mes_str) = 1 Then
          var_mes_str = "0" + var_mes_str
       End If
       var_año_str = CStr(Year(var_fecha))
       If Len(var_año_str) = 2 Then
          var_año_str = "20" + var_año_str
       End If
       var_hora = Hour(var_fecha)
       var_fecha_s = var_dia_str + "/" + var_mes_str + "/" + var_año_str

       rsaux10.Open "select * from tb_oracle_lectura_usuarios where fecha ='" + var_fecha_s + "' and usuario = '" + var_clave_usuario_global + "'", cnn, adOpenDynamic, adLockOptimistic
       If rsaux10.EOF Then
          rsaux11.Open "insert into tb_oracle_lectura_usuarios (fecha, usuario) values ('" + var_fecha_s + "','" + var_clave_usuario_global + "')", cnn, adOpenDynamic, adLockOptimistic
       End If
       rsaux10.Close
       Select Case var_hora
              Case 0
                   var_hora_str = "h_0_1"
              Case 1
                   var_hora_str = "h_1_2"
              Case 2
                   var_hora_str = "h_2_3"
              Case 3
                   var_hora_str = "h_3_4"
              Case 4
                   var_hora_str = "h_4_5"
              Case 5
                   var_hora_str = "h_5_6"
              Case 6
                   var_hora_str = "h_6_7"
              Case 7
                   var_hora_str = "h_7_8"
              Case 8
                   var_hora_str = "h_8_9"
              Case 9
                   var_hora_str = "h_9_10"
              Case 10
                   var_hora_str = "h_10_11"
              Case 11
                   var_hora_str = "h_11_12"
              Case 12
                   var_hora_str = "h_12_13"
              Case 13
                   var_hora_str = "h_13_14"
              Case 14
                   var_hora_str = "h_14_15"
              Case 15
                   var_hora_str = "h_15_16"
              Case 16
                   var_hora_str = "h_16_17"
              Case 17
                   var_hora_str = "h_17_18"
              Case 18
                   var_hora_str = "h_18_19"
              Case 19
                   var_hora_str = "h_19_20"
              Case 20
                   var_hora_str = "h_20_21"
              Case 21
                   var_hora_str = "h_21_22"
              Case 22
                   var_hora_str = "h_22_23"
              Case 23
                   var_hora_str = "h_23_24"
       End Select
       If var_signo = "+" Then
          var_cadena = "update tb_oracle_lectura_usuarios set " + var_hora_str + "=  isnull(" + var_hora_str + ",0) + " + CStr(var_cantidad) + " where fecha = '" + var_fecha_s + "' and usuario = '" + var_clave_usuario_global + "'"
       End If
       If var_signo = "-" Then
          var_cadena = "update tb_oracle_lectura_usuarios set " + var_hora_str + "=  isnull(" + var_hora_str + ",0) - " + CStr(var_cantidad) + " where fecha = '" + var_fecha_s + "' and usuario = '" + var_clave_usuario_global + "'"
       End If
       rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
    'End If
End Sub



Public Sub activa_forma(var_activa_forma As String)
   Select Case var_activa_forma
       Case "MENU"
            Frmmenu2.Enabled = True
       Case "frmsalidas"
       Case "frmsalidas_cajas"
       Case "frmclientes"
       Case "frmarticulos2"
       Case "frmlistatitulares"
       Case "frmcodigo_acceso"
            frmcodigo_acceso.Enabled = True
       Case "frmcajas"
       Case "frmcanalesventas"
       Case "frmcatalogos"
       Case "frmcatempresas"
       Case "frmciudades"
       Case "frmclasearticulos"
       Case "frmmonedas"
       Case "frmmovimientos"
       Case "frmordenescompra"
       Case "frmpaises"
       Case "frmplazos"
       Case "frmproductos2"
       Case "16"
       Case "frmrutas"
       Case "frmsublineas"
       Case "frmsubtipoarticulos"
       Case "frmsubtiposusos"
       Case "frmtallas"
       Case "frmtipoagentes"
       Case "frmtipoarticulos"
       Case "frmtipocambio"
       Case "frmtipoestampados"
       Case "frmtiposclientes"
       Case "frmtitulares"
       Case "frmtonos"
       Case "frmtransportes"
       Case "frmunidades"
       Case "frmusos"
       Case "frmusuarios"
            frmusuarios.Enabled = True
       Case "frmvehiculos"
       Case "frmagentes"
       Case "frmlicencias"
       Case "frmdiseños"
       Case "frmlineas"
       Case "frmmenus"
       Case "frmpuestos"
       Case "frmzonas"
       Case "frmcargapedidos"
       Case "frmtipopedidos"
       Case "frmlistatitulares"
       Case "frmgruposactuales"
       Case "frmgruposreales"
       Case "frmclientes2"
       Case "frmgenerapedido"
       Case "frmgenera_pedidos_multibondeados"
       Case "frmtipopedidos"
       Case "frmautorizapedidos"
       Case "frmordensurtido"
       Case "frmkardex"
       Case "frmnumero_embarque"
       Case "frmfactura_embarques"
       Case "frmcodigo_acceso"
            frmcodigo_acceso.Enabled = True
       Case "frmequipos"
            frmequipos.Enabled = True
       Case "frmcancela_cajas"
       Case "frmcausas_devolucion"
       Case "frmasigna_causa_devolucion"
       Case "frmasignacion_negado"
       Case "frmembarques_activos"
       Case "frmnotas_credito"
       Case "frmestampados"
       Case "frmcolores"
       Case "frmunidadesorganizacionales"
            frmunidadesorganizacionales.Enabled = True
       Case "frmalmacenes"
            frmalmacenes.Enabled = True
       Case "frmproveedores"
       Case "frmlistadeprecios"
       Case "frmdescuentos_catalogos"
       Case "frmdescuentos_promociones"
       Case "frmdescuentos_volumen"
       Case "frmestados_cuenta"
       Case "frmreporte_movimientos"
       Case "frmrelacion_cobranza_listado"
       Case "frmasigna_pagos_no_aplicados"
       Case "frmnotas_cargo"
       Case "frmfacturas"
       Case "frmcheques_devueltos"
       Case "frmbonificaciones"
       Case "frmbonificaciones_financieras"
       Case "frmcancela_facturas"
       Case "frmcheques_devueltos_inicio"
       Case "frmclases_cartera"
       Case "frmdescuentos_volumen_grupo_actual"
       Case "frmdescuentos_volumen_grupo_real"
       Case "frmdescuentos_volumen_titular"
       Case "frmdescuentos_volumen_cliente"
       Case "frmclasificacion_clientes"
       Case "frmcausas_no_otorgamiento"
       Case "frmdescuentos_pronto_pago_cambios"
       Case "frmfamilia_agrupadores"
       Case "frmpromociones_inicio_catalogo"
       Case "frmrangos_promociones_catalogos"
       Case "frmcancela_documentos_existentes"
       Case "frmcancela_facturas_devolucion"
       Case "frmejectuta_sistema"
            frmejectuta_sistema.Enabled = True
       Case "frmequivalencias"
       Case "frminformacion_pedido_sugerido"
       Case "frminformacion_articulos_enviar"
       Case "frmcerrado_pedidos"
       Case "frmconcentrado_orden_surtido"
       Case "frmreporte_nivel_surtido"
       Case "frminventario_documentos"
       Case "frmreporte_comisiones"
       Case "frmpacking_list"
       Case "frmreporte_valuacion_facturas"
       Case "frmreporte_valuacion_devoluciones"
       Case "frmreporte_envios_tiendas"
       Case "frmexistencias_generales"
       Case "frmcostos_predeterminados"
       Case "frmreporte_acumulado_ventas"
       Case "frmreporte_ajustes_reempaque"
       Case "frmreporte_entradas_produccion"
       Case "frmreporte_mercancia_empacada_transito"
       Case "frmreporte_ordenes_surtido_pendientes"
       Case "frmreporte_entradas_salidas"
       Case "frmreporte_catalogo_articulos"
       Case "frmdetalle_documentos_fiscales"
       Case "frmreporte_concentrado_entradas_salidas"
       Case "frmestados"
       Case "frmmunicipios"
       Case "frmcolonias"
       Case "frmmigracion"
       Case "frmestablecimientos"
       Case "frmreporte_totalizador"
       Case "frmreclasificacion_almacen"
       Case "frmreporte_ventas_netas_tipo_reporte"
       Case "froracle_asignacion_embarques"
            froracle_asignacion_embarques.Enabled = True
      End Select
End Sub




Public Sub ejecuta_cambios()
   var_acepta_seguridad = 1
End Sub


 

'************ procedimiento para eviar correo **********************

 

 

Public Sub pro_envio_correo_app(ByVal var_para As String, ByVal var_sujeto As String, ByVal var_mensaje As String, ByVal var_adjunto As String)
End Sub

 


'======== Procedimiento Generar el Menu de Acuerdo al Modulo   ==========================


Public Sub pro_forma_menu(ByVal forma As Form, ByVal tabla As String, ByVal indice As Integer)
   Dim tablas As String
   Select Case indice + 1
   Case 1
      Call pro_menus_consulta(forma, tabla)
   Case 2
      Call pro_menus_consulta(forma, tabla)
   Case 3
      Call pro_menus_consulta(forma, tabla)
   End Select
End Sub

Public Sub pro_menus_consulta(ByVal forma As Form, ByVal tabla As String)
Dim i As Integer, x As Byte
Dim varmenu As String, varsubmenu As String, varcaption As String
Dim var_nivel As String
Dim var_nivelaux2 As String, var_nivelaux3 As String, var_nivelaux4 As String
Dim var_resultado As Integer

With forma

Set .XPsidemenu1.Pimagelist = .ImageList2
Set .XPsidemenu1.hImageList = .ImageList1
        i = 1: x = 0
        rs.Open "select * from " & tabla & " where left(vcha_men_nivel,2) =" & Str(indice + 1) & "order by vcha_men_nivel", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
            rsaux.Open "select TB_PUESTOS.VCHA_PUE_MENUS from TB_USUARIOS,TB_PUESTOS WHERE VCHA_USU_PUESTO = VCHA_PUE_DESCRIPCION AND VCHA_USU_USUARIO = '" & var_usuario_global & "'", cnn, adOpenDynamic, adLockOptimistic
            
            If rsaux.RecordCount <> 0 And rsaux(0) <> "" Then
                var_resultado = InStr(1, rsaux(0).Value, rs(0).Value, vbTextCompare)
                If var_resultado <> 0 Then
                    If Trim(Mid(rs(1).Value, 7, 2)) <> "00" Then GoTo otro:
                        'If Trim(Mid(rs(1).Value, 3, 2)) = "01" Then GoTo otro:
                            varmenu = Str(Trim(Mid(rs(1).Value, 3, 2))): varcaption = Trim(rs(2).Value)
                            varmenu = Trim(varmenu)
                            If Trim(Mid(rs(1).Value, 5, 2)) = "00" Then
 '                               .XPsidemenu1.Addpanel "P" & varmenu, varcaption, Closed, False, Mid(rs(1).Value, 3, 2)
                                x = x + 1
                            End If
                            If Trim(Mid(rs(1).Value, 5, 2)) <> "00" Then
                                var_resultado = False
                                var_resultado = InStr(1, rsaux(0).Value, rs(0).Value + "*", vbTextCompare)
                                'If var_resultado <> 0 Then
                                '    vector_valida_passwords(i) = "*"
                                'End If
                                varsubmenu = Str(Trim(Mid(rs(1).Value, 3, 2))): varcaption = Trim(rs(2).Value)
                                varsubmenu = Trim(varsubmenu)
                                '.XPsidemenu1.AddHyper "H" & Str(i), "P" & varsubmenu, varcaption, True, Hyperlink, 1, "This is tooltip 1"
                        End If
                    End If
                End If
            'End If
otro:
            rsaux.Close
            rs.MoveNext
            i = i + 1
        Wend
        
        For i = 1 To x
        .XPsidemenu1.opeclo (i)
        Next i
        rs.Close
End With
End Sub

'========== Procedimiento para consultar por el campo principal de cada tabla ===========




Public Sub pro_encabezadosView(ByVal forma As Form, ByVal ctl As ListView, ByVal flat As Boolean)
Dim imlItem As ListImage

    With forma
        If flat Then
            'encabezados flat inabilita el click colum
            Dim hHeader As Long
            hHeader = SendMessage(ctl.hwnd, LVM_GETHEADER, 0, ByVal 0&)
            SetWindowLong hHeader, GWL_STYLE, GetWindowLong(hHeader, GWL_STYLE) Xor HDS_BUTTONS
        End If
        SendMessage ctl.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, ByVal (SendMessage(ctl.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, ByVal 0&) Xor LVS_EX_FULLROWSELECT)
        '.icono_encabezado.ImageHeight = 12
        '.icono_encabezado.ImageWidth = 12
        'ctl.ColumnHeaderIcons = .icono_encabezado
        ctl.Arrange = lvwAutoTop
        ctl.LabelEdit = lvwManual
    End With
End Sub




Public Sub pro_ordena_listas(ByVal lv As ListView, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim colvar As ColumnHeader
    If lv.Sorted = True And ColumnHeader.SubItemIndex = lv.SortKey Then
        If lv.SortOrder = lvwAscending Then
            lv.SortOrder = lvwDescending
        Else
            lv.SortOrder = lvwAscending
        End If
    Else
        lv.Sorted = True
        lv.SortKey = ColumnHeader.SubItemIndex
        lv.SortOrder = lvwAscending
    End If
   ' For Each colvar In lv.ColumnHeaders
   '     If colvar.SubItemIndex = lv.SortKey Then
   '         If lv.SortOrder = lvwDescending Then
   '             colvar.Icon = 1
   '         Else
   '             colvar.Icon = 2
   '         End If
   '     Else
   '         colvar.Icon = 0
   '     End If
    'Next colvar
End Sub

Public Sub pro_cambiaseleccion(ByVal forma As Form, ByVal ctl As ListView, ByVal Item As MSComctlLib.ListItem)
Dim x As Byte

    With forma
        Set ctl.selectedItem = Item
        
        If Item.Checked Then
            ' user checked an item
            ctl.ListItems.Item(ctl.selectedItem.Index).SmallIcon = 2
            ctl.ListItems.Item(ctl.selectedItem.Index).Bold = True
            ctl.ListItems.Item(ctl.selectedItem.Index).ForeColor = &H800000
            For x = 1 To ctl.ColumnHeaders.Count - 1
                ctl.ListItems.Item(ctl.selectedItem.Index).ListSubItems(x).Bold = True
                ctl.ListItems.Item(ctl.selectedItem.Index).ListSubItems(x).ForeColor = &H800000
            Next x
        Else
            ctl.ListItems.Item(ctl.selectedItem.Index).SmallIcon = 1
            ctl.ListItems.Item(ctl.selectedItem.Index).ForeColor = vbBlack
            ctl.ListItems.Item(ctl.selectedItem.Index).Bold = False
            For x = 1 To ctl.ColumnHeaders.Count - 1
                ctl.ListItems.Item(ctl.selectedItem.Index).ListSubItems(x).Bold = False
                ctl.ListItems.Item(ctl.selectedItem.Index).ListSubItems(x).ForeColor = vbBlack
            Next x
      
      End If
    End With

End Sub

'========== Procedimiento llenar el grids  ===================================




Public Sub pro_llena_Treeherramientas(ByVal forma As Form)

Dim i As Integer
Dim nivel2 As Node, nivel3 As Node, nivel4 As Node
    With forma
        ' Asigna el TreeView sobre el ImageList.
        For i = 1 To 7
            .imgadmini.ListImages.Add , , .Image1(i - 1).Picture
        Next i
        .Treemenu.ImageList = .imgadmini
        Set nivel2 = .Treemenu.Nodes.Add(, , "f Administracion de Usuarios", "Administracion de Usuarios", 4, 4)
        Set nivel3 = .Treemenu.Nodes.Add(nivel2, tvwChild, "g Usuarios", "Usuarios", 1, 6)
        Set nivel3 = .Treemenu.Nodes.Add(nivel2, tvwChild, "g Plantas", "Plantas", 2, 6)
        Set nivel3 = .Treemenu.Nodes.Add(nivel2, tvwChild, "g Bloques del Sistema", "Bloques del Sistema", 3, 6)
        Set nivel3 = .Treemenu.Nodes.Add(nivel2, tvwChild, "g Puestos", "Puestos", 7, 6)
        Set nivel3 = .Treemenu.Nodes.Add(nivel2, tvwChild, "g Menus del Sistema", "Menus del Sistema", 5, 6)
        
        rs.Open "select * from TB_bloques", cnn, adOpenDynamic, adLockOptimistic
        If rs.RecordCount <> 0 Then
            While Not rs.EOF
                Set nivel4 = .Treemenu.Nodes.Add(nivel3, tvwChild, "g " & Format(rs(0).Value, "00") & " " & rs(1).Value, Format(rs(0).Value, "00") & " " & rs(1).Value, 5, 6)
                rs.MoveNext
            Wend
        End If
        nivel3.EnsureVisible
        rs.Close
    End With
    
End Sub




Public Sub pro_llena_TreeMenus(ByVal forma As Form, ByVal var_num_empresa As String)
Dim i As Integer
Dim nivel2 As Node, nivel3 As Node, nivel4 As Node
Dim x As Integer, Y As Integer, Z As Integer

    With forma
         For i = 1 To 6
             .TreeImages.ListImages.Add , , .TreeImage(i).Picture
             .imgadmini.ListImages.Add , , .Image1(i - 1).Picture
         Next i
         .OrgTree.ImageList = .TreeImages         ' Asigna el TreeView sobre el ImageList.
         x = 1: Y = 1: Z = 1
          
        Select Case var_num_empresa
        Case "01"
              rs.Open "select * from tb_menus where left(VCHA_MEN_NIVEL,2)=" & var_num_empresa & " order by vcha_men_nivel", cnn, adOpenDynamic, adLockOptimistic
              While Not rs.EOF
                  If Trim(Mid(rs(1).Value, 5, 2)) = "00" Then
                      Set nivel2 = .OrgTree.Nodes.Add(, , "f" & rs(2).Value, Format(x, "00") & " " & rs(2).Value, otnivel2, otnivel22)
                      x = x + 1: Y = 1: Z = 1
                  End If
                  If Trim(Mid(rs(1).Value, 5, 2)) <> "00" And Trim(Mid(rs(1).Value, 7, 2)) = "00" Then
                      Set nivel3 = .OrgTree.Nodes.Add(nivel2, tvwChild, "g" & rs(2).Value, Format(Y, "00") & " " & rs(2).Value, otnivel3, otNivel32)
                      Y = Y + 1
                  End If
                  If Trim(Mid(rs(1).Value, 7, 2)) <> "00" Then
                      Set nivel4 = .OrgTree.Nodes.Add(nivel3, tvwChild, "p" & rs(2).Value, Format(Z, "00") & " " & rs(2).Value, otnivel4, otNivel42)
                      Z = Z + 1
                  End If
                  rs.MoveNext
              Wend
              rs.Close
        Case "02"
              rs.Open "select * from tb_menus where left(VCHA_MEN_NIVEL,2)=" & var_num_empresa & " order by vcha_men_nivel", cnn, adOpenDynamic, adLockOptimistic
              While Not rs.EOF
                  If Trim(Mid(rs(1).Value, 5, 2)) = "00" Then
                      Set nivel2 = .OrgTree.Nodes.Add(, , "f" & rs(2).Value, Format(x, "00") & " " & rs(2).Value, otnivel2, otnivel22)
                      x = x + 1: Y = 1: Z = 1
                  End If
                  If Trim(Mid(rs(1).Value, 5, 2)) <> "00" And Trim(Mid(rs(1).Value, 7, 2)) = "00" Then
                      Set nivel3 = .OrgTree.Nodes.Add(nivel2, tvwChild, "g" & rs(2).Value, Format(Y, "00") & " " & rs(2).Value, otnivel3, otNivel32)
                      Y = Y + 1
                  End If
                  If Trim(Mid(rs(1).Value, 7, 2)) <> "00" Then
                      Set nivel4 = .OrgTree.Nodes.Add(nivel3, tvwChild, "p" & rs(2).Value, Format(Z, "00") & " " & rs(2).Value, otnivel4, otNivel42)
                      Z = Z + 1
                  End If
                  rs.MoveNext
              Wend
              rs.Close
        Case "03"
              rs.Open "select * from tb_menus where left(VCHA_MEN_NIVEL,2)=" & var_num_empresa & " order by vcha_men_nivel", cnn, adOpenDynamic, adLockOptimistic
              While Not rs.EOF
                        If Trim(Mid(rs(1).Value, 5, 2)) = "00" Then
                            Set nivel2 = .OrgTree.Nodes.Add(, , "f" & rs(2).Value, Format(x, "00") & " " & rs(2).Value, otnivel2, otnivel22)
                            x = x + 1: Y = 1: Z = 1
                        End If
                        If Trim(Mid(rs(1).Value, 5, 2)) <> "00" And Trim(Mid(rs(1).Value, 7, 2)) = "00" Then
                            Set nivel3 = .OrgTree.Nodes.Add(nivel2, tvwChild, "g" & rs(2).Value, Format(Y, "00") & " " & rs(2).Value, otnivel3, otNivel32)
                            Y = Y + 1
                        End If
                        If Trim(Mid(rs(1).Value, 7, 2)) <> "00" Then
                            Set nivel4 = .OrgTree.Nodes.Add(nivel3, tvwChild, "p" & rs(2).Value, Format(Z, "00") & " " & rs(2).Value, otnivel4, otNivel42)
                            Z = Z + 1
                        End If
                   
                    rs.MoveNext
              Wend
              rs.Close
        Case "04"
        End Select
    
    End With
End Sub




'************************************************************************************

'            RUTINAS O PROCEDIMIENTOS ESPECIALES

'************************************************************************************
 


'========================= Regresa el Objeto Tipo de Nodo. ===================

Public Function NodeType(test_node As Node) As ObjectType
    If test_node Is Nothing Then
        NodeType = otNone
    Else
        Select Case Left$(test_node.Key, 1)
            Case "f"
                NodeType = otnivel2
            Case "g"
                NodeType = otnivel3
            Case "p"
                NodeType = otnivel4
        End Select
    End If
End Function





'=== Limpiar todos los Controles de una forma se le puede mod. p/cualquier control ======

Public Function pro_limpiatextos(ByVal forma As Form)
Dim ctl As Object
    For Each ctl In forma.Controls
        If TypeOf ctl Is TextBox Then ctl = ""
        If TypeOf ctl Is ComboBox Then ctl = ""
    Next ctl
End Function

Public Function pro_limpiatextos2(ByVal forma As Form)
Dim ctl As Object
    For Each ctl In forma.Controls
        If TypeOf ctl Is TextBox And ctl.Name <> "txt_clave" And ctl.Name <> "txt_fecha_transaccion" Then
            ctl = ""
        End If
    Next ctl
End Function


Public Sub pro_combodrop(ByVal Control As ComboBox, ByVal status As Boolean)
    SendMessage Control.hwnd, CB_SHOWDROPDOWN, status, 0
End Sub


Public Sub pro_enfoque(ByVal KeyAscii As Integer)
    If KeyAscii = 13 Then
       Dim WshShell As Object
       Set WshShell = CreateObject("WScript.Shell")
       WshShell.SendKeys "{Tab}"
    End If
End Sub

Public Sub camposbd(ByVal forma As Form, ByVal tabla As String)
Dim i As Integer
Dim vec(10) As Variant

    With forma
    i = 0
    rs.Open "select * from " & tabla, cnn, adOpenDynamic, adLockOptimistic
    While i < 5
        .Combo1.AddItem UCase((rs(i).Name))
        .combo4.AddItem UCase((rs(i).Name))
        .Combo7.AddItem UCase((rs(i).Name))
        vec(i) = rs(i)
        rs.MoveNext
        i = i + 1
    Wend
    End With
    rs.Close
End Sub


'======================= Procedimiento Avanza un Lugar en Listview ===========================

Public Sub pro_avanzar(ByVal forma As Form, ByVal ctl As ListView, ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
Dim x As Long, Y As Long
    With forma
    x = ctl.selectedItem.Index
    If x <> 0 Then
        Select Case Button.Index
        Case 2
            If x = 1 Then x = 0
            If x = 0 Then Exit Sub
            ctl.ListItems.Item(ctl.selectedItem.Index - 1).Selected = True
            ctl.selectedItem.EnsureVisible
            ctl.selectedItem.Selected = True
        Case 3
            If x = ctl.ListItems.Count Then Exit Sub
            ctl.ListItems.Item(ctl.selectedItem.Index + 1).Selected = True
            ctl.selectedItem.EnsureVisible
            ctl.selectedItem.Selected = True
        End Select
    End If
    End With
err0:
End Sub


'=======================     Oculta Todos los Frames de un Forma    ===========================

Sub pro_oculta_frames(forma As Form)
Dim ctl As Object
    For Each ctl In forma.Controls
        If TypeOf ctl Is Frame Then ctl.Visible = False
    Next ctl
End Sub




'=======================  Procedimiento ordenar los flexgrids ===========================

Public Sub SortByColumn(ByVal forma As Form, ByVal sort_column As Integer)
    ' Hide the FlexGrid.
    With forma.cuadro
        .Visible = False
        .Refresh

    ' Sort using the clicked column.
        .col = sort_column
        .ColSel = sort_column
        .row = 0
        .RowSel = 0

    ' If this is a new sort column, sort ascending.
    ' Otherwise switch which sort order we use.
    If m_SortColumn <> sort_column Then
        m_SortOrder = flexSortGenericAscending
    ElseIf m_SortOrder = flexSortGenericAscending Then
        m_SortOrder = flexSortGenericDescending
    Else
        m_SortOrder = flexSortGenericAscending
    End If
    .Sort = m_SortOrder

    ' Restore the previous sort column's name.
    If m_SortColumn >= 0 Then
        .TextMatrix(0, m_SortColumn) = Mid$(.TextMatrix(0, m_SortColumn), 3)
    End If

    ' Display the new sort column's name.
    m_SortColumn = sort_column
    If m_SortOrder = flexSortGenericAscending Then
        .TextMatrix(0, m_SortColumn) = "> " & .TextMatrix(0, m_SortColumn)
    Else
        .TextMatrix(0, m_SortColumn) = "< " & .TextMatrix(0, m_SortColumn)
    End If

    ' Display the FlexGrid.
    .Visible = True
    
    End With
End Sub


'=================== Oculta o Muestra el Menu Principal =================================

Public Sub menuvisible(ByVal forma As Form, ByVal ban As Boolean)
    

End Sub


'=======================  Procedimiento retardar una ventana ==================================

Public Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    
    'this is a callback function.  This means that windows "calls back" to this function
    'when it's time for the timer event to fire
    'first thing we do is kill the timer so that no other timer events will fire
    KillTimer hwnd, idEvent
    
    'select the type of manipulation that we want to perform
    Select Case idEvent
    Case NV_CLOSEMSGBOX '// we want to close this messagebox after 4 seconds
        Dim hMessageBox As Long
        
        'find the messagebox window
        'change the text to whatever the title of the message box is
        hMessageBox = FindWindow("#32770", "TRANSACCIONES [ AVISO ]")
        
        'if we found it make sure it has the keyboard focus and then send it an enter to dismiss it
        If hMessageBox Then
            Call SetForegroundWindow(hMessageBox)
            
            'this will result in the default option being chosen
            SendKeys "{enter}"
        End If
    End Select
End Sub


'=======================  Funcion nombre del Usuario ==================================

Public Function fun_NombreUsuario() As String
Const UNLEN = 256   ' Max user name length.
Dim user_name As String
Dim name_len As Long

    user_name = Space$(UNLEN + 1)
    name_len = Len(user_name)
    If GetUserName(user_name, name_len) = 0 Then
        fun_NombreUsuario = "<unknown>"
    Else
        fun_NombreUsuario = Left$(user_name, name_len - 1)
    End If
End Function

'==================================
'Obtiene Nombre de Pc
'==================================


Public Function fun_NombrePc() As String
Dim compname As String, retval As Long   ' string to use as buffer & return value

compname = Space(255)                    ' set a large enough buffer for the computer name
retval = GetComputerName(compname, 255)  ' get the computer's name
                                         ' Remove the trailing null character from the strong
fun_NombrePc = Left(compname, InStr(compname, vbNullChar) - 1)

  End Function

'==================================
'Procedimiento Valida Numeros
'==================================
Function pro_valida_numeros(KeyAscii As Integer)
    Select Case KeyAscii
    Case 48 To 57, 52, 13, 8, 46
    Case Else
        KeyAscii = 0
    End Select
End Function


Public Sub Autocomplete(Lvw As ListView, sFind, Mytextbox As TextBox)
   Dim Lvfindtm As ListItem
   Dim TempSelStart As Integer
   Dim strTemp As String

   Set Lvfindtm = Lvw.findItem(sFind, lvwText, , lvwPartial)
   If Not Lvfindtm Is Nothing Then
      Lvfindtm.EnsureVisible
      Lvfindtm.Selected = True

      If execute Then
         TempSelStart = Mytextbox.SelStart
         Mytextbox.Text = CStr(Lvfindtm)
         If Not Mytextbox.Text = "" Then
            Mytextbox.SelStart = TempSelStart
            Mytextbox.SelLength = Len(Mytextbox.Text) - TempSelStart
         End If
      End If
   End If
End Sub


Public Sub pro_busca_registro(ByVal lv As ListView, ByVal valor As String, ByVal var_codigo As Boolean)
Dim itmfound As ListItem
    If var_codigo Then
        Set itmfound = lv.findItem(valor, lvwSubItem, , lvwPartial)
    Else
        Set itmfound = lv.findItem(valor, lvwText, , lvwPartial)
    End If
    
    If itmfound Is Nothing Then
       Set itmfound = lv.findItem(valor, lvwText, , lvwPartial)
       If itmfound Is Nothing Then
          Set itmfound = lv.findItem(valor, lvwSubItem, , lvwPartial)
          If itmfound Is Nothing Then
             MsgBox "No se Encontro Informacion", vbExclamation, "ATENCION"
             Exit Sub
          Else
             itmfound.EnsureVisible
             itmfound.Selected = True
             lv.SetFocus
          End If
       Else
          itmfound.EnsureVisible
          itmfound.Selected = True
          lv.SetFocus
       End If
    Else
        itmfound.EnsureVisible
        itmfound.Selected = True
        lv.SetFocus
    End If
End Sub







'***********************************************************************************************
'*
'*                         PROCEDIMIENTOS GLOBALES DE BASE DE DATOS
'*
'*                    _ BUSCAR UN REGISTRO Y REGRESAR CUALQUIER CAMPO
'*                    _ BUSCAR UNA LLAVE SIGUIENTE EN CASO QUE LA LLAVE SEA NUMERICA
'*                    _ ELIMINAR UN REGISTRO DE DE CUALQUIER TABLA MEDIANTE SU PARAMETRO
'*
'*
'***********************************************************************************************

'=============================== PROCEDIMIENTOS PARA BUSQUEDAS =================================

'UBICAR UN REGISTRO
Public Function Obtener_llave(cn As ADODB.Connection, ByVal rs As Recordset, ByVal var_tabla As String _
, ByVal var_campo As String, ByVal var_comparar As String, n As Integer, ByVal var_tipo As String) As String
Dim i As Integer
Dim l As Integer
Dim Cadena As String
Dim var_comparar_2 As String

'On Error GoTo HELL

 '   If var_tipo = "N" Then
 '      rs.Open "select * from " & var_tabla & " where " & var_campo & " = " + Trim(var_comparar), cn, adOpenDynamic, adLockOptimistic
 '       If Not rs.EOF Then
 '         Obtener_llave = IIf(IsNull(rs(n).Value), "", rs(n).Value)
 '       End If
 '       rs.Close
 '   Else
 '        l = Len(Trim(var_comparar))
 '        var_comparar_2 = Trim(var_comparar)
 '        Cadena = ""
 '        For i = 1 To l
 '            If Mid(var_comparar_2, i, 1) = "'" Then
 '               Cadena = Cadena + "'" + Mid(var_comparar_2, i, 1)
 '            Else
 '               Cadena = Cadena + Mid(var_comparar_2, i, 1)
 '            End If
 '        Next i
 '        var_comparar_2 = Cadena
 '        var_comparar = var_comparar_2
 '        'MsgBox cn.ConnectionString
 '        rs.Open "select * from " & var_tabla & " where " & var_campo & " =  '" + Trim(var_comparar) + "'", cn, adOpenDynamic, adLockOptimistic
 '        If Not rs.EOF Then
 '           Obtener_llave = IIf(IsNull(rs(n).Value), "", rs(n).Value)
 '        Else
 '           Obtener_llave = ""
 '        End If
  '       rs.Close
 '     End If
SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    GoTo SIGUE
End Function



'=============================== PROCEDIMEINTOS LLAVE SIGIENTE =================================

Public Function Siguiente(ByVal cn As ADODB.Connection, ByVal rs As Recordset, ByVal var_tabla As String, ByVal var_order_by As String) As String

Siguiente = ""
On Error GoTo HELL

'rs.Open "select * from " & var_tabla & " order by " & var_order_by, cn, adOpenDynamic, adLockOptimistic
'If Not rs.EOF Then
'  rs.MoveLast
'  Siguiente = rs(0) + 1
'Else
'    Siguiente = 1
'End If
'rs.Close

SIGUE:
On Error GoTo 0
Exit Function
HELL:
    MensajeError = Err.Description
    GoTo SIGUE
End Function



'=========================== PROCEDIMEINTOS ELIMINAR UN REGISTRO ===============================


Public Function Eliminar(var_pro_almacenado As String, var_pro_llave As String, var_texto As String) As Boolean

Eliminar = True
'On Error GoTo HELL

Set CMD.ActiveConnection = cnn                      'Esta es la conexión activa
CMD.CommandType = adCmdStoredProc                   'Aquí le indico a ADO que se trata de un PA
    
CMD.CommandText = var_pro_almacenado                     'Abrir Procedimiento Almacenado y Agregar Banco
    CMD(var_pro_llave) = var_texto
          
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


' Procedimiento para llenar grids con el contenido de una tabla, la forma mas rapida
'_____________________________________________________________________________________




Public Function numero_letras(monpag, var_moneda)
'para monpag,canstr

Dim digito As String
Dim unidad As String
Dim decena As String
Dim ceros1 As String
Dim ceros2 As String
Dim posici As Integer
Dim nument As String
Dim lonnum As Integer
Dim lonnum2 As Integer
Dim numfra As String
Dim var_nombre_moneda As String
Dim var_nombre_moneda_plural As String
Dim var_moneda_local As Integer
canstr = ""
digito = ""
unidad = ""
decena = ""
ceros1 = ""
ceros2 = ""
posici = 1

'nument = LTrim(Str(Int(monpag), 9, 0))
'lonnum = Len(nument)
'numfra = Str(((monpag - Int(monpag)) * 100), 2, 0)
canstr = ""

nument = LTrim(Str(Fix(monpag)))
lonnum = Len(nument)
numfra = Str(Fix((monpag - Int(monpag)) * 100))


While lonnum <> 0
   digito = Mid(nument, lonnum, 1)
   Select Case posici
      Case 2
         unidad = Mid(nument, lonnum + 1, 1)
      Case 5
         unidad = Mid(nument, lonnum + 1, 1)
      Case 8
         unidad = Mid(nument, lonnum + 1, 1)
      Case 3
         decena = Mid(nument, lonnum + 1, 1)
         unidad = Mid(nument, lonnum + 2, 1)
      Case 6
         decena = Mid(nument, lonnum + 1, 1)
         unidad = Mid(nument, lonnum + 2, 1)
      Case 9
         decena = Mid(nument, lonnum + 1, 1)
         unidad = Mid(nument, lonnum + 2, 1)
      Case 1
         If lonnum - 1 = 0 Then
            lonnum2 = 1
         Else
            lonnum2 = lonnum - 1
         End If
         decena = Mid(nument, lonnum2, 1)
      Case 4
         If lonnum - 1 = 0 Then
            lonnum2 = 1
         Else
            lonnum2 = lonnum - 1
         End If
         decena = Mid(nument, lonnum2, 1)
      Case 7
         If lonnum - 1 = 0 Then
            lonnum2 = 1
         Else
            lonnum2 = lonnum - 1
         End If
         decena = Mid(nument, lonnum2, 1)
   End Select
   Select Case posici
      Case 4
         If lonnum > 3 And lonnum < 7 Then
            Select Case lonnum
               Case 4
                  ceros1 = Mid(nument, 2, 3)
                  ceros2 = Mid(nument, 5, 3)
               Case 5
                  ceros1 = Mid(nument, 3, 3)
                  ceros2 = Mid(nument, 6, 3)
               Case 6
                  ceros1 = Mid(nument, 4, 3)
                  ceros2 = Mid(nument, 7, 3)
            End Select
            If ceros1 = "000" Then
               If ceros2 = "000" Then
                  canstr = "DE " + canstr
               Else
                  canstr = canstr
               End If
            Else
               canstr = "MIL " + canstr
            End If
         Else
            canstr = "MIL " + canstr
         End If
      Case "1"
         If posici = 7 And lonnum = 1 Then
            canstr = "MILLON " + canstr
         End If
      Case 7
         canstr = "MILLONES " + canstr
   End Select
   If digito = "1" Then
      If (posici = 1 Or posici = 4 Or posici = 7) And decena <> "1" Then
         canstr = "UN " + canstr
      End If
      If (posici = 2 Or posici = 5 Or posici = 8) Then
         Select Case unidad
            Case "0"
               canstr = "DIEZ " + canstr
            Case "1"
               canstr = "ONCE " + canstr
            Case "2"
               canstr = "DOCE " + canstr
            Case "3"
               canstr = "TRECE " + canstr
            Case "4"
               canstr = "CATORCE " + canstr
            Case "5"
               canstr = "QUINCE " + canstr
            Case "6"
               canstr = "DIECISEIS " + canstr
            Case "7"
               canstr = "DIECISIETE " + canstr
            Case "8"
               canstr = "DIECIOCHO " + canstr
            Case "9"
               canstr = "DIECINUEVE " + canstr
         End Select
      End If
      If (posici = 3 Or posici = 6 Or posici = 9) Then
         If unidad = "0" And decena = "0" Then
            canstr = "CIEN " + canstr
         Else
            canstr = "CIENTO " + canstr
         End If
      End If
   End If
   If digito = "2" Then
      If (posici = 1 Or posici = 4 Or posici = 7) And decena <> "1" Then
         canstr = "DOS " + canstr
      End If
      If (posici = 2 Or posici = 5 Or posici = 8) Then
         If unidad = "0" Then
            canstr = "VEINTE " + canstr
         Else
            canstr = "VEINTI" + canstr
         End If
      End If
      If (posici = 3 Or posici = 6 Or posici = 9) Then
         canstr = "DOSCIENTOS " + canstr
      End If
   End If
   If digito = "3" Then
      If (posici = 1 Or posici = 4 Or posici = 7) And decena <> "1" Then
         canstr = "TRES " + canstr
      End If
      If (posici = 2 Or posici = 5 Or posici = 8) Then
         If unidad = "0" Then
            canstr = "TREINTA " + canstr
         Else
            canstr = "TREINTA Y " + canstr
         End If
      End If
      If (posici = 3 Or posici = 6 Or posici = 9) Then
         canstr = "TRECIENTOS " + canstr
      End If
   End If
   If digito = "4" Then
      If (posici = 1 Or posici = 4 Or posici = 7) And decena <> "1" Then
         canstr = "CUATRO " + canstr
      End If
      If (posici = 2 Or posici = 5 Or posici = 8) Then
         If unidad = "0" Then
            canstr = "CUARENTA " + canstr
         Else
            canstr = "CUARENTA Y " + canstr
         End If
      End If
      If (posici = 3 Or posici = 6 Or posici = 9) Then
         canstr = "CUATROCIENTOS " + canstr
      End If
   End If
   If digito = "5" Then
      If (posici = 1 Or posici = 4 Or posici = 7) And decena <> "1" Then
         canstr = "CINCO " + canstr
      End If
      If (posici = 2 Or posici = 5 Or posici = 8) Then
         If unidad = "0" Then
            canstr = "CINCUENTA " + canstr
         Else
            canstr = "CINCUENTA Y " + canstr
         End If
      End If
      If (posici = 3 Or posici = 6 Or posici = 9) Then
         canstr = "QUINIENTOS " + canstr
      End If
   End If
   If digito = "6" Then
      If (posici = 1 Or posici = 4 Or posici = 7) And decena <> "1" Then
         canstr = "SEIS " + canstr
      End If
      If (posici = 2 Or posici = 5 Or posici = 8) Then
         If unidad = "0" Then
            canstr = "SESENTA " + canstr
         Else
            canstr = "SESENTA Y " + canstr
         End If
      End If
      If (posici = 3 Or posici = 6 Or posici = 9) Then
         canstr = "SEISCIENTOS " + canstr
      End If
   End If
   If digito = "7" Then
      If (posici = 1 Or posici = 4 Or posici = 7) And decena <> "1" Then
         canstr = "SIETE " + canstr
      End If
      If (posici = 2 Or posici = 5 Or posici = 8) Then
         If unidad = "0" Then
            canstr = "SETENTA " + canstr
         Else
            canstr = "SETENTA Y " + canstr
         End If
      End If
      If (posici = 3 Or posici = 6 Or posici = 9) Then
         canstr = "SETECIENTOS " + canstr
      End If
   End If
   If digito = "8" Then
      If (posici = 1 Or posici = 4 Or posici = 7) And decena <> "1" Then
         canstr = "OCHO " + canstr
      End If
      If (posici = 2 Or posici = 5 Or posici = 8) Then
         If unidad = "0" Then
            canstr = "OCHENTA " + canstr
         Else
            canstr = "OCHENTA Y " + canstr
         End If
      End If
      If (posici = 3 Or posici = 6 Or posici = 9) Then
         canstr = "OCHOCIENTOS " + canstr
      End If
   End If
   If digito = "9" Then
      If (posici = 1 Or posici = 4 Or posici = 7) And decena <> "1" Then
         canstr = "NUEVE " + canstr
      End If
      If (posici = 2 Or posici = 5 Or posici = 8) Then
         If unidad = "0" Then
            canstr = "NOVENTA " + canstr
         Else
            canstr = "NOVENTA Y " + canstr
         End If
      End If
      If (posici = 3 Or posici = 6 Or posici = 9) Then
         canstr = "NOVECIENTOS " + canstr
      End If
   End If
   If digito = "0" And lonnum = 1 And posici = 1 Then
      canstr = "CERO " + canstr
   End If
   lonnum = lonnum - 1
   posici = posici + 1
Wend
rsaux11.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + var_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
var_moneda_local = IIf(IsNull(rsaux11!inte_mon_moneda_local), 0, rsaux11!inte_mon_moneda_local)
var_nombre_moneda = rsaux11!vcha_mon_nombre_plural
rsaux11.Close
If var_moneda_local = 1 Then
   canstr = "(" + canstr + " PESOS " + numfra + "/100 M.N.)"
Else
   canstr = "(" + canstr + " " + Trim(var_nombre_moneda) + " " + numfra + "/100 " + ")"
End If
End Function


Public Function Fechas(formita As Form) As Date
    Fechas = frmcalendario.mes.Value
End Function


