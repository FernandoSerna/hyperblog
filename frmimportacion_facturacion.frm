VERSION 5.00
Begin VB.Form frmimportacion_facturacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importación de Facturación"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_aceptar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   45
      Picture         =   "frmimportacion_facturacion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Aceptar"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2850
      Picture         =   "frmimportacion_facturacion.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   -15
      TabIndex        =   2
      Top             =   330
      Width           =   3225
   End
   Begin VB.Frame Frame1 
      Caption         =   " Fecha a Importar "
      Height          =   825
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   3105
      Begin VB.TextBox txt_fecha 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   840
         TabIndex        =   1
         Top             =   315
         Width           =   1380
      End
   End
End
Attribute VB_Name = "frmimportacion_facturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tabla As ADODB.Connection


Private Sub cmd_aceptar_Click()
   Set var_tabla = CreateObject("ADODB.connection")
   Set TB_ENCABEZADO_CARTERA_I = New TB_ENCABEZADO_CARTERA_I
   Set TB_ESTADO_CUENTA_INSERTA = New TB_ESTADO_CUENTA_INSERTA
   Dim var_cliente As String
   cnn.CommandTimeout = 6000
   If IsDate(Me.txt_fecha) Then
      If var_clave_usuario_global = "U0000000019" Then
         If var_empresa = "15" Then
            If cnn_importacion.State = 1 Then
               cnn_importacion.Close
            End If
            cnn_importacion.Open "Provider=OraOLEDB.Oracle.1;User ID=INTERFACE;Data Source=VENTAS;Extended Properties=;Persist Security Info=True;Password=INTERFACE"
            cnn_importacion.CursorLocation = adUseClient
            'var_cadena = "select c.faccod, c.facfch, c.clicod, c.facsernum, round(sum(l.facmts*l.facpremts)*1.16,2) as IMPTOTALDO from cfaven@cipic.vianney.com.mx c, lfaven@cipic.vianney.com.mx l Where l.faccod = c.faccod and c.facfch = TO_DATE('" + Me.txt_fecha + "', 'DD/MM/YYYY') group by c.faccod, c.facfch, c.clicod, c.facsernum"
            'var_cadena = "select c.faccod, c.facfch, c.clicod, c.facsernum, round(sum(l.facmts*l.facpremts)*1.16,2) as IMPTOTALDO from cfaven@cipic.vianney.com.mx c, lfaven@cipic.vianney.com.mx l Where l.faccod = c.faccod and c.facfch = TO_DATE('" + Me.txt_fecha + "', 'DD/MM/YYYY') group by c.faccod, c.facfch, c.clicod, c.facsernum"
            var_fecha_fin_1 = CDate(txt_fecha) + 1
            
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
          
            var_fecha_fin = var_dia + "/" + var_mes + "/" + var_año
         
            var_cadena = "select c.faccod, c.facfch, c.clicod, 'CRF' as facsernum, round(sum(l.facmts*l.facpremts)*1.16,2) as IMPTOTALDO from cfaven@cipic.vianney.com.mx c, lfaven@cipic.vianney.com.mx l Where l.faccod = c.faccod and c.facfch between TO_DATE('" + Me.txt_fecha + "', 'DD/MM/YYYY') and TO_DATE('" + Me.txt_fecha + "', 'DD/MM/YYYY')  group by c.faccod, c.facfch, c.clicod, c.facsernum"


            Text1 = var_cadena
            
            rs.Open var_cadena, cnn_importacion, adOpenDynamic, adLockOptimistic
            var_clientes_faltantes = ""
            While Not rs.EOF
                  var_cliente = Trim(rs!clicod)
                  If var_cliente = "100002" Or var_cliente = "100005" Or var_cliente = "100602" Or var_cliente = "100603" Or var_cliente = "100601" Or var_cliente = "100002" Or var_cliente = "100608" Then
                     var_cliente = "100001"
                  End If
                  rsaux.Open "select vcha_cli_nombre from vw_clientes where textilera = '" + Trim(var_cliente) + "'", cnn, adOpenDynamic, adLockOptimistic
                  If rsaux.EOF Then
                     If var_clientes_faltantes = "" Then
                        var_clientes_faltantes = var_cliente
                     Else
                        var_clientes_faltantes = var_clientes_faltantes + ", " + var_cliente
                     End If
                  End If
                  rsaux.Close
                  rs.MoveNext
            Wend
            var_contador_facturas = 0
            If rs.RecordCount > 0 Then
               rs.MoveFirst
               If Trim(var_clientes_faltantes) = "" Then
                  If Trim(var_clientes_faltantes) = "" Then
                     While Not rs.EOF
                           rsaux.Open "SELECT * FROM TB_ENCABEZADO_CARTERA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_cAR_DOCUMENTO = 'FA' AND VCHA_SER_SERIE_ID = '" + rs!facsernum + "' AND INTE_CAR_NUMERO = " + CStr(rs!faccod), cnn, adOpenDynamic, adLockOptimistic
                           If rsaux.EOF Then
                              If rsaux2.State = 1 Then
                                 rsaux2.Close
                              End If
                              var_cliente = Trim(rs!clicod)
                              If var_cliente = "100002" Or var_cliente = "100005" Or var_cliente = "100602" Or var_cliente = "100603" Or var_cliente = "100601" Or var_cliente = "100608" Then
                                 var_cliente = "100001"
                              Else
                                 var_cliente = Trim(rs!clicod)
                              End If
                              If rsaux2.State = 1 Then
                                 rsaux2.Close
                              End If
                              rsaux2.Open "select vcha_cli_clave_id from vw_clientes where textilera = '" + var_cliente + "' and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                              'MsgBox var_cliente
                              var_cliente = IIf(IsNull(rsaux2!vcha_cli_clave_id), "", rsaux2!vcha_cli_clave_id)
                              rsaux2.Close
                              rsaux2.Open "SELECT * FROM VW_CLIENTES WHERE VCHA_CLI_CLAVE_ID = '" + Trim(var_cliente) + "'", cnn, adOpenDynamic, adLockOptimistic
                        
                              txt_nombre_cliente = IIf(IsNull(rsaux2!VCHA_CLI_NOMBRE), "", rsaux2!VCHA_CLI_NOMBRE)
                              txt_plazo = IIf(IsNull(rsaux2!inte_pla_dias), 0, rsaux2!inte_pla_dias)
                              var_grupo_actual = IIf(IsNull(rsaux2!VCHA_GAC_GRUPO_aCTUAL_ID), "", rsaux2!VCHA_GAC_GRUPO_aCTUAL_ID)
                              var_grupo_real = IIf(IsNull(rsaux2!vcha_gre_grupo_real_id), "", rsaux2!vcha_gre_grupo_real_id)
                              var_titular = IIf(IsNull(rsaux2!vcha_tit_titular_id), "", rsaux2!vcha_tit_titular_id)
                              var_clave_moneda = IIf(IsNull(rsaux2!vcha_mon_moneda_id), "", rsaux2!vcha_mon_moneda_id)
                              var_agente = IIf(IsNull(rsaux2!vcha_AGE_aGENTE_ID), "", rsaux2!vcha_AGE_aGENTE_ID)
                              var_plazo = IIf(IsNull(rsaux2!inte_pla_dias), 0, rsaux2!inte_pla_dias)
                              txt_plazo = var_plazo
                              var_iva = IIf(IsNull(rsaux2!FLOA_TPE_IVA), 0, rsaux2!FLOA_TPE_IVA)
                              var_clave_moneda = IIf(IsNull(rsaux2!vcha_mon_moneda_id), "", rsaux2!vcha_mon_moneda_id)
                              var_tipo_Cambio = 1
                        
                              var_numero_folio = rs!faccod
                              var_importe_total = (rs!IMPTOTALDO / (1 + (var_iva / 100))) * var_tipo_Cambio
                              var_importe_neto = rs!IMPTOTALDO * var_tipo_Cambio
                              var_importe_iva = var_importe_neto - var_importe_total
                              var_subimporte = var_importe_total
                              var_insertar = False
                              var_serie = rs!facsernum
                              var_tipo_comision = 0
                              cnn.BeginTrans
                              Cadena = "INSERT INTO TB_ENCABEZADO_CARTERA (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_CAR_CLASE_ID, INTE_CAR_NUMERO, CHAR_CAR_AFECTACION, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO, DTIM_CAR_FECHA, VCHA_AGE_AGENTE_ID, VCHA_GAC_GRUPO_ACTUAL_ID, VCHA_GRE_GRUPO_REAL_ID, VCHA_TIT_TITULAR_ID, VCHA_CLI_CLAVE_ID, VCHA_ESB_ESTABLECIMIENTO_ID, INTE_CAR_PLAZO, FLOA_CAR_PORCENTAJE_IVA, FLOA_CAR_PORCENTAJE_IMPUESTO_1, FLOA_CAR_PORCENTAJE_IMPUESTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_1, "
                              Cadena = Cadena + " FLOA_CAR_PORCENTAJE_DESCUENTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_3, FLOA_CAR_IMPORTE_TOTAL, FLOA_CAR_IMPORTE_IVA, FLOA_CAR_IMPORTE_IMPUESTO_1, FLOA_CAR_IMPORTE_IMPUESTO_2, FLOA_CAR_IMPORTE_DESCUENTO_1, FLOA_CAR_IMPORTE_DESCUENTO_2, FLOA_CAR_IMPORTE_DESCUENTO_3, FLOA_CAR_SUBIMPORTE, FLOA_CAR_IMPORTE_NETO, VCHA_CAR_IMPORTE_LETRA, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA, VCHA_AUD_FECHA, FLOA_CAR_SALDO, DTIM_CAR_FECHA_VENCIMIENTO, DTIM_CAR_FECHA_ENTREGA, VCHA_MON_MONEDA_ID, FLOA_CAR_TIPO_CAMBIO, VCHA_SER_SERIE_ID, CHAR_CAR_ESTATUS, INTE_cAR_TIPO_COMISION) Values "
                              Cadena = Cadena + " ('" + var_empresa + "', '" + var_unidad_organizacional + "', 'FA', 'FA', 'FA', " + CStr(var_numero_folio) + ", '+', '', '', 0, '" + Me.txt_fecha + "', '" + var_agente + "', '" + var_grupo_actual + "', '" + var_grupo_real + "', '" + var_titular + "', '" + Trim(var_cliente) + "', '', " + CStr(var_plazo) + ", " + CStr(var_iva) + ", 0, 0, 0, 0, 0, " + CStr(var_importe_total) + ", " + CStr(var_importe_iva) + ", 0, 0, 0, 0, 0, " + CStr(var_subimporte) + ", " + CStr(var_importe_neto) + ", '', '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', GETDate(), 0, GETDate(), GETDate(), '" + var_clave_moneda + "', " + CStr(var_tipo_Cambio) + ", '" + var_serie + "', ''," + CStr(var_tipo_comision) + ")"
                              rsaux5.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                              var_inserta = TB_ESTADO_CUENTA_INSERTA.Anadir(var_empresa, CStr(Trim(var_serie)), "FA", CDbl(var_numero_folio), "", "", 0, CDbl(var_importe_neto), 0)
                              rsaux9.Open "SELECT * FROM TB_INVENTARIO_DOCUMENTOS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_SER_SERIE_ID = '" + var_serie + "' AND INTE_CAR_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                              If rsaux9.EOF Then
                                 rsaux8.Open "INSERT INTO TB_INVENTARIO_DOCUMENTOS (VCHA_EMP_EMPRESA_ID, VCHA_AGE_AGENTE_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_CAR_CLASE_ID, INTE_CAR_NUMERO, CHAR_CAR_AFECTACION, VCHA_SER_SERIE_ID, CHAR_IDO_ESTATUS, FLOA_IDO_CANTIDAD, FLOA_CAR_IMPORTE_NETO, FLOA_CAR_TIPO_CAMBIO, VCHA_MON_MONEDA_ID, DTIM_IDO_FECHA_ENTRAGA, VCHA_CLI_CLAVE_ID, INTE_EMB_EMBARQUE) VALUES ('" + var_empresa + "','" + var_agente + "', 'FA', 'FA', 'FA', " + CStr(var_numero_folio) + ",'+', '" + var_serie + "','A',0," + CStr(var_importe_neto) + "," + CStr(var_tipo_Cambio) + ",'" + var_clave_moneda + "',GETDATE(),'" + Trim(var_cliente) + "',0)", cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux9.Close
                              
                              
                              cnn.CommitTrans
                              var_contador_facturas = var_contador_facturas + 1
                           End If
                           rsaux.Close
                           rs.MoveNext
                     Wend
                  End If
               Else
                  MsgBox "Faltan las siguientes claves de cliente de dar de alta " + var_clientes_faltantes, vbOKOnly, "ATENCION"
               End If
            End If
            rs.Close
            MsgBox "Se insertaron " + CStr(var_contador_facturas) + " facturas", vbOKOnly, "ATENCION"
         Else
            MsgBox "No se puede migrar la información de la empresa seleccionada", vbOKOnly, "ATENCION"
         End If
      Else
         If var_empresa = "15" Then
            If cnn_importacion.State = 1 Then
               cnn_importacion.Close
            End If
            cnn_importacion.Open "Provider=OraOLEDB.Oracle.1;User ID=INTERFACE;Data Source=VENTAS;Extended Properties=;Persist Security Info=True;Password=INTERFACE"
            cnn_importacion.CursorLocation = adUseClient
            'var_cadena = "select c.faccod, c.facfch, c.clicod, c.facsernum, round(sum(l.facmts*l.facpremts)*1.16,2) as IMPTOTALDO from cfaven@cipic.vianney.com.mx c, lfaven@cipic.vianney.com.mx l Where l.faccod = c.faccod and c.facfch = TO_DATE('" + Me.txt_fecha + "', 'DD/MM/YYYY') group by c.faccod, c.facfch, c.clicod, c.facsernum"
            'var_cadena = "select c.faccod, c.facfch, c.clicod, c.facsernum, round(sum(l.facmts*l.facpremts)*1.16,2) as IMPTOTALDO from cfaven@cipic.vianney.com.mx c, lfaven@cipic.vianney.com.mx l Where l.faccod = c.faccod and c.facfch = TO_DATE('" + Me.txt_fecha + "', 'DD/MM/YYYY') group by c.faccod, c.facfch, c.clicod, c.facsernum"
            var_fecha_fin_1 = CDate(txt_fecha) + 1
            
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
          
            var_fecha_fin = var_dia + "/" + var_mes + "/" + var_año
         
            var_cadena = "select c.faccod, c.facfch, c.clicod, c.facsernum, round(sum(l.facmts*l.facpremts)*1.16,2) as IMPTOTALDO from cfaven@cipic.vianney.com.mx c, lfaven@cipic.vianney.com.mx l Where l.faccod = c.faccod and c.facfch between TO_DATE('" + Me.txt_fecha + "', 'DD/MM/YYYY') and TO_DATE('" + var_fecha_fin + "', 'DD/MM/YYYY')  group by c.faccod, c.facfch, c.clicod, c.facsernum"


            Text1 = var_cadena
            
            rs.Open var_cadena, cnn_importacion, adOpenDynamic, adLockOptimistic
            var_clientes_faltantes = ""
            While Not rs.EOF
                  var_cliente = Trim(rs!clicod)
                  If var_cliente = "100002" Or var_cliente = "100005" Or var_cliente = "100602" Or var_cliente = "100002" Or var_cliente = "100608" Or var_cliente = "100601" Then
                     var_cliente = "100001"
                  End If
                  rsaux.Open "select vcha_cli_nombre from vw_clientes where textilera = '" + Trim(var_cliente) + "'", cnn, adOpenDynamic, adLockOptimistic
                  If rsaux.EOF Then
                     If var_clientes_faltantes = "" Then
                        var_clientes_faltantes = var_cliente
                     Else
                        var_clientes_faltantes = var_clientes_faltantes + ", " + var_cliente
                     End If
                  End If
                  rsaux.Close
                  rs.MoveNext
            Wend
            var_contador_facturas = 0
            If rs.RecordCount > 0 Then
               rs.MoveFirst
               If Trim(var_clientes_faltantes) = "" Then
                  If Trim(var_clientes_faltantes) = "" Then
                     While Not rs.EOF
                           rsaux.Open "SELECT * FROM TB_ENCABEZADO_CARTERA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_cAR_DOCUMENTO = 'FA' AND VCHA_SER_SERIE_ID = '" + rs!facsernum + "' AND INTE_CAR_NUMERO = " + CStr(rs!faccod), cnn, adOpenDynamic, adLockOptimistic
                           If rsaux.EOF Then
                              If rsaux2.State = 1 Then
                                 rsaux2.Close
                              End If
                              var_cliente = Trim(rs!clicod)
                              If var_cliente = "100002" Or var_cliente = "100005" Or var_cliente = "100602" Or var_cliente = "100002" Or var_cliente = "100608" Or var_cliente = "100601" Then
                                 var_cliente = "100001"
                              Else
                                 var_cliente = Trim(rs!clicod)
                              End If
                              If var_cliente = "100609" Then
                                 rsaux2.Open "select vcha_cli_clave_id from vw_clientes where textilera = '" + var_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 rsaux2.Open "select vcha_cli_clave_id from vw_clientes where textilera = '" + var_cliente + "' and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                              End If
                              var_cliente = IIf(IsNull(rsaux2!vcha_cli_clave_id), "", rsaux2!vcha_cli_clave_id)
                              rsaux2.Close
                              rsaux2.Open "SELECT * FROM VW_CLIENTES WHERE VCHA_CLI_CLAVE_ID = '" + Trim(var_cliente) + "'", cnn, adOpenDynamic, adLockOptimistic
                        
                              txt_nombre_cliente = IIf(IsNull(rsaux2!VCHA_CLI_NOMBRE), "", rsaux2!VCHA_CLI_NOMBRE)
                              txt_plazo = IIf(IsNull(rsaux2!inte_pla_dias), 0, rsaux2!inte_pla_dias)
                              var_grupo_actual = IIf(IsNull(rsaux2!VCHA_GAC_GRUPO_aCTUAL_ID), "", rsaux2!VCHA_GAC_GRUPO_aCTUAL_ID)
                              var_grupo_real = IIf(IsNull(rsaux2!vcha_gre_grupo_real_id), "", rsaux2!vcha_gre_grupo_real_id)
                              var_titular = IIf(IsNull(rsaux2!vcha_tit_titular_id), "", rsaux2!vcha_tit_titular_id)
                              var_clave_moneda = IIf(IsNull(rsaux2!vcha_mon_moneda_id), "", rsaux2!vcha_mon_moneda_id)
                              var_agente = IIf(IsNull(rsaux2!vcha_AGE_aGENTE_ID), "", rsaux2!vcha_AGE_aGENTE_ID)
                              var_plazo = IIf(IsNull(rsaux2!inte_pla_dias), 0, rsaux2!inte_pla_dias)
                              txt_plazo = var_plazo
                              var_iva = IIf(IsNull(rsaux2!FLOA_TPE_IVA), 0, rsaux2!FLOA_TPE_IVA)
                              var_clave_moneda = IIf(IsNull(rsaux2!vcha_mon_moneda_id), "", rsaux2!vcha_mon_moneda_id)
                              var_tipo_Cambio = 1
                        
                              var_numero_folio = rs!faccod
                              var_importe_total = (rs!IMPTOTALDO / (1 + (var_iva / 100))) * var_tipo_Cambio
                              var_importe_neto = rs!IMPTOTALDO * var_tipo_Cambio
                              var_importe_iva = var_importe_neto - var_importe_total
                              var_subimporte = var_importe_total
                              var_insertar = False
                              var_serie = rs!facsernum
                              var_tipo_comision = 0
                              cnn.BeginTrans
                              Cadena = "INSERT INTO TB_ENCABEZADO_CARTERA (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_CAR_CLASE_ID, INTE_CAR_NUMERO, CHAR_CAR_AFECTACION, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO, DTIM_CAR_FECHA, VCHA_AGE_AGENTE_ID, VCHA_GAC_GRUPO_ACTUAL_ID, VCHA_GRE_GRUPO_REAL_ID, VCHA_TIT_TITULAR_ID, VCHA_CLI_CLAVE_ID, VCHA_ESB_ESTABLECIMIENTO_ID, INTE_CAR_PLAZO, FLOA_CAR_PORCENTAJE_IVA, FLOA_CAR_PORCENTAJE_IMPUESTO_1, FLOA_CAR_PORCENTAJE_IMPUESTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_1, "
                              Cadena = Cadena + " FLOA_CAR_PORCENTAJE_DESCUENTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_3, FLOA_CAR_IMPORTE_TOTAL, FLOA_CAR_IMPORTE_IVA, FLOA_CAR_IMPORTE_IMPUESTO_1, FLOA_CAR_IMPORTE_IMPUESTO_2, FLOA_CAR_IMPORTE_DESCUENTO_1, FLOA_CAR_IMPORTE_DESCUENTO_2, FLOA_CAR_IMPORTE_DESCUENTO_3, FLOA_CAR_SUBIMPORTE, FLOA_CAR_IMPORTE_NETO, VCHA_CAR_IMPORTE_LETRA, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA, VCHA_AUD_FECHA, FLOA_CAR_SALDO, DTIM_CAR_FECHA_VENCIMIENTO, DTIM_CAR_FECHA_ENTREGA, VCHA_MON_MONEDA_ID, FLOA_CAR_TIPO_CAMBIO, VCHA_SER_SERIE_ID, CHAR_CAR_ESTATUS, INTE_cAR_TIPO_COMISION) Values "
                              Cadena = Cadena + " ('" + var_empresa + "', '" + var_unidad_organizacional + "', 'FA', 'FA', 'FA', " + CStr(var_numero_folio) + ", '+', '', '', 0, '" + Me.txt_fecha + "', '" + var_agente + "', '" + var_grupo_actual + "', '" + var_grupo_real + "', '" + var_titular + "', '" + Trim(var_cliente) + "', '', " + CStr(var_plazo) + ", " + CStr(var_iva) + ", 0, 0, 0, 0, 0, " + CStr(var_importe_total) + ", " + CStr(var_importe_iva) + ", 0, 0, 0, 0, 0, " + CStr(var_subimporte) + ", " + CStr(var_importe_neto) + ", '', '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', GETDate(), 0, GETDate(), GETDate(), '" + var_clave_moneda + "', " + CStr(var_tipo_Cambio) + ", '" + var_serie + "', ''," + CStr(var_tipo_comision) + ")"
                              rsaux5.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                              var_inserta = TB_ESTADO_CUENTA_INSERTA.Anadir(var_empresa, CStr(Trim(var_serie)), "FA", CDbl(var_numero_folio), "", "", 0, CDbl(var_importe_neto), 0)
                              rsaux9.Open "SELECT * FROM TB_INVENTARIO_DOCUMENTOS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_SER_SERIE_ID = '" + var_serie + "' AND INTE_CAR_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                              If rsaux9.EOF Then
                                 
                                 rsaux8.Open "INSERT INTO TB_INVENTARIO_DOCUMENTOS (VCHA_EMP_EMPRESA_ID, VCHA_AGE_AGENTE_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_CAR_CLASE_ID, INTE_CAR_NUMERO, CHAR_CAR_AFECTACION, VCHA_SER_SERIE_ID, CHAR_IDO_ESTATUS, FLOA_IDO_CANTIDAD, FLOA_CAR_IMPORTE_NETO, FLOA_CAR_TIPO_CAMBIO, VCHA_MON_MONEDA_ID, DTIM_IDO_FECHA_ENTRAGA, VCHA_CLI_CLAVE_ID, INTE_EMB_EMBARQUE) VALUES ('" + var_empresa + "','" + var_agente + "', 'FA', 'FA', 'FA', " + CStr(var_numero_folio) + ",'+', '" + var_serie + "','A',0," + CStr(var_importe_neto) + "," + CStr(var_tipo_Cambio) + ",'" + var_clave_moneda + "',GETDATE(),'" + Trim(var_cliente) + "',0)", cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux9.Close
                              
                              
                              cnn.CommitTrans
                              var_contador_facturas = var_contador_facturas + 1
                           End If
                           rsaux.Close
                           rs.MoveNext
                     Wend
                  End If
               Else
                  MsgBox "Faltan las siguientes claves de cliente de dar de alta " + var_clientes_faltantes, vbOKOnly, "ATENCION"
               End If
            End If
            rs.Close
            MsgBox "Se insertaron " + CStr(var_contador_facturas) + " facturas", vbOKOnly, "ATENCION"
         Else
            var_ruta = ""
            rs.Open "select VCHA_EMP_RUTA_FACTURAS_EXTERNAS from tb_empresas where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
            var_ruta = IIf(IsNull(rs(0).Value), "", rs(0).Value)
            rs.Close
            If var_ruta <> "" Then
               If var_tabla.State = 1 Then
                  var_tabla.Close
               End If
               If var_empresa = "18" Then
                  var_ruta = "g:\sistemas\desarrollo\textilera\"
                  var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + var_ruta + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
                  'var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=t;UID=;SourceDB=g:\sistemas\desarrollo\textilera\;SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
               Else
                  var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + var_ruta + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
               End If
            
               var_dia = CStr(Day(CDate(txt_fecha)))
               var_mes = CStr(Month(CDate(txt_fecha)))
               var_año = CStr(Year(CDate(txt_fecha)))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               var_fecha = var_mes + "/" + var_dia + "/" + var_año
               If var_empresa = "16" Then
                  var_cadena = "SELECT     MGP10008.fecvencedo - MGP10008.fecdocto as plazo,  MGP10002.RAZSOCIAL,MGP10008.aniodocto, MGP10008.perdocto, MGP10008.numtipodoc, MGP10008.seriedocto, MGP10008.numdocto, MGP10008.afectestad, MGP10008.clasedoc, MGP10008.codcteprov, MGP10008.diaesp, MGP10008.fecdocto, MGP10008.fecenvio, MGP10008.imptotaldo, MGP10002.represcte fROM MGP10002 INNER JOIN MGP10008 ON MGP10002.codcteprov = MGP10008.codcteprov where MGP10008.fecdocto = CTOD('" + var_fecha + "') and MGP10008.numtipodoc = '100' and MGP10008.cancela <> '1'"
               Else
                  If var_empresa = "17" Then
                     var_cadena = "SELECT     MGP10008.fecvencedo - MGP10008.fecdocto as plazo, MGP10002.RAZSOCIAL,MGP10008.aniodocto, MGP10008.perdocto, MGP10008.numtipodoc, 'LBI' as seriedocto, MGP10008.numdocto, MGP10008.afectestad, MGP10008.clasedoc, MGP10008.codcteprov, MGP10008.diaesp, MGP10008.fecdocto, MGP10008.fecenvio, MGP10008.imptotaldo, MGP10002.represcte fROM MGP10002 INNER JOIN MGP10008 ON MGP10002.codcteprov = MGP10008.codcteprov where MGP10008.fecdocto = CTOD('" + var_fecha + "') and (allt(MGP10008.numtipodoc) = '1' or allt(MGP10008.numtipodoc) = '2')  and MGP10008.cancela <> '1'"
                  End If
                  If var_empresa = "06" Then
                     var_cadena = "SELECT   MGP10008.fecvencedo - MGP10008.fecdocto as plazo,  MGP10002.RAZSOCIAL,MGP10008.aniodocto, MGP10008.perdocto, MGP10008.numtipodoc, 'Q0Z' as seriedocto, MGP10008.numdocto, MGP10008.afectestad, MGP10008.clasedoc, MGP10008.codcteprov, MGP10008.diaesp, MGP10008.fecdocto, MGP10008.fecenvio, MGP10008.imptotaldo, MGP10002.represcte fROM MGP10002 INNER JOIN MGP10008 ON MGP10002.codcteprov = MGP10008.codcteprov where MGP10008.fecdocto = CTOD('" + var_fecha + "') and (allt(MGP10008.numtipodoc) = '1' or allt(MGP10008.numtipodoc) = '2')  and MGP10008.cancela <> '1'"
                  End If
                  If var_empresa = "18" Then
                     var_cadena = "SELECT distinct 'AA' as SERIEDOCTO, a.cvecliente,allt(c.nombre)+' '+allt(c.paterno)+' '+allt(c.materno) as RAZSOCIAL, c.cvecuenta as represcte, a.fecha, a.importe, a.tipo, a.factura, b.fecha_pago - b.fecha as plazo, a.factura as numdocto, a.importe as IMPTOTALDO from enca_fac a, facturas b, clientes c where a.fecha = CTOD('" + var_fecha + "') and a.factura = b.factura and allt(a.cvecliente) = allt(c.cvecliente) and a.estatus <> 'C'"
                  End If
               End If
               ''aqui
               rs.Open var_cadena, var_tabla, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_clientes_faltantes = ""
                  While Not rs.EOF
                        If rsaux.State = 1 Then
                           rsaux.Close
                        End If
                        rsaux.Open "select vcha_cli_nombre from tb_clientes where vcha_cli_clave_id = '" + Trim(rs!represcte) + "'", cnn, adOpenDynamic, adLockOptimistic
                        If rsaux.EOF Then
                           If var_clientes_faltantes = "" Then
                              var_clientes_faltantes = Trim(rs!RAZSOCIAL)
                           Else
                              var_clientes_faltantes = var_clientes_faltantes + ", " + Trim(rs!RAZSOCIAL)
                           End If
                        End If
                        rs.MoveNext
                  Wend
                  rs.MoveFirst
                  var_contador_facturas = 0
                  If rsaux.State = 1 Then
                     rsaux.Close
                  End If
                  If Trim(var_clientes_faltantes) = "" Then
                     While Not rs.EOF
                           rsaux.Open "SELECT * FROM TB_ENCABEZADO_CARTERA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_cAR_DOCUMENTO = 'FA' AND VCHA_SER_SERIE_ID = '" + rs!SERIEDOCTO + "' AND INTE_CAR_NUMERO = " + CStr(rs!numdocto), cnn, adOpenDynamic, adLockOptimistic
                           If rsaux.EOF Then
                              If rsaux2.State = 1 Then
                                 rsaux2.Close
                              End If
                              rsaux2.Open "SELECT * FROM VW_CLIENTES WHERE VCHA_CLI_CLAVE_ID = '" + Trim(rs!represcte) + "'", cnn, adOpenDynamic, adLockOptimistic
                        
                              txt_nombre_cliente = IIf(IsNull(rsaux2!VCHA_CLI_NOMBRE), "", rsaux2!VCHA_CLI_NOMBRE)
                              txt_plazo = IIf(IsNull(rsaux2!inte_pla_dias), 0, rsaux2!inte_pla_dias)
                              var_grupo_actual = IIf(IsNull(rsaux2!VCHA_GAC_GRUPO_aCTUAL_ID), "", rsaux2!VCHA_GAC_GRUPO_aCTUAL_ID)
                              var_grupo_real = IIf(IsNull(rsaux2!vcha_gre_grupo_real_id), "", rsaux2!vcha_gre_grupo_real_id)
                              var_cliente = txt_clave_cliente
                              var_titular = IIf(IsNull(rsaux2!vcha_tit_titular_id), "", rsaux2!vcha_tit_titular_id)
                              var_clave_moneda = IIf(IsNull(rsaux2!vcha_mon_moneda_id), "", rsaux2!vcha_mon_moneda_id)
                              var_agente = IIf(IsNull(rsaux2!vcha_AGE_aGENTE_ID), "", rsaux2!vcha_AGE_aGENTE_ID)
                              var_plazo = IIf(IsNull(rs!PLAZO), 0, rs!PLAZO)
                              txt_plazo = var_plazo
                              var_iva = IIf(IsNull(rsaux2!FLOA_TPE_IVA), 0, rsaux2!FLOA_TPE_IVA)
                              var_clave_moneda = IIf(IsNull(rsaux2!vcha_mon_moneda_id), "", rsaux2!vcha_mon_moneda_id)
                              var_tipo_Cambio = 1
                              
                              var_numero_folio = rs!numdocto
                              var_importe_total = (rs!IMPTOTALDO / (1 + (var_iva / 100))) * var_tipo_Cambio
                              var_importe_neto = rs!IMPTOTALDO * var_tipo_Cambio
                              var_importe_iva = var_importe_neto - var_importe_total
                              var_subimporte = var_importe_total
                              var_insertar = False
                              var_serie = rs!SERIEDOCTO
                              var_tipo_comision = 0
                              cnn.BeginTrans
                              Cadena = "INSERT INTO TB_ENCABEZADO_CARTERA (VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_CAR_CLASE_ID, INTE_CAR_NUMERO, CHAR_CAR_AFECTACION, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO, DTIM_CAR_FECHA, VCHA_AGE_AGENTE_ID, VCHA_GAC_GRUPO_ACTUAL_ID, VCHA_GRE_GRUPO_REAL_ID, VCHA_TIT_TITULAR_ID, VCHA_CLI_CLAVE_ID, VCHA_ESB_ESTABLECIMIENTO_ID, INTE_CAR_PLAZO, FLOA_CAR_PORCENTAJE_IVA, FLOA_CAR_PORCENTAJE_IMPUESTO_1, FLOA_CAR_PORCENTAJE_IMPUESTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_1, "
                              Cadena = Cadena + " FLOA_CAR_PORCENTAJE_DESCUENTO_2, FLOA_CAR_PORCENTAJE_DESCUENTO_3, FLOA_CAR_IMPORTE_TOTAL, FLOA_CAR_IMPORTE_IVA, FLOA_CAR_IMPORTE_IMPUESTO_1, FLOA_CAR_IMPORTE_IMPUESTO_2, FLOA_CAR_IMPORTE_DESCUENTO_1, FLOA_CAR_IMPORTE_DESCUENTO_2, FLOA_CAR_IMPORTE_DESCUENTO_3, FLOA_CAR_SUBIMPORTE, FLOA_CAR_IMPORTE_NETO, VCHA_CAR_IMPORTE_LETRA, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA, VCHA_AUD_FECHA, FLOA_CAR_SALDO, DTIM_CAR_FECHA_VENCIMIENTO, DTIM_CAR_FECHA_ENTREGA, VCHA_MON_MONEDA_ID, FLOA_CAR_TIPO_CAMBIO, VCHA_SER_SERIE_ID, CHAR_CAR_ESTATUS, INTE_cAR_TIPO_COMISION) Values "
                              Cadena = Cadena + " ('" + var_empresa + "', '" + var_unidad_organizacional + "', 'FA', 'FA', 'FA', " + CStr(var_numero_folio) + ", '+', '', '', 0, '" + Me.txt_fecha + "', '" + var_agente + "', '" + var_grupo_actual + "', '" + var_grupo_real + "', '" + var_titular + "', '" + Trim(rs!represcte) + "', '', " + CStr(var_plazo) + ", " + CStr(var_iva) + ", 0, 0, 0, 0, 0, " + CStr(var_importe_total) + ", " + CStr(var_importe_iva) + ", 0, 0, 0, 0, 0, " + CStr(var_subimporte) + ", " + CStr(var_importe_neto) + ", '', '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', GETDate(), 0, GETDate(), GETDate(), '" + var_clave_moneda + "', " + CStr(var_tipo_Cambio) + ", '" + var_serie + "', ''," + CStr(var_tipo_comision) + ")"
                              rsaux5.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                              var_inserta = TB_ESTADO_CUENTA_INSERTA.Anadir(var_empresa, CStr(Trim(var_serie)), "FA", CDbl(var_numero_folio), "", "", 0, CDbl(var_importe_neto), 0)
                              rsaux9.Open "select * from tb_estado_cuenta where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_ecu_Serie_cargo = '" + CStr(Trim(var_serie)) + "' and vcha_ecu_movimiento_cargo = 'FA' and inte_Ecu_numero_cargo = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                              If rsaux9.EOF Then
                                 rsaux8.Open "insert into tb_estado_cuenta (vcha_emp_empresa, vcha_uor_unidad_id, vcha_ecu_serie_cargo, vcha_ecu_movimiento_cargo, inte_Ecu_numero_cargo, floa_ecu_importe_cargo) values ('" + var_empresa + "','" + var_unidad_organizacional + "' ,'" + CStr(Trim(var_serie)) + "', 'FA', " + CStr(CDbl(var_numero_folio)) + "," + CStr(CDbl(var_importe_neto)), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux9.Close
                              
                              rsaux9.Open "SELECT * FROM TB_INVENTARIO_DOCUMENTOS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_SER_SERIE_ID = '" + var_serie + "' AND INTE_CAR_NUMERO = " + CStr(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                              If rsaux9.EOF Then
                                 rsaux8.Open "INSERT INTO TB_INVENTARIO_DOCUMENTOS (VCHA_EMP_EMPRESA_ID, VCHA_AGE_AGENTE_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_CAR_CLASE_ID, INTE_CAR_NUMERO, CHAR_CAR_AFECTACION, VCHA_SER_SERIE_ID, CHAR_IDO_ESTATUS, FLOA_IDO_CANTIDAD, FLOA_CAR_IMPORTE_NETO, FLOA_CAR_TIPO_CAMBIO, VCHA_MON_MONEDA_ID, DTIM_IDO_FECHA_ENTRAGA, VCHA_CLI_CLAVE_ID, INTE_EMB_EMBARQUE) VALUES ('" + var_empresa + "','" + var_agente + "', 'FA', 'FA', 'FA', " + CStr(var_numero_folio) + ",'+', '" + var_serie + "','A',0," + CStr(var_importe_neto) + "," + CStr(var_tipo_Cambio) + ",'" + var_clave_moneda + "',GETDATE(),'" + Trim(rs!represcte) + "',0)", cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux9.Close
                              
                              
                              
                              cnn.CommitTrans
                              var_contador_facturas = var_contador_facturas + 1
                           End If
                           rsaux.Close
                           rs.MoveNext
                     Wend
                 Else
                      MsgBox "Faltan los siguientes clientes por dar de alta en el S.I.D. " + var_clientes_faltantes, vbOKOnly, "ATENCION"
                 End If
                 MsgBox "Se insertaron " + CStr(var_contador_facturas) + " facturas", vbOKOnly, "ATENCION"
               Else
                  MsgBox "No existen facturas en la fecha seleccionada", vbOKOnly, "ATENCION"
               End If
               rs.Close
            Else
               MsgBox "No se a especificado una ruta donde se encuentren los archivos a importar", vbOKOnly, "ATENCION"
            End If
         End If
      End If
   Else
      MsgBox "Fecha incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 3850
   Me.txt_fecha = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub txt_fecha_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fecha) Then
         frmcalendario.mes.Value = CDate(Me.txt_fecha)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_fecha = var_fecha_general
   End If
End Sub
