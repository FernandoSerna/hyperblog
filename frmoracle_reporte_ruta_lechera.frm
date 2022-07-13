VERSION 5.00
Begin VB.Form frmoracle_reporte_ruta_lechera 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ruta lechera"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   3195
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2820
      Picture         =   "frmoracle_reporte_ruta_lechera.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   30
      Picture         =   "frmoracle_reporte_ruta_lechera.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   -15
      TabIndex        =   3
      Top             =   375
      Width           =   3210
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   45
      TabIndex        =   0
      Top             =   360
      Width           =   3105
      Begin VB.TextBox txt_embarque 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1080
         TabIndex        =   2
         Top             =   255
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Embarque:"
         Height          =   210
         Left            =   150
         TabIndex        =   1
         Top             =   330
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmoracle_reporte_ruta_lechera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim comandoORA As New adodb.Command
Dim parametro As adodb.Parameter

Private Sub cmd_imprimir_Click()
      If IsNumeric(Me.txt_embarque) Then
         strconsulta = "SELECT INTE_EMB_EMBARQUE, SOURCE_HEADER_NUMBER, MAX(CUSTOMER_NAME) AS CLIENTE, MAX(ENTREGA) AS ESTABLECIMIENTO, SUM(FLOA_SAL_CANTIDAD_LEIDA) AS CANTIDAD FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = ?  GROUP BY INTE_EMB_EMBARQUE, SOURCE_HEADER_NUMBER ORDER BY INTE_EMB_EMBARQUE, SOURCE_HEADER_NUMBER"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_embarque))
              .Parameters.Append parametro
         End With
         Set rsaux9 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         cnn.BeginTrans
         
         rs.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_ORACLE_REPORTE_RUTA_LECHERA", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
         Else
            var_consecutivo = 0
         End If
         var_consecutivo = var_consecutivo + 1
         rs.Close
         rs.Open "insert into TB_TEMP_ORACLE_REPORTE_RUTA_LECHERA (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
         cnn.CommitTrans
         strconsulta = "SELECT TRANSPORTE, TO_CHAR(FECHA_FIN, 'DD/MM/YYYY HH24:MI:SS') AS FECHA_FIN, TO_CHAR(FECHA_INICIO, 'DD/MM/YYYY HH24:MI:SS') FECHA_INICIO, chofer FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_embarque))
              .Parameters.Append parametro
         End With
         Set rsaux11 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         If Not rsaux11.EOF Then
            var_chofer = IIf(IsNull(rsaux11!chofer), "", rsaux11!chofer)
         Else
            var_chofer = ""
         End If
         If var_chofer = "" Then
            VAR_NOMBRE_CHOFER = "                                                     "
         Else
            rsaux6.Open "select * from tb_choferes where vcha_cho_chofer_id = '" + var_chofer + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux6.EOF Then
               VAR_NOMBRE_CHOFER = IIf(IsNull(rsaux6!vcha_cho_nombre), "                                              ", rsaux6!vcha_cho_nombre)
            Else
               VAR_NOMBRE_CHOFER = "                                             "
            End If
            rsaux6.Close
         End If
         rsaux6.Open "select * from tb_usuarios where vcha_usu_usuario_id = '" + var_clave_usuario_global + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux6.EOF Then
            var_usuario_facturacion = IIf(IsNull(rsaux6!vcha_usu_nombre), "", rsaux6!vcha_usu_nombre) + " " + IIf(IsNull(rsaux6!vcha_usu_apellidos), "", rsaux6!vcha_usu_apellidos)
         Else
            var_usuario_facturacion = ""
         End If
         rsaux6.Close
         If Not rsaux9.EOF Then
            var_fecha_embarque = IIf(IsNull(rsaux11!fecha_fin), rsaux11!fecha_inicio, rsaux11!fecha_fin)
            rsaux5.Open "SELECT * FROM TB_ORACLE_TRANSPORTES where clave = '" + IIf(IsNull(rsaux11!transporte), "", rsaux11!transporte) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux5.EOF Then
               var_transporte = IIf(IsNull(rsaux5!nombre), "", rsaux5!nombre)
            Else
               var_transporte = ""
            End If
            rsaux5.Close
            var_cadena_sellos = ""
            rsaux5.Open "select * from tb_sellos where inte_emb_embarque = " + CStr(CDbl(Me.txt_embarque)), cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux5.EOF
                  If var_cadena_sellos = "" Then
                     var_cadena_sellos = IIf(IsNull(rsaux5!vcha_sel_Sello), "", rsaux5!vcha_sel_Sello)
                  Else
                     var_cadena_sellos = var_cadena_sellos + ", " + IIf(IsNull(rsaux5!vcha_sel_Sello), "", rsaux5!vcha_sel_Sello)
                  End If
                  rsaux5.MoveNext
            Wend
            rsaux5.Close
            
            
            strconsulta = "SELECT DISTINCT  J.SALESREP_ID, J.NAME  FROM OE_ORDER_HEADERS_ALL OHA, XXVIA_TB_SALIDAS_CAJAS, XXVIA_VENDEDORES J WHERE  INTE_EMB_EMBARQUE = ? AND OHA.ORDER_NUMBER = SOURCE_HEADER_NUMBER AND OHA.SALESREP_ID = J.SALESREP_ID AND J.SALESREP_ID <> -3"
            'strconsulta = "SELECT DISTINCT  J.SALESREP_ID, J.NAME  FROM OE_ORDER_HEADERS_ALL OHA, XXVIA_TB_SALIDAS_CAJAS, XXVIA_VENDEDORES J WHERE  INTE_EMB_EMBARQUE = ? AND OHA.ORDER_NUMBER = SOURCE_HEADER_NUMBER AND OHA.SALESREP_ID = J.SALESREP_ID"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_embarque))
                 .Parameters.Append parametro
            End With
            Set rsaux5 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            
            VAR_CADENA_RUTAS = ""
            While Not rsaux5.EOF
                  If VAR_CADENA_RUTAS = "" Then
                     VAR_CADENA_RUTAS = IIf(IsNull(rsaux5!Name), "", rsaux5!Name)
                  Else
                     VAR_CADENA_RUTAS = VAR_CADENA_RUTAS + ", " + IIf(IsNull(rsaux5!Name), "", rsaux5!Name)
                  End If
                  rsaux5.MoveNext
            Wend
            rsaux5.Close
            
            While Not rsaux9.EOF
                  If rsaux7.State = 1 Then
                     rsaux7.Close
                  End If
                  rsaux7.Open "select * from TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES WHERE PEDIDO = " + CStr(rsaux9!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux7.EOF Then
                     var_orden = IIf(IsNull(rsaux7!orden_pedido), 0, rsaux7!orden_pedido)
                  Else
                     var_orden = 0
                  End If
                  rsaux7.Close
                  If rsaux9!source_header_number < 10000012 Then
                     strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rsaux9!source_header_number)
                          .Parameters.Append parametro
                     End With
                     Set rsaux8 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     If rsaux8!order_type_id = 1002 Then
                        var_nombre_cliente = IIf(IsNull(rsaux9!Cliente), "", rsaux9!Cliente)
                        If var_nombre_cliente = "VIANNEY TEXTIL HOGAR SA DE CV" Then
                           var_nombre_cliente = IIf(IsNull(rsaux9!establecimiento), var_nombre_cliente, rsaux9!establecimiento)
                        End If
                        var_tipo = "Nota de envío"
                        var_folio = rsaux9!source_header_number
                        rsaux10.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        strconsulta = "SELECT source_document_id FROM OE_ORDER_HEADERS_ALL WHERE  ORDER_NUMBER = ? "
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(rsaux9!source_header_number))
                             .Parameters.Append parametro
                        End With
                        Set rsaux5 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        If Not rsaux5.EOF Then
                           var_source_document_id = IIf(IsNull(rsaux5!source_document_id), 0, rsaux5!source_document_id)
                        End If
                        rsaux5.Close
'----- -------
                        strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id = ? AND secondary_inventory_name = A.ATTRIBUTE1"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_source_document_id)
                             .Parameters.Append parametro
                        End With
                        Set rsaux7 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                      
                        If Not rsaux7.EOF Then
                           var_clave_cliente = rsaux7!attribute1
                        Else
                           var_clave_cliente = ""
                        End If
                        rsaux7.Close
                        If rsaux3.State = 1 Then
                           rsaux3.Close
                        End If
                        rsaux3.Open "select * from mtl_secondary_inventories where secondary_inventory_name = '" + var_clave_cliente + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux3.EOF Then
                           var_location_id = IIf(IsNull(rsaux3!LOCATION_ID), 0, rsaux3!LOCATION_ID)
                           If var_location_id > 0 Then
                              rsaux4.Open "select ADDRESS_LINE_1, ADDRESS_LINE_2, TOWN_OR_CITY, REGION_1, COUNTRY, POSTAL_CODE  from hr_locations_all where location_id = '" + CStr(CDbl(var_location_id)) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              VAR_CALLE = Mid(IIf(IsNull(rsaux4!ADDRESS_LINE_1), "", rsaux4!ADDRESS_LINE_1), 1, 50)
                              VAR_DIRECCION = Mid(IIf(IsNull(rsaux4!ADDRESS_LINE_1), "", rsaux4!ADDRESS_LINE_1), 1, 50)
                              VAR_COLONIA = IIf(IsNull(rsaux4!ADDRESS_LINE_2), "", rsaux4!ADDRESS_LINE_2)
                              var_ciudad = IIf(IsNull(rsaux4!TOWN_OR_CITY), "", rsaux4!TOWN_OR_CITY)
                              var_estado = IIf(IsNull(rsaux4!REGION_1), "", rsaux4!REGION_1)
                              var_pais = IIf(IsNull(rsaux4!COUNTRY), "", rsaux4!COUNTRY)
                              VAR_CP = IIf(IsNull(rsaux4!POSTAL_CODE), "", rsaux4!POSTAL_CODE)
                              rsaux4.Close
                           End If
                        End If
                        rsaux3.Close
                     
                     
                     Else
                        var_tipo = "Factura"
                        strconsulta = "SELECT A.NAME FROM OE_TRANSACTION_TYPES_TL A, OE_ORDER_HEADERS_ALL B WHERE A.TRANSACTION_TYPE_ID = B.ORDER_TYPE_ID AND B.ORDER_NUMBER = ? AND ROWNUM = 1"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rsaux9!source_header_number)
                             .Parameters.Append parametro
                        End With
                        Set rsaux7 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
'--------
                        strconsulta = "SELECT  hps.pArty_site_number as clave_cliente , HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.SHIP_TO_ORG_ID AND oha.order_number = ? AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux9!source_header_number)
                             .Parameters.Append parametro
                        End With
                        Set rsaux6 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                           
                        If Not rsaux6.EOF Then
                           var_clave_cliente = IIf(IsNull(rsaux6!clave_cliente), "", rsaux6!clave_cliente)
                           var_nombre_cliente = IIf(IsNull(rsaux6!customer_name), "", rsaux6!customer_name)
                           VAR_DIRECCION = Mid(IIf(IsNull(rsaux6!calle), "", rsaux6!calle) + " " + IIf(IsNull(rsaux6!NUMERO), "", rsaux6!NUMERO), 1, 50)
                           VAR_COLONIA = IIf(IsNull(rsaux6!colonia), "", rsaux6!colonia)
                           var_ciudad = IIf(IsNull(rsaux6!ciudad), "", rsaux6!ciudad)
                           VAR_MUNICIPIO = IIf(IsNull(rsaux6!municipio), "", rsaux6!municipio)
                           var_estado = IIf(IsNull(rsaux6!estado), "", rsaux6!estado)
                           var_pais = IIf(IsNull(rsaux6!pais), "", rsaux6!pais)
                           VAR_CP = IIf(IsNull(rsaux6!cp), "", rsaux6!cp)
                           VAR_DIRECCION = VAR_DIRECCION + " " + VAR_COLONIA
                        Else
                           VAR_DIRECCION = ""
                           var_clave_cliente = ""
                           var_nombre_cliente = ""
                           VAR_COLONIA = ""
                           var_ciudad = ""
                           VAR_MUNICIPIO = ""
                           var_estado = ""
                           var_pais = ""
                           VAR_CP = ""
                        End If
                        rsaux6.Close
                        rsaux7.Close
                     End If
                     VAR_CLIENTE = var_clave_cliente + " " + var_nombre_cliente
                  Else
                     rs.Open "select * from TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where pedido = '" + CStr(rsaux9!source_header_number) + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        'AQUI EMPIEZA LA BUSQUEDA
                        rsaux5.Open "select * from XXVIA_TB_CLIENTES_RUTAS_DISTR where nombre_establecimiento like  '" + rs!Cliente + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux5.EOF Then
                        
                        
                           strconsulta = "select * from xxvia_vw_clientes_bcp where  site_use_id = ?"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(rsaux5!establecimiento))
                                .Parameters.Append parametro
                           End With
                           Set rsaux4 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           If Not rsaux4.EOF Then
                              var_clave_cliente = IIf(IsNull(rsaux5!establecimiento), "", rsaux5!establecimiento)
                              var_nombre_cliente = IIf(IsNull(rs!Cliente), "", rs!Cliente)
                              VAR_CALLE = Mid(IIf(IsNull(rsaux4!calle), "", rsaux4!calle), 1, 50) + " " + IIf(IsNull(rsaux4!NUM_CALLE), "", rsaux4!NUM_CALLE)
                              VAR_DIRECCION = Mid(IIf(IsNull(rsaux4!calle), "", rsaux4!calle), 1, 50) + " " + IIf(IsNull(rsaux4!NUM_CALLE), "", rsaux4!NUM_CALLE)
                              VAR_COLONIA = IIf(IsNull(rsaux4!colonia), "", rsaux4!colonia)
                              var_ciudad = IIf(IsNull(rsaux4!ciudad), "", rsaux4!ciudad)
                              var_estado = IIf(IsNull(rsaux4!estado), "", rsaux4!estado)
                              var_pais = IIf(IsNull(rsaux4!pais), "", rsaux4!pais)
                              VAR_CP = IIf(IsNull(rsaux4!codigo_postal), "", rsaux4!codigo_postal)
                              VAR_CLIENTE = var_clave_cliente + " " + var_nombre_cliente
                           End If
                           rsaux4.Close
                        End If
                        rsaux5.Close
                     End If
                     rs.Close
                  
                  End If
                  
                  var_cadena = "INSERT INTO TB_TEMP_ORACLE_REPORTE_RUTA_LECHERA (INTE_TEM_CONSECUTIVO, EMBARQUE, CHOFER, PEDIDO, CLIENTE, CALLE,                                                        NUMERO, COLONIA, CIUDAD, MUNICIPIO, ESTADO, CP, ORDEN)"
                  var_cadena = var_cadena + " VALUES (" + CStr(var_consecutivo) + "," + Me.txt_embarque + ",'" + VAR_NOMBRE_CHOFER + "'," + CStr(rsaux9!source_header_number) + ",'" + var_clave_cliente + " " + var_nombre_cliente + "','" + VAR_DIRECCION + "','" + VAR_NUMERO + "','" + VAR_COLONIA + "','" + var_ciudad + "','" + VAR_MUNICIPIO + "','" + var_estado + "','" + VAR_CP + "'," + CStr(var_orden) + ")"
                  rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  rsaux9.MoveNext
            Wend
                                            
            rs.Open "DELETE FROM TB_TEMP_ORACLE_REPORTE_RUTA_LECHERA WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo) + " and PEDIDO is null", cnn, adOpenDynamic, adLockOptimistic
            
            Set reporte = appl.OpenReport(App.Path + "\rep_oracle_ruta_lechera.rpt")
            reporte.RecordSelectionFormula = "{VW_ORACLE_REPORTE_RUTA_LECHERA.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Relación de documentos"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            rs.Open "DELETE FROM TB_TEMP_ORACLE_REPORTE_RUTA_LECHERA WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         
                    
        Else
           MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
         End If
         rsaux9.Close
         rsaux11.Close
      End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 3200
   Left = 4200
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   End If
End Sub


