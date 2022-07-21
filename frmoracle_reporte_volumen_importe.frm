VERSION 5.00
Begin VB.Form frmoracle_reporte_volumen_importe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importe y volumen por embarque"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   60
      TabIndex        =   3
      Top             =   360
      Width           =   3390
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
         TabIndex        =   4
         Top             =   255
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Embarque:"
         Height          =   210
         Left            =   150
         TabIndex        =   5
         Top             =   330
         Width           =   765
      End
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   0
      TabIndex        =   2
      Top             =   375
      Width           =   3210
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   45
      Picture         =   "frmoracle_reporte_volumen_importe.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3105
      Picture         =   "frmoracle_reporte_volumen_importe.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
End
Attribute VB_Name = "frmoracle_reporte_volumen_importe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter

Private Sub cmd_imprimir_Click()
   If IsNumeric(Me.txt_embarque) Then
      rs.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "SELECT * from xxvia_tb_encabezado_embarques WHERE EMBARQUE = ?"
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
      If Not rsaux9.EOF Then
         VAR_ESTATUS = IIf(IsNull(rsaux9!char_emb_estatus), "", rsaux9!char_emb_estatus)
         If VAR_ESTATUS = "I" Or VAR_ESTATUS = "F" Then
            var_fecha_emabarque = IIf(IsNull(rsaux9!fecha_fin), rsaux9!FECHA_INICIO, rsaux9!fecha_fin)
            strconsulta = "select b.order_type_id, b.pricing_date, b.price_list_id, order_number, c.segment1, c.inventory_item_id, SUM(shipped_quantity) as cantidad ,SUM(shipped_quantity * UNIT_SELLING_PRICE) AS IMPORTE, sum(shipped_quantity *unit_volume) as volumen from oe_order_lines_all a, oe_order_headers_all b, xxvia_system_items_b c where a.header_id = b.header_id and order_number in (select distinct source_header_number from xxvia_tb_salidas_cajas where inte_emb_embarque = ?) and a.inventory_item_id = c.inventory_item_id and a.ship_from_org_id = c.organization_id group by b.order_type_id, b.pricing_date, b.price_list_id, order_number, c.segment1, c.inventory_item_id"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_embarque))
                 .Parameters.Append parametro
            End With
            Set rs = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            cnn.CommandTimeout = 720
            cnn.BeginTrans
            rsaux.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_ORACLE_reporte_IMPORTE_VOLUMEN", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rsaux!numero), 0, rsaux!numero)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rsaux.Close
            rsaux.Open "insert into TB_TEMP_ORACLE_reporte_IMPORTE_VOLUMEN (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            var_dia = CStr(Day(CDate(var_fecha_emabarque)))
            var_mes = CStr(Month(CDate(var_fecha_emabarque)))
            var_año = CStr(Year(CDate(var_fecha_emabarque)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_embarque = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            
            While Not rs.EOF
                  var_dia = CStr(Day(CDate(rs!pricing_date)))
                  var_mes = CStr(Month(CDate(rs!pricing_date)))
                  var_año = CStr(Year(CDate(rs!pricing_date)))
                  If Len(Trim(var_dia)) = 1 Then
                      var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  var_fecha = var_dia + "/" + var_mes + "/" + var_año
                  var_fecha_sql = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                  var_lista_precios = rs!price_list_id
                  var_codigo = rs!SEGMENT1
                  var_codigo_oracle = rs!inventory_item_id
                  
                  If rs!ORDER_TYPE_ID <> 1042 Then
                     strconsulta = "SELECT OPERAND as precio FROM qp_list_lines_v WHERE list_header_id =  ? AND  PRODUCT_ATTR_VAL_DISP = ? and start_date_active <= to_date(?,'DD/MM/YYYY') and (end_date_active is null or end_date_active >= to_date(?,'DD/MM/YYYY')) and Product_Attr_Value = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(var_lista_precios))
                          .Parameters.Append parametro
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, var_codigo)
                          .Parameters.Append parametro
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, var_fecha)
                          .Parameters.Append parametro
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, var_fecha)
                          .Parameters.Append parametro
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(var_codigo_oracle))
                          .Parameters.Append parametro
                     End With
                     Set rsaux1 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     If Not rsaux1.EOF Then
                        VAR_IMPORTE = IIf(IsNull(rsaux1!Precio), 0, rsaux1!Precio) * IIf(IsNull(rs!cantidad), 0, rs!cantidad)
                     Else
                        VAR_IMPORTE = rs!Importe
                     End If
                     rsaux1.Close
                  Else
                     VAR_IMPORTE = IIf(IsNull(rs!Importe), 0, rs!Importe)
                  End If
                  rsaux1.Open "insert into TB_TEMP_ORACLE_REPORTE_IMPORTE_VOLUMEN (inte_tem_consecutivo, pedido, lista_precios, tipo_pedido, codigo, cantidad, importe, volumen) values (" + CStr(var_consecutivo) + "," + CStr(rs!order_number) + "," + CStr(var_lista_precios) + "," + CStr(rs!ORDER_TYPE_ID) + ",'" + rs!SEGMENT1 + "'," + CStr(IIf(IsNull(rs!cantidad), 0, rs!cantidad)) + "," + CStr(IIf(IsNull(VAR_IMPORTE), 0, VAR_IMPORTE)) + "," + CStr(IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN)) + ")", cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
            Wend
            rs.Close
            rs.Open "select distinct pedido, tipo_pedido from TB_TEMP_ORACLE_REPORTE_IMPORTE_VOLUMEN where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido is not null", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  If rs!tipo_pedido = 1002 Then
                     strconsulta = "SELECT source_document_id FROM OE_ORDER_HEADERS_ALL WHERE  ORDER_NUMBER = ? "
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(rs!pedido))
                          .Parameters.Append parametro
                     End With
                     Set rsaux5 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     If Not rsaux5.EOF Then
                        var_source_document_id = IIf(IsNull(rsaux5!source_document_id), 0, rsaux5!source_document_id)
                     End If
                     rsaux5.Close
'------------
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
                        var_nombre_cliente = IIf(IsNull(rsaux3!Description), 0, rsaux3!Description)
                     Else
                        var_nombre_cliente = "VIANNEY TEXTIL HOGAR SA DE CV"
                     End If
                     rsaux3.Close
                  Else
                     strconsulta = "SELECT  hps.pArty_site_number as clave_cliente , HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.SHIP_TO_ORG_ID AND oha.order_number = ? AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rs!pedido)
                          .Parameters.Append parametro
                     End With
                     Set rsaux3 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                          
                     If Not rsaux3.EOF Then
                        var_clave_cliente = IIf(IsNull(rsaux3!clave_cliente), "", rsaux3!clave_cliente)
                        var_nombre_cliente = IIf(IsNull(rsaux3!customer_name), "", rsaux3!customer_name)
                     Else
                        var_clave_cliente = ""
                        var_nombre_cliente = ""
                     End If
                     rsaux3.Close
                  End If
                  rsaux3.Open "update TB_TEMP_ORACLE_REPORTE_IMPORTE_VOLUMEN set clave_cliente = '" + var_clave_cliente + "', Nombre_cliente = '" + var_nombre_cliente + "', embarque = " + Me.txt_embarque + ", fecha_embarque = " + var_fecha_embarque + "  where pedido = " + CStr(rs!pedido) + " and inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
            Wend
            rs.Close
            Set reporte = appl.OpenReport(App.Path + "\rep_oracle_reporte_importe_volumen_embarque.rpt")
            reporte.RecordSelectionFormula = "{VW_ORACLE_REPORTE_IMPORTE_VOLUMEN.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Relación de documentos"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            rs.Open "DELETE FROM TB_TEMP_ORACLE_REPORTE_IMPORTE_VOLUMEN WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            
            
            
         Else
            MsgBox "El embarque no a sido cerrador", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
      End If
      rsaux9.Close
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



