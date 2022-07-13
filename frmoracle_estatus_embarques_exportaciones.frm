VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_estatus_embarques_exportaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estatus de embarques de exportaciones"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   18315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   18315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_embarque 
      Height          =   375
      Left            =   7320
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5520
      Picture         =   "frmoracle_estatus_embarques_exportaciones.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Buscar Movimiento"
      Top             =   240
      Width           =   330
   End
   Begin VB.TextBox txt_fin 
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   210
      Width           =   1575
   End
   Begin VB.TextBox txt_inicio 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   210
      Width           =   1575
   End
   Begin MSComctlLib.ListView lv_pedidos 
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   18255
      _ExtentX        =   32200
      _ExtentY        =   13785
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Embarque"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fecha Inicio"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha fin"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Pedido"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Estatus"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Cliente entrega"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Facturas"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Transportista"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Sello Barril"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha fin:"
      Height          =   195
      Left            =   3000
      TabIndex        =   4
      Top             =   300
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha inicio:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   300
      Width           =   900
   End
End
Attribute VB_Name = "frmoracle_estatus_embarques_exportaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objConn As New ADODB.Connection
Dim objCmd As New ADODB.Command
Dim objParm As ADODB.Parameter
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter


Private Sub cmd_buscar_Click()
   If IsDate(Me.txt_inicio) Then
      If IsDate(Me.txt_fin) Then
         If CDate(Me.txt_inicio) <= CDate(Me.txt_fin) Then
            Me.lv_pedidos.ListItems.Clear
            var_dia_s = CStr(Day(CDate(Me.txt_inicio)))
            var_mes_s = CStr(Month(CDate(Me.txt_inicio)))
            var_año_s = CStr(Year(CDate(Me.txt_inicio)))
            If Len(var_dia_s) = 1 Then
               var_dia_s = "0" + var_dia_s
            End If
            If Len(var_mes_s) = 1 Then
               var_mes_s = "0" + var_mes_s
            End If
            If Len(var_año_s) = 2 Then
               var_año_s = "20" + var_año_s
            End If
            var_fecha_inicio = var_dia_s + "/" + var_mes_s + "/" + var_año_s
            
            
            var_dia_s = CStr(Day(CDate(Me.txt_fin) + 1))
            var_mes_s = CStr(Month(CDate(Me.txt_fin) + 1))
            var_año_s = CStr(Year(CDate(Me.txt_fin) + 1))
            If Len(var_dia_s) = 1 Then
               var_dia_s = "0" + var_dia_s
            End If
            If Len(var_mes_s) = 1 Then
               var_mes_s = "0" + var_mes_s
            End If
            If Len(var_año_s) = 2 Then
               var_año_s = "20" + var_año_s
            End If
            var_fecha_fin = var_dia_s + "/" + var_mes_s + "/" + var_año_s
   
   
            var_cadena = "SELECT DISTINCT inte_Emb_Embarque, SOURCE_HEADER_NUMBER, fecha_inicio, fecha_fin, d.RAZON_SOCIAL_CLIENTE, c.CHAR_EMB_ESTATUS, nvl(transportista,' ') transportista, nvl(sello_barril, ' ') sello_barril "
            var_cadena = var_cadena + " FROM xxvia_tb_salidas_cajas A, OE_ORDER_HEADERS_ALL B, xxvia_Tb_encabezado_embarques c, xxvia_vw_clientes_bcp d "
            var_cadena = var_cadena + " Where ORDER_NUMBER = SOURCE_HEADER_NUMBER "
            var_cadena = var_cadena + " AND b.ORDER_TYPE_ID IN (1048,1051,1464, 2121) "
            var_cadena = var_cadena + " and a.inte_emb_embarque = c.embarque "
            var_cadena = var_cadena + " and b.ship_to_org_id =  d.site_use_id "
            var_cadena = var_cadena + " and C.fecha_inicio >= to_Date(?,'DD/MM/YYYY') anD C.fecha_inicio < to_Date(?,'DD/MM/YYYY') "
            var_cadena = var_cadena + " ORDER BY A.SOURCE_HEADER_NUMBER DESC "
   
            strconsulta = var_cadena
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, var_fecha_inicio)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, var_fecha_fin)
                 .Parameters.Append parametro
            End With
            Set rs = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
   
            While Not rs.EOF
                  Set list_item = lv_pedidos.ListItems.Add(, , rs!inte_Emb_Embarque)
                  list_item.SubItems(1) = rs!FECHA_INICIO
                  list_item.SubItems(2) = IIf(IsNull(rs!fecha_fin), "", rs!fecha_fin)
                  list_item.SubItems(3) = rs!source_header_number
                  var_cadena_facturas = ""
                  strconsulta = "select * from ra_customer_trx_all where ct_Reference = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, rs!source_header_number)
                       .Parameters.Append parametro
                  End With
                  Set rsaux = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  While Not rsaux.EOF
                        If var_cadena_facturas = "" Then
                           var_cadena_facturas = "FAEVII" + CStr(rsaux!trx_number)
                        Else
                           var_cadena_facturas = var_cadena_facturas + ", FAEVII" + CStr(rsaux!trx_number)
                        End If
                        rsaux.MoveNext
                  Wend
                  rsaux.Close
                  
                  list_item.SubItems(4) = IIf(IsNull(rs!char_emb_estatus), "", rs!char_emb_estatus)
                  list_item.SubItems(5) = rs!razon_social_cliente
                  list_item.SubItems(6) = var_cadena_facturas
                  list_item.SubItems(7) = rs!transportista
                  list_item.SubItems(8) = rs!sello_barril
                  rs.MoveNext
            Wend
            rs.Close
         Else
            MsgBox "La fecha de inicio no puede ser mayor a la fecha final.", vbOKOnly, "ATENCIOM"
         End If
      Else
         MsgBox "Fecha fin incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de inicio incorrecta", vbOKOnly, "ATENCION"
   End If

End Sub

Private Sub Form_Load()
   Me.txt_inicio = Date
   Me.txt_fin = Date
   
   var_dia_s = CStr(Day(CDate(Me.txt_inicio)))
   var_mes_s = CStr(Month(CDate(Me.txt_inicio)))
   var_año_s = CStr(Year(CDate(Me.txt_inicio)))
   If Len(var_dia_s) = 1 Then
      var_dia_s = "0" + var_dia_s
   End If
   If Len(var_mes_s) = 1 Then
      var_mes_s = "0" + var_mes_s
   End If
   If Len(var_año_s) = 2 Then
      var_año_s = "20" + var_año_s
   End If
   var_fecha_inicio = var_dia_s + "/" + var_mes_s + "/" + var_año_s
   
   
   var_dia_s = CStr(Day(CDate(Me.txt_fin) + 1))
   var_mes_s = CStr(Month(CDate(Me.txt_fin) + 1))
   var_año_s = CStr(Year(CDate(Me.txt_fin) + 1))
   If Len(var_dia_s) = 1 Then
      var_dia_s = "0" + var_dia_s
   End If
   If Len(var_mes_s) = 1 Then
      var_mes_s = "0" + var_mes_s
   End If
   If Len(var_año_s) = 2 Then
      var_año_s = "20" + var_año_s
   End If
   var_fecha_fin = var_dia_s + "/" + var_mes_s + "/" + var_año_s
   
   
   var_cadena = "SELECT DISTINCT inte_Emb_Embarque, SOURCE_HEADER_NUMBER, fecha_inicio, fecha_fin, d.RAZON_SOCIAL_CLIENTE, c.CHAR_EMB_ESTATUS, nvl(transportista,' ') transportista, nvl(sello_barril, ' ') sello_barril "
   var_cadena = var_cadena + " FROM xxvia_tb_salidas_cajas A, OE_ORDER_HEADERS_ALL B, xxvia_Tb_encabezado_embarques c, xxvia_vw_clientes_bcp d "
   var_cadena = var_cadena + " Where ORDER_NUMBER = SOURCE_HEADER_NUMBER "
   var_cadena = var_cadena + " AND b.ORDER_TYPE_ID IN (1048,1051,1464, 2121) "
   var_cadena = var_cadena + " and a.inte_emb_embarque = c.embarque "
   var_cadena = var_cadena + " and b.ship_to_org_id =  d.site_use_id "
   var_cadena = var_cadena + " and C.fecha_inicio >= to_Date(?,'DD/MM/YYYY') anD C.fecha_inicio < to_Date(?,'DD/MM/YYYY') "
   var_cadena = var_cadena + " ORDER BY A.SOURCE_HEADER_NUMBER DESC "
   
   strconsulta = var_cadena
   With comandoORA
        .ActiveConnection = cnnoracle_4
        .CommandType = adCmdText
        .CommandText = strconsulta
        Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, var_fecha_inicio)
        .Parameters.Append parametro
        Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, var_fecha_fin)
        .Parameters.Append parametro
   End With
   Set rs = comandoORA.execute
   Set comandoORA = Nothing
   Set parametro = Nothing
   
   While Not rs.EOF
         Set list_item = lv_pedidos.ListItems.Add(, , rs!inte_Emb_Embarque)
         list_item.SubItems(1) = rs!FECHA_INICIO
         list_item.SubItems(2) = IIf(IsNull(rs!fecha_fin), "", rs!fecha_fin)
         list_item.SubItems(3) = rs!source_header_number
         var_cadena_facturas = ""
         strconsulta = "select * from ra_customer_trx_all where ct_Reference = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, rs!source_header_number)
              .Parameters.Append parametro
         End With
         Set rsaux = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         While Not rsaux.EOF
               If var_cadena_facturas = "" Then
                  var_cadena_facturas = "FAEVII" + CStr(rsaux!trx_number)
               Else
                  var_cadena_facturas = var_cadena_facturas + ", FAEVII" + CStr(rsaux!trx_number)
               End If
               rsaux.MoveNext
         Wend
         rsaux.Close
         
         list_item.SubItems(4) = IIf(IsNull(rs!char_emb_estatus), "", rs!char_emb_estatus)
         list_item.SubItems(5) = rs!razon_social_cliente
         list_item.SubItems(6) = var_cadena_facturas
         list_item.SubItems(7) = rs!transportista
         list_item.SubItems(8) = rs!sello_barril
         rs.MoveNext
   Wend
   rs.Close
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_pedidos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_pedidos, ColumnHeader)
End Sub

Private Sub lv_pedidos_ItemClick(ByVal Item As MSComctlLib.ListItem)
   If Me.lv_pedidos.ListItems.Count > 0 Then
      Me.txt_embarque = Me.lv_pedidos.selectedItem
   End If
End Sub

Private Sub lv_pedidos_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 116 Then
      If IsNumeric(Me.txt_embarque) Then
         strconsulta = "SELECT EMBARQUE, CHAR_EMB_ESTATUS, CASE CHAR_EMB_ESTATUS WHEN 'F' THEN 'CERRADO' WHEN 'I' THEN 'CERRADO' ELSE 'ABIERTO' END ESTATUS, TRANSPORTISTA, SELLO_BARRIL, SELLO_LAMINA, SELLO_LATERALES, CERTIFICA_ADUANAL, DETALLE_CONTENIDO, COLOR_ETIQUETA, PALET FROM XXVIA_TB_ENCABEZADO_eMBARQUES WHERE EMBARQUE = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_embarque))
              .Parameters.Append parametro
         End With
         Set rsaux14 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         If Not rsaux14.EOF Then
            If IIf(IsNull(rsaux14!estatus), "ABIERTO", rsaux14!estatus) = "CERRADO" Then
               
            VAR_TRANSPORTISTA = IIf(IsNull(rsaux14!transportista), "", rsaux14!transportista)
            VAR_SELLO_BARRIL = IIf(IsNull(rsaux14!sello_barril), "", rsaux14!sello_barril)
            VAR_SELLO_LAMINA = IIf(IsNull(rsaux14!SELLO_LAMINA), "", rsaux14!SELLO_LAMINA)
            VAR_sELLO_LATERALES = IIf(IsNull(rsaux14!SELLO_LATERALes), "", rsaux14!SELLO_LATERALes)
            VAR_CERTIFICA_ADUANAL = IIf(IsNull(rsaux14!CERTIFICA_ADUANAL), "", rsaux14!CERTIFICA_ADUANAL)
            VAR_DETALLE_CONTENIDO = IIf(IsNull(rsaux14!DETALLE_CONTENIDO), "", rsaux14!DETALLE_CONTENIDO)
            VAR_COLOR_eTIQUETA = IIf(IsNull(rsaux14!COLOR_ETIQUETA), "", rsaux14!COLOR_ETIQUETA)
            VAR_PALET = IIf(IsNull(rsaux14!PALET), "", rsaux14!PALET)
               
               
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
               rs.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_ORACLE_REPORTE_CAJAS", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
               Else
                  var_consecutivo = 0
               End If
               var_consecutivo = var_consecutivo + 1
               rs.Close
               rs.Open "insert into TB_TEMP_ORACLE_REPORTE_CAJAS (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
               rs.Open "delete from TB_TEMP_ORACLE_CAJAS_ADUANA_DIVIDIDAS_EN_3 where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
               cnn.CommitTrans
               strconsulta = "SELECT TRANSPORTE, TO_CHAR(FECHA_FIN, 'DD/MM/YYYY HH24:MI:SS') AS FECHA_FIN, TO_CHAR(FECHA_INICIO, 'DD/MM/YYYY HH24:MI:SS') FECHA_INICIO, USUARIO_CERRO FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = ?"
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
               rsaux10.Open "SELECT * FROM TB_USUARIOS WHERE VCHA_USU_USUARIO_ID = '" + IIf(IsNull(rsaux11!USUARIO_CERRO), "", rsaux11!USUARIO_CERRO) + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux10.EOF Then
                  VAR_USUARIO_CERRO = IIf(IsNull(rsaux10!vcha_usu_nombre), "", rsaux10!vcha_usu_nombre) + " " + IIf(IsNull(rsaux10!vcha_usu_apellidos), "", rsaux10!vcha_usu_apellidos)
                  If VAR_USUARIO_CERRO = "" Then
                     rsaux6.Open "SELECT * FROM TB_USUARIOS WHERE VCHA_USU_USUARIO_ID = '" + var_clave_usuario_global + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux6.EOF Then
                        VAR_USUARIO_CERRO = IIf(IsNull(rsaux6!vcha_usu_nombre), "", rsaux6!vcha_usu_nombre) + " " + IIf(IsNull(rsaux6!vcha_usu_apellidos), "", rsaux6!vcha_usu_apellidos)
                     Else
                        VAR_USUARIO_CERRO = ""
                     End If
                     rsaux6.Close
                  End If
               Else
                  rsaux6.Open "SELECT * FROM TB_USUARIOS WHERE VCHA_USU_USUARIO_ID = '" + var_clave_usuario_global + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux6.EOF Then
                     VAR_USUARIO_CERRO = IIf(IsNull(rsaux6!vcha_usu_nombre), "", rsaux6!vcha_usu_nombre) + " " + IIf(IsNull(rsaux6!vcha_usu_apellidos), "", rsaux6!vcha_usu_apellidos)
                  Else
                     VAR_USUARIO_CERRO = ""
                  End If
                  rsaux6.Close
               End If
               rsaux10.Close
               If Not rsaux9.EOF Then
                  var_fecha_embarque = IIf(IsNull(rsaux11!fecha_fin), rsaux11!FECHA_INICIO, rsaux11!fecha_fin)
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
                           var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
'----
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
                              var_nombre_cliente = rsaux7!Description
                           Else
                              var_clave_cliente = ""
                              var_nombre_cliente = ""
                           End If
                           rsaux7.Close
                           If var_almacen_tienda <> "" Then
                        
                              strconsulta = "select * from mtl_secondary_inventories where secondary_inventory_name = ?"
                              With comandoORA
                                   .ActiveConnection = cnnoracle_4
                                   .CommandType = adCmdText
                                   .CommandText = strconsulta
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_almacen_tienda)
                                   .Parameters.Append parametro
                              End With
                              Set rsaux3 = comandoORA.execute
                              Set comandoORA = Nothing
                              Set parametro = Nothing
                           
                              If Not rsaux3.EOF Then
                                 var_location_id = IIf(IsNull(rsaux3!LOCATION_ID), 0, rsaux3!LOCATION_ID)
                                 If var_location_id > 0 Then
                                 
                                    strconsulta = "select ADDRESS_LINE_1, ADDRESS_LINE_2, TOWN_OR_CITY, REGION_1, COUNTRY, POSTAL_CODE  from hr_locations_all where location_id = ?"
                                    With comandoORA
                                         .ActiveConnection = cnnoracle_4
                                         .CommandType = adCmdText
                                         .CommandText = strconsulta
                                         Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, var_location_id)
                                         .Parameters.Append parametro
                                    End With
                                    Set rsaux4 = comandoORA.execute
                                    Set comandoORA = Nothing
                                    Set parametro = Nothing
                                    If Not rsaux4.EOF Then
                                       VAR_DIRECCION = Mid(IIf(IsNull(rsaux4!ADDRESS_LINE_1), "", rsaux4!ADDRESS_LINE_1), 1, 50)
                                       VAR_COLONIA = IIf(IsNull(rsaux4!ADDRESS_LINE_2), "", rsaux4!ADDRESS_LINE_2)
                                       var_ciudad = IIf(IsNull(rsaux4!TOWN_OR_CITY), "", rsaux4!TOWN_OR_CITY)
                                       var_estado = IIf(IsNull(rsaux4!REGION_1), "", rsaux4!REGION_1)
                                       var_pais = IIf(IsNull(rsaux4!COUNTRY), "", rsaux4!COUNTRY)
                                       VAR_CP = IIf(IsNull(rsaux4!POSTAL_CODE), "", rsaux4!POSTAL_CODE)
                                    End If
                                    rsaux4.Close
                                 End If
                              Else
                                 VAR_DIRECCION = ""
                                 VAR_COLONIA = ""
                                 var_ciudad = ""
                                 var_estado = ""
                                 var_pais = ""
                                 VAR_CP = ""
                              End If
                              rsaux3.Close
                           Else
                              VAR_DIRECCION = ""
                              VAR_COLONIA = ""
                              var_ciudad = ""
                              var_estado = ""
                              var_pais = ""
                              VAR_CP = ""
                           End If
                        Else
                     
                           strconsulta = "SELECT  hps.pArty_site_number as clave_cliente , HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.invoice_TO_ORG_ID AND oha.order_number = ? AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
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
                              strconsulta = "SELECT  hps.pArty_site_number as clave_cliente ,HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  city as ciudad, postal_code  as cp, state  as estado, province as municipio, county as colonia, country as pais, address2 as calle, address3 as numero, address4 as colonia_1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.ship_TO_ORG_ID AND oha.order_number = ? AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                              With comandoORA
                                   .ActiveConnection = cnnoracle_4
                                   .CommandType = adCmdText
                                   .CommandText = strconsulta
                                   Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux9!source_header_number)
                                   .Parameters.Append parametro
                              End With
                              Set rsaux5 = comandoORA.execute
                              Set comandoORA = Nothing
                              Set parametro = Nothing
                                    
                              If Not rsaux5.EOF Then
                                 var_clave_cliente = IIf(IsNull(rsaux5!clave_cliente), "", rsaux5!clave_cliente)
                                 VAR_DIRECCION = Mid(IIf(IsNull(rsaux5!calle), "", rsaux5!calle) + " " + IIf(IsNull(rsaux5!NUMERO), "", rsaux5!NUMERO), 1, 50)
                                 VAR_COLONIA = IIf(IsNull(rsaux5!colonia), "", rsaux5!colonia)
                                 var_ciudad = IIf(IsNull(rsaux5!ciudad), "", rsaux5!ciudad)
                                 VAR_MUNICIPIO = IIf(IsNull(rsaux5!municipio), "", rsaux5!municipio)
                                 var_estado = IIf(IsNull(rsaux5!estado), "", rsaux5!estado)
                                 var_pais = IIf(IsNull(rsaux5!pais), "", rsaux5!pais)
                                 VAR_CP = IIf(IsNull(rsaux5!cp), "", rsaux5!cp)
                                 VAR_DIRECCION = IIf(IsNull(rsaux5!customer_name), "", rsaux5!customer_name) + ", Dirección de entrega: " + VAR_DIRECCION
                                 rsaux5.Close
                              Else
                                 rsaux5.Close
                                 var_clave_cliente = IIf(IsNull(rsaux6!clave_cliente), "", rsaux6!clave_cliente)
                                 VAR_DIRECCION = Mid(IIf(IsNull(rsaux6!calle), "", rsaux6!calle) + " " + IIf(IsNull(rsaux6!NUMERO), "", rsaux6!NUMERO), 1, 50)
                                 VAR_COLONIA = IIf(IsNull(rsaux6!colonia), "", rsaux6!colonia)
                                 var_ciudad = IIf(IsNull(rsaux6!ciudad), "", rsaux6!ciudad)
                                 VAR_MUNICIPIO = IIf(IsNull(rsaux6!municipio), "", rsaux6!municipio)
                                 var_estado = IIf(IsNull(rsaux6!estado), "", rsaux6!estado)
                                 var_pais = IIf(IsNull(rsaux6!pais), "", rsaux6!pais)
                                 VAR_CP = IIf(IsNull(rsaux6!cp), "", rsaux6!cp)
                              End If
                           Else
                              VAR_DIRECCION = ""
                              VAR_COLONIA = ""
                              var_ciudad = ""
                              VAR_MUNICIPIO = ""
                              var_estado = ""
                              var_pais = ""
                              VAR_CP = ""
                           End If
                           rsaux6.Close
                        End If
                        rsaux8.Close
                     
                        var_direccion_str = VAR_DIRECCION + ", " + VAR_COLONIA + ", " + var_ciudad + ", " + VAR_MUNICIPIO + ", " + var_estado + ", " + var_pais + ", CP: " + VAR_CP
                     
                     
                     
                     
                        var_nombre_cliente = IIf(IsNull(rsaux9!Cliente), "", rsaux9!Cliente)
                        If var_nombre_cliente = "VIANNEY TEXTIL HOGAR SA DE CV" Then
                           var_nombre_cliente = IIf(IsNull(rsaux9!ESTABLECIMIENTO), var_nombre_cliente, rsaux9!ESTABLECIMIENTO)
                        End If
                     
                        strconsulta = "SELECT SOURCE_HEADER_NUMBER, TIPO_CAJA, COUNT(TIPO_CAJA) AS CANTIDAD FROM XXVIA_VW_CANTIDAD_BULTOS WHERE SOURCE_HEADER_NUMBER = ? and tipo_caja is not null GROUP BY SOURCE_HEADER_NUMBER, TIPO_CAJA"
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
                        rsaux7.Open "select * from TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES WHERE PEDIDO = " + CStr(rsaux9!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux7.EOF Then
                           var_orden = IIf(IsNull(rsaux7!orden_pedido), 0, rsaux7!orden_pedido)
                        Else
                           var_orden = 0
                        End If
                        rsaux7.Close
                        rs.Open "INSERT INTO TB_TEMP_ORACLE_REPORTE_CAJAS (INTE_TEM_CONSECUTIVO, EMBARQUE, CLIENTE, CANTIDAD, PEDIDO, unidad, sellos, FECHA_EMBARQUE, RUTA, direccion_entrega, CLAVE_CLIENTE, ORDEN_ENTREGA, USUARIO_CERRO) VALUES (" + CStr(var_consecutivo) + "," + CStr(rsaux9!inte_Emb_Embarque) + ",'" + Replace(var_nombre_cliente, "'", " ") + "'," + CStr(IIf(IsNull(rsaux9!cantidad), 0, rsaux9!cantidad)) + "," + CStr(rsaux9!source_header_number) + ",'" + var_transporte + "','" + var_cadena_sellos + "', '" + var_fecha_embarque + "','" + VAR_CADENA_RUTAS + "','" + Replace(var_direccion_str, "'", " ") + "','" + var_clave_cliente + "'," + CStr(var_orden) + ",'" + VAR_USUARIO_CERRO + "')", cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux8.EOF
                              If rsaux8!tipo_caja = "CAJA EXTRAGRANDE" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set CAJA_EXTRAGRANDE = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "CAJA GRANDE" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set CAJA_GRANDE = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "CAJA MEDIANA" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set CAJA_MEDIANA = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "CAJA CHICA" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set CAJA_CHICA = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "CAJA MINI/CATALOGO" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set CAJA_MINI_CATALOGO = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "CAJA  SOBRE-CAJA" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set CAJA_SOBRE_CAJA = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "CAJA CORTINERO" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set CAJA_CORTINERO = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "COSTAL GRANDE" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set COSTAL_GRANDE = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "COSTAL CHICO" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set COSTAL_CHICO = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "EMPLAYE CORTINEROS" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set EMPLAYE_CORTINEROS = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "PAQUETE BOLSA" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set PAQUETE_BOLSA = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "PAQUETE PUBLICIDAD" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set PAQUETE_PUBLICIDAD = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "OTROS" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set OTROS = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "OTROS MUEBLES" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set OTROS_MUEBLES = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "CAJA CHICA C/CATALOGO" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set CAJA_CHICA_CATALOGO = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "CAJA BIASI" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set CAJA_GRIS = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              If rsaux8!tipo_caja = "COSTAL EXTRAGRANDE" Then
                                 rsaux10.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set COSTAL_EXTRAGRANDE = " + CStr(rsaux8!cantidad) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux9!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux8.MoveNext
                        Wend
                        
                        'If CDbl(Me.txt_embarque) <> 140832 Then
                        '   strConsulta = "SELECT DISTINCT INTE_EMB_EMBARQUE, SOURCE_HEADER_NUMBER, INTE_PAQ_CAJA, CAJA_PEDIDO, TIPO_CAJA, SELLO, AUDITADA FROM XXVIA_TB_SALIDAS_CAJAS WHERE inte_emb_embarque = ? AND INTE_PAQ_CAJA > 0 and source_header_number = ? order by inte_paq_caja"
                        '   With comandoORA
                        '       .ActiveConnection = cnnoracle_4
                        '       .CommandType = adCmdText
                        '       .CommandText = strConsulta
                        '       Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_embarque))
                        '       .Parameters.Append parametro
                        '       Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rsaux9!SOURCE_HEADER_NUMBER)
                        '        .Parameters.Append parametro
                        '   End With
                        '   Set rsaux7 = comandoORA.execute
                        '   Set comandoORA = Nothing
                        '   Set parametro = Nothing
                        'Else
                           rsaux7.Open "SELECT DISTINCT INTE_EMB_EMBARQUE, SOURCE_HEADER_NUMBER, INTE_PAQ_CAJA, CAJA_PEDIDO, TIPO_CAJA, SELLO, AUDITADA FROM XXVIA_TB_SALIDAS_CAJAS WHERE inte_emb_embarque = " + Me.txt_embarque + " AND INTE_PAQ_CAJA > 0 and source_header_number = " + CStr(rsaux9!source_header_number) + " order by inte_paq_caja", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        'End If
                      
                        var_i = 0
                        While Not rsaux7.EOF
                              var_i = var_i + 1
                              rsaux7.MoveNext
                        Wend
                
                        VAR_Z = Round(var_i / 3, 2)
                        VAR_Y = Round(var_i / 3, 0)
                        var_x = VAR_Z - VAR_Y
                        If var_x < 0.5 Then
                           If var_x = 0 Then
                              VAR_Z = Round(var_i / 3, 0)
                           Else
                              VAR_Z = Round(Round(var_i / 3, 2) + 0.5, 0)
                           End If
                        Else
                           VAR_Z = Round(var_i / 3, 0)
                        End If
               
                        rsaux7.MoveFirst
                        'MsgBox rsaux7.RecordCount
                        var_j = 0
                        var_m = 1
                        While Not rsaux7.EOF
                              If var_j = VAR_Z Then
                                 var_j = 0
                                 var_m = var_m + 1
                              End If
                              var_j = var_j + 1
                              If Not rsaux7.EOF Then
                                 var_numero_caja = rsaux7!INTE_PAQ_CAJA
                                 If Len(Trim(Str(var_numero_caja))) = 1 Then
                                    var_referencia_caja = "00" + Trim(Str(var_numero_caja))
                                 End If
                                 If Len(Trim(Str(var_numero_caja))) = 2 Then
                                    var_referencia_caja = "0" + Trim(Str(var_numero_caja))
                                 End If
                                 If Len(Trim(Str(var_numero_caja))) = 3 Then
                                    var_referencia_caja = Trim(Str(var_numero_caja))
                                 End If
                                 If Len(Trim(Str(txt_embarque))) = 1 Then
                                    var_referencia_embarque = "00000" + Trim(Str(txt_embarque))
                                 End If
                                 If Len(Trim(Str(txt_embarque))) = 2 Then
                                    var_referencia_embarque = "0000" + Trim(Str(txt_embarque))
                                 End If
                                 If Len(Trim(Str(txt_embarque))) = 3 Then
                                    var_referencia_embarque = "000" + Trim(Str(txt_embarque))
                                 End If
                                 If Len(Trim(Str(txt_embarque))) = 4 Then
                                    var_referencia_embarque = "00" + Trim(Str(txt_embarque))
                                 End If
                                 If Len(Trim(Str(txt_embarque))) = 5 Then
                                    var_referencia_embarque = "0" + Trim(Str(txt_embarque))
                                 End If
                                 If Len(Trim(Str(txt_embarque))) = 6 Then
                                    var_referencia_embarque = Trim(Str(txt_embarque))
                                 End If
                                 var_contingencia = 1
                                 If var_contingencia = 1 Then
                                    VAR_ESTATUS = "S"
                                 Else
                                 rsaux6.Open "select * from tb_oracle_cajas_aduana where embarque = " + Me.txt_embarque + " and numero_caja = " + CStr(rsaux7!INTE_PAQ_CAJA), cnn, adOpenDynamic, adLockOptimistic
                                 If Not rsaux6.EOF Then
                                    VAR_ESTATUS = IIf(IsNull(rsaux6!estatus), "", rsaux6!estatus)
                                 Else
                                    VAR_ESTATUS = ""
                                 End If
                                 rsaux6.Close
                                 End If
                                 If var_m = 1 Then
                                    var_cadena = "insert into TB_TEMP_ORACLE_CAJAS_ADUANA_DIVIDIDAS_EN_3 (inte_tem_consecutivo, renglon, pedido, caja_" + CStr(var_m) + ",codigo_" + CStr(var_m) + ", tipo_" + CStr(var_m) + ", sello_" + CStr(var_m) + ", auditada_" + CStr(var_m) + ",estatus_" + CStr(var_m) + ")"
                                    var_cadena = var_cadena + " values (" + CStr(var_consecutivo) + "," + CStr(var_j) + "," + CStr(rsaux7!source_header_number) + "," + CStr(rsaux7!caja_pedido) + ",'C" + var_referencia_embarque + var_referencia_caja + "','" + rsaux7!tipo_caja + "','" + IIf(IsNull(rsaux7!sello), "", rsaux7!sello) + "'," + CStr(IIf(IsNull(rsaux7!auditada), 0, rsaux7!auditada)) + ",'" + VAR_ESTATUS + "')"
                                 Else
                                    var_cadena = "update TB_TEMP_ORACLE_CAJAS_ADUANA_DIVIDIDAS_EN_3 set caja_" + CStr(var_m) + " = " + CStr(rsaux7!caja_pedido) + ", codigo_" + CStr(var_m) + " = 'C" + var_referencia_embarque + var_referencia_caja + "',tipo_" + CStr(var_m) + " = '" + rsaux7!tipo_caja + "', sello_" + CStr(var_m) + "= '" + IIf(IsNull(rsaux7!sello), "", rsaux7!sello) + "', auditada_" + CStr(var_m) + "  = " + CStr(IIf(IsNull(rsaux7!auditada), 0, rsaux7!auditada)) + ", estatus_" + CStr(var_m) + " = '" + VAR_ESTATUS + "'  where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and renglon = " + CStr(var_j) + " and pedido = " + CStr(rsaux7!source_header_number)
                                 End If
                                 'MsgBox var_cadena
                                 rsaux6.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux7.MoveNext
                        Wend
                        rsaux7.Close
                        rsaux9.MoveNext
                  Wend
                  
                        strconsulta = "SELECT INTE_EMB_EMBARQUE, COUNT(*) AS VECES FROM XXVIA_VW_BULTOS_AUDITADOS WHERE INTE_EMB_EMBARQUE = ? GROUP BY INTE_eMB_EMBARQUE"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_embarque))
                             .Parameters.Append parametro
                        End With
                        Set rsaux8 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        If Not rsaux8.EOF Then
                           var_total_bultos_embarques = IIf(IsNull(rsaux8!VECES), 0, rsaux8!VECES)
                        Else
                           var_total_bultos_embarques = 0
                        End If
                  
                        rsaux8.Close
                        rsaux8.Open "update TB_TEMP_ORACLE_REPORTE_CAJAS set TRANSPORTISTA = '" + VAR_TRANSPORTISTA + "', SELLO_BARRIL = '" + VAR_SELLO_BARRIL + "', SELLO_LAMINAS = '" + VAR_SELLO_LAMINA + "', SELLO_LATERALES = '" + VAR_sELLO_LATERALES + "', CERTIFICA_ADUANAL = '" + VAR_CERTIFICA_ADUANAL + "', DETALLE_CONTENIDO = '" + VAR_DETALLE_CONTENIDO + "', TOTAL_BULTOS_AUDITODOS = " + CStr(var_total_bultos_embarques) + ", COLOR_ETIQUETA = '" + VAR_COLOR_eTIQUETA + "', PALET = '" + VAR_PALET + "' where embarque = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
                        
                  rs.Open "DELETE FROM TB_TEMP_ORACLE_REPORTE_CAJAS WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo) + " and ruta is null", cnn, adOpenDynamic, adLockOptimistic
                  Set reporte = appl.OpenReport(App.Path + "\rep_oracle_cajas_en_embarque_EXPORTACIONES.rpt")
                  reporte.RecordSelectionFormula = "{VW_ORACLE_REPORTE_CAJAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Pedidos cargados"
                  frmvistasprevias.Show 1
                  Set reporte = Nothing
                  rs.Open "DELETE FROM TB_TEMP_ORACLE_REPORTE_CAJAS WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
               Else
                 MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
               End If
               rsaux9.Close
               rsaux11.Close
            Else
               MsgBox "El embarque aun no a sido cerrado", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
         End If
         rsaux14.Close
         End If
End If
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_inicio) Then
         frmcalendario.mes = CDate(Me.txt_fin)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      Me.txt_fin = var_fecha_general
   End If
End Sub

Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_inicio) Then
         frmcalendario.mes = CDate(Me.txt_inicio)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      txt_inicio = var_fecha_general
   End If
End Sub
