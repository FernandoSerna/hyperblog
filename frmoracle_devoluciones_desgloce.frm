VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_devoluciones_desgloce 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Desgloce de devoluciones"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   330
      Left            =   540
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8025
      Picture         =   "frmoracle_devoluciones_desgloce.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   300
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmoracle_devoluciones_desgloce.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   60
      TabIndex        =   2
      Top             =   270
      Width           =   8280
   End
   Begin VB.Frame Frame2 
      Height          =   6900
      Left            =   75
      TabIndex        =   0
      Top             =   345
      Width           =   8235
      Begin VB.Frame frm_mensaje 
         Height          =   1215
         Left            =   570
         TabIndex        =   14
         Top             =   3645
         Width           =   7065
         Begin VB.Label lbl_mensaje 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            Caption         =   "estatus"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1200
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   7050
         End
      End
      Begin VB.Frame frm_lista 
         Height          =   2400
         Left            =   1185
         TabIndex        =   10
         Top             =   900
         Width           =   5970
         Begin MSComctlLib.ListView lv_lista 
            Height          =   1950
            Left            =   60
            TabIndex        =   11
            Top             =   405
            Width           =   5865
            _ExtentX        =   10345
            _ExtentY        =   3440
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Clave"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nombre"
               Object.Width           =   7937
            EndProperty
         End
         Begin VB.Label lbl_lista 
            BackColor       =   &H000000C0&
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   30
            TabIndex        =   12
            Top             =   120
            Width           =   5895
         End
      End
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1365
         Picture         =   "frmoracle_devoluciones_desgloce.frx":073C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   705
         Picture         =   "frmoracle_devoluciones_desgloce.frx":0952
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Marcar (Enter)"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1035
         Picture         =   "frmoracle_devoluciones_desgloce.frx":0B9C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   45
         Picture         =   "frmoracle_devoluciones_desgloce.frx":0C6E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   375
         Picture         =   "frmoracle_devoluciones_desgloce.frx":0D70
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   135
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_devoluciones 
         Height          =   6315
         Left            =   45
         TabIndex        =   1
         Top             =   525
         Width           =   8115
         _ExtentX        =   14314
         _ExtentY        =   11139
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
            Text            =   "Código"
            Object.Width           =   2478
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Causa"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Clave causa"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Consecutivo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "inventory_item_id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Localizador"
            Object.Width           =   0
         EndProperty
      End
   End
End
Attribute VB_Name = "frmoracle_devoluciones_desgloce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_guardar_Click()
   Dim clnt As New SoapClient30
   Dim objConn As New ADODB.Connection
   Dim objCmd As New ADODB.Command
   Dim objParm As ADODB.Parameter
   Dim var_header_interface_id  As Double
   Dim var_group_id As Double
   Dim var_clave_tipo_pedido As Integer
   Dim var_concurrente As Double
   'MsgBox var_clave_movimiento
   var_posible = 1
   For var_j = 1 To Me.lv_devoluciones.ListItems.Count
       Me.lv_devoluciones.ListItems(var_j).Selected = True
       If Trim(Me.lv_devoluciones.selectedItem.SubItems(3)) = "" Then
          var_posible = 0
       End If
   Next var_j
   If var_posible = 1 Then
      var_si = MsgBox("¿Desea cerrar el movimiento", vbYesNo, "ATENCION")
      If var_si = 6 Then
         
         If rs.State = 1 Then
            rs.Close
         End If
         If rsaux9.State = 1 Then
            rsaux9.Close
         End If
         rsaux9.Open "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux9.EOF Then
            var_clave_tipo_pedido = IIf(IsNull(rsaux9!inte_int_interface), 0, rsaux9!inte_int_interface)
         Else
            var_clave_tipo_pedido = 0
         End If
         'var_clave_tipo_pedido = 1048
         rsaux9.Close
         If rsaux7.State = 1 Then
            rsaux7.Close
         End If
         If var_clave_movimiento = "SNC" Then
            If var_unidad_organizacional = 93 Then
               var_clave_tipo_pedido = 1941
            End If
            If var_unidad_organizacional = 94 Then
               var_clave_tipo_pedido = 1983
            End If
         End If
         'var_clave_tipo_pedido = 2181
         'var_clave_lista_precios = 9007
         rsaux7.Open "select name from qp_secu_list_headers_v where list_header_id = " + CStr(var_clave_lista_precios), cnnoracle_4, adOpenDynamic, adLockOptimistic
         var_lista_precios = rsaux7(0).Value
         rsaux7.Close

         If var_clave_tipo_pedido > 0 Then
            If var_lista_precios <> "" Then
               If rs.State = 1 Then
                  rs.Close
               End If
               'var_cadena = "SELECT A.FACTURA, A.TIPO_PEDIDO, A.MOVIMIENTO, b.estatus, A.NUMERO, A.ORGANIZACION, A.inventory_item_id, A.almacen, A.titular, A.unidad_medida, A.precio, A.TITULAR, A.CLIENTE, A.ESTABLECIMIENTO, B.CAUSA_DEVOLUCION, b.descripcion_causa, b.localizador, SUM(b.cantidad) AS CANTIDAD FROM XXVIA_TB_DEVOLUCIONES_CLIENTES A, xxvia_tb_dev_clientes_desgloce B WHERE A.numero = b.numero AND a.organizacion = b.organizacion AND a.inventory_item_id = b.inventory_item_id AND A.NUMERO = " + CStr(var_numero_folio_devoluciones) + " AND "
               'var_cadena = var_cadena + " A.ORGANIZACION = " + var_unidad_organizacional + "  AND A.LOCALIZADOR = B.LOCALIZADOR AND A.MOVIMIENTO = B.MOVIMIENTO AND A.MOVIMIENTO = '" + var_clave_movimiento + "' GROUP BY A.FACTURA, A.TIPO_PEDIDO, A.MOVIMIENTO, b.estatus, A.NUMERO, A.ORGANIZACION, A.inventory_item_id, A.almacen, A.titular, A.unidad_medida, A.precio, A.TITULAR, A.CLIENTE, A.ESTABLECIMIENTO, B.CAUSA_DEVOLUCION, B.causa_devolucion, b.descripcion_causa, b.localizador"
               var_cadena = "SELECT a.agente, A.FACTURA, A.TIPO_PEDIDO, A.MOVIMIENTO, b.estatus, A.NUMERO, A.ORGANIZACION, A.inventory_item_id, A.almacen, A.titular, A.unidad_medida, A.precio, A.TITULAR, A.CLIENTE, A.ESTABLECIMIENTO, B.CAUSA_DEVOLUCION, b.descripcion_causa, b.localizador, SUM(b.cantidad) AS CANTIDAD FROM XXVIA_TB_DEVOLUCIONES_CLIENTES A, xxvia_tb_dev_clientes_desgloce B WHERE A.numero = b.numero AND a.organizacion = b.organizacion AND a.inventory_item_id = b.inventory_item_id AND A.NUMERO = " + CStr(var_numero_folio_devoluciones) + " AND "
               var_cadena = var_cadena + " A.ORGANIZACION = " + var_unidad_organizacional + " AND A.MOVIMIENTO = B.MOVIMIENTO AND A.MOVIMIENTO = '" + var_clave_movimiento + "' GROUP BY a.agente, A.FACTURA, A.TIPO_PEDIDO, A.MOVIMIENTO, b.estatus, A.NUMERO, A.ORGANIZACION, A.inventory_item_id, A.almacen, A.titular, A.unidad_medida, A.precio, A.TITULAR, A.CLIENTE, A.ESTABLECIMIENTO, B.CAUSA_DEVOLUCION, B.causa_devolucion, b.descripcion_causa, b.localizador"
               rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If IIf(IsNull(rs!estatus), "", rs!estatus) = "" Then
                  If Not rs.EOF Then
                     var_clave_tipo_pedido = rs!tipo_pedido
                     If var_unidad_organizacional = "93" Then
                         If var_clave_tipo_pedido <> 1048 Then
                            var_clave_tipo_pedido = 1101
                         End If
                     End If
                     If var_unidad_organizacional = "90" Then
                        If rs!Agente = 1016 Then
                           var_clave_tipo_pedido = 1069
                        Else
                           var_clave_tipo_pedido = 1105
                        End If
                     End If
                     If var_unidad_organizacional = "89" Then
                        var_clave_tipo_pedido = 1078
                     End If
                     If var_unidad_organizacional = "85" Then
                        If rs!Agente = 1028 Then
                           var_clave_tipo_pedido = 1102
                        Else
                           var_clave_tipo_pedido = 1050
                        End If
                     End If
                     If var_clave_usuario_global = "U0000000430" Then
                        var_clave_tipo_pedido = 1381
                     End If
                     If var_clave_movimiento = "SNC" Then
                        If var_unidad_organizacional = 93 Then
                           var_clave_tipo_pedido = 1941
                        End If
                        If var_unidad_organizacional = 90 Then
                           var_clave_tipo_pedido = 1983
                        End If
                     End If
                     
                     'var_clave_tipo_pedido = 1051
                     'var_clave_tipo_pedido = 1841
                     rs.MoveFirst
                     var_referencia = ""
                     rsaux1.Open "select * from xxvia_tb_Devoluciones_clientes where numero = " + Trim(CStr((var_numero_folio_devoluciones))) + " and rownum = 1 AND MOVIMIENTO = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux1.EOF Then
                        var_referencia = Mid(IIf(IsNull(rsaux1!Referencia), "", rsaux1!Referencia), 1, 15)
                     End If
                     rsaux1.Close
                     If IIf(IsNull(rs!estatus), "", rs!estatus) = "" Then
                     Dim var_cliente As Double
                     Dim var_establecimiento As Double
                     var_cliente = rs!Cliente
                     If var_cliente = 7086 Then
                        var_cliente = 7080
                     End If
                     If var_cliente = 4223960 Then
                        var_cliente = 1042769
                     End If
                     If var_cliente = 34190 Then
                        var_cliente = 612991
                     End If
                     If var_cliente = 984967 Then
                        var_cliente = 984967
                     End If
                     If var_cliente = 1070229 Then
                        var_cliente = 8650
                     End If
                     var_establecimiento = rs!ESTABLECIMIENTO
                     If var_establecimiento = 610313 Then
                        var_establecimiento = 1082173
                     End If
                     
                     If var_establecimiento = 998699 Or var_establecimiento = 1012556 Then
                        var_establecimiento = 8911
                     End If
                     
                     If var_establecimiento = 307959 Then
                        var_establecimiento = 770480
                     End If
                     If var_establecimiento = 1010054 Then
                        var_establecimiento = 8155
                     End If
                     If var_establecimiento = 7514 Then
                        var_establecimiento = 7512
                     End If
                     If var_establecimiento = 7423 Then
                       var_establecimiento = 1004076
                     End If
                     If var_establecimiento = 1052955 Then
                       var_establecimiento = 8189
                     End If
                     If var_cliente = 312914 Then
                        var_cliente = 984967
                     End If
                     If var_establecimiento = 7572 Then
                        var_establecimiento = 7573
                     End If
                     If var_establecimiento = 34191 Then
                        var_establecimiento = 1005313
                     End If
                     
                     If var_cliente = 34192 Then
                        var_cliente = 612991
                     End If
                     
                     If var_cliente = 34192 Then
                        var_cliente = 612991
                     End If
                     
                     If var_cliente = 397647 Then
                        var_cliente = 4187415
                     End If
                     
                     If var_establecimiento = 1038164 Then
                        var_establecimiento = 4187416
                     End If
                     
                     
                     
                     If var_establecimiento = 887492 Then
                        var_establecimiento = 7568
                     End If
                     
                     If var_establecimiento = 83521 Then
                        var_establecimiento = 795304
                     End If
                     If var_establecimiento = 1005193 Or var_establecimiento = 99 Then
                        var_establecimiento = 8911
                     End If
                     If var_establecimiento = 5220215 Then
                        var_establecimiento = 1037555
                     End If
                     
                     If var_establecimiento = 884546 Then
                        var_establecimiento = 5222656
                     End If
   
                     
                     
                     var_cadena = "INSERT INTO oe_headers_iface_all (ORDER_SOURCE_ID, orig_sys_document_ref, creation_date, created_by, last_update_date, last_updated_by, operation_code , sold_to_org_id        , SHIP_TO_ORG_id                   ,INVOICE_TO_ORG_ID     , Order_type_ID, PRICE_LIST, SHIP_FROM_ORG_ID, attribute7)"
                     var_cadena = var_cadena + "  VALUES (1001,'SID" + var_clave_movimiento + "_" + Trim(CStr((var_numero_folio_devoluciones))) + "',SYSDATE,-1,SYSDATE, -1,'INSERT', " + CStr(rs!TITULAR) + "," + CStr(var_establecimiento) + "," + CStr(var_cliente) + "," + CStr(var_clave_tipo_pedido) + ",'" + var_lista_precios + "'," + var_unidad_organizacional + ",'" + var_referencia + "')"
                     rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     var_i = 0
                     While Not rs.EOF
                           If rs!FACTURA = 0 Then
                              VAR_FACTURA_VALUADA = ""
                           Else
                              rsaux3.Open "select * from RA_CUSTOMER_TRX_ALL where customer_trx_id = " + CStr(IIf(IsNull(rs!FACTURA), 0, rs!FACTURA)), cnnoracle_4
                              If Not rsaux3.EOF Then
                                 VAR_FACTURA_VALUADA = IIf(IsNull(rsaux3!trx_number), "", rsaux3!trx_number)
                              Else
                                 VAR_FACTURA_VALUADA = ""
                              End If
                              rsaux3.Close
                           End If
                           var_i = var_i + 1
                           
                           rsaux10.Open "SELECT PRIMARY_UOM_CODE, primary_unit_of_measure FROM xxvia_system_items_b WHERE INVENTORY_ITEM_ID = " + CStr(rs!inventory_item_id) + " AND ORGANIZATION_ID = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If Not rsaux10.EOF Then
                              VAR_MEDIDA = rsaux10(0).Value
                              VAR_NOMBRE_MEDIDA = rsaux10(1).Value
                           End If
                           rsaux10.Close
                           
                           
                           var_cadena = "INSERT INTO oe_lines_iface_all (ORDER_SOURCE_ID, orig_sys_document_ref,orig_sys_line_ref,inventory_item_id,ordered_quantity, operation_code, created_by, creation_date, last_updated_by, last_update_date, unit_selling_price, unit_list_price, calculate_price_flag, RETURN_REASON_CODE, PRICING_QUANTITY, PRICING_QUANTITY_UOM, ATTRIBUTE1, ATTRIBUTE11, SHIP_FROM_ORG_ID)"
                           If IIf(IsNull(rs!Precio), 0, rs!Precio) = 0 Then
                              var_cadena = var_cadena + " VALUES (1001,'SID" + var_clave_movimiento + "_" + Trim(CStr(var_numero_folio_devoluciones)) + "','" + CStr(var_i) + "', " + CStr(rs!inventory_item_id) + ", -" + CStr(Round(rs!cantidad, 2)) + ",'INSERT', -1,SYSDATE, -1,SYSDATE," + CStr(rs!Precio) + "," + CStr(rs!Precio) + ",'Y', " + rs!causa_devolucion + ",-" + CStr(rs!cantidad) + ", '" + VAR_MEDIDA + "','" + IIf(IsNull(rs!localizador), "", rs!localizador) + "','" + VAR_FACTURA_VALUADA + "'," + var_unidad_organizacional + ")"
                           Else
                              var_cadena = var_cadena + " VALUES (1001,'SID" + var_clave_movimiento + "_" + Trim(CStr(var_numero_folio_devoluciones)) + "','" + CStr(var_i) + "', " + CStr(rs!inventory_item_id) + ", -" + CStr(Round(rs!cantidad, 2)) + ",'INSERT', -1,SYSDATE, -1,SYSDATE," + CStr(rs!Precio) + "," + CStr(rs!Precio) + ",'N', " + rs!causa_devolucion + ",-" + CStr(rs!cantidad) + ", '" + VAR_MEDIDA + "','" + IIf(IsNull(rs!localizador), "", rs!localizador) + "','" + VAR_FACTURA_VALUADA + "'," + var_unidad_organizacional + ")"
                           End If
                           rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                           rsaux.Open "insert into paso (consecutivo, codigo) values (" + CStr(var_i) + "," + CStr(rs!inventory_item_id) + ")", cnn, adOpenDynamic, adLockOptimistic
                           rs.MoveNext
                     Wend
                     On Error GoTo SALIR
                     rsaux.Open "INSERT INTO oe_actions_iface_all (order_source_ID, orig_sys_document_ref, operation_code) VALUES (1001, 'SID" + var_clave_movimiento + "_" + Trim(Trim(CStr(var_numero_folio_devoluciones))) + "','BOOK_ORDER')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     Me.frm_mensaje.Visible = True
                     Me.lbl_mensaje.Caption = "GENERANDO PEDIDO"
                     Me.Refresh
                     rsaux.Open "CALL XXVIA_PK_INTERFACES_OM.importar_pedido('SID" + var_clave_movimiento + "_" + Trim(Trim(CStr(var_numero_folio_devoluciones))) + "'," + var_empresa + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     End If
                     If rsaux.State = 1 Then
                        rsaux.Close
                     End If
                     
                     strconsulta = "SELECT * FROM XXVIA_TB_dEVOLUCIONES_CLIENTES where movimiento = ? and numero = ? and nvl(factura,0) > 0"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, var_clave_movimiento)
                          .Parameters.Append parametro
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, var_numero_folio_devoluciones)
                          .Parameters.Append parametro
                     End With
                     Set rsaux8 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     While Not rsaux8.EOF
                           strconsulta = "select nvl(to_char(cadena_original),' ') as uuid from xxvia_Tb_control_doc_fiscales where nvl(to_char(cadena_original),' ') <> ' ' and customer_trx_id = ?"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, rsaux8!FACTURA)
                                .Parameters.Append parametro
                           End With
                           Set rsaux9 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           If Not rsaux9.EOF Then
                              strconsulta = "update XXVIA_TB_dEVOLUCIONES_CLIENTES set uuid = ? where movimiento = ? and numero = ? and codigo = ?"
                              With comandoORA
                                   .ActiveConnection = cnnoracle_4
                                   .CommandType = adCmdText
                                   .CommandText = strconsulta
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 300, rsaux9!uuid)
                                   .Parameters.Append parametro
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, var_clave_movimiento)
                                   .Parameters.Append parametro
                                   Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, var_numero_folio_devoluciones)
                                   .Parameters.Append parametro
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, rsaux8!codigo)
                                   .Parameters.Append parametro
                              End With
                              Set rsaux10 = comandoORA.execute
                              Set comandoORA = Nothing
                              Set parametro = Nothing
                           Else
                              strconsulta = "update XXVIA_TB_dEVOLUCIONES_CLIENTES set uuid = ? where movimiento = ? and numero = ? and codigo = ?"
                              With comandoORA
                                   .ActiveConnection = cnnoracle_4
                                   .CommandType = adCmdText
                                   .CommandText = strconsulta
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 300, " ")
                                   .Parameters.Append parametro
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, var_clave_movimiento)
                                   .Parameters.Append parametro
                                   Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, var_numero_folio_devoluciones)
                                   .Parameters.Append parametro
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, rsaux8!codigo)
                                   .Parameters.Append parametro
                              End With
                              Set rsaux10 = comandoORA.execute
                              Set comandoORA = Nothing
                              Set parametro = Nothing
                           End If
                           rsaux9.Close
                           rsaux8.MoveNext
                     Wend
                     rsaux8.Close
                     
                     rsaux.Open "select * from oe_order_headers_all where orig_sys_document_ref = 'SID" + var_clave_movimiento + "_" + Trim(CStr(var_numero_folio_devoluciones)) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     var_nUMERO_pedido = rsaux!order_number
                     
                           
                           
                           
                           rsaux9.Open "SELECT * FROM TB_ORACLE_PEDIDOS_CERRADOS WHERE PEDIDO = " + CStr(var_nUMERO_pedido), cnn, adOpenDynamic, adLockOptimistic
                           If rsaux9.EOF Then
                              rsaux10.Open "INSERT INTO TB_ORACLE_PEDIDOS_CERRADOS (PEDIDO, REQUEST_ID) VALUES (" + CStr(var_nUMERO_pedido) + ",0)", cnn, adOpenDynamic, adLockOptimistic
                           End If
                           rsaux9.Close
                     
                     
                     If var_clave_movimiento <> "SNC" Then
                     var_cadena = "SELECT hp.party_type,hca.account_number numero_titular, hp.party_name titular, hp.party_id, hcas.org_id, hcas.cust_acct_site_id, hcsu.site_use_id, hr.NAME unidad_operativa, hca.customer_class_code clasificacion_titular, hca.cust_account_id, decode(hcsu.site_use_code, 'BILL_TO', 'FACTURACION','SHIP_TO', 'ENVIO') proposito, hcsu.LOCATION numero_cliente, hcas.orig_system_reference referencia_bancaria, hl.address1 cliente, hl.address3 numero_externo, hl.address4 numero_interno, hl.address2 calle, hl.city ciudad, hl.county colionia, hl.province delegacion_municipio, hl.state estado, hl.postal_code codigo_postal, hcsu.attribute1 rfc, hcsu.attribute2 curp, hcsu.attribute3 prioridad_envio, hcp.cust_account_profile_id profile_id, hcp.collector_id, arc.NAME agente, arcpc.NAME nombre_perfil, hcsu.price_list_id id_lista_precio, plist.NAME nombre_lista_precio, email.email_address email, phone.phone_number phone, hcas.party_site_id, NVL(hcsu.TERRITORY_ID,0) TERRITORY_ID, hl.location_id "
                     var_cadena = var_cadena + " FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc, qp_secu_list_headers_v plist, hz_contact_points email, hz_contact_points phone Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id AND hcsu.price_list_id = plist.list_header_id(+) AND email.contact_point_type(+) = 'EMAIL' AND email.owner_table_name(+) = 'HZ_PARTY_SITES' "
                     var_cadena = var_cadena + " AND email.owner_table_id(+) = hcas.party_site_id AND phone.contact_point_type(+) = 'PHONE' AND phone.owner_table_name(+) = 'HZ_PARTY_SITES' AND phone.owner_table_id(+) = hcas.party_site_id and hca.cust_account_id = " + CStr(rsaux!SOLD_TO_ORG_ID) + " and hcsu.site_use_id = " + CStr(rsaux!INVOICE_TO_ORG_ID) + " and hcas.org_id = " + var_empresa
                        If rsaux7.State = 1 Then
                           rsaux7.Close
                        End If
                        rsaux7.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                       
                       
                        If Not rsaux.EOF Then
                           If rsaux1.State = 1 Then
                              rsaux1.Close
                           End If
                           rsaux1.Open "select almacen from xxvia_tb_devoluciones_clientes where numero = " + Trim(CStr(var_numero_folio_devoluciones)) + " and movimiento = '" + var_clave_movimiento + "' and almacen is not null and rownum = 1 and organizacion = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
                           var_clave_almacen_devolucion = IIf(IsNull(rsaux1(0).Value), "", rsaux1(0).Value)
                           rsaux1.Close
                           rsaux1.Open "SELECT * FROM OE_ORDER_LINES_ALL WHERE HEADER_ID = " + CStr(rsaux!header_id), cnnoracle_4, adOpenDynamic, adLockOptimistic
                           
                           objConn.Open var_conexion_oracle
                           '… Establecer conexión a la base de datos con el objeto objConn.
                           With objCmd
                                objConn.BeginTrans
                                .ActiveConnection = objConn
                                .CommandText = "xxvia_pk_interfaces_om.generar_encabezado_devolucion"
                                .CommandType = adCmdStoredProc
                                          
                                'p_organization_id IN NUMBER
                                Set objParm = .CreateParameter("p_organization_id", adNumeric, adParamInput, , CDbl(var_unidad_organizacional))
                                .Parameters.Append objParm
                                
                                'MsgBox rsaux1!sold_to_org_id
                                'p_customer_id IN number,
                                Set objParm = .CreateParameter("p_customer_id", adNumeric, adParamInput, , IIf(IsNull(rsaux1!SOLD_TO_ORG_ID), 0, rsaux1!SOLD_TO_ORG_ID))
                                .Parameters.Append objParm
               
                                'p_devolucion_sid IN VARCHAR2,
                                Set objParm = .CreateParameter("p_devolucion_sid", adVarChar, adParamInput, 50, "SID" + var_clave_movimiento + "_" + Trim(var_numero_folio_devoluciones))
                                .Parameters.Append objParm
                                
                                ' x_header_interface_id out number,
                                Set objParm = .CreateParameter("x_header_interface_id", adVarNumeric, adParamOutput, 50, var_header_interface_id)
                                .Parameters.Append objParm
                                  
                                'x_group_id IN VARCHAR,
                                Set objParm = .CreateParameter("x_group_id", adVarNumeric, adParamOutput, 50, var_group_id)
                                .Parameters.Append objParm
                                .execute
                                var_header_interface_id = .Parameters("x_header_interface_id").Value
                                var_group_id = .Parameters("x_group_id").Value
                                objConn.CommitTrans
                           End With
                           'MsgBox var_conexion_oracle
                           Set objConn = Nothing
                           Set objCmd = Nothing
                           var_j = 0
                           While Not rsaux1.EOF
                                 objConn.Open var_conexion_oracle
                                 '… Establecer conexión a la base de datos con el objeto objConn.
                                 With objCmd
                                      objConn.BeginTrans
                                      .ActiveConnection = objConn
                                      .CommandText = "xxvia_pk_interfaces_om.generar_linea_devolucion"
                                      .CommandType = adCmdStoredProc
                                           
                                      'p_header_interface_id IN NUMBER
                                      Set objParm = .CreateParameter("p_header_interface_id", adNumeric, adParamInput, , var_header_interface_id)
                                      .Parameters.Append objParm
                              
                                      'p_group_id IN number,
                                      Set objParm = .CreateParameter("p_group_id", adNumeric, adParamInput, , var_group_id)
                                      .Parameters.Append objParm
                          
                                      'p_quantity IN number,
                                      Set objParm = .CreateParameter("p_quantity", adNumeric, adParamInput, , rsaux1!ORDERED_QUANTITY)
                                      .Parameters.Append objParm
                  
                                      'p_transaction_date IN VARCHAR2,
                                      'Set objParm = .CreateParameter("p_transaction_date", adDate, adParamInput, , Date)
                                      '.Parameters.Append objParm
                             
                                      rsaux2.Open "select primary_uom_code, primary_unit_of_measure from xxvia_system_items_b where organization_id = " + var_unidad_organizacional + " and inventory_item_id = " + CStr(rsaux1!inventory_item_id), cnnoracle_4, adOpenDynamic, adLockOptimistic
                             
                                      'p_unit_of_measure IN VARCHAR2,
                                      Set objParm = .CreateParameter("p_unit_of_measure", adVarChar, adParamInput, 50, rsaux2!PRIMARY_UNIT_OF_MEASURE)
                                      .Parameters.Append objParm
                           
                                      'p_uom_code IN VARCHAR2,
                                      Set objParm = .CreateParameter("p_uom_code", adVarChar, adParamInput, 50, rsaux2!PRIMARY_UOM_CODE)
                                      .Parameters.Append objParm
                             
                                      rsaux2.Close
                             
                                      'p_expected_receipt_date IN VARCHAR2,
                                      'Set objParm = .CreateParameter("p_expected_receipt_date", adDate, adParamInput, , Date)
                                      '.Parameters.Append objParm
                             
                                      'p_employee_id IN number,
                                      Set objParm = .CreateParameter("p_employee_id", adNumeric, adParamInput, , -1)
                                      .Parameters.Append objParm
                               
                                      'p_item_id IN number,
                                      Set objParm = .CreateParameter("p_item_id", adNumeric, adParamInput, , rsaux1!inventory_item_id)
                                      .Parameters.Append objParm
                                  
                                      'p_ship_to_location_id IN number,
                                      Set objParm = .CreateParameter("p_ship_to_location_id", adNumeric, adParamInput, , rsaux7!LOCATION_ID)
                                      .Parameters.Append objParm
                               
                                      'p_to_organization_id IN number,
                                      Set objParm = .CreateParameter("p_to_organization_id", adNumeric, adParamInput, , CDbl(var_unidad_organizacional))
                                      .Parameters.Append objParm
                                  
                                      'p_oe_order_header_id IN number,
                                      Set objParm = .CreateParameter("p_oe_order_header_id", adNumeric, adParamInput, , rsaux1!header_id)
                                      .Parameters.Append objParm
                                  
                                      'p_oe_order_line_id IN number,
                                      Set objParm = .CreateParameter("p_oe_order_line_id", adNumeric, adParamInput, , rsaux1!line_id)
                                      .Parameters.Append objParm
                               
                                      'p_customer_id IN number,
                                      Set objParm = .CreateParameter("p_customer_id", adNumeric, adParamInput, , rsaux!SOLD_TO_ORG_ID)
                                      .Parameters.Append objParm
                               
                                      'p_customer_site_id IN number,
                                      Set objParm = .CreateParameter("p_customer_site_id", adNumeric, adParamInput, , rsaux!INVOICE_TO_ORG_ID)
                                      .Parameters.Append objParm
                                     
                                      'p_subinventory IN VARCHAR2,
                                      
                                      Set objParm = .CreateParameter("p_subinventory", adVarChar, adParamInput, 50, var_clave_almacen_devolucion)
                                      .Parameters.Append objParm
                             
                                      'p_org_id IN number,
                                      Set objParm = .CreateParameter("p_org_id", adNumeric, adParamInput, , CDbl(var_empresa))
                                      .Parameters.Append objParm
                                       
                                      rsaux2.Open "select * from mtl_item_locations where segment1 = '" + IIf(IsNull(rsaux1!attribute1), "", rsaux1!attribute1) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                      If Not rsaux2.EOF Then
                                         Set objParm = .CreateParameter("p_location", adNumeric, adParamInput, , IIf(IsNull(rsaux2!inventory_location_id), 0, rsaux2!inventory_location_id))
                                      Else
                                         Set objParm = .CreateParameter("p_location", adNumeric, adParamInput, , Null)
                                      End If
                                      rsaux2.Close
                                      .Parameters.Append objParm
                                      
                                      Set objParm = .CreateParameter("p_factura_val", adVarChar, adParamInput, 50, IIf(IsNull(rsaux1!attribute11), "", rsaux1!attribute11))
                                      .Parameters.Append objParm
                                      
                                      
                                      
                                      .execute
                                      objConn.CommitTrans
                                 End With
                                 Set objConn = Nothing
                                 Set objCmd = Nothing
                                 rsaux1.MoveNext
                                 var_j = var_j + 1
                                 If var_j = 250 Or rsaux1.EOF Then
                                    var_concurrente = 0
                                    objConn.Open var_conexion_oracle
                                    With objCmd
                                         objConn.BeginTrans
                                         .ActiveConnection = objConn
                                         .CommandText = "XXVIA_PK_INVENTARIOS.XXVIA_SP_CONCURRENTE_MAT"
                                         .CommandType = adCmdStoredProc
                                    
                                         Set objParm = .CreateParameter("x_concurrente", adNumeric, adParamOutput, 50, var_concurrente)
                                         .Parameters.Append objParm
                             
                                         Set objParm = .CreateParameter("p_tipo_movimiento", adVarChar, adParamInput, 200, "Traspasos")
                                         .Parameters.Append objParm
                                         
                                         Set objParm = .CreateParameter("p_organization_id", adNumeric, adParamInput, 200, var_unidad_organizacional)
                                         .Parameters.Append objParm
                                         
                                         Set objParm = .CreateParameter("p_group_id", adNumeric, adParamInput, 200, var_group_id)
                                         .Parameters.Append objParm
                                         
                                         On Error GoTo SALIR
                                        .execute
                            
                                        var_concurrente = .Parameters("x_concurrente").Value
                                        objConn.CommitTrans
                                    End With
                                    Set objConn = Nothing
                                    Set objCmd = Nothing
                                    var_j = 0
                                 End If
                           Wend
                           rsaux1.Close
                           'Dim var_concurrente As Double
                           'var_concurrente = 0
                           'objConn.Open var_conexion_oracle
                           'With objCmd
                           '     objConn.BeginTrans
                           '     .ActiveConnection = objConn
                           '     .CommandText = "XXVIA_PK_INVENTARIOS.XXVIA_SP_CONCURRENTE_MAT"
                           '     .CommandType = adCmdStoredProc
                           '
                           '     Set objParm = .CreateParameter("x_concurrente", adNumeric, adParamOutput, 50, var_concurrente)
                           '     .Parameters.Append objParm
                           '
                            '     Set objParm = .CreateParameter("p_tipo_movimiento", adVarChar, adParamInput, 200, "Traspasos")
                           '     .Parameters.Append objParm
                           '     On Error GoTo salir
                           '     .execute
                           '
                           '     var_concurrente = .Parameters("x_concurrente").Value
                           '     objConn.CommitTrans
                           'End With
                           'Set objConn = Nothing
                           'Set objCmd = Nothing
                        
                        End If
                        rsaux.Close
                        rsaux7.Close
                        rsaux7.Open "update xxvia_tb_dev_clientes_desgloce set estatus = 'I' where numero = " + CStr(var_numero_folio_devoluciones) + " and organizacion = " + var_unidad_organizacional + " and movimiento = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        rsaux7.Open "update xxvia_tb_devoluciones_clientes set estatus = 'I', fecha_fin = '" + CStr(Date) + "' where numero = " + CStr(var_numero_folio_devoluciones) + " and organizacion = " + var_unidad_organizacional + " and movimiento = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        
                                              
                        
                        If rsaux8.State = 1 Then
                           rsaux8.Close
                        End If
                        
                        Me.frm_mensaje.Visible = True
                        Me.lbl_mensaje = "GENERANDO TRANSACCION"
                        Me.Refresh
                        var_encontro = 0
                        While var_encontro = 0
                               rsaux9.Open "select * from rcv_shipment_headers where attribute15 = 'SID" + var_clave_movimiento + "_" + Trim(Trim(CStr(var_numero_folio_devoluciones))) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                               If Not rsaux9.EOF Then
                                  var_encontro = 1
                               End If
                               rsaux9.Close
                        Wend
                        
                        x = 1
                        If x = 0 Then
                        clnt.MSSoapInit var_webservice
                        For var_j = 1 To 2
                           var_con = clnt.ejecutar_autoinvoice("OM_FACTURAS", 4002)
                        Next var_j
                     
                        Me.frm_mensaje.Visible = True
                        Me.lbl_mensaje = "GENERANDO NOTA DE CREDITO"
                        Me.Refresh
                        var_encontro = 0
                        var_i = 0
                        While var_encontro = 0
                              If var_i = 50 Then
                                 Me.frm_mensaje.Visible = True
                                 Me.lbl_mensaje = "FAVOR DE CORRER LA FACTURA ELECTRONICA"
                                 var_i = 0
                                 For var_j = 1 To 2
                                     var_con = clnt.ejecutar_autoinvoice("OM_FACTURAS", 4002)
                                 Next var_j
                              End If
                              var_cadena = "SELECT * FROM RA_CUSTOMER_TRX_LINES_ALL WHERE SALES_ORDER = " + CStr(var_nUMERO_pedido)
                              rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux.EOF Then
                                 var_encontro = 1
                              End If
                              rsaux.Close
                              var_i = var_i + 1
                        Wend
                        Set clint = Nothing
   
                     
                     
                     
                                       
                                       
                        rsaux.Open "SELECT oha.header_id, oha.ordered_date, oha.order_number,  HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  E.NAME, f.orig_system_reference from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU, hz_customer_profiles D, ar_collectors E, hz_cust_acct_sites_all f Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND D.collector_id = e.collector_id and HCSU.site_use_code = 'BILL_TO' and f.cust_acct_site_id = HCAS.CUST_ACCT_SITE_ID and order_number  = '" + CStr(var_nUMERO_pedido) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           var_encontros = 0
                           VAR_Z = 0
                           While var_encontros = 0
                                 If VAR_Z = 1000 Then
                                    VAR_Z = 0
                                 End If
                                 var_cadena = "SELECT RCT.CUSTOMER_TRX_ID, APS.TRX_NUMBER, APS.AMOUNT_DUE_ORIGINAL, APS.STATUS, APS.CLASS, HCAS.CUST_ACCT_SITE_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  E.COLLECTOR_ID, E.NAME From hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, HZ_CUST_SITE_USES_ALL HCSU, RA_CUSTOMER_TRX_ALL RCT, AR_PAYMENT_SCHEDULES_ALL APS, ar_collectors E, hz_customer_profiles D Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND RCT.BILL_TO_SITE_USE_ID = HCSU.SITE_USE_ID AND INTERFACE_HEADER_ATTRIBUTE1 IN ('" + CStr(var_nUMERO_pedido) + "') AND RCT.customer_trx_id = APS.customer_trx_id AND E.collector_id = D.COLLECTOR_ID AND D.site_use_id = HCSU.SITE_USE_ID "
                                 rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 If Not rsaux2.EOF Then
                                    var_encontros = 1
                                    var_customer_trx_id = rsaux2!customer_Trx_id
                                 End If
                                 rsaux2.Close
                                 VAR_Z = VAR_Z + 1
                           Wend
                           objConn.Open var_conexion_oracle
                           '… Establecer conexión a la base de datos con el objeto objConn.
                           With objCmd
                                objConn.BeginTrans
                                .ActiveConnection = objConn
                                .CommandText = "xxvia_pk_fact_pos_ar.ejecuta_conc_fact"
                                .CommandType = adCmdStoredProc
                                   
                                rsaux10.Open "select * from tb_unidadesorganizacionales where vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                                If Not rsaux10.EOF Then
                                   var_responsabilidad_facturacion = IIf(IsNull(rsaux10!RESPONSABILIDAD_FACTURACION), "", rsaux10!RESPONSABILIDAD_FACTURACION)
                                End If
                                rsaux10.Close
                                   
                                
                                Set objParm = .CreateParameter("p_responsabilidad", adVarChar, adParamInput, 100, var_responsabilidad_facturacion)
                                .Parameters.Append objParm
                                
                                'Set objParm = .CreateParameter("p_customer_trx_id", adNumeric, adParamInput, 50, var_customer_trx_id)
                                Set objParm = .CreateParameter("p_customer_trx_id", adNumeric, adParamInput, 50, Null)
                                .Parameters.Append objParm
                                
                                Set objParm = .CreateParameter("p_esperar", adNumeric, adParamInput, 50, 1)
                                .Parameters.Append objParm
                               
                                Set objParm = .CreateParameter("p_fact_pagada", adVarChar, adParamInput, 50, "Y")
                                .Parameters.Append objParm
                                
                                var_estatus_factura = ""
                                Set objParm = .CreateParameter("p_estatus", adVarChar, adParamOutput, 50, var_estatus_factura)
                                .Parameters.Append objParm
                                rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                
                                On Error GoTo SALIR
                                .execute
                              
                                var_estatus_factura = .Parameters("p_estatus").Value
                                'MsgBox var_responsabilidad_facturacion
                                'MsgBox var_estatus_factura
                                objConn.CommitTrans
                           End With
                           Set objConn = Nothing
                           Set objCmd = Nothing
                        
                        End If
                        rsaux.Close
                        End If
                        Me.frm_mensaje.Visible = False
                             
                  
                  
                        rsaux8.Open "select * from rcv_shipment_headers where attribute15 = 'SID" + var_clave_movimiento + "_" + Trim(Trim(CStr(var_numero_folio_devoluciones))) + "'", cnnoracle_4
                        If Not rsaux8.EOF Then
                           rsaux9.Open "select * from rcv_shipment_lines where shipment_header_id = " + CStr(rsaux8!shipment_header_id), cnnoracle_4, adOpenDynamic, adLockOptimistic
                           'rsaux9.Open "select * from rcv_shipment_lines where attribute15 = 'SID" + var_clave_movimiento + "_" +  Trim(Trim(CStr(var_numero_folio_devoluciones))) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If Not rsaux9.EOF Then
                              cnn.BeginTrans
                              If rsaux.State = 1 Then
                                 rsaux.Close
                              End If
                              rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_ORACLE_RECEPCIONES", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux.EOF Then
                                 var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
                              Else
                                 var_consecutivo = 1
                              End If
                              rsaux.Close
                              rsaux.Open "INSERT INTO TB_TEMP_ORACLE_RECEPCIONES (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                              cnn.CommitTrans
                              rsaux7.Open "SELECT * FROM XXVIA_TB_DEVOLUCIONES_CLIENTES WHERE NUMERO = " + Trim(Trim(CStr(var_numero_folio_devoluciones))) + " and movimiento = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux7.EOF Then
                                 var_nombre_cliente = rsaux7!nombre_cliente
                                 var_establecimiento = rsaux7!ESTABLECIMIENTO
                                 var_cliente = rsaux7!Cliente
                                 var_agente = rsaux7!Agente
                                 VAR_USUARIO_MOV = rsaux7!USUARIO
                                 FECHA_INICIO = rsaux7!FECHA_INICIO
                                 var_referencia = rsaux7!Referencia
                                 var_maquina = rsaux7!MAQUINA
                                 fecha_fin = rsaux7!fecha_fin
                                 
                              End If
                              rsaux7.Close
                              rsaux7.Open "SELECT address1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, HZ_CUST_SITE_USES_ALL HCSU, hz_customer_profiles D Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= " + CStr(var_establecimiento) + " AND HCSU.SITE_USE_ID = D.site_use_id AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                              If Not rsaux7.EOF Then
                                 VAR_NOMBRE_ESTABLECIMIENTO = IIf(IsNull(rsaux7!address1), "", rsaux7!address1)
                              End If
                              rsaux7.Close
                              rsaux7.Open "SELECT address1, HCSU.ATTRIBUTE1 as rfc from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, HZ_CUST_SITE_USES_ALL HCSU, hz_customer_profiles D Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= " + CStr(var_cliente) + " AND HCSU.SITE_USE_ID = D.site_use_id AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                              If Not rsaux7.EOF Then
                                 var_rfc = IIf(IsNull(rsaux7!rfc), "XAXX010101000", rsaux7!rfc)
                              Else
                                 var_rfc = "XAXX010101000"
                              End If
                              rsaux7.Close
                              
                              rsaux.Open "SELECT * FROM AR_COLLECTORS WHERE COLLECTOR_ID = '" + CStr(var_agente) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux.EOF Then
                                 var_nombre_agente = IIf(IsNull(rsaux!Name), "", rsaux!Name)
                              End If
                              rsaux.Close
                          
                              var_almacen = rsaux9!TO_SUBINVENTORY
                              rsaux.Open "SELECT * FROM mtl_secondary_inventories WHERE ORGANIZATION_ID = " + var_unidad_organizacional + " AND SECONDARY_INVENTORY_NAME = '" + IIf(IsNull(rsaux9!TO_SUBINVENTORY), "", rsaux9!TO_SUBINVENTORY) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux.EOF Then
                                 var_nombre_almacen_subinventario = IIf(IsNull(rsaux!Description), "", rsaux!Description)
                              End If
                              rsaux.Close
                              rsaux.Open "SELECT * FROM HR_ALL_ORGANIZATION_UNITS WHERE ORGANIZATION_ID = " + CStr(rsaux9!TO_organizaTion_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux.EOF Then
                                 var_nombre_unidad_origen = IIf(IsNull(rsaux!Name), "", rsaux!Name)
                              End If
                              rsaux.Close
                              
                              rsaux.Open "select ORDER_NUMBER from oe_order_headers_all where orig_sys_document_ref = 'SID" + var_clave_movimiento + "_" + Trim(CStr(var_numero_folio_devoluciones)) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              var_pedido = rsaux(0).Value
                              rsaux.Close
                              While Not rsaux9.EOF
                                    rsaux.Open "SELECT * FROM xxvia_system_items_b WHERE ORGANIZATION_ID = " + var_unidad_organizacional + " AND INVENTORY_ITEM_ID = " + CStr(rsaux9!ITEM_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                    VAR_SEGMENT2 = ""
                                    If Not rsaux.EOF Then
                                       VAR_SEGMENT2 = IIf(IsNull(rsaux!SEGMENT1), "", rsaux!SEGMENT1)
                                       var_descripcion = IIf(IsNull(rsaux!Description), "", rsaux!Description)
                                       var_ubicacion = IIf(IsNull(rsaux!attribute2), "", rsaux!attribute2)
                                    End If
                                    rsaux.Close
                                    'var_clave_usuario_movimiento = ""
                                    var_cadena = "INSERT INTO TB_TEMP_ORACLE_RECEPCIONES (INTE_TEM_CONSECUTIVO, ORGANIZACION_DESTINO, ORGANIZACION_ORIGEN, SHIPMENT_NUM, SHIPMENT_HEADER_ID, SHIPMENT_LINE_ID,                                                SUBINVENTARIO, SEGMENT1, CANTIDAD_ENVIADA, CANTIDAD_RECIBIDA, DESCRIPCION, NOMBRE_ORGANIZACION_DESTINO, NOMBRE_ORGANIZACION_ORIGEN,NOMBRE_SUBINVENTARIO,USUARIO, MAQUINA, FECHA_INICIO, FECHA_FIN, CLIENTE_ID, ESTABLECIMIENTO_ID,VENDOR_ID, NOMBRE_CLIENTE,NOMBRE_ESTABLECIMIENTO,NOMBRE_PROVEEDOR,MOVIMIENTO, FOLIO,referencia, PEDIDO, UBICACION, RFC)"
                                    var_cadena = var_cadena + "VALUES (" + CStr(var_consecutivo) + "," + CStr(IIf(IsNull(rsaux9!TO_organizaTion_ID), 0, rsaux9!TO_organizaTion_ID)) + "," + CStr(IIf(IsNull(rsaux9!FROM_ORGANIZATION_ID), 0, rsaux9!FROM_ORGANIZATION_ID)) + ",'" + Trim(Trim(CStr(var_numero_folio_devoluciones))) + "',0,0,'" + IIf(IsNull(rsaux9!TO_SUBINVENTORY), "", rsaux9!TO_SUBINVENTORY) + "','" + VAR_SEGMENT2 + "'," + CStr(rsaux9!QUANTITY_RECEIVED) + "," + CStr(rsaux9!QUANTITY_RECEIVED) + ",'" + var_descripcion + "','" + var_nombre_unidad_origen + "','" + var_nombre_unidad_origen + "','" + var_nombre_almacen_subinventario + "','" + var_clave_usuario_global + "','" + var_maquina + "','" + CStr(FECHA_INICIO) + "','" + CStr(IIf(IsNull(fecha_fin), "", fecha_fin)) + "','" + CStr(var_cliente) + "','" + CStr(var_establecimiento) + "','" + CStr(var_agente) + "','" + var_nombre_cliente + "','" + VAR_NOMBRE_ESTABLECIMIENTO + "','"
                                    If var_clave_movimiento = "DCS" Then
                                       var_cadena = var_cadena + var_nombre_agente + "','DEVOLUCION DE CLIENTES CON REFERENCIA','" + Trim(Trim(CStr(var_numero_folio_devoluciones))) + "','" + var_referencia + " DEV. AGENTE: " + var_sello + "'," + CStr(var_pedido) + ",'" + var_ubicacion + "','" + var_rfc + "')"
                                    Else
                                       var_cadena = var_cadena + var_nombre_agente + "','DEVOLUCION DE CLIENTES','" + Trim(Trim(CStr(var_numero_folio_devoluciones))) + "','" + var_referencia + "'," + CStr(var_pedido) + ",'" + var_ubicacion + "','" + var_rfc + "')"
                                    End If
                                    rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                    rsaux9.MoveNext
                              Wend
                              rsaux.Open "delete from TB_TEMP_ORACLE_RECEPCIONES where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and maquina is null", cnn, adOpenDynamic, adLockOptimistic
                              Set reporte = appl.OpenReport(App.Path + "\rep_oracle_recepciones_devoluciones_clientes.rpt")
                              reporte.RecordSelectionFormula = "{VW_ORACLE_RECEPCIONES.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                              frmvistasprevias.cr.ReportSource = reporte
                              For ntablas = 1 To reporte.Database.Tables.Count
                                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                              Next ntablas
                              frmvistasprevias.cr.ViewReport
                              frmvistasprevias.Caption = "Devolución de clientes"
                              frmvistasprevias.Show 1
                              Set reporte = Nothing
                              If var_estatus_factura = "N" Then
                                 MsgBox "No se pudo generar el documento fiscal", vbOKOnly
                              End If

                           Else
                              MsgBox "El movimiento no se a terminado de generar en ORACLE, espere un momento por favor y vuelvalo a intentar"
                           End If
                           rsaux9.Close
                        Else
                           MsgBox "El movimiento no se a terminado de generar en ORACLE, espere un momento por favor y vuelvalo a intentar"
                        End If
                        rsaux8.Close
                     Else
                        MsgBox "El número de pedido es: " + CStr(var_nUMERO_pedido), vbOKOnly, "ATENCION"
                     End If
                  End If
               Else
                  If rsaux8.State = 1 Then
                     rsaux8.Close
                  End If
                  rsaux8.Open "select * from  oe_order_headers_all where orig_sys_document_ref = 'SID" + var_clave_movimiento + "_" + Trim(Trim(CStr(var_numero_folio_devoluciones))) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  VAR_HEADER_ID = 0
                  If Not rsaux8.EOF Then
                     VAR_HEADER_ID = IIf(IsNull(rsaux8!header_id), 0, rsaux8!header_id)
                  End If
                  rsaux8.Close
                  'rsaux8.Open "select * from rcv_shipment_headers where attribute15 = 'SID" + var_clave_movimiento + "_" +  + Trim(Trim(CStr(var_numero_folio_devoluciones))) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  'If Not rsaux8.EOF Then
                     rsaux9.Open "select * from rcv_shipment_lines where oe_order_header_id = " + CStr(VAR_HEADER_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     'rsaux9.Open "select * from rcv_shipment_lines where attribute15 = 'SID" + var_clave_movimiento + "_" +  Trim(Trim(CStr(var_numero_folio_devoluciones))) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux9.EOF Then
                        cnn.BeginTrans
                        rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_ORACLE_RECEPCIONES", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
                        Else
                           var_consecutivo = 1
                        End If
                        rsaux.Close
                        rsaux.Open "INSERT INTO TB_TEMP_ORACLE_RECEPCIONES (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                        cnn.CommitTrans
                        rsaux7.Open "SELECT * FROM XXVIA_TB_DEVOLUCIONES_CLIENTES WHERE NUMERO = " + Trim(Trim(CStr(var_numero_folio_devoluciones))) + " and movimiento = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux7.EOF Then
                           var_nombre_cliente = rsaux7!nombre_cliente
                           var_establecimiento = rsaux7!ESTABLECIMIENTO
                           var_cliente = rsaux7!Cliente
                           var_agente = rsaux7!Agente
                           VAR_USUARIO_MOV = rsaux7!USUARIO
                           FECHA_INICIO = rsaux7!FECHA_INICIO
                           fecha_fin = rsaux7!fecha_fin
                           var_referencia = rsaux7!Referencia
                           var_maquina = rsaux7!MAQUINA
                        End If
                        rsaux7.Close
                        rsaux7.Open "SELECT address1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, HZ_CUST_SITE_USES_ALL HCSU, hz_customer_profiles D Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= " + CStr(var_establecimiento) + " AND HCSU.SITE_USE_ID = D.site_use_id AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                        If Not rsaux7.EOF Then
                           VAR_NOMBRE_ESTABLECIMIENTO = IIf(IsNull(rsaux7!address1), "", rsaux7!address1)
                        End If
                        rsaux7.Close
                        rsaux7.Open "SELECT address1, HCSU.ATTRIBUTE1 as rfc from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, HZ_CUST_SITE_USES_ALL HCSU, hz_customer_profiles D Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= " + CStr(var_cliente) + " AND HCSU.SITE_USE_ID = D.site_use_id AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID"
                        If Not rsaux7.EOF Then
                           var_rfc = IIf(IsNull(rsaux7!rfc), "XAXX010101000", rsaux7!rfc)
                        Else
                           var_rfc = "XAXX010101000"
                        End If
                        rsaux7.Close
                        rsaux.Open "SELECT * FROM AR_COLLECTORS WHERE COLLECTOR_ID = '" + CStr(var_agente) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           var_nombre_agente = IIf(IsNull(rsaux!Name), "", rsaux!Name)
                        End If
                        rsaux.Close
                     
                        var_almacen = rsaux9!TO_SUBINVENTORY
                        rsaux.Open "SELECT * FROM mtl_secondary_inventories WHERE ORGANIZATION_ID = " + var_unidad_organizacional + " AND SECONDARY_INVENTORY_NAME = '" + IIf(IsNull(rsaux9!TO_SUBINVENTORY), "", rsaux9!TO_SUBINVENTORY) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           var_nombre_almacen_subinventario = IIf(IsNull(rsaux!Description), "", rsaux!Description)
                        End If
                        rsaux.Close
                        rsaux.Open "SELECT * FROM HR_ALL_ORGANIZATION_UNITS WHERE ORGANIZATION_ID = " + CStr(rsaux9!TO_organizaTion_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           var_nombre_unidad_origen = IIf(IsNull(rsaux!Name), "", rsaux!Name)
                        End If
                        rsaux.Close
                        
                        rsaux.Open "select ORDER_NUMBER from oe_order_headers_all where orig_sys_document_ref = 'SID" + var_clave_movimiento + "_" + Trim(CStr(var_numero_folio_devoluciones)) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        var_pedido = rsaux(0).Value
                        rsaux.Close
                        
                        While Not rsaux9.EOF
                              rsaux.Open "SELECT * FROM xxvia_system_items_b WHERE ORGANIZATION_ID = " + var_unidad_organizacional + " AND INVENTORY_ITEM_ID = " + CStr(rsaux9!ITEM_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                              VAR_SEGMENT2 = ""
                              If Not rsaux.EOF Then
                                 VAR_SEGMENT2 = IIf(IsNull(rsaux!SEGMENT1), "", rsaux!SEGMENT1)
                                 var_descripcion = IIf(IsNull(rsaux!Description), "", rsaux!Description)
                                 var_ubicacion = IIf(IsNull(rsaux!attribute2), "", rsaux!attribute2)
                              End If
                              rsaux.Close
                              var_cadena = "INSERT INTO TB_TEMP_ORACLE_RECEPCIONES (INTE_TEM_CONSECUTIVO, ORGANIZACION_DESTINO, ORGANIZACION_ORIGEN, SHIPMENT_NUM, SHIPMENT_HEADER_ID, SHIPMENT_LINE_ID,                                                SUBINVENTARIO, SEGMENT1, CANTIDAD_ENVIADA, CANTIDAD_RECIBIDA, DESCRIPCION, NOMBRE_ORGANIZACION_DESTINO, NOMBRE_ORGANIZACION_ORIGEN,NOMBRE_SUBINVENTARIO,USUARIO, MAQUINA, FECHA_INICIO, FECHA_FIN, CLIENTE_ID, ESTABLECIMIENTO_ID,VENDOR_ID, NOMBRE_CLIENTE,NOMBRE_ESTABLECIMIENTO,NOMBRE_PROVEEDOR,MOVIMIENTO, FOLIO,referencia, PEDIDO, UBICACION, RFC)"
                              var_cadena = var_cadena + "VALUES (" + CStr(var_consecutivo) + "," + CStr(IIf(IsNull(rsaux9!TO_organizaTion_ID), 0, rsaux9!TO_organizaTion_ID)) + "," + CStr(IIf(IsNull(rsaux9!FROM_ORGANIZATION_ID), 0, rsaux9!FROM_ORGANIZATION_ID)) + ",'" + Trim(Trim(CStr(var_numero_folio_devoluciones))) + "',0,0,'" + IIf(IsNull(rsaux9!TO_SUBINVENTORY), "", rsaux9!TO_SUBINVENTORY) + "','" + VAR_SEGMENT2 + "'," + CStr(rsaux9!QUANTITY_RECEIVED) + "," + CStr(rsaux9!QUANTITY_RECEIVED) + ",'" + var_descripcion + "','" + var_nombre_unidad_origen + "','" + var_nombre_unidad_origen + "','" + var_nombre_almacen_subinventario + "','" + var_clave_usuario_global + "','" + var_maquina + "','" + CStr(FECHA_INICIO) + "','" + CStr(IIf(IsNull(fecha_fin), "", fecha_fin)) + "','" + CStr(var_cliente) + "','" + CStr(var_establecimiento) + "','" + CStr(var_agente) + "','" + var_nombre_cliente + "','" + VAR_NOMBRE_ESTABLECIMIENTO + "','"
                              If var_clave_movimiento = "DCS" Then
                                 var_cadena = var_cadena + var_nombre_agente + "','DEVOLUCION DE CLIENTES CON REFERENCIA','" + Trim(Trim(CStr(var_numero_folio_devoluciones))) + "','" + IIf(IsNull(var_referencia), "", var_referencia) + " DEV. AGENTE: " + var_sello + "'," + CStr(var_pedido) + ",'" + var_ubicacion + "','" + var_rfc + "')"
                              Else
                                 var_cadena = var_cadena + var_nombre_agente + "','DEVOLUCION DE CLIENTES','" + Trim(Trim(CStr(var_numero_folio_devoluciones))) + "','" + IIf(IsNull(var_referencia), "", var_referencia) + "'," + CStr(var_pedido) + ",'" + var_ubicacion + "','" + var_rfc + "')"
                              End If
                              rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              rsaux9.MoveNext
                        Wend
                        rsaux.Open "delete from TB_TEMP_ORACLE_RECEPCIONES where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and maquina is null", cnn, adOpenDynamic, adLockOptimistic
                        Set reporte = appl.OpenReport(App.Path + "\rep_oracle_recepciones_devoluciones_clientes.rpt")
                        reporte.RecordSelectionFormula = "{VW_ORACLE_RECEPCIONES.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                        frmvistasprevias.cr.ReportSource = reporte
                        For ntablas = 1 To reporte.Database.Tables.Count
                            reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                        Next ntablas
                        frmvistasprevias.cr.ViewReport
                        frmvistasprevias.Caption = "Devolución de clientes"
                        frmvistasprevias.Show 1
                        Set reporte = Nothing
                     Else
                        MsgBox "El movimiento no se a terminado de generar en ORACLE, espere un momento por favor y vuelvalo a intentar"
                     End If
                     rsaux9.Close
                  'Else
                  '   MsgBox "El movimiento no se a terminado de generar en ORACLE, espere un momento por favor y vuelvalo a intentar"
                  'End If
                  'rsaux8.Close
                  'MsgBox "El movimiento ya fue cerrado", vbOKOnly, "ATENCION"
               End If
               rs.Close
            Else
               MsgBox "No existe una lista de precios", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "No se a definido un tipo de pedido para la empresa seleccionada", vbOKOnly, "ATENCION"
         End If
      End If
   Else
      MsgBox "No se han asignado todas las causas de devolución", vbOKOnly, "ATENCION"
   End If
   Exit Sub
SALIR:
   If Err.Number = -2147217900 Then
      'MsgBox Err.Description
      rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
       rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      Resume
   Else
      MsgBox Err.Description
      'Resume
      If rs.State = 1 Then
         rs.Close
      End If
      If rsaux.State = 1 Then
         rsaux.Close
      End If
      If rsaux1.State = 1 Then
         rsaux1.Close
      End If
      If rsaux2.State = 1 Then
         rsaux2.Close
      End If
      If rsaux3.State = 1 Then
         rsaux3.Close
      End If
      If rsaux4.State = 1 Then
         rsaux4.Close
      End If
      If rsaux5.State = 1 Then
         rsaux5.Close
      End If
      If rsaux6.State = 1 Then
         rsaux6.Close
      End If
      If rsaux7.State = 1 Then
         rsaux7.Close
      End If
      If rsaux8.State = 1 Then
         rsaux8.Close
      End If
      If rsaux9.State = 1 Then
         rsaux9.Close
      End If
   End If
   Exit Sub
salir_factura:
   MsgBox "Surgio un error al generar los documentos electrónicos", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   If rsaux1.State = 1 Then
      rsaux1.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If rsaux3.State = 1 Then
      rsaux3.Close
   End If
   If rsaux4.State = 1 Then
      rsaux4.Close
   End If
   If objConn.State = 1 Then
      objConn.RollbackTrans
      objConn.Close
   End If
End Sub

Private Sub cmd_invertir_Click()
   n = lv_devoluciones.ListItems.Count
   For i = 1 To n
      lv_devoluciones.ListItems.Item(i).Selected = True
      If lv_devoluciones.selectedItem.SubItems(4) = "*" Then
         lv_devoluciones.selectedItem.SubItems(4) = ""
         lv_devoluciones.ListItems.Item(i).Bold = False
         lv_devoluciones.ListItems.Item(i).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(4).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(5).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
      Else
         lv_devoluciones.selectedItem.SubItems(4) = "*"
         lv_devoluciones.ListItems.Item(i).Bold = True
         lv_devoluciones.ListItems.Item(i).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(5).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
      End If
   Next i
   If Me.lv_devoluciones.ListItems.Count > 0 Then
      Me.lv_devoluciones.SetFocus
   End If
End Sub

Private Sub cmd_marcar_Click()
   i = lv_devoluciones.selectedItem.Index
   If lv_devoluciones.selectedItem.SubItems(4) = "*" Then
      lv_devoluciones.selectedItem.SubItems(4) = ""
      lv_devoluciones.ListItems.Item(i).Bold = False
      lv_devoluciones.ListItems.Item(i).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(4).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(5).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
      lv_devoluciones.Refresh
   Else
      lv_devoluciones.selectedItem.SubItems(4) = "*"
      lv_devoluciones.ListItems.Item(i).Bold = True
      lv_devoluciones.ListItems.Item(i).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(3).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(4).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(5).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
      lv_devoluciones.Refresh
   End If
   If Me.lv_devoluciones.ListItems.Count > 0 Then
      Me.lv_devoluciones.SetFocus
   End If
End Sub

Private Sub cmd_ninguno_Click()
   n = lv_devoluciones.ListItems.Count
   For i = 1 To n
      lv_devoluciones.ListItems.Item(i).Selected = True
      lv_devoluciones.selectedItem.SubItems(4) = ""
      lv_devoluciones.ListItems.Item(i).Bold = False
      lv_devoluciones.ListItems.Item(i).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(4).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(5).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
   Next i
   lv_devoluciones.Refresh
   If Me.lv_devoluciones.ListItems.Count > 0 Then
      Me.lv_devoluciones.SetFocus
   End If

End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmd_seleccion_Click()
   n = lv_devoluciones.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_devoluciones.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_devoluciones.selectedItem.SubItems(4) = "" And var_rellena = True Then
         lv_devoluciones.selectedItem.SubItems(4) = "*"
         lv_devoluciones.ListItems.Item(i).Bold = True
         lv_devoluciones.ListItems.Item(i).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(5).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_devoluciones.selectedItem.SubItems(4) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_devoluciones.selectedItem.SubItems(4) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
   If Me.lv_devoluciones.ListItems.Count > 0 Then
      Me.lv_devoluciones.SetFocus
   End If
End Sub

Private Sub cmd_todos_Click()
   n = lv_devoluciones.ListItems.Count
   For i = 1 To n
      lv_devoluciones.ListItems.Item(i).Selected = True
      lv_devoluciones.selectedItem.SubItems(4) = "*"
      lv_devoluciones.ListItems.Item(i).Bold = True
      lv_devoluciones.ListItems.Item(i).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(3).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(4).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(5).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
   Next i
   lv_devoluciones.Refresh
   If Me.lv_devoluciones.ListItems.Count > 0 Then
      Me.lv_devoluciones.SetFocus
   End If
End Sub

Private Sub Command1_Click()
   rs.Open "SELECT DISTINCT NUMERO, REFERENCIA FROM XXVIA_tB_DEVOLUCIONES_CLIENTES", cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         var_referencia = IIf(IsNull(rs!Referencia), "", rs!Referencia)
         VAR_ORIGEN = "SID" + var_clave_movimiento + "_" + CStr(IIf(IsNull(rs!numero), 0, rs!numero))
         'MsgBox VAR_ORIGEN
         rsaux.Open "UPDATE OE_ORDER_HEADERS_ALL SET ATTRIBUTE7 = '" + var_referencia + "' WHERE orig_sys_document_ref = '" + VAR_ORIGEN + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub Form_Load()
   rs.Open "select * from xxvia_tb_dev_clientes_desgloce where numero = " + CStr(var_numero_folio_devoluciones) + " and organizacion = " + var_unidad_organizacional + " AND MOVIMIENTO = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set list_item = Me.lv_devoluciones.ListItems.Add(, , IIf(IsNull(rs!codigo), "", rs!codigo))
         list_item.SubItems(1) = IIf(IsNull(rs!Descripcion), "", rs!Descripcion)
         list_item.SubItems(2) = Format(IIf(IsNull(rs!cantidad), 0, rs!cantidad), "###,###,##0.00")
         list_item.SubItems(3) = IIf(IsNull(rs!descripcion_causa), "", rs!descripcion_causa)
         list_item.SubItems(4) = ""
         list_item.SubItems(5) = IIf(IsNull(rs!causa_devolucion), "", rs!causa_devolucion)
         list_item.SubItems(6) = IIf(IsNull(rs!CONSECUTIVO), "", rs!CONSECUTIVO)
         list_item.SubItems(7) = IIf(IsNull(rs!inventory_item_id), 0, rs!inventory_item_id)
         list_item.SubItems(8) = IIf(IsNull(rs!localizador), "", rs!localizador)
         rs.MoveNext
   Wend
   rs.Close
   Me.frm_lista.Visible = False
   Me.frm_mensaje.Visible = False
End Sub

Private Sub Label1_Click()

End Sub

Private Sub lv_devoluciones_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_devoluciones, ColumnHeader)
End Sub

Private Sub lv_devoluciones_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = 115 Then
      rs.Open "select * from xxvia_tb_dev_clientes_DESGLOCE where NUMERO = " + CStr(var_numero_folio_devoluciones) + " and organizacion = " + var_unidad_organizacional + " and movimiento = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      VAR_ESTATUS = Trim(IIf(IsNull(rs!estatus), "", rs!estatus))
      rs.Close
      If VAR_ESTATUS = "" Then
         Me.lv_lista.ListItems.Clear
         If var_unidad_organizacional = "85" Then
            If var_clave_usuario_global <> "U0000000314" Then
               var_cadena = "select lookup_code as CODIGO, meaning as NOMBRE, description as DESCRIPCION From FND_LOOKUP_VALUES_VL where lookup_type = 'CREDIT_MEMO_REASON' and enabled_flag = 'Y' AND MEANING LIKE '%ROLLO%' ORDER BY 1"
            Else
               var_cadena = "select lookup_code as CODIGO, meaning as NOMBRE, description as DESCRIPCION From FND_LOOKUP_VALUES_VL where lookup_type = 'CREDIT_MEMO_REASON' and enabled_flag = 'Y' ORDER BY 1"
            End If
         Else
            var_cadena = "select lookup_code as CODIGO, meaning as NOMBRE, description as DESCRIPCION From FND_LOOKUP_VALUES_VL where lookup_type = 'CREDIT_MEMO_REASON' and enabled_flag = 'Y' ORDER BY 1"
         End If
         If rs.State = 1 Then
            rs.Close
         End If
         rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = Me.lv_lista.ListItems.Add(, , rs(0).Value)
               list_item.SubItems(1) = rs(1).Value
               rs.MoveNext
         Wend
         rs.Close
         If lv_lista.ListItems.Count > 0 Then
            Me.frm_lista.Visible = True
            Me.lv_lista.SetFocus
         Else
            MsgBox "No existen causas de devolución", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El movimiento ya no puede ser modificado", vbOKOnly, "ATENCION"
      End If
      
   End If
End Sub

Private Sub lv_devoluciones_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      i = lv_devoluciones.selectedItem.Index
      If lv_devoluciones.selectedItem.SubItems(4) = "*" Then
         lv_devoluciones.selectedItem.SubItems(4) = ""
         lv_devoluciones.ListItems.Item(i).Bold = False
         lv_devoluciones.ListItems.Item(i).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(4).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(5).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
         lv_devoluciones.Refresh
      Else
         lv_devoluciones.selectedItem.SubItems(4) = "*"
         lv_devoluciones.ListItems.Item(i).Bold = True
         lv_devoluciones.ListItems.Item(i).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(5).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
         lv_devoluciones.Refresh
      End If
   End If
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      var_si = MsgBox("Se asignara la causa de devolución a los artículos seleccionados", vbYesNo, "ATENCION")
      If var_si = 6 Then
         For var_j = 1 To lv_devoluciones.ListItems.Count
             Me.lv_devoluciones.ListItems(var_j).Selected = True
             If Me.lv_devoluciones.selectedItem.SubItems(4) = "*" Then
                Me.lv_devoluciones.selectedItem.SubItems(3) = Me.lv_lista.selectedItem.SubItems(1)
                Me.lv_devoluciones.selectedItem.SubItems(5) = Me.lv_lista.selectedItem
                rsaux.Open "update xxvia_tb_dev_clientes_desgloce set causa_devolucion = '" + Me.lv_lista.selectedItem + "', descripcion_causa = '" + Me.lv_lista.selectedItem.SubItems(1) + "' where numero = " + CStr(var_numero_folio_devoluciones) + " and inventory_item_id = " + Me.lv_devoluciones.selectedItem.SubItems(7) + " and consecutivo= " + Me.lv_devoluciones.selectedItem.SubItems(6), cnnoracle_4, adOpenDynamic, adLockOptimistic
                lv_devoluciones.selectedItem.SubItems(4) = ""
                lv_devoluciones.ListItems.Item(var_j).Bold = False
                lv_devoluciones.ListItems.Item(var_j).ForeColor = &H80000012
                lv_devoluciones.ListItems.Item(var_j).ListSubItems(1).Bold = False
                lv_devoluciones.ListItems.Item(var_j).ListSubItems(2).Bold = False
                lv_devoluciones.ListItems.Item(var_j).ListSubItems(3).Bold = False
                lv_devoluciones.ListItems.Item(var_j).ListSubItems(4).Bold = False
                lv_devoluciones.ListItems.Item(var_j).ListSubItems(5).Bold = False
                lv_devoluciones.ListItems.Item(var_j).ListSubItems(1).ForeColor = &H80000012
                lv_devoluciones.ListItems.Item(var_j).ListSubItems(2).ForeColor = &H80000012
                lv_devoluciones.ListItems.Item(var_j).ListSubItems(3).ForeColor = &H80000012
                lv_devoluciones.ListItems.Item(var_j).ListSubItems(4).ForeColor = &H80000012
                lv_devoluciones.ListItems.Item(var_j).ListSubItems(5).ForeColor = &H80000012
             End If
         Next var_j
         Me.frm_lista.Visible = False
      Else
         Me.lv_devoluciones.SetFocus
      End If
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub
