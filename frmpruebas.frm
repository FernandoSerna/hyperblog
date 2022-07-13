VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmpruebas 
   Caption         =   "Form1"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15630
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   15630
   Begin VB.CommandButton Command29 
      Caption         =   "timbrar documentos V 4.0"
      Height          =   1335
      Left            =   13200
      TabIndex        =   43
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command28 
      Caption         =   "postgres"
      Height          =   855
      Left            =   12720
      TabIndex        =   42
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Command27"
      Height          =   735
      Left            =   11280
      TabIndex        =   40
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Actualizar FA"
      Height          =   855
      Left            =   11160
      TabIndex        =   39
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Command25"
      Height          =   855
      Left            =   11040
      TabIndex        =   38
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txt_embraue_nota_envio 
      Height          =   375
      Left            =   9360
      TabIndex        =   37
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txt_embarque_pedido 
      Height          =   375
      Left            =   7920
      TabIndex        =   36
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command24 
      Caption         =   "restructurar packing list"
      Height          =   855
      Left            =   11160
      TabIndex        =   35
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Eliminar ubicaciones"
      Height          =   855
      Left            =   11160
      TabIndex        =   34
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command22 
      Caption         =   "complementos"
      Height          =   615
      Left            =   3360
      TabIndex        =   33
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command21 
      Caption         =   "ascii tab"
      Height          =   735
      Left            =   6360
      TabIndex        =   31
      Top             =   120
      Width           =   1455
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   495
      Left            =   4200
      TabIndex        =   30
      Top             =   1920
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   873
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "P##-N##-"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Command20"
      Height          =   615
      Left            =   2400
      TabIndex        =   29
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton Command19 
      Caption         =   "fraccion"
      Height          =   735
      Left            =   8640
      TabIndex        =   28
      Top             =   5520
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   360
      TabIndex        =   27
      Text            =   "Text4"
      Top             =   5640
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Height          =   855
      Left            =   3840
      TabIndex        =   26
      Text            =   "Text3"
      Top             =   5520
      Width           =   3255
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Cargar lista precios SID 2"
      Height          =   735
      Left            =   8400
      TabIndex        =   25
      Top             =   4560
      Width           =   2295
   End
   Begin VB.CommandButton Command17 
      Caption         =   "cargar equivalencias SID 2"
      Height          =   735
      Left            =   8400
      TabIndex        =   24
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CommandButton cmd_cargar_articulos_sid_2 
      Caption         =   "Cargar artículos SID 2"
      Height          =   855
      Left            =   8400
      TabIndex        =   23
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox txt_codigo 
      Height          =   615
      Left            =   7440
      TabIndex        =   22
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton Command16 
      Caption         =   "cargar archivo de texto"
      Height          =   975
      Left            =   8400
      TabIndex        =   21
      Top             =   1800
      Width           =   2295
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Enviar packing list"
      Height          =   615
      Left            =   5280
      TabIndex        =   20
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Command14"
      Height          =   495
      Left            =   6480
      TabIndex        =   19
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Command13"
      Height          =   615
      Left            =   4080
      TabIndex        =   18
      Top             =   4440
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   360
      TabIndex        =   17
      Text            =   "Text2"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command12 
      Caption         =   "equialencias sid2"
      Height          =   735
      Left            =   2400
      TabIndex        =   16
      Top             =   3240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Alta de articulos sid 2"
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   4440
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Ubicaciones"
      Height          =   495
      Left            =   360
      TabIndex        =   14
      Top             =   3840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Fracciones arancelarias"
      Height          =   615
      Left            =   1800
      TabIndex        =   13
      Top             =   2880
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   855
      Left            =   3240
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3240
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   3960
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox txt_pedido 
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   3240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "correo con pedido"
      Height          =   615
      Left            =   360
      TabIndex        =   9
      Top             =   3720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "correo"
      Height          =   615
      Left            =   360
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "packing list"
      Height          =   615
      Left            =   360
      TabIndex        =   7
      Top             =   420
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.CommandButton Command4 
      Caption         =   "mysql"
      Height          =   675
      Left            =   360
      TabIndex        =   6
      Top             =   3870
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmd_alta_rutas 
      Caption         =   "Alta rutas"
      Height          =   765
      Left            =   360
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exportar excel"
      Height          =   645
      Left            =   360
      TabIndex        =   4
      Top             =   1830
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.CommandButton cmd_creacion_ruta_clientes 
      Caption         =   "Ruta --> Clientes"
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   900
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.CommandButton Command2 
      Caption         =   "REVALUACION DE DEVOLUCIONES"
      Height          =   720
      Left            =   360
      TabIndex        =   2
      Top             =   1845
      Visible         =   0   'False
      Width           =   2910
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   570
      Left            =   360
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmd_telnet 
      Caption         =   "Prueba de video"
      Height          =   555
      Left            =   5640
      TabIndex        =   0
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   10920
      TabIndex        =   41
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   8880
      TabIndex        =   32
      Top             =   1320
      Width           =   1575
   End
End
Attribute VB_Name = "frmpruebas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hora_inicio As Date
Dim var_hora_fin As Date
Dim var_i As Integer
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter





Private Sub cmd_alta_rutas_Click()
   rs.Open "select * from rutas_establecimientos_290916", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux.Open "SELECT CUST_ACCOUNT_ID , ACCOUNT_FULL_NAME NOMBRE_TITULAR ,SITE_USE_ID CLAVE, RAZON_SOCIAL_CLIENTE NOMBRE_ESTABLECIMIENTO, PARTY_SITE_NUMBER,CALLE||' '||COLONIA||' '||' '||CIUDAD AS DIRECCION FROM XXVIA_VW_CLIENTES_BCP WHERE site_use_id = " + CStr(rs!ESTABLECIMIENTO) + " and site_use_code = 'SHIP_TO' ", cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            rsaux1.Open "select * from XXVIA_tb_CLIENTES_RUTAS_DISTR where ruta = '" + CStr(rs!ruta) + "' and establecimiento = '" + CStr(rs!ESTABLECIMIENTO) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rsaux1.EOF Then
               rsaux2.Open "INSERT INTO XXVIA_TB_CLIENTES_RUTAS_DISTR (RUTA, TITULAR, NOMBRE_TITULAR, ESTABLECIMIENTO, NOMBRE_ESTABLECIMIENTO, DIRECCION, PRIORIDAD) values  ('" + rs!ruta + "','" + CStr(rsaux!CUST_ACCOUNT_ID) + "','" + rsaux!nombre_titular + "','" + CStr(rs!ESTABLECIMIENTO) + "','" + rsaux!nombre_Establecimiento + "','" + rsaux!DIRECCION + "',0)", cnnoracle_4, adOpenDynamic, adLockOptimistic
            End If
            rsaux1.Close
         End If
         rsaux.Close
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub cmd_cargar_articulos_sid_2_Click()
   rs.Open "SELECT SEGMENT1 AS CODIGO, DESCRIPTION AS DESCRIPCION, NVL(ATTRIBUTE10,'N') AS RECONTABLE, NVL(WEIGHT_UOM_CODE,' ') UNIDAD_PESO_ID, NVL(UNIT_WEIGHT,0) AS PESO, NVL(VOLUME_UOM_CODE,' ') AS UNIDAD_VOLUMEN_ID, NVL(UNIT_VOLUME,0) VOLUMEN, PRIMARY_UOM_CODE AS UNIDAD_MEDIDA_ID, PRIMARY_UNIT_OF_MEASURE UNIDAD_MEDIDA, nvl(CLASIFICACIONSAT, ' ') clasificacion_sat, UOM_SAT   FROM XXVIA_SYSTEM_ITEMS_B WHERE ORGANIZATION_ID = 89", cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux.Open "select * from sqlquezada2.sid2.dbo.tb_articulos where vcha_Articulo_id = '" + rs!codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If rsaux.EOF Then
            var_codigo = rs!codigo
            var_descripcion = rs!Descripcion
            var_recontable = rs!recontable
            var_unidad_peso_id = rs!UNIDAD_PESO_ID
            var_peso = rs!PESO
            VAR_UNIDAD_volumen_ID = rs!UNIDAD_VOLUMEN_ID
            var_volumen = rs!VOLUMEN
            VAR_UNIDAD_MEDIDA_ID = rs!UNIDAD_MEDIDA_ID
            var_unidad_medida = rs!UNIDAD_MEDIDA
            var_clasificacion_Sat = rs!clasificacion_sat
            VAR_UOM_SAT = rs!UOM_SAT
            If var_recontable = "Y" Then
               VAR_rEC = 1
            Else
               VAR_rEC = 0
            End If
            var_cadena = "INSERT INTO sqlquezada2.sid2.dbo.TB_aRTICULOS (vcha_articulo_id, vcha_descripcion_articulo, vcha_unidad_medida_id,           vcha_unidad_medida_id_volumen, numb_volumen,           vcha_unidad_medida_id_peso, numb_peso, numb_recontable, clasificacion_sat, uom_sat)"
            var_cadena = var_cadena + "     VALUES ('" + var_codigo + "','" + Replace(var_descripcion, "'", " ") + "','" + VAR_UNIDAD_MEDIDA_ID + "','" + UNIDAD_VOLUMEN_ID + "', " + CStr(var_volumen) + ",'" + VAR_UNIDAD_volumen_ID + "'," + CStr(var_peso) + "," + CStr(VAR_rEC) + ",'" + var_clasificacion_Sat + "','" + VAR_UOM_SAT + "')"
            rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         End If
         rsaux.Close
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub cmd_creacion_ruta_clientes_Click()
   rs.Open "select * FROM TB_ORACLE_RUTAS_EMBARQUES where embarque = 255297", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            rsaux.Open "select * from TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque = '" + CStr(rs!Embarque) + "'", cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux.EOF
'----
                  strconsulta = "select * from oe_order_headers_all where order_number = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rsaux!pedido))
                       .Parameters.Append parametro
                  End With
                  Set rsaux6 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  VAR_TITULAR = rsaux6!SOLD_TO_ORG_ID
                  var_cliente = rsaux6!INVOICE_TO_ORG_ID
                  If rsaux6!INVOICE_TO_ORG_ID = 1060 Then
                     strconsulta = "SELECT ORDER_TYPE_ID, source_document_id FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rsaux!pedido)
                          .Parameters.Append parametro
                     End With
                     Set rsaux8 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     var_source_document_id = IIf(IsNull(rsaux8!source_document_id), 0, rsaux8!source_document_id)
                     rsaux8.Close
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
                     var_establecimiento = rsaux7!attribute1
                     rsaux7.Close
                     strconsulta = "select secondary_inventory_name VCHA_ALM_ALMACEN_ID, description VCHA_ALM_NOMBRE  from mtl_secondary_inventories where ATTRIBUTE3 LIKE '%PTO%' AND ORGANIZATION_ID = 93 and secondary_inventory_name = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_establecimiento)
                          .Parameters.Append parametro
                     End With
                     Set rsaux7 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     VAR_TITULAR = 2040
                     VAR_NOMBRE_TITULAR = "VIANNEY TEXTIL HOGAR S.A DE C.V."
                     var_establecimiento = IIf(IsNull(rsaux7!vcha_alm_almacen_id), "", rsaux7!vcha_alm_almacen_id)
                     VAR_NOMBRE_ESTABLECIMIENTO = rsaux7!vcha_alm_nombre
                     VAR_DIRECCION = ""
                     rsaux7.Close
                  Else
                     var_establecimiento = rsaux6!SHIP_TO_ORG_ID
                     strconsulta = "SELECT CUST_ACCOUNT_ID TITULAR, ACCOUNT_FULL_NAME NOMBRE_TITULAR, SITE_USE_ID , RAZON_SOCIAL_CLIENTE NOMBRE_ESTABLECIMIENTO, CALLE||' '||NUM_CALLE||' '||COLONIA||' '||CIUDAD AS DIRECCION FROM XXVIA_VW_CLIENTES_BCP WHERE  site_use_id = ? and site_use_code = 'SHIP_TO'"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_establecimiento)
                          .Parameters.Append parametro
                     End With
                     Set rsaux7 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     VAR_TITULAR = rsaux7!TITULAR
                     VAR_NOMBRE_TITULAR = rsaux7!nombre_titular
                     VAR_NOMBRE_ESTABLECIMIENTO = rsaux7!nombre_Establecimiento
                     VAR_DIRECCION = rsaux7!DIRECCION
                     rsaux7.Close
                  End If
                  rsaux6.Close
                  rsaux7.Open "SELECT * FROM XXVIA_TB_CLIENTES_RUTAS_DISTR WHERE ESTABLECIMIENTO = '" + CStr(var_establecimiento) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If rsaux7.EOF Then
                     rsaux6.Open "INSERT INTO XXVIA_TB_CLIENTES_RUTAS_DISTR (RUTA, TITULAR, NOMBRE_TITULAR, ESTABLECIMIENTO, NOMBRE_ESTABLECIMIENTO, DIRECCION, PRIORIDAD) VALUES ('" + rs!ruta + "','" + CStr(VAR_TITULAR) + "','" + VAR_NOMBRE_TITULAR + "','" + CStr(var_establecimiento) + "','" + VAR_NOMBRE_ESTABLECIMIENTO + "','" + VAR_DIRECCION + "',0)", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  Else
                     rsaux6.Open "update XXVIA_TB_CLIENTES_RUTAS_DISTR set ruta = '" + rs!ruta + "' WHERE ESTABLECIMIENTO = '" + CStr(var_establecimiento) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux7.Close
'----
                  rsaux.MoveNext
            Wend
            rsaux.Close
            rs.MoveNext
      Wend
   End If
   rs.Close
End Sub

Private Sub cmd_telnet_Click()
On Error GoTo SALIR:
        Dim clnt As New SoapClient30
        Set clnt = Nothing
        clnt.MSSoapInit var_webservice_texto
        MsgBox fun_NombrePc, vbOKOnly, ""
        MsgBox "dvr: " + CStr(var_dvr_texto_ip) + " puerto: " + CStr(var_puerto_texto)
        var_s = clnt.insertar_texto(CStr(var_dvr_texto_ip), CStr(var_puerto_texto), "PRUEBA DE SISTEMAS " + CStr(Now))
        Set clnt = Nothing
        clnt.MSSoapInit var_webservice_texto
        var_s = clnt.insertar_texto(CStr(var_dvr_texto_ip), CStr(var_puerto_texto), "MAQUINA: " + fun_NombrePc + ", USUARIO: " + var_nombre_usuario + Chr(13) + "EMBARQUE: prueba")
        Set clnt = Nothing
        Exit Sub
SALIR:
        MsgBox "No es posible insertar el texto error: " + Err.Description
        
 End Sub

Private Sub Command1_Click()
                     cnn.BeginTrans
                     rsaux.Open "SELECT MAX(CAJA) FROM TB_ORACLE_BLOQUEO_CAJAS", cnn, adOpenDynamic, adLockOptimistic
                     strconsulta = "select inte_paq_caja inte_paq_caja from XXVIA_TB_SALIDAS_CAJAS where inte_emb_embarque = ? for update "
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, 156437)
                          .Parameters.Append parametro
                     End With
                     Set rs = comandoORA.execute
                     MsgBox CStr(rs!INTE_PAQ_CAJA)
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     rsaux.Close
                     cnn.CommitTrans
                     rs.Close
End Sub

Private Sub Command10_Click()
x = 1
If x = 1 Then
   rs.Open "select * from XXVIA_TB_UBICACIONES where organizacion = 90 and ubicacion is not null", cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         txt_ubicacion_1 = ""
         txt_ubicacion_2 = ""
         txt_ubicacion_3 = ""
         txt_ubicacion_4 = ""
         txt_ubicacion_5 = ""
         txt_ubicacion_6 = ""
         
         If rs!numero = 1 Then
            txt_ubicacion_1 = rs!ubicacion
               strconsulta = "UPDATE mtl_system_items_b SET  attribute2 = ? where  segment1 = ? and organization_id = ? "
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, txt_ubicacion_1)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rs!codigo)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, 90)
                    .Parameters.Append parametro
               End With
               Set rsaux9 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
         
         End If
         If rs!numero = 2 Then
            txt_ubicacion_2 = rs!ubicacion
               strconsulta = "UPDATE mtl_system_items_b SET  attribute3 = ? where  segment1 = ? and organization_id = ? "
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, txt_ubicacion_2)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rs!codigo)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, 90)
                    .Parameters.Append parametro
               End With
               Set rsaux9 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
         End If
         If rs!numero = 3 Then
            txt_ubicacion_3 = rs!ubicacion
               strconsulta = "UPDATE mtl_system_items_b SET  attribute4 = ? where  segment1 = ? and organization_id = ? "
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, txt_ubicacion_3)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rs!codigo)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, 90)
                    .Parameters.Append parametro
               End With
               Set rsaux9 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
            
         End If
         If rs!numero = 4 Then
            txt_ubicacion_4 = rs!ubicacion
               strconsulta = "UPDATE mtl_system_items_b SET  attribute5 = ? where  segment1 = ? and organization_id = ? "
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, txt_ubicacion_4)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rs!codigo)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, 90)
                    .Parameters.Append parametro
               End With
               Set rsaux9 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
            
         End If
         If rs!numero = 5 Then
            txt_ubicacion_5 = rs!ubicacion
               strconsulta = "UPDATE mtl_system_items_b SET  attribute6 = ? where  segment1 = ? and organization_id = ? "
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, txt_ubicacion_5)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rs!codigo)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, 90)
                    .Parameters.Append parametro
               End With
               Set rsaux9 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
            
         End If
         If rs!numero = 6 Then
            txt_ubicacion_6 = rs!ubicacion
               strconsulta = "UPDATE mtl_system_items_b SET  attribute7 = ? where  segment1 = ? and organization_id = ? "
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, txt_ubicacion_6)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rs!codigo)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, 90)
                    .Parameters.Append parametro
               End With
               Set rsaux9 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
            
         End If
         x = 3
         If x = 4 Then
               strconsulta = "UPDATE mtl_system_items_b SET  attribute2 = ?, attribute3 = ?, attribute4 = ?, attribute5 = ?, attribute6 = ?, attribute7 = ? where  segment1 = ? and organization_id = ? "
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, txt_ubicacion_1)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, txt_ubicacion_2)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, txt_ubicacion_3)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, txt_ubicacion_4)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, txt_ubicacion_5)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, txt_ubicacion_6)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rs!codigo)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, 90)
                    .Parameters.Append parametro
               End With
               Set rsaux9 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
         End If
         rs.MoveNext
   Wend
   rs.Close
   End If
   rs.Open " SELECT * FROM XXVIA_SYSTEM_ITEMS_B WHERE ORGANIZATION_ID = 90 AND NVL(attribute3,' ') <> ' '", cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux1.Open "UPDATE XXVIA_TB_UBICACIONES SET UBICACION = '" + rs!attribute3 + "'  WHERE ORGANIZACION = 90 AND CODIGO = '" + rs!SEGMENT1 + "' AND NUMERO = 2", cnnoracle_4, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
   rs.Open " SELECT * FROM XXVIA_SYSTEM_ITEMS_B WHERE ORGANIZATION_ID = 90 AND NVL(attribute4,' ') <> ' '", cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux1.Open "UPDATE XXVIA_TB_UBICACIONES SET UBICACION = '" + rs!attribute4 + "'  WHERE ORGANIZACION = 90 AND CODIGO = '" + rs!SEGMENT1 + "' AND NUMERO = 3", cnnoracle_4, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
   
   rs.Open " SELECT * FROM XXVIA_SYSTEM_ITEMS_B WHERE ORGANIZATION_ID = 90 AND NVL(attribute5,' ') <> ' '", cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux1.Open "UPDATE XXVIA_TB_UBICACIONES SET UBICACION = '" + rs!attribute5 + "'  WHERE ORGANIZACION = 90 AND CODIGO = '" + rs!SEGMENT1 + "' AND NUMERO = 4", cnnoracle_4, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
   
   rs.Open " SELECT * FROM XXVIA_SYSTEM_ITEMS_B WHERE ORGANIZATION_ID = 90 AND NVL(attribute6,' ') <> ' '", cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux1.Open "UPDATE XXVIA_TB_UBICACIONES SET UBICACION = '" + rs!attribute6 + "'  WHERE ORGANIZACION = 90 AND CODIGO = '" + rs!SEGMENT1 + "' AND NUMERO = 5", cnnoracle_4, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
   
   rs.Open " SELECT * FROM XXVIA_SYSTEM_ITEMS_B WHERE ORGANIZATION_ID = 90 AND NVL(attribute7,' ') <> ' '", cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux1.Open "UPDATE XXVIA_TB_UBICACIONES SET UBICACION = '" + rs!attribute7 + "'  WHERE ORGANIZACION = 90 AND CODIGO = '" + rs!SEGMENT1 + "' AND NUMERO = 6", cnnoracle_4, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
   
End Sub

Private Sub Command11_Click()
   rs.Open "select segment1, description, WEIGHT_UOM_CODE, VOLUME_UOM_CODE, PRIMARY_UOM_CODE  from xxvia_system_items_b where organization_id = 93 and creation_date >= to_date('01/07/2017','DD/MM/YYYY')", cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux.Open "select * from sid2hou.dbo.tb_articulos where vcha_articulo_id = '" + rs!SEGMENT1 + "'", cnn, adOpenDynamic, adLockOptimistic
         If rsaux.EOF Then
            rsaux1.Open "insert into sid2hou.dbo.tb_articulos (vcha_articulo_id, vcha_unidad_medida_id, vcha_Descripcion_articulo, vcha_unidad_medida_id_volumen, vcha_unidad_medida_id_peso) values ('" + rs!SEGMENT1 + "','" + Trim(rs!PRIMARY_UOM_CODE) + "','" + rs!Description + "','" + IIf(IsNull(rs!volume_uom_code), "MTO", rs!volume_uom_code) + "','" + IIf(IsNull(rs!WEIGHT_UOM_CODE), "GR", rs!WEIGHT_UOM_CODE) + "')", cnn, adOpenDynamic, adLockOptimistic
         End If
         rsaux.Close
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub Command12_Click()
var_cadena = "SELECT VCHA_CAJ_CAJA_ID cross_reference, VCHA_ART_aRTICULO_ID as segment1, NUMB_CAJ_CANTIDAD FROM XXVIA_TB_CAJAS_PROD where date_caj_Fecha >= to_Date('01/01/2018','DD/MM/YYYY') AND VCHA_CAJ_ORGANIZACION_ID = 93 "
   rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
  
   While Not rs.EOF
         rsaux1.Open "insert into sid2hou.dbo.tb_equivalencias (vcha_articulo_id, vcha_codigo_equivalente) values ('" + rs!SEGMENT1 + "','" + rs!cross_reference + "')", cnn, adOpenDynamic, adLockOptimistic
         rsaux1.Open "insert into sid2hou.dbo.tb_articulos_Segundo_nivel (codigo, equivalencia, cantidad) values ('" + rs!SEGMENT1 + "','" + rs!cross_reference + "'," + CStr(rs!numb_caj_cantidad) + ")"
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub Command14_Click()
   frmoracle_cajas_NO_divididas.Show
End Sub

Private Sub Command15_Click()
Dim clnt As New SoapClient30
Dim clnt2 As New SoapClient30
clnt2.MSSoapInit "http://serviciowebcedisdesa.vianney.com.mx/EnviarCorreos.asmx?wsdl"
var_cadena = "select  ORDER_NUMBER from mtl_secondary_inventories a, hr_locations_all b, xxvia_jv_tb_agentes c, po_requisition_headers_ALL D, OE_ORDER_HEADERS_ALL E Where A.location_id = b.location_id and a.secondary_inventory_name = c.subinventory_code AND E.source_document_id = D.requisition_header_id AND A.secondary_inventory_name = D.ATTRIBUTE1 and order_type_id = 1002 and ordered_date >= to_Date('20/08/2018','DD/MM/YYYY')"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = var_cadena
                           End With
                           Set rsaux6 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           While Not rsaux6.EOF
                           
                           
                           var_cadena = "select  A.SECONDARY_INVENTORY_NAME, A.DESCRIPTION, ADDRESS_LINE_1, ADDRESS_LINE_2, TOWN_OR_CITY, REGION_1, COUNTRY, POSTAL_CODE, loc_information13 EMAIL from mtl_secondary_inventories a, hr_locations_all b, xxvia_jv_tb_agentes c, po_requisition_headers_ALL D, OE_ORDER_HEADERS_ALL E Where A.location_id = b.location_id and a.secondary_inventory_name = c.subinventory_code AND E.source_document_id = D.requisition_header_id AND A.secondary_inventory_name = D.ATTRIBUTE1 AND E.ORDER_NUMBER = ?                 "
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = var_cadena
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rsaux6!order_number))
                                .Parameters.Append parametro
                           End With
                           Set rsaux7 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           If Not rsaux7.EOF Then
                              var_contingencia = 0
                              If var_contingencia = 1 Then
                              VAR_CORREO = ""
                              Else
                              VAR_CORREO = IIf(IsNull(rsaux7!Email), "", rsaux7!Email)
                              End If
                              If VAR_CORREO <> "" Then
                                 ' SE ENVIA CORREO A TIENDA
                                 rs.Open "select * from TB_ORACLE_PEDIDOS_CERRADOS_CN where pedido = " + CStr(rsaux6!order_number) + " and fecha_fin is null order by pedido desc", cnn, adOpenDynamic, adLockOptimistic
                                 If Not rs.EOF Then
                                    On Error GoTo salir2
                                    
                                    var_s = clnt2.CorreoAdjunto(VAR_CORREO, "Packing List del pedido " + CStr(rsaux6!order_number), "Se anexa packing list de pedido " + CStr(rsaux6!order_number) + " del CN. " + rsaux7!Description, CStr(rsaux6!order_number), "1002")
salir2:
                                    rsaux1.Open "update TB_ORACLE_PEDIDOS_CERRADOS_CN set FECHA_FIN = getdate() where pedido = " + CStr(rsaux6!order_number), cnn, adOpenDynamic, adLockOptimistic

                                 End If
                                 rs.Close
                              End If
                           End If
                           rsaux7.Close
                           rsaux6.MoveNext
                           Wend
                           rsaux6.Close

End Sub

Private Sub Command16_Click()
Dim codigo As String
Dim cantidad As Double

On Error GoTo e
Open "c:\sistemas\texto.txt" For Input As #1


Do While Not EOF(1)
   Line Input #1, sCadena1
   palabras = Split(sCadena1, " ")
   MsgBox "codigo: " + Trim(palabras(0)) + "   cantidad: " + Trim(palabras(1))
Loop
Close #1
e:
End Sub

Private Sub Command17_Click()
   rs.Open "SELECT distinct b.segment1, cross_reference, nvl(a.attribute1,1) as cantidad FROM mtl_cross_references_b A, xxvia_system_items_b B WHERE A.inventory_item_id = B.inventory_item_id AND B.organization_id = 89 and nvl(CLASIFICACIONSAT,' ') <>' '", cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux.Open "insert into sqlquezada2.sid2.dbo.tb_equivalencias (vcha_codigo_equivalente, vcha_Articulo_id, numb_cantidad) values ('" + rs!cross_reference + "','" + rs!SEGMENT1 + "'," + CStr(rs!cantidad) + ")", cnn, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub Command18_Click()
   rs.Open "SELECT LIST_HEADER_ID AS LISTA_ID, DESCRIPTION AS DESCRIPCION, START_dATE_ACTIVE FECHA_INICIO, END_DATE_ACTIVE FECHA_FIN, CURRENCY_CODE MONEDA FROM qp_secu_list_headers_v a WHERE LIST_HEADER_ID = 719011", cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux.Open "insert into sqlquezada2.sid2.dbo.tb_lista_precios (vcha_lista_precios_id, vcha_nombre_lista_precios, vcha_moneda) values ('" + rs!LISTA_ID + "','" + rs!Descripcion + "','" + CStr(rs!moneda) + "')", cnn, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
   rs.Open "SELECT LIST_HEADER_ID LISTA_id, B.SEGMENT1 codigo, OPERAND precio FROM qp_list_lines_v A, XXVIA_SYSTEM_ITEMS_B B WHERE LIST_HEADER_ID = 719011 AND PRODUCT_ATTR_VALUE = B.INVENTORY_ITEM_ID AND B.ORGANIZATION_ID = 89 and NVL(a.END_DATE_ACTIVE,SYSDATE) >= SYSDATE", cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux.Open "insert into sqlquezada2.sid2.dbo.tb_detalle_lista_precios (vcha_lista_precios_id, vcha_Articulo_id, floa_precio) values ('" + CStr(rs!LISTA_ID) + "','" + rs!codigo + "'," + CStr(rs!Precio) + ")", cnn, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
   MsgBox "Termino la carga de precios", vbOKOnly, "ATENCION"
End Sub

Private Sub Command19_Click()
    rs.Open "select * from fracciones_sabanas", cnn, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
          rsaux.Open "update xxvia_tb_complementos_pk_list set fraccion_arancelaria = '" + CStr(rs!fraccion) + "' where codigo = '" + rs!codigo + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
          rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub Command20_Click()
    Me.txt_codigo = Replace(Me.txt_codigo, "'", "-")
End Sub

Private Sub Command22_Click()
   rs.Open "select * from complementos_291220", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Me.Text1.Visible = True
         var_cadena = "update xxvia_tb_complementos_pk_list set fraccion_arancelaria = '" + CStr(rs!nueva) + "' where  codigo = '" + rs!codigo + "'"
         Me.Text1 = var_cadena
         
         rsaux.Open "update xxvia_tb_complementos_pk_list set fraccion_arancelaria = '" + CStr(rs!nueva) + "' where codigo = '" + rs!codigo + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         

         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub Command23_Click()
   rsaux10.Open "select * from ubicacione_borrar_110321", cnn, adOpenDynamic, adLockOptimistic
   While Not rsaux10.EOF
         rs.Open "select codigo, ubicacion from xxvia_Tb_ubicaciones where codigo = '" + rsaux10!codigo + "' and numero = 1", cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Me.txt_codigo = rs!codigo
              strconsulta = "UPDATE mtl_system_items_b SET  attribute2 = ? where  segment1 = ? and organization_id = ? "
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rs!ubicacion)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                    .Parameters.Append parametro
               End With
               Set rsaux9 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
         
         rs.MoveNext
   Wend
   rs.Close
   rsaux10.MoveNext
   Wend
End Sub

Private Sub Command24_Click()
    rs.Open "select 337043 as embarque, inventory_item_id, Description,  pedido, caja, codigo, sum(cantidad) as cantidad from xxvia_Tb_bitacora_lectura, xxvia_system_items_b  where pedido = 664251 and caja in(117, 131,132, 134, 136, 144, 145, 146, 147,148) and codigo = segment1 and organization_id = 93 group by 337043, inventory_item_id, Description, pedido, caja, codigo order by caja", cnnoracle_4, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
          If rsaux.State = 1 Then
             rsaux.Close
          End If
          rsaux.Open "select * from tb_oracle_Cajas_aduana where embarque = 337043 and numero_Caja = " + CStr(rs!Caja), cnn, adOpenDynamic, adLockOptimistic
          var_tipo_Caja = rsaux!TIPO_EMPAQUE
          rsaux.Close
          var_cadena = "insert into xxvia_Tb_salidas_Cajas (inte_emb_embarque, source_header_number, inte_paq_Caja, segment1, inventory_item_id, item_description, caja_pedido, floa_sal_cantidad_leida, tipo_caja)"
          var_cadena = var_cadena + " values (337043, 664251, " + CStr(rs!Caja) + ",'" + rs!codigo + "', " + CStr(rs!inventory_item_id) + ",'" + rs!Description + "', " + CStr(rs!Caja) + "," + CStr(rs!cantidad) + ",'" + var_tipo_Caja + "' )"
          rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
          rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub Command25_Click()
    rs.Open "select * from clientes_textilera311221", cnn, adOpenDynamic, adLockOptimistic
    var_i = 50
    While Not rs.EOF
          var_i = var_i + 1
          'rsaux.Open "INSERT INTO xxvia_TB_CHOFERES (id_chofer, nombre,rfc, licencia ) VALUES (" + CStr(var_i) + ",'" + rs!NOMBRE + "')", cnn, adOpenDynamic, adLockOptimistic
          
          'var_cadena = "UPDATE XXVIA_TB_ANES_CARTA_PORTE SET DISTANCIA_MTY = " + CStr(rs!KM) + " WHERE CLAVE = '" + CStr(rs!clave) + "'"
          'var_cadena = var_cadena + " VALUES('" + rs!CLAVE + "','" + IIf(IsNull(rs!CALLE), "", rs!CALLE) + "','" + IIf(IsNull(rs!NUM_EXT), "", rs!NUM_EXT) + "','" + CStr(IIf(IsNull(rs!LOCALIDAD), "", rs!LOCALIDAD)) + "','" + CStr(IIf(IsNull(rs!MUNICIPIO), "", rs!MUNICIPIO)) + "','" + CStr(IIf(IsNull(rs!ESTADO), "", rs!ESTADO)) + "','MX','" + CStr(IIf(IsNull(rs!CODIGO_POSTAL), "", rs!CODIGO_POSTAL)) + "','" + IIf(IsNull(rs!colonia), "", rs!colonia) + "')"
          'var_cadena = var_cadena + " VALUES('" + CStr(rs!site_use_id) + "','" + IIf(IsNull(rs!CALLE), "", rs!CALLE) + "','','" + CStr(IIf(IsNull(rs!LOCALIDAD), "", rs!LOCALIDAD)) + "','" + CStr(IIf(IsNull(rs!MUNICIPIO), "", rs!MUNICIPIO)) + "','" + CStr(IIf(IsNull(rs!ESTADO), "", rs!ESTADO)) + "','MX','" + CStr(IIf(IsNull(rs!POSTAL_code), "", rs!POSTAL_code)) + "','" + CStr(IIf(IsNull(rs!colonia), "", rs!colonia)) + "')"
          'var_cadena = var_cadena + " VALUES('" + rs!CLAVE + "','','','" + CStr(IIf(IsNull(rs!LOCALIDAD), "", rs!LOCALIDAD)) + "','" + CStr(IIf(IsNull(rs!MUNICIPIO), "", rs!MUNICIPIO)) + "','" + CStr(IIf(IsNull(rs!ESTADO), "", rs!ESTADO)) + "','MX','" + CStr(IIf(IsNull(rs!CODIGO_POSTAL), "", rs!CODIGO_POSTAL)) + "','" + CStr(IIf(IsNull(rs!colonia), "", rs!colonia)) + "')"
          'var_cadena = "UPDATE XXVIA_TB_choferes SET rfc = '" + CStr(IIf(IsNull(rs!rfc), "", rs!rfc)) + "', licencia = '" + IIf(IsNull(rs!licencia), "", rs!licencia) + "' WHERE id_chofer = '" + CStr(rs!clave) + "'"
          'var_cadena = "INSERT INTO XXVIA_TB_CHOFERES (ID_CHOFER, NOMBRE,RFC, LICENCIA) VALUES (" + CStr(var_i) + ",'" + rs!nombre + "','" + IIf(IsNull(rs!rfc), "", rs!rfc) + "','" + IIf(IsNull(rs!licencia), "", rs!licencia) + "')"
          var_cadena = "insert into XXVIA_TB_ANES_CARTA_PORTE VALUES(clave,calle,numero_exterior,numero_interior,rs!MUNICIPIO), "", rs!MUNICIPIO) + " ','" + CStr(IIf(IsNull(rs!ESTADO), "", rs!ESTADO)) + "','MX','" + CStr(IIf(IsNull(rs!codigo_postal), "", rs!codigo_postal)) + "','" + CStr(IIf(IsNull(rs!colonia), "", rs!colonia)) + "')"
          var_cadena = "insert into XXVIA_TB_ANES_CARTA_PORTE (clave,calle,numero_exterior,numero_interior,MUNICIPIO,ESTADO,pais,codigo_postal,colonia, distancia_ags)"
          var_cadena = var_cadena + " values ('" + CStr(rs!site_use_id) + "','" + rs!calle + "','" + CStr(IIf(IsNull(rs!num_ext), "", rs!num_ext)) + "','','" + CStr(IIf(IsNull(rs!municipio), "", rs!municipio)) + "','" + CStr(IIf(IsNull(rs!estado), "", rs!estado)) + "','MEX','" + CStr(IIf(IsNull(rs!codigo_postal), "", rs!codigo_postal)) + "','" + CStr(IIf(IsNull(rs!colonia), "", rs!colonia)) + "'," + CStr(rs!distancia) + ")"
          rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
          
          
          rs.MoveNext
    Wend
    rs.Close

End Sub

Private Sub Command26_Click()
    rs.Open "SELECT * FROM fracciones_arancelarias_091221", cnn, adOpenDynamic, adLockOptimistic
    var_i = 50
    While Not rs.EOF
          var_i = var_i + 1
          'rsaux.Open "INSERT INTO xxvia_TB_CHOFERES (id_chofer, nombre,rfc, licencia ) VALUES (" + CStr(var_i) + ",'" + rs!NOMBRE + "')", cnn, adOpenDynamic, adLockOptimistic
          
          'var_cadena = "INSERT INTO XXVIA_TB_ANES_CARTA_PORTE (CLAVE,CALLE, NUMERO_EXTERIOR, LOCALIDAD, MUNICIPIO, ESTADO, PAIS, CODIGO_POSTAL,colonia)"
          'var_cadena = var_cadena + " VALUES('" + rs!CLAVE + "','" + IIf(IsNull(rs!CALLE), "", rs!CALLE) + "','" + IIf(IsNull(rs!NUM_EXT), "", rs!NUM_EXT) + "','" + CStr(IIf(IsNull(rs!LOCALIDAD), "", rs!LOCALIDAD)) + "','" + CStr(IIf(IsNull(rs!MUNICIPIO), "", rs!MUNICIPIO)) + "','" + CStr(IIf(IsNull(rs!ESTADO), "", rs!ESTADO)) + "','MX','" + CStr(IIf(IsNull(rs!CODIGO_POSTAL), "", rs!CODIGO_POSTAL)) + "','" + IIf(IsNull(rs!colonia), "", rs!colonia) + "')"
          'var_cadena = var_cadena + " VALUES('" + CStr(rs!site_use_id) + "','" + IIf(IsNull(rs!CALLE), "", rs!CALLE) + "','','" + CStr(IIf(IsNull(rs!LOCALIDAD), "", rs!LOCALIDAD)) + "','" + CStr(IIf(IsNull(rs!MUNICIPIO), "", rs!MUNICIPIO)) + "','" + CStr(IIf(IsNull(rs!ESTADO), "", rs!ESTADO)) + "','MX','" + CStr(IIf(IsNull(rs!POSTAL_code), "", rs!POSTAL_code)) + "','" + CStr(IIf(IsNull(rs!colonia), "", rs!colonia)) + "')"
          'var_cadena = var_cadena + " VALUES('" + rs!CLAVE + "','','','" + CStr(IIf(IsNull(rs!LOCALIDAD), "", rs!LOCALIDAD)) + "','" + CStr(IIf(IsNull(rs!MUNICIPIO), "", rs!MUNICIPIO)) + "','" + CStr(IIf(IsNull(rs!ESTADO), "", rs!ESTADO)) + "','MX','" + CStr(IIf(IsNull(rs!CODIGO_POSTAL), "", rs!CODIGO_POSTAL)) + "','" + CStr(IIf(IsNull(rs!colonia), "", rs!colonia)) + "')"
          var_cadena = "UPDATE xxvia_tb_complementos_pk_list SET fraccion_arancelaria = '" + CStr(IIf(IsNull(rs!fraccion_arancelaria), "", rs!fraccion_arancelaria)) + "' WHERE codigo = '" + CStr(rs!codigo) + "'"
          'var_cadena = "INSERT INTO XXVIA_TB_CHOFERES (ID_CHOFER, NOMBRE,RFC, LICENCIA) VALUES (" + CStr(var_i) + ",'" + rs!NOMBRE + "','" + IIf(IsNull(rs!rfc), "", rs!rfc) + "','" + IIf(IsNull(rs!licencia), "", rs!licencia) + "')"
          
          rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
          rs.MoveNext
    Wend
    rs.Close


End Sub

Private Sub Command27_Click()
   Dim var_i As Double
   var_i = 0
   rs.Open "select * from xxvia_tb_anes_carta_porte order by nombre", cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux.Open "select razon_social_cliente from xxvia_vw_clientes_bcp where to_char(site_use_id) = '" + CStr(rs!clave) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            rsaux10.Open "update xxvia_tb_anes_carta_porte set nombre = '" + IIf(IsNull(rsaux!razon_social_cliente), "", rsaux!razon_social_cliente) + "'"
         End If
         rsaux.Close
         rs.MoveNext
         var_i = var_i + 1
         Me.Label2 = var_i
         Me.Refresh
         
   Wend
End Sub

Private Sub Command28_Click()
   
   Dim cn As New ADODB.Connection
   Dim DSN As String
   Dim cn2 As New ADODB.Connection
   DSN = "eflow"
   If cn.State = 1 Then
      cn.Close
   End If
   cn.Open ("DSN=" & DSN & ";")
   
   Set rs = cn.execute("SELECT factura, fecha FROM facturas where factura like 'CPV%'")
    Dim iFila As Long, iCol As Integer, i As Integer

 Set oexcel = CreateObject("Excel.Application")
 Set owbook = oexcel.Workbooks.Add
 Set osheet = owbook.Worksheets(1)
 osheet.Name = "ASIENTO"
 Screen.MousePointer = vbHourglass
 iFila = 1
 ifila2 = 1
 icol2 = 1
 iCol = 1
 rs.MoveFirst
 For i = 0 To rs.Fields.Count - 1
 ' pone el nombre de los campos en la primera fila
 osheet.Cells(iFila, i + 1) = rs.Fields(i).Name
 Next
 iFila = iFila + 1
 With osheet
 ' carga los registros del recordset
 .Cells(iFila, iCol).CopyFromRecordset rs
 'oexcel.Columns(7).Select
 'oexcel.Selection.NumberFormat = "#,##0.00"
 'oexcel.Columns(12).Select
 'oexcel.Selection.NumberFormat = "0.00"
 .Columns.AutoFit ' ajusta el ancho de las columnas

 End With
 owbook.SaveAs "c:\reportessid\primer_reporte_excel.xls"
 
 oexcel.Visible = True
 Set oexcel = Nothing
 Screen.MousePointer = vbDefault


   

End Sub

Private Sub Command3_Click()
'Dim oExcel As Excel.Application
 'Dim oWBook As Excel.Workbook
 'Dim oSheet As Excel.Worksheet
 Dim iFila As Long, iCol As Integer, i As Integer

 Set oexcel = CreateObject("Excel.Application")
 Set owbook = oexcel.Workbooks.Add
 Set osheet = owbook.Worksheets(1)
 osheet.Name = "ASIENTO"
 Screen.MousePointer = vbHourglass
 iFila = 1
 ifila2 = 1
 icol2 = 1
 iCol = 1
 rs.Open "select * from tb_articulos", cnn, adOpenDynamic, adLockOptimistic
 rs.MoveFirst
 For i = 0 To rs.Fields.Count - 1
 ' pone el nombre de los campos en la primera fila
 osheet.Cells(iFila, i + 1) = rs.Fields(i).Name
 Next
 iFila = iFila + 1
 With osheet
 ' carga los registros del recordset
 .Cells(iFila, iCol).CopyFromRecordset rs
 oexcel.Columns(7).Select
 oexcel.Selection.NumberFormat = "#,##0.00"
 oexcel.Columns(12).Select
 oexcel.Selection.NumberFormat = "0.00"
 .Columns.AutoFit ' ajusta el ancho de las columnas

 End With
 owbook.SaveAs "c:\reportessid\primer_reporte_excel.xls"
 
 oexcel.Visible = True
 Set oexcel = Nothing
 Screen.MousePointer = vbDefault

End Sub

Private Sub Command4_Click()
   'x = Shell(App.Path + "/envia_texto.exe 10.6.200.70, 9001, PRUEBA DE SISTEMAS " + CStr(Now))}
   'cnn_minegocio.Open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=minegocio.vianney.mx:22;DATABASE=vianney_servicios;UID=UserSer;PWD=3Hty6r2+; OPTION=3"
    cnn_minegocio.Open "DRIVER={MySQL ODBC 3.51 Driver};DATABASE=vianney_servicios;SERVER=minegocio.vianney.mx:22;UID=UserSer;password=3Hty6r2+;PORT=3306;"
    cnn_minegocio.Open "Data Source=MiNegocioProduccion;Initial Catalog=vianney_servicios;User ID=UserSer;Password=3Hty6r2+name=MiNegocio"

   rs.Open "SELECT * FROM vianney_servicios.xxvia_tb_pedido_mayoreo where status=4 and day_shipping=''", cnn_minegocio, adOpenDynamic, adLockOptimistic
   cnn.Close
End Sub

Private Sub Command2_Click()
rs.Open "delete from TB_ORACLE_REVALUACION_DEVOLUCIONES", cnn, adOpenDynamic, adLockOptimistic
'var_cadena = "select A.ct_reference, A.TRX_NUMBER AS NC, C.TITULAR, C.INVENTORY_ITEM_ID,B.ORDER_NUMBER PEDIDO, NUMERO, FECHA_FIN FECHA,CODIGO, CANTIDAD, ESTATUS, REFERENCIA, PRECIO, PRECIO_ENTERO, FACTURA, DESCUENTO_FINANCIERO, NOTA_CREDITO_DESC_FIN,D.TRX_NUMBER, D.TRX_DATE"
'var_cadena = var_cadena + " from ra_customer_Trx_all A, OE_ORDER_HEADERS_ALL B, XXVIA.XXVIA_TB_DEVOLUCIONES_CLIENTES C, RA_CUSTOMER_TRX_ALL D"
'var_cadena = var_cadena + " "
'var_cadena = var_cadena + " where A.trx_number IN ('27847','27873','27874','27883','27884','27885','27886','27898','27899','27909','27975','27976','29183','29184','29190','29191','29192','29194','29204','29205','29206','29207','29208','29210','29212','29213','29216','29223','29225','29226','29227','29237','29238','29260','29327','29333','29334','29335','29841','29860','29861','29872','29873','30181','30396','30406','30420','30524','31205','31206','31209','31210','31211','31212','31213','31216','31217','31218','31219','31225','31226','31228','31241','31242','31243','31244','31247','31248','31249','31251','31252','31253','31254','31256','31257','31258','31259','31260','31268','31269','31341','31342','31561','31562','31563','31564','31565','31566','31568','31569','31570','31571','31572','31573','31586','33039','33040','33041','33042','33043','33044','33045','33046','33260','33310','33047','33048','33414','33457','33458','33476','34224','34226','34227',"
'var_cadena = var_cadena + " '34228','34229','34251','34273','34274','34292','34293','34297','34298','34313','34581')"
'var_cadena = var_cadena + " and A.SOLD_TO_CUSTOMER_ID =  4089 AND A.CT_REFERENCE = B.ORDER_NUMBER AND TO_NUMBER(REPLACE(ORIG_SYS_DOCUMENT_REF,'SIDDC_','')) = C.NUMERO AND C.FACTURA = D.CUSTOMER_TRX_ID"

var_cadena = "select A.ct_reference, A.TRX_NUMBER AS NC, C.TITULAR, C.INVENTORY_ITEM_ID,B.ORDER_NUMBER PEDIDO, NUMERO, FECHA_FIN FECHA,CODIGO, f.quantity_ordered * (-1) CANTIDAD, ESTATUS, REFERENCIA, PRECIO, PRECIO_ENTERO, FACTURA, DESCUENTO_FINANCIERO, NOTA_CREDITO_DESC_FIN,D.TRX_NUMBER, D.TRX_DATE"
var_cadena = var_cadena + " from ra_customer_Trx_all A, OE_ORDER_HEADERS_ALL B, XXVIA.XXVIA_TB_DEVOLUCIONES_CLIENTES C, RA_CUSTOMER_TRX_ALL D, ra_customer_trx_lines_all f "
var_cadena = var_cadena + " where A.trx_number IN ('27847','27873','27874','27883','27884','27885','27886','27898','27899','27909',"
var_cadena = var_cadena + " '27975','27976','29183','29184','29190','29191','29192','29194','29204','29205','29206','29207','29208','29210','29212','29213','29216','29223','29225','29226','29227','29237','29238','29260','29327','29333','29334','29335','29841','29860','29861','29872','29873','30181','30396','30406','30420','30524','31205','31206','31209','31210','31211','31212','31213','31216','31217','31218','31219','31225','31226','31228','31241','31242','31243','31244','31247','31248','31249','31251','31252','31253','31254','31256','31257','31258','31259','31260','31268','31269','31341','31342','31561','31562','31563','31564','31565','31566','31568','31569','31570','31571','31572','31573','31586','33039','33040','33041','33042','33043','33044','33045','33046','33260','33310','33047','33048','33414','33457','33458','33476','34224','34226','34227','34228','34229','34251','34273','34274','34292','34293','34297','34298','34313','34581')"
 var_cadena = var_cadena + " and A.SOLD_TO_CUSTOMER_ID =  4089 AND A.CT_REFERENCE = B.ORDER_NUMBER AND TO_NUMBER(REPLACE(ORIG_SYS_DOCUMENT_REF,'SIDDC_','')) = C.NUMERO AND C.FACTURA = D.CUSTOMER_TRX_ID and f.customer_trx_id = a.customer_Trx_id and c.inventory_item_id = f.INVENTORY_ITEM_ID order by a.trx_number"


rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic

While Not rs.EOF
      var_cadena = "INSERT INTO TB_ORACLE_REVALUACION_DEVOLUCIONES ([INVENTORY_ITEM_ID],[NOTA_CREDITO],[TITULAR],[PEDIDO], [NUMERO] ,[FECHA] ,[CODIGO] ,[CANTIDAD] ,[ESTATUS] ,[REFERENCIA] ,[PRECIO] ,[PRECIO_ENTERIO] ,[FACTURA] ,[CUSTOMER_TRX_ID] ,[TRX_DATE] ,[DESCUENTO_FINANCIARO] ,[NC_DF] ,[CUSTOMER_TRX_ID_NUEVO] ,[TRX_NUMBER_NUEVO] ,[TRX_DATE_NUEVO]  ,[DESCUENTO_FINANCIERO_NUEVO]  ,[NC_DF_NUEVO])"
      var_cadena = var_cadena + " Values (" + CStr(rs!inventory_item_id) + ",'" + rs!NC + "'," + CStr(rs!TITULAR) + ",  " + CStr(rs!pedido) + "," + CStr(rs!numero) + ",'" + CStr(rs!Fecha) + "' ,'" + rs!codigo + "'," + CStr(rs!cantidad) + " ,'" + rs!estatus + "','" + rs!Referencia + "'," + CStr(rs!Precio) + ", " + CStr(rs!PRECIO_ENTERO) + " ,'" + CStr(rs!trx_number) + "'," + CStr(rs!FACTURA) + ",'" + CStr(rs!trx_date) + "'," + CStr(rs!DESCUENTO_FINANCIERO) + "," + CStr(rs!NOTA_CREDITO_DESC_FIN) + ",0,'', NULL, 0, 0)"
      'MsgBox var_cadena
      
      
      rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
      rs.MoveNext
Wend
rs.Close
rsaux9.Open "select * from TB_ORACLE_REVALUACION_DEVOLUCIONES", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux9.EOF
                        var_numero_factura = 0
                        VAR_PORCENTAJE_FIN = 0
                        VAR_NOTA_CREDITO_DF = 0
                        If rsaux.State = 1 Then
                           rsaux.Close
                        End If
                        x = 1
                        var_dia = CStr(Day(rsaux9!Fecha))
                        var_mes = CStr(Month(rsaux9!Fecha))
                        var_año = CStr(Year(rsaux9!Fecha))
                        If Len(var_dia) < 2 Then
                           var_dia = "0" + var_dia
                        End If
                        If Len(var_mes) < 2 Then
                           var_mes = "0" + var_mes
                        End If
                        If Len(var_año) = 2 Then
                           var_año = "20" + var_año
                        End If
                        var_fecha_devolucion = var_dia + "/" + var_mes + "/" + var_año
                        ' esto no recuerdo porque se comento
                        'rsaux.Open "SELECT L.CUSTOMER_TRX_ID, TRX_DATE, TRX_NUMBER, E.sold_to_customer_id,  L.inventory_item_id, NVL(SALES_ORDER_LINE,ROWNUM) AS LINEA, SUM(NVL(L.GROSS_unit_selling_price,l.unit_selling_price)) AS PRECIO FROM RA_CUSTOMER_TRX_LINES_ALL L, RA_CUSTOMER_TRX_ALL E, ra_cust_trx_types_all TYPES Where TYPES.TYPE = 'INV' AND TYPES.cust_trx_type_id = E.cust_trx_type_id AND TYPES.org_id = E.org_id AND l.customer_trx_id = E.customer_trx_id AND L.inventory_item_id = " + CStr(rsaux9!INVENTORY_ITEM_ID) + " AND E.sold_to_customer_id = " + CStr(rsaux9!TITULAR) + "  GROUP BY L.CUSTOMER_TRX_ID, TRX_DATE, TRX_NUMBER, E.sold_to_customer_id,  L.inventory_item_id, NVL(SALES_ORDER_LINE,ROWNUM) ORDER BY trx_date desc", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        rsaux.Open "SELECT L.CUSTOMER_TRX_ID, TRX_DATE, TRX_NUMBER, E.sold_to_customer_id,  L.inventory_item_id, SUM(NVL(L.GROSS_unit_selling_price,l.unit_selling_price)) AS PRECIO FROM RA_CUSTOMER_TRX_LINES_ALL L, RA_CUSTOMER_TRX_ALL E, ra_cust_trx_types_all TYPES Where TYPES.TYPE = 'INV' AND TYPES.cust_trx_type_id = E.cust_trx_type_id AND TYPES.org_id = E.org_id AND l.customer_trx_id = E.customer_trx_id AND L.inventory_item_id = " + CStr(rsaux9!inventory_item_id) + " AND E.sold_to_customer_id = " + CStr(rsaux9!TITULAR) + " and trx_date >= to_date('01/01/2016','DD/MM/YYYY') and trx_date <= to_date('" + var_fecha_devolucion + "','DD/MM/YYYY') GROUP BY L.CUSTOMER_TRX_ID, TRX_DATE, TRX_NUMBER, E.sold_to_customer_id,  L.inventory_item_id  ORDER BY trx_date ASC", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
''''''


                        'rsaux.Open "SELECT L.CUSTOMER_TRX_ID, TRX_DATE, TRX_NUMBER, E.sold_to_customer_id,  L.inventory_item_id, SUM(NVL(L.GROSS_unit_selling_price,l.unit_selling_price)) AS PRECIO FROM RA_CUSTOMER_TRX_LINES_ALL L, RA_CUSTOMER_TRX_ALL E, ra_cust_trx_types_all TYPES Where TYPES.TYPE = 'INV' AND TYPES.cust_trx_type_id = E.cust_trx_type_id AND TYPES.org_id = E.org_id AND l.customer_trx_id = E.customer_trx_id AND L.inventory_item_id = " + CStr(rsaux9!inventory_item_id) + " AND E.sold_to_customer_id = " + CStr(rsaux9!TITULAR) + "  GROUP BY L.CUSTOMER_TRX_ID, TRX_DATE, TRX_NUMBER, E.sold_to_customer_id,  L.inventory_item_id  ORDER BY trx_date desc", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           var_veces = 1
                           If rsaux7.State = 1 Then
                              rsaux7.Close
                           End If
                           rsaux7.Open "select count(*) from RA_CUSTOMER_TRX_LINES_ALL where customer_trx_id= " + CStr(rsaux!customer_Trx_id) + " and inventory_item_id = " + CStr(rsaux9!inventory_item_id) + " and unit_selling_price >0", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If Not rsaux7.EOF Then
                              var_veces = IIf(IsNull(rsaux7(0).Value), 1, rsaux7(0).Value)
                           End If
                           rsaux7.Close
                           If rsaux!Precio = 0 Then
                              var_precio = 0
                           Else
                              var_precio = rsaux!Precio / var_veces
                           End If
                           
                           
                           If rsaux!Precio = 0 Then
                              var_precio_entero = 0
                           Else
                              var_precio_entero = rsaux!Precio / var_veces
                           End If
                           var_numero_factura = rsaux!customer_Trx_id
                           var_factura_trx_number = rsaux!trx_number
                           var_fecha_factura_nueva = rsaux!trx_date
                           rsaux5.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           
                           
                           
                           If var_attribute10 = "0" Then
                              var_cadena = "SELECT ARPA.APPLIED_CUSTOMER_TRX_ID AS FACTURA_ID, ARPA.CUSTOMER_TRX_ID AS NOTA_CREDITO_ID, ARPA.ACCTD_AMOUNT_APPLIED_TO AS MONTO_APLICADO, RCT.CUST_TRX_TYPE_ID, RCTL.ATTRIBUTE11, RCTL.ATTRIBUTE10, ARPA.AMOUNT_APPLIED, acr.amount FROM AR_RECEIVABLE_APPLICATIONS_ALL ARPA, RA_CUSTOMER_TRX_ALL RCT, RA_CUSTOMER_TRX_LINES_ALL RCTL, ar_cash_receipts_all acr WHERE ARPA.APPLICATION_TYPE = 'CM' AND ARPA.CUSTOMER_TRX_ID = RCT.CUSTOMER_TRX_ID AND RCT.CUST_TRX_TYPE_ID IN (SELECT ATTRIBUTE2 From RA_CUST_TRX_TYPES_ALL WHERE ATTRIBUTE2 IS NOT NULL) AND ARPA.CUSTOMER_TRX_ID  = RCTL.CUSTOMER_TRX_ID AND RCTL.ATTRIBUTE11 IS NOT NULL AND ARPA.APPLIED_CUSTOMER_TRX_ID = " + CStr(var_numero_factura) + " and RCTL.ATTRIBUTE10 = acr.cash_receipt_id and ARPA.ACCTD_AMOUNT_APPLIED_TO > 0 order by arpa.last_update_date desc"
                              rsaux5.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                              var_attribute10 = ""
                              If Not rsaux5.EOF Then
                                 While Not rsaux5.EOF
                                      If var_attribute10 = "" Then
                                         var_attribute10 = rsaux5!attribute10
                                      Else
                                         var_attribute10 = var_attribute10 + "," + rsaux5!attribute10
                                      End If
                                       rsaux5.MoveNext
                                 Wend
                              Else
                                 var_attribute10 = 0
                              End If
                              rsaux5.Close
                           End If
                           
                           
                           var_cadena = "select rec.CUSTOMER_TRX_ID, nvl(sum(rec.amount_applied),0) as importe_df from ar_receivable_applications_all rec Inner join ar_payment_schedules_all pay on rec.payment_schedule_id = pay.payment_schedule_id Inner join ra_cust_trx_types_all on pay.cust_trx_type_id = ra_cust_trx_types_all.cust_trx_type_id Where rec.applied_customer_trx_id = " + CStr(var_numero_factura) + " and rec.apply_date < sysdate and rec.display = 'Y' and application_type = 'CM' and ra_cust_trx_types_all.cust_trx_type_id in (1564,1028) group by rec.CUSTOMER_TRX_ID "
                           rsaux5.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If Not rsaux5.EOF Then
                              var_importe_total_df = 0
                              var_notas_credito_df = ""
                              While Not rsaux5.EOF
                                    var_importe_total_df = var_importe_total_df + IIf(IsNull(rsaux5!importe_df), 0, rsaux5!importe_df)
                                    If var_notas_credito_df = "" Then
                                       var_notas_credito_df = CStr(rsaux5!customer_Trx_id)
                                    Else
                                       var_notas_credito_df = var_notas_credito_df + ", " + CStr(rsaux5!customer_Trx_id)
                                    End If
                                    rsaux5.MoveNext
                              Wend
                              rsaux5.MoveFirst
                              'var_cadena = "select sum(amount_applied) amount_applied from ar_receivable_applications_all Where applied_customer_trx_id = " + CStr(VAR_NUMERO_fACTURA) + " and display = 'Y' and application_type = 'CASH' and cash_receipt_id in( " + CStr(var_attribute10) + ")"
                              var_cadena = "select SUM(nvl(gross_extended_amount, extended_amount)) AS amount_applied from ra_customer_trx_lines_all where customer_trx_id = " + CStr(var_numero_factura) + " and line_type = 'LINE'"
                              rsaux6.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux6.EOF Then
                                 'var_importe_total = IIf(IsNull(rsaux6!amount_applied), 0, rsaux6!amount_applied) + var_importe_total_df
                                 VAR_IMPORTE_TOTAL = IIf(IsNull(rsaux6!amount_applied), 0, rsaux6!amount_applied)
                                 If VAR_IMPORTE_TOTAL = 0 Then
                                    VAR_PORCENTAJE_FIN = 0
                                 Else
                                    'VAR_PORCENTAJE_FIN = 100 - (Round((IIf(IsNull(rsaux6!amount_applied), 0, rsaux6!amount_applied) * 100) / var_importe_total, 2))
                                    VAR_PORCENTAJE_FIN = (Round((IIf(IsNull(var_importe_total_df), 0, var_importe_total_df) * 100) / VAR_IMPORTE_TOTAL, 2))
                                 End If
                                 VAR_NOTA_CREDITO_DF = var_notas_credito_df
                                 var_precio = var_precio * (1 - (IIf(IsNull(VAR_PORCENTAJE_FIN), 0, VAR_PORCENTAJE_FIN) / 100))
                              End If
                              rsaux6.Close
                           End If
                           rsaux5.Close
                        Else
                           'MsgBox "SELECT PRODUCT_UOM_CODE, CURRENCY_CODE  * FROM qp_list_lines_v A, qp_list_HEADERS_v B WHERE A.list_header_id = B.list_header_id AND  A.list_header_id = " + CStr(var_clave_lista_precios) + " and  A.product_attr_value = " + CStr(rsaux9!INVENTORY_ITEM_ID) + " AND A.product_attr_val_disp = '" + rsaux9!CODIGO + "'"
                           If rsaux11.State = 1 Then
                              rsaux11.Close
                           End If
                           rsaux11.Open "SELECT PRODUCT_UOM_CODE, CURRENCY_CODE  FROM qp_list_lines_v A, qp_list_HEADERS_v B WHERE A.list_header_id = B.list_header_id AND  A.list_header_id = " + CStr(var_clave_lista_precios) + " and  A.product_attr_value = " + CStr(rsaux9!inventory_item_id) + " AND A.product_attr_val_disp = '" + rsaux9!codigo + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           var_unidad_medida = rsaux11!PRODUCT_UOM_CODE
                           VAR_CURRENCY = rsaux11!CURRENCY_CODE
                           rsaux11.Close
                           VAR_ZZ = 0
                           If var_unidad_organizacional = "93" And VAR_ZZ = 1 Then
                              objConn.Open var_conexion_oracle
                              '… Establecer conexión a la base de datos con el objeto objConn.
                              With objCmd
                                   objConn.BeginTrans
                                   .ActiveConnection = objConn
                                   .CommandText = "APPS.xxvia_descuento_linea"
                                   .CommandType = adCmdStoredProc
                                
                                   rsaux10.Open "select * from tb_unidadesorganizacionales where vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                                   If Not rsaux10.EOF Then
                                      var_responsabilidad_facturacion = IIf(IsNull(rsaux10!RESPONSABILIDAD_FACTURACION), "", rsaux10!RESPONSABILIDAD_FACTURACION)
                                   End If
                                   rsaux10.Close
                                         
                                         
                                   Set objParm = .CreateParameter("p_responsabilidad", adVarChar, adParamInput, 100, var_responsabilidad_facturacion)
                                   .Parameters.Append objParm
      
                                        
                                   Set objParm = .CreateParameter("p_org_id", adNumeric, adParamInput, 50, CDbl(var_empresa))
                                   .Parameters.Append objParm
                                   
                                   Set objParm = .CreateParameter("p_uom", adVarChar, adParamInput, 50, var_unidad_medida)
                                   .Parameters.Append objParm
                                
                                   Set objParm = .CreateParameter("p_currency", adVarChar, adParamInput, 50, VAR_CURRENCY)
                                   .Parameters.Append objParm
                                
                                   var_estatus_factura = ""
                                   Set objParm = .CreateParameter("p_inventory_item_id", adNumeric, adParamInput, 50, rsaux9!inventory_item_id)
                                   .Parameters.Append objParm
                                   
                                   Set objParm = .CreateParameter("p_price_list_id", adNumeric, adParamInput, 50, var_clave_lista_precios)
                                   .Parameters.Append objParm
                                   
                                   Set objParm = .CreateParameter("p_cust_account_id", adNumeric, adParamInput, 50, var_clave_titular)
                                   .Parameters.Append objParm
                                
                                   Set objParm = .CreateParameter("xx_unit_price", adNumeric, adParamOutput, 50, var_precio_inflado)
                                   .Parameters.Append objParm
                                   
                                   Set objParm = .CreateParameter("xx_adjusted_price", adNumeric, adParamOutput, 50, var_precio_descuento)
                                   .Parameters.Append objParm
                                   
                                   rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                   rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                   On Error GoTo salir2:
                                   .execute
                                 
                                   var_precio_inflado = .Parameters("xx_unit_price").Value
                                   var_precio_entero = var_precio_inflado
                                   var_precio_descuento = .Parameters("xx_adjusted_price").Value
                                   var_precio = var_precio_descuento
                                   objConn.CommitTrans
                              End With
                              Set objConn = Nothing
                              Set objCmd = Nothing
                           Else
                              x = 0
                              If x = 0 Then
                                 If rsaux10.State = 1 Then
                                    rsaux10.Close
                                 End If
                                 rsaux11.Open "select * from xxvia_system_items_b where segment1 = '" + rsaux9!codigo + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 If Not rsaux11.EOF Then
                                    rsaux10.Open "SELECT OPERAND FROM qp_list_lines_v WHERE list_header_id =  " + CStr(var_clave_lista_precios) + " AND  PRODUCT_ATTR_VAL_DISP = '" + rsaux9!codigo + "' and start_date_active <= sysdate and (end_date_active is null or end_date_active >= sysdate) and Product_Attr_Value = " + CStr(rsaux11!inventory_item_id), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                    'rsaux10.Open "SELECT OPERAND FROM qp_list_lines_v WHERE list_header_id =  " + CStr(var_clave_lista_precios) + " AND  PRODUCT_ATTR_VAL_DISP = '" + rsaux9!codigo + "' AND Product_Attr_Value = " + CStr(rsaux11!INVENTORY_ITEM_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 Else
                                    rsaux10.Open "SELECT OPERAND FROM qp_list_lines_v WHERE list_header_id =  " + CStr(var_clave_lista_precios) + " AND  PRODUCT_ATTR_VAL_DISP = '" + rsaux9!codigo + "' and start_date_active <= sysdate and (end_date_active is null or end_date_active >= sysdate) and ", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 End If
                                 rsaux11.Close
                                 Product_Attr_Value = 16217
                                 var_precio_entero = IIf(IsNull(rsaux10!OPERAND), 0, rsaux10!OPERAND)
                                 var_precio = IIf(IsNull(rsaux10!OPERAND), 0, rsaux10!OPERAND)
                                 rsaux10.Close
                                 VAR_DESCUENTO = 0
                                 rsaux11.Open "SELECT DISTINCT(list_header_id) as calificador FROM qp_qualifiers_v WHERE list_header_id IS NOT NULL", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 While Not rsaux11.EOF
                                       rsaux10.Open "select xxvia_fn_descuento_titular(" + CStr(rsaux11!calificador) + ",'" + CStr(rsaux9!TITULAR) + "') as descuento from dual ", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                       If Not rsaux10.EOF Then
                                          If rsaux10!DESCUENTO > VAR_DESCUENTO Then
                                             VAR_DESCUENTO = rsaux10!DESCUENTO
                                          End If
                                       End If
                                       rsaux10.Close
                                       rsaux11.MoveNext
                                 Wend
                                 rsaux11.Close
                                 var_precio = var_precio * (1 - (VAR_DESCUENTO / 100))
                              Else
                                 var_precio_entero = IIf(IsNull(rsaux9!Precio), 0, rsaux9!Precio)
                                 var_precio = IIf(IsNull(rsaux9!Precio), 0, rsaux9!Precio)
                              End If
                           End If
                       End If






                        
''''''
                        Else
                        rsaux.Close
                        rsaux.Open "SELECT L.CUSTOMER_TRX_ID, TRX_DATE, TRX_NUMBER, E.sold_to_customer_id,  L.inventory_item_id, SUM(NVL(L.GROSS_unit_selling_price,l.unit_selling_price)) AS PRECIO FROM RA_CUSTOMER_TRX_LINES_ALL L, RA_CUSTOMER_TRX_ALL E, ra_cust_trx_types_all TYPES Where TYPES.TYPE = 'INV' AND TYPES.cust_trx_type_id = E.cust_trx_type_id AND TYPES.org_id = E.org_id AND l.customer_trx_id = E.customer_trx_id AND L.inventory_item_id = " + CStr(rsaux9!inventory_item_id) + " AND E.sold_to_customer_id = " + CStr(rsaux9!TITULAR) + " AND trx_date <= to_date('" + var_fecha_devolucion + "','DD/MM/YYYY')  GROUP BY L.CUSTOMER_TRX_ID, TRX_DATE, TRX_NUMBER, E.sold_to_customer_id,  L.inventory_item_id  ORDER BY trx_date desc", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           var_veces = 1
                           If rsaux7.State = 1 Then
                              rsaux7.Close
                           End If
                           rsaux7.Open "select count(*) from RA_CUSTOMER_TRX_LINES_ALL where customer_trx_id= " + CStr(rsaux!customer_Trx_id) + " and inventory_item_id = " + CStr(rsaux9!inventory_item_id) + " and unit_selling_price >0", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If Not rsaux7.EOF Then
                              var_veces = IIf(IsNull(rsaux7(0).Value), 1, rsaux7(0).Value)
                           End If
                           rsaux7.Close
                           If rsaux!Precio = 0 Then
                              var_precio = 0
                           Else
                              var_precio = rsaux!Precio / var_veces
                           End If
                           
                           
                           If rsaux!Precio = 0 Then
                              var_precio_entero = 0
                           Else
                              var_precio_entero = rsaux!Precio / var_veces
                           End If
                           var_numero_factura = rsaux!customer_Trx_id
                           var_factura_trx_number = rsaux!trx_number
                           var_fecha_factura_nueva = rsaux!trx_date
                           
                           rsaux5.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           
                           
                           
                           If var_attribute10 = "0" Then
                              var_cadena = "SELECT ARPA.APPLIED_CUSTOMER_TRX_ID AS FACTURA_ID, ARPA.CUSTOMER_TRX_ID AS NOTA_CREDITO_ID, ARPA.ACCTD_AMOUNT_APPLIED_TO AS MONTO_APLICADO, RCT.CUST_TRX_TYPE_ID, RCTL.ATTRIBUTE11, RCTL.ATTRIBUTE10, ARPA.AMOUNT_APPLIED, acr.amount FROM AR_RECEIVABLE_APPLICATIONS_ALL ARPA, RA_CUSTOMER_TRX_ALL RCT, RA_CUSTOMER_TRX_LINES_ALL RCTL, ar_cash_receipts_all acr WHERE ARPA.APPLICATION_TYPE = 'CM' AND ARPA.CUSTOMER_TRX_ID = RCT.CUSTOMER_TRX_ID AND RCT.CUST_TRX_TYPE_ID IN (SELECT ATTRIBUTE2 From RA_CUST_TRX_TYPES_ALL WHERE ATTRIBUTE2 IS NOT NULL) AND ARPA.CUSTOMER_TRX_ID  = RCTL.CUSTOMER_TRX_ID AND RCTL.ATTRIBUTE11 IS NOT NULL AND ARPA.APPLIED_CUSTOMER_TRX_ID = " + CStr(var_numero_factura) + " and RCTL.ATTRIBUTE10 = acr.cash_receipt_id and ARPA.ACCTD_AMOUNT_APPLIED_TO > 0 order by arpa.last_update_date desc"
                              rsaux5.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                              var_attribute10 = ""
                              If Not rsaux5.EOF Then
                                 While Not rsaux5.EOF
                                      If var_attribute10 = "" Then
                                         var_attribute10 = rsaux5!attribute10
                                      Else
                                         var_attribute10 = var_attribute10 + "," + rsaux5!attribute10
                                      End If
                                       rsaux5.MoveNext
                                 Wend
                              Else
                                 var_attribute10 = 0
                              End If
                              rsaux5.Close
                           End If
                           
                           
                           var_cadena = "select rec.CUSTOMER_TRX_ID, nvl(sum(rec.amount_applied),0) as importe_df from ar_receivable_applications_all rec Inner join ar_payment_schedules_all pay on rec.payment_schedule_id = pay.payment_schedule_id Inner join ra_cust_trx_types_all on pay.cust_trx_type_id = ra_cust_trx_types_all.cust_trx_type_id Where rec.applied_customer_trx_id = " + CStr(var_numero_factura) + " and rec.apply_date < sysdate and rec.display = 'Y' and application_type = 'CM' and ra_cust_trx_types_all.cust_trx_type_id in (1564,1028) group by rec.CUSTOMER_TRX_ID "
                           rsaux5.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If Not rsaux5.EOF Then
                              var_importe_total_df = 0
                              var_notas_credito_df = ""
                              While Not rsaux5.EOF
                                    var_importe_total_df = var_importe_total_df + IIf(IsNull(rsaux5!importe_df), 0, rsaux5!importe_df)
                                    If var_notas_credito_df = "" Then
                                       var_notas_credito_df = CStr(rsaux5!customer_Trx_id)
                                    Else
                                       var_notas_credito_df = var_notas_credito_df + ", " + CStr(rsaux5!customer_Trx_id)
                                    End If
                                    rsaux5.MoveNext
                              Wend
                              rsaux5.MoveFirst
                              'var_cadena = "select sum(amount_applied) amount_applied from ar_receivable_applications_all Where applied_customer_trx_id = " + CStr(VAR_NUMERO_fACTURA) + " and display = 'Y' and application_type = 'CASH' and cash_receipt_id in( " + CStr(var_attribute10) + ")"
                              var_cadena = "select SUM(nvl(gross_extended_amount, extended_amount)) AS amount_applied from ra_customer_trx_lines_all where customer_trx_id = " + CStr(var_numero_factura) + " and line_type = 'LINE'"
                              rsaux6.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux6.EOF Then
                                 'var_importe_total = IIf(IsNull(rsaux6!amount_applied), 0, rsaux6!amount_applied) + var_importe_total_df
                                 VAR_IMPORTE_TOTAL = IIf(IsNull(rsaux6!amount_applied), 0, rsaux6!amount_applied)
                                 If VAR_IMPORTE_TOTAL = 0 Then
                                    VAR_PORCENTAJE_FIN = 0
                                 Else
                                    'VAR_PORCENTAJE_FIN = 100 - (Round((IIf(IsNull(rsaux6!amount_applied), 0, rsaux6!amount_applied) * 100) / var_importe_total, 2))
                                    VAR_PORCENTAJE_FIN = (Round((IIf(IsNull(var_importe_total_df), 0, var_importe_total_df) * 100) / VAR_IMPORTE_TOTAL, 2))
                                 End If
                                 VAR_NOTA_CREDITO_DF = var_notas_credito_df
                                 var_precio = var_precio * (1 - (IIf(IsNull(VAR_PORCENTAJE_FIN), 0, VAR_PORCENTAJE_FIN) / 100))
                              End If
                              rsaux6.Close
                           End If
                           rsaux5.Close
                        Else
                           'MsgBox "SELECT PRODUCT_UOM_CODE, CURRENCY_CODE  * FROM qp_list_lines_v A, qp_list_HEADERS_v B WHERE A.list_header_id = B.list_header_id AND  A.list_header_id = " + CStr(var_clave_lista_precios) + " and  A.product_attr_value = " + CStr(rsaux9!INVENTORY_ITEM_ID) + " AND A.product_attr_val_disp = '" + rsaux9!CODIGO + "'"
                           If rsaux11.State = 1 Then
                              rsaux11.Close
                           End If
                           rsaux11.Open "SELECT PRODUCT_UOM_CODE, CURRENCY_CODE  FROM qp_list_lines_v A, qp_list_HEADERS_v B WHERE A.list_header_id = B.list_header_id AND  A.list_header_id = " + CStr(var_clave_lista_precios) + " and  A.product_attr_value = " + CStr(rsaux9!inventory_item_id) + " AND A.product_attr_val_disp = '" + rsaux9!codigo + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           var_unidad_medida = rsaux11!PRODUCT_UOM_CODE
                           VAR_CURRENCY = rsaux11!CURRENCY_CODE
                           rsaux11.Close
                           VAR_ZZ = 0
                           If var_unidad_organizacional = "93" And VAR_ZZ = 1 Then
                              objConn.Open var_conexion_oracle
                              '… Establecer conexión a la base de datos con el objeto objConn.
                              With objCmd
                                   objConn.BeginTrans
                                   .ActiveConnection = objConn
                                   .CommandText = "APPS.xxvia_descuento_linea"
                                   .CommandType = adCmdStoredProc
                                
                                   rsaux10.Open "select * from tb_unidadesorganizacionales where vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                                   If Not rsaux10.EOF Then
                                      var_responsabilidad_facturacion = IIf(IsNull(rsaux10!RESPONSABILIDAD_FACTURACION), "", rsaux10!RESPONSABILIDAD_FACTURACION)
                                   End If
                                   rsaux10.Close
                                         
                                         
                                   Set objParm = .CreateParameter("p_responsabilidad", adVarChar, adParamInput, 100, var_responsabilidad_facturacion)
                                   .Parameters.Append objParm
      
                                        
                                   Set objParm = .CreateParameter("p_org_id", adNumeric, adParamInput, 50, CDbl(var_empresa))
                                   .Parameters.Append objParm
                                   
                                   Set objParm = .CreateParameter("p_uom", adVarChar, adParamInput, 50, var_unidad_medida)
                                   .Parameters.Append objParm
                                
                                   Set objParm = .CreateParameter("p_currency", adVarChar, adParamInput, 50, VAR_CURRENCY)
                                   .Parameters.Append objParm
                                
                                   var_estatus_factura = ""
                                   Set objParm = .CreateParameter("p_inventory_item_id", adNumeric, adParamInput, 50, rsaux9!inventory_item_id)
                                   .Parameters.Append objParm
                                   
                                   Set objParm = .CreateParameter("p_price_list_id", adNumeric, adParamInput, 50, var_clave_lista_precios)
                                   .Parameters.Append objParm
                                   
                                   Set objParm = .CreateParameter("p_cust_account_id", adNumeric, adParamInput, 50, var_clave_titular)
                                   .Parameters.Append objParm
                                
                                   Set objParm = .CreateParameter("xx_unit_price", adNumeric, adParamOutput, 50, var_precio_inflado)
                                   .Parameters.Append objParm
                                   
                                   Set objParm = .CreateParameter("xx_adjusted_price", adNumeric, adParamOutput, 50, var_precio_descuento)
                                   .Parameters.Append objParm
                                   
                                   rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                   rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                   On Error GoTo salir2:
                                   .execute
                                 
                                   var_precio_inflado = .Parameters("xx_unit_price").Value
                                   var_precio_entero = var_precio_inflado
                                   var_precio_descuento = .Parameters("xx_adjusted_price").Value
                                   var_precio = var_precio_descuento
                                   objConn.CommitTrans
                              End With
                              Set objConn = Nothing
                              Set objCmd = Nothing
                           Else
                              x = 0
                              If x = 0 Then
                                 If rsaux10.State = 1 Then
                                    rsaux10.Close
                                 End If
                                 rsaux11.Open "select * from xxvia_system_items_b where segment1 = '" + rsaux9!codigo + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 If Not rsaux11.EOF Then
                                    rsaux10.Open "SELECT OPERAND FROM qp_list_lines_v WHERE list_header_id =  " + CStr(var_clave_lista_precios) + " AND  PRODUCT_ATTR_VAL_DISP = '" + rsaux9!codigo + "' and start_date_active <= sysdate and (end_date_active is null or end_date_active >= sysdate) and Product_Attr_Value = " + CStr(rsaux11!inventory_item_id), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                    'rsaux10.Open "SELECT OPERAND FROM qp_list_lines_v WHERE list_header_id =  " + CStr(var_clave_lista_precios) + " AND  PRODUCT_ATTR_VAL_DISP = '" + rsaux9!codigo + "' AND Product_Attr_Value = " + CStr(rsaux11!INVENTORY_ITEM_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 Else
                                    rsaux10.Open "SELECT OPERAND FROM qp_list_lines_v WHERE list_header_id =  " + CStr(var_clave_lista_precios) + " AND  PRODUCT_ATTR_VAL_DISP = '" + rsaux9!codigo + "' and start_date_active <= sysdate and (end_date_active is null or end_date_active >= sysdate) and ", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 End If
                                 rsaux11.Close
                                 Product_Attr_Value = 16217
                                 var_precio_entero = IIf(IsNull(rsaux10!OPERAND), 0, rsaux10!OPERAND)
                                 var_precio = IIf(IsNull(rsaux10!OPERAND), 0, rsaux10!OPERAND)
                                 rsaux10.Close
                                 VAR_DESCUENTO = 0
                                 rsaux11.Open "SELECT DISTINCT(list_header_id) as calificador FROM qp_qualifiers_v WHERE list_header_id IS NOT NULL", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 While Not rsaux11.EOF
                                       rsaux10.Open "select xxvia_fn_descuento_titular(" + CStr(rsaux11!calificador) + ",'" + CStr(rsaux9!TITULAR) + "') as descuento from dual ", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                       If Not rsaux10.EOF Then
                                          If rsaux10!DESCUENTO > VAR_DESCUENTO Then
                                             VAR_DESCUENTO = rsaux10!DESCUENTO
                                          End If
                                       End If
                                       rsaux10.Close
                                       rsaux11.MoveNext
                                 Wend
                                 rsaux11.Close
                                 var_precio = var_precio * (1 - (VAR_DESCUENTO / 100))
                              Else
                                 var_precio_entero = IIf(IsNull(rsaux9!Precio), 0, rsaux9!Precio)
                                 var_precio = IIf(IsNull(rsaux9!Precio), 0, rsaux9!Precio)
                              End If
                           End If
                        End If
                        End If 'PARA QUE TOME EN CUENTA EL PRECIO VIEJO DE ARTICULOS QUE CAMBIARON DE PRECIO
                        'quitar cuando se termine lod el cambio de cantia
                        If rsaux.State = 1 Then
                           rsaux.Close
                        End If
                                                  ' var_factura_trx_number = rsaux!trx_number
                           'var_fecha_factura_nueva = rsaux!trx_date

                        rsaux.Open "UPDATE TB_ORACLE_REVALUACION_DEVOLUCIONES SET PRECIO_nuevo = " + CStr(var_precio) + ",customer_trx_id_nuevo = '" + CStr(var_numero_factura) + "', DESCUENTO_FINANCIERO_nuevo = '" + CStr(IIf(IsNull(VAR_PORCENTAJE_FIN), 0, VAR_PORCENTAJE_FIN)) + "', NC_DF_NUEVO = '" + Mid(CStr(VAR_NOTA_CREDITO_DF), 1, 100) + "', precio_entero_NUEVO = " + CStr(var_precio_entero) + ", trx_number_nuevo = '" + var_factura_trx_number + "',trx_date_nuevo = '" + CStr(var_fecha_factura_nueva) + "' WHERE NUMERO = " + CStr(rsaux9!numero) + "  AND INVENTORY_ITEM_ID = " + CStr(rsaux9!inventory_item_id), cnn, adOpenDynamic, adLockOptimistic
                        var_cantidad = rsaux9!cantidad
                        x = 0
                        If x = 0 Then
                           While var_cantidad > 0
                                 VAR_CANTIDAD_RESTANTE = var_cantidad - 1
                                 If VAR_CANTIDAD_RESTANTE >= 1 Then
                                    var_cantidad_LEER = 1
                                    var_cantidad = var_cantidad - 1
                                 Else
                                    var_cantidad_LEER = var_cantidad
                                    var_cantidad = 0
                                 End If
                                 If Mid("q23q23", 1, 9) = "DC BULTOS" Then
                                    var_devolucion_costales = 1
                                 Else
                                    var_devolucion_costales = 0
                                 End If
                                 var_consecutivo = var_consecutivo + 1
                           Wend
                        Else
                             If Mid("123123", 1, 9) = "DC BULTOS" Then
                                var_devolucion_costales = 1
                             Else
                                var_devolucion_costales = 0
                             End If
                             
                             var_consecutivo = var_consecutivo + 1
                             var_cantidad_LEER = var_cantidad
                        End If
                        rsaux9.MoveNext
                  Wend
                  rsaux9.Close
salir2:
   'MsgBox Err.Number
   Exit Sub
   If Err.Number = -2147217900 Then
      'MsgBox Err.Description
      rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      Resume
   Else
      '
      'MsgBox Err.Description
      Resume
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
   MsgBox "No se pudo generar el documento electrónico", vbOKOnly, "ATENCION"
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
                  



End Sub

Private Sub Command5_Click()
      var_posible = 1
      If var_posible = 1 Then
         If rsaux8.State = 1 Then
            rsaux8.Close
         End If
         rsaux8.Open "SELECT DISTINCT SOURCE_HEADER_NUMBER AS PEDIDO FROM XXVIA_TB_salidas_cajas WHERE source_header_number = " + "376863", cnnoracle_4, adOpenDynamic, adLockOptimistic
         'rs.Open "SELECT inte_emb_embarque as embarque, inte_paq_caja as caja, source_header_number  as pedido, segment1 as codigo, collector_id as agente,name as nombre_agente, customer_id as cliente, customer_name as nombre_cliente, inventory_item_id, caja_pedido, sello, item_description as descripcion, floa_sal_cantidad_leida as cantidad   FROM XXVIA_TB_salidas_cajas where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rsaux8.EOF Then
            cnn.BeginTrans
            rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) AS CONSECUTIVO FROM TB_TEMP_ORACLE_DETALLE_CAJAS", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
            Else
               var_consecutivo = 1
            End If
            rsaux.Close
            rsaux.Open "INSERT INTO TB_TEMP_ORACLE_DETALLE_CAJAS (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            While Not rsaux8.EOF
                  'rs.Close
                  rs.Open "SELECT floa_sal_Cantidad_leida as cantidad, organizacion,  inte_emb_embarque as embarque, inte_paq_caja as caja, source_header_number  as pedido, a.segment1 as codigo, collector_id as agente,name as nombre_agente, customer_id as cliente, customer_name as nombre_cliente, a.inventory_item_id, caja_pedido, sello, UNIT_WEIGHT as peso, item_description as descripcion, tipo_caja    FROM XXVIA_TB_salidas_cajas a, xxvia_tb_encabezado_embarques, xxvia_system_items_b b, oe_order_headers_all oh where inte_emb_embarque = embarque and organizacion = b.organization_id and a.inventory_item_id = b.inventory_item_id and order_number = a.source_header_number and  oh.ship_from_org_id = organizacion and floa_sal_Cantidad_leida > 0 AND SOURCE_HEADER_NUMBER = " + "376863", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  rsaux.Open "select * from oe_order_headers_all where order_number = " + CStr(rs!pedido), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If rsaux!ORDER_TYPE_ID = 1002 Then
                     var_agente = rs!Agente
                     var_nombre_cliente = rs!nombre_cliente
                     var_nombre_agente = rs!NOMBRE_AGENTE
                     var_cliente = rs!Cliente
                     rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(rsaux!source_document_id) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux2.EOF Then
                        var_nombre_cliente = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                     End If
                     rsaux2.Close
                  Else
                     var_agente = rs!Agente
                     var_nombre_cliente = rs!nombre_cliente
                     var_nombre_agente = rs!NOMBRE_AGENTE
                     var_cliente = rs!Cliente
                  End If
                  rsaux.Close
                  rsaux.Open "alter session set nls_language= 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  While Not rs.EOF
                        'var_cadena = "select ic.category_concat_segs,fft.description from mtl_item_categories_v ic, fnd_flex_value_sets   ffvs, fnd_flex_values_vl ffv,  fnd_flex_values_tl fft,mtl_parameters mtp where UPPER (ic.category_set_name) LIKE '%VIANNEY%EXPORTACION%' AND ic.inventory_item_id = " + CStr(rs!INVENTORY_ITEM_ID) + " AND ic.organization_id = mtp.organization_id AND mtp.organization_code = 'MTO' AND ic.category_concat_segs =  ffv.flex_value_meaning AND ffvs.flex_value_set_name = 'VIANNEY_INV_EXPORTACION' AND ffvs.flex_value_set_id  =  ffv.flex_value_set_id AND ffv.flex_value_id  =  fft.flex_value_id AND fft.language  =  USERENV('LANG') "
                        'If rsaux.State = 1 Then
                        '   rsaux.Close
                        'End If
                        'rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        'If Not rsaux.EOF Then
                        '   var_arancel = CStr(IIf(IsNull(rsaux!Description), "", rsaux!Description))
                        'Else
                        '   var_arancel = ""
                        'End If
                        'rsaux.Close
               
                        var_peso = IIf(IsNull(rs!PESO), 0, rs!PESO)
               
                        'var_pedido = rs!pedido
                        'rsaux.Open "select unit_selling_price from oe_order_headers_all oh, oe_order_lines_all ol where order_number = " + CStr(var_pedido) + " and oh.header_id = ol.header_id and oh.ship_from_org_id = " + var_unidad_organizacional + " and ol.inventory_item_id = " + CStr(IIf(IsNull(rs!INVENTORY_ITEM_ID), 0, rs!INVENTORY_ITEM_ID)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        'If Not rsaux.EOF Then
                        '   var_precio = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
                        'Else
                        '   var_precio = 0
                        'End If
                        'rsaux.Close
                        var_precio = 0
                        
                        'MsgBox Len(rs!sello)
                        var_cadena = "INSERT INTO TB_TEMP_ORACLE_DETALLE_CAJAS (INTE_TEM_CONSECUTIVO, EMBARQUE, AGENTE, NOMBRE_AGENTE, CLIENTE, NOMBRE_CLIENTE, CODIGO, DESCRIPCION, CANTIDAD, PRECIO, PESO, CAJA, CONTENIDO, INVENTORY_ITEM_ID, ARANCEL, CODIGO_BARRAS, SELLO, CAJA_PEDIDO, PEDIDO, bulto)"
                        var_cadena = var_cadena + " VALUES (" + CStr(var_consecutivo) + "," + CStr(rs!Embarque) + ",'" + CStr(IIf(IsNull(var_agente), "", var_agente)) + "','" + IIf(IsNull(var_nombre_agente), "", var_nombre_agente) + "', '" + CStr(var_cliente) + "','" + IIf(IsNull(var_nombre_cliente), "", var_nombre_cliente) + "','" + rs!codigo + "','" + Mid(rs!Descripcion, 1, 100) + "'," + CStr(rs!cantidad) + "," + CStr(var_precio) + "," + CStr(var_peso) + "," + CStr(rs!Caja) + ",''," + CStr(rs!inventory_item_id) + ", '" + IIf(IsNull(var_arancel), "", var_arancel) + "','" + IIf(IsNull(var_codigo_barras), "", var_codigo_barras) + "','" + Trim(IIf(IsNull(rs!sello), "", rs!sello)) + "'," + CStr(IIf(IsNull(rs!caja_pedido), rs!Caja, rs!caja_pedido)) + "," + CStr(rs!pedido) + ",'" + IIf(IsNull(rs!tipo_caja), "", rs!tipo_caja) + "')"
                        'MsgBox var_cadena
                        rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        
                        
                        rs.MoveNext
                  Wend
                  rsaux1.Open "DELETE FROM TB_TEMP_ORACLE_DETALLE_CAJAS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND EMBARQUE IS NULL", cnn, adOpenDynamic, adLockOptimistic
             
                  rsaux1.Open "SELECT DISTINCT CAJA FROM TB_TEMP_ORACLE_DETALLE_CAJAS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO = " + CStr(rsaux8!pedido), cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveFirst
                  VAR_CANTIDAD_CAJAS = 0
                  While Not rsaux1.EOF
                        VAR_CANTIDAD_CAJAS = VAR_CANTIDAD_CAJAS + 1
                        rsaux1.MoveNext
                  Wend
                  rsaux1.Close
                  rsaux1.Open "UPDATE TB_TEMP_ORACLE_DETALLE_CAJAS SET NUMERO_CAJAS = " + CStr(VAR_CANTIDAD_CAJAS) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO = " + CStr(rsaux8!pedido), cnn, adOpenDynamic, adLockOptimistic
                  rsaux8.MoveNext
                  rs.Close
            Wend
            rsaux8.Close
            rsaux8.Open "SELECT DISTINCT PEDIDO FROM TB_TEMP_ORACLE_DETALLE_CAJAS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO IS NOT NULL", cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux8.EOF
                  VAR_CADENA_BULTOS = ""
                  strconsulta = "SELECT tipo_caja, COUNT(*) CANTIDAD FROM XXVIA_VW_CAJAS_POR_PEDIDO WHERE SOURCE_HEADER_NUMBER = ? GROUP BY TIPO_CAJA"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 20, rsaux8!pedido)
                       .Parameters.Append parametro
                  End With
                  Set rsaux9 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  While Not rsaux9.EOF
                        If VAR_CADENA_BULTOS = "" Then
                           VAR_CADENA_BULTOS = rsaux9!tipo_caja + ": " + CStr(rsaux9!cantidad)
                        Else
                           VAR_CADENA_BULTOS = VAR_CADENA_BULTOS + ",    " + rsaux9!tipo_caja + ": " + CStr(rsaux9!cantidad)
                        End If
                        rsaux9.MoveNext
                  Wend
                  rsaux9.Close
                  rsaux9.Open "UPDATE TB_TEMP_ORACLE_DETALLE_CAJAS SET VCHA_PAQ_TIPO_BULTOS = '" + IIf(IsNull(VAR_CADENA_BULTOS), "", VAR_CADENA_BULTOS) + "' WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO = " + CStr(rsaux8!pedido), cnn, adOpenDynamic, adLockOptimistic
                  rsaux8.MoveNext
            Wend
            rsaux8.Close
         
         
         
            
            
            
            Set reporte = appl.OpenReport(App.Path + "\rep_oracle_packing_list.rpt")
            reporte.RecordSelectionFormula = "{VW_ORACLE_DETALLE_CAJAS.EMBARQUE} = " + txt_embarque + " and {VW_ORACLE_DETALLE_CAJAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
               reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Packing List"
            frmvistasprevias.Show 1
            Set reporte = Nothing
       
            var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_packing_list.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_DETALLE_CAJAS.EMBARQUE} = " + txt_embarque + " and {VW_ORACLE_DETALLE_CAJAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\packing_list" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               
               MsgBox "Se a terminado de guardar el archivo " + archivo
               var_si = MsgBox("Desea enviar el packing list por correo?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  VAR_CORREO_ELECTRONICO = "vluna@vianney.com.mx"
                  If Trim(VAR_CORREO_ELECTRONICO) <> "" Then
                     If MAPISession1.SessionID = 0 Then
                        MAPISession1.SignOn
                     End If
                     MAPIMessages1.SessionID = MAPISession1.SessionID
                     MAPIMessages1.Compose
                     MAPIMessages1.RecipDisplayName = VAR_CORREO_ELECTRONICO
                     MAPIMessages1.RecipAddress = VAR_CORREO_ELECTRONICO
                     MAPIMessages1.AddressResolveUI = True
                     MAPIMessages1.ResolveName
                     MAPIMessages1.MsgSubject = "Packing list"
                     MAPIMessages1.MsgNoteText = "Se anexa archivo de packing list"
                     MAPIMessages1.AttachmentPathName = archivo
                     MAPIMessages1.send True
                     If MAPISession1.SessionID > 0 Then
                        MAPISession1.SignOff
                     End If
                  End If
               End If
            End If
            rsaux.Open "delete from TB_TEMP_ORACLE_DETALLE_CAJAS where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
         End If
      Else
         If var_posible = 0 Then
            MsgBox "El embarque no a sido cerrado.", vbOKOnly, "ATENCION"
         End If
         If var_posible = 2 Then
            MsgBox "El embarque no existe.", vbOKOnly, "ATENCION"
         End If
      End If

End Sub

Private Sub Command6_Click()
   Dim clnt As New SoapClient30
   'clnt.MSSoapInit "http://serviciowebcedisdesa.vianney.com.mx/EnviarCorreos.asmx"
   clnt.MSSoapInit "http://serviciowebcedisdesa.vianney.com.mx/EnviarCorreos.asmx?wsdl"
   var_s = clnt.CorreoAdjunto("fserna@vianney.com.mx", "prueba", "mensaje prueba", "380179", "1002")
   'Set clnt = Nothing

End Sub

Private Sub Command7_Click()
   Dim clnt2 As New SoapClient30
   rsaux6.Open "select pedido as source_header_number from tb_oracle_pedidos_asignados_embarques where pedido = " + Me.txt_pedido, cnn, adOpenDynamic, adLockOptimistic
    var_cadena = "select order_type_id from oe_order_headers_all where order_number = ?"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = var_cadena
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rsaux6!source_header_number))
                             .Parameters.Append parametro
                        End With
                        Set rsaux7 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        If rsaux7!ORDER_TYPE_ID = 1002 Then
                           rsaux7.Close
                           var_cadena = "select  A.SECONDARY_INVENTORY_NAME, A.DESCRIPTION, ADDRESS_LINE_1, ADDRESS_LINE_2, TOWN_OR_CITY, REGION_1, COUNTRY, POSTAL_CODE, EMAIL from mtl_secondary_inventories a, hr_locations_all b, xxvia_jv_tb_agentes c, po_requisition_headers_ALL D, OE_ORDER_HEADERS_ALL E Where A.location_id = b.location_id and a.secondary_inventory_name = c.subinventory_code AND E.source_document_id = D.requisition_header_id AND A.secondary_inventory_name = D.ATTRIBUTE1 AND E.ORDER_NUMBER = ?                 "
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = var_cadena
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rsaux6!source_header_number))
                                .Parameters.Append parametro
                           End With
                           Set rsaux7 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           If Not rsaux7.EOF Then
                              VAR_CORREO = IIf(IsNull(rsaux7!Email), "", rsaux7!Email)
                              If VAR_CORREO <> "" Then
                                 ' SE ENVIA CORREO A TIENDA
                                 clnt2.MSSoapInit "http://serviciowebcedisdesa.vianney.com.mx/EnviarCorreos.asmx?wsdl"
                                 var_s = clnt2.CorreoAdjunto(VAR_CORREO, "Packing List de pedido " + CStr(rsaux6!source_header_number), "Se anexa packing list de pedido " + CStr(rsaux6!source_header_number) + " del CN. " + rsaux7!Description, CStr(rsaux6!source_header_number), "1002")
                              End If
                           End If
                           rsaux7.Close
                        Else
                                                     
                           var_cadena = "select razon_social_cliente as description, email_Address as email from xxvia_vw_clientes_bcp a, oe_order_headers_all b where order_number = ? and a.site_use_id = b.invoice_to_org_id"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = var_cadena
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rsaux6!source_header_number))
                                .Parameters.Append parametro
                           End With
                           Set rsaux7 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           If Not rsaux7.EOF Then
                              VAR_CORREO = IIf(IsNull(rsaux7!Email), "", rsaux7!Email)
                              If VAR_CORREO <> "" Then
                                 ' SE ENVIA CORREO A TIENDA
                                 clnt2.MSSoapInit "http://serviciowebcedisdesa.vianney.com.mx/EnviarCorreos.asmx?wsdl"
                                 var_s = clnt2.CorreoAdjunto(VAR_CORREO, "Packing List de pedido " + CStr(rsaux6!source_header_number), "Se anexa packing list de pedido " + CStr(rsaux6!source_header_number) + " del cliente " + rsaux7!Description, CStr(rsaux6!source_header_number), "9999")
                                 'Me.Text1 = VAR_CORREO + ", Packing List de pedido " + CStr(rsaux6!source_header_number) + ", Se anexa packing list de pedido " + CStr(rsaux6!source_header_number) + " del cliente " + rsaux7!Description + ", " + CStr(rsaux6!source_header_number) + ", 9999"
                              End If
                           End If
                           rsaux7.Close
                        End If
   rsaux6.Close
End Sub

Private Sub Command8_Click()
   rs.Close
   rs.Open "SELECT * FROM TB_CHOFERES", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux.Open "INSERT INTO XXVIA_TB_CHOFERES (ID_CHOFER, NOMBRE) VALUES ('" + rs!vcha_cho_chofer_id + "','" + rs!vcha_cho_nombre + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
End Sub

Private Sub Command9_Click()
   rs.Open "select * from fa_281217_B", cnn, adOpenDynamic, adLockOptimistic
   
   While Not rs.EOF
         rsaux1.Open "select * from XXVIA_TB_COMPLEMENTOS_PK_LIST where codigo = '" + rs!codigo + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rsaux1.EOF Then
            
            'rsaux11.Open "update XXVIA_TB_COMPLEMENTOS_PK_LIST set fraccion_arancelaria = '" + CStr(rs!fraccion) + "', contenido = '" + IIf(IsNull(rs!contenido), "", rs!contenido) + "', composicion = '" + IIf(IsNull(rs!composicion), "", rs!composicion) + "', hecho_en = '" + rs!hecho_en + "' where codigo = '" + rs!CODIGO + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            rsaux11.Open "update XXVIA_TB_COMPLEMENTOS_PK_LIST set fraccion_arancelaria = '" + CStr(rs!fraccion) + "' where codigo = '" + rs!codigo + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         Else
            var_cadena = var_cadena + " " + rs!codigo
         End If
         rsaux1.Close
         rs.MoveNext
   Wend
   rs.Close
   MsgBox var_cadena
End Sub

Private Sub Form_Load()
   
   Top = 0
   Left = 0
   var_i = 0
   'Me.Text5.Text = Format(Date, "YYYY/MM/DD")
   Me.txt_embarque_pedido = 343980
   Me.txt_embraue_nota_envio = 699896
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 78 Then
      x = 1
   End If
   If Shift = 4 And KeyCode = 66 Then
      x = 1
   End If
   If Shift = 4 And KeyCode = 73 Then
      x = 1
   End If
   If Shift = 4 And KeyCode = 67 Then
      x = 1
   End If
   If Shift = 4 And KeyCode = 77 Then
      x = 1
   End If

End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
    x = 1 + 1
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    x = 1 + 1

End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
    MsgBox Str(KeyCode)
End Sub

Private Sub Text5_Change()

End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   If Len(Me.txt_codigo) = 0 Then
      var_hora_inicio = Now
   End If
   If Len(Me.txt_codigo) = 4 Then
      var_hora_fin = Now
      var_diferencia = Round(CDbl(var_hora_fin - var_hora_inicio), 5)
      If var_diferencia >= 0.00002 Then
         Me.txt_codigo = ""
      End If
   End If
   If KeyAscii = 13 Then
      Me.txt_codigo.Enabled = False
      Sleep 3000
      Me.txt_codigo.Enabled = True
      Me.txt_codigo.SetFocus
   End If
End Sub

Private Sub txt_embraue_nota_envio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim dl As Long                                 ' Valor devuelto por la función API
      Dim sAttributes As String                  ' Aributos
      Dim sDriver As String                       ' Nombre del controlador
      Dim sDescription As String                ' Descripción del DSN
      Dim sDsnName As String                  ' Nombre del DSN

      Const ODBC_ADD_SYS_DSN As Long = 4         ' Se creará un DSN de sistema
      Const vbAPINull As Long = 0&                         ' Puntero NULL

      ' se elimina
      Const ODBC_REMOVE_SYS_DSN As Long = 6    ' Se eliminará un DSN de sistema
      sDsnName = "DSN=sqlquezada2"
      sDriver = "SQL Server"
      dl = SQLConfigDataSource(0&, ODBC_REMOVE_SYS_DSN, sDriver, sDsnName)

      'se crea
      sDsnName = "sqlsistema"
      sDescription = "sqlsistema"
      sDriver = "SQL Server"
      sAttributes = "DSN=" & sDsnName & Chr(0)
      sAttributes = sAttributes & "Server=" + parametros(0) & Chr$(0)
      sAttributes = sAttributes & "Description=" & sDescription & Chr(0)
      sAttributes = sAttributes & "Database=" + var_bd_reportes & Chr(0)
      strAttributes = strAttributes & "UID=sa" & Chr$(0)
      strAttributes = strAttributes & "PWD=elia" & Chr$(0)
      dl = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYS_DSN, sDriver, sAttributes)
      
      Dim var_location_id As Double
      Dim VAR_CLAVE_USUARIO_MOV As String
      Dim var_fecha_inicio As String
      Dim var_fecha_fin As String
      Dim var_consignacion As String
      If IsNumeric(Me.txt_embraue_nota_envio) Then
         var_posible_embarque = 1
         var_Cadena_pedidos = Me.txt_embraue_nota_envio
         var_j = 0
         If var_posible_embarque = 1 Then
            rsaux.Open "alter session set nls_languAge = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            
            var_cadena = "SELECT  A.ATTRIBUTE1, B.description as nombre_almacen, g.organization_id, oh.ordered_date, oh.source_document_id, oh.header_id, oh.order_number, oh.transactional_curr_code, NVL(ol.ordered_quantity,0) AS CANTIDAD_PEDIDA, NVL(ol.cancelled_quantity,0) AS CANTIDAD_NEGADA, NVL(ol.shipped_quantity,0)   AS CANTIDAD_surtida, ol.line_id, ol.ordered_item, g.description, ol.order_quantity_uom, ol.inventory_item_id, ol.price_list_id, ol.unit_selling_price, DECODE(ol.cancelled_flag,'Y','CANCELADA','SURTIDA') line_status, ol.flow_status_code, h.linea FROM oe_order_headers_all oh, oe_order_lines_all ol, OE_ORDER_LINES_HISTORY OLH, xxvia_system_items_b g, xxvia_vw_articulos_cat h, po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE order_number  IN (" + var_Cadena_pedidos + ") "
            var_cadena = var_cadena + " AND oh.header_id = ol.header_id AND ol.ship_from_org_id = " + var_unidad_organizacional + " AND oL.header_id = oLh.header_id(+) AND OL.LINE_ID = OLH.LINE_ID(+) AND ol.inventory_item_id = g.inventory_item_id"
            var_cadena = var_cadena + " AND g.organization_id = ol.ship_from_org_id AND h.item_id = g.inventory_item_id AND h.organization_id = g.organization_id AND requisition_header_id = OH.source_document_id AND B.secondary_inventory_name = A.ATTRIBUTE1 "
               
            If rsaux.State = 1 Then
               rsaux.Close
            End If
            rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_posible_embarque = 0
            If Not rsaux.EOF Then
               var_posible_embarque = 1
            End If
            rsaux.Close
            If var_posible_embarque = 1 Then
            
            var_j = 1
            var_i = 1
            If rsaux.State = 1 Then
               rsaux.Close
            End If
            var_embarque_pedido = Me.txt_embarque_pedido.Text
            rsaux.Open "select distinct INTE_EMB_EMBARQUE from xxvia_tb_salidas where SOURCE_HEADER_NUMBER = " + Me.txt_embraue_nota_envio, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               var_numero_embarque = IIf(IsNull(rsaux!inte_emb_embarque), 0, rsaux!inte_emb_embarque)
            End If
            rsaux.Close
            rsaux.Open "select distinct INTE_EMB_EMBARQUE from xxvia_tb_SAlidas_cajas where SOURCE_HEADER_NUMBER = " + Me.txt_embraue_nota_envio + " AND INTE_EMB_EMBARQUE = " + CStr(Me.txt_embarque_pedido), cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               var_si = 1
            Else
               var_si = 0
            End If
            rsaux.Close
            var_numero_embarque = var_embarque_pedido
            rsaux.Open "alter session set nls_languAge = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            If var_si = 1 Then
            If var_j = var_i Then
               var_cadena = "SELECT  A.ATTRIBUTE1, B.description as nombre_almacen, g.organization_id, oh.ordered_date, oh.source_document_id, oh.header_id, oh.order_number, oh.transactional_curr_code, NVL(ol.ordered_quantity,0) AS CANTIDAD_PEDIDA, NVL(ol.cancelled_quantity,0) AS CANTIDAD_NEGADA, NVL(ol.shipped_quantity,0)   AS CANTIDAD_surtida, ol.line_id, ol.ordered_item, g.description, ol.order_quantity_uom, ol.inventory_item_id, ol.price_list_id, ol.unit_selling_price, DECODE(ol.cancelled_flag,'Y','CANCELADA','SURTIDA') line_status, ol.flow_status_code, h.linea FROM oe_order_headers_all oh, oe_order_lines_all ol, OE_ORDER_LINES_HISTORY OLH, xxvia_system_items_b g, xxvia_vw_articulos_cat h, po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE order_number  IN (" + var_Cadena_pedidos + ") "
               var_cadena = var_cadena + " AND oh.header_id = ol.header_id AND ol.ship_from_org_id = " + var_unidad_organizacional + " AND oL.header_id = oLh.header_id(+) AND OL.LINE_ID = OLH.LINE_ID(+) AND ol.inventory_item_id = g.inventory_item_id"
               var_cadena = var_cadena + " AND g.organization_id = ol.ship_from_org_id AND h.item_id = g.inventory_item_id AND h.organization_id = g.organization_id AND requisition_header_id = OH.source_document_id AND B.secondary_inventory_name = A.ATTRIBUTE1 "
               
               If rsaux.State = 1 Then
                  rsaux.Close
               End If
               rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  'MsgBox rsaux!ATTRIBUTE1
                  cnn.BeginTrans
                  If rs.State = 1 Then
                     rs.Close
                  End If
                  rs.Open "select max(inte_tem_consecutivo) from tb_Temp_oracle_NOTA_ENVIO", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
                  Else
                     var_consecutivo = 1
                  End If
                  rs.Close
                  rs.Open "insert into tb_Temp_oracle_NOTA_ENVIO (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                  cnn.CommitTrans
                  
                  
                  rsaux1.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  rsaux1.Open "SELECT LAST_UPDATE_dATE FROM WSH_DELIVERABLES_V WHERE SOURCE_HEADER_NUMBER = '" + CStr(rsaux!order_number) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  
                  
                  VAR_FECHA_MOVIMIENTO = CStr(Date)
                  VAR_FECHA_MOVIMIENTO = CStr(rsaux1!LAST_UPDATE_DATE)
                  If rsaux3.State = 1 Then
                     rsaux3.Close
                  End If
                  rsaux3.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  rsaux3.Open "SELECT * FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = " + CStr(rsaux!order_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  var_header = rsaux3!header_id
                  var_dia = CStr(Day(IIf(IsNull(rsaux3!pricing_date), Date, rsaux3!pricing_date)))
                  var_mes = CStr(Month(IIf(IsNull(rsaux3!pricing_date), Date, rsaux3!pricing_date)))
                  var_año = CStr(Year(IIf(IsNull(rsaux3!pricing_date), Date, rsaux3!pricing_date)))
                  If Len(var_dia) = 1 Then
                     var_dia = "0" + var_dia
                  End If
                  If Len(var_mes) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  If Len(var_año) = 2 Then
                     var_año = "20" + var_año
                  End If
                  VAR_FECHA_PRECIO = var_dia + "/" + var_mes + "/" + var_año
                  rsaux3.Close
                  'SELECT DISTINCT b.secondary_inventory_name AS CLAVE_ALMACEN, B.DESCRIPTION AS NOMBRE_ALMACEN  FROM WSH_DELIVERABLES_V A, mtl_subinventories_all_v B WHERE source_header_number = 74124 AND B.secondary_inventory_name = A.subinventory AND A.organization_id = b.organization_id AND A.SOURCE_HEADER_ID = 4355290
                  rsaux3.Open "SELECT DISTINCT b.secondary_inventory_name AS CLAVE_ALMACEN, B.DESCRIPTION AS NOMBRE_ALMACEN  FROM WSH_DELIVERABLES_V A, mtl_subinventories_all_v B WHERE source_header_number = " + CStr(rsaux!order_number) + " AND B.secondary_inventory_name = A.subinventory AND A.organization_id = b.organization_id AND A.SOURCE_HEADER_ID = " + CStr(var_header), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux3.EOF Then
                     var_almacen = rsaux3!CLAVE_ALMACEN
                     var_nombre_almacen = rsaux3!nombre_almacen
                  End If
                  rsaux3.Close
                  If var_almacen = "CDISTEX_PT" Then
                     var_almacen = "TEX_PT_QL"
                     var_nombre_almacen = "EL VERGEL PRODUCTO TERMINADO TEX"
                  End If
                  
                  rsaux3.Open "select * from mtl_secondary_inventories where secondary_inventory_name = '" + rsaux!attribute1 + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  var_consignacion = IIf(IsNull(rsaux3!attribute3), "", rsaux3!attribute3)
                  var_almacen_icg = IIf(IsNull(rsaux!attribute1), "", rsaux!attribute1)
                  If Not rsaux3.EOF Then
                     var_location_id = IIf(IsNull(rsaux3!LOCATION_ID), 0, rsaux3!LOCATION_ID)
                     If var_location_id > 0 Then
                        rsaux4.Open "select ADDRESS_LINE_1, ADDRESS_LINE_2, TOWN_OR_CITY, REGION_1, COUNTRY, POSTAL_CODE  from hr_locations_all where location_id = '" + CStr(CDbl(var_location_id)) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        VAR_DIRECCION = IIf(IsNull(rsaux4!ADDRESS_LINE_1), "", rsaux4!ADDRESS_LINE_1)
                        VAR_COLONIA = IIf(IsNull(rsaux4!ADDRESS_LINE_2), "", rsaux4!ADDRESS_LINE_2)
                        var_ciudad = IIf(IsNull(rsaux4!TOWN_OR_CITY), "", rsaux4!TOWN_OR_CITY)
                        var_estado = IIf(IsNull(rsaux4!REGION_1), "", rsaux4!REGION_1)
                        var_pais = IIf(IsNull(rsaux4!COUNTRY), "", rsaux4!COUNTRY)
                        VAR_CP = IIf(IsNull(rsaux4!POSTAL_code), "", rsaux4!POSTAL_code)
                        rsaux4.Close
                     Else
                        VAR_DIRECCION = ""
                        VAR_COLONIA = ""
                        var_ciudad = ""
                        var_estado = ""
                        var_pais = ""
                        VAR_CP = ""
                     End If
                  End If
                  rsaux3.Close
                  rsaux3.Open "SELECT * FROM TB_USUARIOS WHERE VCHA_USU_USUARIO_ID = '" + VAR_CLAVE_USUARIO_MOV + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux3.EOF Then
                     VAR_NOMBRE_USUARIO_ENTREGO = IIf(IsNull(rsaux3!vcha_usu_nombre), "", rsaux3!vcha_usu_nombre) + " " + IIf(IsNull(rsaux3!vcha_usu_apellidos), "", rsaux3!vcha_usu_apellidos)
                  Else
                  End If
                  rsaux3.Close
                  
                  var_clave_Destino = rsaux!attribute1
                  var_nombre_destino = rsaux!nombre_almacen
                  
                  
                  rsaux.Close
                  rsaux.Open "select * from xxvia_Tb_Salidas_cajas a, xxvia_vw_categorias_item_b b where source_header_number = " + Me.txt_embraue_nota_envio + " and inte_emb_embarque = " + Me.txt_embarque_pedido + " and a.segment1 = codigo and organization_id = " + CStr(var_unidad_organizacional) + " and floa_Sal_Cantidad_leida > 0", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  
                  
                  While Not rsaux.EOF
                  
                        If var_consignacion = "PTO_CONS2" Then
                           If rsaux1.State = 1 Then
                              rsaux1.Close
                           End If
                           'rsaux1.Open "SELECT OPERAND FROM qp_list_lines_v WHERE list_header_id =  9007 AND  PRODUCT_ATTR_VAL_DISP = '" + rsaux!ORDERED_ITEM + "' and start_date_active <= to_date('" + CStr(VAR_FECHA_PRECIO) + "','DD/MM/YYYY') and (end_date_active is null or end_date_active >= to_date('" + CStr(VAR_FECHA_PRECIO) + "','DD/MM/YYYY')) and Product_Attr_Value = " + CStr(rsaux!inventory_item_id), cnnoracle_4, adOpenDynamic, adLockOptimistic
                           rsaux1.Open "SELECT OPERAND FROM qp_list_lines_v WHERE list_header_id =  9007 and start_date_active <= to_date('" + CStr(VAR_FECHA_PRECIO) + "','DD/MM/YYYY') and (end_date_active is null or end_date_active >= to_date('" + CStr(VAR_FECHA_PRECIO) + "','DD/MM/YYYY')) and Product_Attr_Value = " + CStr(rsaux!inventory_item_id), cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If Not rsaux1.EOF Then
                              'var_precio = IIf(IsNull(rsaux1!OPERAND), rsaux!unit_selling_price, rsaux1!OPERAND) / 1.16
                              var_precio = rsaux1!OPERAND / 1.16
                           Else
                              var_precio = 0
                           End If
                           rsaux1.Close
                        Else
                           'var_precio = rsaux!unit_selling_price
                           var_precio = 0
                        End If
                        
                        var_cadena = "INSERT INTO tb_Temp_oracle_NOTA_ENVIO (INTE_TEM_CONSECUTIVO, PEDIDO,                  FECHA,                 ALMACEN,                       CLIENTE,            NOMBRE_CLIENTE,                                            EMBARQUE,                 LINEA, NOMBRE_LINEA,            CODIGO,                     NOMBRE_ARTICULO,              ENTREGO, INICIO, TERMINO, DIRECCION, COLONIA, CIUDAD, ESTADO, PAIS, CP, CANTIDAD, PRECIO, INTE_EMB_EMBARQUE) VALUES "
                        'var_cadena = var_cadena + " (" + CStr(var_consecutivo) + ", " + CStr(rsaux!order_number) + ",'" + VAR_FECHA_MOVIMIENTO + "', '" + var_nombre_almacen + "','" + CStr(rsaux!attribute1) + "', '" + rsaux!nombre_almacen + "'," + Me.txt_embraue_nota_envio + ",'','" + rsaux!Linea + "','" + rsaux!ORDERED_ITEM + "', '" + rsaux!Description + "','" + VAR_NOMBRE_USUARIO_ENTREGO + "','" + CStr(VAR_FECHA_PRECIO) + "','" + CStr(VAR_FECHA_MOVIMIENTO) + "','" + VAR_DIRECCION + "','" + VAR_COLONIA + "','" + var_ciudad + "','" + var_estado + "','" + var_pais + "','" + VAR_CP + "'," + CStr(rsaux!CANTIDAD_SURTIDA) + "," + CStr(var_precio) + "," + CStr(var_numero_embarque) + ")"
                        var_cadena = var_cadena + " (" + CStr(var_consecutivo) + ", " + CStr(rsaux!source_header_number) + ",'" + VAR_FECHA_MOVIMIENTO + "', '" + var_nombre_almacen + "','" + CStr(var_clave_Destino) + "', '" + var_nombre_destino + "'," + Me.txt_embraue_nota_envio + ",'','" + rsaux!LINEA + "','" + rsaux!SEGMENT1 + "', '" + rsaux!Descripcion + "','" + VAR_NOMBRE_USUARIO_ENTREGO + "','" + CStr(VAR_FECHA_PRECIO) + "','" + CStr(VAR_FECHA_MOVIMIENTO) + "','" + VAR_DIRECCION + "','" + VAR_COLONIA + "','" + var_ciudad + "','" + var_estado + "','" + var_pais + "','" + VAR_CP + "'," + CStr(rsaux!FLOA_SAL_CANTIDAD_LEIDA) + "," + CStr(var_precio) + "," + CStr(var_numero_embarque) + ")"
                        'var_nombre_almacen_consignacioN = rsaux!nombre_almacen
                        var_nombre_almacen_consignacioN = ""
                        'MsgBox var_cadena
                        If rsaux1.State = 1 Then
                           rsaux1.Close
                        End If
                        rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        
                        rsaux.MoveNext
                  Wend
                  
                  
                  rsaux1.Open "DELETE FROM tb_Temp_oracle_NOTA_ENVIO WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO IS NULL", cnn, adOpenDynamic, adLockOptimistic
                  rsaux.MoveFirst
                  'If rsaux!order_number = 125410 Then
                  '   rsaux1.Open "select segment1, sum(floa_Sal_Cantidad_leida) as cantidad from xxvia_tb_salidas_cajas where source_header_number = 125410 group by segment1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  '   While Not rsaux1.EOF
                  '         rsaux9.Open "update tb_Temp_oracle_NOTA_ENVIO set cantidad = " + CStr(rsaux1!cantidad) + " where pedido = 125410 and inte_tem_consecutivo = " + CStr(var_consecutivo) + " and codigo = '" + rsaux1!segment1 + "'", cnn, adOpenDynamic, adLockOptimistic
                  '         rsaux1.MoveNext
                  '   Wend
                  '   rsaux1.Close
                  'End If
                  rsaux1.Open "select sum(cantidad) from tb_Temp_oracle_NOTA_ENVIO where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux1.EOF Then
                     var_cantidad_oracle = IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value)
                  Else
                     var_cantidad_oracle = 0
                  End If
                  rsaux1.Close
                  rsaux1.Open "select sum(floa_sal_Cantidad_leida) from xxvia_tb_salidas where source_header_number = " + CStr(Me.txt_embraue_nota_envio) + " and inte_emb_embarque = " + Me.txt_embarque_pedido, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux1.EOF Then
                     var_cantidad_leida = IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value)
                  Else
                     var_cantidad_leida = 0
                  End If
                  rsaux1.Close
                  rsaux1.Open "select sum(floa_sal_Cantidad_leida) from xxvia_tb_salidas_cajas where source_header_number = " + CStr(Me.txt_embraue_nota_envio) + " and inte_emb_embarque = " + Me.txt_embarque_pedido, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux1.EOF Then
                     var_cantidad_leida = var_cantidad_leida + IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value)
                  Else
                     var_cantidad_leida = var_cantidad_leida + 0
                  End If
                  rsaux1.Close
                  'If Round(var_cantidad_leida, 2) = Round(var_cantidad_leida, 2) Then
                  If Round(var_cantidad_leida, 2) <= Round(var_cantidad_oracle, 2) Then
                     strconsulta = "select nvl(attribute7,'N') activa_pos_icg from mtl_secondary_inventories where secondary_inventory_name=?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_almacen_icg)
                          .Parameters.Append parametro
                     End With
                     Set rsaux9 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     If Not rsaux9.EOF Then
                        var_posible_cgi = IIf(IsNull(rsaux9(0).Value), "N", rsaux9(0))
                     Else
                        var_posible_cgi = "N"
                     End If
                     rsaux9.Close
                     var_posible_cgi = "N"
                     rsaux9.Open "select * from TB_ORACLE_NOTAS_IMPRESAS_ICG where pedido = " + Me.txt_embraue_nota_envio, cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux9.EOF Then
                        var_posible_cgi = "N"
                     End If
                     rsaux9.Close
                     var_posible_cgi = "N"
                     If var_posible_cgi = "Y" Then
                        If CDbl(Me.txt_embraue_nota_envio) >= 161037 Then
                        If cnnicg_sql.State = 1 Then
                           cnnicg_sql.Close
                        End If
                        cnnicg_sql.Open "Provider=SQLOLEDB.1;Password=icgfront2013;Persist Security Info=True;User ID=ICGAdmin;Initial Catalog=general;Data Source=sqlposprod"
                        
                        rsaux1.Open "SELECT source_header_number, inte_paq_caja, segment1, sum(floa_sal_Cantidad_leida) as FLOA_SAL_CANTIDAD_LEIDA FROM XXVIA_TB_SALIDAS_CAJAS WHERE SOURCE_HEADER_NUMBER = " + Me.txt_embraue_nota_envio + " AND FLOA_SAL_CANTIDAD_LEIDA >0 group by source_header_number, inte_paq_caja, segment1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        While Not rsaux1.EOF
                              'VAR_EMBARQUE_ICG = rsaux1!inte_emb_embarque
                              VAR_CAJA_ICG = rsaux1!INTE_PAQ_CAJA
                              If Len(Trim(Str(VAR_CAJA_ICG))) = 1 Then
                                 var_referencia_caja = "00" + Trim(Str(VAR_CAJA_ICG))
                              End If
                              If Len(Trim(Str(VAR_CAJA_ICG))) = 2 Then
                                 var_referencia_caja = "0" + Trim(Str(VAR_CAJA_ICG))
                              End If
                              If Len(Trim(Str(VAR_CAJA_ICG))) = 3 Then
                                 var_referencia_caja = Trim(Str(VAR_CAJA_ICG))
                              End If
                              VAR_CAJA_S = var_referencia_caja
                              var_dia_s = CStr(Day(Now))
                              var_mes_s = CStr(Month(Now))
                              var_año_s = CStr(Year(Now))
                              If Len(var_dia_s) = 1 Then
                                 var_dia_s = "0" + var_dia_s
                              End If
                              If Len(var_mes_s) = 1 Then
                                 var_mes_s = "0" + var_mes_s
                              End If
                              If Len(var_año_s) = 2 Then
                                 var_año_s = "20" + var_año_s
                              End If
                              var_fecha = var_dia_s + "-" + var_mes_s + "-" + var_año_s
                              
                              
                              
                              strconsulta = "select * from XXVIA_TB_ICG_TRAN_CEDIS_TIENDA where NUMB_ORGANIZATION_ID = ? and  VCHA_SUBINVENTORY_CODE = ? and VCHA_TRANSFER_SUBINVENTORY = ? and VCHA_NOTA_ENVIO = ? and VCHA_NUMERO_CAJA = ? and VCHA_CODIGO = ? "
                              With comandoORA
                                   .ActiveConnection = cnnicg
                                   .CommandType = adCmdText
                                   .CommandText = strconsulta
                                   Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(var_unidad_organizacional))
                                   .Parameters.Append parametro
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_almacen)
                                   .Parameters.Append parametro
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_almacen_icg)
                                   .Parameters.Append parametro
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, CStr(rsaux1!source_header_number))
                                   .Parameters.Append parametro
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, VAR_CAJA_S)
                                   .Parameters.Append parametro
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rsaux1!SEGMENT1)
                                   .Parameters.Append parametro
                              End With
                              Set rsaux9 = comandoORA.execute
                              Set comandoORA = Nothing
                              Set parametro = Nothing
                              If rsaux9.EOF Then
                                 strconsulta = "INSERT INTO XXVIA_TB_ICG_TRAN_CEDIS_TIENDA (NUMB_ORGANIZATION_ID, VCHA_SUBINVENTORY_CODE, VCHA_TRANSFER_SUBINVENTORY, DATE_FECHA, VCHA_NOTA_ENVIO, VCHA_NUMERO_CAJA, VCHA_CODIGO, NUMB_CANTIDAD, NUMB_STATUS, NUMB_ORG_ORIGEN) VALUES (?, ?, ?, SYSDATE, ?, ?, ?, ?, 3, ?)"
                                 With comandoORA
                                      .ActiveConnection = cnnicg
                                      .CommandType = adCmdText
                                      .CommandText = strconsulta
                                      Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(var_unidad_organizacional))
                                      .Parameters.Append parametro
                                      Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_almacen)
                                      .Parameters.Append parametro
                                      Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_almacen_icg)
                                      .Parameters.Append parametro
   '                                   Set parametro = .CreateParameter(, adDate, adParamInput, 200, CDate(var_fecha))
'                                      .Parameters.Append parametro
                                      Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, CStr(rsaux1!source_header_number))
                                      .Parameters.Append parametro
                                      Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, VAR_CAJA_S)
                                      .Parameters.Append parametro
                                      Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rsaux1!SEGMENT1)
                                      .Parameters.Append parametro
                                      Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, rsaux1!FLOA_SAL_CANTIDAD_LEIDA)
                                      .Parameters.Append parametro
                                      Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(var_unidad_organizacional))
                                      .Parameters.Append parametro
                                 End With
                                 Set rsaux8 = comandoORA.execute
                                 Set comandoORA = Nothing
                                 Set parametro = Nothing
                              End If
                              rsaux9.Close
                              rsaux1.MoveNext
                        Wend
                        rsaux1.MoveFirst
                        
                        strconsulta = "UPDATE XXVIA_TB_ICG_TRAN_CEDIS_TIENDA SET NUMB_STATUS = 0 where NUMB_ORGANIZATION_ID = ? and  VCHA_SUBINVENTORY_CODE = ? and VCHA_TRANSFER_SUBINVENTORY = ? and VCHA_NOTA_ENVIO = ? and NUMB_STATUS = 3"
                        With comandoORA
                             .ActiveConnection = cnnicg
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(var_unidad_organizacional))
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_almacen)
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_almacen_icg)
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, CStr(rsaux1!source_header_number))
                             .Parameters.Append parametro
                        End With
                        Set rsaux9 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        
'''''''''''''''''''''''' comienza traspaso de costales
                        x = 1
                        If x = 1 Then
                           var_posible_pedido = 1
                           var_pedido_tienda = Me.txt_embraue_nota_envio
                           If rsaux8.State = 1 Then
                              rsaux8.Close
                           End If
                           rsaux8.Open "SELECT * FROM TB_ORACLE_PEDIDOS_TIENDAS_COSTALES WHERE PEDIDO = " + CStr(var_pedido_tienda), cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux8.EOF Then
                              var_posible_pedido = 0
                           End If
                           rsaux8.Close
                           If var_posible_pedido = 1 Then
                              strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B,  OE_ORDER_HEADERS_ALL OHA Where requisition_header_id = OHA.SOURCE_DOCUMENT_ID AND secondary_inventory_name = A.ATTRIBUTE1 AND OHA.ORDER_NUMBER = ?"
                              With comandoORA
                                   .ActiveConnection = cnnoracle_4
                                   .CommandType = adCmdText
                                   .CommandText = strconsulta
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_pedido_tienda)
                                   .Parameters.Append parametro
                              End With
                              Set rsaux8 = comandoORA.execute
                              Set comandoORA = Nothing
                              Set parametro = Nothing
                              var_almacen_tienda = IIf(IsNull(rsaux8!attribute1), "", rsaux8!attribute1)
                              p_almacendestinofinal = var_almacen_tienda
                              rsaux8.Close
                              If var_almacen_tienda <> "" Then
                                 var_i = 0
                                 rsaux8.Open "SELECT XXVIA_SQ_LINEA_TM.nextval FROM dual", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 If Not rsaux8.EOF Then
                                    p_origenencabezadoid = rsaux8(0).Value
                                 End If
                                 rsaux8.Close
                                 rsaux8.Open "select XXVIA_SQ_ENCABEZADO_MT_ID.nextval from dual", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 If Not rsaux8.EOF Then
                                    P_ENCABEZADO_MT_ID = rsaux8(0).Value
                                 End If
                                 rsaux8.Close
                     
                                 strconsulta = "SELECT TIPO_CAJA, COUNT(*) AS CANTIDAD FROM XXVIA_VW_CAJAS_POR_PEDIDO WHERE SOURCE_HEADER_NUMBER = ? AND (TIPO_CAJA LIKE '%COSTAL%' OR TIPO_CAJA LIKE '%CAJA BIASI%') and INTE_EMB_EMBARQUE = " + Me.txt_embarque_pedido + " GROUP BY TIPO_CAJA"
                                 With comandoORA
                                     .ActiveConnection = cnnoracle_4
                                     .CommandType = adCmdText
                                     .CommandText = strconsulta
                                     Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, Me.txt_embraue_nota_envio)
                                     .Parameters.Append parametro
                                 End With
                                 Set rsaux6 = comandoORA.execute
                                 Set comandoORA = Nothing
                                 Set parametro = Nothing



                                 While Not rsaux6.EOF
                                       var_i = var_i + 1
                                       rs.Open "select * from tb_oracle_empaques where empaque = '" + rsaux6!tipo_caja + "' and codigo is not null", cnn, adOpenDynamic, adLockOptimistic
                                       If Not rs.EOF Then
                                          strconsulta = "select PRIMARY_UOM_CODE, INVENTORY_ITEM_ID from xxvia_system_items_b where SEGMENT1 = ? AND ORGANIZATION_ID = ?"
                                          With comandoORA
                                               .ActiveConnection = cnnoracle_4
                                               .CommandType = adCmdText
                                               .CommandText = strconsulta
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, rs!codigo)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, var_unidad_organizacional)
                                               .Parameters.Append parametro
                                          End With
                                          Set rsaux8 = comandoORA.execute
                                          Set comandoORA = Nothing
                                          Set parametro = Nothing
                                          var_inventory_item_id = rsaux8!inventory_item_id
                                          p_um = rsaux8!PRIMARY_UOM_CODE
                                          rsaux8.Close
                                          p_organizacion_id = var_unidad_organizacional
                                          p_organizacion_destino = var_unidad_organizacional
                                          If var_empresa = 92 Then
                                             p_subinventario = "CDI_ALMPT"
                                          End If
                                          If var_empresa = 83 Then
                                             p_subinventario = "TEX_PT_QL"
                                          End If
                                          p_subinventario_destino = "TRANS"
                                          p_codigoarticulo = rs!codigo
                                          p_cantidadorigen = rsaux6!cantidad
                                          p_Cantidadrecibida = 0
                                          p_origentransaccion = "SID_COSTALES_" + CStr(var_pedido_tienda)
                                          p_referencia_transaccion = var_pedido_tienda
                                          p_mensajeerror = ""
                                          strconsulta = "call xxvia_pk_inventarios.xxvia_sp_inventarios4 (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
                                          With comandoORA
                                               .ActiveConnection = cnnoracle_4
                                               .CommandType = adCmdText
                                               .CommandText = strconsulta
                                               Set parametro = .CreateParameter(, adInteger, adParamInput, 100, p_organizacion_id)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adInteger, adParamInput, 100, p_organizacion_destino)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, p_subinventario)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, p_subinventario_destino)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adInteger, adParamInput, 100, 2)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, p_codigoarticulo)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adInteger, adParamInput, 100, p_cantidadorigen)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adInteger, adParamInput, 100, p_cantidadorigen)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, p_origentransaccion)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adInteger, adParamInput, 100, p_origenencabezadoid)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adDouble, adParamInput, 100, Null)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adDouble, adParamInput, 100, Null)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, p_referencia_transaccion)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adDouble, adParamInput, 100, Null)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, p_almacendestinofinal)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adDate, adParamInput, 100, Date)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, Null)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, p_um)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adDouble, adParamInput, 100, P_ENCABEZADO_MT_ID)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_clave_usuario_global)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, fun_NombrePc)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, Null)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, "")
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, "")
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamOutput, 100, p_mensajeerror)
                                               .Parameters.Append parametro
                                          End With
                                          'MsgBox strconsulta
                                          Set rsaux9 = comandoORA.execute
                                          Set comandoORA = Nothing
                                          Set parametro = Nothing
                                          
                                       End If
                                       rs.Close
                                       rsaux6.MoveNext
                                 Wend
                                 strconsulta = "call xxvia_pk_inventarios.xxvia_valida_interface (1,?,?)"
                                 With comandoORA
                                      .ActiveConnection = cnnoracle_4
                                      .CommandType = adCmdText
                                      .CommandText = strconsulta
                                      Set parametro = .CreateParameter(, adDouble, adParamInput, 100, p_origenencabezadoid)
                                      .Parameters.Append parametro
                                      Set parametro = .CreateParameter(, adDouble, adParamInput, 200, 0)
                                      .Parameters.Append parametro
                                 End With
                                 rsaux9.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 On Error GoTo salir2
                                 Set rsaux9 = comandoORA.execute
                                 Set comandoORA = Nothing
                                 Set parametro = Nothing
                                 rsaux8.Open "insert into tb_oracle_pedidos_tiendas_costales (pedido) values (" + CStr(var_pedido_tienda) + ")", cnn, adOpenDynamic, adLockOptimistic
                              End If
                           End If
                        End If
''''''''''''''''''''''''  fin de traspaso de costales
                        If rsaux6.State = 1 Then
                           rsaux6.Close
                        End If
                        rsaux10.Open "SELECT * FROM TB_ORACLE_NOTAS_IMPRESAS_ICG WHERE PEDIDO = " + Me.txt_embraue_nota_envio, cnn, adOpenDynamic, adLockOptimistic
                        If rsaux10.EOF Then
                           'MsgBox cnnicg_sql.ConnectionString
                           cnn.CommandTimeout = 360
                           rsaux11.Open "INSERT INTO TB_ORACLE_NOTAS_IMPRESAS_ICG (PEDIDO, FECHA) VALUES (" + Me.txt_embraue_nota_envio + ",GETDATE())", cnn, adOpenDynamic, adLockOptimistic
                           x = 1
                           If x = 1 Then
                              If cnnicg_sql.State = 1 Then
                                 cnnicg_sql.Close
                              End If
                              cnnicg_sql.Open "Provider=SQLOLEDB.1;Password=icgfront2013;Persist Security Info=True;User ID=ICGAdmin;Initial Catalog=general;Data Source=sqlposprod"
                              rsaux9.Open "exec vyt_crea_pedido_cedis " + var_unidad_organizacional + ", '" + CStr(Me.txt_embraue_nota_envio) + "'", cnnicg_sql, adOpenDynamic, adLockOptimistic
                              rsaux9.Open "call xxpos.xxvia_pk_motor_logistico.xxvia_sp_senales_eviandas_a_cn (" + Me.txt_embraue_nota_envio + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           End If
                        End If
                        If rsaux10.State = 1 Then
                           rsaux10.Close
                        End If
                        If rsaux1.State = 1 Then
                           rsaux1.Close
                        End If
                        If cnnicg_sql.State = 1 Then
                           cnnicg_sql.Close
                        End If
                        End If
                     End If
                     rsaux11.Open "INSERT INTO TB_ORACLE_NOTAS_IMPRESAS_ICG (PEDIDO, FECHA) VALUES (" + Me.txt_embraue_nota_envio + ",GETDATE())", cnn, adOpenDynamic, adLockOptimistic
                     'error aqui
                     strconsulta = "select nvl(attribute7,'N') activa_pos_icg from mtl_secondary_inventories where secondary_inventory_name=?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_almacen_icg)
                          .Parameters.Append parametro
                     End With
                     Set rsaux9 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     If Not rsaux9.EOF Then
                        var_posible_cgi = IIf(IsNull(rsaux9(0).Value), "N", rsaux9(0))
                     Else
                        var_posible_cgi = "N"
                     End If
                     rsaux9.Close
                     var_posible_cgi = "Y"
                     If var_posible_cgi = "Y" Then
                        x = 0
                        If x = 1 Then
                           var_conexion_string_p = "Provider=SQLOLEDB.1;Password=icgfront2013;Persist Security Info=True;User ID=ICGAdmin;Initial Catalog=general;Data Source=sqlposprod"
                           If cnn_icg_posprod.State = 1 Then
                              cnn_icg_posprod.Close
                           End If
                           cnn_icg_posprod.Open var_conexion_string_p
                           cnn_icg_posprod.CommandTimeout = 360
                           'rsaux9.Open "exec [sqlposprod.vianney.com.mx].general.dbo.vyt_crea_pedido_cedis " + var_unidad_organizacional + ", '" + CStr(Me.txt_embraue_nota_envio) + "'", cnn, adOpenDynamic, adLockOptimistic
                           'MsgBox cnn_icg_posprod.ConnectionString
                           'rsaux9.Open "exec vyt_crea_pedido_cedis " + var_unidad_organizacional + ", '" + CStr(Me.txt_embraue_nota_envio) + "'", cnn_icg_posprod, adOpenDynamic, adLockOptimistic
                         End If
                     End If
                     
                     
                     
                     rsaux1.Open "DELETE FROM TB_ORACLE_CAJAS_EMBARQUES WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                     rsaux1.Open "select  inte_paq_caja, tipo_Caja, MAX(sello)  as cantidad, MAX(NVL(TRANSPORTE,' ')) AS TRANSPORTE from xxvia_tb_salidas_Cajas where source_header_number = " + Me.txt_embraue_nota_envio + " and inte_emb_embarque = " + Me.txt_embarque_pedido + " and floa_sal_cantidad_leida > 0 GROUP BY inte_paq_caja, tipo_Caja order by inte_paq_caja ", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     While Not rsaux1.EOF
                           If rsaux2.State = 1 Then
                              rsaux2.Close
                           End If
                           rsaux2.Open "INSERT INTO TB_ORACLE_CAJAS_EMBARQUES (INTE_TEM_CONSECUTIVO, PEDIDO, CAJA, TIPO_CAJA, SELLO, TRANSPORTE) VALUES (" + CStr(var_consecutivo) + "," + Me.txt_embraue_nota_envio + "," + CStr(rsaux1!INTE_PAQ_CAJA) + ",'" + IIf(IsNull(rsaux1!tipo_caja), "", rsaux1!tipo_caja) + "','" + IIf(IsNull(rsaux1!cantidad), "", rsaux1!cantidad) + "','" + rsaux1!transporte + "')", cnn, adOpenDynamic, adLockOptimistic
                           rsaux1.MoveNext
                     Wend
                     rsaux1.Close
                     x = 0
                     If x = 0 Then
                        'select a.description, ADDRESS_LINE_1, ADDRESS_LINE_2, TOWN_OR_CITY, REGION_1, POSTAL_CODE  into lv_nombre_subinventario, lv_direccion_1, lv_direccion_2 , lv_ciudad, lv_Estado, lv_cp  from mtl_secondary_inventories A, hr_locations_all B where A.ORGANIZATION_ID = LV_ORGANIZACION_DESTINO AND A.secondary_inventory_name = var_destino AND A.LOCATION_ID = B.LOCATION_ID;
                        
                        
                        If var_location_id > 0 Then
                           'strconsulta = "call XXVIA_SP_TIMBRAR_TRASPASOS(?,?,?)"
                           strconsulta = "call XXVIA_SP_TIMBRAR_TRASPASOS_3(?,?,?,?)"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, Me.txt_embraue_nota_envio)
                                .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 50, 3)
                                .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 50, var_unidad_organizacional)
                                .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 50, var_numero_embarque)
                                .Parameters.Append parametro
                           End With
                           Set rsaux2 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           var_serie = "TRX" + Me.txt_embarque_pedido + "_"
                           strconsulta = "select customer_trx_id, cadena as cadena, numero from xxvia_tb_control_doc_fiscales where serie = '" + var_serie + "' and numero = ? AND ORGANIZACION = ?"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, Me.txt_embraue_nota_envio)
                                .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, var_unidad_organizacional)
                                .Parameters.Append parametro
                           End With
                           Set rsaux2 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           If Not rsaux2.EOF Then
                              var_cadena = Replace(rsaux2!Cadena, " ", "")
                              var_cadena_rfc = Mid(var_cadena, 34, 12)
                              VAR_CADENA_STR = ""
                              Dim var_nN As Double
                              Open ("C:\SISTEMAS\TRX" + Trim(Me.txt_embarque_pedido) + "_" + Trim(Me.txt_embraue_nota_envio) + ".FAC") For Output As #1
                              VAR_Z = Len(var_cadena)
                              'For var_i = 1 To Len(var_cadena)
                              For var_nN = 1 To VAR_Z
                                  If Asc(Mid(var_cadena, var_nN, 1)) = 63 Then
                                     Print #1, VAR_CADENA_STR
                                     VAR_CADENA_STR = ""
                                  Else
                                     VAR_CADENA_STR = VAR_CADENA_STR + Mid(var_cadena, var_nN, 1)
                                  End If
                              Next var_nN
                              Print #1, "FIN:"
                              Close #1
                              var_archivo = "C:\SISTEMAS\sube_fact" + Trim(Str(Me.txt_embarque_pedido)) + "_" + Trim(Me.txt_embraue_nota_envio) + ".bat"
                              x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\TRX" + Me.txt_embarque_pedido + "_" + Trim(Me.txt_embraue_nota_envio) + ".FAC" + "|https://facturas2.vianney.mx/cgi-bin/cfds/timbrarGR33|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
                              rsaux2.Close
                              rsaux1.Open "select *  from tb_oracle_tiempo_impresion_documentos where pedido = " + Me.txt_embraue_nota_envio, cnn, adOpenDynamic, adLockOptimistic
                              If rsaux1.EOF Then
                                 strconsulta = "select oha.source_document_id, A.ATTRIBUTE1, B.description  from oe_order_headers_all oha, po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B  where order_number = ? and requisition_header_id = oha.source_document_id and secondary_inventory_name = A.ATTRIBUTE1"
                                 With comandoORA
                                      .ActiveConnection = cnnoracle_4
                                      .CommandType = adCmdText
                                      .CommandText = strconsulta
                                      Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, Me.txt_embraue_nota_envio)
                                      .Parameters.Append parametro
                                 End With
                                 Set rsaux2 = comandoORA.execute
                                 Set comandoORA = Nothing
                                 Set parametro = Nothing
                                 If Not rsaux2.EOF Then
                                    var_clave_almacen = rsaux2!attribute1
                                    var_nombre_almacen = rsaux2!Description
                                 Else
                                    var_clave_almacen = ""
                                    var_nombre_almacen = ""
                                 End If
                                 rsaux2.Close
                                 rsaux2.Open "insert into tb_oracle_tiempo_impresion_documentos (pedido,fecha, tienda, nombre) values (" + Me.txt_embraue_nota_envio + ",GETDATE(),'" + var_clave_almacen + "','" + var_nombre_almacen + "')", cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux1.Close
                           Else
                           
                           End If
                        'rsaux2.Close
                        Else
                           MsgBox "El subinventario " + var_almacen_icg + "   " + var_nombre_almacen_consignacioN + ", no tiene una dirección asignada, favor de validarlo con el departamento de costos o contraloria", vbOKOnly, "ATENCION"
                        End If
                     'Else
                        rsaux1.Open "select *  from tb_oracle_tiempo_impresion_documentos where pedido = " + Me.txt_embraue_nota_envio, cnn, adOpenDynamic, adLockOptimistic
                        If rsaux1.EOF Then
                           strconsulta = "select oha.source_document_id, A.ATTRIBUTE1, B.description  from oe_order_headers_all oha, po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B  where order_number = ? and requisition_header_id = oha.source_document_id and secondary_inventory_name = A.ATTRIBUTE1"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, Me.txt_embraue_nota_envio)
                                .Parameters.Append parametro
                           End With
                           Set rsaux2 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           If Not rsaux2.EOF Then
                              var_clave_almacen = rsaux2!attribute1
                              var_nombre_almacen = rsaux2!Description
                           Else
                              var_clave_almacen = ""
                              var_nombre_almacen = ""
                           End If
                           rsaux2.Close
                           rsaux2.Open "insert into tb_oracle_tiempo_impresion_documentos (pedido,fecha, tienda, nombre) values (" + Me.txt_embraue_nota_envio + ",GETDATE(),'" + var_clave_almacen + "','" + var_nombre_almacen + "')", cnn, adOpenDynamic, adLockOptimistic
                        End If
                        rsaux1.Close
                        var_x = 1
                        If var_x = 0 Then
                        rsaux1.Open "select distinct pedido from tb_Temp_oracle_NOTA_ENVIO where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux1.EOF
                              Set reporte = appl.OpenReport(App.Path + "\rep_oracle_nota_envio_linea.rpt")
                              reporte.RecordSelectionFormula = "{VW_ORACLE_NOTA_ENVIO_LINEA.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_ORACLE_NOTA_ENVIO_LINEA.cantidad} > 0 and {VW_ORACLE_NOTA_ENVIO_LINEA.pedido} = '" + CStr(rsaux1!pedido) + "'"
                              frmvistasprevias.cr.ReportSource = reporte
                              For ntablas = 1 To reporte.Database.Tables.Count
                                  reporte.Database.Tables(ntablas).SetLogOnInfo "admcdindustrial", var_bd_reportes, parametros(4), parametros(5)
                              Next ntablas
                              frmvistasprevias.cr.ViewReport
                              frmvistasprevias.Caption = "Nota de envio a tiendas"
                              frmvistasprevias.Show 1
                              Set reporte = Nothing
                              var_si = MsgBox("¿Desea el reporte a detalle?", vbYesNo, "ATENCION")
                              If var_si = 6 Then
                                 If var_consignacion = "PTO_CONS2" Then
                                    Set reporte = appl.OpenReport(App.Path + "\rep_oracle_nota_envio_detalle_consig.rpt")
                                    reporte.RecordSelectionFormula = "{VW_ORACLE_NOTA_ENVIO_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_ORACLE_NOTA_ENVIO_DETALLE.cantidad} > 0 and {VW_ORACLE_NOTA_ENVIO_DETALLE.pedido} = '" + CStr(rsaux1!pedido) + "'"
                                    frmvistasprevias.cr.ReportSource = reporte
                                    For ntablas = 1 To reporte.Database.Tables.Count
                                        reporte.Database.Tables(ntablas).SetLogOnInfo "admcdindustrial", var_bd_reportes, parametros(4), parametros(5)
                                    Next ntablas
                                    frmvistasprevias.cr.ViewReport
                                    frmvistasprevias.Caption = "Nota de envio a tiendas"
                                    frmvistasprevias.Show 1
                                    Set reporte = Nothing
                                 
                                    Set reporte = appl.OpenReport(App.Path + "\rep_oracle_nota_envio_detalle_consig.rpt")
                                    reporte.RecordSelectionFormula = "{VW_ORACLE_NOTA_ENVIO_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_ORACLE_NOTA_ENVIO_DETALLE.cantidad} > 0 and {VW_ORACLE_NOTA_ENVIO_DETALLE.pedido} = '" + CStr(rsaux1!pedido) + "'"
                                    For ntablas = 1 To reporte.Database.Tables.Count
                                        reporte.Database.Tables(ntablas).SetLogOnInfo "admcdindustrial", var_bd_reportes, parametros(4), parametros(5)
                                    Next ntablas
                                    reporte.ExportOptions.FormatType = crEFTExcel80
                                    reporte.ExportOptions.DestinationType = crEDTDiskFile
                                    archivo = "c:\reportessid\NOTA_ENVIO_" + Me.txt_embraue_nota_envio + "_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + ".xls"
                                    reporte.ExportOptions.DiskFileName = archivo
                                    reporte.Export False
                                    Set reporte = Nothing
                                    
                                    var_si = MsgBox("¿Desea enviar la nota por correo?", vbYesNo, "ATENCION")
                                    If var_si = 6 Then
                                       VAR_CORREO_ELECTRONICO = "vluna@vianney.com.mx"
                                       If Trim(VAR_CORREO_ELECTRONICO) <> "" Then
                                          If MAPISession1.SessionID = 0 Then
                                             MAPISession1.SignOn
                                          End If
                                          MAPIMessages1.SessionID = MAPISession1.SessionID
                                          MAPIMessages1.Compose
                                          MAPIMessages1.RecipDisplayName = VAR_CORREO_ELECTRONICO
                                          MAPIMessages1.RecipAddress = VAR_CORREO_ELECTRONICO
                                          MAPIMessages1.AddressResolveUI = True
                                          MAPIMessages1.ResolveName
                                          MAPIMessages1.MsgSubject = "Nota de envio " + Str(Me.txt_embraue_nota_envio)
                                          MAPIMessages1.MsgNoteText = "Se adjunta nota de envio número " + Str(Me.txt_embraue_nota_envio) + " del cliente " + var_nombre_almacen_consignacioN
                                          MAPIMessages1.AttachmentPathName = archivo
                                          MAPIMessages1.send True
                                          If MAPISession1.SessionID > 0 Then
                                             MAPISession1.SignOff
                                          End If
                                       End If
                                    End If
                                    
                                    MsgBox "Se a terminado de guardar el archivo " + archivo
                                 Else
                                    Set reporte = appl.OpenReport(App.Path + "\rep_oracle_nota_envio_detalle.rpt")
                                    reporte.RecordSelectionFormula = "{VW_ORACLE_NOTA_ENVIO_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_ORACLE_NOTA_ENVIO_DETALLE.cantidad} > 0 and {VW_ORACLE_NOTA_ENVIO_DETALLE.pedido} = '" + CStr(rsaux1!pedido) + "'"
                                    frmvistasprevias.cr.ReportSource = reporte
                                    For ntablas = 1 To reporte.Database.Tables.Count
                                        reporte.Database.Tables(ntablas).SetLogOnInfo "admcdindustrial", var_bd_reportes, parametros(4), parametros(5)
                                    Next ntablas
                                    frmvistasprevias.cr.ViewReport
                                    frmvistasprevias.Caption = "Nota de envio a tiendas"
                                    frmvistasprevias.Show 1
                                    Set reporte = Nothing
                                 End If
                              End If
                              rsaux1.MoveNext
                        Wend
                        rsaux1.Close
                        End If
                     End If
                  Else
                     MsgBox "No se a terminado de procesar la información en oracle, vuelva a intentar", vbOKOnly, "ATENCION"
                  End If
                  If rsaux1.State = 1 Then
                     rsaux1.Close
                  End If
                  rsaux1.Open "delete from tb_Temp_oracle_NOTA_ENVIO where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  'rsaux.Close
                  'Me.frm_embarque_nota_envio.Visible = False
               Else
                  MsgBox "No existen pedidos para el embarque", vbOKOnly, "ATENCION"
               End If
               rsaux.Close
            Else
               MsgBox "No se han generado todos los pedidos", vbOKOnly, "ATENCION"
            End If
            Else
               MsgBox "El pedido no existe o no corresponde al embarque seleccionado", vbOKOnly, "ATENCION"
            End If
            Else
               MsgBox "El pedido no ha sido cerrado", vbOKOnly, "ATENCION"
            End If
         Else
            rsaux.Close
            MsgBox "No existen movimientos para el embarque seleccionado", vbOKOnly, "ATENCION"
         End If
      End If
   End If
   If KeyAscii = 27 Then
      'Me.frm_embarque_nota_envio.Visible = False
   End If
   Exit Sub
salir2:
   If Err.Number = -2147217900 Then
      If rsaux10.State = 1 Then
         rsaux10.Close
         'MsgBox Err.Description
      End If
      rsaux10.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      rsaux10.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      'MsgBox Err.Description
      Resume
   End If

End Sub
