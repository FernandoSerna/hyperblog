VERSION 5.00
Begin VB.Form frmoracle_cerrar_embarque 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cerrar embarques"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   2835
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2445
      Picture         =   "frmoracle_cerrar_embarque.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_cerrar_embarque 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   375
      Picture         =   "frmoracle_cerrar_embarque.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cerrar Embarque"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   45
      Picture         =   "frmoracle_cerrar_embarque.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   45
      TabIndex        =   1
      Top             =   420
      Width           =   2730
      Begin VB.TextBox txt_embarque 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   135
         TabIndex        =   2
         Top             =   180
         Width           =   2445
      End
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   30
      TabIndex        =   0
      Top             =   330
      Width           =   2820
   End
End
Attribute VB_Name = "frmoracle_cerrar_embarque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objConn As New ADODB.Connection
Dim objCmd As New ADODB.Command
Dim objParm As ADODB.Parameter
Dim clnt As New SoapClient30
Dim var_arreglo() As String
Dim var_trip_id As String
Dim var_b As Boolean
Dim var_con As String


Private Sub cmd_cerrar_embarque_Click()
   Dim clnt As New SoapClient30
   Dim objConn As New ADODB.Connection
   Dim objCmd As New ADODB.Command
   Dim objParm As ADODB.Parameter
   rsaux.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      VAR_X_TRIP_ID = rs!ARREGLO_0
      var_x_trip_name = rs!ARREGLO_1
      var_estatus = IIf(IsNull(rs!char_emb_estatus), "", rs!char_emb_estatus)
      If IIf(IsNull(rs!char_emb_estatus), "", rs!char_emb_estatus) = "I" Then
         If rs!tipo_embarque = 1 Then
            rsaux.Open "select distinct source_header_number from xxvia_tb_salidas where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
         End If
         If rs!tipo_embarque = 2 Then
            rsaux.Open "select distinct source_header_number from xxvia_tb_SAlidas_cajas where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
         End If
         VAR_CADENA_PEDIDOS_M = ""
         While Not rsaux.EOF
               If VAR_CADENA_PEDIDOS_M = "" Then
                  VAR_CADENA_PEDIDOS_M = CStr(rsaux!source_header_number)
               Else
                  VAR_CADENA_PEDIDOS_M = VAR_CADENA_PEDIDOS_M + ", " + CStr(rsaux!source_header_number)
               End If
               rsaux.MoveNext
         Wend
         VAR_CADENA_PEDIDOS = ""
         rsaux.MoveFirst
         While Not rsaux.EOF
               rsaux1.Open "select distinct delivery_id from wsh_deliverables_v where SOURCE_HEADER_NUMBER = " + CStr(rsaux!source_header_number) + " AND delivery_id IS NOT NULL", cnnoracle_4, adOpenDynamic, adLockOptimistic
               VAR_ENTREGA = rsaux1!delivery_id
               rsaux1.Close
               rsaux1.Open "select distinct source_header_number from wsh_deliverables_v where delivery_id = " + CStr(VAR_ENTREGA), cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux1.EOF Then
                  var_j = 0
                  While Not rsaux1.EOF
                        var_j = var_j + 1
                        rsaux1.MoveNext
                  Wend
                  If var_j > 1 Then
                     If VAR_CADENA_PEDIDOS = "" Then
                        VAR_CADENA_PEDIDOS = CStr(rsaux!source_header_number) + " ENTREGA: " + CStr(VAR_ENTREGA)
                     Else
                        VAR_CADENA_PEDIDOS = VAR_CADENA_PEDIDOS + ", " + CStr(rsaux!source_header_number) + " ENTREGA: " + CStr(VAR_ENTREGA)
                     End If
                  End If
               End If
               rsaux1.Close
               rsaux.MoveNext
         Wend
         rsaux.MoveFirst
         
         If VAR_CADENA_PEDIDOS <> "" Then
            MsgBox "Los pedidos siguientes tienen dos entregas " + VAR_CADENA_PEDIDOS
         Else
            cnn.BeginTrans
            rsaux8.Open "SELECT MAX(CONSECUTIVO) FROM TB_ORACLE_COMPARACION_PEDIDO_AFECTACION", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux8.EOF Then
               var_consecutivo = IIf(IsNull(rsaux8(0).Value), 0, rsaux8(0).Value) + 1
            Else
               var_consecutivo = 1
            End If
            rsaux8.Close
            rsaux8.Open "insert into TB_ORACLE_COMPARACION_PEDIDO_AFECTACION (consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            
            
            
            rsaux8.Open "SELECT inte_emb_embarque, SOURCE_HEADER_NUMBER, SUM(FLOA_sal_cANTIDAD_LEIDA) AS CANTIDAD FROM XXVIA_TB_sALIDAS where source_header_number in (" + VAR_CADENA_PEDIDOS_M + ") GROUP BY inte_emb_embarque, SOURCE_HEADER_NUMBER", cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rsaux8.EOF
                  rsaux2.Open "select * from xxvia_tb_encabezado_embarques where embarque = " + CStr(rsaux8!INTE_EMB_EMBARQUE), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux2.EOF Then
                     rsaux10.Open "insert INTO TB_ORACLE_COMPARACION_PEDIDO_AFECTACION (PEDIDO, CANTIDAD_LEIDA, CANTIDAD_AFECTADA, fecha, CONSECUTIVO) VALUES (" + CStr(rsaux8!source_header_number) + "," + CStr(rsaux8!Cantidad) + ",0, '" + CStr(rsaux2!FECHA_INICIO) + "'," + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                  Else
                     rsaux10.Open "insert INTO TB_ORACLE_COMPARACION_PEDIDO_AFECTACION (PEDIDO, CANTIDAD_LEIDA, CANTIDAD_AFECTADA, fecha, CONSECUTIVO) VALUES (" + CStr(rsaux8!source_header_number) + "," + CStr(rsaux8!Cantidad) + ",0, ''," + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux2.Close
                  rsaux8.MoveNext
            Wend
            rsaux8.Close
            rsaux8.Open "SELECT inte_emb_embarque, SOURCE_HEADER_NUMBER, SUM(FLOA_sal_cANTIDAD_LEIDA) AS CANTIDAD FROM XXVIA_TB_sALIDAS_CAJAS where source_header_number in (" + VAR_CADENA_PEDIDOS_M + ") GROUP BY inte_emb_embarque, SOURCE_HEADER_NUMBER", cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rsaux8.EOF
                  rsaux2.Open "select * from xxvia_tb_encabezado_embarques where embarque = " + CStr(rsaux8!INTE_EMB_EMBARQUE), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux2.EOF Then
                     rsaux10.Open "insert INTO TB_ORACLE_COMPARACION_PEDIDO_AFECTACION (PEDIDO, CANTIDAD_LEIDA, CANTIDAD_AFECTADA, fecha, CONSECUTIVO) VALUES (" + CStr(rsaux8!source_header_number) + "," + CStr(rsaux8!Cantidad) + ",0, '" + CStr(rsaux2!FECHA_INICIO) + "'," + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                  Else
                     rsaux10.Open "insert INTO TB_ORACLE_COMPARACION_PEDIDO_AFECTACION (PEDIDO, CANTIDAD_LEIDA, CANTIDAD_AFECTADA, fecha, CONSECUTIVO) VALUES (" + CStr(rsaux8!source_header_number) + "," + CStr(rsaux8!Cantidad) + ",0, ''," + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux2.Close
                  rsaux8.MoveNext
            Wend
            rsaux8.Close
            rsaux8.Open "SELECT pedido FROM TB_ORACLE_COMPARACION_PEDIDO_AFECTACION WHERE CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux8.EOF
                  rsaux1.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  rsaux10.Open "SELECT SOURCE_HEADER_NUMBER, SUM(SHIPPED_QUANTITY) AS CANTIDAD FROM WSH_DELIVERABLES_V WHERE SOURCE_HEADER_NUMBER = " + CStr(IIf(IsNull(rsaux8!pedido), 0, rsaux8!pedido)) + " GROUP BY SOURCE_HEADER_NUMBER", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux10.EOF Then
                     rsaux1.Open "UPDATE TB_ORACLE_COMPARACION_PEDIDO_AFECTACION SET CANTIDAD_AFECTADA = " + CStr(IIf(IsNull(rsaux10!Cantidad), 0, rsaux10!Cantidad)) + " WHERE PEDIDO = " + CStr(rsaux8!pedido) + " AND CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux10.Close
                  rsaux8.MoveNext
            Wend
            rsaux8.Close
            rsaux8.Open "SELECT *  FROM TB_ORACLE_COMPARACION_PEDIDO_AFECTACION where cantidad_afectada > 0 and CANTIDAD_LEIDA <> cantidad_afectada AND CONSECUTIVO = " + CStr(var_consecutivo) + " order by PEDIDO desc "
            If Not rsaux8.EOF Then
               var_cadena_pedidos_mal = ""
               While Not rsaux8.EOF
                     If var_cadena_pedidos_mal = "" Then
                        var_cadena_pedidos_mal = CStr(IIf(IsNull(rsaux8!pedido), 0, rsaux8!pedido))
                     Else
                        var_cadena_pedidos_mal = var_cadena_pedidos_mal + ", " + CStr(IIf(IsNull(rsaux8!pedido), 0, rsaux8!pedido))
                     End If
                     rsaux8.MoveNext
               Wend
               MsgBox "Los siguientes pedidos tienen errores entra la cantidad leida y la cantidad afectada: " + CStr(var_cadena_pedidos_mal), vbOKOnly, "ATENCION"
            Else
               clnt.MSSoapInit var_webservice
               While Not rsaux.EOF
                     rsaux1.Open "select distinct delivery_id from wsh_deliverables_v where SOURCE_HEADER_NUMBER = " + CStr(rsaux!source_header_number) + " AND delivery_id IS NOT NULL", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     VAR_ENTREGA = rsaux1!delivery_id
                     rsaux1.Close
                     rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     var_estatus = 0
                     On Error GoTo salir:
                     'var_arreglo = clnt.ASIGNAR_embarque(VAR_ENTREGA, Val(VAR_X_TRIP_ID), "CONFIRM")
                     
                     
                        objConn.Open var_conexion_oracle
                        '… Establecer conexión a la base de datos con el objeto objConn.
                        With objCmd
                             objConn.BeginTrans
                             .ActiveConnection = objConn
                             .CommandText = "xxvia_pk_interfaces_om.asignar_embarque"
                             .CommandType = adCmdStoredProc
                                       
                             'p_organization_id IN NUMBER
                             Set objParm = .CreateParameter("p_api_version_number", adNumeric, adParamInput, , 1)
                             .Parameters.Append objParm
                             
                             'MsgBox rsaux1!sold_to_org_id
                             'p_customer_id IN number,
                             Set objParm = .CreateParameter("p_action_code", adVarChar, adParamInput, 50, "CONFIRM")
                             .Parameters.Append objParm
            
                             'p_devolucion_sid IN VARCHAR2,
                             Set objParm = .CreateParameter("p_delivery_id", adNumeric, adParamInput, , VAR_ENTREGA)
                             .Parameters.Append objParm
                             
                             ' x_header_interface_id out number,
                             Set objParm = .CreateParameter("p_asg_trip_id", adNumeric, adParamInput, , Val(VAR_X_TRIP_ID))
                             .Parameters.Append objParm
                               
                             'x_group_id IN VARCHAR,
                             Set objParm = .CreateParameter("x_trip_id", adVarNumeric, adParamOutput, 50, 0)
                             .Parameters.Append objParm
                             
                             Set objParm = .CreateParameter("x_trip_name", adVarChar, adParamOutput, 50, "")
                             .Parameters.Append objParm
                             
                             .execute
                             objConn.CommitTrans
                        End With
                        'MsgBox var_conexion_oracle
                        Set objConn = Nothing
                        Set objCmd = Nothing
                     
                     
                     
                     
                     
                     rsaux1.Open "insert into tb_oracle_pedidos_confirmados (pedido, fecha, maquina, error) values (" + CStr(rsaux!source_header_number) + ", getdate(), '" + fun_NombrePc + "'," + CStr(var_estatus) + ")", cnn, adOpenDynamic, adLockOptimistic
                     rsaux.MoveNext
               Wend
               Set clnt = Nothing
               MsgBox "Se termino de cerrar el embarque", vbOKOnly, "ATENCION"
            End If
            rsaux8.Close
         End If
         rsaux.Close
      Else
         If var_estatus = "F" Then
            MsgBox "EL embarque ya fue facturado"
         Else
            MsgBox "El embarque NO a sido cerrado", vbOKOnly, "ATENCION"
         End If
      End If
   End If
   rs.Close
   Exit Sub
salir:
   If Err.Number = -2147217900 Then
      'MsgBox Err.Description
      rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      Resume
   Else
      If Err.Number = -2147467259 Then
         'MsgBox Err.Description
         Resume Next
         var_estatus = 1
      Else
         MsgBox Err.Description
      End If
   End If
   
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_embarque = ""
   Me.txt_embarque.SetFocus
End Sub

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub txt_cerrar_embarque_Change()

End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 3850
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_cerrar_embarque.SetFocus
   End If
End Sub
