VERSION 5.00
Begin VB.Form frmoracle_embarques_cerrados_pedidos_abiertos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Embarques con pedidos abiertos"
   ClientHeight    =   960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   825
      Left            =   75
      TabIndex        =   0
      Top             =   30
      Width           =   3360
      Begin VB.TextBox txt_embarque 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1830
         TabIndex        =   1
         Top             =   225
         Width           =   1425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Embarque:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   225
         TabIndex        =   2
         Top             =   285
         Width           =   1560
      End
   End
End
Attribute VB_Name = "frmoracle_embarques_cerrados_pedidos_abiertos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comandoORA As New adodb.Command
Dim parametro As adodb.Parameter

Private Sub Form_Load()
   Top = 3300
   Left = 4250
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
   If IsNumeric(Me.txt_embarque) Then
      var_cadena = "SELECT * FROM XXVIA_tB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = ?"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = var_cadena
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_embarque))
           .Parameters.Append parametro
      End With
      Set rs = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rs.EOF Then
         var_estatus = IIf(IsNull(rs!char_emb_estatus), "", rs!char_emb_estatus)
         If var_estatus = "I" Or var_estatus = "F" Then
            rsaux.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_cadena = "SELECT distinct source_header_number from xxvia_tb_Salidas_cajas WHERE inte_emb_EMBARQUE = ?"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = var_cadena
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_embarque))
                 .Parameters.Append parametro
            End With
            Set rsaux = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            var_cadena_pedidos = ""
            While Not rsaux.EOF
                  If var_cadena_pedidos = "" Then
                     var_cadena_pedidos = CStr(rsaux!SOURCE_HEADER_NUMBER)
                  Else
                     var_cadena_pedidos = var_cadena_pedidos + "," + CStr(rsaux!SOURCE_HEADER_NUMBER)
                  End If
                  rsaux.MoveNext
            Wend
            rsaux.Close
            If var_cadena_pedidos <> "" Then
               rsaux.Open "SELECT source_header_number, sum(src_requested_quantity) as cantidad FROM WSH_DELIVERABLES_V WHERE SOURCE_HEADER_NUMBER in (" + var_cadena_pedidos + ") and released_status = 'Y' group by source_header_number", cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  var_cadena = ""
                  While Not rsaux.EOF
                        If var_cadena = "" Then
                           var_cadena = "Pedido " + CStr(rsaux!SOURCE_HEADER_NUMBER) + " cantidad " + CStr(rsaux!Cantidad)
                        Else
                           var_cadena = var_cadena + ", Pedido " + CStr(rsaux!SOURCE_HEADER_NUMBER) + " cantidad " + CStr(rsaux!Cantidad)
                        End If
                        rsaux.MoveNext
                  Wend
                  var_si = MsgBox("Faltan pedidos por cerrar: " + var_cadena + " ¿Desea abrir el embarque para volverlo a cerrar?", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     var_cadena = "update xxvia_Tb_encabezado_embarques set char_emb_estatus = 'E', maquina = ? where  EMBARQUE = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = var_cadena
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, fun_NombrePc)
                          .Parameters.Append parametro
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_embarque))
                          .Parameters.Append parametro
                     End With
                     Set rsaux1 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     
                     
                     var_cadena = "select distinct source_header_number from xxvia_tb_salidas_cajas where  inte_emb_EMBARQUE = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = var_cadena
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_embarque))
                          .Parameters.Append parametro
                     End With
                     Set rsaux1 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                                          
                     While Not rsaux1.EOF
                           On Error GoTo salir2:
                           var_cadena = "call xxvia_sp_act_det_pedido_2 (?)"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = var_cadena
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rsaux1!SOURCE_HEADER_NUMBER))
                                .Parameters.Append parametro
                           End With
                           Set rsaux2 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           rsaux1.MoveNext
                     Wend
                     rsaux1.Close
                                          
                     rsaux2.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     var_cadena = "select SOURCE_HEADER_NUMBER, delivery_detail_id from xxvia_Tb_salidas_cajas where  inte_emb_EMBARQUE = ? and floa_sal_cantidad_leida = 0"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = var_cadena
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_embarque))
                          .Parameters.Append parametro
                     End With
                     Set rsaux1 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     var_cadena_ceros = ""
                     While Not rsaux1.EOF
                           var_cadena = "select released_status, SHIPPED_QUANTITY from WSH_DELIVERABLES_V where delivery_detail_id = ?"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = var_cadena
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rsaux1!delivery_detail_id))
                                .Parameters.Append parametro
                           End With
                           Set rsaux2 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           If Not rsaux2.EOF Then
                              var_cantidad = IIf(IsNull(rsaux2!SHIPPED_QUANTITY), -3, rsaux2!SHIPPED_QUANTITY)
                           Else
                              var_cantidad = -3
                           End If
                           If rsaux2!released_status = "Y" Then
                              If var_cantidad = -3 Then
                                  If var_cadena_ceros = "" Then
                                     var_cadena_ceros = CStr(rsaux1!SOURCE_HEADER_NUMBER)
                                  Else
                                     var_cadena_ceros = var_cadena_ceros + ", " + CStr(rsaux1!SOURCE_HEADER_NUMBER)
                                  End If
                              End If
                           End If
                           rsaux1.MoveNext
                     Wend
                     rsaux1.Close
                     If var_cadena_ceros <> "" Then
                        var_cadena = "update xxvia_Tb_encabezado_embarques set char_emb_estatus = 'I', maquina = ? where  EMBARQUE = ?"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = var_cadena
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, fun_NombrePc)
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_embarque))
                             .Parameters.Append parametro
                        End With
                        Set rsaux1 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        
                        MsgBox "El embarque no puede ser cerrado porque los siguientes pedidos no estan correctos " + var_cadena_ceros, vbOKOnly, "ATENCION"
                     Else
                        MsgBox "Vuelva a cerrar el embarque", vbOKOnly, "ATENCION"
                     End If
                  End If
               Else
                  MsgBox "No existen pedidos abiertos", vbOKOnly, "ATENCION"
               End If
               rsaux.Close
               
            Else
               MsgBox "El embarque no tiene pedidos asignados"
            End If
         Else
            MsgBox "El embarque no a sido cerrado", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      MsgBox "Número de embarque incorrecto", vbOKOnly, "ATENCION"
   End If
   End If
   Exit Sub
salir2:
   'MsgBox Err.Description
   If Err.Number = -2147217900 Then
      'MsgBox Err.Description
      rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      'MsgBox Err.Description
      Resume
   End If
   
End Sub
