VERSION 5.00
Begin VB.Form frmoracle_factura_pedido 
   Caption         =   "Factura pedido"
   ClientHeight    =   630
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   2430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   630
   ScaleWidth      =   2430
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_pedido 
      Height          =   405
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Pedido:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   225
      Width           =   540
   End
End
Attribute VB_Name = "frmoracle_factura_pedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter

Private Sub txt_pedido_KeyPress(KeyAscii As Integer)
    On Error GoTo salir2:

   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_pedido) Then
         If rsaux15.State = 1 Then
            rsaux15.Close
         End If
         rsaux15.Open "select * from oe_order_headers_all where order_number = '" + Me.txt_pedido + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rsaux15.EOF Then
            If rsaux14.State = 1 Then
               rsaux14.Close
            End If
            rsaux14.Open "SELECT * FROM TB_ORACLE_PEDIDOS_CERRADOS WHERE PEDIDO = '" + Me.txt_pedido + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux14.EOF Then
               rsaux13.Open "update TB_ORACLE_PEDIDOS_CERRADOS set request_id = 0, customer_Trx_id = null where pedido = '" + Me.txt_pedido + "'", cnn, adOpenDynamic, adLockOptimistic
            Else
               rsaux13.Open "insert into tb_oracle_pedidos_cerrados (pedido, request_id) values ('" + Me.txt_pedido + "',0)", cnn, adOpenDynamic, adLockOptimistic
            End If
            rs.Open "select * from tb_oracle_pedidos_cerrados where pedido = '" + Me.txt_pedido + "'", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  If rsaux.State = 1 Then
                     rsaux.Close
                  End If
                  var_encontro = 0
                  While var_encontro = 0
                        var_cadena = "select * from ra_interface_lines_all where INTERFACE_LINE_ATTRIBUTE1 = ?"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = var_cadena
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                             .Parameters.Append parametro
                        End With
                        Set rsaux = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        If Not rsaux.EOF Then
                           var_cadena = "call xxvia_sp_FS_FACTURA_PEDIDO (?)"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = var_cadena
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rs!pedido))
                                .Parameters.Append parametro
                           End With
                           Set rsaux2 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           rsaux2.Open "UPDATE TB_ORACLE_PEDIDOS_CERRADOS SET REQUEST_ID = 1 WHERE PEDIDO =  " + CStr(rs!pedido), cnn, adOpenDynamic, adLockOptimistic
                        End If
                        var_cadena = "select * from ra_customer_Trx_all where ct_reference = ?"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = var_cadena
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                             .Parameters.Append parametro
                        End With
                        Set rsaux3 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        If Not rsaux3.EOF Then
                           var_encontro = 1
                        End If
                        rsaux3.Close
                  Wend
                  var_encontro = 0
                  While var_encontro = 0
                        var_cadena = "select * from ra_Customer_trx_all where ct_reference = ? and creation_date >= to_date('01/08/2017','DD/MM/YYYY')"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = var_cadena
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!pedido))
                             .Parameters.Append parametro
                        End With
                        Set rsaux2 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        If Not rsaux2.EOF Then
                        
                           var_customer_trx_id = rsaux2!customer_Trx_id
                           rsaux2.Close
                           
                           
                           var_cadena = "select nvl(b.name,' ') SERIE, replace(TRX_NUMBER,'_D',''), A.ORG_ID  from ra_customer_trx_all a, fnd_document_sequences b where customer_Trx_id = ? and a.doc_sequence_id = b.doc_sequence_id(+) and rownum = 1"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = var_cadena
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, var_customer_trx_id)
                                .Parameters.Append parametro
                           End With
                           Set rsaux4 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           If Not rsaux4.EOF Then
                              var_serie = IIf(IsNull(rsaux4!Serie), "", rsaux4!Serie)
                              rsaux4.Close
             
                              var_cadena = "SELECT * FROM XXVIA_TB_SERIES_33 WHERE SERIE = ?"
                              With comandoORA
                                   .ActiveConnection = cnnoracle_4
                                   .CommandType = adCmdText
                                   .CommandText = var_cadena
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_serie)
                                   .Parameters.Append parametro
                              End With
                              Set rsaux4 = comandoORA.execute
                              Set comandoORA = Nothing
                              Set parametro = Nothing
                              If Not rsaux4.EOF Then
                                 var_cadena = "CALL XXVIA_SP_CREA_EFLOW_33 (?)"
                                 With comandoORA
                                      .ActiveConnection = cnnoracle_4
                                      .CommandType = adCmdText
                                      .CommandText = var_cadena
                                      Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, var_customer_trx_id)
                                      .Parameters.Append parametro
                                 End With
                                 Set rsaux2 = comandoORA.execute
                                 Set comandoORA = Nothing
                                 Set parametro = Nothing
                                 rsaux2.Open "UPDATE TB_ORACLE_PEDIDOS_CERRADOS SET FECHA_FIN_EFLOW = GETDATE() WHERE PEDIDO =  " + CStr(rs!pedido), cnn, adOpenDynamic, adLockOptimistic
                              Else
                                 rsaux2.Open "UPDATE TB_ORACLE_PEDIDOS_CERRADOS SET REQUEST_ID = 1,FECHA_INICIO_EFLOW = GETDATE(), CUSTOMER_TRX_ID = " + CStr(var_customer_trx_id) + " WHERE PEDIDO =  " + CStr(rs!pedido), cnn, adOpenDynamic, adLockOptimistic
                                 var_cadena = "CALL XXVIA_SP_CREA_EFLOW (?)"
                                 With comandoORA
                                      .ActiveConnection = cnnoracle_4
                                      .CommandType = adCmdText
                                      .CommandText = var_cadena
                                      Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, var_customer_trx_id)
                                      .Parameters.Append parametro
                                 End With
                                 Set rsaux2 = comandoORA.execute
                                 Set comandoORA = Nothing
                                 Set parametro = Nothing
                                 rsaux2.Open "UPDATE TB_ORACLE_PEDIDOS_CERRADOS SET FECHA_FIN_EFLOW = GETDATE() WHERE PEDIDO =  " + CStr(rs!pedido), cnn, adOpenDynamic, adLockOptimistic
                              End If
                           Else
                              rsaux4.Close
                           End If
                           If rsaux4.State = 1 Then
                              rsaux4.Close
                           End If
                        Else
                           rsaux2.Close
                        End If
                        rsaux3.Open "select NUMERO, SERIE from xxvia_Tb_control_doc_fiscales where customer_trx_id = " + CStr(var_customer_trx_id), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux3.EOF Then
                           var_numero_factura = rsaux3!numero
                           var_serie = rsaux3!Serie
                           var_encontro = 1
                        End If
                  Wend
                  rs.MoveNext
            Wend
            rs.Close
            var_j = var_numero_factura
                  For var_j = var_numero_factura To var_numero_factura
                      rsaux1.Open "select customer_trx_id, cadena as cadena, numero from xxvia_tb_control_doc_fiscales where serie = '" + var_serie + "' and numero = " + CStr(var_numero_factura), cnnoracle_4, adOpenDynamic, adLockOptimistic
                      If rsaux1.EOF Then
                         var_posible = 1
                         MsgBox "No existe aun la factura " + var_serie + CStr(var_numero_factura) + ", ejecute nuevamente el concurrente", vbOKOnly, "ATENCION"
                      Else
                     
                         var_cadena = rsaux1!Cadena
                         var_cadena_rfc = Mid(var_cadena, 34, 12)
                         VAR_CADENA_STR = ""
                         Open ("C:\SISTEMAS\" + Trim(var_serie) + Trim(Str(var_numero_factura)) + ".FAC") For Output As #1
                         For var_i = 1 To Len(var_cadena)
                             'MsgBox Asc(Mid(var_cadena, var_i, 1))
                             If Asc(Mid(var_cadena, var_i, 1)) = 63 Then
                                If Trim(VAR_CADENA_STR) = "CONDICIONES_PAGO:- -" And Mid(var_serie, 1, 2) = "NC" Then
                                   VAR_CADENA_STR = "CONDICIONES_PAGO:NO IDENTIFICADO"
                                End If
                                'If Mid(var_cadena_str, 1, 6) = "FECHA:" Or Mid(var_cadena_str, 1, 16) = "FECHA:2014-08-08" Or Mid(var_cadena_str, 1, 16) = "FECHA:2014-08-09" Or Mid(var_cadena_str, 1, 16) = "FECHA:2014-08-10" Or Mid(var_cadena_str, 1, 16) = "FECHA:2014-08-11" Or Mid(var_cadena_str, 1, 16) = "FECHA:2014-08-11" Then
                                If Mid(VAR_CADENA_STR, 1, 6) = "FECHA:" Then
                                   'var_cadena_str = "FECHA:2016-01-31T23:59:00 "
                                   'var_cadena_str = "FECHA:2016-08-28T23:01:00 "
                                End If
                                Print #1, VAR_CADENA_STR
                                VAR_CADENA_STR = ""
                             Else
                                VAR_CADENA_STR = VAR_CADENA_STR + Mid(var_cadena, var_i, 1)
                             End If
                         Next var_i
                         Print #1, "FIN:"
                         Close #1
                         var_j = var_numero_factura
                         
                         var_archivo = "C:\SISTEMAS\sube_fact" + Trim(Str(var_j)) + ".bat"
                         
                         strconsulta = "SELECT SERIE FROM XXVIA_tB_cONTROL_DOC_FISCALES WHERE SERIE = ? and numero = ? AND NUMERO_TIENDA = '3.3'"
                         With comandoORA
                              .ActiveConnection = cnnoracle_4
                              .CommandType = adCmdText
                              .CommandText = strconsulta
                              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Trim(var_serie))
                              .Parameters.Append parametro
                              Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, var_j)
                              .Parameters.Append parametro
                        End With
                        Set rsaux9 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        If Not rsaux9.EOF Then
                           VAR_33 = 1
                        Else
                           VAR_33 = 0
                        End If
                        rsaux9.Close
                         If VAR_33 = 1 Then
                            x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\" + Trim(var_serie) + Trim(Str(var_j)) + ".FAC" + "|https://facturas2.vianney.mx/cgi-bin/cfds/timbrarGR33|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
                         Else
                            x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\" + Trim(var_serie) + Trim(Str(var_j)) + ".FAC" + "|https://facturas.vianney.mx/cgi-bin/cfds/timbrarGR|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
                         End If
                                          
                         'URL = "https://facturas.vianney.mx/cgi-bin/cfds/cfdsORACLE?cmd=download_pdf&rfc_emisor=" + var_cadena_rfc + "&serie=" + Trim(Me.txt_serie) + "&folio=" + Trim(CStr(var_j))
                         'buf = Split(URL, ".")
                         'ext = buf(UBound(buf))
                         'strSavePath = "C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(CStr(var_j)) + ".pdf"
                         'ret = URLDownloadToFile(0, URL, strSavePath, 0, 0)
                      End If
                      rsaux1.Close
                      
                      
                
                      Sleep 5000

                      Open "C:\sistemas\" + Trim(var_serie) + CStr(var_j) + ".txt" For Input As #1
                      'Open "C:\sistemas\FAEVXX_prueba" + CStr(var_j) + ".txt" For Input As #1
                      Dim Linea As String, Total As String
                      Do Until EOF(1)
                         Line Input #1, Linea
                         If Mid(Linea, 1, 2) <> "OK" Then
                            var_linea = ""
                            For var_jF = 1 To Len(Linea)
                                If Mid(Linea, var_jF, 16) = "Valor Esperado: " Then
                                   var_i = 0
                                   var_linea = ""
                                   VAR_Z = 1
                                   While var_i = 0
                                         If Mid(Linea, var_jF + 15 + VAR_Z, 1) <> " " Then
                                            var_linea = var_linea + Mid(Linea, var_jF + 15 + VAR_Z, 1)
                                            VAR_Z = VAR_Z + 1
                                         Else
                                            var_i = 1
                                         End If
                   
                                   Wend
                                   'MsgBox "cantidad esperada " + var_linea, vbOKOnly, ""
                                End If
                                If Mid(Linea, var_jF, 15) = "ValorEsperado: " Then
                                   var_i = 0
                                   var_linea = ""
                                   VAR_Z = 1
                                   While var_i = 0
                                         If Mid(Linea, var_jF + 14 + VAR_Z, 1) <> " " Then
                                            var_linea = var_linea + Mid(Linea, var_jF + 14 + VAR_Z, 1)
                                            VAR_Z = VAR_Z + 1
                                         Else
                                            var_i = 1
                                         End If
                   
                                   Wend
                                   'MsgBox "cantidad esperada " + var_linea, vbOKOnly, ""
                                End If
                                
                            Next var_jF
   
                            For var_jF = 1 To Len(Linea)
                                If Mid(Linea, var_jF, 17) = "Valor Reportado: " Then
                                   var_i = 0
                                   var_linea_2 = ""
                                   VAR_Z = 1
                                   While var_i = 0
                                         If Mid(Linea, var_jF + 15 + VAR_Z, 1) <> "<" Then
                                            var_linea_2 = var_linea_2 + Mid(Linea, var_jF + 15 + VAR_Z, 1)
                                            VAR_Z = VAR_Z + 1
                                         Else
                                            var_i = 1
                                         End If
                   
                                   Wend
                                   'MsgBox "cantidad esperada " + var_linea + " cantidad reportada " + VAR_LINEA_2, vbOKOnly, ""
                                End If
                                If Mid(Linea, var_jF, 16) = "ValorReportado: " Then
                                   var_i = 0
                                   var_linea_2 = ""
                                   VAR_Z = 1
                                   While var_i = 0
                                         'MsgBox Mid(Linea, var_jF + 14 + VAR_Z, 1)
                                         If Mid(Linea, var_jF + 14 + VAR_Z, 1) <> "a" Then
                                            var_linea_2 = var_linea_2 + Mid(Linea, var_jF + 15 + VAR_Z, 1)
                                            VAR_Z = VAR_Z + 1
                                         Else
                                            var_i = 1
                                         End If
                   
                                  Wend
                                  'MsgBox "cantidad esperada " + var_linea + " cantidad reportada " + VAR_LINEA_2, vbOKOnly, ""
                               End If
                               var_linea_2 = Replace(var_linea_2, " a", "")
                            Next var_jF
                            'MsgBox "Factura incorrecta", vbCritical, ""
                         End If
                      Loop
                      Close #1
                      If rsaux1.State = 1 Then
                         rsaux1.Close
                      End If
                      'var_serie = Me.txt_serie
                      var_cadena = "update xxvia_Tb_control_doc_fiscales set cadena = replace(cadena,'|" + Trim(var_linea_2) + "|','|" + Trim(var_linea) + "|') where serie = '" + Trim(var_serie) + "' and numero = " + CStr(var_j)
                      rsaux1.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                      rsaux1.Open "select customer_trx_id, cadena as cadena, numero from xxvia_tb_control_doc_fiscales where serie = '" + Trim(var_serie) + "' and numero = " + CStr(var_j), cnnoracle_4, adOpenDynamic, adLockOptimistic
                      var_cadena = Replace(rsaux1!Cadena, "T23:", "T00:")
                      var_cadena_rfc = Mid(var_cadena, 34, 12)
                      VAR_CADENA_STR = ""
                      Open ("C:\SISTEMAS\" + Trim(var_serie) + CStr(var_j) + ".FAC") For Output As #1
                      For var_i = 1 To Len(var_cadena)
                          'MsgBox Asc(Mid(var_cadena, var_i, 1))
                          If Asc(Mid(var_cadena, var_i, 1)) = 63 Then
                             If Mid(VAR_CADENA_STR, 1, 11) = "COMPROBANTE" Then
                                var_linea_2 = "|" + Trim(var_linea_2) + "|"
                                var_linea = "|" + Trim(var_linea) + "|"
                                VAR_CADENA_STR = Replace(VAR_CADENA_STR, var_linea_2, var_linea)
                             End If
                             'MsgBox VAR_CADENA_STR
                             Print #1, VAR_CADENA_STR
                             VAR_CADENA_STR = ""
                          Else
                             VAR_CADENA_STR = VAR_CADENA_STR + Mid(var_cadena, var_i, 1)
                          End If
                      Next var_i
                      Print #1, "FIN:"
                      Close #1
                      rsaux1.Close
                      If Trim(var_serie) <> "FAEVII" Then
                         If VAR_33 = 1 Then
                            x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\" + Trim(Trim(var_serie)) + Trim(Str(var_j)) + ".FAC" + "|https://facturas2.vianney.mx/cgi-bin/cfds/timbrarGR33|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
                         Else
                            x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\" + Trim(Trim(var_serie)) + Trim(Str(var_j)) + ".FAC" + "|https://facturas.vianney.mx/cgi-bin/cfds/timbrarGR|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
                         End If
                      
                      
                      Else
                         MsgBox "Selecciona la opción de exporta", vbOKOnly, "ATENCION"
                      End If
                      
                      
                      
                  Next var_j
            
            
            
            'MsgBox "Imprimir factura " + VAR_SERIE + CStr(var_numero_factura), vbOKOnly, "ATENCION"
            
            
            
            
            
            
            rsaux14.Close
         Else
            MsgBox "El pedido no existe", vbOKOnly, "ATENCION"
         End If
         rsaux15.Close
      End If
   End If
Exit Sub
salir2:
   'MsgBox Err.Description
   If Err.Number = -2147217900 Then
      'MsgBox Err.Description
      rsaux2.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      rsaux2.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      Resume
   End If
   'MsgBox Err.Description
   
   
End Sub
