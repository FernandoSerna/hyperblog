VERSION 5.00
Begin VB.Form frmoracle_cargar_pedido_guatemala 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cargar pedido Guatemala"
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_pedido 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1230
      TabIndex        =   1
      Top             =   120
      Width           =   2385
   End
   Begin VB.TextBox txt_embarque 
      Height          =   555
      Left            =   150
      TabIndex        =   0
      Top             =   105
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Pedido:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   165
      TabIndex        =   2
      Top             =   180
      Width           =   975
   End
End
Attribute VB_Name = "frmoracle_cargar_pedido_guatemala"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter

Private Sub Form_Load()
   Top = 3200
   Left = 4200

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_pedido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_pedido) Then
         strconsulta = "select ship_to_org_id from oe_order_headers_all where order_number = ?"
         With comandoORA
              .ActiveConnection = cnnicg
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_pedido)
              .Parameters.Append parametro
         End With
         Set rsaux11 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         If Not rsaux11.EOF Then


            If rsaux11!ship_to_org_id = 694559 Or rsaux11!ship_to_org_id = 985388 Or rsaux11!ship_to_org_id = 1019346 Or rsaux11!ship_to_org_id = 1061430 Or rsaux11!ship_to_org_id = 1060967 Or rsaux11!ship_to_org_id = 1060927 Then
               If cnnicg_sql.State = 1 Then
                  cnnicg_sql.Close
               End If
               var_almacen = "CDI_ALMPT"

               If rsaux11!ship_to_org_id = 694559 Then
                  var_consignacion = "PTO_VTA"
                  var_almacen_icg = "VGT_TD6501"
               End If
               If rsaux11!ship_to_org_id = 985388 Then
                  var_consignacion = "PTO_VTA"
                  var_almacen_icg = "VGT_TD6502"
               End If
               If rsaux11!ship_to_org_id = 1019346 Then
                  var_consignacion = "PTO_VTA"
                  var_almacen_icg = "VGT_TD6503"
               End If
               If rsaux11!ship_to_org_id = 1061430 Then
                  var_consignacion = "PTO_VTA"
                  var_almacen_icg = "CRI_PAVAS"
               End If
               If rsaux11!ship_to_org_id = 1060967 Then
                  var_consignacion = "PTO_VTA"
                  var_almacen_icg = "CRI_COLON"
               End If
               If rsaux11!ship_to_org_id = 1060927 Then
                  var_consignacion = "PTO_VTA"
                  var_almacen_icg = "VGT_TD6504"
               End If
               
               cnnicg_sql.Open "Provider=SQLOLEDB.1;Password=icgfront2013;Persist Security Info=True;User ID=ICGAdmin;Initial Catalog=general;Data Source=sqlposprod"
                        
               rsaux1.Open "SELECT source_header_number, inte_paq_caja, segment1, sum(floa_sal_Cantidad_leida) as FLOA_SAL_CANTIDAD_LEIDA FROM XXVIA_TB_SALIDAS_CAJAS WHERE SOURCE_HEADER_NUMBER = " + Me.txt_pedido + " AND FLOA_SAL_CANTIDAD_LEIDA >0 group by source_header_number, inte_paq_caja, segment1", cnnoracle_4, adOpenDynamic, adLockOptimistic
               While Not rsaux1.EOF
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
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(93))
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
                        strconsulta = "INSERT INTO XXVIA_TB_ICG_TRAN_CEDIS_TIENDA (NUMB_ORGANIZATION_ID, VCHA_SUBINVENTORY_CODE, VCHA_TRANSFER_SUBINVENTORY, DATE_FECHA, VCHA_NOTA_ENVIO, VCHA_NUMERO_CAJA, VCHA_CODIGO, NUMB_CANTIDAD, NUMB_STATUS) VALUES (?, ?, ?, SYSDATE, ?, ?, ?, ?, 3)"
                        With comandoORA
                             .ActiveConnection = cnnicg
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             If var_almacen_icg = "CRI_PAVAS" Or var_almacen_icg = "CRI_COLON" Then
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(10))
                             Else
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(281))
                             End If
                             .Parameters.Append parametro
                             If var_almacen_icg = "CRI_PAVAS" Or var_almacen_icg = "CRI_COLON" Then
                             
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_almacen_icg)
                                .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_almacen)
                                .Parameters.Append parametro
                             
                             Else
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_almacen)
                                .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_almacen_icg)
                                .Parameters.Append parametro
                             
                             End If
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, CStr(rsaux1!source_header_number))
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, VAR_CAJA_S)
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rsaux1!SEGMENT1)
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, rsaux1!FLOA_SAL_CANTIDAD_LEIDA)
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
                    If var_almacen_icg = "CRI_PAVAS" Or var_almacen_icg = "CRI_COLON" Then
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(10))
                    Else
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(281))
                    End If
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
               If rsaux10.State = 1 Then
                  rsaux10.Close
               End If
               rsaux10.Open "SELECT * FROM TB_ORACLE_NOTAS_IMPRESAS_ICG WHERE PEDIDO = " + Me.txt_pedido, cnn, adOpenDynamic, adLockOptimistic
               If rsaux10.EOF Then
                  cnn.CommandTimeout = 360
                  rsaux12.Open "INSERT INTO TB_ORACLE_NOTAS_IMPRESAS_ICG (PEDIDO, FECHA) VALUES (" + Me.txt_pedido + ",GETDATE())", cnn, adOpenDynamic, adLockOptimistic
                  If var_almacen_icg = "CRI_PAVAS" Or var_almacen_icg = "CRI_COLON" Then
                     rsaux9.Open "exec [sqlposprod.vianney.com.mx].general.dbo.vyt_crea_pedido_cedis 10, '" + CStr(rsaux1!source_header_number) + "'", cnn, adOpenDynamic, adLockOptimistic
                  Else
                     rsaux9.Open "exec [sqlposprod.vianney.com.mx].general.dbo.vyt_crea_pedido_cedis 281, '" + CStr(rsaux1!source_header_number) + "'", cnn, adOpenDynamic, adLockOptimistic
                  End If
                  'rsaux9.Open "call xxpos.xxvia_pk_motor_logistico.xxvia_sp_senales_eviandas_a_cn (" + Me.txt_embraue_nota_envio + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
               End If
               rsaux1.Close
               rsaux10.Close
               MsgBox "Se a terminado de cargar el pedido", vbOKOnly, "ATENCION"
            Else
               MsgBox "El pedido no pertenece a Guatemala", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El pedido no existe", vbOKOnly, "ATENCION"
         End If
         rsaux11.Close
      Else
          MsgBox "Número de pedido incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
End Sub
