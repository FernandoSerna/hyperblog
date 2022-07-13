VERSION 5.00
Begin VB.Form frmoracle_impresion_documentos_fiscales 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresión de documentos fiscales"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_carta_porte 
      Caption         =   "Carta Porte"
      Height          =   315
      Left            =   240
      TabIndex        =   29
      Top             =   30
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   4800
      TabIndex        =   28
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txt_cfdi 
      Height          =   375
      Left            =   240
      TabIndex        =   27
      Top             =   3240
      Width           =   5895
   End
   Begin VB.TextBox txt_factura 
      Height          =   405
      Left            =   1680
      TabIndex        =   26
      Top             =   2760
      Width           =   2535
   End
   Begin VB.ComboBox cmb_metodo_2 
      Height          =   315
      ItemData        =   "frmoracle_impresion_documentos_fiscales.frx":0000
      Left            =   960
      List            =   "frmoracle_impresion_documentos_fiscales.frx":000A
      TabIndex        =   22
      Top             =   2040
      Width           =   3315
   End
   Begin VB.ComboBox cmb_uso 
      Height          =   315
      ItemData        =   "frmoracle_impresion_documentos_fiscales.frx":0055
      Left            =   960
      List            =   "frmoracle_impresion_documentos_fiscales.frx":0071
      TabIndex        =   21
      Top             =   2400
      Width           =   3315
   End
   Begin VB.ComboBox cmb_metodos 
      Height          =   315
      ItemData        =   "frmoracle_impresion_documentos_fiscales.frx":0170
      Left            =   960
      List            =   "frmoracle_impresion_documentos_fiscales.frx":0186
      TabIndex        =   19
      Top             =   1680
      Width           =   3315
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exporta"
      Height          =   315
      Left            =   3480
      TabIndex        =   15
      Top             =   30
      Width           =   1080
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Bajar facturas"
      Height          =   315
      Left            =   4560
      TabIndex        =   4
      Top             =   30
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Archivo"
      Height          =   315
      Left            =   2490
      TabIndex        =   3
      Top             =   30
      Width           =   990
   End
   Begin VB.CommandButton cmd_cantia 
      Caption         =   "Cantia"
      Height          =   315
      Left            =   1425
      TabIndex        =   2
      Top             =   30
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
      Height          =   315
      Left            =   -1110
      TabIndex        =   1
      Top             =   30
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmd_concurrente 
      Appearance      =   0  'Flat
      Caption         =   "15/05/2016"
      Height          =   315
      Left            =   -1200
      Picture         =   "frmoracle_impresion_documentos_fiscales.frx":0213
      TabIndex        =   0
      ToolTipText     =   "Ejecuta concurrente"
      Top             =   30
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5910
      Picture         =   "frmoracle_impresion_documentos_fiscales.frx":0315
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   75
      TabIndex        =   14
      Top             =   450
      Width           =   6195
   End
   Begin VB.Frame frmseri 
      Height          =   1050
      Left            =   90
      TabIndex        =   9
      Top             =   495
      Width           =   6150
      Begin VB.CommandButton cmd_timbrar_traspaso 
         Caption         =   "Timbrar traspaso"
         Height          =   315
         Left            =   4080
         TabIndex        =   18
         Top             =   1080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txt_traspaso 
         Height          =   390
         Left            =   900
         TabIndex        =   16
         Top             =   1080
         Visible         =   0   'False
         Width           =   3120
      End
      Begin VB.TextBox txt_a 
         Height          =   390
         Left            =   4965
         TabIndex        =   8
         Top             =   555
         Width           =   1080
      End
      Begin VB.TextBox txt_de 
         Height          =   390
         Left            =   2895
         TabIndex        =   7
         Top             =   555
         Width           =   1080
      End
      Begin VB.TextBox txt_serie 
         Height          =   390
         Left            =   900
         TabIndex        =   6
         Top             =   555
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Traspaso:"
         Height          =   195
         Left            =   90
         TabIndex        =   17
         Top             =   1185
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label Label4 
         BackColor       =   &H000000FF&
         Caption         =   " Documento "
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   30
         TabIndex        =   13
         Top             =   120
         Width           =   6075
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Folio final:"
         Height          =   195
         Left            =   4095
         TabIndex        =   12
         Top             =   660
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Folio inicio:"
         Height          =   195
         Left            =   2100
         TabIndex        =   11
         Top             =   660
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   90
         TabIndex        =   10
         Top             =   660
         Width           =   405
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "CFDI Relacionado"
      Height          =   195
      Left            =   240
      TabIndex        =   25
      Top             =   2880
      Width           =   1305
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "Metodo:"
      Height          =   195
      Left            =   360
      TabIndex        =   24
      Top             =   2130
      Width           =   585
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "Uso:"
      Height          =   195
      Left            =   600
      TabIndex        =   23
      Top             =   2490
      Width           =   330
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "FP:"
      Height          =   195
      Left            =   720
      TabIndex        =   20
      Top             =   1770
      Width           =   240
   End
End
Attribute VB_Name = "frmoracle_impresion_documentos_fiscales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objConn As New ADODB.Connection
Dim objCmd As New ADODB.Command
Dim objParm As ADODB.Parameter
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const SW_SHOWNORMAL = 1
Dim var_ruta_facturas As String



Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
  "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal _
    szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Private Sub cmd_cantia_Click()
If IsNumeric(Me.txt_de) Then
      If IsNumeric(Me.txt_a) Then
         If CDbl(Me.txt_de) <= CDbl(Me.txt_a) Then
            If Me.txt_serie <> "" Then
               If Me.txt_serie = "FAEII" Or Me.txt_serie = "NCTEII" Then
                  If Me.cmb_uso <> "" Then
                     var_uso_cfdi_r = Mid(Me.cmb_uso, 1, 3)
                     If Me.cmb_metodos <> "" Then
                        var_forma_pago_r = Mid(Me.cmb_metodos, 1, 2)
                        If Me.cmb_metodo_2 <> "" Then
                           var_metodo_pago_r = Mid(Me.cmb_metodo_2, 1, 3)
                           For var_j = CDbl(Me.txt_de) To CDbl(Me.txt_a)
                               If rsaux1.State = 1 Then
                                  rsaux1.Close
                               End If
                               rsaux1.Open "select customer_trx_id, cadena as cadena, numero from xxvia_tb_control_doc_fiscales where serie = '" + Me.txt_serie + "' and numero = " + CStr(var_j), cnnoracle_4, adOpenDynamic, adLockOptimistic
                               If rsaux1.EOF Then
                                  On Error GoTo salir1:
                                  var_posible = 1
                                  CC = 1
                                  If CC = 1 Then
                                     var_cadena = "select CUSTOMER_TRX_ID from ra_customer_trx_all a, fnd_document_sequences b where b.name = ? AND TRX_NUMBER = ? and a.doc_sequence_id = b.doc_sequence_id(+) and rownum = 1"
                                     With comandoORA
                                          .ActiveConnection = cnnoracle_4
                                          .CommandType = adCmdText
                                          .CommandText = var_cadena
                                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Trim(Me.txt_serie))
                                          .Parameters.Append parametro
                                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, CStr(var_j))
                                          .Parameters.Append parametro
                                     End With
                                     Set rsaux9 = comandoORA.execute
                                     Set comandoORA = Nothing
                                     Set parametro = Nothing
                                     If Not rsaux9.EOF Then
                             
                                        var_cadena = "SELECT * FROM XXVIA_TB_SERIES_33 WHERE SERIE = ?"
                                        With comandoORA
                                             .ActiveConnection = cnnoracle_4
                                             .CommandType = adCmdText
                                             .CommandText = var_cadena
                                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_serie)
                                             .Parameters.Append parametro
                                        End With
                                        Set rsaux8 = comandoORA.execute
                                        Set comandoORA = Nothing
                                        Set parametro = Nothing
                                        If rsaux8.EOF Then
                                           CC = 2
                                        Else
                                           CC = 1
                                        End If
                                        If CC = 1 Then
                                           var_cadena = "CALL XXVIA_SP_CREA_EFLOW_33(?)"
                                           With comandoORA
                                                .ActiveConnection = cnnoracle_4
                                                .CommandType = adCmdText
                                                .CommandText = var_cadena
                                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, rsaux9!customer_Trx_id)
                                                .Parameters.Append parametro
                                           End With
                                           Set rsaux8 = comandoORA.execute
                                           Set comandoORA = Nothing
                                           Set parametro = Nothing
                                           rsaux1.Close
                                           rsaux1.Open "select customer_trx_id, cadena as cadena, numero from xxvia_tb_control_doc_fiscales where serie = '" + Me.txt_serie + "' and numero = " + CStr(var_j), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                           GoTo SIGUE
                               
                                        End If
                                        If CC = 2 Then
                                           var_cadena = "CALL XXVIA_SP_CREA_EFLOW(?)"
                                           With comandoORA
                                                .ActiveConnection = cnnoracle_4
                                                .CommandType = adCmdText
                                                .CommandText = var_cadena
                                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, rsaux9!customer_Trx_id)
                                                .Parameters.Append parametro
                                           End With
                                           Set rsaux8 = comandoORA.execute
                                           Set comandoORA = Nothing
                                           Set parametro = Nothing
                                           rsaux1.Close
                                           rsaux1.Open "select customer_trx_id, cadena  as cadena, numero from xxvia_tb_control_doc_fiscales where serie = '" + Me.txt_serie + "' and numero = " + CStr(var_j), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                           GoTo SIGUE
                                        End If
                                        'rsaux8.Open "CALL XXVIA_SP_CREA_EFLOW_33(" + CStr(rsaux9!CUSTOMER_tRX_ID) + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                     End If
                                     rsaux9.Close
                                  End If
                                  MsgBox "No existe aun la factura " + Me.txt_serie + CStr(var_j) + ", ejecute nuevamente el concurrente", vbOKOnly, "ATENCION"
                               Else
SIGUE:
                                  var_cadena = Replace(rsaux1!Cadena, "T23:", "T00:")
                                  var_cadena_rfc = Mid(var_cadena, 34, 12)
                                  VAR_CADENA_STR = ""
                                  Open ("C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".FAC") For Output As #1
                                  For var_i = 1 To Len(var_cadena)
                                      'MsgBox Asc(Mid(var_cadena, var_i, 1))
                                      If Asc(Mid(var_cadena, var_i, 1)) = 63 Then
                                         If Mid(VAR_CADENA_STR, 1, 15) = "COMPROBANTE:3.3" Then
                                            var_f = 1
                                            var_forma_pago = ""
                                            var_contador_pip = 0
                                            While var_contador_pip < 4
                                                  'var_f = 1
                                                  If Mid(VAR_CADENA_STR, var_f, 1) = "|" Then
                                                     var_contador_pip = var_contador_pip + 1
                                                  End If
                                                  var_f = var_f + 1
                                            Wend
                                            If var_contador_pip = 4 Then
                                               var_forma_pago = Mid(VAR_CADENA_STR, var_f, 2)
                                               var_forma_pago = "|" + var_forma_pago
                                               var_forma_pago_r = "|" + var_forma_pago_r
                                            End If
                                         End If
                                         VAR_CADENA_STR = Replace(VAR_CADENA_STR, var_forma_pago, var_forma_pago_r)
                                         If Mid(VAR_CADENA_STR, 1, 15) = "COMPROBANTE:3.3" Then
                                            var_m = 1
                                            VAR_METODO_PAGO = ""
                                            var_contador_met = 0
                                            While var_contador_met < 12
                                                  'var_f = 1
                                                  If Mid(VAR_CADENA_STR, var_m, 1) = "|" Then
                                                     var_contador_met = var_contador_met + 1
                                                  End If
                                                  var_m = var_m + 1
                                            Wend
                                            If var_contador_met = 12 Then
                                               VAR_METODO_PAGO = Mid(VAR_CADENA_STR, var_m, 3)
                                               VAR_METODO_PAGO = "|" + VAR_METODO_PAGO
                                               var_metodo_pago_r = "|" + var_metodo_pago_r
                                            End If
                                         End If
                                         VAR_CADENA_STR = Replace(VAR_CADENA_STR, VAR_METODO_PAGO, var_metodo_pago_r)
                                         If Mid(VAR_CADENA_STR, 1, 9) = "RECEPTOR:" Then
                                            var_u = 1
                                            var_uso_cfdi = ""
                                            var_contador_uso = 0
                                            While var_contador_uso < 4
                                                  'var_f = 1
                                                  If Mid(VAR_CADENA_STR, var_u, 1) = "|" Then
                                                     var_contador_uso = var_contador_uso + 1
                                                  End If
                                                  var_u = var_u + 1
                                            Wend
                                            If var_contador_uso = 4 Then
                                               var_uso_cfdi = Mid(VAR_CADENA_STR, var_u, 3)
                                               var_uso_cfdi = "|" + var_uso_cfdi
                                               var_uso_cfdi_r = "|" + var_uso_cfdi_r
                                            End If
                                         End If
                                         VAR_CADENA_STR = Replace(VAR_CADENA_STR, var_uso_cfdi, var_uso_cfdi_r)
                                         If Me.txt_serie = "NCTEII" Then
                                            If Mid(VAR_CADENA_STR, 1, 21) = "CFDIS_RELACIONADO:01|" Then
                                               VAR_CADENA_STR = Replace(VAR_CADENA_STR + Me.txt_cfdi, " ", "")
                                            End If
                                         End If
                                         
                                         
                                         If Trim(VAR_CADENA_STR) = "CONDICIONES_PAGO:- -" And Mid(Me.txt_serie, 1, 2) = "NC" Then
                                            VAR_CADENA_STR = "CONDICIONES_PAGO:NO IDENTIFICADO"
                                         End If
                                         'If Mid(var_cadena_str, 1, 6) = "FECHA:" Or Mid(var_cadena_str, 1, 16) = "FECHA:2014-08-08" Or Mid(var_cadena_str, 1, 16) = "FECHA:2014-08-09" Or Mid(var_cadena_str, 1, 16) = "FECHA:2014-08-10" Or Mid(var_cadena_str, 1, 16) = "FECHA:2014-08-11" Or Mid(var_cadena_str, 1, 16) = "FECHA:2014-08-11" Then
                                         If Mid(VAR_CADENA_STR, 1, 6) = "LINEA:" Then
                                            VAR_Z = 0
                                            var_codigo = ""
                                            VAR_39 = 38
                                            While VAR_Z = 0
                                                  If Mid(VAR_CADENA_STR, VAR_39 + 1, 1) = " " Then
                                                     VAR_Z = 1
                                                  Else
                                                     var_codigo = var_codigo + Mid(VAR_CADENA_STR, VAR_39 + 1, 1)
                                                  End If
                                                  VAR_39 = VAR_39 + 1
                                            Wend
                                            'var_codigo
                                            If cnn_compucaja_f.State = 1 Then
                                               cnn_compucaja_f.Close
                                            End If
                                            cnn_compucaja_f.Open "Provider=SQLOLEDB.1;Password=compucaja;Persist Security Info=True;User ID=sa;Initial Catalog=TdaAgsCC2013;Data Source=sqlcantia2.vianney.com.mx"
                                            rsaux11.Open "SELECT B.Art_Codigo, B.SAT_ProdServClave, C.SAT_UnidadMedidaClave UOM FROM UnidadesMedidaArticulo A, ArticulosSAT B, UnidadesMedida C WHERE A.Art_Codigo = B.Art_Codigo AND C.UM_Codigo = A.UM_Codigo AND A.Art_Codigo = '" + Replace(IIf(IsNull(var_codigo), "", var_codigo), "|", "") + "'", cnn_compucaja_f, adOpenDynamic, adLockOptimistic
                                            If Not rsaux11.EOF Then
                                               var_codigo_Sat = IIf(IsNull(rsaux11!SAT_ProdServClave), "01010101", rsaux11!SAT_ProdServClave)
                                               VAR_UOM_SAT = IIf(IsNull(rsaux11!UOM), "H87", rsaux11!UOM)
                                               If var_codigo_Sat = "01010101" Then
                                                  
                                                  
                                                  var_cadena_o = "select clasificacionsat, uom_sat from MTL_CROSS_REFERENCES_B a, XXVIA_SYSTEM_ITEMS_B b where cross_reference = ? and a.inventory_item_id = b.inventory_item_id and b.organization_id = 90"
                                                  With comandoORA
                                                       .ActiveConnection = cnnoracle_4
                                                       .CommandType = adCmdText
                                                       .CommandText = var_cadena_o
                                                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(var_codigo), "", var_codigo))
                                                       .Parameters.Append parametro
                                                  End With
                                                  Set rsaux8 = comandoORA.execute
                                                  Set comandoORA = Nothing
                                                  Set parametro = Nothing
                                                  If Not rsaux8.EOF Then
                                                     var_codigo_Sat = IIf(IsNull(rsaux8!CLASIFICACIONSAT), "01010101", rsaux8!CLASIFICACIONSAT)
                                                     VAR_UOM_SAT = IIf(IsNull(rsaux8!UOM_SAT), "H87", rsaux8!UOM_SAT)
                                                  Else
                                                     var_codigo_Sat = "01010101"
                                                     VAR_UOM_SAT = "H87"
                                                  End If
                                                  rsaux8.Close
                                               End If
                                            Else
                                               var_cadena_o = "select clasificacionsat, uom_sat from MTL_CROSS_REFERENCES_B a, XXVIA_SYSTEM_ITEMS_B b where cross_reference = ? and a.inventory_item_id = b.inventory_item_id and b.organization_id = 90"
                                               With comandoORA
                                                    .ActiveConnection = cnnoracle_4
                                                    .CommandType = adCmdText
                                                    .CommandText = var_cadena_o
                                                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(var_codigo), "", var_codigo))
                                                    .Parameters.Append parametro
                                               End With
                                               Set rsaux8 = comandoORA.execute
                                               Set comandoORA = Nothing
                                               Set parametro = Nothing
                                               If Not rsaux8.EOF Then
                                                  var_codigo_Sat = IIf(IsNull(rsaux8!CLASIFICACIONSAT), "01010101", rsaux8!CLASIFICACIONSAT)
                                                  VAR_UOM_SAT = IIf(IsNull(rsaux8!UOM_SAT), "H87", rsaux8!UOM_SAT)
                                               Else
                                                  var_codigo_Sat = "01010101"
                                                  VAR_UOM_SAT = "H87"
                                               End If
                                               rsaux8.Close
                                            End If
                                            rsaux11.Close
                                  
                                  
                                  
                                            VAR_CADENA_STR = Replace(VAR_CADENA_STR, "84111506", var_codigo_Sat)
                                            VAR_CADENA_STR = Replace(VAR_CADENA_STR, "H87", VAR_UOM_SAT)
                                         End If
                                         Print #1, VAR_CADENA_STR
                                         VAR_CADENA_STR = ""
                                      Else
                                         VAR_CADENA_STR = VAR_CADENA_STR + Mid(var_cadena, var_i, 1)
                                      End If
                                  Next var_i
                                  Print #1, "FIN:"
                                  Close #1
                                  var_archivo = "C:\SISTEMAS\sube_fact_" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".bat"
                                  'Open ("C:\SISTEMAS\sube_fact_" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".bat") For Output As #2
                                  'Print #2, "c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS|" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".FAC|https://facturas2.vianney.mx/cgi-bin/cfds/timbrarGR33|cfdsvianney|9y3jv^TI;4g#|1"
                                  ''Print #2, "facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\" + Trim(var_serie) + Trim(Str(rsaux3!inte_Car_numero)) + ".FAC" + "|https://facturas2.vianney.mx/cgi-bin/cfds/timbrarGR|cfdsvianney|9y3jv^TI;4g#|1" + """"
                                  'Close #2
                                  'var_Archivo = "C:\SISTEMAS\sube_fact_" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".bat"
                                  'x = Shell(var_Archivo, vbHide)
                                  'x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".FAC" + "|https://facturas2.vianney.mx/cgi-bin/cfds/timbrarGR33|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
                                  'x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".FAC" + "|https://facturas2.vianney.mx/cgi-bin/cfds/timbrarGR33|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
                                  strconsulta = "SELECT SERIE FROM XXVIA_tB_cONTROL_DOC_FISCALES WHERE SERIE = ? and numero = ? AND NUMERO_TIENDA = '3.3'"
                                  With comandoORA
                                       .ActiveConnection = cnnoracle_4
                                       .CommandType = adCmdText
                                       .CommandText = strconsulta
                                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Trim(Me.txt_serie))
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
                                  If Me.txt_serie <> "FAEVII" Then
                                     If VAR_33 = 1 Then
                                        x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".FAC" + "|https://facturas2.vianney.mx/cgi-bin/cfds/timbrarGR33|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
                                     Else
                                        x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".FAC" + "|https://facturas.vianney.mx/cgi-bin/cfds/timbrarGR|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
                                     End If
                                  Else
                                     MsgBox "Selecciona la opción de exporta", vbOKOnly, "ATENCION"
                                  End If
                                  'URL = "https://facturas.vianney.mx/cgi-bin/cfds/cfdsORACLE?cmd=download_pdf&rfc_emisor=" + var_cadena_rfc + "&serie=" + Trim(Me.txt_serie) + "&folio=" + Trim(CStr(var_j))
                                  'buf = Split(URL, ".")
                                  'ext = buf(UBound(buf))
                                  'strSavePath = "C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(CStr(var_j)) + ".pdf"
                                  'ret = URLDownloadToFile(0, URL, strSavePath, 0, 0)
                               End If
                               rsaux1.Close
                           Next var_j
                           MsgBox "Se a terminado el proceso", vbOKOnly, "ATENCION"
                        Else
                           MsgBox "Selecciona la opción de exporta", vbOKOnly, "ATENCION"
                        End If
                     Else
                     End If
                  Else
                  End If
               Else
               End If
            Else
               MsgBox "Serie invalida", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El número final de factura debe de ser mayor al inicial"
         End If
      Else
         MsgBox "Número de factura final incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Número de factura inical incorrecto", vbOKOnly, "ATENCION"
   End If
   Exit Sub
salir1:
      If Err.Number = -2147217900 Then
         'MsgBox Err.Description
         rsaux14.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         rsaux14.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         Resume
      End If

End Sub

Private Sub cmd_carta_porte_Click()
If IsNumeric(Me.txt_de) Then
      If IsNumeric(Me.txt_a) Then
         If CDbl(Me.txt_de) <= CDbl(Me.txt_a) Then
            If Me.txt_serie <> "" Then
               If rsaux1.State = 1 Then
                  rsaux1.Close
               End If
               var_posible_embarque = 1
               var_Cadena_pedidos = Me.txt_de
               var_j = 0
               rsaux.Open "alter session set nls_languAge = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               var_cadena = "SELECT  oh.ordered_date, oh.source_document_id, oh.header_id, oh.order_number, oh.transactional_curr_code, NVL(ol.ordered_quantity,0) AS CANTIDAD_PEDIDA, NVL(ol.cancelled_quantity,0) AS CANTIDAD_NEGADA, NVL(ol.shipped_quantity,0)   AS CANTIDAD_surtida, ol.line_id, ol.ordered_item, ol.order_quantity_uom, ol.inventory_item_id, ol.price_list_id, ol.unit_selling_price, DECODE(ol.cancelled_flag,'Y','CANCELADA','SURTIDA') line_status, ol.flow_status_code"
               var_cadena = var_cadena + " FROM oe_order_headers_all oh, oe_order_lines_all ol, OE_ORDER_LINES_HISTORY OLH WHERE order_number  = " + var_Cadena_pedidos
               var_cadena = var_cadena + " AND oh.header_id = ol.header_id AND ol.ship_from_org_id = 93 AND oL.header_id = oLh.header_id(+) AND OL.LINE_ID = OLH.LINE_ID(+) and  NVL(ol.shipped_quantity,0) > 0"
               If rsaux.State = 1 Then
                  rsaux.Close
               End If
               rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               var_posible_embarque = 0
               If Not rsaux.EOF Then
                  var_posible_embarque = 1
               End If
               rsaux.Close
               rsaux.Open "select * from xxvia_tb_encabezado_embarques where embarque = " + CStr(Me.txt_serie), cnnoracle_4, adOpenDynamic, adLockOptimistic
               var_chofer = ""
               If Not rsaux.EOF Then
                  var_chofer = IIf(IsNull(rsaux!CHOFER), "", rsaux!CHOFER)
               Else
                  var_chofer = ""
               End If
               If var_chofer = "" Then
                  var_posible_embarque = 2
               End If
               rsaux.Close
               rsaux.Open "select * from xxvia_tb_encabezado_embarques where embarque = " + CStr(Me.txt_serie), cnnoracle_4, adOpenDynamic, adLockOptimistic
               var_transporte = ""
               If Not rsaux.EOF Then
                  var_transporte = IIf(IsNull(rsaux!transporte), "", rsaux!transporte)
               Else
                  var_transporte = ""
               End If
               If var_transporte = "" Then
                  var_posible_embarque = 3
               End If
               rsaux.Close
               If var_posible_embarque = 1 Then
                  If rsaux1.State = 1 Then
                     rsaux1.Close
                  End If
                  var_cadena = "select customer_trx_id, cadena  as cadena, numero from xxvia_tb_control_doc_fiscales where serie = 'TRX" + Me.txt_serie + "' and numero = " + CStr(Me.txt_de)
                  rsaux1.Open "select customer_trx_id, cadena  as cadena, numero from xxvia_tb_control_doc_fiscales where serie = 'TRX" + Me.txt_serie + "_' and numero = " + CStr(Me.txt_de), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  var_cadena = Replace(Replace(rsaux1!Cadena, "T23:", "T00:"), "AUTORIZADO  ", "AUTORIZADO ")
                  var_cadena_rfc = Mid(var_cadena, 34, 12)
                  VAR_CADENA_STR = ""
                  Open ("C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".FAC") For Output As #1
                  For var_i = 1 To Len(var_cadena)
                      If Asc(Mid(var_cadena, var_i, 1)) = 63 Then
                         Print #1, VAR_CADENA_STR
                         VAR_CADENA_STR = ""
                      Else
                         VAR_CADENA_STR = VAR_CADENA_STR + Mid(var_cadena, var_i, 1)
                      End If
                  Next var_i
                  Print #1, "FIN:"
                  Close #1
                        
                  var_archivo = "C:\SISTEMAS\sube_fact_" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".bat"
                  'x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".FAC" + "|https://facturas2.vianney.mx/cgi-bin/cfds/timbrarGR33|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
                  x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\" + Trim(Me.txt_serie) + ".FAC" + "|https://facturas2.vianney.mx/cgi-bin/cfds/timbrarGR33|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
               Else
                  If var_posible_embarque = 0 Then
                     MsgBox "El pedido no ha sido cerrado", vbOKOnly, "ATENCION"
                  End If
                  If var_posible_embarque = 2 Then
                     MsgBox "No se ha asignado un chofer al embarque", vbOKOnly, "ATENCION"
                  End If
                  If var_posible_embarque = 3 Then
                     MsgBox "El embarque no tiene un transporte asignado.", vbOKOnly, "ATENCION"
                  End If
               End If
            Else
               MsgBox "No se indico el embarque.", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El número final de factura debe de ser mayor al inicial"
         End If
      Else
         MsgBox "Número de factura final incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Número de factura inical incorrecto", vbOKOnly, "ATENCION"
   End If
   Exit Sub
salir1:
      If Err.Number = -2147217900 Then
         'MsgBox Err.Description
         rsaux14.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         rsaux14.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         Resume
      End If

End Sub

Private Sub cmd_concurrente_Click()
If IsNumeric(Me.txt_de) Then
      If IsNumeric(Me.txt_a) Then
         If CDbl(Me.txt_de) <= CDbl(Me.txt_a) Then
            If Me.txt_serie <> "" Then
               For var_j = CDbl(Me.txt_de) To CDbl(Me.txt_a)
                   rsaux1.Open "select customer_trx_id, cadena as cadena, numero from xxvia_tb_control_doc_fiscales where serie = '" + Me.txt_serie + "' and numero = " + CStr(var_j), cnnoracle_4, adOpenDynamic, adLockOptimistic
                   If rsaux1.EOF Then
                      var_posible = 1
                      MsgBox "No existe aun la factura " + Me.txt_serie + CStr(var_j) + ", ejecute nuevamente el concurrente", vbOKOnly, "ATENCION"
                   Else
                     
                      var_cadena = rsaux1!Cadena
                      var_cadena_rfc = Mid(var_cadena, 34, 12)
                      VAR_CADENA_STR = ""
                      
                      
                      Open ("C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".FAC") For Output As #1
                      For var_i = 1 To Len(var_cadena)
                          'MsgBox Asc(Mid(var_cadena, var_i, 1))
                          If Asc(Mid(var_cadena, var_i, 1)) = 63 Then
                             If Trim(VAR_CADENA_STR) = "CONDICIONES_PAGO:- -" And Mid(Me.txt_serie, 1, 2) = "NC" Then
                                VAR_CADENA_STR = "CONDICIONES_PAGO:NO IDENTIFICADO"
                             End If
                             'If Mid(var_cadena_str, 1, 6) = "FECHA:" Or Mid(var_cadena_str, 1, 16) = "FECHA:2014-08-08" Or Mid(var_cadena_str, 1, 16) = "FECHA:2014-08-09" Or Mid(var_cadena_str, 1, 16) = "FECHA:2014-08-10" Or Mid(var_cadena_str, 1, 16) = "FECHA:2014-08-11" Or Mid(var_cadena_str, 1, 16) = "FECHA:2014-08-11" Then
                             If Mid(VAR_CADENA_STR, 1, 6) = "FECHA:" Then
                                VAR_CADENA_STR = "FECHA:2017-12-31T23:59:00 "
                                'var_cadena_str = "FECHA:2015-09-27T00:01:00 "
                             End If
                             Print #1, VAR_CADENA_STR
                             VAR_CADENA_STR = ""
                          Else
                             VAR_CADENA_STR = VAR_CADENA_STR + Mid(var_cadena, var_i, 1)
                          End If
                      Next var_i
                      Print #1, "FIN:"
                      Close #1
                      var_archivo = "C:\SISTEMAS\sube_fact" + Trim(Str(var_j)) + ".bat"
                      x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".FAC" + "|https://facturas.vianney.mx/cgi-bin/cfds/timbrarGR|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
                                       
                      'URL = "https://facturas.vianney.mx/cgi-bin/cfds/cfdsORACLE?cmd=download_pdf&rfc_emisor=" + var_cadena_rfc + "&serie=" + Trim(Me.txt_serie) + "&folio=" + Trim(CStr(var_j))
                      'buf = Split(URL, ".")
                      'ext = buf(UBound(buf))
                      'strSavePath = "C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(CStr(var_j)) + ".pdf"
                      'ret = URLDownloadToFile(0, URL, strSavePath, 0, 0)
                   End If
                   rsaux1.Close
               Next var_j
               MsgBox "Se a terminado el proceso", vbOKOnly, "ATENCION"
            Else
               MsgBox "Serie invalida", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El número final de factura debe de ser mayor al inicial"
         End If
      Else
         MsgBox "Número de factura final incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Número de factura inical incorrecto", vbOKOnly, "ATENCION"
   End If

End Sub


Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmd_timbrar_traspaso_Click()
   If Me.txt_traspaso <> "" Then
      rsaux1.Open "select customer_trx_id, cadena as cadena, numero from xxvia_tb_control_doc_fiscales where serie = '" + Me.txt_traspaso + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If rsaux1.EOF Then
         var_posible = 1
         MsgBox "No existe el traspaso " + Me.txt_traspaso, vbOKOnly, "ATENCION"
      Else
         var_cadena = rsaux1!Cadena
         var_cadena_rfc = Mid(var_cadena, 34, 12)
         VAR_CADENA_STR = ""
         Open ("C:\SISTEMAS\" + Trim(Me.txt_traspaso) + ".FAC") For Output As #1
         For var_i = 1 To Len(var_cadena)
             'MsgBox Asc(Mid(var_cadena, var_i, 1))
             If Asc(Mid(var_cadena, var_i, 1)) = 63 Then
                If Trim(VAR_CADENA_STR) = "CONDICIONES_PAGO:- -" And Mid(Me.txt_serie, 1, 2) = "NC" Then
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
         var_archivo = "C:\SISTEMAS\sube_fact" + Trim(Str(var_j)) + ".bat"
         x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".FAC" + "|https://facturas2.vianney.mx/cgi-bin/cfds/timbrarGR33|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
     End If
     rsaux1.Close
   End If
End Sub

Private Sub cmd_version_40_Click()

End Sub

Private Sub Command1_Click()
If IsNumeric(Me.txt_de) Then
      If IsNumeric(Me.txt_a) Then
         If CDbl(Me.txt_de) <= CDbl(Me.txt_a) Then
            If Me.txt_serie <> "" Then
               For var_j = CDbl(Me.txt_de) To CDbl(Me.txt_a)
                   rsaux1.Open "select customer_trx_id from xxvia_Tb_control_doc_fiscales where serie = '" + Me.txt_serie + "' and numero = " + CStr(var_j), cnnoracle_4, adOpenDynamic, adLockOptimistic
                   If rsaux1.EOF Then
                      var_posible = 1
                   Else
                      'rsaux2.Open "CALL XXVIA_SEND_POST('" + CStr(rsaux1!CUSTOMER_TRX_ID) + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                   End If
                   rsaux1.Close
               Next var_j
               
               For var_j = CDbl(Me.txt_de) To CDbl(Me.txt_a)
                   URL = "https://facturas.vianney.mx/cgi-bin/cfds/cfdsORACLE?cmd=download_pdf&rfc_emisor=VTH981105F90&serie=" + Trim(Me.txt_serie) + "&folio=" + Trim(CStr(var_j))
                   buf = Split(URL, ".")
                   ext = buf(UBound(buf))
                   strSavePath = "C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(CStr(var_j)) + ".pdf"
                   ret = URLDownloadToFile(0, URL, strSavePath, 0, 0)
                   If ret = 0 Then
                      Call ShellExecute(Me.hwnd, "print", "C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(CStr(var_j)) + ".PDF", vbNullString, vbNullString, SW_SHOWNORMAL = 1)
                   Else
                      MsgBox "Error en la factura " + Me.txt_serie + Trim(CStr(var_j))
                   End If
               Next var_j
            Else
               MsgBox "Se debe de indicar una serie", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El número final de factura debe de ser mayor al inicial"
         End If
      Else
         MsgBox "Número de factura final incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Número de factura inical incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Command2_Click()

   Open "C:\sistemas\" + Me.txt_serie + CStr(var_j) + ".txt" For Input As #1
   Dim Linea As String, Total As String
   Do Until EOF(1)
   Line Input #1, Linea
   If Mid(Linea, 1, 2) = "OK" Then
      MsgBox "Factura correcta", vbCritical, ""
   Else
      var_linea = ""
      For var_j = 1 To Len(Linea)
          If Mid(Linea, var_j, 16) = "Valor Esperado: " Then
             var_i = 0
             var_linea = ""
             VAR_Z = 1
             While var_i = 0
                   If Mid(Linea, var_j + 15 + VAR_Z, 1) <> " " Then
                      var_linea = var_linea + Mid(Linea, var_j + 15 + VAR_Z, 1)
                      VAR_Z = VAR_Z + 1
                   Else
                      var_i = 1
                   End If
                   
             Wend
             MsgBox "cantidad esperada " + var_linea, vbOKOnly, ""
          End If
      Next var_j
   
      For var_j = 1 To Len(Linea)
          If Mid(Linea, var_j, 17) = "Valor Reportado: " Then
             var_i = 0
             var_linea_2 = ""
             VAR_Z = 1
             While var_i = 0
                   If Mid(Linea, var_j + 15 + VAR_Z, 1) <> "<" Then
                      var_linea_2 = var_linea_2 + Mid(Linea, var_j + 15 + VAR_Z, 1)
                      VAR_Z = VAR_Z + 1
                   Else
                      var_i = 1
                   End If
                   
              Wend
             MsgBox "cantidad esperada " + var_linea + " cantidad reportada " + var_linea_2, vbOKOnly, ""
          End If
      Next var_j
      MsgBox "Factura incorrecta", vbCritical, ""
   End If
   Loop
   Close #1
   If rsaux1.State = 1 Then
      rsaux1.Close
   End If
   rsaux1.Open "select customer_trx_id, cadena as cadena, numero from xxvia_tb_control_doc_fiscales where serie = 'FAEVXX' and numero = 48388", cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_cadena = Replace(rsaux1!Cadena, "T23:", "T00:")
   var_cadena_rfc = Mid(var_cadena, 34, 12)
   VAR_CADENA_STR = ""
   Open ("C:\SISTEMAS\FAEVXX48388.FAC") For Output As #1
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
End Sub


Private Sub Command3_Click()
Dim var_version As String
If IsNumeric(Me.txt_de) Then
      If IsNumeric(Me.txt_a) Then
         If CDbl(Me.txt_de) <= CDbl(Me.txt_a) Then
            If Me.txt_serie <> "" Then
               If Me.txt_serie <> "FAEVII" Then
                  For var_j = CDbl(Me.txt_de) To CDbl(Me.txt_a)
                      If rsaux1.State = 1 Then
                         rsaux1.Close
                      End If
                      rsaux1.Open "select customer_trx_id, cadena as cadena, numero, NVL(numero_tienda,'3.3') AS numero_tienda from xxvia_tb_control_doc_fiscales where serie = '" + Me.txt_serie + "' and numero = " + CStr(var_j), cnnoracle_4, adOpenDynamic, adLockOptimistic
                      
                      'rsaux1.Open "select customer_trx_id, translate(substr(cadena,1,4000),chr(9),' ') cadena, numero from xxvia_tb_control_doc_fiscales where serie = '" + Me.txt_serie + "' and numero = " + CStr(var_j), cnnoracle_4, adOpenDynamic, adLockOptimistic
                      'Me.txt_serie = "FAI"
                      If rsaux1.EOF Then
                         
                         
                         On Error GoTo salir1:
                         var_posible = 1
                         CC = 1
                         If CC = 1 Then
                         var_cadena = "select CUSTOMER_TRX_ID from ra_customer_trx_all a, fnd_document_sequences b where b.name = ? AND TRX_NUMBER = ? and a.doc_sequence_id = b.doc_sequence_id(+) and rownum = 1"
                         With comandoORA
                              .ActiveConnection = cnnoracle_4
                              .CommandType = adCmdText
                              .CommandText = var_cadena
                              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Trim(Me.txt_serie))
                              .Parameters.Append parametro
                              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, CStr(var_j))
                              .Parameters.Append parametro
                         End With
                         Set rsaux9 = comandoORA.execute
                         Set comandoORA = Nothing
                         Set parametro = Nothing
                         If Not rsaux9.EOF Then
                             
                            var_cadena = "SELECT * FROM XXVIA_TB_SERIES_33 WHERE SERIE = ?"
                            With comandoORA
                                 .ActiveConnection = cnnoracle_4
                                 .CommandType = adCmdText
                                 .CommandText = var_cadena
                                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_serie)
                                 .Parameters.Append parametro
                            End With
                            Set rsaux8 = comandoORA.execute
                            Set comandoORA = Nothing
                            Set parametro = Nothing
                            If rsaux8.EOF Then
                               CC = 2
                            Else
                               CC = 1
                            End If
                            If CC = 1 Then
                               var_cadena = "CALL XXVIA_SP_CREA_EFLOW_33(?)"
                               With comandoORA
                                    .ActiveConnection = cnnoracle_4
                                    .CommandType = adCmdText
                                    .CommandText = var_cadena
                                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, rsaux9!customer_Trx_id)
                                    .Parameters.Append parametro
                               End With
                               Set rsaux8 = comandoORA.execute
                               Set comandoORA = Nothing
                               Set parametro = Nothing
                               rsaux1.Close
                               rsaux1.Open "select customer_trx_id, cadena as cadena, numero from xxvia_tb_control_doc_fiscales where serie = '" + Me.txt_serie + "' and numero = " + CStr(var_j), cnnoracle_4, adOpenDynamic, adLockOptimistic
                               GoTo SIGUE
                               
                            End If
                            If CC = 2 Then
                               var_cadena = "CALL XXVIA_SP_CREA_EFLOW(?)"
                               With comandoORA
                                    .ActiveConnection = cnnoracle_4
                                    .CommandType = adCmdText
                                    .CommandText = var_cadena
                                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, rsaux9!customer_Trx_id)
                                    .Parameters.Append parametro
                               End With
                               Set rsaux8 = comandoORA.execute
                               Set comandoORA = Nothing
                               Set parametro = Nothing
                               rsaux1.Close
                               rsaux1.Open "select customer_trx_id, cadena  as cadena, numero from xxvia_tb_control_doc_fiscales where serie = '" + Me.txt_serie + "' and numero = " + CStr(var_j), cnnoracle_4, adOpenDynamic, adLockOptimistic
                               GoTo SIGUE
                               
                            End If
                            
                            'rsaux8.Open "CALL XXVIA_SP_CREA_EFLOW_33(" + CStr(rsaux9!CUSTOMER_tRX_ID) + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                         End If
                         rsaux9.Close
                         End If
                         
                         MsgBox "No existe aun la factura " + Me.txt_serie + CStr(var_j) + ", ejecute nuevamente el concurrente", vbOKOnly, "ATENCION"
                         
                      Else
SIGUE:
                         var_version = rsaux1!numero_tienda
                         var_cadena = Replace(Replace(rsaux1!Cadena, "T23:", "T00:"), "AUTORIZADO  ", "AUTORIZADO ")
                         var_cadena_rfc = Mid(var_cadena, 34, 12)
                         VAR_CADENA_STR = ""
                         Open ("C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".FAC") For Output As #1
                         For var_i = 1 To Len(var_cadena)
                             'MsgBox Asc(Mid(var_cadena, var_i, 1))
                             If Asc(Mid(var_cadena, var_i, 1)) = 63 Then
                                If Trim(VAR_CADENA_STR) = "CONDICIONES_PAGO:- -" And Mid(Me.txt_serie, 1, 2) = "NC" Then
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
                         var_archivo = "C:\SISTEMAS\sube_fact_" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".bat"
                         
                         
                         'Open ("C:\SISTEMAS\sube_fact_" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".bat") For Output As #2
                         'Print #2, "c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS|" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".FAC|https://facturas2.vianney.mx/cgi-bin/cfds/timbrarGR33|cfdsvianney|9y3jv^TI;4g#|1"
                         ''Print #2, "facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\" + Trim(var_serie) + Trim(Str(rsaux3!inte_Car_numero)) + ".FAC" + "|https://facturas2.vianney.mx/cgi-bin/cfds/timbrarGR|cfdsvianney|9y3jv^TI;4g#|1" + """"
                         'Close #2
                         'var_Archivo = "C:\SISTEMAS\sube_fact_" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".bat"
                         'x = Shell(var_Archivo, vbHide)
                         'x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".FAC" + "|https://facturas2.vianney.mx/cgi-bin/cfds/timbrarGR33|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
                         'x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".FAC" + "|https://facturas2.vianney.mx/cgi-bin/cfds/timbrarGR33|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
                         
                         
                         'quitar esto por la version 4.0
                         'strconsulta = "SELECT SERIE FROM XXVIA_tB_cONTROL_DOC_FISCALES WHERE SERIE = ? and numero = ? AND NUMERO_TIENDA = '3.3'"
                         'With comandoORA
                         '     .ActiveConnection = cnnoracle_4
                         '     .CommandType = adCmdText
                         '     .CommandText = strconsulta
                         '     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Trim(Me.txt_serie))
                         '     .Parameters.Append parametro
                         '     Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, var_j)
                         '     .Parameters.Append parametro
                         'End With
                         'Set rsaux9 = comandoORA.execute
                         'Set comandoORA = Nothing
                         'Set parametro = Nothing
                        'If Not rsaux9.EOF Then
                           VAR_33 = 1
                        'Else
                        '   VAR_33 = 0
                        'End If
                        'rsaux9.Close
                        If Me.txt_serie <> "FAEVII" Then
                         If VAR_33 = 1 Then
                            If var_version = "3.3" Then
                               x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".FAC" + "|https://facturas2.vianney.mx/cgi-bin/cfds/timbrarGR33|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
                            Else
                               If var_version = "4.0" Then
                                  x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".FAC" + "|https://facturas2.vianney.mx/cgi-bin/cfds/timbrarGR40|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
                               End If
                            End If
                         Else
                            x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".FAC" + "|https://facturas.vianney.mx/cgi-bin/cfds/timbrarGR|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
                         End If
                         Else
                            MsgBox "Selecciona la opción de exporta", vbOKOnly, "ATENCION"
                         End If
                         'Me.txt_serie = "FAEVII"
                         'URL = "https://facturas.vianney.mx/cgi-bin/cfds/cfdsORACLE?cmd=download_pdf&rfc_emisor=" + var_cadena_rfc + "&serie=" + Trim(Me.txt_serie) + "&folio=" + Trim(CStr(var_j))
                         'buf = Split(URL, ".")
                         'ext = buf(UBound(buf))
                         'strSavePath = "C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(CStr(var_j)) + ".pdf"
                         'ret = URLDownloadToFile(0, URL, strSavePath, 0, 0)
                      End If
                      rsaux1.Close
                      Sleep 5000
                      var_zzz = 0
                      On Error GoTo archivo
                      Open "C:\sistemas\" + Me.txt_serie + CStr(var_j) + ".txt" For Input As #1
                      'Open "C:\sistemas\FAEVXX_prueba" + CStr(var_j) + ".txt" For Input As #1
                      
                      Dim Linea As String, Total As String
                      Do Until EOF(1)
                         Line Input #1, Linea
                         VAR_IOU = 0
                         
                         'If Mid(Linea, 1, 2) <> "OK" Then
                         If VAR_IOU = 0 Then
                            var_zzz = 1
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
                            
                            
                            
                            For var_jF = 1 To Len(Linea)
                                If Mid(Linea, var_jF, 17) = "Valor Reportado: " Or Mid(Linea, var_jF, 16) = "ValorReportado: " Then
                                   var_i = 0
                                   var_linea_2 = ""
                                   VAR_Z = 1
                                   If Mid(Linea, var_jF, 17) = "Valor Reportado: " Then
                                   While var_i = 0
                                         If Mid(Linea, var_jF + 15 + VAR_Z, 1) <> "<" Then
                                            var_linea_2 = var_linea_2 + Mid(Linea, var_jF + 15 + VAR_Z, 1)
                                            VAR_Z = VAR_Z + 1
                                         Else
                                            var_i = 1
                                         End If
                   
                                   Wend
                                   Else
                                   
                                   While var_i = 0
                                         If Mid(Linea, var_jF + 15 + VAR_Z, 1) <> " " Then
                                            var_linea_2 = var_linea_2 + Mid(Linea, var_jF + 15 + VAR_Z, 1)
                                            VAR_Z = VAR_Z + 1
                                         Else
                                            var_i = 1
                                         End If
                   
                                   Wend
                                   
                                   
                                   End If
                                  'MsgBox "cantidad esperada " + var_linea + " cantidad reportada " + VAR_LINEA_2, vbOKOnly, ""
                               End If
                            Next var_jF
                            
                            
                            
                            'MsgBox "Factura incorrecta", vbCritical, ""
                         End If
                      Loop
                      Close #1
                      If var_zzz = 1 Then
                      If rsaux1.State = 1 Then
                         rsaux1.Close
                      End If
                      var_cadena = "update xxvia_Tb_control_doc_fiscales set cadena = replace(cadena,'|" + Trim(var_linea_2) + "|','|" + Trim(var_linea) + "|') where serie = '" + Trim(Me.txt_serie) + "' and numero = " + CStr(var_j)
                      rsaux1.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                      
                      rsaux1.Open "select customer_trx_id, cadena as cadena, numero from xxvia_tb_control_doc_fiscales where serie = '" + Me.txt_serie + "' and numero = " + CStr(var_j), cnnoracle_4, adOpenDynamic, adLockOptimistic
                      var_cadena = Replace(Replace(rsaux1!Cadena, "T23:", "T00:"), "AUTORIZADO  ", "AUTORIZADO ")
                      var_cadena_rfc = Mid(var_cadena, 34, 12)
                      VAR_CADENA_STR = ""
                      Open ("C:\SISTEMAS\" + Me.txt_serie + CStr(var_j) + ".FAC") For Output As #1
                      For var_i = 1 To Len(var_cadena)
                          'MsgBox Asc(Mid(var_cadena, var_i, 1))
                          If Asc(Mid(var_cadena, var_i, 1)) = 63 Then
                             If Mid(VAR_CADENA_STR, 1, 11) = "COMPROBANTE" Then
                                var_linea_2 = "|" + Trim(var_linea_2) + "|"
                                var_linea = "|" + Trim(var_linea) + "|"
                                VAR_CADENA_STR = Replace(VAR_CADENA_STR, var_linea_2, var_linea)
                             End If
                             'MsgBox VAR_CADENA_STR
                             VAR_CADENA_STR = Replace(VAR_CADENA_STR, "    |", " |")
                             Print #1, VAR_CADENA_STR
                             VAR_CADENA_STR = ""
                          Else
                             VAR_CADENA_STR = VAR_CADENA_STR + Mid(var_cadena, var_i, 1)
                          End If
                          
                      Next var_i
                      Print #1, "FIN:"
                      Close #1
                      rsaux1.Close
                      If Me.txt_serie <> "FAEVII" Then
                         If VAR_33 = 1 Then
                            'x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".FAC" + "|https://facturas2.vianney.mx/cgi-bin/cfds/timbrarGR33|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
                         Else
                            x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".FAC" + "|https://facturas.vianney.mx/cgi-bin/cfds/timbrarGR|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
                         End If
                      Else
                         MsgBox "Selecciona la opción de exporta", vbOKOnly, "ATENCION"
                      End If
                      End If
                      
                  Next var_j
                  
                  MsgBox "Se a terminado el proceso", vbOKOnly, "ATENCION"
               Else
                  MsgBox "Selecciona la opción de exporta", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "Serie invalida", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El número final de factura debe de ser mayor al inicial"
         End If
      Else
         MsgBox "Número de factura final incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Número de factura inical incorrecto", vbOKOnly, "ATENCION"
   End If
   Exit Sub
salir1:
      If Err.Number = -2147217900 Then
         'MsgBox Err.Description
         rsaux14.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         rsaux14.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         Resume
      End If
archivo:

End Sub

Private Sub Command4_Click()
If IsNumeric(Me.txt_de) Then
      If IsNumeric(Me.txt_a) Then
         If CDbl(Me.txt_de) <= CDbl(Me.txt_a) Then
            If Me.txt_serie <> "" Then
               var_posible = 0
               If var_posible = 0 Then
                  For var_j = CDbl(Me.txt_de) To CDbl(Me.txt_a)
                  URL = "https://facturas.vianney.mx/cgi-bin/cfds/cfdsORACLE33?cmd=download_pdf&rfc_emisor=CAN060117PP1&serie=" + Trim(Me.txt_serie) + "&folio=" + Trim(CStr(var_j))
                      'URL = "https://facturas2.vianney.mx/cgi-bin/cfds/cfdsORACLE?cmd=download_pdf&rfc_emisor=CAN60117PP1&serie=" + Trim(Me.txt_serie) + "&folio=" + Trim(CStr(var_j))
                       'URL = "https://facturas.vianney.mx/cgi-bin/cfds/cfdsORACLE33?cmd=download_pdf&rfc_emisor=VTH981105F90&serie=" + Trim(Me.txt_serie) + "&folio=" + Trim(CStr(var_j))
                      buf = Split(URL, ".")
                      ext = buf(UBound(buf))
                      strSavePath = "C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(CStr(var_j)) + ".pdf"
                      ret = URLDownloadToFile(0, URL, strSavePath, 0, 0)
                      If ret = 0 Then
                         Call ShellExecute(Me.hwnd, "print", "C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(CStr(var_j)) + ".PDF", vbNullString, vbNullString, SW_SHOWNORMAL = 1)
                         tamano = FileLen("C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(CStr(var_j)) + ".PDF")
                         If tamano > 1000 Then
                            Call ShellExecute(Me.hwnd, "open", "C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(CStr(var_j)) + ".PDF", vbNullString, vbNullString, SW_SHOWNORMAL = 1)
                         Else
                            MsgBox "El documento " + Trim(Me.txt_serie) + Trim(CStr(var_j)) + " no existe", vbOKOnly, "ATENCION"
                         End If
                      Else
                         MsgBox "Error en la factura " + Me.txt_serie + Trim(CStr(var_j))
                      End If
                  Next var_j
               Else
                  MsgBox "No se han generado todas las facturas, puede que todavia no existan, por favor oprima el bóton de ARCHIVO", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "Se debe de indicar una serie", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El número final de factura debe de ser mayor al inicial"
         End If
      Else
         MsgBox "Número de factura final incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Número de factura inical incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Command5_Click()
Dim VAR_UOM_ADUANA As String
If IsNumeric(Me.txt_de) Then
      If IsNumeric(Me.txt_a) Then
         If CDbl(Me.txt_de) <= CDbl(Me.txt_a) Then
            If Me.txt_serie = "FAEVII" Then
               For var_j = CDbl(Me.txt_de) To CDbl(Me.txt_a)
                   'rsaux1.Open "select customer_trx_id, cadena as cadena, numero from xxvia_tb_control_doc_fiscales where serie = '" + Me.txt_serie + "' and numero = " + CStr(var_j), cnnoracle_4, adOpenDynamic, adLockOptimistic
                   rsaux1.Open "select A.customer_trx_id, cadena as cadena, numero, c.PARTY_SITE_NUMBER cliente, TOTAL_PIEZAS from xxvia_tb_control_doc_fiscales A, RA_CUSTOMER_tRX_ALL B, XXVIA_VW_CLIENTES_BCP C  where serie = '" + Me.txt_serie + "' and numero = " + CStr(var_j) + " AND A.CUSTOMER_tRX_ID = B.CUSTOMER_tRX_ID AND B.BILL_TO_SITE_USE_ID = C.SITE_USE_ID", cnnoracle_4, adOpenDynamic, adLockOptimistic
                   If rsaux1.EOF Then
                      
                      On Error GoTo salir1:
                      var_posible = 1
                      CC = 1
                      If CC = 1 Then
                         var_cadena = "select CUSTOMER_TRX_ID from ra_customer_trx_all a, fnd_document_sequences b where b.name = ? AND TRX_NUMBER = ? and a.doc_sequence_id = b.doc_sequence_id(+) and rownum = 1"
                         With comandoORA
                              .ActiveConnection = cnnoracle_4
                              .CommandType = adCmdText
                              .CommandText = var_cadena
                              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Trim(Me.txt_serie))
                              .Parameters.Append parametro
                              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, CStr(var_j))
                              .Parameters.Append parametro
                         End With
                         Set rsaux9 = comandoORA.execute
                         Set comandoORA = Nothing
                         Set parametro = Nothing
                         If Not rsaux9.EOF Then
                             
                            var_cadena = "SELECT * FROM XXVIA_TB_SERIES_33 WHERE SERIE = ?"
                            With comandoORA
                                 .ActiveConnection = cnnoracle_4
                                 .CommandType = adCmdText
                                 .CommandText = var_cadena
                                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_serie)
                                 .Parameters.Append parametro
                            End With
                            Set rsaux8 = comandoORA.execute
                            Set comandoORA = Nothing
                            Set parametro = Nothing
                            If rsaux8.EOF Then
                               CC = 2
                            Else
                               CC = 1
                            End If
                            If CC = 1 Then
                               var_cadena = "CALL XXVIA_SP_CREA_EFLOW_33(?)"
                               With comandoORA
                                    .ActiveConnection = cnnoracle_4
                                    .CommandType = adCmdText
                                    .CommandText = var_cadena
                                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, rsaux9!customer_Trx_id)
                                    .Parameters.Append parametro
                               End With
                               Set rsaux8 = comandoORA.execute
                               Set comandoORA = Nothing
                               Set parametro = Nothing
                               rsaux1.Close
                               rsaux1.Open "select customer_trx_id, cadena as cadena, numero from xxvia_tb_control_doc_fiscales where serie = '" + Me.txt_serie + "' and numero = " + CStr(var_j), cnnoracle_4, adOpenDynamic, adLockOptimistic
                               GoTo SIGUE
                               
                            End If
                         End If
                      End If
                      
                      
                      
                      var_posible = 1
                      MsgBox "No existe aun la factura " + Me.txt_serie + CStr(var_j) + ", ejecute nuevamente el concurrente", vbOKOnly, "ATENCION"
                   Else
SIGUE:
                      cnn.BeginTrans
                      rsaux2.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_COMPLEMENTO_CE_33", cnn, adOpenDynamic, adLockOptimistic
                      If Not rsaux2.EOF Then
                         var_consecutivo = IIf(IsNull(rsaux2(0).Value), 0, rsaux2(0).Value) + 1
                      End If
                      If rsaux3.State = 1 Then
                         rsaux3.Close
                      End If
                      rsaux3.Open "INSERT INTO TB_TEMP_COMPLEMENTO_CE_33 (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                      rsaux2.Close
                      cnn.CommitTrans
                      
                      var_cadena = rsaux1!Cadena
                      var_cadena = Replace(var_cadena, "VIANNEY CATALOG,LLC|USA|XEXX010101000|G01|", "VIANNEY CATALOG,LLC|USA|760739149|G01|")
                      var_cadena_rfc = Mid(var_cadena, 34, 12)
                      VAR_CADENA_STR = ""
                      Open ("C:\SISTEMAS\COMPL" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".TXT") For Output As #3
                      Open ("C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".FAC") For Output As #1
                      var_pedido = ""
                      For var_i = 1 To Len(var_cadena)
                          'MsgBox Asc(Mid(var_cadena, var_i, 1))
                          
                          
                                strconsulta = "select * from XXVIA_TB_DOC_CON_DIFERENCIA  where serie = ? and numero = ?"
                                With comandoORA
                                     .ActiveConnection = cnnoracle_4
                                     .CommandType = adCmdText
                                     .CommandText = strconsulta
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_serie)
                                     .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, var_j)
                                     .Parameters.Append parametro
                                End With
                                Set rsaux9 = comandoORA.execute
                                Set comandoORA = Nothing
                                Set parametro = Nothing
                          
                                var_diferencia = 0
                                
                                If Not rsaux9.EOF Then
                                   VAR_IMPORTE_TOTAL = rsaux1!TOTAL_PIEZAS + (rsaux9!diferencia)
                                Else
                                   VAR_IMPORTE_TOTAL = rsaux1!TOTAL_PIEZAS
                                End If
                                rsaux9.Close
                          
                          If Asc(Mid(var_cadena, var_i, 1)) = 63 Then
                             If Trim(VAR_CADENA_STR) = "CONDICIONES_PAGO:- -" And Mid(Me.txt_serie, 1, 2) = "NC" Then
                                VAR_CADENA_STR = "CONDICIONES_PAGO:NO IDENTIFICADO"
                             End If
                             'If Mid(var_cadena_str, 1, 6) = "FECHA:" Or Mid(var_cadena_str, 1, 16) = "FECHA:2014-08-08" Or Mid(var_cadena_str, 1, 16) = "FECHA:2014-08-09" Or Mid(var_cadena_str, 1, 16) = "FECHA:2014-08-10" Or Mid(var_cadena_str, 1, 16) = "FECHA:2014-08-11" Or Mid(var_cadena_str, 1, 16) = "FECHA:2014-08-11" Then
                             If Mid(VAR_CADENA_STR, 1, 46) = "LUGAR_EXPEDICION:AGUASCALIENTES,AGUASCALIENTES" Then
                                'var_cadena_str = "LUGAR_EXPEDICION:AGU,001"
                                VAR_CADENA_STR = "LUGAR_EXPEDICION:20290"
                             End If
                             If Mid(VAR_CADENA_STR, 1, 6) = "FECHA:" Then
                                'var_cadena_str = "FECHA:2016-01-31T23:59:00 "
                                'var_cadena_str = "FECHA:2016-08-28T23:01:00 "
                             End If
                             If Mid(VAR_CADENA_STR, 1, 33) = "COLONIA_EXPEDICION:Cd. Industrial" Then
                                VAR_CADENA_STR = "MUNICIPIO_EXPEDICION:001"
                                VAR_CADENA_STR = "COLONIA_EXPEDICION:0027"
                             End If
                             If Mid(VAR_CADENA_STR, 1, 32) = "CIUDAD_EXPEDICION:AGUASCALIENTES" Then
                                VAR_CADENA_STR = "CIUDAD_EXPEDICION:01"
                             End If
                             If Mid(VAR_CADENA_STR, 1, 38) = "ESTADO_EXPEDICION:AGUASCALIENTES|20290" Then
                                VAR_CADENA_STR = "ESTADO_EXPEDICION:AGU|20290 "
                             End If
                             If Mid(VAR_CADENA_STR, 1, 11) = "PDF_VIANNEY" Then
                                var_cliente = ""
                                var_jj = Len(VAR_CADENA_STR)
                                While VAR_ZZ <> "|"
                                      var_cliente = Mid(VAR_CADENA_STR, var_jj, 1) + var_cliente
                                      var_jj = var_jj - 1
                                      VAR_ZZ = Mid(VAR_CADENA_STR, var_jj, 1)
                                Wend
                                
                                
                                'var_cliente = Mid(var_Cadena_str, 9, 100)
                                var_cliente = rsaux1!Cliente
                                'rsaux9.Open "select * from tb_oracle_incoterms where clave = '" + VAR_CLIENTE + "'", cnn, adOpenDynamic, adLockOptimistic
                                
                                strconsulta = "SELECT nvl(incoterm,' ') incoterm, nvl(pais,' ') pais, nvl(estado,' ') estado, nvl(cp,' ') cp, nvl(reg_tributario,' ') reg_tributario, nvl(incoterm2,' ') incoterm2, nvl(observaciones,' ') observaciones FROM XXVIA_tB_COMERCIO_eXTERIOR where clave = ? and rownum = 1"
                                With comandoORA
                                     .ActiveConnection = cnnoracle_4
                                     .CommandType = adCmdText
                                     .CommandText = strconsulta
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_cliente)
                                     .Parameters.Append parametro
                                End With
                                Set rsaux9 = comandoORA.execute
                                Set comandoORA = Nothing
                                Set parametro = Nothing
                                
                                
                                var_cliente_e = var_cliente
                                If Not rsaux9.EOF Then
                                   var_incoterm = IIf(IsNull(rsaux9!incoterm), "", rsaux9!incoterm)
                                   var_pais_destino = IIf(IsNull(rsaux9!pais), "", rsaux9!pais)
                                   var_estado_Destino = IIf(IsNull(rsaux9!estado), "", rsaux9!estado)
                                   var_rfc_e = IIf(IsNull(rsaux9!reg_tributario), "", rsaux9!reg_tributario)
                                Else
                                   var_incoterm = ""
                                   var_pais_destino = ""
                                   var_estado_Destino = ""
                                   var_rfc_e = ""
                                End If
                                rsaux9.Close
                                strconsulta = "select * from XXVIA_VW_CLIENTES_BCP WHERE PARTY_SITE_NUMBER = ?"
                                With comandoORA
                                     .ActiveConnection = cnnoracle_4
                                     .CommandType = adCmdText
                                     .CommandText = strconsulta
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_cliente)
                                     .Parameters.Append parametro
                                End With
                                Set rsaux8 = comandoORA.execute
                                Set comandoORA = Nothing
                                Set parametro = Nothing
                                If Not rsaux8.EOF Then
                                   var_calle_e = IIf(IsNull(rsaux8!calle), "", rsaux8!calle)
                                   var_numero_e = IIf(IsNull(rsaux8!NUM_CALLE), "", rsaux8!NUM_CALLE)
                                   var_numero_interior_e = IIf(IsNull(rsaux8!num_interior), "", rsaux8!num_interior)
                                   var_colonia_e = IIf(IsNull(rsaux8!colonia), "", rsaux8!colonia)
                                   var_ciudad_e = IIf(IsNull(rsaux8!ciudad), "", rsaux8!ciudad)
                                   var_municipio_e = IIf(IsNull(rsaux8!municipio), "", rsaux8!municipio)
                                   VAR_CP_E = IIf(IsNull(rsaux8!codigo_postal), "", rsaux8!codigo_postal)
                                   var_nombre_e = IIf(IsNull(rsaux8!razon_social_cliente), "", rsaux8!razon_social_cliente)
                                   If var_estado_Destino = " " Then
                                      var_estado_Destino = IIf(IsNull(rsaux8!estado), "", rsaux8!estado)
                                   End If
                                Else
                                   var_calle_e = ""
                                   var_numero_e = ""
                                   var_numero_interior_e = ""
                                   var_colonia_e = ""
                                   var_ciudad_e = ""
                                   var_municipio_e = ""
                                   VAR_CP_E = ""
                                   var_nombre_e = ""
                                End If
                                rsaux8.Close
                                var_codigo_postal_e = VAR_CP_E
                             End If
                             If Mid(VAR_CADENA_STR, 1, 9) = "PAIS_CTE:" Then
                                VAR_CADENA_STR = "PAIS_CTE:" + var_pais_destino
                             End If
                             If Mid(VAR_CADENA_STR, 1, 7) = "PEDIDO:" Then
                                   var_pedido = var_pedido + Mid(VAR_CADENA_STR, 8, Len(VAR_CADENA_STR))
                             End If
                             If Mid(VAR_CADENA_STR, 1, 6) = "TOTAL:" Then
                                var_total_s = Mid(VAR_CADENA_STR, 7, 100)
                             End If
                             If Mid(VAR_CADENA_STR, 1, 6) = "LINEA:" Then
                                'LINEA:90.00|00002635|CATÁLOGO VNG 17 GUATEMALA *|Pieza|1.140900|102.680000
                                var_precio_s = ""
                                VAR_CODIGO_S = ""
                                VAR_cANTIDAD_S = ""
                                var_descripcion_s = ""
                                var_importe_s = ""
                                var_unidad_s = ""
                                var_Cantidad_completa = 0
                                var_codigo_completo = 1
                                var_precio_completo = 1
                                var_importe_completo = 1
                                var_descripcion_completa = 1
                                var_unidad_completa = 1
                                If Mid(VAR_CADENA_STR, 1, 5) = "LINEA" Then
                                var_codigo_completo = 0
                                For VAR_Z = 15 To Len(VAR_CADENA_STR)
                                    If var_codigo_completo = 0 Then
                                       If Mid(VAR_CADENA_STR, VAR_Z, 1) = "|" Then
                                          var_descripcion_completa = 0
                                          var_codigo_completo = 1
                                          GoTo completo:
                                       Else
                                          VAR_cANTIDAD_S = VAR_cANTIDAD_S + Mid(VAR_CADENA_STR, VAR_Z, 1)
                                       End If
                                    End If
                                    If var_Cantidad_completa = 0 Then
                                       If Mid(VAR_CADENA_STR, VAR_Z, 1) = "|" Then
                                          var_Cantidad_completa = 1
                                          var_codigo_completo = 0
                                          GoTo completo:
                                       Else
                                          VAR_CODIGO_S = VAR_CODIGO_S + Mid(VAR_CADENA_STR, VAR_Z, 1)
                                       End If
                                    End If
                                    If var_unidad_completa = 0 Then
                                       If Mid(VAR_CADENA_STR, VAR_Z, 1) = "|" Then
                                          var_unidad_completa = 1
                                          var_precio_completo = 0
                                          GoTo completo:
                                       Else
                                          var_unidad_s = var_unidad_s + Mid(VAR_CADENA_STR, VAR_Z, 1)
                                       End If
                                    End If
                                    
                                    
                                    
                                    If var_descripcion_completa = 0 Then
                                       If Mid(VAR_CADENA_STR, VAR_Z, 1) = "|" Then
                                          var_unidad_completa = 0
                                          var_descripcion_completa = 1
                                          GoTo completo:
                                       Else
                                          var_descripcion_s = var_descripcion_s + Mid(VAR_CADENA_STR, VAR_Z, 1)
                                       End If
                                    End If
                                    
                                    
                                    If var_precio_completo = 0 Then
                                       If Mid(VAR_CADENA_STR, VAR_Z, 1) = "|" Then
                                          var_importe_completo = 0
                                          var_precio_completo = 1
                                          GoTo completo:
                                       Else
                                          var_precio_s = var_precio_s + Mid(VAR_CADENA_STR, VAR_Z, 1)
                                       End If
                                    End If
                                    
                                    If var_importe_completo = 0 Then
                                       If Mid(VAR_CADENA_STR, VAR_Z, 1) = "|" Then
                                          var_importe_completo = 1
                                          var_precio_completo = 1
                                          GoTo completo:
                                       Else
                                          var_importe_s = var_importe_s + Mid(VAR_CADENA_STR, VAR_Z, 1)
                                       End If
                                    End If
                                    
                                    
completo:
                                Next VAR_Z
                                
                                
                                
                                
                                strconsulta = "select * from XXVIA_TB_COMPLEMENTOS_PK_LIST WHERE CODIGO = ?"
                                With comandoORA
                                     .ActiveConnection = cnnoracle_4
                                     .CommandType = adCmdText
                                     .CommandText = strconsulta
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, VAR_CODIGO_S)
                                     .Parameters.Append parametro
                                End With
                                Set rsaux8 = comandoORA.execute
                                Set comandoORA = Nothing
                                Set parametro = Nothing
                                If Not rsaux8.EOF Then
                                   VAR_FRACCION_A = CStr(IIf(IsNull(rsaux8!fraccion_arancelaria), "", rsaux8!fraccion_arancelaria))
                                Else
                                   VAR_FRACCION_A = ""
                                End If
                                'se quita la fraccion arancelaria
                                'Print #3, "MERCANCIA:" + VAR_CODIGO_S + "|" + var_fraccion_a + "|06|" + var_Cantidad_s + "|" + CStr(Round(CDbl(var_precio_s), 2)) + "|" + CStr(Round(CDbl(Trim(var_importe_s)), 2)) + "|VIANNEY|||"
                                '3.3
                                'MERCANCIA:num identificacion|c_fraccionarancelaria|cantidad aduana|c_unidadaduana|valor unitario aduana|valor dolares
                                'DESC_ESPECIFICA:marca|modelo|submodelo|num serie
                                'se hizo el cambio porque es distinto el layout
                                'Print #3, "MERCANCIA:" + var_cantidad_s + "||06|" + var_codigo_s + "|" + CStr(Round(CDbl(var_importe_s), 2)) + "|" + CStr(Round(CDbl(Trim(var_importe_s) * CDbl(var_cantidad_s)), 2)) + "|VIANNEY|||"
                                rsaux12.Open "SELECT * FROM TB_ORACLE_FRACCIONES_aRANCELARIAS_33 WHERE fraccion_arancelaria = '" + VAR_FRACCION_A + "'", cnn, adOpenDynamic, adLockOptimistic
                                If Not rsaux12.EOF Then
                                   VAR_UOM_ADUANA = CStr(IIf(IsNull(rsaux12!UOM), "06", rsaux12!UOM))
                                Else
                                   VAR_UOM_ADUANA = "06"
                                End If
                                If Len(VAR_UOM_ADUANA) = 1 Then
                                   VAR_UOM_ADUANA = "0" + VAR_UOM_ADUANA
                                End If
                                rsaux12.Close
                                'SE QUITA PARA QUE NO SE DUPLIQUEN CODIGOS
                                rsaux9.Open "SELECT * FROM TB_TEMP_COMPLEMENTO_CE_33 WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND CODIGO = '" + VAR_CODIGO_S + "'", cnn, adOpenDynamic, adLockOptimistic
                                If Not rsaux9.EOF Then
                                   rsaux10.Open "UPDATE TB_TEMP_COMPLEMENTO_CE_33 SET CANTIDAD = CANTIDAD + " + CStr(VAR_cANTIDAD_S) + ", IMPORTE = IMPORTE + " + CStr(Round(CDbl(Trim(var_importe_s)) * CDbl(VAR_cANTIDAD_S), 2)) + " WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND CODIGO = '" + VAR_CODIGO_S + "'", cnn, adOpenDynamic, adLockOptimistic
                                Else
                                   var_cadena_2 = "INSERT INTO TB_TEMP_COMPLEMENTO_CE_33 (INTE_TEM_CONSECUTIVO, CODIGO, FRACCION, CANTIDAD, UNIDAD, PRECIO, IMPORTE)"
                                   'var_cantidad =
                                   var_cadena_2 = var_cadena_2 + " VALUES                               (" + CStr(var_consecutivo) + ",'" + VAR_CODIGO_S + "','" + CStr(VAR_FRACCION_A) + "'," + CStr(Round(VAR_cANTIDAD_S, 2)) + ",'" + VAR_UOM_ADUANA + "'," + CStr((CDbl(var_importe_s))) + "," + CStr(Round(CDbl(Trim(var_importe_s)) * CDbl(VAR_cANTIDAD_S), 6)) + ")"
                                   'MsgBox var_cadena
                                   rsaux10.Open var_cadena_2, cnn, adOpenDynamic, adLockOptimistic
                                End If
                                rsaux9.Close
                                'Print #3, "MERCANCIA:" + var_codigo_s + "|" + CStr(VAR_FRACCION_A) + "|" + Format(Round(var_cantidad_s, 2), "#######0.00") + "|" + VAR_UOM_ADUANA + "|" + Format(CStr((CDbl(var_importe_s))), "#######0.00") + "|" + Format(CStr((Round(CDbl(Trim(var_importe_s)) * CDbl(var_cantidad_s), 2))), "########0.00") + "|" + "VIANNEY|" + "|" + "|"
                                rsaux8.Close
                                End If
                                
                                
                                
                             End If
                             Print #1, VAR_CADENA_STR
                             If Mid(VAR_CADENA_STR, 1, 33) = "COLONIA_EXPEDICION:0027" Then
                                VAR_CADENA_STR = "MUNICIPIO_EXPEDICION:001"
                                'var_cadena_str = "COLONIA_EXPEDICION:0027"
                                Print #1, VAR_CADENA_STR
                             End If
                             
                             VAR_CADENA_STR = ""
                          Else
                             VAR_CADENA_STR = VAR_CADENA_STR + Mid(var_cadena, var_i, 1)
                          End If
                      Next var_i
                      Close #3
                      strconsulta = "SELECT A.CUSTOMER_tRX_ID, EXCHANGE_RATE FROM XXVIA_tB_CONTROL_DOC_FISCALES A, RA_CUSTOMER_tRX_ALL B WHERE SERIE = 'FAEVII' AND NUMERO = ? AND A.CUSTOMER_tRX_ID = B.CUSTOMER_tRX_ID"
                      With comandoORA
                           .ActiveConnection = cnnoracle_4
                           .CommandType = adCmdText
                           .CommandText = strconsulta
                           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_j)
                           .Parameters.Append parametro
                      End With
                      Set rsaux8 = comandoORA.execute
                      Set comandoORA = Nothing
                      Set parametro = Nothing
                      If Not rsaux8.EOF Then
                         var_tipo_cambio_s = CStr(Round(IIf(IsNull(rsaux8!EXCHANGE_RATE), 1, rsaux8!EXCHANGE_RATE), 6))
                      Else
                         var_tipo_cambio_s = 1
                      End If
                      'Print #3, "MERCANCIA:" + VAR_CODIGO_S + "|" + var_fraccion_a + "|06|" + var_Cantidad_s + "|" + var_precio_s + "|" + Trim(var_importe_s) + "|VIANNEY|||"
                      rsaux8.Close
                      
                      'c_incoterm = "DAT"
                      Print #1, "COMPLEMENTO_CE_INICIO:"
                      'Print #1, "COMERCIO_EXTERIOR:1.1|01|2|A1|0|||" + var_incoterm + "|0|" + var_tipo_cambio_s + "|" + Trim(var_total_s)
                      Print #1, "COMERCIO_EXTERIOR:1.1||2|A1|0|||" + var_incoterm + "|0|" + CStr(var_tipo_cambio_s) + "|" + Trim(Format(VAR_IMPORTE_TOTAL, "#########.00"))
                      Print #1, "OBSERVACIONES_INICIO:"
                      Print #1, "OBSERVACIONES_FIN:"
                      Print #1, "EMISOR:|SALVADOR QUEZADA LIMON|1512||0027|001|01|AGU|MEX|20040"
                      Print #1, "PROPIETARIO:" + var_rfc_e + "|" + var_pais_destino
                      'RECEPTOR:num reg trubitario|calle|numero_externo|numero_interno|c_colonia|c_municipio|c_ciudad|c_estado|c_pais|cp (la parte de dirección es requerida para CFDI 3.3, se ignora en CFDI 3.2)
                      'Print #1, "RECEPTOR:" + var_rfc_e + "|" + var_calle_e + "|" + var_numero_e + "|" + var_numero_interior_e + "|" + var_colonia_e + "|" + var_municipio_e + "|" + var_ciudad_e + "|" + var_estado_Destino + "|" + var_pais_destino + "|" + var_codigo_postal_e
                      Print #1, "RECEPTOR:" + "" + "|" + var_calle_e + "|" + var_numero_e + "|" + var_numero_interior_e + "|" + var_colonia_e + "|" + var_municipio_e + "|" + var_ciudad_e + "|" + var_estado_Destino + "|" + var_pais_destino + "|" + var_codigo_postal_e
                      'DESTINATARIO:numreg tributario|nombre|calle|numero_externo|numero_interno|c_colonia|c_municipio|c_ciudad|c_estado|c_pais|cp
                      Print #1, "DESTINATARIO:" + var_rfc_e + "|" + var_nombre_e + "|" + var_calle_e + "|" + var_numero_e + "|" + var_numero_interior_e + "|" + var_colonia_e + "|" + var_municipio_e + "|" + var_ciudad_e + "|" + var_estado_Destino + "|" + var_pais_destino + "|" + var_codigo_postal_e
                      x = 1
                      If x = 0 Then
                         Open "c:\sistemas\COMPL" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".TXT" For Input As #3
                         While Not EOF(3)
                               Line Input #3, mivariable
                               Print #1, mivariable
                         Wend
                         Close #3
                      Else
                         rsaux9.Open "SELECT * FROM TB_TEMP_COMPLEMENTO_CE_33 WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND CODIGO IS NOT NULL", cnn, adOpenDynamic, adLockOptimistic
                         While Not rsaux9.EOF
                               If rsaux9!codigo = "00052274" Then
                                  SF = 0
                               End If
                               'var_cantidad = Round(rsaux9!Importe / Round(rsaux9!Precio, 2), 3)
                               'var_cantidad = 1
                               'var_cantidad = Round(rsaux9!Importe / Round(rsaux9!Precio, 3), 3)
                               'var_precio = Round(rsaux9!Importe, 3)
                               Print #1, "MERCANCIA:" + rsaux9!codigo + "|" + rsaux9!fraccion + "|" + Format(Round(rsaux9!cantidad, 2), "#######0.00") + "|" + rsaux9!unidad + "|" + Format(CStr(rsaux9!Precio), "#######0.00") + "|" + Format(CStr(rsaux9!Importe), "########0.00") + "|" + "VIANNEY|" + "|" + "|"
                               'Print #1, "MERCANCIA:" + rsaux9!codigo + "|" + rsaux9!fraccion + "|" + Format(Round(rsaux9!cantidad, 2), "#######0.000") + "|" + rsaux9!unidad + "|" + Format(CStr(var_precio), "#######0.00") + "|" + Format(CStr(rsaux9!Importe), "########0.00") + "|" + "VIANNEY|" + "|" + "|"
                               rsaux9.MoveNext
                         Wend
                         rsaux9.Close
                      End If
                      Print #1, "COMPLEMENTO_CE_FIN:"
                      
                      Print #1, "FIN:"
                      Close #1
                      var_archivo = "C:\SISTEMAS\sube_fact" + Trim(Str(var_j)) + ".bat"
                      x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".FAC" + "|https://facturas2.vianney.mx/cgi-bin/cfds/timbrarGR33|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
                      
                      'x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".FAC" + "|https://facturas.vianney.mx/cgi-bin/cfds/timbrarGR|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
                                       
                      'URL = "https://facturas.vianney.mx/cgi-bin/cfds/cfdsORACLE?cmd=download_pdf&rfc_emisor=" + var_cadena_rfc + "&serie=" + Trim(Me.txt_serie) + "&folio=" + Trim(CStr(var_j))
                      'buf = Split(URL, ".")
                      'ext = buf(UBound(buf))
                      'strSavePath = "C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(CStr(var_j)) + ".pdf"
                      'ret = URLDownloadToFile(0, URL, strSavePath, 0, 0)
                   End If
                   rsaux1.Close
               Next var_j
               MsgBox "Se a terminado el proceso", vbOKOnly, "ATENCION"
            Else
               MsgBox "Serie invalida para exportaciones", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El número final de factura debe de ser mayor al inicial"
         End If
      Else
         MsgBox "Número de factura final incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Número de factura inical incorrecto", vbOKOnly, "ATENCION"
   End If
salir1:
      If Err.Number = -2147217900 Then
         MsgBox Err.Description
         rsaux14.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         rsaux14.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         Resume
Else
   MsgBox "intente de nuevo."
   
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
   If rsaux10.State = 1 Then
      rsaux10.Close
   End If
   If rsaux11.State = 1 Then
      rsaux11.Close
   End If
   If rsaux12.State = 1 Then
      rsaux12.Close
   End If
   If rsaux13.State = 1 Then
      rsaux13.Close
   End If
   If rsaux14.State = 1 Then
      rsaux14.Close
   End If
   If rsaux15.State = 1 Then
      rsaux15.Close
   End If
      End If

End Sub

Private Sub Command6_Click()
            If Me.txt_serie <> "" Then
               If Me.txt_serie <> "FAEVIooI" Then
                  For var_j = CDbl(Me.txt_de) To CDbl(Me.txt_a)
                      rsaux1.Open "select customer_trx_id, cadena as cadena, numero from xxvia_tb_control_doc_fiscales where serie = '" + Me.txt_serie + "' and numero = " + CStr(var_j), cnnoracle_4, adOpenDynamic, adLockOptimistic
                      If rsaux1.EOF Then
                         var_posible = 1
                         MsgBox "No existe aun la factura " + Me.txt_serie + CStr(var_j) + ", ejecute nuevamente el concurrente", vbOKOnly, "ATENCION"
                      Else
                     
                         var_cadena = rsaux1!Cadena
                         var_cadena_rfc = Mid(var_cadena, 34, 12)
                         VAR_CADENA_STR = ""
                         Open ("C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".FAC") For Output As #1
                         For var_i = 1 To Len(var_cadena)
                             'MsgBox Asc(Mid(var_cadena, var_i, 1))
                             If Asc(Mid(var_cadena, var_i, 1)) = 63 Then
                                If Trim(VAR_CADENA_STR) = "CONDICIONES_PAGO:- -" And Mid(Me.txt_serie, 1, 2) = "NC" Then
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
                         var_archivo = "C:\SISTEMAS\sube_fact" + Trim(Str(var_j)) + ".bat"
                         If var_prueba = 2 Then
                                                                                                                                                                 
                            x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".FAC" + "|https://facturas2.vianney.mx/cgi-bin/cfds/cfdsORACLE33|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
                         Else
                            x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".FAC" + "|https://facturas.vianney.mx/cgi-bin/cfds/timbrarGR|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
                         End If
                         'URL = "https://facturas.vianney.mx/cgi-bin/cfds/cfdsORACLE?cmd=download_pdf&rfc_emisor=" + var_cadena_rfc + "&serie=" + Trim(Me.txt_serie) + "&folio=" + Trim(CStr(var_j))
                         'buf = Split(URL, ".")
                         'ext = buf(UBound(buf))
                         'strSavePath = "C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(CStr(var_j)) + ".pdf"
                         'ret = URLDownloadToFile(0, URL, strSavePath, 0, 0)
                      End If
                      rsaux1.Close
                  Next var_j
                  MsgBox "Se a terminado el proceso", vbOKOnly, "ATENCION"
               Else
                  MsgBox "Selecciona la opción de exporta", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "Serie invalida", vbOKOnly, "ATENCION"
            End If

End Sub

Private Sub Form_Load()
   Left = 3000
   If var_clave_usuario_global = "U0000000020" Or var_clave_usuario_global = "8" Or var_clave_usuario_global = "U0000000845" Then
      Me.Height = 4170
      Top = 2000
      Me.cmb_metodo_2.Enabled = True
      Me.cmb_metodos.Enabled = True
      Me.cmb_metodos.Enabled = True
      Me.txt_factura.Enabled = True
      Me.txt_cfdi.Enabled = True
   Else
      Me.cmb_metodo_2.Enabled = False
      Me.cmb_metodos.Enabled = False
      Me.cmb_metodos.Enabled = False
      Me.txt_factura.Enabled = False
      Me.txt_cfdi.Enabled = False
      Top = 3000
      Me.Height = 2010
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_lineas)
End Sub

Private Sub txt_a_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      'Command1.SetFocus
   End If
End Sub


Private Sub txt_de_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Me.txt_a.SetFocus
   End If
End Sub

Private Sub txt_de_LostFocus()
   Me.txt_a = Me.txt_de
End Sub

Private Sub txt_factura_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim cn As New ADODB.Connection
      Dim DSN As String
      Dim cn2 As New ADODB.Connection
      DSN = "eflow"
      cn.Open ("DSN=" & DSN & ";")
      
      Set rsaux1 = cn.execute("SELECT * FROM facturas where factura = '" + Me.txt_factura + "'")
      If Not rsaux1.EOF Then
         Me.txt_cfdi = IIf(IsNull(rsaux1!sat_uuid), "", rsaux1!sat_uuid)
      Else
         MsgBox "La factura no existe", vbOKOnly, "ATENCION"
         Me.txt_cfdi = ""
      End If
   
   End If
End Sub

Private Sub txt_serie_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   
   If KeyAscii = 13 Then
      Me.txt_de.SetFocus
   End If
End Sub
