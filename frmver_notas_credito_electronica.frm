VERSION 5.00
Begin VB.Form frmver_notas_credito_electronica 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imprimir documentos electronicos"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command6 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frmver_notas_credito_electronica.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Reimprimir facturas"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   0
      TabIndex        =   11
      Top             =   345
      Width           =   4530
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmver_notas_credito_electronica.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4125
      Picture         =   "frmver_notas_credito_electronica.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   1170
      Left            =   45
      TabIndex        =   5
      Top             =   375
      Width           =   4500
      Begin VB.TextBox txt_serie 
         Height          =   390
         Left            =   795
         TabIndex        =   0
         Top             =   210
         Width           =   1575
      End
      Begin VB.TextBox txt_de 
         Height          =   390
         Left            =   795
         TabIndex        =   1
         Top             =   660
         Width           =   1575
      End
      Begin VB.TextBox txt_a 
         Height          =   390
         Left            =   2760
         TabIndex        =   2
         Top             =   660
         Width           =   1575
      End
      Begin VB.TextBox txt_copias 
         Height          =   345
         Left            =   3075
         TabIndex        =   6
         Top             =   225
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   300
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "De:"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "a:"
         Height          =   195
         Left            =   2520
         TabIndex        =   8
         Top             =   750
         Width           =   135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Copias:"
         Height          =   195
         Left            =   2520
         TabIndex        =   7
         Top             =   270
         Visible         =   0   'False
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmver_notas_credito_electronica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_imprimir_Click()
   Me.txt_copias = 1
   If IsNumeric(Me.txt_copias) Then
      If IsNumeric(Me.txt_de) Then
         If IsNumeric(Me.txt_a) Then
            If CDbl(Me.txt_de) <= CDbl(Me.txt_a) Then
               If Trim(Me.txt_serie) <> "" Then
                     For var_j = CDbl(Me.txt_de) To CDbl(Me.txt_a)
                         var_posible = 0
                         var_Archivo = var_ruta_documentos_electronicos_pdf + "\" + Trim(Me.txt_serie) + "\" + Trim(Me.txt_serie) + Trim(CStr(var_j)) + ".pdf"
                         Archivoabuscar = Dir(var_Archivo)
                         If Archivoabuscar = "" Then
                            var_posible = 1
                         End If
                     Next var_j
                     If var_posible = 0 Then
                        var_si = MsgBox("Se van a imprimir los documentos del " + Me.txt_de + " al " + Me.txt_a, vbYesNo, "ATENCION")
                        If var_si = 6 Then
                           var_si = MsgBox("Confirmar la impresi?n de las facturas", vbYesNo, "ATENCION")
                           If var_si = 6 Then
                              For var_j = CDbl(Me.txt_de) To CDbl(Me.txt_a)
                                  If rs.State = 1 Then
                                     rs.Close
                                  End If
                                  rs.Open "select * from tb_encabezado_cartera where vcha_Ser_serie_id = '" + Me.txt_serie + "' and inte_Car_numero = " + CStr(var_j) + " and isnull(char_car_estatus,'') <> 'C'", cnn, adOpenDynamic, adLockOptimistic
                                  If Not rs.EOF Then
                                     var_Archivo = var_ruta_documentos_electronicos_pdf + "\" + Me.txt_serie + "\" + Trim(Me.txt_serie) + Trim(CStr(var_j)) + ".pdf"
                                     Archivoabuscar = Dir(var_Archivo)
                                     If Archivoabuscar <> "" Then
                                        var_Archivo = var_ruta_documentos_electronicos_pdf + "\" + Me.txt_serie + "\" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".PDF"
                                        Archivoabuscar = Dir("c:\archivos de programa\adobe\reader 8.0\reader\acrord32.exe")
                                        If Archivoabuscar <> "" Then
                                           Call Shell("c:\archivos de programa\adobe\reader 8.0\reader\acrord32.exe  /p /h " + var_Archivo, vbMaximizedFocus)
                                        Else
                                           Archivoabuscar = Dir("c:\Program files\adobe\reader 8.0\reader\acrord32.exe")
                                           If Archivoabuscar <> "" Then
                                              Call Shell("c:\Program files\adobe\reader 8.0\reader\acrord32.exe  /p /h " + var_Archivo, vbMaximizedFocus)
                                           Else
                                              Open (App.Path & "\EJPDF" + Trim(Me.txt_serie) + Trim(CStr(var_j)) + ".bat") For Output As #2
                                              Print #2, "START " + var_ruta_documentos_electronicos_pdf + "\" + Me.txt_serie + "\" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".PDF"
                                              Close #2
                                              var_Archivo = App.Path & "\EJPDF" + Trim(Me.txt_serie) + Trim(CStr(var_j)) + ".bat"
                                              x = Shell(var_Archivo, vbHide)
                                           End If
                                        End If
                                     End If
                                  End If
                                  rs.Close
                              Next var_j
                           Else
                              MsgBox "Se a cancelado la impresi?n de las facturas", vbOKOnly, "ATENCION"
                           End If
                        Else
                           MsgBox "Se a cancelado la impresi?n de las facturas", vbOKOnly, "ATENCION"
                        End If
                  Else
                     MsgBox "Las notas de cr?dito no existen", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "Debe de seleccionar una serie", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "La factura inicial no puede ser mayor a la factura final", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "N?mero de factura final incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "N?mero de factura inicial incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "N?mero de copias incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command6_Click()
            If rs.State = 1 Then
               rs.Close
            End If
            var_serie = Me.txt_serie
            
'''''''''''''''''' DEVOLUCIONES


                     rs.Open "select * from VW_DOCUMENTOS_IMPRESION where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'DV' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(txt_de), cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        txt_movimiento = rs!VCHA_MOV_MOVIMIENTO_ID
                        txt_numero = rs!INTE_EMO_NUMERO
''' ''''''''      ''' IMPRESION DE LA NOTA DE CARGO
                        var_k = var_numero_nota_inicio
                        'Close #1
                        Open (var_ruta_documentos_electronicos & "\" + Trim(var_serie) + Trim(Str(rs!inte_Car_numero)) + ".fI") For Output As #1
                        var_cadena = "Outputmode=" + Chr(13) + "<Factura>" + Chr(13) + "<Comprobante>" + Chr(13) + "Version=2.0" + Chr(13) + "Serie=" + rs!vcha_Ser_Serie_id + Chr(13) + "folio=" + CStr(rs!inte_Car_numero) + Chr(13)
                        var_a?o = CStr(Year(rs!dtim_Car_fecha))
                        var_mes = CStr(Month(rs!dtim_Car_fecha))
                        VAR_DIA = CStr(Day(rs!dtim_Car_fecha))
                        var_hora = CStr(Hour(rs!dtim_Car_fecha))
                        var_minuto = CStr(Minute(rs!dtim_Car_fecha))
                        var_segundo = CStr(Second(rs!dtim_Car_fecha))
                        If Len(var_a?o) = 2 Then
                           var_a?o = "20" + var_a?o
                        End If
                        If Len(var_mes) = 1 Then
                           var_mes = "0" + var_mes
                        End If
                        If Len(VAR_DIA) = 1 Then
                           VAR_DIA = "0" + VAR_DIA
                        End If
                        If Len(var_hora) = 1 Then
                           var_hora = "0" + var_hora
                        End If
                        If Len(var_minuto) = 1 Then
                           var_minuto = "0" + var_minuto
                        End If
                        If Len(var_segundo) = 1 Then
                           var_segundo = "0" + var_segundo
                        End If
                        
                        
                        
                        var_contador_renglones = 0
                        If rsaux3.State = 1 Then
                           rsaux3.Close
                        End If
                        rsaux3.Open "select * from TB_DETALLE_DEVOLUCION_IMPORTES_ASIGNADOS where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'DV' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(txt_de), cnn, adOpenDynamic, adLockOptimistic
                        var_linea = ""
                        While Not rsaux3.EOF
                              If Len(Trim(var_linea)) = 0 Then
                                 var_linea = var_linea + CStr(rsaux3!inte_dia_numero) + rsaux3!vcha_dia_serie + " " + Format(rsaux3!floa_dia_importe, "###,###,##0.00")
                              Else
                                 var_linea = var_linea + ", " + CStr(rsaux3!inte_dia_numero) + rsaux3!vcha_dia_serie + " " + Format(rsaux3!floa_dia_importe, "###,###,##0.00")
                              End If
                              rsaux3.MoveNext
                        Wend
                        rsaux3.Close
                                                
                        
                        
                        var_rfc_cliente_1 = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                        var_rfc_cliente = ""
                        If var_rfc_cliente_1 = "" Then
                           var_rfc_cliente = "XAXX010101000"
                        Else
                           For var_j = 1 To Len(var_rfc_cliente_1)
                               If Mid(var_rfc_cliente_1, var_j, 1) <> "-" Then
                                  If Mid(var_rfc_cliente_1, var_j, 1) <> "" Then
                                     If Mid(var_rfc_cliente_1, var_j, 1) <> " " Then
                                        var_rfc_cliente = var_rfc_cliente + Mid(var_rfc_cliente_1, var_j, 1)
                                     End If
                                  End If
                               End If
                           Next var_j
                        End If
                        
                        
                        var_cadena_fecha = var_a?o + "-" + var_mes + "-" + VAR_DIA + "T" + var_hora + ":" + var_minuto + ":" + var_segundo
                        var_cadena = var_cadena + "fecha=" + var_cadena_fecha + Chr(13)
                        var_cadena = var_cadena + "noAprobacion=" + Chr(13)
                        var_cadena = var_cadena + "anoAprobacion=" + Chr(13)
                        var_cadena = var_cadena + "tipoDeComprobante=NOTA DE CREDITO" + Chr(13)
                        var_cadena = var_cadena + "formaDePago=CONTADO" + Chr(13)
                        var_cadena = var_cadena + "condicionesDePago=" + Chr(13)
                        If var_rfc_cliente = "XAXX010101000" Then
                           var_cadena = var_cadena + "subtotal=" + Format(CStr(rs!floa_Car_importe_neto / rs!floa_car_tipo_cambio), "###,###,##0.000000") + Chr(13)
                        Else
                           var_cadena = var_cadena + "subtotal=" + Format(CStr(rs!floa_car_subimporte / rs!floa_car_tipo_cambio), "###,###,##0.000000") + Chr(13)
                        End If
                        var_cadena = var_cadena + "descuento=" + Chr(13)
                        var_cadena = var_cadena + "descuento1=" + Chr(13)
                        var_cadena = var_cadena + "descuento2=" + Chr(13)
                        var_cadena = var_cadena + "conceptodescuento1=" + Chr(13)
                        var_cadena = var_cadena + "conceptodescuento2=" + Chr(13)
                        var_cadena = var_cadena + "tasadescuento1=" + Chr(13)
                        var_cadena = var_cadena + "tasadescuento2=" + Chr(13)
                        If rsaux2.State = 1 Then
                           rsaux2.Close
                        End If
                        rsaux2.Open "select * from tb_empresa_FACTURA_ELECTRONICA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_certificado = rsaux2!vcha_emp_certificado
                        var_expedido = rsaux2!vcha_emp_expedido
                        If var_rfc_cliente = "XAXX010101000" Then
                           var_cadena = var_cadena + "iva=" + Format(CStr(0), "###,###,##0.000000") + Chr(13)
                        Else
                           var_cadena = var_cadena + "iva=" + Format(CStr(rs!floa_car_importe_iva / rs!floa_car_tipo_cambio), "###,###,##0.000000") + Chr(13)
                        End If
                        var_cadena = var_cadena + "total=" + Format(CStr(rs!floa_Car_importe_neto / rs!floa_car_tipo_cambio), "###,###,##0.000000") + Chr(13)
                        var_cadena = var_cadena + "retencion=" + Chr(13)
                        var_cadena = var_cadena + "factorretencioniva=" + Chr(13)
                        var_cadena = var_cadena + "</Comprobante>" + Chr(13) + Chr(13)
                        var_cadena = var_cadena + "<Emisor>" + Chr(13)
                        var_cadena = var_cadena + "erfc=" + rsaux2!VCHA_eMP_RFC + Chr(13)
                        var_cadena = var_cadena + "enombre=" + rsaux2!VCHA_EMP_NOMBRE + Chr(13)
                        var_cadena = var_cadena + "</Emisor>" + Chr(13) + Chr(13)
                        var_cadena = var_cadena + "<DomicilioFiscal>" + Chr(13)
                        var_cadena = var_cadena + "ecalle=" + rsaux2!VCHA_eMP_CALLE + Chr(13)
                        var_cadena = var_cadena + "enoExterior=" + rsaux2!VCHA_eMP_exterior + Chr(13)
                        var_cadena = var_cadena + "enoInterior=" + Chr(13)
                        var_cadena = var_cadena + "ecolonia=" + rsaux2!VCHA_eMP_COLONIA + Chr(13)
                        var_cadena = var_cadena + "elocalidad=" + rsaux2!VCHA_EMP_LOCALIDAD + Chr(13)
                        var_cadena = var_cadena + "ereferencia=" + Chr(13)
                        var_cadena = var_cadena + "emunicipio=" + rsaux2!VCHA_EMP_MUNICIPIO + Chr(13)
                        var_cadena = var_cadena + "eestado=" + rsaux2!VCHA_EMP_ESTADO + Chr(13)
                        var_cadena = var_cadena + "epais=" + rsaux2!VCHA_eMP_PAIS + Chr(13)
                        var_cadena = var_cadena + "ecodigoPostal=" + rsaux2!VCHA_EMP_CODIGO_POSTAL + Chr(13)
                        var_cadena = var_cadena + "etel=" + IIf(IsNull(rsaux2!VCHA_EMP_TELEFONO), "", rsaux2!VCHA_EMP_TELEFONO) + Chr(13)
                        var_cadena = var_cadena + "eemail=" + IIf(IsNull(rsaux2!VCHA_EMP_EMAIL), "", rsaux2!VCHA_EMP_EMAIL) + Chr(13)
                        var_cadena = var_cadena + "</DomicilioFiscal>" + Chr(13) + Chr(13)
                        var_cadena = var_cadena + "<Receptor>" + Chr(13)
                        var_cadena = var_cadena + "noCliente=" + rs!vcha_cli_clave_id + Chr(13)
                        rsaux2.Close
                                       
                        var_rfc_cliente_1 = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                        var_rfc_cliente = ""
                        If var_rfc_cliente_1 = "" Then
                           var_rfc_cliente = "XAXX010101000"
                        Else
                           For var_j = 1 To Len(var_rfc_cliente_1)
                               If Mid(var_rfc_cliente_1, var_j, 1) <> "-" Then
                                  If Mid(var_rfc_cliente_1, var_j, 1) <> "" Then
                                     If Mid(var_rfc_cliente_1, var_j, 1) <> " " Then
                                        var_rfc_cliente = var_rfc_cliente + Mid(var_rfc_cliente_1, var_j, 1)
                                     End If
                                  End If
                               End If
                           Next var_j
                        End If
                        If var_empresa = "03" Or var_empresa = "28" Then
                            var_rfc_cliente = "XEXX010101000"
                        End If
                        var_cadena = var_cadena + "rfc=" + var_rfc_cliente + Chr(13)
                        var_cadena = var_cadena + "nombre=" + rs!VCHA_CLI_NOMBRE + Chr(13)
                        var_cadena = var_cadena + "</Receptor>" + Chr(13) + Chr(13)
                        var_cadena = var_cadena + "<Cliente>" + Chr(13)
                        var_cadena = var_cadena + "domicilio=" + IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + Chr(13)
                        var_cadena = var_cadena + "calle=" + Chr(13)
                        var_cadena = var_cadena + "noExterior=" + Chr(13)
                        var_cadena = var_cadena + "noInterior=" + Chr(13)
                        var_cadena = var_cadena + "colonia=" + IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre) + Chr(13)
                        var_cadena = var_cadena + "localidad=MONTERREY" + Chr(13)
                        rsaux2.Open "select * from vw_clientes where vcha_Cli_clave_id = '" + rs!vcha_cli_clave_id + "'"
                        var_cadena = var_cadena + "referencia=" + Chr(13)
                        var_cadena = var_cadena + "municipio=" + IIf(IsNull(rsaux2!vcha_mun_nombre), "", rsaux2!vcha_mun_nombre) + Chr(13)
                        var_cadena = var_cadena + "estado=" + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre) + Chr(13)
                        VAR_NOMBRE_PAIS = IIf(IsNull(rs!vcha_pai_nombre), "MEXICO", rs!vcha_pai_nombre)
                        If Trim(VAR_NOMBRE_PAIS) = "" Then
                           VAR_NOMBRE_PAIS = "MEXICO"
                        End If
                        var_cadena = var_cadena + "pais=" + VAR_NOMBRE_PAIS + Chr(13)
                        var_cadena = var_cadena + Chr(13)
                        var_cadena = var_cadena + "codigoPostal=" + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP) + Chr(13)
                        var_cadena = var_cadena + "tel=" + Chr(13)
                        var_cadena = var_cadena + "email=" + IIf(IsNull(rs!vcha_cli_email), "", rs!vcha_cli_email) + Chr(13)
                        var_cadena = var_cadena + "</Cliente>" + Chr(13) + Chr(13)
                                      
                        var_cadena = var_cadena + "<EntregarEn>" + Chr(13)
                        var_cadena = var_cadena + "endomicilio=" + Chr(13)
                        var_cadena = var_cadena + "encalle=" + Chr(13)
                        var_cadena = var_cadena + "ennoExterior=" + Chr(13)
                        var_cadena = var_cadena + "ennoInterior=" + Chr(13)
                        var_cadena = var_cadena + "encolonia=" + Chr(13)
                        var_cadena = var_cadena + "enlocalidad=" + Chr(13)
                        var_cadena = var_cadena + "enreferencia=" + Chr(13)
                        var_cadena = var_cadena + "enmunicipio=" + Chr(13)
                        var_cadena = var_cadena + "enestado=" + Chr(13)
                        var_cadena = var_cadena + "enpais=" + Chr(13)
                        var_cadena = var_cadena + "encodigoPostal=" + Chr(13)
                        var_cadena = var_cadena + "entel=" + Chr(13)
                        var_cadena = var_cadena + "enemail=" + Chr(13)
                        var_cadena = var_cadena + "</EntregarEn>" + Chr(13) + Chr(13)
                        var_cadena = var_cadena + "<Concepto>" + Chr(13)
                        
                        
                        
                        
                        
                        
                        If rsaux3.State = 1 Then
                           rsaux3.Close
                        End If
                        rsaux3.Open "select * from vw_devolucion_nota_credito where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + txt_movimiento + "' and inte_emo_numero = " + CStr(txt_numero), cnn, adOpenDynamic, adLockOptimistic
                        
                        var_i = 1
                        While Not rsaux3.EOF
                              pxx = CStr(var_i)
                              If Len(pxx) = 1 Then
                                 pxx = "0" + pxx
                              End If
                              var_cadena = var_cadena + "p" + pxx + "_cantidad=" + CStr(IIf(IsNull(rsaux3!Cantidad), 0, rsaux3!Cantidad)) + Chr(13)
                              If rsaux4.State = 1 Then
                                 rsaux4.Close
                              End If
                              rsaux4.Open "SELECT dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPA?OL, dbo.TB_UNIDADES.VCHA_UNI_UNIDAD_ID, dbo.TB_UNIDADES.VCHA_UNI_NOMBRE, dbo.TB_Articulos.VCHA_ART_ARTICULO_ID FROM dbo.TB_ARTICULOS LEFT OUTER JOIN dbo.TB_UNIDADES ON dbo.TB_ARTICULOS.VCHA_UNI_UNIDAD_ID = dbo.TB_UNIDADES.VCHA_UNI_UNIDAD_ID WHERE (dbo.TB_ARTICULOS.vcha_art_Articulo_id = '" + Trim(IIf(IsNull(rsaux3!VCHA_ART_ARTICULO_ID), "", rsaux3!VCHA_ART_ARTICULO_ID)) + "')", cnn, adOpenDynamic, adLockOptimistic
                              var_linea = var_factura + " " + rsaux3!VCHA_ART_ARTICULO_ID + " " + Trim(rsaux4!vcha_Art_nombre_espa?ol)
                              var_cadena = var_cadena + "p" + pxx + "_unidad=" + IIf(IsNull(rsaux4!VCHA_UNI_NOMBRE), "", rsaux4!VCHA_UNI_NOMBRE) + Chr(13)
                              var_cadena = var_cadena + "p" + pxx + "_noIdentificacion=" + Chr(13)
                              
                              var_factura = IIf(IsNull(rsaux3!vcha_Ser_Serie_id), "", rsaux3!vcha_Ser_Serie_id) + CStr(rsaux3!inte_fac_factura)
                              rsaux4.Close
                              var_cadena = var_cadena + "p" + pxx + "_descripcion=" + var_linea + Chr(13)
                              'var_importe_str = var_importe_str = Format(((IIf(IsNull(rsaux3!FLOA_dbo_IMPORTE), 0, rsaux3!FLOA_dbo_IMPORTE)) / (1 + (rsaux3!floa_dbo_iva / 100)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio))), "###,###,##0.00")
                              
                              var_precio = IIf(IsNull(rsaux3!floa_cde_precio), 0, rsaux3!floa_cde_precio) / IIf(IsNull(rsaux3!Cantidad), 0, rsaux3!Cantidad)
                              var_descuento_1 = IIf(IsNull(rsaux3!floa_cde_descuento_1), 0, rsaux3!floa_cde_descuento_1)
                              var_descuento_2 = IIf(IsNull(rsaux3!floa_cde_descuento_2), 0, rsaux3!floa_cde_descuento_2)
                              var_descuento_3 = IIf(IsNull(rsaux3!floa_cde_descuento_3), 0, rsaux3!floa_cde_descuento_3)
                              var_tipo_Cambio = IIf(IsNull(rsaux3!floa_dev_tipo_cambio), 1, rsaux3!floa_dev_tipo_cambio)
                              var_precio = var_precio * (1 - (var_descuento_1 / 100))
                              var_precio = var_precio * (1 - (var_descuento_2 / 100))
                              var_precio = var_precio * (1 - (var_descuento_3 / 100))
                              var_precio = var_precio / var_tipo_Cambio
                              var_iva = IIf(IsNull(rsaux3!floa_cde_iva), 0, rsaux3!floa_cde_iva)
                              If var_rfc_cliente = "XAXX010101000" Then
                                 var_precio = var_precio * (1 + (var_iva / 100))
                              End If
                              var_cadena = var_cadena + "p" + pxx + "_valorUnitario=" + Format(CStr(var_precio), "###,###,##0.000000") + Chr(13)
                              var_cadena = var_cadena + "p" + pxx + "_importe=" + Format(CStr(var_precio * IIf(IsNull(rsaux3!Cantidad), 0, rsaux3!Cantidad)), "###,###,##0.000000") + Chr(13)
                              rsaux3.MoveNext
                              var_i = var_i + 1
                        Wend
                        rsaux3.Close
                        
                        
                        
                        var_cadena = var_cadena + "</Concepto>" + Chr(13) + Chr(13)
                        var_cadena = var_cadena + "<Otros>" + Chr(13)
                        var_cadena = var_cadena + "certificado=" + IIf(IsNull(var_certificado), "", var_certificado) + Chr(13)
                        rs.MoveFirst
                        var_cadena = var_cadena + "cant_letra=" + rs!vcha_car_importe_letra + Chr(13)
                        var_cadena = var_cadena + "factoriva=" + CStr(rs!floa_car_porcentaje_iva) + "%" + Chr(13)
                        rsaux1.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                        var_cadena = var_cadena + "moneda=" + IIf(IsNull(rsaux1!vcha_mon_nombre_plural), "", rsaux1!vcha_mon_nombre_plural) + Chr(13)
                        rsaux1.Close
                        var_cadena = var_cadena + "tipodeCambio=" + CStr(IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)) + Chr(13)
                        var_cadena = var_cadena + "pedido=" + Chr(13)
                        var_cadena = var_cadena + "Embarque=" + Chr(13)
                        var_referencia_Bancaria = ""
                        var_cadena = var_cadena + "referenciabancaria=" + Chr(13)
                        var_cadena = var_cadena + "fechaPedido=" + Chr(13)
                        var_cadena = var_cadena + "expedicion=" + Chr(13)
                        var_cadena = var_cadena + "observaciones=" + Chr(13)
                        var_cadena = var_cadena + "conceptoExtra1=" + Chr(13)
                        var_cadena = var_cadena + "montoconceptoExtra1=" + Chr(13)
                        var_cadena = var_cadena + "conceptoExtra2=" + Chr(13)
                        var_cadena = var_cadena + "montoconceptoExtra2=" + Chr(13)
                        var_cadena = var_cadena + "tipoimpresion=2" + Chr(13)
                        
                        rsaux11.Open "select * from vw_clientes where vcha_cli_clave_id = '" + rs!vcha_cli_clave_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux11.EOF Then
                           var_cadena = var_cadena + "agente=" + rsaux11!VCHA_AGE_AGENTE_ID + " " + rsaux11!VCHA_AGE_NOMBRE + Chr(13)
                        End If
                        rsaux11.Close
                        
                        If var_empresa = "02" Or var_empresa = "03" Or var_empresa = "18" Or var_empresa = "17" Or var_empresa = "06" Then
                           var_cadena = var_cadena + "formato=MHNCVTH_V01.dat" + Chr(13)
                        End If
                        If var_empresa = "07" Then
                           var_cadena = var_cadena + "formato=MHNCARE_V01.dat" + Chr(13)
                        End If
                        If var_empresa = "31" Then
                           var_cadena = var_cadena + "formato=MHNCCAN_V01.dat" + Chr(13)
                        End If
                        If var_empresa = "42" Then
                           var_cadena = var_cadena + "formato=MHNCCMA_V01.dat" + Chr(13)
                        End If
                        If var_empresa = "41" Then
                           var_cadena = var_cadena + "formato=MHNCCOP_V01.dat" + Chr(13)
                        End If
                        If var_empresa = "15" Then
                           var_cadena = var_cadena + "formato=MHNCERE_V01.dat" + Chr(13)
                        End If
                        If var_empresa = "33" Then
                           var_cadena = var_cadena + "formato=MHNCMPU_V01.dat" + Chr(13)
                        End If
                        If var_empresa = "34" Then
                           var_cadena = var_cadena + "formato=MHNCMYG_V01.dat" + Chr(13)
                        End If
                        If var_empresa = "16" Then
                           var_cadena = var_cadena + "formato=MHNCMYG_V01.dat" + Chr(13)
                        End If
                        If var_empresa = "36" Then
                           var_cadena = var_cadena + "formato=MHNCSME_V01.dat" + Chr(13)
                        End If
                        If var_empresa = "30" Then
                           var_cadena = var_cadena + "formato=MHNCTUR_V01.dat" + Chr(13)
                        End If
                        If var_empresa = "44" Then
                           var_cadena = var_cadena + "formato=MHNCUTV_V01.dat" + Chr(13)
                        End If
                        If var_empresa = "38" Then
                           var_cadena = var_cadena + "formato=MHNCVIA_V01.dat" + Chr(13)
                        End If
                        If var_empresa = "40" Then
                           var_cadena = var_cadena + "formato=MHNCVIN_V01.dat" + Chr(13)
                        End If
                        If var_empresa = "43" Then
                           var_cadena = var_cadena + "formato=MHNCVOP_V01.dat" + Chr(13)
                        End If
                      
                        var_cadena = var_cadena + "</Otros>" + Chr(13) + Chr(13)
                        var_cadena = var_cadena + "<addenda>" + Chr(13)
                        var_cadena = var_cadena + "</addenda>" + Chr(13) + Chr(13)
                        var_cadena = var_cadena + "</Factura>"
                        Print #1, var_cadena
                        Close #1
                        
                        var_Archivo = App.Path & "\renombra" + var_serie + Trim(Str(rs!inte_Car_numero)) + ".bat"
                        Open (App.Path & "\renombra" + var_serie + Trim(Str(rs!inte_Car_numero)) + ".bat") For Output As #2
                        Print #2, "ren " + var_ruta_documentos_electronicos + "\" + Trim(var_serie) + Trim(Str(rs!inte_Car_numero)) + ".fi " + Trim(var_serie) + Trim(Str(rs!inte_Car_numero)) + ".ff"
                        Close #2
            
                         x = Shell(var_Archivo, vbHide)
                    End If
                    rs.Close
                     
            
            
            
            
            ''''''''''''''' bonificaciones
            rs.Open "select * from tb_encabezado_Cartera where vcha_car_tipo_documento = 'NC' and vcha_ser_serie_id = '" + var_serie + "' and inte_Car_numero = " + Me.txt_de, cnn, adOpenDynamic, adLockOptimistic
            txt_clase = rs!vcha_car_documento
            rs.Close
                        rs.Open "select * from VW_DOCUMENTOS_IMPRESION where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = '" + txt_clase + "' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(Me.txt_de), cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           Open (App.Path & "\renombra" + Trim(var_serie) + Trim(Str(rs!inte_Car_numero)) + ".bat") For Output As #2
                           Print #2, "ren " + var_ruta_documentos_electronicos + "\" + Trim(var_serie) + Trim(Str(rs!inte_Car_numero)) + ".fi " + Trim(var_serie) + Trim(Str(rs!inte_Car_numero)) + ".ff"
                           Close #2
                        
                           Open (var_ruta_documentos_electronicos & "\" + Trim(var_serie) + Trim(Str(rs!inte_Car_numero)) + ".fi") For Output As #1
                           'Open ("c:\NC_" + Trim(var_serie) + Trim(Str(rs!inte_car_numero)) + ".fi") For Output As #1
                           var_cadena = "Outputmode=" + Chr(13) + "<Factura>" + Chr(13) + "<Comprobante>" + Chr(13) + "Version=2.0" + Chr(13) + "Serie=" + rs!vcha_Ser_Serie_id + Chr(13) + "folio=" + CStr(rs!inte_Car_numero) + Chr(13)
                           var_a?o = CStr(Year(rs!dtim_Car_fecha))
                           var_mes = CStr(Month(rs!dtim_Car_fecha))
                           VAR_DIA = CStr(Day(rs!dtim_Car_fecha))
                           var_hora = CStr(Hour(rs!dtim_Car_fecha))
                           var_minuto = CStr(Minute(rs!dtim_Car_fecha))
                           var_segundo = CStr(Second(rs!dtim_Car_fecha))
                           If Len(var_a?o) = 2 Then
                              var_a?o = "20" + var_a?o
                           End If
                           If Len(var_mes) = 1 Then
                              var_mes = "0" + var_mes
                           End If
                           If Len(VAR_DIA) = 1 Then
                              VAR_DIA = "0" + VAR_DIA
                           End If
                           If Len(var_hora) = 1 Then
                              var_hora = "0" + var_hora
                           End If
                           If Len(var_minuto) = 1 Then
                              var_minuto = "0" + var_minuto
                           End If
                           If Len(var_segundo) = 1 Then
                              var_segundo = "0" + var_segundo
                           End If
                           
                           var_rfc_cliente_1 = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                           var_rfc_cliente = ""
                           If var_rfc_cliente_1 = "" Then
                              var_rfc_cliente = "XAXX010101000"
                           Else
                              For var_j = 1 To Len(var_rfc_cliente_1)
                                  If Mid(var_rfc_cliente_1, var_j, 1) <> "-" Then
                                     If Mid(var_rfc_cliente_1, var_j, 1) <> "" Then
                                        If Mid(var_rfc_cliente_1, var_j, 1) <> " " Then
                                           var_rfc_cliente = var_rfc_cliente + Mid(var_rfc_cliente_1, var_j, 1)
                                        End If
                                     End If
                                  End If
                              Next var_j
                           End If
                           
                           
                           var_cadena_fecha = var_a?o + "-" + var_mes + "-" + VAR_DIA + "T" + var_hora + ":" + var_minuto + ":" + var_segundo
                           var_cadena = var_cadena + "fecha=" + var_cadena_fecha + Chr(13)
                           var_cadena = var_cadena + "noAprobacion=" + Chr(13)
                           var_cadena = var_cadena + "anoAprobacion=" + Chr(13)
                           var_cadena = var_cadena + "tipoDeComprobante=NOTA DE CREDITO" + Chr(13)
                           var_cadena = var_cadena + "formaDePago=PAGO HECHO EN UNA SOLA EXHIBICION" + Chr(13)
                           var_cadena = var_cadena + "condicionesDePago=" + Chr(13)
                           If var_rfc_cliente = "XAXX010101000" Then
                              var_cadena = var_cadena + "subtotal=" + Format(CStr(rs!floa_Car_importe_neto / rs!floa_car_tipo_cambio), "###,###,##0.000000") + Chr(13)
                           Else
                              var_cadena = var_cadena + "subtotal=" + Format(CStr(rs!floa_car_subimporte / rs!floa_car_tipo_cambio), "###,###,##0.000000") + Chr(13)
                           End If
                           var_cadena = var_cadena + "descuento=" + Chr(13)
                           var_cadena = var_cadena + "descuento1=" + Chr(13)
                           var_cadena = var_cadena + "descuento2=" + Chr(13)
                           var_cadena = var_cadena + "conceptodescuento1=" + Chr(13)
                           var_cadena = var_cadena + "conceptodescuento2=" + Chr(13)
                           var_cadena = var_cadena + "tasadescuento1=" + Chr(13)
                           var_cadena = var_cadena + "tasadescuento2=" + Chr(13)
                           If rsaux2.State = 1 Then
                              rsaux2.Close
                           End If
                           rsaux2.Open "select * from tb_empresa_FACTURA_ELECTRONICA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
                           var_certificado = rsaux2!vcha_emp_certificado
                           var_expedido = rsaux2!vcha_emp_expedido
                           If var_rfc_cliente = "XAXX010101000" Then
                              var_cadena = var_cadena + "iva=" + Format(CStr(0), "###,###,##0.000000") + Chr(13)
                           Else
                              var_cadena = var_cadena + "iva=" + Format(CStr(rs!floa_car_importe_iva / rs!floa_car_tipo_cambio), "###,###,##0.000000") + Chr(13)
                           End If
                           var_cadena = var_cadena + "total=" + Format(CStr(rs!floa_Car_importe_neto / rs!floa_car_tipo_cambio), "###,###,##0.000000") + Chr(13)
                           var_cadena = var_cadena + "retencion=" + Chr(13)
                           var_cadena = var_cadena + "factorretencioniva=" + Chr(13)
                           var_cadena = var_cadena + "</Comprobante>" + Chr(13) + Chr(13)
                           var_cadena = var_cadena + "<Emisor>" + Chr(13)
                           var_cadena = var_cadena + "erfc=" + rsaux2!VCHA_eMP_RFC + Chr(13)
                           var_cadena = var_cadena + "enombre=" + rsaux2!VCHA_EMP_NOMBRE + Chr(13)
                           var_cadena = var_cadena + "</Emisor>" + Chr(13) + Chr(13)
                           var_cadena = var_cadena + "<DomicilioFiscal>" + Chr(13)
                           var_cadena = var_cadena + "ecalle=" + rsaux2!VCHA_eMP_CALLE + Chr(13)
                           var_cadena = var_cadena + "enoExterior=" + rsaux2!VCHA_eMP_exterior + Chr(13)
                           var_cadena = var_cadena + "enoInterior=" + Chr(13)
                           var_cadena = var_cadena + "ecolonia=" + rsaux2!VCHA_eMP_COLONIA + Chr(13)
                           var_cadena = var_cadena + "elocalidad=" + rsaux2!VCHA_EMP_LOCALIDAD + Chr(13)
                           var_cadena = var_cadena + "ereferencia=" + Chr(13)
                           var_cadena = var_cadena + "emunicipio=" + rsaux2!VCHA_EMP_MUNICIPIO + Chr(13)
                           var_cadena = var_cadena + "eestado=" + rsaux2!VCHA_EMP_ESTADO + Chr(13)
                           var_cadena = var_cadena + "epais=" + rsaux2!VCHA_eMP_PAIS + Chr(13)
                           var_cadena = var_cadena + "ecodigoPostal=" + rsaux2!VCHA_EMP_CODIGO_POSTAL + Chr(13)
                           var_cadena = var_cadena + "etel=" + IIf(IsNull(rsaux2!VCHA_EMP_TELEFONO), "", rsaux2!VCHA_EMP_TELEFONO) + Chr(13)
                           var_cadena = var_cadena + "eemail=" + IIf(IsNull(rsaux2!VCHA_EMP_EMAIL), "", rsaux2!VCHA_EMP_EMAIL) + Chr(13)
                           var_cadena = var_cadena + "</DomicilioFiscal>" + Chr(13) + Chr(13)
                           var_cadena = var_cadena + "<Receptor>" + Chr(13)
                           var_cadena = var_cadena + "noCliente=" + rs!vcha_cli_clave_id + Chr(13)
                           rsaux2.Close
                                         
                           var_rfc_cliente_1 = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                           var_rfc_cliente = ""
                           If var_rfc_cliente_1 = "" Then
                              var_rfc_cliente = "XAXX010101000"
                           Else
                              For var_j = 1 To Len(var_rfc_cliente_1)
                                  If Mid(var_rfc_cliente_1, var_j, 1) <> "-" Then
                                     If Mid(var_rfc_cliente_1, var_j, 1) <> "" Then
                                        If Mid(var_rfc_cliente_1, var_j, 1) <> " " Then
                                           var_rfc_cliente = var_rfc_cliente + Mid(var_rfc_cliente_1, var_j, 1)
                                        End If
                                     End If
                                  End If
                              Next var_j
                           End If
                           If var_empresa = "03" Or var_empresa = "28" Then
                              var_rfc_cliente = "XEXX010101000"
                           End If
                           var_cadena = var_cadena + "rfc=" + var_rfc_cliente + Chr(13)
                           var_cadena = var_cadena + "nombre=" + rs!VCHA_CLI_NOMBRE + Chr(13)
                           var_cadena = var_cadena + "</Receptor>" + Chr(13) + Chr(13)
                           var_cadena = var_cadena + "<Cliente>" + Chr(13)
                           var_cadena = var_cadena + "domicilio=" + IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + Chr(13)
                           var_cadena = var_cadena + "calle=" + Chr(13)
                           var_cadena = var_cadena + "noExterior=" + Chr(13)
                           var_cadena = var_cadena + "noInterior=" + Chr(13)
                           var_cadena = var_cadena + "colonia=" + IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre) + Chr(13)
                           var_cadena = var_cadena + "localidad=" + IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre) + Chr(13)
                           rsaux2.Open "select * from vw_clientes where vcha_Cli_clave_id = '" + rs!vcha_cli_clave_id + "'"
                           var_cadena = var_cadena + "referencia=" + Chr(13)
                           var_cadena = var_cadena + "municipio=" + IIf(IsNull(rsaux2!vcha_mun_nombre), "", rsaux2!vcha_mun_nombre) + Chr(13)
                           var_cadena = var_cadena + "estado=" + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre) + Chr(13)
                           VAR_NOMBRE_PAIS = IIf(IsNull(rs!vcha_pai_nombre), "MEXICO", rs!vcha_pai_nombre)
                           If Trim(VAR_NOMBRE_PAIS) = "" Then
                              VAR_NOMBRE_PAIS = "MEXICO"
                           End If
                           var_cadena = var_cadena + "pais=" + VAR_NOMBRE_PAIS + Chr(13)
                           var_cadena = var_cadena + Chr(13)
                           var_cadena = var_cadena + "codigoPostal=" + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP) + Chr(13)
                           var_cadena = var_cadena + "tel=" + Chr(13)
                           var_cadena = var_cadena + "email=" + IIf(IsNull(rs!vcha_cli_email), "", rs!vcha_cli_email) + Chr(13)
                           var_cadena = var_cadena + "</Cliente>" + Chr(13) + Chr(13)
                                         
                           var_cadena = var_cadena + "<EntregarEn>" + Chr(13)
                           var_cadena = var_cadena + "endomicilio=" + Chr(13)
                           var_cadena = var_cadena + "encalle=" + Chr(13)
                           var_cadena = var_cadena + "ennoExterior=" + Chr(13)
                           var_cadena = var_cadena + "ennoInterior=" + Chr(13)
                           var_cadena = var_cadena + "encolonia=" + Chr(13)
                           var_cadena = var_cadena + "enlocalidad=" + Chr(13)
                           var_cadena = var_cadena + "enreferencia=" + Chr(13)
                           var_cadena = var_cadena + "enmunicipio=" + Chr(13)
                           var_cadena = var_cadena + "enestado=" + Chr(13)
                           var_cadena = var_cadena + "enpais=" + Chr(13)
                           var_cadena = var_cadena + "encodigoPostal=" + Chr(13)
                           var_cadena = var_cadena + "entel=" + Chr(13)
                           var_cadena = var_cadena + "enemail=" + Chr(13)
                           var_cadena = var_cadena + "</EntregarEn>" + Chr(13) + Chr(13)
                           var_cadena = var_cadena + "<Concepto>" + Chr(13)
                           
                           
                           rsaux3.Open "select * from tb_detalle_bonificaciones where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = '" + txt_clase + "' and vcha_ser_serie_id = '" + var_serie + "' and inte_car_numero = " + CStr(Me.txt_de), cnn, adOpenDynamic, adLockOptimistic
                           var_i = 1
                           While Not rsaux3.EOF
                                 pxx = CStr(var_i)
                                 If Len(pxx) = 1 Then
                                    pxx = "0" + pxx
                                 End If
                                 var_cadena = var_cadena + "p" + pxx + "_cantidad=1" + Chr(13)
                                 var_cadena = var_cadena + "p" + pxx + "_unidad=" + txt_clase + Chr(13)
                                 var_cadena = var_cadena + "p" + pxx + "_noIdentificacion=" + Chr(13)
                                 var_linea = txt_clase + Str(rs!inte_Car_numero) + " " + rs!vcha_Car_nombre + " FACTURA " + CStr(rsaux3!inte_car_factura)
                                 var_cadena = var_cadena + "p" + pxx + "_descripcion=" + var_linea + Chr(13)
                                 If var_rfc_cliente = "XAXX010101000" Then
                                    var_importe_str = ((IIf(IsNull(rsaux3!FLOA_dbo_IMPORTE), 0, rsaux3!FLOA_dbo_IMPORTE)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)))
                                 Else
                                    var_importe_str = ((IIf(IsNull(rsaux3!FLOA_dbo_IMPORTE), 0, rsaux3!FLOA_dbo_IMPORTE)) / (1 + (rsaux3!floa_dbo_iva / 100)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)))
                                 End If
                                 var_cadena = var_cadena + "p" + pxx + "_valorUnitario=" + Format(CStr(var_importe_str), "###,###,##0.000000") + Chr(13)
                                 var_cadena = var_cadena + "p" + pxx + "_importe=" + Format(CStr(var_importe_str), "###,###,##0.000000") + Chr(13)
                                 rsaux3.MoveNext
                                 var_i = var_i + 1
                           Wend
                           rsaux3.Close
                           'MsgBox var_cadena
                           var_cadena = var_cadena + "</Concepto>" + Chr(13) + Chr(13)
                           var_cadena = var_cadena + "<Otros>" + Chr(13)
                           var_cadena = var_cadena + "certificado=" + IIf(IsNull(var_certificado), "", var_certificado) + Chr(13)
                           rs.MoveFirst
                           var_cadena = var_cadena + "cant_letra=" + rs!vcha_car_importe_letra + Chr(13)
                           var_cadena = var_cadena + "factoriva=" + CStr(rs!floa_car_porcentaje_iva) + "%" + Chr(13)
                           rsaux1.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                           var_cadena = var_cadena + "moneda=" + IIf(IsNull(rsaux1!vcha_mon_nombre_plural), "", rsaux1!vcha_mon_nombre_plural) + Chr(13)
                           rsaux1.Close
                           var_cadena = var_cadena + "tipodeCambio=" + CStr(IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)) + Chr(13)
                           var_cadena = var_cadena + "pedido=" + Chr(13)
                           var_cadena = var_cadena + "Embarque=" + Chr(13)
                           var_referencia_Bancaria = ""
                           var_cadena = var_cadena + "referenciabancaria=" + Chr(13)
                           var_cadena = var_cadena + "fechaPedido=" + Chr(13)
                           var_cadena = var_cadena + "expedicion=" + Chr(13)
                           var_cadena = var_cadena + "observaciones=" + Chr(13)
                           var_cadena = var_cadena + "conceptoExtra1=" + Chr(13)
                           var_cadena = var_cadena + "montoconceptoExtra1=" + Chr(13)
                           var_cadena = var_cadena + "conceptoExtra2=" + Chr(13)
                           var_cadena = var_cadena + "montoconceptoExtra2=" + Chr(13)
                           var_cadena = var_cadena + "tipoimpresion=2" + Chr(13)
                           
                           rsaux11.Open "select * from vw_clientes where vcha_cli_clave_id = '" + rs!vcha_cli_clave_id + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux11.EOF Then
                              var_cadena = var_cadena + "agente=" + rsaux11!VCHA_AGE_AGENTE_ID + " " + rsaux11!VCHA_AGE_NOMBRE + Chr(13)
                           End If
                           rsaux11.Close
                           
                           If var_empresa = "02" Or var_empresa = "03" Or var_empresa = "18" Or var_empresa = "17" Or var_empresa = "06" Then
                              var_cadena = var_cadena + "formato=MHNCVTH_V01.dat" + Chr(13)
                           End If
                           If var_empresa = "07" Then
                              var_cadena = var_cadena + "formato=MHNCARE_V01.dat" + Chr(13)
                           End If
                           If var_empresa = "31" Then
                              var_cadena = var_cadena + "formato=MHNCCAN_V01.dat" + Chr(13)
                           End If
                           If var_empresa = "42" Then
                              var_cadena = var_cadena + "formato=MHNCCMA_V01.dat" + Chr(13)
                           End If
                           If var_empresa = "41" Then
                              var_cadena = var_cadena + "formato=MHNCCOP_V01.dat" + Chr(13)
                           End If
                           If var_empresa = "15" Then
                              var_cadena = var_cadena + "formato=MHNCERE_V01.dat" + Chr(13)
                           End If
                           If var_empresa = "33" Then
                              var_cadena = var_cadena + "formato=MHNCMPU_V01.dat" + Chr(13)
                           End If
                           If var_empresa = "34" Then
                              var_cadena = var_cadena + "formato=MHNCMYG_V01.dat" + Chr(13)
                           End If
                           If var_empresa = "16" Then
                              var_cadena = var_cadena + "formato=MHNCMYG_V01.dat" + Chr(13)
                           End If
                           If var_empresa = "36" Then
                              var_cadena = var_cadena + "formato=MHNCSME_V01.dat" + Chr(13)
                           End If
                           If var_empresa = "30" Then
                              var_cadena = var_cadena + "formato=MHNCTUR_V01.dat" + Chr(13)
                           End If
                           If var_empresa = "44" Then
                              var_cadena = var_cadena + "formato=MHNCUTV_V01.dat" + Chr(13)
                           End If
                           If var_empresa = "38" Then
                              var_cadena = var_cadena + "formato=MHNCVIA_V01.dat" + Chr(13)
                           End If
                           If var_empresa = "40" Then
                              var_cadena = var_cadena + "formato=MHNCVIN_V01.dat" + Chr(13)
                           End If
                           If var_empresa = "43" Then
                              var_cadena = var_cadena + "formato=MHNCVOP_V01.dat" + Chr(13)
                           End If
                           
                           
                           var_cadena = var_cadena + "</Otros>" + Chr(13) + Chr(13)
                           var_cadena = var_cadena + "<addenda>" + Chr(13)
                           var_cadena = var_cadena + "</addenda>" + Chr(13) + Chr(13)
                           var_cadena = var_cadena + "</Factura>"
                           Print #1, var_cadena
                           Close #1
                        End If
                        rs.Close
                           var_Archivo = App.Path & "\renombra" + Trim(var_serie) + Trim(Str(Me.txt_de)) + ".bat"
                           x = Shell(var_Archivo, vbHide)
            
            Exit Sub
            
            
            
            
            
            ''''''''''''''''descuento financiero
            rs.Open "select * from VW_DOCUMENTOS_IMPRESION where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'DF' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(Me.txt_de), cnn, adOpenDynamic, adLockOptimistic
            Open (var_ruta_documentos_electronicos & "\" + Trim(var_serie) + Trim(Str(Me.txt_de)) + ".fi") For Output As #1
            
            
            var_rfc_cliente_1 = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
            var_rfc_cliente = ""
            If var_rfc_cliente_1 = "" Then
               var_rfc_cliente = "XAXX010101000"
            Else
               For var_j = 1 To Len(var_rfc_cliente_1)
                   If Mid(var_rfc_cliente_1, var_j, 1) <> "-" Then
                      If Mid(var_rfc_cliente_1, var_j, 1) <> "" Then
                         If Mid(var_rfc_cliente_1, var_j, 1) <> " " Then
                            var_rfc_cliente = var_rfc_cliente + Mid(var_rfc_cliente_1, var_j, 1)
                         End If
                      End If
                   End If
               Next var_j
            End If
            
            
            
            var_cadena = "Outputmode=" + Chr(13) + "<Factura>" + Chr(13) + "<Comprobante>" + Chr(13) + "Version=2.0" + Chr(13) + "Serie=" + rs!vcha_Ser_Serie_id + Chr(13) + "folio=" + CStr(rs!inte_Car_numero) + Chr(13)
            var_a?o = CStr(Year(rs!dtim_Car_fecha))
            var_mes = CStr(Month(rs!dtim_Car_fecha))
            VAR_DIA = CStr(Day(rs!dtim_Car_fecha))
            var_hora = CStr(Hour(rs!dtim_Car_fecha))
            var_minuto = CStr(Minute(rs!dtim_Car_fecha))
            var_segundo = CStr(Second(rs!dtim_Car_fecha))
            If Len(var_a?o) = 2 Then
               var_a?o = "20" + var_a?o
            End If
            If Len(var_mes) = 1 Then
               var_mes = "0" + var_mes
            End If
            If Len(VAR_DIA) = 1 Then
               VAR_DIA = "0" + VAR_DIA
            End If
            If Len(var_hora) = 1 Then
               var_hora = "0" + var_hora
            End If
            If Len(var_minuto) = 1 Then
               var_minuto = "0" + var_minuto
            End If
            If Len(var_segundo) = 1 Then
               var_segundo = "0" + var_segundo
            End If
            var_rfc_cliente_1 = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
            var_rfc_cliente = ""
            If var_rfc_cliente_1 = "" Then
               var_rfc_cliente = "XAXX010101000"
            Else
               For var_j = 1 To Len(var_rfc_cliente_1)
                   If Mid(var_rfc_cliente_1, var_j, 1) <> "-" Then
                      If Mid(var_rfc_cliente_1, var_j, 1) <> "" Then
                         If Mid(var_rfc_cliente_1, var_j, 1) <> " " Then
                            var_rfc_cliente = var_rfc_cliente + Mid(var_rfc_cliente_1, var_j, 1)
                         End If
                      End If
                   End If
               Next var_j
            End If
            
            var_cadena_fecha = var_a?o + "-" + var_mes + "-" + VAR_DIA + "T" + var_hora + ":" + var_minuto + ":" + var_segundo
            var_cadena = var_cadena + "fecha=" + var_cadena_fecha + Chr(13)
            var_cadena = var_cadena + "noAprobacion=" + Chr(13)
            var_cadena = var_cadena + "anoAprobacion=" + Chr(13)
            var_cadena = var_cadena + "tipoDeComprobante=NOTA DE CREDITO" + Chr(13)
            var_cadena = var_cadena + "formaDePago=PAGO HECHO EN UNA SOLA EXHIBICION" + Chr(13)
            var_cadena = var_cadena + "condicionesDePago=" + Chr(13)
            If var_rfc_cliente = "XAXX010101000" Then
               var_cadena = var_cadena + "subtotal=" + Format(CStr(rs!floa_Car_importe_neto / rs!floa_car_tipo_cambio), "###,###,##0.000000") + Chr(13)
            Else
               var_cadena = var_cadena + "subtotal=" + Format(CStr(rs!floa_car_subimporte / rs!floa_car_tipo_cambio), "###,###,##0.000000") + Chr(13)
            End If
            var_cadena = var_cadena + "descuento=" + Chr(13)
            var_cadena = var_cadena + "descuento1=" + Chr(13)
            var_cadena = var_cadena + "descuento2=" + Chr(13)
            var_cadena = var_cadena + "conceptodescuento1=" + Chr(13)
            var_cadena = var_cadena + "conceptodescuento2=" + Chr(13)
            var_cadena = var_cadena + "tasadescuento1=" + Chr(13)
            var_cadena = var_cadena + "tasadescuento2=" + Chr(13)
            If rsaux2.State = 1 Then
               rsaux2.Close
            End If
            rsaux2.Open "select * from tb_empresa_FACTURA_ELECTRONICA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
            var_certificado = rsaux2!vcha_emp_certificado
            var_expedido = rsaux2!vcha_emp_expedido
            If var_rfc_cliente = "XAXX010101000" Then
               var_cadena = var_cadena + "iva=" + Format(0, "###,###,##0.000000") + Chr(13)
            Else
               var_cadena = var_cadena + "iva=" + Format(CStr(rs!floa_car_importe_iva / rs!floa_car_tipo_cambio), "###,###,##0.000000") + Chr(13)
            End If
            var_cadena = var_cadena + "total=" + Format(CStr(rs!floa_Car_importe_neto / rs!floa_car_tipo_cambio), "###,###,##0.000000") + Chr(13)
            var_cadena = var_cadena + "retencion=" + Chr(13)
            var_cadena = var_cadena + "factorretencioniva=" + Chr(13)
            var_cadena = var_cadena + "</Comprobante>" + Chr(13) + Chr(13)
            var_cadena = var_cadena + "<Emisor>" + Chr(13)
            var_cadena = var_cadena + "erfc=" + rsaux2!VCHA_eMP_RFC + Chr(13)
            var_cadena = var_cadena + "enombre=" + rsaux2!VCHA_EMP_NOMBRE + Chr(13)
            var_cadena = var_cadena + "</Emisor>" + Chr(13) + Chr(13)
            var_cadena = var_cadena + "<DomicilioFiscal>" + Chr(13)
            var_cadena = var_cadena + "ecalle=" + rsaux2!VCHA_eMP_CALLE + Chr(13)
            var_cadena = var_cadena + "enoExterior=" + rsaux2!VCHA_eMP_exterior + Chr(13)
            var_cadena = var_cadena + "enoInterior=" + Chr(13)
            var_cadena = var_cadena + "ecolonia=" + rsaux2!VCHA_eMP_COLONIA + Chr(13)
            var_cadena = var_cadena + "elocalidad=" + rsaux2!VCHA_EMP_LOCALIDAD + Chr(13)
            var_cadena = var_cadena + "ereferencia=" + Chr(13)
            var_cadena = var_cadena + "emunicipio=" + rsaux2!VCHA_EMP_MUNICIPIO + Chr(13)
            var_cadena = var_cadena + "eestado=" + rsaux2!VCHA_EMP_ESTADO + Chr(13)
            var_cadena = var_cadena + "epais=" + rsaux2!VCHA_eMP_PAIS + Chr(13)
            var_cadena = var_cadena + "ecodigoPostal=" + rsaux2!VCHA_EMP_CODIGO_POSTAL + Chr(13)
            var_cadena = var_cadena + "etel=" + IIf(IsNull(rsaux2!VCHA_EMP_TELEFONO), "", rsaux2!VCHA_EMP_TELEFONO) + Chr(13)
            var_cadena = var_cadena + "eemail=" + IIf(IsNull(rsaux2!VCHA_EMP_EMAIL), "", rsaux2!VCHA_EMP_EMAIL) + Chr(13)
            var_cadena = var_cadena + "</DomicilioFiscal>" + Chr(13) + Chr(13)
            var_cadena = var_cadena + "<Receptor>" + Chr(13)
            var_cadena = var_cadena + "noCliente=" + rs!vcha_cli_clave_id + Chr(13)
            rsaux2.Close
                                         
            var_rfc_cliente_1 = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
            var_rfc_cliente = ""
            If var_rfc_cliente_1 = "" Then
               var_rfc_cliente = "XAXX010101000"
            Else
               For var_j = 1 To Len(var_rfc_cliente_1)
                   If Mid(var_rfc_cliente_1, var_j, 1) <> "-" Then
                      If Mid(var_rfc_cliente_1, var_j, 1) <> "" Then
                         If Mid(var_rfc_cliente_1, var_j, 1) <> " " Then
                            var_rfc_cliente = var_rfc_cliente + Mid(var_rfc_cliente_1, var_j, 1)
                         End If
                      End If
                   End If
               Next var_j
            End If
            If var_empresa = "03" Or var_empresa = "28" Then
               var_rfc_cliente = "XEXX010101000"
            End If
            var_cadena = var_cadena + "rfc=" + var_rfc_cliente + Chr(13)
            var_cadena = var_cadena + "nombre=" + rs!VCHA_CLI_NOMBRE + Chr(13)
            var_cadena = var_cadena + "</Receptor>" + Chr(13) + Chr(13)
            var_cadena = var_cadena + "<Cliente>" + Chr(13)
            var_cadena = var_cadena + "domicilio=" + IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + Chr(13)
            var_cadena = var_cadena + "calle=" + Chr(13)
            var_cadena = var_cadena + "noExterior=" + Chr(13)
            var_cadena = var_cadena + "noInterior=" + Chr(13)
            var_cadena = var_cadena + "colonia=" + IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre) + Chr(13)
            var_cadena = var_cadena + "localidad=" + IIf(IsNull(rs!vcha_ciu_nombre), "", rs!VCHA_CLI_NOMBRE) + Chr(13)
            rsaux2.Open "select * from vw_clientes where vcha_Cli_clave_id = '" + rs!vcha_cli_clave_id + "'"
            var_cadena = var_cadena + "referencia=" + Chr(13)
            var_cadena = var_cadena + "municipio=" + IIf(IsNull(rsaux2!vcha_mun_nombre), "", rsaux2!vcha_mun_nombre) + Chr(13)
            var_cadena = var_cadena + "estado=" + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre) + Chr(13)
            VAR_NOMBRE_PAIS = IIf(IsNull(rs!vcha_pai_nombre), "MEXICO", rs!vcha_pai_nombre)
            If Trim(VAR_NOMBRE_PAIS) = "" Then
               VAR_NOMBRE_PAIS = "MEXICO"
            End If
            var_cadena = var_cadena + "pais=" + VAR_NOMBRE_PAIS + Chr(13)
            var_cadena = var_cadena + Chr(13)
            var_cadena = var_cadena + "codigoPostal=" + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP) + Chr(13)
            var_cadena = var_cadena + "tel=" + Chr(13)
            var_cadena = var_cadena + "email=" + IIf(IsNull(rs!vcha_cli_email), "", rs!vcha_cli_email) + Chr(13)
            var_cadena = var_cadena + "</Cliente>" + Chr(13) + Chr(13)
                                         
            var_cadena = var_cadena + "<EntregarEn>" + Chr(13)
            var_cadena = var_cadena + "endomicilio=" + Chr(13)
            var_cadena = var_cadena + "encalle=" + Chr(13)
            var_cadena = var_cadena + "ennoExterior=" + Chr(13)
            var_cadena = var_cadena + "ennoInterior=" + Chr(13)
            var_cadena = var_cadena + "encolonia=" + Chr(13)
            var_cadena = var_cadena + "enlocalidad=" + Chr(13)
            var_cadena = var_cadena + "enreferencia=" + Chr(13)
            var_cadena = var_cadena + "enmunicipio=" + Chr(13)
            var_cadena = var_cadena + "enestado=" + Chr(13)
            var_cadena = var_cadena + "enpais=" + Chr(13)
            var_cadena = var_cadena + "encodigoPostal=" + Chr(13)
            var_cadena = var_cadena + "entel=" + Chr(13)
            var_cadena = var_cadena + "enemail=" + Chr(13)
            var_cadena = var_cadena + "</EntregarEn>" + Chr(13) + Chr(13)
            var_cadena = var_cadena + "<Concepto>" + Chr(13)
            
            
            var_i = 1
            rsaux3.Open "select * from TB_DETALLE_DESCUENTOS_FINANCIEROS where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'DF' and vcha_ser_serie_id  = '" + var_serie + "' and vcha_car_clase_id = 'DF' and inte_car_numero =  " + Str(Me.txt_de), cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux3.EOF
                  pxx = CStr(var_i)
                  If Len(pxx) = 1 Then
                     pxx = "0" + pxx
                  End If
                  var_cadena = var_cadena + "p" + pxx + "_cantidad=1" + Chr(13)
                  var_cadena = var_cadena + "p" + pxx + "_unidad=DF" + Chr(13)
                  var_cadena = var_cadena + "p" + pxx + "_noIdentificacion=" + Chr(13)
                  var_linea = "DF" + Str(rs!inte_Car_numero) + " " + rs!vcha_Car_nombre + " Factura " + Str(rsaux3!inte_ddf_factura)
                  'If Round(rsaux3!floa_ddf_descuento_otorgado, 2) <> Round(rsaux3!floa_ddf_descuento_aplicado, 2) Then
                  '   var_linea = var_linea + " (" + Format(rsaux3!floa_ddf_descuento_aplicado, "###,###,##0.0000") + "%)"
                  'End If
                  var_cadena = var_cadena + "p" + pxx + "_descripcion=" + var_linea + Chr(13)
                  If var_rfc_cliente = "XAXX010101000" Then
                     var_importe_str = IIf(IsNull(rsaux3!FLOA_DDF_IMPORTE), 0, rsaux3!FLOA_DDF_IMPORTE)
                  Else
                     var_importe_str = IIf(IsNull(rsaux3!FLOA_DDF_IMPORTE), 0, rsaux3!FLOA_DDF_IMPORTE) / (1 + (rs!floa_car_porcentaje_iva / 100))
                  End If
                     
                  var_cadena = var_cadena + "p" + pxx + "_valorUnitario=" + Format(CStr(var_importe_str), "###,###,##0.000000") + Chr(13)
                  var_cadena = var_cadena + "p" + pxx + "_importe=" + Format(CStr(var_importe_str), "###,###,##0.000000") + Chr(13)
                  rsaux3.MoveNext
                  var_i = var_i + 1
            Wend
            rsaux3.Close
            
            
            
            
            
            var_cadena = var_cadena + "</Concepto>" + Chr(13) + Chr(13)
            var_cadena = var_cadena + "<Otros>" + Chr(13)
            var_cadena = var_cadena + "certificado=" + IIf(IsNull(var_certificado), "", var_certificado) + Chr(13)
            rs.MoveFirst
            var_cadena = var_cadena + "cant_letra=" + rs!vcha_car_importe_letra + Chr(13)
            var_cadena = var_cadena + "factoriva=" + CStr(rs!floa_car_porcentaje_iva) + "%" + Chr(13)
            rsaux1.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id) + "'", cnn, adOpenDynamic, adLockOptimistic
            var_cadena = var_cadena + "moneda=" + IIf(IsNull(rsaux1!vcha_mon_nombre_plural), "", rsaux1!vcha_mon_nombre_plural) + Chr(13)
            rsaux1.Close
            var_cadena = var_cadena + "tipodeCambio=" + CStr(IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)) + Chr(13)
            var_cadena = var_cadena + "pedido=" + Chr(13)
            var_cadena = var_cadena + "Embarque=" + Chr(13)
            var_referencia_Bancaria = ""
            var_cadena = var_cadena + "referenciabancaria=" + Chr(13)
            var_cadena = var_cadena + "fechaPedido=" + Chr(13)
            var_cadena = var_cadena + "expedicion=" + Chr(13)
            var_cadena = var_cadena + "observaciones=" + Chr(13)
            var_cadena = var_cadena + "conceptoExtra1=" + Chr(13)
            var_cadena = var_cadena + "montoconceptoExtra1=" + Chr(13)
            var_cadena = var_cadena + "conceptoExtra2=" + Chr(13)
            var_cadena = var_cadena + "montoconceptoExtra2=" + Chr(13)
            var_cadena = var_cadena + "tipoimpresion=2" + Chr(13)
            If var_empresa = "15" Then
               var_cadena = var_cadena + "formato=MHESTAMPADOS_V01.DAT" + Chr(13)
            End If
            If var_empresa = "16" Then
               var_cadena = var_cadena + "formato=MHMULTIBONDEADOS_V01.DAT" + Chr(13)
            End If
            If var_empresa = "02" Or var_empresa = "03" Or var_empresa = "18" Or var_empresa = "17" Or var_empresa = "06" Then
               var_cadena = var_cadena + "formato=MHVTH_V01.DAT" + Chr(13)
            End If
            
            var_cadena = var_cadena + "</Otros>" + Chr(13) + Chr(13)
            var_cadena = var_cadena + "<addenda>" + Chr(13)
            var_cadena = var_cadena + "</addenda>" + Chr(13) + Chr(13)
            var_cadena = var_cadena + "</Factura>"
            Print #1, var_cadena
            Close #1
            var_Archivo = App.Path & "\renombra" + var_serie + Trim(Str(rs!inte_Car_numero)) + ".bat"
            Open (App.Path & "\renombra" + var_serie + Trim(Str(rs!inte_Car_numero)) + ".bat") For Output As #2
            Print #2, "ren " + var_ruta_documentos_electronicos + "\" + Trim(var_serie) + Trim(Str(rs!inte_Car_numero)) + ".fi " + Trim(var_serie) + Trim(Str(rs!inte_Car_numero)) + ".ff"
            Close #2
            
            x = Shell(var_Archivo, vbHide)
             
            rs.Close
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 3500
   If var_empresa = "15" Then
      var_conexion = "Provider=SQLOLEDB.1;Password=" & parametros(3) & ";Persist Security Info=True;User ID=sa;Initial Catalog=sid;Data Source=admcdindustrial"
   End If
   If var_empresa = "16" Then
      var_conexion = "Provider=SQLOLEDB.1;Password=" & parametros(3) & ";Persist Security Info=True;User ID=sa;Initial Catalog=sid;Data Source=admcdindustrial"
   End If
   If var_empresa = "30" Then
      var_conexion = "Provider=SQLOLEDB.1;Password=" & parametros(3) & ";Persist Security Info=True;User ID=sa;Initial Catalog=vianney;Data Source=DISTRIBUCION"
   End If
   If var_empresa = "31" Then
      var_conexion = "Provider=SQLOLEDB.1;Password=" & parametros(3) & ";Persist Security Info=True;User ID=sa;Initial Catalog=sidcantia;Data Source=sqlquezada2"
   End If
   If var_empresa = "38" Then
      var_conexion = "Provider=SQLOLEDB.1;Password=" & parametros(3) & ";Persist Security Info=True;User ID=sa;Initial Catalog=vianney;Data Source=DISTRIBUCION"
   End If
   If var_empresa = "02" Or var_empresa = "03" Then
      var_conexion = "Provider=SQLOLEDB.1;Password=" & parametros(3) & ";Persist Security Info=True;User ID=sa;Initial Catalog=vianney;Data Source=DISTRIBUCION"
   End If
   If var_empresa = "06" Or var_empresa = "17" Then
      var_conexion = "Provider=SQLOLEDB.1;Password=" & parametros(3) & ";Persist Security Info=True;User ID=sa;Initial Catalog=sid;Data Source=admcdindustrial"
   End If
   If var_empresa = "18" Then
      var_conexion = "Provider=SQLOLEDB.1;Password=" & parametros(3) & ";Persist Security Info=True;User ID=sa;Initial Catalog=sidtextilera;Data Source=sqlquezada2"
   End If
   If cnn_ver_factura_electronica.State = 1 Then
      cnn_ver_factura_electronica.Close
   End If
   cnn_ver_factura_electronica.Open var_conexion
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub txt_a_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   End If
End Sub

Private Sub txt_de_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_a.SetFocus
   End If
End Sub


Private Sub txt_serie_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_de.SetFocus
   End If
End Sub

