VERSION 5.00
Begin VB.Form frmnota_credito_electronica_cantia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nota electronica CANTIA"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5640
      Picture         =   "frmnota_credito_electronica_cantia.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_nc_electronica 
      Caption         =   "NC electrónica"
      Height          =   345
      Left            =   45
      TabIndex        =   13
      Top             =   30
      Width           =   1605
   End
   Begin VB.Frame Frame4 
      Height          =   1170
      Left            =   75
      TabIndex        =   10
      Top             =   2145
      Width           =   5910
      Begin VB.TextBox txt_correo 
         Height          =   405
         Left            =   840
         TabIndex        =   15
         Top             =   645
         Width           =   4950
      End
      Begin VB.TextBox txt_cliente 
         Height          =   405
         Left            =   840
         TabIndex        =   12
         Top             =   195
         Width           =   4950
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Correo:"
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   750
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   300
         Width           =   525
      End
   End
   Begin VB.Frame Frame3 
      Height          =   675
      Left            =   75
      TabIndex        =   7
      Top             =   1425
      Width           =   5895
      Begin VB.TextBox txt_importe 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   2895
         TabIndex        =   8
         Top             =   150
         Width           =   2895
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   270
         Width           =   570
      End
   End
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   0
      TabIndex        =   6
      Top             =   390
      Width           =   6075
   End
   Begin VB.Frame Frame1 
      Height          =   915
      Left            =   75
      TabIndex        =   0
      Top             =   465
      Width           =   5895
      Begin VB.CommandButton cmd_buscar 
         Caption         =   "Buscar"
         Height          =   450
         Left            =   4530
         TabIndex        =   5
         Top             =   285
         Width           =   960
      End
      Begin VB.TextBox txt_tipo_4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2715
         TabIndex        =   4
         Top             =   263
         Width           =   1740
      End
      Begin VB.TextBox txt_tipo_3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   1890
         TabIndex        =   3
         Top             =   263
         Width           =   810
      End
      Begin VB.TextBox txt_tipo_2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   1140
         TabIndex        =   2
         Top             =   263
         Width           =   735
      End
      Begin VB.TextBox txt_tipo_1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   450
         TabIndex        =   1
         Top             =   263
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmnota_credito_electronica_cantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_buscar_Click()
   If IsNumeric(Me.txt_tipo_1) Then
      If IsNumeric(Me.txt_tipo_2) Then
         If IsNumeric(Me.txt_tipo_3) Then
            If IsNumeric(Me.txt_tipo_4) Then
               var_cadena = "select * from vw_NotasDeCredito where folTda_codigo=" + Me.txt_tipo_1 + " and FolEst_codigo =" + Me.txt_tipo_2 + " and FolDoc_codigo = " + Me.txt_tipo_3 + " and FolConsecutivo=" + Me.txt_tipo_4
               rs.Open var_cadena, cnn_compucaja, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  Me.txt_cliente = IIf(IsNull(rs!cli_nombre), "", rs!cli_nombre)
                  Me.txt_importe = Format(IIf(IsNull(rs!NC_TOTALIMPORTE), 0, rs!NC_TOTALIMPORTE), "###,###,##0.00")
               Else
                  MsgBox "La nota de crédito no existe", vbOKOnly, "ATENCION"
               End If
               rs.Close
            Else
               MsgBox "Parametro 4 incorrecto", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Parametro 3 incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Parametro 2 incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Parametro 1 incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nc_electronica_Click()
   var_ruta_documentos_electronicos = "\\facelectronica\fefiles\Conectorcan\envio\por_enviar"
   If IsNumeric(Me.txt_importe) Then
      rsaux10.Open "select * from TB_notas_credito_COMPUCAJA_SID where vcha_FOL_folio_compucaja = '" + Me.txt_tipo_1 + "-" + Me.txt_tipo_2 + "-" + Me.txt_tipo_3 + "-" + Me.txt_tipo_4 + "'", cnn, adOpenDynamic, adLockOptimistic
      If rsaux10.EOF Then
         cnn.BeginTrans
         rsaux1.Open "SELECT * FROM TB_sERIES WHERE VCHA_UOR_UNIDAD_ID = 'CANTIANC'", cnn, adOpenDynamic, adLockOptimistic
         var_serie = IIf(IsNull(rsaux1!vcha_Ser_Serie_id), "", rsaux1!vcha_Ser_Serie_id)
         var_numero_factura = IIf(IsNull(rsaux1!inte_ser_nota_credito), 0, rsaux1!inte_ser_nota_credito)
         rsaux1.Close
         Open (App.Path & "\renombra" + Trim(Str(var_numero_factura)) + ".bat") For Output As #2
         If var_empresa = "31" Then
            Print #2, "ren " + var_ruta_documentos_electronicos + "\" + Trim(var_serie) + Trim(Str(var_numero_factura)) + ".fi " + Trim(var_serie) + Trim(Str(var_numero_factura)) + ".ff"
         End If
         Close #2
         rsaux3.Open "select * from vw_NotasDeCredito where folTda_codigo=" + Me.txt_tipo_1 + " and FolEst_codigo =" + Me.txt_tipo_2 + " and FolDoc_codigo = " + Me.txt_tipo_3 + " and FolConsecutivo=" + Me.txt_tipo_4, cnn_compucaja, adOpenDynamic, adLockOptimistic
         Open (var_ruta_documentos_electronicos & "\" + Trim(var_serie) + Trim(Str(var_numero_factura)) + ".fi") For Output As #1
         var_cadena = "Outputmode=" + Chr(13) + "<Factura>" + Chr(13) + "<Comprobante>" + Chr(13) + "Version=2.0" + Chr(13) + "Serie=" + var_serie + Chr(13) + "folio=" + CStr(var_numero_factura) + Chr(13)
         var_año = CStr(Year(rsaux3!nc_FECHA))
         var_mes = CStr(Month(rsaux3!nc_FECHA))
         var_dia = CStr(Day(rsaux3!nc_FECHA))
         var_hora = CStr(Hour(rsaux3!nc_FECHA))
         var_minuto = CStr(Minute(rsaux3!nc_FECHA))
         var_segundo = CStr(Second(rsaux3!nc_FECHA))
         If Len(var_año) = 2 Then
            var_año = "20" + var_año
         End If
         If Len(var_mes) = 1 Then
            var_mes = "0" + var_mes
         End If
         If Len(var_dia) = 1 Then
            var_dia = "0" + var_dia
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
         var_cadena_fecha = var_año + "-" + var_mes + "-" + var_dia + "T" + var_hora + ":" + var_minuto + ":" + var_segundo
         If rsaux1.State = 1 Then
            rsaux1.Close
         End If
         
         var_rfc_cliente_1 = IIf(IsNull(rsaux3!cli_registrotributario), "", rsaux3!cli_registrotributario)
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
         
         
         If var_rfc_cliente = "XAXX010101000" Then
            var_desglose = 1
         Else
            var_desglose = 1.16
         End If

         var_cadena = var_cadena + "fecha=" + var_cadena_fecha + Chr(13)
         var_cadena = var_cadena + "noAprobacion=" + Chr(13)
         var_cadena = var_cadena + "anoAprobacion=" + Chr(13)
         var_cadena = var_cadena + "tipoDeComprobante=NOTA DE CREDITO" + Chr(13)
         var_cadena = var_cadena + "formaDePago=PAGO HECHO EN UNA SOLA EXHIBICION" + Chr(13)
         var_cadena = var_cadena + "condicionesDePago=" + Chr(13)
         var_total = rsaux3!NC_TOTALIMPORTE
         Call numero_letras(var_total, "1")
         var_cantidad_letra = canstr
         var_importe_iva = var_total - (var_total / var_desglose)
         var_subtotal = var_total - var_importe_iva
         var_cadena = var_cadena + "subtotal=" + Format(CStr(var_subtotal), "###,###,##0.000000") + Chr(13)
         'rsaux1.Open "select sum(ddt_importe) from via_facturas where foltda_codigo=" + Me.txt_tipo_1 + " and folest_codigo=" + Me.txt_tipo_2 + " and foldoc_codigo=" + Me.txt_tipo_3 + " and folconsecutivo=" + Me.txt_tipo_4, cnn_compucaja, adOpenDynamic, adLockOptimistic
         'If Not rsaux1.EOF Then
         '   var_importe_descuento = IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value)
            var_cadena = var_cadena + "descuento=" + Chr(13)
         'Else
         '   var_cadena = var_cadena + "descuento=0" + Chr(13)
         'End If
         'rsaux1.Close
         var_cadena = var_cadena + "descuento1=" + Chr(13)
         var_cadena = var_cadena + "descuento2=" + Chr(13)
         var_cadena = var_cadena + "conceptodescuento1=" + Chr(13)
         var_cadena = var_cadena + "conceptodescuento2=" + Chr(13)
         var_cadena = var_cadena + "tasadescuento1=" + Chr(13)
         var_cadena = var_cadena + "tasadescuento2=" + Chr(13)
         If rsaux1.State = 1 Then
            rsaux1.Close
         End If
         rsaux1.Open "select * from tb_empresa_FACTURA_ELECTRONICA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
         var_certificado = rsaux1!vcha_emp_certificado
         var_expedido = rsaux1!vcha_emp_expedido
         var_cadena = var_cadena + "iva=" + Format(CStr(var_importe_iva), "###,###,##0.000000") + Chr(13)
         var_cadena = var_cadena + "total=" + Format(CStr(var_total), "###,###,##0.000000") + Chr(13)
         var_cadena = var_cadena + "retencion=" + Chr(13)
         var_cadena = var_cadena + "factorretencioniva=" + Chr(13)
         var_cadena = var_cadena + "</Comprobante>" + Chr(13) + Chr(13)
         var_cadena = var_cadena + "<Emisor>" + Chr(13)
         var_cadena = var_cadena + "erfc=" + rsaux1!VCHA_eMP_RFC + Chr(13)
         var_cadena = var_cadena + "enombre=" + rsaux1!VCHA_EMP_NOMBRE + Chr(13)
         var_cadena = var_cadena + "</Emisor>" + Chr(13) + Chr(13)
         var_cadena = var_cadena + "<DomicilioFiscal>" + Chr(13)
         var_cadena = var_cadena + "ecalle=" + rsaux1!VCHA_eMP_CALLE + Chr(13)
         var_cadena = var_cadena + "enoExterior=" + rsaux1!VCHA_eMP_exterior + Chr(13)
         var_cadena = var_cadena + "enoInterior=" + Chr(13)
         var_cadena = var_cadena + "ecolonia=" + rsaux1!VCHA_eMP_COLONIA + Chr(13)
         var_cadena = var_cadena + "elocalidad=" + rsaux1!VCHA_EMP_LOCALIDAD + Chr(13)
         var_cadena = var_cadena + "ereferencia=" + Chr(13)
         var_cadena = var_cadena + "emunicipio=" + rsaux1!VCHA_EMP_MUNICIPIO + Chr(13)
         var_cadena = var_cadena + "eestado=" + rsaux1!VCHA_EMP_ESTADO + Chr(13)
         var_cadena = var_cadena + "epais=" + rsaux1!VCHA_eMP_PAIS + Chr(13)
         var_cadena = var_cadena + "ecodigoPostal=" + rsaux1!VCHA_EMP_CODIGO_POSTAL + Chr(13)
         var_cadena = var_cadena + "etel=" + IIf(IsNull(rsaux1!VCHA_EMP_TELEFONO), "", rsaux1!VCHA_EMP_TELEFONO) + Chr(13)
         var_cadena = var_cadena + "eemail=" + IIf(IsNull(rsaux1!VCHA_EMP_EMAIL), "", rsaux1!VCHA_EMP_EMAIL) + Chr(13)
         correo = IIf(IsNull(rsaux1!VCHA_EMP_EMAIL), "", rsaux1!VCHA_EMP_EMAIL)
         var_cadena = var_cadena + "</DomicilioFiscal>" + Chr(13) + Chr(13)
         
         
         var_cadena = var_cadena + "<ExpedidoEn>" + Chr(13) + Chr(13)
         var_cadena = var_cadena + "ex_calle=" + rsaux1!VCHA_eMP_CALLE + Chr(13)
         var_cadena = var_cadena + "ex_noExterior=" + rsaux1!VCHA_eMP_exterior + Chr(13)
         var_cadena = var_cadena + "ex_noInterior=" + Chr(13)
         var_cadena = var_cadena + "ex_colonia=" + rsaux1!VCHA_eMP_COLONIA + Chr(13)
         var_cadena = var_cadena + "ex_localidad=" + rsaux1!VCHA_EMP_LOCALIDAD + Chr(13)
         var_cadena = var_cadena + "ex_referencia=" + Chr(13)
         var_cadena = var_cadena + "ex_municipio=" + rsaux1!VCHA_EMP_MUNICIPIO + Chr(13)
         var_cadena = var_cadena + "ex_estado=" + rsaux1!VCHA_EMP_ESTADO + Chr(13)
         var_cadena = var_cadena + "ex_pais=" + rsaux1!VCHA_eMP_PAIS + Chr(13)
         var_cadena = var_cadena + "ex_codigoPostal=" + rsaux1!VCHA_EMP_CODIGO_POSTAL + Chr(13)
         var_cadena = var_cadena + "</ExpedidoEn>"
         
         
         
         var_cadena = var_cadena + "<Receptor>" + Chr(13)
         var_cadena = var_cadena + "noCliente=" + rsaux3!cli_codigo + Chr(13)
                                         
         var_rfc_cliente_1 = IIf(IsNull(rsaux3!cli_registrotributario), "", rsaux3!cli_registrotributario)
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
         var_cadena = var_cadena + "rfc=" + var_rfc_cliente + Chr(13)
         var_cadena = var_cadena + "nombre=" + rsaux3!cli_nombre + Chr(13)
         var_cadena = var_cadena + "</Receptor>" + Chr(13) + Chr(13)
         var_cadena = var_cadena + "<Cliente>" + Chr(13)
         var_cadena = var_cadena + "domicilio=" + IIf(IsNull(rsaux3!cli_domicilio), "", rsaux3!cli_domicilio) + Chr(13)
         var_cadena = var_cadena + "calle=" + Chr(13)
         var_cadena = var_cadena + "noExterior=" + Chr(13)
         var_cadena = var_cadena + "noInterior=" + Chr(13)
         var_cadena = var_cadena + "colonia=" + Chr(13)
         var_cadena = var_cadena + "localidad=" + IIf(IsNull(rsaux3!cli_asentamiento), "", rsaux3!cli_asentamiento) + Chr(13)
         rsaux1.Close
         var_cadena = var_cadena + "referencia=" + Chr(13)
         var_cadena = var_cadena + "municipio=" + IIf(IsNull(rsaux3!cli_municipio), "", rsaux3!cli_municipio) + Chr(13)
         var_cadena = var_cadena + "estado=" + IIf(IsNull(rsaux3!cli_estado), "", rsaux3!cli_estado) + Chr(13)
         VAR_NOMBRE_PAIS = IIf(IsNull(rsaux3!cli_nombre), "MEXICO", rsaux3!cli_nombre)
         If Trim(VAR_NOMBRE_PAIS) = "" Then
            VAR_NOMBRE_PAIS = "MEXICO"
         End If
         var_cadena = var_cadena + "pais=" + VAR_NOMBRE_PAIS + Chr(13)
         var_cadena = var_cadena + Chr(13)
         var_cadena = var_cadena + "codigoPostal=" + IIf(IsNull(rsaux3!cli_cp), "", rsaux3!cli_cp) + Chr(13)
         var_cadena = var_cadena + "tel=" + Chr(13)
         var_cadena = var_cadena + "email=" + Me.txt_correo + Chr(13)
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

         var_k = 0
         While Not rsaux3.EOF
               var_k = var_k + 1
               pxx = CStr(var_k)
               If Len(pxx) = 1 Then
                  pxx = "0" + pxx
               End If
               var_cadena = var_cadena + "p" + pxx + "_cantidad=" + CStr(IIf(IsNull(rsaux3!DNC_cantidad), 0, rsaux3!DNC_cantidad)) + Chr(13)
               var_cadena = var_cadena + "p" + pxx + "_unidad=PZA" + Chr(13)
               var_cadena = var_cadena + "p" + pxx + "_noIdentificacion=" + IIf(IsNull(rsaux3!art_codigo), "", rsaux3!art_codigo) + Chr(13)
               var_linea = IIf(IsNull(rsaux3!art_codigo), "", rsaux3!art_codigo) + " " + IIf(IsNull(rsaux3!art_descripcion), "", rsaux3!art_descripcion)
               var_cadena = var_cadena + "p" + pxx + "_descripcion=" + var_linea + Chr(13)
               var_precio = IIf(IsNull(rsaux3!DNC_preciounitario), 0, rsaux3!DNC_preciounitario)
               var_descuento_1 = 0
               var_descuento_2 = 0
               var_descuento_3 = 0
               var_porcentaje = (100 - var_descuento_1) / 100
               If var_porcentaje = 0 Then
                  var_precio = 0
               Else
                  var_precio = (var_precio / var_porcentaje) / var_desglose
               End If
               var_importe_descuento_1_2 = 0
               var_importe_descuento_1 = 0
               var_cadena = var_cadena + "p" + pxx + "_valorUnitario=" + Format(CStr(var_precio), "###,###,##0.000000") + Chr(13)
               var_cadena = var_cadena + "p" + pxx + "_importe=" + Format(CStr(var_precio * CStr(IIf(IsNull(rsaux3!DNC_cantidad), 0, rsaux3!DNC_cantidad))), "###,###,##0.000000") + Chr(13)
                                            
               rsaux3.MoveNext
         Wend
         rsaux3.MoveFirst
         var_cadena = var_cadena + "</Concepto>" + Chr(13) + Chr(13)
         var_cadena = var_cadena + "<Otros>" + Chr(13)
         var_cadena = var_cadena + "certificado=" + IIf(IsNull(var_certificado), "", var_certificado) + Chr(13)
         var_cadena = var_cadena + "cant_letra=" + var_cantidad_letra + Chr(13)
         var_cadena = var_cadena + "factoriva=16%" + Chr(13)
         var_cadena = var_cadena + "moneda=PESOS" + Chr(13)
         var_cadena = var_cadena + "tipodeCambio=1" + Chr(13)
         var_cadena = var_cadena + "pedido=" + Chr(13)
         var_cadena = var_cadena + "Embarque=" + Me.txt_tipo_1 + "-" + Me.txt_tipo_2 + "-" + Me.txt_tipo_3 + "-" + Me.txt_tipo_4 + Chr(13)
         var_referencia_Bancaria = ""
         var_referencia_Bancaria = ""
         var_cadena = var_cadena + "referenciabancaria=" + var_referencia_Bancaria + Chr(13)
         var_cadena = var_cadena + "fechaPedido=" + var_cadena_fecha + Chr(13)
         var_cadena = var_cadena + "expedicion=" + var_expedido + Chr(13)
         var_cadena = var_cadena + "observaciones=" + Chr(13)
         var_cadena = var_cadena + "conceptoExtra1=" + Chr(13)
         var_cadena = var_cadena + "montoconceptoExtra1=" + Chr(13)
         var_cadena = var_cadena + "conceptoExtra2=" + Chr(13)
         var_cadena = var_cadena + "montoconceptoExtra2=" + Chr(13)
         var_cadena = var_cadena + "tipoimpresion=2" + Chr(13)
         var_cadena = var_cadena + "agente=" + Chr(13)
         If var_empresa = "15" Then
            var_cadena = var_cadena + "formato=MHESTAMPADOS_V01.DAT" + Chr(13)
         End If
         If var_empresa = "16" Then
            var_cadena = var_cadena + "formato=MHMULTIBONDEADOS_V01.DAT" + Chr(13)
         End If
         If var_empresa = "31" Then
            var_cadena = var_cadena + "formato=MHCANTIA_V01.DAT" + Chr(13)
         End If
                              
         var_cadena = var_cadena + "</Otros>" + Chr(13) + Chr(13)
         var_cadena = var_cadena + "<addenda>" + Chr(13)
         var_cadena = var_cadena + "</addenda>" + Chr(13) + Chr(13)
         var_cadena = var_cadena + "</Factura>"
         Print #1, var_cadena
         Close #1
         var_Archivo = App.Path & "\renombra" + Trim(Str(var_numero_factura)) + ".bat"
         x = Shell(var_Archivo, vbHide)
         rsaux1.Open "INSERT INTO TB_NOTAS_CREDITO_COMPUCAJA_SID (VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, VCHA_FOL_FOLIO_COMPUCAJA) VALUES ('" + var_serie + "', '" + CStr(var_numero_factura) + "','" + Me.txt_tipo_1 + "-" + Me.txt_tipo_2 + "-" + Me.txt_tipo_3 + "-" + Me.txt_tipo_4 + "')", cnn, adOpenDynamic, adLockOptimistic
         rsaux1.Open "UPDATE TB_sERIES SET inte_ser_nota_Credito = inte_ser_nota_Credito + 1 WHERE VCHA_UOR_UNIDAD_ID = 'CANTIANC'", cnn, adOpenDynamic, adLockOptimistic
         rsaux3.Close
         cnn.CommitTrans
      Else
         MsgBox "El folio " + Me.txt_tipo_1 + "-" + Me.txt_tipo_2 + "-" + Me.txt_tipo_3 + "-" + Me.txt_tipo_4 + ", se encuentra en la nota de crédito " + CStr(rsaux10!inte_Car_numero), vbOKOnly, "ATENCION"
      End If
      rsaux10.Close
   Else
      MsgBox "Nota de crédito incorrecta incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 1500
   Left = 3200
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub txt_tipo_1_Change()
   Me.txt_cliente = ""
   Me.txt_importe = ""
End Sub

Private Sub txt_tipo_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_tipo_2.SetFocus
   End If
End Sub

Private Sub txt_tipo_2_Change()
   Me.txt_cliente = ""
   Me.txt_importe = ""
End Sub

Private Sub txt_tipo_2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_tipo_3.SetFocus
   End If
End Sub

Private Sub txt_tipo_3_Change()
   Me.txt_cliente = ""
   Me.txt_importe = ""
End Sub

Private Sub txt_tipo_3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_tipo_4.SetFocus
   End If
End Sub

Private Sub txt_tipo_4_Change()
   Me.txt_cliente = ""
   Me.txt_importe = ""
End Sub

Private Sub txt_tipo_4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_buscar.SetFocus
   End If
End Sub
