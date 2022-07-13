VERSION 5.00
Begin VB.Form frmcancela_documentos_existentes 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancelación de documentos fiscales existentes"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   6990
   Begin VB.CommandButton cmd_reimprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   795
      Picture         =   "frmcancelacion_documentos_existentes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Reimprimir notas de crédito"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   465
      Picture         =   "frmcancelacion_documentos_existentes.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Refacturar"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Caption         =   " Documento a cancelar "
      Height          =   1455
      Left            =   90
      TabIndex        =   8
      Top             =   435
      Width           =   6825
      Begin VB.TextBox txt_serie 
         Height          =   315
         Left            =   1500
         TabIndex        =   3
         Top             =   645
         Width           =   1035
      End
      Begin VB.TextBox txt_documento 
         Height          =   315
         Left            =   1500
         TabIndex        =   1
         Top             =   300
         Width           =   1035
      End
      Begin VB.ComboBox cmb_documentos 
         Height          =   315
         ItemData        =   "frmcancelacion_documentos_existentes.frx":0204
         Left            =   2550
         List            =   "frmcancelacion_documentos_existentes.frx":0217
         TabIndex        =   2
         Top             =   300
         Width           =   4155
      End
      Begin VB.TextBox txt_numero 
         Height          =   315
         Left            =   1500
         TabIndex        =   4
         Top             =   990
         Width           =   1035
      End
      Begin VB.Label lbl_estatus 
         Caption         =   "Label3"
         Height          =   210
         Left            =   2745
         TabIndex        =   12
         Top             =   1050
         Width           =   3465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Documento: "
         Height          =   195
         Left            =   495
         TabIndex        =   11
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   495
         TabIndex        =   10
         Top             =   1050
         Width           =   600
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   495
         TabIndex        =   9
         Top             =   705
         Width           =   405
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6540
      Picture         =   "frmcancelacion_documentos_existentes.frx":0253
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir Esc"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      Picture         =   "frmcancelacion_documentos_existentes.frx":088D
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cancelar Alt + C"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   75
      Left            =   60
      TabIndex        =   0
      Top             =   330
      Width           =   6885
   End
End
Attribute VB_Name = "frmcancela_documentos_existentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_serie As String
Dim var_estatus As String
Dim var_tipo_facturacion As String
Dim var_tipo_nota_credito As String
Dim var_numero_renglones As Integer
Dim var_factura_nueva As Double
Dim var_numero_movimiento As Double


Private Sub imprime_factura()
   If var_empresa = "02" Or var_empresa = "18" Or var_empresa = "31" Then
      var_serie = Me.txt_serie
                              Open (App.Path & "\factura" + Trim(Str(var_factura_nueva)) + ".txt") For Output As #1
                              Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
                              'Print #1, Chr(27) + Chr(67) + Chr(44) + Chr(13)
                              'Print #1, ""
                              Print #1, Spc(105); Str(var_factura_nueva)
                              Print #1, ""
                              Print #1, ""
                              Print #1, Spc(105); Str(rs!INTE_CAR_PLAZO) + " DIAS DE VENCIMIENTO" + "                  " + Format(rs!dtim_Car_fecha, "Short Date")
                              Print #1, ""
                              'Print #1, Spc(92); Str(rs!inte_car_PLAZO) + " DIAS DE VENCIMIENTO"
                              var_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                              var_cliente_coppel = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
                              For var_j = 1 + Len(Trim(var_cliente)) To 83
                                  var_cliente = var_cliente + " "
                              Next var_j
                              var_cliente = var_cliente + "               AGUASCALIENTES, AGS."
                              Print #1, Spc(10); var_cliente
                              var_domicilio = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " COLONIA: " + IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
                              For var_j = 1 + Len(Trim(var_domicilio)) To 83
                                  var_domicilio = var_domicilio + " "
                              Next var_j
                              var_agente = ""
                              var_agente = IIf(IsNull(rs!vcha_AGE_aGENTE_ID), "", rs!vcha_AGE_aGENTE_ID)
                              rsaux4.Open "select * from tb_agentes where vcha_age_agente_id = '" + var_agente + "'", cnn, adOpenDynamic, adLockOptimistic
                              var_nombre_agenteS = IIf(IsNull(rsaux4!VCHA_AGE_NOMBRE), "", rsaux4!VCHA_AGE_NOMBRE)
                              rsaux4.Close
                              For var_j = 1 + Len(Trim(var_agente)) To 8
                                  var_agente = var_agente + " "
                              Next var_j
                              
                              var_agente = var_agente + var_nombre_agenteS
                              var_domicilio = var_domicilio
                              'Print #1, Spc(111); var_agente
                              Print #1, Spc(10); var_domicilio
                              var_ciudad = ""
                              var_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                              For var_j = 1 + Len(Trim(var_ciudad)) To 37
                                 var_ciudad = var_ciudad + " "
                              Next var_j
                              
                              var_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                              var_ciudad = var_ciudad
                              var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                              var_ciudad = var_ciudad + var_rfc
                              
                              For var_j = 1 + Len(Trim(var_estado)) To 46
                                 var_estado = var_estado + " "
                              Next var_j
                              

                              For var_j = 1 + Len(Trim(var_ciudad)) To 14
                                 var_ciudad = var_ciudad + " "
                              Next var_j
                               
                              var_ciudad = var_ciudad + "                                                      " + var_agente
                              
                              VAR_EMBARQUE = "EMB.: " + CStr(var_numero_embarque)
                              var_ordern_surtido = x
                              Print #1, Spc(10); var_ciudad
                              var_rfc = "RFC:  " + var_rfc
                              var_rfc = IIf(IsNull(rs!VCHA_ESB_ESTABLECIMIENTO_ID), "", rs!VCHA_ESB_ESTABLECIMIENTO_ID) + " " + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre) + " " + " C.P. " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
                              For var_j = 1 + Len(Trim(var_rfc)) To 89
                                 var_rfc = var_rfc + " "
                              Next var_j
                                 If Trim(var_cliente_coppel) = "C000002947" Then
                                    rsaux5.Open "select * from tb_encabezado_pedidos where inte_ped_numero = " + Trim(Str(IIf(IsNull(rs!inte_ped_numero), 0, rs!inte_ped_numero))), cnn, adOpenDynamic, adLockOptimistic
                                    If Not rsaux5.EOF Then
                                       var_rfc = var_rfc + "               PED.: " + Trim(CStr(IIf(IsNull(rsaux5!vcha_ped_pedido_Externo), "", rsaux5!vcha_ped_pedido_Externo))) + " "
                                    End If
                                    rsaux5.Close
                                 Else
                                    var_rfc = var_rfc + "                   PED.: " + Trim(Str(IIf(IsNull(rs!inte_ped_numero), 0, rs!inte_ped_numero))) + " "
                                 End If
                              var_rfc = var_rfc + " O.S.: " + Trim(Str(IIf(IsNull(rs!INTE_ORS_ORDEN_SURTIDO), 0, rs!INTE_ORS_ORDEN_SURTIDO))) + " " + VAR_EMBARQUE
                              Print #1, var_rfc
                              'Print #1, Spc(10); IIf(IsNull(rs!vcha_esb_establecimiento_id), "", rs!vcha_esb_establecimiento_id)
                              Print #1, ""
                              Print #1, ""
                              var_importe_descuento_1 = 0
                              var_importe_descuento_2 = 0
                              var_importe_descuento_3 = 0
                              var_contador_promociones = 0
                              var_cantidad_total = 0
                              For var_k = 1 To var_renglones_factura
                                 If Not rs.EOF Then
                                    var_linea = ""
                                    var_marca_promocion = " "
                                    var_promocion_1 = IIf(IsNull(rs!floa_sal_promocion_1), 0, rs!floa_sal_promocion_1)
                                    var_promocion_2 = IIf(IsNull(rs!FLOA_SAL_PROMOCION_2), 0, rs!FLOA_SAL_PROMOCION_2)
                                    If var_promocion_1 > 0 Then
                                       var_marca_promocion = "*"
                                       var_contador_promociones = var_contador_promociones + 1
                                    End If
                                    If var_promocion_2 > 0 Then
                                       var_marca_promocion = "*"
                                       var_contador_promociones = var_contador_promociones + 1
                                    End If
                                    var_linea = IIf(IsNull(rs!vcha_agr_agrupador_id), "", rs!vcha_agr_agrupador_id)
                                    For var_j = 1 + Len(Trim(var_linea)) To 15
                                        var_linea = var_linea + " "
                                    Next var_j
                                    var_linea = var_linea + IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura)
                                    
                                    
                                       
                                       ''' imprimir cantidad en la orilla
                                       var_cantidad_nueva = Format(IIf(IsNull(rs!Cantidad), 0, rs!Cantidad), "###,###,##0.00")
                                       If Len(Trim(var_cantidad_nueva)) < 14 Then
                                          For var_j = 1 + Len(Trim(var_cantidad_nueva)) To 14
                                             var_cantidad_nueva = " " + var_cantidad_nueva
                                          Next var_j
                                       End If
                                       While Len((var_linea)) < 60
                                             var_linea = var_linea + " "
                                       Wend
                                       var_linea = var_linea + var_cantidad_nueva
                                       
                                       ''' imprimir cantidad en la orilla
                                       
                                    
                                    
                                    var_i = 0
                                    While Len((var_linea)) < 115
                                          var_linea = var_linea + " "
                                    Wend
                                    var_linea = var_linea + " "
                                    var_linea = var_linea + var_marca_promocion
                                    var_cantidad = Format(IIf(IsNull(rs!Cantidad), 0, rs!Cantidad), "###,###,##0.00")
                                    var_cantidad_total = var_cantidad_total + IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)
                                    If Len(Trim(var_cantidad)) < 14 Then
                                       For var_j = 1 + Len(Trim(var_cantidad)) To 14
                                          var_cantidad = " " + var_cantidad
                                       Next var_j
                                    End If
                                    var_precio = IIf(IsNull(rs!Importe), 0, rs!Importe)
                                    var_descuento_1 = IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1)
                                    var_descuento_2 = IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2)
                                    var_descuento_3 = IIf(IsNull(rs!floa_car_porcentaje_descuento_3), 0, rs!floa_car_porcentaje_descuento_3)
                                    var_porcentaje = (100 - var_descuento_1) / 100
                                    var_precio = var_precio * var_porcentaje
                                    var_importe_descuento_1_2 = (IIf(IsNull(rs!Importe), 0, rs!Importe) - var_precio)
                                    var_importe_descuento_1 = var_importe_descuento_1 + (IIf(IsNull(rs!Importe), 0, rs!Importe) - var_precio)
                                    var_precio = var_precio * ((100 - var_descuento_2) / 100)
                                    var_importe_descuento_2 = var_importe_descuento_2 + (IIf(IsNull(rs!Importe), 0, rs!Importe) - (var_importe_descuento_1_2 + var_precio))
                                    var_precio = var_precio * ((100 - var_descuento_3) / 100)
                                    var_precio = var_precio / IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)
                                    'var_precio_str = Format(var_precio / IIf(IsNull(rs!cantidad), 0, rs!cantidad), "###,###,##0.00")
                                    var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                                    If Len(Trim(var_rfc)) > 0 Then
                                       var_precio_str = Format(IIf(IsNull(rs!Importe), 0, rs!Importe) / IIf(IsNull(rs!Cantidad), 0, rs!Cantidad), "###,###,##0.00")
                                    Else
                                       var_precio_str = Format((IIf(IsNull(rs!Importe), 0, rs!Importe) / IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)) * (1 + (rs!floa_car_porcentaje_iva / 100)), "###,###,##0.00")
                                    End If
                                    If Len(Trim(var_precio_str)) < 14 Then
                                       For var_j = 1 + Len(Trim(var_precio_str)) To 14
                                           var_precio_str = " " + var_precio_str
                                       Next var_j
                                    End If
                                    var_linea = var_linea + var_cantidad + var_precio_str
                                    If Len(Trim(var_rfc)) > 0 Then
                                       var_importe = Format((IIf(IsNull(rs!Importe), 0, rs!Importe)), "###,###,##0.00")
                                       If Len(Trim(var_importe)) < 14 Then
                                           For var_j = 1 + Len(Trim(var_importe)) To 14
                                              var_importe = " " + var_importe
                                           Next var_j
                                       End If
                                    Else
                                       var_importe = Format((IIf(IsNull(rs!Importe), 0, rs!Importe) * (1 + (rs!floa_car_porcentaje_iva / 100))), "###,###,##0.00")
                                       If Len(Trim(var_importe)) < 14 Then
                                           For var_j = 1 + Len(Trim(var_importe)) To 14
                                              var_importe = " " + var_importe
                                           Next var_j
                                       End If
                                    End If
                                    var_linea = var_linea + var_importe
                                     
                                    Print #1, var_linea
                                    rs.MoveNext
                                 Else
                                    Print #1, ""
                                 End If
                              Next var_k
                              Print #1, ""
                              'Print #1, ""
                              rs.MoveFirst
                              
                              var_cantidad_total_str = Format(var_cantidad_total, "###,###,##0.00")
                              var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                              If Len(Trim(var_rfc)) > 0 Then
                                 var_cantidad_letra = rs!vcha_car_importe_letra
                                 var_importe_descuento_1_str = Format(IIf(IsNull(rs!FLOA_CAR_IMPORTE_DESCUENTO_1), 0, rs!FLOA_CAR_IMPORTE_DESCUENTO_1) / (rs!floa_car_tipo_cambio), "###,###,##0.00")
                                 If Len(Trim(var_importe_descuento_1_str)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_importe_descuento_1_str)) To 14
                                         var_importe_descuento_1_str = " " + var_importe_descuento_1_str
                                    Next var_j
                                 End If
                                 var_importe_descuento_2_str = Format(IIf(IsNull(rs!FLOA_CAR_IMPORTE_DESCUENTO_2), 0, rs!FLOA_CAR_IMPORTE_DESCUENTO_2) / (rs!floa_car_tipo_cambio), "###,###,##0.00")
                                 If Len(Trim(var_importe_descuento_2_str)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_importe_descuento_2_str)) To 14
                                        var_importe_descuento_2_str = " " + var_importe_descuento_2_str
                                    Next var_j
                                 End If
                              Else
                                 var_cantidad_letra = rs!vcha_car_importe_letra
                                 var_importe_descuento_1_str = Format((IIf(IsNull(rs!FLOA_CAR_IMPORTE_DESCUENTO_1), 0, rs!FLOA_CAR_IMPORTE_DESCUENTO_1)) * (1 + (rs!floa_car_porcentaje_iva / 100)) / rs!floa_car_tipo_cambio, "###,###,##0.00")
                                 If Len(Trim(var_importe_descuento_1_str)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_importe_descuento_1_str)) To 14
                                         var_importe_descuento_1_str = " " + var_importe_descuento_1_str
                                    Next var_j
                                 End If
                                 var_importe_descuento_2_str = Format((IIf(IsNull(rs!FLOA_CAR_IMPORTE_DESCUENTO_2), 0, rs!FLOA_CAR_IMPORTE_DESCUENTO_2)) * (1 + (rs!floa_car_porcentaje_iva / 100)) / rs!floa_car_tipo_cambio, "###,###,##0.00")
                                 If Len(Trim(var_importe_descuento_2_str)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_importe_descuento_2_str)) To 14
                                        var_importe_descuento_2_str = " " + var_importe_descuento_2_str
                                    Next var_j
                                 End If
                              End If
                              var_linea = "- DESCUENTO DEL " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1))) + "%"
                              If Len(Trim(var_linea)) < 145 Then
                                 For var_j = 1 + Len(Trim(var_linea)) To 145
                                     var_linea = var_linea + " "
                                 Next var_j
                              End If
                              Print #1, var_linea + var_importe_descuento_1_str
                              If var_empresa = "18" Then
                                 var_linea = ""
                              Else
                                 If Trim(var_cliente_coppel) = "C000002947" Then
                                    var_linea = ""
                                 Else
                                    var_linea = "- DESCUENTO POR PAGO OPORTUNO " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2))) + "%"
                                 End If
                              End If
                              If Len(Trim(var_linea)) < 145 Then
                                 For var_j = 1 + Len(Trim(var_linea)) To 145
                                     var_linea = var_linea + " "
                                 Next var_j
                              End If
                              If Trim(var_cliente_coppel) <> "C000002947" Then
                                 var_linea = var_linea + var_importe_descuento_2_str
                              End If
                              Print #1, var_linea
                              If var_contador_promociones > 0 Then
                                 Print #1, "PROMOCION EN ARTICULOS MARCADOS CON *"
                              Else
                                 Print #1, ""
                              End If
                              
                              var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                              var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
                              
                              If Len(Trim(var_linea)) < 117 Then
                                 For var_j = 1 + Len(Trim(var_linea)) To 117
                                     var_x = var_j Mod 2
                                     If var_x >= 1 Then
                                        var_linea = " " + var_linea
                                     Else
                                        var_linea = var_linea + " "
                                     End If
                                 Next var_j
                              End If
                              
                              If Len(Trim(var_rfc)) = 0 Then
                                 var_subimporte = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                 If Len(Trim(var_subimporte)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_subimporte)) To 14
                                        var_subimporte = " " + var_subimporte
                                    Next var_j
                                 End If
                                 var_iva = "-"
                                 For var_j = 1 + Len(Trim(var_iva)) To 11
                                     var_iva = " " + var_iva
                                  Next var_j
                              Else
                                 var_subimporte = Format(((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                 If Len(Trim(var_subimporte)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_subimporte)) To 14
                                        var_subimporte = " " + var_subimporte
                                    Next var_j
                                 End If
                                 var_iva = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                 If Len(Trim(var_iva)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_iva)) To 14
                                        var_iva = " " + var_iva
                                    Next var_j
                                 End If
                              End If
                              
                              If Len(Trim(var_subimporte)) < 14 Then
                                 For var_j = 1 + Len(Trim(var_subimporte)) To 14
                                     var_subimporte = " " + var_subimporte
                                 Next var_j
                              End If
                              var_espacios = 131 - Len(var_cantidad_total_str)
                              var_cantidad_total_str = Trim(var_cantidad_total_str)
                              If Len(Trim(var_cantidad_total_str)) < 14 Then
                                 For var_j = 1 + Len(Trim(var_cantidad_total_str)) To 14
                                     var_cantidad_total_str = " " + var_cantidad_total_str
                                 Next var_j
                              End If
                              var_subimporte = Trim(var_subimporte)
                              If Len(Trim(var_subimporte)) < 24 Then
                                 For var_j = 1 + Len(Trim(var_subimporte)) To 24
                                     var_subimporte = " " + var_subimporte
                                 Next var_j
                              End If
                              
                              var_cantidad_total_str = var_linea + var_cantidad_total_str + "    " + var_subimporte
                              'Print #1, Spc(var_espacios); var_cantidad_total_str; Spc(8); var_subimporte
                              Print #1, var_cantidad_total_str
                              var_linea = "                                                                          ESTA FACTURA SERA PAGADA EN UNA SOLA EXHIBICION                        " + var_iva
                              Print #1, var_linea
                              var_dia = Day(rs!dtim_Car_fecha)
                              var_mes = Month(rs!dtim_Car_fecha)
                              var_año = Year(rs!dtim_Car_fecha)
                              If var_empresa = "31" Then
                                 var_linea = "                                                       " + CStr(var_dia) + "     " + CStr(var_mes)
                              Else
                                 var_linea = "                                                             " + CStr(var_dia) + "     " + CStr(var_mes)
                              End If
                              
                              If Len(var_linea) < 145 Then
                                 For var_j = 1 + Len(var_linea) To 145
                                     var_linea = var_linea + " "
                                 Next var_j
                              End If
                              
                              var_importe = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                              
                              If Len(Trim(var_importe)) < 14 Then
                                 For var_j = 1 + Len(Trim(var_importe)) To 14
                                     var_importe = " " + var_importe
                                 Next var_j
                              End If
                              
                              'var_linea = "                                                                   ESTA FACTURA SERA PAGADA EN UNA SOLA EXHIBICION                               " + var_iva
                              'var_linea = "                                                                                                                                                 " + var_importe
                              
                              var_linea = var_linea + var_importe
                              Print #1, var_linea
                              
                              var_linea = var_importe
                              If Len(Trim(var_linea)) < 20 Then
                                 For var_j = 1 + Len(Trim(var_linea)) To 20
                                     var_linea = " " + var_linea
                                 Next var_j
                              End If
                              var_linea = var_linea + " " + var_cantidad_letra
                              Print #1, Spc(2); CStr(var_año); var_linea
                              
                              var_linea = ""
                              Print #1, ""
                              Print #1, ""
                              If var_empresa = "31" Then
                                 Print #1, Spc(10); Trim(IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE))
                                 Print #1, Spc(10); Trim(IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " " + IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre))
                                 Print #1, Spc(10); Trim(IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre) + " " + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre))
                              Else
                                 Print #1, Spc(5); Trim(IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE))
                                 Print #1, Spc(5); Trim(IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " " + IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre))
                                 Print #1, Spc(5); Trim(IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre) + " " + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre))
                              End If
                              If var_empresa <> "03" Then
                                 Print #1, ""
                                 Print #1, ""
                              Else
                                 Print #1, ""
                                 Print #1, ""
                              End If
                              Print #1, ""
                              Print #1, ""
                              
                              Close #1
                              Open (App.Path & "\factura" + Trim(Str(var_factura_nueva)) + ".bat") For Output As #2
                              var_Archivo = App.Path & "\factura" + Trim(Str(var_factura_nueva)) + ".bat"
                              If Trim(var_empresa) = "02" Or var_empresa = "18" Then
                                 Print #2, "copy " + App.Path + "\factura" + Trim(Str(var_factura_nueva)) + ".txt lpt1"
                              Else
                                 Print #2, "copy " + App.Path + "\factura" + Trim(Str(var_factura_nueva)) + ".txt lpt1"
                              End If
                              Close #2
                              x = Shell(var_Archivo, vbHide)
End If


If var_empresa = "03" Then
                              Open (App.Path & "\factura" + Trim(Str(var_factura_nueva)) + ".txt") For Output As #1
                              Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
                              'Print #1, Chr(27) + Chr(67) + Chr(44) + Chr(13)
                              'Print #1, ""
                              Print #1, Spc(105); Str(var_factura_nueva)
                              Print #1, ""
                              Print #1, ""
                              Print #1, Spc(105); Str(rs!INTE_CAR_PLAZO) + " DIAS DE VENCIMIENTO" + "                  " + Format(rs!dtim_Car_fecha, "Short Date")
                              Print #1, ""
                              'Print #1, Spc(92); Str(rs!inte_car_PLAZO) + " DIAS DE VENCIMIENTO"
                              var_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                              For var_j = 1 + Len(Trim(var_cliente)) To 83
                                  var_cliente = var_cliente + " "
                              Next var_j
                              var_cliente = var_cliente + "               AGUASCALIENTES, AGS."
                              Print #1, Spc(10); var_cliente
                              var_domicilio = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " C.P. " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
                              For var_j = 1 + Len(Trim(var_domicilio)) To 83
                                  var_domicilio = var_domicilio + " "
                              Next var_j
                              var_agente = ""
                              var_agente = IIf(IsNull(rs!vcha_AGE_aGENTE_ID), "", rs!vcha_AGE_aGENTE_ID)
                              For var_j = 1 + Len(Trim(var_agente)) To 8
                                  var_agente = var_agente + " "
                              Next var_j
                              var_agente = var_agente + IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
                              var_domicilio = var_domicilio
                              'Print #1, Spc(111); var_agente
                              Print #1, Spc(10); var_domicilio
                              var_ciudad = ""
                              var_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                              For var_j = 1 + Len(Trim(var_ciudad)) To 37
                                 var_ciudad = var_ciudad + " "
                              Next var_j
                              
                              var_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                              var_ciudad = var_ciudad
                              var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                              var_ciudad = var_ciudad + var_rfc
                              
                              For var_j = 1 + Len(Trim(var_estado)) To 46
                                 var_estado = var_estado + " "
                              Next var_j
                              

                              For var_j = 1 + Len(Trim(var_ciudad)) To 14
                                 var_ciudad = var_ciudad + " "
                              Next var_j
                               
                              var_ciudad = var_ciudad + "                                                      " + var_agente
                              
                              VAR_EMBARQUE = "EMB.: " + CStr(var_numero_embarque)
                              var_ordern_surtido = x
                              Print #1, Spc(10); var_ciudad
                              var_rfc = "RFC:  " + var_rfc
                              var_rfc = IIf(IsNull(rs!VCHA_ESB_ESTABLECIMIENTO_ID), "", rs!VCHA_ESB_ESTABLECIMIENTO_ID) + " " + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre) + ", " + IIf(IsNull(rs!vcha_pai_nombre), "", rs!vcha_pai_nombre)
                              For var_j = 1 + Len(Trim(var_rfc)) To 89
                                 var_rfc = var_rfc + " "
                              Next var_j
                              var_rfc = var_rfc + "                   PED.: " + Trim(Str(IIf(IsNull(rs!inte_ped_numero), 0, rs!inte_ped_numero))) + " "
                              var_rfc = var_rfc + " O.S.: " + Trim(Str(IIf(IsNull(rs!INTE_ORS_ORDEN_SURTIDO), 0, rs!INTE_ORS_ORDEN_SURTIDO))) + " " + VAR_EMBARQUE
                              Print #1, var_rfc
                              'Print #1, Spc(10); IIf(IsNull(rs!vcha_esb_establecimiento_id), "", rs!vcha_esb_establecimiento_id)
                              Print #1, ""
                              Print #1, ""
                              var_importe_descuento_1 = 0
                              var_importe_descuento_2 = 0
                              var_importe_descuento_3 = 0
                              var_contador_promociones = 0
                              var_cantidad_total = 0
                              For var_k = 1 To var_renglones_factura
                                 If Not rs.EOF Then
                                    var_linea = ""
                                    var_marca_promocion = " "
                                    var_promocion_1 = IIf(IsNull(rs!floa_sal_promocion_1), 0, rs!floa_sal_promocion_1)
                                    var_promocion_2 = IIf(IsNull(rs!FLOA_SAL_PROMOCION_2), 0, rs!FLOA_SAL_PROMOCION_2)
                                    If var_promocion_1 > 0 Then
                                       var_marca_promocion = "*"
                                       var_contador_promociones = var_contador_promociones + 1
                                    End If
                                    If var_promocion_2 > 0 Then
                                       var_marca_promocion = "*"
                                       var_contador_promociones = var_contador_promociones + 1
                                    End If
                                    var_linea = IIf(IsNull(rs!vcha_agr_agrupador_id), "", rs!vcha_agr_agrupador_id)
                                    For var_j = 1 + Len(Trim(var_linea)) To 15
                                        var_linea = var_linea + " "
                                    Next var_j
                                    var_linea = var_linea + IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura)
                                    var_i = 0
                                    While Len((var_linea)) < 115
                                          var_linea = var_linea + " "
                                    Wend
                                    var_linea = var_linea + " "
                                    var_linea = var_linea + var_marca_promocion
                                    var_cantidad = Format(IIf(IsNull(rs!Cantidad), 0, rs!Cantidad), "###,###,##0.00")
                                    var_cantidad_total = var_cantidad_total + IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)
                                    If Len(Trim(var_cantidad)) < 14 Then
                                       For var_j = 1 + Len(Trim(var_cantidad)) To 14
                                          var_cantidad = " " + var_cantidad
                                       Next var_j
                                    End If
                                    var_precio = IIf(IsNull(rs!Importe), 0, rs!Importe)
                                    var_descuento_1 = IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1)
                                    var_descuento_2 = IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2)
                                    var_descuento_3 = IIf(IsNull(rs!floa_car_porcentaje_descuento_3), 0, rs!floa_car_porcentaje_descuento_3)
                                    var_porcentaje = (100 - var_descuento_1) / 100
                                    var_precio = var_precio * var_porcentaje
                                    var_importe_descuento_1_2 = (IIf(IsNull(rs!Importe), 0, rs!Importe) - var_precio)
                                    var_importe_descuento_1 = var_importe_descuento_1 + (IIf(IsNull(rs!Importe), 0, rs!Importe) - var_precio)
                                    var_precio = var_precio * ((100 - var_descuento_2) / 100)
                                    var_importe_descuento_2 = var_importe_descuento_2 + (IIf(IsNull(rs!Importe), 0, rs!Importe) - (var_importe_descuento_1_2 + var_precio))
                                    var_precio = var_precio * ((100 - var_descuento_3) / 100)
                                    var_precio = var_precio / IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)
                                    'var_precio_str = Format(var_precio / IIf(IsNull(rs!cantidad), 0, rs!cantidad), "###,###,##0.00")
                                    var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                                    If Len(Trim(var_rfc)) > 0 Then
                                       var_importe_precio = IIf(IsNull(rs!Importe), 0, rs!Importe) / IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)
                                       var_importe_precio = var_importe_precio * ((100 - var_descuento_1) / 100)
                                       var_importe_precio = var_importe_precio * ((100 - var_descuento_2) / 100)
                                       var_importe_precio = var_importe_precio * ((100 - var_descuento_3) / 100)
                                       var_precio_str = Format(var_importe_precio, "###,###,##0.00")
                                       
                                    Else
                                       var_importe_precio = (IIf(IsNull(rs!Importe), 0, rs!Importe) / IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)) * (1 + (rs!floa_car_porcentaje_iva / 100))
                                       var_importe_precio = var_importe_precio * ((100 - var_descuento_1) / 100)
                                       var_importe_precio = var_importe_precio * ((100 - var_descuento_2) / 100)
                                       var_importe_precio = var_importe_precio * ((100 - var_descuento_3) / 100)
                                       var_precio_str = Format(var_importe_precio, "###,###,##0.00")
                                    End If
                                    If Len(Trim(var_precio_str)) < 14 Then
                                       For var_j = 1 + Len(Trim(var_precio_str)) To 14
                                           var_precio_str = " " + var_precio_str
                                       Next var_j
                                    End If
                                    var_linea = var_linea + var_cantidad + var_precio_str
                                    If Len(Trim(var_rfc)) > 0 Then
                                       var_importe_tot = Format((IIf(IsNull(rs!Importe), 0, rs!Importe)), "###,###,##0.00")
                                       var_importe_tot = var_importe_tot * ((100 - var_descuento_1) / 100)
                                       var_importe_tot = var_importe_tot * ((100 - var_descuento_2) / 100)
                                       var_importe_tot = var_importe_tot * ((100 - var_descuento_3) / 100)
                                       var_importe = Format(var_importe_tot, "###,###,##0.00")
                                       If Len(Trim(var_importe)) < 14 Then
                                           For var_j = 1 + Len(Trim(var_importe)) To 14
                                              var_importe = " " + var_importe
                                           Next var_j
                                       End If
                                    Else
                                       var_importe_tot = Format((IIf(IsNull(rs!Importe), 0, rs!Importe) * (1 + (rs!floa_car_porcentaje_iva / 100))), "###,###,##0.00")
                                       var_importe_tot = var_importe_tot * ((100 - var_descuento_1) / 100)
                                       var_importe_tot = var_importe_tot * ((100 - var_descuento_2) / 100)
                                       var_importe_tot = var_importe_tot * ((100 - var_descuento_3) / 100)
                                       var_importe = Format(var_importe_tot, "###,###,##0.00")
                                       If Len(Trim(var_importe)) < 14 Then
                                           For var_j = 1 + Len(Trim(var_importe)) To 14
                                              var_importe = " " + var_importe
                                           Next var_j
                                       End If
                                    End If
                                    var_linea = var_linea + var_importe
                                     
                                    Print #1, var_linea
                                    rs.MoveNext
                                 Else
                                    Print #1, ""
                                 End If
                              Next var_k
                              Print #1, ""
                              'Print #1, ""
                              rs.MoveFirst
                              
                              var_cantidad_total_str = Format(var_cantidad_total, "###,###,##0.00")
                              var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                              If Len(Trim(var_rfc)) > 0 Then
                                 var_cantidad_letra = rs!vcha_car_importe_letra
                                 'var_importe_descuento_1_str = Format(IIf(IsNull(rs!FLOA_CAR_IMPORTE_DESCUENTO_1), 0, rs!FLOA_CAR_IMPORTE_DESCUENTO_1) / (rs!floa_car_tipo_cambio), "###,###,##0.00")
                                 var_importe_descuento_1_str = "0"
                                 If Len(Trim(var_importe_descuento_1_str)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_importe_descuento_1_str)) To 14
                                         var_importe_descuento_1_str = " " + var_importe_descuento_1_str
                                    Next var_j
                                 End If
                                 'var_importe_descuento_2_str = Format(IIf(IsNull(rs!FLOA_CAR_IMPORTE_DESCUENTO_2), 0, rs!FLOA_CAR_IMPORTE_DESCUENTO_2) / (rs!floa_car_tipo_cambio), "###,###,##0.00")
                                 var_importe_descuento_2_str = "0"
                                 If Len(Trim(var_importe_descuento_2_str)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_importe_descuento_2_str)) To 14
                                        var_importe_descuento_2_str = " " + var_importe_descuento_2_str
                                    Next var_j
                                 End If
                              Else
                                 var_cantidad_letra = rs!vcha_car_importe_letra
                                 'var_importe_descuento_1_str = Format((IIf(IsNull(rs!FLOA_CAR_IMPORTE_DESCUENTO_1), 0, rs!FLOA_CAR_IMPORTE_DESCUENTO_1)) * (1 + (rs!floa_car_porcentaje_iva / 100)) / rs!floa_car_tipo_cambio, "###,###,##0.00")
                                 var_importe_descuento_1_str = "0"
                                 If Len(Trim(var_importe_descuento_1_str)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_importe_descuento_1_str)) To 14
                                         var_importe_descuento_1_str = " " + var_importe_descuento_1_str
                                    Next var_j
                                 End If
                                 'var_importe_descuento_2_str = Format((IIf(IsNull(rs!FLOA_CAR_IMPORTE_DESCUENTO_2), 0, rs!FLOA_CAR_IMPORTE_DESCUENTO_2)) * (1 + (rs!floa_car_porcentaje_iva / 100)) / rs!floa_car_tipo_cambio, "###,###,##0.00")
                                 var_importe_descuento_2_str = "0"
                                 If Len(Trim(var_importe_descuento_2_str)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_importe_descuento_2_str)) To 14
                                        var_importe_descuento_2_str = " " + var_importe_descuento_2_str
                                    Next var_j
                                 End If
                              End If
                              var_linea = "- DESCUENTO DEL " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1))) + "%"
                              If Len(Trim(var_linea)) < 145 Then
                                 For var_j = 1 + Len(Trim(var_linea)) To 145
                                     var_linea = var_linea + " "
                                 Next var_j
                              End If
                              Print #1, var_linea + var_importe_descuento_1_str
                              If var_empresa = "18" Then
                                 var_linea = ""
                              Else
                                 If Trim(var_cliente_coppel) = "C000002947" Then
                                    var_linea = ""
                                 Else
                                    var_linea = "- DESCUENTO POR PAGO OPORTUNO " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2))) + "%"
                                 End If
                              End If
                              If Len(Trim(var_linea)) < 145 Then
                                 For var_j = 1 + Len(Trim(var_linea)) To 145
                                     var_linea = var_linea + " "
                                 Next var_j
                              End If
                              var_linea = var_linea + var_importe_descuento_2_str
                              Print #1, var_linea
                              If var_contador_promociones > 0 Then
                                 Print #1, "PROMOCION EN ARTICULOS MARCADOS CON *"
                              Else
                                 Print #1, ""
                              End If
                              
                              var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                              var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
                              
                              If Len(Trim(var_linea)) < 117 Then
                                 For var_j = 1 + Len(Trim(var_linea)) To 117
                                     var_x = var_j Mod 2
                                     If var_x >= 1 Then
                                        var_linea = " " + var_linea
                                     Else
                                        var_linea = var_linea + " "
                                     End If
                                 Next var_j
                              End If
                              
                              If Len(Trim(var_rfc)) = 0 Then
                                 var_subimporte = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                 If Len(Trim(var_subimporte)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_subimporte)) To 14
                                        var_subimporte = " " + var_subimporte
                                    Next var_j
                                 End If
                                 var_iva = "-"
                                 For var_j = 1 + Len(Trim(var_iva)) To 11
                                     var_iva = " " + var_iva
                                  Next var_j
                              Else
                                 var_subimporte = Format(((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                 If Len(Trim(var_subimporte)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_subimporte)) To 14
                                        var_subimporte = " " + var_subimporte
                                    Next var_j
                                 End If
                                 var_iva = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                 If Len(Trim(var_iva)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_iva)) To 14
                                        var_iva = " " + var_iva
                                    Next var_j
                                 End If
                              End If
                              
                              If Len(Trim(var_subimporte)) < 14 Then
                                 For var_j = 1 + Len(Trim(var_subimporte)) To 14
                                     var_subimporte = " " + var_subimporte
                                 Next var_j
                              End If
                              var_espacios = 131 - Len(var_cantidad_total_str)
                              var_cantidad_total_str = Trim(var_cantidad_total_str)
                              If Len(Trim(var_cantidad_total_str)) < 14 Then
                                 For var_j = 1 + Len(Trim(var_cantidad_total_str)) To 14
                                     var_cantidad_total_str = " " + var_cantidad_total_str
                                 Next var_j
                              End If
                              var_subimporte = Trim(var_subimporte)
                              If Len(Trim(var_subimporte)) < 24 Then
                                 For var_j = 1 + Len(Trim(var_subimporte)) To 24
                                     var_subimporte = " " + var_subimporte
                                 Next var_j
                              End If
                              
                              var_cantidad_total_str = var_linea + var_cantidad_total_str + "    " + var_subimporte
                              'Print #1, Spc(var_espacios); var_cantidad_total_str; Spc(8); var_subimporte
                              Print #1, var_cantidad_total_str
                              var_linea = "                                                                          ESTA FACTURA SERA PAGADA EN UNA SOLA EXHIBICION                        " + var_iva
                              Print #1, var_linea
                              var_dia = Day(rs!dtim_Car_fecha)
                              var_mes = Month(rs!dtim_Car_fecha)
                              var_año = Year(rs!dtim_Car_fecha)
                              
                              var_linea = "                                                             " + CStr(var_dia) + "     " + CStr(var_mes)
                              
                              If Len(var_linea) < 145 Then
                                 For var_j = 1 + Len(var_linea) To 145
                                     var_linea = var_linea + " "
                                 Next var_j
                              End If
                              
                              var_importe = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                              
                              If Len(Trim(var_importe)) < 14 Then
                                 For var_j = 1 + Len(Trim(var_importe)) To 14
                                     var_importe = " " + var_importe
                                 Next var_j
                              End If
                              
                              'var_linea = "                                                                   ESTA FACTURA SERA PAGADA EN UNA SOLA EXHIBICION                               " + var_iva
                              'var_linea = "                                                                                                                                                 " + var_importe
                              
                              var_linea = var_linea + var_importe
                              Print #1, var_linea
                              
                              var_linea = var_importe
                              If Len(Trim(var_linea)) < 20 Then
                                 For var_j = 1 + Len(Trim(var_linea)) To 20
                                     var_linea = " " + var_linea
                                 Next var_j
                              End If
                              var_linea = var_linea + " " + var_cantidad_letra
                              Print #1, Spc(2); CStr(var_año); var_linea
                              
                              var_linea = ""
                              Print #1, ""
                              Print #1, ""
                              Print #1, Spc(5); Trim(IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE))
                              Print #1, Spc(5); Trim(IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " " + IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre))
                              Print #1, Spc(5); Trim(IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre) + " " + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre))
                              If var_empresa <> "03" Then
                                 Print #1, ""
                                 Print #1, ""
                              Else
                                 Print #1, ""
                                 Print #1, ""
                              End If
                              Print #1, ""
                              Print #1, ""
                              
                              Close #1
                              Open (App.Path & "\factura" + Trim(Str(var_factura_nueva)) + ".bat") For Output As #2
                              var_Archivo = App.Path & "\factura" + Trim(Str(var_factura_nueva)) + ".bat"
                              If Trim(var_empresa) = "02" Or var_empresa = "18" Then
                                 Print #2, "copy " + App.Path + "\factura" + Trim(Str(var_factura_nueva)) + ".txt lpt1"
                              Else
                                 Print #2, "copy " + App.Path + "\factura" + Trim(Str(var_factura_nueva)) + ".txt lpt1"
                              End If
                              Close #2
                              x = Shell(var_Archivo, vbHide)
End If


End Sub
Private Sub reimprimir_notas_credito()
   Dim var_rfc As String
   var_serie = Me.txt_serie
   rs.Open "select * from tb_principal where vcha_emp_empresa_id = '" + var_empresa + "'"
   var_numero_renglones = rs!INTE_PRI_RENGLONES_NOTA_CREDITO
   var_tolerancia_saldo = rs!FLOA_PRI_TOLERANCIA_SALDOS
   rs.Close
   si = MsgBox("¿Deseas reimprimir la nota de crédito " + txt_numero + "?", vbYesNo, "ATENCION")
   If si = 6 Then
      rs.Open "select * from tb_encabezado_cartera where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_tipo_documento = '" + txt_documento + "' and vcha_Ser_serie_id = '" + var_serie + "' and inte_Car_numero = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         rsaux2.Open "select * from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
         var_factura_nueva = rsaux2!inte_ser_nota_credito
         rsaux2.Close
         si = MsgBox("¿Se va a imprimir la nota de crédito " + Str(var_factura_nueva) + "?", vbYesNo, "ATENCION")
         If si = 6 Then
            si = MsgBox("Confirmar la reimpresión de la nota de crédito " + Str(var_factura_nueva), vbYesNo, "ATENCION")
            If si = 6 Then
               rsaux2.Open "select * from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
               If var_factura_nueva <> rsaux2!inte_ser_nota_credito Then
                  MsgBox "El número de la nota de crédito ya cambio y el proceso de reimpresión a cambiado", vbOKOnly, "ATENCION"
                  rsaux2.Close
               Else
                  rsaux2.Close
                  var_cadena = ""
                  var_cadena = "INSERT INTO [TB_ENCABEZADO_CARTERA] ([VCHA_EMP_EMPRESA_ID], [VCHA_UOR_UNIDAD_ID], [VCHA_CAR_TIPO_DOCUMENTO], [VCHA_CAR_DOCUMENTO], [VCHA_CAR_CLASE_ID], [INTE_CAR_NUMERO], [CHAR_CAR_AFECTACION], [VCHA_ALM_ALMACEN_ID], [VCHA_MOV_MOVIMIENTO_ID], [INTE_EMO_NUMERO], [DTIM_CAR_FECHA], [VCHA_AGE_AGENTE_ID], [VCHA_GAC_GRUPO_ACTUAL_ID], [VCHA_GRE_GRUPO_REAL_ID], [VCHA_TIT_TITULAR_ID], [VCHA_CLI_CLAVE_ID], [VCHA_ESB_ESTABLECIMIENTO_ID], [INTE_CAR_PLAZO], [FLOA_CAR_PORCENTAJE_IVA], [FLOA_CAR_PORCENTAJE_IMPUESTO_1], [FLOA_CAR_PORCENTAJE_IMPUESTO_2], [FLOA_CAR_PORCENTAJE_DESCUENTO_1], [FLOA_CAR_PORCENTAJE_DESCUENTO_2], [FLOA_CAR_PORCENTAJE_DESCUENTO_3], [FLOA_CAR_IMPORTE_TOTAL], [FLOA_CAR_IMPORTE_IVA], [FLOA_CAR_IMPORTE_IMPUESTO_1], [FLOA_CAR_IMPORTE_IMPUESTO_2], [FLOA_CAR_IMPORTE_DESCUENTO_1], [FLOA_CAR_IMPORTE_DESCUENTO_2], [FLOA_CAR_IMPORTE_DESCUENTO_3], [FLOA_CAR_SUBIMPORTE], [FLOA_CAR_IMPORTE_NETO], [VCHA_CAR_IMPORTE_LETRA], [VCHA_AUD_USUARIO], [VCHA_AUD_MAQUINA],"
                  var_cadena = var_cadena + "[VCHA_AUD_FECHA], [FLOA_CAR_SALDO], [DTIM_CAR_FECHA_VENCIMIENTO], [DTIM_CAR_FECHA_ENTREGA], [VCHA_MON_MONEDA_ID], [FLOA_CAR_TIPO_CAMBIO], [VCHA_SER_SERIE_ID], [CHAR_CAR_ESTATUS], [CHAR_CAR_TIPO_FACTURACION]) Values ('" + IIf(IsNull(rs!VCHA_EMP_EMPRESA_ID), "", rs!VCHA_EMP_EMPRESA_ID) + "', '" + rs!VCHA_UOR_UNIDAD_ID + "', 'NC', '" + rs!vcha_Car_documento + "', '" + rs!vcha_Car_clase_id + "', " + CStr(var_factura_nueva) + ", '" + rs!char_car_afectacion
                  var_cadena = var_cadena + "', '" + rs!vcha_alm_almacen_id + "', '" + rs!vcha_mov_movimiento_id + "', " + CStr(rs!INTE_EMO_NUMERO) + ", getdate(),  '" + rs!vcha_AGE_aGENTE_ID + "', '" + rs!VCHA_GAC_GRUPO_aCTUAL_ID + "', '" + rs!vcha_gre_grupo_real_id + "', '" + rs!vcha_tit_titular_id + "', '" + rs!vcha_cli_clave_id + "', '" + rs!VCHA_ESB_ESTABLECIMIENTO_ID + "', " + CStr(rs!INTE_CAR_PLAZO) + ", " + CStr(rs!floa_car_porcentaje_iva) + ", " + CStr(rs!FLOA_CAR_PORCENTAJE_IMPUESTO_1) + ", " + CStr(rs!FLOA_CAR_PORCENTAJE_IMPUESTO_2) + ", " + CStr(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1) + ", " + CStr(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2) + ", " + CStr(rs!floa_car_porcentaje_descuento_3) + ", " + CStr(rs!FLOA_CAR_IMPORTE_TOTAL) + ", " + CStr(rs!floa_car_importe_iva) + ", " + CStr(rs!FLOA_CAR_IMPORTE_IMPUESTO_1) + ", " + CStr(rs!FLOA_CAR_IMPORTE_IMPUESTO_2) + ", " + CStr(rs!FLOA_CAR_IMPORTE_DESCUENTO_1) + ", " + CStr(rs!FLOA_CAR_IMPORTE_DESCUENTO_2) + ", "
                  var_cadena = var_cadena + CStr(rs!FLOA_CAR_IMPORTE_DESCUENTO_3) + "," + CStr(rs!floa_car_subimporte) + ", " + CStr(rs!floa_Car_importe_neto) + ", '" + rs!vcha_car_importe_letra + "', '" + rs!vcha_aud_usuario + "', '" + rs!vcha_aud_maquina + "', getdate(), 0, null, null, '" + rs!vcha_mon_moneda_id + "', " + CStr(rs!floa_car_tipo_cambio) + ", '" + rs!VCHA_SER_SERIE_ID + "', '', 'V') "
                  cnn.BeginTrans
                  rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  rsaux3.Open "select * from tb_estado_cuenta where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_ecu_serie_abono = '" + var_serie + "' and vcha_ecu_movimiento_abono = '" + var_tipo_nota_credito + "' and inte_ecu_numero_abono = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux3.EOF Then
                     While Not rsaux3.EOF
                           var_cadena = ""
                           var_cadena = "INSERT INTO TB_ESTADO_CUENTA ([VCHA_EMP_EMPRESA_ID], [VCHA_UOR_UNIDAD_ID], [VCHA_ECU_SERIE_CARGO], [VCHA_ECU_MOVIMIENTO_CARGO], [INTE_ECU_NUMERO_CARGO], [VCHA_ECU_SERIE_ABONO], [VCHA_ECU_MOVIMIENTO_ABONO], [INTE_ECU_NUMERO_ABONO], [FLOA_ECU_IMPORTE_CARGO], [FLOA_ECU_IMPORTE_ABONO]) values ( '" + IIf(IsNull(rsaux3!VCHA_EMP_EMPRESA_ID), "", rsaux3!VCHA_EMP_EMPRESA_ID) + "', '" + IIf(IsNull(rsaux3!VCHA_UOR_UNIDAD_ID), "", rsaux3!VCHA_UOR_UNIDAD_ID) + "', '" + IIf(IsNull(rsaux3!vcha_ecu_serie_cargo), "", rsaux3!vcha_ecu_serie_cargo) + "',  '" + IIf(IsNull(rsaux3!vcha_ecu_movimiento_cargo), "", rsaux3!vcha_ecu_movimiento_cargo) + "', " + CStr(IIf(IsNull(rsaux3!inte_ecu_numero_cargo), 0, rsaux3!inte_ecu_numero_cargo)) + ",  '" + IIf(IsNull(rsaux3!VCHA_ECU_SERIE_ABONO), "", rsaux3!VCHA_ECU_SERIE_ABONO) + "', '" + IIf(IsNull(rsaux3!VCHA_ECU_MOVIMIENTO_ABONO), "", rsaux3!VCHA_ECU_MOVIMIENTO_ABONO) + "', "
                           var_cadena = var_cadena + CStr(var_factura_nueva)
                           var_cadena = var_cadena + ", " + CStr(IIf(IsNull(rsaux3!floa_ecu_importe_cargo), 0, rsaux3!floa_ecu_importe_cargo)) + ", " + CStr(IIf(IsNull(rsaux3!floa_ecu_importe_abono), 0, rsaux3!floa_ecu_importe_abono)) + ")"
                           rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                           rsaux3.MoveNext
                     Wend
                  End If
                  If var_tipo_nota_credito = "DF" Then
                     rsaux2.Open "update TB_DETALLE_DESCUENTOS_FINANCIEROS set inte_Car_numero = " + CStr(var_factura_nueva) + " where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'DF' and inte_Car_numero = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
                  End If
                  If var_tipo_nota_credito = "" Then
                     rsaux2.Open "update TB_DETALLE_BONIFICACION_FINANCIERA set inte_Car_numero = " + CStr(var_factura_nueva) + " where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'BF' and inte_Car_numero = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
                  End If
                  If var_tipo_nota_credito = "BO" Or var_tipo_nota_credito = "BF" Then
                     rsaux2.Open "update TB_DETALLE_BONIFICACIONES set inte_Car_numero = " + CStr(var_factura_nueva) + " where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = '" + var_tipo_nota_credito + "' and inte_Car_numero = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
                  End If
                  If var_tipo_nota_credito = "DV" Then
                     rsaux2.Open "update TB_DETALLE_DEVOLUCION_IMPORTES_ASIGNADOS set inte_Car_numero = " + CStr(var_factura_nueva) + " where vcha_Emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'DV' and inte_Car_numero = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux2.Open "INSERT INTO TB_SECUENCIA_NOTAS_CREDITO (VCHA_EMP_EMPRESA_ID, VCHA_SER_SERIE_ID, INTE_SNC_NUMERO_ANTERIOR, INTE_SNC_NUMERO_ACTUAL) VALUES ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!VCHA_SER_SERIE_ID + "', " + CStr(var_factura_nueva) + ", " + CStr(var_factura_nueva) + ")", cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.Open "UPDATE TB_SECUENCIA_NOTAS_CREDITO SET INTE_SNC_NUMERO_ACTUAL = " + CStr(var_factura_nueva) + " WHERE VCHA_EMP_EMPRESA_ID = '" + rs!VCHA_EMP_EMPRESA_ID + "' AND VCHA_SER_SERIE_ID = '" + rs!VCHA_SER_SERIE_ID + "' AND  INTE_SNC_NUMERO_ANTERIOR = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.Open "update tb_series set inte_ser_nota_credito = isnull(inte_ser_nota_credito,0) + 1 where vcha_emp_empresa_id = '" + rs!VCHA_EMP_EMPRESA_ID + "' and vcha_uor_unidad_id = '" + rs!VCHA_UOR_UNIDAD_ID + "' and vcha_Ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                  rsaux3.Close
                  cnn.CommitTrans
               End If
            End If
         End If
      End If
      rs.Close
   End If
   If var_tipo_nota_credito = "DF" Then
      rs.Open "select * from VW_DOCUMENTOS_IMPRESION where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'DF' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(var_factura_nueva), cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Open (App.Path & "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".txt") For Output As #1
         Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
         Print #1, Spc(92); Str(rs!INTE_CAR_NUMERO)
         Print #1, ""
         Print #1, Spc(93); "FECHA: "; Format(rs!dtim_Car_fecha, "Short Date")
         var_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
         For var_j = 1 + Len(Trim(var_cliente)) To 83
             var_cliente = var_cliente + " "
         Next var_j
         var_cliente = var_cliente + "AGUASCALIENTES, AGS."
         Print #1, ""
         Print #1, Spc(10); var_cliente
         var_domicilio = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " C.P. " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
         For var_j = 1 + Len(Trim(var_domicilio)) To 83
             var_domicilio = var_domicilio + " "
         Next var_j
         var_agente = ""
         var_agente = IIf(IsNull(rs!vcha_AGE_aGENTE_ID), "", rs!vcha_AGE_aGENTE_ID)
         For var_j = 1 + Len(Trim(var_agente)) To 8
             var_agente = var_agente + " "
         Next var_j
         var_agente = var_agente + IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
         var_domicilio = var_domicilio
         Print #1, Spc(10); var_domicilio
         var_ciudad = ""
         var_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
         For var_j = 1 + Len(Trim(var_ciudad)) To 37
             var_ciudad = var_ciudad + " "
         Next var_j
         var_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
         For var_j = 1 + Len(Trim(var_estado)) To 46
             var_estado = var_estado + " "
         Next var_j
         var_ciudad = var_ciudad + var_estado
                               
         For var_j = 1 + Len(Trim(var_ciudad)) To 14
             var_ciudad = var_ciudad + " "
         Next var_j
                       
         var_ciudad = var_ciudad + var_agente
                         
         Print #1, Spc(10); var_ciudad
         var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
         var_rfc = "RFC:  " + var_rfc
         For var_j = 1 + Len(Trim(var_rfc)) To 89
             var_rfc = var_rfc + " "
         Next var_j
         var_rfc = var_rfc
         Print #1, Spc(4); var_rfc
         Print #1, ""
         Print #1, ""
         Print #1, ""
         Print #1, ""
         If rsaux3.State = 1 Then
            rsaux3.Close
         End If
         var_iva = IIf(IsNull(rs!floa_car_porcentaje_iva), 0, rs!floa_car_porcentaje_iva)
         rsaux3.Open "select * from TB_DETALLE_DESCUENTOS_FINANCIEROS where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'DF' and vcha_ser_serie_id  = '" + var_serie + "' and vcha_car_clase_id = 'DF' and inte_car_numero =  " + Str(var_factura_nueva), cnn, adOpenDynamic, adLockOptimistic
         var_contador_lineas = 0
         var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
         While Not rsaux3.EOF
            var_linea = "DF" + Str(rs!INTE_CAR_NUMERO) + " " + rs!vcha_Car_nombre + " Factura " + Str(rsaux3!inte_ddf_factura) + " " + CStr(rsaux3!floa_ddf_descuento_otorgado) + "%"
            If Round(rsaux3!floa_ddf_descuento_otorgado, 2) <> Round(rsaux3!floa_ddf_descuento_aplicado, 2) Then
               var_linea = var_linea + " (" + Format(rsaux3!floa_ddf_descuento_aplicado, "###,###,##0.0000") + "%)"
            End If
            If Len(Trim(var_linea)) < 120 Then
               For var_j = 1 + Len(Trim(var_linea)) To 120
                   var_linea = var_linea + " "
               Next var_j
            End If
            If Len(Trim(var_rfc)) = 0 Then
               var_importe_str = Format((IIf(IsNull(rsaux3!FLOA_DDF_IMPORTE), 0, rsaux3!FLOA_DDF_IMPORTE)), "###,###,##0.00")
               If Len(Trim(var_importe_str)) < 14 Then
                  For var_j = 1 + Len(Trim(var_importe_str)) To 14
                      var_importe_str = " " + var_importe_str
                  Next var_j
               End If
            Else
               var_importe_str = Format((IIf(IsNull(rsaux3!FLOA_DDF_IMPORTE), 0, rsaux3!FLOA_DDF_IMPORTE) / (1 + (var_iva / 100))), "###,###,##0.00")
               If Len(Trim(var_importe_str)) < 14 Then
                  For var_j = 1 + Len(Trim(var_importe_str)) To 14
                      var_importe_str = " " + var_importe_str
                  Next var_j
               End If
            End If
            var_linea = var_linea + var_importe_str
            Print #1, var_linea
            rsaux3.MoveNext
            var_contador_lineas = var_contador_lineas + 1
         Wend
         If var_contador_lineas < var_numero_renglones Then
            For var_l = var_contador_lineas To var_numero_renglones
                Print #1, ""
            Next var_l
         End If
         rsaux3.Close
         var_cantidad_letra = rs!vcha_car_importe_letra
         var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
         If Len(Trim(var_linea)) < 105 Then
            For var_j = 1 + Len(Trim(var_linea)) To 105
                var_linea = var_linea + " "
            Next var_j
         End If
         If Len(Trim(var_rfc)) = 0 Then
            var_subimporte_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
            If Len(Trim(var_subimporte_str)) < 14 Then
               For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                   var_subimporte_str = " " + var_subimporte_str
               Next var_j
            End If
            var_iva = "-        "
            For var_j = 1 + Len(Trim(var_iva_str)) To 14
                var_iva_str = " " + var_iva_str
            Next var_j
         Else
            var_subimporte_str = Format(((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
            If Len(Trim(var_subimporte_str)) < 14 Then
               For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                   var_subimporte_str = " " + var_subimporte_str
               Next var_j
            End If
            var_iva_str = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
            If Len(Trim(var_iva_str)) < 14 Then
               For var_j = 1 + Len(Trim(var_iva_str)) To 14
                   var_iva_str = " " + var_iva_str
               Next var_j
            End If
         End If
         var_linea = var_linea + "           " + var_subimporte_str
         Print #1, Spc(4); var_linea
         Print #1, Spc(120); var_iva_str
         var_importe_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
         If Len(Trim(var_importe_str)) < 14 Then
            For var_j = 1 + Len(Trim(var_importe_str)) To 14
                var_importe_str = " " + var_importe_str
            Next var_j
         End If
         Print #1, Spc(120); var_importe_str
         Print #1, ""
         Print #1, ""
         Print #1, ""
         Print #1, Spc(85); "SISTEMAS"
         Close #1
                   
         Open (App.Path & "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".bat") For Output As #2
         var_Archivo = App.Path & "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".bat"
         Print #2, "copy " + App.Path + "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".txt lpt1"
         Close #2
         x = Shell(var_Archivo, vbHide)
      End If
      rs.Close
   End If
   If var_tipo_nota_credito = "" Then
      rs.Open "select * from VW_DOCUMENTOS_IMPRESION where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'BF' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + CStr(var_factura_nueva), cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
''''''''IMPRESION DE LA NOTA DE CARGO
         Open (App.Path & "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".txt") For Output As #1
         'Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
         Print #1, Chr(27) + Chr(64)
         Print #1, ""
         Print #1, Spc(92); Str(rs!INTE_CAR_NUMERO)
         Print #1, ""
         Print #1, Spc(92); "FECHA: "; Format(rs!dtim_Car_fecha, "Short Date")
         var_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
         For var_j = 1 + Len(Trim(var_cliente)) To 83
             var_cliente = var_cliente + " "
         Next var_j
         var_cliente = var_cliente + "AGUASCALIENTES, AGS."
         Print #1, ""
         Print #1, Spc(12); var_cliente
         var_domicilio = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " C.P. " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
         For var_j = 1 + Len(Trim(var_domicilio)) To 83
             var_domicilio = var_domicilio + " "
         Next var_j
         var_agente = ""
         var_agente = IIf(IsNull(rs!vcha_AGE_aGENTE_ID), "", rs!vcha_AGE_aGENTE_ID)
         For var_j = 1 + Len(Trim(var_agente)) To 8
             var_agente = var_agente + " "
         Next var_j
         var_agente = var_agente + IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
         var_domicilio = var_domicilio
         Print #1, Spc(12); var_domicilio
         var_ciudad = ""
         var_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
         For var_j = 1 + Len(Trim(var_ciudad)) To 37
             var_ciudad = var_ciudad + " "
         Next var_j
         var_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
         For var_j = 1 + Len(Trim(var_estado)) To 46
             var_estado = var_estado + " "
         Next var_j
         var_ciudad = var_ciudad + var_estado
                            
         For var_j = 1 + Len(Trim(var_ciudad)) To 14
             var_ciudad = var_ciudad + " "
         Next var_j
                        
         var_ciudad = var_ciudad + var_agente
                       
         Print #1, Spc(12); var_ciudad
         var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
         var_rfc = "      " + var_rfc
         For var_j = 1 + Len(Trim(var_rfc)) To 89
             var_rfc = var_rfc + " "
         Next var_j
         var_rfc = var_rfc
         Print #1, Spc(6); var_rfc
         Print #1, ""
         Print #1, ""
         Print #1, ""
         var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
         If rsaux.State = 1 Then
            rsaux.Close
         End If
                     
         rsaux.Open "select * from tb_detalle_bonificacion_financiera where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'BF' and vcha_ser_serie_id = '" + var_serie + "' and inte_car_numero = " + CStr(var_factura_nueva), cnn, adOpenDynamic, adLockOptimistic
         var_contador_renglones_nota = 0
         While Not rsaux.EOF
               var_linea = "BF" + Str(rs!INTE_CAR_NUMERO) + " " + rs!vcha_Car_nombre + " FACTURA " + Str(rsaux!inte_dbf_factura) + " " + Format(rsaux!floa_dbf_porcentaje, "###,###,##0.0000") + "%"
               If Len(Trim(var_linea)) < 120 Then
                  For var_j = 1 + Len(Trim(var_linea)) To 120
                      var_linea = var_linea + " "
                  Next var_j
               End If
               If Len(Trim(var_rfc)) = 0 Then
                  var_importe_str = Format((IIf(IsNull(rsaux!FLOA_DBF_IMPORTE), 0, rsaux!FLOA_DBF_IMPORTE)), "###,###,##0.00")
               Else
                  var_importe_str = Format(((IIf(IsNull(rsaux!FLOA_DBF_IMPORTE), 0, rsaux!FLOA_DBF_IMPORTE)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)) / (1 + ((IIf(IsNull(rsaux!floa_dbf_iva), 1, rsaux!floa_dbf_iva) / 100)))), "###,###,##0.00")
               End If
               If Len(Trim(var_importe_str)) < 14 Then
                  For var_j = 1 + Len(Trim(var_importe_str)) To 14
                      var_importe_str = " " + var_importe_str
                  Next var_j
               End If
               var_linea = var_linea + var_importe_str
               Print #1, var_linea
               rsaux.MoveNext
               var_contador_renglones_nota = var_contador_renglones_nota + 1
         Wend
         rsaux.Close
         If var_contador_renglones_nota < var_numero_renglones Then
            For var_l = var_contador_renglones_nota To var_numero_renglones
                Print #1, ""
            Next var_l
         End If
         var_cantidad_letra = rs!vcha_car_importe_letra
         var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
         If Len(Trim(var_linea)) < 105 Then
            For var_j = 1 + Len(Trim(var_linea)) To 105
                var_linea = var_linea + " "
            Next var_j
         End If
                     
                     
         If Len(Trim(var_rfc)) = 0 Then
            var_subimporte_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
            If Len(Trim(var_subimporte_str)) < 14 Then
               For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                   var_subimporte_str = " " + var_subimporte_str
               Next var_j
            End If
            var_iva = "      -        "
            For var_j = 1 + Len(Trim(var_iva_str)) To 14
                var_iva_str = " " + var_iva_str
            Next var_j
         Else
            var_subimporte_str = Format(((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
            If Len(Trim(var_subimporte_str)) < 14 Then
               For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                   var_subimporte_str = " " + var_subimporte_str
               Next var_j
            End If
            var_iva_str = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
            If Len(Trim(var_iva_str)) < 14 Then
               For var_j = 1 + Len(Trim(var_iva_str)) To 14
                   var_iva_str = " " + var_iva_str
               Next var_j
            End If
         End If
         var_linea = var_linea + "           " + var_subimporte_str
         Print #1, Spc(4); var_linea
         Print #1, Spc(120); var_iva_str
         var_importe_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
         If Len(Trim(var_importe_str)) < 14 Then
            For var_j = 1 + Len(Trim(var_importe_str)) To 14
                var_importe_str = " " + var_importe_str
            Next var_j
         End If
         Print #1, Spc(120); var_importe_str
         Print #1, ""
         Print #1, ""
         Print #1, ""
         Print #1, Spc(85); "SISTEMAS"
         Print #1, ""
         Print #1, ""
         Print #1, ""
         Print #1, ""
         Print #1, ""
         Print #1, ""
         Print #1, ""
         Close #1
               
         Open (App.Path & "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".bat") For Output As #2
         var_Archivo = App.Path & "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".bat"
         Print #2, "copy " + App.Path + "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".txt lpt1"
         Close #2
         x = Shell(var_Archivo, vbHide)
''''''''''''
      End If
      rs.Close
   End If
   If var_tipo_nota_credito = "BO" Or var_tipo_nota_credito = "BF" Then
      rs.Open "select * from VW_DOCUMENTOS_IMPRESION where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = '" + var_tipo_nota_credito + "' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(var_factura_nueva), cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Open (App.Path & "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".txt") For Output As #1
         Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
         Print #1, Spc(92); Str(rs!INTE_CAR_NUMERO)
         Print #1, ""
         Print #1, Spc(92); "FECHA: "; Format(rs!dtim_Car_fecha, "Short Date")
         var_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
         For var_j = 1 + Len(Trim(var_cliente)) To 83
             var_cliente = var_cliente + " "
         Next var_j
         var_cliente = var_cliente + "AGUASCALIENTES, AGS."
         Print #1, ""
         Print #1, Spc(12); var_cliente
         var_domicilio = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " C.P. " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
         For var_j = 1 + Len(Trim(var_domicilio)) To 83
             var_domicilio = var_domicilio + " "
         Next var_j
         var_agente = ""
         var_agente = IIf(IsNull(rs!vcha_AGE_aGENTE_ID), "", rs!vcha_AGE_aGENTE_ID)
         For var_j = 1 + Len(Trim(var_agente)) To 8
             var_agente = var_agente + " "
         Next var_j
         var_agente = var_agente + IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
         var_domicilio = var_domicilio
         Print #1, Spc(12); var_domicilio
         var_ciudad = ""
         var_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
         For var_j = 1 + Len(Trim(var_ciudad)) To 37
             var_ciudad = var_ciudad + " "
         Next var_j
         var_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
         For var_j = 1 + Len(Trim(var_estado)) To 46
             var_estado = var_estado + " "
         Next var_j
         var_ciudad = var_ciudad + var_estado
                             
         For var_j = 1 + Len(Trim(var_ciudad)) To 14
             var_ciudad = var_ciudad + " "
         Next var_j
                
         var_ciudad = var_ciudad + var_agente
                            
         Print #1, Spc(12); var_ciudad
         var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
         var_rfc = "      " + var_rfc
         For var_j = 1 + Len(Trim(var_rfc)) To 89
             var_rfc = var_rfc + " "
         Next var_j
         var_rfc = var_rfc
         Print #1, Spc(6); var_rfc
         Print #1, ""
         Print #1, ""
         Print #1, ""
         var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
         rsaux.Open "select * from tb_detalle_bonificaciones where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = '" + var_tipo_nota_credito + "' and vcha_ser_serie_id = '" + var_serie + "' and inte_car_numero = " + CStr(var_factura_nueva), cnn, adOpenDynamic, adLockOptimistic
         var_contador_renglones_nota = 0
         While Not rsaux.EOF
               var_linea = "BO" + Str(rs!INTE_CAR_NUMERO) + " " + rs!vcha_Car_nombre + " FACTURA " + CStr(rsaux!inte_car_factura)
               If Len(Trim(var_linea)) < 108 Then
                  For var_j = 1 + Len(Trim(var_linea)) To 120
                      var_linea = var_linea + " "
                  Next var_j
               End If
               If Len(Trim(var_rfc)) = 0 Then
                  var_importe_str = Format(((IIf(IsNull(rsaux!FLOA_dbo_IMPORTE), 0, rsaux!FLOA_dbo_IMPORTE)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio))), "###,###,##0.00")
               Else
                  var_importe_str = Format(((IIf(IsNull(rsaux!FLOA_dbo_IMPORTE), 0, rsaux!FLOA_dbo_IMPORTE)) / (1 + (rsaux!floa_dbo_iva / 100)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio))), "###,###,##0.00")
               End If
               If Len(Trim(var_importe_str)) < 14 Then
                  For var_j = 1 + Len(Trim(var_importe_str)) To 14
                      var_importe_str = " " + var_importe_str
                  Next var_j
               End If
               var_linea = var_linea + var_importe_str
               Print #1, var_linea
               rsaux.MoveNext
               var_contador_renglones_nota = var_contador_renglones_nota + 1
         Wend
         rsaux.Close
         If var_contador_renglones_nota < var_numero_renglones Then
            For var_l = var_contador_renglones_nota To var_numero_renglones
                Print #1, ""
            Next var_l
         End If
         var_cantidad_letra = rs!vcha_car_importe_letra
         var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
         If Len(Trim(var_linea)) < 105 Then
            For var_j = 1 + Len(Trim(var_linea)) To 105
                var_linea = var_linea + " "
            Next var_j
         End If
                   
                       
         If Len(Trim(var_rfc)) = 0 Then
            var_subimporte_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
            If Len(Trim(var_subimporte_str)) < 14 Then
               For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                   var_subimporte_str = " " + var_subimporte_str
               Next var_j
            End If
            var_iva = "      -        "
            For var_j = 1 + Len(Trim(var_iva_str)) To 14
                var_iva_str = " " + var_iva_str
            Next var_j
         Else
            var_subimporte_str = Format(((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
            If Len(Trim(var_subimporte_str)) < 14 Then
               For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                   var_subimporte_str = " " + var_subimporte_str
               Next var_j
            End If
            var_iva_str = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
            If Len(Trim(var_iva_str)) < 14 Then
               For var_j = 1 + Len(Trim(var_iva_str)) To 14
                   var_iva_str = " " + var_iva_str
               Next var_j
            End If
         End If
         var_linea = var_linea + "           " + var_subimporte_str
         Print #1, Spc(4); var_linea
         Print #1, Spc(120); var_iva_str
         var_importe_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
         If Len(Trim(var_importe_str)) < 14 Then
            For var_j = 1 + Len(Trim(var_importe_str)) To 14
                var_importe_str = " " + var_importe_str
            Next var_j
         End If
         Print #1, Spc(120); var_importe_str
         Print #1, ""
         Print #1, ""
         Print #1, ""
         Print #1, Spc(85); "SISTEMAS"
         Close #1
                       
         Open (App.Path & "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".bat") For Output As #2
         var_Archivo = App.Path & "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".bat"
         Print #2, "copy " + App.Path + "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".txt lpt1"
         Close #2
         x = Shell(var_Archivo, vbHide)
      End If
      rs.Close
   End If
   
   If var_tipo_nota_credito = "DV" Then
      rs.Open "select * from VW_DOCUMENTOS_IMPRESION where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'DV' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(var_factura_nueva), cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
''''' IMPRESION DE LA NOTA DE CARGO
         var_clave_movimiento = rs!vcha_mov_movimiento_id
         var_numero_movimiento = rs!INTE_EMO_NUMERO
         Open (App.Path & "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".txt") For Output As #1
         Print #1, Chr(27) + Chr(64)
         Print #1, ""
         Print #1, Spc(92); Str(rs!INTE_CAR_NUMERO)
         Print #1, ""
         Print #1, Spc(92); "FECHA: "; Format(rs!dtim_Car_fecha, "Short Date")
         var_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
         For var_j = 1 + Len(Trim(var_cliente)) To 83
             var_cliente = var_cliente + " "
         Next var_j
         var_cliente = var_cliente + "AGUASCALIENTES, AGS."
         Print #1, ""
         Print #1, Spc(12); var_cliente
         var_domicilio = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " C.P. " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
         For var_j = 1 + Len(Trim(var_domicilio)) To 83
             var_domicilio = var_domicilio + " "
         Next var_j
         var_agente = ""
         var_agente = IIf(IsNull(rs!vcha_AGE_aGENTE_ID), "", rs!vcha_AGE_aGENTE_ID)
         For var_j = 1 + Len(Trim(var_agente)) To 8
             var_agente = var_agente + " "
         Next var_j
         var_agente = var_agente + IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
         var_domicilio = var_domicilio
         Print #1, Spc(12); var_domicilio
         var_ciudad = ""
         var_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
         For var_j = 1 + Len(Trim(var_ciudad)) To 37
             var_ciudad = var_ciudad + " "
         Next var_j
         var_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
         For var_j = 1 + Len(Trim(var_estado)) To 46
             var_estado = var_estado + " "
         Next var_j
         var_ciudad = var_ciudad + var_estado
                            
         For var_j = 1 + Len(Trim(var_ciudad)) To 14
             var_ciudad = var_ciudad + " "
         Next var_j
                            
         var_ciudad = var_ciudad + var_agente
                           
         Print #1, Spc(12); var_ciudad
         var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
         var_rfc = "      " + var_rfc
         For var_j = 1 + Len(Trim(var_rfc)) To 89
             var_rfc = var_rfc + " "
         Next var_j
         rsaux3.Open "select * from vw_devolucion_nota_credito where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_movimiento), cnn, adOpenDynamic, adLockOptimistic
         var_referencia = rsaux3!vcha_cde_referencia
         var_rfc = var_rfc + "ENTRADA: " + var_clave_movimiento + " " + CStr(var_numero_movimiento) + " RELACION: " + var_referencia
         var_rfc = var_rfc
         Print #1, Spc(6); var_rfc
         Print #1, ""
         Print #1, ""
         Print #1, ""
         var_contador_renglones = 0
         var_cantidad_total = 0
         While Not rsaux3.EOF
               var_factura = CStr(rsaux3!inte_fac_factura) + IIf(IsNull(rsaux3!VCHA_SER_SERIE_ID), "", rsaux3!VCHA_SER_SERIE_ID)
               If Len(Trim(var_factura)) < 14 Then
                  For var_j = Len(Trim(var_factura)) To 14
                      var_factura = var_factura + " "
                  Next var_j
               End If
               var_linea = var_factura + rsaux3!vcha_Art_articulo_id + " " + rsaux3!vcha_art_nombre_Español
                    
               If Len(Trim(var_linea)) < 88 Then
                  For var_j = Len(Trim(var_linea)) To 88
                      var_linea = var_linea + " "
                  Next var_j
               End If
                      
               var_cantidad_str = Format(IIf(IsNull(rsaux3!Cantidad), 0, rsaux3!Cantidad), "###,###,##0.00")
               var_tipo_Cambio = IIf(IsNull(rsaux3!floa_dev_tipo_cambio), 1, rsaux3!floa_dev_tipo_cambio)
               var_cantidad = Format(IIf(IsNull(rsaux3!Cantidad), 0, rsaux3!Cantidad), "###,###,##0.00")
               var_cantidad_total = var_cantidad_total + IIf(IsNull(rsaux3!Cantidad), 0, rsaux3!Cantidad)
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
               If Len(Trim(var_rfc)) = 0 Then
                  var_precio = var_precio * (1 + var_iva / 100)
               End If
               var_precio_str = Format(var_precio, "###,###,##0.00")
               var_importe_str = Format(var_precio * var_cantidad, "###,###,##0.00")
               If Len(Trim(var_cantidad_str)) < 14 Then
                  For var_j = Len(Trim(var_cantidad_str)) To 14
                      var_cantidad_str = " " + var_cantidad_str
                  Next var_j
               End If
               var_linea = var_linea + var_cantidad_str
               If Len(Trim(var_linea)) < 104 Then
                  For var_j = 1 + Len(Trim(var_linea)) To 104
                      var_linea = var_linea + " "
                  Next var_j
               End If
               If Len(Trim(var_precio_str)) < 14 Then
                  For var_j = Len(Trim(var_precio_str)) To 14
                      var_precio_str = " " + var_precio_str
                  Next var_j
               End If
               If Len(Trim(var_importe_str)) < 14 Then
                  For var_j = Len(Trim(var_importe_str)) To 14
                      var_importe_str = " " + var_importe_str
                  Next var_j
               End If
               var_linea = var_linea + var_precio_str + var_importe_str
               Print #1, var_linea
               rsaux3.MoveNext
               var_contador_renglones = var_contador_renglones + 1
         Wend
         var_cantidad_total_str = Format(var_cantidad_total, "###,###,##0.00")
         If Len(Trim(var_cantidad_total_str)) < 14 Then
            For var_j = Len(Trim(var_cantidad_total_str)) To 14
                var_cantidad_total_str = " " + var_cantidad_total_str
            Next var_j
         End If
         If var_contador_renglones < 30 Then
            For var_j = var_contador_renglones To 30
                Print #1, ""
            Next var_j
         End If
         rsaux3.Close
         Print #1, ""
         Print #1, ""
         Print #1, ""
         var_contador_renglones = 0
         rsaux3.Open "select * from TB_DETALLE_DEVOLUCION_IMPORTES_ASIGNADOS where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'DV' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(var_factura_nueva), cnn, adOpenDynamic, adLockOptimistic
         var_linea = ""
         While Not rsaux3.EOF
               If Len(Trim(var_linea + ", " + CStr(rsaux3!inte_dia_numero) + rsaux3!vcha_dia_serie + " " + Format(rsaux3!floa_dia_importe, "###,###,##0.00"))) < 108 Then
                  If Len(Trim(var_linea)) = 0 Then
                     var_linea = var_linea + CStr(rsaux3!inte_dia_numero) + rsaux3!vcha_dia_serie + " " + Format(rsaux3!floa_dia_importe, "###,###,##0.00")
                  Else
                     var_linea = var_linea + ", " + CStr(rsaux3!inte_dia_numero) + rsaux3!vcha_dia_serie + " " + Format(rsaux3!floa_dia_importe, "###,###,##0.00")
                  End If
               Else
                  Print #1, var_linea
                  var_contador_renglones = var_contador_renglones + 1
                  var_linea = ""
                  var_linea = CStr(rsaux3!inte_dia_numero) + rsaux3!vcha_dia_serie + " " + Format(rsaux3!floa_dia_importe, "###,###,##0.00")
               End If
               rsaux3.MoveNext
               If rsaux3.EOF And Len(var_linea) < 118 Then
                  Print #1, var_linea
                  var_contador_renglones = var_contador_renglones + 1
               End If
         Wend
         If var_contador_renglones < 4 Then
            For var_j = var_contador_renglones To 4
                Print #1, ""
            Next var_j
         End If
         rsaux3.Close
         var_cantidad_letra = rs!vcha_car_importe_letra
                
         var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
             
         If Len(Trim(var_linea)) < 62 Then
            For var_j = 1 + Len(Trim(var_linea)) To 62
                var_linea = var_linea + " "
            Next var_j
         End If
         var_linea = var_linea + var_cantidad_total_str
         var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                
         If Len(Trim(var_rfc)) = 0 Then
            var_subimporte_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
            If Len(Trim(var_subimporte_str)) < 14 Then
               For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                   var_subimporte_str = " " + var_subimporte_str
               Next var_j
            End If
            var_iva = "      -        "
            For var_j = 1 + Len(Trim(var_iva_str)) To 14
                var_iva_str = " " + var_iva_str
            Next var_j
         Else
            var_subimporte_str = Format(((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
            If Len(Trim(var_subimporte_str)) < 14 Then
               For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                   var_subimporte_str = " " + var_subimporte_str
               Next var_j
            End If
            var_iva_str = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
            If Len(Trim(var_iva_str)) < 14 Then
               For var_j = 1 + Len(Trim(var_iva_str)) To 14
                   var_iva_str = " " + var_iva_str
               Next var_j
            End If
         End If
         If Len(Trim(var_linea)) < 115 Then
            For var_j = Len(Trim(var_linea)) To 115
                var_linea = var_linea + " "
            Next var_j
         End If
         var_linea = var_linea + var_subimporte_str
         Print #1, Spc(4); var_linea
         Print #1, Spc(120); var_iva_str
         var_importe_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
         If Len(Trim(var_importe_str)) < 14 Then
            For var_j = 1 + Len(Trim(var_importe_str)) To 14
                var_importe_str = " " + var_importe_str
            Next var_j
         End If
         Print #1, Spc(120); var_importe_str
         Print #1, ""
         Print #1, ""
         Print #1, ""
         Print #1, Spc(85); "SISTEMAS"
         Print #1, ""
         Print #1, ""
         Print #1, ""
         Print #1, ""
         Print #1, ""
         Print #1, ""
         Close #1
                    
         Open (App.Path & "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".bat") For Output As #2
         var_Archivo = App.Path & "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".bat"
         Print #2, "copy " + App.Path + "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".txt lpt1"
         Close #2
         x = Shell(var_Archivo, vbHide)
      End If
   End If
End Sub

Private Sub cmb_documentos_Click()
   txt_numero = ""
   var_estatus = ""
   lbl_estatus = ""
   If cmb_documentos = "FACTURA" Then
      txt_documento = "FA"
   End If
   If cmb_documentos = "NOTA DE CREDITO" Then
      txt_documento = "NC"
   End If
   If cmb_documentos = "NOTA DE CARGO" Then
      txt_documento = "NG"
   End If
   If cmb_documentos = "CANCELACION DE SALDOS" Then
      txt_documento = "CS"
   End If
   If cmb_documentos = "CHEQUES" Then
      txt_documento = "CH"
   End If
   If cmb_documentos = "PAGO" Then
      txt_documento = "PA"
   End If
End Sub

Private Sub cmb_documentos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_serie.SetFocus
   End If
End Sub



Private Sub cmd_cancelar_Click()
Dim si As Integer
Dim var_documento As String
Dim var_clase_documento As String
Dim var_afectacion As String
Dim var_cadena As String
Dim var_tipo_cancelacion As String
Dim var_estatus As String
Dim var_importe As Double
Dim var_tipo_Cambio As Double
Dim ndo As New aClsNodoArbolTrazabilidad

Set TB_ENCABEZA_CARTERA_I = New TB_ENCABEZADO_CARTERA_I
Set TB_ESTADO_CUENTA_INSERTA = New TB_ESTADO_CUENTA_INSERTA
var_serie = Me.txt_serie
If Trim(txt_documento) <> "" Then
   If Trim(txt_numero) <> "" Then
      rs.Open "select * from tb_encabezado_cartera where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_Car_tipo_documento = '" + txt_documento + "' and inte_Car_numero = " + txt_numero + " and vcha_Ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockBatchOptimistic
      If Not rs.EOF Then
         var_tipo_Cambio = rs!floa_car_tipo_cambio
         var_clase_documento = rs!vcha_Car_documento
         rs.Close
         rs.Open "SELECT * FROM VW_DOCUMENTOS_DEL_DIA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_CAR_TIPO_DOCUMENTO = '" + txt_documento + "' AND INTE_CAR_NUMERO = " + txt_numero + " and vcha_Ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
         'If Not rs.EOF Or Me.txt_documento = "NC" Or Me.txt_documento = "CS" Then
         If Not rs.EOF Then
            var_estatus = IIf(IsNull(rs!CHAR_CAR_ESTATUS), "", rs!CHAR_CAR_ESTATUS)
            If var_estatus <> "C" Then
               si = MsgBox("¿Deseas cancelar el documento " + Trim(cmb_documentos) + " serie " + Trim(txt_serie) + " número " + txt_numero, vbYesNo, "ATENCION")
               If si = 6 Then
                  si = MsgBox("Confirmar la cancelación del documento", vbYesNo, "ATENCION")
                  If si = 6 Then
                     var_documento = rs!vcha_Car_documento
                     var_clase_documento = rs!vcha_Car_clase_id
                     var_afectacion = rs!char_car_afectacion
                     If var_afectacion = "+" Then
                        rsaux.Open "select * from tb_estado_cuenta where vcha_Emp_empresa_id = '" + var_empresa + "'  and vcha_ecu_serie_cargo = '" + var_serie + "' and vcha_Ecu_movimiento_cargo = '" + txt_documento + "' and inte_Ecu_numero_cargo = " + txt_numero + " and floa_ecu_importe_abono > 0 and (char_ecu_estatus is null or char_ecu_estatus <> 'C')", cnn, adOpenDynamic, adLockOptimistic
                        If rsaux.EOF Then
                           cnn.BeginTrans
                           rsaux2.Open "update tb_encabezado_cartera set char_car_estatus = 'C' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_Car_documento =  '" + var_documento + "' and inte_car_numero = " + txt_numero + " and vcha_Ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                           rsaux2.Open "insert into TB_BITACORA_CANCELACION_DOCUMENTOS (vcha_emp_empresa_id, vcha_Car_documento, vcha_ser_serie_id, inte_car_numero, dtim_bit_fecha_Cancelacion, vcha_usu_usuario_id, VCHA_BIT_MAQUINA) values ('" + var_empresa + "', '" + var_documento + "', '" + var_serie + "', " + Me.txt_numero + ",getdate(),'" + var_clave_usuario_global + "', '" + fun_NombrePc + "')", cnn, adOpenDynamic, adLockOptimistic
                           var_estatus = "C"
                           lbl_estatus = "ESTATUS: CANCELADA"
                           cnn.CommitTrans
                           
                           If var_trazabilidad = 1000 Then
                              If cnn_trazabilidad.State = 0 Then
                                 cnn_trazabilidad.Open
                              End If
                        
                              rsaux10.Open "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                              var_nombre_unidad = ""
                              If Not rsaux10.EOF Then
                                 var_nombre_unidad = IIf(IsNull(rsaux10!VCHA_UOR_NOMBRE), "", rsaux10!VCHA_UOR_NOMBRE)
                              End If
                              rsaux10.Close
                                
                              ndo.organizacion = var_nombre_unidad
                              ndo.eventoNumero = Me.txt_numero
                              If ndo.cancelarFactura(cnn_trazabilidad) Then
                                    
                              Else
                                 MsgBox "No se pudo ejecutar la trazabilidad", vbOKOnly, "ATENCION"
                              End If
                              cnn_trazabilidad.Close
                           End If
                         Else
                           MsgBox "El documento ya no puede ser cancelado ya que tiene abonos", vbOKOnly, "ATENCION"
                         End If
                         rsaux.Close
                     End If
                     If var_afectacion = "-" Then
                        cnn.BeginTrans

                        rsaux2.Open "update tb_encabezado_cartera set char_car_estatus = 'C' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_Car_documento =  '" + var_documento + "' and inte_car_numero = " + txt_numero + " AND VCHA_SER_SERIE_ID = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                        If var_documento = "DF" Then
                           rsaux10.Open "select * from tb_Relacion_cobranza  WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_RCO_NUMERO_DESCUENTO_FINANCIERO = " + Me.txt_numero + " and vcha_Ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                           While Not rsaux10.EOF
                                 rsaux2.Open "UPDATE TB_RELACION_COBRANZA SET INTE_RCO_NUMERO_DESCUENTO_FINANCIERO = 0 WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_RCO_NUMERO_DESCUENTO_FINANCIERO = " + Me.txt_numero + " and vcha_Ser_serie_id = '" + var_serie + "' and inte_rco_partida  = " + CStr(IIf(IsNull(rsaux10!inte_rco_partida), 0, rsaux10!inte_rco_partida)), cnn, adOpenDynamic, adLockOptimistic
                                 rsaux10.MoveNext
                           Wend
                           rsaux10.Close
                        End If
                        
                        rsaux2.Open "insert into TB_BITACORA_CANCELACION_DOCUMENTOS (vcha_emp_empresa_id, vcha_Car_documento, vcha_ser_serie_id, inte_car_numero, dtim_bit_fecha_Cancelacion, vcha_usu_usuario_id, VCHA_BIT_MAQUINA) values ('" + var_empresa + "', '" + var_documento + "', '" + var_serie + "', " + Me.txt_numero + ",getdate(),'" + var_clave_usuario_global + "', '" + fun_NombrePc + "')", cnn, adOpenDynamic, adLockOptimistic
                        If Me.txt_documento = "PA" Then
                           rsaux2.Open "UPDATE TB_RELACION_COBRANZA SET CHAR_RCO_APLICADA = '' WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_RCO_PAGO = " + Me.txt_numero + " and vcha_Ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                        End If
                        
                        If var_documento = "DF" Then
                           rsaux9.Open "select * from tb_relacion_cobranza where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_ser_Serie_id = '" + var_serie + "' and inte_rco_numero_descuento_financiero = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
                           While Not rsaux9.EOF
                                 rsaux8.Open "update tb_relacion_cobranza set inte_rco_numero_descuento_financiero = 0 where vcha_emp_empresa_id = '" + rsaux9!VCHA_EMP_EMPRESA_ID + "' and vcha_rco_folio = '" + rsaux9!VCHA_RCO_FOLIO + "' and inte_car_numero = " + CStr(rsaux9!INTE_CAR_NUMERO) + " and vcha_ser_Serie_id = '" + Me.txt_serie + "' and inte_rco_partida = " + CStr(rsaux9!inte_rco_partida), cnn, adOpenDynamic, adLockOptimistic
                                 rsaux9.MoveNext
                           Wend
                           rsaux9.Close
                        End If
                        
                        
                        
                        'rsaux4.Open "select * from tb_estado_cuenta where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_ecu_serie_abono = '" + var_serie + "' and inte_ecu_numero_abono = " + txt_numero + " and vcha_ecu_movimiento_abono = '" + var_clase_documento + "'", cnn, adOpenDynamic, adLockOptimistic
                        'While Not rsaux4.EOF
                        '   var_importe = rsaux4!floa_ecu_importe_abono / var_tipo_Cambio
                        '   rsaux2.Open "update tb_saldos set floa_sal_importe =  floa_sal_importe + " + CStr(var_importe) + " where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = '" + rsaux4!vcha_ecu_movimiento_cargo + "' and inte_car_numero = " + CStr(rsaux4!inte_ecu_numero_cargo) + " and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                        '   rsaux4.MoveNext
                        'Wend
                        'rsaux4.Close
                        
                        cnn.CommitTrans
                        var_estatus = "C"
                        lbl_estatus = "ESTATUS: CANCELADA"
                     End If
                  Else
                     MsgBox "Se a cancelado la cancelación del documento", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "Se a cancelado la cancelación del documento", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "El documento ya fue cancelado con anterioridad", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El documento no existe o fue elaborado otro dia", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         rs.Close
         si = MsgBox("¿Deseas cancelar el documento " + Trim(cmb_documentos) + " serie " + Trim(txt_serie) + " número " + txt_numero, vbYesNo, "ATENCION")
         If si = 6 Then
            si = MsgBox("Confirmar la cancelación del documento", vbYesNo, "ATENCION")
            If si = 6 Then
               If txt_documento = "FA" Then
                  var_tipo_cancelacion = "CF"
                  var_afectacion = "+"
               End If
               If txt_documento = "NC" Then
                  var_tipo_cancelacion = "CN"
                  var_afectacion = "-"
               End If
               If txt_documento = "NG" Then
                  var_tipo_cancelacion = "CG"
                  var_afectacion = "+"
               End If
               lbl_estatus = "ESTATUS: CANCELADA"
               var_documento = txt_documento
               var_clase_documento = Var_clase
               If var_afectacion = "+" Then
                  var_inserta = TB_ENCABEZA_CARTERA_I.Anadir(var_empresa, var_unidad_organizacional, txt_documento, var_tipo_cancelacion, var_tipo_cancelacion, Val(txt_numero), var_afectacion, _
                  "", "", 0, Date, "", "", "", "", "", "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", var_clave_usuario_global, fun_NombrePc, Date, 0, Date, Date, "", 0, var_serie, "C")
               End If
               If var_afectacion = "-" Then
                  var_inserta = TB_ENCABEZA_CARTERA_I.Anadir(var_empresa, var_unidad_organizacional, txt_documento, var_tipo_cancelacion, var_tipo_cancelacion, Val(txt_numero), var_afectacion, _
                  "", "", 0, Date, "", "", "", "", "", "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", var_clave_usuario_global, fun_NombrePc, Date, 0, Date, Date, "", 0, var_serie, "C")
               End If
            Else
               MsgBox "Se a cancelado la cancelación del documento", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Se a cancelado la cancelación del documento", vbOKOnly, "ATENCION"
         End If
      End If
   Else
      MsgBox "Número de documento incorrecto", vbOKOnly, "ATENCION"
   End If
Else
   MsgBox "Documento incorrecto", vbOKOnly, "ATENCION"
End If
End Sub


Private Sub cmd_imprimir_Click()
Dim si As Integer
Dim var_movimiento As String
Dim var_almacen As String
Dim var_linea As String
Dim var_mes_str As String
Dim var_dia_str As String
Dim var_anio_str As String

   If var_estatus = "C" Then
      If txt_documento = "FA" Then
         If var_tipo_facturacion = "E" Then
            si = MsgBox("¿Deseas reimprimir la factura " + txt_numero + "?", vbYesNo, "ATENCION")
            If si = 6 Then
               If var_empresa = "02" Or var_empresa = "18" Or var_empresa = "31" Then
                  rs.Open "select vcha_mov_movimiento_id from tb_encabezado_cartera where vcha_emp_empresa_id = '" + var_empresa + "' AND VCHA_SER_SERIE_ID = '" + var_serie + "' and inte_Car_numero = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
                  var_clave_movimiento_factura = IIf(IsNull(rs!vcha_mov_movimiento_id), "", rs!vcha_mov_movimiento_id)
                  rs.Close
                  If var_clave_movimiento_factura = "FV" Then
                     rs.Open "select * from vw_facturas_embarque_vistas where vcha_emp_empresa_id = '" + var_empresa + "' AND VCHA_SER_SERIE_ID = '" + var_serie + "' and inte_Car_numero = " + txt_numero + " ORDER BY INTE_SAL_CONSECUTIVO_FACTURA", cnn, adOpenDynamic, adLockOptimistic
                  Else
                     rs.Open "select * from vw_facturas_embarque where vcha_emp_empresa_id = '" + var_empresa + "' AND VCHA_SER_SERIE_ID = '" + var_serie + "' and inte_Car_numero = " + txt_numero + " ORDER BY INTE_SAL_CONSECUTIVO_FACTURA", cnn, adOpenDynamic, adLockOptimistic
                  End If
               End If
               If var_empresa = "03" Then
                  rs.Open "select * from vw_facturas_embarque where vcha_emp_empresa_id = '" + var_empresa + "' AND VCHA_SER_SERIE_ID = '" + var_serie + "' and inte_Car_numero = " + txt_numero + " ORDER BY vcha_sal_descripcion_factura", cnn, adOpenDynamic, adLockOptimistic
               End If
               If Not rs.EOF Then
                  var_movimiento = IIf(IsNull(rs!vcha_mov_movimiento_id), "", rs!vcha_mov_movimiento_id)
                  var_almacen = IIf(IsNull(rs!vcha_alm_almacen_id), "", rs!vcha_alm_almacen_id)
                  var_numero_movimiento = IIf(IsNull(rs!INTE_SAL_NUMERO), 0, rs!INTE_SAL_NUMERO)
                  var_numero_embarque = IIf(IsNull(rs!inte_emb_embarque), 0, rs!inte_emb_embarque)
                  rsaux.Open "select * from tb_encabezado_cartera where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_mov_movimiento_id = '" + var_movimiento + "' and inte_emo_numero = " + Str(var_numero_movimiento) + " and vcha_alm_almacen_id = '" + IIf(IsNull(rs!vcha_alm_almacen_id), "", rs!vcha_alm_almacen_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     rsaux2.Open "select * from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                     var_factura_nueva = rsaux2!inte_ser_factura
                     rsaux2.Close
                     si = MsgBox("¿Se va a imprimir la factura " + Str(var_factura_nueva) + "?", vbYesNo, "ATENCION")
                     If si = 6 Then
                        si = MsgBox("Confirmar la reimpresión de la factura " + Str(var_factura_nueva), vbYesNo, "ATENCION")
                        If si = 6 Then
                           rsaux2.Open "select * from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                           If var_factura_nueva <> rsaux2!inte_ser_factura Then
                              MsgBox "El número de la factura ya cambio y el proceso de reimpresión a cambiado", vbOKOnly, "ATENCION"
                              rsaux2.Close
                           Else
                              rsaux2.Close
                              var_cadena = ""
                              var_cadena = "INSERT INTO [TB_ENCABEZADO_CARTERA] ([VCHA_EMP_EMPRESA_ID], [VCHA_UOR_UNIDAD_ID], [VCHA_CAR_TIPO_DOCUMENTO], [VCHA_CAR_DOCUMENTO], [VCHA_CAR_CLASE_ID], [INTE_CAR_NUMERO], [CHAR_CAR_AFECTACION], [VCHA_ALM_ALMACEN_ID], [VCHA_MOV_MOVIMIENTO_ID], [INTE_EMO_NUMERO], [DTIM_CAR_FECHA], [VCHA_AGE_AGENTE_ID], [VCHA_GAC_GRUPO_ACTUAL_ID], [VCHA_GRE_GRUPO_REAL_ID], [VCHA_TIT_TITULAR_ID], [VCHA_CLI_CLAVE_ID], [VCHA_ESB_ESTABLECIMIENTO_ID], [INTE_CAR_PLAZO], [FLOA_CAR_PORCENTAJE_IVA], [FLOA_CAR_PORCENTAJE_IMPUESTO_1], [FLOA_CAR_PORCENTAJE_IMPUESTO_2], [FLOA_CAR_PORCENTAJE_DESCUENTO_1], [FLOA_CAR_PORCENTAJE_DESCUENTO_2], [FLOA_CAR_PORCENTAJE_DESCUENTO_3], [FLOA_CAR_IMPORTE_TOTAL], [FLOA_CAR_IMPORTE_IVA], [FLOA_CAR_IMPORTE_IMPUESTO_1], [FLOA_CAR_IMPORTE_IMPUESTO_2], [FLOA_CAR_IMPORTE_DESCUENTO_1], [FLOA_CAR_IMPORTE_DESCUENTO_2], [FLOA_CAR_IMPORTE_DESCUENTO_3], [FLOA_CAR_SUBIMPORTE], [FLOA_CAR_IMPORTE_NETO], [VCHA_CAR_IMPORTE_LETRA], [VCHA_AUD_USUARIO], [VCHA_AUD_MAQUINA], "
                              var_cadena = var_cadena + "[VCHA_AUD_FECHA], [FLOA_CAR_SALDO], [DTIM_CAR_FECHA_VENCIMIENTO], [DTIM_CAR_FECHA_ENTREGA], [VCHA_MON_MONEDA_ID], [FLOA_CAR_TIPO_CAMBIO], [VCHA_SER_SERIE_ID], [CHAR_CAR_ESTATUS], [CHAR_CAR_TIPO_FACTURACION], [INTE_CAR_FACTURA_CEROS], [FLOA_CAR_COSTO]) Values ('" + IIf(IsNull(rs!VCHA_EMP_EMPRESA_ID), "", rs!VCHA_EMP_EMPRESA_ID) + "', '" + rs!VCHA_UOR_UNIDAD_ID + "', '" + rs!vcha_Car_tipo_documento + "', '" + rs!vcha_Car_documento + "', '" + rs!vcha_Car_clase_id + "', " + CStr(var_factura_nueva) + ", '" + rs!char_car_afectacion
                              var_cadena = var_cadena + "', '" + rs!vcha_alm_almacen_id + "', '" + rs!vcha_mov_movimiento_id + "', " + CStr(rs!INTE_EMO_NUMERO) + ", getdate(),  '" + rs!vcha_AGE_aGENTE_ID + "', '" + rs!VCHA_GAC_GRUPO_aCTUAL_ID + "', '" + rs!vcha_gre_grupo_real_id + "', '" + rs!vcha_tit_titular_id + "', '" + rs!vcha_cli_clave_id + "', '" + rs!VCHA_ESB_ESTABLECIMIENTO_ID + "', " + CStr(rs!INTE_CAR_PLAZO) + ", " + CStr(rs!floa_car_porcentaje_iva) + ", " + CStr(rs!FLOA_CAR_PORCENTAJE_IMPUESTO_1) + ", " + CStr(rs!FLOA_CAR_PORCENTAJE_IMPUESTO_2) + ", " + CStr(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1) + ", " + CStr(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2) + ", "
                              var_cadena = var_cadena + CStr(IIf(IsNull(rs!floa_car_porcentaje_descuento_3), 0, rs!floa_car_porcentaje_descuento_3)) + ", " + CStr(rs!FLOA_CAR_IMPORTE_TOTAL) + ", " + CStr(rs!floa_car_importe_iva) + ", " + CStr(rs!FLOA_CAR_IMPORTE_IMPUESTO_1) + ", " + CStr(rs!FLOA_CAR_IMPORTE_IMPUESTO_2) + ", " + CStr(rs!FLOA_CAR_IMPORTE_DESCUENTO_1) + ", " + CStr(rs!FLOA_CAR_IMPORTE_DESCUENTO_2) + ", "
                              var_cadena = var_cadena + CStr(rs!FLOA_CAR_IMPORTE_DESCUENTO_3) + "," + CStr(rs!floa_car_subimporte) + ", " + CStr(rs!floa_Car_importe_neto) + ", '" + rs!vcha_car_importe_letra + "', '" + rs!vcha_aud_usuario + "', '" + rs!vcha_aud_maquina + "', getdate(), 0, null, null, '" + rs!vcha_mon_moneda_id + "', " + CStr(rs!floa_car_tipo_cambio) + ", '" + rs!VCHA_SER_SERIE_ID + "', '', 'E', " + CStr(IIf(IsNull(rs!INTE_CAR_FACTURA_CEROS), 0, rs!INTE_CAR_FACTURA_CEROS)) + ", " + CStr(IIf(IsNull(rs!FLOA_CAR_COSTO), 0, rs!FLOA_CAR_COSTO)) + ") "
                              cnn.BeginTrans
                              rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              rsaux2.Open "INSERT INTO [TB_ESTADO_CUENTA] ([VCHA_EMP_EMPRESA_ID], [VCHA_ECU_SERIE_CARGO], [VCHA_ECU_MOVIMIENTO_CARGO], [INTE_ECU_NUMERO_CARGO], [FLOA_ECU_IMPORTE_CARGO], [FLOA_ECU_IMPORTE_ABONO]) Values ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!VCHA_SER_SERIE_ID + "', 'FA', " + CStr(var_factura_nueva) + ", " + CStr(rs!floa_Car_importe_neto) + ", 0) ", cnn, adOpenDynamic, adLockOptimistic
                              rsaux2.Open "INSERT INTO TB_SECUENCIA_FACTURACION (VCHA_EMP_EMPRESA_ID, VCHA_SER_SERIE_ID, INTE_SFA_NUMERO_ANTERIOR, INTE_SFA_NUMERO_ACTUAL) VALUES ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!VCHA_SER_SERIE_ID + "', " + CStr(var_factura_nueva) + ", " + CStr(var_factura_nueva) + ")", cnn, adOpenDynamic, adLockOptimistic
                              rsaux2.Open "UPDATE TB_SECUENCIA_FACTURACION SET INTE_SFA_NUMERO_ACTUAL = " + CStr(var_factura_nueva) + " WHERE VCHA_EMP_EMPRESA_ID = '" + rs!VCHA_EMP_EMPRESA_ID + "' AND VCHA_SER_SERIE_ID = '" + rs!VCHA_SER_SERIE_ID + "' AND  INTE_SFA_NUMERO_ANTERIOR = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
                              rsaux2.Open "INSERT INTO TB_INVENTARIO_DOCUMENTOS (VCHA_EMP_EMPRESA_ID, VCHA_AGE_AGENTE_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_CAR_CLASE_ID, INTE_CAR_NUMERO, CHAR_CAR_AFECTACION, VCHA_SER_SERIE_ID, CHAR_IDO_ESTATUS, FLOA_IDO_CANTIDAD, FLOA_CAR_IMPORTE_NETO, FLOA_CAR_TIPO_CAMBIO, VCHA_MON_MONEDA_ID, DTIM_IDO_FECHA_ENTRAGA, VCHA_CLI_CLAVE_ID, INTE_EMB_EMBARQUE) VALUES ('" + rs!VCHA_EMP_EMPRESA_ID + "','" + rs!vcha_AGE_aGENTE_ID + "', 'FA', 'FA', '" + rs!vcha_Car_clase_id + "', " + CStr(var_factura_nueva) + ",'+', '" + rs!VCHA_SER_SERIE_ID + "','A',0," + CStr(rs!floa_Car_importe_neto) + "," + CStr(rs!floa_car_tipo_cambio) + ",'" + rs!vcha_mon_moneda_id + "',GETDATE(),'" + rs!vcha_cli_clave_id + "',0)"
                              
                              rsaux3.Open "select * from tb_salidas where  VCHA_EMP_EMPRESA_ID = '" + rs!VCHA_EMP_EMPRESA_ID + "' AND VCHA_UOR_UNIDAD_ID = '" + rs!VCHA_UOR_UNIDAD_ID + "' AND VCHA_CAR_DOCUMENTO = 'FA' AND VCHA_SER_SERIE_ID = '" + rs!VCHA_SER_SERIE_ID + "' AND INTE_CAR_NUMERO = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
                              While Not rsaux3.EOF
                                    If rsaux2.State = 1 Then
                                       rsaux2.Close
                                    End If
                                    rsaux2.Open "UPDATE TB_SALIDAS SET INTE_CAR_NUMERO = " + CStr(var_factura_nueva) + " WHERE VCHA_EMP_EMPRESA_ID = '" + rs!VCHA_EMP_EMPRESA_ID + "' AND VCHA_UOR_UNIDAD_ID = '" + rs!VCHA_UOR_UNIDAD_ID + "' AND VCHA_CAR_DOCUMENTO = 'FA' AND VCHA_SER_SERIE_ID = '" + rs!VCHA_SER_SERIE_ID + "' AND INTE_CAR_NUMERO = " + txt_numero + " and vcha_art_articulo_id = '" + rsaux3!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                    rsaux3.MoveNext
                              Wend
                              rsaux3.Close
                              If rsaux2.State = 1 Then
                                 rsaux2.Close
                              End If
                              rsaux2.Open "update tb_series set inte_ser_factura = isnull(inte_ser_factura,0) + 1 where vcha_emp_empresa_id = '" + rs!VCHA_EMP_EMPRESA_ID + "' and vcha_uor_unidad_id = '" + rs!VCHA_UOR_UNIDAD_ID + "' and vcha_Ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                              cnn.CommitTrans
                              Call imprime_factura
                              
                              
                              
                              If var_trazabilidad = 1000 Then
                                 rsaux10.Open "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                                 var_nombre_unidad = ""
                                 If Not rsaux10.EOF Then
                                    var_nombre_unidad = IIf(IsNull(rsaux10!VCHA_UOR_NOMBRE), "", rsaux10!VCHA_UOR_NOMBRE)
                                 End If
                                 rsaux10.Close
                                   
                                 var_cadena = "SELECT dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_SALIDAS.VCHA_CAR_DOCUMENTO, dbo.TB_SALIDAS.VCHA_SER_SERIE_ID, dbo.TB_SALIDAS.INTE_CAR_NUMERO FROM dbo.TB_DETALLE_EMBARQUES INNER JOIN dbo.TB_SALIDAS ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND"
                                 var_cadena = var_cadena + " dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID = dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO = dbo.TB_SALIDAS.INTE_SAL_NUMERO WHERE     (dbo.TB_SALIDAS.VCHA_CAR_DOCUMENTO = 'FA') AND (dbo.TB_SALIDAS.VCHA_SER_SERIE_ID = '" + Me.txt_serie + "') AND (dbo.TB_SALIDAS.INTE_CAR_NUMERO = " + Me.txt_numero + ")"
                                 rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                 If Not rsaux10.EOF Then
                                    var_numero_embarque = rsaux10!inte_emb_embarque
                                    If cnn_trazabilidad.State = 0 Then
                                       cnn_trazabilidad.Open
                                    End If
                                    var_cadena = "SELECT     dbo.TB_SALIDAS.VCHA_SER_SERIE_ID, dbo.TB_SALIDAS.INTE_CAR_NUMERO, dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE FROM dbo.TB_DETALLE_EMBARQUES INNER JOIN dbo.TB_SALIDAS ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID = dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID AND"
                                    var_cadena = var_cadena + " dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO = dbo.TB_SALIDAS.INTE_SAL_NUMERO WHERE     (dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = " + var_numero_embarque + ")"
                                    rsaux9.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                    While Not rsaux9.EOF
                                          If Not ndo.generarIdentificadorInformacionNodo(cnn_trazabilidad) Then
                                             MsgBox "No se pudo generar la trazabilidad", vbOKOnly, "ATENCION"
                                          Else
                                             ndo.nodoTipo = "E"
                                             ndo.organizacion = var_nombre_unidad
                                             ndo.eventoTipo = "FACTURACIÓN"
                                             ndo.eventoNumero = IIf(IsNull(rsaux9!INTE_CAR_NUMERO), 0, rsaux9!INTE_CAR_NUMERO)
                                             If Not ndo.registrarInformacionNodo(cnn_trazabilidad) Then
                                                MsgBox "No se pudo registrar la información del nodo de trazabilidad", vbOKOnly, "ATENCION"
                                             Else
                                                ndo.nodoTipo = "E"
                                                ndo.nodoPadreTipo = "N"
                                                ndo.informacionNodoPadreIdentificador = "0"
                                                If Not ndo.registrarNodo(cnn_trazabilidad) Then
                                                   MsgBox "No se pudo registrar nodo de trazabilidad", vbOKOnly, "ATENCION"
                                                Else
                                                   ndo.nodoPadreTipo = "E"
                                                   ndo.informacionNodoPadreIdentificador = ndo.informacionNodoIdentificador
                                                   ndo.nodoTipo = "L"
                                                   var_cadena = "SELECT DISTINCT dbo.TB_SALIDAS.VCHA_SER_SERIE_ID, dbo.TB_SALIDAS.INTE_CAR_NUMERO, dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE,dbo.TB_TRAZABILIDAD_CODIGOS.INTE_TRA_NODO_IDENTIFICADOR , dbo.TB_TRAZABILIDAD_CODIGOS.INTE_TRA_LOTE FROM dbo.TB_DETALLE_EMBARQUES INNER JOIN dbo.TB_SALIDAS ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID = dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID = dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID AND"
                                                   var_cadena = var_cadena + " dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_SALIDAS.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO = dbo.TB_SALIDAS.INTE_SAL_NUMERO INNER JOIN dbo.TB_TRAZABILIDAD_CODIGOS ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_TRAZABILIDAD_CODIGOS.VCHA_EMP_EMPRESA_ID AND dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = dbo.TB_TRAZABILIDAD_CODIGOS.INTE_EMB_EMBARQUE AND dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID = dbo.TB_TRAZABILIDAD_CODIGOS.VCHA_ART_ARTICULO_ID WHERE (dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = " + var_numero_embarque + ") and (dbo.tb_salidas.inte_car_numero = " + CStr(IIf(IsNull(rsaux9!INTE_CAR_NUMERO), 0, rsaux9!INTE_CAR_NUMERO)) + ")"
                                                   rsaux7.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                        
                                                   While Not rsaux7.EOF
                                                         ndo.informacionNodoIdentificador = rsaux7!inte_tra_nodo_identificador
                                                         If Not ndo.registrarNodo(cnn_trazabilidad) Then
                                                            MsgBox "No se pudo registrar nodo de trazabilidad", vbOKOnly, "ATENCION"
                                                         End If
                                                         rsaux7.MoveNext
                                                   Wend
                                                   rsaux7.Close
                                                End If
                                                
                                             End If
                                          End If
                                          rsaux9.MoveNext
                                    Wend
                                    rsaux9.Close
                                    cnn_trazabilidad.Close
                                 End If
                              End If
                              
                              
                              
                              
                              
                              
                           End If
                        End If
                     End If
                  End If
                  rsaux.Close
               Else
                  MsgBox "Probablemente la factura ya fue reimpresa", vbOKOnly, "ATENCION"
               End If
               rs.Close
            End If
         End If
         'MsgBox var_tipo_facturacion
         If var_tipo_facturacion = "V" Then
            si = MsgBox("¿Deseas reimprimir la factura " + txt_numero + "?", vbYesNo, "ATENCION")
            If si = 6 Then
               rs.Open "select * from vw_facturas_vistas where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'FA' and vcha_ser_serie_id = '" + var_serie + "' and inte_Car_numero = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_movimiento = IIf(IsNull(rs!vcha_mov_movimiento_id), "", rs!vcha_mov_movimiento_id)
                  var_almacen = IIf(IsNull(rs!vcha_alm_almacen_id), "", rs!vcha_alm_almacen_id)
                  var_numero_movimiento = IIf(IsNull(rs!INTE_EMO_NUMERO), 0, rs!INTE_EMO_NUMERO)
'                  var_numero_embarque = IIf(IsNull(rs!inte_emb_embarque), 0, rs!inte_emb_embarque)
                  rsaux.Open "select * from tb_encabezado_cartera where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_mov_movimiento_id = '" + var_movimiento + "' and inte_emo_numero = " + Str(var_numero_movimiento) + " and vcha_alm_almacen_id = '" + IIf(IsNull(rs!vcha_alm_almacen_id), "", rs!vcha_alm_almacen_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     rsaux2.Open "select * from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                     var_factura_nueva = rsaux2!inte_ser_factura
                     rsaux2.Close
                     si = MsgBox("¿Se va a imprimir la factura " + Str(var_factura_nueva) + "?", vbYesNo, "ATENCION")
                     If si = 6 Then
                        si = MsgBox("Confirmar la reimpresión de la factura " + Str(var_factura_nueva), vbYesNo, "ATENCION")
                        If si = 6 Then
                           rsaux2.Open "select * from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                           If var_factura_nueva <> rsaux2!inte_ser_factura Then
                              MsgBox "El número de la factura ya cambio y el proceso de reimpresión a cambiado", vbOKOnly, "ATENCION"
                              rsaux2.Close
                           Else
                              rsaux2.Close
                              var_cadena = ""
                              var_cadena = "INSERT INTO [TB_ENCABEZADO_CARTERA] ([VCHA_EMP_EMPRESA_ID], [VCHA_UOR_UNIDAD_ID], [VCHA_CAR_TIPO_DOCUMENTO], [VCHA_CAR_DOCUMENTO], [VCHA_CAR_CLASE_ID], [INTE_CAR_NUMERO], [CHAR_CAR_AFECTACION], [VCHA_ALM_ALMACEN_ID], [VCHA_MOV_MOVIMIENTO_ID], [INTE_EMO_NUMERO], [DTIM_CAR_FECHA], [VCHA_AGE_AGENTE_ID], [VCHA_GAC_GRUPO_ACTUAL_ID], [VCHA_GRE_GRUPO_REAL_ID], [VCHA_TIT_TITULAR_ID], [VCHA_CLI_CLAVE_ID], [VCHA_ESB_ESTABLECIMIENTO_ID], [INTE_CAR_PLAZO], [FLOA_CAR_PORCENTAJE_IVA], [FLOA_CAR_PORCENTAJE_IMPUESTO_1], [FLOA_CAR_PORCENTAJE_IMPUESTO_2], [FLOA_CAR_PORCENTAJE_DESCUENTO_1], [FLOA_CAR_PORCENTAJE_DESCUENTO_2], [FLOA_CAR_PORCENTAJE_DESCUENTO_3], [FLOA_CAR_IMPORTE_TOTAL], [FLOA_CAR_IMPORTE_IVA], [FLOA_CAR_IMPORTE_IMPUESTO_1], [FLOA_CAR_IMPORTE_IMPUESTO_2], [FLOA_CAR_IMPORTE_DESCUENTO_1], [FLOA_CAR_IMPORTE_DESCUENTO_2], [FLOA_CAR_IMPORTE_DESCUENTO_3], [FLOA_CAR_SUBIMPORTE], [FLOA_CAR_IMPORTE_NETO], [VCHA_CAR_IMPORTE_LETRA], [VCHA_AUD_USUARIO], [VCHA_AUD_MAQUINA], "
                              var_cadena = var_cadena + "[VCHA_AUD_FECHA], [FLOA_CAR_SALDO], [DTIM_CAR_FECHA_VENCIMIENTO], [DTIM_CAR_FECHA_ENTREGA], [VCHA_MON_MONEDA_ID], [FLOA_CAR_TIPO_CAMBIO], [VCHA_SER_SERIE_ID], [CHAR_CAR_ESTATUS], [CHAR_CAR_TIPO_FACTURACION]) Values ('" + IIf(IsNull(rs!VCHA_EMP_EMPRESA_ID), "", rs!VCHA_EMP_EMPRESA_ID) + "', '" + rs!VCHA_UOR_UNIDAD_ID + "', 'FA', '" + rs!vcha_Car_documento + "', '" + rs!vcha_Car_clase_id + "', " + CStr(var_factura_nueva) + ", '" + rs!char_car_afectacion
                              var_cadena = var_cadena + "', '" + rs!vcha_alm_almacen_id + "', '" + rs!vcha_mov_movimiento_id + "', " + CStr(rs!INTE_EMO_NUMERO) + ", getdate(),  '" + rs!vcha_AGE_aGENTE_ID + "', '" + rs!VCHA_GAC_GRUPO_aCTUAL_ID + "', '" + rs!vcha_gre_grupo_real_id + "', '" + rs!vcha_tit_titular_id + "', '" + rs!vcha_cli_clave_id + "', '" + rs!VCHA_ESB_ESTABLECIMIENTO_ID + "', " + CStr(rs!INTE_CAR_PLAZO) + ", " + CStr(rs!floa_car_porcentaje_iva) + ", " + CStr(rs!FLOA_CAR_PORCENTAJE_IMPUESTO_1) + ", " + CStr(rs!FLOA_CAR_PORCENTAJE_IMPUESTO_2) + ", " + CStr(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1) + ", " + CStr(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2) + ", " + CStr(rs!floa_car_porcentaje_descuento_3) + ", " + CStr(rs!FLOA_CAR_IMPORTE_TOTAL) + ", " + CStr(rs!floa_car_importe_iva) + ", " + CStr(rs!FLOA_CAR_IMPORTE_IMPUESTO_1) + ", " + CStr(rs!FLOA_CAR_IMPORTE_IMPUESTO_2) + ", " + CStr(rs!FLOA_CAR_IMPORTE_DESCUENTO_1) + ", " + CStr(rs!FLOA_CAR_IMPORTE_DESCUENTO_2) + ", "
                              var_cadena = var_cadena + CStr(rs!FLOA_CAR_IMPORTE_DESCUENTO_3) + "," + CStr(rs!floa_car_subimporte) + ", " + CStr(rs!floa_Car_importe_neto) + ", '" + rs!vcha_car_importe_letra + "', '" + rs!vcha_aud_usuario + "', '" + rs!vcha_aud_maquina + "', getdate(), 0, null, null, '" + rs!vcha_mon_moneda_id + "', " + CStr(rs!floa_car_tipo_cambio) + ", '" + rs!VCHA_SER_SERIE_ID + "', '" + rs!CHAR_CAR_ESTATUS + "', 'V') "
                              cnn.BeginTrans
                              rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              rsaux2.Open "INSERT INTO [TB_ESTADO_CUENTA] ([VCHA_EMP_EMPRESA_ID], [VCHA_ECU_SERIE_CARGO], [VCHA_ECU_MOVIMIENTO_CARGO], [INTE_ECU_NUMERO_CARGO], [FLOA_ECU_IMPORTE_CARGO], [FLOA_ECU_IMPORTE_ABONO]) Values ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!VCHA_SER_SERIE_ID + "', 'FA', " + CStr(var_factura_nueva) + ", " + CStr(rs!floa_Car_importe_neto) + ", 0) ", cnn, adOpenDynamic, adLockOptimistic
                              rsaux2.Open "INSERT INTO TB_SECUENCIA_FACTURACION (VCHA_EMP_EMPRESA_ID, VCHA_SER_SERIE_ID, INTE_SFA_NUMERO_ANTERIOR, INTE_SFA_NUMERO_ACTUAL) VALUES ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!VCHA_SER_SERIE_ID + "', " + CStr(var_factura_nueva) + ", " + CStr(var_factura_nueva) + ")", cnn, adOpenDynamic, adLockOptimistic
                              rsaux2.Open "UPDATE TB_SECUENCIA_FACTURACION SET INTE_SFA_NUMERO_ACTUAL = " + CStr(var_factura_nueva) + " WHERE VCHA_EMP_EMPRESA_ID = '" + rs!VCHA_EMP_EMPRESA_ID + "' AND VCHA_SER_SERIE_ID = '" + rs!VCHA_SER_SERIE_ID + "' AND  INTE_SFA_NUMERO_ANTERIOR = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
                              
                              rsaux2.Open "INSERT INTO TB_INVENTARIO_DOCUMENTOS (VCHA_EMP_EMPRESA_ID, VCHA_AGE_AGENTE_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_CAR_CLASE_ID, INTE_CAR_NUMERO, CHAR_CAR_AFECTACION, VCHA_SER_SERIE_ID, CHAR_IDO_ESTATUS, FLOA_IDO_CANTIDAD, FLOA_CAR_IMPORTE_NETO, FLOA_CAR_TIPO_CAMBIO, VCHA_MON_MONEDA_ID, DTIM_IDO_FECHA_ENTRAGA, VCHA_CLI_CLAVE_ID, INTE_EMB_EMBARQUE) VALUES ('" + rs!VCHA_EMP_EMPRESA_ID + "','" + rs!vcha_AGE_aGENTE_ID + "', 'FA', 'FA', '" + rs!vcha_Car_clase_id + "', " + CStr(var_factura_nueva) + ",'+', '" + rs!VCHA_SER_SERIE_ID + "','A',0," + CStr(rs!floa_Car_importe_neto) + "," + CStr(rs!floa_car_tipo_cambio) + ",'" + rs!vcha_mon_moneda_id + "',GETDATE(),'" + rs!vcha_cli_clave_id + "',0)"

                              rsaux3.Open "select * from tb_salidas where  VCHA_EMP_EMPRESA_ID = '" + rs!VCHA_EMP_EMPRESA_ID + "' AND VCHA_UOR_UNIDAD_ID = '" + rs!VCHA_UOR_UNIDAD_ID + "' AND VCHA_CAR_DOCUMENTO = 'FA' AND VCHA_SER_SERIE_ID = '" + rs!VCHA_SER_SERIE_ID + "' AND INTE_CAR_NUMERO = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
                              While Not rsaux3.EOF
                                    rsaux2.Open "UPDATE TB_SALIDAS SET INTE_CAR_NUMERO = " + CStr(var_factura_nueva) + " WHERE VCHA_EMP_EMPRESA_ID = '" + rs!VCHA_EMP_EMPRESA_ID + "' AND VCHA_UOR_UNIDAD_ID = '" + rs!VCHA_UOR_UNIDAD_ID + "' AND VCHA_CAR_DOCUMENTO = 'FA' AND VCHA_SER_SERIE_ID = '" + rs!VCHA_SER_SERIE_ID + "' AND INTE_CAR_NUMERO = " + txt_numero + " and vcha_art_articulo_id = '" + rsaux3!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                    rsaux3.MoveNext
                              Wend
                              rsaux3.Close
                              rsaux2.Open "update tb_series set inte_ser_factura = isnull(inte_ser_factura,0) + 1 where vcha_emp_empresa_id = '" + rs!VCHA_EMP_EMPRESA_ID + "' and vcha_uor_unidad_id = '" + rs!VCHA_UOR_UNIDAD_ID + "' and vcha_Ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                              cnn.CommitTrans
                              
                              Open (App.Path & "\factura" + Trim(Str(var_factura_nueva)) + ".txt") For Output As #1
                              'Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
                              'Print #1, Chr(27) + Chr(67) + Chr(44) + Chr(13)
                              'Print #1, ""
                              Print #1, Chr(27) + Chr(64)
                              If var_empresa = "18" Then
                                 Print #1, ""
                              End If
                              Print #1, Spc(92); Str(var_factura_nueva)
                              Print #1, ""
                              Print #1, ""
                              Print #1, Spc(93); "FECHA: "; Format(rs!dtim_Car_fecha, "Short Date")
                              Print #1, ""
                              Print #1, Spc(92); Str(rs!INTE_CAR_PLAZO) + " DIAS DE VENCIMIENTO"
                              var_cliente_str = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                              For var_j = 1 + Len(Trim(var_cliente_str)) To 83
                                  var_cliente_str = var_cliente_str + " "
                              Next var_j
                              var_cliente_str = var_cliente_str + "AGUASCALIENTES, AGS."
                              Print #1, ""
                              Print #1, Spc(10); var_cliente_str
                              var_domicilio_str = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
                              For var_j = 1 + Len(Trim(var_domicilio_str)) To 83
                                  var_domicilio_str = var_domicilio_str + " "
                              Next var_j
                              var_agente_str = ""
                              var_agente_str = IIf(IsNull(rs!vcha_AGE_aGENTE_ID), "", rs!vcha_AGE_aGENTE_ID)
                              For var_j = 1 + Len(Trim(var_agente_str)) To 8
                                  var_agente_str = var_agente_str + " "
                              Next var_j
                              var_agente_str = var_agente_str + IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
                              var_domicilio_str = var_domicilio_str
                              'Print #1, Spc(111); var_agente
                              Print #1, Spc(10); var_domicilio_str
                              var_ciudad_str = ""
                              var_ciudad_str = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                              For var_j = 1 + Len(Trim(var_ciudad_str)) To 37
                                  var_ciudad_str = var_ciudad_str + " "
                              Next var_j
                              var_estado_str = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                              For var_j = 1 + Len(Trim(var_estado_str)) To 46
                                  var_estado_str = var_estado_str + " "
                              Next var_j
                              var_ciudad_str = var_ciudad_str + var_estado_str
                      
                              For var_j = 1 + Len(Trim(var_ciudad_str)) To 14
                                  var_ciudad_str = var_ciudad_str + " "
                              Next var_j
                      
                              var_ciudad_str = var_ciudad_str + var_agente_str
                              var_relacion = "RMV: " + CStr(IIf(IsNull(rs!INTE_EMO_NUMERO_ORIGEN), "", rs!INTE_EMO_NUMERO_ORIGEN))
                              Print #1, Spc(10); var_ciudad_str
                              var_rfc_str = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                              var_rfc_str = "RFC:  " + var_rfc_str
                              For var_j = 1 + Len(Trim(var_rfc_str)) To 89
                                  var_rfc_str = var_rfc_str + " "
                              Next var_j
                              var_rfc_str = var_rfc_str + var_relacion
                              Print #1, Spc(4); var_rfc_str
                              Print #1, Spc(10); IIf(IsNull(rs!VCHA_ESB_ESTABLECIMIENTO_ID), "", rs!VCHA_ESB_ESTABLECIMIENTO_ID)
                              Print #1, ""
                              Print #1, ""
                              var_importe_descuento_1 = 0
                              var_importe_descuento_2 = 0
                              var_importe_descuento_3 = 0
                              var_contador_promociones = 0
                              var_cantidad_total = 0
                              For var_k = 1 To var_renglones_factura
                                  If Not rs.EOF Then
                                     var_linea = ""
                                     var_marca_promocion = " "
                                     var_promocion_1 = IIf(IsNull(rs!floa_sal_promocion_1), 0, rs!floa_sal_promocion_1)
                                     var_promocion_2 = IIf(IsNull(rs!FLOA_SAL_PROMOCION_2), 0, rs!FLOA_SAL_PROMOCION_2)
                                     If var_promocion_1 > 0 Then
                                        var_marca_promocion = "*"
                                        var_contador_promociones = var_contador_promociones + 1
                                     End If
                                     If var_promocion_2 > 0 Then
                                        var_marca_promocion = "*"
                                        var_contador_promociones = var_contador_promociones + 1
                                     End If
                                     var_linea = IIf(IsNull(rs!vcha_agr_agrupador_id), "", rs!vcha_agr_agrupador_id)
                                     For var_j = 1 + Len(Trim(var_linea)) To 15
                                         var_linea = var_linea + " "
                                     Next var_j
                                     var_linea = var_linea + IIf(IsNull(rs!vcha_sal_descripcion_factura), "", rs!vcha_sal_descripcion_factura)
                                     For var_j = 1 + Len(Trim(var_linea)) To 91
                                         var_linea = var_linea + " "
                                     Next var_j
                                     var_linea = var_linea + var_marca_promocion
                                     var_cantidad_str = Format(IIf(IsNull(rs!FLOA_SAL_CANTIDAD), 0, rs!FLOA_SAL_CANTIDAD), "###,###,##0.00")
                                     var_cantidad_total = var_cantidad_total + IIf(IsNull(rs!FLOA_SAL_CANTIDAD), 0, rs!FLOA_SAL_CANTIDAD)
                                     If Len(Trim(var_cantidad_str)) < 14 Then
                                        For var_j = 1 + Len(Trim(var_cantidad_str)) To 14
                                            var_cantidad_str = " " + var_cantidad_str
                                        Next var_j
                                     End If
                                     var_precio = IIf(IsNull(rs!FLOA_sAL_PRECIO), 0, rs!FLOA_sAL_PRECIO)
                                     var_descuento_1 = IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1)
                                     var_descuento_2 = IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2)
                                     var_descuento_3 = IIf(IsNull(rs!floa_car_porcentaje_descuento_3), 0, rs!floa_car_porcentaje_descuento_3)
                                     var_porcentaje = (100 - var_descuento_1) / 100
                                     var_precio = var_precio * var_porcentaje
                                     var_importe_descuento_1_2 = (IIf(IsNull(rs!FLOA_sAL_PRECIO), 0, rs!FLOA_sAL_PRECIO) - var_precio)
                                     var_importe_descuento_1 = var_importe_descuento_1 + (IIf(IsNull(rs!FLOA_sAL_PRECIO), 0, rs!FLOA_sAL_PRECIO) - var_precio)
                                     var_precio = var_precio * ((100 - var_descuento_2) / 100)
                                     var_importe_descuento_2 = var_importe_descuento_2 + (IIf(IsNull(rs!FLOA_sAL_PRECIO), 0, rs!FLOA_sAL_PRECIO) - (var_importe_descuento_1_2 + var_precio))
                                     var_precio = var_precio * ((100 - var_descuento_3) / 100)
                                     var_precio = var_precio / IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)
                                     'var_precio_str = Format(var_precio / IIf(IsNull(rs!cantidad), 0, rs!cantidad), "###,###,##0.00")
                                     var_precio_str = Format(IIf(IsNull(rs!FLOA_sAL_PRECIO), 0, rs!FLOA_sAL_PRECIO) / IIf(IsNull(rs!FLOA_SAL_CANTIDAD), 0, rs!FLOA_SAL_CANTIDAD), "###,###,##0.00")
                                     If Len(Trim(var_precio_str)) < 14 Then
                                        For var_j = 1 + Len(Trim(var_precio_str)) To 14
                                            var_precio_str = " " + var_precio_str
                                        Next var_j
                                     End If
                                     var_linea = var_linea + var_cantidad_str + var_precio_str
                                     var_importe_str = Format((IIf(IsNull(rs!FLOA_sAL_PRECIO), 0, rs!FLOA_sAL_PRECIO)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                     If Len(Trim(var_importe_str)) < 14 Then
                                        For var_j = 1 + Len(Trim(var_importe_str)) To 14
                                            var_importe_str = " " + var_importe_str
                                        Next var_j
                                     End If
                                     var_linea = var_linea + var_importe_str
                              
                                     Print #1, var_linea
                                     rs.MoveNext
                                  Else
                                     Print #1, ""
                                  End If
                               Next var_k
                               Print #1, ""
                               Print #1, ""
                               rs.MoveFirst
                               var_cantidad_total_str = Format(var_cantidad_total, "###,###,##0.00")
                               var_cantidad_letra_str = rs!vcha_car_importe_letra
                               var_importe_descuento_1_str = Format((var_importe_descuento_1 / rs!floa_car_tipo_cambio), "###,###,##0.00")
                         
                               If Len(Trim(var_importe_descuento_1_str)) < 14 Then
                                  For var_j = 1 + Len(Trim(var_importe_descuento_1_str)) To 14
                                      var_importe_descuento_1_str = " " + var_importe_descuento_1_str
                                  Next var_j
                               End If
                               var_importe_descuento_2_str = Format((var_importe_descuento_2 / rs!floa_car_tipo_cambio), "###,###,##0.00")
                               If Len(Trim(var_importe_descuento_2_str)) < 14 Then
                                  For var_j = 1 + Len(Trim(var_importe_descuento_2_str)) To 14
                                      var_importe_descuento_2_str = " " + var_importe_descuento_2_str
                                  Next var_j
                               End If
                               var_linea = "- DESCUENTO DEL " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1))) + "%"
                               If Len(Trim(var_linea)) < 120 Then
                                  For var_j = 1 + Len(Trim(var_linea)) To 120
                                      var_linea = var_linea + " "
                                  Next var_j
                               End If
                               Print #1, var_linea + var_importe_descuento_1_str
                               If var_empresa = "18" Then
                                  var_linea = ""
                               Else
                                  var_linea = "- DESCUENTO POR PAGO OPORTUNO " + Trim(Str(IIf(IsNull(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2), 0, rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2))) + "%"
                               End If
                               If Len(Trim(var_linea)) < 120 Then
                                  For var_j = 1 + Len(Trim(var_linea)) To 120
                                      var_linea = var_linea + " "
                                  Next var_j
                               End If
                               var_linea = var_linea + var_importe_descuento_2_str
                               Print #1, var_linea
                               If var_contador_promociones > 0 Then
                                  Print #1, "PROMOCION EN ARTICULOS MARCADOS CON *"
                               Else
                                  Print #1, ""
                               End If
                               var_rfc_str = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                               var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
                               If Len(Trim(var_linea)) < 120 Then
                                  For var_j = 1 + Len(Trim(var_linea)) To 120
                                      var_x = var_j Mod 2
                                      If var_x >= 1 Then
                                         var_linea = " " + var_linea
                                      Else
                                         var_linea = var_linea + " "
                                      End If
                                  Next var_j
                               End If
                         
                               If Len(Trim(var_rfc_str)) = 0 Then
                                  var_subimporte_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                  If Len(Trim(var_subimporte_str)) < 14 Then
                                     For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                         var_subimporte_str = " " + var_subimporte_str
                                     Next var_j
                                  End If
                                  var_iva_str = "      -        "
                                  For var_j = 1 + Len(Trim(var_iva_str)) To 14
                                      var_iva_str = " " + var_iva_str
                                  Next var_j
                               Else
                                  var_subimporte_str = Format(((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                  If Len(Trim(var_subimporte_str)) < 14 Then
                                     For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                         var_subimporte_str = " " + var_subimporte_str
                                     Next var_j
                                  End If
                                  var_iva_str = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                  If Len(Trim(var_iva_str)) < 14 Then
                                     For var_j = 1 + Len(Trim(var_iva_str)) To 14
                                         var_iva_str = " " + var_iva_str
                                     Next var_j
                                  End If
                               End If
                               var_linea = var_linea + var_iva_str
                        
                               If Len(Trim(var_subimporte_str)) < 14 Then
                                  For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                      var_subimporte_str = " " + var_subimporte_str
                                  Next var_j
                               End If
                               Print #1, Spc(101); var_cantidad_total_str; Spc(6); var_subimporte_str
                       
                               Print #1, var_linea
                               var_importe_str = Format(Round((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), 2), "###,###,##0.00")
                               If Len(Trim(var_importe_str)) < 120 Then
                                  For var_j = 1 + Len(Trim(var_importe_str)) To 14
                                      var_importe_str = " " + var_importe_str
                                  Next var_j
                               End If
                               var_linea = "                                             ESTA FACTURA SERA PAGADA EN UNA SOLA EXHIBICION                            "
                               var_linea = var_linea + var_importe_str
                               Print #1, var_linea
                               
                               Print #1, Spc(4); "AGUASCALIENTES, AGS"; Spc(3); Format(rs!dtim_Car_fecha, "Short Date")
                               
                               var_linea = ""
                               Print #1, Spc(45); var_linea
                               var_dia_str = Day(rs!dtim_Car_fecha)
                               var_mes_str = Month(rs!dtim_Car_fecha)
                               var_año_str = Year(rs!dtim_Car_fecha)
                               var_linea = var_dia
                               If Len(Trim(var_linea)) < 14 Then
                                  For var_j = 1 + Len(Trim(var_linea)) To 14
                                      var_linea = var_linea + " "
                                  Next var_j
                               End If
                               var_linea = var_linea + var_mes_str
                               If Len(Trim(var_linea)) < 50 Then
                                  For var_j = 1 + Len(Trim(var_linea)) To 50
                                      var_linea = var_linea + " "
                                  Next var_j
                               End If
                               Print #1, Spc(70); var_linea
                               var_linea = ""
                               var_linea = var_año_str
                               If Len(Trim(var_linea)) < 15 Then
                                  For var_j = 1 + Len(Trim(var_linea)) To 15
                                      var_linea = var_linea + " "
                                  Next var_j
                               End If
                               var_linea = var_linea + var_importe_str
                               If Len(Trim(var_linea)) < 24 Then
                                  For var_j = 1 + Len(Trim(var_linea)) To 24
                                      var_linea = " " + var_linea
                                  Next var_j
                               End If
                               var_linea = var_linea + " " + var_cantidad_letra_str
                               Print #1, Spc(2); var_linea
                               Print #1, ""
                               Print #1, ""
                               Print #1, ""
                               Print #1, Spc(5); Trim(IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE))
                               Print #1, Spc(5); Trim(IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " " + IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre))
                               Print #1, Spc(5); Trim(IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre) + " " + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre))
                               Close #1
                               Open (App.Path & "\factura" + Trim(Str(var_factura_nueva)) + ".bat") For Output As #2
                               var_Archivo = App.Path & "\factura" + Trim(Str(var_factura_nueva)) + ".bat"
                               Print #2, "copy " + App.Path + "\factura" + Trim(Str(var_factura_nueva)) + ".txt lpt1"
                               Close #2
                               x = Shell(var_Archivo, vbHide)
                           End If
                        End If
                     End If
                  End If
                  rsaux.Close
               End If
               rs.Close
            End If
         End If
         If var_tipo_facturacion = "" Or var_tipo_facturacion = " " Then
            si = MsgBox("¿Deseas reimprimir la factura " + txt_numero + "?", vbYesNo, "ATENCION")
            If si = 6 Then
               rs.Open "select * from tb_encabezado_cartera where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'FA' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  rsaux2.Open "select * from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                  var_factura_nueva = rsaux2!inte_ser_factura
                  rsaux2.Close
                  si = MsgBox("¿Se va a imprimir la factura " + Str(var_factura_nueva) + "?", vbYesNo, "ATENCION")
                  If si = 6 Then
                     si = MsgBox("Confirmar la reimpresión de la factura " + Str(var_factura_nueva), vbYesNo, "ATENCION")
                     If si = 6 Then
                        rsaux2.Open "select * from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                        If var_factura_nueva <> rsaux2!inte_ser_factura Then
                           MsgBox "El número de la factura ya cambio y el proceso de reimpresión a cambiado", vbOKOnly, "ATENCION"
                           rsaux2.Close
                        Else
                           rsaux2.Close
                           var_cadena = ""
                           var_cadena = "INSERT INTO [TB_ENCABEZADO_CARTERA] ([VCHA_EMP_EMPRESA_ID], [VCHA_UOR_UNIDAD_ID], [VCHA_CAR_TIPO_DOCUMENTO], [VCHA_CAR_DOCUMENTO], [VCHA_CAR_CLASE_ID], [INTE_CAR_NUMERO], [CHAR_CAR_AFECTACION], [VCHA_ALM_ALMACEN_ID], [VCHA_MOV_MOVIMIENTO_ID], [INTE_EMO_NUMERO], [DTIM_CAR_FECHA], [VCHA_AGE_AGENTE_ID], [VCHA_GAC_GRUPO_ACTUAL_ID], [VCHA_GRE_GRUPO_REAL_ID], [VCHA_TIT_TITULAR_ID], [VCHA_CLI_CLAVE_ID], [VCHA_ESB_ESTABLECIMIENTO_ID], [INTE_CAR_PLAZO], [FLOA_CAR_PORCENTAJE_IVA], [FLOA_CAR_PORCENTAJE_IMPUESTO_1], [FLOA_CAR_PORCENTAJE_IMPUESTO_2], [FLOA_CAR_PORCENTAJE_DESCUENTO_1], [FLOA_CAR_PORCENTAJE_DESCUENTO_2], [FLOA_CAR_PORCENTAJE_DESCUENTO_3], [FLOA_CAR_IMPORTE_TOTAL], [FLOA_CAR_IMPORTE_IVA], [FLOA_CAR_IMPORTE_IMPUESTO_1], [FLOA_CAR_IMPORTE_IMPUESTO_2], [FLOA_CAR_IMPORTE_DESCUENTO_1], [FLOA_CAR_IMPORTE_DESCUENTO_2], [FLOA_CAR_IMPORTE_DESCUENTO_3], [FLOA_CAR_SUBIMPORTE], [FLOA_CAR_IMPORTE_NETO], [VCHA_CAR_IMPORTE_LETRA], [VCHA_AUD_USUARIO], [VCHA_AUD_MAQUINA], "
                           var_cadena = var_cadena + "[VCHA_AUD_FECHA], [FLOA_CAR_SALDO], [DTIM_CAR_FECHA_VENCIMIENTO], [DTIM_CAR_FECHA_ENTREGA], [VCHA_MON_MONEDA_ID], [FLOA_CAR_TIPO_CAMBIO], [VCHA_SER_SERIE_ID], [CHAR_CAR_ESTATUS], [CHAR_CAR_TIPO_FACTURACION]) Values ('" + IIf(IsNull(rs!VCHA_EMP_EMPRESA_ID), "", rs!VCHA_EMP_EMPRESA_ID) + "', '" + rs!VCHA_UOR_UNIDAD_ID + "', 'FA', '" + rs!vcha_Car_documento + "', '" + rs!vcha_Car_clase_id + "', " + CStr(var_factura_nueva) + ", '" + rs!char_car_afectacion
                           var_cadena = var_cadena + "', '" + rs!vcha_alm_almacen_id + "', '" + rs!vcha_mov_movimiento_id + "', " + CStr(rs!INTE_EMO_NUMERO) + ", getdate(),  '" + rs!vcha_AGE_aGENTE_ID + "', '" + rs!VCHA_GAC_GRUPO_aCTUAL_ID + "', '" + rs!vcha_gre_grupo_real_id + "', '" + rs!vcha_tit_titular_id + "', '" + rs!vcha_cli_clave_id + "', '" + rs!VCHA_ESB_ESTABLECIMIENTO_ID + "', " + CStr(rs!INTE_CAR_PLAZO) + ", " + CStr(rs!floa_car_porcentaje_iva) + ", " + CStr(rs!FLOA_CAR_PORCENTAJE_IMPUESTO_1) + ", " + CStr(rs!FLOA_CAR_PORCENTAJE_IMPUESTO_2) + ", " + CStr(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1) + ", " + CStr(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2) + ", " + CStr(rs!floa_car_porcentaje_descuento_3) + ", " + CStr(rs!FLOA_CAR_IMPORTE_TOTAL) + ", " + CStr(rs!floa_car_importe_iva) + ", " + CStr(rs!FLOA_CAR_IMPORTE_IMPUESTO_1) + ", " + CStr(rs!FLOA_CAR_IMPORTE_IMPUESTO_2) + ", " + CStr(rs!FLOA_CAR_IMPORTE_DESCUENTO_1) + ", " + CStr(rs!FLOA_CAR_IMPORTE_DESCUENTO_2) + ", "
                           var_cadena = var_cadena + CStr(rs!FLOA_CAR_IMPORTE_DESCUENTO_3) + "," + CStr(rs!floa_car_subimporte) + ", " + CStr(rs!floa_Car_importe_neto) + ", '" + rs!vcha_car_importe_letra + "', '" + rs!vcha_aud_usuario + "', '" + rs!vcha_aud_maquina + "', getdate(), 0, null, null, '" + rs!vcha_mon_moneda_id + "', " + CStr(rs!floa_car_tipo_cambio) + ", '" + rs!VCHA_SER_SERIE_ID + "', '', '') "
                           cnn.BeginTrans
                           rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                           rsaux2.Open "INSERT INTO [TB_ESTADO_CUENTA] ([VCHA_EMP_EMPRESA_ID], [VCHA_ECU_SERIE_CARGO], [VCHA_ECU_MOVIMIENTO_CARGO], [INTE_ECU_NUMERO_CARGO], [FLOA_ECU_IMPORTE_CARGO], [FLOA_ECU_IMPORTE_ABONO]) Values ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!VCHA_SER_SERIE_ID + "', 'FA', " + CStr(var_factura_nueva) + ", " + CStr(rs!floa_Car_importe_neto) + ", 0) ", cnn, adOpenDynamic, adLockOptimistic
                           rsaux2.Open "INSERT INTO TB_SECUENCIA_FACTURACION (VCHA_EMP_EMPRESA_ID, VCHA_SER_SERIE_ID, INTE_SFA_NUMERO_ANTERIOR, INTE_SFA_NUMERO_ACTUAL) VALUES ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!VCHA_SER_SERIE_ID + "', " + CStr(var_factura_nueva) + ", " + CStr(var_factura_nueva) + ")", cnn, adOpenDynamic, adLockOptimistic
                           rsaux2.Open "UPDATE TB_SECUENCIA_FACTURACION SET INTE_SFA_NUMERO_ACTUAL = " + CStr(var_factura_nueva) + " WHERE VCHA_EMP_EMPRESA_ID = '" + rs!VCHA_EMP_EMPRESA_ID + "' AND VCHA_SER_SERIE_ID = '" + rs!VCHA_SER_SERIE_ID + "' AND  INTE_SFA_NUMERO_ANTERIOR = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
                           rsaux2.Open "INSERT INTO TB_INVENTARIO_DOCUMENTOS (VCHA_EMP_EMPRESA_ID, VCHA_AGE_AGENTE_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, VCHA_CAR_CLASE_ID, INTE_CAR_NUMERO, CHAR_CAR_AFECTACION, VCHA_SER_SERIE_ID, CHAR_IDO_ESTATUS, FLOA_IDO_CANTIDAD, FLOA_CAR_IMPORTE_NETO, FLOA_CAR_TIPO_CAMBIO, VCHA_MON_MONEDA_ID, DTIM_IDO_FECHA_ENTRAGA, VCHA_CLI_CLAVE_ID, INTE_EMB_EMBARQUE) VALUES ('" + rs!VCHA_EMP_EMPRESA_ID + "','" + rs!vcha_AGE_aGENTE_ID + "', 'FA', 'FA', '" + rs!vcha_Car_clase_id + "', " + CStr(var_factura_nueva) + ",'+', '" + rs!VCHA_SER_SERIE_ID + "','A',0," + CStr(rs!floa_Car_importe_neto) + "," + CStr(rs!floa_car_tipo_cambio) + ",'" + rs!vcha_mon_moneda_id + "',GETDATE(),'" + rs!vcha_cli_clave_id + "',0)"
                           
                           rsaux2.Open "update tb_series set inte_ser_factura = isnull(inte_ser_factura,0) + 1 where vcha_emp_empresa_id = '" + rs!VCHA_EMP_EMPRESA_ID + "' and vcha_uor_unidad_id = '" + rs!VCHA_UOR_UNIDAD_ID + "' and vcha_Ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                           cnn.CommitTrans
                           rs.Close
                           rs.Open "select * from VW_DOCUMENTOS_IMPRESION where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'FA' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + CStr(var_factura_nueva), cnn, adOpenDynamic, adLockOptimistic
                           If Not rs.EOF Then
                              Open (App.Path & "\factura" + Trim(var_factura_nueva) + ".txt") For Output As #1
                              Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
                              'Print #1, Chr(27) + Chr(67) + Chr(44) + Chr(13)
                              'Print #1, ""
                              Print #1, Spc(92); Str(rs!INTE_CAR_NUMERO)
                              Print #1, ""
                              Print #1, ""
                              Print #1, Spc(93); "FECHA: "; Format(rs!dtim_Car_fecha, "Short Date")
                              Print #1, ""
                              Print #1, Spc(92); Str(rs!INTE_CAR_PLAZO) + " DIAS DE VENCIMIENTO"
                              var_cliente_str = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                              For var_j = 1 + Len(Trim(var_cliente_str)) To 83
                                  var_cliente_str = var_cliente_str + " "
                              Next var_j
                              If var_unidad_organizacional = "21" Then
                                 var_cliente_str = var_cliente_str + "MEXICO, D.F."
                              Else
                                 var_cliente_str = var_cliente_str + "AGUASCALIENTES, AGS."
                              End If
                              Print #1, ""
                              Print #1, Spc(10); var_cliente_str
                              var_domicilio_str = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
                              For var_j = 1 + Len(Trim(var_domicilio_str)) To 83
                                  var_domicilio_str = var_domicilio_str + " "
                              Next var_j
                              var_agente_str = ""
                              var_agente_str = IIf(IsNull(rs!vcha_AGE_aGENTE_ID), "", rs!vcha_AGE_aGENTE_ID)
                              For var_j = 1 + Len(Trim(var_agente_str)) To 8
                                  var_agente_str = var_agente_str + " "
                              Next var_j
                              var_agente_str = var_agente_str + IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
                              var_domicilio_str = var_domicilio_str
                              'Print #1, Spc(111); var_agente
                              Print #1, Spc(10); var_domicilio_str
                              var_ciudad_str = ""
                              var_ciudad_str = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                              For var_j = 1 + Len(Trim(var_ciudad_str)) To 37
                                  var_ciudad_str = var_ciudad_str + " "
                              Next var_j
                              var_estado_str = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                              For var_j = 1 + Len(Trim(var_estado_str)) To 46
                                  var_estado_str = var_estado_str + " "
                              Next var_j
                              var_ciudad_str = var_ciudad_str + var_estado_str
                         
                              For var_j = 1 + Len(Trim(var_ciudad_str)) To 14
                                  var_ciudad_str = var_ciudad_str + " "
                              Next var_j
                      
                              var_ciudad_str = var_ciudad_str + var_agente_str
                              var_relacion = ""
                              Print #1, Spc(10); var_ciudad_str
                              var_rfc_str = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                              var_rfc_str = "RFC:  " + var_rfc_str
                              For var_j = 1 + Len(Trim(var_rfc_str)) To 89
                                  var_rfc_str = var_rfc_str + " "
                              Next var_j
                              var_rfc_str = var_rfc_str + var_relacion
                              Print #1, Spc(4); var_rfc_str
                              Print #1, Spc(10); IIf(IsNull(rs!VCHA_ESB_ESTABLECIMIENTO_ID), "", rs!VCHA_ESB_ESTABLECIMIENTO_ID)
                              Print #1, ""
                              Print #1, ""
                              var_importe_descuento_1 = 0
                              var_importe_descuento_2 = 0
                              var_importe_descuento_3 = 0
                              var_contador_promociones = 0
                              var_cantidad_total = 0
                              For var_k = 1 To var_renglones_factura
                                  var_linea = ""
                                  Print #1, ""
                              Next var_k
                              Print #1, ""
                              Print #1, ""
                              rs.MoveFirst
                              var_cantidad_total_str = Format(var_cantidad_total, "###,###,##0.00")
                              var_cantidad_letra_str = rs!vcha_car_importe_letra
                              var_importe_descuento_1_str = Format(var_importe_descuento_1, "###,###,##0.00")
                             
                              If Len(Trim(var_importe_descuento_1_str)) < 14 Then
                                 For var_j = 1 + Len(Trim(var_importe_descuento_1_str)) To 14
                                     var_importe_descuento_1_str = " " + var_importe_descuento_1_str
                                 Next var_j
                              End If
                              var_importe_descuento_2_str = Format(var_importe_descuento_2, "###,###,##0.00")
                              If Len(Trim(var_importe_descuento_2_str)) < 14 Then
                                 For var_j = 1 + Len(Trim(var_importe_descuento_2_str)) To 14
                                     var_importe_descuento_2_str = " " + var_importe_descuento_2_str
                                 Next var_j
                              End If
                              var_linea = ""
                              If Len(Trim(var_linea)) < 120 Then
                                 For var_j = 1 + Len(Trim(var_linea)) To 120
                                     var_linea = var_linea + " "
                                 Next var_j
                              End If
                              Print #1, var_linea + var_importe_descuento_1_str
                              var_linea = ""
                              If Len(Trim(var_linea)) < 120 Then
                                 For var_j = 1 + Len(Trim(var_linea)) To 120
                                     var_linea = var_linea + " "
                                 Next var_j
                              End If
                              var_linea = var_linea + var_importe_descuento_2_str
                              Print #1, var_linea
                              Print #1, ""
                              var_rfc_str = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                              var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
                              If Len(Trim(var_linea)) < 120 Then
                                 For var_j = 1 + Len(Trim(var_linea)) To 120
                                     var_x = var_j Mod 2
                                     If var_x >= 1 Then
                                        var_linea = " " + var_linea
                                     Else
                                        var_linea = var_linea + " "
                                     End If
                                  Next var_j
                              End If
                         
                              If Len(Trim(var_rfc_str)) = 0 Then
                                 var_subimporte_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                 If Len(Trim(var_subimporte_str)) < 14 Then
                                    For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                        var_subimporte_str = " " + var_subimporte_str
                                    Next var_j
                                  End If
                                  var_iva_str = "      -        "
                                  For var_j = 1 + Len(Trim(var_iva_str)) To 14
                                      var_iva_str = " " + var_iva_str
                                  Next var_j
                               Else
                                  var_subimporte_str = Format(((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                  If Len(Trim(var_subimporte_str)) < 14 Then
                                     For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                         var_subimporte_str = " " + var_subimporte_str
                                     Next var_j
                                  End If
                                  var_iva_str = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                  If Len(Trim(var_iva_str)) < 14 Then
                                     For var_j = 1 + Len(Trim(var_iva_str)) To 14
                                         var_iva_str = " " + var_iva_str
                                     Next var_j
                                  End If
                               End If
                               var_linea = var_linea + var_iva_str
                         
                               If Len(Trim(var_subimporte_str)) < 14 Then
                                  For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                      var_subimporte_str = " " + var_subimporte_str
                                  Next var_j
                               End If
                               Print #1, Spc(101); var_cantidad_total_str; Spc(6); var_subimporte_str
                      
                               Print #1, var_linea
                               var_importe_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                               If Len(Trim(var_importe_str)) < 120 Then
                                  For var_j = 1 + Len(Trim(var_importe_str)) To 14
                                      var_importe_str = " " + var_importe_str
                                  Next var_j
                               End If
                               var_linea = "                                             ESTA FACTURA SERA PAGADA EN UNA SOLA EXHIBICION                            "
                               var_linea = var_linea + var_importe_str
                               Print #1, var_linea
                               If var_unidad_organizacional = "21" Then
                                  Print #1, Spc(4); "MEXICO, D.F."; Spc(3); Format(rs!dtim_Car_fecha, "Short Date")
                               Else
                                  Print #1, Spc(4); "AGUASCALIENTES, AGS."; Spc(3); Format(rs!dtim_Car_fecha, "Short Date")
                               End If
                               var_linea = ""
                               Print #1, Spc(45); var_linea
                               var_dia_str = Day(rs!dtim_Car_fecha)
                               var_mes_str = Month(rs!dtim_Car_fecha)
                               var_año_str = Year(rs!dtim_Car_fecha)
                               var_linea = var_dia
                               If Len(Trim(var_linea)) < 14 Then
                                  For var_j = 1 + Len(Trim(var_linea)) To 14
                                      var_linea = var_linea + " "
                                  Next var_j
                               End If
                               var_linea = var_linea + var_mes_str
                               If Len(Trim(var_linea)) < 50 Then
                                  For var_j = 1 + Len(Trim(var_linea)) To 50
                                      var_linea = var_linea + " "
                                  Next var_j
                               End If
                               Print #1, Spc(70); var_linea
                               var_linea = ""
                               var_linea = var_año_str
                               If Len(Trim(var_linea)) < 15 Then
                                  For var_j = 1 + Len(Trim(var_linea)) To 15
                                      var_linea = var_linea + " "
                                  Next var_j
                               End If
                               var_linea = var_linea + var_importe_str
                               If Len(Trim(var_linea)) < 24 Then
                                  For var_j = 1 + Len(Trim(var_linea)) To 24
                                      var_linea = " " + var_linea
                                  Next var_j
                               End If
                               var_linea = var_linea + " " + var_cantidad_letra_str
                               Print #1, Spc(2); var_linea
                               Print #1, ""
                               Print #1, ""
                               Print #1, ""
                               Print #1, Spc(5); Trim(IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE))
                               Print #1, Spc(5); Trim(IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " " + IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre))
                               Print #1, Spc(5); Trim(IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre) + " " + IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre))
                               Close #1
                               Open (App.Path & "\factura" + Trim(var_factura_nueva) + ".bat") For Output As #2
                               var_Archivo = App.Path & "\factura" + Trim(var_factura_nueva) + ".bat"
                               Print #2, "copy " + App.Path + "\factura" + Trim(var_factura_nueva) + ".txt lpt1"
                               Close #2
                               x = Shell(var_Archivo, vbHide)
                               '''factura vieja
                            End If
                         End If
                      End If
                   End If
                   rs.Close
               Else
                  rs.Close
                  MsgBox "La factura " + txt_numero + " no existe", vbOKOnly, "ATENCION"
               End If
           End If
         End If
      End If
      If txt_documento = "NG" Then
         si = MsgBox("¿Deseas reimprimir la nota de cargo " + txt_numero + "?", vbYesNo, "ATENCION")
         If si = 6 Then
            rs.Open "select * from tb_encabezado_cartera where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_tipo_documento = 'NG' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               rsaux2.Open "select * from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
               var_factura_nueva = IIf(IsNull(rsaux2!inte_ser_nota_Cargo), 0, rsaux2!inte_ser_nota_Cargo)
               var_factura_nueva = var_factura_nueva + 1
               rsaux2.Close
               si = MsgBox("¿Se va a imprimir la nota de cargo " + Str(var_factura_nueva) + "?", vbYesNo, "ATENCION")
               If si = 6 Then
                  si = MsgBox("Confirmar la reimpresión de la nota de cargo " + Str(var_factura_nueva), vbYesNo, "ATENCION")
                  If si = 6 Then
                     rsaux2.Open "select * from tb_series where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                     If var_factura_nueva <> rsaux2!inte_ser_nota_Cargo Then
                        MsgBox "El número de la nota de cargo ya cambio y el proceso de reimpresión se a cancelado", vbOKOnly, "ATENCION"
                        rsaux2.Close
                     Else
                        rsaux2.Close
                        var_cadena = ""
                        var_cadena = "INSERT INTO [TB_ENCABEZADO_CARTERA] ([VCHA_EMP_EMPRESA_ID], [VCHA_UOR_UNIDAD_ID], [VCHA_CAR_TIPO_DOCUMENTO], [VCHA_CAR_DOCUMENTO], [VCHA_CAR_CLASE_ID], [INTE_CAR_NUMERO], [CHAR_CAR_AFECTACION], [VCHA_ALM_ALMACEN_ID], [VCHA_MOV_MOVIMIENTO_ID], [INTE_EMO_NUMERO], [DTIM_CAR_FECHA], [VCHA_AGE_AGENTE_ID], [VCHA_GAC_GRUPO_ACTUAL_ID], [VCHA_GRE_GRUPO_REAL_ID], [VCHA_TIT_TITULAR_ID], [VCHA_CLI_CLAVE_ID], [VCHA_ESB_ESTABLECIMIENTO_ID], [INTE_CAR_PLAZO], [FLOA_CAR_PORCENTAJE_IVA], [FLOA_CAR_PORCENTAJE_IMPUESTO_1], [FLOA_CAR_PORCENTAJE_IMPUESTO_2], [FLOA_CAR_PORCENTAJE_DESCUENTO_1], [FLOA_CAR_PORCENTAJE_DESCUENTO_2], [FLOA_CAR_PORCENTAJE_DESCUENTO_3], [FLOA_CAR_IMPORTE_TOTAL], [FLOA_CAR_IMPORTE_IVA], [FLOA_CAR_IMPORTE_IMPUESTO_1], [FLOA_CAR_IMPORTE_IMPUESTO_2], [FLOA_CAR_IMPORTE_DESCUENTO_1], [FLOA_CAR_IMPORTE_DESCUENTO_2], [FLOA_CAR_IMPORTE_DESCUENTO_3], [FLOA_CAR_SUBIMPORTE], [FLOA_CAR_IMPORTE_NETO], [VCHA_CAR_IMPORTE_LETRA], [VCHA_AUD_USUARIO], [VCHA_AUD_MAQUINA], "
                        var_cadena = var_cadena + "[VCHA_AUD_FECHA], [FLOA_CAR_SALDO], [DTIM_CAR_FECHA_VENCIMIENTO], [DTIM_CAR_FECHA_ENTREGA], [VCHA_MON_MONEDA_ID], [FLOA_CAR_TIPO_CAMBIO], [VCHA_SER_SERIE_ID], [CHAR_CAR_ESTATUS], [CHAR_CAR_TIPO_FACTURACION]) Values ('" + IIf(IsNull(rs!VCHA_EMP_EMPRESA_ID), "", rs!VCHA_EMP_EMPRESA_ID) + "', '" + rs!VCHA_UOR_UNIDAD_ID + "', 'NG', '" + rs!vcha_Car_documento + "', '" + rs!vcha_Car_clase_id + "', " + CStr(var_factura_nueva) + ", '" + rs!char_car_afectacion
                        var_cadena = var_cadena + "', '" + rs!vcha_alm_almacen_id + "', '" + rs!vcha_mov_movimiento_id + "', " + CStr(rs!INTE_EMO_NUMERO) + ", getdate(),  '" + rs!vcha_AGE_aGENTE_ID + "', '" + rs!VCHA_GAC_GRUPO_aCTUAL_ID + "', '" + rs!vcha_gre_grupo_real_id + "', '" + rs!vcha_tit_titular_id + "', '" + rs!vcha_cli_clave_id + "', '" + rs!VCHA_ESB_ESTABLECIMIENTO_ID + "', " + CStr(rs!INTE_CAR_PLAZO) + ", " + CStr(rs!floa_car_porcentaje_iva) + ", " + CStr(rs!FLOA_CAR_PORCENTAJE_IMPUESTO_1) + ", " + CStr(rs!FLOA_CAR_PORCENTAJE_IMPUESTO_2) + ", " + CStr(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_1) + ", " + CStr(rs!FLOA_CAR_PORCENTAJE_DESCUENTO_2) + ", " + CStr(rs!floa_car_porcentaje_descuento_3) + ", " + CStr(rs!FLOA_CAR_IMPORTE_TOTAL) + ", " + CStr(rs!floa_car_importe_iva) + ", " + CStr(rs!FLOA_CAR_IMPORTE_IMPUESTO_1) + ", " + CStr(rs!FLOA_CAR_IMPORTE_IMPUESTO_2) + ", " + CStr(rs!FLOA_CAR_IMPORTE_DESCUENTO_1) + ", " + CStr(rs!FLOA_CAR_IMPORTE_DESCUENTO_2) + ", "
                        var_cadena = var_cadena + CStr(rs!FLOA_CAR_IMPORTE_DESCUENTO_3) + "," + CStr(rs!floa_car_subimporte) + ", " + CStr(rs!floa_Car_importe_neto) + ", '" + rs!vcha_car_importe_letra + "', '" + rs!vcha_aud_usuario + "', '" + rs!vcha_aud_maquina + "', getdate(), 0, null, null, '" + rs!vcha_mon_moneda_id + "', " + CStr(rs!floa_car_tipo_cambio) + ", '" + rs!VCHA_SER_SERIE_ID + "', '', '') "
                        cnn.BeginTrans
                        rsaux2.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                        rsaux2.Open "INSERT INTO [TB_ESTADO_CUENTA] ([VCHA_EMP_EMPRESA_ID], [VCHA_ECU_SERIE_CARGO], [VCHA_ECU_MOVIMIENTO_CARGO], [INTE_ECU_NUMERO_CARGO], [FLOA_ECU_IMPORTE_CARGO], [FLOA_ECU_IMPORTE_ABONO]) Values ('" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!VCHA_SER_SERIE_ID + "', 'NC', " + CStr(var_factura_nueva) + ", " + CStr(rs!floa_Car_importe_neto) + ", 0) ", cnn, adOpenDynamic, adLockOptimistic
                        rsaux2.Open "update tb_series set inte_ser_nota_cargo = isnull(inte_ser_nota_cargo,0) + 1 where vcha_emp_empresa_id = '" + rs!VCHA_EMP_EMPRESA_ID + "' and vcha_uor_unidad_id = '" + rs!VCHA_UOR_UNIDAD_ID + "' and vcha_Ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                        cnn.CommitTrans
                        rs.Close
                        rs.Open "select * from vw_notas_cargo where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'NC' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + CStr(var_factura_nueva), cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
'''''''''''''''  IMPRESION DE LA NOTA DE CARGO
                          Open (App.Path & "\nota_cargo" + Trim(CStr(var_factura_nueva)) + ".txt") For Output As #1
                          Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
                          Print #1, Spc(92); Str(rs!INTE_CAR_NUMERO)
                          Print #1, ""
                          Print #1, Spc(93); "FECHA: "; Format(rs!dtim_Car_fecha, "Short Date")
                          var_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                          For var_j = 1 + Len(Trim(var_cliente)) To 83
                              var_cliente = var_cliente + " "
                          Next var_j
                          var_cliente = var_cliente + "AGUASCALIENTES, AGS."
                          Print #1, ""
                          Print #1, Spc(10); var_cliente
                          var_domicilio = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " C.P. " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
                          For var_j = 1 + Len(Trim(var_domicilio)) To 83
                              var_domicilio = var_domicilio + " "
                          Next var_j
                          var_agente = ""
                          var_agente = IIf(IsNull(rs!vcha_AGE_aGENTE_ID), "", rs!vcha_AGE_aGENTE_ID)
                          For var_j = 1 + Len(Trim(var_agente)) To 8
                              var_agente = var_agente + " "
                          Next var_j
                          var_agente = var_agente + IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
                          var_domicilio = var_domicilio
                          Print #1, Spc(10); var_domicilio
                          var_ciudad = ""
                          var_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                          For var_j = 1 + Len(Trim(var_ciudad)) To 37
                              var_ciudad = var_ciudad + " "
                          Next var_j
                          var_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
                          For var_j = 1 + Len(Trim(var_estado)) To 46
                              var_estado = var_estado + " "
                          Next var_j
                          var_ciudad = var_ciudad + var_estado
                                 
                          For var_j = 1 + Len(Trim(var_ciudad)) To 14
                              var_ciudad = var_ciudad + " "
                          Next var_j
                           
                          var_ciudad = var_ciudad + var_agente
                            
                          Print #1, Spc(10); var_ciudad
                          var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                          var_rfc = "RFC:  " + var_rfc
                          For var_j = 1 + Len(Trim(var_rfc)) To 89
                              var_rfc = var_rfc + " "
                          Next var_j
                          var_rfc = var_rfc
                          Print #1, Spc(4); var_rfc
                          Print #1, ""
                          Print #1, ""
                          var_linea = "NC" + Str(rs!INTE_CAR_NUMERO) + " " + rs!vcha_Car_nombre
                          If Len(Trim(var_linea)) < 108 Then
                             For var_j = 1 + Len(Trim(var_linea)) To 108
                                 var_linea = var_linea + " "
                             Next var_j
                          End If
                          If Len(Trim(var_rfc)) = 0 Then
                             var_importe_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                          Else
                             var_importe_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)) / (1 + (rs!floa_car_porcentaje_iva / 100)), "###,###,##0.00")
                          End If
                          If Len(Trim(var_importe_str)) < 14 Then
                             For var_j = 1 + Len(Trim(var_importe_str)) To 14
                                 var_importe_str = " " + var_importe_str
                             Next var_j
                          End If
                          var_linea = var_linea + var_importe_str
                          Print #1, var_linea
                          Print #1, ""
                          Print #1, ""
                          Print #1, ""
                          Print #1, ""
                          Print #1, ""
                          Print #1, ""
                          var_cantidad_letra = rs!vcha_car_importe_letra
                          
                          var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
                          If Len(Trim(var_linea)) < 93 Then
                             For var_j = 1 + Len(Trim(var_linea)) To 93
                                 var_linea = var_linea + " "
                             Next var_j
                          End If
                          
                          var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                          
                          If Len(Trim(var_rfc)) = 0 Then
                             var_subimporte_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                             If Len(Trim(var_subimporte_str)) < 14 Then
                                For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                    var_subimporte_str = " " + var_subimporte_str
                                Next var_j
                             End If
                             var_iva = "      -        "
                             For var_j = 1 + Len(Trim(var_iva_str)) To 14
                                 var_iva_str = " " + var_iva_str
                             Next var_j
                          Else
                             var_subimporte_str = Format(((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                             If Len(Trim(var_subimporte_str)) < 14 Then
                                For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                    var_subimporte_str = " " + var_subimporte_str
                                Next var_j
                             End If
                             var_iva_str = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                             If Len(Trim(var_iva_str)) < 14 Then
                                For var_j = 1 + Len(Trim(var_iva_str)) To 14
                                    var_iva_str = " " + var_iva_str
                                Next var_j
                             End If
                          End If
                          var_linea = var_linea + "           " + var_subimporte_str
                          Print #1, Spc(4); var_linea
                          Print #1, Spc(108); var_iva_str
                          var_importe_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                          If Len(Trim(var_importe_str)) < 14 Then
                             For var_j = 1 + Len(Trim(var_importe_str)) To 14
                                 var_importe_str = " " + var_importe_str
                             Next var_j
                          End If
                          Print #1, Spc(108); var_importe_str
                          Print #1, ""
                          Print #1, Spc(4); "ESTA DOCUMENTO SERA PAGADO EN UNA SOLA EXHIBICION"
                          Print #1, ""
                          Print #1, Spc(85); "SISTEMAS"
                          Close #1
                          
                          Open (App.Path & "\nota_cargo" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".bat") For Output As #2
                          var_Archivo = App.Path & "\nota_cargo" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".bat"
                          Print #2, "copy " + App.Path + "\nota_cargo" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".txt lpt1"
                          Close #2
                          x = Shell(var_Archivo, vbHide)
 ''''''''''''
                        End If
                        rs.Close
                     End If
                  End If
               End If
            End If
         End If
      End If
      If txt_documento = "NC" Then
         Call reimprimir_notas_credito
      End If
   Else
      MsgBox "La " + cmb_documentos + " número " + txt_numero + " no puede ser reimpresa ya que no a sido cancelada", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_reimprimir_Click()
   Dim var_rfc As String
   var_serie = Me.txt_serie
   rs.Open "select * from tb_principal where vcha_emp_empresa_id = '" + var_empresa + "'"
   var_numero_renglones = rs!INTE_PRI_RENGLONES_NOTA_CREDITO
   var_tolerancia_saldo = rs!FLOA_PRI_TOLERANCIA_SALDOS
   var_serie = Me.txt_serie
   var_factura_nueva = CDbl(Me.txt_numero)
   rs.Close
   If Me.txt_documento = "NC" Then
   si = MsgBox("¿Deseas reimprimir la nota de crédito " + txt_numero + "?", vbYesNo, "ATENCION")
   If si = 6 Then
      rs.Open "select * from tb_encabezado_cartera where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_tipo_documento = '" + txt_documento + "' and vcha_Ser_serie_id = '" + var_serie + "' and inte_Car_numero = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         rs.Close
         If var_tipo_nota_credito = "DF" Then
            rs.Open "select * from VW_DOCUMENTOS_IMPRESION where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'DF' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(var_factura_nueva), cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               Open (App.Path & "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".txt") For Output As #1
               Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
               Print #1, Spc(92); Str(rs!INTE_CAR_NUMERO)
               Print #1, ""
               Print #1, Spc(93); "FECHA: "; Format(rs!dtim_Car_fecha, "Short Date")
               var_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
               For var_j = 1 + Len(Trim(var_cliente)) To 83
                  var_cliente = var_cliente + " "
               Next var_j
               var_cliente = var_cliente + "AGUASCALIENTES, AGS."
               Print #1, ""
               Print #1, Spc(10); var_cliente
               var_domicilio = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " C.P. " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
               For var_j = 1 + Len(Trim(var_domicilio)) To 83
                   var_domicilio = var_domicilio + " "
               Next var_j
               var_agente = ""
               var_agente = IIf(IsNull(rs!vcha_AGE_aGENTE_ID), "", rs!vcha_AGE_aGENTE_ID)
               For var_j = 1 + Len(Trim(var_agente)) To 8
                   var_agente = var_agente + " "
               Next var_j
               var_agente = var_agente + IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
               var_domicilio = var_domicilio
               Print #1, Spc(10); var_domicilio
               var_ciudad = ""
               var_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
               For var_j = 1 + Len(Trim(var_ciudad)) To 37
                   var_ciudad = var_ciudad + " "
               Next var_j
               var_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
               For var_j = 1 + Len(Trim(var_estado)) To 46
                  var_estado = var_estado + " "
               Next var_j
               var_ciudad = var_ciudad + var_estado
                               
               For var_j = 1 + Len(Trim(var_ciudad)) To 14
                  var_ciudad = var_ciudad + " "
               Next var_j
                       
               var_ciudad = var_ciudad + var_agente
                         
               Print #1, Spc(10); var_ciudad
               var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
               var_rfc = "RFC:  " + var_rfc
               For var_j = 1 + Len(Trim(var_rfc)) To 89
                   var_rfc = var_rfc + " "
               Next var_j
               var_rfc = var_rfc
               Print #1, Spc(4); var_rfc
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               If rsaux3.State = 1 Then
                  rsaux3.Close
               End If
               var_iva = IIf(IsNull(rs!floa_car_porcentaje_iva), 0, rs!floa_car_porcentaje_iva)
               rsaux3.Open "select * from TB_DETALLE_DESCUENTOS_FINANCIEROS where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'DF' and vcha_ser_serie_id  = '" + var_serie + "' and vcha_car_clase_id = 'DF' and inte_car_numero =  " + Str(var_factura_nueva), cnn, adOpenDynamic, adLockOptimistic
               var_contador_lineas = 0
               var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
               While Not rsaux3.EOF
                     var_linea = "DF" + Str(rs!INTE_CAR_NUMERO) + " " + rs!vcha_Car_nombre + " Factura " + Str(rsaux3!inte_ddf_factura) + " " + CStr(rsaux3!floa_ddf_descuento_otorgado) + "%"
                     If Round(rsaux3!floa_ddf_descuento_otorgado, 2) <> Round(rsaux3!floa_ddf_descuento_aplicado, 2) Then
                        var_linea = var_linea + " (" + Format(rsaux3!floa_ddf_descuento_aplicado, "###,###,##0.0000") + "%)"
                     End If
                     If Len(Trim(var_linea)) < 120 Then
                        For var_j = 1 + Len(Trim(var_linea)) To 120
                            var_linea = var_linea + " "
                        Next var_j
                     End If
                     If Len(Trim(var_rfc)) = 0 Then
                        var_importe_str = Format((IIf(IsNull(rsaux3!FLOA_DDF_IMPORTE), 0, rsaux3!FLOA_DDF_IMPORTE)), "###,###,##0.00")
                        If Len(Trim(var_importe_str)) < 14 Then
                           For var_j = 1 + Len(Trim(var_importe_str)) To 14
                               var_importe_str = " " + var_importe_str
                           Next var_j
                        End If
                     Else
                        var_importe_str = Format((IIf(IsNull(rsaux3!FLOA_DDF_IMPORTE), 0, rsaux3!FLOA_DDF_IMPORTE) / (1 + (var_iva / 100))), "###,###,##0.00")
                        If Len(Trim(var_importe_str)) < 14 Then
                           For var_j = 1 + Len(Trim(var_importe_str)) To 14
                               var_importe_str = " " + var_importe_str
                           Next var_j
                        End If
                     End If
                     var_linea = var_linea + var_importe_str
                     Print #1, var_linea
                     rsaux3.MoveNext
                     var_contador_lineas = var_contador_lineas + 1
               Wend
               If var_contador_lineas < var_numero_renglones Then
                  For var_l = var_contador_lineas To var_numero_renglones
                      Print #1, ""
                  Next var_l
               End If
               rsaux3.Close
               var_cantidad_letra = rs!vcha_car_importe_letra
               var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
               If Len(Trim(var_linea)) < 105 Then
                  For var_j = 1 + Len(Trim(var_linea)) To 105
                      var_linea = var_linea + " "
                  Next var_j
               End If
               If Len(Trim(var_rfc)) = 0 Then
                  var_subimporte_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                  If Len(Trim(var_subimporte_str)) < 14 Then
                     For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                         var_subimporte_str = " " + var_subimporte_str
                     Next var_j
                  End If
                  var_iva = "-        "
                  For var_j = 1 + Len(Trim(var_iva_str)) To 14
                      var_iva_str = " " + var_iva_str
                  Next var_j
               Else
                  var_subimporte_str = Format(((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                  If Len(Trim(var_subimporte_str)) < 14 Then
                     For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                         var_subimporte_str = " " + var_subimporte_str
                     Next var_j
                  End If
                  var_iva_str = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                  If Len(Trim(var_iva_str)) < 14 Then
                     For var_j = 1 + Len(Trim(var_iva_str)) To 14
                         var_iva_str = " " + var_iva_str
                     Next var_j
                  End If
               End If
               var_linea = var_linea + "           " + var_subimporte_str
               Print #1, Spc(4); var_linea
               Print #1, Spc(120); var_iva_str
               var_importe_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
               If Len(Trim(var_importe_str)) < 14 Then
                  For var_j = 1 + Len(Trim(var_importe_str)) To 14
                      var_importe_str = " " + var_importe_str
                  Next var_j
               End If
               Print #1, Spc(120); var_importe_str
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, Spc(85); "SISTEMAS"
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Close #1
                      
               Open (App.Path & "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".bat") For Output As #2
               var_Archivo = App.Path & "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".bat"
               Print #2, "copy " + App.Path + "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".txt lpt1"
               Close #2
               x = Shell(var_Archivo, vbHide)
            End If
            rs.Close
         End If
         If var_tipo_nota_credito = "" Then
            rs.Open "select * from VW_DOCUMENTOS_IMPRESION where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'BF' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + CStr(var_factura_nueva), cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
      ''''''''IMPRESION DE LA NOTA DE CARGO
               Open (App.Path & "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".txt") For Output As #1
               'Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
               Print #1, Chr(27) + Chr(64)
               Print #1, ""
               Print #1, Spc(92); Str(rs!INTE_CAR_NUMERO)
               Print #1, ""
               Print #1, Spc(92); "FECHA: "; Format(rs!dtim_Car_fecha, "Short Date")
               var_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
               For var_j = 1 + Len(Trim(var_cliente)) To 83
                   var_cliente = var_cliente + " "
               Next var_j
               var_cliente = var_cliente + "AGUASCALIENTES, AGS."
               Print #1, ""
               Print #1, Spc(12); var_cliente
               var_domicilio = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " C.P. " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
               For var_j = 1 + Len(Trim(var_domicilio)) To 83
                   var_domicilio = var_domicilio + " "
               Next var_j
               var_agente = ""
               var_agente = IIf(IsNull(rs!vcha_AGE_aGENTE_ID), "", rs!vcha_AGE_aGENTE_ID)
               For var_j = 1 + Len(Trim(var_agente)) To 8
                   var_agente = var_agente + " "
               Next var_j
               var_agente = var_agente + IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
               var_domicilio = var_domicilio
               Print #1, Spc(12); var_domicilio
               var_ciudad = ""
               var_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
               For var_j = 1 + Len(Trim(var_ciudad)) To 37
                   var_ciudad = var_ciudad + " "
               Next var_j
               var_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
               For var_j = 1 + Len(Trim(var_estado)) To 46
                   var_estado = var_estado + " "
               Next var_j
               var_ciudad = var_ciudad + var_estado
                            
               For var_j = 1 + Len(Trim(var_ciudad)) To 14
                   var_ciudad = var_ciudad + " "
               Next var_j
                        
               var_ciudad = var_ciudad + var_agente
                       
               Print #1, Spc(12); var_ciudad
               var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
               var_rfc = "      " + var_rfc
               For var_j = 1 + Len(Trim(var_rfc)) To 89
                   var_rfc = var_rfc + " "
               Next var_j
               var_rfc = var_rfc
               Print #1, Spc(6); var_rfc
               Print #1, ""
               Print #1, ""
               Print #1, ""
               var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
               If rsaux.State = 1 Then
                  rsaux.Close
               End If
                     
               rsaux.Open "select * from tb_detalle_bonificacion_financiera where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'BF' and vcha_ser_serie_id = '" + var_serie + "' and inte_car_numero = " + CStr(var_factura_nueva), cnn, adOpenDynamic, adLockOptimistic
               var_contador_renglones_nota = 0
               While Not rsaux.EOF
                     var_linea = "BF" + Str(rs!INTE_CAR_NUMERO) + " " + rs!vcha_Car_nombre + " FACTURA " + Str(rsaux!inte_dbf_factura) + " " + Format(rsaux!floa_dbf_porcentaje, "###,###,##0.0000") + "%"
                     If Len(Trim(var_linea)) < 120 Then
                        For var_j = 1 + Len(Trim(var_linea)) To 120
                            var_linea = var_linea + " "
                        Next var_j
                     End If
                     If Len(Trim(var_rfc)) = 0 Then
                        var_importe_str = Format((IIf(IsNull(rsaux!FLOA_DBF_IMPORTE), 0, rsaux!FLOA_DBF_IMPORTE)), "###,###,##0.00")
                     Else
                        var_importe_str = Format(((IIf(IsNull(rsaux!FLOA_DBF_IMPORTE), 0, rsaux!FLOA_DBF_IMPORTE)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)) / (1 + ((IIf(IsNull(rsaux!floa_dbf_iva), 1, rsaux!floa_dbf_iva) / 100)))), "###,###,##0.00")
                     End If
                     If Len(Trim(var_importe_str)) < 14 Then
                        For var_j = 1 + Len(Trim(var_importe_str)) To 14
                            var_importe_str = " " + var_importe_str
                        Next var_j
                     End If
                     var_linea = var_linea + var_importe_str
                     Print #1, var_linea
                     rsaux.MoveNext
                     var_contador_renglones_nota = var_contador_renglones_nota + 1
               Wend
               rsaux.Close
               If var_contador_renglones_nota < var_numero_renglones Then
                  For var_l = var_contador_renglones_nota To var_numero_renglones
                      Print #1, ""
                  Next var_l
               End If
               var_cantidad_letra = rs!vcha_car_importe_letra
               var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
               If Len(Trim(var_linea)) < 105 Then
                  For var_j = 1 + Len(Trim(var_linea)) To 105
                      var_linea = var_linea + " "
                  Next var_j
               End If
                     
                     
               If Len(Trim(var_rfc)) = 0 Then
                  var_subimporte_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                  If Len(Trim(var_subimporte_str)) < 14 Then
                     For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                         var_subimporte_str = " " + var_subimporte_str
                     Next var_j
                  End If
                  var_iva = "      -        "
                  For var_j = 1 + Len(Trim(var_iva_str)) To 14
                      var_iva_str = " " + var_iva_str
                  Next var_j
               Else
                  var_subimporte_str = Format(((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                  If Len(Trim(var_subimporte_str)) < 14 Then
                     For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                         var_subimporte_str = " " + var_subimporte_str
                     Next var_j
                  End If
                  var_iva_str = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                  If Len(Trim(var_iva_str)) < 14 Then
                     For var_j = 1 + Len(Trim(var_iva_str)) To 14
                         var_iva_str = " " + var_iva_str
                     Next var_j
                  End If
               End If
               var_linea = var_linea + "           " + var_subimporte_str
               Print #1, Spc(4); var_linea
               Print #1, Spc(120); var_iva_str
               var_importe_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
               If Len(Trim(var_importe_str)) < 14 Then
                  For var_j = 1 + Len(Trim(var_importe_str)) To 14
                      var_importe_str = " " + var_importe_str
                  Next var_j
               End If
               Print #1, Spc(120); var_importe_str
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, Spc(85); "SISTEMAS"
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Close #1
               
               Open (App.Path & "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".bat") For Output As #2
               var_Archivo = App.Path & "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".bat"
               Print #2, "copy " + App.Path + "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".txt lpt1"
               Close #2
               x = Shell(var_Archivo, vbHide)
''''''''''''
            End If
            rs.Close
         End If
         If var_tipo_nota_credito = "BO" Or var_tipo_nota_credito = "BF" Then
            rs.Open "select * from VW_DOCUMENTOS_IMPRESION where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = '" + var_tipo_nota_credito + "' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(var_factura_nueva), cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               Open (App.Path & "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".txt") For Output As #1
               Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
               Print #1, Spc(92); Str(rs!INTE_CAR_NUMERO)
               Print #1, ""
               Print #1, Spc(92); "FECHA: "; Format(rs!dtim_Car_fecha, "Short Date")
               var_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
               For var_j = 1 + Len(Trim(var_cliente)) To 83
                   var_cliente = var_cliente + " "
               Next var_j
               var_cliente = var_cliente + "AGUASCALIENTES, AGS."
               Print #1, ""
               Print #1, Spc(12); var_cliente
               var_domicilio = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " C.P. " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
               For var_j = 1 + Len(Trim(var_domicilio)) To 83
                   var_domicilio = var_domicilio + " "
               Next var_j
               var_agente = ""
               var_agente = IIf(IsNull(rs!vcha_AGE_aGENTE_ID), "", rs!vcha_AGE_aGENTE_ID)
               For var_j = 1 + Len(Trim(var_agente)) To 8
                   var_agente = var_agente + " "
               Next var_j
               var_agente = var_agente + IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
               var_domicilio = var_domicilio
               Print #1, Spc(12); var_domicilio
               var_ciudad = ""
               var_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
               For var_j = 1 + Len(Trim(var_ciudad)) To 37
                   var_ciudad = var_ciudad + " "
               Next var_j
               var_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
               For var_j = 1 + Len(Trim(var_estado)) To 46
                   var_estado = var_estado + " "
               Next var_j
               var_ciudad = var_ciudad + var_estado
                             
               For var_j = 1 + Len(Trim(var_ciudad)) To 14
                   var_ciudad = var_ciudad + " "
               Next var_j
                 
               var_ciudad = var_ciudad + var_agente
                            
               Print #1, Spc(12); var_ciudad
               var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
               var_rfc = "      " + var_rfc
               For var_j = 1 + Len(Trim(var_rfc)) To 89
                   var_rfc = var_rfc + " "
               Next var_j
               var_rfc = var_rfc
               Print #1, Spc(6); var_rfc
               Print #1, ""
               Print #1, ""
               Print #1, ""
               var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
               rsaux.Open "select * from tb_detalle_bonificaciones where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = '" + var_tipo_nota_credito + "' and vcha_ser_serie_id = '" + var_serie + "' and inte_car_numero = " + CStr(var_factura_nueva), cnn, adOpenDynamic, adLockOptimistic
               var_contador_renglones_nota = 0
               While Not rsaux.EOF
                     var_linea = "BO" + Str(rs!INTE_CAR_NUMERO) + " " + rs!vcha_Car_nombre + " FACTURA " + CStr(rsaux!inte_car_factura)
                     If Len(Trim(var_linea)) < 108 Then
                        For var_j = 1 + Len(Trim(var_linea)) To 120
                            var_linea = var_linea + " "
                        Next var_j
                     End If
                     If Len(Trim(var_rfc)) = 0 Then
                        var_importe_str = Format(((IIf(IsNull(rsaux!FLOA_dbo_IMPORTE), 0, rsaux!FLOA_dbo_IMPORTE)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio))), "###,###,##0.00")
                     Else
                        var_importe_str = Format(((IIf(IsNull(rsaux!FLOA_dbo_IMPORTE), 0, rsaux!FLOA_dbo_IMPORTE)) / (1 + (rsaux!floa_dbo_iva / 100)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio))), "###,###,##0.00")
                     End If
                     If Len(Trim(var_importe_str)) < 14 Then
                        For var_j = 1 + Len(Trim(var_importe_str)) To 14
                            var_importe_str = " " + var_importe_str
                        Next var_j
                     End If
                     var_linea = var_linea + var_importe_str
                     Print #1, var_linea
                     rsaux.MoveNext
                     var_contador_renglones_nota = var_contador_renglones_nota + 1
               Wend
               rsaux.Close
               If var_contador_renglones_nota < var_numero_renglones Then
                  For var_l = var_contador_renglones_nota To var_numero_renglones
                      Print #1, ""
                  Next var_l
               End If
               var_cantidad_letra = rs!vcha_car_importe_letra
               var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
               If Len(Trim(var_linea)) < 105 Then
                  For var_j = 1 + Len(Trim(var_linea)) To 105
                      var_linea = var_linea + " "
                  Next var_j
               End If
                    
                       
               If Len(Trim(var_rfc)) = 0 Then
                  var_subimporte_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                  If Len(Trim(var_subimporte_str)) < 14 Then
                     For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                         var_subimporte_str = " " + var_subimporte_str
                     Next var_j
                  End If
                  var_iva = "      -        "
                  For var_j = 1 + Len(Trim(var_iva_str)) To 14
                      var_iva_str = " " + var_iva_str
                  Next var_j
               Else
                  var_subimporte_str = Format(((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                  If Len(Trim(var_subimporte_str)) < 14 Then
                     For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                         var_subimporte_str = " " + var_subimporte_str
                     Next var_j
                  End If
                  var_iva_str = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                  If Len(Trim(var_iva_str)) < 14 Then
                     For var_j = 1 + Len(Trim(var_iva_str)) To 14
                         var_iva_str = " " + var_iva_str
                     Next var_j
                  End If
               End If
               var_linea = var_linea + "           " + var_subimporte_str
               Print #1, Spc(4); var_linea
               Print #1, Spc(120); var_iva_str
               var_importe_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
               If Len(Trim(var_importe_str)) < 14 Then
                  For var_j = 1 + Len(Trim(var_importe_str)) To 14
                      var_importe_str = " " + var_importe_str
                  Next var_j
               End If
               
               Print #1, Spc(120); var_importe_str
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, Spc(85); "SISTEMAS"
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Close #1
                       
               Open (App.Path & "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".bat") For Output As #2
               var_Archivo = App.Path & "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".bat"
               Print #2, "copy " + App.Path + "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".txt lpt1"
               Close #2
               x = Shell(var_Archivo, vbHide)
            End If
            rs.Close
         End If
   
         If var_tipo_nota_credito = "DV" Then
            rs.Open "select * from VW_DOCUMENTOS_IMPRESION where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'DV' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(var_factura_nueva), cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
''''' IMPRESION DE LA NOTA DE CARGO
               var_clave_movimiento = rs!vcha_mov_movimiento_id
               var_numero_movimiento = rs!INTE_EMO_NUMERO
               Open (App.Path & "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".txt") For Output As #1
               Print #1, Chr(27) + Chr(64)
               Print #1, ""
               Print #1, Spc(92); Str(rs!INTE_CAR_NUMERO)
               Print #1, ""
               Print #1, Spc(92); "FECHA: "; Format(rs!dtim_Car_fecha, "Short Date")
               var_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
               For var_j = 1 + Len(Trim(var_cliente)) To 83
                   var_cliente = var_cliente + " "
               Next var_j
               var_cliente = var_cliente + "AGUASCALIENTES, AGS."
               Print #1, ""
               Print #1, Spc(12); var_cliente
               var_domicilio = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " C.P. " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
               For var_j = 1 + Len(Trim(var_domicilio)) To 83
                   var_domicilio = var_domicilio + " "
               Next var_j
               var_agente = ""
               var_agente = IIf(IsNull(rs!vcha_AGE_aGENTE_ID), "", rs!vcha_AGE_aGENTE_ID)
               For var_j = 1 + Len(Trim(var_agente)) To 8
                   var_agente = var_agente + " "
               Next var_j
               var_agente = var_agente + IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
               var_domicilio = var_domicilio
               Print #1, Spc(12); var_domicilio
               var_ciudad = ""
               var_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
               For var_j = 1 + Len(Trim(var_ciudad)) To 37
                   var_ciudad = var_ciudad + " "
               Next var_j
               var_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
               For var_j = 1 + Len(Trim(var_estado)) To 46
                   var_estado = var_estado + " "
               Next var_j
               var_ciudad = var_ciudad + var_estado
                            
               For var_j = 1 + Len(Trim(var_ciudad)) To 14
                   var_ciudad = var_ciudad + " "
               Next var_j
                            
               var_ciudad = var_ciudad + var_agente
                                     
               Print #1, Spc(12); var_ciudad
               var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
               var_rfc = "      " + var_rfc
               For var_j = 1 + Len(Trim(var_rfc)) To 89
                   var_rfc = var_rfc + " "
               Next var_j
               rsaux3.Open "select * from vw_devolucion_nota_credito where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' and inte_emo_numero = " + CStr(var_numero_movimiento), cnn, adOpenDynamic, adLockOptimistic
               var_referencia = rsaux3!vcha_cde_referencia
               var_rfc = var_rfc + "ENTRADA: " + var_clave_movimiento + " " + CStr(var_numero_movimiento) + " RELACION: " + var_referencia
               var_rfc = var_rfc
               Print #1, Spc(6); var_rfc
               Print #1, ""
               Print #1, ""
               Print #1, ""
               var_contador_renglones = 0
               var_cantidad_total = 0
               While Not rsaux3.EOF
                     var_factura = CStr(rsaux3!inte_fac_factura) + IIf(IsNull(rsaux3!VCHA_SER_SERIE_ID), "", rsaux3!VCHA_SER_SERIE_ID)
                     If Len(Trim(var_factura)) < 14 Then
                        For var_j = Len(Trim(var_factura)) To 14
                            var_factura = var_factura + " "
                        Next var_j
                     End If
                     var_linea = var_factura + rsaux3!vcha_Art_articulo_id + " " + rsaux3!vcha_art_nombre_Español
                    
                     If Len(Trim(var_linea)) < 88 Then
                        For var_j = Len(Trim(var_linea)) To 88
                            var_linea = var_linea + " "
                        Next var_j
                     End If
                     var_cantidad_str = Format(IIf(IsNull(rsaux3!Cantidad), 0, rsaux3!Cantidad), "###,###,##0.00")
                     var_tipo_Cambio = IIf(IsNull(rsaux3!floa_dev_tipo_cambio), 1, rsaux3!floa_dev_tipo_cambio)
                     var_cantidad = Format(IIf(IsNull(rsaux3!Cantidad), 0, rsaux3!Cantidad), "###,###,##0.00")
                     var_cantidad_total = var_cantidad_total + IIf(IsNull(rsaux3!Cantidad), 0, rsaux3!Cantidad)
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
                     If Len(Trim(var_rfc)) = 0 Then
                        var_precio = var_precio * (1 + var_iva / 100)
                     End If
                     var_precio_str = Format(var_precio, "###,###,##0.00")
                     var_importe_str = Format(var_precio * var_cantidad, "###,###,##0.00")
                     If Len(Trim(var_cantidad_str)) < 14 Then
                        For var_j = Len(Trim(var_cantidad_str)) To 14
                            var_cantidad_str = " " + var_cantidad_str
                        Next var_j
                     End If
                     var_linea = var_linea + var_cantidad_str
                     If Len(Trim(var_linea)) < 104 Then
                        For var_j = 1 + Len(Trim(var_linea)) To 104
                            var_linea = var_linea + " "
                        Next var_j
                     End If
                     If Len(Trim(var_precio_str)) < 14 Then
                        For var_j = Len(Trim(var_precio_str)) To 14
                            var_precio_str = " " + var_precio_str
                        Next var_j
                     End If
                     If Len(Trim(var_importe_str)) < 14 Then
                        For var_j = Len(Trim(var_importe_str)) To 14
                            var_importe_str = " " + var_importe_str
                        Next var_j
                     End If
                     var_linea = var_linea + var_precio_str + var_importe_str
                     Print #1, var_linea
                     rsaux3.MoveNext
                     var_contador_renglones = var_contador_renglones + 1
               Wend
               var_cantidad_total_str = Format(var_cantidad_total, "###,###,##0.00")
               If Len(Trim(var_cantidad_total_str)) < 14 Then
                  For var_j = Len(Trim(var_cantidad_total_str)) To 14
                      var_cantidad_total_str = " " + var_cantidad_total_str
                  Next var_j
               End If
               If var_contador_renglones < 30 Then
                  For var_j = var_contador_renglones To 30
                      Print #1, ""
                  Next var_j
               End If
               rsaux3.Close
               Print #1, ""
               Print #1, ""
               Print #1, ""
               var_contador_renglones = 0
               rsaux3.Open "select * from TB_DETALLE_DEVOLUCION_IMPORTES_ASIGNADOS where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'DV' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(var_factura_nueva), cnn, adOpenDynamic, adLockOptimistic
               var_linea = ""
               While Not rsaux3.EOF
                     If Len(Trim(var_linea + ", " + CStr(rsaux3!inte_dia_numero) + rsaux3!vcha_dia_serie + " " + Format(rsaux3!floa_dia_importe, "###,###,##0.00"))) < 108 Then
                        If Len(Trim(var_linea)) = 0 Then
                           var_linea = var_linea + CStr(rsaux3!inte_dia_numero) + rsaux3!vcha_dia_serie + " " + Format(rsaux3!floa_dia_importe, "###,###,##0.00")
                        Else
                           var_linea = var_linea + ", " + CStr(rsaux3!inte_dia_numero) + rsaux3!vcha_dia_serie + " " + Format(rsaux3!floa_dia_importe, "###,###,##0.00")
                        End If
                     Else
                        Print #1, var_linea
                        var_contador_renglones = var_contador_renglones + 1
                        var_linea = ""
                        var_linea = CStr(rsaux3!inte_dia_numero) + rsaux3!vcha_dia_serie + " " + Format(rsaux3!floa_dia_importe, "###,###,##0.00")
                     End If
                     rsaux3.MoveNext
                     If rsaux3.EOF And Len(var_linea) < 118 Then
                        Print #1, var_linea
                        var_contador_renglones = var_contador_renglones + 1
                     End If
               Wend
               If var_contador_renglones < 4 Then
                  For var_j = var_contador_renglones To 4
                      Print #1, ""
                  Next var_j
               End If
               rsaux3.Close
               var_cantidad_letra = rs!vcha_car_importe_letra
                
               var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
                
               If Len(Trim(var_linea)) < 62 Then
                  For var_j = 1 + Len(Trim(var_linea)) To 62
                      var_linea = var_linea + " "
                  Next var_j
               End If
               var_linea = var_linea + var_cantidad_total_str
               var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                  
               If Len(Trim(var_rfc)) = 0 Then
                  var_subimporte_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                  If Len(Trim(var_subimporte_str)) < 14 Then
                     For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                         var_subimporte_str = " " + var_subimporte_str
                     Next var_j
                  End If
                  var_iva = "      -        "
                  For var_j = 1 + Len(Trim(var_iva_str)) To 14
                      var_iva_str = " " + var_iva_str
                  Next var_j
               Else
                  var_subimporte_str = Format(((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) - (IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva))) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                  If Len(Trim(var_subimporte_str)) < 14 Then
                     For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                         var_subimporte_str = " " + var_subimporte_str
                     Next var_j
                  End If
                  var_iva_str = Format((IIf(IsNull(rs!floa_car_importe_iva), 0, rs!floa_car_importe_iva)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                  If Len(Trim(var_iva_str)) < 14 Then
                     For var_j = 1 + Len(Trim(var_iva_str)) To 14
                         var_iva_str = " " + var_iva_str
                     Next var_j
                  End If
               End If
               If Len(Trim(var_linea)) < 115 Then
                  For var_j = Len(Trim(var_linea)) To 115
                      var_linea = var_linea + " "
                  Next var_j
               End If
               var_linea = var_linea + var_subimporte_str
               Print #1, Spc(4); var_linea
               Print #1, Spc(120); var_iva_str
               var_importe_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
               If Len(Trim(var_importe_str)) < 14 Then
                  For var_j = 1 + Len(Trim(var_importe_str)) To 14
                      var_importe_str = " " + var_importe_str
                  Next var_j
               End If
               Print #1, Spc(120); var_importe_str
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, Spc(85); "SISTEMAS"
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Print #1, ""
               Close #1
                            
               Open (App.Path & "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".bat") For Output As #2
               var_Archivo = App.Path & "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".bat"
               Print #2, "copy " + App.Path + "\nota_credito" + Trim(Str(rs!INTE_CAR_NUMERO)) + ".txt lpt1"
               Close #2
               x = Shell(var_Archivo, vbHide)
            End If
         End If
      Else
         rs.Close
         MsgBox "La nota de crédito no existe", vbOKOnly, "ATENCION"
      End If
   End If
   Else
      MsgBox "Solo se pueden reimprimir notas de credito", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 2500
   Left = 2200
   lbl_estatus = ""
   If var_empresa = "15" Or var_empresa = "06" Then
      Me.cmd_imprimir.Visible = False
      Me.cmd_reimprimir.Visible = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_despliega_menu = True Then
      var_swpassword = False
      var_modifica_registro = False
   End If
   Call activa_forma(var_activa_forma_cancela_documentos_existentes)
End Sub

Private Sub txt_documento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      cmb_documentos.SetFocus
   End If
End Sub

Private Sub txt_documento_LostFocus()
   lbl_estatus = ""
   txt_numero = ""
   var_tipo_facturacion = ""
   If Trim(txt_documento) <> "" Then
      If txt_documento = "FA" Then
         cmb_documentos = "FACTURA"
         var_tipo_facturacion = ""
      Else
         If txt_documento = "NC" Then
            cmb_documentos = "NOTA DE CREDITO"
         Else
            If txt_documento = "NG" Then
               cmb_documentos = "NOTA DE CARGO"
            Else
               If txt_documento = "CS" Then
                  cmb_documentos = "CANCELACION DE SALDOS"
               Else
                  If txt_documento = "CH" Then
                     cmb_documentos = "CHEQUES"
                  Else
                     If txt_documento = "PA" Then
                        cmb_documentos = "PAGO"
                     Else
                        MsgBox "Clave de documento incorrecta", vbOKOnly, "ATENCION"
                        txt_documento = ""
                        cmb_documentos = ""
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_cancelar.SetFocus
   End If
End Sub

Private Sub txt_numero_LostFocus()
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If Trim(txt_documento) <> "" Then
      If Not IsNumeric(txt_numero) Then
         MsgBox "Número de documento incorrecto", vbOKOnly, "ATENCION"
         txt_numero = ""
      Else
         var_estatus = ""
         If rs.State = 1 Then
            rs.Close
         End If
         var_serie = Me.txt_serie
         rs.Open "select isnull(char_car_estatus,'') as char_car_estatus, isnull(char_car_tipo_facturacion,'') as char_car_tipo_facturacion, isnull(vcha_car_documento,'') as vcha_car_documento from tb_encabezado_cartera where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_tipo_documento = '" + txt_documento + "' and vcha_ser_serie_id = '" + var_serie + "' and inte_Car_numero = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            If rs!CHAR_CAR_ESTATUS = "" Or rs!CHAR_CAR_ESTATUS = " " Or rs!CHAR_CAR_ESTATUS = "I" Then
               lbl_estatus = "ESTATUS: IMPRESA"
               var_estatus = "I"
               If txt_documento = "FA" Then
                  var_tipo_facturacion = rs!char_Car_tipo_facturacion
               End If
               If txt_documento = "NC" Then
                  var_tipo_nota_credito = rs!vcha_Car_documento
               End If
            End If
            If rs!CHAR_CAR_ESTATUS = "C" Then
               lbl_estatus = "ESTATUS: CANCELADA"
               var_estatus = "C"
               If txt_documento = "FA" Then
                  var_tipo_facturacion = rs!char_Car_tipo_facturacion
               End If
               If txt_documento = "NC" Then
                  var_tipo_nota_credito = rs!vcha_Car_documento
               End If
            End If
         Else
            lbl_estatus = "ESTATUS: NO IMPRESA"
            var_estatus = "N"
         End If
         rs.Close
      End If
   Else
      MsgBox "Se debe de seleccionar un tipo de documento", vbOKOnly, "ATENCION"
      txt_numero = ""
   End If
End Sub


Private Sub txt_serie_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub
