VERSION 5.00
Begin VB.Form frmimpresion_notas_cargo_1_por_ciento_iva 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresión de notas de cargo "
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Impresión de relación de facturas"
      Height          =   525
      Left            =   180
      TabIndex        =   1
      Top             =   840
      Width           =   4320
   End
   Begin VB.CommandButton cmd_vth 
      Caption         =   "Impresión de notas de cargo"
      Height          =   525
      Left            =   180
      TabIndex        =   0
      Top             =   225
      Width           =   4320
   End
End
Attribute VB_Name = "frmimpresion_notas_cargo_1_por_ciento_iva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub Command1_Click()
            Set reporte = appl.OpenReport(App.Path + "\rep_saldos_con_importe.rpt")
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de relación de facturas por complemento del 1% de IVA"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            
            
            var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_saldos_con_importe.rpt")
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\relacion_facturar_1_por_ciento_iva_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If

End Sub

Private Sub cmd_vth_Click()
   Set TB_ENCABEZADO_CARTERA_I = New TB_ENCABEZADO_CARTERA_I
   Set TB_ESTADO_CUENTA_INSERTA = New TB_ESTADO_CUENTA_INSERTA
   Dim si As Integer
   Dim var_importe_iva As Double
   Dim var_importe_total As Double
   Dim var_importe_neto As Double
   Dim var_subimporte As Double
   Dim var_tipo_Cambio As Double
   Dim var_numero_folio As Double
   Dim var_moneda_local As Integer
   Dim var_posible_tipo_cambio As Boolean
   Dim var_j As Integer
   Dim var_subimporte_str As String
   Dim var_importe_str As String
   Dim var_iva_str As String
   Dim var_ciudad As String
   Dim var_rfc As String
   Dim var_linea As String
   Dim var_estado As String
   Dim var_almacen As String
   Dim var_grupo_actual As String
   Dim var_grupo_real As String
   Dim var_cliente As String
   Dim var_titular As String
   Dim var_establecimiento As String
   Dim var_clave_moneda As String
   Dim var_agente As String
   Dim var_imprimir As Boolean
   Dim var_contador As Integer
   Dim var_contador_notas As Integer
   Dim var_iva As Double
   Dim var_importe As Double
   Dim var_plazo As Integer
   Dim i, n As Integer
   Dim var_serie As String
   Dim var_tipo_lista As Integer
   rs.Open "select * from tb_empresas where vcha_emp_Empresa_id = '" + var_empresa + "' ", cnn, adOpenDynamic, adLockOptimistic
   var_nombre_empresa = IIf(IsNull(rs!VCHA_EMP_NOMBRE), "", rs!VCHA_EMP_NOMBRE)
   rs.Close
   var_si = MsgBox("¿Se imprimiran las notas de cargo de la empresa " + var_empresa + " " + var_nombre_empresa + "?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      var_si = MsgBox("Confirmar la impresión de las notas de cargo de la empresa ", vbYesNo, "ATENCION")
      If var_si = 6 Then
         'var_cadena = "SELECT TOP 10 dbo.TB_SALDOS.VCHA_EMP_EMPRESA_ID, dbo.TB_SALDOS.VCHA_SER_SERIE_ID, dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO, dbo.TB_SALDOS.VCHA_CLI_CLAVE_ID, dbo.TB_SALDOS.INTE_CAR_NUMERO, dbo.TB_SALDOS.FLOA_SAL_IMPORTE, dbo.TB_SALDOS.VCHA_MON_MONEDA_ID, dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA, dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_IMPORTE_NETO / dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_TIPO_CAMBIO AS importe, (dbo.TB_SALDOS.FLOA_SAL_IMPORTE * 100) / (dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_IMPORTE_NETO / dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_TIPO_CAMBIO) AS porcentaje, (dbo.TB_SALDOS.FLOA_SAL_IMPORTE/1.15) * .01 AS IMPORTE_1_POR_CIENTO FROM dbo.TB_SALDOS INNER JOIN dbo.TB_ENCABEZADO_CARTERA ON dbo.TB_SALDOS.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_SALDOS.VCHA_SER_SERIE_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID AND dbo.TB_SALDOS.VCHA_CAR_DOCUMENTO = dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO AND"
         'var_cadena = var_cadena + " dbo.TB_SALDOS.INTE_CAR_NUMERO = dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO INNER JOIN dbo.TB_CLIENTES ON dbo.TB_SALDOS.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID WHERE ((dbo.TB_SALDOS.FLOA_SAL_IMPORTE * 100) / (dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_IMPORTE_NETO / dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_TIPO_CAMBIO) > 4.001) AND (dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_IMPORTE_NETO > 0) AND (dbo.TB_SALDOS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA < CONVERT(DATETIME, '2010-01-01', 102)) AND (dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS <> 'C') OR ((dbo.TB_SALDOS.FLOA_SAL_IMPORTE * 100) / (dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_IMPORTE_NETO / dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_TIPO_CAMBIO) > 4.001) AND (dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_IMPORTE_NETO > 0) AND (dbo.TB_SALDOS.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA < CONVERT(DATETIME, '2010-01-01', 102)) AND "
         'var_cadena = var_cadena + " (dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS IS NULL) ORDER BY dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA, (dbo.TB_SALDOS.FLOA_SAL_IMPORTE * 100) / (dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_IMPORTE_NETO / dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_TIPO_CAMBIO)"
         var_cadena = "SELECT CAST(INTE_cAR_NUMERO as integer) as inte_Car_numero, VCHA_CLI_CLAVE_ID, SUM(importe) AS importe, CAST(MAX(fecha - GETDATE()) AS INTEGER) AS PLAZO, VCHA_EMP_EMPRESA_ID From dbo.saldos_con_importe Where (marca Is Not Null) and vcha_Emp_empresa_id = '" + var_empresa + "' AND FECHA < {d '2010-01-01'}GROUP BY INTE_cAR_NUMERO, VCHA_CLI_CLAVE_ID, VCHA_EMP_EMPRESA_ID"
         rsaux11.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         txt_clase_cartera = "IV"
         txt_serie = ""
         txt_numero = ""
         rs.Open "select * from tb_series where vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
         txt_serie = rs!VCHA_SER_SERIE_ID
         txt_numero = rs!inte_ser_nota_Cargo
         'txt_importe = rsaux11!IMPORTE_1_POR_CIENTO
         txt_importe = rsaux11!Importe
         rs.Close
         var_archivo_bat = txt_numero
         If txt_serie <> "" Then
            Open (App.Path & "\nota_cargo" + CStr(txt_numero) + ".bat") For Output As #2
            var_archivo_bat = App.Path & "\nota_cargo" + CStr(txt_numero) + ".bat"
            While Not rsaux11.EOF
                  If rsaux11!Importe > 0 Then
                  txt_clave_cliente = rsaux11!vcha_cli_clave_id
                  If Trim(txt_clave_cliente) <> "" Then
                     rs.Open "select * from vw_clientes where vcha_cli_clave_id = '" + txt_clave_cliente + "' and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_gac_grupo_Actual_id is not null", cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        txt_nombre_cliente = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                        var_grupo_actual = IIf(IsNull(rs!VCHA_GAC_GRUPO_aCTUAL_ID), "", rs!VCHA_GAC_GRUPO_aCTUAL_ID)
                        var_grupo_real = IIf(IsNull(rs!vcha_gre_grupo_real_id), "", rs!vcha_gre_grupo_real_id)
                        var_cliente = txt_clave_cliente
                        var_titular = IIf(IsNull(rs!vcha_tit_titular_id), "", rs!vcha_tit_titular_id)
                        var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
                        var_agente = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
                        var_plazo = IIf(IsNull(rs!inte_pla_dias), 0, rs!inte_pla_dias)
                        txt_plazo = 0
                        var_iva = IIf(IsNull(rs!FLOA_TPE_IVA), 0, rs!FLOA_TPE_IVA)
                        var_clave_moneda = IIf(IsNull(rs!vcha_mon_moneda_id), "", rs!vcha_mon_moneda_id)
                        lbl_moneda = IIf(IsNull(rs!vcha_mon_nombre_plural), "", rs!vcha_mon_nombre_plural)
                     Else
                        txt_clave_cliente = ""
                        txt_nombre_cliente = ""
                        var_grupo_actual = ""
                        var_grupo_real = ""
                        var_cliente = ""
                        var_titular = ""
                        var_clave_moneda = ""
                        var_agente = ""
                        var_plazo = 0
                        txt_plazo = var_plazo
                        var_iva = 0
                        var_clave_moneda = ""
                        lbl_moneda = ""
                     End If
                     rs.Close
                  End If
                  
                  
                  
                  var_moneda_local = 1
                  If Trim(txt_serie) <> "" Then
                     If Trim(txt_numero) <> "" Then
                        If IsNumeric(txt_numero) Then
                           txt_clave_cliente = rsaux11!vcha_cli_clave_id
                           If Trim(txt_clave_cliente) <> "" Then
                              txt_clase_cartera = "IV"
                              If Trim(txt_clase_cartera) <> "" Then
                                 'txt_importe = rsaux11!IMPORTE_1_POR_CIENTO
                                 txt_importe = rsaux11!Importe
                                 If IsNumeric(txt_importe) Then
                                    var_moneda_local = 1
                                    rs.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
                                    If Not rs.EOF Then
                                       var_moneda_local = IIf(IsNull(rs!inte_mon_moneda_local), 0, rs!inte_mon_moneda_local)
                                    End If
                                    rs.Close
                                    var_tipo_Cambio = 1
                                    If var_moneda_local = 0 Then
                                       rs.Open "select * from vw_tipocambio_fecha where vcha_mon_moneda_id = '" + var_clave_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
                                       If Not rs.EOF Then
                                          var_tipo_Cambio = IIf(IsNull(rs!mone_tca_importe), 1, rs!mone_tca_importe)
                                          var_posible_tipo_cambio = True
                                       Else
                                          var_posible_tipo_cambio = False
                                       End If
                                       rs.Close
                                    Else
                                       var_posible_tipo_cambio = True
                                    End If
                                    If var_posible_tipo_cambio = True Then
                                       si = 6
                                       If si = 6 Then
                                          si = 6
                                          If si = 6 Then
                                             var_serie = txt_serie
                                             var_numero_folio = CDbl(txt_numero)
                                             rs.Open "SELECT * FROM TB_ENCABEZADO_CARTERA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_CAR_TIPO_DOCUMENTO = 'NG' AND VCHA_SER_SERIE_ID = '" + var_serie + "' AND INTE_CAR_NUMERO = " + CStr(txt_numero), cnn, adOpenDynamic, adLockOptimistic
                                             If rs.EOF Then
                                                rs.Close
                                                si = 6
                                                If si = 6 Then
                                                   cnn.BeginTrans
                                                   var_importe_total = (txt_importe / (1 + (var_iva / 100))) * var_tipo_Cambio
                                                   var_importe_neto = txt_importe * var_tipo_Cambio
                                                   var_importe_iva = var_importe_neto - var_importe_total
                                                   var_subimporte = var_importe_total
                                                   var_insertar = False
                                                   var_insertar = TB_ENCABEZADO_CARTERA_I.Anadir(var_empresa, var_unidad_organizacional, "NG", "NC", CStr(txt_clase_cartera), CDbl(var_numero_folio), "+", "", "", 0, CStr(Date + 19), CStr(var_agente), CStr(var_grupo_actual), CStr(var_grupo_real), CStr(var_titular), CStr(txt_clave_cliente), "", CDbl(var_plazo), CDbl(var_iva), 0, 0, 0, 0, 0, CDbl(var_importe_total), CDbl(var_importe_iva), 0, 0, 0, 0, 0, CDbl(var_subimporte), CDbl(var_importe_neto), "", var_clave_usuario_global, fun_NombrePc, Date, 0, Date, Date, CStr(var_clave_moneda), CStr(var_tipo_Cambio), CStr(var_serie), "")
                                                   var_inserta = TB_ESTADO_CUENTA_INSERTA.Anadir(var_empresa, var_serie, "NC", var_numero_folio, "", "", 0, var_importe_neto, 0)
                                                   rsaux3.Open "update tb_principal set inte_pri_nota_cargo = inte_pri_nota_cargo + 1", cnn, adOpenDynamic, adLockOptimistic
                                                   rsaux3.Open "update tb_series set inte_ser_nota_cargo =  inte_ser_nota_cargo + 1 where vcha_emp_empresa_id= '" + var_emppresa + "' and vcha_ser_serie_id = '" + var_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                                                   cnn.CommitTrans
                        
                                                   rs.Open "select * from vw_notas_cargo where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = 'NC' and vcha_ser_Serie_id = '" + var_serie + "' and inte_Car_numero = " + Str(var_numero_folio), cnn, adOpenDynamic, adLockOptimistic
                                                   
                                                   If Not rs.EOF Then
                                                      rsaux.Open "INSERT INTO TB_IMPRESION_NOTAS_CARGO_1_CLIENTES (VCHA_EMP_EMPRESA_ID, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, VCHA_CLI_CLAVE_ID) VALUES ('" + var_empresa + "', 'NC','" + var_serie + "'," + CStr(var_numero_folio) + ",'" + txt_clave_cliente + "')", cnn, adOpenDynamic, adLockOptimistic
                                                      ''''''''' '''  IMPRESION DE LA NOTA DE CARGO
                                                      If var_empresa = "16" Then
                                                         ''''' nota de cargo para otras empresas
                                                         var_Archivo = App.Path & "\nota_cargo" + Trim(Str(rs!inte_Car_numero)) + ".txt"
                                                         Open (App.Path & "\nota_cargo" + Trim(Str(rs!inte_Car_numero)) + ".txt") For Output As #1
                                                         Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
                                                         Print #1, Chr(27) + Chr(64)
                                                         Print #1, Spc(92); Str(rs!inte_Car_numero)
                                                         Print #1, ""
                                                         Print #1, Spc(92); "       "; Format(rs!dtim_Car_fecha, "Short Date")
                                                         var_cliente = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                                                         For var_j = 1 + Len(Trim(var_cliente)) To 63
                                                             var_cliente = var_cliente + " "
                                                         Next var_j
                                                         var_cliente = var_cliente + " "
                                                         Print #1, ""
                                                         Print #1, Spc(12); var_cliente
                                                         var_domicilio = Trim(IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION))
                                                         var_j = 1 + Len(Trim(var_domicilio))
                                                         For var_j = var_j To 70
                                                             var_domicilio = var_domicilio + " "
                                                         Next var_j
                                                         var_domicilio = var_domicilio + " AGUASCALIENTES, AGS"
                                                         var_j = Len(var_domicilio)
                                                         var_agente = ""
                                                         var_agente = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
                                                         For var_j = 1 + Len(Trim(var_agente)) To 8
                                                             var_agente = var_agente + " "
                                                         Next var_j
                                                         var_agente = var_agente
                                                         var_domicilio = var_domicilio
                                                         Print #1, Spc(12); var_domicilio
                                                         Print #1, Spc(12); IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
                                                         var_ciudad = ""
                                                         var_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                      
                                                         For var_j = 1 + Len(Trim(var_ciudad)) To 14
                                                             var_ciudad = var_ciudad + " "
                                                         Next var_j
                           
                                                         var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                                                         var_ciudad = var_ciudad
                       
                                                         For var_j = 1 + Len(Trim(var_rfc)) To 79
                                                             var_rfc = var_rfc + " "
                                                         Next var_j
                                                         var_rfc = var_rfc + IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
                                                         For var_j = 1 + Len(Trim(var_rfc)) To 103
                                                             var_rfc = var_rfc + " "
                                                         Next var_j
                                                         var_rfc = var_rfc + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
                                                         Print #1, Spc(12); var_ciudad
                                                         Print #1, Spc(12); var_rfc
                                                         Print #1, ""
                                                         Print #1, ""
                                                         var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                                    
                                                         var_linea = "1         " + rs!vcha_Car_nombre + " " + CStr(rsaux11!inte_Car_numero)
                        
                                                         If Len(Trim(var_linea)) < 109 Then
                                                            For var_j = 1 + Len(Trim(var_linea)) To 109
                                                                var_linea = var_linea + " "
                                                            Next var_j
                                                         End If
                                                         If Len(Trim(var_rfc)) = 0 Then
                                                            var_importe_str = Format(0, "###,###,##0.00")
                                                         Else
                                                            var_importe_str = Format(0, "###,###,##0.00")
                                                         End If
                                                         If Len(Trim(var_importe_str)) < 14 Then
                                                            For var_j = 1 + Len(Trim(var_importe_str)) To 14
                                                                var_importe_str = " " + var_importe_str
                                                            Next var_j
                                                         End If
                                                         var_linea = var_linea + " " + var_importe_str
                                                         Print #1, var_linea
                                                          'var_año_str = CStr(Year(rsaux11!DTIM_CAR_FECHA))
                                                          'If Len(var_año_str) = 2 Then
                                                          '   var_año_str = "20" + var_año_str
                                                          'End If
                                                          'var_mes_str = CStr(Month(rsaux11!DTIM_CAR_FECHA))
                                                          'If Len(var_mes_str) = 1 Then
                                                          '   var_mes_str = "0" + var_mes_str
                                                          'End If
                                                          'var_diaS_str = CStr(Day(rsaux11!DTIM_CAR_FECHA))
                                                          'If Len(var_diaS_str) = 1 Then
                                                          '   var_diaS_str = "0" + var_diaS_str
                                                          'End If
                                                          Print #1, "          "
                                                         Print #1, ""
                                                         Print #1, ""
                                                         Print #1, ""
                                                         var_cantidad_letra = rs!vcha_car_importe_letra
                                                         var_linea = IIf(IsNull(rs!vcha_car_importe_letra), "", rs!vcha_car_importe_letra)
                                                         If Len(Trim(var_linea)) < 91 Then
                                                            For var_j = 1 + Len(Trim(var_linea)) To 91
                                                                var_linea = var_linea + " "
                                                            Next var_j
                                                         End If
                                                         Print #1, ""
                                                         If Len(Trim(var_rfc)) = 0 Then
                                                             var_subimporte_str = Format(0, "###,###,##0.00")
                                                             If Len(Trim(var_subimporte_str)) < 14 Then
                                                                For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                                                    var_subimporte_str = " " + var_subimporte_str
                                                                Next var_j
                                                             End If
                                                             var_iva_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                                             If Len(Trim(var_iva_str)) < 14 Then
                                                                For var_j = 1 + Len(Trim(var_iva_str)) To 14
                                                                    var_iva_str = " " + var_iva_str
                                                                Next var_j
                                                             End If
                                                          Else
                                                             var_subimporte_str = Format(0, "###,###,##0.00")
                                                             If Len(Trim(var_subimporte_str)) < 14 Then
                                                                For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                                                    var_subimporte_str = " " + var_subimporte_str
                                                                Next var_j
                                                             End If
                                                             var_iva_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                                             If Len(Trim(var_iva_str)) < 14 Then
                                                                For var_j = 1 + Len(Trim(var_iva_str)) To 14
                                                                    var_iva_str = " " + var_iva_str
                                                                Next var_j
                                                             End If
                                                          End If
                                                          var_linea = var_linea
                                                          Print #1, ""
                                                          Print #1, ""
                                                          Print #1, Spc(8); var_linea
                                                          Print #1, ""
                                                          Print #1, Spc(110); var_subimporte_str
                                                          Print #1, ""
                                                          Print #1, Spc(110); var_iva_str
                                                          Print #1, ""
                                                          var_importe_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                                          If Len(Trim(var_importe_str)) < 14 Then
                                                             For var_j = 1 + Len(Trim(var_importe_str)) To 14
                                                                 var_importe_str = " " + var_importe_str
                                                             Next var_j
                                                          End If
                                                          Print #1, Spc(110); var_importe_str
                                                          Print #1, ""
                                                          Print #1, ""
                                                          Print #1, ""
                                                          Print #1, ""
                                                          Print #1, ""

                                                          Close #1
                                                          rsaux10.Open "update tb_series set inte_ser_nota_cargo =  inte_ser_nota_cargo + 1 where vcha_emp_empresa_id= '" + var_empresa + "' and vcha_ser_serie_id = '" + var_serie + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                                                          txt_numero = CDbl(txt_numero) + 1
                                                          Print #2, "copy " + App.Path + "\nota_cargo" + Trim(Str(rs!inte_Car_numero)) + ".txt lpt1"
                                                       Else
                                                          Open (App.Path & "\nota_cargo" + Trim(Str(rs!inte_Car_numero)) + ".txt") For Output As #1
                                                          Print #1, Chr(15) + Chr(13) + Chr(27) + Chr(67) + Chr(44) + Chr(13)
                                                          Print #1, Spc(92); Str(rs!inte_Car_numero)
                                                          Print #1, ""
                                                          Print #1, ""
                                                          'Print #1, Spc(70); Format(rs!DTIM_CAR_FECHA, "Short Date")
                                                          var_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + " " + IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                                                          For var_j = 1 + Len(Trim(var_cliente)) To 73
                                                              var_cliente = var_cliente + " "
                                                          Next var_j
                                                          var_cliente = var_cliente + Format(rs!dtim_Car_fecha, "Short Date")
                                                          Print #1, Spc(12); var_cliente
                                                          var_domicilio = IIf(IsNull(rs!VCHA_CLI_DIRECCION), "", rs!VCHA_CLI_DIRECCION) + " C.P. " + IIf(IsNull(rs!VCHA_CLI_CP), "", rs!VCHA_CLI_CP)
                                                          For var_j = 1 + Len(Trim(var_domicilio)) To 73
                                                              var_domicilio = var_domicilio + " "
                                                          Next var_j
                                                          var_agente = ""
                                                          'var_agente = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
                                                          'For var_j = 1 + Len(Trim(var_agente)) To 8
                                                          '    var_agente = var_agente + " "
                                                          'Next var_j
                                                          ' var_agente = var_agente + IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
                                                          var_domicilio = var_domicilio
                                                          Print #1, Spc(5); var_domicilio
                                                          var_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
                                                          For var_j = 1 + Len(Trim(var_ciudad)) To 37
                                                              var_ciudad = var_ciudad + " "
                                                          Next var_j
                                                          var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                                                          var_rfc = "RFC:  " + var_rfc
                                                          var_rfc = var_rfc
                                                          var_ciudad = var_ciudad + var_rfc
                                                          Print #1, Spc(5); var_ciudad
                                                          Print #1, ""
                                                          Print #1, ""
                                                          var_linea = "1         " + rs!vcha_Car_nombre + " " + CStr(rsaux11!inte_Car_numero)
                                                          If Len(Trim(var_linea)) < 108 Then
                                                             For var_j = 1 + Len(Trim(var_linea)) To 108
                                                                 var_linea = var_linea + " "
                                                             Next var_j
                                                          End If
                                                          var_rfc = IIf(IsNull(rs!VCHA_CLI_RFC), "", rs!VCHA_CLI_RFC)
                                                          If Len(Trim(var_rfc)) = 0 Then
                                                             var_importe_str = Format(0, "###,###,##0.00")
                                                          Else
                                                             var_importe_str = Format(0, "###,###,##0.00")
                                                          End If
                                                          If Len(Trim(var_importe_str)) < 14 Then
                                                             For var_j = 1 + Len(Trim(var_importe_str)) To 14
                                                                 var_importe_str = " " + var_importe_str
                                                             Next var_j
                                                          End If
                                                          var_linea = var_linea + var_importe_str
                                                          Print #1, var_linea
                                                          'var_año_str = CStr(Year(rsaux11!DTIM_CAR_FECHA))
                                                          'If Len(var_año_str) = 2 Then
                                                          '   var_año_str = "20" + var_año_str
                                                          'End If
                                                          'var_mes_str = CStr(Month(rsaux11!DTIM_CAR_FECHA))
                                                          'If Len(var_mes_str) = 1 Then
                                                          '   var_mes_str = "0" + var_mes_str
                                                          'End If
                                                          'var_diaS_str = CStr(Day(rsaux11!DTIM_CAR_FECHA))
                                                          'If Len(var_diaS_str) = 1 Then
                                                          '   var_diaS_str = "0" + var_diaS_str
                                                          'End If
                                                          'Print #1, "       " + CStr(rsaux11!INTE_CAR_NUMERO) + " DE FECHA " + var_año_str + " " + var_mes_str + " " + var_diaS_str
                                                          Print #1, "       "
                                                          Print #1, ""
                                                          Print #1, ""
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
                                                             var_subimporte_str = Format(0, "###,###,##0.00")
                                                             If Len(Trim(var_subimporte_str)) < 14 Then
                                                                For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                                                    var_subimporte_str = " " + var_subimporte_str
                                                                Next var_j
                                                             End If
                                                             var_iva_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
                                                             If Len(Trim(var_iva_str)) < 14 Then
                                                                For var_j = 1 + Len(Trim(var_iva_str)) To 14
                                                                    var_iva_str = " " + var_iva_str
                                                                Next var_j
                                                             End If
                                                          Else
                                                             var_subimporte_str = Format(0, "###,###,##0.00")
                                                             If Len(Trim(var_subimporte_str)) < 14 Then
                                                                For var_j = 1 + Len(Trim(var_subimporte_str)) To 14
                                                                    var_subimporte_str = " " + var_subimporte_str
                                                                Next var_j
                                                             End If
                                                             var_iva_str = Format((IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto)) / (IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)), "###,###,##0.00")
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
                                                          'Print #1, Spc(4); "ESTA DOCUMENTO SERA PAGADO EN UNA SOLA EXHIBICION"
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
                                                          Print #2, "copy " + App.Path + "\nota_cargo" + Trim(Str(rs!inte_Car_numero)) + ".txt lpt1"
                                                          txt_numero = CDbl(txt_numero) + 1
                                                          rsaux10.Open "update tb_series set inte_ser_nota_cargo =  inte_ser_nota_cargo + 1 where vcha_emp_empresa_id= '" + var_empresa + "' and vcha_ser_serie_id = '" + var_serie + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                                                       End If
'''''''''
                                                       'txt_clase_cartera.Enabled = True
                                                       'txt_nombre_clase_cartera.Enabled = True
                                                       'txt_plazo.Enabled = True
                                                       'txt_importe.Enabled = True
                                                       cmb_clientes = ""
                                                       txt_clave_cliente = ""
                                                       txt_plazo = ""
                                                       txt_importe = ""
                                                       txt_clave_empresa = ""
                                                       lbl_moneda = ""
                                                       txt_nombre_clase_cartera = ""
                                                       txt_clase_cartera = ""
                                                       txt_nombre_cliente = ""
                                                    End If
                                                    If rs.State = 1 Then
                                                       rs.Close
                                                    End If
                                                 Else
                                                    MsgBox "La impresión de la Nota de Cargo a sido cancelada", vbOKOnly, "ATENCION"
                                                 End If
                                              Else
                                                 rs.Close
                                                 MsgBox "La nota de cargo ya existe", vbOKOnly, "ATENCION"
                                              End If
                                           End If
                                        End If
                                     Else
                                        MsgBox "No se a asignado el tipo de cambio del dia de hoy", vbOKOnly, "ATENCION"
                                     End If
                                  Else
                                     MsgBox "Importe Incorrecto", vbOKOnly, "ATENCION"
                                  End If
                               Else
                                  MsgBox "Clave de cartera incorrecta", vbOKOnly, "ATENCION"
                               End If
                            Else
                               MsgBox "No se a seleccionado un cliente", vbOKOnly, "ATENCION"
                            End If
                         Else
                            MsgBox "Numero de documento incorecto", vbOKOnly, "ATENCION"
                         End If
                      Else
                         MsgBox "Se debe de indicar un numero de documento", vbOKOnly, "ATENCION"
                      End If
                   Else
                      MsgBox "Se debe de indicar una serie", vbOKOnly, "ATENCION"
                   End If
                   'fin de la rutina de impresion de facturas cuando el iva ya es del 16%
                   End If
                   rsaux11.MoveNext
            Wend
            Close #2
            x = Shell(var_archivo_bat, vbHide)
            MsgBox "A terminado la impresion de las notas de cargo", vbOKOnly, "ATENCION"
            rsaux11.Close
         Else
            MsgBox "La unidad organizacional no tiene una serie", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub

Private Sub Form_Load()
   Top = 2500
   Left = 3200
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_articulos2)
End Sub
