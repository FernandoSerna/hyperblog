VERSION 5.00
Begin VB.Form frmreporte_Facturacion_multibondeados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de facturación"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "A"
      Height          =   315
      Left            =   1110
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Reporte de facturación general"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "G"
      Height          =   315
      Left            =   780
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Reporte de facturación general"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "L"
      Height          =   315
      Left            =   450
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Reporte de facturación por linea"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Caption         =   "D"
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Reporte de facturación a detalle"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4020
      Picture         =   "frmreporte_Facturacion_multibondeados.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   840
      Left            =   135
      TabIndex        =   0
      Top             =   435
      Width           =   4245
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   2865
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   300
         Width           =   1140
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2325
         TabIndex        =   4
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   375
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   45
      TabIndex        =   7
      Top             =   270
      Width           =   4395
   End
End
Attribute VB_Name = "frmreporte_Facturacion_multibondeados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim año As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_consecutivo As Integer
   Dim var_afectacion_movimiento As String
   Dim var_vistas As String
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_FACTURACION_MULTIBONDEADOS", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into TB_TEMP_FACTURACION_MULTIBONDEADOS (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            var_fecha_fin_1 = CDate(txt_fin) + 1
            var_dia = CStr(Day(CDate(txt_inicio)))
            var_mes = CStr(Month(CDate(txt_inicio)))
            var_año = CStr(Year(CDate(txt_inicio)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            'rs.Open "select * from vw_encabezado_movimientos where dtim_emo_fecha >= " + var_fecha_inicio + " and dtim_emo_fecha <= " + var_fecha_fin + "-.00001", cnn, adOpenDynamic, adLockOptimistic
            
            var_fecha_fin_1 = CDate(txt_fin)
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin_2 = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            rs.Open "delete from TB_TEMP_FACTURACION_MULTIBONDEADOS where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and dtim_tem_fecha_inicio is null", cnn, adOpenDynamic, adLockOptimistic
            var_cadena = " INSERT INTO TB_TEMP_FACTURACION_MULTIBONDEADOS (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, VCHA_ESB_ESTABLECIMIENTO_ID, VCHA_ESB_NOMBRE, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, FLOA_SAL_CANTIDAD, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO_1, FLOA_SAL_DESCUENTO_2, VCHA_LIN_LINEA_ID, VCHA_LIN_NOMBRE, VCHA_EMP_EMPRESA_ID, VCHA_EMP_NOMBRE, VCHA_ART_ARTICULO_ID, VCHA_ART_NOMBRE, CANTIDAD_KILOS, DTIM_CAR_FECHA, INTE_CAR_PLAZO, CHAR_CAR_eSTATUS, FLOA_TEM_VALOR_1, FLOA_TEM_VALOR_2)"
            var_cadena = var_cadena + " select " + CStr(var_consecutivo) + ",    " + var_fecha_inicio + "," + var_fecha_fin + ", VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, VCHA_ESB_ESTABLECIMIENTO_ID,VCHA_ESB_NOMBRE, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, FLOA_SAL_CANTIDAD, FLOA_SAL_PRECIO* (1 +(floa_iva_iva/100)), FLOA_SAL_DESCUENTO_1, FLOA_SAL_DESCUENTO_2, VCHA_LIN_LINEA_ID, VCHA_LIN_NOMBRE, VCHA_EMP_EMPRESA_ID, VCHA_EMP_NOMBRE, VCHA_ART_ARTICULO_ID, VCHA_ART_NOMBRE_español, CANTIDAD_KILOS, DTIM_CAR_FECHA, INTE_CAR_PLAZO, ISNULL(CHAR_CAR_ESTATUS,'I'), VALOR_1, VALOR_2 from VW_FACTURACION_MULTIBONDEADOS where dtim_car_fecha >=" + var_fecha_inicio + " and dtim_car_fecha <= " + var_fecha_fin + " order by inte_Car_numero, vcha_Art_articulo_id"
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            rs.Open "select * from TB_TEMP_FACTURACION_MULTIBONDEADOS where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_Art_articulo_id is not null", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  var_codigo = rs!vcha_Art_articulo_id
                  var_codigo_1 = Mid(var_codigo, 1, 3)
                  var_codigo_2 = Mid(var_codigo, 5, 3)
                  If IsNumeric(var_codigo_1) Then
                     If IsNumeric(var_codigo_2) Then
                        If Mid(var_codigo, 4, 1) = "-" And Mid(var_codigo, 8, 1) = "-" Then
                           var_cantidad_kilos = (CDbl(var_codigo_1) * CDbl(var_codigo_2) * rs!floa_sal_cantidad) / 100000
                        Else
                           var_cantidad_kilos = 0
                        End If
                        rsaux2.Open "update TB_TEMP_FACTURACION_MULTIBONDEADOS set CANTIDAD_KILOS = " + CStr(var_cantidad_kilos) + " where inte_tem_consecutivo_tabla = " + CStr(rs!inte_tem_consecutivo_tabla), cnn, adOpenDynamic, adLockOptimistic
                     End If
                  End If
                  rs.MoveNext
            Wend
            rs.Close
            
            
            rs.Open "select * from tb_empresas where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
            var_nombre_empresa = ""
            If Not rs.EOF Then
               var_nombre_empresa = IIf(IsNull(rs!VCHA_EMP_NOMBRE), "", rs!VCHA_EMP_NOMBRE)
            End If
            rs.Close
            rs.Open "select vcha_emp_empresa_id, vcha_Car_documento, dtim_Car_fecha, inte_car_numero, vcha_ser_serie_id from tb_encabezado_cartera where vcha_emp_empresa_id = '16' and vcha_Car_documento = 'FA' and dtim_car_fecha >=" + var_fecha_inicio + " and dtim_car_fecha <= " + var_fecha_fin + " and char_car_estatus = 'C'", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  rsaux.Open "select * from TB_TEMP_FACTURACION_MULTIBONDEADOS where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_emp_empresa_id = '16' and vcha_Ser_serie_id = '" + IIf(IsNull(rs!vcha_ser_Serie_id), "", rs!vcha_ser_Serie_id) + "' and inte_Car_numero = " + CStr(IIf(IsNull(rs!inte_car_numero), 0, rs!inte_car_numero)), cnn, adOpenDynamic, adLockOptimistic
                  If rsaux.EOF Then
                     var_cadena = "INSERT INTO TB_TEMP_FACTURACION_MULTIBONDEADOS (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, VCHA_ESB_ESTABLECIMIENTO_ID, VCHA_ESB_NOMBRE, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, FLOA_SAL_CANTIDAD, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO_1, FLOA_SAL_DESCUENTO_2, VCHA_LIN_LINEA_ID, VCHA_LIN_NOMBRE, VCHA_EMP_EMPRESA_ID, VCHA_EMP_NOMBRE, VCHA_ART_ARTICULO_ID, VCHA_ART_NOMBRE, CANTIDAD_KILOS, DTIM_CAR_FECHA, INTE_CAR_PLAZO, CHAR_CAR_eSTATUS, FLOA_TEM_VALOR_1, FLOA_TEM_VALOR_2) "
                     var_cadena = var_cadena + " values (" + CStr(var_consecutivo) + ",    " + var_fecha_inicio + "," + var_fecha_fin + ", '', '', '','', '" + IIf(IsNull(rs!vcha_ser_Serie_id), "", rs!vcha_ser_Serie_id) + "', " + CStr(IIf(IsNull(rs!inte_car_numero), 0, rs!inte_car_numero)) + ", 0, 0, 0, 0, '', '', '" + var_empresa + "', '" + var_nombre_empresa + "', '', '', 0, " + Format(IIf(IsNull(rs!dtim_Car_fecha), "", rs!dtim_Car_fecha), "Short Date") + ", 0, 'C', '', '')"
                     rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux.Close
                  rs.MoveNext
            Wend
            rs.Close
            
            rs.Open "select distinct vcha_cli_clave_id from TB_TEMP_FACTURACION_MULTIBONDEADOS where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  rsaux1.Open "select vcha_Age_Agente_id from tb_clientes where vcha_cli_clave_id = '" + IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                  var_agente = ""
                  If Not rsaux1.EOF Then
                     var_agente = IIf(IsNull(rsaux1!VCHA_AGE_AGENTE_ID), "", rsaux1!VCHA_AGE_AGENTE_ID)
                  End If
                  rsaux1.Close
                  rsaux1.Open "update TB_TEMP_FACTURACION_MULTIBONDEADOS set vcha_Age_agente_id = '" + var_agente + "' where vcha_cli_clave_id = '" + IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
            Wend
            rs.Close
            
            rs.Open "UPDATE TB_TEMP_FACTURACION_MULTIBONDEADOS SET INTE_TEM_TERCEROS = 1 WHERE VCHA_aGE_AGENTE_ID = '00100' AND INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            
            
            Set reporte = appl.OpenReport(App.Path + "\rep_facturacion_multibondeados_detalle.rpt")
            reporte.RecordSelectionFormula = "{TB_TEMP_FACTURACION_MULTIBONDEADOS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Entradas concentrado"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            
            
            var_si = MsgBox("żDesea exportar el reporte a excel?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_facturacion_multibondeados_detalle.rpt")
               reporte.RecordSelectionFormula = "{TB_TEMP_FACTURACION_MULTIBONDEADOS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\reporte_facturacion_detalle_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
            End If
            
            
            
            rs.Open "delete from TB_TEMP_FACTURACION_MULTIBONDEADOS where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "La fecha de inicio debe de ser menor o igual a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de Inicio incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim año As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_consecutivo As Integer
   Dim var_afectacion_movimiento As String
   Dim var_vistas As String
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_FACTURACION_MULTIBONDEADOS", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into TB_TEMP_FACTURACION_MULTIBONDEADOS (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            var_fecha_fin_1 = CDate(txt_fin) + 1
            var_dia = CStr(Day(CDate(txt_inicio)))
            var_mes = CStr(Month(CDate(txt_inicio)))
            var_año = CStr(Year(CDate(txt_inicio)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            'rs.Open "select * from vw_encabezado_movimientos where dtim_emo_fecha >= " + var_fecha_inicio + " and dtim_emo_fecha <= " + var_fecha_fin + "-.00001", cnn, adOpenDynamic, adLockOptimistic
            
            var_fecha_fin_1 = CDate(txt_fin)
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin_2 = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            rs.Open "delete from TB_TEMP_FACTURACION_MULTIBONDEADOS where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and dtim_tem_fecha_inicio is null", cnn, adOpenDynamic, adLockOptimistic
            var_cadena = " INSERT INTO TB_TEMP_FACTURACION_MULTIBONDEADOS (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, VCHA_ESB_ESTABLECIMIENTO_ID, VCHA_ESB_NOMBRE, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, FLOA_SAL_CANTIDAD, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO_1, FLOA_SAL_DESCUENTO_2, VCHA_LIN_LINEA_ID, VCHA_LIN_NOMBRE, VCHA_EMP_EMPRESA_ID, VCHA_EMP_NOMBRE, VCHA_ART_ARTICULO_ID, VCHA_ART_NOMBRE, CANTIDAD_KILOS, DTIM_CAR_FECHA, INTE_CAR_PLAZO,FLOA_TEM_VALOR_1, FLOA_TEM_VALOR_2)"
            var_cadena = var_cadena + " select " + CStr(var_consecutivo) + ",    " + var_fecha_inicio + "," + var_fecha_fin + ", VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, VCHA_ESB_ESTABLECIMIENTO_ID,VCHA_ESB_NOMBRE, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, FLOA_SAL_CANTIDAD, FLOA_SAL_PRECIO* (1 +(floa_iva_iva/100)), FLOA_SAL_DESCUENTO_1, FLOA_SAL_DESCUENTO_2, VCHA_LIN_LINEA_ID, VCHA_LIN_NOMBRE, VCHA_EMP_EMPRESA_ID, VCHA_EMP_NOMBRE, VCHA_ART_ARTICULO_ID, VCHA_ART_NOMBRE_español, CANTIDAD_KILOS, DTIM_CAR_FECHA, INTE_CAR_PLAZO, VALOR_1, VALOR_2 from VW_FACTURACION_MULTIBONDEADOS where dtim_car_fecha >=" + var_fecha_inicio + " and dtim_car_fecha <= " + var_fecha_fin + " order by inte_Car_numero, vcha_Art_articulo_id"
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            
            
            
            rs.Open "select * from TB_TEMP_FACTURACION_MULTIBONDEADOS where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_Art_articulo_id is not null", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  var_codigo = rs!vcha_Art_articulo_id
                  var_codigo_1 = Mid(var_codigo, 1, 3)
                  var_codigo_2 = Mid(var_codigo, 5, 3)
                  If IsNumeric(var_codigo_1) Then
                     If IsNumeric(var_codigo_2) Then
                        If Mid(var_codigo, 4, 1) = "-" And Mid(var_codigo, 8, 1) = "-" Then
                           var_cantidad_kilos = (CDbl(var_codigo_1) * CDbl(var_codigo_2) * rs!floa_sal_cantidad) / 100000
                        Else
                           var_cantidad_kilos = 0
                        End If
                        rsaux2.Open "update TB_TEMP_FACTURACION_MULTIBONDEADOS set CANTIDAD_KILOS = " + CStr(var_cantidad_kilos) + " where inte_tem_consecutivo_tabla = " + CStr(rs!inte_tem_consecutivo_tabla), cnn, adOpenDynamic, adLockOptimistic
                     End If
                  End If
                  rs.MoveNext
            Wend
            rs.Close
            
            
            
            rs.Open "select * from tb_empresas where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
            var_nombre_empresa = ""
            If Not rs.EOF Then
               var_nombre_empresa = IIf(IsNull(rs!VCHA_EMP_NOMBRE), "", rs!VCHA_EMP_NOMBRE)
            End If
            rs.Close
            rs.Open "select vcha_emp_empresa_id, vcha_Car_documento, dtim_Car_fecha, inte_car_numero, vcha_ser_serie_id from tb_encabezado_cartera where vcha_emp_empresa_id = '16' and vcha_Car_documento = 'FA' and dtim_car_fecha >=" + var_fecha_inicio + " and dtim_car_fecha <= " + var_fecha_fin + " and char_car_estatus = 'C'", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  rsaux.Open "select * from TB_TEMP_FACTURACION_MULTIBONDEADOS where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_emp_empresa_id = '16' and vcha_Ser_serie_id = '" + IIf(IsNull(rs!vcha_ser_Serie_id), "", rs!vcha_ser_Serie_id) + "' and inte_Car_numero = " + CStr(IIf(IsNull(rs!inte_car_numero), 0, rs!inte_car_numero)), cnn, adOpenDynamic, adLockOptimistic
                  If rsaux.EOF Then
                     var_cadena = "INSERT INTO TB_TEMP_FACTURACION_MULTIBONDEADOS (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, VCHA_ESB_ESTABLECIMIENTO_ID, VCHA_ESB_NOMBRE, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, FLOA_SAL_CANTIDAD, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO_1, FLOA_SAL_DESCUENTO_2, VCHA_LIN_LINEA_ID, VCHA_LIN_NOMBRE, VCHA_EMP_EMPRESA_ID, VCHA_EMP_NOMBRE, VCHA_ART_ARTICULO_ID, VCHA_ART_NOMBRE, CANTIDAD_KILOS, DTIM_CAR_FECHA, INTE_CAR_PLAZO, CHAR_CAR_eSTATUS, FLOA_TEM_VALOR_1, FLOA_TEM_VALOR_2) "
                     var_cadena = var_cadena + " values (" + CStr(var_consecutivo) + ",    " + var_fecha_inicio + "," + var_fecha_fin + ", '', '', '','', '" + IIf(IsNull(rs!vcha_ser_Serie_id), "", rs!vcha_ser_Serie_id) + "', " + CStr(IIf(IsNull(rs!inte_car_numero), 0, rs!inte_car_numero)) + ", 0, 0, 0, 0, '', '', '" + var_empresa + "', '" + var_nombre_empresa + "', '', '', 0, " + Format(IIf(IsNull(rs!dtim_Car_fecha), "", rs!dtim_Car_fecha), "Short Date") + ", 0, 'C', '', '')"
                     rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux.Close
                  rs.MoveNext
            Wend
            rs.Close
            
            
            rs.Open "select distinct vcha_cli_clave_id from TB_TEMP_FACTURACION_MULTIBONDEADOS where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  rsaux1.Open "select vcha_Age_Agente_id from tb_clientes where vcha_cli_clave_id = '" + IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                  var_agente = ""
                  If Not rsaux1.EOF Then
                     var_agente = IIf(IsNull(rsaux1!VCHA_AGE_AGENTE_ID), "", rsaux1!VCHA_AGE_AGENTE_ID)
                  End If
                  rsaux1.Close
                  rsaux1.Open "update TB_TEMP_FACTURACION_MULTIBONDEADOS set vcha_Age_agente_id = '" + var_agente + "' where vcha_cli_clave_id = '" + IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
            Wend
            rs.Close
            
            rs.Open "UPDATE TB_TEMP_FACTURACION_MULTIBONDEADOS SET INTE_TEM_TERCEROS = 1 WHERE VCHA_aGE_AGENTE_ID = '00100' AND INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            
            
            
            Set reporte = appl.OpenReport(App.Path + "\rep_facturacion_multibondeados_linea.rpt")
            reporte.RecordSelectionFormula = "{VW_FACTURACION_MULTIBONDEADOS_LINEA.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Entradas concentrado"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            
            
            var_si = MsgBox("żDesea exportar el reporte a excel?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_facturacion_multibondeados_linea.rpt")
               reporte.RecordSelectionFormula = "{VW_FACTURACION_MULTIBONDEADOS_LINEA.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\reporte_facturacion_linea_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
            End If
            
            
            
            
            rs.Open "delete from TB_TEMP_FACTURACION_MULTIBONDEADOS where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "La fecha de inicio debe de ser menor o igual a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de Inicio incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Command2_Click()
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim año As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_consecutivo As Integer
   Dim var_afectacion_movimiento As String
   Dim var_vistas As String
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_FACTURACION_MULTIBONDEADOS", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into TB_TEMP_FACTURACION_MULTIBONDEADOS (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            var_fecha_fin_1 = CDate(txt_fin) + 1
            var_dia = CStr(Day(CDate(txt_inicio)))
            var_mes = CStr(Month(CDate(txt_inicio)))
            var_año = CStr(Year(CDate(txt_inicio)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            'rs.Open "select * from vw_encabezado_movimientos where dtim_emo_fecha >= " + var_fecha_inicio + " and dtim_emo_fecha <= " + var_fecha_fin + "-.00001", cnn, adOpenDynamic, adLockOptimistic
            
            var_fecha_fin_1 = CDate(txt_fin)
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin_2 = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            rs.Open "delete from TB_TEMP_FACTURACION_MULTIBONDEADOS where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and dtim_tem_fecha_inicio is null", cnn, adOpenDynamic, adLockOptimistic
            var_cadena = " INSERT INTO TB_TEMP_FACTURACION_MULTIBONDEADOS (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, VCHA_ESB_ESTABLECIMIENTO_ID, VCHA_ESB_NOMBRE, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, FLOA_SAL_CANTIDAD, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO_1, FLOA_SAL_DESCUENTO_2, VCHA_LIN_LINEA_ID, VCHA_LIN_NOMBRE, VCHA_EMP_EMPRESA_ID, VCHA_EMP_NOMBRE, VCHA_ART_ARTICULO_ID, VCHA_ART_NOMBRE, CANTIDAD_KILOS, DTIM_CAR_FECHA, INTE_CAR_PLAZO, FLOA_TEM_VALOR_1, FLOA_TEM_VALOR_2)"
            var_cadena = var_cadena + " select " + CStr(var_consecutivo) + ",    " + var_fecha_inicio + "," + var_fecha_fin + ", VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, VCHA_ESB_ESTABLECIMIENTO_ID,VCHA_ESB_NOMBRE, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, FLOA_SAL_CANTIDAD, FLOA_SAL_PRECIO* (1 +(floa_iva_iva/100)), FLOA_SAL_DESCUENTO_1, FLOA_SAL_DESCUENTO_2, VCHA_LIN_LINEA_ID, VCHA_LIN_NOMBRE, VCHA_EMP_EMPRESA_ID, VCHA_EMP_NOMBRE, VCHA_ART_ARTICULO_ID, VCHA_ART_NOMBRE_español, CANTIDAD_KILOS, DTIM_CAR_FECHA, INTE_CAR_PLAZO, VALOR_1, VALOR_2 from VW_FACTURACION_MULTIBONDEADOS where dtim_car_fecha >=" + var_fecha_inicio + " and dtim_car_fecha <= " + var_fecha_fin + " order by inte_Car_numero, vcha_Art_articulo_id"
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            
            rs.Open "select * from TB_TEMP_FACTURACION_MULTIBONDEADOS where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_Art_articulo_id is not null", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  var_codigo = rs!vcha_Art_articulo_id
                  var_codigo_1 = Mid(var_codigo, 1, 3)
                  var_codigo_2 = Mid(var_codigo, 5, 3)
                  If IsNumeric(var_codigo_1) Then
                     If IsNumeric(var_codigo_2) Then
                        If Mid(var_codigo, 4, 1) = "-" And Mid(var_codigo, 8, 1) = "-" Then
                           var_cantidad_kilos = (CDbl(var_codigo_1) * CDbl(var_codigo_2) * rs!floa_sal_cantidad) / 100000
                        Else
                           var_cantidad_kilos = 0
                        End If
                        rsaux2.Open "update TB_TEMP_FACTURACION_MULTIBONDEADOS set CANTIDAD_KILOS = " + CStr(var_cantidad_kilos) + " where inte_tem_consecutivo_tabla = " + CStr(rs!inte_tem_consecutivo_tabla), cnn, adOpenDynamic, adLockOptimistic
                     End If
                  End If
                  rs.MoveNext
            Wend
            rs.Close
            
            rs.Open "select * from tb_empresas where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
            var_nombre_empresa = ""
            If Not rs.EOF Then
               var_nombre_empresa = IIf(IsNull(rs!VCHA_EMP_NOMBRE), "", rs!VCHA_EMP_NOMBRE)
            End If
            rs.Close
            rs.Open "select vcha_emp_empresa_id, vcha_Car_documento, dtim_Car_fecha, inte_car_numero, vcha_ser_serie_id from tb_encabezado_cartera where vcha_emp_empresa_id = '16' and vcha_Car_documento = 'FA' and dtim_car_fecha >=" + var_fecha_inicio + " and dtim_car_fecha <= " + var_fecha_fin + " and char_car_estatus = 'C'", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  rsaux.Open "select * from TB_TEMP_FACTURACION_MULTIBONDEADOS where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_emp_empresa_id = '16' and vcha_Ser_serie_id = '" + IIf(IsNull(rs!vcha_ser_Serie_id), "", rs!vcha_ser_Serie_id) + "' and inte_Car_numero = " + CStr(IIf(IsNull(rs!inte_car_numero), 0, rs!inte_car_numero)), cnn, adOpenDynamic, adLockOptimistic
                  If rsaux.EOF Then
                     var_cadena = "INSERT INTO TB_TEMP_FACTURACION_MULTIBONDEADOS (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, VCHA_ESB_ESTABLECIMIENTO_ID, VCHA_ESB_NOMBRE, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, FLOA_SAL_CANTIDAD, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO_1, FLOA_SAL_DESCUENTO_2, VCHA_LIN_LINEA_ID, VCHA_LIN_NOMBRE, VCHA_EMP_EMPRESA_ID, VCHA_EMP_NOMBRE, VCHA_ART_ARTICULO_ID, VCHA_ART_NOMBRE, CANTIDAD_KILOS, DTIM_CAR_FECHA, INTE_CAR_PLAZO, CHAR_CAR_eSTATUS, FLOA_TEM_VALOR_1, FLOA_TEM_VALOR_2) "
                     var_cadena = var_cadena + " values (" + CStr(var_consecutivo) + ",    " + var_fecha_inicio + "," + var_fecha_fin + ", '', '', '','', '" + IIf(IsNull(rs!vcha_ser_Serie_id), "", rs!vcha_ser_Serie_id) + "', " + CStr(IIf(IsNull(rs!inte_car_numero), 0, rs!inte_car_numero)) + ", 0, 0, 0, 0, '', '', '" + var_empresa + "', '" + var_nombre_empresa + "', '', '', 0, " + Format(IIf(IsNull(rs!dtim_Car_fecha), "", rs!dtim_Car_fecha), "Short Date") + ", 0, 'C', '', '')"
                     rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux.Close
                  rs.MoveNext
            Wend
            rs.Close
            
            
            rs.Open "select distinct vcha_cli_clave_id from TB_TEMP_FACTURACION_MULTIBONDEADOS where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  rsaux1.Open "select vcha_Age_Agente_id from tb_clientes where vcha_cli_clave_id = '" + IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                  var_agente = ""
                  If Not rsaux1.EOF Then
                     var_agente = IIf(IsNull(rsaux1!VCHA_AGE_AGENTE_ID), "", rsaux1!VCHA_AGE_AGENTE_ID)
                  End If
                  rsaux1.Close
                  rsaux1.Open "update TB_TEMP_FACTURACION_MULTIBONDEADOS set vcha_Age_agente_id = '" + var_agente + "' where vcha_cli_clave_id = '" + IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
            Wend
            rs.Close
            
            rs.Open "UPDATE TB_TEMP_FACTURACION_MULTIBONDEADOS SET INTE_TEM_TERCEROS = 1 WHERE VCHA_aGE_AGENTE_ID = '00100' AND INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            
            
            
            Set reporte = appl.OpenReport(App.Path + "\rep_facturacion_multibondeados_general.rpt")
            reporte.RecordSelectionFormula = "{VW_FACTURACION_MULTIBONDEADOS_GENERAL.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Entradas concentrado"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            
            
            var_si = MsgBox("żDesea exportar el reporte a excel?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_facturacion_multibondeados_general.rpt")
               reporte.RecordSelectionFormula = "{VW_FACTURACION_MULTIBONDEADOS_GENERAL.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\reporte_facturacion_general_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
            End If
            
            
            
            
            rs.Open "delete from TB_TEMP_FACTURACION_MULTIBONDEADOS where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "La fecha de inicio debe de ser menor o igual a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de Inicio incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Command3_Click()
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim año As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_consecutivo As Integer
   Dim var_afectacion_movimiento As String
   Dim var_vistas As String
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_FACTURACION_MULTIBONDEADOS", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into TB_TEMP_FACTURACION_MULTIBONDEADOS (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            var_fecha_fin_1 = CDate(txt_fin) + 1
            var_dia = CStr(Day(CDate(txt_inicio)))
            var_mes = CStr(Month(CDate(txt_inicio)))
            var_año = CStr(Year(CDate(txt_inicio)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            'rs.Open "select * from vw_encabezado_movimientos where dtim_emo_fecha >= " + var_fecha_inicio + " and dtim_emo_fecha <= " + var_fecha_fin + "-.00001", cnn, adOpenDynamic, adLockOptimistic
            
            var_fecha_fin_1 = CDate(txt_fin)
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin_2 = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            rs.Open "delete from TB_TEMP_FACTURACION_MULTIBONDEADOS where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and dtim_tem_fecha_inicio is null", cnn, adOpenDynamic, adLockOptimistic
            var_cadena = " INSERT INTO TB_TEMP_FACTURACION_MULTIBONDEADOS (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, VCHA_ESB_ESTABLECIMIENTO_ID, VCHA_ESB_NOMBRE, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, FLOA_SAL_CANTIDAD, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO_1, FLOA_SAL_DESCUENTO_2, VCHA_LIN_LINEA_ID, VCHA_LIN_NOMBRE, VCHA_EMP_EMPRESA_ID, VCHA_EMP_NOMBRE, VCHA_ART_ARTICULO_ID, VCHA_ART_NOMBRE, CANTIDAD_KILOS, DTIM_CAR_FECHA, INTE_CAR_PLAZO, FLOA_TEM_VALOR_1, FLOA_TEM_VALOR_2)"
            var_cadena = var_cadena + " select " + CStr(var_consecutivo) + ",    " + var_fecha_inicio + "," + var_fecha_fin + ", VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, VCHA_ESB_ESTABLECIMIENTO_ID,VCHA_ESB_NOMBRE, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, FLOA_SAL_CANTIDAD, FLOA_SAL_PRECIO* (1 +(floa_iva_iva/100)), FLOA_SAL_DESCUENTO_1, FLOA_SAL_DESCUENTO_2, VCHA_LIN_LINEA_ID, VCHA_LIN_NOMBRE, VCHA_EMP_EMPRESA_ID, VCHA_EMP_NOMBRE, VCHA_ART_ARTICULO_ID, VCHA_ART_NOMBRE_español, CANTIDAD_KILOS, DTIM_CAR_FECHA, INTE_CAR_PLAZO, VALOR_1, VALOR_2 from VW_FACTURACION_MULTIBONDEADOS where dtim_car_fecha >=" + var_fecha_inicio + " and dtim_car_fecha <= " + var_fecha_fin + " order by inte_Car_numero, vcha_Art_articulo_id"
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            
            rs.Open "select * from TB_TEMP_FACTURACION_MULTIBONDEADOS where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_Art_articulo_id is not null", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  var_codigo = rs!vcha_Art_articulo_id
                  var_codigo_1 = Mid(var_codigo, 1, 3)
                  var_codigo_2 = Mid(var_codigo, 5, 3)
                  If IsNumeric(var_codigo_1) Then
                     If IsNumeric(var_codigo_2) Then
                        If Mid(var_codigo, 4, 1) = "-" And Mid(var_codigo, 8, 1) = "-" Then
                           var_cantidad_kilos = (CDbl(var_codigo_1) * CDbl(var_codigo_2) * rs!floa_sal_cantidad) / 100000
                        Else
                           var_cantidad_kilos = 0
                        End If
                        rsaux2.Open "update TB_TEMP_FACTURACION_MULTIBONDEADOS set CANTIDAD_KILOS = " + CStr(var_cantidad_kilos) + " where inte_tem_consecutivo_tabla = " + CStr(rs!inte_tem_consecutivo_tabla), cnn, adOpenDynamic, adLockOptimistic
                     End If
                  End If
                  rs.MoveNext
            Wend
            rs.Close
            
            
            
            rs.Open "select * from tb_empresas where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
            var_nombre_empresa = ""
            If Not rs.EOF Then
               var_nombre_empresa = IIf(IsNull(rs!VCHA_EMP_NOMBRE), "", rs!VCHA_EMP_NOMBRE)
            End If
            rs.Close
            rs.Open "select vcha_emp_empresa_id, vcha_Car_documento, dtim_Car_fecha, inte_car_numero, vcha_ser_serie_id from tb_encabezado_cartera where vcha_emp_empresa_id = '16' and vcha_Car_documento = 'FA' and dtim_car_fecha >=" + var_fecha_inicio + " and dtim_car_fecha <= " + var_fecha_fin + " and char_car_estatus = 'C'", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  rsaux.Open "select * from TB_TEMP_FACTURACION_MULTIBONDEADOS where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_emp_empresa_id = '16' and vcha_Ser_serie_id = '" + IIf(IsNull(rs!vcha_ser_Serie_id), "", rs!vcha_ser_Serie_id) + "' and inte_Car_numero = " + CStr(IIf(IsNull(rs!inte_car_numero), 0, rs!inte_car_numero)), cnn, adOpenDynamic, adLockOptimistic
                  If rsaux.EOF Then
                     var_cadena = "INSERT INTO TB_TEMP_FACTURACION_MULTIBONDEADOS (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, VCHA_ESB_ESTABLECIMIENTO_ID, VCHA_ESB_NOMBRE, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, FLOA_SAL_CANTIDAD, FLOA_SAL_PRECIO, FLOA_SAL_DESCUENTO_1, FLOA_SAL_DESCUENTO_2, VCHA_LIN_LINEA_ID, VCHA_LIN_NOMBRE, VCHA_EMP_EMPRESA_ID, VCHA_EMP_NOMBRE, VCHA_ART_ARTICULO_ID, VCHA_ART_NOMBRE, CANTIDAD_KILOS, DTIM_CAR_FECHA, INTE_CAR_PLAZO, CHAR_CAR_eSTATUS, FLOA_TEM_VALOR_1, FLOA_TEM_VALOR_2) "
                     var_cadena = var_cadena + " values (" + CStr(var_consecutivo) + ",    " + var_fecha_inicio + "," + var_fecha_fin + ", '', '', '','', '" + IIf(IsNull(rs!vcha_ser_Serie_id), "", rs!vcha_ser_Serie_id) + "', " + CStr(IIf(IsNull(rs!inte_car_numero), 0, rs!inte_car_numero)) + ", 0, 0, 0, 0, '', '', '" + var_empresa + "', '" + var_nombre_empresa + "', '', '', 0, " + Format(IIf(IsNull(rs!dtim_Car_fecha), "", rs!dtim_Car_fecha), "Short Date") + ", 0, 'C', '', '')"
                     rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux.Close
                  rs.MoveNext
            Wend
            rs.Close
            
            
            
            
            Set reporte = appl.OpenReport(App.Path + "\rep_facturacion_multibondeados_articulo.rpt")
            reporte.RecordSelectionFormula = "{VW_FACTURACION_MULTIBONDEADOS_ARTICULOS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Entradas concentrado"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            
            
            var_si = MsgBox("żDesea exportar el reporte a excel?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_facturacion_multibondeados_articulo.rpt")
               reporte.RecordSelectionFormula = "{VW_FACTURACION_MULTIBONDEADOS_ARTICULOS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\reporte_facturacion_articulo_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
            End If
            rs.Open "delete from TB_TEMP_FACTURACION_MULTIBONDEADOS where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "La fecha de inicio debe de ser menor o igual a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de Inicio incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Form_Load()
Dim dl As Long                                 ' Valor devuelto por la función API
Dim sAttributes As String                  ' Aributos
Dim sDriver As String                       ' Nombre del controlador
Dim sDescription As String                ' Descripción del DSN
Dim sDsnName As String                  ' Nombre del DSN





   cnn.Close
   cnn.Open var_conexion_string_distribucion

   Const ODBC_ADD_SYS_DSN As Long = 4         ' Se creará un DSN de sistema
   Const vbAPINull As Long = 0&                         ' Puntero NULL

   ' se elimina
   Const ODBC_REMOVE_SYS_DSN As Long = 6    ' Se eliminará un DSN de sistema
   sDsnName = "DSN=sqlsistema"
   sDriver = "SQL Server"
   dl = SQLConfigDataSource(0&, ODBC_REMOVE_SYS_DSN, sDriver, sDsnName)

   'se crea
   sDsnName = "sqlsistema"
   sDescription = "sqlsistema"
   sDriver = "SQL Server"
   sAttributes = "DSN=" & sDsnName & Chr(0)
   If UCase(Trim(parametros(0))) = "ADMCDINDUSTRIAL" Then
      sAttributes = sAttributes & "Server=ADMCDINDUSTRIAL" & Chr$(0)
   Else
      sAttributes = sAttributes & "Server=" + parametros(0) & Chr$(0)
   End If
   sAttributes = sAttributes & "Description=" & sDescription & Chr(0)
   sAttributes = sAttributes & "Database=" + var_bd_reportes & Chr(0)
   strAttributes = strAttributes & "UID=sa" & Chr$(0)
   strAttributes = strAttributes & "PWD=elia" & Chr$(0)
   dl = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYS_DSN, sDriver, sAttributes)
   
   var_cadena_seguridad = ""
   Top = 3000
   Left = 3500
   txt_inicio = Date
   txt_fin = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
    var_swpassword = False
    var_modifica_registro = False
    Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub txt_fin_GotFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = "Presione F5 para seleccionar la fecha"
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fin) Then
         frmcalendario.mes = CDate(Me.txt_fin)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      txt_fin = var_fecha_general
   End If
End Sub

Private Sub txt_fin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   End If
End Sub

Private Sub txt_fin_LostFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = ""
End Sub

Private Sub txt_inicio_GotFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = "Presione F5 para seleccionar la fecha"
End Sub

Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_inicio) Then
         frmcalendario.mes = CDate(Me.txt_inicio)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      txt_inicio = var_fecha_general
   End If
End Sub

Private Sub txt_inicio_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_inicio_LostFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = ""
End Sub


