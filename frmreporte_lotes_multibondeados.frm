VERSION 5.00
Begin VB.Form frmreporte_lotes_multibondeados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes de entradas por lote"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   840
      Left            =   90
      TabIndex        =   5
      Top             =   435
      Width           =   4245
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   315
         Width           =   1140
      End
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   2865
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   300
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   375
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2325
         TabIndex        =   8
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3975
      Picture         =   "frmreporte_lotes_multibondeados.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Caption         =   "D"
      Height          =   315
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Reporte de facturación a detalle"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "L"
      Height          =   315
      Left            =   405
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Reporte de facturación por linea"
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "G"
      Height          =   315
      Left            =   735
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Reporte de facturación general"
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "A"
      Height          =   315
      Left            =   1065
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Reporte de facturación general"
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   0
      TabIndex        =   10
      Top             =   270
      Width           =   4395
   End
End
Attribute VB_Name = "frmreporte_lotes_multibondeados"
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
            rs.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_REPORTE_ENTRADAS_LOTES_MULTIBONDEADOS", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into TB_TEMP_REPORTE_ENTRADAS_LOTES_MULTIBONDEADOS (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
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
            rs.Open "delete from TB_TEMP_REPORTE_ENTRADAS_LOTES_MULTIBONDEADOS where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and dtim_tem_fecha_inicio is null", cnn, adOpenDynamic, adLockOptimistic
            var_cadena = "INSERT INTO TB_TEMP_REPORTE_ENTRADAS_LOTES_MULTIBONDEADOS (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_EMP_EMPRESA_ID, VCHA_EMP_NOMBRE, INTE_EMO_NUMERO, VCHA_EMO_REFERENCIA, VCHA_ART_ARTICULO_ID, VCHA_ART_NOMBRE_ESPAÑOL, FLOA_ENT_CANTIDAD, FLOA_ENT_COSTO, FLOA_ENT_PRECIO, VCHA_LIN_LINEA_ID, VCHA_LIN_NOMBRE, DTIM_EMO_FECHA, VCHA_MOV_NOMBRE)"
            var_cadena = var_cadena + " select " + CStr(var_consecutivo) + ",    " + var_fecha_inicio + "," + var_fecha_fin + ", VCHA_EMP_EMPRESA_ID, VCHA_EMP_NOMBRE, INTE_EMO_NUMERO, VCHA_EMO_REFERENCIA, VCHA_ART_ARTICULO_ID, VCHA_ART_NOMBRE_ESPAÑOL, FLOA_ENT_CANTIDAD, FLOA_ENT_COSTO, FLOA_ENT_PRECIO, VCHA_LIN_LINEA_ID, VCHA_LIN_NOMBRE, DTIM_EMO_FECHA, VCHA_MOV_NOMBRE from VW_ENTRADAS_PT_MULTIBONDEADOS  where dtim_emo_fecha >=" + var_fecha_inicio + " and dtim_emo_fecha <= " + var_fecha_fin + " order by inte_emo_numero, vcha_Art_articulo_id"
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            
            Set reporte = appl.OpenReport(App.Path + "\rep_entradas_lotes_multibondeados_general.rpt")
            reporte.RecordSelectionFormula = "{TB_TEMP_REPORTE_ENTRADAS_LOTES_MULTIBONDEADOS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Entradas concentrado"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            
            
            var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_entradas_lotes_multibondeados_general.rpt")
               reporte.RecordSelectionFormula = "{TB_TEMP_REPORTE_ENTRADAS_LOTES_MULTIBONDEADOS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\reporte_entradas_lote_detalle_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
            End If
            
            
            
            rs.Open "delete from TB_TEMP_REPORTE_ENTRADAS_LOTES_MULTIBONDEADOS where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
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
            var_cadena = var_cadena + " select " + CStr(var_consecutivo) + ",    " + var_fecha_inicio + "," + var_fecha_fin + ", VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, VCHA_ESB_ESTABLECIMIENTO_ID,VCHA_ESB_NOMBRE, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, FLOA_SAL_CANTIDAD, FLOA_SAL_PRECIO*floa_iva_iva, FLOA_SAL_DESCUENTO_1, FLOA_SAL_DESCUENTO_2, VCHA_LIN_LINEA_ID, VCHA_LIN_NOMBRE, VCHA_EMP_EMPRESA_ID, VCHA_EMP_NOMBRE, VCHA_ART_ARTICULO_ID, VCHA_ART_NOMBRE_español, CANTIDAD_KILOS, DTIM_CAR_FECHA, INTE_CAR_PLAZO, VALOR_1, VALOR_2 from VW_FACTURACION_MULTIBONDEADOS where dtim_car_fecha >=" + var_fecha_inicio + " and dtim_car_fecha <= " + var_fecha_fin + " order by inte_Car_numero, vcha_Art_articulo_id"
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            
            
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
            
            
            var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
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
            var_cadena = var_cadena + " select " + CStr(var_consecutivo) + ",    " + var_fecha_inicio + "," + var_fecha_fin + ", VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, VCHA_ESB_ESTABLECIMIENTO_ID,VCHA_ESB_NOMBRE, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, FLOA_SAL_CANTIDAD, FLOA_SAL_PRECIO*floa_iva_iva, FLOA_SAL_DESCUENTO_1, FLOA_SAL_DESCUENTO_2, VCHA_LIN_LINEA_ID, VCHA_LIN_NOMBRE, VCHA_EMP_EMPRESA_ID, VCHA_EMP_NOMBRE, VCHA_ART_ARTICULO_ID, VCHA_ART_NOMBRE_español, CANTIDAD_KILOS, DTIM_CAR_FECHA, INTE_CAR_PLAZO, VALOR_1, VALOR_2 from VW_FACTURACION_MULTIBONDEADOS where dtim_car_fecha >=" + var_fecha_inicio + " and dtim_car_fecha <= " + var_fecha_fin + " order by inte_Car_numero, vcha_Art_articulo_id"
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            
            
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
            
            
            var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
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
            var_cadena = var_cadena + " select " + CStr(var_consecutivo) + ",    " + var_fecha_inicio + "," + var_fecha_fin + ", VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, VCHA_ESB_ESTABLECIMIENTO_ID,VCHA_ESB_NOMBRE, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, FLOA_SAL_CANTIDAD, FLOA_SAL_PRECIO*floa_iva_iva, FLOA_SAL_DESCUENTO_1, FLOA_SAL_DESCUENTO_2, VCHA_LIN_LINEA_ID, VCHA_LIN_NOMBRE, VCHA_EMP_EMPRESA_ID, VCHA_EMP_NOMBRE, VCHA_ART_ARTICULO_ID, VCHA_ART_NOMBRE_español, CANTIDAD_KILOS, DTIM_CAR_FECHA, INTE_CAR_PLAZO, VALOR_1, VALOR_2 from VW_FACTURACION_MULTIBONDEADOS where dtim_car_fecha >=" + var_fecha_inicio + " and dtim_car_fecha <= " + var_fecha_fin + " order by inte_Car_numero, vcha_Art_articulo_id"
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            
            
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
            
            
            var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
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
   sAttributes = sAttributes & "Server=" + parametros(0) & Chr$(0)
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



