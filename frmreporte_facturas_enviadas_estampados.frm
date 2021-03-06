VERSION 5.00
Begin VB.Form frmreporte_facturas_enviadas_estampados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de facturas enviadas de estampados"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   720
      Left            =   75
      TabIndex        =   3
      Top             =   420
      Width           =   4335
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   690
         TabIndex        =   5
         Top             =   255
         Width           =   1080
      End
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   2700
         TabIndex        =   4
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   315
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2385
         TabIndex        =   6
         Top             =   315
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   15
      TabIndex        =   2
      Top             =   330
      Width           =   4485
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmreporte_facturas_enviadas_estampados.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4050
      Picture         =   "frmreporte_facturas_enviadas_estampados.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "frmreporte_facturas_enviadas_estampados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
   Dim var_consecutivo As Double
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim a?o As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   
   'On Error GoTo salir:
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
             cnn.BeginTrans
             rs.Open "select max(inte_tem_consecutivo) from TB_TEMP_REPORTE_FACTURAS_ENVIADAS_ESTAMPADOS", cnn, adOpenDynamic, adLockOptimistic
             If Not rs.EOF Then
                var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
             Else
                var_consecutivo = 0
             End If
             var_consecutivo = var_consecutivo + 1
             rs.Close
             rs.Open "Insert into TB_TEMP_REPORTE_FACTURAS_ENVIADAS_ESTAMPADOS (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
             cnn.CommitTrans
             var_fecha_fin_1 = CDate(txt_fin) + 1
             var_dia = CStr(Day(CDate(txt_inicio)))
             var_mes = CStr(Month(CDate(txt_inicio)))
             var_a?o = CStr(Year(CDate(txt_inicio)))
             If Len(Trim(var_dia)) = 1 Then
                var_dia = "0" + var_dia
             End If
             If Len(Trim(var_mes)) = 1 Then
                var_mes = "0" + var_mes
             End If
             var_fecha_inicio = "{d '" + var_a?o + "-" + var_mes + "-" + var_dia + "'}"
             
             
             var_dia = CStr(Day(var_fecha_fin_1))
             var_mes = CStr(Month(var_fecha_fin_1))
             var_a?o = CStr(Year(var_fecha_fin_1))
             If Len(Trim(var_dia)) = 1 Then
                var_dia = "0" + var_dia
             End If
             If Len(Trim(var_mes)) = 1 Then
                var_mes = "0" + var_mes
             End If
             var_fecha_fin = "{d '" + var_a?o + "-" + var_mes + "-" + var_dia + "'}"
             
             var_cadena = "insert into admcdindustrial.sid.dbo.TB_TEMP_REPORTE_FACTURAS_ENVIADAS_ESTAMPADOS "
             var_cadena = var_cadena + "SELECT     " + CStr(var_consecutivo) + ", " + var_fecha_inicio + " AS Expr2, " + var_fecha_fin + " AS Expr3, dbo.TB_SALIDAS.VCHA_SER_SERIE_ID + CAST(dbo.TB_SALIDAS.INTE_CAR_NUMERO AS varchar(50)) AS FACTURA, dbo.TB_SALIDAS.INTE_CAR_NUMERO, dbo.TB_SALIDAS.VCHA_SER_SERIE_ID, dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA, dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID, dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD, 0 AS cantidad, dbo.TB_Articulos.VCHA_ART_NOMBRE_ESPA?OL FROM dbo.TB_ENCABEZADO_CARTERA INNER JOIN dbo.TB_SALIDAS ON dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_UOR_UNIDAD_ID = dbo.TB_SALIDAS.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID = dbo.TB_SALIDAS.VCHA_SER_SERIE_ID AND dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO = dbo.TB_SALIDAS.VCHA_CAR_DOCUMENTO AND dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO = dbo.TB_SALIDAS.INTE_CAR_NUMERO INNER JOIN "
             var_cadena = var_cadena + " dbo.TB_ARTICULOS ON dbo.TB_SALIDAS.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID WHERE (dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA >= " + var_fecha_inicio + ") AND (dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA < " + var_fecha_fin + ") AND (dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS <> 'C') AND (dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = '15') AND (dbo.TB_ENCABEZADO_CARTERA.VCHA_ESB_ESTABLECIMIENTO_ID = 'E000001752'  or dbo.TB_ENCABEZADO_CARTERA.VCHA_ESB_ESTABLECIMIENTO_ID = 'E000001754' ) OR (dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA >= " + var_fecha_inicio + ") AND (dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA < " + var_fecha_fin + ") AND (dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS IS NULL) AND (dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = '15') AND (dbo.TB_ENCABEZADO_CARTERA.VCHA_ESB_ESTABLECIMIENTO_ID = 'E000001752' or dbo.TB_ENCABEZADO_CARTERA.VCHA_ESB_ESTABLECIMIENTO_ID = 'E000001754') order by dtim_Car_Fecha"
             cnn_sqlquezada2.CommandTimeout = 360
             rs.Open var_cadena, cnn_sqlquezada2, adOpenDynamic, adLockOptimistic
             rs.Open "select * from admcdindustrial.sid.dbo.TB_TEMP_REPORTE_FACTURAS_ENVIADAS_ESTAMPADOS where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and dtim_tem_fecha_inicio is not null", cnn, adOpenDynamic, adLockOptimistic
             While Not rs.EOF
                   var_longitud_serie = Len(rs!vcha_ser_Serie_id)
                   rsaux.Open "select * from tb_archivo_comparacion where vcha_emp_empresa_id = '06' and vcha_mov_movimiento_id = 'EI' and inte_com_numero = " + CStr(rs!inte_car_numero) + " and substring(vcha_com_referencia,1," + CStr(var_longitud_serie) + ") = '" + rs!vcha_ser_Serie_id + "' and vcha_Art_articulo_id = '" + rs!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                   If Not rsaux.EOF Then
                      rsaux1.Open "update TB_TEMP_REPORTE_FACTURAS_ENVIADAS_ESTAMPADOS set floa_tem_Cantidad_recibida = " + CStr(rsaux!floa_com_cantidad_recibida) + " where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_tem_Factura = '" + rs!vcha_tem_factura + "' and vcha_art_Articulo_id = '" + rs!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                   End If
                   rsaux.Close
                   rs.MoveNext
             Wend
             rs.Close
             rs.Open "DELETE FROM TB_TEMP_REPORTE_FACTURAS_ENVIADAS_ESTAMPADOS WHERE DTIM_TEM_FECHA_INICIO IS NULL AND INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
             Set reporte = appl.OpenReport(App.Path + "\REP_FACTURAS_ENVIADAS_ESTAMPADOS_Q0Z.rpt")
             reporte.RecordSelectionFormula = "{VW_REPORTE_FACTURAS_ENVIADAS_ESTAMPADOS_Q0Z.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
             
             frmvistasprevias.cr.ReportSource = reporte
             For ntablas = 1 To reporte.Database.Tables.Count
                 reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
             Next ntablas
             frmvistasprevias.cr.ViewReport
             frmvistasprevias.Caption = "Reporte de Cedula de Saldos"
             frmvistasprevias.Show 1
             Set reporte = Nothing
             var_si = MsgBox("?Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
             If var_si = 6 Then
                Set reporte = appl.OpenReport(App.Path + "\REP_FACTURAS_ENVIADAS_ESTAMPADOS_Q0Z.rpt")
                reporte.RecordSelectionFormula = "{VW_REPORTE_FACTURAS_ENVIADAS_ESTAMPADOS_Q0Z.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                For ntablas = 1 To reporte.Database.Tables.Count
                    reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                Next ntablas
                reporte.ExportOptions.FormatType = crEFTExcel80
                reporte.ExportOptions.DestinationType = crEDTDiskFile
                archivo = "c:\reportessid\facturas_enviadas_estampados_q0z_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                reporte.ExportOptions.DiskFileName = archivo
                reporte.Export False
                Set reporte = Nothing
                MsgBox "Se a terminado de guardar el archivo " + archivo
             End If
             rs.Open "delete from TB_TEMP_REPORTE_FACTURAS_ENVIADAS_ESTAMPADOS where INTE_Tem_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "La fecha de inicio debe de ser mayor", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha Final Incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de Inicio Incorrecta", vbOKOnly, "ATENCION"
   End If
   Exit Sub
salir:
   MsgBox "A surgido un error al generar el reporte", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 3000
   Left = 3500
   txt_inicio = Date
   txt_fin = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
   var_swpassword = False
   var_modifica_registro = False
   Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fin) Then
         frmcalendario.mes.Value = CDate(Me.txt_fin)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_fin = var_fecha_general
   End If
End Sub

Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_inicio) Then
         frmcalendario.mes.Value = CDate(Me.txt_inicio)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_inicio = var_fecha_general
   End If
End Sub



