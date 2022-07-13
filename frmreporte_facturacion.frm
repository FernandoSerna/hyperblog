VERSION 5.00
Begin VB.Form frmreporte_facturacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de facturación"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   720
      Left            =   135
      TabIndex        =   3
      Top             =   405
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
      Left            =   90
      TabIndex        =   2
      Top             =   315
      Width           =   4440
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmreporte_facturacion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4110
      Picture         =   "frmreporte_facturacion.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "frmreporte_facturacion"
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
   Dim año As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   
   'On Error GoTo salir:
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) from tb_temp_reporte_facturacion_2", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "Insert into tb_temp_reporte_facturacion_2 (INTE_tem_CONSECUTIVO) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
             
            var_fecha_fin_1 = CDate(txt_fin)
             
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
            
            var_cadena = "INSERT INTO TB_TEMP_REPORTE_FACTURACION_2 (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, DTIM_CAR_FECHA, FLOA_CAR_IMPORTE_NETO, floa_Car_tipo_cambio,CHAR_CAR_ESTATUS, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, VCHA_AGE_AGENTE_ID, VCHA_AGE_NOMBRE, VCHA_EMP_NOMBRE, VCHA_UOR_NOMBRE)"
            var_cadena = var_cadena + " SELECT TOP 100 PERCENT " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + ",dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_UOR_UNIDAD_ID, dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_DOCUMENTO, dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID, dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO, dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA, dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_IMPORTE_NETO, dbo.TB_ENCABEZADO_CARTERA.FLOA_CAR_TIPO_CAMBIO, ISNULL(dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS, 'I') AS CHAR_CAR_ESTATUS, dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_ENCABEZADO_CARTERA.VCHA_AGE_AGENTE_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_EMPRESAS.VCHA_EMP_NOMBRE , dbo.TB_UNIDADESORGANIZACIONALES.VCHA_UOR_NOMBRE FROM dbo.TB_ENCABEZADO_CARTERA LEFT OUTER JOIN dbo.TB_UNIDADESORGANIZACIONALES ON dbo.TB_ENCABEZADO_CARTERA.VCHA_UOR_UNIDAD_ID = dbo.TB_UNIDADESORGANIZACIONALES.VCHA_UOR_UNIDAD_ID LEFT OUTER JOIN "
            var_cadena = var_cadena + " dbo.TB_EMPRESAS ON dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = dbo.TB_EMPRESAS.VCHA_EMP_EMPRESA_ID LEFT OUTER JOIN dbo.TB_AGENTES ON dbo.TB_ENCABEZADO_CARTERA.VCHA_AGE_AGENTE_ID = dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID LEFT OUTER JOIN dbo.TB_CLIENTES ON dbo.TB_ENCABEZADO_CARTERA.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID WHERE (dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA >= " + var_fecha_inicio + ") AND (dbo.TB_ENCABEZADO_CARTERA.DTIM_CAR_FECHA <= " + var_fecha_fin + " + 1 -.0000001) AND (dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID = '15') AND (dbo.TB_ENCABEZADO_CARTERA.VCHA_CAR_TIPO_DOCUMENTO = 'FA') ORDER BY dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO "
                         
            If rsaux.State = 1 Then
               rsaux.Close
            End If
            rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            rs.Open "delete from TB_TEMP_REPORTE_FACTURACION_2 where INTE_Tem_CONSECUTIVO = " + CStr(var_consecutivo) + " AND VCHA_eMP_eMPRESA_ID IS NULL", cnn, adOpenDynamic, adLockOptimistic
            
            Set reporte = appl.OpenReport(App.Path + "\rep_facturacion.rpt")
            reporte.RecordSelectionFormula = "{vw_temp_reporte_facturacion_2.INTE_tem_CONSECUTIVO} = '" + CStr(var_consecutivo) + "'"
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Valuación de Facturas"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_facturacion.rpt")
               reporte.RecordSelectionFormula = "{vw_temp_reporte_facturacion_2.INTE_tem_CONSECUTIVO} = '" + CStr(var_consecutivo) + "'"
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\reporte_valuacion_facturas" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
            rs.Open "delete from TB_TEMP_REPORTE_FACTURACION_2 where INTE_Tem_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
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
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(txt_fin) Then
         frmcalendario.mes.Value = CDate(txt_fin)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_fin = var_fecha_general
   End If
End Sub

Private Sub txt_fin_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(txt_inicio) Then
         frmcalendario.mes.Value = CDate(txt_inicio)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_inicio = var_fecha_general
   End If
End Sub

Private Sub txt_inicio_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub


