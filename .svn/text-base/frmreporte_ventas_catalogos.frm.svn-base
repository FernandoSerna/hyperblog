VERSION 5.00
Begin VB.Form frmreporte_ventas_catalogos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Ventas Netas por Catálogo"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4035
      Picture         =   "frmreporte_ventas_catalogos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   45
      Picture         =   "frmreporte_ventas_catalogos.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   5
      Top             =   345
      Width           =   4485
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   720
      Left            =   60
      TabIndex        =   0
      Top             =   435
      Width           =   4335
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   2700
         TabIndex        =   2
         Top             =   255
         Width           =   1095
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   690
         TabIndex        =   1
         Top             =   255
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2385
         TabIndex        =   4
         Top             =   315
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   315
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmreporte_ventas_catalogos"
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
             rs.Open "select max(inte_tem_consecutivo) from TB_TEMP_REPORTE_VENTAS_CATALOGOS", cnn, adOpenDynamic, adLockOptimistic
             If Not rs.EOF Then
                var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
             Else
                var_consecutivo = 0
             End If
             var_consecutivo = var_consecutivo + 1
             rs.Close
             rs.Open "Insert into TB_TEMP_REPORTE_VENTAS_CATALOGOS (INTE_tem_CONSECUTIVO) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
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
             
             var_cadena = " INSERT INTO TB_TEMP_REPORTE_VENTAS_CATALOGOS (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_EMP_EMPRESA_ID, VCHA_ALM_ALMACEN_ID, VCHA_CAN_NOMBRE, VCHA_ART_CATALOGO_VIGENTE, FLOA_TEM_PIEZAS, FLOA_TEM_IMPORTE) "
             var_cadena = var_cadena + " SELECT " + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin + " - 0.000001, tb_salidas.vcha_emp_empresa_id,dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID, vcha_can_nombre,tb_articulos.vcha_art_catalogo_vigente, SUM(dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD) AS piezas, SUM((isnull(dbo.TB_SALIDAS.FLOA_SAL_PRECIO,0) * isnull(dbo.TB_SALIDAS.FLOA_SAL_CANTIDAD,0) * (1 - isnull(dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_1,0) / 100)) * (1 - isnull(dbo.TB_SALIDAS.FLOA_SAL_DESCUENTO_2,0) / 100)) AS total "
             var_cadena = var_cadena + " FROM tb_agentes, tb_canalesventas, tb_encabezado_cartera, TB_SALIDAS, tB_ARTICULOS wHERE dbo.TB_SALIDAS.VCHA_car_documento = tb_encabezado_cartera.vcha_car_tipo_documento and  dbo.TB_SALIDAS.inte_car_numero = tb_encabezado_cartera.inte_car_numero and  (dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID = '8') AND (dbo.TB_SALIDAS.DTIM_SAL_FECHA BETWEEN " + var_fecha_inicio + " AND  " + var_fecha_fin + "-.00001) AND"
             var_cadena = var_cadena + " (dbo.TB_SALIDAS.VCHA_EMP_EMPRESA_ID='02' or tb_salidas.vcha_emp_empresa_id='03') AND (dbo.TB_SALIDAS.VCHA_car_documento = 'FA') and  tb_encabezado_cartera.vcha_age_agente_id = tb_agentes.vcha_age_agente_id and tb_agentes.vcha_can_canal_venta_id=tb_canalesventas.vcha_can_canal_venta_id and  tb_salidas.vcha_art_articulo_id=tb_articulos.vcha_art_articulo_id grOUP BY dbo.tb_articulos.VCHA_ART_catalogo_vigente,dbo.TB_SALIDAS.VCHA_ALM_ALMACEN_ID,vcha_can_nombre,tb_salidas.vcha_emp_empresa_id"
             var_cadena = var_cadena + " ORDER BY vcha_can_nombre,dbo.TB_articulos.vcha_art_catalogo_vigente"
             
             Text1 = var_cadena
             cnn.CommandTimeout = 6000
             rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
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
             var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
             
         Set reporte = appl.OpenReport(App.Path + "\rep_ventas_catalogos.rpt")
         reporte.RecordSelectionFormula = "{VW_REPORTE_VENTAS_CATALOGOS.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo)
         frmvistasprevias.cr.ReportSource = reporte
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Reporte de ventas"
         frmvistasprevias.Show 1
         Set reporte = Nothing
         var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            Set reporte = appl.OpenReport(App.Path + "\rep_ventas_catalogos.rpt")
            reporte.RecordSelectionFormula = "{VW_REPORTE_VENTAS_CATALOGOS.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\reporte_ventas_catalogos" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
            MsgBox "Se a terminado de guardar el archivo " + archivo
         End If
         
         rs.Open "delete from TB_TEMP_REPORTE_VENTAS_CATALOGOS where INTE_tem_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
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
   Call activa_forma(var_activa_forma_reporte_valuacion_devoluciones)
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

