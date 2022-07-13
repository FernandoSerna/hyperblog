VERSION 5.00
Begin VB.Form frmreporte_existencias_cajas_pedidos_tiendas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Existencias de cajas"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   4515
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4050
      Picture         =   "frmreporte_existencias_cajas_pedidos_tiendas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmreporte_existencias_cajas_pedidos_tiendas.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   15
      TabIndex        =   5
      Top             =   330
      Width           =   4485
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   720
      Left            =   75
      TabIndex        =   0
      Top             =   420
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
Attribute VB_Name = "frmreporte_existencias_cajas_pedidos_tiendas"
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
             cnn.BeginTrans
             rs.Open "select max(inte_tem_consecutivo) from TB_TEMP_EXISTENCIAS_CAJAS_PEDIDOS_TIENDAS", cnn, adOpenDynamic, adLockOptimistic
             If Not rs.EOF Then
                var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
             Else
                var_consecutivo = 0
             End If
             var_consecutivo = var_consecutivo + 1
             rs.Close
             rs.Open "Insert into TB_TEMP_EXISTENCIAS_CAJAS_PEDIDOS_TIENDAS (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
             cnn.CommitTrans
             
             If CDate(Me.txt_inicio) <= CDate("2007-12-31") Then
                If CDate(Me.txt_fin) <= CDate("2007-12-31") Then
                   rs.Open "insert into TB_TEMP_EXISTENCIAS_CAJAS_PEDIDOS_TIENDAS (inte_tem_consecutivo, dtim_tem_fecha_inicio, dtim_tem_fecha_fin, inte_emb_embarque, INTE_PAQ_CAJA, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, FLOA_PAQ_CANTIDAD, VCHA_PAQ_NOMBRE, DTIM_ORS_FECHA_CARGA, vcha_age_Agente_id, vcha_age_nombre) select distinct " + CStr(var_consecutivo) + ", " + var_fecha_inicio + "," + var_fecha_fin + "-.000001, INTE_EMB_EMBARQUE, INTE_PAQ_CAJA, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, FLOA_PAQ_CANTIDAD, VCHA_PAQ_NOMBRE, DTIM_ORS_FECHA_CARGA, vcha_age_agente_id, vcha_age_nombre from VW_EXISTENCIAS_CAJAS_PEDIDOS_TIENDAS_2007 WHERE DTIM_ORS_FECHA_CARGA >= " + var_fecha_inicio + " AND DTIM_ORS_FECHA_CARGA <= " + var_fecha_fin, cnn, adOpenDynamic, adLockOptimistic
                Else
                   MsgBox "La fecha final debe de ser menor o igual al 31 de Diciembre del 2007", vbOKOnly, "ATENCION"
                End If
             Else
                rs.Open "insert into TB_TEMP_EXISTENCIAS_CAJAS_PEDIDOS_TIENDAS (inte_tem_consecutivo, dtim_tem_fecha_inicio, dtim_tem_fecha_fin, inte_emb_embarque, INTE_PAQ_CAJA, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, FLOA_PAQ_CANTIDAD, VCHA_PAQ_NOMBRE, DTIM_ORS_FECHA_CARGA, vcha_age_Agente_id, vcha_age_nombre) select distinct " + CStr(var_consecutivo) + ", " + var_fecha_inicio + "," + var_fecha_fin + "-.000001, INTE_EMB_EMBARQUE, INTE_PAQ_CAJA, VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE, FLOA_PAQ_CANTIDAD, VCHA_PAQ_NOMBRE, DTIM_ORS_FECHA_CARGA, vcha_age_agente_id, vcha_age_nombre from VW_EXISTENCIAS_CAJAS_PEDIDOS_TIENDAS WHERE DTIM_ORS_FECHA_CARGA >= " + var_fecha_inicio + " AND DTIM_ORS_FECHA_CARGA <= " + var_fecha_fin, cnn, adOpenDynamic, adLockOptimistic
             End If
             
             Set reporte = appl.OpenReport(App.Path + "\rep_existencias_pedidos_cajas_tiendas.rpt")
             reporte.RecordSelectionFormula = "{TB_TEMP_EXISTENCIAS_CAJAS_PEDIDOS_TIENDAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {TB_TEMP_EXISTENCIAS_CAJAS_PEDIDOS_TIENDAS.inte_emb_embarque} >0"
             frmvistasprevias.cr.ReportSource = reporte
             For ntablas = 1 To reporte.Database.Tables.Count
                 reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
             Next ntablas
             frmvistasprevias.cr.ViewReport
             frmvistasprevias.Caption = "Reporte de existencias en cajas"
             frmvistasprevias.Show 1
             Set reporte = Nothing
         
             var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
             If var_si = 6 Then
                Set reporte = appl.OpenReport(App.Path + "\rep_existencias_pedidos_cajas_tiendas.rpt")
                reporte.RecordSelectionFormula = "{TB_TEMP_EXISTENCIAS_CAJAS_PEDIDOS_TIENDAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {TB_TEMP_EXISTENCIAS_CAJAS_PEDIDOS_TIENDAS.inte_emb_embarque} > 0"
                For ntablas = 1 To reporte.Database.Tables.Count
                    reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                Next ntablas
                reporte.ExportOptions.FormatType = crEFTExcel80
                reporte.ExportOptions.DestinationType = crEDTDiskFile
                archivo = "c:\reportessid\existencias_cajas_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                reporte.ExportOptions.DiskFileName = archivo
                reporte.Export False
                              Set reporte = Nothing
                              MsgBox "Se a terminado de guardar el archivo " + archivo
                           End If
      
         
         rs.Open "delete from TB_TEMP_REPORTE_DEVOLUCIONES where INTE_Tem_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
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
   Call activa_forma(var_activa_forma_existencias_generales)
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


