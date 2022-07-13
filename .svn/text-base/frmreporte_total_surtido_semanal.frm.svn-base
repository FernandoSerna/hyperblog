VERSION 5.00
Begin VB.Form frmreporte_total_surtido_semanal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de total surtido semanal"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   840
      Left            =   135
      TabIndex        =   3
      Top             =   480
      Width           =   4245
      Begin VB.TextBox txt_año 
         Height          =   315
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   300
         Width           =   885
      End
      Begin VB.TextBox txt_semana 
         Height          =   315
         Left            =   2700
         TabIndex        =   4
         Top             =   300
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
         Height          =   195
         Left            =   465
         TabIndex        =   7
         Top             =   360
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Semana:"
         Height          =   195
         Left            =   2025
         TabIndex        =   6
         Top             =   360
         Width           =   630
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4050
      Picture         =   "frmreporte_total_surtido_semanal.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmreporte_total_surtido_semanal.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   15
      TabIndex        =   0
      Top             =   315
      Width           =   4485
   End
End
Attribute VB_Name = "frmreporte_total_surtido_semanal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
   cnn.BeginTrans
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   rsaux.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_REPORTE_TOTAL_SURTIDO_ALMACEN_SEMANAL", cnn, adOpenDynamic, adLockOptimistic
   If Not rsaux.EOF Then
      var_consecutivo = IIf(IsNull(rsaux!NUMERO), 0, rsaux!NUMERO)
   Else
      var_consecutivo = 0
   End If
   rsaux.Close
   var_consecutivo = var_consecutivo + 1
   rsaux.Open "insert into TB_TEMP_REPORTE_TOTAL_SURTIDO_ALMACEN_SEMANAL (INTE_TEM_CONSECUTIVO) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
   cnn.CommitTrans
   cnn.CommandTimeout = 360
   rs.Open "exec SP_REPORTE_TOTAL_SURTIDO_ALMACEN_SEMANAL " + CStr(var_consecutivo) + "," + Me.txt_año + ", " + Me.txt_semana, cnn, adOpenDynamic, adLockOptimistic
   'Set reporte = appl.OpenReport(App.Path + "\rep_total_surtido_almacen_semanal_transportes.rpt")
   'reporte.RecordSelectionFormula = "{TB_TEMP_REPORTE_TOTAL_SURTIDO_ALMACEN_SEMANAL.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
   'frmvistasprevias.cr.ReportSource = reporte
   'For ntablas = 1 To reporte.Database.Tables.Count
   '    reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
   'Next ntablas
   'frmvistasprevias.cr.ViewReport
   'frmvistasprevias.Caption = "Reporte Consecutivo de Movimientos por Agente (cargos)"
   'frmvistasprevias.Show 1
   'Set reporte = Nothing
   'var_si = MsgBox("¿Desea importar el reporte?", vbYesNo, "ATENCION")
   'If var_si = 6 Then
      Set reporte = appl.OpenReport(App.Path + "\rep_total_surtido_almacen_general_transportes.rpt")
      reporte.RecordSelectionFormula = "{VW_TEMP_REPORTE_TOTAL_SURTIDO_ALMACEN_SEMANAL.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      reporte.ExportOptions.FormatType = crEFTExcel80
      reporte.ExportOptions.DestinationType = crEDTDiskFile
      archivo = "c:\reportessid\reporte_total_surtido_semanal_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
      reporte.ExportOptions.DiskFileName = archivo
      reporte.Export False
      Set reporte = Nothing
      MsgBox "Se a terminado de guardar el archivo " + archivo
   'End If
   rs.Open "delete from TB_TEMP_REPORTE_TOTAL_SURTIDO_ALMACEN_SEMANAL where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
   rs.Open "delete from TB_TEMP_REPORTE_TOTAL_SURTIDO_ALMACEN_SEMANAL_ORDENES where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 3000
   Left = 3500
   Me.txt_año = CStr(Year(Date))
   var_dia = CStr(Day(Date))
   var_mes = CStr(Month(Date))
   var_año = CStr(Year(Date))
   If Len(Trim(var_dia)) = 1 Then
      var_dia = "0" + var_dia
   End If
   If Len(Trim(var_mes)) = 1 Then
      var_mes = "0" + var_mes
   End If
   var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
   
   rs.Open "select * from TB_SEMANAS_LABORABLES where date_sem_fecha_inicio <= " + var_fecha + " and date_sem_fecha_fin >= " + var_fecha, cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      Me.txt_semana = IIf(IsNull(rs!numb_sem_semana_id), "", rs!numb_sem_semana_id)
   End If
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
    activa_forma (var_activa_forma_packing_list)
End Sub

