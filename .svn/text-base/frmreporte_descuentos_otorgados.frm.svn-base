VERSION 5.00
Begin VB.Form frmreporte_descuentos_otorgados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Descuentos Otorgados"
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_periodo 
      Height          =   675
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   4770
      Begin VB.CommandButton cmd_cambiar_periodo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4350
         Picture         =   "frmreporte_descuentos_otorgados.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Aplicar Pagos Alt + A"
         Top             =   240
         Width           =   330
      End
      Begin VB.ListBox lst_años 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frmreporte_descuentos_otorgados.frx":014A
         Left            =   3405
         List            =   "frmreporte_descuentos_otorgados.frx":018D
         TabIndex        =   2
         Top             =   247
         Width           =   900
      End
      Begin VB.ComboBox cmb_meses 
         Height          =   315
         ItemData        =   "frmreporte_descuentos_otorgados.frx":020F
         Left            =   630
         List            =   "frmreporte_descuentos_otorgados.frx":0237
         TabIndex        =   1
         Top             =   240
         Width           =   2280
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
         Height          =   195
         Left            =   3015
         TabIndex        =   5
         Top             =   300
         Width           =   330
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Mes:"
         Height          =   195
         Left            =   195
         TabIndex        =   4
         Top             =   300
         Width           =   345
      End
   End
End
Attribute VB_Name = "frmreporte_descuentos_otorgados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim fecha_inicio As Date
Dim fecha_fin As Date
Private Sub cmd_cambiar_periodo_Click()
   Dim fecha_anterior As Date
   Dim dia_anterior As Integer
   Dim mes_anterior As Integer
   Dim año_anterior As Integer
   Dim dia As Integer
   Dim mes As Integer
   Dim año As Integer
   Dim periodo As String
   
   If cmb_meses = "Enero" Then
      mes_anterior = 1
   End If
   If cmb_meses = "Febrero" Then
      mes_anterior = 2
   End If
   If cmb_meses = "Marzo" Then
      mes_anterior = 3
   End If
   If cmb_meses = "Abril" Then
      mes_anterior = 4
   End If
   If cmb_meses = "Mayo" Then
      mes_anterior = 5
   End If
   If cmb_meses = "Junio" Then
      mes_anterior = 6
   End If
   If cmb_meses = "Julio" Then
      mes_anterior = 7
   End If
   If cmb_meses = "Agosto" Then
      mes_anterior = 8
   End If
   If cmb_meses = "Septiembre" Then
      mes_anterior = 9
   End If
   If cmb_meses = "Octubre" Then
      mes_anterior = 10
   End If
   If cmb_meses = "Noviembre" Then
      mes_anterior = 11
   End If
   If cmb_meses = "Diciembre" Then
      mes_anterior = 12
   End If
   año_anterior = lst_años
   If mes_anterior = 1 Or mes_anterior = 3 Or mes_anterior = 5 Or mes_anterior = 7 Or mes_anterior = 8 Or mes_anterior = 10 Or mes_anterior = 12 Then
      fecha_inicio = CDate("1/" + Str(mes_anterior) + "/" + Str(año_anterior))
      fecha_fin = CDate("31/" + Str(mes_anterior) + "/" + Str(año_anterior))
      dia_anterior = 31
   End If
   If mes_anterior = 4 Or mes_anterior = 6 Or mes_anterior = 9 Or mes_anterior = 11 Then
      fecha_inicio = CDate("1/" + Str(mes_anterior) + "/" + Str(año_anterior))
      fecha_fin = CDate("30/" + Str(mes_anterior) + "/" + Str(año_anterior))
      dia_anterior = 30
   End If
   
   If mes_anterior = 2 Then
      fecha_inicio = CDate("1/" + Str(mes_anterior) + "/" + Str(año_anterior))
      If año_anterior = 2004 Or año_anterior = 2008 Or año_anterior = 2012 Or año_anterior = 2016 Or año_anterior = 2020 Or año_anterior = 2024 Then
         fecha_fin = CDate("29/" + Str(mes_anterior) + "/" + Str(año_anterior))
         dia_anterior = 29
      Else
         fecha_fin = CDate("28/" + Str(mes_anterior) + "/" + Str(año_anterior))
         dia_anterior = 28
      End If
   End If
   
   mes = mes_anterior
   año = año_anterior
  
   If mes = 1 Then
      periodo = "Enero"
   End If
   If mes = 2 Then
      periodo = "Febrero"
   End If
   If mes = 3 Then
      periodo = "Marzo"
   End If
   If mes = 4 Then
      periodo = "Abril"
   End If
   If mes = 5 Then
      periodo = "Mayo"
   End If
   If mes = 6 Then
      periodo = "Junio"
   End If
   If mes = 7 Then
      periodo = "Julio"
   End If
   If mes = 8 Then
      periodo = "Agosto"
   End If
   If mes = 9 Then
      periodo = "Septiembre"
   End If
   If mes = 10 Then
      periodo = "Octubre"
   End If
   If mes = 11 Then
      periodo = "Noviembre"
   End If
   If mes = 12 Then
      periodo = "Diciembre"
   End If
   txt_periodo = "1 de " + periodo + " al " + Str(dia_anterior) + " de " + periodo + " del " + Str(año)
   
   
   Set reporte = appl.OpenReport(App.Path + "\rep_descuento_otorgados.rpt")
   reporte.RecordSelectionFormula = "{vw_descuentos_volumen.DTIM_dvo_PERIODO_INICIO} = date('" + CStr(fecha_inicio) + "') and {vw_descuentos_volumen.DTIM_dvo_PERIODO_fin} = date('" + CStr(fecha_fin + 1) + "')  and {vw_descuentos_volumen.vcha_emp_empresa_id} = '" + var_empresa + "'"
   frmvistasprevias.cr.ReportSource = reporte
   For ntablas = 1 To reporte.Database.Tables.Count
       reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
   Next ntablas
   frmvistasprevias.cr.ViewReport
   frmvistasprevias.Caption = "Reporte de Movimientos"
   frmvistasprevias.Show 1
   Set reporte = Nothing
   var_si = MsgBox("¿Desea importar el reporte?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      Set reporte = appl.OpenReport(App.Path + "\rep_descuento_otorgados.rpt")
      reporte.RecordSelectionFormula = "{vw_descuentos_volumen.DTIM_dvo_PERIODO_INICIO} = date('" + CStr(fecha_inicio) + "') and {vw_descuentos_volumen.DTIM_dvo_PERIODO_fin} = date('" + CStr(fecha_fin + 1) + "')  and {vw_descuentos_volumen.vcha_emp_empresa_id} = '" + var_empresa + "'"
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      reporte.ExportOptions.FormatType = crEFTExcel80
      reporte.ExportOptions.DestinationType = crEDTDiskFile
      archivo = "c:\reportessid\REPORTE_DESCUENTOS_OTORGADOS" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
      reporte.ExportOptions.DiskFileName = archivo
      reporte.Export False
      Set reporte = Nothing
      MsgBox "Se a terminado de guardar el archivo " + archivo
   End If
End Sub

Private Sub Form_Load()
   Dim mes As Integer
   Dim año As Integer
   var_cadena_seguridad = ""
   Top = 3000
   Left = 3000
   mes = Month(Date)
   año = Year(Date)
   lst_años = año
   If mes = 1 Then
      cmb_meses = "Enero"
   End If
   If mes = 2 Then
      cmb_meses = "Febrero"
   End If
   If mes = 3 Then
      cmb_meses = "Marzo"
   End If
   If mes = 4 Then
      cmb_meses = "Abril"
   End If
   If mes = 5 Then
      cmb_meses = "Mayo"
   End If
   If mes = 6 Then
      cmb_meses = "Junio"
   End If
   If mes = 7 Then
      cmb_meses = "Julio"
   End If
   If mes = 8 Then
      cmb_meses = "Agosto"
   End If
   If mes = 9 Then
      cmb_meses = "Septiembre"
   End If
   If mes = 10 Then
      cmb_meses = "Octubre"
   End If
   If mes = 11 Then
      cmb_meses = "Noviembre"
   End If
   If mes = 12 Then
      cmb_meses = "Diciembre"
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_despliega_menu = True Then
      var_swpassword = False
      var_modifica_registro = False
   End If
   Call activa_forma(var_activa_forma_clasificacion_clientes)
End Sub

