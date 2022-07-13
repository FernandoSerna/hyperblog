VERSION 5.00
Begin VB.Form frmreporte_ventas_netas_cantia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de ventas netas de cantia"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   " Tipo de reporte "
      Height          =   1665
      Left            =   60
      TabIndex        =   8
      Top             =   390
      Width           =   4335
      Begin VB.OptionButton opt_articulo 
         Caption         =   "Artículo"
         Height          =   300
         Left            =   210
         TabIndex        =   12
         Top             =   1290
         Width           =   1905
      End
      Begin VB.OptionButton opt_canal 
         Caption         =   "Canal de venta"
         Height          =   300
         Left            =   210
         TabIndex        =   11
         Top             =   955
         Width           =   1905
      End
      Begin VB.OptionButton opt_agente 
         Caption         =   "Agente"
         Height          =   300
         Left            =   210
         TabIndex        =   10
         Top             =   620
         Width           =   1905
      End
      Begin VB.OptionButton opt_cliente 
         Caption         =   "Cliente"
         Height          =   300
         Left            =   210
         TabIndex        =   9
         Top             =   285
         Width           =   1905
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   720
      Left            =   60
      TabIndex        =   3
      Top             =   2130
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
      Left            =   0
      TabIndex        =   2
      Top             =   345
      Width           =   4485
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   30
      Picture         =   "frmreporte_ventas_netas_cantia.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4035
      Picture         =   "frmreporte_ventas_netas_cantia.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
End
Attribute VB_Name = "frmreporte_ventas_netas_cantia"
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
            rs.Open "select max(inte_tem_consecutivo) from TB_TEMP_REPORTE_VENTAS_NETAS_CANTIA", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "Insert into TB_TEMP_REPORTE_VENTAS_NETAS_CANTIA (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
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
            
            
            var_dia = CStr(Day(CDate(Me.txt_fin) + 1))
            var_mes = CStr(Month(CDate(Me.txt_fin) + 1))
            var_año = CStr(Year(CDate(Me.txt_fin) + 1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
             
            rs.Open "exec SP_REPORTE_VENTAS_NETAS " + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin, cnn, adOpenDynamic, adLockOptimistic
            'rs.Open "delete from TB_TEMP_REPORTE_VENTAS_NETAS_CANTIA where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn_sqlquezada2, adOpenDynamic, adLockOptimistic
            'rs.Open "exec SP_REPORTE_VENTAS_NETAS " + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin, cnn_sqlquezada2, adOpenDynamic, adLockOptimistic
            'rs.Open "select * from TB_TEMP_REPORTE_VENTAS_NETAS_CANTIA where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn_sqlquezada2, adOpenDynamic, adLockOptimistic
            'While Not rs.EOF
            '      var_cadena = "insert into TB_TEMP_REPORTE_VENTAS_NETAS_CANTIA (inte_tem_consecutivo, dtim_tem_fecha_inicio, dtim_tem_Fecha_fin, vcha_can_canal_Venta_id,                                                vcha_age_agente_id,                                               vcha_cli_clave_id,                                                vcha_art_Articulo_id,                                                        floa_tem_Cantidad_facturada,                                                                 floa_tem_precio_facturado,                                                                  floa_tem_costo_facturado,                                                                       floa_tem_cantidad_devuelta,                                                      floa_tem_precio_devuelto,                                                               floa_tem_costo_devuelto, inte_tem_planta, vcha_Art_nombre_Español)   values"
            '      var_cadena = var_cadena + "  (" + CStr(var_consecutivo) + "," + var_fecha_inicio + ", " + var_fecha_fin + ",'" + IIf(IsNull(rs!vcha_can_canal_venta_id), "", rs!vcha_can_canal_venta_id) + "', '" + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID) + "','" + IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id) + "', "
            '      var_cadena = var_cadena + "'" + IIf(IsNull(rs!vcha_Art_Articulo_id), "", rs!vcha_Art_Articulo_id) + "', " + CStr(IIf(IsNull(rs!floa_tem_Cantidad_facturada), 0, rs!floa_tem_Cantidad_facturada)) + "," + CStr(IIf(IsNull(rs!floa_tem_precio_facturado), 0, rs!floa_tem_precio_facturado)) + "," + CStr(IIf(IsNull(rs!floa_tem_costo_facturado), 0, rs!floa_tem_costo_facturado)) + ", " + CStr(IIf(IsNull(rs!floa_tem_cantidad_devuelta), 0, rs!floa_tem_cantidad_devuelta)) + ", " + CStr(IIf(IsNull(rs!floa_tem_precio_devuelto), 0, rs!floa_tem_precio_devuelto)) + ", " + CStr(IIf(IsNull(rs!floa_tem_costo_devuelto), 0, rs!floa_tem_costo_devuelto)) + ",2, '" + IIf(IsNull(rs!vcha_Art_nombre_español), "", rs!vcha_Art_nombre_español) + "')"
            '      rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            '      rs.MoveNext
            'Wend
            'rs.Close
            rs.Open "delete from TB_TEMP_REPORTE_VENTAS_NETAS_CANTIA where INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " and vcha_Art_articulo_id is null", cnn, adOpenDynamic, adLockOptimistic
            'rs.Open "delete from TB_TEMP_REPORTE_VENTAS_NETAS_CANTIA where INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " and vcha_Art_articulo_id is null", cnn_sqlquezada2, adOpenDynamic, adLockOptimistic
            If Me.opt_articulo = True Then
               Set reporte = appl.OpenReport(App.Path + "\rep_ventas_netas_cantia_articulo.rpt")
               reporte.RecordSelectionFormula = "{VW_TEMP_REPORTE_VENTAS_NETAS_CANTIA_ARTICULO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\reporte_ventas_netas_cantia" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
            If Me.opt_agente = True Then
               Set reporte = appl.OpenReport(App.Path + "\rep_ventas_netas_cantia_agente.rpt")
               reporte.RecordSelectionFormula = "{VW_TEMP_REPORTE_VENTAS_NETAS_AGENTE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\reporte_ventas_netas_cantia" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
            If Me.opt_canal = True Then
               Set reporte = appl.OpenReport(App.Path + "\rep_ventas_netas_cantia_CANAL.rpt")
               reporte.RecordSelectionFormula = "{VW_TEMP_REPORTE_VENTAS_NETAS_CANTIA_CANAL.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\reporte_ventas_netas_cantia" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
            rs.Open "delete from TB_TEMP_REPORTE_VENTAS_NETAS_CANTIA where INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            rs.Open "delete from TB_TEMP_REPORTE_VENTAS_NETAS_CANTIA where INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn_sqlquezada2, adOpenDynamic, adLockOptimistic
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
   Top = 2700
   Left = 3500
   txt_inicio = Date
   txt_fin = Date
   Me.opt_cliente.Value = True
   
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
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      txt_inicio = var_fecha_general
   End If
End Sub

