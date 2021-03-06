VERSION 5.00
Begin VB.Form frmreporte_entradas_salidas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Entradas y de Salidas"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4485
   Begin VB.CommandButton Command1 
      Caption         =   "SV"
      Height          =   315
      Left            =   435
      TabIndex        =   11
      ToolTipText     =   "Entradas al almacen de calidad"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_Calidad 
      Caption         =   "14"
      Height          =   315
      Left            =   1425
      TabIndex        =   10
      ToolTipText     =   "Entradas al almacen de calidad"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_valuacion_entradas_produccion 
      Appearance      =   0  'Flat
      Caption         =   "EP"
      Height          =   315
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Reporte de entradas Concentrado"
      Top             =   30
      Width           =   345
   End
   Begin VB.CommandButton cmd_reporte_entradas 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   765
      Picture         =   "frmreporte_entradas_salidas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Reporte de entradas Concentrado"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frmreporte_entradas_salidas.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4005
      Picture         =   "frmreporte_entradas_salidas.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   840
      Left            =   120
      TabIndex        =   0
      Top             =   465
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
      Left            =   30
      TabIndex        =   7
      Top             =   300
      Width           =   4395
   End
End
Attribute VB_Name = "frmreporte_entradas_salidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_Calidad_Click()
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim a?o As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_consecutivo As Integer
   Dim var_afectacion_movimiento As String
   Dim var_vistas As String
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) as numero from tb_temp_reporte_entradas_salidas", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into tb_temp_reporte_entradas_salidas (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
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
            var_fecha_fin_1 = CDate(txt_fin)
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_a?o = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin_2 = "{d '" + var_a?o + "-" + var_mes + "-" + var_dia + "'}"
            rs.Open "EXEC SP_REPORTE_ENTRADAS_SALIDAS " + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin + ",'" + var_clave_usuario_global + "', '" + fun_NombrePc + "'", cnn, adOpenDynamic, adLockOptimistic
            var_fecha_fin_1 = CDate(txt_fin)
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
            

            Set reporte = appl.OpenReport(App.Path + "\rep_entradas_salidas_entradas_detalle_sin_agrupar.rpt")
            reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_DETALLE.VCHA_aLM_ALMACEN_ID} = '14'"
            For ntablas = 1 To reporte.Database.Tables.Count
               reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\REPORTE_ENTRADAS_DETALLE" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
      
            MsgBox "Se a terminado de guardar el archivo " + archivo
            rs.Open "delete from TB_TEMP_REPORTE_ENTRADAS_SALIDAS where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
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

Private Sub cmd_imprimir_Click()
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim a?o As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_consecutivo As Integer
   Dim var_afectacion_movimiento As String
   Dim var_vistas As String
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) as numero from tb_temp_reporte_entradas_salidas", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into tb_temp_reporte_entradas_salidas (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
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
            
            'rs.Open "select * from vw_encabezado_movimientos where dtim_emo_fecha >= " + var_fecha_inicio + " and dtim_emo_fecha <= " + var_fecha_fin + "-.00001", cnn, adOpenDynamic, adLockOptimistic
            
            var_fecha_fin_1 = CDate(txt_fin)
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_a?o = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin_2 = "{d '" + var_a?o + "-" + var_mes + "-" + var_dia + "'}"
            
             rs.Open "EXEC SP_REPORTE_ENTRADAS_SALIDAS " + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin + ",'" + var_clave_usuario_global + "', '" + fun_NombrePc + "'", cnn, adOpenDynamic, adLockOptimistic
             rs.Open "DELETE FROM TB_TEMP_REPORTE_ENTRADAS_SALIDAS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND VCHA_eMP_EMPRESA_ID <> '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
            
            var_fecha_fin_1 = CDate(txt_fin)
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
            
            If var_empresa = "18" Then
               Set reporte = appl.OpenReport(App.Path + "\rep_entradas_salidas_entradas_concentrado_textilera.rpt")
            Else
               Set reporte = appl.OpenReport(App.Path + "\rep_entradas_salidas_entradas_concentrado.rpt")
            End If
            If var_empresa = "06" Or var_empresa = "31" Then
               reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_CONCENTRADO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_CONCENTRADO.vcha_emp_empresa_id} = '" + var_empresa + "'"
            Else
               reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_CONCENTRADO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            End If
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Entradas concentrado"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            var_si = MsgBox("?Desea importar el reporte?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               If var_empresa = "18" Then
                  Set reporte = appl.OpenReport(App.Path + "\rep_entradas_salidas_entradas_concentrado_TEXTILERA_excel.rpt")
               Else
                  Set reporte = appl.OpenReport(App.Path + "\rep_entradas_salidas_entradas_concentrado_excel.rpt")
               End If
               If var_empresa = "06" Or var_empresa = "31" Then
                  reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_CONCENTRADO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_CONCENTRADO.vcha_emp_empresa_id} = '" + var_empresa + "'"
               Else
                  reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_CONCENTRADO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               End If
               For ntablas = 1 To reporte.Database.Tables.Count
                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\REPORTE_ENTRADAS_CONCENTRADO" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
            
            Set reporte = appl.OpenReport(App.Path + "\rep_entradas_salidas_salidas_concentrado.rpt")
            If var_empresa = "06" Or var_empresa = "31" Then
               reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_SALIDAS_SALIDAS_CONCENTRADO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_REPORTE_ENTRADAS_SALIDAS_SALIDAS_CONCENTRADO.vcha_emp_empresa_id} = '" + var_empresa + "'"
            Else
               reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_SALIDAS_SALIDAS_CONCENTRADO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            End If
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Salidas concentrado"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            
            var_si = MsgBox("?Desea importar el reporte?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_entradas_salidas_salidas_concentrado_excel.rpt")
               If var_empresa = "06" Or var_empresa = "31" Then
                  reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_SALIDAS_SALIDAS_CONCENTRADO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_REPORTE_ENTRADAS_SALIDAS_SALIDAS_CONCENTRADO.vcha_emp_empresa_id} = '" + var_empresa + "'"
               Else
                  reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_SALIDAS_SALIDAS_CONCENTRADO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               End If
               For ntablas = 1 To reporte.Database.Tables.Count
                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\REPORTE_SALIDAS_CONCENTRADO" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
            If var_empresa = "18" Then
               Set reporte = appl.OpenReport(App.Path + "\rep_entradas_salidas_entradas_detalle_textilera.rpt")
            Else
               Set reporte = appl.OpenReport(App.Path + "\rep_entradas_salidas_entradas_detalle.rpt")
            End If
            If var_empresa = "06" Or var_empresa = "31" Then
               reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_DETALLE.vcha_emp_empresa_id} = '" + var_empresa + "'"
            Else
               reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            End If
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Entradas Detalle"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            
            var_si = MsgBox("?Desea importar el reporte?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               If var_empresa = "18" Then
                  Set reporte = appl.OpenReport(App.Path + "\rep_entradas_salidas_entradas_detalle_TEXTILERA.rpt")
               Else
                  Set reporte = appl.OpenReport(App.Path + "\rep_entradas_salidas_entradas_detalle_excel.rpt")
               End If
               If var_empresa = "06" Or var_empresa = "31" Then
                  reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_DETALLE.vcha_emp_empresa_id} = '" + var_empresa + "'"
               Else
                  reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               End If
               For ntablas = 1 To reporte.Database.Tables.Count
                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\REPORTE_ENTRADAS_DETALLE" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
         
            Set reporte = appl.OpenReport(App.Path + "\rep_entradas_salidas_salidas_detalle.rpt")
            If var_empresa = "06" Or var_empresa = "31" Then
               reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_SALIDAS_SALIDAS_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_REPORTE_ENTRADAS_SALIDAS_SALIDAS_DETALLE.vcha_Emp_empresa_id} = '" + var_empresa + "'"
            Else
               reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_SALIDAS_SALIDAS_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            End If
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Salidas Detalle"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            
            var_si = MsgBox("?Desea importar el reporte?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_entradas_salidas_salidas_detalle_excel.rpt")
               If var_empresa = "06" Or var_empresa = "31" Then
                  reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_SALIDAS_SALIDAS_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_REPORTE_ENTRADAS_SALIDAS_SALIDAS_DETALLE.vcha_emp_empresa_id} = '" + var_empresa + "'"
               Else
                  reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_SALIDAS_SALIDAS_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               End If
               For ntablas = 1 To reporte.Database.Tables.Count
                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\REPORTE_SALIDAS_DETALLE" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
            rs.Open "delete from TB_TEMP_REPORTE_ENTRADAS_SALIDAS where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
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

Private Sub cmd_reporte_entradas_Click()
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim a?o As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_consecutivo As Integer
   Dim var_afectacion_movimiento As String
   Dim var_vistas As String
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) as numero from tb_temp_reporte_entradas_salidas", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into tb_temp_reporte_entradas_salidas (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
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
            
            
            var_fecha_fin_1 = CDate(txt_fin)
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_a?o = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin_2 = "{d '" + var_a?o + "-" + var_mes + "-" + var_dia + "'}"
            
                          
             rs.Open "EXEC SP_REPORTE_ENTRADAS_SALIDAS " + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin + ",'" + var_clave_usuario_global + "', '" + fun_NombrePc + "'", cnn, adOpenDynamic, adLockOptimistic
            
            
            var_fecha_fin_1 = CDate(txt_fin)
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
            
            
            Set reporte = appl.OpenReport(App.Path + "\rep_entradas_concentrado.rpt")
            reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_CONCENTRADO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and ({VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_CONCENTRADO.VCHA_MOV_MOVIMIENTO_ID} = 'EP' or {VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_CONCENTRADO.VCHA_MOV_MOVIMIENTO_ID} = 'EC' or {VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_CONCENTRADO.VCHA_MOV_MOVIMIENTO_ID} = 'RE')"
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Entradas concentrado"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            var_si = MsgBox("?Desea importar el reporte?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_entradas_concentrado.rpt")
               reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_CONCENTRADO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\REPORTE_ENTRADAS_CONCENTRADO" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
            
            
            Set reporte = appl.OpenReport(App.Path + "\rep_entradas_salidas_salidas_concentrado.rpt")
            reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_SALIDAS_SALIDAS_CONCENTRADO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_REPORTE_ENTRADAS_SALIDAS_SALIDAS_CONCENTRADO.vcha_mov_movimiento_id} = 'SR'"
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Salidas concentrado"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            
            var_si = MsgBox("?Desea importar el reporte?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_entradas_salidas_salidas_concentrado.rpt")
               reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_SALIDAS_SALIDAS_CONCENTRADO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_REPORTE_ENTRADAS_SALIDAS_SALIDAS_CONCENTRADO.vcha_mov_movimiento_id} = 'SR'"
               For ntablas = 1 To reporte.Database.Tables.Count
                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\REPORTE_SALIDAS_CONCENTRADO" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
            
            
            
            
            rs.Open "delete from TB_TEMP_REPORTE_ENTRADAS_SALIDAS where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            
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
   Dim a?o As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_consecutivo As Integer
   Dim var_afectacion_movimiento As String
   Dim var_vistas As String
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) as numero from tb_temp_reporte_entradas_salidas", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into tb_temp_reporte_entradas_salidas (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
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
            
            var_fecha_fin_1 = CDate(txt_fin)
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_a?o = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin_2 = "{d '" + var_a?o + "-" + var_mes + "-" + var_dia + "'}"
            
                          
            rs.Open "EXEC SP_REPORTE_ENTRADAS_SALIDAS " + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin + ",'" + var_clave_usuario_global + "', '" + fun_NombrePc + "'", cnn, adOpenDynamic, adLockOptimistic
            rs.Open "DELETE FROM TB_TEMP_REPORTE_ENTRADAS_SALIDAS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND VCHA_eMP_EMPRESA_ID <> '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
            
            var_fecha_fin_1 = CDate(txt_fin)
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
                                                       
            Set reporte = appl.OpenReport(App.Path + "\rep_entradas_salidas_entradas_detalle_textilera_3.rpt")
            reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_DETALLE.vcha_emp_empresa_id} = '" + var_empresa + "' AND ({VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_DETALLE.VCHA_MOV_MOVIMIENTO_ID} = 'SV' OR {VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_DETALLE.VCHA_MOV_MOVIMIENTO_ID} = 'EC' OR {VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_DETALLE.VCHA_MOV_MOVIMIENTO_ID} = 'ETP'  OR {VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_DETALLE.VCHA_MOV_MOVIMIENTO_ID} = 'ETA')"
            For ntablas = 1 To reporte.Database.Tables.Count
               reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\REPORTE_ENTRADAS_DETALLE_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
            MsgBox "Se a terminado de guardar el archivo " + archivo
            rs.Open "delete from TB_TEMP_REPORTE_ENTRADAS_SALIDAS where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
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

Private Sub cmd_valuacion_entradas_produccion_Click()
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim a?o As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_consecutivo As Integer
   Dim var_afectacion_movimiento As String
   Dim var_vistas As String
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) as numero from tb_temp_reporte_entradas_salidas", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into tb_temp_reporte_entradas_salidas (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
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
            
            
            var_fecha_fin_1 = CDate(txt_fin)
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_a?o = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin_2 = "{d '" + var_a?o + "-" + var_mes + "-" + var_dia + "'}"
            
            rs.Open "EXEC SP_REPORTE_ENTRADAS_SALIDAS " + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin + ",'" + var_clave_usuario_global + "', '" + fun_NombrePc + "'", cnn, adOpenDynamic, adLockOptimistic
            
            var_fecha_fin_1 = CDate(txt_fin)
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
            
            
            Set reporte = appl.OpenReport(App.Path + "\rep_entradas_salidas_entradas_detalle_produccion.rpt")
            If var_empresa = "18" Then
               reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_DETALLE.vcha_mov_movimiento_id} = 'EP' and {VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_DETALLE.vcha_Emp_empresa_id} = '" + var_empresa + "'"
            Else
               reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_DETALLE.vcha_mov_movimiento_id} = 'EP'"
            End If
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Entradas Detalle"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            
            var_si = MsgBox("?Desea importar el reporte?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_entradas_salidas_entradas_detalle_produccion.rpt")
               If var_empresa = "18" Then
                  reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_DETALLE.vcha_mov_movimiento_id} = 'EP' and {VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_DETALLE.vcha_Emp_empresa_id} = '" + var_empresa + "'"
               Else
                  reporte.RecordSelectionFormula = "{VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_REPORTE_ENTRADAS_SALIDAS_ENTRADAS_DETALLE.vcha_mov_movimiento_id} = 'EP'"
               End If
               For ntablas = 1 To reporte.Database.Tables.Count
                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\REPORTE_ENTRADAS_PRODUCCION_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
         
            rs.Open "delete from TB_TEMP_REPORTE_ENTRADAS_SALIDAS where INTE_TEM_CONSECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
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
   var_cadena_seguridad = ""
   Top = 3000
   Left = 3500
   txt_inicio = Date
   txt_fin = Date
   If var_empresa = "18" Then
     Me.cmd_valuacion_entradas_produccion.Enabled = True
     Me.cmd_reporte_entradas.Enabled = False
   Else
     Me.cmd_valuacion_entradas_produccion.Visible = False
     If var_empresa <> "06" Then
        Me.cmd_reporte_entradas.Enabled = True
     Else
        Me.cmd_reporte_entradas.Enabled = False
        Me.cmd_Calidad.Enabled = False
     End If
     If var_empresa = "31" Then
        Me.cmd_reporte_entradas.Enabled = False
        Me.cmd_valuacion_entradas_produccion.Enabled = False
        Me.cmd_Calidad.Enabled = False
     End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    var_swpassword = False
    var_modifica_registro = False
    Call activa_forma(var_activa_forma_reporte_entradas_salidas)
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
