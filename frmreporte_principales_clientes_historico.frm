VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmreporte_principales_clientes_historico 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Principales clientes historico"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   2910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2490
      Picture         =   "frmreporte_principales_clientes_historico.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmreporte_principales_clientes_historico.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   105
      TabIndex        =   10
      Top             =   360
      Width           =   2805
   End
   Begin VB.Frame Frame3 
      Caption         =   "  Agentes "
      Height          =   150
      Left            =   600
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   90
      Begin VB.Frame Frame6 
         Height          =   120
         Left            =   30
         TabIndex        =   8
         Top             =   540
         Width           =   5565
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   450
         Picture         =   "frmreporte_principales_clientes_historico.frx":073C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   120
         Picture         =   "frmreporte_principales_clientes_historico.frx":0952
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1110
         Picture         =   "frmreporte_principales_clientes_historico.frx":0A54
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   780
         Picture         =   "frmreporte_principales_clientes_historico.frx":0B26
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Marcar (Enter)"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         Picture         =   "frmreporte_principales_clientes_historico.frx":0D70
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   225
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_agentes 
         Height          =   2025
         Left            =   45
         TabIndex        =   9
         Top             =   690
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   3572
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fecha "
      Height          =   645
      Left            =   120
      TabIndex        =   0
      Top             =   375
      Width           =   2730
      Begin VB.TextBox txt_fecha 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   540
         TabIndex        =   1
         Top             =   210
         Width           =   1620
      End
   End
End
Attribute VB_Name = "frmreporte_principales_clientes_historico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_mes As Integer
Private Sub cmd_imprimir_Click()
   Dim var_cadena As String
   Dim var_mes As String
   Dim var_dia As String
   Dim var_año As String
   var_cadena = ""
   'On Error GoTo salir:
   If IsDate(Me.txt_fecha) Then
      var_contador = 0
      var_fecha_fin_1 = CDate(txt_fecha)
      var_dia = CStr(Day(CDate(txt_fecha)))
      var_mes = CStr(Month(CDate(txt_fecha)))
      var_año = CStr(Year(CDate(txt_fecha)))
      If Len(Trim(var_dia)) = 1 Then
         var_dia = "0" + var_dia
      End If
      If Len(Trim(var_mes)) = 1 Then
         var_mes = "0" + var_mes
      End If
      var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
      cnn.CommandTimeout = 360
      cnn.BeginTrans
      rs.Open "select max(inte_tem_consecutivo) from tb_temp_antiguedad_saldos", cnn, adOpenDynamic, adLockOptimistic
      If rs.EOF Then
         var_consecutivo = 1
      Else
         var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
      End If
      rs.Close
      rs.Open "insert into tb_temp_antiguedad_saldos (inte_tem_consecutivo, dtim_tem_fecha) values (" + CStr(var_consecutivo) + ", " + var_fecha + ")", cnn, adOpenDynamic, adLockOptimistic
      cnn.CommitTrans
      'rs.Open "exec SP_ANTIGUEDAD_SALDOS " + var_fecha + "," + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      If var_empresa = "03" Then
         rs.Open "EXEC SP_PRINCIPALES_CLIENTES_03 " + CStr(var_consecutivo) + "," + var_fecha
      Else
         If var_empresa = "06" Then
            rs.Open "EXEC SP_PRINCIPALES_CLIENTES_06 " + CStr(var_consecutivo) + "," + var_fecha
         Else
            var_cadena_2 = "INSERT INTO TB_TEMP_ANTIGUEDAD_SALDOS (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA, VCHA_EMP_EMPRESA_ID, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, VCHA_AGE_AGENTE_ID, VCHA_CLI_CLAVE_ID, INTE_CAR_NUMERO, FLOA_TEM_IMPORTE, floa_tem_saldo, DTIM_CAR_FECHA, INTE_TEM_DIFERENCIA, FLOA_CAR_TIPO_CAMBIO, FLOA_CAR_IMPORTE_NETO, INTE_CAR_PLAZO)"
            var_cadena_2 = var_cadena_2 + "  select " + CStr(var_consecutivo) + ", " + var_fecha + ", a.vcha_emp_empresa_id, a.vcha_Car_documento, a.vcha_Ser_Serie_id, a.vcha_age_agente_id, a.vcha_cli_clave_id, a.inte_Car_numero, "
            var_cadena_2 = var_cadena_2 + " (round((a.floa_Car_importe_neto/a.floa_Car_tipo_cambio),2) - isnull((select round(sum(floa_car_importe_neto/floa_Car_tipo_cambio),2) from vw_abonos with (nolock) where vcha_emp_empresa_id = a.vcha_emp_empresa_id and vcha_ecu_movimiento_cargo = a.vcha_car_documento and vcha_ecu_serie_cargo = a.vcha_ser_Serie_id and inte_Ecu_numero_cargo = a.inte_car_numero and dtim_Car_fecha <= ((" + var_fecha + " + 1) - .000001)) ,0)) * a.floa_car_tipo_cambio, ((a.floa_Car_importe_neto/a.floa_Car_tipo_cambio) - isnull((select sum(floa_car_importe_neto/floa_Car_tipo_cambio) from vw_abonos with (nolock) where vcha_emp_empresa_id = a.vcha_emp_empresa_id and vcha_ecu_movimiento_cargo = a.vcha_car_documento and vcha_ecu_serie_cargo = a.vcha_ser_Serie_id and inte_Ecu_numero_cargo = a.inte_car_numero and dtim_car_fecha <= ((" + var_fecha + " + 1) - .000001)) ,0)),"
            'var_cadena_2 = var_cadena_2 + " a.dtim_Car_fecha, datediff(day, a.dtim_Car_fecha+A.INTE_cAR_PLAZO, " + var_fecha + "), a.floa_car_tipo_Cambio, a.floa_car_importe_neto, a.inte_Car_plazo from tb_encabezado_cartera a where a.dtim_Car_fecha <= ((" + var_fecha + " + 1) - .0000001) and a.char_Car_afectacion = '+' and (a.char_car_Estatus <> 'C' or a.char_Car_estatus is null) and a.vcha_emp_empresa_id = '" + var_empresa + "'"
            var_cadena_2 = var_cadena_2 + " a.dtim_Car_fecha, datediff(day, a.dtim_Car_fecha+A.INTE_cAR_PLAZO, " + var_fecha + "), a.floa_car_tipo_Cambio, a.floa_car_importe_neto, a.inte_Car_plazo from tb_encabezado_cartera a with (nolock) where a.dtim_Car_fecha <= ((" + var_fecha + " + 1) - .0000001) and a.char_Car_afectacion = '+' and (a.char_car_Estatus <> 'C' or a.char_Car_estatus is null) "
            rs.Open var_cadena_2, cnn, adOpenDynamic, adLockOptimistic
         End If
      End If
      For var_i = 1 To lv_agentes.ListItems.Count
          lv_agentes.ListItems.Item(var_i).Selected = True
          If lv_agentes.selectedItem.SubItems(2) = "*" Then
             If Len(Trim(var_cadena)) = 0 Then
                var_cadena = "({VW_ANTIGUEDAD_SALDOS_HISTORICO.vcha_age_agente_id} = '" + lv_agentes.selectedItem + "'"
             Else
                var_cadena = var_cadena + " or {VW_ANTIGUEDAD_SALDOS_HISTORICO.vcha_age_agente_id} = '" + lv_agentes.selectedItem + "'"
             End If
          End If
      Next var_i
      rsaux5.Open "EXEC SP_PRINCIPALES_CLIENTES_HISTORICO " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      If var_empresa = "06" Or var_empresa = "03" Then
         Set reporte = appl.OpenReport(App.Path + "\rep_principales_clientes_historico.rpt")
         reporte.RecordSelectionFormula = "{VW_PRINCIPALES_CLIENTES_HISTORICO.inte_tem_consecutivo}= " + CStr(var_consecutivo) + " and {VW_PRINCIPALES_CLIENTES_HISTORICO.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "'"
         frmvistasprevias.cr.ReportSource = reporte
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Reporte de Antigüedad de Saldos Historico"
         frmvistasprevias.Show 1
         Set reporte = Nothing
         var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
         If var_si = 6 Then
           Set reporte = appl.OpenReport(App.Path + "\rep_principales_clientes_historico.rpt")
           reporte.RecordSelectionFormula = "{VW_PRINCIPALES_CLIENTES_HISTORICO.inte_tem_consecutivo}= " + CStr(var_consecutivo) + " and {VW_PRINCIPALES_CLIENTES_HISTORICO.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "'"
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\" + var_nombre_empresa + "_Reporte_principales_clientes_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
            MsgBox "Se a guardado el archivo " + archivo, vbOKOnly, "ATENCION"
         End If
      
      
      
      Else
         rsaux5.Open "select vcha_emp_empresa_id, vcha_emp_nombre from tb_empresas with (nolock) where vcha_emp_empresa_id not in ('06','03')", cnn, adOpenDynamic, adLockOptimistic
         While Not rsaux5.EOF
               rsaux.Open "select * from VW_PRINCIPALES_CLIENTES_HISTORICO with (nolock) where inte_tem_consecutivo =" + CStr(var_consecutivo) + " and VCHA_EMP_EMPRESA_ID = '" + rsaux5!VCHA_EMP_EMPRESA_ID + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
               'Set reporte = appl.OpenReport(App.Path + "\rep_principales_clientes_historico.rpt")
               'reporte.RecordSelectionFormula = "{VW_PRINCIPALES_CLIENTES_HISTORICO.inte_tem_consecutivo}= " + CStr(var_consecutivo) + " and {VW_PRINCIPALES_CLIENTES_HISTORICO.VCHA_EMP_EMPRESA_ID} = '" + rsaux5!VCHA_EMP_EMPRESA_ID + "'"
               'frmvistasprevias.cr.ReportSource = reporte
               'For ntablas = 1 To reporte.Database.Tables.Count
               '    reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               'Next ntablas
               'frmvistasprevias.cr.ViewReport
               'frmvistasprevias.Caption = "Reporte de Antigüedad de Saldos Historico"
               'frmvistasprevias.Show 1
               'Set reporte = Nothing
               ''var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
               var_si = 6
               If var_si = 6 Then
                  Set reporte = appl.OpenReport(App.Path + "\rep_principales_clientes_historico.rpt")
                  reporte.RecordSelectionFormula = "{VW_PRINCIPALES_CLIENTES_HISTORICO.inte_tem_consecutivo}= " + CStr(var_consecutivo) + " and {VW_PRINCIPALES_CLIENTES_HISTORICO.VCHA_EMP_EMPRESA_ID} = '" + rsaux5!VCHA_EMP_EMPRESA_ID + "'"
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  reporte.ExportOptions.FormatType = crEFTExcel80
                  reporte.ExportOptions.DestinationType = crEDTDiskFile
                  archivo = "c:\reportessid\" + rsaux5!VCHA_EMP_NOMBRE + "_Reporte_principales_clientes_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                  reporte.ExportOptions.DiskFileName = archivo
                  reporte.Export False
                  Set reporte = Nothing
               '   MsgBox "Se a guardado el archivo " + archivo, vbOKOnly, "ATENCION"
               End If
               End If
               rsaux.Close
               rsaux5.MoveNext
         Wend
         rsaux5.Close
         MsgBox "Se generaron los reportes", vbOKOnly, "ATENCION"
      End If
      rs.Open "delete from tb_temp_antiguedad_saldos where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      rs.Open "delete from TB_TEMP_PRINCIPALES_CLIENTES_HISTORICO where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
   End If
   Exit Sub
salir:
If Err.Number = -2147217871 Then
   var_contador = var_contador + 1
   If var_contador <= 3 Then
      Resume
   Else
      MsgBox "A surgido un error al generar el reporte", vbOKOnly, "ATENCION"
      rs.Open "delete from tb_temp_antiguedad_saldos where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      rs.Open "delete from tb_temp_antiguedad_saldos_historico where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
   End If
End If
End Sub

Private Sub cmd_invertir_Click()
   n = lv_agentes.ListItems.Count
   For i = 1 To n
      lv_agentes.ListItems.Item(i).Selected = True
      If lv_agentes.selectedItem.SubItems(2) = "*" Then
         lv_agentes.selectedItem.SubItems(2) = ""
         lv_agentes.ListItems.Item(i).Bold = False
         lv_agentes.ListItems.Item(i).ForeColor = &H80000012
         lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_agentes.selectedItem.SubItems(2) = "*"
         lv_agentes.ListItems.Item(i).Bold = True
         lv_agentes.ListItems.Item(i).ForeColor = &HFF0000
         lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      End If
   Next i

End Sub

Private Sub cmd_marcar_Click()
   i = lv_agentes.selectedItem.Index
   If lv_agentes.selectedItem.SubItems(2) = "*" Then
      lv_agentes.selectedItem.SubItems(2) = ""
      lv_agentes.ListItems.Item(i).Bold = False
      lv_agentes.ListItems.Item(i).ForeColor = &H80000012
      lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_agentes.Refresh
   Else
      lv_agentes.selectedItem.SubItems(2) = "*"
      lv_agentes.ListItems.Item(i).Bold = True
      lv_agentes.ListItems.Item(i).ForeColor = &HFF0000
      lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_agentes.Refresh
   End If
End Sub

Private Sub cmd_ninguno_Click()
   n = lv_agentes.ListItems.Count
   For i = 1 To n
      lv_agentes.ListItems.Item(i).Selected = True
      lv_agentes.selectedItem.SubItems(2) = ""
      lv_agentes.ListItems.Item(i).Bold = False
      lv_agentes.ListItems.Item(i).ForeColor = &H80000012
      lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   lv_agentes.Refresh
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmd_seleccion_Click()
   n = lv_agentes.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_agentes.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_agentes.selectedItem.SubItems(2) = "" And var_rellena = True Then
         lv_agentes.selectedItem.SubItems(2) = "*"
         lv_agentes.ListItems.Item(i).Bold = True
         lv_agentes.ListItems.Item(i).ForeColor = &HFF0000
         lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_agentes.selectedItem.SubItems(2) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_agentes.selectedItem.SubItems(2) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i

End Sub

Private Sub cmd_todos_Click()
   n = lv_agentes.ListItems.Count
   For i = 1 To n
      lv_agentes.ListItems.Item(i).Selected = True
      lv_agentes.selectedItem.SubItems(2) = "*"
      lv_agentes.ListItems.Item(i).Bold = True
      lv_agentes.ListItems.Item(i).ForeColor = &HFF0000
      lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
   Next i
   lv_agentes.Refresh
End Sub



Private Sub Command1_Click()
   Dim var_cadena As String
   var_cadena = ""
   For var_i = 1 To lv_agentes.ListItems.Count
       lv_agentes.ListItems.Item(var_i).Selected = True
       If lv_agentes.selectedItem.SubItems(2) = "*" Then
          If Len(Trim(var_cadena)) = 0 Then
             var_cadena = "({VW_ANTIGUEDAD_SALDOS.vcha_age_agente_id} = '" + lv_agentes.selectedItem + "'"
          Else
             var_cadena = var_cadena + "or {VW_ANTIGUEDAD_SALDOS.vcha_age_agente_id} = '" + lv_agentes.selectedItem + "'"
          End If
       End If
   Next var_i
   'Set reporte = appl.OpenReport(App.Path + "\repl_antiguedad_saldos_arqueo.rpt")
   'reporte.RecordSelectionFormula = "{VW_ANTIGUEDAD_SALDOS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' and  " + var_cadena + ")"
   'frmvistasprevias.cr.ReportSource = reporte
   'For ntablas = 1 To reporte.Database.Tables.Count
   '    reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
   'Next ntablas
   'frmvistasprevias.cr.ViewReport
   'frmvistasprevias.Caption = "Reporte de Antigüedad de Saldos"
   'frmvistasprevias.Show 1
   'Set reporte = Nothing
   var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      var_si = MsgBox("¿Por agente?", vbYesNo, "ATENCION")
      If var_si = 6 Then
      For var_i = 1 To lv_agentes.ListItems.Count
          lv_agentes.ListItems.Item(var_i).Selected = True
          If lv_agentes.selectedItem.SubItems(2) = "*" Then
             Set reporte = appl.OpenReport(App.Path + "\repl_antiguedad_saldos_arqueo.rpt")
             reporte.RecordSelectionFormula = "{VW_ANTIGUEDAD_SALDOS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' and  {VW_ANTIGUEDAD_SALDOS.VCHA_AGE_AGENTE_ID} = '" + Me.lv_agentes.selectedItem + "'"
             For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
             Next ntablas
             reporte.ExportOptions.FormatType = crEFTExcel80
             reporte.ExportOptions.DestinationType = crEDTDiskFile
             archivo = "c:\reportessid\" + var_nombre_empresa + "_REPORTE_ANTIGuedad_saldos_arqueo_" + Me.lv_agentes.selectedItem + "_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
             reporte.ExportOptions.DiskFileName = archivo
             reporte.Export False
             Set reporte = Nothing
          End If
      Next var_i
      Else
          If lv_agentes.selectedItem.SubItems(2) = "*" Then
             Set reporte = appl.OpenReport(App.Path + "\repl_antiguedad_saldos_arqueo.rpt")
             reporte.RecordSelectionFormula = "{VW_ANTIGUEDAD_SALDOS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "'"
             For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
             Next ntablas
             reporte.ExportOptions.FormatType = crEFTExcel80
             reporte.ExportOptions.DestinationType = crEDTDiskFile
             archivo = "c:\reportessid\" + var_nombre_empresa + "_REPORTE_ANTIGuedad_saldos_arqueo_" + Me.lv_agentes.selectedItem + "_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
             reporte.ExportOptions.DiskFileName = archivo
             reporte.Export False
             Set reporte = Nothing
          End If
      End If
      MsgBox "Se a terminado de guardar el archivo "
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
   Left = 4000
   txt_inicio = Date
   txt_fin = Date
   txt_fecha = Date
   'opt_linea = True
   rs.Open "select * from tb_agentes order by vcha_age_nombre", cnn, adOpenDynamic, adLockOptimistic
   numero_items_ALMACENES = 0
   While Not rs.EOF
      Set list_item = lv_agentes.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
      list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
      list_item.SubItems(2) = ""
      rs.MoveNext:
   Wend
   rs.Close
   If lv_agentes.ListItems.Count > 7 Then
      lv_agentes.ColumnHeaders(2).Width = 4220
   Else
      lv_agentes.ColumnHeaders(2).Width = 4499.71
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_despliega_menu = True Then
      var_swpassword = False
      var_modifica_registro = False
   End If
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub lv_agentes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_agentes, ColumnHeader)
End Sub

Private Sub lv_agentes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      i = lv_agentes.selectedItem.Index
      If lv_agentes.selectedItem.SubItems(2) = "*" Then
         lv_agentes.selectedItem.SubItems(2) = ""
         lv_agentes.ListItems.Item(i).Bold = False
         lv_agentes.ListItems.Item(i).ForeColor = &H80000012
         lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_agentes.Refresh
      Else
         lv_agentes.selectedItem.SubItems(2) = "*"
         lv_agentes.ListItems.Item(i).Bold = True
         lv_agentes.ListItems.Item(i).ForeColor = &HFF0000
         lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_agentes.Refresh
      End If
   End If
End Sub



Private Sub mes_DateDblClick(ByVal DateDblClicked As Date)
   Me.txt_fecha = mes.Value
   Me.txt_fecha.SetFocus
End Sub

Private Sub mes_LostFocus()
   mes.Visible = False
End Sub

Private Sub txt_fecha_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fecha) Then
         frmcalendario.mes = CDate(Me.txt_fecha)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      Me.txt_fecha = var_fecha_general
   End If
End Sub

