VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmreporte_clientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Clientes"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   10455
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9975
      Picture         =   "frmreporte_clientes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   45
      Picture         =   "frmreporte_clientes.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   45
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Height          =   45
      Left            =   0
      TabIndex        =   16
      Top             =   375
      Width           =   10335
   End
   Begin VB.Frame Frame1 
      Caption         =   " Nivel de Información "
      Height          =   4275
      Left            =   5790
      TabIndex        =   8
      Top             =   480
      Width           =   4530
      Begin VB.OptionButton opt_nivel_7 
         Caption         =   "Clientes / Establecimientos"
         Height          =   345
         Left            =   165
         TabIndex        =   15
         Top             =   3780
         Width           =   3195
      End
      Begin VB.OptionButton opt_nivel_6 
         Caption         =   "Titulares / Establecimientos"
         Height          =   345
         Left            =   165
         TabIndex        =   14
         Top             =   3225
         Width           =   3195
      End
      Begin VB.OptionButton Opt_Nivel_5 
         Caption         =   "Grupos / Titulares / Clientes"
         Height          =   345
         Left            =   165
         TabIndex        =   13
         Top             =   2655
         Width           =   3195
      End
      Begin VB.OptionButton Opt_nivel_4 
         Caption         =   "Grupos / Titulares"
         Height          =   345
         Left            =   165
         TabIndex        =   12
         Top             =   2100
         Width           =   3195
      End
      Begin VB.OptionButton opt_nivel_3 
         Caption         =   "Rutas / Titulares / Clientes"
         Height          =   345
         Left            =   165
         TabIndex        =   11
         Top             =   1545
         Width           =   3195
      End
      Begin VB.OptionButton opt_nivel_2 
         Caption         =   "Rutas / Titulares"
         Height          =   345
         Left            =   165
         TabIndex        =   10
         Top             =   990
         Width           =   3195
      End
      Begin VB.OptionButton opt_nivel_1 
         Caption         =   "Rutas"
         Height          =   345
         Left            =   165
         TabIndex        =   9
         Top             =   435
         Width           =   3195
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "  Agentes "
      Height          =   4290
      Left            =   75
      TabIndex        =   0
      Top             =   465
      Width           =   5625
      Begin VB.Frame Frame6 
         Height          =   120
         Left            =   30
         TabIndex        =   6
         Top             =   540
         Width           =   5565
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   450
         Picture         =   "frmreporte_clientes.frx":073C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   120
         Picture         =   "frmreporte_clientes.frx":0952
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1110
         Picture         =   "frmreporte_clientes.frx":0A54
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   780
         Picture         =   "frmreporte_clientes.frx":0B26
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Marcar (Enter)"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         Picture         =   "frmreporte_clientes.frx":0D70
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   225
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_agentes 
         Height          =   3510
         Left            =   45
         TabIndex        =   7
         Top             =   690
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   6191
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
End
Attribute VB_Name = "frmreporte_clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
      VAR_AGENTE_FILTRO = ""
      var_cadena = ""
      var_cadena_2 = ""
      If Me.opt_nivel_1 = True Then
         VAR_AGENTE_FILTRO = "{VW_REPORTE_CLIENTES_RUTAS"
      End If
      If Me.opt_nivel_2 = True Then
         VAR_AGENTE_FILTRO = "{VW_REPORTE_CLIENTES_RUTAS_TITULARES"
      End If
      If Me.opt_nivel_3 = True Then
         VAR_AGENTE_FILTRO = "{VW_CLIENTES"
      End If
      If Me.Opt_nivel_4 = True Then
         VAR_AGENTE_FILTRO = "{VW_REPORTE_CLIENTES_GRUPOS_TITULARES"
      End If
      If Me.Opt_Nivel_5 = True Then
          VAR_AGENTE_FILTRO = "{VW_REPORTE_CLIENTES_GRUPOS_TITULARES_CLIENTES"
      End If
      If Me.opt_nivel_6 = True Then
          VAR_AGENTE_FILTRO = "{VW_REPORTE_CLIENTES_TITULARES_eSTABLECIMIENTOS"
      End If
      If Me.opt_nivel_7 = True Then
          VAR_AGENTE_FILTRO = "{VW_REPORTE_CLIENTES_ESTABLECIMIENTOS"
      End If
      For var_i = 1 To lv_agentes.ListItems.Count
             lv_agentes.ListItems.Item(var_i).Selected = True
             If lv_agentes.selectedItem.SubItems(2) = "*" Then
                var_contador = var_contador + 1
                If Len(Trim(var_cadena)) = 0 Then
                   var_cadena = var_cadena + VAR_AGENTE_FILTRO + ".VCHA_AGE_AGENTE_ID} = '" + lv_agentes.selectedItem + "'"
                Else
                   var_cadena = var_cadena + " or " + VAR_AGENTE_FILTRO + ".VCHA_AGE_AGENTE_ID} = '" + lv_agentes.selectedItem + "'"
                End If
             End If
         Next var_i
         If Len(Trim(var_cadena)) > 0 Then
         If Me.opt_nivel_1.Value = True = True Then
            Set reporte = appl.OpenReport(App.Path + "\REP_CLIENTES_RUTAS.rpt")
            reporte.RecordSelectionFormula = var_cadena
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Clientes"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\REP_CLIENTES_RUTAS.rpt")
               reporte.RecordSelectionFormula = var_cadena
               For ntablas = 1 To reporte.Database.Tables.Count
                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\Reporte_clientes_rutas_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
         End If
         If Me.opt_nivel_2.Value = True = True Then
            Set reporte = appl.OpenReport(App.Path + "\REP_CLIENTES_RUTAS_TITULARES.rpt")
            reporte.RecordSelectionFormula = var_cadena
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Clientes"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\REP_CLIENTES_RUTAS_TITULARES.rpt")
               reporte.RecordSelectionFormula = var_cadena
               For ntablas = 1 To reporte.Database.Tables.Count
                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\Reporte_clientes_rutas_titulares_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
         End If

         If Me.opt_nivel_3.Value = True = True Then
            Set reporte = appl.OpenReport(App.Path + "\REP_CLIENTES_RUTAS_TITULARES_clientes.rpt")
            reporte.RecordSelectionFormula = "(" + var_cadena + ") and LEN({VW_CLIENTES.VCHA_TIT_TITULAR_ID}) = 10"
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Clientes"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\REP_CLIENTES_RUTAS_TITULARES_clientes.rpt")
               reporte.RecordSelectionFormula = "(" + var_cadena + ") and LEN({VW_CLIENTES.VCHA_TIT_TITULAR_ID}) = 10"
               For ntablas = 1 To reporte.Database.Tables.Count
                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\Reporte_clientes_rutas_titulares_clientes_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
         End If

         If Me.Opt_nivel_4.Value = True = True Then
            Set reporte = appl.OpenReport(App.Path + "\REP_CLIENTES_GRUPOS_TITULARES.rpt")
            reporte.RecordSelectionFormula = "(" + var_cadena + ") and LEN({vw_reporte_clientes_grupos_titulares.VCHA_TIT_TITULAR_ID}) = 10"
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Clientes"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\REP_CLIENTES_GRUPOS_TITULARES.rpt")
               reporte.RecordSelectionFormula = "(" + var_cadena + ") and LEN({vw_reporte_clientes_grupos_titulares.VCHA_TIT_TITULAR_ID}) = 10"
               For ntablas = 1 To reporte.Database.Tables.Count
                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\Reporte_clientes_grupos_titulares_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
         End If
         
         If Me.Opt_Nivel_5.Value = True = True Then
            Set reporte = appl.OpenReport(App.Path + "\REP_CLIENTES_GRUPOS_TITULARES_clientes.rpt")
            reporte.RecordSelectionFormula = "(" + var_cadena + ") and LEN({VW_REPORTE_CLIENTES_GRUPOS_TITULARES_CLIENTES.VCHA_TIT_TITULAR_ID}) = 10"
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Clientes"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\REP_CLIENTES_GRUPOS_TITULARES_clientes.rpt")
               reporte.RecordSelectionFormula = "(" + var_cadena + ") and LEN({VW_REPORTE_CLIENTES_GRUPOS_TITULARES_CLIENTES.VCHA_TIT_TITULAR_ID}) = 10"
               For ntablas = 1 To reporte.Database.Tables.Count
                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\Reporte_clientes_grupos_titulares_clientes_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
         End If
         
         If Me.opt_nivel_6.Value = True = True Then
            Set reporte = appl.OpenReport(App.Path + "\rep_clientes_titulares_establecimientos.rpt")
            reporte.RecordSelectionFormula = "(" + var_cadena + ") and LEN({VW_REPORTE_CLIENTES_TITULARES_establecimientos.VCHA_TIT_TITULAR_ID}) = 10"
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Clientes"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_clientes_titulares_establecimientos.rpt")
               reporte.RecordSelectionFormula = "(" + var_cadena + ") and LEN({VW_REPORTE_CLIENTES_TITULARES_establecimientos.VCHA_TIT_TITULAR_ID}) = 10"
               For ntablas = 1 To reporte.Database.Tables.Count
                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\Reporte_titulares_establecimientos_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
         End If
         
         
         If Me.opt_nivel_7.Value = True = True Then
            Set reporte = appl.OpenReport(App.Path + "\rep_clientes_establecimientos.rpt")
            reporte.RecordSelectionFormula = "(" + var_cadena + ")"
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Clientes"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_clientes_establecimientos.rpt")
               reporte.RecordSelectionFormula = "(" + var_cadena + ")"
               For ntablas = 1 To reporte.Database.Tables.Count
                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\Reporte_clientes_establecimientos_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
         End If
         Else
            MsgBox "No se a seleccionado un agente", vbOKOnly, "ATENCION"
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

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   Me.opt_nivel_1.Value = True
   Top = 1000
   Left = 500
   rs.Open "select * from tb_agentes where vcha_emp_empresa_id = '" + var_empresa + "' order by vcha_age_nombre", cnn, adOpenDynamic, adLockOptimistic
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
   Call activa_forma(var_activa_forma_reporte_comisiones)
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
