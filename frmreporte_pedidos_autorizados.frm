VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmreporte_pedidos_autorizados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de pedidos autorizados"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.MonthView mes 
      Height          =   2370
      Left            =   1770
      TabIndex        =   0
      Top             =   2325
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   83230721
      CurrentDate     =   38148
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Reporte "
      Height          =   705
      Left            =   0
      TabIndex        =   19
      Top             =   3315
      Width           =   5685
      Begin VB.OptionButton opt_general 
         Caption         =   "General"
         Height          =   195
         Left            =   1215
         TabIndex        =   21
         Top             =   308
         Width           =   930
      End
      Begin VB.OptionButton opt_linea 
         Caption         =   "Por Linea"
         Height          =   270
         Left            =   3060
         TabIndex        =   20
         Top             =   270
         Width           =   1410
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5325
      Picture         =   "frmreporte_pedidos_autorizados.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      Picture         =   "frmreporte_pedidos_autorizados.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   45
      TabIndex        =   16
      Top             =   360
      Width           =   5685
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   720
      Left            =   15
      TabIndex        =   9
      Top             =   4155
      Width           =   5670
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   3690
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   270
         Width           =   1095
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   1110
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   270
         Width           =   1080
      End
      Begin VB.CommandButton Command11 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2220
         Picture         =   "frmreporte_pedidos_autorizados.frx":073C
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fecha Inicial"
         Top             =   270
         Width           =   330
      End
      Begin VB.CommandButton Command12 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4830
         Picture         =   "frmreporte_pedidos_autorizados.frx":19AE
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fecha Final"
         Top             =   270
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   3375
         TabIndex        =   15
         Top             =   330
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   660
         TabIndex        =   14
         Top             =   330
         Width           =   420
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "  Agentes "
      Height          =   2880
      Left            =   60
      TabIndex        =   1
      Top             =   405
      Width           =   5625
      Begin VB.Frame Frame6 
         Height          =   120
         Left            =   30
         TabIndex        =   7
         Top             =   540
         Width           =   5565
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   450
         Picture         =   "frmreporte_pedidos_autorizados.frx":2C20
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   120
         Picture         =   "frmreporte_pedidos_autorizados.frx":2E36
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1110
         Picture         =   "frmreporte_pedidos_autorizados.frx":2F38
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   780
         Picture         =   "frmreporte_pedidos_autorizados.frx":300A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Marcar (Enter)"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         Picture         =   "frmreporte_pedidos_autorizados.frx":3254
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   225
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_agentes 
         Height          =   2025
         Left            =   45
         TabIndex        =   8
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
End
Attribute VB_Name = "frmreporte_pedidos_autorizados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_mes As Integer
Private Sub cmd_imprimir_Click()
   Dim pError As ADODB.Error
   'On Error GoTo salir:
   Dim var_consecutivo As Double
   Dim var_contador As Double
   Dim var_cadena As String
   Dim var_cadena_2 As String
   Dim var_contador_errores As Integer
   var_contador_errores = 0
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If var_empresa = "02" Or var_empresa = "03" Then
            var_contador = 0
            var_cadena = ""
            var_cadena_2 = ""
            For var_i = 1 To lv_agentes.ListItems.Count
                lv_agentes.ListItems.Item(var_i).Selected = True
                If lv_agentes.selectedItem.SubItems(2) = "*" Then
                   var_contador = var_contador + 1
                   If Len(Trim(var_cadena)) = 0 Then
                      var_cadena = var_cadena + "{VW_TEMP_REPORTE_COMISIONES.VCHA_AGE_AGENTE_ID} = '" + lv_agentes.selectedItem + "'"
                   Else
                      var_cadena = var_cadena + " or {VW_TEMP_REPORTE_COMISIONES.VCHA_AGE_AGENTE_ID} = '" + lv_agentes.selectedItem + "'"
                   End If
                   If Len(Trim(var_cadena_2)) = 0 Then
                      var_cadena_2 = var_cadena_2 + " {VW_TEMP_REPORTE_COMISIONES_GENERAL.VCHA_AGE_AGENTE_ID} = '" + lv_agentes.selectedItem + "'"
                   Else
                      var_cadena_2 = var_cadena_2 + " or {VW_TEMP_REPORTE_COMISIONES_GENERAL.VCHA_AGE_AGENTE_ID} = '" + lv_agentes.selectedItem + "'"
                   End If
                End If
            Next var_i
            If var_contador > 0 Then
               cnn.CommandTimeout = 360
               cnn.BeginTrans
               rs.Open "select max(inte_com_consecutivo) from tb_temP_reporte_comisiones", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
               Else
                  var_consecutivo = 1
               End If
               rs.Close
               rs.Open "insert into tb_temP_reporte_comisiones (inte_com_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
               cnn.CommitTrans
               var_n = Me.lv_agentes.ListItems.Count
               For var_i = 1 To var_n
                   lv_agentes.ListItems.Item(var_i).Selected = True
                   If lv_agentes.selectedItem.SubItems(2) = "*" Then
                     rs.Open "insert into tb_temp_agentes_comisiones (inte_tem_consecutivo, vcha_age_agente_id) values (" + CStr(var_consecutivo) + ",'" + lv_agentes.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
                   End If
               Next var_i
               rs.Open "EXEC SP_CALCULO_COMISIONES " + CStr(var_consecutivo) + ", '" + CStr(CDate(Me.txt_inicio)) + "', '" + CStr(CDate(Me.txt_fin) + 1) + "'", cnn, adOpenDynamic, adLockOptimistic
               
               If opt_linea.Value = True Then
                  Set reporte = appl.OpenReport(App.Path + "\rep_comisiones_linea_2.rpt")
                  reporte.RecordSelectionFormula = "{VW_TEMP_REPORTE_COMISIONES.INTE_COM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and (" + var_cadena + ")"
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Reporte de Comisiones por Linea"
                  frmvistasprevias.Show 1
                  Set reporte = Nothing
                  var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_comisiones_linea_2.rpt")
                     reporte.RecordSelectionFormula = "{VW_TEMP_REPORTE_COMISIONES.INTE_COM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and (" + var_cadena + ")"
                     For ntablas = 1 To reporte.Database.Tables.Count
                        reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     reporte.ExportOptions.FormatType = crEFTExcel80
                     reporte.ExportOptions.DestinationType = crEDTDiskFile
                     archivo = "c:\reportessid\Reporte_comisiones_linea" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     reporte.ExportOptions.DiskFileName = archivo
                     reporte.Export False
                     Set reporte = Nothing
                     MsgBox "Se a terminado de guardar el archivo " + archivo
                  End If
               End If
               If opt_general.Value = True Then
                  Set reporte = appl.OpenReport(App.Path + "\rep_comisiones_general_2.rpt")
                  reporte.RecordSelectionFormula = "{VW_TEMP_REPORTE_COMISIONES_GENERAL.INTE_COM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and(" + var_cadena_2 + ")"
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Reporte de Comisiones General"
                  frmvistasprevias.Show 1
                  Set reporte = Nothing
                  var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_comisiones_general_2.rpt")
                     reporte.RecordSelectionFormula = "{VW_TEMP_REPORTE_COMISIONES_GENERAL.INTE_COM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and(" + var_cadena_2 + ")"
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     reporte.ExportOptions.FormatType = crEFTExcel80
                     reporte.ExportOptions.DestinationType = crEDTDiskFile
                     archivo = "c:\reportessid\Reporte_comisiones_general" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     reporte.ExportOptions.DiskFileName = archivo
                     reporte.Export False
                     Set reporte = Nothing
                     MsgBox "Se a terminado de guardar el archivo " + archivo
                  End If
               End If
               rs.Open "delete from tb_temp_REPORTE_comisiones where inte_com_consecutivo = " + CStr(var_consecutivo)
               rs.Open "delete from tb_temp_agentes_comisiones where inte_tem_consecutivo = " + CStr(var_consecutivo)
            Else
               MsgBox "No se a seleccionado un agente", vbOKOnly, "ATENCION"
            End If
         Else
            If var_empresa = "18" Then
            
               var_hora_inicio = CStr(Now)
               var_contador = 0
               var_cadena = ""
               var_cadena_2 = ""
               For var_i = 1 To lv_agentes.ListItems.Count
                   lv_agentes.ListItems.Item(var_i).Selected = True
                   If lv_agentes.selectedItem.SubItems(2) = "*" Then
                      var_contador = var_contador + 1
                      If Len(Trim(var_cadena)) = 0 Then
                         var_cadena = var_cadena + "{VW_TEMP_REPORTE_COMISIONES.VCHA_AGE_AGENTE_ID} = '" + lv_agentes.selectedItem + "'"
                      Else
                         var_cadena = var_cadena + " or {VW_TEMP_REPORTE_COMISIONES.VCHA_AGE_AGENTE_ID} = '" + lv_agentes.selectedItem + "'"
                      End If
                      If Len(Trim(var_cadena_2)) = 0 Then
                         var_cadena_2 = var_cadena_2 + " {VW_TEMP_REPORTE_COMISIONES_GENERAL.VCHA_AGE_AGENTE_ID} = '" + lv_agentes.selectedItem + "'"
                      Else
                         var_cadena_2 = var_cadena_2 + " or {VW_TEMP_REPORTE_COMISIONES_GENERAL.VCHA_AGE_AGENTE_ID} = '" + lv_agentes.selectedItem + "'"
                      End If
                   End If
               Next var_i
               If var_contador > 0 Then
                  cnn.CommandTimeout = 360
                  cnn.BeginTrans
                  cnn_sqlquezada2.BeginTrans
                  rs.Open "delete from tb_temP_reporte_comisiones", cnn_sqlquezada2, adOpenDynamic, adLockOptimistic
                  rs.Open "delete from tb_temp_agentes_comisiones", cnn_sqlquezada2, adOpenDynamic, adLockOptimistic
                  rs.Open "delete from tb_temP_reporte_comisiones", cnn, adOpenDynamic, adLockOptimistic
                  rs.Open "delete from tb_temp_agentes_comisiones", cnn, adOpenDynamic, adLockOptimistic
                  
                  rs.Open "select max(inte_com_consecutivo) from tb_temP_reporte_comisiones", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
                  Else
                     var_consecutivo = 1
                  End If
                  rs.Close
                  rs.Open "insert into tb_temP_reporte_comisiones (inte_com_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                  rs.Open "insert into tb_temP_reporte_comisiones (inte_com_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn_sqlquezada2, adOpenDynamic, adLockOptimistic
                  cnn.CommitTrans
                  cnn_sqlquezada2.CommitTrans
                  var_n = Me.lv_agentes.ListItems.Count
                  For var_i = 1 To var_n
                      lv_agentes.ListItems.Item(var_i).Selected = True
                      If lv_agentes.selectedItem.SubItems(2) = "*" Then
                         rs.Open "insert into tb_temp_agentes_comisiones (inte_tem_consecutivo, vcha_age_agente_id) values (" + CStr(var_consecutivo) + ",'" + lv_agentes.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
                         rs.Open "insert into tb_temp_agentes_comisiones (inte_tem_consecutivo, vcha_age_agente_id) values (" + CStr(var_consecutivo) + ",'" + lv_agentes.selectedItem + "')", cnn_sqlquezada2, adOpenDynamic, adLockOptimistic
                      End If
                  Next var_i
                  rs.Open "DELETE FROM TB_TEMP_PARTICIPACION", cnn, adOpenDynamic, adLockOptimistic
                  var_cadena = "insert into TB_TEMP_PARTICIPACION ( INTE_TEMP_CONSECUTIVO, VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO, "
                  var_cadena = var_cadena + " VCHA_CAR_CLASE_ID, INTE_CAR_NUMERO, VCHA_LIN_LINEA_ID, CANTIDAD, VCHA_LIN_NOMBRE, IMPORTE, FLOA_CAR_PORCENTAJE_DESCUENTO_1, FLOA_CAR_PORCENTAJE_DESCUENTO_2, FLOA_CAR_PORCENTAJE_IVA, FLOA_CAR_TIPO_CAMBIO,"
                  var_cadena = var_cadena + " FLOA_CAR_SUBIMPORTE, FLOA_CAR_IMPORTE_NETO, DTIM_CAR_FECHA, VCHA_AGE_AGENTE_ID, PARTICIPACION, VCHA_SER_SERIE_ID, FLOA_COM_PORCENTAJE_IVA) SELECT " + CStr(var_consecutivo) + ", VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_CAR_DOCUMENTO,"
                  var_cadena = var_cadena + " VCHA_CAR_CLASE_ID, INTE_CAR_NUMERO, VCHA_LIN_LINEA_ID, CANTIDAD, VCHA_LIN_NOMBRE, IMPORTE,"
                  var_cadena = var_cadena + " FLOA_CAR_PORCENTAJE_DESCUENTO_1, FLOA_CAR_PORCENTAJE_DESCUENTO_2, FLOA_CAR_PORCENTAJE_IVA, FLOA_CAR_TIPO_CAMBIO,"
                  var_cadena = var_cadena + " FLOA_CAR_SUBIMPORTE , FLOA_CAR_IMPORTE_NETO, DTIM_CAR_FECHA, a.VCHA_AGE_AGENTE_ID, PARTICIPACION, VCHA_SER_SERIE_ID, FLOA_CAR_PORCENTAJE_IVA"
                  var_cadena = var_cadena + " from sqlquezada2.VIANNEY.DBO.VW_DETALLE_FACTURACION_LINEAS a, tb_temp_agentes_comisiones b where a.vcha_age_agente_id = b.vcha_age_agente_id and b.inte_tem_consecutivo = " + CStr(var_consecutivo)
                  rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  rs.Open "EXEC SP_CALCULO_COMISIONES_TEXTILERA " + CStr(var_consecutivo) + ", '" + CStr(CDate(Me.txt_inicio)) + "', '" + CStr(CDate(Me.txt_fin) + 1) + "'", cnn, adOpenDynamic, adLockOptimistic
                  'rs.Open "INSERT INTO TB_TEMP_REPORTE_COMISIONES SELECT * FROM sqlquezada2.VIANNEY.DBO.TB_TEMP_REPORTE_COMISIONES", cnn, adOpenDynamic, adLockOptimistic
                  If opt_linea.Value = True Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_comisiones_linea_2.rpt")
                     reporte.RecordSelectionFormula = "{VW_TEMP_REPORTE_COMISIONES.INTE_COM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and (" + var_cadena + ")"
                     frmvistasprevias.cr.ReportSource = reporte
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     frmvistasprevias.cr.ViewReport
                     frmvistasprevias.Caption = "Reporte de Comisiones por Linea"
                     frmvistasprevias.Show 1
                     Set reporte = Nothing
                     var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
                     If var_si = 6 Then
                        Set reporte = appl.OpenReport(App.Path + "\rep_comisiones_linea_2.rpt")
                        reporte.RecordSelectionFormula = "{VW_TEMP_REPORTE_COMISIONES.INTE_COM_CONSECUTIVO} = " + CStr(var_consecutivo) + var_cadena
                        For ntablas = 1 To reporte.Database.Tables.Count
                           reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                        Next ntablas
                        reporte.ExportOptions.FormatType = crEFTExcel80
                        reporte.ExportOptions.DestinationType = crEDTDiskFile
                        archivo = "c:\reportessid\Reporte_comisiones_linea" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                        reporte.ExportOptions.DiskFileName = archivo
                        reporte.Export False
                        Set reporte = Nothing
                        MsgBox "Se a terminado de guardar el archivo " + archivo
                     End If
                  End If
                  If opt_general.Value = True Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_comisiones_general_2.rpt")
                     reporte.RecordSelectionFormula = "{VW_TEMP_REPORTE_COMISIONES_GENERAL.INTE_COM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and(" + var_cadena_2 + ")"
                     frmvistasprevias.cr.ReportSource = reporte
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     frmvistasprevias.cr.ViewReport
                     frmvistasprevias.Caption = "Reporte de Comisiones General"
                     frmvistasprevias.Show 1
                     Set reporte = Nothing
                     var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
                     If var_si = 6 Then
                        Set reporte = appl.OpenReport(App.Path + "\rep_comisiones_general_2.rpt")
                        reporte.RecordSelectionFormula = "{VW_TEMP_REPORTE_COMISIONES_GENERAL.INTE_COM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and(" + var_cadena_2 + ")"
                        For ntablas = 1 To reporte.Database.Tables.Count
                            reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                        Next ntablas
                        reporte.ExportOptions.FormatType = crEFTExcel80
                        reporte.ExportOptions.DestinationType = crEDTDiskFile
                        archivo = "c:\reportessid\Reporte_comisiones_general" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                        reporte.ExportOptions.DiskFileName = archivo
                        reporte.Export False
                        Set reporte = Nothing
                        MsgBox "Se a terminado de guardar el archivo " + archivo
                     End If
                  End If
                  rs.Open "delete from tb_temp_REPORTE_comisiones where inte_com_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  rs.Open "delete from tb_temp_agentes_comisiones where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  rs.Open "delete from tb_temp_REPORTE_comisiones where inte_com_consecutivo = " + CStr(var_consecutivo), cnn_sqlquezada2, adOpenDynamic, adLockOptimistic
                  rs.Open "delete from tb_temp_agentes_comisiones where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn_sqlquezada2, adOpenDynamic, adLockOptimistic
               Else
                  MsgBox "No se a seleccionado un agente", vbOKOnly, "ATENCION"
               End If
            Else
               var_contador = 0
               var_cadena = ""
               var_cadena_2 = ""
               For var_i = 1 To lv_agentes.ListItems.Count
                   lv_agentes.ListItems.Item(var_i).Selected = True
                   If lv_agentes.selectedItem.SubItems(2) = "*" Then
                      var_contador = var_contador + 1
                      If Len(Trim(var_cadena)) = 0 Then
                         var_cadena = var_cadena + "{VW_REPORTE_COMISIONES_EMPRESAS.VCHA_AGE_AGENTE_ID} = '" + lv_agentes.selectedItem + "'"
                      Else
                         var_cadena = var_cadena + " or {VW_REPORTE_COMISIONES_EMPRESAS.VCHA_AGE_AGENTE_ID} = '" + lv_agentes.selectedItem + "'"
                      End If
                      If Len(Trim(var_cadena_2)) = 0 Then
                         var_cadena_2 = var_cadena_2 + " {VW_REPORTE_COMISIONES_EMPRESAS.VCHA_AGE_AGENTE_ID} = '" + lv_agentes.selectedItem + "'"
                      Else
                         var_cadena_2 = var_cadena_2 + " or {VW_REPORTE_COMISIONES_EMPRESAS.VCHA_AGE_AGENTE_ID} = '" + lv_agentes.selectedItem + "'"
                      End If
                   End If
               Next var_i
               If var_contador > 0 Then
                  cnn.CommandTimeout = 360
                  cnn.BeginTrans
                  rs.Open "select max(inte_com_consecutivo) from tb_temP_reporte_comisiones", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
                  Else
                     var_consecutivo = 1
                  End If
                  rs.Close
                  rs.Open "insert into tb_temP_reporte_comisiones (inte_com_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                  cnn.CommitTrans
                  var_n = Me.lv_agentes.ListItems.Count
                  For var_i = 1 To var_n
                      lv_agentes.ListItems.Item(var_i).Selected = True
                      If lv_agentes.selectedItem.SubItems(2) = "*" Then
                        rs.Open "insert into tb_temp_agentes_comisiones (inte_tem_consecutivo, vcha_age_agente_id) values (" + CStr(var_consecutivo) + ",'" + lv_agentes.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
                      End If
                  Next var_i
                  rs.Open "EXEC SP_CALCULO_COMISIONES_EMPRESAS " + CStr(var_consecutivo) + ", '" + CStr(CDate(Me.txt_inicio)) + "', '" + CStr(CDate(Me.txt_fin) + 1) + "'", cnn, adOpenDynamic, adLockOptimistic
                  Set reporte = appl.OpenReport(App.Path + "\rep_comisiones_EMPRESAS.rpt")
                  reporte.RecordSelectionFormula = "{VW_REPORTE_COMISIONES_EMPRESAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Reporte de Comisiones General"
                  frmvistasprevias.Show 1
                  Set reporte = Nothing
                  var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_comisiones_EMPRESAS.rpt")
                     reporte.RecordSelectionFormula = "{VW_REPORTE_COMISIONES_EMPRESAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and(" + var_cadena_2 + ")"
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     reporte.ExportOptions.FormatType = crEFTExcel80
                     reporte.ExportOptions.DestinationType = crEDTDiskFile
                     archivo = "c:\reportessid\Reporte_comisiones_general" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     reporte.ExportOptions.DiskFileName = archivo
                     reporte.Export False
                     Set reporte = Nothing
                     MsgBox "Se a terminado de guardar el archivo " + archivo
                  End If
                  rs.Open "delete from TB_TEMP_COMISIONES_EMPRESAS where inte_tem_consecutivo = " + CStr(var_consecutivo)
                  rs.Open "delete from tb_temp_agentes_comisiones where inte_tem_consecutivo = " + CStr(var_consecutivo)
               Else
                  MsgBox "No se a seleccionado un agente", vbOKOnly, "ATENCION"
               End If
            End If
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de inicio incorrecta", vbOKOnly, "ATENCION"
   End If
    Exit Sub
    
salir:
If Err.Number = -2147217871 Then
   var_si = MsgBox("El sistema a marcado tiempo de espera agotado, ¿Desea continuar?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      Resume
      var_contador_errores = var_contador_errores + 1
      If var_contador_errores = 4 Then
         MsgBox "A surgido un error al conectarce a la base de datos", vbOKOnly, "ATENCION"
         Exit Sub
      End If
   Else
      Exit Sub
   End If
  
Else
   MsgBox "A surgido un error", vbOKOnly, "ATENCION"
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

Private Sub Command11_Click()
   If IsDate(Me.txt_inicio) Then
      Me.mes.Value = CDate(Me.txt_inicio)
   Else
      mes.Value = Date
   End If
   var_mes = 1
   mes.Visible = True
   mes.SetFocus
End Sub

Private Sub Command12_Click()
   If IsDate(Me.txt_fin) Then
      mes.Value = CDate(Me.txt_fin)
   Else
      mes.Value = Date
   End If
   var_mes = 2
   mes.Visible = True
   mes.SetFocus
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
   Top = 1500
   Left = 3200
   txt_inicio = Date
   txt_fin = Date
   'opt_linea = True
   Me.opt_general = True
   mes.Visible = False
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
   Call activa_forma(var_activa_forma_reporte_comisiones)
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
   If var_mes = 1 Then
      txt_inicio = mes.Value
   End If
   If var_mes = 2 Then
      txt_fin = mes.Value
   End If
   mes.Visible = False
End Sub

Private Sub mes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      mes.Visible = False
   End If
End Sub

Private Sub mes_LostFocus()
   mes.Visible = False
End Sub


