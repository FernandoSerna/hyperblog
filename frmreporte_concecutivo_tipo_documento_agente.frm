VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmreporte_concecutivo_tipo_documento_agente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de consecutivo de tipo de documento por agente"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_dolares 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   420
      Picture         =   "frmreporte_concecutivo_tipo_documento_agente.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Imprimir Movimiento en Dolares"
      Top             =   75
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Caption         =   "  Documentos "
      Height          =   3420
      Left            =   5835
      TabIndex        =   16
      Top             =   480
      Width           =   5715
      Begin VB.Frame Frame5 
         Height          =   120
         Left            =   30
         TabIndex        =   22
         Top             =   540
         Width           =   5640
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   450
         Picture         =   "frmreporte_concecutivo_tipo_documento_agente.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   120
         Picture         =   "frmreporte_concecutivo_tipo_documento_agente.frx":0318
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1110
         Picture         =   "frmreporte_concecutivo_tipo_documento_agente.frx":041A
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Invertir Selecci?n Alt + V"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   780
         Picture         =   "frmreporte_concecutivo_tipo_documento_agente.frx":04EC
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Marcar (Enter)"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         Picture         =   "frmreporte_concecutivo_tipo_documento_agente.frx":0736
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   225
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_documentos 
         Height          =   2670
         Left            =   45
         TabIndex        =   23
         Top             =   690
         Width           =   5610
         _ExtentX        =   9895
         _ExtentY        =   4710
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Docto"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Clase"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nombre"
            Object.Width           =   6527
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "  Agentes "
      Height          =   4335
      Left            =   90
      TabIndex        =   8
      Top             =   465
      Width           =   5625
      Begin VB.Frame Frame6 
         Height          =   120
         Left            =   30
         TabIndex        =   14
         Top             =   540
         Width           =   5565
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   450
         Picture         =   "frmreporte_concecutivo_tipo_documento_agente.frx":094C
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   120
         Picture         =   "frmreporte_concecutivo_tipo_documento_agente.frx":0B62
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1110
         Picture         =   "frmreporte_concecutivo_tipo_documento_agente.frx":0C64
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Invertir Selecci?n Alt + V"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   780
         Picture         =   "frmreporte_concecutivo_tipo_documento_agente.frx":0D36
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Marcar (Enter)"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         Picture         =   "frmreporte_concecutivo_tipo_documento_agente.frx":0F80
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   225
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_agentes 
         Height          =   3585
         Left            =   45
         TabIndex        =   15
         Top             =   690
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   6324
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
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmreporte_concecutivo_tipo_documento_agente.frx":1196
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   75
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11280
      Picture         =   "frmreporte_concecutivo_tipo_documento_agente.frx":1298
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   75
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   825
      Left            =   5835
      TabIndex        =   0
      Top             =   3975
      Width           =   5730
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   3645
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   300
         Width           =   1140
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   3105
         TabIndex        =   4
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   1020
         TabIndex        =   3
         Top             =   375
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   75
      TabIndex        =   7
      Top             =   345
      Width           =   11550
   End
End
Attribute VB_Name = "frmreporte_concecutivo_tipo_documento_agente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_dolares_Click()
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
            var_cadena_agentes = ""
            
            For var_j = 1 To lv_agentes.ListItems.Count
                lv_agentes.ListItems.Item(var_j).Selected = True
                If lv_agentes.selectedItem.SubItems(2) = "*" Then
                   If var_cadena_agentes = "" Then
                      var_cadena_agentes = "(VCHA_AGE_AGENTE_ID = '" + lv_agentes.selectedItem + "'"
                   Else
                      var_cadena_agentes = var_cadena_agentes + " OR VCHA_AGE_AGENTE_ID = '" + lv_agentes.selectedItem + "'"
                   End If
                End If
            Next var_j
            var_cadena_documentos = ""
            For var_j = 1 To lv_documentos.ListItems.Count
                lv_documentos.ListItems.Item(var_j).Selected = True
                If lv_documentos.selectedItem.SubItems(3) = "*" Then
                   If var_cadena_documentos = "" Then
                      var_cadena_documentos = "((vcha_car_documento = '" + Me.lv_documentos.selectedItem + "' and vcha_car_clase_id = '" + Me.lv_documentos.selectedItem.SubItems(1) + "') "
                   Else
                      var_cadena_documentos = var_cadena_documentos + " or (vcha_Car_documento = '" + Me.lv_documentos.selectedItem + "' and vcha_car_clase_id = '" + Me.lv_documentos.selectedItem.SubItems(1) + "')"
                   End If
                End If
            Next var_j
            If Trim(var_cadena_agentes) = "" Then
               MsgBox "No se a seleccionado ning?n agente", vbOKOnly, "ATENCION"
            Else
               If Trim(var_cadena_documentos) = "" Then
                  MsgBox "No se a seleccionado ning?n documento", vbOKOnly, "ATENCION"
               Else
                  var_cadena_agentes = var_cadena_agentes + ")"
                  var_cadena_documentos = var_cadena_documentos + ")"
                  cnn.BeginTrans
                  rs.Open "select max(inte_temp_consecutivo) as numero from TB_TEMP_CONSECUTIVO_CARTERA", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
                  Else
                     var_consecutivo = 0
                  End If
                  var_consecutivo = var_consecutivo + 1
                  rs.Close
                  rs.Open "insert into TB_TEMP_CONSECUTIVO_CARTERA (inte_temp_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
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
                  'Me.Text1 = "select " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + "-.000001, VCHA_EMP_EMPRESA_ID, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, VCHA_CLI_CLAVE_ID from tb_encabezado_cartera where dtim_car_fecha >= " + var_fecha_inicio + " and dtim_car_fecha <= " + var_fecha_fin + "-.0000001 and char_car_afectacion = '+' AND " + var_cadena_agentes + " AND " + var_cadena_documentos
                  rs.Open " INSERT INTO TB_TEMP_CONSECUTIVO_CARTERA ( INTE_TEMP_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_EMP_EMPRESA_ID, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, VCHA_CLI_CLAVE_ID) select " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + "-.000001, VCHA_EMP_EMPRESA_ID, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, VCHA_CLI_CLAVE_ID from tb_encabezado_cartera where dtim_car_fecha >= " + var_fecha_inicio + " and dtim_car_fecha <= " + var_fecha_fin + "-.0000001 and char_car_afectacion = '+' AND " + var_cadena_agentes + " AND " + var_cadena_documentos, cnn, adOpenDynamic, adLockOptimistic
               
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
            
                  If var_empresa = "18" Then
                     var_si = MsgBox("?Deseas el reporte dividido por agente?", vbYesNo, "ATENCION")
                  Else
                     var_si = 6
                  End If
                  If var_si = 6 Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_cartera_consecutivo_CARGOS_2_dolares.rpt")
                     reporte.RecordSelectionFormula = "{VW_CARTERA_CONSECUTIVO_CARGOS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_CARTERA_CONSECUTIVO_CARGOS.INTE_TEMP_CONSECUTIVO} = " + CStr(var_consecutivo)
                     frmvistasprevias.cr.ReportSource = reporte
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     frmvistasprevias.cr.ViewReport
                     frmvistasprevias.Caption = "Reporte Consecutivo de Movimientos por Agente (cargos)"
                     frmvistasprevias.Show 1
                     Set reporte = Nothing
                     var_si = MsgBox("?Desea importar el reporte?", vbYesNo, "ATENCION")
                     If var_si = 6 Then
                        Set reporte = appl.OpenReport(App.Path + "\rep_cartera_consecutivo_CARGOS_2_dolares.rpt")
                        reporte.RecordSelectionFormula = "{VW_CARTERA_CONSECUTIVO_CARGOS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_CARTERA_CONSECUTIVO_CARGOS.INTE_TEMP_CONSECUTIVO} = " + CStr(var_consecutivo)
                        For ntablas = 1 To reporte.Database.Tables.Count
                           reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                        Next ntablas
                        reporte.ExportOptions.FormatType = crEFTExcel80
                        reporte.ExportOptions.DestinationType = crEDTDiskFile
                        archivo = "c:\reportessid\reporte_consecutivo_movimientos_cargos" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                        reporte.ExportOptions.DiskFileName = archivo
                        reporte.Export False
                        Set reporte = Nothing
                        MsgBox "Se a terminado de guardar el archivo " + archivo
                     End If
                     rs.Open "delete from TB_TEMP_CONSECUTIVO_CARTERA where INTE_TEMP_CONSECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  Else
                     Set reporte = appl.OpenReport(App.Path + "\rep_cartera_consecutivo_CARGOS_2_NODIVIDIDO_dolares.rpt")
                     reporte.RecordSelectionFormula = "{VW_CARTERA_CONSECUTIVO_CARGOS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_CARTERA_CONSECUTIVO_CARGOS.INTE_TEMP_CONSECUTIVO} = " + CStr(var_consecutivo)
                     frmvistasprevias.cr.ReportSource = reporte
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     frmvistasprevias.cr.ViewReport
                     frmvistasprevias.Caption = "Reporte Consecutivo de Movimientos por Agente (cargos)"
                     frmvistasprevias.Show 1
                     Set reporte = Nothing
                     var_si = MsgBox("?Desea importar el reporte?", vbYesNo, "ATENCION")
                     If var_si = 6 Then
                        Set reporte = appl.OpenReport(App.Path + "\rep_cartera_consecutivo_CARGOS_2_NO_DIVIDIDO_dolares.rpt")
                        reporte.RecordSelectionFormula = "{VW_CARTERA_CONSECUTIVO_CARGOS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_CARTERA_CONSECUTIVO_CARGOS.INTE_TEMP_CONSECUTIVO} = " + CStr(var_consecutivo)
                        For ntablas = 1 To reporte.Database.Tables.Count
                           reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                        Next ntablas
                        reporte.ExportOptions.FormatType = crEFTExcel80
                        reporte.ExportOptions.DestinationType = crEDTDiskFile
                        archivo = "c:\reportessid\reporte_consecutivo_movimientos_cargos" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                        reporte.ExportOptions.DiskFileName = archivo
                        reporte.Export False
                        Set reporte = Nothing
                        MsgBox "Se a terminado de guardar el archivo " + archivo
                     End If
                     rs.Open "delete from TB_TEMP_CONSECUTIVO_CARTERA where INTE_TEMP_CONSECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rs.Open " INSERT INTO TB_TEMP_CONSECUTIVO_CARTERA ( INTE_TEMP_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_EMP_EMPRESA_ID, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, VCHA_CLI_CLAVE_ID) select " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + "-.000001, VCHA_EMP_EMPRESA_ID, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, VCHA_CLI_CLAVE_ID from tb_encabezado_cartera where dtim_car_fecha >= " + var_fecha_inicio + " and dtim_car_fecha <= " + var_fecha_fin + "-.0000001 and char_car_afectacion = '-' AND " + var_cadena_agentes + " AND " + var_cadena_documentos, cnn, adOpenDynamic, adLockOptimistic
            
                  Set reporte = appl.OpenReport(App.Path + "\rep_cartera_consecutivo_abonos_2_dolares.rpt")
                  reporte.RecordSelectionFormula = "{VW_CONSECUTIVO_CARTERA_ABONOS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_CONSECUTIVO_CARTERA_ABONOS.INTE_TEMP_CONSECUTIVO} = " + CStr(var_consecutivo)
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Reporte Consecutivo de Movimientos por Agente (abonos)"
                  frmvistasprevias.Show 1
                  Set reporte = Nothing
                  var_si = MsgBox("?Desea importar el reporte?", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_cartera_consecutivo_abonos_2_dolares.rpt")
                     reporte.RecordSelectionFormula = "{VW_CONSECUTIVO_CARTERA_ABONOS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_CONSECUTIVO_CARTERA_ABONOS.INTE_TEMP_CONSECUTIVO} = " + CStr(var_consecutivo)
                     For ntablas = 1 To reporte.Database.Tables.Count
                        reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     reporte.ExportOptions.FormatType = crEFTExcel80
                     reporte.ExportOptions.DestinationType = crEDTDiskFile
                     archivo = "c:\reportessid\reporte_consecutivo_movimientos_abonos" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     reporte.ExportOptions.DiskFileName = archivo
                     reporte.Export False
                     Set reporte = Nothing
                     MsgBox "Se a terminado de guardar el archivo " + archivo
                  End If
                  rs.Open "delete from TB_TEMP_CONSECUTIVO_CARTERA where INTE_TEMp_COnsECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
               End If
            End If
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
            var_cadena_agentes = ""
            
            For var_j = 1 To lv_agentes.ListItems.Count
                lv_agentes.ListItems.Item(var_j).Selected = True
                If lv_agentes.selectedItem.SubItems(2) = "*" Then
                   If var_cadena_agentes = "" Then
                      var_cadena_agentes = "(VCHA_AGE_AGENTE_ID = '" + lv_agentes.selectedItem + "'"
                   Else
                      var_cadena_agentes = var_cadena_agentes + " OR VCHA_AGE_AGENTE_ID = '" + lv_agentes.selectedItem + "'"
                   End If
                End If
            Next var_j
            var_cadena_documentos = ""
            For var_j = 1 To lv_documentos.ListItems.Count
                lv_documentos.ListItems.Item(var_j).Selected = True
                If lv_documentos.selectedItem.SubItems(3) = "*" Then
                   If var_cadena_documentos = "" Then
                      var_cadena_documentos = "((vcha_car_documento = '" + Me.lv_documentos.selectedItem + "' and vcha_car_clase_id = '" + Me.lv_documentos.selectedItem.SubItems(1) + "') "
                   Else
                      var_cadena_documentos = var_cadena_documentos + " or (vcha_Car_documento = '" + Me.lv_documentos.selectedItem + "' and vcha_car_clase_id = '" + Me.lv_documentos.selectedItem.SubItems(1) + "')"
                   End If
                End If
            Next var_j
            If Trim(var_cadena_agentes) = "" Then
               MsgBox "No se a seleccionado ning?n agente", vbOKOnly, "ATENCION"
            Else
               If Trim(var_cadena_documentos) = "" Then
                  MsgBox "No se a seleccionado ning?n documento", vbOKOnly, "ATENCION"
               Else
                  var_cadena_agentes = var_cadena_agentes + ")"
                  var_cadena_documentos = var_cadena_documentos + ")"
                  cnn.BeginTrans
                  rs.Open "select max(inte_temp_consecutivo) as numero from TB_TEMP_CONSECUTIVO_CARTERA", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
                  Else
                     var_consecutivo = 0
                  End If
                  var_consecutivo = var_consecutivo + 1
                  rs.Close
                  rs.Open "insert into TB_TEMP_CONSECUTIVO_CARTERA (inte_temp_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
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
                  'Me.Text1 = "select " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + "-.000001, VCHA_EMP_EMPRESA_ID, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, VCHA_CLI_CLAVE_ID from tb_encabezado_cartera where dtim_car_fecha >= " + var_fecha_inicio + " and dtim_car_fecha <= " + var_fecha_fin + "-.0000001 and char_car_afectacion = '+' AND " + var_cadena_agentes + " AND " + var_cadena_documentos
                  rs.Open " INSERT INTO TB_TEMP_CONSECUTIVO_CARTERA ( INTE_TEMP_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_EMP_EMPRESA_ID, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, VCHA_CLI_CLAVE_ID) select " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + "-.000001, VCHA_EMP_EMPRESA_ID, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, VCHA_CLI_CLAVE_ID from tb_encabezado_cartera where dtim_car_fecha >= " + var_fecha_inicio + " and dtim_car_fecha <= " + var_fecha_fin + "-.0000001 and char_car_afectacion = '+' AND " + var_cadena_agentes + " AND " + var_cadena_documentos, cnn, adOpenDynamic, adLockOptimistic
                  
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
            
                  If var_empresa = "18" Then
                     var_si = MsgBox("?Deseas el reporte dividido por agente?", vbYesNo, "ATENCION")
                  Else
                     var_si = 6
                  End If
                  If var_si = 6 Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_cartera_consecutivo_CARGOS_2.rpt")
                     reporte.RecordSelectionFormula = "{VW_CARTERA_CONSECUTIVO_CARGOS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_CARTERA_CONSECUTIVO_CARGOS.INTE_TEMP_CONSECUTIVO} = " + CStr(var_consecutivo)
                     frmvistasprevias.cr.ReportSource = reporte
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     frmvistasprevias.cr.ViewReport
                     frmvistasprevias.Caption = "Reporte Consecutivo de Movimientos por Agente (cargos)"
                     frmvistasprevias.Show 1
                     Set reporte = Nothing
                     var_si = MsgBox("?Desea importar el reporte?", vbYesNo, "ATENCION")
                     If var_si = 6 Then
                        Set reporte = appl.OpenReport(App.Path + "\rep_cartera_consecutivo_CARGOS_2.rpt")
                        reporte.RecordSelectionFormula = "{VW_CARTERA_CONSECUTIVO_CARGOS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_CARTERA_CONSECUTIVO_CARGOS.INTE_TEMP_CONSECUTIVO} = " + CStr(var_consecutivo)
                        For ntablas = 1 To reporte.Database.Tables.Count
                           reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                        Next ntablas
                        reporte.ExportOptions.FormatType = crEFTExcel80
                        reporte.ExportOptions.DestinationType = crEDTDiskFile
                        archivo = "c:\reportessid\reporte_consecutivo_movimientos_cargos" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                        reporte.ExportOptions.DiskFileName = archivo
                        reporte.Export False
                        Set reporte = Nothing
                        MsgBox "Se a terminado de guardar el archivo " + archivo
                     End If
                     rs.Open "delete from TB_TEMP_CONSECUTIVO_CARTERA where INTE_TEMP_CONSECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  Else
                     Set reporte = appl.OpenReport(App.Path + "\rep_cartera_consecutivo_CARGOS_2_NODIVIDIDO.rpt")
                     reporte.RecordSelectionFormula = "{VW_CARTERA_CONSECUTIVO_CARGOS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_CARTERA_CONSECUTIVO_CARGOS.INTE_TEMP_CONSECUTIVO} = " + CStr(var_consecutivo)
                     frmvistasprevias.cr.ReportSource = reporte
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     frmvistasprevias.cr.ViewReport
                     frmvistasprevias.Caption = "Reporte Consecutivo de Movimientos por Agente (cargos)"
                     frmvistasprevias.Show 1
                     Set reporte = Nothing
                     var_si = MsgBox("?Desea importar el reporte?", vbYesNo, "ATENCION")
                     If var_si = 6 Then
                        Set reporte = appl.OpenReport(App.Path + "\rep_cartera_consecutivo_CARGOS_2_NO_DIVIDIDO.rpt")
                        reporte.RecordSelectionFormula = "{VW_CARTERA_CONSECUTIVO_CARGOS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_CARTERA_CONSECUTIVO_CARGOS.INTE_TEMP_CONSECUTIVO} = " + CStr(var_consecutivo)
                        For ntablas = 1 To reporte.Database.Tables.Count
                           reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                        Next ntablas
                        reporte.ExportOptions.FormatType = crEFTExcel80
                        reporte.ExportOptions.DestinationType = crEDTDiskFile
                        archivo = "c:\reportessid\reporte_consecutivo_movimientos_cargos" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                        reporte.ExportOptions.DiskFileName = archivo
                        reporte.Export False
                        Set reporte = Nothing
                        MsgBox "Se a terminado de guardar el archivo " + archivo
                     End If
                     rs.Open "delete from TB_TEMP_CONSECUTIVO_CARTERA where INTE_TEMP_CONSECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rs.Open " INSERT INTO TB_TEMP_CONSECUTIVO_CARTERA ( INTE_TEMP_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_EMP_EMPRESA_ID, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, VCHA_CLI_CLAVE_ID) select " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + "-.000001, VCHA_EMP_EMPRESA_ID, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, VCHA_CLI_CLAVE_ID from tb_encabezado_cartera where dtim_car_fecha >= " + var_fecha_inicio + " and dtim_car_fecha <= " + var_fecha_fin + "-.0000001 and char_car_afectacion = '-' AND " + var_cadena_agentes + " AND " + var_cadena_documentos, cnn, adOpenDynamic, adLockOptimistic
            
                  Set reporte = appl.OpenReport(App.Path + "\rep_cartera_consecutivo_abonos_2.rpt")
                  reporte.RecordSelectionFormula = "{VW_CONSECUTIVO_CARTERA_ABONOS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_CONSECUTIVO_CARTERA_ABONOS.INTE_TEMP_CONSECUTIVO} = " + CStr(var_consecutivo)
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Reporte Consecutivo de Movimientos por Agente (abonos)"
                  frmvistasprevias.Show 1
                  Set reporte = Nothing
                  var_si = MsgBox("?Desea importar el reporte?", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_cartera_consecutivo_abonos_2.rpt")
                     reporte.RecordSelectionFormula = "{VW_CONSECUTIVO_CARTERA_ABONOS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_CONSECUTIVO_CARTERA_ABONOS.INTE_TEMP_CONSECUTIVO} = " + CStr(var_consecutivo)
                     For ntablas = 1 To reporte.Database.Tables.Count
                        reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     reporte.ExportOptions.FormatType = crEFTExcel80
                     reporte.ExportOptions.DestinationType = crEDTDiskFile
                     archivo = "c:\reportessid\reporte_consecutivo_movimientos_abonos" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     reporte.ExportOptions.DiskFileName = archivo
                     reporte.Export False
                     Set reporte = Nothing
                     MsgBox "Se a terminado de guardar el archivo " + archivo
                  End If
                  rs.Open "delete from TB_TEMP_CONSECUTIVO_CARTERA where INTE_TEMp_COnsECUTIVO  = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
               End If
            End If
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
   n = lv_documentos.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_documentos.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_documentos.selectedItem.SubItems(3) = "" And var_rellena = True Then
         lv_documentos.selectedItem.SubItems(3) = "*"
         lv_documentos.ListItems.Item(i).Bold = True
         lv_documentos.ListItems.Item(i).ForeColor = &HFF0000
         lv_documentos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_documentos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_documentos.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_documentos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_documentos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_documentos.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_documentos.selectedItem.SubItems(2) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_documentos.selectedItem.SubItems(3) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub Command2_Click()
   i = lv_documentos.selectedItem.Index
   If lv_documentos.selectedItem.SubItems(3) = "*" Then
      lv_documentos.selectedItem.SubItems(3) = ""
      lv_documentos.ListItems.Item(i).Bold = False
      lv_documentos.ListItems.Item(i).ForeColor = &H80000012
      lv_documentos.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_documentos.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_documentos.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_documentos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_documentos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_documentos.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      lv_documentos.Refresh
   Else
      lv_documentos.selectedItem.SubItems(3) = "*"
      lv_documentos.ListItems.Item(i).Bold = True
      lv_documentos.ListItems.Item(i).ForeColor = &HFF0000
      lv_documentos.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_documentos.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_documentos.ListItems.Item(i).ListSubItems(3).Bold = True
      lv_documentos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_documentos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_documentos.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
      lv_documentos.Refresh
   End If
End Sub

Private Sub Command3_Click()
   n = lv_documentos.ListItems.Count
   For i = 1 To n
      lv_documentos.ListItems.Item(i).Selected = True
      If lv_documentos.selectedItem.SubItems(3) = "*" Then
         lv_documentos.selectedItem.SubItems(3) = ""
         lv_documentos.ListItems.Item(i).Bold = False
         lv_documentos.ListItems.Item(i).ForeColor = &H80000012
         lv_documentos.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_documentos.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_documentos.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_documentos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_documentos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_documentos.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      Else
         lv_documentos.selectedItem.SubItems(3) = "*"
         lv_documentos.ListItems.Item(i).Bold = True
         lv_documentos.ListItems.Item(i).ForeColor = &HFF0000
         lv_documentos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_documentos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_documentos.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_documentos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_documentos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_documentos.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub Command4_Click()
   n = lv_documentos.ListItems.Count
   For i = 1 To n
      lv_documentos.ListItems.Item(i).Selected = True
      lv_documentos.selectedItem.SubItems(3) = ""
      lv_documentos.ListItems.Item(i).Bold = False
      lv_documentos.ListItems.Item(i).ForeColor = &H80000012
      lv_documentos.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_documentos.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_documentos.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_documentos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_documentos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_documentos.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
   Next i
   lv_agentes.Refresh
End Sub

Private Sub Command5_Click()
   n = lv_documentos.ListItems.Count
   For i = 1 To n
      lv_documentos.ListItems.Item(i).Selected = True
      lv_documentos.selectedItem.SubItems(3) = "*"
      lv_documentos.ListItems.Item(i).Bold = True
      lv_documentos.ListItems.Item(i).ForeColor = &HFF0000
      lv_documentos.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_documentos.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_documentos.ListItems.Item(i).ListSubItems(3).Bold = True
      lv_documentos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_documentos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_documentos.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
   Next i
   lv_documentos.Refresh
End Sub

Private Sub Command6_Click()

End Sub

Private Sub Form_Load()
   Left = 0
   Top = 1300
   rs.Open "select * from tb_agentes where vcha_emp_empresa_id = '" + var_empresa + "' or vcha_age_agente_id = '00083' or vcha_age_Agente_id = '00100'  order by vcha_age_nombre", cnn, adOpenDynamic, adLockOptimistic
   numero_items_ALMACENES = 0
   While Not rs.EOF
      Set list_item = lv_agentes.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
      list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
      list_item.SubItems(2) = ""
      rs.MoveNext:
   Wend
   rs.Close
   rs.Open "select distinct vcha_Car_documento, vcha_car_clase_id, vcha_car_nombre from tb_clases_cartera where len(vcha_car_nombre) > 0 order by vcha_car_nombre", cnn, adOpenDynamic, adLockOptimistic
   numero_items_ALMACENES = 0
   While Not rs.EOF
      Set list_item = lv_documentos.ListItems.Add(, , rs!vcha_Car_documento)
      list_item.SubItems(1) = IIf(IsNull(rs!vcha_Car_clase_id), "", rs!vcha_Car_clase_id)
      list_item.SubItems(2) = IIf(IsNull(rs!vcha_Car_nombre), "", rs!vcha_Car_nombre)
      list_item.SubItems(3) = ""
      rs.MoveNext:
   Wend
   rs.Close
   
   
   If lv_agentes.ListItems.Count > 7 Then
      lv_agentes.ColumnHeaders(2).Width = 4220
   Else
      lv_agentes.ColumnHeaders(2).Width = 4499.71
   End If
   Me.txt_fin = Date
   Me.txt_inicio = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
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

Private Sub lv_documentos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_documentos, ColumnHeader)
End Sub

Private Sub lv_documentos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      i = lv_documentos.selectedItem.Index
      If lv_documentos.selectedItem.SubItems(3) = "*" Then
         lv_documentos.selectedItem.SubItems(3) = ""
         lv_documentos.ListItems.Item(i).Bold = False
         lv_documentos.ListItems.Item(i).ForeColor = &H80000012
         lv_documentos.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_documentos.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_documentos.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_documentos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_documentos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_documentos.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
         lv_documentos.Refresh
      Else
         lv_documentos.selectedItem.SubItems(3) = "*"
         lv_documentos.ListItems.Item(i).Bold = True
         lv_documentos.ListItems.Item(i).ForeColor = &HFF0000
         lv_documentos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_documentos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_documentos.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_documentos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_documentos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_documentos.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_documentos.Refresh
      End If
   End If
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
