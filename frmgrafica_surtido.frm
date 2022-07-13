VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmgrafica_surtido 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordenes de Surtido"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   " Reporte "
      Height          =   720
      Left            =   5835
      TabIndex        =   9
      Top             =   6570
      Width           =   5730
      Begin VB.CommandButton Command1 
         Height          =   345
         Left            =   2475
         Picture         =   "frmgrafica_surtido.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Actualiza Grafica"
         Top             =   270
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Height          =   345
         Left            =   2850
         Picture         =   "frmgrafica_surtido.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Exportar Gráfica"
         Top             =   270
         Width           =   375
      End
   End
   Begin MSComCtl2.MonthView mes 
      Height          =   2370
      Left            =   2190
      TabIndex        =   8
      Top             =   4395
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   83034113
      CurrentDate     =   38148
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   720
      Left            =   90
      TabIndex        =   1
      Top             =   6570
      Width           =   5670
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   3690
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   270
         Width           =   1095
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   1110
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   270
         Width           =   1080
      End
      Begin VB.CommandButton Command11 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2220
         Picture         =   "frmgrafica_surtido.frx":0414
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Fecha Inicial"
         Top             =   270
         Width           =   330
      End
      Begin VB.CommandButton Command12 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4830
         Picture         =   "frmgrafica_surtido.frx":1686
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Fecha Final"
         Top             =   270
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   3375
         TabIndex        =   7
         Top             =   330
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   660
         TabIndex        =   6
         Top             =   330
         Width           =   420
      End
   End
   Begin MSComctlLib.ListView lv_grafica 
      Height          =   6525
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   11509
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "O.S."
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fecha"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Agente"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Ruta"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Cliente"
         Object.Width           =   3087
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Surtir"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Surtido"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Empacado"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Porcentaje"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "tipo"
         Object.Width           =   1058
      EndProperty
   End
End
Attribute VB_Name = "frmgrafica_surtido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_fecha_inicio As String
Dim var_fecha_fin As String
Dim var_tipo_mes As Integer

Private Sub Command1_Click()
   If IsDate(Me.txt_inicio) Then
      If IsDate(Me.txt_fin) Then
         lv_grafica.ListItems.Clear
         
         var_fecha_fin_1 = CDate(txt_fin) + 1
         var_fecha_inicio = CDate(txt_inicio)
         
         var_dia = CStr(Day(var_fecha_inicio))
         var_mes = CStr(Month(var_fecha_inicio))
         var_año = CStr(Year(var_fecha_inicio))
          
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
         
         var_cadena = "SELECT dbo.TB_ENC_ORDEN_SURTIDO.dtim_ors_fecha_carga, dbo.TB_DET_ORDEN_SURTIDO.char_ped_tipo, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO , SUM(dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIR) AS cantidad_surtir, dbo.TB_RUTAS.VCHA_RUT_NOMBRE, SUM(dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIDA) AS cantidad_surtida,  "
         var_cadena = var_cadena + " SUM(dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_EMPACADA) As cantidad_empacada FROM dbo.TB_ENC_ORDEN_SURTIDO INNER JOIN dbo.TB_ENCABEZADO_PEDIDOS ON dbo.TB_ENC_ORDEN_SURTIDO.INTE_PED_NUMERO = dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_NUMERO INNER JOIN dbo.TB_DET_ORDEN_SURTIDO ON dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO = dbo.TB_DET_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO INNER JOIN "
         var_cadena = var_cadena + " dbo.TB_CLIENTES ON dbo.TB_ENC_ORDEN_SURTIDO.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID INNER JOIN dbo.TB_AGENTES ON dbo.TB_ENCABEZADO_PEDIDOS.VCHA_AGE_AGENTE_ID = dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID INNER JOIN dbo.TB_RUTAS ON dbo.TB_CLIENTES.VCHA_RUT_RUTA_ID = dbo.TB_RUTAS.VCHA_RUT_RUTA_ID WHERE (dbo.TB_ENC_orden_surtido.DTIM_ors_FECHA_carga >= " + var_fecha_inicio + ")  and (dbo.TB_ENC_orden_surtido.DTIM_ors_FECHA_carga <= " + var_fecha_fin + " -.000001)"
         var_cadena = var_cadena + " GROUP BY dbo.TB_ENC_ORDEN_SURTIDO.dtim_ors_fecha_carga, dbo.TB_DET_ORDEN_SURTIDO.char_ped_tipo, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO, dbo.TB_RUTAS.VCHA_RUT_NOMBRE ORDER BY dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_RUTAS.VCHA_RUT_NOMBRE, dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO"
         
         Text1 = var_cadena
         rs.Open var_cadena, cnn, adOpenDynamic, adLockBatchOptimistic
         numero_items_titulares = 0
         While Not rs.EOF
               If rs!cantidad_surtir > 0 Then
                  Set list_item = Me.lv_grafica.ListItems.Add(, , rs!INTE_ORS_ORDEN_SURTIDO)
                  list_item.SubItems(1) = Format(IIf(IsNull(rs!DTIM_ORS_FECHA_CARGA), "", rs!DTIM_ORS_FECHA_CARGA), "Short Date")
                  list_item.SubItems(2) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
                  list_item.SubItems(3) = IIf(IsNull(rs!vcha_rut_nombre), "", rs!vcha_rut_nombre)
                  list_item.SubItems(4) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
                  list_item.SubItems(5) = Format(IIf(IsNull(rs!cantidad_surtir), 0, rs!cantidad_surtir), "###,###,##0")
                  list_item.SubItems(6) = Format(IIf(IsNull(rs!cantidad_surtida), 0, rs!cantidad_surtida), "###,###,##0")
                  list_item.SubItems(7) = Format(IIf(IsNull(rs!cantidad_empacada), 0, rs!cantidad_empacada), "###,###,##0")
                  If rs!cantidad_surtir = 0 Then
                     var_porcentaje = 0
                  Else
                     var_porcentaje = (IIf(IsNull(rs!cantidad_surtida), 0, rs!cantidad_surtida) / IIf(IsNull(rs!cantidad_surtir), 0, rs!cantidad_surtir)) * 100
                  End If
                  list_item.SubItems(8) = Format(CStr(var_porcentaje), "###,###,##0")
                  list_item.SubItems(9) = IIf(IsNull(rs!char_ped_tipo), 0, rs!char_ped_tipo)
               End If
               rs.MoveNext:
               numero_items_titulares = numero_items_titulares + 1
         Wend
         rs.Close
         For var_i = 1 To lv_grafica.ListItems.Count
             lv_grafica.ListItems(var_i).Selected = True
             If (lv_grafica.selectedItem.SubItems(8) * 1) > 25 Then
                lv_grafica.ListItems.Item(var_i).ForeColor = vbBlue
                'lv_grafica.ListItems(var_i).ListSubItems(1).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(1).ForeColor = vbBlue
                lv_grafica.ListItems(var_i).ListSubItems(2).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(2).ForeColor = vbBlue
                lv_grafica.ListItems(var_i).ListSubItems(3).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(3).ForeColor = vbBlue
                lv_grafica.ListItems(var_i).ListSubItems(4).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(4).ForeColor = vbBlue
                lv_grafica.ListItems(var_i).ListSubItems(5).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(5).ForeColor = vbBlue
                lv_grafica.ListItems(var_i).ListSubItems(6).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(6).ForeColor = vbBlue
                lv_grafica.ListItems(var_i).ListSubItems(7).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(7).ForeColor = vbBlue
                lv_grafica.ListItems(var_i).ListSubItems(8).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(8).ForeColor = vbBlue
                lv_grafica.ListItems(var_i).ListSubItems(9).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(9).ForeColor = vbBlue
                lv_grafica.selectedItem.Bold = True
             End If
             If (lv_grafica.selectedItem.SubItems(8) * 1) > 50 Then
                lv_grafica.ListItems.Item(var_i).ForeColor = &HC000C0
                'lv_grafica.ListItems(var_i).ListSubItems(1).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(1).ForeColor = &HC000C0
                lv_grafica.ListItems(var_i).ListSubItems(2).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(2).ForeColor = &HC000C0
                lv_grafica.ListItems(var_i).ListSubItems(3).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(3).ForeColor = &HC000C0
                lv_grafica.ListItems(var_i).ListSubItems(4).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(4).ForeColor = &HC000C0
                lv_grafica.ListItems(var_i).ListSubItems(5).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(5).ForeColor = &HC000C0
                lv_grafica.ListItems(var_i).ListSubItems(6).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(6).ForeColor = &HC000C0
                lv_grafica.ListItems(var_i).ListSubItems(7).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(7).ForeColor = &HC000C0
                lv_grafica.ListItems(var_i).ListSubItems(8).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(8).ForeColor = &HC000C0
                lv_grafica.ListItems(var_i).ListSubItems(9).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(9).ForeColor = &HC000C0
                lv_grafica.selectedItem.Bold = True
             End If
             If (lv_grafica.selectedItem.SubItems(8) * 1) = 100 Then
                lv_grafica.ListItems.Item(var_i).ForeColor = vbRed
                'lv_grafica.ListItems(var_i).ListSubItems(1).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(1).ForeColor = vbRed
                lv_grafica.ListItems(var_i).ListSubItems(2).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(2).ForeColor = vbRed
                lv_grafica.ListItems(var_i).ListSubItems(3).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(3).ForeColor = vbRed
                lv_grafica.ListItems(var_i).ListSubItems(4).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(4).ForeColor = vbRed
                lv_grafica.ListItems(var_i).ListSubItems(5).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(5).ForeColor = vbRed
                lv_grafica.ListItems(var_i).ListSubItems(6).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(6).ForeColor = vbRed
                lv_grafica.ListItems(var_i).ListSubItems(7).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(7).ForeColor = vbRed
                lv_grafica.ListItems(var_i).ListSubItems(8).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(8).ForeColor = vbRed
                lv_grafica.ListItems(var_i).ListSubItems(9).Bold = True
                lv_grafica.ListItems(var_i).ListSubItems(9).ForeColor = vbRed
                lv_grafica.selectedItem.Bold = True
             End If
          Next var_i
      Else
         MsgBox "Fecha final inorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de inicio inorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Command11_Click()
   If IsDate(Me.txt_inicio) Then
      Me.mes.Value = CDate(Me.txt_inicio)
   Else
      mes.Value = Date
   End If
   var_tipo_mes = 1
   mes.Visible = True
   mes.SetFocus
End Sub

Private Sub Command12_Click()
   If IsDate(Me.txt_fin) Then
      mes.Value = CDate(Me.txt_fin)
   Else
      mes.Value = Date
   End If
   var_tipo_mes = 2
   mes.Visible = True
   mes.SetFocus
End Sub

Private Sub Command2_Click()
   If IsDate(Me.txt_inicio) Then
      If IsDate(Me.txt_fin) Then
         var_fecha_fin_1 = CDate(txt_fin) + 1
         var_fecha_inicio = CDate(txt_inicio)
         
         var_dia = CStr(Day(var_fecha_inicio))
         var_mes = CStr(Month(var_fecha_inicio))
         var_año = CStr(Year(var_fecha_inicio))
          
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
         rs.Open "select max(inte_TEM_consecutivo) from TB_TEMP_REPORTE_GRAFICA_SURTIDO", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
         Else
            var_consecutivo = 1
         End If
         rs.Close
         rs.Open "insert into TB_TEMP_REPORTE_GRAFICA_SURTIDO (inte_TEM_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
         cnn.CommitTrans
         rs.Open "insert into TB_TEMP_REPORTE_GRAFICA_SURTIDO (inte_tem_consecutivo, dtim_tem_fecha_inicio, dtim_tem_fecha_fin, inte_ors_orden_surtido) select " + CStr(var_consecutivo) + ", " + var_fecha_inicio + "," + var_fecha_fin + "-.00001,inte_ors_orden_surtido from tb_enc_orden_surtido where dtim_ors_fecha_carga > = " + var_fecha_inicio + " and dtim_ors_fecha_carga <= " + var_fecha_fin + "-.00001", cnn, adOpenDynamic, adLockOptimistic
         If var_empresa = "31" Then
            Set reporte = appl.OpenReport(App.Path + "\rep_grafica_surtido_2.rpt")
         Else
            Set reporte = appl.OpenReport(App.Path + "\rep_grafica_surtido.rpt")
         End If
         reporte.RecordSelectionFormula = "{VW_REPORTE_GRAFICA_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         reporte.ExportOptions.FormatType = crEFTExcel80
         reporte.ExportOptions.DestinationType = crEDTDiskFile
         archivo = "c:\reportessid\Grafica_surtido" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
         reporte.ExportOptions.DiskFileName = archivo
         reporte.Export False
         Set reporte = Nothing
         MsgBox "Se a terminado de guardar el archivo " + archivo
         rs.Open "delete from TB_TEMP_REPORTE_GRAFICA_SURTIDO where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de inicio incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Form_Load()
   mes.Visible = False
   Top = 0
   Left = 0
   Me.txt_fin = Date
   Me.txt_inicio = Date
    var_fecha_fin_1 = Date + 1
    
    var_dia = CStr(Day(Date))
    var_mes = CStr(Month(Date))
    var_año = CStr(Year(Date))
    
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

   var_cadena = "SELECT dbo.TB_ENC_ORDEN_SURTIDO.dtim_ors_fecha_carga, dbo.TB_DET_ORDEN_SURTIDO.char_ped_tipo, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO , SUM(dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIR) AS cantidad_surtir, dbo.TB_RUTAS.VCHA_RUT_NOMBRE, SUM(dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_SURTIDA) AS cantidad_surtida,  "
   var_cadena = var_cadena + " SUM(dbo.TB_DET_ORDEN_SURTIDO.FLOA_ORS_CANTIDAD_EMPACADA) As cantidad_empacada FROM dbo.TB_ENC_ORDEN_SURTIDO INNER JOIN dbo.TB_ENCABEZADO_PEDIDOS ON dbo.TB_ENC_ORDEN_SURTIDO.INTE_PED_NUMERO = dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_NUMERO INNER JOIN dbo.TB_DET_ORDEN_SURTIDO ON dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO = dbo.TB_DET_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO INNER JOIN "
   var_cadena = var_cadena + " dbo.TB_CLIENTES ON dbo.TB_ENC_ORDEN_SURTIDO.VCHA_CLI_CLAVE_ID = dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID INNER JOIN dbo.TB_AGENTES ON dbo.TB_ENCABEZADO_PEDIDOS.VCHA_AGE_AGENTE_ID = dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID INNER JOIN dbo.TB_RUTAS ON dbo.TB_CLIENTES.VCHA_RUT_RUTA_ID = dbo.TB_RUTAS.VCHA_RUT_RUTA_ID WHERE (dbo.TB_ENC_orden_surtido.DTIM_ors_FECHA_carga >= " + var_fecha_inicio + ")  and (dbo.TB_ENC_orden_surtido.DTIM_ors_FECHA_carga <= " + var_fecha_fin + " -.000001)"
   var_cadena = var_cadena + " GROUP BY dbo.TB_ENC_ORDEN_SURTIDO.dtim_ors_fecha_carga,dbo.TB_DET_ORDEN_SURTIDO.char_ped_tipo, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO, dbo.TB_RUTAS.VCHA_RUT_NOMBRE ORDER BY dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_RUTAS.VCHA_RUT_NOMBRE, dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO"
   Text1 = var_cadena
   rs.Open var_cadena, cnn, adOpenDynamic, adLockBatchOptimistic
   numero_items_titulares = 0
   While Not rs.EOF
         If rs!cantidad_surtir > 0 Then
            Set list_item = Me.lv_grafica.ListItems.Add(, , rs!INTE_ORS_ORDEN_SURTIDO)
            list_item.SubItems(1) = Format(IIf(IsNull(rs!DTIM_ORS_FECHA_CARGA), "", rs!DTIM_ORS_FECHA_CARGA), "Short Date")
            list_item.SubItems(2) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            list_item.SubItems(3) = IIf(IsNull(rs!vcha_rut_nombre), "", rs!vcha_rut_nombre)
            list_item.SubItems(4) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            list_item.SubItems(5) = Format(IIf(IsNull(rs!cantidad_surtir), 0, rs!cantidad_surtir), "###,###,##0")
            list_item.SubItems(6) = Format(IIf(IsNull(rs!cantidad_surtida), 0, rs!cantidad_surtida), "###,###,##0")
            list_item.SubItems(7) = Format(IIf(IsNull(rs!cantidad_empacada), 0, rs!cantidad_empacada), "###,###,##0")
            If rs!cantidad_surtir = 0 Then
               var_porcentaje = 0
            Else
               var_porcentaje = (IIf(IsNull(rs!cantidad_surtida), 0, rs!cantidad_surtida) / IIf(IsNull(rs!cantidad_surtir), 0, rs!cantidad_surtir)) * 100
            End If
            list_item.SubItems(8) = Format(CStr(var_porcentaje), "###,###,##0")
            list_item.SubItems(9) = IIf(IsNull(rs!char_ped_tipo), 0, rs!char_ped_tipo)
         End If
          
         rs.MoveNext:
         numero_items_titulares = numero_items_titulares + 1
         
   Wend
   rs.Close
   For var_i = 1 To lv_grafica.ListItems.Count
       lv_grafica.ListItems(var_i).Selected = True
       If (lv_grafica.selectedItem.SubItems(8) * 1) > 25 Then
          lv_grafica.ListItems.Item(var_i).ForeColor = vbBlue
          'lv_grafica.ListItems(var_i).ListSubItems(1).Bold = True
          lv_grafica.ListItems(var_i).ListSubItems(1).ForeColor = vbBlue
          lv_grafica.ListItems(var_i).ListSubItems(2).Bold = True
          lv_grafica.ListItems(var_i).ListSubItems(2).ForeColor = vbBlue
          lv_grafica.ListItems(var_i).ListSubItems(3).Bold = True
          lv_grafica.ListItems(var_i).ListSubItems(3).ForeColor = vbBlue
          lv_grafica.ListItems(var_i).ListSubItems(4).Bold = True
          lv_grafica.ListItems(var_i).ListSubItems(4).ForeColor = vbBlue
          lv_grafica.ListItems(var_i).ListSubItems(5).Bold = True
          lv_grafica.ListItems(var_i).ListSubItems(5).ForeColor = vbBlue
          lv_grafica.ListItems(var_i).ListSubItems(6).Bold = True
          lv_grafica.ListItems(var_i).ListSubItems(6).ForeColor = vbBlue
          lv_grafica.ListItems(var_i).ListSubItems(7).Bold = True
          lv_grafica.ListItems(var_i).ListSubItems(7).ForeColor = vbBlue
          lv_grafica.ListItems(var_i).ListSubItems(8).Bold = True
          lv_grafica.ListItems(var_i).ListSubItems(8).ForeColor = vbBlue
          lv_grafica.ListItems(var_i).ListSubItems(9).Bold = True
          lv_grafica.ListItems(var_i).ListSubItems(9).ForeColor = vbBlue
          lv_grafica.selectedItem.Bold = True
       End If
       If (lv_grafica.selectedItem.SubItems(8) * 1) > 50 Then
          lv_grafica.ListItems.Item(var_i).ForeColor = &HC000C0
          'lv_grafica.ListItems(var_i).ListSubItems(1).Bold = True
          lv_grafica.ListItems(var_i).ListSubItems(1).ForeColor = &HC000C0
          lv_grafica.ListItems(var_i).ListSubItems(2).Bold = True
          lv_grafica.ListItems(var_i).ListSubItems(2).ForeColor = &HC000C0
          lv_grafica.ListItems(var_i).ListSubItems(3).Bold = True
          lv_grafica.ListItems(var_i).ListSubItems(3).ForeColor = &HC000C0
          lv_grafica.ListItems(var_i).ListSubItems(4).Bold = True
          lv_grafica.ListItems(var_i).ListSubItems(4).ForeColor = &HC000C0
          lv_grafica.ListItems(var_i).ListSubItems(5).Bold = True
          lv_grafica.ListItems(var_i).ListSubItems(5).ForeColor = &HC000C0
          lv_grafica.ListItems(var_i).ListSubItems(6).Bold = True
          lv_grafica.ListItems(var_i).ListSubItems(6).ForeColor = &HC000C0
          lv_grafica.ListItems(var_i).ListSubItems(7).Bold = True
          lv_grafica.ListItems(var_i).ListSubItems(7).ForeColor = &HC000C0
          lv_grafica.ListItems(var_i).ListSubItems(8).Bold = True
          lv_grafica.ListItems(var_i).ListSubItems(8).ForeColor = &HC000C0
          lv_grafica.ListItems(var_i).ListSubItems(9).Bold = True
          lv_grafica.ListItems(var_i).ListSubItems(9).ForeColor = &HC000C0
          lv_grafica.selectedItem.Bold = True
       End If
       If (lv_grafica.selectedItem.SubItems(8) * 1) = 100 Then
          lv_grafica.ListItems.Item(var_i).ForeColor = vbRed
          'lv_grafica.ListItems(var_i).ListSubItems(1).Bold = True
          lv_grafica.ListItems(var_i).ListSubItems(1).ForeColor = vbRed
          lv_grafica.ListItems(var_i).ListSubItems(2).Bold = True
          lv_grafica.ListItems(var_i).ListSubItems(2).ForeColor = vbRed
          lv_grafica.ListItems(var_i).ListSubItems(3).Bold = True
          lv_grafica.ListItems(var_i).ListSubItems(3).ForeColor = vbRed
          lv_grafica.ListItems(var_i).ListSubItems(4).Bold = True
          lv_grafica.ListItems(var_i).ListSubItems(4).ForeColor = vbRed
          lv_grafica.ListItems(var_i).ListSubItems(5).Bold = True
          lv_grafica.ListItems(var_i).ListSubItems(5).ForeColor = vbRed
          lv_grafica.ListItems(var_i).ListSubItems(6).Bold = True
          lv_grafica.ListItems(var_i).ListSubItems(6).ForeColor = vbRed
          lv_grafica.ListItems(var_i).ListSubItems(7).Bold = True
          lv_grafica.ListItems(var_i).ListSubItems(7).ForeColor = vbRed
          lv_grafica.ListItems(var_i).ListSubItems(8).Bold = True
          lv_grafica.ListItems(var_i).ListSubItems(8).ForeColor = vbRed
          lv_grafica.ListItems(var_i).ListSubItems(9).Bold = True
          lv_grafica.ListItems(var_i).ListSubItems(9).ForeColor = vbRed
          lv_grafica.selectedItem.Bold = True
       End If
   Next var_i
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_reporte_valuacion_devoluciones)
End Sub

Private Sub lv_grafica_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_grafica, ColumnHeader)
End Sub

Private Sub mes_DateDblClick(ByVal DateDblClicked As Date)
   If var_tipo_mes = 1 Then
      txt_inicio = mes.Value
      
   End If
   If var_tipo_mes = 2 Then
      txt_fin = mes.Value
   End If
   mes.Visible = False
End Sub

Private Sub mes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      mes.Visible = False
   End If
End Sub

Private Sub txt_fin_Change()
    Me.lv_grafica.ListItems.Clear
End Sub

Private Sub txt_inicio_Change()
    Me.lv_grafica.ListItems.Clear
End Sub
