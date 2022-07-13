VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmreporte_envios_tiendas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Envios a Tiendas"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   5295
   Begin MSComCtl2.MonthView mes 
      Height          =   2370
      Left            =   1380
      TabIndex        =   0
      Top             =   360
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   72548353
      CurrentDate     =   38148
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   720
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   5070
      Begin VB.CommandButton Command12 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4290
         Picture         =   "frmreporte_envios_tiendas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fecha Final"
         Top             =   270
         Width           =   330
      End
      Begin VB.CommandButton Command11 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1800
         Picture         =   "frmreporte_envios_tiendas.frx":1272
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fecha Inicial"
         Top             =   255
         Width           =   330
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   690
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   255
         Width           =   1080
      End
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   3150
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   315
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2835
         TabIndex        =   12
         Top             =   315
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   75
      TabIndex        =   6
      Top             =   360
      Width           =   5190
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmreporte_envios_tiendas.frx":24E4
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4845
      Picture         =   "frmreporte_envios_tiendas.frx":25E6
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Reporte "
      Height          =   1680
      Left            =   120
      TabIndex        =   1
      Top             =   495
      Width           =   5070
      Begin VB.OptionButton opt_detalle 
         Caption         =   "A detalle"
         Height          =   345
         Left            =   135
         TabIndex        =   15
         Top             =   1260
         Width           =   1950
      End
      Begin VB.OptionButton opt_agrupado_tienda 
         Caption         =   "Agrupado por Tienda"
         Height          =   345
         Left            =   135
         TabIndex        =   14
         Top             =   945
         Width           =   1950
      End
      Begin VB.OptionButton opt_orden_numero 
         Caption         =   "Ordenado Por Número"
         Height          =   390
         Left            =   135
         TabIndex        =   3
         Top             =   270
         Width           =   1950
      End
      Begin VB.OptionButton opt_orden_tienda 
         Caption         =   "Ordenado por Tienda"
         Height          =   390
         Left            =   135
         TabIndex        =   2
         Top             =   585
         Width           =   1950
      End
   End
End
Attribute VB_Name = "frmreporte_envios_tiendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tipo_mes As Integer
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
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            cnn.BeginTrans
            rs.Open "select max(inte_eti_consecutivo) from tb_temp_envios_tiendas", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "Insert into tb_temp_envios_tiendas (INTE_ETI_CONSECUTIVO, vcha_aud_usuario, vcha_aud_maquina) values (" + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "')", cnn, adOpenDynamic, adLockOptimistic
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
            'rs.Open "select * from tb_encabezado_movimientos where CHAR_EMO_TIPO_CLIENTE_PROVEEDOR = 'T' and VCHA_EMO_AFECTACION = '-' and dtim_emo_fecha >= " + var_fecha_inicio + " and dtim_emo_fecha <= " + var_fecha_fin + " AND CHAR_EMO_ESTATUS <> 'C'", cnn, adOpenDynamic, adLockOptimistic
            var_cadena = " INSERT INTO TB_TEMP_ENVIOS_TIENDAS (INTE_ETI_CONSECUTIVO, DTIM_ETI_FECHA_INICIO, DTIM_ETI_FECHA_FIN, VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID,INTE_EMO_NUMERO, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA)"
            var_cadena = var_cadena + "select " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + ", VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO, '" + var_clave_usuario_global + "', '" + fun_NombrePc + "' from tb_encabezado_movimientos where CHAR_EMO_TIPO_CLIENTE_PROVEEDOR = 'T' and VCHA_EMO_AFECTACION = '-' and dtim_emo_fecha >= " + var_fecha_inicio + " and dtim_emo_fecha <= " + var_fecha_fin + "-0.00001 AND CHAR_EMO_ESTATUS <> 'C'"
            rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            var_fecha_fin_1 = CDate(txt_fin)
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
            
            rs.Open "select * from TB_TEMP_ENVIOS_TIENDAS where INTE_ETI_CONSECUTIVO = " + CStr(var_consecutivo) + " and vcha_emp_empresa_id is not null", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
            '   While Not rs.EOF
            '         var_cadena = " INSERT INTO TB_TEMP_ENVIOS_TIENDAS (INTE_ETI_CONSECUTIVO, DTIM_ETI_FECHA_INICIO, DTIM_ETI_FECHA_FIN, VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID,INTE_EMO_NUMERO, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA)"
            '         var_cadena = var_cadena + "Values ( " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + ", '" + rs!VCHA_EMP_EMPRESA_ID + "','" + rs!VCHA_UOR_UNIDAD_ID + "', '" + rs!VCHA_ALM_ALMACEN_ID + "',  '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_EMO_NUMERO) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "')"
            '         rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            '         rs.MoveNext
            '   Wend
               rs.Close
            
               If opt_orden_numero = True Then
                  Set reporte = appl.OpenReport(App.Path + "\rep_envio_tiendas_numero.rpt")
                  reporte.RecordSelectionFormula = "{VW_ENVIO_TIENDAS_RESUMEN.INTE_ETI_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_ENVIO_TIENDAS_RESUMEN.VCHA_AUD_USUARIO} = '" + var_clave_usuario_global + "' and {VW_ENVIO_TIENDAS_RESUMEN.VCHA_AUD_MAQUINA} = '" + fun_NombrePc + "'"
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Reporte de Envios a Tiendas"
                  frmvistasprevias.Show 1
                  Set reporte = Nothing
                  var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_envio_tiendas_numero.rpt")
                     reporte.RecordSelectionFormula = "{VW_ENVIO_TIENDAS_RESUMEN.INTE_ETI_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_ENVIO_TIENDAS_RESUMEN.VCHA_AUD_USUARIO} = '" + var_clave_usuario_global + "' and {VW_ENVIO_TIENDAS_RESUMEN.VCHA_AUD_MAQUINA} = '" + fun_NombrePc + "'"
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     reporte.ExportOptions.FormatType = crEFTExcel80
                     reporte.ExportOptions.DestinationType = crEDTDiskFile
                     archivo = "c:\reportessid\Envios_tiendas" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     reporte.ExportOptions.DiskFileName = archivo
                     reporte.Export False
                     Set reporte = Nothing
                     MsgBox "Se a terminado de guardar el archivo " + archivo
                  End If
               End If
               If opt_orden_tienda = True Then
                  Set reporte = appl.OpenReport(App.Path + "\rep_envio_tiendas_tienda.rpt")
                  reporte.RecordSelectionFormula = "{VW_ENVIO_TIENDAS_RESUMEN.INTE_ETI_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_ENVIO_TIENDAS_RESUMEN.VCHA_AUD_USUARIO} = '" + var_clave_usuario_global + "' and {VW_ENVIO_TIENDAS_RESUMEN.VCHA_AUD_MAQUINA} = '" + fun_NombrePc + "'"
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Reporte de Envios a Tiendas"
                  frmvistasprevias.Show 1
                  Set reporte = Nothing
                  var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_envio_tiendas_tienda.rpt")
                     reporte.RecordSelectionFormula = "{VW_ENVIO_TIENDAS_RESUMEN.INTE_ETI_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_ENVIO_TIENDAS_RESUMEN.VCHA_AUD_USUARIO} = '" + var_clave_usuario_global + "' and {VW_ENVIO_TIENDAS_RESUMEN.VCHA_AUD_MAQUINA} = '" + fun_NombrePc + "'"
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     reporte.ExportOptions.FormatType = crEFTExcel80
                     reporte.ExportOptions.DestinationType = crEDTDiskFile
                     archivo = "c:\reportessid\Envios_tiendas" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     reporte.ExportOptions.DiskFileName = archivo
                     reporte.Export False
                     Set reporte = Nothing
                     MsgBox "Se a terminado de guardar el archivo " + archivo
                  End If
               End If
               If opt_agrupado_tienda = True Then
                  Set reporte = appl.OpenReport(App.Path + "\rep_envio_tiendas_agente_fecha.rpt")
                  reporte.RecordSelectionFormula = "{VW_ENVIO_TIENDAS_RESUMEN.INTE_ETI_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_ENVIO_TIENDAS_RESUMEN.VCHA_AUD_USUARIO} = '" + var_clave_usuario_global + "' and {VW_ENVIO_TIENDAS_RESUMEN.VCHA_AUD_MAQUINA} = '" + fun_NombrePc + "'"
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Reporte de Envios a Tiendas"
                  frmvistasprevias.Show 1
                  Set reporte = Nothing
                  var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_envio_tiendas_agente_fecha.rpt")
                     reporte.RecordSelectionFormula = "{VW_ENVIO_TIENDAS_RESUMEN.INTE_ETI_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_ENVIO_TIENDAS_RESUMEN.VCHA_AUD_USUARIO} = '" + var_clave_usuario_global + "' and {VW_ENVIO_TIENDAS_RESUMEN.VCHA_AUD_MAQUINA} = '" + fun_NombrePc + "'"
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     reporte.ExportOptions.FormatType = crEFTExcel80
                     reporte.ExportOptions.DestinationType = crEDTDiskFile
                     archivo = "c:\reportessid\Envios_tiendas" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     reporte.ExportOptions.DiskFileName = archivo
                     reporte.Export False
                     Set reporte = Nothing
                     MsgBox "Se a terminado de guardar el archivo " + archivo
                  End If
               End If
               If opt_detalle = True Then
                  Set reporte = appl.OpenReport(App.Path + "\rep_envio_tiendas_detalle.rpt")
                  reporte.RecordSelectionFormula = "{VW_ENVIO_TIENDAS_DESGLOSADO.INTE_ETI_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_ENVIO_TIENDAS_DESGLOSADO.VCHA_AUD_USUARIO} = '" + var_clave_usuario_global + "' and {VW_ENVIO_TIENDAS_DESGLOSADO.VCHA_AUD_MAQUINA} = '" + fun_NombrePc + "'"
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Reporte de Envios a Tiendas"
                  frmvistasprevias.Show 1
                  Set reporte = Nothing
                  var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     Set reporte = appl.OpenReport(App.Path + "\rep_envio_tiendas_detalle.rpt")
                     reporte.RecordSelectionFormula = "{VW_ENVIO_TIENDAS_DESGLOSADO.INTE_ETI_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_ENVIO_TIENDAS_DESGLOSADO.VCHA_AUD_USUARIO} = '" + var_clave_usuario_global + "' and {VW_ENVIO_TIENDAS_DESGLOSADO.VCHA_AUD_MAQUINA} = '" + fun_NombrePc + "'"
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     reporte.ExportOptions.FormatType = crEFTExcel80
                     reporte.ExportOptions.DestinationType = crEDTDiskFile
                     archivo = "c:\reportessid\Envios_tiendas" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     reporte.ExportOptions.DiskFileName = archivo
                     reporte.Export False
                     Set reporte = Nothing
                     MsgBox "Se a terminado de guardar el archivo " + archivo
                  End If
               End If
               rs.Open "delete from tb_temp_envios_tiendas where inte_eti_consecutivo = " + CStr(var_consecutivo) + " and vcha_aud_usuario = '" + var_clave_usuario_global + "' and vcha_aud_maquina = '" + fun_NombrePc + "'", cnn, adOpenDynamic, adLockOptimistic
            Else
               rs.Close
               MsgBox "No existen movimientos", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "La fecha de inicio debe de ser menor a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de Inicio Incorrecta"
   End If
End Sub




Private Sub cmd_salir_Click()
   Unload Me
End Sub



Private Sub Command11_Click()
      If IsDate(Me.txt_inicio) Then
         mes = CDate(Me.txt_inicio)
      Else
         mes = Date
      End If
   var_tipo_mes = 1
   mes.Visible = True
   mes.SetFocus
End Sub

Private Sub Command12_Click()
      If IsDate(Me.txt_fin) Then
         mes = CDate(Me.txt_fin)
      Else
         mes = Date
      End If
   var_tipo_mes = 2
   mes.Visible = True
   mes.SetFocus
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 2000
   Left = 3200
   txt_inicio = Date
   txt_fin = Date
   mes.Visible = False
   opt_orden_numero.Value = True
   Dim list_item As ListItem
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_despliega_menu = True Then
      var_swpassword = False
      var_modifica_registro = False
   End If
   Call activa_forma(var_activa_forma_reporte_envios_tiendas)
End Sub






Private Sub mes_DateDblClick(ByVal DateDblClicked As Date)
   If var_tipo_mes = 1 Then
      txt_inicio = mes.Value
   End If
   If var_tipo_mes = 2 Then
      txt_fin = mes
   End If
   mes.Visible = False
End Sub

Private Sub mes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_tipo_mes = 1 Then
         txt_inicio = mes.Value
      End If
      If txt_fin = 2 Then
         txt_fin = mes
      End If
      mes.Visible = False
   End If
   If KeyAscii = 27 Then
      mes.Visible = False
   End If

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
