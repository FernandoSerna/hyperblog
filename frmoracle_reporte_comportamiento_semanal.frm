VERSION 5.00
Begin VB.Form frmoracle_reporte_comportamiento_semanal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte comportamiento semanal"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   15
      Picture         =   "frmoracle_reporte_comportamiento_semanal.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3945
      Picture         =   "frmoracle_reporte_comportamiento_semanal.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   840
      Left            =   60
      TabIndex        =   0
      Top             =   435
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
      Top             =   270
      Width           =   4275
   End
End
Attribute VB_Name = "frmoracle_reporte_comportamiento_semanal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim año As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_consecutivo As Integer
   Dim var_afectacion_movimiento As String
   Dim var_vistas As String
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            cnn.CommandTimeout = 720
            cnn.BeginTrans
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_ORACLE_PIEZAS_LEIDAS_USUARIO_SEMANAL", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into TB_TEMP_ORACLE_PIEZAS_LEIDAS_USUARIO_SEMANAL (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
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
            
            var_fecha_fin_1 = CDate(txt_fin) + 1
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
            rs.Open "select distinct USUARIO, datepart(year, FECHA) año, DATEPART(MONTH,fecha) mes, datepart(Week, FECHA) SEMANA from TB_ORACLE_LECTURA_USUARIOS where fecha >= " + var_fecha_inicio + " and fecha < " + var_fecha_fin
            If Not rs.EOF Then
               While Not rs.EOF
                     rsaux.Open "INSERT INTO TB_TEMP_ORACLE_PIEZAS_LEIDAS_USUARIO_SEMANAL (INTE_TEM_CONSECUTIVO, USUARIO, AÑO, MES, SEMANA, FECHA_INICIO, FECHA_FIN, LUNES, MARTES, MIERCOLES, JUEVES, VIERNES, SABADO, DOMINGO) VALUES (" + CStr(var_consecutivo) + ",'" + rs!USUARIO + "'," + CStr(rs!año) + "," + CStr(rs!mes) + "," + CStr(rs!semana) + "," + var_fecha_inicio + "," + var_fecha_fin + "-1,0,0,0,0,0,0,0)", cnn, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               If rsaux.State = 1 Then
                  rsaux.Close
               End If
               rsaux.Open "SELECT * FROM TB_TEMP_ORACLE_PIEZAS_LEIDAS_USUARIO_SEMANAL WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND AÑO IS NOT NULL", cnn, adOpenDynamic, adLockOptimistic
               While Not rsaux.EOF
                     rsaux1.Open "select datepart(year, FECHA) año, datepart(Week, FECHA) semana, datepart(DW, FECHA) dia,  USUARIO, isnull(H_0_1,0) + isnull(H_1_2,0)+ISNULL(h_2_3,0)+isnull(H_3_4,0)+isnull(H_4_5,0)+isnull(H_5_6,0)+isnull(H_6_7,0)+isnull(H_7_8,0)+isnull(H_8_9,0)+isnull(H_9_10,0)+isnull(H_10_11,0)+isnull(H_11_12,0) +isnull(H_12_13,0)+isnull(H_13_14,0)+isnull(H_14_15,0)+isnull(H_15_16,0)+isnull(H_16_17,0)+isnull(H_17_18,0)+isnull(H_18_19,0)+isnull(H_19_20,0)+isnull(H_20_21,0)+isnull(H_21_22,0)+isnull(H_22_23,0)+isnull(H_23_24,0) as piezas  from TB_ORACLE_LECTURA_USUARIOS where DATEPART(YEAR,FECHA) = " + CStr(rsaux!año) + " AND DATEPART(MONTH,FECHA) = " + CStr(rsaux!mes) + " AND DATEPART(WEEK,FECHA) = " + CStr(rsaux!semana) + " AND USUARIO = '" + rsaux!USUARIO + "' AND fecha >= " + var_fecha_inicio + " and fecha < " + var_fecha_fin, cnn, adOpenDynamic, adLockOptimistic
                     While Not rsaux1.EOF
                           If rsaux1!dia = 1 Then
                              rsaux2.Open "update TB_TEMP_ORACLE_PIEZAS_LEIDAS_USUARIO_SEMANAL set lunes = " + CStr(rsaux1!PIEZAS) + " where inte_Tem_Consecutivo = " + CStr(var_consecutivo) + " and semana = " + CStr(rsaux1!semana) + " and año = " + CStr(rsaux1!año) + " AND USUARIO = '" + rsaux1!USUARIO + "'", cnn, adOpenDynamic, adLockOptimistic
                           End If
                           If rsaux1!dia = 2 Then
                              rsaux2.Open "update TB_TEMP_ORACLE_PIEZAS_LEIDAS_USUARIO_SEMANAL set martes = " + CStr(rsaux1!PIEZAS) + " where inte_Tem_Consecutivo = " + CStr(var_consecutivo) + " and semana = " + CStr(rsaux1!semana) + " and año = " + CStr(rsaux1!año) + " AND USUARIO = '" + rsaux1!USUARIO + "'", cnn, adOpenDynamic, adLockOptimistic
                           End If
                           If rsaux1!dia = 3 Then
                              rsaux2.Open "update TB_TEMP_ORACLE_PIEZAS_LEIDAS_USUARIO_SEMANAL set miercoles = " + CStr(rsaux1!PIEZAS) + " where inte_Tem_Consecutivo = " + CStr(var_consecutivo) + " and semana = " + CStr(rsaux1!semana) + " and año = " + CStr(rsaux1!año) + " AND USUARIO = '" + rsaux1!USUARIO + "'", cnn, adOpenDynamic, adLockOptimistic
                           End If
                           If rsaux1!dia = 4 Then
                              rsaux2.Open "update TB_TEMP_ORACLE_PIEZAS_LEIDAS_USUARIO_SEMANAL set jueves = " + CStr(rsaux1!PIEZAS) + " where inte_Tem_Consecutivo = " + CStr(var_consecutivo) + " and semana = " + CStr(rsaux1!semana) + " and año = " + CStr(rsaux1!año) + " AND USUARIO = '" + rsaux1!USUARIO + "'", cnn, adOpenDynamic, adLockOptimistic
                           End If
                           If rsaux1!dia = 5 Then
                              rsaux2.Open "update TB_TEMP_ORACLE_PIEZAS_LEIDAS_USUARIO_SEMANAL set viernes = " + CStr(rsaux1!PIEZAS) + " where inte_Tem_Consecutivo = " + CStr(var_consecutivo) + " and semana = " + CStr(rsaux1!semana) + " and año = " + CStr(rsaux1!año) + " AND USUARIO = '" + rsaux1!USUARIO + "'", cnn, adOpenDynamic, adLockOptimistic
                           End If
                           If rsaux1!dia = 6 Then
                              rsaux2.Open "update TB_TEMP_ORACLE_PIEZAS_LEIDAS_USUARIO_SEMANAL set sabado = " + CStr(rsaux1!PIEZAS) + " where inte_Tem_Consecutivo = " + CStr(var_consecutivo) + " and semana = " + CStr(rsaux1!semana) + " and año = " + CStr(rsaux1!año) + " AND USUARIO = '" + rsaux1!USUARIO + "'", cnn, adOpenDynamic, adLockOptimistic
                           End If
                           If rsaux1!dia = 7 Then
                              rsaux2.Open "update TB_TEMP_ORACLE_PIEZAS_LEIDAS_USUARIO_SEMANAL set domingo = " + CStr(rsaux1!PIEZAS) + " where inte_Tem_Consecutivo = " + CStr(var_consecutivo) + " and semana = " + CStr(rsaux1!semana) + " and año = " + CStr(rsaux1!año) + " AND USUARIO = '" + rsaux1!USUARIO + "'", cnn, adOpenDynamic, adLockOptimistic
                           End If
                           rsaux1.MoveNext
                     Wend
                     rsaux1.Close
                     rsaux.MoveNext
               Wend
               rsaux.Close
               
               rsaux.Open "delete from TB_TEMP_ORACLE_PIEZAS_LEIDAS_USUARIO_SEMANAL where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and semana is null", cnn, adOpenDynamic, adLockOptimistic
               
               Set reporte = appl.OpenReport(App.Path + "\REP_ORACLE_PIEZAS_LEIDAS_SEMANA.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_PIEZAS_LEIDAS_SEMANAL.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_ORACLE_PIEZAS_LEIDAS_SEMANAL.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "'"
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Pedidos cargados"
               frmvistasprevias.Show 1
               Set reporte = Nothing
    
               var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  Set reporte = appl.OpenReport(App.Path + "\REP_ORACLE_PIEZAS_LEIDAS_SEMANA_EXCEL.rpt")
                  reporte.RecordSelectionFormula = "{VW_ORACLE_PIEZAS_LEIDAS_SEMANAL.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_ORACLE_PIEZAS_LEIDAS_SEMANAL.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "'"
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  reporte.ExportOptions.FormatType = crEFTExcel80
                  reporte.ExportOptions.DestinationType = crEDTDiskFile
                  archivo = "c:\reportessid\comportamiento_semanal_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                  reporte.ExportOptions.DiskFileName = archivo
                  reporte.Export False
                  Set reporte = Nothing
                  MsgBox "Se a terminado de guardar el archivo " + archivo
               End If
            Else
               MsgBox "No existen resultados", vbOKOnly, "ATENCION"
            End If
            rs.Close
            rs.Open "delete from TB_TEMP_ORACLE_PIEZAS_LEIDAS_USUARIO_SEMANAL WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
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

End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 3000
   Left = 3500
   txt_inicio = Date
   txt_fin = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
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




