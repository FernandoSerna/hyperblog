VERSION 5.00
Begin VB.Form frmreporte_antiguedad_saldos_dolares 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Antigüedad de saldos dolares"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   3420
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   30
      TabIndex        =   4
      Top             =   330
      Width           =   3285
   End
   Begin VB.Frame Frame2 
      Caption         =   " Fecha "
      Height          =   645
      Left            =   45
      TabIndex        =   2
      Top             =   405
      Width           =   3270
      Begin VB.TextBox txt_fecha 
         Height          =   345
         Left            =   900
         TabIndex        =   3
         Top             =   195
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   0
      Picture         =   "frmreporte_antiguedad_saldos_dolares.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2970
      Picture         =   "frmreporte_antiguedad_saldos_dolares.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "frmreporte_antiguedad_saldos_dolares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_mes As Integer

Private Sub cmd_imprimir_Click()
      Dim dl As Long                                 ' Valor devuelto por la función API
      Dim sAttributes As String                  ' Aributos
      Dim sDriver As String                       ' Nombre del controlador
      Dim sDescription As String                ' Descripción del DSN
      Dim sDsnName As String                  ' Nombre del DSN

      Const ODBC_ADD_SYS_DSN As Long = 4         ' Se creará un DSN de sistema
      Const vbAPINull As Long = 0&                         ' Puntero NULL

      ' se elimina
      Const ODBC_REMOVE_SYS_DSN As Long = 6    ' Se eliminará un DSN de sistema

      Dim var_dia, var_mes, var_año As String
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
      cnn_reportes.CommandTimeout = 3600
      cnn_reportes.BeginTrans
      rs.Open "select max(inte_tem_consecutivo) from TB_TEMP_ANTIGUEDAD_SALDOS_DOLARES", cnn_reportes, adOpenDynamic, adLockOptimistic
      If rs.EOF Then
         var_consecutivo = 1
      Else
         var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
      End If
      rs.Close
         rs.Open "insert into TB_TEMP_ANTIGUEDAD_SALDOS_DOLARES (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn_reportes, adOpenDynamic, adLockOptimistic
      cnn_reportes.CommitTrans

      If var_empresa = "03" Then
         rs.Open "exec SP_REPORTE_ANTIGUEDAD_SALDOS_DOLARES_telasdelhogar " + var_fecha + ", " + CStr(var_consecutivo), cnn_reportes, adOpenDynamic, adLockOptimistic
      Else
         rs.Open "exec SP_REPORTE_ANTIGUEDAD_SALDOS_DOLARES " + var_fecha + ", " + CStr(var_consecutivo), cnn_reportes, adOpenDynamic, adLockOptimistic
      End If
      sDsnName = "sqlsistema"
      sDescription = "sqlsistema"
      sDriver = "SQL Server"
      sAttributes = "DSN=" & sDsnName & Chr(0)
      sAttributes = sAttributes & "Server=" + var_sr_reportes & Chr$(0)
      sAttributes = sAttributes & "Description=" & sDescription & Chr(0)
      sAttributes = sAttributes & "Database=" + var_bd_reportes & Chr(0)
      strAttributes = strAttributes & "UID=sa" & Chr$(0)
      strAttributes = strAttributes & "PWD=elia" & Chr$(0)
      dl = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYS_DSN, sDriver, sAttributes)
   
      Set reporte = appl.OpenReport(App.Path + "\rep_antiguedad_saldos_dolares.rpt")
      reporte.RecordSelectionFormula = "{VW_REPORTE_ANTIGUEDAD_SALDOS_DOLARES.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {@SALDO}> 0.01"
      frmvistasprevias.cr.ReportSource = reporte
      For ntablas = 1 To reporte.Database.Tables.Count
      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
           frmvistasprevias.cr.ViewReport
      frmvistasprevias.Caption = "Reporte de valuacion de devoluciones a detalle"
      frmvistasprevias.Show 1
      
      
      Set reporte = Nothing
      var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         Set reporte = appl.OpenReport(App.Path + "\rep_antiguedad_saldos_dolares.rpt")
         reporte.RecordSelectionFormula = "{VW_REPORTE_ANTIGUEDAD_SALDOS_DOLARES.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {@SALDO}> 0.01"
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         reporte.ExportOptions.FormatType = crEFTExcel80
         reporte.ExportOptions.DestinationType = crEDTDiskFile
         archivo = "c:\reportessid\Antiguedad_saldos_dolares_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
         reporte.ExportOptions.DiskFileName = archivo
         reporte.Export False
         Set reporte = Nothing
         MsgBox "Se a terminado de guardar el archivo " + archivo
      End If
      rs.Open "delete from TB_TEMP_ANTIGUEDAD_SALDOS_DOLARES where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn_reportes, adOpenDynamic, adLockOptimistic
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 4000
   txt_fecha = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub txt_fecha_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fecha) Then
         frmcalendario.mes = CDate(Me.txt_fecha)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      txt_fecha = var_fecha_general
   End If
End Sub

Private Sub txt_fecha_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

