VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmreporte_comisiones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Comisiones"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.MonthView mes 
      Height          =   2370
      Left            =   1080
      TabIndex        =   13
      Top             =   0
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   72417281
      CurrentDate     =   38148
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Reporte "
      Height          =   960
      Left            =   90
      TabIndex        =   10
      Top             =   480
      Width           =   4335
      Begin VB.OptionButton opt_general 
         Caption         =   "General"
         Height          =   195
         Left            =   165
         TabIndex        =   12
         Top             =   570
         Width           =   930
      End
      Begin VB.OptionButton opt_linea 
         Caption         =   "Por Linea"
         Height          =   270
         Left            =   165
         TabIndex        =   11
         Top             =   270
         Width           =   1680
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4095
      Picture         =   "frmreporte_comisiones.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   45
      Picture         =   "frmreporte_comisiones.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   7
      Top             =   345
      Width           =   4485
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   720
      Left            =   90
      TabIndex        =   0
      Top             =   1470
      Width           =   4335
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   255
         Width           =   1095
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   690
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   255
         Width           =   1080
      End
      Begin VB.CommandButton Command11 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1800
         Picture         =   "frmreporte_comisiones.frx":073C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Fecha Inicial"
         Top             =   255
         Width           =   330
      End
      Begin VB.CommandButton Command12 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3840
         Picture         =   "frmreporte_comisiones.frx":19AE
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Fecha Final"
         Top             =   270
         Width           =   330
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   315
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmreporte_comisiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_mes As Integer

Private Sub cmd_imprimir_Click()
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If opt_linea.Value = True Then
            Set reporte = appl.OpenReport(App.Path + "\rep_comisiones_linea.rpt")
            reporte.RecordSelectionFormula = "{VW_COMISIONES_LINEA.DTIM_CAP_FECHA_PAGO} >= cdate('" + txt_inicio + "') and {VW_COMISIONES_LINEA.DTIM_CAP_FECHA_PAGO} <= cdate('" + txt_fin + "')"
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Comisiones por Linea"
            frmvistasprevias.Show
            Set reporte = Nothing
         End If
         If opt_general.Value = True Then
            Set reporte = appl.OpenReport(App.Path + "\rep_comisiones_general.rpt")
            reporte.RecordSelectionFormula = "{VW_COMISIONES_general.DTIM_CAP_FECHA_PAGO} >= cdate('" + txt_inicio + "') and {VW_COMISIONES_general.DTIM_CAP_FECHA_PAGO} <= cdate('" + txt_fin + "')"
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Comisiones General"
            frmvistasprevias.Show
            Set reporte = Nothing
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de inicio incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command11_Click()
   var_mes = 1
   mes.Visible = True
   mes.SetFocus
End Sub

Private Sub Command12_Click()
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
   Top = 4000
   Left = 5200
   txt_inicio = Date
   txt_fin = Date
   opt_linea = True
   mes.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_despliega_menu = True Then
      var_swpassword = False
      var_modifica_registro = False
   End If
   Call activa_forma(var_activa_forma_reporte_comisiones)
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
   If KeyAscii = 13 Then
      If var_mes = 1 Then
         txt_inicio = mes.Value
      End If
      If var_mes = 2 Then
         txt_fin = mes.Value
      End If
      mes.Visible = False
   End If
End Sub

Private Sub mes_LostFocus()
   mes.Visible = False
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frmcalendario.Show 1
      txt_fin = var_fecha_general
   End If
End Sub

Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frmcalendario.Show 1
      txt_inicio = var_fecha_general
   End If
End Sub
