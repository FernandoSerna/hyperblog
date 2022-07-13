VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmnavidad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Invitaciones navideñas"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   390
      Left            =   1650
      Top             =   1395
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   688
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Boletos"
      Height          =   735
      Left            =   105
      TabIndex        =   1
      Top             =   945
      Width           =   4155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Invitaciones"
      Height          =   735
      Left            =   105
      TabIndex        =   0
      Top             =   135
      Width           =   4155
   End
End
Attribute VB_Name = "frmnavidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub Command1_Click()
               Set reporte = appl.OpenReport(App.Path + "\navidad_invitacion.rpt")
               frmvistasprevias.cr.ReportSource = reporte
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Invitaciones"
               frmvistasprevias.Show
               Set reporte = Nothing

End Sub

Private Sub Command2_Click()
               Set reporte = appl.OpenReport(App.Path + "\boletos.rpt")
               frmvistasprevias.cr.ReportSource = reporte
               
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Boletos"
               frmvistasprevias.Show
               Set reporte = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub
