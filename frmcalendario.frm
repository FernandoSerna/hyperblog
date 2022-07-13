VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmcalendario 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   2595
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.MonthView mes 
      Height          =   2370
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   62717953
      CurrentDate     =   38196
   End
End
Attribute VB_Name = "frmcalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim f As Form
Dim texto As TextBox
Private Sub MonthView1_DateClick(ByVal DateClicked As Date)

End Sub

Private Sub Form_Load()
   Me.Caption = CStr(mes.Value)
End Sub

Private Sub mes_Click()
   Me.Caption = CStr(mes.Value)
End Sub

Private Sub mes_DateDblClick(ByVal DateDblClicked As Date)
   Me.Caption = CStr(mes.Value)
   var_fecha_general = mes.Value
   Unload Me
End Sub


Private Sub mes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      var_fecha_general = mes.Value
      Unload Me
   End If
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub


