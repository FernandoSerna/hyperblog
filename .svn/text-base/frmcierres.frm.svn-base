VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmcierres 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cierre"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3270
   Icon            =   "frmcierres.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   3270
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.MonthView mon_mes 
      Height          =   2370
      Left            =   0
      TabIndex        =   2
      Top             =   1650
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      ShowWeekNumbers =   -1  'True
      StartOfWeek     =   50593793
      CurrentDate     =   37758
   End
   Begin VB.CommandButton cmd_mes 
      Caption         =   "Mes"
      Height          =   450
      Left            =   2175
      TabIndex        =   3
      Top             =   1110
      Width           =   495
   End
   Begin VB.TextBox txt_mes 
      Height          =   360
      Left            =   825
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1140
      Width           =   1305
   End
   Begin VB.CommandButton cmd_cierre 
      Caption         =   "Ejecutar Cierre"
      Height          =   615
      Left            =   750
      TabIndex        =   0
      Top             =   405
      Width           =   1905
   End
End
Attribute VB_Name = "frmcierres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cierre_Click()
   rs.Open "select "
End Sub

Private Sub cmd_mes_Click()
   mon_mes.Visible = True
   mon_mes.SetFocus
End Sub

Private Sub Form_Load()
   mon_mes.Visible = False
   txt_mes = Date
End Sub

Private Sub mon_mes_DblClick()
   txt_mes = mon_mes.Value
   mon_mes.Visible = False
End Sub

Private Sub mon_mes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      mon_mes.Visible = False
   End If
End Sub

Private Sub mon_mes_LostFocus()
   mon_mes.Visible = False
End Sub
