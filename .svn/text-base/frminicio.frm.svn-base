VERSION 5.00
Begin VB.Form frminicio 
   BackColor       =   &H00800080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2400
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrOut 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2685
      Top             =   5100
   End
   Begin VB.Timer tmrIn 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2265
      Top             =   5100
   End
   Begin VB.Timer tmrIn2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1845
      Top             =   5100
   End
   Begin VB.Timer tmrOut2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1425
      Top             =   5100
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4605
      Left            =   -585
      Picture         =   "frminicio.frx":0000
      ScaleHeight     =   4605
      ScaleWidth      =   5400
      TabIndex        =   0
      Top             =   -975
      Visible         =   0   'False
      Width           =   5400
      Begin VB.Timer Timer1 
         Interval        =   3000
         Left            =   720
         Top             =   2295
      End
      Begin VB.PictureBox pic2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         DrawMode        =   1  'Blackness
         ForeColor       =   &H80000008&
         Height          =   2305
         Left            =   20
         ScaleHeight     =   154
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   358
         TabIndex        =   1
         Top             =   2280
         Visible         =   0   'False
         Width           =   5370
      End
      Begin VB.Label lblCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "Bienvenido al Sistema"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1320
         TabIndex        =   2
         Top             =   1560
         Width           =   3360
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   1005
         Left            =   240
         Picture         =   "frminicio.frx":4062
         Top             =   645
         Width           =   1170
      End
   End
   Begin VB.Image Image2 
      Height          =   5880
      Left            =   -1305
      Picture         =   "frminicio.frx":5A4E
      Stretch         =   -1  'True
      Top             =   -465
      Width           =   5640
   End
End
Attribute VB_Name = "frminicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'* Program and graphics by Bob Shull
'* bobby_shull@yahoo.com
'* 12-16-01
'**********************************************

Private Sub cmd1_Click()
    
'    tmrOut2.Enabled = True
'    pic2.Height = 15
'    pic2.Left = (Me.Width - pic2.Width) / 2
'    pic2.Top = (Me.Height - pic2.Height) / 2
    
End Sub

Private Sub Form_Load()
   Top = 2600
   Left = 3600
   tmrOut.Enabled = True
   pic1.Width = 5295
   pic1.Height = 15
   pic2.Width = 5295
   pic2.Height = 2300
   pic2.Top = 0
   pic2.Left = 0
   pic1.Left = (Me.Width - pic1.Width) / 2
   pic1.Top = (Me.Height - pic1.Height) / 2
End Sub


Private Sub Timer1_Timer()
    If Timer1.Interval = 3000 Then
        Unload Me
        Frmacceso.Show
    End If
End Sub

Private Sub tmrIn_Timer()

    If pic1.Height > 15 Then pic1.Height = pic1.Height - 50
    If pic1.Height <= 15 Then
        tmrIn.Enabled = False
        pic1.Visible = False
    End If
    
    pic1.Left = (Me.Width - pic1.Width) / 2
    pic1.Top = (Me.Height - pic1.Height) / 2
        
End Sub

Private Sub tmrIn2_Timer()

    On Error Resume Next
    If pic2.Height > 100 Then pic2.Height = pic2.Height - 50
    If pic2.Height <= 100 Then
        tmrIn2.Enabled = False
        pic2.Visible = False
    End If

    pic2.Left = (pic1.Width - pic2.Width) / 2
    pic2.Top = (pic1.Height - pic2.Height) / 2

End Sub

Private Sub tmrOut_Timer()

    pic1.Visible = True
    pic2.Visible = True
    
    If pic1.Height < 2300 Then pic1.Height = pic1.Height + 50
    If pic1.Height >= 2300 Then
        tmrOut.Enabled = False
        tmrIn2.Enabled = True
    End If
    
    pic1.Left = (Me.Width - pic1.Width) / 2
    pic1.Top = (Me.Height - pic1.Height) / 2

End Sub

Private Sub tmrOut2_Timer()

    pic2.Visible = True
    
    If pic2.Height < 2300 Then
        pic2.Height = pic2.Height + 50
    Else
        tmrOut2.Enabled = False
        tmrIn.Enabled = True
    End If
        
    pic2.Left = (pic1.Width - pic2.Width) / 2
    pic2.Top = (pic1.Height - pic2.Height) / 2

End Sub
