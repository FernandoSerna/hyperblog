VERSION 5.00
Begin VB.Form frm1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
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
   Begin VB.CommandButton cmd2 
      Caption         =   "MsgBox"
      Height          =   495
      Left            =   180
      TabIndex        =   1
      Top             =   5055
      Width           =   1200
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4605
      Left            =   150
      Picture         =   "frm1.frx":0000
      ScaleHeight     =   4605
      ScaleWidth      =   5265
      TabIndex        =   0
      Top             =   270
      Visible         =   0   'False
      Width           =   5265
      Begin VB.PictureBox pic2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         DrawMode        =   1  'Blackness
         ForeColor       =   &H80000008&
         Height          =   2305
         Left            =   0
         ScaleHeight     =   154
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   351
         TabIndex        =   3
         Top             =   2295
         Visible         =   0   'False
         Width           =   5265
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "OK"
         Height          =   375
         Left            =   3960
         TabIndex        =   2
         Top             =   1710
         Width           =   1095
      End
      Begin VB.Label lblCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "This is a sample Message Box"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1620
         TabIndex        =   4
         Top             =   945
         Width           =   3360
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   1005
         Left            =   240
         Picture         =   "frm1.frx":3B47
         Top             =   645
         Width           =   1170
      End
   End
End
Attribute VB_Name = "frm1"
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
    
    tmrOut2.Enabled = True
    pic2.Height = 15
    pic2.Left = (Me.Width - pic2.Width) / 2
    pic2.Top = (Me.Height - pic2.Height) / 2
    
End Sub

Private Sub cmd2_Click()
    
    tmrOut.Enabled = True

End Sub

Private Sub Form_Load()
    
    pic1.Width = 5295
    pic1.Height = 15
    pic2.Width = 5295
    pic2.Height = 2300
    pic2.Top = 0
    pic2.Left = 0

    pic1.Left = (Me.Width - pic1.Width) / 2
    pic1.Top = (Me.Height - pic1.Height) / 2
    
End Sub

Private Sub Form_Resize()

    cmd2.Top = Me.Height - 1080
    cmd2.Left = Me.Width - 1400
    
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
