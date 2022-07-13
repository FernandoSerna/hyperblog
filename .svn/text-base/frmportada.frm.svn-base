VERSION 5.00
Begin VB.Form frmportada 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   240
      Top             =   1920
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando ..."
      Height          =   195
      Left            =   2160
      TabIndex        =   1
      Top             =   3000
      Width           =   870
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   ".................."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   3870
      Left            =   0
      Picture         =   "frmportada.frx":0000
      Top             =   0
      Width           =   5070
   End
End
Attribute VB_Name = "frmportada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer1_Timer()
    If Timer1.Interval = 5000 Then
        Frmacceso.Show
        Unload Me
    End If
End Sub
