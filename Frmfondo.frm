VERSION 5.00
Begin VB.Form Frmfondo 
   BorderStyle     =   0  'None
   ClientHeight    =   4350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6630
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Height          =   4345
      Left            =   0
      Top             =   0
      Width           =   6630
   End
   Begin VB.Image Image1 
      Height          =   4335
      Left            =   0
      Picture         =   "Frmfondo.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6645
   End
End
Attribute VB_Name = "Frmfondo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   Frmacceso.Show
End Sub

Private Sub Form_Load()
   Frmacceso.Show
End Sub
