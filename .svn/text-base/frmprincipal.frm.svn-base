VERSION 5.00
Begin VB.Form frmprincipal 
   BackColor       =   &H80000009&
   Caption         =   "Sistema Integral"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   11415
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   5865
      Left            =   240
      Top             =   720
      Width           =   11100
   End
End
Attribute VB_Name = "frmprincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public primero As Integer

Private Sub Form_Activate()
   If primero = 0 Then
      primero = 1
      Frmacceso.Show
   End If
End Sub

Private Sub Form_Load()
   primero = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   End
End Sub
