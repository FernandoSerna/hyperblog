VERSION 5.00
Begin VB.Form frmoracle_aplica_PL 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   555
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   555
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_aplica_AP 
      Caption         =   "Aplica AP"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aplica CA"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmd_aplica_USA 
      Caption         =   "Aplica USA"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "frmoracle_aplica_PL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aplica_AP_Click()
   var_aplica_PL = 3
   Unload Me
End Sub

Private Sub cmd_aplica_USA_Click()
   var_aplica_PL = 1
   Unload Me
End Sub

Private Sub Command1_Click()
   var_aplica_PL = 2
   Unload Me
End Sub
