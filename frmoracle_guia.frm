VERSION 5.00
Begin VB.Form frmoracle_guia 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Guia"
   ClientHeight    =   1140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txt_codigo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   80
      TabIndex        =   0
      Top             =   390
      Width           =   7065
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000C0&
      Caption         =   "   Guia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   340
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7290
   End
End
Attribute VB_Name = "frmoracle_guia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      var_guia_aduana = Replace(Me.txt_codigo, "'", "-")
      Unload Me
   End If
   If KeyAscii = 27 Then
      var_guia_aduana = Replace(Me.txt_codigo, "'", "-")
      Unload Me
   End If
End Sub
