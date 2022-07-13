VERSION 5.00
Begin VB.Form frmsellos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sello"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   990
      Left            =   60
      TabIndex        =   0
      Top             =   135
      Width           =   4635
      Begin VB.TextBox txt_sello 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   675
         Left            =   60
         TabIndex        =   1
         Top             =   210
         Width           =   4500
      End
   End
End
Attribute VB_Name = "frmsellos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txt_sello_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(Me.txt_sello) <> "" Then
         var_sello_caja = Me.txt_sello
         Unload Me
      Else
         var_sello_caja = ""
         Unload Me
      End If
   End If
   If KeyAscii = 27 Then
      If Trim(Me.txt_sello) <> "" Then
         var_sello_caja = Me.txt_sello
         Unload Me
      Else
         var_sello_caja = ""
         Unload Me
      End If
   End If
End Sub
