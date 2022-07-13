VERSION 5.00
Begin VB.Form frmtipo_cambio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipo de cambio"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txt_tipo_cambio 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   195
      TabIndex        =   0
      Top             =   165
      Width           =   3390
   End
End
Attribute VB_Name = "frmtipo_Cambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txt_tipo_cambio_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_tipo_cambio) Then
         var_tipo_cambio_global = CDbl(Me.txt_tipo_cambio)
         Unload Me
      Else
         MsgBox "Debe de indicar el tipo de cambio", vbOKOnly, "ATENCION"
      End If
   Else
      KeyAscii = 0
   End If
End Sub
