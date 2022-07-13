VERSION 5.00
Begin VB.Form frmoracle_sello_caja 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sello de caja"
   ClientHeight    =   900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtoracle_sello_caja 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   705
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   4080
   End
End
Attribute VB_Name = "frmoracle_sello_caja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
   var_sello_caja = Me.txtoracle_sello_caja
End Sub
Private Sub txtoracle_sello_caja_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      var_sello_caja = Me.txtoracle_sello_caja
      Unload Me
   End If
   If KeyAscii = 27 Then
      var_sello_caja = Me.txtoracle_sello_caja
      Unload Me
   End If
End Sub
