VERSION 5.00
Begin VB.Form frmoracle_video 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insertar texto"
   ClientHeight    =   915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   3195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   660
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   3000
      Begin VB.CommandButton cmd_guardar 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2535
         Picture         =   "frmoracle_video.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Guardar Alt + G"
         Top             =   195
         Width           =   345
      End
      Begin VB.ComboBox cmb_video 
         Height          =   315
         ItemData        =   "frmoracle_video.frx":0102
         Left            =   90
         List            =   "frmoracle_video.frx":010C
         TabIndex        =   1
         Top             =   210
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmoracle_video"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_guardar_Click()
   var_j = 0
   If Me.cmb_video = "NO INSERTAR TEXTO" Then
      var_j = 0
   Else
      If Me.cmb_video = "INSERTAR TEXTO" Then
         var_j = 1
      End If
   End If
   rs.Open "UPDATE TB_VIDEO SET VIDEO = " + CStr(var_j), cnn, adOpenDynamic, adLockOptimistic
End Sub

Private Sub Form_Load()
   rs.Open "select * from tb_video", cnn, adOpenDynamic
   If rs(0).Value = 0 Then
      Me.cmb_video = "NO INSERTAR TEXTO"
   Else
      Me.cmb_video = "INSERTAR TEXTO"
   End If
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub
