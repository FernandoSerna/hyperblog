VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmmensaje 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atención"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7050
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   7050
   StartUpPosition =   1  'CenterOwner
   Begin WMPLibCtl.WindowsMediaPlayer wmp2 
      Height          =   870
      Left            =   0
      TabIndex        =   4
      Top             =   1350
      Visible         =   0   'False
      Width           =   435
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   0   'False
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   100
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   767
      _cy             =   1535
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   645
      Left            =   6555
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   765
      URL             =   "C:\sistemas\desarrollo\integral\Articulo no existe.wav"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   0   'False
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   100
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   0   'False
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   1349
      _cy             =   1138
   End
   Begin VB.Label lbl_articulo 
      Alignment       =   2  'Center
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
      Height          =   495
      Left            =   45
      TabIndex        =   2
      Top             =   165
      Width           =   6855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Presione Shift + F5 para salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   2070
      TabIndex        =   1
      Top             =   2295
      Width           =   3045
   End
   Begin VB.Label lbl_mensaje 
      Alignment       =   2  'Center
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
      Height          =   1395
      Left            =   105
      TabIndex        =   0
      Top             =   810
      Width           =   6855
   End
End
Attribute VB_Name = "frmmensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cdm_sonido_Click()
   wmp2.Controls.Play
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = 116 Then
      Unload Me
   End If
   
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   Call cdm_sonido_Click
End Sub

Private Sub Form_Load()
   Me.wmp2.URL = App.Path + "\Mec_Alarm_10.wav"
   wmp2.Controls.Play
End Sub
