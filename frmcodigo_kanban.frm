VERSION 5.00
Begin VB.Form frmcodigo_kanban 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1050
      Left            =   75
      TabIndex        =   1
      Top             =   315
      Width           =   4530
      Begin VB.TextBox txt_codigo_kanban 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   720
         Left            =   90
         TabIndex        =   2
         Top             =   195
         Width           =   4335
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "Codigo Kanban"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   4620
   End
End
Attribute VB_Name = "frmcodigo_kanban"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Me.txt_codigo_kanban = ""
End Sub

Private Sub txt_codigo_kanban_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_tarjeta_kanban = ""
      Unload Me
   End If
End Sub

Private Sub txt_codigo_kanban_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.txt_codigo_kanban <> "" Then
         rsaux.Open "SELECT * FROM TB_TARJETAS_KANBAN WHERE VCHA_KAN_KANBAN_ID = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            var_tarjeta_kanban = txt_codigo
         Else
            var_tarjeta_kanban = ""
            frmmensaje.lbl_mensaje = "Número de tarjeta Kanban no existe"
            frmmensaje.Show 1
         End If
         rsaux.Close
      End If
   End If
End Sub
