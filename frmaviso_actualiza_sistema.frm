VERSION 5.00
Begin VB.Form frmaviso_actualiza_sistema 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   225
      Left            =   255
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1575
      Width           =   840
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Height          =   1335
      Left            =   75
      TabIndex        =   1
      Top             =   -15
      Width           =   4485
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "El Sistema Integral de Distribución SID esta actualizandose, favor de esperar un momento."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   1035
         Left            =   135
         TabIndex        =   2
         Top             =   210
         Width           =   4260
      End
   End
End
Attribute VB_Name = "frmaviso_actualiza_sistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub

Private Sub Text1_GotFocus()
   On Error GoTo salir:
   Me.Refresh
   FileCopy var_archivo_servidor, var_archivo_local
   Unload Me
   Exit Sub
salir:
   MsgBox "Existe una nueva actualizacion en el servidor, salga de todas las instancias del sistema para poder actualizar.", vbOKOnly, "ATENCION"
   Unload Me
End Sub
