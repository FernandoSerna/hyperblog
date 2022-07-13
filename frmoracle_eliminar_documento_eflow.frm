VERSION 5.00
Begin VB.Form frmoracle_eliminar_documento_eflow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Eliminar documento EFLOW"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   15
      Picture         =   "frmoracle_eliminar_documento_eflow.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2655
      Picture         =   "frmoracle_eliminar_documento_eflow.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   -15
      TabIndex        =   5
      Top             =   375
      Width           =   3045
   End
   Begin VB.Frame Frame1 
      Height          =   1170
      Left            =   45
      TabIndex        =   0
      Top             =   390
      Width           =   2940
      Begin VB.TextBox txt_numero 
         Height          =   375
         Left            =   825
         TabIndex        =   4
         Top             =   630
         Width           =   1995
      End
      Begin VB.TextBox txt_serie 
         Height          =   360
         Left            =   825
         TabIndex        =   2
         Top             =   225
         Width           =   1980
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   135
         TabIndex        =   3
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   308
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmoracle_eliminar_documento_eflow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_eliminar_Click()
   If Me.txt_serie <> "" Then
      If IsNumeric(Me.txt_numero) Then
         MsgBox "select * from xxvia_tb_control_doc_fiscales where serie = '" + Me.txt_serie + "' and numero = '" + Me.txt_numero + "'"
         rs.Open "select * from xxvia_tb_control_doc_fiscales where serie = '" + Me.txt_serie + "' and numero = '" + Me.txt_numero + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            MsgBox "Si existe"
         Else
            MsgBox "No existe el documento en EFLOW", vbOKOnly, "ATENCION"
         End If
         rs.Close
      End If
   End If
End Sub

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 3200
   Left = 4200
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub
