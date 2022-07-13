VERSION 5.00
Begin VB.Form frmoracle_jaulas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Andenes"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   375
      Picture         =   "frmoracle_jaulas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3885
      Picture         =   "frmoracle_jaulas.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   45
      Picture         =   "frmoracle_jaulas.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   60
      TabIndex        =   11
      Top             =   270
      Width           =   4155
   End
   Begin VB.Frame Frame1 
      Height          =   1725
      Left            =   75
      TabIndex        =   0
      Top             =   330
      Width           =   4125
      Begin VB.TextBox txt_maquina_2 
         Height          =   375
         Left            =   1185
         TabIndex        =   7
         Top             =   1155
         Width           =   2835
      End
      Begin VB.TextBox txt_maquina_1 
         Height          =   390
         Left            =   1185
         TabIndex        =   6
         Top             =   735
         Width           =   2835
      End
      Begin VB.TextBox txt_descripcion 
         Height          =   390
         Left            =   1725
         TabIndex        =   5
         Top             =   330
         Width           =   2295
      End
      Begin VB.TextBox txt_jaula 
         Height          =   390
         Left            =   1185
         TabIndex        =   4
         Top             =   330
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Máquina 2:"
         Height          =   195
         Left            =   225
         TabIndex        =   10
         Top             =   1245
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Máquina 1:"
         Height          =   195
         Left            =   225
         TabIndex        =   9
         Top             =   833
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Anden:"
         Height          =   195
         Left            =   225
         TabIndex        =   8
         Top             =   435
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmoracle_jaulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_guardar_Click()
   Me.txt_jaula = ""
   Me.txt_descripcion = ""
   Me.txt_maquina_1 = ""
   Me.txt_maquina_2 = ""
   Me.txt_jaula.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   If IsNumeric(Me.txt_jaula) Then
      rs.Open "select * from tb_jaulas where INTE_JAU_JAULA_ID = " + Me.txt_jaula, cnn, adOpenDynamic, adLockOptimistic
      If rs.EOF Then
         rsaux.Open "insert into tb_jaulas (INTE_JAU_JAULA_ID, VCHA_JAU_NOMBRE, VCHA_JAU_MAQUINA_1, VCHA_JAU_MAQUINA_2) values (" + Me.txt_jaula + ",'" + Me.txt_descripcion + "','" + Me.txt_maquina_1 + "','" + Me.txt_maquina_2 + "')", cnn, adOpenDynamic, adLockOptimistic
         MsgBox "Se inserto el registro correctamente", vbOKOnly, "ATENCION"
      Else
         rsaux.Open "update tb_jaulas set VCHA_JAU_NOMBRE = '" + Me.txt_descripcion + "', VCHA_JAU_MAQUINA_1 = '" + Me.txt_maquina_1 + "', VCHA_JAU_MAQUINA_2 = '" + Me.txt_maquina_2 + "' where INTE_JAU_JAULA_ID = " + Me.txt_jaula, cnn, adOpenDynamic, adLockOptimistic
         MsgBox "Se han aplicado los cambios", vbOKOnly, "ATENCION"
      End If
      rs.Close
      
   Else
      MsgBox "Número de jaula incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Form_Load()
   Top = 2500
   Left = 3300
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_lineas)
End Sub

Private Sub txt_descripcion_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_jaula_Change()
   Me.txt_descripcion = ""
   Me.txt_maquina_1 = ""
   Me.txt_maquina_2 = ""
End Sub

Private Sub txt_jaula_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_jaula_LostFocus()
   If Me.txt_jaula <> "" Then
      If IsNumeric(Me.txt_jaula) Then
         rs.Open "select * from tb_jaulas where inte_jau_jaula_id  = " + Me.txt_jaula, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_descripcion = IIf(IsNull(rs!vcha_jau_nombre), "", rs!vcha_jau_nombre)
            Me.txt_maquina_1 = IIf(IsNull(rs!VCHA_JAU_MAQUINA_1), "", rs!VCHA_JAU_MAQUINA_1)
            Me.txt_maquina_2 = IIf(IsNull(rs!VCHA_JAU_MAQUINA_2), "", rs!VCHA_JAU_MAQUINA_2)
         End If
         rs.Close
      Else
         MsgBox "Número de jaula incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_maquina_1_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_maquina_2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.Command1.SetFocus
   End If
End Sub
