VERSION 5.00
Begin VB.Form frmoracle_cambiar_maquina_embarque 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambiar la máquina al embarque"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   3645
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   45
      Picture         =   "frmoracle_cambiar_maquina_embarque.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3225
      Picture         =   "frmoracle_cambiar_maquina_embarque.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   1185
      Left            =   30
      TabIndex        =   5
      Top             =   375
      Width           =   3525
      Begin VB.TextBox txt_maquina 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1275
         TabIndex        =   1
         Top             =   660
         Width           =   2145
      End
      Begin VB.TextBox txt_embarque 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1275
         TabIndex        =   0
         Top             =   210
         Width           =   2145
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Máquina:"
         Height          =   195
         Left            =   330
         TabIndex        =   7
         Top             =   765
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Embarque:"
         Height          =   195
         Left            =   330
         TabIndex        =   6
         Top             =   315
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   15
      TabIndex        =   4
      Top             =   345
      Width           =   3570
   End
End
Attribute VB_Name = "frmoracle_cambiar_maquina_embarque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_guardar_Click()
   If Me.txt_maquina <> "" Then
      var_si = MsgBox("¿Desea cambiar la máquina del embarque", vbYesNo, "ATENCION")
      If var_si = 6 Then
         var_si = MsgBox("Confirmar el cambio de la máquina", vbYesNo, "ATENCION")
         If var_si = 6 Then
            rs.Open "UPDATE XXVIA_TB_ENCABEZADO_EMBARQUES SET MAQUINA = '" + Me.txt_maquina + "' WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
            MsgBox "Se a cambiado la máquina del embarque", vbOKOnly, "ATENCION"
         End If
      End If
   Else
      MsgBox "Nombre de máquina incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   var_top = 3000
   var_left = 3850
   Top = var_top
   Left = var_left
   var_encontro = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_embarque_Change()
   Me.txt_maquina = ""
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If IsNumeric(Me.txt_embarque) Then
        rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
           Me.txt_maquina = IIf(IsNull(rs!MAQUINA), "", rs!MAQUINA)
           Me.txt_maquina.SetFocus
        Else
           MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
        End If
        rs.Close
     Else
        MsgBox "Número de embarque incorrecto", vbOKOnly, "ATENCION"
     End If
  End If
End Sub

Private Sub txt_maquina_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub
