VERSION 5.00
Begin VB.Form frmembarques_paquetes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmes"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3345
   Icon            =   "frmembarques_paquetes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   3345
   Begin VB.TextBox txt_acceso 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   225
      TabIndex        =   0
      Top             =   345
      Width           =   2910
   End
End
Attribute VB_Name = "frmembarques_paquetes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 3000
   Left = 3850
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_embarques_paquetes)
End Sub

Private Sub txt_acceso_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If Trim(txt_acceso) <> "" Then
         rs.Open "select * from tb_encabezado_paquetes where inte_paq_numero = " + txt_acceso, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            rs.Close
            var_numero_embarque_paquete = txt_acceso
            var_paquete = True
            frmcodigo_acceso.Show 1
         Else
            rs.Close
            si = MsgBox("El embarque no existe, ¿Desea dar uno de alta?", vbYesNo, "ATENCION")
            If si = 6 Then
               frmembarques_paquetes_2.Show 1
            End If
         End If
      Else
         si = MsgBox("¿Desea dar un embarque de alta?", vbYesNo, "ATENCION")
         If si = 6 Then
            frmembarques_paquetes_2.Show 1
         End If
      End If
   End If
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub
