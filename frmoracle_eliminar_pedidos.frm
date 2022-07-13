VERSION 5.00
Begin VB.Form frmoracle_eliminar_pedidos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Eliminar pedidos"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3120
      Picture         =   "frmoracle_eliminar_pedidos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmoracle_eliminar_pedidos.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Eliminar"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   135
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3375
      Begin VB.TextBox txt_embarque 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmoracle_eliminar_pedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_guardar_Click()
   If IsNumeric(Me.txt_embarque) Then
      rs.Open "SELECT * FROM XXVIA_TB_salidas_cajas WHERE INTE_eMB_EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
      If rs.EOF = True Then
         var_si = MsgBox("¿Desea eliminar los pedidos del emberque?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_si = MsgBox("Confirmar la eliminación de los pedidos", vbYesNo, "ATENCION")
            If var_si = 6 Then
               rsaux.Open "SELECT * FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  While Not rsaux.EOF
                        rsaux1.Open "DELETE FROM XXVIA_TB_PEDIDOS_DIVIDIDOS WHERE SOURCE_HEADER_NUMBER = " + CStr(rsaux!PEDIDO), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  
                        rsaux.MoveNext
                  Wend
                  rsaux1.Open "INSERT INTO TB_ORACLE_BITACORA_ELIMINA_PEDIDOS (USUARIO, MAQUINA, FECHA, EMBARQUE) VALUES ('" + var_clave_usuario_global + "','" + fun_NombrePc + "',GETDATE(),'" + Me.txt_embarque + "')", cnn, adOpenDynamic, adLockOptimistic
                  MsgBox "Los pedidos del embarque han sido eliminados, favor de imprimirlo de nuevo", vbOKOnly, "ATENCION"
               Else
                  MsgBox "El embarque no contiene pedidos", vbOKOnly, "ATENCION"
               End If
               rsaux.Close
            End If
         End If
      Else
         MsgBox "El embarque ya no puede ser eliminado porque ya esta siendo trabajado", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      MsgBox "Número de embarque incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 3850
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub
