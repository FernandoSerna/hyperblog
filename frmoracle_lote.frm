VERSION 5.00
Begin VB.Form frmoracle_lote 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lote"
   ClientHeight    =   870
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   3150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txt_embarque 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   2655
   End
   Begin VB.TextBox txt_lote 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmoracle_lote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txt_lote_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If IsNumeric(Me.txt_lote) Then
          If Len(Me.txt_lote) >= 9 Then
             var_lote_global = Me.txt_lote
             
             var_pedido = Mid(Me.txt_lote, 1, Len(Me.txt_lote) - 3)
             var_lote = Mid(Me.txt_lote, Len(Me.txt_lote) - 2, 3)
             rs.Open "select * from tb_oracle_pedidos_asignados_embarques where pedido = " + CStr(CDbl(var_pedido)), cnn, adOpenDynamic, adLockOptimistic
             If Not rs.EOF Then
                If rs!Embarque = Me.txt_embarque Then
                   var_lote_global = var_lote_global
                Else
                   var_lote_global = ""
                   MsgBox "El pedido seleccionado pertenece al embarque " + CStr(rs!Embarque), vbOKOnly, "ATENCION"
                End If
             Else
                var_lote_global = ""
                MsgBox "El pedido seleccionado pertenece al embarque " + CStr(rs!Embarque), vbOKOnly, "ATENCION"
             End If
             rs.Close
             Unload Me
          Else
             MsgBox "Lote incorrecto", vbOKOnly, "ATENCION"
          End If
       Else
          MsgBox "Número de lote incorrecto.", vbOKOnly, "ATENCION"
       End If
    End If
End Sub
