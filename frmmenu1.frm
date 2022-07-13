VERSION 5.00
Begin VB.Form frmmenu1 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   Icon            =   "frmmenu1.frx":0000
   LinkTopic       =   "frmmenu1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   9060
   Begin VB.Timer Timer1 
      Interval        =   1500
      Left            =   15
      Top             =   30
   End
   Begin VB.CommandButton cmd_opciones_bloques 
      Enabled         =   0   'False
      Height          =   195
      Index           =   3
      Left            =   6360
      Picture         =   "frmmenu1.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.CommandButton cmd_opciones_bloques 
      Enabled         =   0   'False
      Height          =   195
      Index           =   2
      Left            =   4680
      Picture         =   "frmmenu1.frx":1E4B
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.CommandButton cmd_opciones_bloques 
      Enabled         =   0   'False
      Height          =   195
      Index           =   1
      Left            =   3000
      Picture         =   "frmmenu1.frx":3207
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.CommandButton cmd_opciones_bloques 
      Enabled         =   0   'False
      Height          =   195
      Index           =   0
      Left            =   1335
      Picture         =   "frmmenu1.frx":4817
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1215
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image Image9 
      Height          =   6435
      Left            =   165
      Picture         =   "frmmenu1.frx":5F96
      Top             =   0
      Visible         =   0   'False
      Width           =   8850
   End
   Begin VB.Image Image8 
      Height          =   6435
      Left            =   165
      Picture         =   "frmmenu1.frx":A02A
      Top             =   0
      Visible         =   0   'False
      Width           =   8850
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   480
      Left            =   780
      TabIndex        =   4
      Top             =   6405
      Width           =   7425
   End
   Begin VB.Image Image7 
      Height          =   6435
      Left            =   165
      Picture         =   "frmmenu1.frx":E0B8
      Top             =   0
      Visible         =   0   'False
      Width           =   8850
   End
   Begin VB.Image Image6 
      Height          =   210
      Left            =   6360
      Picture         =   "frmmenu1.frx":11866
      Top             =   2040
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Image Image5 
      Height          =   210
      Left            =   4680
      Picture         =   "frmmenu1.frx":1281D
      Top             =   2040
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Image Image4 
      Height          =   210
      Left            =   3000
      Picture         =   "frmmenu1.frx":13A2E
      Top             =   2040
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Image Image3 
      Height          =   210
      Left            =   1320
      Picture         =   "frmmenu1.frx":14A72
      Top             =   2040
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Image Image2 
      Height          =   690
      Left            =   1080
      Picture         =   "frmmenu1.frx":15C40
      Top             =   360
      Visible         =   0   'False
      Width           =   6915
   End
   Begin VB.Image Image1 
      Height          =   3960
      Left            =   1080
      Picture         =   "frmmenu1.frx":1DB61
      Top             =   2520
      Visible         =   0   'False
      Width           =   6900
   End
End
Attribute VB_Name = "frmmenu1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_opciones_bloques_Click(Index As Integer)
   var_global_menu = ""
   If Index = 0 Then
      rs.Open "select * from VW_RELACIONES_BLOQUES_PUESTOS where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_uor_unidad_id = '" + var_empresa_global + "' and vcha_blo_bloque_id = '1'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_global_menu = rs(5).Value
      Else
         var_global_menu = ""
      End If
      rs.Close
      var_bloque_global = "1"
   End If
   If Index = 1 Then
      rs.Open "select * from VW_RELACIONES_BLOQUES_PUESTOS where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_uor_unidad_id = '" + var_empresa_global + "' and vcha_blo_bloque_id = '2'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_global_menu = rs(5).Value
      Else
         var_global_menu = ""
      End If
      rs.Close
      var_bloque_global = "2"
   End If
   If Index = 2 Then
      rs.Open "select * from VW_RELACIONES_BLOQUES_PUESTOS where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_uor_unidad_id = '" + var_empresa_global + "' and vcha_blo_bloque_id = '3'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_global_menu = rs(5).Value
      Else
         var_global_menu = ""
      End If
      rs.Close
      var_bloque_global = "3"
   End If
   If Index = 3 Then
      rs.Open "select * from VW_RELACIONES_BLOQUES_PUESTOS where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_uor_unidad_id = '" + var_empresa_global + "' and vcha_blo_bloque_id = '4'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_global_menu = rs(5).Value
      Else
         var_global_menu = ""
      End If
      rs.Close
      var_bloque_global = "4"
   End If
   Unload Me
   Frmmenu2.Show
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   If var_empresa = "31" Then
      Me.Image8.Visible = True
   Else
      If var_empresa = "18" Then
         Me.Image9.Visible = True
      Else
         Me.Image7.Visible = True
      End If
   End If
   
   Timer1.Enabled = True
   var_cadena_seguridad = ""
   var_top = 3772.5 - (frmmenu1.Height / 2)
   var_left = 5002.5 - (frmmenu1.Height / 2)
   frmmenu1.Top = var_top
   frmmenu1.Left = var_left
   cmd_opciones_bloques(0).Enabled = False
   cmd_opciones_bloques(1).Enabled = False
   cmd_opciones_bloques(2).Enabled = False
   cmd_opciones_bloques(3).Enabled = False
   If var_clave_usuario_global = "1" Then
      cmd_opciones_bloques(0).Enabled = False
      cmd_opciones_bloques(1).Enabled = False
      cmd_opciones_bloques(2).Enabled = False
      cmd_opciones_bloques(3).Enabled = True
   Else
      rs.Open "select * from VW_RELACIONES_BLOQUES_UNIDADES where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_uor_unidad_id = '" + var_empresa_global + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         While Not rs.EOF
            If rs(2).Value = "1" Then
               cmd_opciones_bloques(0).Enabled = True
            End If
            If rs(2).Value = "2" Then
               cmd_opciones_bloques(1).Enabled = True
            End If
            If rs(2).Value = "3" Then
               cmd_opciones_bloques(2).Enabled = True
            End If
            If rs(2).Value = "4" Then
               cmd_opciones_bloques(3).Enabled = True
            End If
            rs.MoveNext
         Wend
      End If
      rs.Close
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
End Sub

Private Sub Timer1_Timer()
   If Timer1.Interval = 1500 Then
      rs.Open "select * from VW_RELACIONES_BLOQUES_PUESTOS where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_uor_unidad_id = '" + var_empresa_global + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_global_menu = rs(5).Value
      Else
         var_global_menu = ""
      End If
      rs.Close
      var_bloque_global = "1"
      Unload Me
      Frmmenu2.Show
   End If
End Sub
