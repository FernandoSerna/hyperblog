VERSION 5.00
Begin VB.Form frmcodigos_recontables 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Código recontable"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5625
      Picture         =   "frmcodigos_recontables.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   45
      Picture         =   "frmcodigos_recontables.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Guardar Alt + G"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   90
      Left            =   0
      TabIndex        =   4
      Top             =   300
      Width           =   6060
   End
   Begin VB.Frame Frame1 
      Caption         =   " Artículo "
      Height          =   1095
      Left            =   75
      TabIndex        =   0
      Top             =   480
      Width           =   5910
      Begin VB.CheckBox chk_recontable 
         Caption         =   "Recontable"
         Height          =   330
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2235
      End
      Begin VB.TextBox txt_descripcion 
         Height          =   345
         Left            =   1575
         TabIndex        =   2
         Top             =   315
         Width           =   4260
      End
      Begin VB.TextBox txt_codigo 
         Height          =   345
         Left            =   135
         TabIndex        =   1
         Top             =   315
         Width           =   1410
      End
   End
End
Attribute VB_Name = "frmcodigos_recontables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_guardar_Click()
   rs.Open "update tb_Articulos set inte_art_salida_masiva = " + CStr(Me.chk_recontable) + " where vcha_Art_Articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
   MsgBox "Se a actualizado el artículo", vbOKOnly, "ATENCION"
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 3000
End Sub

Private Sub Form_Unload(Cancel As Integer)
   activa_forma (var_activa_forma_packing_list)
End Sub

Private Sub txt_codigo_Change()
   Me.txt_descripcion = ""
   Me.chk_recontable = 0
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_descripcion.SetFocus
   End If
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub txt_codigo_LostFocus()
   If Me.txt_codigo <> "" Then
      rs.Open "select * from tb_articulos where vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_descripcion = IIf(IsNull(rs!vcha_Art_nombre_español), "", rs!vcha_Art_nombre_español)
         Me.chk_recontable = IIf(IsNull(rs!inte_Art_salida_masiva), 0, rs!inte_Art_salida_masiva)
      Else
         MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
         Me.txt_descripcion = ""
         Me.chk_recontable = 0
      End If
      rs.Close
   Else
      Me.txt_descripcion = ""
      Me.chk_recontable = 0
   End If
End Sub

Private Sub txt_descripcion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.chk_recontable.SetFocus
   Else
      If KeyAscii = 27 Then
         Unload Me
      Else
         KeyAscii = 0
      End If
   End If
End Sub
