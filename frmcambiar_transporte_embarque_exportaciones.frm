VERSION 5.00
Begin VB.Form frmcambiar_transporte_embarque_exportaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambio de unidad a embarques de exportaciones"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_nombre_unidad 
      Height          =   435
      Left            =   2000
      MaxLength       =   50
      TabIndex        =   2
      Top             =   960
      Width           =   4140
   End
   Begin VB.TextBox txt_clave_unidad 
      Height          =   435
      Left            =   1290
      MaxLength       =   50
      TabIndex        =   1
      Top             =   960
      Width           =   660
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
      Height          =   435
      Left            =   1290
      MaxLength       =   50
      TabIndex        =   0
      Top             =   480
      Width           =   1500
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6000
      Picture         =   "frmcambiar_transporte_embarque_exportaciones.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   0
      Picture         =   "frmcambiar_transporte_embarque_exportaciones.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   0
      TabIndex        =   4
      Top             =   250
      Width           =   6375
   End
   Begin VB.Label lab_paises 
      AutoSize        =   -1  'True
      Caption         =   "Unidad:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   1020
      Width           =   825
   End
   Begin VB.Label lab_paises 
      AutoSize        =   -1  'True
      Caption         =   "Embarque:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   540
      Width           =   1140
   End
End
Attribute VB_Name = "frmcambiar_transporte_embarque_exportaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objConn As New ADODB.Connection
Dim objCmd As New ADODB.Command
Dim objParm As ADODB.Parameter
Dim var_contador_porcentaje As Integer
Dim var_cubicaje As Double
Dim var_ventana As Integer
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub txt_clave_unidad_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_lista_transportes = 1
      frmlista.Show 1
      var_lista_transportes = 0
   End If
End Sub

Private Sub txt_embarque_Change()
   Me.txt_clave_unidad = ""
   Me.txt_nombre_unidad = ""
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)

   If IsNumeric(Me.txt_embarque) Then
      If KeyAscii = 13 Then
         strconsulta = "select embarque, clave, nombre from xxvia_Tb_encabezado_embarques a, xxvia_tb_transportes b where a.transporte = b.clave and embarque = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_embarque))
              .Parameters.Append parametro
         End With
         Set rs = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         If Not rs.EOF Then
            Me.txt_clave_unidad = IIf(IsNull(rs!clave), "", rs!clave)
            Me.txt_nombre_unidad = IIf(IsNull(rs!nombre), "", rs!nombre)
         Else
            MsgBox "El embarque no existe.", vbOKOnly, "ATENCION"
         End If
         rs.Close
      End If
   Else
      MsgBox "Número de embarque incorrecto.", vbOKOnly, "ATENCION"
   End If
End Sub
