VERSION 5.00
Begin VB.Form frmoracle_eliminar_guia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Eliminar guia"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_paqueteria 
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
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1200
      Width           =   4215
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmoracle_desbloquear_lotes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Eliminar"
      Top             =   120
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5880
      Picture         =   "frmoracle_desbloquear_lotes.frx":0532
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   120
      Width           =   330
   End
   Begin VB.TextBox txt_guia 
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
      Left            =   1920
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.TextBox txt_pedido 
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
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   4215
   End
   Begin VB.Frame Frame2 
      Height          =   135
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Width           =   6255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Paqueteria:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   1605
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Caja:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Bulto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   720
   End
End
Attribute VB_Name = "frmoracle_eliminar_guia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objConn As New ADODB.Connection
Dim objCmd As New ADODB.Command
Dim objParm As ADODB.Parameter
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter

Private Sub cmd_guardar_Click()
   If Me.txt_pedido <> "" Then
      If Me.txt_paqueteria <> "" Then
         var_si = MsgBox("¿Desea eliminar las guia del bulto seleccionado?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_si = MsgBox("Confirmar eliminar las guia del bulto seleccionado", vbYesNo, "ATENCION")
            If var_si = 6 Then
               strconsulta = "select * FROM XXVIA.XXVIA_TB_PAQUETERIAS_GUIAS where vcha_Caja_id = ? "
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_pedido)
                    .Parameters.Append parametro
               End With
               Set rs = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               If Not rs.EOF Then
                  'MsgBox "DELETE FROM XXVIA.XXVIA_TB_PAQUETERIAS_GUIAS where vcha_caja_id = " + Me.txt_pedido, vbOKOnly, "ATENCION"
                  strconsulta = "DELETE FROM XXVIA.XXVIA_TB_PAQUETERIAS_GUIAS where vcha_Caja_id = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_pedido)
                       .Parameters.Append parametro
                  End With
                  Set rs = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  
                  
                  rsaux2.Open "INSERT INTO TB_ORACLE_ELIMINAR_GUIA (USUARIO, PEDIDO) VALUES ('" + var_clave_usuario_global + "','" + Me.txt_pedido + "')", cnn, adOpenDynamic, adLockOptimistic
               Else
                  MsgBox "El bulto seleccionado no tiene guia asignada.", vbOKOnly, "ATENCION"
               End If
               If rs.State = 1 Then
                  rs.Close
               End If
            End If
         End If
      Else
         MsgBox "El pedido seleccionado no tiene ninguna guia relacionada"
      End If
   Else
      MsgBox "No se indico un pedido", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Top = 2000
    Left = 2500

End Sub

Private Sub Form_Unload(Cancel As Integer)
   activa_forma (var_activa_forma_existencias_generales)
End Sub

Private Sub txt_guia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_pedido) Then
         var_caja = ""
         If Len(Me.txt_guia) > 9 Then
            var_caja = "C" + Mid(Me.txt_guia, (Len(Me.txt_guia) - 8), 9)
            MsgBox var_caja
            Me.txt_guia = var_caja
            strconsulta = "select *  FROM XXVIA.XXVIA_TB_PAQUETERIAS_GUIAS where numb_pedido = ? and vcha_caja_id =  ? "
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_pedido)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_guia)
                 .Parameters.Append parametro
            End With
            Set rs = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
         Else
            strconsulta = "select *  FROM XXVIA.XXVIA_TB_PAQUETERIAS_GUIAS where numb_pedido = ? and vcha_caja_id =  ? "
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_pedido)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_guia)
                 .Parameters.Append parametro
            End With
            Set rs = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
         End If
         If Not rs.EOF Then
            Me.txt_paqueteria = IIf(IsNull(rs!vcha_paqueteria), "", rs!vcha_paqueteria)
         Else
            MsgBox "No existe la caja seleccionada.", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         MsgBox "La caja no existe para el pedido seleccionado.", vbOKOnly, "ATENCION"
         Me.txt_guia = ""
      End If
   End If
End Sub

Private Sub txt_pedido_Change()
   Me.txt_guia = ""
   Me.txt_paqueteria = ""
End Sub

Private Sub txt_pedido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_pedido) Then
         strconsulta = "select * FROM XXVIA.XXVIA_TB_PAQUETERIAS_GUIAS where vcha_Caja_id = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_pedido)
              .Parameters.Append parametro
         End With
         Set rs = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         If Not rs.EOF Then
            Me.txt_guia = ""
            Me.cmd_guardar.SetFocus
            Me.txt_paqueteria = IIf(IsNull(rs!vcha_paqueteria), "", rs!vcha_paqueteria)
         Else
            MsgBox "No existe ninguna guia para el bulto seleccionado.", vbOKOnly, "ATENCION"
            Me.txt_guia = ""
            Me.txt_paqueteria = ""
         End If
         rs.Close
      Else
         MsgBox "No existen guias para el bulto seleccionado.", vbOKOnly, "ATENCION"
         Me.txt_pedido = ""
      End If
   End If
End Sub
