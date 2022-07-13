VERSION 5.00
Begin VB.Form frmoracle_eliminar_bultos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Eliminación de bultos"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1710
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   3435
      Begin VB.CommandButton cmd_eliminar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3000
         Picture         =   "frmoracle_eliminar_bultos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Eliminar Alt + E"
         Top             =   1245
         Width           =   330
      End
      Begin VB.TextBox txt_caja 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1725
         TabIndex        =   6
         Top             =   1185
         Width           =   1200
      End
      Begin VB.TextBox txt_pedido 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1725
         TabIndex        =   5
         Top             =   720
         Width           =   1605
      End
      Begin VB.TextBox txt_embarque 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1725
         TabIndex        =   4
         Top             =   255
         Width           =   1605
      End
      Begin VB.Label Label3 
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
         Left            =   120
         TabIndex        =   3
         Top             =   1215
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Pedido:"
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
         Left            =   120
         TabIndex        =   2
         Top             =   750
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Embarque:"
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
         Left            =   120
         TabIndex        =   1
         Top             =   285
         Width           =   1560
      End
   End
End
Attribute VB_Name = "frmoracle_eliminar_bultos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comandoORA As New adodb.Command
Dim parametro As adodb.Parameter

Private Sub cmd_eliminar_Click()
   If IsNumeric(Me.txt_embarque) Then
      If IsNumeric(Me.txt_pedido) Then
         If IsNumeric(Me.txt_caja) Then
            strconsulta = "select * from xxvia_tb_Salidas_cajas where inte_emb_embarque = ? AND SOURCE_HEADER_NUMBER = ? and inte_paq_caja = ?"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_embarque))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_pedido))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_caja))
                 .Parameters.Append parametro
            End With
            Set rs = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            If Not rs.EOF Then
               var_si = MsgBox("¿Se eliminara la caja del embarque?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  var_si = MsgBox("¿Se eliminara la caja del embarque?", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     strconsulta = "select * from xxvia_tb_encabezado_embarques where embarque = ? "
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_embarque))
                          .Parameters.Append parametro
                     End With
                     Set rsaux = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     If Not rsaux.EOF Then
                        var_estatus = IIf(IsNull(rsaux!char_emb_estatus), "", rsaux!char_emb_estatus)
                        If var_estatus <> "I" Then
                           If var_estatus <> "F" Then
                              var_usuario_cerrar_pantalla = ""
                              frmoracle_autoriza_cerrar_pantalla.Show 1
                              If var_usuario_cerrar_pantalla <> "" Then
                                 If var_contraseña_cerrar_pantalla <> "" Then
                                    rsaux1.Open "INSERT INTO TB_ORACLE_cAJAS_ELIMINADAS (USUARIO, EMBARQUE, PEDIDO,CAJA, FECHA) VALUES ('" + var_usuario_cerrar_pantalla + "'," + Me.txt_embarque + "," + Me.txt_pedido + "," + Me.txt_caja + ",GETDATE())", cnn, adOpenDynamic, adLockOptimistic
                                    strconsulta = "delete xxvia_tb_Salidas_cajas where inte_emb_embarque = ? AND SOURCE_HEADER_NUMBER = ? and inte_paq_caja = ?"
                                    With comandoORA
                                         .ActiveConnection = cnnoracle_4
                                         .CommandType = adCmdText
                                         .CommandText = strconsulta
                                         Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_embarque))
                                         .Parameters.Append parametro
                                         Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_pedido))
                                         .Parameters.Append parametro
                                         Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_caja))
                                         .Parameters.Append parametro
                                    End With
                                    Set rsaux1 = comandoORA.execute
                                    Set comandoORA = Nothing
                                    Set parametro = Nothing
                                    
                                    
                                    MsgBox "Se a eliminado la caja", vbOKOnly, "ATENCION"
                                  End If
                              End If
                           Else
                              MsgBox "La caja ya no puede ser eliminado ya que el embarque fue cerrado", vbOKOnly, "ATENCION"
                           End If
                        Else
                           MsgBox "La caja ya no puede ser eliminado ya que el embarque fue cerrado", vbOKOnly, "ATENCION"
                        End If
                        
                     End If
                     rsaux.Close
                  End If
               End If
            Else
               MsgBox "La caja no pertenece al pedido seleccionado", vbOKOnly, "ATENCION"
               Me.txt_caja = ""
            End If
            rs.Close
         Else
            MsgBox "Número de caja incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Número de pedido incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Número de embarque incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Form_Load()
   Top = 2800
   Left = 4000
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_caja_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.txt_caja <> "" Then
         If IsNumeric(Me.txt_caja) Then
            Me.cmd_eliminar.SetFocus
         Else
            MsgBox "Número de caja incorrecto", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub

Private Sub txt_caja_LostFocus()
   If Me.txt_embarque <> "" Then
      If Me.txt_pedido <> "" Then
         If Me.txt_caja <> "" Then
            If IsNumeric(Me.txt_embarque) Then
               If IsNumeric(Me.txt_pedido) Then
                  If IsNumeric(Me.txt_caja) Then
                     strconsulta = "select * from xxvia_tb_Salidas_cajas where inte_emb_embarque = ? AND SOURCE_HEADER_NUMBER = ? and inte_paq_caja = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_embarque))
                          .Parameters.Append parametro
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_pedido))
                          .Parameters.Append parametro
                          Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_caja))
                          .Parameters.Append parametro
                     End With
                     Set rs = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     If Not rs.EOF Then
                                  
                     Else
                        MsgBox "La caja no pertenece al pedido seleccionado", vbOKOnly, "ATENCION"
                        Me.txt_caja = ""
                     End If
                     rs.Close
                  Else
                     MsgBox "Número de caja incorrecto", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "Número de pedido incorrecto", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "Número de embarque incorrecto", vbOKOnly, "ATENCION"
            End If
         End If
      End If
   End If
End Sub

Private Sub txt_embarque_Change()
   Me.txt_pedido = ""
   Me.txt_caja = ""
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.txt_embarque <> "" Then
         If IsNumeric(Me.txt_embarque) Then
            Me.txt_pedido.SetFocus
         Else
            MsgBox "Número de embarque incorrecto", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub

Private Sub txt_embarque_LostFocus()
   If Me.txt_embarque <> "" Then
      If IsNumeric(Me.txt_embarque) Then
         strconsulta = "select * from xxvia_tb_Salidas_cajas where inte_emb_embarque = ? "
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_embarque))
              .Parameters.Append parametro
         End With
         Set rs = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         If Not rs.EOF Then
      
         Else
            MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
            Me.txt_embarque = ""
            Me.txt_pedido = ""
            Me.txt_caja = ""
         End If
         rs.Close
      Else
         MsgBox "Número de embarque incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_pedido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.txt_pedido <> "" Then
         If IsNumeric(Me.txt_pedido) Then
            Me.txt_caja.SetFocus
         Else
            MsgBox "Número de pedido incorrecto", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub

Private Sub txt_pedido_LostFocus()
   If Me.txt_embarque <> "" Then
      If Me.txt_pedido <> "" Then
         If IsNumeric(Me.txt_embarque) Then
            If IsNumeric(Me.txt_pedido) Then
               strconsulta = "select * from xxvia_tb_Salidas_cajas where inte_emb_embarque = ? AND SOURCE_HEADER_NUMBER = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_embarque))
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_pedido))
                    .Parameters.Append parametro
               End With
               Set rs = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               If Not rs.EOF Then
         
               Else
                  MsgBox "El pedido no existe o no pertenece al embarque seleccionado", vbOKOnly, "ATENCION"
                  Me.txt_pedido = ""
                  Me.txt_caja = ""
               End If
               rs.Close
            Else
               MsgBox "Número de pedido incorrecto", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Número de embarque incorrecto", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub
