VERSION 5.00
Begin VB.Form frmoracle_unir_bulto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bulto a unir"
   ClientHeight    =   780
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   3555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txt_pedido 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txt_embarque 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txt_caja_unir 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3390
   End
End
Attribute VB_Name = "frmoracle_unir_bulto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter


Private Sub Form_Load()
   Me.txt_embarque = var_embarque_unir
   Me.txt_pedido = var_pedido_unir
End Sub

Private Sub txt_codigo_Change()

End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)

End Sub

Private Sub txt_caja_unir_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      rsaux.Open "select * from TB_ORACLE_CAJAS_ADUANA where CAJA_ANTERIOR =  '" + Me.txt_caja_unir + "' AND EMBARQUE = '" + Me.txt_embarque + "'", cnn, adOpenDynamic, adLockOptimistic
      If rsaux.EOF Then
         var_posible = 1
         If var_posible = 1 Then
            rs.Open "select * from TB_ORACLE_CAJAS_ADUANA where CAJA =  '" + Me.txt_caja_unir + "' AND EMBARQUE = '" + Me.txt_embarque + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               
               var_embarque = rs!Embarque
               var_pedido = rs!pedido
               var_lote_anterior = IIf(IsNull(rs!lote), 0, rs!lote)
               If var_pedido = CDbl(Me.txt_pedido) Then
                  var_posible = 1
               Else
                  var_posible = 0
                  MsgBox "La bulto no corresponde al pedido del bulto padre.", vbOKOnly, "ATENCION"
               End If
            Else
               var_posible = 0
               MsgBox "El bulto no existe o no pertenece al embarque del bulto padre.", vbOKOnly, "ATENCION"
            End If
            rs.Close
         Else
            var_posible = 0
         End If
      Else
         var_posible = 0
         MsgBox "El bulto ya fue agrupado al bulto " + rsaux!caja_actual, vbOKOnly, "ATENCION"
      End If
      rsaux.Close
      
      If var_posible = 1 Then
         var_si = MsgBox("¿Desea agrupar los bultos?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_si = MsgBox("Confirmar agrupar los bultos", vbYesNo, "ATENCION")
            If var_si = 6 Then
               rs.Open "update TB_ORACLE_CAJAS_ADUANA set estatus = 'L', caja_actual = " + CStr(var_caja_padre) + ",caja_anterior = '" + Me.txt_caja_unir + "', pedido_anterior = '" + Me.txt_pedido + "', embarque_anterior = '" + Me.txt_embarque + "'  WHERE EMBARQUE = " + Me.txt_embarque + " and pedido = " + Me.txt_pedido + " and caja = '" + Me.txt_caja_unir + "'", cnn, adOpenDynamic, adLockOptimistic
               'rs.Open "update TB_ORACLE_CAJAS_ADUANA set embarque = 0, PEDIDO  = 0 WHERE EMBARQUE = " + Me.txt_embarque + " and pedido = " + Me.txt_pedido + " and caja = '" + Me.txt_caja_unir + "'", cnn, adOpenDynamic, adLockOptimistic
               var_caja_actual = var_caja_padre
               var_caja = CDbl(Mid(Me.txt_caja_unir, 8, 3))
               strconsulta = "update xxvia_tb_salidas_cajas set  caja_pedido = ?, tipo_caja = ?, char_paq_estatus = '', inte_paq_caja_Anterior = ?, inte_paq_caja = ?, lote = ?, lote_anterior = ? where inte_emb_embarque = ? and source_header_number = ? and inte_paq_caja = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 6, CDbl(var_caja_pedido_padre))
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_tipo_caja_padre)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 6, CDbl(var_caja))
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 6, CDbl(var_caja_actual))
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 6, CDbl(var_lote_padre))
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 6, CDbl(var_lote_anterior))
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 6, CDbl(Me.txt_embarque))
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 6, CDbl(Me.txt_pedido))
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 6, CDbl(var_caja))
                    .Parameters.Append parametro
               End With
               Set rsaux8 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               MsgBox "Se a agrupado la caja", vbOKOnly, "ATENCION"
               Unload Me
            Else
               Unload Me
            End If
         Else
            Unload Me
         End If
      Else
         MsgBox "No se puede agrupar el bulto.", vbOKOnly, "ATENCION"
         Unload Me
      End If
   End If
End Sub
