VERSION 5.00
Begin VB.Form frmnumero_embarque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Número de Embarque"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3345
   Icon            =   "frmnumero_embarque.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   3345
   Begin VB.TextBox txt_embarque 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   555
      Left            =   240
      TabIndex        =   0
      Top             =   195
      Width           =   2910
   End
End
Attribute VB_Name = "frmnumero_embarque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_habilita_forma As Boolean
Dim var_bloqueado As Integer
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter
Dim objConn As New ADODB.Connection
Dim objCmd As New ADODB.Command
Dim objParm As ADODB.Parameter


Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
   txt_embarque.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   var_habilita_forma = True
   Top = 3000
   Left = 3850
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_bloqueado = 1 And var_numero_embarque > 0 Then
      rs.Open "UPDATE TB_ENCABEZADO_EMBARQUES SET INTE_EMB_BLOQUEADO = 0, VCHA_EMB_BLOQUEADO_POR = '' WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_EMB_EMBARQUE = " + CStr(var_numero_embarque), cnn, adOpenDynamic, adLockOptimistic
   End If
   If var_bloqueado = 0 And var_numero_embarque > 0 Then
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open "UPDATE TB_ENCABEZADO_EMBARQUES SET INTE_EMB_BLOQUEADO = 0, VCHA_EMB_BLOQUEADO_POR = '' WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_EMB_EMBARQUE = " + CStr(var_numero_embarque), cnn, adOpenDynamic, adLockOptimistic
   End If
   var_es_embarque = False
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_embarque_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione Shift + F5 para ver la información de los embarques"
End Sub

Private Sub txt_embarque_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = 116 Then
      frmbusqueda_embarque.Show 1
   End If
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   Dim var_bloqueado As Integer
   Dim var_nombre_bloqueado As String
   Dim var_maquina_embarque As String
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   
   If KeyAscii = 13 Then
      If Trim(txt_embarque) <> "" Then
         If rsaux.State = 1 Then
            rsaux.Close
         End If
         rsaux.Open "select * from XXVIA_tb_encabezado_embarques where EMBARQUE = " + txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            If var_bandera_asignacion = 0 Then
               If var_tipo_embarque = 2 Then
                  rs.Open "select * from tb_oracle_maquinas_asignadas where embarque= " + Me.txt_embarque + " and uso = 'S' and maquina = '" + fun_NombrePc + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_maquina_embarque = rs!MAQUINA
                  End If
                  rs.Close
               Else
                  rs.Open "select * from tb_oracle_maquinas_asignadas where embarque= " + Me.txt_embarque + " and uso = 'E' and maquina = '" + fun_NombrePc + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_maquina_embarque = rs!MAQUINA
                  Else
                     If rsaux1.State = 1 Then
                        rsaux1.Close
                     End If
                     rsaux1.Open "select * from tb_oracle_maquinas_asignadas where embarque= " + Me.txt_embarque + " and uso = 'E'", cnn, adOpenDynamic, adLockOptimistic
                     var_maquina_embarque = ""
                     While Not rsaux1.EOF
                           If var_maquina_embarque = "" Then
                              var_maquina_embarque = rsaux1!MAQUINA
                           Else
                              var_maquina_embarque = var_maquina_embarque + ", " + rsaux1!MAQUINA
                           End If
                           rsaux1.MoveNext
                     Wend
                     rsaux1.Close
                  End If
                  rs.Close
               End If
            Else
               var_maquina_embarque = IIf(IsNull(rsaux!MAQUINA), "", rsaux!MAQUINA)
            End If
            var_estatus_embarque = Trim(IIf(IsNull(rsaux!char_emb_estatus), "", rsaux!char_emb_estatus))
            var_tipo_embarque = IIf(IsNull(rsaux!tipo_embarque), 0, rsaux!tipo_embarque)
            var_tipo_embarque = 2
            var_numero_jaula = IIf(IsNull(rsaux!JAULA), 0, rsaux!JAULA)
            var_vf = fun_NombrePc
            var_maquina_embarque = fun_NombrePc
            If fun_NombrePc = UCase(var_maquina_embarque) Then
               If IIf(IsNull(rsaux!char_emb_estatus), "", rsaux!char_emb_estatus) = "F" Then
                  MsgBox "El embarque ya fue cerrado y no puede ser modificado", vbOKOnly, "ATENCION"
               Else
                  If var_tipo_embarque = 2 Then
                     If var_bandera_asignacion = 0 Then
                        If var_salida_cajas = 1 Then
                           If rsaux1.State = 1 Then
                              rsaux1.Close
                           End If
                           rsaux1.Open "SELECT * FROM tb_oracle_maquinas_asignadas where embarque = " + Me.txt_embarque + " and uso = 'S'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux1.EOF Then
                              var_maquina_embarque = rsaux1!MAQUINA
                           End If
                           rsaux1.Close
                           If UCase(var_maquina_embarque) = fun_NombrePc Then
                              var_numero_embarque = CDbl(Me.txt_embarque)
                              
   
                               
                               
                               
                               
                              var_embarque_global = var_numero_embarque
                              frmoracle_salida_cajas_aduana.txt_embarque = Me.txt_embarque
                                 
                                 
                              var_transporte_global = ""
                              frmoracle_transortes.Show 1
                              If var_transporte_global <> "" Then
                                 rs.Open "UPDATE XXVIA_TB_ENCABEZADO_EMBARQUES SET TRANSPORTE = '" + var_transporte_global + "' WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 MsgBox "Se a actualizado el transporte", vbOKOnly, "ATENCION"
                              Else
                                 MsgBox "No se selecciono un transporte", vbOKOnly, "ATENCION"
                              End If
                              If IsNumeric(var_embarque_global) Then
                                 
                                 strconsulta = "SELECT TRANSPORTE FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = ?"
                                 With comandoORA
                                      'MsgBox cnnoracle_4.ConnectionString
                                      .ActiveConnection = cnnoracle_4
                                      .CommandType = adCmdText
                                      .CommandText = strconsulta
                                      Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(var_embarque_global))
                                      .Parameters.Append parametro
                                 End With
                                 Set rsaux8 = comandoORA.execute
                                 Set comandoORA = Nothing
                                 Set parametro = Nothing
                                 var_clave_transporte = ""
                                 If Not rsaux8.EOF Then
                                    var_clave_transporte = IIf(IsNull(rsaux8!transporte), "", rsaux8!transporte)
                                 End If
                                 rsaux8.Close
                                 rsaux8.Open "SELECT isnull(NOMBRE,'') nombre FROM TB_ORACLE_TRANSPORTES WHERE CLAVE = '" + var_clave_transporte + "'", cnn, adOpenDynamic, adLockOptimistic
                                 frmoracle_salida_cajas_aduana.lbl_transporte = ""
                                 
                                 If Not rsaux8.EOF Then
                                    frmoracle_salida_cajas_aduana.lbl_transporte = rsaux8!NOMBRE
                                 Else
                                    frmoracle_salida_cajas_aduana.lbl_transporte = ""
                                 End If
                                 rsaux8.Close
                              End If
                              
                              frmoracle_salida_cajas_aduana.Show 1
                           
                           Else
                              MsgBox "El embarque fue asignado a la máquina " + var_maquina_embarque, vbOKOnly, "ATENCION"
                           End If
                        Else
                           var_numero_embarque = CDbl(Me.txt_embarque)
                           var_embarque_global = var_numero_embarque
                           frmcodigo_acceso.lbl_embarque.Caption = "Embarque:" + Str(var_numero_embarque)
                           frmcodigo_acceso.Show
                        End If
                     Else
                        If IIf(IsNull(rsaux!char_emb_estatus), "", rsaux!char_emb_estatus) = "E" Then
                           var_numero_embarque = CDbl(Me.txt_embarque)
                           var_embarque_global = var_numero_embarque
                           frmoracle_salidas_cajas.txt_embarque = Me.txt_embarque
                           frmoracle_salidas_cajas.Show
                        Else
                           var_numero_embarque = CDbl(Me.txt_embarque)
                           var_embarque_global = var_numero_embarque
                           If var_prueba_2 = 1 Then
                              frmcodigo_acceso.lbl_embarque.Caption = "Embarque:" + Str(var_numero_embarque)
                              frmcodigo_acceso.Show
                           Else
                              frmcodigo_acceso.lbl_embarque.Caption = "Embarque:" + Str(var_numero_embarque)
                              frmcodigo_acceso.Show
                           End If
                        End If
                     End If
                  Else
                     var_numero_embarque = CDbl(Me.txt_embarque)
                     frmcodigo_acceso.lbl_embarque.Caption = "Embarque:" + Str(var_numero_embarque)
                     frmcodigo_acceso.Show
                  End If
               End If
            Else
               If var_bandera_asignacion = 0 Then
                  MsgBox "El embarque fue asignado a la(s) máquina(s) " + var_maquina_embarque, vbOKOnly, "ATENCION"
               Else
                  MsgBox "No se puede abrir el embarque ya que fue hecho en la máquina " + var_maquina_embarque, vbOKOnly, "ATENCION"
               End If
            End If
         Else
            If var_bandera_asignacion = 1 Then
               var_si = MsgBox("El embarque no existe ¿Desea dar uno de alta?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  frmembarques.Show
               End If
            Else
               MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
            End If
         End If
         If rsaux.State = 1 Then
            rsaux.Close
          End If
      Else
         If var_bandera_asignacion = 1 Then
            si = MsgBox("¿Desea dar un embarque de alta?", vbYesNo, "ATENCION")
            If si = 6 Then
               frmembarques.Show 1
               Me.txt_embarque.SetFocus
            End If
         Else
            MsgBox "No se a indicado el número de embarque", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub

Private Sub txt_embarque_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

