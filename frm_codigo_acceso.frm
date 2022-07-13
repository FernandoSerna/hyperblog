VERSION 5.00
Begin VB.Form frmcodigo_acceso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Código de Acceso a Movimientos"
   ClientHeight    =   1005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   Icon            =   "frm_codigo_acceso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
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
      Left            =   240
      TabIndex        =   0
      Top             =   330
      Width           =   2910
   End
   Begin VB.Label lbl_embarque 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   240
      TabIndex        =   1
      Top             =   30
      Width           =   2910
   End
End
Attribute VB_Name = "frmcodigo_acceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_habilita_forma As Boolean
Private Sub Text1_KeyPress(KeyAscii As Integer)
End Sub


Private Sub Command1_Click()
   If mm_sonido.Command = "open" Then
      mm_sonido.Command = "close"
   Else
      mm_sonido.Command = "open"
   End If
End Sub

Private Sub mm_sonido_Done(NotifyCode As Integer)
   x = x + 1
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   var_habilita_forma = True
   frmcodigo_acceso.Top = 3000
   frmcodigo_acceso.Left = 3850
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Frmmenu2.StatusBar1.Panels(1) = ""
   If var_activa_menu = True And var_habilita_forma = True And var_es_embarque = False Then
      Frmmenu2.Enabled = True
   End If
End Sub

Private Sub txt_acceso_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F6 para ver los movimientos disponibles"
End Sub

Private Sub txt_acceso_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 117 Then
      If rs.State = 1 Then
         rs.Close
      End If
      If var_tipo_embarque = 1 Then
         var_activa_forma_listamovimientos = "MENU"
         If var_unidad_organizacional <> 93 Then
            frmoracle_salidas.Show
            var_habilita_forma = False
            Unload Me
            Frmmenu2.Enabled = False
            var_activa_menu = True
         Else
            MsgBox "Opción no valida para esta organización", vbOKOnly, "ATENCION"
         End If
      Else
         var_activa_forma_listamovimientos = "MENU"
         If var_prueba_2 = 1 Then
            frmoracle_cajas_divididas.Show
         Else
            frmoracle_cajas.Show
         End If
         var_habilita_forma = False
         Unload Me
         Frmmenu2.Enabled = False
         var_activa_menu = True
      End If
   End If
   If KeyCode = 116 Then
      If var_bandera_asignacion = 0 Then
         frmoracle_seleccion_pedido.Show 1
         Me.txt_acceso = var_pedido_global
      End If
   End If
End Sub

Private Sub txt_acceso_KeyPress(KeyAscii As Integer)
   Dim var_clave_movimiento As String
   Dim var_requiere_referencia As Integer
   Dim var_requiere_factura As Integer
   Dim var_movimiento As String
   Dim var_nombre_movimiento As String
   Dim var_tipo_movimiento As String
   Dim var_titulo_origen As String
   Dim var_resto As String
   Dim var_orden_surtido As Double
   Dim i As Integer
   If KeyAscii = 39 Or KeyAscii = 61 Or KeyAscii = 44 Then
      KeyAscii = 0
   End If
   If var_bandera_asignacion = 0 Then
      If KeyAscii = 13 Then
         
         txt_acceso = (UCase(txt_acceso))
         var_subir_directo = False
         If var_tipo_embarque = 1 Then
            If IsNumeric(Me.txt_acceso) Then
               var_archivo_lote = CDbl(Me.txt_acceso)
               If var_unidad_organizacional = 93 Then
                  MsgBox "Opción no valida para esta organización", vbOKOnly, "ATENCION"
               Else
                  frmoracle_salidas.txt_archivo = Me.txt_acceso
                  frmoracle_salidas.Show
               End If
            Else
               MsgBox "Número de orden de surtido incorrecta", vbOKOnly, "ATENCION"
            End If
         Else
            var_archivo_lote = CDbl(Me.txt_acceso)
   
            var_oracle_tipo_movimiento = "FA"
            If var_oracle_tipo_movimiento = "FA" Then
               If IsNumeric(Me.txt_acceso) Then
                  If var_prueba_2 = 1 Then
                     If var_metodo_fraccionado = 0 Then
                        'validar que ya este cerrado el pedido anterior
                        
                        var_pedido = Mid(Me.txt_acceso, 1, Len(Me.txt_acceso) - 3)

                        
                        rs.Open "select pedido, isnull(estatus_pedido,0) as estatus_pedido from TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque = " + Me.lbl_embarque + " order by orden_pedido"
                        var_posible = 1
                        While Not rs.EOF
                              var_cerrado = IIf(IsNull(rs!estatus_pedido), 0, rs!estatus_pedido)
                              rs.MoveNext
                              If rs!pedido = CDbl(var_pedido) Then
                                If var_cerrado = 0 Then
                                   var_posible = 0
                                End If
                              End If
                              rs.MovePrevious
                              rs.MoveNext
                        Wend
                        rs.Close
                        
                        frmoracle_cajas_divididas.txt_archivo = Me.txt_acceso
                        frmoracle_cajas_divididas.Show
                     Else
                        
                        'validar que ya este cerrado el pedido anterior
                        
                        var_pedido = Mid(Me.txt_acceso, 1, Len(Me.txt_acceso) - 3)

                        
                        rs.Open "select pedido, isnull(estatus_pedido,0) as estatus_pedido from TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque = " + Replace(Me.lbl_embarque, "Embarque: ", "") + " order by orden_pedido", cnn, adOpenDynamic, adLockOptimistic
                        var_posible = 1
                        While Not rs.EOF
                              var_cerrado = IIf(IsNull(rs!estatus_pedido), 0, rs!estatus_pedido)
                              rs.MoveNext
                              If Not rs.EOF Then
                                 If rs!pedido = CDbl(var_pedido) Then
                                    If var_cerrado = 0 Then
                                       var_posible = 0
                                    End If
                                  End If
                                  rs.MovePrevious
                              End If
                              If Not rs.EOF Then
                                 rs.MoveNext
                              End If
                        Wend
                        rs.Close
                        
                        var_posible = 1
                        If var_posible = 0 Then
                           MsgBox "No es posible comenzar a surtir el pedido ya que no a sido cerrado el pedido anterior", vbOKOnly, "ATENCION"
                        Else
                           frmoracle_cajas_NO_divididas.txt_archivo = Me.txt_acceso
                           frmoracle_cajas_NO_divididas.Show
                        End If
                     End If
                  Else
                     frmoracle_cajas.txt_archivo = Me.txt_acceso
                     frmoracle_cajas.Show
                  End If
               Else
                  MsgBox "Número de orden de surtido incorrecta", vbOKOnly, "ATENCION"
               End If
            End If
         End If
      Else
         If KeyAscii = 27 Then
            Unload Me
         Else
            'KeyAscii = 0
         End If
      End If
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
      If KeyAscii = 13 Then
         txt_acceso = (UCase(txt_acceso))
         var_subir_directo = False
         var_archivo_lote = CDbl(Me.txt_acceso)
    
         If var_tipo_embarque = 1 Then
            If IsNumeric(Me.txt_acceso) Then
               If var_unidad_organizacional = 93 Then
                  MsgBox "Opción no valida para esta organización", vbOKOnly, "ATENCION"
               Else
                  frmoracle_salidas.txt_archivo = Me.txt_acceso
                  frmoracle_salidas.Show
               End If
            Else
               MsgBox "Número de orden de surtido incorrecta", vbOKOnly, "ATENCION"
            End If
         Else
            var_oracle_tipo_movimiento = "FA"
            If var_oracle_tipo_movimiento = "FA" Then
               If IsNumeric(Me.txt_acceso) Then
                  If var_prueba_2 = 1 Then
                     frmoracle_cajas_divididas.txt_archivo = Me.txt_acceso
                     frmoracle_cajas_divididas.Show
                  Else
                     frmoracle_cajas.txt_archivo = Me.txt_acceso
                     frmoracle_cajas.Show
                  End If
               Else
                  MsgBox "Número de orden de surtido incorrecta", vbOKOnly, "ATENCION"
               End If
            End If
         End If
      End If
   End If
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub txt_acceso_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub
