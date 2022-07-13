VERSION 5.00
Begin VB.Form frmcambio_precios_devolucion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambio de valor de devoluciones"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   0
      TabIndex        =   11
      Top             =   330
      Width           =   7920
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmcambio_precios_devolucion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frmcambio_precios_devolucion.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_cancelar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   735
      Picture         =   "frmcambio_precios_devolucion.frx":024C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Cancelar Alt + C"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Caption         =   " Devolucion "
      Height          =   2325
      Left            =   90
      TabIndex        =   10
      Top             =   450
      Width           =   7800
      Begin VB.TextBox txt_piezas 
         Enabled         =   0   'False
         Height          =   350
         Left            =   6075
         TabIndex        =   18
         Top             =   1425
         Width           =   1500
      End
      Begin VB.TextBox txt_importe_correcto 
         Height          =   350
         Left            =   1395
         TabIndex        =   6
         Top             =   1800
         Width           =   1995
      End
      Begin VB.TextBox txt_importe 
         Height          =   350
         Left            =   3600
         TabIndex        =   5
         Top             =   1425
         Width           =   1500
      End
      Begin VB.TextBox txt_fecha 
         Height          =   350
         Left            =   1395
         TabIndex        =   4
         Top             =   1425
         Width           =   1260
      End
      Begin VB.TextBox txt_cliente 
         Height          =   350
         Left            =   1395
         TabIndex        =   3
         Top             =   1050
         Width           =   6180
      End
      Begin VB.TextBox txt_numero 
         Height          =   350
         Left            =   1395
         TabIndex        =   2
         Top             =   675
         Width           =   1740
      End
      Begin VB.TextBox txt_nombre_movimiento 
         Height          =   350
         Left            =   2625
         TabIndex        =   1
         Top             =   300
         Width           =   5025
      End
      Begin VB.TextBox txt_movimiento 
         Height          =   350
         Left            =   1395
         TabIndex        =   0
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Piezas:"
         Height          =   195
         Left            =   5355
         TabIndex        =   19
         Top             =   1500
         Width           =   510
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Importe correcto:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1875
         Width           =   1200
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
         Height          =   195
         Left            =   2880
         TabIndex        =   16
         Top             =   1500
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1500
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1125
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   750
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Movimiento:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   375
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmcambio_precios_devolucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_pedidos_Click()
   If Me.txt_movimiento <> "" Then
      If Me.txt_nombre_movimiento <> "" Then
         If IsNumeric(Me.txt_numero) Then
            If Me.txt_cliente <> "" Then
               If IsDate(Me.txt_fecha) Then
                  If IsNumeric(Me.txt_importe) Then
                     If IsNumeric(Me.txt_importe_correcto) Then
                        If CDbl(Me.txt_importe) <> CDbl(Me.txt_importe_correcto) Then
                           If IsNumeric(Me.txt_piezas) Then
                              var_si = MsgBox("¿Deseas cambiar el importe de la devolución", vbYesNo, "ATENCION")
                              If var_si = 6 Then
                                 var_si = MsgBox("Confirmar la cancelación de la devolución", vbYesNo, "ATENCION")
                                 If var_si = 6 Then
                                    var_precio = CDbl(Me.txt_importe_correcto) / CDbl(Me.txt_piezas)
                                    rs.Open "SELECT * FROM TB_DEVOLUCIONES WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + Me.txt_movimiento + "' AND INTE_EMO_NUMERO = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
                                    While Not rs.EOF
                                          var_cadena = " INSERT INTO TB_BITACORA_DEVOLUCIONES ([VCHA_EMP_EMPRESA_ID],[VCHA_UOR_UNIDAD_ID],[VCHA_ALM_ALMACEN_ID],[VCHA_MOV_MOVIMIENTO_ID],[INTE_EMO_NUMERO], [VCHA_ART_ARTICULO_ID], [INTE_CDE_CONSECUTIVO], [FLOA_CDE_PRECIO],[FLOA_CDE_DESCUENTO_1], [FLOA_CDE_DESCUENTO_2], [FLOA_CDE_DESCUENTO_3], [VCHA_BIT_USUARIO], [VCHA_BIT_MAQUINA], [DTIM_BIT_FECHA]) "
                                          var_cadena = var_cadena + " Values ('" + rs!vcha_emp_empresa_id + "','" + rs!vcha_uor_unidad_id + "', '" + rs!vcha_alm_almacen_id + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "'," + CStr(rs!INTE_EMO_NUMERO) + ",'" + rs!VCHA_aRT_ARTICULO_ID + "',  " + CStr(rs!INTE_CDE_CONSECUTIVO) + ", " + CStr(rs!floa_cde_precio) + ", " + CStr(rs!floa_cde_descuento_1) + ", " + CStr(rs!floa_cde_descuento_2) + "," + CStr(rs!floa_cde_descuento_3) + ", '" + var_clave_usuario_global + "','" + fun_NombrePc + "',getdate())"
                                          rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                          rs.MoveNext
                                    Wend
                                    rs.Close
                                    rs.Open "UPDATE TB_DEVOLUCIONES SET FLOA_CDE_PRECIO = " + CStr(var_precio) + "/ (1 +(FLOA_CDE_IVA/100)), FLOA_CDE_DESCUENTO_1 = 0, FLOA_CDE_DESCUENTO_2 = 0, FLOA_CDE_DESCUENTO_3 = 0  WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + Me.txt_movimiento + "' AND INTE_EMO_NUMERO = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
                                    MsgBox "Se a cambiado el importe de la devolución satisfactoriamente", vbOKOnly, "ATENCION"
                                 Else
                                    MsgBox "Se a cancelado el cambio de importe de la devolución", vbOKOnly, "ATENCION"
                                 End If
                              Else
                                 MsgBox "Se a cancelado el cambio de importe de la devolución", vbOKOnly, "ATENCION"
                              End If
                           Else
                              MsgBox "Cantidad de piezas incorrectas", vbOKOnly, "ATENCION"
                           End If
                        Else
                           MsgBox "El importe de la devolución debe de ser diferente al importe a actualizar", vbOKOnly, "ATENCION"
                        End If
                     Else
                        MsgBox "Importe a actualizar incorrecto", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "El movimiento no existe", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "El movimiento no existe", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "El movimiento no existe", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Número de movimiento incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Movimiento incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Se debe de seleccionar un movimiento", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_cancelar_pedidos_Click()
   Unload Me
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_movimiento = ""
   Me.txt_nombre_movimiento = ""
   Me.txt_numero = ""
   Me.txt_cliente = ""
   Me.txt_fecha = ""
   Me.txt_importe = ""
   Me.txt_importe_correcto = ""
   Me.txt_piezas = ""
   Me.txt_movimiento = "CA"
   Me.txt_nombre_movimiento = "ENTRADAS A CALIDAD"
   Me.txt_movimiento.SetFocus
End Sub

Private Sub Form_Load()
   Top = 2200
   Left = 2000
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_fecha_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_importe_correcto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_aceptar_pedidos.SetFocus
   End If
End Sub

Private Sub txt_importe_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_movimiento_Change()
   Me.txt_nombre_movimiento = ""
   Me.txt_numero = ""
   Me.txt_cliente = ""
   Me.txt_fecha = ""
   Me.txt_importe = ""
   Me.txt_importe_correcto = ""
   Me.txt_piezas = ""
End Sub

Private Sub txt_movimiento_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_movimiento_LostFocus()
   If Trim(Me.txt_movimiento) <> "" Then
      rs.Open "select * from tb_movimientos where vcha_mov_movimiento_id = '" + Me.txt_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_movimiento = IIf(IsNull(rs!vcha_mov_nombre), "", rs!vcha_mov_nombre)
         Me.txt_numero = ""
         Me.txt_cliente = ""
         Me.txt_fecha = ""
         Me.txt_importe = ""
         Me.txt_importe_correcto = ""
      Else
         MsgBox "El movimiento no existe", vbOKOnly, "ATENCION"
         Me.txt_nombre_movimiento = ""
         Me.txt_numero = ""
         Me.txt_cliente = ""
         Me.txt_fecha = ""
         Me.txt_importe = ""
         Me.txt_importe_correcto = ""
      End If
      rs.Close
   Else
      Me.txt_nombre_movimiento = ""
      Me.txt_numero = ""
      Me.txt_cliente = ""
      Me.txt_fecha = ""
      Me.txt_importe = ""
      Me.txt_importe_correcto = ""
   End If
End Sub

Private Sub txt_nombre_movimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_numero_Change()
   Me.txt_cliente = ""
   Me.txt_fecha = ""
   Me.txt_importe = ""
   Me.txt_importe_correcto = ""
   Me.txt_piezas = ""
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_numero_LostFocus()
   If Me.txt_numero <> "" Then
      If IsNumeric(Me.txt_numero) Then
         If Me.txt_movimiento <> "" Then
            rs.Open "SELECT * FROM TB_DEVOLUCIONES WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + Me.txt_movimiento + "' AND INTE_EMO_NUMERO = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_estatus = IIf(IsNull(rs!CHAR_CDE_ESTATUS), "", rs!CHAR_CDE_ESTATUS)
               If var_estatus = "I" Then
                  If var_empresa = "18" Then
                     rsaux.Open "SELECT SUM(((((FLOA_CDE_PRECIO * FLOA_DEV_CANTIDAD) * (1 -(ISNULL(FLOA_CDE_DESCUENTO_1,0)/100))) * (1-(ISNULL(FLOA_CDE_DESCUENTO_2,0)/100)))  *(1-(ISNULL(FLOA_CDE_DESCUENTO_3,0)/100))) *(1 + (FLOA_CDE_IVA/100)) ), SUM(FLOA_DEV_CANTIDAD)  FROM TB_DEVOLUCIONES WHERE VCHA_MOV_MOVIMIENTO_ID = '" + Me.txt_movimiento + "' AND INTE_EMO_NUMERO = " + Me.txt_numero + " AND VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                  Else
                     rsaux.Open "SELECT SUM(((((FLOA_CDE_PRECIO * FLOA_DEV_CANTIDAD) * (1 -(ISNULL(FLOA_CDE_DESCUENTO_1,0)/100))) * (1-(ISNULL(FLOA_CDE_DESCUENTO_2,0)/100)))  *(1-(ISNULL(FLOA_CDE_DESCUENTO_3,0)/100))) *(1 + (FLOA_CDE_IVA/100)) ), SUM(FLOA_DEV_CANTIDAD)  FROM TB_DEVOLUCIONES WHERE VCHA_MOV_MOVIMIENTO_ID = '" + Me.txt_movimiento + "' AND INTE_EMO_NUMERO = " + Me.txt_numero + " AND INTE_DEV_RECHAZADO <> 1 AND VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                  End If
                  If Not rsaux.EOF Then
                     rsaux2.Open "SELECT * FROM TB_ENCABEZADO_MOVIMIENTOS WHERE  VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + Me.txt_movimiento + "' AND INTE_EMO_NUMERO = " + Me.txt_numero, cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux2.EOF Then
                        rsaux3.Open "SELECT * FROM TB_CLIENTES WHERE VCHA_CLI_CLAVE_ID = '" + IIf(IsNull(rsaux2!VCHA_CLI_CLAVE_ID), "", rsaux2!VCHA_CLI_CLAVE_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux3.EOF Then
                           Me.txt_cliente = IIf(IsNull(rsaux3!VCHA_CLI_NOMBRE), "", rsaux3!VCHA_CLI_NOMBRE)
                           Me.txt_fecha = IIf(IsNull(rsaux2!dtim_emo_fecha), "", rsaux2!dtim_emo_fecha)
                           Me.txt_importe = Format((rsaux(0).Value), "###,###,##0.00")
                           Me.txt_piezas = Format((rsaux(1).Value), "###,###,##0.00")
                           Me.txt_importe_correcto = ""
                        Else
                           MsgBox "El cliente no existe", vbOKOnly, "ATENCION"
                           Me.txt_cliente = ""
                           Me.txt_fecha = ""
                           Me.txt_importe = ""
                           Me.txt_importe_correcto = ""
                           Me.txt_piezas = ""
                        End If
                        rsaux3.Close
                     Else
                     End If
                     rsaux2.Close
                  End If
                  rsaux.Close
               Else
                  If var_estatus = "" Then
                     MsgBox "El movimiento aun no puede ser corregido ya que no a sido cerrado", vbOKOnly, "ATENCION"
                     Me.txt_cliente = ""
                     Me.txt_fecha = ""
                     Me.txt_importe = ""
                     Me.txt_importe_correcto = ""
                     Me.txt_piezas = ""
                  End If
                  If var_estatus = "N" Then
                     MsgBox "El movimiento no puede ser modificado ya que ya se imprimio la nota de crédito", vbOKOnly, "ATENCION"
                     Me.txt_cliente = ""
                     Me.txt_fecha = ""
                     Me.txt_importe = ""
                     Me.txt_importe_correcto = ""
                     Me.txt_piezas = ""
                  End If
               End If
            Else
               MsgBox "El movimiento no existe", vbOKOnly, "ATENCION"
            End If
            rs.Close
         Else
            MsgBox "Se debe de indicar un movimiento", vbOKOnly, "ATENCION"
            Me.txt_cliente = ""
            Me.txt_fecha = ""
            Me.txt_importe = ""
            Me.txt_importe_correcto = ""
            Me.txt_numero = ""
            Me.txt_piezas = ""
         End If
      Else
         MsgBox "Número incorrecto", vbOKOnly, "ATENCION"
         Me.txt_cliente = ""
         Me.txt_fecha = ""
         Me.txt_importe = ""
         Me.txt_importe_correcto = ""
         Me.txt_numero = ""
         Me.txt_piezas = ""
      End If
   Else
      Me.txt_cliente = ""
      Me.txt_fecha = ""
      Me.txt_importe = ""
      Me.txt_importe_correcto = ""
      Me.txt_piezas = ""
   End If
End Sub
