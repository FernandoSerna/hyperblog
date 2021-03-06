VERSION 5.00
Begin VB.Form frmcancelacion_facturacion_eliminacion_multibondeados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancelacion de embarques y eliminacion de movimiento"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   30
      Picture         =   "frmcancelacion_facturacion_eliminacion_multibondeados.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Cancelar Factura"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5265
      Picture         =   "frmcancelacion_facturacion_eliminacion_multibondeados.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Cancelar Factura"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos del Embarque "
      Height          =   2595
      Left            =   60
      TabIndex        =   0
      Top             =   465
      Width           =   5580
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   870
         TabIndex        =   5
         Top             =   2130
         Width           =   1695
      End
      Begin VB.TextBox txt_fecha 
         Height          =   360
         Left            =   870
         TabIndex        =   4
         Top             =   1671
         Width           =   1665
      End
      Begin VB.TextBox txt_cliente 
         Height          =   360
         Left            =   870
         TabIndex        =   3
         Top             =   1214
         Width           =   4635
      End
      Begin VB.TextBox txt_agente 
         Height          =   360
         Left            =   870
         TabIndex        =   2
         Top             =   757
         Width           =   4635
      End
      Begin VB.TextBox txt_embarque 
         Height          =   360
         Left            =   870
         TabIndex        =   1
         Top             =   300
         Width           =   1050
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   2213
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1754
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1297
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N?mero:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   383
         Width           =   600
      End
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   0
      TabIndex        =   13
      Top             =   270
      Width           =   5625
   End
End
Attribute VB_Name = "frmcancelacion_facturacion_eliminacion_multibondeados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Set TB_FOLIOS_MOVIMIENTOS = New TB_FOLIOS_MOVIMIENTOS
   Dim var_si As Integer
   Dim var_posible As Boolean
   Dim var_fecha_embarque As Date
   Dim var_fecha As Date
   var_posible = True
   var_fecha_embarque = CDate(Me.txt_fecha)
   
   rs.Open "select getdate() from tb_lineas", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      txt_inicio = rs(0).Value
   End If
   rs.Close
   
   var_dia = CStr(Day(CDate(txt_inicio)))
   var_mes = CStr(Month(CDate(txt_inicio)))
   var_a?o = CStr(Year(CDate(txt_inicio)))
   
   If Len(Trim(var_dia)) = 1 Then
      var_dia = "0" + var_dia
   End If
   If Len(Trim(var_mes)) = 1 Then
      var_mes = "0" + var_mes
   End If
   var_fecha_inicio = "{d '" + var_a?o + "-" + var_mes + "-" + var_dia + "'}"

   rs.Open "select " + var_fecha_inicio + " from tb_lineas", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_fecha = rs(0).Value
   End If
   rs.Close
   
   If var_fecha_embarque >= var_fecha Then
      var_posible = True
      If IsNumeric(Me.txt_embarque) Then
         If var_posible = True Then
            var_si = MsgBox("?Desea cancelar las facturas del embarque y eliminar el movimiento?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               var_si = MsgBox("Confirmar la cancelacion de las facturas del embarque " + Me.txt_embarque + " y la eliminaci?n del movimiento", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  If rsaux8.State = 1 Then
                     rsaux8.Close
                  End If
                  rsaux8.Open "select * from vw_embarques_facturacion_multibondeados where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_embarque + " AND CHAR_CAR_ESTATUS <>'C'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux8.EOF Then
                     While Not rsaux8.EOF
                           var_pedido_credito = IIf(IsNull(rsaux8!inte_ped_pedido_credito), 0, rsaux8!inte_ped_pedido_credito)
                           rs.Open "select * from tb_salidas where vcha_emp_empresa_id = '" + var_empresa + "' and VCHA_MOV_MOVIMIENTO_ID = 'FA' AND inte_Car_numero = " + CStr(rsaux8!inte_car_numero) + " and vcha_ser_serie_id = '" + rsaux8!vcha_ser_serie_id + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rs.EOF Then
                              rsaux2.Open "update tb_encabezado_cartera set char_car_estatus = 'C' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_Car_documento =  'FA' and inte_car_numero = " + CStr(rsaux8!inte_car_numero) + " and vcha_Ser_serie_id = '" + rsaux8!vcha_ser_serie_id + "'", cnn, adOpenDynamic, adLockOptimistic
                              While Not rs.EOF
                                    rsaux2.Open "delete from TB_SALIDAS WHERE VCHA_EMP_EMPRESA_ID = '" + rs!VCHA_EMP_EMPRESA_ID + "' AND VCHA_UOR_UNIDAD_ID = '" + rs!VCHA_UOR_UNIDAD_ID + "' AND VCHA_CAR_DOCUMENTO = 'FA' AND VCHA_SER_SERIE_ID = '" + rs!vcha_ser_serie_id + "' AND INTE_CAR_NUMERO = " + CStr((rs!inte_car_numero)) + " and vcha_art_articulo_id = '" + rs!VCHA_aRT_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                                    rs.MoveNext
                              Wend
                              rsaux2.Open "update tb_saldos set floa_sal_importe = 0 where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_Car_documento =  'FA' and inte_car_numero = " + CStr(rsaux8!inte_car_numero) + " and vcha_Ser_serie_id = '" + rsaux8!vcha_ser_serie_id + "'", cnn, adOpenDynamic, adLockOptimistic
                           End If
                           rs.Close
                           rsaux8.MoveNext
                     Wend
                  Else
                     MsgBox "El embarque no existe o no tiene movimientos", vbOKOnly, "ATENCION"
                  End If
                  rsaux8.Close
               
               End If
            End If
         Else
            MsgBox "La factura no puede ser cancelada ya que pertenece a otro dia", vbOKOnly, "ATENCION"
         End If
      Else
      End If
   Else
      MsgBox "El embarque no puede ser cancelado ya que no es del dia", vbOKOnly, "ATENCION"
   End If
   Me.txt_agente = ""
   Me.txt_cliente = ""
   Me.txt_embarque = ""
   Me.txt_fecha = ""
   Me.txt_importe = ""
End Sub

Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 1500
   Left = 3200
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub txt_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 13
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_embarque_LostFocus()
   If Trim(Me.txt_embarque) <> "" Then
      If IsNumeric(Me.txt_embarque) Then
         rs.Open "SELECT * FROM VW_EMBARQUES_FACTURACION_multibondeados where vcha_emp_empresA_id = '" + var_empresa + "' and inte_emb_embarque = " + Me.txt_embarque + " AND CHAR_CAR_ESTATUS <> 'C'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_agente = IIf(IsNull(rs!vcha_age_nombre), "", rs!vcha_age_nombre)
            Me.txt_cliente = IIf(IsNull(rs!vcha_cli_nombre), "", rs!vcha_cli_nombre)
            Me.txt_fecha = IIf(IsNull(rs!DTIM_car_FECHA), "", rs!DTIM_car_FECHA)
            var_importe = 0
            While Not rs.EOF
                  var_importe = var_importe + IIf(IsNull(rs!floa_car_importe_neto), 0, rs!floa_car_importe_neto)
                  rs.MoveNext
            Wend
            Me.txt_importe = Format(var_importe, "###,###,##0.00")
         Else
            Me.txt_agente = ""
            Me.txt_cliente = ""
            Me.txt_embarque = ""
            Me.txt_fecha = ""
            Me.txt_importe = ""
            MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         Me.txt_agente = ""
         Me.txt_cliente = ""
         Me.txt_embarque = ""
         Me.txt_fecha = ""
         Me.txt_importe = ""
         MsgBox "N?mero de embarque incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_fecha_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_importe_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub


