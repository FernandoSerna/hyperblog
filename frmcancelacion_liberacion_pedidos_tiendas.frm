VERSION 5.00
Begin VB.Form frmcancelacion_liberacion_pedidos_tiendas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancelación de la liberación de pedidos de tiendas"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7140
      Picture         =   "frmcancelacion_liberacion_pedidos_tiendas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      Picture         =   "frmcancelacion_liberacion_pedidos_tiendas.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   7
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   105
      TabIndex        =   8
      Top             =   360
      Width           =   7410
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos del pedido "
      Height          =   2025
      Left            =   135
      TabIndex        =   7
      Top             =   435
      Width           =   7335
      Begin VB.TextBox txt_orden_surtido 
         Height          =   350
         Left            =   3855
         TabIndex        =   18
         Top             =   330
         Width           =   1650
      End
      Begin VB.TextBox txt_estatus 
         Height          =   350
         Left            =   3840
         TabIndex        =   6
         Top             =   1485
         Width           =   1650
      End
      Begin VB.TextBox txt_importe 
         Height          =   350
         Left            =   900
         TabIndex        =   5
         Top             =   1485
         Width           =   1380
      End
      Begin VB.TextBox txt_nombre_tienda 
         Height          =   350
         Left            =   2310
         TabIndex        =   4
         Top             =   1095
         Width           =   4905
      End
      Begin VB.TextBox txt_clave_tienda 
         Height          =   350
         Left            =   900
         TabIndex        =   3
         Top             =   1095
         Width           =   1380
      End
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   350
         Left            =   2310
         TabIndex        =   2
         Top             =   705
         Width           =   4905
      End
      Begin VB.TextBox txt_clave_cliente 
         Height          =   350
         Left            =   900
         TabIndex        =   1
         Top             =   705
         Width           =   1380
      End
      Begin VB.TextBox txt_pedido 
         Height          =   350
         Left            =   900
         TabIndex        =   0
         Top             =   330
         Width           =   1380
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Orden de Surtido:"
         Height          =   195
         Left            =   2565
         TabIndex        =   17
         Top             =   375
         Width           =   1245
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Pedido:"
         Height          =   195
         Left            =   270
         TabIndex        =   16
         Top             =   405
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Estatus:"
         Height          =   195
         Left            =   3165
         TabIndex        =   15
         Top             =   1560
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
         Height          =   195
         Left            =   270
         TabIndex        =   14
         Top             =   1560
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tienda:"
         Height          =   195
         Left            =   270
         TabIndex        =   13
         Top             =   1170
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   270
         TabIndex        =   12
         Top             =   780
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pedido:"
         Height          =   195
         Left            =   1050
         TabIndex        =   11
         Top             =   420
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmcancelacion_liberacion_pedidos_tiendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_pedidos_Click()
   If Trim(Me.txt_pedido) <> "" Then
      If Trim(Me.txt_estatus) = "S" Then
         var_si = MsgBox("Desea cancelar la liberación de la orden de surtido", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_si = MsgBox("Confirmar la cancelación de la liberación de la orden de surtido", vbYesNo, "ATENCION")
            If var_si = 6 Then
               rs.Open "update tb_enc_orden_surtido set inte_ors_liberada = 0 where inte_ors_orden_surtido = " + Me.txt_orden_surtido, cnn, adOpenDynamic, adLockOptimistic
               rs.Open "select vcha_cli_referencia from tb_clientes where vcha_Cli_clave_id = '" + Me.txt_clave_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
               var_referencia = Trim(rs(0).Value)
               rs.Close
               rs.Open "CALL SP_AGREGA_ABONO('" + var_referencia + "',0, " + CStr(CDbl(Me.txt_importe)) + ",SYSDATE,SYSDATE,'','','CL','Cancelación de liberación de ordenes de surtido')", cnn_clientes_tiendas, adOpenDynamic, adLockOptimistic
            End If
         End If
      Else
         MsgBox "Ya no ser puede cancelar la liberación del pedido ya que este fue cerrado", vbOKOnly, "ATENCION"
      End If
      Me.txt_clave_cliente = ""
      Me.txt_clave_tienda = ""
      Me.txt_estatus = ""
      Me.txt_importe = ""
      Me.txt_nombre_cliente = ""
      Me.txt_nombre_tienda = ""
      Me.txt_pedido = ""
      Me.txt_orden_surtido = ""
   Else
      MsgBox "No se a seleccionado un pedido", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 2500
   Left = 2200
   If cnn_clientes_tiendas.State = 0 Then
      cnn_clientes_tiendas.Open var_conexion_pedidos_tiendas
      cnn_clientes_tiendas.CursorLocation = adUseClient
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub


Private Sub txt_clave_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_clave_tienda_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_estatus_KeyPress(KeyAscii As Integer)
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

Private Sub txt_nombre_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_tienda_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_orden_surtido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_pedido_KeyPress(KeyAscii As Integer)
  Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_pedido_LostFocus()
   If Trim(Me.txt_pedido) <> "" Then
      If IsNumeric(Me.txt_pedido) Then
         rs.Open "select * from vw_pedidos_tiendas where char_tpe_tipo_pedido_id = 'FT' AND inte_ped_pedido_Credito =  0 and inte_ped_numero = " + Me.txt_pedido + " and inte_ors_liberada = 1", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_clave_cliente = rs!vcha_cli_clave_id
            Me.txt_nombre_cliente = rs!vcha_cli_nombre
            Me.txt_clave_tienda = rs!vcha_age_agente_id
            Me.txt_nombre_tienda = rs!vcha_age_nombre
            Me.txt_importe = Format(IIf(IsNull(rs!importe_pedido), 0, rs!importe_pedido) + IIf(IsNull(rs!importe_seguro), 0, rs!importe_seguro) + IIf(IsNull(rs!importe_paqueteria), 0, rs!importe_paqueteria) + IIf(IsNull(rs!floa_paq_costo_referencia), 0, rs!floa_paq_costo_referencia), "###,###,##0.00")
            Me.txt_estatus = rs!char_ped_estatus
            Me.txt_orden_surtido = rs!inte_ors_orden_surtido
         Else
            MsgBox "El pedido no existe", vbOKOnly, "ATENCION"
            Me.txt_clave_cliente = ""
            Me.txt_clave_tienda = ""
            Me.txt_estatus = ""
            Me.txt_importe = ""
            Me.txt_nombre_cliente = ""
            Me.txt_nombre_tienda = ""
            Me.txt_pedido = ""
            Me.txt_orden_surtido = ""
         End If
         rs.Close
      Else
         MsgBox "Número de pedido incorrecto", vbOKOnly, "ATENCION"
         Me.txt_clave_cliente = ""
         Me.txt_clave_tienda = ""
         Me.txt_estatus = ""
         Me.txt_importe = ""
         Me.txt_nombre_cliente = ""
         Me.txt_nombre_tienda = ""
         Me.txt_pedido = ""
         Me.txt_orden_surtido = ""
      End If
   Else
      Me.txt_clave_cliente = ""
      Me.txt_clave_tienda = ""
      Me.txt_estatus = ""
      Me.txt_importe = ""
      Me.txt_nombre_cliente = ""
      Me.txt_nombre_tienda = ""
      Me.txt_pedido = ""
      Me.txt_orden_surtido = ""
   End If
End Sub
