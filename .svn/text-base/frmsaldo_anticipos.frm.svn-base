VERSION 5.00
Begin VB.Form frmsaldo_anticipos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Saldo de anticipos"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1200
      Left            =   75
      TabIndex        =   3
      Top             =   60
      Width           =   3735
      Begin VB.CommandButton cmd_aceptar_pedidos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3270
         Picture         =   "frmsaldo_anticipos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Aceptar Alt + A"
         Top             =   705
         Width           =   330
      End
      Begin VB.TextBox txt_aplicar 
         Alignment       =   1  'Right Justify
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
         Left            =   975
         TabIndex        =   0
         Top             =   645
         Width           =   2250
      End
      Begin VB.TextBox txt_saldo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   975
         TabIndex        =   2
         Top             =   165
         Width           =   2250
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Aplicar:"
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
         Left            =   120
         TabIndex        =   5
         Top             =   742
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Saldo:"
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
         Left            =   105
         TabIndex        =   4
         Top             =   262
         Width           =   690
      End
   End
End
Attribute VB_Name = "frmsaldo_anticipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_pedidos_Click()
   If IsNumeric(Me.txt_aplicar) Then
      If CDbl(Me.txt_aplicar) <= CDbl(Me.txt_saldo) Then
         var_importe_anticipo = CDbl(Me.txt_aplicar)
         Unload Me
      Else
         MsgBox "El importe del anticipo a aplicar debe de ser menor o igual al saldo del anticipo", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Importe de anticipo es incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Form_Load()
   rs.Open "select sum(a.floa_sal_cantidad * a.floa_sal_precio * (1-a.floa_sal_descuento_1/100) * (1-a.floa_sal_descuento_2/100) * (1+a.floa_iva_iva/100)) as importe from  tb_salidas a inner join tb_encabezado_movimientos b on a.vcha_emp_empresa_id = b.vcha_emp_empresa_id and a.vcha_uor_unidad_id = b.vcha_uor_unidad_id and a.vcha_mov_movimiento_id = b.vcha_mov_movimiento_id And a.inte_sal_numero = b.inte_emo_numero where (a.vcha_art_articulo_id = 'S1003' or a.vcha_art_articulo_id = 'S1005')  and b.vcha_cli_clave_id = '" + var_cliente_anticipo + "' and b.vcha_mov_movimiento_id = 'VDI' group by b.vcha_mov_movimiento_id,b.vcha_emo_afectacion"
   If Not rs.EOF Then
      var_a = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
   Else
      var_a = 0
   End If
   rs.Close
   
   rs.Open "select sum(a.floa_ent_cantidad * a.floa_ent_precio * (1-c.floa_cde_descuento_1/100) * (1-c.floa_cde_descuento_2/100)* (1+c.floa_cde_iva/100)) as importe  from  tb_entradas a inner join tb_encabezado_movimientos b on a.vcha_emp_empresa_id = b.vcha_emp_empresa_id and a.vcha_uor_unidad_id = b.vcha_uor_unidad_id and a.vcha_mov_movimiento_id = b.vcha_mov_movimiento_id And a.inte_ent_numero = b.inte_emo_numero inner join tb_devoluciones c on a.vcha_emp_empresa_id = c.vcha_emp_empresa_id and a.vcha_uor_unidad_id = c.vcha_uor_unidad_id and a.vcha_mov_movimiento_id = c.vcha_mov_movimiento_id And a.inte_ent_numero = c.inte_emo_numero where (a.vcha_art_articulo_id = 'S1003' or a.vcha_art_articulo_id = 'S1005')  and b.vcha_cli_clave_id = '" + var_cliente_anticipo + "' and b.vcha_mov_movimiento_id = 'CA' group by b.vcha_mov_movimiento_id,b.vcha_emo_afectacion "
   If Not rs.EOF Then
      var_b = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
   Else
      var_b = 0
   End If
   rs.Close


   Me.txt_saldo = Format(var_a - var_b, "###,###,##0.00")
End Sub

Private Sub txt_aplicar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Me.cmd_aceptar_pedidos.SetFocus
    Else
       If KeyAscii = 27 Then
          Unload Me
       End If
    End If
End Sub

Private Sub txt_saldo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_aplicar.SetFocus
   Else
      If KeyAscii = 27 Then
         Unload Me
      Else
          KeyAscii = 0
      End If
   End If
End Sub
