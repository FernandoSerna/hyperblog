VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_peso_volumen_embarques 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Volumen y peso por pedido"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   885
      Left            =   75
      TabIndex        =   2
      Top             =   0
      Width           =   5475
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
         Height          =   480
         Left            =   2115
         TabIndex        =   4
         Top             =   285
         Width           =   1890
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pedido:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   600
         TabIndex        =   3
         Top             =   345
         Width           =   915
      End
   End
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   30
      TabIndex        =   1
      Top             =   375
      Width           =   5520
   End
   Begin VB.Frame Frame1 
      Height          =   2880
      Left            =   90
      TabIndex        =   0
      Top             =   870
      Width           =   5475
      Begin VB.TextBox txt_cantidad 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1590
         TabIndex        =   9
         Top             =   2325
         Width           =   1155
      End
      Begin VB.TextBox txt_total_volumen 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3915
         TabIndex        =   8
         Top             =   2325
         Width           =   1110
      End
      Begin VB.TextBox txt_total_peso 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2775
         TabIndex        =   6
         Top             =   2325
         Width           =   1110
      End
      Begin MSComctlLib.ListView lv_pedidos 
         Height          =   2130
         Left            =   60
         TabIndex        =   5
         Top             =   150
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   3757
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Pedido"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Piezas"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Peso"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Volumen"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         TabIndex        =   7
         Top             =   2385
         Width           =   690
      End
   End
End
Attribute VB_Name = "frmoracle_peso_volumen_embarques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comandoORA As New adodb.Command
Dim parametro As adodb.Parameter
Dim objConn As New adodb.Connection
Dim objCmd As New adodb.Command
Dim objParm As adodb.Parameter

Private Sub Form_Load()
   Top = 2000
   Left = 2000
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_embarque_Change()
   Me.lv_pedidos.ListItems.Clear
   Me.txt_total_peso = ""
   Me.txt_total_volumen = ""
   Me.txt_cantidad = ""
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_embarque) Then
         
         'strconsulta = "select source_header_number, sum(nvl(floa_sal_Cantidad_leida,0)) as cantidad,sum(floa_sal_cantidad_leida * unit_weight) as peso, sum(floa_sal_cantidad_leida * unit_volume) as volumen from xxvia_tb_salidas_cajas a, xxvia_system_items_b  b where a.segment1 = b.segment1 and b.organization_id = ? and inte_emb_embarque = ? group by source_header_number"
         strconsulta = "select header_id from oe_order_headers_all where order_number = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(Me.txt_embarque))
              .Parameters.Append parametro
         End With
         Set rsaux = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         If Not rsaux.EOF Then
            var_header = rsaux(0).Value
         
            strconsulta = "select sum(nvl(ordered_quantity + cancelled_quantity,0)) as cantidad,sum((ordered_quantity + cancelled_quantity) * unit_weight) as peso, sum((ordered_quantity + cancelled_quantity) * unit_volume) as volumen from oe_order_lines_all a, xxvia_system_items_b  b where a.ordered_item = b.segment1 and b.organization_id = ? and header_id = ? group by header_id"
            'strconsulta = "select sum(nvl(ordered_quantity,0)) as cantidad,sum((ordered_quantity + cancelled_quantity) * unit_weight) as peso, sum((ordered_quantity + cancelled_quantity) * unit_volume) as volumen from oe_order_lines_all a, xxvia_system_items_b  b where a.ordered_item = b.segment1 and b.organization_id = ? and header_id = ? group by header_id"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(var_unidad_organizacional))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(var_header))
                 .Parameters.Append parametro
            End With
            Set rs = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            var_cantidad = 0
            var_peso = 0
            var_volumen = 0
            If Not rs.EOF Then
               While Not rs.EOF
                     Set list_item = lv_pedidos.ListItems.Add(, , Me.txt_embarque)
                     list_item.SubItems(1) = IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)
                     var_cantidad = var_cantidad + IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)
                     list_item.SubItems(2) = IIf(IsNull(rs!PESO), 0, rs!PESO)
                     var_peso = var_peso + IIf(IsNull(rs!PESO), 0, rs!PESO)
                     list_item.SubItems(3) = Format(IIf(IsNull(rs!vOLUMEN), 0, rs!vOLUMEN), "###,###,##0.0000")
                     var_volumen = var_volumen + IIf(IsNull(rs!vOLUMEN), 0, rs!vOLUMEN)
                     rs.MoveNext
               Wend
               Me.txt_cantidad = Format(var_cantidad, "###,###,##0.00")
               Me.txt_total_peso = Format(var_peso, "###,###,##0.00")
               Me.txt_total_volumen = Format(var_volumen, "###,###,##0.0000")
            Else
               MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
            End If
            rs.Close
         Else
            MsgBox "El pedido no existe", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub

Private Sub txt_total_peso_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_total_volumen.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_total_volumen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_embarque.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub
