VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmdepositos_mal_aplicados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Depositos mal aplicados"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9645
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9045
      Picture         =   "frmdepositos_mal_aplicados.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      Picture         =   "frmdepositos_mal_aplicados.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Actualizar Alt + A"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   120
      Left            =   105
      TabIndex        =   4
      Top             =   300
      Width           =   9300
   End
   Begin VB.Frame Frame2 
      Caption         =   " Aplicaciones "
      Height          =   3705
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   9285
      Begin MSComctlLib.ListView lv_aplicaciones 
         Height          =   3420
         Left            =   90
         TabIndex        =   3
         Top             =   210
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   6033
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
            Text            =   "Referencia"
            Object.Width           =   4022
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Importe"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha de aplicación"
            Object.Width           =   4022
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Factura"
            Object.Width           =   4022
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Depositos"
      Height          =   2625
      Left            =   120
      TabIndex        =   0
      Top             =   465
      Width           =   9285
      Begin MSComctlLib.ListView lv_depositos 
         Height          =   2295
         Left            =   90
         TabIndex        =   2
         Top             =   225
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   4048
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Deposito"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Referencia"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cliente"
            Object.Width           =   5115
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Importe"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Saldo"
            Object.Width           =   2117
         EndProperty
      End
   End
End
Attribute VB_Name = "frmdepositos_mal_aplicados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim numero_items_lineas As Integer
Private Sub cmd_guardar_Click()
   Me.lv_depositos.ListItems.Clear
   rs.Open "SELECT * FROM VW_ESTATUS_APLICACIONES", cnnoracle_2, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set list_item = lv_depositos.ListItems.Add(, , rs(0).Value)
         list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
         list_item.SubItems(2) = IIf(IsNull(rs(2).Value), "", rs(2).Value)
         rsaux.Open "select * from tb_clientes where vcha_cli_referencia = '" + rs(2).Value + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            list_item.SubItems(3) = IIf(IsNull(rsaux!vcha_cli_nombre), "", rsaux!vcha_cli_nombre)
         Else
            list_item.SubItems(3) = ""
         End If
         rsaux.Close
         list_item.SubItems(4) = Format(IIf(IsNull(rs(3).Value), 0, rs(3).Value), "###,###,##0.00")
         list_item.SubItems(5) = IIf(IsNull(rs(4).Value), 0, rs(4).Value)
         rs.MoveNext
   Wend
   rs.Close
   Me.lv_aplicaciones.ListItems.Clear
   
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   Top = 200
   Left = 1000
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub lv_depositos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.lv_aplicaciones.ListItems.Clear
      If Me.lv_depositos.ListItems.Count > 0 Then
         var_importe = CDbl(Me.lv_depositos.selectedItem.SubItems(4))
         rs.Open "SELECT vcha_car_referencia as Referencia, numb_car_importe as importe, date_car_fecha_cargo , vcha_car_num_docum  FROM TB_CARGO WHERE INTE_CAR_ABONO_ID = " + Me.lv_depositos.selectedItem, cnnoracle_2, adOpenDynamic, adLockOptimistic
         var_importe_facturas = 0
         While Not rs.EOF
               Set list_item = lv_aplicaciones.ListItems.Add(, , rs(0).Value)
               list_item.SubItems(1) = Format(IIf(IsNull(rs(1).Value), 0, rs(1).Value), "###,###,##0.00")
               list_item.SubItems(2) = IIf(IsNull(rs(2).Value), "", rs(2).Value)
               list_item.SubItems(3) = IIf(IsNull(rs(3).Value), "", rs(3).Value)
               var_importe_facturas = var_importe_facturas + IIf(IsNull(rs(1).Value), 0, rs(1).Value)
               If var_importe_facturas > var_importe Then
                  list_item.ForeColor = &HC0&
                  list_item.ListSubItems.Item(1).ForeColor = &HC0&
                  list_item.ListSubItems.Item(2).ForeColor = &HC0&
                  list_item.ListSubItems.Item(3).ForeColor = &HC0&
               End If
               rs.MoveNext
         Wend
         rs.Close
      End If
   End If
End Sub
