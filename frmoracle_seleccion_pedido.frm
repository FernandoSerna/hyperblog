VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_seleccion_pedido 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección de pedidos"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lv_salidas 
      Height          =   3345
      Left            =   30
      TabIndex        =   0
      Top             =   45
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   5900
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "   Agente"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cliente"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Pedido"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Máquina"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Orden"
         Object.Width           =   1235
      EndProperty
   End
End
Attribute VB_Name = "frmoracle_seleccion_pedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   rs.Open "select * from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(var_embarque_global), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            var_agente_nombre = ""
            If rsaux1.State = 1 Then
               rsaux1.Close
            End If
            rsaux1.Open "select * from  tb_oracle_pedidos_maquinas where pedido = " + CStr(rs!pedido), cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux1.EOF Then
               If rsaux1!maquina = fun_NombrePc Then
                  var_agente_nombre = IIf(IsNull(rs!nombre_agente), "", rs!nombre_agente)
                  Set list_item = lv_salidas.ListItems.Add(, , var_agente_nombre)
                  list_item.SubItems(1) = IIf(IsNull(rs!Cliente), "", rs!Cliente)
                  list_item.SubItems(2) = IIf(IsNull(rs!pedido), 0, rs!pedido)
                  list_item.SubItems(3) = IIf(IsNull(rsaux1!maquina), "", rsaux1!maquina)
                  list_item.SubItems(4) = IIf(IsNull(rs!orden_pedido), "", rs!orden_pedido)
               End If
            Else
               var_agente_nombre = IIf(IsNull(rs!nombre_agente), "", rs!nombre_agente)
               Set list_item = lv_salidas.ListItems.Add(, , var_agente_nombre)
               list_item.SubItems(1) = IIf(IsNull(rs!Cliente), "", rs!Cliente)
               list_item.SubItems(2) = IIf(IsNull(rs!pedido), 0, rs!pedido)
               list_item.SubItems(3) = ""
               list_item.SubItems(4) = IIf(IsNull(rs!orden_pedido), "", rs!orden_pedido)
            End If
            rsaux1.Close
            rs.MoveNext
      Wend
   End If
   rs.Close
End Sub

Private Sub lv_salidas_KeyPress(KeyAscii As Integer)
   If Me.lv_salidas.ListItems.Count > 0 Then
      var_pedido_global = CDbl(Me.lv_salidas.selectedItem.SubItems(2))
      Unload Me
   End If
End Sub
