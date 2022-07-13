VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_lineas_depurar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lineas a depurar"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmoracle_lineas_depurar.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   60
      TabIndex        =   2
      Top             =   360
      Width           =   11565
   End
   Begin VB.Frame Frame1 
      Height          =   5745
      Left            =   60
      TabIndex        =   0
      Top             =   390
      Width           =   11505
      Begin VB.Frame frm_causas_negado 
         Height          =   2835
         Left            =   6630
         TabIndex        =   9
         Top             =   1305
         Width           =   4560
         Begin MSComctlLib.ListView lv_causas_negado 
            Height          =   2430
            Left            =   45
            TabIndex        =   10
            Top             =   330
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   4286
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Clave"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Descripción"
               Object.Width           =   5292
            EndProperty
         End
         Begin VB.Label Label1 
            BackColor       =   &H000000C0&
            Caption         =   " Causas de negado"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   0
            TabIndex        =   11
            Top             =   15
            Width           =   4545
         End
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   390
         Picture         =   "frmoracle_lineas_depurar.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   60
         Picture         =   "frmoracle_lineas_depurar.frx":0318
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1050
         Picture         =   "frmoracle_lineas_depurar.frx":041A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   720
         Picture         =   "frmoracle_lineas_depurar.frx":04EC
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Marcar (Enter)"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1380
         Picture         =   "frmoracle_lineas_depurar.frx":0736
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   135
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_depurar 
         Height          =   5190
         Left            =   30
         TabIndex        =   1
         Top             =   480
         Width           =   11430
         _ExtentX        =   20161
         _ExtentY        =   9155
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "   Código"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Causa negado"
            Object.Width           =   6526
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Clave causa de negado"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "inventory item id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "delivery detail id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "marca"
            Object.Width           =   0
         EndProperty
      End
   End
End
Attribute VB_Name = "frmoracle_lineas_depurar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter

Private Sub cmd_guardar_Click()
   Unload Me
End Sub

Private Sub cmd_invertir_Click()
   n = lv_depurar.ListItems.Count
   For i = 1 To n
      lv_depurar.ListItems.Item(i).Selected = True
      If lv_depurar.selectedItem.SubItems(7) = "*" Then
         lv_depurar.selectedItem.SubItems(7) = ""
         lv_depurar.ListItems.Item(i).Bold = False
         lv_depurar.ListItems.Item(i).ForeColor = &H80000012
         lv_depurar.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_depurar.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_depurar.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_depurar.ListItems.Item(i).ListSubItems(4).Bold = False
         lv_depurar.ListItems.Item(i).ListSubItems(5).Bold = False
         lv_depurar.ListItems.Item(i).ListSubItems(6).Bold = False
         lv_depurar.ListItems.Item(i).ListSubItems(7).Bold = False
         lv_depurar.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_depurar.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_depurar.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
         lv_depurar.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
         lv_depurar.ListItems.Item(i).ListSubItems(5).Bold = False
         lv_depurar.ListItems.Item(i).ListSubItems(6).Bold = False
         lv_depurar.ListItems.Item(i).ListSubItems(7).Bold = False
      Else
         lv_depurar.selectedItem.SubItems(7) = "*"
         lv_depurar.ListItems.Item(i).Bold = True
         lv_depurar.ListItems.Item(i).ForeColor = &HFF0000
         lv_depurar.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_depurar.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_depurar.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_depurar.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_depurar.ListItems.Item(i).ListSubItems(5).Bold = True
         lv_depurar.ListItems.Item(i).ListSubItems(6).Bold = True
         lv_depurar.ListItems.Item(i).ListSubItems(7).Bold = True
         lv_depurar.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_depurar.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_depurar.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_depurar.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
         lv_depurar.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
         lv_depurar.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
         lv_depurar.ListItems.Item(i).ListSubItems(7).ForeColor = &HFF0000
      End If
   Next i
   If Me.lv_depurar.ListItems.Count > 0 Then
      Me.lv_depurar.SetFocus
   End If
End Sub

Private Sub cmd_marcar_Click()
   i = lv_depurar.selectedItem.Index
   If lv_depurar.selectedItem.SubItems(7) = "*" Then
      lv_depurar.selectedItem.SubItems(7) = ""
      lv_depurar.ListItems.Item(i).Bold = False
      lv_depurar.ListItems.Item(i).ForeColor = &H80000012
      lv_depurar.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_depurar.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_depurar.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_depurar.ListItems.Item(i).ListSubItems(4).Bold = False
      lv_depurar.ListItems.Item(i).ListSubItems(5).Bold = False
      lv_depurar.ListItems.Item(i).ListSubItems(6).Bold = False
      lv_depurar.ListItems.Item(i).ListSubItems(7).Bold = False
      lv_depurar.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_depurar.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_depurar.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      lv_depurar.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
      lv_depurar.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
      lv_depurar.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
      lv_depurar.ListItems.Item(i).ListSubItems(7).ForeColor = &H80000012
      lv_depurar.Refresh
   Else
      lv_depurar.selectedItem.SubItems(7) = "*"
      lv_depurar.ListItems.Item(i).Bold = True
      lv_depurar.ListItems.Item(i).ForeColor = &HFF0000
      lv_depurar.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_depurar.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_depurar.ListItems.Item(i).ListSubItems(3).Bold = True
      lv_depurar.ListItems.Item(i).ListSubItems(4).Bold = True
      lv_depurar.ListItems.Item(i).ListSubItems(5).Bold = True
      lv_depurar.ListItems.Item(i).ListSubItems(6).Bold = True
      lv_depurar.ListItems.Item(i).ListSubItems(7).Bold = True
      lv_depurar.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_depurar.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_depurar.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
      lv_depurar.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
      lv_depurar.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
      lv_depurar.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
      lv_depurar.ListItems.Item(i).ListSubItems(7).ForeColor = &HFF0000
      lv_depurar.Refresh
   End If
   If Me.lv_depurar.ListItems.Count > 0 Then
      Me.lv_depurar.SetFocus
   End If
End Sub

Private Sub cmd_ninguno_Click()
   n = lv_depurar.ListItems.Count
   For i = 1 To n
      lv_depurar.ListItems.Item(i).Selected = True
      lv_depurar.selectedItem.SubItems(7) = ""
      lv_depurar.ListItems.Item(i).Bold = False
      lv_depurar.ListItems.Item(i).ForeColor = &H80000012
      lv_depurar.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_depurar.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_depurar.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_depurar.ListItems.Item(i).ListSubItems(4).Bold = False
      lv_depurar.ListItems.Item(i).ListSubItems(5).Bold = False
      lv_depurar.ListItems.Item(i).ListSubItems(6).Bold = False
      lv_depurar.ListItems.Item(i).ListSubItems(7).Bold = False
      lv_depurar.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_depurar.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_depurar.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      lv_depurar.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
      lv_depurar.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
      lv_depurar.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
      lv_depurar.ListItems.Item(i).ListSubItems(7).ForeColor = &H80000012
   Next i
   lv_depurar.Refresh
   If Me.lv_depurar.ListItems.Count > 0 Then
      Me.lv_depurar.SetFocus
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmd_seleccion_Click()
   n = lv_depurar.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_depurar.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_depurar.selectedItem.SubItems(7) = "" And var_rellena = True Then
         lv_depurar.selectedItem.SubItems(7) = "*"
         lv_depurar.ListItems.Item(i).Bold = True
         lv_depurar.ListItems.Item(i).ForeColor = &HFF0000
         lv_depurar.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_depurar.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_depurar.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_depurar.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_depurar.ListItems.Item(i).ListSubItems(5).Bold = True
         lv_depurar.ListItems.Item(i).ListSubItems(6).Bold = True
         lv_depurar.ListItems.Item(i).ListSubItems(7).Bold = True
         lv_depurar.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_depurar.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_depurar.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_depurar.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
         lv_depurar.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
         lv_depurar.ListItems.Item(i).ListSubItems(7).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_depurar.selectedItem.SubItems(7) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_depurar.selectedItem.SubItems(7) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
   If Me.lv_depurar.ListItems.Count > 0 Then
      Me.lv_depurar.SetFocus
   End If
End Sub

Private Sub cmd_todos_Click()
   n = lv_depurar.ListItems.Count
   For i = 1 To n
      lv_depurar.ListItems.Item(i).Selected = True
      lv_depurar.selectedItem.SubItems(7) = "*"
      lv_depurar.ListItems.Item(i).Bold = True
      lv_depurar.ListItems.Item(i).ForeColor = &HFF0000
      lv_depurar.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_depurar.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_depurar.ListItems.Item(i).ListSubItems(3).Bold = True
      lv_depurar.ListItems.Item(i).ListSubItems(4).Bold = True
      lv_depurar.ListItems.Item(i).ListSubItems(5).Bold = True
      lv_depurar.ListItems.Item(i).ListSubItems(6).Bold = True
      lv_depurar.ListItems.Item(i).ListSubItems(7).Bold = True
      lv_depurar.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_depurar.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_depurar.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
      lv_depurar.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
      lv_depurar.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
      lv_depurar.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
      lv_depurar.ListItems.Item(i).ListSubItems(7).ForeColor = &HFF0000
   Next i
   lv_depurar.Refresh
   If Me.lv_depurar.ListItems.Count > 0 Then
      Me.lv_depurar.SetFocus
   End If
End Sub

Private Sub Form_Load()
   Me.frm_causas_negado.Visible = False
   If var_lote_depurar = 0 Then
      strconsulta = "select a.DELIVERY_DETAIL_ID, a.INVENTORY_ITEM_ID, a.SOURCE_HEADER_NUMBER, a.SEGMENT1 as codigo, a.FECHA_NEGADO, a.CAUSA_NEGADO, a.NOMBRE_CAUSA_NEGADO, a.Cantidad, a.ORGANIZATION_ID, a.LOTE, b.description as descripcion from xxvia_tb_negado_distribucion a, xxvia_system_items_b b where SOURCE_HEADER_NUMBER = ? and a.inventory_item_id = b.inventory_item_id and a.organization_id = b.organization_id and cantidad > 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(var_orden_depurar))
           .Parameters.Append parametro
      End With
      Set rsaux8 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      While Not rsaux8.EOF
            Set list_item = Me.lv_depurar.ListItems.Add(, , IIf(IsNull(rsaux8!CODIGO), "", rsaux8!CODIGO))
            list_item.SubItems(1) = IIf(IsNull(rsaux8!DESCRIPCION), "", rsaux8!DESCRIPCION)
            list_item.SubItems(2) = Format(IIf(IsNull(rsaux8!Cantidad), 0, rsaux8!Cantidad), "###,###,##0.00")
            list_item.SubItems(3) = IIf(IsNull(rsaux8!nombre_causa_negado), "", rsaux8!nombre_causa_negado)
            list_item.SubItems(4) = IIf(IsNull(rsaux8!CAUSA_NEGADO), "", rsaux8!CAUSA_NEGADO)
            list_item.SubItems(5) = IIf(IsNull(rsaux8!inventory_item_id), 0, rsaux8!inventory_item_id)
            list_item.SubItems(6) = IIf(IsNull(rsaux8!delivery_detail_id), "", rsaux8!delivery_detail_id)
            rsaux8.MoveNext
      Wend
      rsaux8.Close

   Else
      strconsulta = "select a.DELIVERY_DETAIL_ID, a.INVENTORY_ITEM_ID, a.SOURCE_HEADER_NUMBER, a.SEGMENT1 as codigo, a.FECHA_NEGADO, a.CAUSA_NEGADO, a.NOMBRE_CAUSA_NEGADO, a.Cantidad, a.ORGANIZATION_ID, a.LOTE, b.description as descripcion from xxvia_tb_negado_distribucion a, xxvia_system_items_b b where SOURCE_HEADER_NUMBER = ? and a.inventory_item_id = b.inventory_item_id and a.organization_id = b.organization_id and cantidad > 0 and lote = ?"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(var_orden_depurar))
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(var_lote_depurar))
           .Parameters.Append parametro
      End With
      Set rsaux8 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      While Not rsaux8.EOF
            Set list_item = Me.lv_depurar.ListItems.Add(, , IIf(IsNull(rsaux8!CODIGO), "", rsaux8!CODIGO))
            list_item.SubItems(1) = IIf(IsNull(rsaux8!DESCRIPCION), "", rsaux8!DESCRIPCION)
            list_item.SubItems(2) = Format(IIf(IsNull(rsaux8!Cantidad), 0, rsaux8!Cantidad), "###,###,##0.00")
            list_item.SubItems(3) = IIf(IsNull(rsaux8!nombre_causa_negado), "", rsaux8!nombre_causa_negado)
            list_item.SubItems(4) = IIf(IsNull(rsaux8!CAUSA_NEGADO), "", rsaux8!CAUSA_NEGADO)
            list_item.SubItems(5) = IIf(IsNull(rsaux8!inventory_item_id), 0, rsaux8!inventory_item_id)
            list_item.SubItems(6) = IIf(IsNull(rsaux8!delivery_detail_id), "", rsaux8!delivery_detail_id)
            rsaux8.MoveNext
      Wend
      rsaux8.Close
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If var_lote_depurar = 0 Then
       For var_j = 1 To Me.lv_depurar.ListItems.Count
           Me.lv_depurar.ListItems.Item(var_j).Selected = True
           strconsulta = "UPDATE xxvia_tb_negado_distribucion SET CAUSA_NEGADO = ?, NOMBRE_CAUSA_NEGADO = ? where SOURCE_HEADER_NUMBER = ? AND inventory_item_id = ? and DELIVERY_DETAIL_ID = ?"
           With comandoORA
                .ActiveConnection = cnnoracle_4
                .CommandType = adCmdText
                .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, Me.lv_depurar.selectedItem.SubItems(3))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, Me.lv_depurar.selectedItem.SubItems(4))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(var_orden_depurar))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(Me.lv_depurar.selectedItem.SubItems(5)))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(Me.lv_depurar.selectedItem.SubItems(6)))
                 .Parameters.Append parametro
            End With
            Set rsaux8 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
       Next var_j
    Else
       For var_j = 1 To Me.lv_depurar.ListItems.Count
           Me.lv_depurar.ListItems.Item(var_j).Selected = True
           strconsulta = "UPDATE xxvia_tb_negado_distribucion SET CAUSA_NEGADO = ?, NOMBRE_CAUSA_NEGADO = ? where SOURCE_HEADER_NUMBER = ? AND inventory_item_id = ? and DELIVERY_DETAIL_ID = ? and lote = ?"
           With comandoORA
                .ActiveConnection = cnnoracle_4
                .CommandType = adCmdText
                .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, Me.lv_depurar.selectedItem.SubItems(3))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, Me.lv_depurar.selectedItem.SubItems(4))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(var_orden_depurar))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(Me.lv_depurar.selectedItem.SubItems(5)))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(Me.lv_depurar.selectedItem.SubItems(6)))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(var_lote_depurar))
                 .Parameters.Append parametro
            End With
            Set rsaux8 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
       Next var_j
    
    End If
End Sub

Private Sub lv_causas_negado_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_causas_negado, ColumnHeader)
End Sub

Private Sub lv_causas_negado_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      If Me.lv_depurar.ListItems.Count > 0 Then
         Me.lv_depurar.SetFocus
      Else
         Me.frm_causas_negado.Visible = False
      End If
   End If
   If KeyAscii = 13 Then
      var_si = MsgBox("¿Desea actualizar el registos?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         For var_j = 1 To Me.lv_depurar.ListItems.Count
             Me.lv_depurar.ListItems.Item(var_j).Selected = True
             If Me.lv_depurar.selectedItem.SubItems(7) = "*" Then
                Me.lv_depurar.selectedItem.SubItems(4) = Me.lv_causas_negado.selectedItem.SubItems(1)
                Me.lv_depurar.selectedItem.SubItems(3) = Me.lv_causas_negado.selectedItem
                lv_depurar.selectedItem.SubItems(7) = ""
                lv_depurar.ListItems.Item(var_j).Bold = False
                lv_depurar.ListItems.Item(var_j).ForeColor = &H80000012
                lv_depurar.ListItems.Item(var_j).ListSubItems(1).Bold = False
                lv_depurar.ListItems.Item(var_j).ListSubItems(2).Bold = False
                lv_depurar.ListItems.Item(var_j).ListSubItems(3).Bold = False
                lv_depurar.ListItems.Item(var_j).ListSubItems(4).Bold = False
                lv_depurar.ListItems.Item(var_j).ListSubItems(5).Bold = False
                lv_depurar.ListItems.Item(var_j).ListSubItems(6).Bold = False
                lv_depurar.ListItems.Item(var_j).ListSubItems(7).Bold = False
                lv_depurar.ListItems.Item(var_j).ListSubItems(1).ForeColor = &H80000012
                lv_depurar.ListItems.Item(var_j).ListSubItems(2).ForeColor = &H80000012
                lv_depurar.ListItems.Item(var_j).ListSubItems(3).ForeColor = &H80000012
                lv_depurar.ListItems.Item(var_j).ListSubItems(4).ForeColor = &H80000012
                lv_depurar.ListItems.Item(var_j).ListSubItems(5).ForeColor = &H80000012
                lv_depurar.ListItems.Item(var_j).ListSubItems(6).ForeColor = &H80000012
                lv_depurar.ListItems.Item(var_j).ListSubItems(7).ForeColor = &H80000012
                lv_depurar.Refresh
             End If
         Next var_j
         Me.frm_causas_negado.Visible = False
      End If
   End If
End Sub

Private Sub lv_causas_negado_LostFocus()
   Me.frm_causas_negado.Visible = False
End Sub

Private Sub lv_depurar_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_depurar, ColumnHeader)
End Sub

Private Sub lv_depurar_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 115 Then
      Me.frm_causas_negado.Visible = True
      Me.lv_causas_negado.ListItems.Clear
      rsaux.Open "select lookup_code as causa_negado, meaning as descripcion from fnd_lookup_values where lookup_type = 'CANCEL_CODE' and language = 'US' AND UPPER(lookup_code) NOT IN  ('EDI CANCELLATION','1','ADMIN ERROR','CONFIGURATOR','CREDIT PROBLEM','DISCONTINUED','IR_ISO_CMS_CHG','LATE','NOT PROVIDED','2','EDI CANCELLATION','SYSTEM','VIGENCIA','XXVIA PEDIDO CERRADO') AND lookup_code LIKE 'XXVIA%' ORDER BY meaning", cnnoracle_4, adOpenDynamic, adLockOptimistic
      While Not rsaux.EOF
         Set list_item = Me.lv_causas_negado.ListItems.Add(, , rsaux!CAUSA_NEGADO)
         list_item.SubItems(1) = IIf(IsNull(rsaux!DESCRIPCION), "", rsaux!DESCRIPCION)
         rsaux.MoveNext
      Wend
      rsaux.Close
      Me.lv_causas_negado.SetFocus
   End If
End Sub

Private Sub lv_depurar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      i = lv_depurar.selectedItem.Index
      If lv_depurar.selectedItem.SubItems(7) = "*" Then
         lv_depurar.selectedItem.SubItems(7) = ""
         lv_depurar.ListItems.Item(i).Bold = False
         lv_depurar.ListItems.Item(i).ForeColor = &H80000012
         lv_depurar.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_depurar.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_depurar.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_depurar.ListItems.Item(i).ListSubItems(4).Bold = False
         lv_depurar.ListItems.Item(i).ListSubItems(5).Bold = False
         lv_depurar.ListItems.Item(i).ListSubItems(6).Bold = False
         lv_depurar.ListItems.Item(i).ListSubItems(7).Bold = False
         lv_depurar.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_depurar.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_depurar.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
         lv_depurar.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
         lv_depurar.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
         lv_depurar.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
         lv_depurar.ListItems.Item(i).ListSubItems(7).ForeColor = &H80000012
         lv_depurar.Refresh
      Else
         lv_depurar.selectedItem.SubItems(7) = "*"
         lv_depurar.ListItems.Item(i).Bold = True
         lv_depurar.ListItems.Item(i).ForeColor = &HFF0000
         lv_depurar.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_depurar.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_depurar.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_depurar.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_depurar.ListItems.Item(i).ListSubItems(5).Bold = True
         lv_depurar.ListItems.Item(i).ListSubItems(6).Bold = True
         lv_depurar.ListItems.Item(i).ListSubItems(7).Bold = True
         lv_depurar.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_depurar.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_depurar.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_depurar.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
         lv_depurar.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
         lv_depurar.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
         lv_depurar.ListItems.Item(i).ListSubItems(7).ForeColor = &HFF0000
         lv_depurar.Refresh
      End If
      If Me.lv_depurar.ListItems.Count > 0 Then
         Me.lv_depurar.SetFocus
      End If
   End If
End Sub
