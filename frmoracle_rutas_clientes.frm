VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_rutas_clientes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignar clientes a rutas"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frm_rutas 
      Height          =   3150
      Left            =   0
      TabIndex        =   22
      Top             =   960
      Width           =   8400
      Begin MSComctlLib.ListView lv_rutas 
         Height          =   2670
         Left            =   45
         TabIndex        =   23
         Top             =   435
         Width           =   8310
         _ExtentX        =   14658
         _ExtentY        =   4710
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Ruta"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre Ruta"
            Object.Width           =   10583
         EndProperty
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         Caption         =   " Rutas"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   24
         Top             =   120
         Width           =   8325
      End
   End
   Begin VB.Frame frm_lista 
      Height          =   3150
      Left            =   75
      TabIndex        =   19
      Top             =   1965
      Width           =   8400
      Begin VB.CommandButton cmd_pasar 
         Height          =   315
         Left            =   1725
         Picture         =   "frmoracle_rutas_clientes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   390
         Width           =   345
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   390
         Picture         =   "frmoracle_rutas_clientes.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   390
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   60
         Picture         =   "frmoracle_rutas_clientes.frx":0318
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   390
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1050
         Picture         =   "frmoracle_rutas_clientes.frx":041A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   390
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   720
         Picture         =   "frmoracle_rutas_clientes.frx":04EC
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Marcar (Enter)"
         Top             =   390
         Width           =   330
      End
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1380
         Picture         =   "frmoracle_rutas_clientes.frx":0736
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   390
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_lista 
         Height          =   2325
         Left            =   45
         TabIndex        =   20
         Top             =   750
         Width           =   8310
         _ExtentX        =   14658
         _ExtentY        =   4101
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cliente"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre cliente"
            Object.Width           =   11994
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         Caption         =   " Clientes"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   21
         Top             =   120
         Width           =   8325
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5280
      Left            =   75
      TabIndex        =   17
      Top             =   1950
      Width           =   8310
      Begin MSComctlLib.ListView lv_clientes 
         Height          =   4725
         Left            =   60
         TabIndex        =   5
         Top             =   495
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   8334
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre Ruta"
            Object.Width           =   10142
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Prioridad"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000FF&
         Caption         =   "  Clientes"
         ForeColor       =   &H8000000E&
         Height          =   315
         Left            =   45
         TabIndex        =   18
         Top             =   135
         Width           =   8205
      End
   End
   Begin VB.Frame Frame6 
      Height          =   75
      Left            =   30
      TabIndex        =   16
      Top             =   330
      Width           =   8340
   End
   Begin VB.CommandButton com_nuevo_orden 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   45
      Picture         =   "frmoracle_rutas_clientes.frx":094C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton com_guardar 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   375
      Picture         =   "frmoracle_rutas_clientes.frx":0A4E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   1545
      Left            =   75
      TabIndex        =   12
      Top             =   405
      Width           =   8295
      Begin VB.TextBox txt_clave 
         Height          =   390
         Left            =   945
         TabIndex        =   2
         Top             =   165
         Width           =   1815
      End
      Begin VB.TextBox txt_nombre 
         Height          =   420
         Left            =   945
         TabIndex        =   3
         Top             =   585
         Width           =   7230
      End
      Begin VB.TextBox txt_prioridad 
         Height          =   390
         Left            =   945
         TabIndex        =   4
         Top             =   1035
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   263
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   705
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Prioridad:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1140
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmoracle_rutas_clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim comandoORA As New adodb.Command
Dim parametro As adodb.Parameter


Private Sub cmd_invertir_Click()
   n = lv_lista.ListItems.Count
   For i = 1 To n
      lv_lista.ListItems.Item(i).Selected = True
      If lv_lista.selectedItem.SubItems(2) = "*" Then
         lv_lista.selectedItem.SubItems(2) = ""
         lv_lista.ListItems.Item(i).Bold = False
         lv_lista.ListItems.Item(i).ForeColor = &H80000012
         lv_lista.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_lista.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_lista.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_lista.selectedItem.SubItems(2) = "*"
         lv_lista.ListItems.Item(i).Bold = True
         lv_lista.ListItems.Item(i).ForeColor = &HFF0000
         lv_lista.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_lista.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_lista.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub cmd_marcar_Click()
   i = lv_lista.selectedItem.Index
   If lv_lista.selectedItem.SubItems(2) = "*" Then
      lv_lista.selectedItem.SubItems(2) = ""
      lv_lista.ListItems.Item(i).Bold = False
      lv_lista.ListItems.Item(i).ForeColor = &H80000012
      lv_lista.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_lista.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_lista.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_lista.Refresh
   Else
      lv_lista.selectedItem.SubItems(2) = "*"
      lv_lista.ListItems.Item(i).Bold = True
      lv_lista.ListItems.Item(i).ForeColor = &HFF0000
      lv_lista.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_lista.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_lista.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_lista.Refresh
   End If
End Sub

Private Sub cmd_ninguno_Click()
   n = lv_lista.ListItems.Count
   For i = 1 To n
      lv_lista.ListItems.Item(i).Selected = True
      lv_lista.selectedItem.SubItems(2) = ""
      lv_lista.ListItems.Item(i).Bold = False
      lv_lista.ListItems.Item(i).ForeColor = &H80000012
      lv_lista.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_lista.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_lista.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   lv_lista.Refresh

End Sub

Private Sub cmd_pasar_Click()
   var_si = MsgBox("¿Desea asignar los cliente?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      var_si = MsgBox("Confirmar la asignación de los cliente", vbYesNo, "ATENCION")
      If var_si = 6 Then
         For var_j = 1 To Me.lv_lista.ListItems.Count
             Me.lv_lista.ListItems.Item(var_j).Selected = True
             If Me.lv_lista.selectedItem.SubItems(2) = "*" Then
                rs.Open "DELETE FROM TB_ORACLE_CLIENTES_RUTAS WHERE CLIENTE = '" + Me.lv_lista.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                rs.Open "INSERT INTO TB_ORACLE_CLIENTES_RUTAS (RUTA, CLIENTE) VALUES ('" + var_ruta_cliente + "','" + Me.lv_lista.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
             End If
         Next var_j
         Me.lv_clientes.ListItems.Clear
         rs.Open "select * from tb_oracle_clientes_rutas where ruta = '" + var_ruta_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               strconsulta = "select party_site_number, razon_social from xxvia_jv_tb_clientes where party_site_number = ? order by razon_social"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rs!Cliente)
                    .Parameters.Append parametro
               End With
               Set rsaux = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               Set list_item = lv_clientes.ListItems.Add(, , rs!Cliente)
               list_item.SubItems(1) = IIf(IsNull(rsaux!razon_social), "", rsaux!razon_social)
               list_item.SubItems(2) = IIf(IsNull(rs!prioridad), "", rs!prioridad)
               rs.MoveNext
         Wend
         rs.Close
         Me.frm_lista.Visible = False
      End If
   End If
End Sub

Private Sub cmd_seleccion_Click()
   n = lv_lista.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_lista.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_lista.selectedItem.SubItems(2) = "" And var_rellena = True Then
         lv_lista.selectedItem.SubItems(2) = "*"
         lv_lista.ListItems.Item(i).Bold = True
         lv_lista.ListItems.Item(i).ForeColor = &HFF0000
         lv_lista.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_lista.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_lista.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_lista.selectedItem.SubItems(2) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_lista.selectedItem.SubItems(2) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub cmd_todos_Click()
   n = lv_lista.ListItems.Count
   For i = 1 To n
      lv_lista.ListItems.Item(i).Selected = True
      lv_lista.selectedItem.SubItems(2) = "*"
      lv_lista.ListItems.Item(i).Bold = True
      lv_lista.ListItems.Item(i).ForeColor = &HFF0000
      lv_lista.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_lista.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_lista.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
   Next i
   lv_lista.Refresh
End Sub

Private Sub com_guardar_Click()
   If Trim(Me.txt_clave) <> "" Then
      If Trim(Me.txt_nombre) <> "" Then
         If IsNumeric(Me.txt_prioridad) Then
            rs.Open "UPDATE TB_ORACLE_CLIENTES_RUTAS SET PRIORIDAD = " + Me.txt_prioridad + " WHERE CLIENTE = '" + Me.txt_clave + "'", cnn, adOpenDynamic, adLockOptimistic
            Me.lv_clientes.selectedItem.SubItems(2) = Me.txt_prioridad
         Else
            MsgBox "Número de prioridad incorrecto", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 1500
   frm_lista.Visible = False
   Me.frm_rutas.Visible = False
   Me.Caption = Me.Caption + " " + var_ruta_cliente + " " + var_nombre_ruta_cliente
   rs.Open "select * from tb_oracle_clientes_rutas where ruta = '" + var_ruta_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         strconsulta = "select party_site_number, razon_social from xxvia_jv_tb_clientes where party_site_number = ? order by razon_social"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rs!Cliente)
              .Parameters.Append parametro
         End With
         Set rsaux = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         Set list_item = lv_clientes.ListItems.Add(, , rs!Cliente)
         list_item.SubItems(1) = IIf(IsNull(rsaux!razon_social), "", rsaux!razon_social)
         list_item.SubItems(2) = IIf(IsNull(rs!prioridad), "", rs!prioridad)
         rs.MoveNext
   Wend
   rs.Close
   
End Sub

Private Sub lv_clientes_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Me.txt_clave = Me.lv_clientes.selectedItem
   Me.txt_nombre = Me.lv_clientes.selectedItem.SubItems(1)
   Me.txt_prioridad = Me.lv_clientes.selectedItem.SubItems(2)
End Sub

Private Sub lv_clientes_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      var_si = MsgBox("¿Desea eliminar el registro?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         var_si = MsgBox("Confirmar la eliminación del registro", vbYesNo, "ATENCION")
         If var_si = 6 Then
            rs.Open "DELETE FROM TB_ORACLE_CLIENTES_RUTAS WHERE CLIENTE = '" + Me.lv_clientes.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
            lv_clientes.ListItems.Remove (lv_clientes.selectedItem.Index)
            MsgBox "Se a eliminado el registro", vbOKOnly, "ATENCION"
         End If
      End If
   End If
   If KeyCode = 116 Then
      Me.lv_lista.ListItems.Clear
      rs.Open "select distinct salesrep_id, resource_name  from xxvia_jv_tb_clientes order by resource_name", cnnoracle_4, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_rutas.ListItems.Add(, , rs(0).Value)
            list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
            rs.MoveNext
      Wend
      rs.Close
      Me.frm_rutas.Visible = True
      Me.lv_rutas.SetFocus
   End If
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      i = lv_lista.selectedItem.Index
      If lv_lista.selectedItem.SubItems(2) = "*" Then
         lv_lista.selectedItem.SubItems(2) = ""
         lv_lista.ListItems.Item(i).Bold = False
         lv_lista.ListItems.Item(i).ForeColor = &H80000012
         lv_lista.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_lista.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_lista.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_lista.Refresh
      Else
         lv_lista.selectedItem.SubItems(2) = "*"
         lv_lista.ListItems.Item(i).Bold = True
         lv_lista.ListItems.Item(i).ForeColor = &HFF0000
         lv_lista.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_lista.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_lista.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_lista.Refresh
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_lista.Visible = False
   End If
End Sub

Private Sub lv_rutas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      strconsulta = "select party_site_number, razon_social from xxvia_jv_tb_clientes where salesrep_id = ? order by razon_social"
      With comandoORA
          .ActiveConnection = cnnoracle_4
          .CommandType = adCmdText
          .CommandText = strconsulta
          Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.lv_rutas.selectedItem))
          .Parameters.Append parametro
          
      End With
      Set rs = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs(0).Value)
            list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
            rs.MoveNext
      Wend
      rs.Close
      Me.frm_rutas.Visible = False
      Me.frm_lista.Visible = True
   End If
   If KeyAscii = 27 Then
      Me.frm_rutas.Visible = False
   End If
End Sub

Private Sub txt_clave_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Me.lv_lista.ListItems.Clear
      rs.Open "select distinct salesrep_id, resource_name  from xxvia_jv_tb_clientes order by resource_name", cnnoracle_4, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_rutas.ListItems.Add(, , rs(0).Value)
            list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
            rs.MoveNext
      Wend
      rs.Close
      Me.frm_rutas.Visible = True
      Me.lv_rutas.SetFocus
   End If
End Sub

Private Sub txt_nombre_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Me.lv_lista.ListItems.Clear
      rs.Open "select distinct salesrep_id, resource_name  from xxvia_jv_tb_clientes order by resource_name", cnnoracle_4, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_rutas.ListItems.Add(, , rs(0).Value)
            list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
            rs.MoveNext
      Wend
      rs.Close
      Me.frm_rutas.Visible = True
      Me.lv_rutas.SetFocus
   End If
End Sub

Private Sub txt_prioridad_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Me.lv_lista.ListItems.Clear
      rs.Open "select distinct salesrep_id, resource_name  from xxvia_jv_tb_clientes order by resource_name", cnnoracle_4, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_rutas.ListItems.Add(, , rs(0).Value)
            list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
            rs.MoveNext
      Wend
      rs.Close
      Me.frm_rutas.Visible = True
      Me.lv_rutas.SetFocus
   End If
End Sub

Private Sub txt_prioridad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.com_guardar.SetFocus
   End If
End Sub
