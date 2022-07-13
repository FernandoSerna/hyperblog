VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_asignar_maquinas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignar máquinas"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frm_lista_2 
      Height          =   2310
      Left            =   105
      TabIndex        =   15
      Top             =   2880
      Width           =   3750
      Begin MSComctlLib.ListView lv_lista_2 
         Height          =   1755
         Left            =   45
         TabIndex        =   16
         Top             =   480
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   3096
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
            Text            =   "Máquinas"
            Object.Width           =   5733
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label3 
         BackColor       =   &H000000FF&
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   30
         TabIndex        =   17
         Top             =   135
         Width           =   3675
      End
   End
   Begin VB.Frame frm_lista 
      Height          =   4185
      Left            =   90
      TabIndex        =   11
      Top             =   825
      Width           =   3750
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   360
         Picture         =   "frmoracle_asignar_maquinas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   435
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   30
         Picture         =   "frmoracle_asignar_maquinas.frx":0216
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   435
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1020
         Picture         =   "frmoracle_asignar_maquinas.frx":0318
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   435
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   690
         Picture         =   "frmoracle_asignar_maquinas.frx":03EA
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Marcar (Enter)"
         Top             =   435
         Width           =   330
      End
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1350
         Picture         =   "frmoracle_asignar_maquinas.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   435
         Width           =   330
      End
      Begin VB.CommandButton cmd_asignar 
         Caption         =   "Asignar máquinas"
         Height          =   360
         Left            =   60
         TabIndex        =   14
         Top             =   3765
         Width           =   3645
      End
      Begin MSComctlLib.ListView lv_lista 
         Height          =   2970
         Left            =   45
         TabIndex        =   12
         Top             =   765
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   5239
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
            Text            =   "Máquinas"
            Object.Width           =   5733
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H000000FF&
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   30
         TabIndex        =   13
         Top             =   135
         Width           =   3675
      End
   End
   Begin VB.Frame Frame5 
      Height          =   690
      Left            =   90
      TabIndex        =   9
      Top             =   4470
      Width           =   3765
      Begin VB.TextBox txt_maquina_salida 
         Height          =   390
         Left            =   870
         TabIndex        =   2
         Top             =   165
         Width           =   2820
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Anden:"
         Height          =   195
         Left            =   75
         TabIndex        =   10
         Top             =   270
         Width           =   510
      End
   End
   Begin VB.Frame Frame3 
      Height          =   810
      Left            =   90
      TabIndex        =   4
      Top             =   30
      Width           =   3765
      Begin VB.Label lbl_embarque 
         Alignment       =   2  'Center
         Caption         =   "Embarque 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   60
         TabIndex        =   5
         Top             =   210
         Width           =   3570
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   90
      TabIndex        =   3
      Top             =   840
      Width           =   3765
      Begin VB.CommandButton cmd_guardar 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3300
         Picture         =   "frmoracle_asignar_maquinas.frx":084A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Guardar Alt + G"
         Top             =   210
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txt_maquina_entrada 
         Height          =   390
         Left            =   915
         TabIndex        =   0
         Top             =   180
         Width           =   2775
      End
      Begin VB.Frame Frame4 
         Height          =   120
         Left            =   15
         TabIndex        =   6
         Top             =   645
         Width           =   3705
      End
      Begin MSComctlLib.ListView lv_maquinas 
         Height          =   2715
         Left            =   45
         TabIndex        =   7
         Top             =   840
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   4789
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Maquina"
            Object.Width           =   6209
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Estación:"
         Height          =   195
         Left            =   90
         TabIndex        =   8
         Top             =   285
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmoracle_asignar_maquinas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_ventana As Integer

Private Sub cmd_asignar_Click()
   var_si = MsgBox("¿Desea asignar la(s) máquina(s) al embarque?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      For var_j = 1 To Me.lv_lista.ListItems.Count
          Me.lv_lista.ListItems.Item(var_j).Selected = True
          If Me.lv_lista.selectedItem.SubItems(1) = "*" Then
             Me.txt_maquina_entrada = Me.lv_lista.selectedItem
             rs.Open "select * from tb_oracle_maquinas_asignadas where embarque =" + CStr(var_embarque_asignar) + " and maquina = '" + Me.txt_maquina_entrada + "' AND USO = 'E'", cnn, adOpenDynamic, adLockOptimistic
             If rs.EOF Then
                rsaux.Open "INSERT INTO tb_oracle_maquinas_asignadas (EMBARQUE, MAQUINA, USO) VALUES (" + CStr(var_embarque_asignar) + ",'" + Me.txt_maquina_entrada + "','E')", cnn, adOpenDynamic, adLockOptimistic
                Set list_item = Me.lv_maquinas.ListItems.Add(, , Me.txt_maquina_entrada)
             End If
             rs.Close
          End If
      Next var_j
      Me.frm_lista.Visible = False
   End If
End Sub

Private Sub cmd_guardar_Click()
   If Trim(Me.txt_maquina_entrada) <> "" Then
      var_si = MsgBox("¿Desea asignar la máquina al embarque?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         rs.Open "select * from tb_oracle_maquinas_asignadas where embarque =" + CStr(var_embarque_asignar) + " and maquina = '" + Me.txt_maquina_entrada + "' AND USO = 'E'", cnn, adOpenDynamic, adLockOptimistic
         If rs.EOF Then
            rsaux.Open "INSERT INTO tb_oracle_maquinas_asignadas (EMBARQUE, MAQUINA, USO) VALUES (" + CStr(var_embarque_asignar) + ",'" + Me.txt_maquina_entrada + "','E')", cnn, adOpenDynamic, adLockOptimistic
            Set list_item = Me.lv_maquinas.ListItems.Add(, , Me.txt_maquina_entrada)
         End If
         rs.Close
      End If
   Else
      MsgBox "No se a indicado una máquina", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_invertir_Click()
   n = lv_lista.ListItems.Count
   For i = 1 To n
      lv_lista.ListItems.Item(i).Selected = True
      If lv_lista.selectedItem.SubItems(1) = "*" Then
         lv_lista.selectedItem.SubItems(1) = ""
         lv_lista.ListItems.Item(i).Bold = False
         lv_lista.ListItems.Item(i).ForeColor = &H80000012
         lv_lista.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      Else
         lv_lista.selectedItem.SubItems(1) = "*"
         lv_lista.ListItems.Item(i).Bold = True
         lv_lista.ListItems.Item(i).ForeColor = &HFF0000
         lv_lista.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub cmd_marcar_Click()
   i = lv_lista.selectedItem.Index
   If lv_lista.selectedItem.SubItems(1) = "*" Then
      lv_lista.selectedItem.SubItems(1) = ""
      lv_lista.ListItems.Item(i).Bold = False
      lv_lista.ListItems.Item(i).ForeColor = &H80000012
      lv_lista.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_lista.Refresh
   Else
      lv_lista.selectedItem.SubItems(1) = "*"
      lv_lista.ListItems.Item(i).Bold = True
      lv_lista.ListItems.Item(i).ForeColor = &HFF0000
      lv_lista.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_lista.Refresh
   End If
End Sub

Private Sub cmd_ninguno_Click()
   n = lv_lista.ListItems.Count
   For i = 1 To n
      lv_lista.ListItems.Item(i).Selected = True
      lv_lista.selectedItem.SubItems(1) = ""
      lv_lista.ListItems.Item(i).Bold = False
      lv_lista.ListItems.Item(i).ForeColor = &H80000012
      lv_lista.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
   Next i
   lv_lista.Refresh
End Sub

Private Sub cmd_seleccion_Click()
   n = lv_lista.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_lista.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_lista.selectedItem.SubItems(1) = "" And var_rellena = True Then
         lv_lista.selectedItem.SubItems(1) = "*"
         lv_lista.ListItems.Item(i).Bold = True
         lv_lista.ListItems.Item(i).ForeColor = &HFF0000
         lv_lista.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_lista.selectedItem.SubItems(1) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_lista.selectedItem.SubItems(1) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub cmd_todos_Click()
   n = lv_lista.ListItems.Count
   For i = 1 To n
      lv_lista.ListItems.Item(i).Selected = True
      lv_lista.selectedItem.SubItems(1) = "*"
      lv_lista.ListItems.Item(i).Bold = True
      lv_lista.ListItems.Item(i).ForeColor = &HFF0000
      lv_lista.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
   Next i
   lv_lista.Refresh
End Sub

Private Sub Form_Load()
   Me.frm_lista.Visible = False
   Me.frm_lista_2.Visible = False
   Me.lbl_embarque = "Embarque: " + CStr(var_embarque_asignar)
   'MsgBox cnn.ConnectionString
   rs.Open "select * from tb_oracle_maquinas_asignadas where embarque = " + CStr(var_embarque_asignar) + " and uso = 'E'", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set list_item = lv_maquinas.ListItems.Add(, , rs!maquina)
         rs.MoveNext
   Wend
   rs.Close
   rs.Open "select * from tb_oracle_maquinas_asignadas where embarque = " + CStr(var_embarque_asignar) + " and uso = 'S'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      Me.txt_maquina_salida = rs!maquina
   End If
   rs.Close
End Sub

Private Sub Frame2_DragDrop(Source As Control, x As Single, Y As Single)

End Sub

Private Sub lv_lista_2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_lista_2, ColumnHeader)
End Sub

Private Sub lv_lista_2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_ventana = 2 Then
         var_si = MsgBox("¿Desea asignar la máquina de salida?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            rs.Open "DELETE FROM TB_ORACLE_MAQUINAS_ASIGNADAS WHERE EMBARQUE = " + CStr(var_embarque_asignar) + " AND USO = 'S'", cnn, adOpenDynamic, adLockOptimistic
            rs.Open "INSERT INTO TB_ORACLE_MAQUINAS_ASIGNADAS (EMBARQUE, MAQUINA, USO) VALUES (" + CStr(var_embarque_asignar) + ",'" + Me.lv_lista_2.selectedItem + "','S')"
            Me.txt_maquina_salida = Me.lv_lista_2.selectedItem
            Me.txt_maquina_salida.SetFocus
         End If
      End If
   End If
End Sub

Private Sub lv_lista_2_LostFocus()
   Me.frm_lista_2.Visible = False
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_ventana = 1 Then
         If lv_lista.ListItems.Count > 0 Then
            i = lv_lista.selectedItem.Index
            If lv_lista.selectedItem.SubItems(1) = "*" Then
               lv_lista.selectedItem.SubItems(1) = ""
               lv_lista.ListItems.Item(i).Bold = False
               lv_lista.ListItems.Item(i).ForeColor = &H80000012
               lv_lista.ListItems.Item(i).ListSubItems(1).Bold = False
               lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
               lv_lista.Refresh
            Else
               lv_lista.selectedItem.SubItems(1) = "*"
               lv_lista.ListItems.Item(i).Bold = True
               lv_lista.ListItems.Item(i).ForeColor = &HFF0000
               lv_lista.ListItems.Item(i).ListSubItems(1).Bold = True
               lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
               lv_lista.Refresh
            End If
         End If
         'Me.txt_maquina_entrada = Me.lv_lista.selectedItem
         'Me.txt_maquina_entrada.SetFocus
      End If
      If var_ventana = 2 Then
         var_si = MsgBox("¿Desea asignar la máquina de salida?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            rs.Open "DELETE FROM TB_ORACLE_MAQUINAS_ASIGNADAS WHERE EMBARQUE = " + CStr(var_embarque_asignar) + " AND USO = 'S'", cnn, adOpenDynamic, adLockOptimistic
            rs.Open "INSERT INTO TB_ORACLE_MAQUINAS_ASIGNADAS (EMBARQUE, MAQUINA, USO) VALUES (" + CStr(var_embarque_asignar) + ",'" + Me.lv_lista.selectedItem + "','S')"
            Me.txt_maquina_salida = Me.lv_lista.selectedItem
            Me.txt_maquina_salida.SetFocus
         End If
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_lista.Visible = False
   End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_ventana = 1
      lv_lista.ListItems.Clear
      rs.Open "SELECT * FROM TB_ORACLE_MAQUINAS WHERE USO = 'S'", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!maquina)
            rs.MoveNext
      Wend
      rs.Close
      Me.frm_lista.Visible = True
      Me.lv_lista.SetFocus
   End If
End Sub

Private Sub lv_maquinas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_ventana = 1
      lv_lista.ListItems.Clear
      rs.Open "SELECT * FROM TB_ORACLE_MAQUINAS WHERE USO = 'E'", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!maquina)
            rs.MoveNext
      Wend
      rs.Close
      Me.frm_lista.Visible = True
      Me.lv_lista.SetFocus
      For var_j = 1 To lv_lista.ListItems.Count
          Me.lv_lista.ListItems.Item(var_j).Selected = True
          For var_i = 1 To Me.lv_maquinas.ListItems.Count
              Me.lv_maquinas.ListItems.Item(var_i).Selected = True
              If Me.lv_lista.selectedItem = Me.lv_maquinas.selectedItem Then
                 lv_lista.selectedItem.SubItems(1) = "*"
                 lv_lista.ListItems.Item(var_j).Bold = True
                 lv_lista.ListItems.Item(var_j).ForeColor = &HFF0000
                 lv_lista.ListItems.Item(var_j).ListSubItems(1).Bold = True
                 lv_lista.ListItems.Item(var_j).ListSubItems(1).ForeColor = &HFF0000
              End If
          Next var_i
      Next var_j
      If Me.lv_lista.ListItems.Count > 0 Then
         Me.lv_lista.ListItems.Item(1).Selected = True
         Me.lv_lista.SetFocus
      End If
   End If

   If KeyCode = 114 Then
      If Me.lv_maquinas.ListItems.Count > 0 Then
         var_si = MsgBox("¿Deseas eliminar la máquina del embarque?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            rs.Open "DELETE FROM TB_ORACLE_MAQUINAS_ASIGNADAS WHERE EMBARQUE = " + CStr(var_embarque_asignar) + " AND MAQUINA = '" + Me.lv_maquinas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
            lv_maquinas.ListItems.Remove (lv_maquinas.selectedItem.Index)
         End If
      End If
   End If
End Sub

Private Sub lv_maquinas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_maquina_salida.SetFocus
   End If
End Sub

Private Sub txt_maquina_entrada_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_ventana = 1
      lv_lista.ListItems.Clear
      rs.Open "SELECT * FROM TB_ORACLE_MAQUINAS WHERE USO = 'E'", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!maquina)
            rs.MoveNext
      Wend
      rs.Close
      Me.frm_lista.Visible = True
      Me.lv_lista.SetFocus
      For var_j = 1 To lv_lista.ListItems.Count
          Me.lv_lista.ListItems.Item(var_j).Selected = True
          For var_i = 1 To Me.lv_maquinas.ListItems.Count
              Me.lv_maquinas.ListItems.Item(var_i).Selected = True
              If Me.lv_lista.selectedItem = Me.lv_maquinas.selectedItem Then
                 lv_lista.selectedItem.SubItems(1) = "*"
                 lv_lista.ListItems.Item(var_j).Bold = True
                 lv_lista.ListItems.Item(var_j).ForeColor = &HFF0000
                 lv_lista.ListItems.Item(var_j).ListSubItems(1).Bold = True
                 lv_lista.ListItems.Item(var_j).ListSubItems(1).ForeColor = &HFF0000
              End If
          Next var_i
      Next var_j
      If Me.lv_lista.ListItems.Count > 0 Then
         Me.lv_lista.ListItems.Item(1).Selected = True
         Me.lv_lista.SetFocus
      End If
   End If
   
End Sub

Private Sub txt_maquina_entrada_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      'Me.cmd_guardar.SetFocus
   End If
End Sub

Private Sub txt_maquina_entrada_LostFocus()
   If Trim(Me.txt_maquina_entrada) <> "" Then
      rs.Open "select * from tb_oracle_maquinas where maquina = '" + Me.txt_maquina_entrada + "' and uso = 'E'", cnn, adOpenDynamic, adLockOptimistic
      If rs.EOF Then
         MsgBox "La máquina no existe", vbOKOnly, "ATENCION"
         Me.txt_maquina_entrada = ""
      End If
      rs.Close
   End If
End Sub

Private Sub txt_maquina_salida_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_ventana = 2
      lv_lista_2.ListItems.Clear
      rs.Open "SELECT * FROM TB_ORACLE_MAQUINAS WHERE USO = 'S'", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista_2.ListItems.Add(, , rs!maquina)
            rs.MoveNext
      Wend
      rs.Close
      Me.frm_lista_2.Visible = True
      Me.lv_lista_2.SetFocus
   End If
End Sub
