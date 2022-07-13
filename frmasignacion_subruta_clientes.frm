VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmasignacion_subruta_clientes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignación de subrutas a clientes"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11160
      Picture         =   "frmasignacion_subruta_clientes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frmasignacion_subruta_clientes.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   45
      TabIndex        =   8
      Top             =   300
      Width           =   11520
   End
   Begin VB.Frame Frame3 
      Caption         =   " Subruta "
      Height          =   720
      Left            =   5850
      TabIndex        =   2
      Top             =   495
      Width           =   5640
      Begin VB.TextBox txt_nombre_subruta 
         Height          =   390
         Left            =   915
         TabIndex        =   7
         Top             =   240
         Width           =   4530
      End
      Begin VB.TextBox txt_subruta 
         Height          =   390
         Left            =   105
         TabIndex        =   6
         Top             =   240
         Width           =   780
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Clientes "
      Height          =   5940
      Left            =   120
      TabIndex        =   1
      Top             =   1245
      Width           =   11400
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   435
         Picture         =   "frmasignacion_subruta_clientes.frx":0784
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton Command9 
         Height          =   315
         Left            =   105
         Picture         =   "frmasignacion_subruta_clientes.frx":099A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton Command8 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1095
         Picture         =   "frmasignacion_subruta_clientes.frx":0A9C
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton Command7 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   765
         Picture         =   "frmasignacion_subruta_clientes.frx":0B6E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Marcar (Enter)"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1425
         Picture         =   "frmasignacion_subruta_clientes.frx":0DB8
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   225
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_clientes 
         Height          =   5325
         Left            =   90
         TabIndex        =   5
         Top             =   540
         Width           =   11205
         _ExtentX        =   19764
         _ExtentY        =   9393
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
            Text            =   "Clave"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Subruta"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nombre Subruta"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Ruta  "
      Height          =   720
      Left            =   135
      TabIndex        =   0
      Top             =   495
      Width           =   5640
      Begin VB.TextBox txt_nombre_ruta 
         Height          =   390
         Left            =   930
         TabIndex        =   4
         Top             =   240
         Width           =   4530
      End
      Begin VB.TextBox txt_ruta 
         Height          =   390
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmasignacion_subruta_clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_Click()
   If Me.txt_subruta <> "" Then
      If Me.lv_clientes.ListItems.Count > 0 Then
         var_posible = 0
         For var_i = 1 To lv_clientes.ListItems.Count
             lv_clientes.ListItems.Item(var_i).Selected = True
             If Me.lv_clientes.selectedItem.SubItems(4) = "*" Then
                var_posible = 1
             End If
         Next var_i
         If var_posible = 1 Then
            var_si = MsgBox("¿Desea asignarle la subruta a los clientes seleccionados?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               var_si = MsgBox("Confirmar la seleccion de la subruta a los clientes seleccionados", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  For var_i = 1 To lv_clientes.ListItems.Count
                      lv_clientes.ListItems.Item(var_i).Selected = True
                      If lv_clientes.selectedItem.SubItems(4) = "*" Then
                         rs.Open "update tb_clientes set vcha_sru_subruta_id = '" + Me.txt_subruta + "' where vcha_cli_clave_id = '" + Me.lv_clientes.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
                      End If
                  Next var_i
                  Dim list_item As ListItem
                  lv_clientes.ListItems.Clear
                  rsaux.Open "select * from vw_clientes where vcha_rut_ruta_id = '" + Me.txt_ruta + "'", cnn, adOpenDynamic, adLockOptimistic
                  While Not rsaux.EOF
                        Set list_item = lv_clientes.ListItems.Add(, , rsaux!vcha_Cli_clave_id)
                        list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_Cli_nombre), "", rsaux!vcha_Cli_nombre)
                        list_item.SubItems(2) = IIf(IsNull(rsaux!vcha_sru_subruta_id), "", rsaux!vcha_sru_subruta_id)
                        list_item.SubItems(3) = IIf(IsNull(rsaux!vcha_sru_nombre), "", rsaux!vcha_sru_nombre)
                        list_item.SubItems(4) = ""
                        rsaux.MoveNext
                  Wend
                  rsaux.Close

               End If
            End If
         Else
            MsgBox "No se a seleccionado ningun cliente", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "No existen clientes para asignar", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command10_Click()
   var_todos_lineas = 1
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   n = lv_clientes.ListItems.Count
   For i = 1 To n
       lv_clientes.ListItems.Item(i).SubItems(4) = "*"
       lv_clientes.ListItems.Item(i).Bold = True
       lv_clientes.ListItems.Item(i).ForeColor = &HFF0000
       lv_clientes.ListItems.Item(i).ListSubItems(1).Bold = True
       lv_clientes.ListItems.Item(i).ListSubItems(2).Bold = True
       lv_clientes.ListItems.Item(i).ListSubItems(3).Bold = True
       lv_clientes.ListItems.Item(i).ListSubItems(4).Bold = True
       lv_clientes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
       lv_clientes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
       lv_clientes.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
       lv_clientes.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
   Next
   lv_clientes.Refresh
End Sub

Private Sub Command6_Click()
   If var_todos_lineas = 1 Then
   Else
         var_todos_lineas = 0
   End If
   n = lv_clientes.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_clientes.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_clientes.selectedItem.SubItems(4) = "" And var_rellena = True Then
         lv_clientes.selectedItem.SubItems(4) = "*"
         lv_clientes.ListItems.Item(i).Bold = True
         lv_clientes.ListItems.Item(i).ForeColor = &HFF0000
         lv_clientes.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_clientes.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_clientes.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_clientes.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_clientes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_clientes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_clientes.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_clientes.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_clientes.selectedItem.SubItems(4) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_clientes.selectedItem.SubItems(4) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i

End Sub

Private Sub Command7_Click()
   var_todos_lineas = 0
   i = lv_clientes.selectedItem.Index
   If lv_clientes.selectedItem.SubItems(4) = "*" Then
      lv_clientes.selectedItem.SubItems(4) = ""
      lv_clientes.ListItems.Item(i).Bold = False
      lv_clientes.ListItems.Item(i).ForeColor = &H80000012
      lv_clientes.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_clientes.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_clientes.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_clientes.ListItems.Item(i).ListSubItems(4).Bold = False
      lv_clientes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_clientes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_clientes.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      lv_clientes.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
      lv_clientes.Refresh
   Else
      lv_clientes.selectedItem.SubItems(4) = "*"
      lv_clientes.ListItems.Item(i).Bold = True
      lv_clientes.ListItems.Item(i).ForeColor = &HFF0000
      lv_clientes.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_clientes.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_clientes.ListItems.Item(i).ListSubItems(3).Bold = True
      lv_clientes.ListItems.Item(i).ListSubItems(4).Bold = True
      lv_clientes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_clientes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_clientes.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
      lv_clientes.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
      lv_clientes.Refresh
   End If
End Sub

Private Sub Command8_Click()
   If var_todos_lineas = 1 Then
   Else
        var_todos_lineas = 0
   End If
   n = lv_clientes.ListItems.Count
   For i = 1 To n
      lv_clientes.ListItems.Item(i).Selected = True
      If lv_clientes.selectedItem.SubItems(4) = "*" Then
         lv_clientes.selectedItem.SubItems(4) = ""
         lv_clientes.ListItems.Item(i).Bold = False
         lv_clientes.ListItems.Item(i).ForeColor = &H80000012
         lv_clientes.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_clientes.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_clientes.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_clientes.ListItems.Item(i).ListSubItems(4).Bold = False
         lv_clientes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_clientes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_clientes.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
         lv_clientes.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
      Else
         lv_clientes.selectedItem.SubItems(4) = "*"
         lv_clientes.ListItems.Item(i).Bold = True
         lv_clientes.ListItems.Item(i).ForeColor = &HFF0000
         lv_clientes.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_clientes.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_clientes.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_clientes.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_clientes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_clientes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_clientes.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_clientes.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub Command9_Click()
   var_todos_lineas = 0
   n = lv_clientes.ListItems.Count
   For i = 1 To n
      lv_clientes.ListItems.Item(i).Selected = True
      lv_clientes.selectedItem.SubItems(4) = ""
      lv_clientes.ListItems.Item(i).Bold = False
      lv_clientes.ListItems.Item(i).ForeColor = &H80000012
      lv_clientes.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_clientes.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_clientes.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_clientes.ListItems.Item(i).ListSubItems(4).Bold = False
      lv_clientes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_clientes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_clientes.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      lv_clientes.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
   Next i
   lv_clientes.Refresh
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   activa_forma (var_activa_forma_packing_list)
End Sub

Private Sub lv_clientes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_clientes, ColumnHeader)
End Sub

Private Sub lv_clientes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim numero_lineas As Integer
      Dim numero_seleccionado1 As Integer
      Dim numero_seleccionado2 As Integer
      Dim primera_vez As Boolean
      Dim segunda_vez As Boolean
      Dim i As Integer
      Dim n As Integer
      Dim list_item As ListItem
      n = lv_clientes.ListItems.Count
      i = lv_clientes.selectedItem.Index
      If lv_clientes.ListItems.Item(i).SubItems(4) = "*" Then
         lv_clientes.ListItems.Item(i).SubItems(4) = " "
         lv_clientes.ListItems.Item(i).Bold = False
         lv_clientes.ListItems.Item(i).ForeColor = &H80000012
         lv_clientes.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_clientes.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_clientes.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_clientes.ListItems.Item(i).ListSubItems(4).Bold = False
         lv_clientes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_clientes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_clientes.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
         lv_clientes.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
      Else
         lv_clientes.ListItems.Item(i).SubItems(4) = "*"
         lv_clientes.ListItems.Item(i).Bold = True
         lv_clientes.ListItems.Item(i).ForeColor = &HFF0000
         lv_clientes.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_clientes.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_clientes.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_clientes.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_clientes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_clientes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_clientes.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_clientes.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
      End If
      lv_clientes.Refresh
   End If
End Sub

Private Sub txt_nombre_ruta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_subruta.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_subruta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.lv_clientes.ListItems.Count > 0 Then
         Me.lv_clientes.SetFocus
      End If
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_ruta_Change()
   Me.lv_clientes.ListItems.Clear
End Sub

Private Sub txt_ruta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_nombre_ruta.SetFocus
   End If
End Sub

Private Sub txt_ruta_LostFocus()
   If Me.txt_ruta <> "" Then
      Me.lv_clientes.ListItems.Clear
      rs.Open "select * from tb_rutas where vcha_rut_ruta_id = '" + Me.txt_ruta + "' ", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_ruta = IIf(IsNull(rs!vcha_rut_nombre), "", rs!vcha_rut_nombre)
         Dim list_item As ListItem
         rsaux.Open "select * from vw_clientes where vcha_rut_ruta_id = '" + Me.txt_ruta + "'", cnn, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               Set list_item = lv_clientes.ListItems.Add(, , rsaux!vcha_Cli_clave_id)
               list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_Cli_nombre), "", rsaux!vcha_Cli_nombre)
               list_item.SubItems(2) = IIf(IsNull(rsaux!vcha_sru_subruta_id), "", rsaux!vcha_sru_subruta_id)
               list_item.SubItems(3) = IIf(IsNull(rsaux!vcha_sru_nombre), "", rsaux!vcha_sru_nombre)
               list_item.SubItems(4) = ""
               rsaux.MoveNext
         Wend
         rsaux.Close

      Else
         MsgBox "La ruta no existe", vbOKOnly, "ATENCION"
         Me.txt_nombre_ruta = ""
      End If
      rs.Close
   End If
End Sub

Private Sub txt_subruta_Change()
   Me.txt_nombre_subruta = ""
End Sub

Private Sub txt_subruta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_nombre_subruta.SetFocus
   End If
End Sub

Private Sub txt_subruta_LostFocus()
   If Trim(Me.txt_subruta) <> "" Then
      rs.Open "select * from tb_subrutas where vcha_sru_subruta_id = '" + Me.txt_subruta + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_subruta = IIf(IsNull(rs!vcha_sru_nombre), "", rs!vcha_sru_nombre)
      Else
         MsgBox "La subruta no existe", vbOKOnly, "ATENCION"
         Me.txt_subruta = ""
         Me.txt_nombre_subruta = ""
      End If
      rs.Close
   Else
      Me.txt_nombre_subruta = ""
   End If
End Sub
