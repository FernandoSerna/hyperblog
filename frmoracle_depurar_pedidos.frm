VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_depurar_pedidos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Depurar ordenes de surtido"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frm_causas_negado 
      Height          =   2835
      Left            =   3750
      TabIndex        =   5
      Top             =   1365
      Width           =   4560
      Begin MSComctlLib.ListView lv_causas_negado 
         Height          =   2430
         Left            =   45
         TabIndex        =   7
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
            Text            =   "Pedido"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Código"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000C0&
         Caption         =   " Causas de negado"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   15
         Width           =   4545
      End
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmoracle_depurar_pedidos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8325
      Picture         =   "frmoracle_depurar_pedidos.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   30
      TabIndex        =   2
      Top             =   330
      Width           =   8670
   End
   Begin VB.Frame Frame1 
      Height          =   4380
      Left            =   45
      TabIndex        =   0
      Top             =   420
      Width           =   8640
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   390
         Picture         =   "frmoracle_depurar_pedidos.frx":073C
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   60
         Picture         =   "frmoracle_depurar_pedidos.frx":0952
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1050
         Picture         =   "frmoracle_depurar_pedidos.frx":0A54
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   720
         Picture         =   "frmoracle_depurar_pedidos.frx":0B26
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Marcar (Enter)"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1380
         Picture         =   "frmoracle_depurar_pedidos.frx":0D70
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   135
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_pedidos 
         Height          =   3855
         Left            =   45
         TabIndex        =   1
         Top             =   435
         Width           =   8550
         _ExtentX        =   15081
         _ExtentY        =   6800
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
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Pedido"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Código"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descripcion"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Cantidad"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Depurado"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "delivery detail"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Source Header ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Inventory Item ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Causa Negado"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Estatus"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "CODIGO NEGADO"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   570
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   30
      Visible         =   0   'False
      Width           =   7665
   End
End
Attribute VB_Name = "frmoracle_depurar_pedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_guardar_Click()
   Dim clnt As New SoapClient30
   Dim var_arreglo() As String
   Dim var_s As String
   Dim var_paso As Boolean
   VAR_ESTATUS = ""
   For var_j = 1 To lv_pedidos.ListItems.Count
       Me.lv_pedidos.ListItems.Item(var_j).Selected = True
       If Trim(Me.lv_pedidos.selectedItem.SubItems(10)) = "C" Then
          VAR_ESTATUS = "C"
       End If
   Next var_j
   var_posible = 0
   For var_j = 1 To lv_pedidos.ListItems.Count
       Me.lv_pedidos.ListItems.Item(var_j).Selected = True
       If Trim(Me.lv_pedidos.selectedItem.SubItems(8)) = "" Then
          var_posible = 1
       End If
   Next var_j
   
   If VAR_ESTATUS = "" Then
      If var_posible = 0 Then
         var_si = MsgBox("¿Desea cerrar el negado?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_si = MsgBox("Confirmar el cerrado del negado", vbYesNo, "ATENCION")
            If var_si = 6 Then
               rsaux10.Open "select * from tb_unidadesorganizacionales where vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux10.EOF Then
                  VAR_USER_ID = rsaux10!user_id
                  VAR_RESP_ID = rsaux10!resp_id
                  VAR_RESP_APPL_ID = rsaux10!resp_appl_id
               End If
               rsaux10.Close
               For var_j = 1 To Me.lv_pedidos.ListItems.Count
                   Me.lv_pedidos.ListItems.Item(var_j).Selected = True
                   rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                   rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                   On Error GoTo SALIR
                   Text1 = "call XXVIA_SP_DEPURA_ORDEN_SURTIDO (" + CStr(CDbl(Me.lv_pedidos.selectedItem.SubItems(6))) + ", " + CStr(CDbl(Me.lv_pedidos.selectedItem.SubItems(5))) + ", '" + Trim(Me.lv_pedidos.selectedItem.SubItems(11)) + "'," + CStr(VAR_USER_ID) + "," + CStr(VAR_RESP_ID) + "," + CStr(VAR_RESP_APPL_ID) + ")"
                   If rsaux8.State = 1 Then
                      rsaux8.Close
                   End If
                   rsaux8.Open "call XXVIA_SP_DEPURA_ORDEN_SURTIDO (" + CStr(CDbl(Me.lv_pedidos.selectedItem.SubItems(6))) + ", " + CStr(CDbl(Me.lv_pedidos.selectedItem.SubItems(5))) + ", '" + Trim(Me.lv_pedidos.selectedItem.SubItems(11)) + "'," + CStr(VAR_USER_ID) + "," + CStr(VAR_RESP_ID) + "," + CStr(VAR_RESP_APPL_ID) + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                   If var_tipo_depurado = 1 Then
                       rsaux.Open "INSERT INTO TB_ORACLE_NEGADO (PEDIDO, INVENTORY_ITEM_ID, CAUSA_NEGADO) VALUES (" + Me.lv_pedidos.selectedItem + "," + Me.lv_pedidos.selectedItem.SubItems(7) + ",'" + Me.lv_pedidos.selectedItem.SubItems(11) + "')", cnn, adOpenDynamic, adLockOptimistic
                   End If
               Next var_j
               Me.lv_pedidos.ListItems.Clear
               MsgBox "Se a terminado el depurado de los pedidos", vbOKOnly, "ATENCION"
               'Unload Me
            End If
         End If
      Else
         MsgBox "Faltan artículos por asignar", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Las causas de negado ya habian sido asignadas con anterioridad", vbOKOnly, "ATENCION"
   End If
   Exit Sub
SALIR:
   If Err.Number = -2147217900 Then
      MsgBox Err.Description
      Resume
   Else
      MsgBox Err.Description
      
      If rs.State = 1 Then
         rs.Close
      End If
   End If
End Sub

Private Sub cmd_invertir_Click()
   n = lv_pedidos.ListItems.Count
   For i = 1 To n
      lv_pedidos.ListItems.Item(i).Selected = True
      If lv_pedidos.selectedItem.SubItems(9) = "*" Then
         lv_pedidos.selectedItem.SubItems(9) = ""
         lv_pedidos.ListItems.Item(i).Bold = False
         lv_pedidos.ListItems.Item(i).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(4).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(5).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(6).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(7).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(8).Bold = False
         lv_pedidos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(7).ForeColor = &H80000012
         lv_pedidos.ListItems.Item(i).ListSubItems(8).ForeColor = &H80000012
      Else
         lv_pedidos.selectedItem.SubItems(9) = "*"
         lv_pedidos.ListItems.Item(i).Bold = True
         lv_pedidos.ListItems.Item(i).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(5).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(6).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(7).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(8).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(7).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(8).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub cmd_marcar_Click()
   i = lv_pedidos.selectedItem.Index
   If lv_pedidos.selectedItem.SubItems(9) = "*" Then
      lv_pedidos.selectedItem.SubItems(9) = ""
      lv_pedidos.ListItems.Item(i).Bold = False
      lv_pedidos.ListItems.Item(i).ForeColor = &H80000012
      lv_pedidos.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_pedidos.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_pedidos.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_pedidos.ListItems.Item(i).ListSubItems(4).Bold = False
      lv_pedidos.ListItems.Item(i).ListSubItems(5).Bold = False
      lv_pedidos.ListItems.Item(i).ListSubItems(6).Bold = False
      lv_pedidos.ListItems.Item(i).ListSubItems(7).Bold = False
      lv_pedidos.ListItems.Item(i).ListSubItems(8).Bold = False
      lv_pedidos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_pedidos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_pedidos.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      lv_pedidos.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
      lv_pedidos.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
      lv_pedidos.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
      lv_pedidos.ListItems.Item(i).ListSubItems(7).ForeColor = &H80000012
      lv_pedidos.ListItems.Item(i).ListSubItems(8).ForeColor = &H80000012
      lv_pedidos.Refresh
   Else
      lv_pedidos.selectedItem.SubItems(9) = "*"
      lv_pedidos.ListItems.Item(i).Bold = True
      lv_pedidos.ListItems.Item(i).ForeColor = &HFF0000
      lv_pedidos.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_pedidos.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_pedidos.ListItems.Item(i).ListSubItems(3).Bold = True
      lv_pedidos.ListItems.Item(i).ListSubItems(4).Bold = True
      lv_pedidos.ListItems.Item(i).ListSubItems(5).Bold = True
      lv_pedidos.ListItems.Item(i).ListSubItems(6).Bold = True
      lv_pedidos.ListItems.Item(i).ListSubItems(7).Bold = True
      lv_pedidos.ListItems.Item(i).ListSubItems(8).Bold = True
      lv_pedidos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_pedidos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_pedidos.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
      lv_pedidos.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
      lv_pedidos.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
      lv_pedidos.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
      lv_pedidos.ListItems.Item(i).ListSubItems(7).ForeColor = &HFF0000
      lv_pedidos.ListItems.Item(i).ListSubItems(8).ForeColor = &HFF0000
      lv_pedidos.Refresh
   End If

End Sub

Private Sub cmd_ninguno_Click()
   n = lv_pedidos.ListItems.Count
   For i = 1 To n
      lv_pedidos.ListItems.Item(i).Selected = True
      lv_pedidos.selectedItem.SubItems(9) = ""
      lv_pedidos.ListItems.Item(i).Bold = False
      lv_pedidos.ListItems.Item(i).ForeColor = &H80000012
      lv_pedidos.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_pedidos.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_pedidos.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_pedidos.ListItems.Item(i).ListSubItems(4).Bold = False
      lv_pedidos.ListItems.Item(i).ListSubItems(5).Bold = False
      lv_pedidos.ListItems.Item(i).ListSubItems(6).Bold = False
      lv_pedidos.ListItems.Item(i).ListSubItems(7).Bold = False
      lv_pedidos.ListItems.Item(i).ListSubItems(8).Bold = False
      lv_pedidos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_pedidos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_pedidos.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      lv_pedidos.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
      lv_pedidos.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
      lv_pedidos.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
      lv_pedidos.ListItems.Item(i).ListSubItems(7).ForeColor = &H80000012
      lv_pedidos.ListItems.Item(i).ListSubItems(8).ForeColor = &H80000012
   Next i
   lv_pedidos.Refresh
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub lv_almacenes_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub cmd_seleccion_Click()
   n = lv_pedidos.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_pedidos.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_pedidos.selectedItem.SubItems(9) = "" And var_rellena = True Then
         lv_pedidos.selectedItem.SubItems(9) = "*"
         lv_pedidos.ListItems.Item(i).Bold = True
         lv_pedidos.ListItems.Item(i).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(5).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(6).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(7).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(8).Bold = True
         lv_pedidos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(7).ForeColor = &HFF0000
         lv_pedidos.ListItems.Item(i).ListSubItems(8).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_pedidos.selectedItem.SubItems(9) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_pedidos.selectedItem.SubItems(9) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i

End Sub

Private Sub cmd_todos_Click()
   n = lv_pedidos.ListItems.Count
   For i = 1 To n
      lv_pedidos.ListItems.Item(i).Selected = True
      lv_pedidos.selectedItem.SubItems(9) = "*"
      lv_pedidos.ListItems.Item(i).Bold = True
      lv_pedidos.ListItems.Item(i).ForeColor = &HFF0000
      lv_pedidos.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_pedidos.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_pedidos.ListItems.Item(i).ListSubItems(3).Bold = True
      lv_pedidos.ListItems.Item(i).ListSubItems(4).Bold = True
      lv_pedidos.ListItems.Item(i).ListSubItems(5).Bold = True
      lv_pedidos.ListItems.Item(i).ListSubItems(6).Bold = True
      lv_pedidos.ListItems.Item(i).ListSubItems(7).Bold = True
      lv_pedidos.ListItems.Item(i).ListSubItems(8).Bold = True
      lv_pedidos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_pedidos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_pedidos.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
      lv_pedidos.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
      lv_pedidos.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
      lv_pedidos.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
      lv_pedidos.ListItems.Item(i).ListSubItems(7).ForeColor = &HFF0000
      lv_pedidos.ListItems.Item(i).ListSubItems(8).ForeColor = &HFF0000
   Next i
   lv_pedidos.Refresh
End Sub

Private Sub Form_Load()
   Me.frm_causas_negado.Visible = False
   rsaux.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_cadena_pedidos_global = "662753,662754,662755,662756,662758,662759,662760,662761,662763,662764,662967,669092,669097,669138,669331,669374,669757,669758,669806,670158,670168,670195,670415,672671,673003,675397,675762,680629,681737,681744,682339,682399,691039,691071"
   var_cadena_pedidos_global = "695186,693365,696094,696165,696217,696403,696611,696636,697024"
   var_cadena = "SELECT a.source_line_id, OHA.HEADER_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,A.item_description,A.source_line_number,A.requested_quantity,A.released_status, c.segment1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID   AND A.SOURCE_HEADER_NUMBER in (" + var_cadena_pedidos_global + ")"
   var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND released_status IN ('B','S') order by A.source_header_number"
   'Text1 = var_cadena
   rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rsaux.EOF
         Set list_item = Me.lv_pedidos.ListItems.Add(, , rsaux!source_header_number)
         list_item.SubItems(1) = IIf(IsNull(rsaux!SEGMENT1), "", rsaux!SEGMENT1)
         list_item.SubItems(2) = IIf(IsNull(rsaux!item_description), "", rsaux!item_description)
         list_item.SubItems(3) = Format(IIf(IsNull(rsaux!requested_quantity), 0, rsaux!requested_quantity), "###,###,##0.00")
         list_item.SubItems(4) = 0
         'list_item.SubItems(5) = rsaux!delivery_detail_id
         list_item.SubItems(5) = rsaux!source_LINE_ID
         list_item.SubItems(6) = rsaux!header_id
         list_item.SubItems(7) = rsaux!inventory_item_id
         list_item.SubItems(8) = ""
         list_item.SubItems(9) = ""
         rsaux.MoveNext
   Wend
   rsaux.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   var_tipo_depurado = 0
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub Frame3_DragDrop(Source As Control, x As Single, Y As Single)

End Sub


Private Sub lv_causas_negado_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      If Me.lv_pedidos.ListItems.Count > 0 Then
         Me.lv_pedidos.SetFocus
      Else
         Me.frm_causas_negado.Visible = False
      End If
   End If
   If KeyAscii = 13 Then
      var_si = MsgBox("¿Desea actualizar el registos?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         For var_j = 1 To Me.lv_pedidos.ListItems.Count
             Me.lv_pedidos.ListItems.Item(var_j).Selected = True
             Me.lv_pedidos.selectedItem.SubItems(8) = Me.lv_causas_negado.selectedItem.SubItems(1)
             Me.lv_pedidos.selectedItem.SubItems(11) = Me.lv_causas_negado.selectedItem
             lv_pedidos.selectedItem.SubItems(9) = ""
             lv_pedidos.ListItems.Item(var_j).Bold = False
             lv_pedidos.ListItems.Item(var_j).ForeColor = &H80000012
             lv_pedidos.ListItems.Item(var_j).ListSubItems(1).Bold = False
             lv_pedidos.ListItems.Item(var_j).ListSubItems(2).Bold = False
             lv_pedidos.ListItems.Item(var_j).ListSubItems(3).Bold = False
             lv_pedidos.ListItems.Item(var_j).ListSubItems(4).Bold = False
             lv_pedidos.ListItems.Item(var_j).ListSubItems(5).Bold = False
             lv_pedidos.ListItems.Item(var_j).ListSubItems(6).Bold = False
             lv_pedidos.ListItems.Item(var_j).ListSubItems(7).Bold = False
             lv_pedidos.ListItems.Item(var_j).ListSubItems(8).Bold = False
             lv_pedidos.ListItems.Item(var_j).ListSubItems(1).ForeColor = &H80000012
             lv_pedidos.ListItems.Item(var_j).ListSubItems(2).ForeColor = &H80000012
             lv_pedidos.ListItems.Item(var_j).ListSubItems(3).ForeColor = &H80000012
             lv_pedidos.ListItems.Item(var_j).ListSubItems(4).ForeColor = &H80000012
             lv_pedidos.ListItems.Item(var_j).ListSubItems(5).ForeColor = &H80000012
             lv_pedidos.ListItems.Item(var_j).ListSubItems(6).ForeColor = &H80000012
             lv_pedidos.ListItems.Item(var_j).ListSubItems(7).ForeColor = &H80000012
             lv_pedidos.ListItems.Item(var_j).ListSubItems(8).ForeColor = &H80000012
             lv_pedidos.Refresh
         Next var_j
         
         Me.frm_causas_negado.Visible = False
      End If
   End If
   
End Sub

Private Sub lv_causas_negado_LostFocus()
   Me.frm_causas_negado.Visible = False
End Sub

Private Sub lv_pedidos_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 115 Then
      Me.frm_causas_negado.Visible = True
      Me.lv_causas_negado.ListItems.Clear
      rsaux.Open "select lookup_code as causa_negado, meaning as descripcion from fnd_lookup_values where lookup_type = 'CANCEL_CODE' and language = 'US' ORDER BY LANGUAGE", cnnoracle_4, adOpenDynamic, adLockOptimistic
      While Not rsaux.EOF
         Set list_item = Me.lv_causas_negado.ListItems.Add(, , rsaux!CAUSA_NEGADO)
         list_item.SubItems(1) = IIf(IsNull(rsaux!Descripcion), "", rsaux!Descripcion)
         rsaux.MoveNext
      Wend
      rsaux.Close
      Me.lv_causas_negado.SetFocus
   End If
End Sub

Private Sub lv_pedidos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_pedidos.ListItems.Count > 0 Then
         i = lv_pedidos.selectedItem.Index
         If lv_pedidos.selectedItem.SubItems(9) = "*" Then
            lv_pedidos.selectedItem.SubItems(9) = ""
            lv_pedidos.ListItems.Item(i).Bold = False
            lv_pedidos.ListItems.Item(i).ForeColor = &H80000012
            lv_pedidos.ListItems.Item(i).ListSubItems(1).Bold = False
            lv_pedidos.ListItems.Item(i).ListSubItems(2).Bold = False
            lv_pedidos.ListItems.Item(i).ListSubItems(3).Bold = False
            lv_pedidos.ListItems.Item(i).ListSubItems(4).Bold = False
            lv_pedidos.ListItems.Item(i).ListSubItems(5).Bold = False
            lv_pedidos.ListItems.Item(i).ListSubItems(6).Bold = False
            lv_pedidos.ListItems.Item(i).ListSubItems(7).Bold = False
            lv_pedidos.ListItems.Item(i).ListSubItems(8).Bold = False
            lv_pedidos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
            lv_pedidos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
            lv_pedidos.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
            lv_pedidos.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
            lv_pedidos.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
            lv_pedidos.ListItems.Item(i).ListSubItems(6).ForeColor = &H80000012
            lv_pedidos.ListItems.Item(i).ListSubItems(7).ForeColor = &H80000012
            lv_pedidos.ListItems.Item(i).ListSubItems(8).ForeColor = &H80000012
            lv_pedidos.Refresh
         Else
            lv_pedidos.selectedItem.SubItems(9) = "*"
            lv_pedidos.ListItems.Item(i).Bold = True
            lv_pedidos.ListItems.Item(i).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(1).Bold = True
            lv_pedidos.ListItems.Item(i).ListSubItems(2).Bold = True
            lv_pedidos.ListItems.Item(i).ListSubItems(3).Bold = True
            lv_pedidos.ListItems.Item(i).ListSubItems(4).Bold = True
            lv_pedidos.ListItems.Item(i).ListSubItems(5).Bold = True
            lv_pedidos.ListItems.Item(i).ListSubItems(6).Bold = True
            lv_pedidos.ListItems.Item(i).ListSubItems(7).Bold = True
            lv_pedidos.ListItems.Item(i).ListSubItems(8).Bold = True
            lv_pedidos.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(6).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(7).ForeColor = &HFF0000
            lv_pedidos.ListItems.Item(i).ListSubItems(8).ForeColor = &HFF0000
            lv_pedidos.Refresh
         End If
      End If
   End If
End Sub

