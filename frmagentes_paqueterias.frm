VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmagentes_paqueterias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignación de paqueteria a agentes"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   1710
      Left            =   30
      TabIndex        =   12
      Top             =   180
      Width           =   6405
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1155
         Left            =   45
         TabIndex        =   13
         Top             =   465
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   2037
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
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   8819
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   14
         Top             =   135
         Width           =   6330
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6045
      Picture         =   "frmagentes_paqueterias.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_cancelar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frmagentes_paqueterias.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Cancelar Alt + C"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmagentes_paqueterias.frx":0784
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   1515
      Left            =   60
      TabIndex        =   8
      Top             =   390
      Width           =   6300
      Begin VB.TextBox txt_contable 
         Height          =   345
         Left            =   1035
         TabIndex        =   4
         Top             =   990
         Width           =   2550
      End
      Begin VB.TextBox txt_nombre_paqueteria 
         Height          =   345
         Left            =   1890
         TabIndex        =   3
         Top             =   630
         Width           =   4305
      End
      Begin VB.TextBox txt_paqueteria 
         Height          =   345
         Left            =   1035
         TabIndex        =   2
         Top             =   630
         Width           =   825
      End
      Begin VB.TextBox txt_nombre_agente 
         Height          =   345
         Left            =   1890
         TabIndex        =   1
         Top             =   270
         Width           =   4305
      End
      Begin VB.TextBox txt_agente 
         Height          =   345
         Left            =   1035
         TabIndex        =   0
         Top             =   270
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Contable:"
         Height          =   195
         Left            =   150
         TabIndex        =   15
         Top             =   1065
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Paqueteria:"
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   705
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   345
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   0
      TabIndex        =   9
      Top             =   255
      Width           =   6345
   End
End
Attribute VB_Name = "frmagentes_paqueterias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VAR_TIPO_LISTA As Integer
Private Sub cmd_aceptar_pedidos_Click()
   If Me.txt_agente <> "" Then
      If Me.txt_paqueteria <> "" Then
         var_si = MsgBox("Desea asignarle la paqueteria al agente", vbYesNo, "ATENCION")
         If var_si = 6 Then
            rs.Open "update tb_agentes set VCHA_PAQ_PAQUETERIA_ID = '" + Me.txt_paqueteria + "', vcha_age_contable = '" + Me.txt_contable + "' WHERE VCHA_AGE_AGENTE_ID = '" + Me.txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
            MsgBox "Se a actualizado el registro", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Debe indicar una paqueteria", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Deve de seleccionar un agente", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_cancelar_pedidos_Click()
   Me.txt_agente = ""
   Me.txt_nombre_agente = ""
   Me.txt_paqueteria = ""
   Me.txt_nombre_paqueteria = ""
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Me.frm_lista.Visible = False
   Top = 2700
   Left = 2500
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
   
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     If VAR_TIPO_LISTA = 1 Then
        Me.txt_agente = Me.lv_lista.selectedItem
        Me.txt_nombre_agente = Me.lv_lista.selectedItem.SubItems(1)
        Me.txt_agente.SetFocus
     End If
     If VAR_TIPO_LISTA = 2 Then
        Me.txt_paqueteria = Me.lv_lista.selectedItem
        Me.txt_nombre_paqueteria = Me.lv_lista.selectedItem.SubItems(1)
        Me.txt_paqueteria.SetFocus
     End If
   End If
   If KeyAscii = 27 Then
      If VAR_TIPO_LISTA = 1 Then
         Me.txt_agente.SetFocus
      End If
      If VAR_TIPO_LISTA = 2 Then
         Me.txt_paqueteria.SetFocus
      End If
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub txt_agente_Change()
   Me.txt_nombre_agente = ""
   Me.txt_paqueteria = ""
   Me.txt_nombre_paqueteria = ""
   Me.txt_contable = ""
End Sub

Private Sub txt_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_agentes order by vcha_age_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_age_agente_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
      VAR_TIPO_LISTA = 1
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_agente_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_agente_LostFocus()
   If Me.txt_agente <> "" Then
      rs.Open "SELECT * FROM TB_AGENTES WHERE VCHA_aGE_aGENTE_ID = '" + Me.txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_agente = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
         Me.txt_contable = IIf(IsNull(rs!VCHA_AGE_CONTABLE), "", rs!VCHA_AGE_CONTABLE)
         rsaux2.Open "SELECT * FROM TB_PAQUETERIA WHERE VCHA_PAQ_CLAVE_ID = '" + IIf(IsNull(rs!VCHA_PAQ_PAQUETERIA_ID), "", rs!VCHA_PAQ_PAQUETERIA_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux2.EOF Then
            Me.txt_paqueteria = IIf(IsNull(rsaux2!vcha_paq_clave_id), "", rsaux2!vcha_paq_clave_id)
            Me.txt_nombre_paqueteria = IIf(IsNull(rsaux2!vcha_paq_nombre), "", rsaux2!vcha_paq_nombre)
         Else
            Me.txt_paqueteria = ""
            Me.txt_nombre_paqueteria = ""
         End If
         rsaux2.Close
      Else
         MsgBox "Clave de agente incorrecta", vbOKOnly, "ATENCION"
         Me.txt_nombre_agente = ""
         Me.txt_agente = ""
      End If
      rs.Close
   Else
      Me.txt_nombre_agente = ""
   End If

End Sub

Private Sub txt_contable_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_aceptar_pedidos.SetFocus
   End If
End Sub

Private Sub txt_nombre_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_agentes order by vcha_age_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_age_agente_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
      VAR_TIPO_LISTA = 1
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
     KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_paqueteria_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_PAQUETERIA where vcha_paq_nombre like '%ALMACEN%'  order by vcha_PAQ_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_paq_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_paq_nombre), "", rs!vcha_paq_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "PAQUETERIAS"
      VAR_TIPO_LISTA = 2
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_paqueteria_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_paqueteria_Change()
   Me.txt_nombre_paqueteria = ""
End Sub

Private Sub txt_paqueteria_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_PAQUETERIA where vcha_paq_nombre like '%ALMACEN%' order by vcha_PAQ_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_paq_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_paq_nombre), "", rs!vcha_paq_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "PAQUETERIAS"
      VAR_TIPO_LISTA = 2
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_paqueteria_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_paqueteria_LostFocus()
   If Me.txt_paqueteria <> "" Then
      rs.Open "select * from TB_PAQUETERIA where vcha_paq_CLAVE_ID = '" + Me.txt_paqueteria + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_paqueteria = IIf(IsNull(rs!vcha_paq_nombre), "", rs!vcha_paq_nombre)
      Else
         MsgBox "Paqueteria incorrecta", vbOKOnly, "ATENCION"
         Me.txt_nombre_paqueteria = ""
         Me.txt_paqueteria = ""
      End If
      rs.Close
   Else
      Me.txt_nombre_paqueteria = ""
   End If
End Sub
