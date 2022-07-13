VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmcambio_agente_cartera 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambio de agente en cartera"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      Picture         =   "frmcambio_agente_cartera.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Cambiar agente a clientes"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame frm_lista 
      Height          =   1740
      Left            =   1245
      TabIndex        =   10
      Top             =   15
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1200
         Left            =   30
         TabIndex        =   11
         Top             =   465
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   2117
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
            Object.Width           =   7937
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   45
         TabIndex        =   12
         Top             =   150
         Width           =   5610
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Cambio de Agente en Cartera "
      Height          =   1245
      Left            =   150
      TabIndex        =   7
      Top             =   465
      Width           =   7500
      Begin VB.TextBox txt_nombre_nuevo 
         Height          =   315
         Left            =   2445
         TabIndex        =   4
         Top             =   750
         Width           =   4935
      End
      Begin VB.TextBox txt_agente_nuevo 
         Height          =   315
         Left            =   1500
         TabIndex        =   3
         Top             =   750
         Width           =   900
      End
      Begin VB.TextBox txt_nombre_anterior 
         Height          =   315
         Left            =   2445
         TabIndex        =   2
         Top             =   345
         Width           =   4935
      End
      Begin VB.TextBox txt_agente_anterior 
         Height          =   315
         Left            =   1515
         TabIndex        =   1
         Top             =   345
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Agente Nuevo:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   810
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Agente Anterior:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   405
         Width           =   1140
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7230
      Picture         =   "frmcambio_agente_cartera.frx":04F2
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   360
   End
   Begin VB.CommandButton cmd_aceptar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      Picture         =   "frmcambio_agente_cartera.frx":0B2C
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Height          =   60
      Left            =   0
      TabIndex        =   0
      Top             =   330
      Width           =   7710
   End
End
Attribute VB_Name = "frmcambio_agente_cartera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tipo_lista As Integer
Private Sub cmd_aceptar_pedidos_Click()
   If Trim(txt_agente_anterior) <> "" Then
      If Trim(txt_agente_nuevo) <> "" Then
         If txt_agente_anterior <> txt_agente_nuevo Then
            var_si = MsgBox("¿Desea cambiar el agente?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               var_si = MsgBox("Confirmar el cambio de agente", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  rs.Open "EXEC SP_ACTUALIZA_AGENTE_CARTERA '" + txt_agente_anterior + "', '" + txt_agente_nuevo + "'", cnn, adOpenDynamic, adLockOptimistic
                  MsgBox "Se a terminado el cambio del agente", vbOKOnly, "ATENCION"
               End If
            End If
         Else
            MsgBox "Claves de agentes iguales", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Se debe de seleccionar un agente nuevo", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Se debe de seleccionar un agente anterior", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   If Trim(txt_agente_anterior) <> "" Then
      If Trim(txt_agente_nuevo) <> "" Then
         If txt_agente_anterior <> txt_agente_nuevo Then
            var_si = MsgBox("¿Desea cambiar el agente?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               var_si = MsgBox("Confirmar el cambio de agente", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  rs.Open "UPDATE TB_CLIENTES set vcha_age_agente_id = '" + txt_agente_nuevo + "' where vcha_age_agente_id = '" + txt_agente_anterior + "'", cnn, adOpenDynamic, adLockOptimistic
                  MsgBox "Se a terminado el cambio del agente", vbOKOnly, "ATENCION"
               End If
            End If
         Else
            MsgBox "Claves de agentes iguales", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Se debe de seleccionar un agente nuevo", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Se debe de seleccionar un agente anterior", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 2500
   Left = 2000
   frm_lista.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   var_swpassword = False
   var_modifica_registro = False
   Call activa_forma(var_activa_forma_reporte_catalogo_articulos)
End Sub

Private Sub Text2_Change()

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_tipo_lista = 1 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_agente_anterior = lv_lista.selectedItem
            txt_nombre_anterior = lv_lista.selectedItem.SubItems(1)
         Else
            txt_agente_anterior = ""
            txt_nombre_anterior = ""
         End If
         txt_agente_anterior.SetFocus
      End If
      If var_tipo_lista = 2 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_agente_nuevo = lv_lista.selectedItem
            txt_nombre_nuevo = lv_lista.selectedItem.SubItems(1)
         Else
            txt_agente_nuevo = ""
            txt_nombre_nuevo = ""
         End If
         txt_agente_nuevo.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub txt_agente_anterior_KeyDown(KeyCode As Integer, Shift As Integer)
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
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 4 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_agente_anterior_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_agente_anterior_LostFocus()
   If Trim(txt_agente_anterior) <> "" Then
      rs.Open "select * from tb_agentes where vcha_age_agente_id = '" + txt_agente_anterior + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_anterior = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
      Else
         MsgBox "Clave de agente incorrecta", vbOKOnly, "ATENCION"
         txt_agente_anterior = ""
         txt_nombre_anterior = ""
      End If
      rs.Close
   Else
      txt_nombre_anterior = ""
   End If
End Sub

Private Sub txt_agente_nuevo_Change()
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_agente_nuevo_KeyDown(KeyCode As Integer, Shift As Integer)
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
      var_tipo_lista = 2
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 4 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_agente_nuevo_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_agente_nuevo_LostFocus()
   If Trim(txt_agente_nuevo) <> "" Then
      rs.Open "select * from tb_agentes where vcha_age_agente_id = '" + txt_agente_nuevo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_nuevo = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
      Else
         MsgBox "Clave de agente incorrecta", vbOKOnly, "ATENCION"
         txt_agente_nuevo = ""
         txt_nombre_nuevo = ""
      End If
      rs.Close
   Else
      txt_nombre_nuevo = ""
   End If
End Sub

Private Sub txt_nombre_anterior_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frm_lista.Left = 1530
      frm_lista.Top = 2220
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_agentes order by vcha_age_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_age_agente_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 4 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_anterior_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_nuevo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frm_lista.Left = 1530
      frm_lista.Top = 2220
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_agentes order by vcha_age_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_age_agente_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
      var_tipo_lista = 2
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 4 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_nuevo_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub
