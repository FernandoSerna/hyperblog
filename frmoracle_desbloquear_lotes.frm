VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_desbloquear_lotes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Desbloquear lotes"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   5565
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin MSComctlLib.ListView lv_lotes_bloqueados 
         Height          =   5340
         Left            =   60
         TabIndex        =   1
         Top             =   180
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   9419
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
            Text            =   "Lote"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Usuario"
            Object.Width           =   8555
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Maquina"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Clave usuario"
            Object.Width           =   0
         EndProperty
      End
   End
End
Attribute VB_Name = "frmoracle_desbloquear_lotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Dim list_item As ListItem

Private Sub Form_Load()
   Top = 1000
   Left = 1300
   rs.Open "SELECT * FROM TB_ORACLE_BLOQUEO_PEDIDOS_LOTES", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         var_pedido_s = CStr(rs!pedido)
         var_lote_s = CStr(rs!lote)
         If Len(var_pedido_s) = 6 Then
            var_pedido_s = "0" + var_pedido_s
         End If
         If Len(var_lote_s) = 1 Then
            var_lote_s = "00" + var_lote_s
         Else
              If Len(var_lote_s) = 2 Then
                 var_lote_s = "0" + var_lote_s
              End If
         End If
         var_lote_str = var_pedido_s + var_lote_s
         rsaux.Open "select * from tb_usuarios where vcha_usu_usuario_id = '" + IIf(IsNull(rs!usuario), "", rs!usuario) + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            var_nombre_usuario = IIf(IsNull(rsaux!vcha_usu_nombre), "", rsaux!vcha_usu_nombre) + " " + IIf(IsNull(rsaux!vcha_usu_apellidos), "", rsaux!vcha_usu_apellidos)
         Else
            var_nombre_usuario = ""
         End If
         rsaux.Close
         Set list_item = Me.lv_lotes_bloqueados.ListItems.Add(, , var_lote_str)
         list_item.SubItems(1) = var_nombre_usuario
         list_item.SubItems(2) = IIf(IsNull(rs!maquina), "", rs!maquina)
         list_item.SubItems(3) = IIf(IsNull(rs!usuario), "", rs!usuario)
         
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_lotes_bloqueados_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_lotes_bloqueados, ColumnHeader)
End Sub

Private Sub lv_lotes_bloqueados_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      If Me.lv_lotes_bloqueados.ListItems.Count > 0 Then
         var_si_permiso = 0
         frmoracle_permiso_cerrar_pedidos.Show 1
         If var_si_permiso = 1 Then
            var_pedido = Mid(Me.lv_lotes_bloqueados.selectedItem, 1, Len(Me.lv_lotes_bloqueados.selectedItem) - 3)
            var_lote = Mid(Me.lv_lotes_bloqueados.selectedItem, Len(Me.lv_lotes_bloqueados.selectedItem) - 2, 3)
            rs.Open "insert into tb_bitacora_desbloqueo (lote, usuario, fecha) values ('" + Me.lv_lotes_bloqueados.selectedItem + "','" + var_clave_usuario_global + "', getdate())", cnn, adOpenDynamic, adLockOptimistic
            rs.Open "delete from TB_ORACLE_BLOQUEO_PEDIDOS_LOTES where pedido = " + CStr(var_pedido) + " and lote = " + CStr(var_lote) + " and maquina = '" + Me.lv_lotes_bloqueados.selectedItem.SubItems(2) + "' and usuario = '" + Me.lv_lotes_bloqueados.selectedItem.SubItems(3) + "'", cnn, adOpenDynamic, adLockOptimistic
            Me.lv_lotes_bloqueados.ListItems.Remove (Me.lv_lotes_bloqueados.selectedItem.Index)
         Else
            MsgBox "No tiene autorizacion para desbloquear lotes", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub

Private Sub lv_lotes_bloqueados_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

