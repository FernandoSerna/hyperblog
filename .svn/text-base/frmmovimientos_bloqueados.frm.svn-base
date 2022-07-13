VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmovimientos_bloqueados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos Bloqueados"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   11640
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   450
      Picture         =   "frmmovimientos_bloqueados.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancelar Alt + C"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmmovimientos_bloqueados.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   6825
      Left            =   60
      TabIndex        =   0
      Top             =   450
      Width           =   11475
      Begin MSComctlLib.ListView lv_bloqueos 
         Height          =   6615
         Left            =   45
         TabIndex        =   1
         Top             =   150
         Width           =   11340
         _ExtentX        =   20003
         _ExtentY        =   11668
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Empresa"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Unidad"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Clave Almacen"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Almacen"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Clave movimiento"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Movimiento"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Número "
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Usuario"
            Object.Width           =   6879
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Máquina"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   120
      TabIndex        =   4
      Top             =   270
      Width           =   11430
   End
End
Attribute VB_Name = "frmmovimientos_bloqueados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_Click()
   If lv_bloqueos.ListItems.Count > 0 Then
      var_si = MsgBox("Deseas desbloquear el movimiento " + Trim(lv_bloqueos.selectedItem.SubItems(5)) + " Número " + Trim(lv_bloqueos.selectedItem.SubItems(6)), vbYesNo, "ATENCION")
      If var_si = 6 Then
         var_si = MsgBox("Confirmar el desbloqueo del movimiento", vbYesNo, "ATENCION")
         If var_si = 6 Then
            rs.Open "UPDATE TB_ENCABEZADO_MOVIMIENTOS SET INTE_EMO_BLOQUEADO = 0 WHERE VCHA_EMP_EMPRESA_ID = '" + lv_bloqueos.selectedItem + "' AND VCHA_UOR_UNIDAD_ID = '" + lv_bloqueos.selectedItem.SubItems(1) + "' AND VCHA_ALM_ALMACEN_ID = '" + lv_bloqueos.selectedItem.SubItems(2) + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + lv_bloqueos.selectedItem.SubItems(4) + "' AND INTE_EMO_NUMERO =" + lv_bloqueos.selectedItem.SubItems(6), cnn, adOpenDynamic, adLockOptimistic
            lv_bloqueos.ListItems.Clear
            rs.Open "select * from vw_movimientos_bloqueados where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
            Dim list_item As ListItem
            If Not rs.EOF Then
               While Not rs.EOF
                     Set list_item = lv_bloqueos.ListItems.Add(, , rs!VCHA_EMP_EMPRESA_ID)
                     list_item.SubItems(1) = IIf(IsNull(rs!vcha_uor_unidad_id), "", rs!vcha_uor_unidad_id)
                     list_item.SubItems(2) = IIf(IsNull(rs!VCHA_ALM_ALMACEN_ID), "", rs!VCHA_ALM_ALMACEN_ID)
                     list_item.SubItems(3) = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
                     list_item.SubItems(4) = IIf(IsNull(rs!VCHA_MOV_MOVIMIENTO_ID), "", rs!VCHA_MOV_MOVIMIENTO_ID)
                     list_item.SubItems(5) = IIf(IsNull(rs!vcha_mov_nombre), 0, rs!vcha_mov_nombre)
                     list_item.SubItems(6) = IIf(IsNull(rs!INTE_EMO_NUMERO), 0, rs!INTE_EMO_NUMERO)
                     list_item.SubItems(7) = IIf(IsNull(rs!VCHA_USU_NOMBRE), "", rs!VCHA_USU_NOMBRE) + " " + IIf(IsNull(rs!VCHA_USU_APELLIDOS), "", rs!VCHA_USU_APELLIDOS)
                     list_item.SubItems(8) = IIf(IsNull(rs!vcha_aud_maquina), "", rs!vcha_aud_maquina)
                     rs.MoveNext:
               Wend
               rs.MoveFirst
            End If
            rs.Close
         End If
      End If
   End If
End Sub

Private Sub cmd_cancelar_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 0
   rs.Open "select * from vw_movimientos_bloqueados where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
   Dim list_item As ListItem
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = lv_bloqueos.ListItems.Add(, , rs!VCHA_EMP_EMPRESA_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_uor_unidad_id), "", rs!vcha_uor_unidad_id)
            list_item.SubItems(2) = IIf(IsNull(rs!VCHA_ALM_ALMACEN_ID), "", rs!VCHA_ALM_ALMACEN_ID)
            list_item.SubItems(3) = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
            list_item.SubItems(4) = IIf(IsNull(rs!VCHA_MOV_MOVIMIENTO_ID), "", rs!VCHA_MOV_MOVIMIENTO_ID)
            list_item.SubItems(5) = IIf(IsNull(rs!vcha_mov_nombre), 0, rs!vcha_mov_nombre)
            list_item.SubItems(6) = IIf(IsNull(rs!INTE_EMO_NUMERO), 0, rs!INTE_EMO_NUMERO)
            list_item.SubItems(7) = IIf(IsNull(rs!VCHA_USU_NOMBRE), "", rs!VCHA_USU_NOMBRE) + " " + IIf(IsNull(rs!VCHA_USU_APELLIDOS), "", rs!VCHA_USU_APELLIDOS)
            list_item.SubItems(8) = IIf(IsNull(rs!vcha_aud_maquina), "", rs!vcha_aud_maquina)
            rs.MoveNext:
      Wend
      rs.MoveFirst
   End If
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_movimientos_bloqueados)
End Sub

Private Sub lv_bloqueos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_bloqueos, ColumnHeader)
End Sub

Private Sub lv_bloqueos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_aceptar.SetFocus
   End If
End Sub
