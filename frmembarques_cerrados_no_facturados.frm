VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmembarques_cerrados_no_facturados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Embarques cerrados no facturados"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7785
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSComctlLib.ListView lv_embarques 
      Height          =   4920
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   8678
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
         Text            =   "Empresa"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Embarque"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Agente"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Creado"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Cerrado"
         Object.Width           =   2117
      EndProperty
   End
End
Attribute VB_Name = "frmembarques_cerrados_no_facturados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Dim list_item As ListItem
   Top = 1000
   Left = 1000
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   If var_empresa = "30" Or var_empresa = "15" Or var_empresa = "31" Or var_empresa = "28" Then
      If var_empresa = "30" Then
         rsaux.Open "select * from VW_EMBARQUES_CERRADOS_NO_FACTURADOS where vcha_emp_empresa_id = '30'", cnn, adOpenDynamic, adLockOptimistic
      Else
         If var_empresa = "15" Then
            rsaux.Open "select * from VW_EMBARQUES_CERRADOS_NO_FACTURADOS where vcha_emp_empresa_id = '15'", cnn, adOpenDynamic, adLockOptimistic
         Else
            rsaux.Open "select * from VW_EMBARQUES_CERRADOS_NO_FACTURADOS where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
         End If
      End If
   Else
      If var_empresa = "18" Then
         rsaux.Open "select * from VW_EMBARQUES_CERRADOS_NO_FACTURADOS where vcha_emp_empresa_id = '18'", cnn, adOpenDynamic, adLockOptimistic
      Else
         rsaux.Open "select * from VW_EMBARQUES_CERRADOS_NO_FACTURADOS where vcha_emp_empresa_id <> '16'", cnn, adOpenDynamic, adLockOptimistic
      End If
   End If
   If Not rsaux.EOF Then
      While Not rsaux.EOF
            Set list_item = lv_embarques.ListItems.Add(, , rsaux!VCHA_EMP_EMPRESA_ID)
            list_item.SubItems(1) = IIf(IsNull(rsaux!inte_emb_embarque), "", rsaux!inte_emb_embarque)
            list_item.SubItems(2) = IIf(IsNull(rsaux!VCHA_AGE_NOMBRE), "", rsaux!VCHA_AGE_NOMBRE)
            list_item.SubItems(3) = Format(IIf(IsNull(rsaux!DTIM_EMB_FECHA_INICIO), "", rsaux!DTIM_EMB_FECHA_INICIO), "Short date")
            list_item.SubItems(4) = Format(IIf(IsNull(rsaux!DTIM_EMB_FECHA_FINAL), "", rsaux!DTIM_EMB_FECHA_FINAL), "Short date")
            rsaux.MoveNext
      Wend
   Else
      Unload Me
   End If
   rsaux.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_detalle_cajas)
End Sub

Private Sub lv_movimientos_KeyPress(KeyAscii As Integer)
End Sub

Private Sub lv_embarques_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_embarques, ColumnHeader)
End Sub

Private Sub lv_embarques_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      If Me.lv_embarques.ListItems.Count > 0 Then
         var_numero_embarque_global = CDbl(Me.lv_embarques.selectedItem.SubItems(1))
      End If
      Unload Me
   End If
   If KeyAscii = 13 Then
      If Me.lv_embarques.ListItems.Count > 0 Then
         
         var_numero_embarque_global = CDbl(Me.lv_embarques.selectedItem.SubItems(1))
      End If
      Unload Me
   End If
End Sub
