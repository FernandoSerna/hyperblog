VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmembarques_activos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Embarques Activos"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   11880
   Begin VB.Frame Frame1 
      Height          =   5205
      Left            =   30
      TabIndex        =   0
      Top             =   -45
      Width           =   11790
      Begin MSComctlLib.ListView lv_embarques 
         Height          =   5040
         Left            =   30
         TabIndex        =   1
         Top             =   120
         Width           =   11730
         _ExtentX        =   20690
         _ExtentY        =   8890
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Embarque"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Jaula     "
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "       Fecha     "
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "                                       Agente"
            Object.Width           =   6826
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Cantidad Surtir     "
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Cantidad Surtida  "
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Progreso %     "
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "frmembarques_activos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   Dim list_item As ListItem
   Dim var_cantidad_surtir As Double
   Dim var_cantidad_surtida As Double
   Dim var_progreso As Double
   var_cadena_seguridad = ""
   If var_empresa = "30" Then
      rs.Open "select * from vw_embarques_activos where char_emb_estatus = '' AND VCHA_EMP_EMPRESA_ID = '30'", cnn, adOpenDynamic, adLockOptimistic
   Else
      rs.Open "select * from vw_embarques_activos where char_emb_estatus = ''", cnn, adOpenDynamic, adLockOptimistic
   End If
   If Not rs.EOF Then
      While Not rs.EOF
         var_progreso = 0
         Set list_item = lv_embarques.ListItems.Add(, , rs!INTE_EMB_EMBARQUE)
         list_item.SubItems(1) = IIf(IsNull(rs!inte_jau_jaula_id), "", rs!inte_jau_jaula_id)
         list_item.SubItems(2) = IIf(IsNull(rs!DTIM_EMB_FECHA_INICIO), "", rs!DTIM_EMB_FECHA_INICIO)
         list_item.SubItems(3) = IIf(IsNull(rs!vcha_age_nombre), "", rs!vcha_age_nombre)
         list_item.SubItems(4) = Format(IIf(IsNull(rs!cantidad_surtir), 0, rs!cantidad_surtir), "###,###,##0.00")
         list_item.SubItems(5) = Format(IIf(IsNull(rs!cantidad_surtida), 0, rs!cantidad_surtida), "###,###,##0.00")
         var_cantidad_surtir = IIf(IsNull(rs!cantidad_surtir), "", rs!cantidad_surtir)
         var_cantidad_surtida = IIf(IsNull(rs!cantidad_surtida), 0, rs!cantidad_surtida)
         var_progreso = (var_cantidad_surtida * 100) / var_cantidad_surtir
         list_item.SubItems(6) = Format(var_progreso, "###,###,##0.00")
         rs.MoveNext
      Wend
   Else
   End If
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_embarques_activos)
End Sub

Private Sub lv_embarques_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_embarques, ColumnHeader)
End Sub
