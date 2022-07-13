VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmanticipos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anticipos"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   " Anticipos  "
      Height          =   2565
      Left            =   45
      TabIndex        =   0
      Top             =   105
      Width           =   3735
      Begin MSComctlLib.ListView lv_lista 
         Height          =   2265
         Left            =   75
         TabIndex        =   1
         Top             =   225
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   3995
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
            Text            =   "Fecha"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Importe"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Consecutivo"
            Object.Width           =   0
         EndProperty
      End
   End
End
Attribute VB_Name = "frmanticipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Dim list_item As ListItem
   If Trim(var_cliente_anticipo) <> "" Then
      rs.Open "select * from tb_anticipos where vcha_cli_clave_id = '" + var_cliente_anticipo + "' and INTE_ant_cargado = 0", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , Trim(rs!dtim_Sal_fecha))
            list_item.SubItems(1) = Format(IIf(IsNull(rs!floa_sal_cantidad * rs!floa_Sal_precio), 0, rs!floa_sal_cantidad * rs!floa_Sal_precio) * 1.16, "###,###,##0.00")
            list_item.SubItems(2) = IIf(IsNull(rs!INTE_ANT_CONSECUTIVO), 0, rs!INTE_ANT_CONSECUTIVO)
            rs.MoveNext
      Wend
      rs.Close
   End If
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      var_consecutivo_anticipo = Me.lv_lista.selectedItem.SubItems(2)
      var_importe_anticipo = CDbl(Me.lv_lista.selectedItem.SubItems(1))
      Unload Me
   End If
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub
