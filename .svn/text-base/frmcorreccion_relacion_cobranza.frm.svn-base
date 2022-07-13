VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcorreccion_relacion_cobranza 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   " Folio "
      Height          =   660
      Left            =   120
      TabIndex        =   1
      Top             =   165
      Width           =   5070
      Begin VB.TextBox txt_relacion 
         Height          =   360
         Left            =   1620
         TabIndex        =   2
         Top             =   195
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Detalle relación "
      Height          =   2670
      Left            =   120
      TabIndex        =   0
      Top             =   900
      Width           =   5085
      Begin VB.Frame frm_nuevo_folio 
         Height          =   915
         Left            =   315
         TabIndex        =   4
         Top             =   840
         Width           =   2145
         Begin VB.TextBox txt_nuevo_folio 
            Height          =   360
            Left            =   210
            TabIndex        =   5
            Top             =   405
            Width           =   1695
         End
         Begin VB.Label Label1 
            BackColor       =   &H8000000D&
            Caption         =   " Nuevo Folio"
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   0
            TabIndex        =   6
            Top             =   15
            Width           =   2130
         End
      End
      Begin MSComctlLib.ListView lv_detalle_relacion 
         Height          =   2325
         Left            =   90
         TabIndex        =   3
         Top             =   225
         Width           =   4860
         _ExtentX        =   8573
         _ExtentY        =   4101
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
            Text            =   "Folio"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Factura"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Consecutivo"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "frmcorreccion_relacion_cobranza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Me.frm_nuevo_folio.Visible = False
   Top = 2500
   Left = 3000
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_LostFocus()
   Me.frm_nuevo_folio.Visible = False
End Sub

Private Sub lv_detalle_relacion_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Me.frm_nuevo_folio.Visible = True
      Me.txt_nuevo_folio = ""
      Me.txt_nuevo_folio.SetFocus
   End If
End Sub

Private Sub txt_nuevo_folio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      var_si = MsgBox("¿Desea cambiar el folio al registro seleccionado?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         rs.Open "UPDATE TB_RELACION_COBRANZA SET VCHA_rCO_FOLIO = '" + Me.txt_nuevo_folio + "' WHERE VCHA_rCO_FOLIO = '" + Me.lv_detalle_relacion.selectedItem + "' AND INTE_RCO_CONSECUTIVO = " + CStr(CDbl(Me.lv_detalle_relacion.selectedItem.SubItems(2))), cnn, adOpenDynamic, adLockOptimistic
         Me.lv_detalle_relacion.ListItems.Clear
         cnn.CommandTimeout = 360
         rs.Open "select * from tb_Relacion_cobranza with (nolock) where vcha_rco_folio = '" + Me.txt_relacion + "'", cnn, adOpenDynamic, adLockBatchOptimistic
         While Not rs.EOF
               Set list_item = Me.lv_detalle_relacion.ListItems.Add(, , rs!vcha_Rco_folio)
               list_item.SubItems(1) = IIf(IsNull(rs!inte_Car_numero), "", rs!inte_Car_numero)
               list_item.SubItems(2) = IIf(IsNull(rs!inte_rco_Consecutivo), "", rs!inte_rco_Consecutivo)
               rs.MoveNext
         Wend
         rs.Close
         Me.lv_detalle_relacion.SetFocus
         MsgBox "Se a cambiado el registro", vbOKCancel, "ATENCION"
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_nuevo_folio.Visible = False
   End If
End Sub

Private Sub txt_nuevo_folio_LostFocus()
   Me.frm_nuevo_folio.Visible = False
End Sub

Private Sub txt_relacion_Change()
   Me.lv_detalle_relacion.ListItems.Clear
End Sub

Private Sub txt_relacion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.lv_detalle_relacion.SetFocus
   End If
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub txt_relacion_LostFocus()
   If Trim(Me.txt_relacion) <> "" Then
      Me.lv_detalle_relacion.ListItems.Clear
      cnn.CommandTimeout = 360
      rs.Open "select * from tb_Relacion_cobranza with (nolock) where vcha_rco_folio = '" + Me.txt_relacion + "'", cnn, adOpenDynamic, adLockBatchOptimistic
      While Not rs.EOF
            Set list_item = Me.lv_detalle_relacion.ListItems.Add(, , rs!vcha_Rco_folio)
            list_item.SubItems(1) = IIf(IsNull(rs!inte_Car_numero), "", rs!inte_Car_numero)
            list_item.SubItems(2) = IIf(IsNull(rs!inte_rco_Consecutivo), "", rs!inte_rco_Consecutivo)
            rs.MoveNext
      Wend
      rs.Close
   Else
      Me.lv_detalle_relacion.ListItems.Clear
   End If
End Sub
