VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_asignacion_causas_devolucion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignación de causas de devolución "
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   15330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txt_folio_ANC 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   18
      Top             =   413
      Width           =   2775
   End
   Begin VB.TextBox txt_codigo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   17
      Top             =   413
      Width           =   2775
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   30
      Picture         =   "frmoracle_asignacion_causas_devolucion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   14925
      Picture         =   "frmoracle_asignacion_causas_devolucion.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   300
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   330
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   0
      TabIndex        =   3
      Top             =   270
      Width           =   15240
   End
   Begin VB.Frame Frame2 
      Height          =   7500
      Left            =   15
      TabIndex        =   4
      Top             =   945
      Width           =   15315
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   375
         Picture         =   "frmoracle_asignacion_causas_devolucion.frx":073C
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   45
         Picture         =   "frmoracle_asignacion_causas_devolucion.frx":0952
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1035
         Picture         =   "frmoracle_asignacion_causas_devolucion.frx":0A54
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   705
         Picture         =   "frmoracle_asignacion_causas_devolucion.frx":0B26
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Marcar (Enter)"
         Top             =   135
         Width           =   330
      End
      Begin VB.Frame frm_lista 
         Height          =   2400
         Left            =   3000
         TabIndex        =   7
         Top             =   3720
         Width           =   5970
         Begin MSComctlLib.ListView lv_lista 
            Height          =   1950
            Left            =   60
            TabIndex        =   8
            Top             =   405
            Width           =   5865
            _ExtentX        =   10345
            _ExtentY        =   3440
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
            BackColor       =   &H000000C0&
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   30
            TabIndex        =   9
            Top             =   120
            Width           =   5895
         End
      End
      Begin VB.Frame frm_mensaje 
         Height          =   1215
         Left            =   570
         TabIndex        =   5
         Top             =   3645
         Width           =   7065
         Begin VB.Label lbl_mensaje 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            Caption         =   "estatus"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1200
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   7050
         End
      End
      Begin MSComctlLib.ListView lv_devoluciones 
         Height          =   6915
         Left            =   45
         TabIndex        =   14
         Top             =   480
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   12197
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
            Text            =   "Código"
            Object.Width           =   2478
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Causa"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Clave causa"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Consecutivo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "inventory_item_id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Localizador"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Causa Real"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Clave causa real"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Folio ANC"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Folio ANC:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8040
      TabIndex        =   16
      Top             =   480
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código artículo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      TabIndex        =   15
      Top             =   480
      Width           =   2205
   End
End
Attribute VB_Name = "frmoracle_asignacion_causas_devolucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter
Dim objConn As New ADODB.Connection
Dim objCmd As New ADODB.Command
Dim objParm As ADODB.Parameter
Dim var_orden As String
Dim var_posicion As Double

Dim var_encontro As Double

Private Sub cmd_invertir_Click()
   n = lv_devoluciones.ListItems.Count
   For i = 1 To n
      lv_devoluciones.ListItems.Item(i).Selected = True
      If lv_devoluciones.selectedItem.SubItems(4) = "*" Then
         lv_devoluciones.selectedItem.SubItems(4) = ""
         lv_devoluciones.ListItems.Item(i).Bold = False
         lv_devoluciones.ListItems.Item(i).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(4).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(5).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(9).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(10).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(9).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(10).ForeColor = &H80000012
      Else
         lv_devoluciones.selectedItem.SubItems(4) = "*"
         lv_devoluciones.ListItems.Item(i).Bold = True
         lv_devoluciones.ListItems.Item(i).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(5).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(9).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(10).ForeColor = &HFF0000
      End If
   Next i
   If Me.lv_devoluciones.ListItems.Count > 0 Then
      Me.lv_devoluciones.SetFocus
   End If
End Sub

Private Sub cmd_marcar_Click()
   i = lv_devoluciones.selectedItem.Index
   If lv_devoluciones.selectedItem.SubItems(4) = "*" Then
      lv_devoluciones.selectedItem.SubItems(4) = ""
      lv_devoluciones.ListItems.Item(i).Bold = False
      lv_devoluciones.ListItems.Item(i).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(4).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(5).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(9).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(10).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(9).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(10).ForeColor = &H80000012
      lv_devoluciones.Refresh
   Else
      lv_devoluciones.selectedItem.SubItems(4) = "*"
      lv_devoluciones.ListItems.Item(i).Bold = True
      lv_devoluciones.ListItems.Item(i).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(3).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(4).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(5).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(9).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(10).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(9).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(10).ForeColor = &HFF0000
      lv_devoluciones.Refresh
   End If
   If Me.lv_devoluciones.ListItems.Count > 0 Then
      Me.lv_devoluciones.SetFocus
   End If
End Sub

Private Sub cmd_ninguno_Click()
   n = lv_devoluciones.ListItems.Count
   For i = 1 To n
      lv_devoluciones.ListItems.Item(i).Selected = True
      lv_devoluciones.selectedItem.SubItems(4) = ""
      lv_devoluciones.ListItems.Item(i).Bold = False
      lv_devoluciones.ListItems.Item(i).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(4).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(5).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(9).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(10).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(9).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(10).ForeColor = &H80000012
   Next i
   lv_devoluciones.Refresh
   If Me.lv_devoluciones.ListItems.Count > 0 Then
      Me.lv_devoluciones.SetFocus
   End If

End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmd_todos_Click()
   n = lv_devoluciones.ListItems.Count
   For i = 1 To n
      lv_devoluciones.ListItems.Item(i).Selected = True
      lv_devoluciones.selectedItem.SubItems(4) = "*"
      lv_devoluciones.ListItems.Item(i).Bold = True
      lv_devoluciones.ListItems.Item(i).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(3).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(4).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(5).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(9).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(10).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(11).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(9).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(10).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(11).ForeColor = &HFF0000
   Next i
   lv_devoluciones.Refresh
   If Me.lv_devoluciones.ListItems.Count > 0 Then
      Me.lv_devoluciones.SetFocus
   End If

End Sub

Private Sub Form_Load()
   rs.Open "select * from xxvia_tb_dev_clientes_desgloce where numero = " + CStr(var_numero_folio_devoluciones) + " and organizacion = " + var_unidad_organizacional + " AND MOVIMIENTO = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set list_item = Me.lv_devoluciones.ListItems.Add(, , IIf(IsNull(rs!codigo), "", rs!codigo))
         list_item.SubItems(1) = IIf(IsNull(rs!descripcion), "", rs!descripcion)
         list_item.SubItems(2) = Format(IIf(IsNull(rs!cantidad), 0, rs!cantidad), "###,###,##0.00")
         list_item.SubItems(3) = IIf(IsNull(rs!descripcion_causa), "", rs!descripcion_causa)
         list_item.SubItems(4) = ""
         list_item.SubItems(5) = IIf(IsNull(rs!causa_devolucion), "", rs!causa_devolucion)
         list_item.SubItems(6) = IIf(IsNull(rs!CONSECUTIVO), "", rs!CONSECUTIVO)
         list_item.SubItems(7) = IIf(IsNull(rs!inventory_item_id), 0, rs!inventory_item_id)
         list_item.SubItems(8) = IIf(IsNull(rs!localizador), "", rs!localizador)
         list_item.SubItems(9) = IIf(IsNull(rs!causa_real_descripcion), "", rs!causa_real_descripcion)
         list_item.SubItems(10) = IIf(IsNull(rs!causa_real), "", rs!causa_real)
         list_item.SubItems(11) = IIf(IsNull(rs!folio_anc), "", rs!folio_anc)
         rs.MoveNext
   Wend
   rs.Close
   Me.frm_lista.Visible = False
   Me.frm_mensaje.Visible = False
End Sub

Private Sub lv_devoluciones_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_devoluciones, ColumnHeader)
End Sub

Private Sub lv_devoluciones_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = 115 Then
      rs.Open "select * from xxvia_tb_dev_clientes_DESGLOCE where NUMERO = " + CStr(var_numero_folio_devoluciones) + " and organizacion = " + var_unidad_organizacional + " and movimiento = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      VAR_ESTATUS = IIf(IsNull(rs!estatus), "", rs!estatus)
      rs.Close
      If VAR_ESTATUS = "I" Then
         Me.lv_lista.ListItems.Clear
         var_cadena = "select lookup_code as CODIGO, meaning as NOMBRE, description as DESCRIPCION From FND_LOOKUP_VALUES_VL where lookup_type = 'CREDIT_MEMO_REASON' and enabled_flag = 'Y' ORDER BY 1"
         rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = Me.lv_lista.ListItems.Add(, , rs(0).Value)
               list_item.SubItems(1) = rs(1).Value
               rs.MoveNext
         Wend
         rs.Close
         If lv_lista.ListItems.Count > 0 Then
            Me.frm_lista.Visible = True
            Me.lv_lista.SetFocus
         Else
            MsgBox "No existen causas de devolución", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El movimiento ya no puede ser modificado", vbOKOnly, "ATENCION"
      End If
      
   End If
End Sub

Private Sub lv_devoluciones_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      i = lv_devoluciones.selectedItem.Index
      If lv_devoluciones.selectedItem.SubItems(4) = "*" Then
         lv_devoluciones.selectedItem.SubItems(4) = ""
         lv_devoluciones.ListItems.Item(i).Bold = False
         lv_devoluciones.ListItems.Item(i).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(4).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(5).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(9).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(10).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(11).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(9).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(10).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(11).ForeColor = &H80000012
         lv_devoluciones.Refresh
      Else
         lv_devoluciones.selectedItem.SubItems(4) = "*"
         lv_devoluciones.ListItems.Item(i).Bold = True
         lv_devoluciones.ListItems.Item(i).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(5).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(9).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(10).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(11).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(9).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(10).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(11).ForeColor = &HFF0000
         lv_devoluciones.Refresh
      End If
   End If

End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frm_lista.Visible = False
   End If
   If KeyAscii = 13 Then
      var_si = MsgBox("Se asignara la causa de devolución a los artículos seleccionados", vbYesNo, "ATENCION")
      If var_si = 6 Then
         If var_encontro = 0 Then
            For var_j = 1 To lv_devoluciones.ListItems.Count
                Me.lv_devoluciones.ListItems(var_j).Selected = True
                If Me.lv_devoluciones.selectedItem.SubItems(4) = "*" Then
                   lv_devoluciones.selectedItem.SubItems(4) = ""
                   lv_devoluciones.ListItems.Item(var_j).Bold = False
                   lv_devoluciones.ListItems.Item(var_j).ForeColor = &H80000012
                   lv_devoluciones.ListItems.Item(var_j).ListSubItems(1).Bold = False
                   lv_devoluciones.ListItems.Item(var_j).ListSubItems(2).Bold = False
                   lv_devoluciones.ListItems.Item(var_j).ListSubItems(3).Bold = False
                   lv_devoluciones.ListItems.Item(var_j).ListSubItems(4).Bold = False
                   lv_devoluciones.ListItems.Item(var_j).ListSubItems(5).Bold = False
                   lv_devoluciones.ListItems.Item(var_j).ListSubItems(6).Bold = False
                   lv_devoluciones.ListItems.Item(var_j).ListSubItems(7).Bold = False
                   lv_devoluciones.ListItems.Item(var_j).ListSubItems(8).Bold = False
                   lv_devoluciones.ListItems.Item(var_j).ListSubItems(9).Bold = False
                   lv_devoluciones.ListItems.Item(var_j).ListSubItems(10).Bold = False
                   lv_devoluciones.ListItems.Item(var_j).ListSubItems(11).Bold = False
                   lv_devoluciones.ListItems.Item(var_j).ListSubItems(1).ForeColor = &H80000012
                   lv_devoluciones.ListItems.Item(var_j).ListSubItems(2).ForeColor = &H80000012
                   lv_devoluciones.ListItems.Item(var_j).ListSubItems(3).ForeColor = &H80000012
                   lv_devoluciones.ListItems.Item(var_j).ListSubItems(4).ForeColor = &H80000012
                   lv_devoluciones.ListItems.Item(var_j).ListSubItems(5).ForeColor = &H80000012
                   lv_devoluciones.ListItems.Item(var_j).ListSubItems(6).ForeColor = &H80000012
                   lv_devoluciones.ListItems.Item(var_j).ListSubItems(7).ForeColor = &H80000012
                   lv_devoluciones.ListItems.Item(var_j).ListSubItems(8).ForeColor = &H80000012
                   lv_devoluciones.ListItems.Item(var_j).ListSubItems(9).ForeColor = &H80000012
                   lv_devoluciones.ListItems.Item(var_j).ListSubItems(10).ForeColor = &H80000012
                   lv_devoluciones.ListItems.Item(var_j).ListSubItems(11).ForeColor = &H80000012
                
                
                
                
                   strconsulta = "update xxvia_tb_dev_clientes_desgloce set causa_real = ?, causa_real_descripcion = ?, folio_anc = ? where movimiento = ? and  numero = ? and consecutivo = ?"
                   With comandoORA
                        .ActiveConnection = cnnoracle_4
                        .CommandType = adCmdText
                        .CommandText = strconsulta
                        Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.lv_lista.selectedItem)
                        .Parameters.Append parametro
                        Set parametro = .CreateParameter(, adVarChar, adParamInput, 500, Me.lv_lista.selectedItem.SubItems(1))
                        .Parameters.Append parametro
                        Set parametro = .CreateParameter(, adVarChar, adParamInput, 500, Me.txt_folio_ANC)
                        .Parameters.Append parametro
                        Set parametro = .CreateParameter(, adVarChar, adParamInput, 500, var_clave_movimiento)
                        .Parameters.Append parametro
                        Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_numero_folio_devoluciones)
                        .Parameters.Append parametro
                        Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.lv_devoluciones.ListItems.Item(var_j).SubItems(6)))
                        .Parameters.Append parametro
                   End With
                   Set rsaux9 = comandoORA.execute
                   Set comandoORA = Nothing
                   Set parametro = Nothing
                   Me.lv_devoluciones.ListItems.Item(var_j).SubItems(9) = Me.lv_lista.selectedItem.SubItems(1)
                   Me.lv_devoluciones.ListItems.Item(var_j).SubItems(10) = Me.lv_lista.selectedItem
                   Me.lv_devoluciones.ListItems.Item(var_j).SubItems(11) = Me.txt_folio_ANC
                
                
                
                
                End If
            Next var_j
         Else
            Me.lv_devoluciones.ListItems.Item(var_encontro).SubItems(9) = Me.lv_lista.selectedItem.SubItems(1)
            Me.lv_devoluciones.ListItems.Item(var_encontro).SubItems(10) = Me.lv_lista.selectedItem
            Me.lv_devoluciones.ListItems.Item(var_encontro).SubItems(11) = Me.txt_folio_ANC
            strconsulta = "update xxvia_tb_dev_clientes_desgloce set causa_real = ?, causa_real_descripcion = ?, folio_anc = ? where movimiento = ? and  numero = ? and consecutivo = ?"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.lv_lista.selectedItem)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 500, Me.lv_lista.selectedItem.SubItems(1))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 500, Me.txt_folio_ANC)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 500, var_clave_movimiento)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, var_numero_folio_devoluciones)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.lv_devoluciones.ListItems.Item(var_encontro).SubItems(6)))
                 .Parameters.Append parametro
            End With
            Set rsaux9 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            Me.txt_codigo = ""
            Me.txt_folio_ANC = ""
            Me.txt_codigo.SetFocus
         End If
         Me.frm_lista.Visible = False
      Else
         Me.lv_devoluciones.SetFocus
      End If
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub txt_codigo_GotFocus()
   Me.txt_codigo = ""
   var_encontro = 0
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Me.txt_codigo <> "" Then
         strconsulta = "select * from xxvia_system_items_b where segment1 = ? and organization_id = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
              .Parameters.Append parametro
         End With
         Set rsaux9 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         
         
         If rsaux9.EOF Then
              strconsulta = "SELECT A.INVENTORY_ITEM_ID, B.DESCRIPTION, cross_reference, b.segment1, nvl(a.description,'') as localizador, b.UNIT_WEIGHT, nvl(a.attribute1,1) as cantidad FROM mtl_cross_references_b A, xxvia_system_items_b B WHERE A.inventory_item_id = B.inventory_item_id AND B.organization_id = ? AND CROSS_REFERENCE = ?"
             With comandoORA
                  .ActiveConnection = cnnoracle_4
                  .CommandType = adCmdText
                  .CommandText = strconsulta
                  Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
                  .Parameters.Append parametro
                  Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
                  .Parameters.Append parametro
            End With
            Set rsaux8 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            var_cantidad_leida = 1
            If Not rsaux8.EOF Then
               var_cantidad_leida = IIf(IsNull(rsaux8!cantidad), 1, rsaux8!cantidad)
               var_peso = IIf(IsNull(rsaux8!UNIT_WEIGHT), 0, rsaux8!UNIT_WEIGHT)
               If IIf(IsNull(rsaux8!localizador), "", rsaux8!localizador) <> "" Then
                  var_localizador_subinventario = txt_almacen + IIf(IsNull(rsaux8!localizador), "", rsaux8!localizador)
                  If var_localizador_subinventario <> "" Then
                     Me.txt_codigo = rsaux8!SEGMENT1
                  Else
                     Me.txt_codigo = ""
                     Me.txt_codigo = rsaux8!SEGMENT1
                  End If
               Else
                  Me.txt_codigo = ""
                  Me.txt_codigo = rsaux8!SEGMENT1
               End If
            Else
               Me.txt_codigo = ""
            End If
            rsaux8.Close
         Else
            var_cantidad_leida = 1
         End If
         rsaux9.Close
       End If
       If Me.txt_codigo <> "" Then
          If Me.lv_devoluciones.ListItems.Count > 0 Then
             var_encontro = 0
             For var_j = 1 To Me.lv_devoluciones.ListItems.Count
                 Me.lv_devoluciones.ListItems.Item(var_j).Selected = True
                 If Me.lv_devoluciones.selectedItem = Me.txt_codigo And Me.lv_devoluciones.selectedItem.SubItems(11) = "" Then
                    var_encontro = var_j
                 End If
             Next var_j
          End If
          If var_encontro > 0 Then
             Me.txt_folio_ANC.SetFocus
          Else
             MsgBox "El código leido no se encuentra en la relación", vbOKOnly, "ATENCION"
          End If
       End If
    End If
End Sub

Private Sub txt_folio_ANC_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.txt_folio_ANC <> "" Then
         rs.Open "select * from xxvia_tb_dev_clientes_DESGLOCE where NUMERO = " + CStr(var_numero_folio_devoluciones) + " and organizacion = " + var_unidad_organizacional + " and movimiento = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         VAR_ESTATUS = IIf(IsNull(rs!estatus), "", rs!estatus)
         rs.Close
         If VAR_ESTATUS = "I" Then
            Me.lv_lista.ListItems.Clear
            var_cadena = "select lookup_code as CODIGO, meaning as NOMBRE, description as DESCRIPCION From FND_LOOKUP_VALUES_VL where lookup_type = 'CREDIT_MEMO_REASON' and enabled_flag = 'Y' ORDER BY 1"
            rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  Set list_item = Me.lv_lista.ListItems.Add(, , rs(0).Value)
                  list_item.SubItems(1) = rs(1).Value
                  rs.MoveNext
            Wend
            rs.Close
            If lv_lista.ListItems.Count > 0 Then
               Me.frm_lista.Visible = True
               Me.lv_lista.SetFocus
            Else
               MsgBox "No existen causas de devolución", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El movimiento ya no puede ser modificado", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub
