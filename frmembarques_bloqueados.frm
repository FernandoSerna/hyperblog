VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmembarques_bloqueados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Embarques Bloqueados"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   7710
   Begin VB.Frame frm_bloqueos 
      Caption         =   "Embarque Bloqueado "
      Height          =   1020
      Left            =   75
      TabIndex        =   5
      Top             =   435
      Width           =   7560
      Begin VB.TextBox txt_embarque 
         Height          =   315
         Left            =   1515
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   2025
      End
      Begin VB.TextBox txt_usuario 
         Height          =   315
         Left            =   1515
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   570
         Width           =   5925
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Embarque:"
         Height          =   195
         Left            =   90
         TabIndex        =   9
         Top             =   300
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Bloqueado por:"
         Height          =   195
         Left            =   90
         TabIndex        =   8
         Top             =   630
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frmembarques_bloqueados.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancelar Esc"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmembarques_bloqueados.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   5730
      Left            =   75
      TabIndex        =   0
      Top             =   1455
      Width           =   7560
      Begin MSComctlLib.ListView lv_bloqueos 
         Height          =   5535
         Left            =   45
         TabIndex        =   1
         Top             =   135
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   9763
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Embarque"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Usuario"
            Object.Width           =   9525
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Clave usuario"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   75
      TabIndex        =   4
      Top             =   255
      Width           =   7560
   End
End
Attribute VB_Name = "frmembarques_bloqueados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_Click()
   If Trim(txt_embarque) <> "" Then
      var_si = MsgBox("¿Deseas desbloquear el embaruqe?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         rs.Open "UPDATE TB_ENCABEZADO_EMBARQUES SET INTE_EMB_BLOQUEADO = 0, VCHA_EMB_BLOQUEADO_POR = '' WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_EMB_EMBARQUE = " + txt_embarque, cnn, adOpenDynamic, adLockOptimistic
         lv_bloqueos.ListItems.Remove (lv_bloqueos.selectedItem.Index)
         If lv_bloqueos.ListItems.Count > 0 Then
            Me.txt_embarque = ""
            Me.txt_usuario = ""
            lv_bloqueos.SetFocus
         Else
            Me.txt_embarque = ""
            Me.txt_usuario = ""
         End If
      End If
   Else
      MsgBox "No se a seleccionado un embarque", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_cancelar_Click()
   Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 65 Then
      cmd_aceptar_Click
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 2000
   Cadena = "SELECT     dbo.TB_ENCABEZADO_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_USUARIOS.VCHA_USU_NOMBRE, dbo.TB_USUARIOS.VCHA_USU_APELLIDOS, dbo.TB_ENCABEZADO_EMBARQUES.INTE_EMB_BLOQUEADO fROM dbo.TB_ENCABEZADO_EMBARQUES INNER JOIN dbo.TB_USUARIOS ON dbo.TB_ENCABEZADO_EMBARQUES.VCHA_EMB_BLOQUEADO_POR = dbo.TB_USUARIOS.VCHA_USU_USUARIO_ID wHERE (dbo.TB_ENCABEZADO_EMBARQUES.INTE_EMB_BLOQUEADO = 1 and dbo.TB_ENCABEZADO_EMBARQUES.vcha_emp_empresa_id = '" + var_empresa + "')"
   
   rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set list_item = lv_bloqueos.ListItems.Add(, , rs!inte_emb_embarque)
         list_item.SubItems(1) = IIf(IsNull(rs!VCHA_USU_NOMBRE), "", rs!VCHA_USU_NOMBRE) + " " + IIf(IsNull(rs!VCHA_USU_APELLIDOS), "", rs!VCHA_USU_APELLIDOS)
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_embarques_bloqueados)
End Sub

Private Sub lv_bloqueos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call pro_ordena_listas(Me.lv_bloqueos, ColumnHeader)
End Sub

Private Sub lv_bloqueos_GotFocus()
   If lv_bloqueos.ListItems.Count > 0 Then
      txt_embarque = lv_bloqueos.selectedItem
      txt_usuario = lv_bloqueos.selectedItem.SubItems(1)
   End If
End Sub

Private Sub lv_bloqueos_ItemClick(ByVal Item As MSComctlLib.ListItem)
   If lv_bloqueos.ListItems.Count > 0 Then
      txt_embarque = lv_bloqueos.selectedItem
      txt_usuario = lv_bloqueos.selectedItem.SubItems(1)
   End If
End Sub

Private Sub lv_bloqueos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_aceptar.SetFocus
   End If
End Sub
