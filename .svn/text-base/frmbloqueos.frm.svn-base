VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmbloqueos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bloqueos"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7755
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   7755
   Begin VB.Frame Frame3 
      Height          =   5100
      Left            =   90
      TabIndex        =   9
      Top             =   2100
      Width           =   7560
      Begin MSComctlLib.ListView lv_bloqueos 
         Height          =   4890
         Left            =   45
         TabIndex        =   1
         Top             =   150
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   8625
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Modulo"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Usuario"
            Object.Width           =   6879
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Máquina"
            Object.Width           =   2558
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fecha"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Clave usuario"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.CommandButton cmd_aceptar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmbloqueos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   420
      Picture         =   "frmbloqueos.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancelar Esc"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   90
      TabIndex        =   8
      Top             =   285
      Width           =   7560
   End
   Begin VB.Frame frm_bloqueos 
      Caption         =   "Modulo Bloqueado "
      Height          =   1635
      Left            =   90
      TabIndex        =   0
      Top             =   465
      Width           =   7560
      Begin VB.TextBox txt_fecha 
         Height          =   315
         Left            =   1515
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1230
         Width           =   2460
      End
      Begin VB.TextBox txt_maquina 
         Height          =   315
         Left            =   1515
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   900
         Width           =   2460
      End
      Begin VB.TextBox txt_usuario 
         Height          =   315
         Left            =   1515
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   570
         Width           =   5925
      End
      Begin VB.TextBox txt_modulo 
         Height          =   315
         Left            =   1515
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   2025
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha del Bloqueo:"
         Height          =   195
         Left            =   90
         TabIndex        =   13
         Top             =   1290
         Width           =   1380
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Máquina:"
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   960
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Bloqueado por:"
         Height          =   195
         Left            =   90
         TabIndex        =   11
         Top             =   630
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Modulo:"
         Height          =   195
         Left            =   90
         TabIndex        =   10
         Top             =   300
         Width           =   570
      End
   End
End
Attribute VB_Name = "frmbloqueos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_Click()
   Dim var_si As Integer
   If txt_modulo <> "" Then
      var_si = MsgBox("¿Deseas eliminar el bloqueo?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         var_si = MsgBox("Confirmar la eliminación del bloqueo", vbYesNo, "ATENCION")
         If var_si = 6 Then
            rs.Open "delete from tb_bloqueos where vcha_blo_bloqueado_por = '" + Me.txt_modulo + "' and vcha_usu_usuario_id = '" + Me.lv_bloqueos.selectedItem.SubItems(4) + "' and vcha_blo_maquina ='" + txt_maquina + "'", cnn, adOpenDynamic, adLockOptimistic
            lv_bloqueos.ListItems.Remove (lv_bloqueos.selectedItem.Index)
            If lv_bloqueos.ListItems.Count > 0 Then
               lv_bloqueos.SetFocus
            Else
               txt_modulo = ""
               txt_usuario = ""
               txt_maquina = ""
               txt_fecha = ""
            End If
         End If
      End If
   Else
      MsgBox "No se a seleccionado un modulo para desbloquear", vbOKOnly, "ATENCION"
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
   rs.Open "select * from VW_BLOQUEOS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
   Dim list_item As ListItem
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = lv_bloqueos.ListItems.Add(, , rs!VCHA_BLO_BLOQUEADO_POR)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_USU_NOMBRE), "", rs!VCHA_USU_NOMBRE) + " " + IIf(IsNull(rs!VCHA_USU_APELLIDOS), "", rs!VCHA_USU_APELLIDOS)
            list_item.SubItems(2) = IIf(IsNull(rs!vcha_blo_maquina), "", rs!vcha_blo_maquina)
            list_item.SubItems(3) = IIf(IsNull(rs!dtim_blo_fecha), "", rs!dtim_blo_fecha)
            list_item.SubItems(4) = IIf(IsNull(rs!vcha_usu_usuario_id), "", rs!vcha_usu_usuario_id)
            rs.MoveNext:
      Wend
      rs.MoveFirst
      txt_modulo = rs!VCHA_BLO_BLOQUEADO_POR
      txt_usuario = IIf(IsNull(rs!VCHA_USU_NOMBRE), "", rs!VCHA_USU_NOMBRE) + " " + IIf(IsNull(rs!VCHA_USU_APELLIDOS), "", rs!VCHA_USU_APELLIDOS)
      txt_maquina = IIf(IsNull(rs!vcha_blo_maquina), "", rs!vcha_blo_maquina)
      txt_fecha = IIf(IsNull(rs!dtim_blo_fecha), "", rs!dtim_blo_fecha)
   End If
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_bloqueos)
End Sub

Private Sub lv_bloqueos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_bloqueos, ColumnHeader)
End Sub

Private Sub lv_bloqueos_GotFocus()
   If lv_bloqueos.ListItems.Count > 0 Then
      txt_modulo = lv_bloqueos.selectedItem
      txt_usuario = lv_bloqueos.selectedItem.SubItems(1)
      txt_maquina = lv_bloqueos.selectedItem.SubItems(2)
      txt_fecha = lv_bloqueos.selectedItem.SubItems(3)
   End If
End Sub

Private Sub lv_bloqueos_ItemClick(ByVal Item As MSComctlLib.ListItem)
   If lv_bloqueos.ListItems.Count > 0 Then
      txt_modulo = lv_bloqueos.selectedItem
      txt_usuario = lv_bloqueos.selectedItem.SubItems(1)
      txt_maquina = lv_bloqueos.selectedItem.SubItems(2)
      txt_fecha = lv_bloqueos.selectedItem.SubItems(3)
   End If
End Sub
