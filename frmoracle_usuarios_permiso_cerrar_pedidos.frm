VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_usuarios_permiso_cerrar_pedidos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Supervisores para autorizacion de cerrado de pedidos incompletos"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_seleccionusuarios 
      Height          =   2235
      Left            =   1275
      TabIndex        =   0
      Top             =   375
      Width           =   7110
      Begin MSComctlLib.ListView lv_seleccion_usuarios 
         Height          =   1785
         Left            =   30
         TabIndex        =   1
         Top             =   360
         Width           =   7020
         _ExtentX        =   12383
         _ExtentY        =   3149
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
            Object.Width           =   8467
         EndProperty
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000002&
         Caption         =   " Seleccione el usuario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   30
         TabIndex        =   2
         Top             =   120
         Width           =   7020
      End
   End
   Begin VB.Frame Frame5 
      Height          =   6015
      Left            =   60
      TabIndex        =   3
      Top             =   0
      Width           =   9195
      Begin VB.CommandButton cmd_salir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8805
         Picture         =   "frmoracle_usuarios_permiso_cerrar_pedidos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Salir"
         Top             =   375
         Width           =   330
      End
      Begin VB.CommandButton cmd_eliminar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   705
         Picture         =   "frmoracle_usuarios_permiso_cerrar_pedidos.frx":063A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Eliminar Alt + E"
         Top             =   375
         Width           =   330
      End
      Begin VB.CommandButton cmd_guardar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   375
         Picture         =   "frmoracle_usuarios_permiso_cerrar_pedidos.frx":073C
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Guardar Alt + G"
         Top             =   375
         Width           =   330
      End
      Begin VB.CommandButton cmd_nuevo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   45
         Picture         =   "frmoracle_usuarios_permiso_cerrar_pedidos.frx":083E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Nuevo Alt + N"
         Top             =   375
         Width           =   330
      End
      Begin VB.Frame Frame1 
         Caption         =   " Usuarios "
         Height          =   1020
         Index           =   0
         Left            =   45
         TabIndex        =   7
         Top             =   750
         Width           =   9090
         Begin VB.TextBox txt_nombre_usuario 
            Height          =   315
            Left            =   1230
            MaxLength       =   50
            TabIndex        =   9
            Top             =   600
            Width           =   7770
         End
         Begin VB.TextBox txt_usuario 
            Height          =   315
            Left            =   1230
            MaxLength       =   50
            TabIndex        =   8
            Top             =   255
            Width           =   1365
         End
         Begin VB.Label lab_paises 
            AutoSize        =   -1  'True
            Caption         =   "Nombre:"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   11
            Top             =   660
            Width           =   600
         End
         Begin VB.Label lab_paises 
            AutoSize        =   -1  'True
            Caption         =   "Clave:"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   10
            Top             =   315
            Width           =   450
         End
      End
      Begin VB.Frame Frame4 
         Height          =   120
         Left            =   45
         TabIndex        =   6
         Top             =   630
         Width           =   9090
      End
      Begin VB.Frame Frame3 
         Height          =   4230
         Left            =   45
         TabIndex        =   4
         Top             =   1740
         Width           =   9090
         Begin MSComctlLib.ListView lv_usuarios 
            Height          =   3990
            Left            =   30
            TabIndex        =   5
            Top             =   165
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   7038
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Clave"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nombre"
               Object.Width           =   12876
            EndProperty
         End
      End
      Begin MSComctlLib.ImageList icono_encabezado 
         Left            =   3810
         Top             =   105
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   12
         ImageHeight     =   12
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmoracle_usuarios_permiso_cerrar_pedidos.frx":0940
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmoracle_usuarios_permiso_cerrar_pedidos.frx":121A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   3195
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   11
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmoracle_usuarios_permiso_cerrar_pedidos.frx":1AF4
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmoracle_usuarios_permiso_cerrar_pedidos.frx":23CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmoracle_usuarios_permiso_cerrar_pedidos.frx":2CA8
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmoracle_usuarios_permiso_cerrar_pedidos.frx":3244
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmoracle_usuarios_permiso_cerrar_pedidos.frx":3B20
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmoracle_usuarios_permiso_cerrar_pedidos.frx":43FA
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmoracle_usuarios_permiso_cerrar_pedidos.frx":4CD4
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmoracle_usuarios_permiso_cerrar_pedidos.frx":4DE6
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmoracle_usuarios_permiso_cerrar_pedidos.frx":4EF8
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmoracle_usuarios_permiso_cerrar_pedidos.frx":500A
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmoracle_usuarios_permiso_cerrar_pedidos.frx":511C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000002&
         Caption         =   " Usuarios "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   30
         TabIndex        =   16
         Top             =   120
         Width           =   9120
      End
   End
End
Attribute VB_Name = "frmoracle_usuarios_permiso_cerrar_pedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_eliminar_Click()
   var_si = MsgBox("¿Desea eliminar el usuario?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      rs.Open "DELETE FROM TB_ORACLE_USUARIOS_PERMISO_CERRAR_PEDIDOS WHERE VCHA_USU_USUARIO_ID = '" + Me.txt_usuario + "'", cnn, adOpenDynamic, adLockOptimistic
      lv_usuarios.ListItems.Remove (lv_usuarios.selectedItem.Index)

   End If
End Sub

Private Sub cmd_guardar_Click()
   If Me.txt_usuario <> "" Then
      rs.Open "select * from TB_ORACLE_USUARIOS_PERMISO_CERRAR_PEDIDOS where vcha_usu_usuario_id = '" + Me.txt_usuario + "'", cnn, adOpenDynamic, adLockOptimistic
      If rs.EOF Then
         rsaux.Open "insert into TB_ORACLE_USUARIOS_PERMISO_CERRAR_PEDIDOS (vcha_usu_usuario_id) values ('" + Me.txt_usuario + "')", cnn, adOpenDynamic, adLockOptimistic
         Set list_item = lv_usuarios.ListItems.Add(, , Me.txt_usuario)
         list_item.SubItems(1) = Me.txt_nombre_usuario
      End If
      rs.Close
   End If
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_usuario = ""
   Me.txt_nombre_usuario = ""
   Me.txt_usuario.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 800
   Left = 1000
   Me.frm_seleccionusuarios.Visible = False
   rs.Open "select a.vcha_usu_usuario_id, ISNULL(vcha_usu_nombre,'')+' '+ISNULL(vcha_usu_apellidos,'') as nombre from TB_ORACLE_USUARIOS_PERMISO_CERRAR_PEDIDOS a, tb_usuarios b where a.VCHA_USU_USUARIO_ID = b.VCHA_USU_USUARIO_ID ", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set list_item = Me.lv_usuarios.ListItems.Add(, , rs!vcha_usu_usuario_id)
         list_item.SubItems(1) = rs!NOMBRE
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_seleccion_usuarios_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_usuario = Me.lv_seleccion_usuarios.selectedItem
      Me.txt_nombre_usuario = Me.lv_seleccion_usuarios.selectedItem.SubItems(1)
      Me.txt_usuario.SetFocus
      Me.frm_seleccionusuarios.Visible = False
   End If
   If KeyAscii = 27 Then
      Me.frm_seleccionusuarios.Visible = False
   End If
End Sub

Private Sub lv_seleccion_usuarios_LostFocus()
   Me.frm_seleccionusuarios.Visible = False
End Sub

Private Sub lv_usuarios_Click()
   Me.txt_usuario = Me.lv_usuarios.selectedItem
   Me.txt_nombre_usuario = Me.lv_usuarios.selectedItem.SubItems(1)
End Sub

Private Sub lv_usuarios_GotFocus()
   If Me.lv_usuarios.ListItems.Count > 0 Then
      Me.txt_usuario = Me.lv_usuarios.selectedItem
      Me.txt_nombre_usuario = Me.lv_usuarios.selectedItem.SubItems(1)
   End If
End Sub

Private Sub txt_nombre_usuario_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   Else
      KeyAscii = 27
   End If
End Sub

Private Sub txt_usuario_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Me.lv_seleccion_usuarios.ListItems.Clear
      rs.Open "select vcha_usu_usuario_id, isnull(vcha_usu_nombre,'')+' '+isnull(vcha_usu_apellidos,'') as nombre from tb_usuarios order by isnull(vcha_usu_nombre,'')+' '+isnull(vcha_usu_apellidos,'')", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_seleccion_usuarios.ListItems.Add(, , rs!vcha_usu_usuario_id)
            list_item.SubItems(1) = rs!NOMBRE
            rs.MoveNext
      Wend
      rs.Close
      Me.frm_seleccionusuarios.Visible = True
      Me.lv_seleccion_usuarios.SetFocus
   End If
End Sub

Private Sub txt_usuario_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_nombre_usuario.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub
