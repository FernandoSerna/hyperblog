VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsubrutas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Subrutas"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   5850
      Left            =   150
      TabIndex        =   10
      Top             =   1440
      Width           =   5655
      Begin MSComctlLib.ImageList icono_encabezado 
         Left            =   0
         Top             =   0
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
               Picture         =   "frmsubrutas.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmsubrutas.frx":08DA
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lv_subrutas 
         Height          =   5595
         Left            =   45
         TabIndex        =   11
         Top             =   165
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   9869
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
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   6879
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "  Subrutas "
      Height          =   1020
      Left            =   150
      TabIndex        =   7
      Top             =   405
      Width           =   5655
      Begin VB.TextBox txt_subruta 
         Height          =   315
         Left            =   1155
         MaxLength       =   50
         TabIndex        =   4
         Top             =   225
         Width           =   960
      End
      Begin VB.TextBox txt_nombre_subruta 
         Height          =   315
         Left            =   1155
         MaxLength       =   50
         TabIndex        =   5
         Top             =   585
         Width           =   4425
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   9
         Top             =   270
         Width           =   450
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Index           =   2
         Left            =   195
         TabIndex        =   8
         Top             =   630
         Width           =   885
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2775
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   6
      Top             =   90
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5460
      Picture         =   "frmsubrutas.frx":11B4
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   810
      Picture         =   "frmsubrutas.frx":17EE
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      Picture         =   "frmsubrutas.frx":18F0
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Guardar Alt + G"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmsubrutas.frx":19F2
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubrutas.frx":1AF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubrutas.frx":23CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubrutas.frx":2CA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubrutas.frx":3244
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubrutas.frx":3B20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubrutas.frx":43FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubrutas.frx":4CD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubrutas.frx":4DE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubrutas.frx":4EF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubrutas.frx":500A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   12
      Top             =   285
      Width           =   5655
   End
End
Attribute VB_Name = "frmsubrutas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_eliminar_Click()
   rs.Open "SELECT * from TB_CLIENTES WHERE VCHA_SRU_SUBRUTA_ID = '" + Me.txt_subruta + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      MsgBox "No se puede eliminar la subruta ya que hay clientes que la tienen asignada", vbOKOnly, "ATENCION"
   Else
      var_si = MsgBox("¿Desea eliminar la subruta?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         var_si = MsgBox("Confirmar la eliminación de la subruta", vbYesNo, "ATENCION")
         If var_si = 6 Then
            rsaux.Open "DELETE FROM TB_SUBRUTAS WHERE VCHA_SRU_SUBRUTA_ID = '" + Me.txt_subruta + "'", cnn, adOpenDynamic, adLockOptimistic
            Dim list_item As ListItem
            lv_subrutas.ListItems.Clear
            Me.txt_subruta = ""
            Me.txt_nombre_subruta = ""
            rsaux.Open "select * from tb_subrutas", cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux.EOF
                  Set list_item = lv_subrutas.ListItems.Add(, , rsaux(0).Value)
                  list_item.SubItems(1) = IIf(IsNull(rsaux(1).Value), "", rsaux(1).Value)
                  rsaux.MoveNext
            Wend
            rsaux.Close
            If Me.lv_subrutas.ListItems.Count > 0 Then
               Me.lv_subrutas.ListItems.Item(1).Selected = True
            End If
            MsgBox "Se a eliminado el registro", vbOKOnly, "ATENCION"
         
         End If
      End If
   End If
   rs.Close
End Sub

Private Sub cmd_guardar_Click()
   If Me.txt_subruta <> "" Then
      rs.Open "select * from tb_subrutas where vcha_sru_subruta_id = '" + Me.txt_subruta + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_si = MsgBox("Desea efectuar los cambios a la subruta", vbYesNo, "ATENCION")
         If var_si = 6 Then
            rsaux.Open "update tb_subrutas set vcha_sru_nombre = '" + Me.txt_nombre_subruta + "' where vcha_sru_subruta_id = '" + Me.txt_subruta + "'", cnn, adOpenDynamic, adLockOptimistic
            MsgBox "Se a actualizado el registro", vbOKOnly, "ATENCION"
         End If
      Else
         'MsgBox "INSERT INTO TB_SUBRUTAS (TB_SRU_SUBRUTA_ID,TB_SRU_NOMBRE) VALUES ('" + Me.txt_subruta + "','" + Me.txt_nombre_subruta + "')"
         rsaux.Open "INSERT INTO TB_SUBRUTAS (VCHA_SRU_SUBRUTA_ID,vcha_SRU_NOMBRE) VALUES ('" + Me.txt_subruta + "','" + Me.txt_nombre_subruta + "')", cnn, adOpenDynamic, adLockOptimistic
      End If
      rs.Close
      Dim list_item As ListItem
      lv_subrutas.ListItems.Clear
      rs.Open "select * from tb_subrutas", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_subrutas.ListItems.Add(, , rs(0).Value)
            list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
            rs.MoveNext
      Wend
      rs.Close
   Else
      MsgBox "No se a seleccionado una subruta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_subruta = ""
   Me.txt_nombre_subruta = ""
   Me.txt_subruta.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 78 Then
      cmd_nuevo_Click
   End If
   If Shift = 4 And KeyCode = 71 Then
      cmd_guardar_Click
   End If
   'If Shift = 4 And KeyCode = 68 Then
   '   cmd_deshacer_Click
   'End If
   If Shift = 4 And KeyCode = 69 Then
      cmd_eliminar_Click
   End If
   'If Shift = 4 And KeyCode = 73 Then
   '   cmd_imprimir_Click
   'End If
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 2900
   Dim list_item As ListItem
   rs.Open "select * from tb_subrutas", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set list_item = lv_subrutas.ListItems.Add(, , rs(0).Value)
         list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub lv_subrutas_Click()
   If Me.lv_subrutas.ListItems.Count > 0 Then
      Me.txt_subruta = Me.lv_subrutas.selectedItem
      Me.txt_nombre_subruta = Me.lv_subrutas.selectedItem.SubItems(1)
   End If
End Sub

Private Sub lv_subrutas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_subrutas, ColumnHeader)
End Sub

Private Sub lv_subrutas_ItemClick(ByVal Item As MSComctlLib.ListItem)
   If Me.lv_subrutas.ListItems.Count > 0 Then
      Me.txt_subruta = Me.lv_subrutas.selectedItem
      Me.txt_nombre_subruta = Me.lv_subrutas.selectedItem.SubItems(1)
   End If
End Sub

Private Sub txt_nombre_subruta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub

Private Sub txt_subruta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_nombre_subruta.SetFocus
   End If
End Sub
