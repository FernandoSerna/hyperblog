VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmdescuentos_pago_correcto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Descuentos por pronto pago"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2910
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5460
      Picture         =   "frmdescuentos_pago_correcto.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1470
      Picture         =   "frmdescuentos_pago_correcto.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      Picture         =   "frmdescuentos_pago_correcto.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   810
      Picture         =   "frmdescuentos_pago_correcto.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      Picture         =   "frmdescuentos_pago_correcto.frx":0910
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmdescuentos_pago_correcto.frx":0A12
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   165
      TabIndex        =   11
      Top             =   270
      Width           =   5655
   End
   Begin VB.Frame Frame1 
      Caption         =   " Catálogos "
      Height          =   1455
      Left            =   150
      TabIndex        =   7
      Top             =   375
      Width           =   5655
      Begin VB.TextBox txt_descuento 
         Height          =   315
         Left            =   1305
         TabIndex        =   23
         Top             =   975
         Width           =   1110
      End
      Begin VB.TextBox txt_limite_superior 
         Height          =   315
         Left            =   4065
         TabIndex        =   22
         Top             =   630
         Width           =   1110
      End
      Begin VB.TextBox txt_limite_inferior 
         Height          =   315
         Left            =   1305
         TabIndex        =   21
         Top             =   630
         Width           =   1110
      End
      Begin VB.TextBox txt_empresa 
         Enabled         =   0   'False
         Height          =   315
         Left            =   915
         TabIndex        =   9
         Top             =   270
         Width           =   1110
      End
      Begin VB.TextBox txt_nombre_empresa 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2055
         TabIndex        =   8
         Top             =   270
         Width           =   3495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Descuento:"
         Height          =   195
         Left            =   165
         TabIndex        =   20
         Top             =   1035
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Limite Superior:"
         Height          =   195
         Left            =   2850
         TabIndex        =   19
         Top             =   690
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Limite Inferior:"
         Height          =   195
         Left            =   165
         TabIndex        =   18
         Top             =   690
         Width           =   975
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Empresa:"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   10
         Top             =   330
         Width           =   660
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   150
      TabIndex        =   12
      Top             =   1845
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   1890
         TabIndex        =   13
         Top             =   150
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3795
         TabIndex        =   14
         Top             =   150
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ir al primero"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Un Registro Atras"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Un registro adelante"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ir al ultimo"
               ImageIndex      =   10
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Busqueda de catalogo:"
         Height          =   195
         Left            =   195
         TabIndex        =   15
         Top             =   195
         Width           =   1650
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4770
      Left            =   150
      TabIndex        =   16
      Top             =   2400
      Width           =   5655
      Begin MSComctlLib.ListView lv_descuentos 
         Height          =   4560
         Left            =   45
         TabIndex        =   17
         Top             =   150
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   8043
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Inferior"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Superior"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descuento"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "fecha inicio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "fecha fin"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "AGRUPADOR"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   45
         Top             =   2505
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
               Picture         =   "frmdescuentos_pago_correcto.frx":0B14
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdescuentos_pago_correcto.frx":13EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdescuentos_pago_correcto.frx":1CC8
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdescuentos_pago_correcto.frx":2264
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdescuentos_pago_correcto.frx":2B40
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdescuentos_pago_correcto.frx":341A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdescuentos_pago_correcto.frx":3CF4
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdescuentos_pago_correcto.frx":3E06
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdescuentos_pago_correcto.frx":3F18
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdescuentos_pago_correcto.frx":402A
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdescuentos_pago_correcto.frx":413C
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmdescuentos_pago_correcto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_deshacer_Click()
   Me.txt_descuento.Enabled = False
   Me.txt_limite_inferior.Enabled = False
   Me.txt_limite_superior.Enabled = False
End Sub

Private Sub cmd_eliminar_Click()
   Dim var_si As Integer
   var_si = MsgBox("¿Desea eliminar el registro?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      rs.Open "delete from TB_DESCUENTOS_PRONTO_PAGO where vcha_emp_empresa_id = '" + var_empresa + "' and INTE_DPG_LIMITE_INFERIOR = " + txt_limite_inferior + " and INTE_DPG_LIMITE_SUPERIOR = " + txt_limite_superior, cnn, adOpenDynamic, adLockOptimistic
      lv_descuentos.ListItems.Remove (lv_descuentos.selectedItem.Index)
      Call pro_limpiatextos(Me)
      If lv_descuentos.ListItems.Count > 0 Then
         lv_descuentos.SetFocus
      End If
   End If
   Me.txt_descuento.Enabled = False
   Me.txt_limite_inferior.Enabled = False
   Me.txt_limite_superior.Enabled = False
End Sub

Private Sub cmd_guardar_Click()
   If IsNumeric(Me.txt_descuento) Then
      If IsNumeric(Me.txt_limite_inferior) Then
         If IsNumeric(Me.txt_limite_superior) Then
            rs.Open "SELECT * FROM TB_DESCUENTOS_PRONTO_PAGO where vcha_emp_empresa_id = '" + var_empresa + "' and INTE_DPG_LIMITE_INFERIOR = " + txt_limite_inferior + " and INTE_DPG_LIMITE_SUPERIOR = " + txt_limite_superior, cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               MsgBox "Ya existe un descuento con los limites indicados", vbOKOnly, "ATENCION"
            Else
               var_si = MsgBox("¿Deseas insertar el registro?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  rsaux.Open "INSERT INTO TB_DESCUENTOS_PRONTO_PAGO (VCHA_EMP_EMPRESA_ID,INTE_DPG_LIMITE_INFERIOR,INTE_DPG_LIMITE_SUPERIOR, FLOA_DPG_DESCUENTO) VALUES ('" + var_empresa + "', " + Me.txt_limite_inferior + ", " + Me.txt_limite_superior + ", " + Me.txt_descuento + ")", cnn, adOpenDynamic, adLockOptimistic
                  Dim list_item As ListItem
                  Set list_item = lv_descuentos_catalogos.ListItems.Add(, , Me.txt_limite_inferior)
                  list_item.SubItems(1) = Me.txt_limite_superior
                  list_item.SubItems(2) = Me.txt_descuento
                  Me.txt_descuento.Enabled = False
                  Me.txt_limite_inferior.Enabled = False
                  Me.txt_limite_superior.Enabled = False
               End If
            End If
            rs.Close
         Else
            MsgBox "Limite superior incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Limite inferior incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Descuento incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
x = 1
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_descuento = ""
   Me.txt_limite_inferior = ""
   Me.txt_limite_superior = ""
   Me.txt_limite_inferior.Enabled = True
   Me.txt_limite_superior.Enabled = True
   Me.txt_descuento.Enabled = True
   Me.cmd_guardar.Enabled = True
   Me.cmd_deshacer.Enabled = False
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
   If Shift = 4 And KeyCode = 68 Then
      cmd_deshacer_Click
   End If
   If Shift = 4 And KeyCode = 69 Then
      cmd_eliminar_Click
   End If
   If Shift = 4 And KeyCode = 73 Then
      cmd_imprimir_Click
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 0
   Left = 2900
   var_modifica_registro_descuentos_pago_correcto = True
   rs.Open "select * from tb_empresas where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      txt_empresa = rs!vcha_emp_empresa_id
      Me.txt_nombre_empresa = rs!vcha_emp_nombre
   End If
   rs.Close
   rs.Open "select * from TB_DESCUENTOS_PRONTO_PAGO where vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      Me.txt_limite_inferior = rs!INTE_DPG_LIMITE_INFERIOR
      Me.txt_limite_superior = rs!INTE_DPG_LIMITE_SUPERIOR
      Me.txt_descuento = rs!FLOA_DPG_DESCUENTO
      cmd_eliminar.Enabled = True
   Else
      cmd_eliminar.Enabled = False
   End If
   cmd_guardar.Enabled = False
   cmd_deshacer.Enabled = False
   Me.txt_limite_inferior.Enabled = False
   Me.txt_descuento.Enabled = False
   Me.txt_limite_superior.Enabled = False
   Dim list_item As ListItem
   numero_items_cajas = 0
   While Not rs.EOF
      Set list_item = lv_descuentos.ListItems.Add(, , rs!INTE_DPG_LIMITE_INFERIOR)
      list_item.SubItems(1) = IIf(IsNull(rs!INTE_DPG_LIMITE_SUPERIOR), "", rs!INTE_DPG_LIMITE_SUPERIOR)
      list_item.SubItems(2) = IIf(IsNull(rs!FLOA_DPG_DESCUENTO), "", rs!FLOA_DPG_DESCUENTO)
      rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
    var_swpassword = False
    Call activa_forma(var_activa_forma_descuentos_pago_correcto)
End Sub

Private Sub lv_descuentos_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Me.txt_limite_inferior = lv_descuentos.selectedItem
   Me.txt_limite_superior = lv_descuentos.selectedItem.SubItems(1)
   Me.txt_descuento = lv_descuentos.selectedItem.SubItems(2)
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      Me.lv_descuentos.SetFocus
      Call pro_avanzar(Me, lv_descuentos, Button)
      lv_descuentos.selectedItem.EnsureVisible
      Me.txt_limite_inferior = lv_descuentos.selectedItem
      Me.txt_limite_superior = lv_descuentos.selectedItem.SubItems(1)
      Me.txt_descuento = lv_descuentos.selectedItem.SubItems(2)
   End If
   If Button.Index = 1 Then
      lv_descuentos.ListItems(1).Selected = True
      Me.txt_limite_inferior = lv_descuentos.selectedItem
      Me.txt_limite_superior = lv_descuentos.selectedItem.SubItems(1)
      Me.txt_descuento = lv_descuentos.selectedItem.SubItems(2)
      lv_descuentos.selectedItem.EnsureVisible
   End If
   If Button.Index = 4 Then
      numero_items_diseños = lv_descuentos.ListItems.Count
      lv_descuentos.ListItems(numero_items_diseños).Selected = True
      lv_descuentos.selectedItem.EnsureVisible
      Me.txt_limite_inferior = lv_descuentos.selectedItem
      Me.txt_limite_superior = lv_descuentos.selectedItem.SubItems(1)
      Me.txt_descuento = lv_descuentos.selectedItem.SubItems(2)
   End If
err0:
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_descuentos, txt_buscar, False)
      txt_buscar = ""
   End If
End Sub

Private Sub txt_descuento_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
End Sub

Private Sub txt_empresa_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_limite_inferior_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_limite_inferior_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
End Sub

Private Sub txt_limite_superior_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
End Sub

Private Sub txt_nombre_empresa_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub
