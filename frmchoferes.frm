VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmchoferes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choferes"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   4575
      Left            =   90
      TabIndex        =   13
      Top             =   2895
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
               Picture         =   "frmchoferes.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmchoferes.frx":08DA
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lv_transportes 
         Height          =   4245
         Left            =   45
         TabIndex        =   14
         Top             =   150
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   7488
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Licencia"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Choferes "
      Height          =   1860
      Left            =   90
      TabIndex        =   8
      Top             =   480
      Width           =   5655
      Begin VB.TextBox txt_rfc 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   18
         Top             =   1320
         Width           =   2355
      End
      Begin VB.TextBox txt_licencia 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   16
         Top             =   960
         Width           =   2355
      End
      Begin VB.TextBox txt_numero 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   10
         Top             =   255
         Width           =   900
      End
      Begin VB.TextBox txt_nombre 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   9
         Top             =   585
         Width           =   4275
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "RFC:"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   19
         Top             =   1380
         Width           =   360
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Licencia:"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   17
         Top             =   1020
         Width           =   645
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   12
         Top             =   315
         Width           =   450
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   11
         Top             =   645
         Width           =   600
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   90
      TabIndex        =   4
      Top             =   2340
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   2055
         TabIndex        =   5
         Top             =   150
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3870
         TabIndex        =   6
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
               Object.ToolTipText     =   "Nuevo Registro"
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
         Caption         =   "Busqueda de choferes:"
         Height          =   195
         Left            =   195
         TabIndex        =   7
         Top             =   195
         Width           =   1650
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2715
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   3
      Top             =   105
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5400
      Picture         =   "frmchoferes.frx":11B4
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   420
      Picture         =   "frmchoferes.frx":17EE
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Guardar Alt + G"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmchoferes.frx":18F0
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   45
      Width           =   330
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   -60
      Top             =   15
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
            Picture         =   "frmchoferes.frx":19F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmchoferes.frx":22CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmchoferes.frx":2BA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmchoferes.frx":3142
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmchoferes.frx":3A1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmchoferes.frx":42F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmchoferes.frx":4BD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmchoferes.frx":4CE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmchoferes.frx":4DF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmchoferes.frx":4F08
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   90
      TabIndex        =   15
      Top             =   300
      Width           =   5655
   End
End
Attribute VB_Name = "frmchoferes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim numero_items_transportes As Integer






Private Sub cmd_deshacer_Click()
   Call pro_textos
End Sub

Private Sub cmd_eliminar_Click()
   var_opcion_seguridad = 2
   var_acepta_seguridad = 1
   If var_global_permiso3 = 1 Then
      var_acepta_seguridad = 2
      If var_global_permiso4 = 1 Then
         frmpasswords2.Show 1
      Else
         frmpasswords.Show 1
      End If
   End If
   If var_acepta_seguridad = 1 Then
      Call pro_elimina_transportes
      Call pro_llena_listview1
      rs.Open "select * from TB_CHOFERES", cnn, adOpenDynamic, adLockOptimistic
      If rs.BOF Then
         cmd_guardar.Enabled = False
         cmd_deshacer.Enabled = False
         cmd_eliminar.Enabled = False
      Else
         cmd_guardar.Enabled = True
         cmd_deshacer.Enabled = True
         cmd_eliminar.Enabled = True
      End If
      rs.Close
   End If
End Sub

Private Sub cmd_guardar_Click()
Dim var_posible As Boolean
   var_posible = True
   rs.Open "select * from xxvia_tb_choferes where id_chofer = '" + Me.txt_numero + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      rsaux1.Open "update xxvia_Tb_choferes set nombre = '" + Me.txt_nombre + "', rfc = '" + Me.txt_RFC + "', licencia = '" + Me.txt_licencia + "' where id_chofer = '" + Me.txt_numero + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   Else
      rsaux1.Open "insert into xxvia_tb_choferes (id_chofer, nombre, rfc, licencia) values ('" + Me.txt_numero + "','" + Me.txt_nombre + "','" + Me.txt_RFC + "','" + Me.txt_licencia + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
   End If
   rs.Close
   Me.lv_transportes.ListItems.Clear
   Call pro_llena_listview1
End Sub

Private Sub cmd_imprimir_Click()

End Sub

Private Sub cmd_nuevo_Click()
        Call pro_limpiatextos(Me)
        txt_numero.Enabled = True
        cmd_guardar.Enabled = True
        rs.Open "select max(to_number(id_chofer)) from xxvia_tb_choferes", cnnoracle_4, adOpenDynamic, adLockOptimistic
        Me.txt_numero = rs(0).Value + 1
        rs.Close
        txt_numero.SetFocus: var_modifica_registro_transporte = False
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
   var_modifica_registro_transporte = True
   lv_transportes.SmallIcons = ImageList1
   Call pro_encabezadosView(Me, lv_transportes, False)
   Call pro_llena_listview1
   pro_textos
    Me.txt_numero.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   var_swpassword = False
   var_modifica_registro_transporte = False
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_transportes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_transportes, ColumnHeader)
End Sub

Private Sub lv_transportes_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_transportes.selectedItem = Item
        pro_textos
        var_modifica_registro_transporte = True
        txt_numero.Enabled = False

End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_transportes.SetFocus
      Call pro_avanzar(Me, lv_transportes, Button)
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_transportes.ListItems(1).Selected = True
      lv_transportes.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 4 Then
      numero_items_transportes = lv_transportes.ListItems.Count
      lv_transportes.ListItems(numero_items_transportes).Selected = True
      lv_transportes.selectedItem.EnsureVisible
      pro_textos
   End If
err0:
End Sub


Sub pro_guardar_transportes()
End Sub

Sub pro_elimina_transportes()
   var_si = MsgBox("¿Desea eliminar el registro?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      rs.Open "DELETE FROM xxvia_TB_CHOFERES WHERE id_chofer = '" + Me.txt_numero + "'", cnn, adOpenDynamic, adLockOptimistic
   End If
End Sub


Sub pro_llena_listview1()
   Dim list_item As ListItem
   Me.lv_transportes.ListItems.Clear
   rs.Open "select * from xxvia_tb_choferes", cnnoracle_4, adOpenDynamic, adLockOptimistic
   numero_items_transportes = 0
   While Not rs.EOF
      Set list_item = lv_transportes.ListItems.Add(, , rs!id_chofer)
      list_item.SubItems(1) = IIf(IsNull(rs!nombre), "", rs!nombre)
      list_item.SubItems(2) = IIf(IsNull(rs!licencia), "", rs!licencia)
      list_item.SubItems(3) = IIf(IsNull(rs!rfc), "", rs!rfc)
      rs.MoveNext:
      numero_items_transportes = numero_items_transportes + 1
    Wend
    rs.Close
End Sub


Sub pro_textos()
On Error GoTo err0:
Dim var_n As Double
   var_n = lv_transportes.ListItems.Count
   If var_n > 0 Then
      txt_numero = lv_transportes.selectedItem
      txt_nombre = lv_transportes.selectedItem.SubItems(1)
      txt_licencia = lv_transportes.selectedItem.SubItems(2)
      txt_RFC = lv_transportes.selectedItem.SubItems(3)
   End If
   var_numero_renglones = lv_transportes.Height / 312.5
   var_modifica_registro_transporte = True
   var_hubo_cambios = False
   Me.txt_numero.Enabled = False
err0:
End Sub

Private Sub pro_actualiza_ListView()
'    lv_transportes.SetFocus
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_transportes, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub

Private Sub txt_cubicaje_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_cubicaje_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
        KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub

Private Sub txt_licencia_Change()
var_hubo_cambios = True
End Sub

Private Sub txt_licencia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_RFC.SetFocus
   End If
End Sub

Private Sub txt_nombre_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_licencia.SetFocus
      
   End If
End Sub

Private Sub txt_numero_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub

Private Sub txt_placas_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_placas_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub




Private Sub txt_rfc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub
