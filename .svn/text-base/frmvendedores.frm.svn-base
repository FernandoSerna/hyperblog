VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmvendedores 
   Caption         =   "Catálogo de vendedores"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6075
   Icon            =   "frmvendedores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   6075
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   15
      Top             =   6495
      Width           =   255
   End
   Begin VB.TextBox txt_buscar 
      Height          =   285
      Left            =   2145
      TabIndex        =   8
      Top             =   2490
      Width           =   1350
   End
   Begin VB.Frame Frame1 
      Caption         =   "Vendedores"
      Height          =   1875
      Left            =   210
      TabIndex        =   4
      Top             =   465
      Width           =   5655
      Begin VB.ComboBox cmb_vendedores 
         Height          =   315
         Left            =   1335
         TabIndex        =   21
         Top             =   855
         Width           =   4170
      End
      Begin VB.TextBox txt_vendedores 
         Height          =   285
         Index           =   4
         Left            =   1350
         MaxLength       =   12
         TabIndex        =   20
         Top             =   1500
         Width           =   1590
      End
      Begin VB.TextBox txt_vendedores 
         Height          =   285
         Index           =   3
         Left            =   1350
         MaxLength       =   12
         TabIndex        =   3
         Top             =   1185
         Width           =   1590
      End
      Begin VB.TextBox txt_vendedores 
         Height          =   285
         Index           =   1
         Left            =   1335
         MaxLength       =   50
         TabIndex        =   1
         Top             =   555
         Width           =   4140
      End
      Begin VB.TextBox txt_vendedores 
         Height          =   285
         Index           =   0
         Left            =   1350
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   765
      End
      Begin VB.TextBox txt_vendedores 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1335
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   2
         Top             =   870
         Width           =   1590
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Canal de Venta:"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   19
         Top             =   915
         Width           =   1140
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Comisión:"
         Height          =   195
         Index           =   4
         Left            =   645
         TabIndex        =   13
         Top             =   1530
         Width           =   675
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono:"
         Height          =   195
         Index           =   2
         Left            =   645
         TabIndex        =   7
         Top             =   1215
         Width           =   675
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   6
         Top             =   585
         Width           =   600
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   0
         Left            =   870
         TabIndex        =   5
         Top             =   255
         Width           =   450
      End
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   360
      Top             =   6210
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
            Picture         =   "frmvendedores.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmvendedores.frx":11A4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tool_atras_siguiente 
      Height          =   330
      Left            =   3885
      TabIndex        =   16
      Top             =   2460
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   180
      TabIndex        =   17
      Top             =   0
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo Registro"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Guardar Registro"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Deshacer"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar Regsitro"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir Catálogo"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salir de Esta Ventana"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   165
      TabIndex        =   12
      Top             =   270
      Width           =   5685
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   195
      TabIndex        =   9
      Top             =   2295
      Width           =   5655
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Busqueda de vendedor:"
         Height          =   195
         Left            =   195
         TabIndex        =   10
         Top             =   195
         Width           =   1710
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3780
      Left            =   210
      TabIndex        =   11
      Top             =   2865
      Width           =   5670
      Begin MSComctlLib.ListView lv_vendedores 
         Height          =   3585
         Left            =   30
         TabIndex        =   18
         Top             =   135
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   6324
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "telefono"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "comision"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "fecha captura"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "fecha alta"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1425
      Top             =   6240
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
            Picture         =   "frmvendedores.frx":1A7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmvendedores.frx":2358
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmvendedores.frx":2C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmvendedores.frx":31CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmvendedores.frx":3AAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmvendedores.frx":4384
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmvendedores.frx":4C5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmvendedores.frx":4D70
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmvendedores.frx":4E82
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmvendedores.frx":4F94
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lab_paises 
      AutoSize        =   -1  'True
      Caption         =   "Busqueda de pais:"
      Height          =   195
      Index           =   3
      Left            =   195
      TabIndex        =   14
      Top             =   3240
      Width           =   1980
   End
End
Attribute VB_Name = "frmvendedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim opcion_fecha As Integer
Dim numero_items_vendedores As Integer



Private Sub cmd_buscar_Click()
 '   Call pro_busca_registro(lv_vendedores, txt_buscar)
 '   txt_buscar = ""
 '   pro_textos

End Sub

Private Sub Combo1_Click()
   txt_vendedores(0) = Obtener_llave(cnn, rs, "TB_paises", "VCHA_PAI_NOMBRE", Combo1, 0, "T")
End Sub

Private Sub Combo2_Click()
   txt_vendedores(1) = Obtener_llave(cnn, rs, "TB_ESTADOS", "VCHA_EST_NOMBRE", Combo2, 1, "T")
End Sub

Private Sub Combo3_Click()
   txt_vendedores(2) = Obtener_llave(cnn, rs, "TB_CIUDADES", "VCHA_CIU_NOMBRE", Combo3, 2, "T")
End Sub

Private Sub cmb_vendedores_Change()
      txt_vendedores(2) = Obtener_llave(cnn, rs, "TB_CANALESVENTAS", "VCHA_CAN_NOMBRE", cmb_vendedores, 0, "T")
End Sub

Private Sub cmb_vendedores_Click()
      txt_vendedores(2) = Obtener_llave(cnn, rs, "TB_CANALESVENTAS", "VCHA_CAN_NOMBRE", cmb_vendedores, 0, "T")
End Sub

Private Sub cmb_vendedores_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub Form_Load()
    rs.Open "select * from tb_canalesventas", cnn, adOpenDynamic, adLockBatchOptimistic
       Call RecsetToCombo(cmb_vendedores.hwnd, rs, 1)
    rs.Close
    var_modifica_registro = True
    lv_vendedores.SmallIcons = ImageList1
    Call pro_encabezadosView(Me, lv_vendedores, False)
    Call pro_llena_listview1
    pro_textos

    Call pro_AsignarAViewColor(lv_vendedores, Picture1, vbWhite, vbGray)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    var_modifica_registro = False
End Sub

Private Sub lv_vendedores_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_vendedores.SelectedItem = Item
        pro_textos
        var_modifica_registro = True
        txt_vendedores(0).Enabled = True
End Sub

Private Sub mon_vendedores_DblClick()
   If opcion_fecha = 1 Then
      txt_vendedores(4) = Format(mon_vendedores.Value, "dd/mm/yyyy")
   End If
   If opcion_fecha = 2 Then
      txt_vendedores(5) = Format(mon_vendedores.Value, "dd/mm/yyyy")
   End If
   mon_vendedores.Visible = False
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
  If Button.Index = 2 Or Button.Index = 3 Then
      lv_vendedores.SetFocus
      Call pro_avanzar(Me, lv_vendedores, Button)
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_vendedores.ListItems(1).Selected = True
      pro_textos
   End If
   If Button.Index = 4 Then
      lv_vendedores.ListItems(numero_items_vendedores).Selected = True
      pro_textos
   End If
err0:
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
        Call pro_limpiatextos(Me)
        txt_vendedores(0).Enabled = True
        txt_vendedores(0).SetFocus: var_modifica_registro = False
        Toolbar1.Buttons.Item(2).Enabled = True
        Toolbar1.Buttons.Item(3).Enabled = True
    Case 2
        var_resultado = InStr(1, var_menus, Me.Caption)
        var_inicio = var_resultado + Len(Me.Caption) + 3
        If Mid(var_menus, var_inicio, 1) = "1" Then
            Set var_forma = frmvendedores
            var_swpassword = True
            sw_primera_validacion = False
            frmpasswords.Show 1
        Else
            If Mid(var_menus, var_inicio, 2) = "01" Then
                Set var_forma = frmvendedores
                var_swpassword = True
                sw_primera_validacion = False
                frmpasswords2.txt_supervisor = var_supervisor
                frmpasswords2.Show 1
            Else
               Call pro_guardar_vendedores
               rs.Open "select * from tb_vendedores", cnn, adOpenDynamic, adLockOptimistic
               If rs.BOF Then
                  Toolbar1.Buttons.Item(2).Enabled = False
                  Toolbar1.Buttons.Item(3).Enabled = False
                  Toolbar1.Buttons.Item(4).Enabled = False
               Else
                  Toolbar1.Buttons.Item(2).Enabled = True
                  Toolbar1.Buttons.Item(3).Enabled = True
                  Toolbar1.Buttons.Item(4).Enabled = True
               End If
               rs.Close
            End If
        End If
    Case 3
       Call pro_textos
    Case 4
        var_resultado = InStr(1, var_menus, Me.Caption)
        var_inicio = var_resultado + Len(Me.Caption) + 3
        If Mid(var_menus, var_inicio, 1) = "1" Then
            Set var_forma = frmvendedores
            var_swpassword = True
            sw_primera_validacion = False
            frmpasswords.Show 1
        Else
            If Mid(var_menus, var_inicio, 2) = "01" Then
                Set var_forma = frmvendedores
                var_swpassword = True
                sw_primera_validacion = False
                frmpasswords2.txt_supervisor = var_supervisor
                frmpasswords2.Show 1
            Else
               Call pro_elimina_vendedores
               rs.Open "select * from tb_vendedores", cnn, adOpenDynamic, adLockOptimistic
               If rs.BOF Then
                  Toolbar1.Buttons.Item(2).Enabled = False
                  Toolbar1.Buttons.Item(3).Enabled = False
                  Toolbar1.Buttons.Item(4).Enabled = False
               Else
                  Toolbar1.Buttons.Item(2).Enabled = True
                  Toolbar1.Buttons.Item(3).Enabled = True
                  Toolbar1.Buttons.Item(4).Enabled = True
               End If
               rs.Close
            End If
        End If
    Case 6
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call gPrintListView(lv_vendedores, "LISTADO DE vendedores")
        End If
    Case 8
        Unload Me
    End Select

End Sub

Sub pro_guardar_vendedores()
   Dim ok As Boolean
   Set TB_VENDEDORES = New TB_VENDEDORES
   Set TB_BITACORA_VENDEDORES = New TB_BITACORA_VENDEDORES
   If txt_vendedores(0) <> "" And txt_vendedores(1) <> "" Then
      If var_hubo_cambios Then
         rs.Open "SELECT * FROM TB_VENDEDORES WHERE VCHA_VEN_VENDEDOR_ID = '" + txt_vendedores(0) + "'", cnn, adOpenDynamic, adLockOptimistic
         ok = TB_VENDEDORES.Anadir(txt_vendedores(0), txt_vendedores(1), txt_vendedores(2), txt_vendedores(3), txt_vendedores(4))
         If ok Then
            bitacora = True
            If var_modifica_registro = False Then
               var_operacion_bitacora = "I"
               bitacora = TB_BITACORA_VENDEDORES.Anadir(txt_vendedores(0), "VCHA_VEN_NOMBRE", var_operacion_bitacora, "", txt_vendedores(1), var_clave_usuario_global, fun_NombrePc, Date)
            Else
               var_operacion_bitacora = "M"
               If rs(0) <> txt_vendedores(0) Then
                  bitacora = TB_BITACORA_VENDEDORES.Anadir(txt_vendedores(0), "VCHA_VEN_VENDEDOR_ID", var_operacion_bitacora, rs(0), txt_vendedores(0), var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(1) <> txt_vendedores(1) Then
                  bitacora = TB_BITACORA_VENDEDORES.Anadir(txt_vendedores(0), "VCHA_VEN_NOMBRE", var_operacion_bitacora, rs(1), txt_vendedores(1), var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(2) <> txt_vendedores(2) Then
                  bitacora = TB_BITACORA_VENDEDORES.Anadir(txt_vendedores(0), "VCHA_CAN_CANAL_VENTA_ID", var_operacion_bitacora, rs(2), txt_vendedores(2), var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(3) <> txt_vendedores(3) Then
                  bitacora = TB_BITACORA_VENDEDORES.Anadir(txt_vendedores(0), "VCHA_VEN_TELEFONO", var_operacion_bitacora, rs(3), txt_vendedores(3), var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(4) <> txt_vendedores(4) Then
                  bitacora = TB_BITACORA_VENDEDORES.Anadir(txt_vendedores(0), "INTE_VEN_COMISION", var_operacion_bitacora, rs(4), txt_vendedores(4), var_clave_usuario_global, fun_NombrePc, Date)
               End If
            End If
            rs.Close
                pro_actualiza_ListView
                txt_vendedores(0).Enabled = False
                MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
                txt_registros = lv_vendedores.ListItems.Count
                'var_modifica_registro = True
            Else
                MsgBox "No se puede grabar registro: " + TB_VENDEDORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
        End If
    End If
    
Set TB_VENDEDORES = Nothing: var_hubo_cambios = False

End Sub

Sub pro_elimina_vendedores()
   Dim var_llave_usuarios As String
   Set TB_VENDEDORES = New TB_VENDEDORES
   Set TB_BITACORA_VENDEDORES = New TB_BITACORA_VENDEDORES
   ok = True
   If txt_vendedores(0) <> "" And txt_vendedores(1) <> "" Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_VENDEDORES.Eliminar(txt_vendedores(0))
      Else
         GoTo salir:
      End If
      If ok Then
         var_operacion_bitacora = "E"
         bitacora = TB_BITACORA_VENDEDORES.Anadir(txt_vendedores(0), "VCHA_VEN_NOMBRE", var_operacion_bitacora, txt_vendedores(1), "", var_clave_usuario_global, fun_NombrePc, Date)
         MsgBox "Se Elimino Correctamente el Registro", vbInformation
         lv_vendedores.ListItems.Remove (lv_vendedores.SelectedItem.Index)
         Call pro_limpiatextos(Me)
         txt_registros = lv_vendedores.ListItems.Count
         lv_vendedores.SelectedItem.Selected = True
         pro_textos
         numero_items_vendedores = numero_items_vendedores - 1
      Else
         MsgBox "No se puede grabar registro: " + TB_VENDEDORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
   Set TB_VENDEDORES = Nothing
End Sub


Sub pro_llena_listview1()

Dim list_item As ListItem
    numero_items_vendedores = 0
    rs.Open "select * from TB_vendedores", cnn, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
        Set list_item = lv_vendedores.ListItems.Add(, , rs(0).Value)
        list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
        list_item.SubItems(2) = IIf(IsNull(rs(2).Value), "", rs(2).Value)
        list_item.SubItems(3) = IIf(IsNull(rs(3).Value), "", rs(3).Value)
        list_item.SubItems(4) = IIf(IsNull(rs(4).Value), "", rs(4).Value)
        numero_items_vendedores = numero_items_vendedores + 1
    rs.MoveNext:
    Wend
    rs.Close

End Sub


Sub pro_textos()
On Error GoTo err0:
        txt_vendedores(0) = lv_vendedores.SelectedItem
        txt_vendedores(1) = lv_vendedores.SelectedItem.SubItems(1)
        txt_vendedores(2) = lv_vendedores.SelectedItem.SubItems(2)
        txt_vendedores(3) = lv_vendedores.SelectedItem.SubItems(3)
        txt_vendedores(4) = lv_vendedores.SelectedItem.SubItems(4)
        cmb_vendedores = Obtener_llave(cnn, rs, "TB_CANALESVENTAS", "VCHA_CAN_CANAL_VENTA_ID", txt_vendedores(2), 1, "T")
        
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro = False Then
        Set list_item = lv_vendedores.ListItems.Add(, , txt_vendedores(0))
        list_item.SubItems(1) = txt_vendedores(1)
        list_item.SubItems(2) = txt_vendedores(2)
        list_item.SubItems(3) = txt_vendedores(3)
        list_item.SubItems(4) = txt_vendedores(4)
        list_item.EnsureVisible
        list_item.Selected = True
        numero_items_vendedores = numero_items_vendedores + 1
    Else
  '      lv_vendedores.ListItems.Item(lv_vendedores.SelectedItem.Index).Checked = False
        lv_vendedores.ListItems.Item(lv_vendedores.SelectedItem.Index) = txt_vendedores(0)
        lv_vendedores.ListItems.Item(lv_vendedores.SelectedItem.Index).ListSubItems(1) = txt_vendedores(1)
        lv_vendedores.ListItems.Item(lv_vendedores.SelectedItem.Index).ListSubItems(2) = txt_vendedores(2)
        lv_vendedores.ListItems.Item(lv_vendedores.SelectedItem.Index).ListSubItems(3) = txt_vendedores(3)
        lv_vendedores.ListItems.Item(lv_vendedores.SelectedItem.Index).ListSubItems(4) = txt_vendedores(4)
        lv_vendedores.ListItems.Item(lv_vendedores.SelectedItem.Index).Selected = True
    End If
    lv_vendedores.SetFocus
End Sub

Private Sub txt_buscar_LostFocus()
   Call pro_busca_registro(lv_vendedores, txt_buscar, False)
   txt_buscar = ""
   pro_textos
End Sub

Private Sub txt_vendedores_Change(Index As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub

Private Sub txt_vendedores_KeyPress(Index As Integer, KeyAscii As Integer)
    var_hubo_cambios = True
    If (KeyAscii = 13 And Index = 4) Or (KeyAscii = 13 And Index = 5) Then
        If Index = 4 Then
           opcion_fecha = 1
        End If
        If Index = 5 Then
           opcion_fecha = 2
        End If
        mon_vendedores.Visible = True: mon_vendedores.SetFocus
    End If
    
End Sub
