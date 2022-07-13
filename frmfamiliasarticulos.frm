VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#11.0#0"; "VBSKFREE.OCX"
Begin VB.Form frmfamiliasarticulos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Familias de artículos"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   5850
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2775
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   10
      Top             =   90
      Width           =   255
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   135
      TabIndex        =   5
      Top             =   1425
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   285
         Left            =   1785
         TabIndex        =   9
         Top             =   165
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3870
         TabIndex        =   13
         Top             =   135
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList"
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
         Caption         =   "Busqueda de familia:"
         Height          =   195
         Left            =   195
         TabIndex        =   6
         Top             =   195
         Width           =   1470
      End
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   165
      Top             =   4935
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
            Picture         =   "frmfamiliasarticulos.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfamiliasarticulos.frx":08DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   105
      TabIndex        =   11
      Top             =   15
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
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
            Object.ToolTipText     =   "Deshacer cambios"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar Registro"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Lista"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir de Esta Ventana"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   105
      TabIndex        =   8
      Top             =   300
      Width           =   5685
   End
   Begin VB.Frame Frame1 
      Caption         =   "Familias de artículos "
      Height          =   1020
      Left            =   135
      TabIndex        =   0
      Top             =   405
      Width           =   5655
      Begin VB.TextBox txt_familiasarticulos 
         Height          =   285
         Index           =   1
         Left            =   1275
         TabIndex        =   2
         Top             =   630
         Width           =   4155
      End
      Begin VB.TextBox txt_familiasarticulos 
         Height          =   285
         Index           =   0
         Left            =   1275
         MaxLength       =   3
         TabIndex        =   1
         Top             =   270
         Width           =   690
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Index           =   2
         Left            =   345
         TabIndex        =   4
         Top             =   645
         Width           =   885
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Familia:"
         Height          =   195
         Index           =   1
         Left            =   705
         TabIndex        =   3
         Top             =   285
         Width           =   525
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3825
      Left            =   120
      TabIndex        =   7
      Top             =   1935
      Width           =   5670
      Begin MSComctlLib.ListView lv_familiasarticulos 
         Height          =   3615
         Left            =   45
         TabIndex        =   12
         Top             =   135
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   6376
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
   End
   Begin vbskfree.Skinner Skinner1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1270
      _ExtentY        =   1270
      MaxButton       =   0
      MinButton       =   0
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
   End
   Begin MSComctlLib.ImageList ImageList 
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
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfamiliasarticulos.frx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfamiliasarticulos.frx":1A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfamiliasarticulos.frx":2368
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfamiliasarticulos.frx":2904
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfamiliasarticulos.frx":31E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfamiliasarticulos.frx":3ABA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfamiliasarticulos.frx":4394
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfamiliasarticulos.frx":44A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfamiliasarticulos.frx":45B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfamiliasarticulos.frx":46CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfamiliasarticulos.frx":47DC
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmfamiliasarticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim numero_items_familiasarticulos As Integer

Private Sub cmd_buscar_Click()
    Call pro_busca_registro(lv_familiasarticulos, txt_buscar, False)
    txt_buscar = ""
    pro_textos

End Sub



Private Sub Adodc1_WillMove(ByVal adReason As adodb.EventReasonEnum, adStatus As adodb.EventStatusEnum, ByVal pRecordset As adodb.Recordset)

End Sub

Private Sub Combo1_Click()
   txt_familiasarticulos(0) = Obtener_llave(cnn, rsaux, "TB_EMPRESAS", "VCHA_EMP_NOMBRE", Combo1, 0, "T")
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_familiasarticulos(1).SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub Form_Activate()
Dim var_resultado As Variant
Dim mientras As Integer
mientras = 0
If mientras = 0 Then

    If sw_primera_validacion = False Then
    
        If var_swpassword = False Then
        Call menuvisible(Frmmenu2, False)
            var_resultado = InStr(1, var_menus, Me.caption & "*1")
            If var_resultado <> 0 Then
                Set var_forma = frmfamiliasarticulos
                var_swpassword = True
                sw_primera_validacion = True
                frmfamiliasarticulos.Hide
                frmpasswords.Show 1
            End If
        End If
        If var_swpassword = False Then
            var_resultado = InStr(1, var_menus, Me.caption & "*01")
            If var_resultado <> 0 Then
                Set var_forma = frmfamiliasarticulos
                var_swpassword = True
                sw_primera_validacion = True
                frmfamiliasarticulos.Hide
                frmpasswords2.txt_supervisor = var_supervisor
                frmpasswords2.Show
            End If
        End If
    End If
 End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      End
   End If
End Sub

Private Sub Form_Load()
    var_modifica_registro = True
    'lv_familiasarticulos.SmallIcons = ImageList1
    Call pro_encabezadosView(Me, lv_familiasarticulos, False)
    Call pro_llena_listview1
    pro_textos
    Call pro_AsignarAViewColor(lv_familiasarticulos, Picture1, vbWhite, vbGray)
    rs.Open "select * from tb_familiasarticulos", cnn, adOpenDynamic, adLockOptimistic
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

End Sub

Private Sub Form_Unload(Cancel As Integer)
    var_swpassword = False
    var_modifica_registro = False
    Call menuvisible(Frmmenu2, True)
End Sub

Private Sub lv_familiasarticulos_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_familiasarticulos.SelectedItem = Item
        pro_textos
        var_modifica_registro = True
        txt_familiasarticulos(0).Enabled = False

End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_familiasarticulos.SetFocus
      Call pro_avanzar(Me, lv_familiasarticulos, Button)
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_familiasarticulos.ListItems(1).Selected = True
      pro_textos
   End If
   If Button.Index = 4 Then
      lv_familiasarticulos.ListItems(numero_items_familiasarticulos).Selected = True
      pro_textos
   End If
err0:
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
        Call pro_limpiatextos(Me)
        txt_familiasarticulos(0).Enabled = True
        txt_familiasarticulos(0).SetFocus: var_modifica_registro = False
        Toolbar1.Buttons.Item(2).Enabled = True
        Toolbar1.Buttons.Item(3).Enabled = True
    Case 2
        var_movimiento = 1
        var_resultado = InStr(1, var_menus, Me.caption)
        var_inicio = var_resultado + Len(Me.caption) + 3
        If Mid(var_menus, var_inicio, 1) = "1" Then
            Set var_forma = frmfamiliasarticulos
            var_swpassword = True
            sw_primera_validacion = False
            frmpasswords.Show 1
        Else
            If Mid(var_menus, var_inicio, 2) = "01" Then
                Set var_forma = frmfamiliasarticulos
                var_swpassword = True
                sw_primera_validacion = False
                frmpasswords2.txt_supervisor = var_supervisor
                frmpasswords2.Show 1
            Else
            Call pro_guardar_familiasarticulos
               rs.Open "select * from tb_familiasarticulos", cnn, adOpenDynamic, adLockOptimistic
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
        var_movimiento = 2
        var_resultado = InStr(1, var_menus, Me.caption)
        var_inicio = var_resultado + Len(Me.caption) + 3
        If Mid(var_menus, var_inicio, 1) = "1" Then
            Set var_forma = frmfamiliasarticulos
            var_swpassword = True
            sw_primera_validacion = False
            frmpasswords.Show 1
        Else
            If Mid(var_menus, var_inicio, 2) = "01" Then
                Set var_forma = frmfamiliasarticulos
                var_swpassword = True
                sw_primera_validacion = False
                frmpasswords2.Show 1
            Else
               Call pro_elimina_familiasarticulos
               rs.Open "select * from tb_familiasarticulos", cnn, adOpenDynamic, adLockOptimistic
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
           Call gPrintListView(lv_familiasarticulos, "LISTADO DE familiasarticulos")
    Case 8
        Unload Me
    End Select

End Sub

Sub pro_guardar_familiasarticulos()

Dim ok As Boolean

Set TB_FAMILIASARTICULOS = New TB_FAMILIASARTICULOS
    
    ok = True
    If txt_familiasarticulos(0) <> "" And txt_familiasarticulos(1) <> "" Then
        If var_hubo_cambios Then
            ok = TB_FAMILIASARTICULOS.Anadir(txt_familiasarticulos(0), txt_familiasarticulos(1), Date, fun_NombreUsuario, fun_NombrePc)
            If ok Then
                pro_actualiza_ListView
                txt_familiasarticulos(0).Enabled = False
                MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
                txt_registros = lv_familiasarticulos.ListItems.Count
                var_modifica_registro = True
            Else
                MsgBox "No se puede grabar registro: " + TB_FAMILIASARTICULOS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
        End If
    End If
    
Set TB_FAMILIASARTICULOS = Nothing: var_hubo_cambios = False

End Sub

Sub pro_elimina_familiasarticulos()
   Dim var_llave_usuarios As String
   On Error GoTo SALIR:
   Set TB_FAMILIASARTICULOS = New TB_FAMILIASARTICULOS
   ok = True
   If txt_familiasarticulos(0) <> "" And txt_familiasarticulos(1) <> "" And var_modifica_registro = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_FAMILIASARTICULOS.Eliminar(txt_familiasarticulos(0))
      Else
         GoTo SALIR:
      End If
      If ok Then
         numero_items_familiasarticulos = numero_items_familiasarticulos - 1
        MsgBox "Se Elimino Correctamente el Registro", vbInformation
        lv_familiasarticulos.ListItems.Remove (lv_familiasarticulos.SelectedItem.Index)
        Call pro_limpiatextos(Me)
        txt_registros = lv_familiasarticulos.ListItems.Count
        lv_familiasarticulos.SelectedItem.Selected = True
        pro_textos
      Else
        MsgBox "No se puede eliminar registro: " + TB_FAMILIASARTICULOS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
SALIR:
   Set TB_FAMILIASARTICULOS = Nothing
End Sub


Sub pro_llena_listview1()
   Dim list_item As ListItem
   rs.Open "select * from TB_familiasarticulos", cnn, adOpenDynamic, adLockOptimistic
   numero_items_familiasarticulos = 0
   While Not rs.EOF
      Set list_item = lv_familiasarticulos.ListItems.Add(, , rs(0).Value): list_item.SmallIcon = 9
      list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
      rs.MoveNext:
      numero_items_familiasarticulos = numero_items_familiasarticulos + 1
    Wend
    rs.Close
End Sub


Sub pro_textos()
On Error GoTo err0:
        txt_familiasarticulos(0) = lv_familiasarticulos.SelectedItem
        txt_familiasarticulos(1) = lv_familiasarticulos.SelectedItem.SubItems(1)
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem

    If var_modifica_registro = False Then
        Set list_item = lv_familiasarticulos.ListItems.Add(, , txt_familiasarticulos(0)): list_item.SmallIcon = 9
        list_item.SubItems(1) = txt_familiasarticulos(1)
        list_item.EnsureVisible
        list_item.Selected = True
        numero_items_familiasarticulos = numero_items_familiasarticulos + 1
    Else
        lv_familiasarticulos.ListItems.Item(lv_familiasarticulos.SelectedItem.Index).Checked = False
        lv_familiasarticulos.ListItems.Item(lv_familiasarticulos.SelectedItem.Index) = txt_familiasarticulos(0)
        lv_familiasarticulos.ListItems.Item(lv_familiasarticulos.SelectedItem.Index).ListSubItems(1) = txt_familiasarticulos(1)
        lv_familiasarticulos.ListItems.Item(lv_familiasarticulos.SelectedItem.Index).Selected = True
    End If
'    lv_familiasarticulos.SetFocus
End Sub

Private Sub txt_familiasarticulos_Change(Index As Integer)
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
       var_hubo_cambios = True
End Sub

Private Sub txt_familiasarticulos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Index < 1 Then
          txt_familiasarticulos(Index + 1).SetFocus
       End If
    Else
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
       var_hubo_cambios = True
    End If
    
End Sub

