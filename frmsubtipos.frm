VERSION 5.00
Begin VB.Form frmsubtipoarticulos 
   Caption         =   "Subtipo de usos"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_buscar 
      Height          =   285
      Left            =   2700
      TabIndex        =   5
      Top             =   1875
      Width           =   1350
   End
   Begin VB.Frame Frame1 
      Caption         =   " Subtipo de usos "
      Height          =   1320
      Left            =   105
      TabIndex        =   0
      Top             =   435
      Width           =   5655
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1260
         TabIndex        =   12
         Text            =   "Combo2"
         Top             =   240
         Width           =   4110
      End
      Begin VB.TextBox txt_subtiposusos 
         Height          =   285
         Index           =   2
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   10
         Top             =   900
         Width           =   4020
      End
      Begin VB.TextBox txt_subtiposusos 
         Height          =   285
         Index           =   1
         Left            =   1260
         MaxLength       =   3
         TabIndex        =   2
         Top             =   600
         Width           =   780
      End
      Begin VB.TextBox txt_subtiposusos 
         Height          =   285
         Index           =   0
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   270
         Width           =   765
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion:"
         Height          =   195
         Index           =   4
         Left            =   315
         TabIndex        =   11
         Top             =   930
         Width           =   885
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave subtipo:"
         Height          =   195
         Index           =   2
         Left            =   195
         TabIndex        =   4
         Top             =   615
         Width           =   1005
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de uso:"
         Height          =   195
         Index           =   1
         Left            =   315
         TabIndex        =   3
         Top             =   300
         Width           =   885
      End
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   75
      TabIndex        =   9
      Top             =   315
      Width           =   5685
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   120
      TabIndex        =   6
      Top             =   1710
      Width           =   5655
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Busqueda de subtipo de artículo:"
         Height          =   195
         Left            =   195
         TabIndex        =   7
         Top             =   195
         Width           =   2355
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3705
      Left            =   105
      TabIndex        =   8
      Top             =   2220
      Width           =   5670
   End
End
Attribute VB_Name = "frmsubtipoarticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean



Private Sub cmd_buscar_Click()
    Call pro_busca_registro(lv_subtiposusos, txt_buscar, False)
    txt_buscar = ""
    pro_textos

End Sub

Private Sub Combo1_Click()
   txt_subtiposusos(0) = Obtener_llave(cnn, rs, "TB_EMPRESAS", "VCHA_EMP_NOMBRE", Combo1, 0, "T")
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Combo2.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub Combo2_Click()
   txt_subtiposusos(0) = Obtener_llave(cnn, rs, "TB_usos", "VCHA_uso_NOMBRE", Combo2, 0, "T")
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_subtiposusos(1).SetFocus
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
            var_resultado = InStr(1, var_menus, Me.Caption & "*1")
            If var_resultado <> 0 Then
                Set var_forma = frmsubtiposusos
                var_swpassword = True
                sw_primera_validacion = True
                frmsubtiposusos.Hide
                frmpasswords.Show 1
            End If
        End If
        If var_swpassword = False Then
            var_resultado = InStr(1, var_menus, Me.Caption & "*01")
            If var_resultado <> 0 Then
                Set var_forma = frmsubtiposusos
                var_swpassword = True
                sw_primera_validacion = True
                frmsubtiposusos.Hide
                frmpasswords2.txt_supervisor = var_supervisor
                frmpasswords2.Show
            End If
        End If
    End If
End If
End Sub

Private Sub Form_Load()
    rs.Open "select * from tb_usos", cnn, adOpenDynamic, adLockBatchOptimistic
    Call RecsetToCombo(Combo2.hWnd, rs, 1)
    rs.Close
    var_modifica_registro = True
    lv_subtiposusos.SmallIcons = ImageList1
    Call pro_encabezadosView(Me, lv_subtiposusos, False)
    Call pro_llena_listview1
    pro_textos

    Call pro_AsignarAViewColor(lv_subtiposusos, Picture1, vbWhite, vbGray)
    rs.Open "select * from tb_subtiposusos", cnn, adOpenDynamic, adLockOptimistic
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

Private Sub lv_subtiposusos_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_subtiposusos.SelectedItem = Item
        pro_textos
        var_modifica_registro = True
        txt_subtiposusos(0).Enabled = True
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
   If Button.Index < 3 Then
      lv_subtiposusos.SetFocus
      Call pro_avanzar(Me, lv_subtiposusos, Button)
      pro_textos
   Else
      Call pro_busca_registro(lv_subtiposusos, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
        Call pro_limpiatextos(Me)
        txt_subtiposusos(0).Enabled = True
        txt_subtiposusos(0).SetFocus: var_modifica_registro = False
        Toolbar1.Buttons.Item(2).Enabled = True
        Toolbar1.Buttons.Item(3).Enabled = True
    Case 2
        var_resultado = InStr(1, var_menus, Me.Caption)
        var_inicio = var_resultado + Len(Me.Caption) + 3
        If Mid(var_menus, var_inicio, 1) = "1" Then
            Set var_forma = frmsubtiposusos
            var_swpassword = True
            sw_primera_validacion = False
            frmpasswords.Show 1
        Else
            If Mid(var_menus, var_inicio, 2) = "01" Then
                Set var_forma = frmsubtiposusos
                var_swpassword = True
                sw_primera_validacion = False
                frmpasswords2.txt_supervisor = var_supervisor
                frmpasswords2.Show 1
            Else
               Call pro_guardar_subtiposusos
               rs.Open "select * from tb_subtiposusos", cnn, adOpenDynamic, adLockOptimistic
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
            Set var_forma = frmsubtiposusos
            var_swpassword = True
            sw_primera_validacion = False
            frmpasswords.Show 1
        Else
            If Mid(var_menus, var_inicio, 2) = "01" Then
                Set var_forma = frmsubtiposusos
                var_swpassword = True
                sw_primera_validacion = False
                frmpasswords2.txt_supervisor = var_supervisor
                frmpasswords2.Show 1
            Else
               Call pro_elimina_subtiposusos
               rs.Open "select * from tb_subtiposusos", cnn, adOpenDynamic, adLockOptimistic
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
           Call gPrintListView(lv_subtiposusos, "LISTADO DE subtiposusos")
        End If
    Case 8
        Unload Me
    End Select

End Sub

Sub pro_guardar_subtiposusos()

Dim ok As Boolean

Set TB_subtiposusos = New TB_subtiposusos
    
    
    If txt_subtiposusos(0) <> "" And txt_subtiposusos(1) <> "" Then
        If var_hubo_cambios Then
            ok = TB_subtiposusos.Anadir(txt_subtiposusos(0), txt_subtiposusos(1), txt_subtiposusos(2))
            If ok Then
                pro_actualiza_ListView
                txt_subtiposusos(0).Enabled = False
                MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
                txt_registros = lv_subtiposusos.ListItems.Count
                var_modifica_registro = True
            Else
                MsgBox "No se puede grabar registro: " + TB_subtiposusos.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
        End If
    End If
    
Set TB_subtiposusos = Nothing: var_hubo_cambios = False

End Sub

Sub pro_elimina_subtiposusos()
Dim var_llave_usuarios As String

Set TB_subtiposusos = New TB_subtiposusos

    On Error GoTo SALIR
    ok = True
    'rs.Open "select * from TB_ARTICULOS,TB_DETALLE where TB_ARTICULOS.VCHA_ART_ARTICULO_ID = TB_DETALLE.VCHA_ART_ARTICULO_ID AND TB_ARTICULOS.VCHA_ART_LINEA = '" & txt_subtiposusos(1) & "'", cnn, adOpenForwardOnly, adLockOptimistic
    'If rs.RecordCount = 0 Then
    '    rs.Close
        If txt_subtiposusos(0) <> "" And txt_subtiposusos(1) <> "" And var_modifica_registro = True Then
            If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
                ok = TB_subtiposusos.Eliminar(txt_subtiposusos(1))
            Else
                GoTo SALIR:
            End If
            If ok Then
                MsgBox "Se Elimino Correctamente el Registro", vbInformation
                lv_subtiposusos.ListItems.Remove (lv_subtiposusos.SelectedItem.Index)
                Call pro_limpiatextos(Me)
                txt_registros = lv_subtiposusos.ListItems.Count
                lv_subtiposusos.SelectedItem.Selected = True
                pro_textos
            Else
                MsgBox "No se puede grabar registro: " + TB_subtiposusos.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
        End If
    'Else
    '    SetTimer hwnd, NV_CLOSEMSGBOX, 1800, AddressOf TimerProc
    '    MsgBox "No se Puede Borrar Este Registro, Existen Dependencias", , "TRANSACCIONES [ AVISO ]"
    '    rs.Close
    'End If

SALIR:
Set TB_subtiposusos = Nothing

End Sub


Sub pro_llena_listview1()

Dim list_item As ListItem

    rs.Open "select * from TB_subtiposusos", cnn, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
        Set list_item = lv_subtiposusos.ListItems.Add(, , rs(1).Value): list_item.SmallIcon = 9
        list_item.SubItems(1) = IIf(IsNull(rs(2).Value), "", rs(2).Value)
        list_item.SubItems(2) = IIf(IsNull(rs(0).Value), "", rs(0).Value)
    rs.MoveNext:
    Wend
    rs.Close

End Sub


Sub pro_textos()
On Error GoTo err0:
        txt_subtiposusos(1) = lv_subtiposusos.SelectedItem
        txt_subtiposusos(2) = lv_subtiposusos.SelectedItem.SubItems(1)
        txt_subtiposusos(0) = lv_subtiposusos.SelectedItem.SubItems(2)
        Combo2 = Obtener_llave(cnn, rs, "TB_usos", "VCHA_uso_uso_ID", txt_subtiposusos(0), 1, "T")
        
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro = False Then
        Set list_item = lv_subtiposusos.ListItems.Add(, , txt_subtiposusos(1)): list_item.SmallIcon = 9
        list_item.SubItems(1) = txt_subtiposusos(2)
        list_item.SubItems(2) = txt_subtiposusos(0)
        list_item.EnsureVisible
        list_item.Selected = True
    Else
        lv_subtiposusos.ListItems.Item(lv_subtiposusos.SelectedItem.Index).Checked = False
        lv_subtiposusos.ListItems.Item(lv_subtiposusos.SelectedItem.Index) = txt_subtiposusos(1)
        lv_subtiposusos.ListItems.Item(lv_subtiposusos.SelectedItem.Index).ListSubItems(1) = txt_subtiposusos(2)
        lv_subtiposusos.ListItems.Item(lv_subtiposusos.SelectedItem.Index).ListSubItems(2) = txt_subtiposusos(0)
        lv_subtiposusos.ListItems.Item(lv_subtiposusos.SelectedItem.Index).Selected = True
    End If
    lv_subtiposusos.SetFocus
End Sub

Private Sub txt_subtiposusos_Change(Index As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub

Private Sub txt_subtiposusos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Index < 2 Then
          txt_subtiposusos(Index + 1).SetFocus
       Else
          Combo2.SetFocus
       End If
    Else
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
       var_hubo_cambios = True
    End If
    
End Sub
