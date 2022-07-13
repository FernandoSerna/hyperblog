VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean



Private Sub cmd_buscar_Click()
    Call pro_busca_registro(lv_ciudades, txt_buscar, False)
    txt_buscar = ""
    pro_textos

End Sub

Private Sub Combo1_Click()
Dim var_clave_pais As String
    txt_ciudades(0) = Obtener_llave(cnn, rs, "TB_PAISES", "VCHA_PAI_NOMBRE", Combo1, 0, "T")
    rs.Open "select * from tb_estados where vcha_pai_pais = '" & txt_ciudades(0) & "'", cnn, adOpenDynamic, adLockBatchOptimistic
    Call RecsetToCombo(Combo2.hWnd, rs, 2)
    rs.Close
    If Combo2.ListIndex <> -1 Then
        Combo2.ListIndex = 0
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Combo2.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub Combo2_Click()
   txt_ciudades(1) = Obtener_llave(cnn, rs, "TB_ESTADOS", "VCHA_EST_NOMBRE", Combo2, 1, "T")
   
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_ciudades(2).SetFocus
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
                Set var_forma = frmciudades
                var_swpassword = True
                sw_primera_validacion = True
                frmciudades.Hide
                frmpasswords.Show 1
            End If
        End If
        If var_swpassword = False Then
            var_resultado = InStr(1, var_menus, Me.Caption & "*01")
            If var_resultado <> 0 Then
                Set var_forma = frmciudades
                var_swpassword = True
                sw_primera_validacion = True
                frmciudades.Hide
                frmpasswords2.txt_supervisor = var_supervisor
                frmpasswords2.Show
            End If
        End If
    End If
End If
End Sub

Private Sub Form_Load()
    rs.Open "select * from tb_paises", cnn, adOpenDynamic, adLockBatchOptimistic
    Call RecsetToCombo(Combo1.hWnd, rs, 1)
    rs.Close
    var_modifica_registro = True
    lv_ciudades.SmallIcons = ImageList1
    Call pro_encabezadosView(Me, lv_ciudades, False)
    Call pro_llena_listview1
    pro_textos

    Call pro_AsignarAViewColor(lv_ciudades, Picture1, vbWhite, vbGray)
    rs.Open "select * from tb_ciudades", cnn, adOpenDynamic, adLockOptimistic
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

Private Sub lv_ciudades_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_ciudades.SelectedItem = Item
        pro_textos
        var_modifica_registro = True
        txt_ciudades(0).Enabled = True
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
   If Button.Index < 3 Then
      lv_ciudades.SetFocus
      Call pro_avanzar(Me, lv_ciudades, Button)
      pro_textos
   Else
      Call pro_busca_registro(lv_ciudades, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
        Call pro_limpiatextos(Me)
        txt_ciudades(0).Enabled = True
        txt_ciudades(0).SetFocus: var_modifica_registro = False
        Toolbar1.Buttons.Item(2).Enabled = True
        Toolbar1.Buttons.Item(3).Enabled = True
    Case 2
        var_resultado = InStr(1, var_menus, Me.Caption)
        var_inicio = var_resultado + Len(Me.Caption) + 3
        If Mid(var_menus, var_inicio, 1) = "1" Then
            Set var_forma = frmciudades
            var_swpassword = True
            sw_primera_validacion = False
            frmpasswords.Show 1
        Else
            If Mid(var_menus, var_inicio, 2) = "01" Then
                Set var_forma = frmciudades
                var_swpassword = True
                sw_primera_validacion = False
                frmpasswords2.txt_supervisor = var_supervisor
                frmpasswords2.Show 1
            Else
               Call pro_guardar_ciudades
               rs.Open "select * from tb_ciudades", cnn, adOpenDynamic, adLockOptimistic
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
            Set var_forma = frmciudades
            var_swpassword = True
            sw_primera_validacion = False
            frmpasswords.Show 1
        Else
            If Mid(var_menus, var_inicio, 2) = "01" Then
                Set var_forma = frmciudades
                var_swpassword = True
                sw_primera_validacion = False
                frmpasswords2.txt_supervisor = var_supervisor
                frmpasswords2.Show 1
            Else
               Call pro_elimina_ciudades
               rs.Open "select * from tb_ciudades", cnn, adOpenDynamic, adLockOptimistic
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
           Call gPrintListView(lv_ciudades, "LISTADO DE ciudades")
        End If
    Case 8
        Unload Me
    End Select

End Sub

Sub pro_guardar_ciudades()

Dim ok As Boolean

Set TB_CIUDADES = New TB_CIUDADES
    
    
    If txt_ciudades(0) <> "" And txt_ciudades(1) <> "" Then
        If var_hubo_cambios Then
            ok = TB_CIUDADES.Anadir(txt_ciudades(0), txt_ciudades(1), txt_ciudades(2), txt_ciudades(4), txt_ciudades(3), fun_NombreUsuario, fun_NombrePc, Date)
            If ok Then
                pro_actualiza_ListView
                txt_ciudades(0).Enabled = False
                MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
                txt_registros = lv_ciudades.ListItems.Count
                var_modifica_registro = True
            Else
                MsgBox "No se puede grabar registro: " + TB_CIUDADES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
        End If
    End If
    
Set TB_CIUDADES = Nothing: var_hubo_cambios = False

End Sub

Sub pro_elimina_ciudades()
Dim var_llave_usuarios As String

Set TB_CIUDADES = New TB_CIUDADES
On Error GoTo SALIR:
    
    ok = True
    rs.Open "select * from TB_ARTICULOS,TB_DETALLE where TB_ARTICULOS.VCHA_ART_ARTICULO_ID = TB_DETALLE.VCHA_ART_ARTICULO_ID AND TB_ARTICULOS.VCHA_ART_LINEA = '" & txt_ciudades(1) & "'", cnn, adOpenForwardOnly, adLockOptimistic
    If rs.RecordCount = 0 Then
        rs.Close
        If txt_ciudades(0) <> "" And txt_ciudades(1) <> "" And var_modifica_registro = True Then
            If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
                ok = TB_CIUDADES.Eliminar(txt_ciudades(2))
            Else
                GoTo SALIR:
            End If
            If ok Then
                MsgBox "Se Elimino Correctamente el Registro", vbInformation
                lv_ciudades.ListItems.Remove (lv_ciudades.SelectedItem.Index)
                Call pro_limpiatextos(Me)
                txt_registros = lv_ciudades.ListItems.Count
                lv_ciudades.SelectedItem.Selected = True
                pro_textos
            Else
                MsgBox "No se puede grabar registro: " + TB_CIUDADES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
        End If
    Else
        SetTimer hWnd, NV_CLOSEMSGBOX, 1800, AddressOf TimerProc
        MsgBox "No se Puede Borrar Este Registro, Existen Dependencias", , "TRANSACCIONES [ AVISO ]"
        rs.Close
    End If

SALIR:
Set TB_CIUDADES = Nothing

End Sub


Sub pro_llena_listview1()

Dim list_item As ListItem

    rs.Open "select * from TB_ciudades", cnn, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
        Set list_item = lv_ciudades.ListItems.Add(, , rs(2).Value): list_item.SmallIcon = 9
        list_item.SubItems(1) = IIf(IsNull(rs(4).Value), "", rs(4).Value)
        list_item.SubItems(2) = IIf(IsNull(rs(0).Value), "", rs(0).Value)
        list_item.SubItems(3) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
        list_item.SubItems(4) = IIf(IsNull(rs(3).Value), "", rs(3).Value)
    rs.MoveNext:
    Wend
    rs.Close

End Sub


Sub pro_textos()
On Error GoTo err0:
        txt_ciudades(2) = lv_ciudades.SelectedItem
        txt_ciudades(3) = lv_ciudades.SelectedItem.SubItems(1)
        txt_ciudades(0) = lv_ciudades.SelectedItem.SubItems(2)
        txt_ciudades(1) = lv_ciudades.SelectedItem.SubItems(3)
        txt_ciudades(4) = lv_ciudades.SelectedItem.SubItems(4)
        Combo1 = Obtener_llave(cnn, rs, "TB_PAISES", "VCHA_PAI_PAIS", txt_ciudades(0), 1, "T")
        Combo2 = Obtener_llave(cnn, rs, "TB_ESTADOS", "VCHA_EST_ESTADO", txt_ciudades(1), 2, "T")
        
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro = False Then
        Set list_item = lv_ciudades.ListItems.Add(, , txt_ciudades(2)): list_item.SmallIcon = 9
        list_item.SubItems(1) = txt_ciudades(3)
        list_item.SubItems(2) = txt_ciudades(0)
        list_item.SubItems(3) = txt_ciudades(1)
        list_item.SubItems(4) = txt_ciudades(4)
        list_item.EnsureVisible
        list_item.Selected = True
    Else
        lv_ciudades.ListItems.Item(lv_ciudades.SelectedItem.Index).Checked = False
        lv_ciudades.ListItems.Item(lv_ciudades.SelectedItem.Index) = txt_ciudades(2)
        lv_ciudades.ListItems.Item(lv_ciudades.SelectedItem.Index).ListSubItems(1) = txt_ciudades(3)
        lv_ciudades.ListItems.Item(lv_ciudades.SelectedItem.Index).ListSubItems(2) = txt_ciudades(0)
        lv_ciudades.ListItems.Item(lv_ciudades.SelectedItem.Index).ListSubItems(3) = txt_ciudades(1)
        lv_ciudades.ListItems.Item(lv_ciudades.SelectedItem.Index).ListSubItems(4) = txt_ciudades(4)
        lv_ciudades.ListItems.Item(lv_ciudades.SelectedItem.Index).Selected = True
    End If
    lv_ciudades.SetFocus
End Sub

Private Sub txt_ciudades_Change(Index As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub

Private Sub txt_ciudades_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Index < 4 Then
          txt_ciudades(Index + 1).SetFocus
       Else
          Combo1.SetFocus
       End If
    Else
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
       var_hubo_cambios = True
    End If
    
End Sub

