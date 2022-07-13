VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmtipocambio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo cambio"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   Icon            =   "frmtipocambio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5955
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   90
      TabIndex        =   24
      Top             =   2985
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   25
         Top             =   480
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   3228
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
            Object.Width           =   7937
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   26
         Top             =   120
         Width           =   5610
      End
   End
   Begin MSComCtl2.MonthView mes 
      Height          =   2370
      Left            =   2175
      TabIndex        =   23
      Top             =   525
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   85196801
      CurrentDate     =   37581
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmtipocambio.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      Picture         =   "frmtipocambio.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Guardar Alt + G"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   810
      Picture         =   "frmtipocambio.frx":0ACE
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      Picture         =   "frmtipocambio.frx":0BA0
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1470
      Picture         =   "frmtipocambio.frx":0CA2
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5460
      Picture         =   "frmtipocambio.frx":0DA4
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2775
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   19
      Top             =   90
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   " Tipocambio "
      Height          =   1335
      Left            =   150
      TabIndex        =   0
      Top             =   450
      Width           =   5655
      Begin VB.TextBox txt_nombre_moneda 
         Height          =   315
         Left            =   1905
         TabIndex        =   8
         Top             =   225
         Width           =   3660
      End
      Begin VB.CommandButton cmdfecha 
         Height          =   285
         Index           =   0
         Left            =   1935
         Picture         =   "frmtipocambio.frx":13DE
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Seleccione la fecha"
         Top             =   585
         Width           =   315
      End
      Begin VB.TextBox txt_importe 
         Height          =   315
         Left            =   885
         TabIndex        =   10
         Top             =   915
         Width           =   1590
      End
      Begin VB.TextBox txt_fecha 
         Height          =   315
         Left            =   885
         TabIndex        =   9
         Top             =   570
         Width           =   1005
      End
      Begin VB.TextBox txt_moneda 
         Height          =   315
         Left            =   885
         TabIndex        =   7
         Top             =   225
         Width           =   1005
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   13
         Top             =   915
         Width           =   570
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   12
         Top             =   570
         Width           =   495
      End
      Begin VB.Label lab_estados 
         AutoSize        =   -1  'True
         Caption         =   "Moneda:"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   11
         Top             =   240
         Width           =   630
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   150
      TabIndex        =   14
      Top             =   1740
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   1815
         TabIndex        =   18
         Top             =   165
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3885
         TabIndex        =   21
         Top             =   180
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
         Caption         =   "Busqueda de moneda:"
         Height          =   195
         Left            =   195
         TabIndex        =   15
         Top             =   195
         Width           =   1605
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4935
      Left            =   150
      TabIndex        =   16
      Top             =   2280
      Width           =   5655
      Begin MSComctlLib.ListView lv_tipocambio 
         Height          =   4725
         Left            =   60
         TabIndex        =   20
         Top             =   150
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   8334
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Importe"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   17
      Top             =   300
      Width           =   5655
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   45
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
            Picture         =   "frmtipocambio.frx":14E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtipocambio.frx":1DBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtipocambio.frx":2694
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtipocambio.frx":2C30
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtipocambio.frx":350C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtipocambio.frx":3DE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtipocambio.frx":46C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtipocambio.frx":47D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtipocambio.frx":48E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtipocambio.frx":49F6
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmtipocambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean

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
      Call pro_elimina_tipocambio
      rs.Open "select * from tb_tipocambio", cnn, adOpenDynamic, adLockOptimistic
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
   var_resultado = InStr(1, var_menus, Me.Caption)
   var_inicio = var_resultado + Len(Me.Caption) + 3
   If Not IsNumeric(txt_importe) Then
      MsgBox "Importe incorrecto", vbOKOnly, "ATENCION"
      txt_importe = 1
   End If
   If Not IsDate(Me.txt_fecha) Then
      txt_fecha = Date
   End If
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
   If var_acepta_seguridad > 0 Then
      Call pro_guardar_tipocambio
      rs.Open "select * from tb_tipocambio", cnn, adOpenDynamic, adLockOptimistic
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

Private Sub cmd_imprimir_Click()
   x = x + 1
End Sub

Private Sub cmd_nuevo_Click()
   Call pro_limpiatextos(Me)
   txt_moneda.Enabled = True
   txt_moneda.SetFocus: var_modifica_registro_tipocambio = False
   cmd_guardar.Enabled = True
   cmd_deshacer.Enabled = True
   txt_fecha = Date
   txt_fecha.Enabled = True
   txt_moneda.Enabled = True
   txt_moneda.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_tipocambio = False Then
      var_si = MsgBox("No se han guardado los cambios, ¿Desea salir?", vbYesNo, "ATENCION")
      If var_si <> 6 Then
         GoTo salir:
      End If
   Else
      If var_hubo_cambios = True Then
         var_si = MsgBox("No se han guardado los cambios, ¿Desea salir?", vbYesNo, "ATENCION")
         If var_si <> 6 Then
            GoTo salir:
         End If
      End If
   End If
   Unload Me
   Exit Sub
salir:
End Sub

Private Sub cmdfecha_Click(Index As Integer)
   mes.Value = txt_fecha
   mes.Visible = True
   mes.SetFocus
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
   mes.Visible = False
   frm_lista.Visible = False
   var_modifica_registro_tipocambio = True
   lv_tipocambio.SmallIcons = ImageList1
   Call pro_encabezadosView(Me, lv_tipocambio, False)
   Call pro_llena_listview1
   pro_textos
   rs.Open "select * from tb_tipocambio", cnn, adOpenDynamic, adLockOptimistic
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
   var_modifica_registro_tipocambio = False
   Call activa_forma(var_activa_forma_tipocambio)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         txt_moneda = lv_lista.selectedItem
         txt_nombre_moneda = lv_lista.selectedItem.SubItems(1)
      Else
         txt_moneda = ""
         txt_nombre_moneda = ""
      End If
      txt_moneda.SetFocus
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub lv_tipocambio_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_tipocambio, ColumnHeader)
End Sub

Private Sub lv_tipocambio_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_tipocambio.selectedItem = Item
        pro_textos
        var_modifica_registro_tipocambio = True
        txt_moneda.Enabled = False

End Sub

Private Sub mes_DblClick()
   txt_fecha = mes.Value
   mes.Visible = False
End Sub

Private Sub mes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_fecha = mes.Value
      mes.Visible = False
   End If
   If KeyAscii = 27 Then
      mes.Visible = False
   End If
End Sub

Private Sub mes_LostFocus()
   mes.Visible = False
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      Me.lv_tipocambio.SetFocus
      Call pro_avanzar(Me, lv_tipocambio, Button)
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_tipocambio.ListItems(1).Selected = True
      lv_tipocambio.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 4 Then
      numero_items_tipoarticulos = lv_tipocambio.ListItems.Count
      lv_tipocambio.ListItems(numero_items_tipoarticulos).Selected = True
      lv_tipocambio.selectedItem.EnsureVisible
      pro_textos
   End If
err0:
End Sub


Sub pro_guardar_tipocambio()

Dim ok As Boolean

Set TB_TIPOCAMBIO = New TB_TIPOCAMBIO
    
    ok = True
    If txt_moneda <> "" And txt_fecha <> "" And txt_importe <> "" Then
        If var_hubo_cambios Then
            ok = TB_TIPOCAMBIO.Anadir(txt_moneda, txt_fecha, txt_importe)
            If ok Then
                pro_actualiza_ListView
                txt_moneda.Enabled = False
                txt_fecha.Enabled = False
                MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
                txt_registros = lv_tipocambio.ListItems.Count
                var_modifica_registro_tipocambio = True
            Else
                MsgBox "No se puede grabar registro: " + TB_TIPOCAMBIO.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
        End If
    End If
    
Set TB_TIPOCAMBIO = Nothing: var_hubo_cambios = False

End Sub

Sub pro_elimina_tipocambio()
   Dim var_llave_usuarios As String
   Set TB_TIPOCAMBIO = New TB_TIPOCAMBIO
   ok = True
   'On Error GoTo salir:
   If txt_moneda <> "" And txt_fecha <> "" And txt_importe _
      <> "" And txt_importe <> "" And var_modifica_registro_tipocambio = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_TIPOCAMBIO.Eliminar(txt_moneda, txt_fecha)
      Else
         GoTo salir:
      End If
      If ok Then
        MsgBox "Se Elimino Correctamente el Registro", vbInformation
        lv_tipocambio.ListItems.Remove (lv_tipocambio.selectedItem.Index)
        Call pro_limpiatextos(Me)
        txt_registros = lv_tipocambio.ListItems.Count
        lv_tipocambio.selectedItem.Selected = True
        pro_textos
      Else
        MsgBox "No se puede eliminar registro: " + TB_TIPOCAMBIO.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
   Set TB_TIPOCAMBIO = Nothing
End Sub


Sub pro_llena_listview1()
   Dim list_item As ListItem
   rs.Open "select * from TB_tipocambio", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
      Set list_item = lv_tipocambio.ListItems.Add(, , rs(0).Value)
      list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
      list_item.SubItems(2) = IIf(IsNull(rs(2).Value), "", rs(2).Value)
      rs.MoveNext:
    Wend
    rs.Close
End Sub


Sub pro_textos()
   On Error GoTo err0:
   Dim var_n As Integer
   var_n = lv_tipocambio.ListItems.Count
   If var_n > 0 Then
      txt_fecha = lv_tipocambio.selectedItem.SubItems(1)
      txt_moneda = lv_tipocambio.selectedItem
      txt_importe = lv_tipocambio.selectedItem.SubItems(2)
      txt_moneda.Enabled = False
      txt_fecha.Enabled = False
      rs.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + txt_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_moneda = IIf(IsNull(rs!vcha_mon_nombre), "", rs!vcha_mon_nombre)
      Else
         txt_nombre_moneda = ""
      End If
      rs.Close
      If var_n > 11 Then
         lv_tipocambio.ColumnHeaders(2).Width = 1750
      Else
         lv_tipocambio.ColumnHeaders(2).Width = 1550
      End If
      
   End If
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem

    If var_modifica_registro_tipocambio = False Then
        Set list_item = lv_tipocambio.ListItems.Add(, , txt_moneda)
        list_item.SubItems(1) = txt_fecha
        list_item.SubItems(2) = txt_importe
        list_item.EnsureVisible
        list_item.Selected = True
    Else
        lv_tipocambio.ListItems.Item(lv_tipocambio.selectedItem.Index).Checked = False
        lv_tipocambio.ListItems.Item(lv_tipocambio.selectedItem.Index) = txt_moneda
        lv_tipocambio.ListItems.Item(lv_tipocambio.selectedItem.Index).ListSubItems(1) = txt_fecha
        lv_tipocambio.ListItems.Item(lv_tipocambio.selectedItem.Index).ListSubItems(2) = txt_importe
        lv_tipocambio.ListItems.Item(lv_tipocambio.selectedItem.Index).Selected = True
    End If
'    lv_tipocambio.SetFocus
End Sub



Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_tipocambio, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub

Private Sub txt_fecha_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_fecha_LostFocus()
   If Not IsDate(txt_fecha) Then
      MsgBox "Fecha Incorrecta", vbOKOnly, "ATENCION"
      txt_fecha.SetFocus
   End If
End Sub

Private Sub txt_importe_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If Me.cmd_guardar.Enabled = True Then
         Me.cmd_guardar.SetFocus
      End If
   End If
End Sub

Private Sub txt_importe_LostFocus()
   If Not IsNumeric(txt_importe) Then
      MsgBox "Importe incorrecto", vbOKOnly, "ATENCION"
      txt_importe = 1
   End If
End Sub

Private Sub txt_moneda_Change()
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub

Private Sub txt_moneda_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para accesar al catálogo"
End Sub

Private Sub txt_moneda_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_MONEDAS order by vcha_mon_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_mon_moneda_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_mon_nombre), "", rs!vcha_mon_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "MONEDAS"
      var_tipo_lista = 3
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
   If KeyCode = 117 Then
      var_activa_forma_monedas = Me.Name
      Me.Enabled = False
      frmmonedas.Show
   End If
End Sub

Private Sub txt_moneda_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_moneda_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_moneda) <> "" Then
      rs.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + txt_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_moneda = IIf(IsNull(rs!vcha_mon_nombre), "", rs!vcha_mon_nombre)
         rs.Close
      Else
         rs.Close
         txt_moneda = ""
         txt_nombre_moneda = ""
         MsgBox "Clave de moneda incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_nombre_moneda_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para accesar al catálogo"
End Sub

Private Sub txt_nombre_moneda_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_MONEDAS order by vcha_mon_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_mon_moneda_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_mon_nombre), "", rs!vcha_mon_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "MONEDAS"
      var_tipo_lista = 3
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
   If KeyCode = 117 Then
      Me.Enabled = False
      var_activa_forma_monedas = Me.Name
      frmmonedas.Show
   End If
End Sub

Private Sub txt_nombre_moneda_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_moneda_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub
