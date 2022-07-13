VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmgruposfamilias 
   Caption         =   "Grupos de clientes o familias"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1365
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   16
      Top             =   6045
      Width           =   255
   End
   Begin VB.TextBox txt_buscar 
      Height          =   285
      Left            =   1605
      TabIndex        =   7
      Top             =   2370
      Width           =   1350
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   765
      Top             =   5760
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
            Picture         =   "frmgruposfamilias.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgruposfamilias.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgruposfamilias.frx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgruposfamilias.frx":1A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgruposfamilias.frx":2368
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgruposfamilias.frx":2C42
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgruposfamilias.frx":351C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgruposfamilias.frx":3AB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgruposfamilias.frx":3DD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgruposfamilias.frx":40EC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   165
      Top             =   5775
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
            Picture         =   "frmgruposfamilias.frx":4406
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgruposfamilias.frx":4CE0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tool_atras_siguiente 
      Height          =   330
      Left            =   2985
      TabIndex        =   17
      Top             =   2340
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Un Registro Atras"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Un Registro Adelante"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo Registro"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Guardar Registro"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar Registro"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Lista"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   6
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir de Esta Ventana"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   -30
      TabIndex        =   11
      Top             =   270
      Width           =   5685
   End
   Begin VB.Frame Frame1 
      Caption         =   "Grupos de clientes o familias "
      Height          =   1860
      Left            =   0
      TabIndex        =   0
      Top             =   390
      Width           =   5655
      Begin VB.TextBox txt_gruposfamilias 
         Height          =   285
         Index           =   4
         Left            =   1335
         MaxLength       =   1
         TabIndex        =   15
         Top             =   1485
         Width           =   300
      End
      Begin VB.TextBox txt_gruposfamilias 
         Height          =   285
         Index           =   3
         Left            =   1335
         MaxLength       =   12
         TabIndex        =   12
         Top             =   1185
         Width           =   1590
      End
      Begin VB.TextBox txt_gruposfamilias 
         Height          =   285
         Index           =   1
         Left            =   1335
         MaxLength       =   50
         TabIndex        =   2
         Top             =   555
         Width           =   4140
      End
      Begin VB.TextBox txt_gruposfamilias 
         Height          =   285
         Index           =   0
         Left            =   1335
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   765
      End
      Begin VB.TextBox txt_gruposfamilias 
         Height          =   285
         Index           =   2
         Left            =   1335
         MaxLength       =   12
         TabIndex        =   3
         Top             =   870
         Width           =   1590
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clasificación:"
         Height          =   195
         Index           =   5
         Left            =   375
         TabIndex        =   14
         Top             =   1515
         Width           =   930
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Cobranza:"
         Height          =   195
         Index           =   4
         Left            =   585
         TabIndex        =   13
         Top             =   1215
         Width           =   720
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Limite crédito:"
         Height          =   195
         Index           =   2
         Left            =   330
         TabIndex        =   6
         Top             =   900
         Width           =   975
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   1
         Left            =   705
         TabIndex        =   5
         Top             =   585
         Width           =   600
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   0
         Left            =   855
         TabIndex        =   4
         Top             =   255
         Width           =   450
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   0
      TabIndex        =   8
      Top             =   2190
      Width           =   5655
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Busqueda de grupo:"
         Height          =   195
         Left            =   195
         TabIndex        =   9
         Top             =   195
         Width           =   1440
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3780
      Left            =   0
      TabIndex        =   10
      Top             =   2685
      Width           =   5670
      Begin MSComctlLib.ListView lv_gruposfamilias 
         Height          =   3585
         Left            =   45
         TabIndex        =   19
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
         NumItems        =   5
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
            Text            =   "limite"
            Object.Width           =   35
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "cobranza"
            Object.Width           =   35
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "clasificacion"
            Object.Width           =   35
         EndProperty
      End
   End
   Begin VB.Label lab_paises 
      AutoSize        =   -1  'True
      Caption         =   "Busqueda de pais:"
      Height          =   195
      Index           =   3
      Left            =   0
      TabIndex        =   20
      Top             =   2790
      Width           =   1980
   End
End
Attribute VB_Name = "frmgruposfamilias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean



Private Sub cmd_buscar_Click()
    Call pro_busca_registro(lv_gruposfamilias, txt_buscar)
    txt_buscar = ""
    pro_textos

End Sub

Private Sub Combo1_Click()
   txt_gruposfamilias(0) = Obtener_llave(cnn, rs, "TB_paises", "VCHA_PAI_NOMBRE", Combo1, 0, "T")
End Sub

Private Sub Combo2_Click()
   txt_gruposfamilias(1) = Obtener_llave(cnn, rs, "TB_ESTADOS", "VCHA_EST_NOMBRE", Combo2, 1, "T")
End Sub

Private Sub Combo3_Click()
   txt_gruposfamilias(2) = Obtener_llave(cnn, rs, "TB_CIUDADES", "VCHA_CIU_NOMBRE", Combo3, 2, "T")
End Sub

Private Sub Form_Load()
    var_modifica_registro = True
    lv_gruposfamilias.SmallIcons = ImageList1
    Call pro_encabezadosView(Me, lv_gruposfamilias, False)
    Call pro_llena_listview1
    pro_textos

    Call pro_AsignarAViewColor(lv_gruposfamilias, Picture1, vbWhite, vbGray)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    var_modifica_registro = False
    Call menuvisible(Frmmenu2, True)
End Sub

Private Sub lv_gruposfamilias_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_gruposfamilias.SelectedItem = Item
        pro_textos
        var_modifica_registro = True
        txt_gruposfamilias(0).Enabled = True
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
   If Button.Index < 3 Then
      lv_gruposfamilias.SetFocus
      Call pro_avanzar(Me, lv_gruposfamilias, Button)
      pro_textos
   Else
      Call pro_busca_registro(lv_gruposfamilias, txt_buscar)
      txt_buscar = ""
      pro_textos
   End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
        Call pro_limpiatextos(Me)
        txt_gruposfamilias(0).Enabled = True
        txt_gruposfamilias(0).SetFocus: var_modifica_registro = False
    Case 2
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call pro_guardar_gruposfamilias
        End If
    Case 3
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call pro_elimina_gruposfamilias
        End If
    Case 5
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call gPrintListView(lv_gruposfamilias, "LISTADO DE gruposfamilias")
        End If
    Case 7
        Unload Me
    End Select

End Sub

Sub pro_guardar_gruposfamilias()

Dim ok As Boolean

Set TB_GRUPOSFAMILIAS = New TB_GRUPOSFAMILIAS
    
    
    If txt_gruposfamilias(0) <> "" And txt_gruposfamilias(1) <> "" Then
        If var_hubo_cambios Then
            ok = TB_GRUPOSFAMILIAS.Anadir(txt_gruposfamilias(0), txt_gruposfamilias(1), txt_gruposfamilias(2), txt_gruposfamilias(3), txt_gruposfamilias(4), fun_NombreUsuario, fun_NombrePc, Date)
            If ok Then
                pro_actualiza_ListView
                txt_gruposfamilias(0).Enabled = False
                MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
                txt_registros = lv_gruposfamilias.ListItems.Count
                var_modifica_registro = True
            Else
                MsgBox "No se puede grabar registro: " + TB_GRUPOSFAMILIAS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
        End If
    End If
    
Set TB_GRUPOSFAMILIAS = Nothing: var_hubo_cambios = False

End Sub

Sub pro_elimina_gruposfamilias()
Dim var_llave_usuarios As String

Set TB_GRUPOSFAMILIAS = New TB_GRUPOSFAMILIAS

    
    ok = True
    rs.Open "select * from TB_ARTICULOS,TB_DETALLE where TB_ARTICULOS.VCHA_ART_ARTICULO_ID = TB_DETALLE.VCHA_ART_ARTICULO_ID AND TB_ARTICULOS.VCHA_ART_LINEA = '" & txt_gruposfamilias(1) & "'", cnn, adOpenForwardOnly, adLockOptimistic
    If rs.RecordCount = 0 Then
        rs.Close
        If txt_gruposfamilias(0) <> "" And txt_gruposfamilias(1) <> "" And var_modifica_registro = True Then
            If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
                ok = TB_GRUPOSFAMILIAS.Eliminar(txt_gruposfamilias(0))
            Else
                GoTo SALIR:
            End If
            If ok Then
                MsgBox "Se Elimino Correctamente el Registro", vbInformation
                lv_gruposfamilias.ListItems.Remove (lv_gruposfamilias.SelectedItem.Index)
                Call pro_limpiatextos(Me)
                txt_registros = lv_gruposfamilias.ListItems.Count
                lv_gruposfamilias.SelectedItem.Selected = True
                pro_textos
            Else
                MsgBox "No se puede grabar registro: " + TB_GRUPOSFAMILIAS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
        End If
    Else
        SetTimer hwnd, NV_CLOSEMSGBOX, 1800, AddressOf TimerProc
        MsgBox "No se Puede Borrar Este Registro, Existen Dependencias", , "TRANSACCIONES [ AVISO ]"
        rs.Close
    End If

SALIR:
Set TB_GRUPOSFAMILIAS = Nothing

End Sub


Sub pro_llena_listview1()

Dim list_item As ListItem

    rs.Open "select * from TB_gruposfamilias", cnn, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
        Set list_item = lv_gruposfamilias.ListItems.Add(, , rs(0).Value): list_item.SmallIcon = 9
        list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
        list_item.SubItems(2) = IIf(IsNull(rs(2).Value), "", rs(2).Value)
        list_item.SubItems(3) = IIf(IsNull(rs(3).Value), "", rs(3).Value)
        list_item.SubItems(4) = IIf(IsNull(rs(4).Value), "", rs(4).Value)
    rs.MoveNext:
    Wend
    rs.Close

End Sub


Sub pro_textos()
On Error GoTo err0:
        txt_gruposfamilias(0) = lv_gruposfamilias.SelectedItem
        txt_gruposfamilias(1) = lv_gruposfamilias.SelectedItem.SubItems(1)
        txt_gruposfamilias(2) = lv_gruposfamilias.SelectedItem.SubItems(2)
        txt_gruposfamilias(3) = lv_gruposfamilias.SelectedItem.SubItems(3)
        txt_gruposfamilias(4) = lv_gruposfamilias.SelectedItem.SubItems(4)
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro = False Then
        Set list_item = lv_gruposfamilias.ListItems.Add(, , txt_gruposfamilias(0)): list_item.SmallIcon = 9
        list_item.SubItems(1) = txt_gruposfamilias(1)
        list_item.SubItems(2) = txt_gruposfamilias(2)
        list_item.SubItems(3) = txt_gruposfamilias(3)
        list_item.SubItems(4) = txt_gruposfamilias(4)
        list_item.EnsureVisible
        list_item.Selected = True
    Else
        lv_gruposfamilias.ListItems.Item(lv_gruposfamilias.SelectedItem.Index).Checked = False
        lv_gruposfamilias.ListItems.Item(lv_gruposfamilias.SelectedItem.Index) = txt_gruposfamilias(0)
        lv_gruposfamilias.ListItems.Item(lv_gruposfamilias.SelectedItem.Index).ListSubItems(1) = txt_gruposfamilias(1)
        lv_gruposfamilias.ListItems.Item(lv_gruposfamilias.SelectedItem.Index).ListSubItems(2) = txt_gruposfamilias(2)
        lv_gruposfamilias.ListItems.Item(lv_gruposfamilias.SelectedItem.Index).ListSubItems(3) = txt_gruposfamilias(3)
        lv_gruposfamilias.ListItems.Item(lv_gruposfamilias.SelectedItem.Index).ListSubItems(4) = txt_gruposfamilias(4)
        lv_gruposfamilias.ListItems.Item(lv_gruposfamilias.SelectedItem.Index).Selected = True
    End If
    lv_gruposfamilias.SetFocus
End Sub

Private Sub txt_gruposfamilias_Change(Index As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub

Private Sub txt_gruposfamilias_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Index < 4 Then
          txt_gruposfamilias(Index + 1).SetFocus
       Else
          txt_gruposfamilias(0).Enabled = True
          txt_gruposfamilias(0).SetFocus
       End If
    Else
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
       var_hubo_cambios = True
    End If
    
End Sub
