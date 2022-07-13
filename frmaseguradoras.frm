VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmaseguradoras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo de aseguradoras"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   Icon            =   "frmaseguradoras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5955
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2775
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   17
      Top             =   90
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txt_buscar 
      Height          =   315
      Left            =   2295
      TabIndex        =   7
      Top             =   2460
      Width           =   1350
   End
   Begin VB.Frame Frame1 
      Caption         =   " Aseguradoras "
      Height          =   1905
      Left            =   150
      TabIndex        =   0
      Top             =   420
      Width           =   5655
      Begin VB.TextBox txt_aseguradoras 
         Height          =   285
         Index           =   3
         Left            =   1320
         TabIndex        =   14
         Top             =   1185
         Width           =   2010
      End
      Begin VB.TextBox txt_aseguradoras 
         Height          =   285
         Index           =   4
         Left            =   1320
         TabIndex        =   13
         Top             =   1485
         Width           =   4230
      End
      Begin VB.TextBox txt_aseguradoras 
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   3
         Top             =   870
         Width           =   4230
      End
      Begin VB.TextBox txt_aseguradoras 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   2
         Top             =   555
         Width           =   4215
      End
      Begin VB.TextBox txt_aseguradoras 
         Height          =   285
         Index           =   0
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   765
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono:"
         Height          =   195
         Index           =   5
         Left            =   495
         TabIndex        =   16
         Top             =   1245
         Width           =   675
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Referencia:"
         Height          =   195
         Index           =   4
         Left            =   345
         TabIndex        =   15
         Top             =   1530
         Width           =   825
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Direccion:"
         Height          =   195
         Index           =   2
         Left            =   450
         TabIndex        =   6
         Top             =   900
         Width           =   720
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   1
         Left            =   570
         TabIndex        =   5
         Top             =   615
         Width           =   600
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   4
         Top             =   270
         Width           =   450
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   150
      TabIndex        =   9
      Top             =   2310
      Width           =   5655
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3900
         TabIndex        =   20
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
         Caption         =   "Busqueda de aseguradora:"
         Height          =   195
         Left            =   195
         TabIndex        =   10
         Top             =   195
         Width           =   1920
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4410
      Left            =   150
      TabIndex        =   12
      Top             =   2805
      Width           =   5670
      Begin MSComctlLib.ListView lv_aseguradoras 
         Height          =   4215
         Left            =   45
         TabIndex        =   19
         Top             =   135
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   7435
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
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   6879
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Direccion"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Telefono"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Referencia"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   150
      TabIndex        =   18
      Top             =   15
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
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
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   3750
      Top             =   255
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
            Picture         =   "frmaseguradoras.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmaseguradoras.frx":11A4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   120
      TabIndex        =   11
      Top             =   300
      Width           =   5685
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
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmaseguradoras.frx":1A7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmaseguradoras.frx":2358
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmaseguradoras.frx":2C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmaseguradoras.frx":31CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmaseguradoras.frx":3AAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmaseguradoras.frx":4384
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmaseguradoras.frx":4C5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmaseguradoras.frx":4D70
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmaseguradoras.frx":4E82
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmaseguradoras.frx":4F94
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmaseguradoras.frx":50A6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lab_paises 
      AutoSize        =   -1  'True
      Caption         =   "Busqueda de pais:"
      Height          =   195
      Index           =   3
      Left            =   150
      TabIndex        =   8
      Top             =   2535
      Width           =   1320
   End
End
Attribute VB_Name = "frmaseguradoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim numero_items_aseguradoras As Integer




Private Sub cmd_buscar_Click()
    Call pro_busca_registro(lv_aseguradoras, txt_buscar, False)
    txt_buscar = ""
    pro_textos

End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 0
   Left = 2900
    var_modifica_registro = True
    lv_aseguradoras.SmallIcons = ImageList1
    
    Call pro_encabezadosView(Me, lv_aseguradoras, False)
    Call pro_llena_listview1
    pro_textos

    Call pro_AsignarAViewColor(lv_aseguradoras, Picture1, vbWhite, vbGray)
    rs.Open "select * from tb_aseguradoras", cnn, adOpenDynamic, adLockOptimistic
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
    If var_activa_menu = True Then
       Frmmenu2.Enabled = True
    End If
End Sub

Private Sub lv_aseguradoras_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_aseguradoras, ColumnHeader)
End Sub

Private Sub lv_aseguradoras_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_aseguradoras.selectedItem = Item
        pro_textos
        var_modifica_registro = True
        txt_aseguradoras(0).Enabled = False

End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_aseguradoras.SetFocus
      Call pro_avanzar(Me, lv_aseguradoras, Button)
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_aseguradoras.ListItems(1).Selected = True
      pro_textos
   End If
   If Button.Index = 4 Then
      lv_aseguradoras.ListItems(numero_items_aseguradoras).Selected = True
      pro_textos
   End If
err0:
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
        Call pro_limpiatextos(Me)
        txt_aseguradoras(0).Enabled = True
        txt_aseguradoras(0).SetFocus: var_modifica_registro = False
        Toolbar1.Buttons.Item(2).Enabled = True
        Toolbar1.Buttons.Item(3).Enabled = True
    Case 2
        var_resultado = InStr(1, var_menus, Me.Caption)
        var_inicio = var_resultado + Len(Me.Caption) + 3
        If Mid(var_menus, var_inicio, 1) = "1" Then
            Set var_forma = frmaseguradoras
            var_swpassword = True
            sw_primera_validacion = False
            frmpasswords.Show 1
        Else
            If Mid(var_menus, var_inicio, 2) = "01" Then
                Set var_forma = frmaseguradoras
                var_swpassword = True
                sw_primera_validacion = False
                frmpasswords2.txt_supervisor = var_supervisor
                frmpasswords2.Show 1
            Else
               Call pro_guardar_aseguradoras
               rs.Open "select * from tb_aseguradoras", cnn, adOpenDynamic, adLockOptimistic
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
            Set var_forma = frmaseguradoras
            var_swpassword = True
            sw_primera_validacion = False
            frmpasswords.Show 1
        Else
            If Mid(var_menus, var_inicio, 2) = "01" Then
                Set var_forma = frmaseguradoras
                var_swpassword = True
                sw_primera_validacion = False
                frmpasswords2.txt_supervisor = var_supervisor
                frmpasswords2.Show 1
            Else
               Call pro_elimina_aseguradoras
               rs.Open "select * from tb_aseguradoras", cnn, adOpenDynamic, adLockOptimistic
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
           Call gPrintListView(lv_aseguradoras, "LISTADO DE aseguradoras")
        End If
    Case 8
        Unload Me
    End Select

End Sub

Sub pro_guardar_aseguradoras()

End Sub

Sub pro_elimina_aseguradoras()
Dim var_llave_usuarios As String


salir:

End Sub


Sub pro_llena_listview1()

Dim list_item As ListItem
   numero_items_aseguradoras = 0
    rs.Open "select * from TB_aseguradoras", cnn, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
        Set list_item = lv_aseguradoras.ListItems.Add(, , rs(0).Value): list_item.SmallIcon = 9
        list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
        list_item.SubItems(2) = IIf(IsNull(rs(2).Value), "", rs(2).Value)
        list_item.SubItems(3) = IIf(IsNull(rs(3).Value), "", rs(3).Value)
        list_item.SubItems(4) = IIf(IsNull(rs(4).Value), "", rs(4).Value)
    rs.MoveNext:
    numero_items_aseguradoras = numero_items_aseguradoras + 1
    Wend
    rs.Close

End Sub


Sub pro_textos()
On Error GoTo err0:
        txt_aseguradoras(0) = lv_aseguradoras.selectedItem
        txt_aseguradoras(1) = lv_aseguradoras.selectedItem.SubItems(1)
        txt_aseguradoras(2) = lv_aseguradoras.selectedItem.SubItems(2)
        txt_aseguradoras(3) = lv_aseguradoras.selectedItem.SubItems(3)
        txt_aseguradoras(4) = lv_aseguradoras.selectedItem.SubItems(4)
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro = False Then
        Set list_item = lv_aseguradoras.ListItems.Add(, , txt_aseguradoras(0)): list_item.SmallIcon = 9
        list_item.SubItems(1) = txt_aseguradoras(1)
        list_item.SubItems(2) = txt_aseguradoras(2)
        list_item.SubItems(3) = txt_aseguradoras(3)
        list_item.SubItems(4) = txt_aseguradoras(4)
        list_item.EnsureVisible
        list_item.Selected = True
        numero_items_aseguradoras = numero_items_aseguradoras + 1
    Else
        lv_aseguradoras.ListItems.Item(lv_aseguradoras.selectedItem.Index).Checked = False
        lv_aseguradoras.ListItems.Item(lv_aseguradoras.selectedItem.Index) = txt_aseguradoras(0)
        lv_aseguradoras.ListItems.Item(lv_aseguradoras.selectedItem.Index).ListSubItems(1) = txt_aseguradoras(1)
        lv_aseguradoras.ListItems.Item(lv_aseguradoras.selectedItem.Index).ListSubItems(2) = txt_aseguradoras(2)
        lv_aseguradoras.ListItems.Item(lv_aseguradoras.selectedItem.Index).ListSubItems(3) = txt_aseguradoras(3)
        lv_aseguradoras.ListItems.Item(lv_aseguradoras.selectedItem.Index).ListSubItems(4) = txt_aseguradoras(4)
        lv_aseguradoras.ListItems.Item(lv_aseguradoras.selectedItem.Index).Selected = True
    End If
    lv_aseguradoras.SetFocus
End Sub

Private Sub txt_aseguradoras_Change(Index As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub

Private Sub txt_aseguradoras_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Index < 4 Then
          txt_aseguradoras(Index + 1).SetFocus
       Else
          txt_aseguradoras(0).Enabled = True
          txt_aseguradoras(0).SetFocus
       End If
    Else
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
       var_hubo_cambios = True
    End If
    
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_aseguradoras, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub
