VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmequivalencias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Equivalencias de códigos"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   Icon            =   "frmequivalencias.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5955
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmequivalencias.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      Picture         =   "frmequivalencias.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Guardar Alt + G"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   810
      Picture         =   "frmequivalencias.frx":0ACE
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      Picture         =   "frmequivalencias.frx":0BA0
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1470
      Picture         =   "frmequivalencias.frx":0CA2
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5475
      Picture         =   "frmequivalencias.frx":0DA4
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   4695
      Left            =   165
      TabIndex        =   14
      Top             =   2430
      Width           =   5655
      Begin MSComctlLib.ListView lv_equivalencias 
         Height          =   4470
         Left            =   45
         TabIndex        =   15
         Top             =   150
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   7885
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
            Text            =   "Código Interno"
            Object.Width           =   4851
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Código Externo"
            Object.Width           =   4851
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "interno"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Canale de ventas"
      Height          =   1395
      Left            =   180
      TabIndex        =   10
      Top             =   420
      Width           =   5655
      Begin VB.CheckBox chk_interno 
         Caption         =   "Código Interno"
         Height          =   270
         Left            =   3465
         TabIndex        =   21
         Top             =   1065
         Width           =   1950
      End
      Begin VB.TextBox txt_equivalencias 
         Height          =   315
         Index           =   2
         Left            =   1275
         MaxLength       =   50
         TabIndex        =   8
         Top             =   990
         Width           =   2070
      End
      Begin VB.TextBox txt_equivalencias 
         Height          =   315
         Index           =   0
         Left            =   1275
         MaxLength       =   50
         TabIndex        =   6
         Top             =   300
         Width           =   1260
      End
      Begin VB.TextBox txt_equivalencias 
         Height          =   315
         Index           =   1
         Left            =   1275
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   7
         Top             =   645
         Width           =   4155
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Código Externo:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   1050
         Width           =   1125
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Código Interno:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   705
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
      TabIndex        =   9
      Top             =   30
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   5085
      Top             =   1140
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
            Picture         =   "frmequivalencias.frx":13DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequivalencias.frx":1CB8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   -30
      Top             =   1095
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
            Picture         =   "frmequivalencias.frx":2592
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequivalencias.frx":2E6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequivalencias.frx":3746
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequivalencias.frx":3CE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequivalencias.frx":45BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequivalencias.frx":4E98
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequivalencias.frx":5772
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequivalencias.frx":5884
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequivalencias.frx":5996
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequivalencias.frx":5AA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequivalencias.frx":5BBA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Index           =   0
      Left            =   150
      TabIndex        =   13
      Top             =   300
      Width           =   5655
   End
   Begin VB.TextBox txt_buscar 
      Height          =   315
      Left            =   1905
      TabIndex        =   17
      Top             =   1995
      Width           =   1350
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   165
      TabIndex        =   18
      Top             =   1845
      Width           =   5655
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3600
         TabIndex        =   19
         Top             =   135
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
         Caption         =   "Busqueda del Código:"
         Height          =   195
         Left            =   165
         TabIndex        =   20
         Top             =   195
         Width           =   1560
      End
   End
End
Attribute VB_Name = "frmequivalencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim numero_items_canalesventas As Integer
Dim bitacora As Boolean

Private Sub chk_interno_Click()
   var_hubo_cambios = True
End Sub

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
      Call pro_elimina_equivalencias
      rs.Open "select * from tb_equivalencias", cnn, adOpenDynamic, adLockOptimistic
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
   rsaux4.Open "SELECT * FROM TB_EQUIVALENCIAS WHERE VCHA_ART_ARTICULO_ID = '" + Me.txt_equivalencias(0) + "' AND VCHA_EQU_CODIGO_EQUIVALENTE = '" + Me.txt_equivalencias(2) + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rsaux4.EOF And var_modifica_registro_equivalencia = False Then
      MsgBox "Ya existe la equivalencia", vbOKOnly, "ATENCION"
   Else
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
         Call pro_guardar_equivalencias
         rs.Open "select * from tb_equivalencias", cnn, adOpenDynamic, adLockOptimistic
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
   End If
   rsaux4.Close
End Sub

Private Sub cmd_imprimir_Click()
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call gPrintListView(lv_equivalencias, "LISTADO DE equivalencias")
        End If

End Sub

Private Sub cmd_nuevo_Click()
   Call pro_limpiatextos(Me)
   txt_equivalencias(0).Enabled = True
   txt_equivalencias(0).SetFocus: var_modifica_registro_equivalencia = False
   cmd_guardar.Enabled = True
   cmd_deshacer.Enabled = True
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_equivalencia = False Then
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
   var_modifica_registro_equivalencia = True
   var_cadena_seguridad = ""
   Top = 0
   Left = 2900
   var_modifica_registro_equivalencia = True
   lv_equivalencias.SmallIcons = ImageList1
   Call pro_encabezadosView(Me, lv_equivalencias, False)
   Call pro_llena_listview1
   pro_textos
   rs.Open "select * from tb_equivalencias", cnn, adOpenDynamic, adLockOptimistic
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
   Call activa_forma(var_activa_forma_equivalencias)
End Sub

Private Sub lv_equivalencias_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_equivalencias, ColumnHeader)
End Sub

Private Sub lv_equivalencias_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_equivalencias.selectedItem = Item
        pro_textos
        var_modifica_registro_equivalencia = True
        txt_equivalencias(0).Enabled = False

End Sub



Sub pro_guardar_equivalencias()
Dim ok As Boolean
Set TB_EQUIVALENCIAS = New TB_EQUIVALENCIAS
   ok = True
   If txt_equivalencias(0) <> "" And txt_equivalencias(2) <> "" Then
      If var_hubo_cambios Then
         ok = TB_EQUIVALENCIAS.Anadir(txt_equivalencias(0), txt_equivalencias(2))

         If ok Then
            rs.Open "update tb_equivalencias set inte_equ_codigo_interno = " + CStr(Me.chk_interno.Value) + " where vcha_art_articulo_id = '" + Me.txt_equivalencias(0) + "' and vcha_equ_codigo_equivalente = '" + Me.txt_equivalencias(2) + "'", cnn, adOpenDynamic, adLockOptimistic
            pro_actualiza_ListView
            txt_equivalencias(0).Enabled = False
            MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
            txt_registros = lv_equivalencias.ListItems.Count
            var_modifica_registro_equivalencia = True
         Else
            MsgBox "No se puede grabar registro: " + TB_EQUIVALENCIAS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
         End If
      End If
   End If
   Set TB_EQUIVALENCIAS = Nothing: var_hubo_cambios = False
End Sub

Sub pro_elimina_equivalencias()
   Dim var_llave_usuarios As String
   Set TB_EQUIVALENCIAS = New TB_EQUIVALENCIAS
   Set TB_BITACORA_EQUIVALENCIAS = New TB_BITACORA_EQUIVALENCIAS
   On Error GoTo salir:
   ok = True
   If txt_equivalencias(0) <> "" And txt_equivalencias(2) <> "" And var_modifica_registro_equivalencia = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_EQUIVALENCIAS.Eliminar(txt_equivalencias(0), txt_equivalencias(2))
      Else
         GoTo salir:
      End If
      If ok Then
         bitacora = True
         var_operacion_bitacora = "E"
         bitacora = TB_BITACORA_EQUIVALENCIAS.Anadir(txt_equivalencias(0), "VCHA_CAN_NOMBRE", var_operacion_bitacora, "", txt_equivalencias(1), var_clave_usuario_global, fun_NombrePc, Date)
         numero_items_EQUIVALENCIAS = numero_items_EQUIVALENCIAS - 1
         MsgBox "Se Elimino Correctamente el Registro", vbInformation
         lv_equivalencias.ListItems.Remove (lv_equivalencias.selectedItem.Index)
         Call pro_limpiatextos(Me)
         txt_registros = lv_equivalencias.ListItems.Count
         lv_equivalencias.selectedItem.Selected = True
         pro_textos
       Else
         MsgBox "No se puede eliminar registro: " + TB_EQUIVALENCIAS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
   Set TB_EQUIVALENCIAS = Nothing
End Sub


Sub pro_llena_listview1()
   Dim list_item As ListItem
   rs.Open "select * from TB_equivalencias ", cnn, adOpenDynamic, adLockOptimistic
   numero_items_EQUIVALENCIAS = 0
   While Not rs.EOF
      Set list_item = lv_equivalencias.ListItems.Add(, , rs(0).Value)
      list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
      list_item.SubItems(2) = IIf(IsNull(rs(2).Value), 0, rs(2).Value)
      rs.MoveNext:
      numero_items_EQUIVALENCIAS = numero_items_EQUIVALENCIAS + 1
    Wend
    rs.Close
End Sub


Sub pro_textos()

'On Error GoTo err0:
Dim var_n As Double
   var_n = lv_equivalencias.ListItems.Count
   If var_n > 0 Then
      txt_equivalencias(0) = lv_equivalencias.selectedItem
      'MsgBox txt_equivalencias(0)
      rs.Open "select * from tb_Articulos where vcha_art_articulo_id = '" + txt_equivalencias(0) + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_equivalencias(1) = rs(1).Value
      End If
      rs.Close
      txt_equivalencias(2) = lv_equivalencias.selectedItem.SubItems(1)
      Me.chk_interno = Me.lv_equivalencias.selectedItem.SubItems(2)
   End If
   var_modifica_registro_equivalencia = True
   var_numero_renglones = lv_equivalencias.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_equivalencias.ColumnHeaders(2).Width = 2500.17
   Else
      lv_equivalencias.ColumnHeaders(2).Width = 2700.17
   End If
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem

    If var_modifica_registro_equivalencia = False Then
        Set list_item = lv_equivalencias.ListItems.Add(, , txt_equivalencias(0))
        list_item.SubItems(1) = txt_equivalencias(2)
        list_item.SubItems(2) = Me.chk_interno.Value
        list_item.EnsureVisible
        list_item.Selected = True
        numero_items_EQUIVALENCIAS = numero_items_EQUIVALENCIAS + 1
    Else
        lv_equivalencias.ListItems.Item(lv_equivalencias.selectedItem.Index).Checked = False
        lv_equivalencias.ListItems.Item(lv_equivalencias.selectedItem.Index) = txt_equivalencias(0)
        lv_equivalencias.ListItems.Item(lv_equivalencias.selectedItem.Index).ListSubItems(1) = txt_equivalencias(2)
        lv_equivalencias.ListItems.Item(lv_equivalencias.selectedItem.Index).ListSubItems(2) = Me.chk_interno.Value
        lv_equivalencias.ListItems.Item(lv_equivalencias.selectedItem.Index).Selected = True
    End If
'    lv_equivalencias.SetFocus
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      Me.lv_equivalencias.SetFocus
      Call pro_avanzar(Me, Me.lv_equivalencias, Button)
      lv_equivalencias.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 1 Then
      Me.lv_equivalencias.ListItems(1).Selected = True
      lv_equivalencias.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 4 Then
      numero_items_EQUIVALENCIAS = lv_equivalencias.ListItems.Count
      Me.lv_equivalencias.ListItems(numero_items_EQUIVALENCIAS).Selected = True
      lv_equivalencias.selectedItem.EnsureVisible
      pro_textos
   End If
err0:
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_equivalencias, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub

Private Sub txt_equivalencias_Change(Index As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub

Private Sub txt_equivalencias_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
   If Index < 2 Then
      Call pro_enfoque(KeyAscii)
   Else
      If KeyAscii = 13 Then
         If Me.cmd_guardar.Enabled = True Then
            Me.cmd_guardar.SetFocus
         End If
      End If
   End If
End Sub


Private Sub txt_equivalencias_LostFocus(Index As Integer)
    rs.Open "select * from tb_Articulos where vcha_art_articulo_id = '" + txt_equivalencias(0) + "'", cnn, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
       txt_equivalencias(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
    Else
       MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
    End If
    rs.Close
End Sub
