VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmubicaciones_almacen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ubicaciones de artículos por almacen"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   8115
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   825
      TabIndex        =   0
      Top             =   2895
      Width           =   6915
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   15
         Top             =   480
         Width           =   6810
         _ExtentX        =   12012
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
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   8819
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   16
         Top             =   120
         Width           =   6840
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4080
      Left            =   105
      TabIndex        =   25
      Top             =   3120
      Width           =   7875
      Begin MSComctlLib.ListView lv_articulos 
         Height          =   3870
         Left            =   45
         TabIndex        =   26
         Top             =   135
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   6826
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
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "agente"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "zona"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "empresa"
            Object.Width           =   0
         EndProperty
      End
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
               Picture         =   "frmubicaciones_almacen.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmubicaciones_almacen.frx":08DA
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   105
      TabIndex        =   22
      Top             =   2565
      Width           =   7875
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   1845
         TabIndex        =   14
         Top             =   150
         Width           =   1860
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   4815
         TabIndex        =   23
         Top             =   165
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
         Caption         =   "Busqueda de Artículo:"
         Height          =   195
         Left            =   195
         TabIndex        =   24
         Top             =   195
         Width           =   1590
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Rutas "
      Height          =   2130
      Left            =   120
      TabIndex        =   18
      Top             =   420
      Width           =   7875
      Begin VB.TextBox txt_ubicacion_3 
         Height          =   315
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   13
         Top             =   1650
         Width           =   2955
      End
      Begin VB.TextBox txt_ubicacion_2 
         Height          =   315
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   12
         Top             =   1320
         Width           =   2955
      End
      Begin VB.TextBox txt_almacen 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   7
         Top             =   255
         Width           =   1710
      End
      Begin VB.TextBox txt_nombre_almacen 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3045
         MaxLength       =   50
         TabIndex        =   8
         Top             =   255
         Width           =   4620
      End
      Begin VB.TextBox txt_articulo 
         Height          =   315
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   9
         Top             =   585
         Width           =   1710
      End
      Begin VB.TextBox txt_ubicacion_1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   11
         Top             =   915
         Width           =   2955
      End
      Begin VB.TextBox txt_nombre_articulo 
         Height          =   315
         Left            =   3045
         TabIndex        =   10
         Top             =   585
         Width           =   4620
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Ubicación 3:"
         Height          =   195
         Index           =   2
         Left            =   345
         TabIndex        =   29
         Top             =   1710
         Width           =   900
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Ubicación 2:"
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   28
         Top             =   1380
         Width           =   900
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Almacén:"
         Height          =   195
         Index           =   1
         Left            =   330
         TabIndex        =   21
         Top             =   315
         Width           =   660
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Artículo:"
         Height          =   195
         Index           =   5
         Left            =   330
         TabIndex        =   20
         Top             =   645
         Width           =   600
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Ubicación 1:"
         Height          =   195
         Index           =   7
         Left            =   330
         TabIndex        =   19
         Top             =   1020
         Width           =   900
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
      TabIndex        =   17
      Top             =   75
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmubicaciones_almacen.frx":11B4
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
      Picture         =   "frmubicaciones_almacen.frx":12B6
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
      Picture         =   "frmubicaciones_almacen.frx":13B8
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
      Picture         =   "frmubicaciones_almacen.frx":148A
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
      Picture         =   "frmubicaciones_almacen.frx":158C
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7710
      Picture         =   "frmubicaciones_almacen.frx":168E
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
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
            Picture         =   "frmubicaciones_almacen.frx":1CC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmubicaciones_almacen.frx":25A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmubicaciones_almacen.frx":2E7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmubicaciones_almacen.frx":3418
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmubicaciones_almacen.frx":3CF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmubicaciones_almacen.frx":45CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmubicaciones_almacen.frx":4EA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmubicaciones_almacen.frx":4FBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmubicaciones_almacen.frx":50CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmubicaciones_almacen.frx":51DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmubicaciones_almacen.frx":52F0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   27
      Top             =   285
      Width           =   7920
   End
End
Attribute VB_Name = "frmubicaciones_almacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim numero_items_rutas As Integer
Dim var_tipo_lista As Integer
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
      Call pro_elimina_rutas
      rs.Open "select * from TB_UBICACIONES_ALMACEN WHERE VCHA_ALM_ALMACEN_ID = '" + txt_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
      If rs.BOF Then
         cmd_guardar.Enabled = False
         cmd_deshacer.Enabled = False
         cmd_eliminar.Enabled = False
         txt_articulo = ""
         txt_nombre_articulo = ""
         txt_ubicacion_1 = ""
         txt_ubicacion_2 = ""
         txt_ubicacion_3 = ""
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
   If var_modifica_registro_ubicacion_almacen = False Then
      rs.Open "select * from tb_ubicaciones_almacen where vcha_alm_almacen_id = '" + Me.txt_almacen + "' and vcha_art_articulo_id = '" + txt_articulo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_posible = False
      End If
      rs.Close
   End If
   If var_posible = True Then
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
         Call pro_guardar_ubicaciones_almacen
         rs.Open "select * from tb_ubicaciones_almacen where vcha_alm_almacen_id = '" + txt_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
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
   Else
      MsgBox "La clave de ruta ya existe", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call gPrintListView(lv_articulos, "LISTADO DE rutas")
        End If

End Sub

Private Sub cmd_nuevo_Click()
   txt_articulo = ""
   txt_nombre_articulo = ""
   txt_ubicacion_1 = ""
   txt_ubicacion_2 = ""
   txt_ubicacion_3 = ""
   txt_articulo.Enabled = True
   txt_articulo.SetFocus: var_modifica_registro_ubicacion_almacen = False
   cmd_guardar.Enabled = True
   cmd_deshacer.Enabled = True
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_ubicacion_almacen = False Then
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


Private Sub Form_Activate()
   Call pro_llena_listview1
   Call pro_textos
End Sub

Private Sub Form_Initialize()
   Call pro_llena_listview1
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
   Left = 2300
   frm_lista.Visible = False
   var_modifica_registro_ubicacion_almacen = True
   lv_articulos.SmallIcons = ImageList
   Call pro_encabezadosView(Me, lv_articulos, False)
   If txt_almacen <> "" Then
      rs.Open "select * from tb_ubicaciones_almacen where vcha_alm_almacen_id = " + txt_almacen + ""
      If Not rs.EOF Then
         cmd_guardar.Enabled = True
         cmd_deshacer.Enabled = True
         cmd_eliminar.Enabled = True
      Else
         cmd_guardar.Enabled = False
         cmd_deshacer.Enabled = False
         cmd_eliminar.Enabled = False
      End If
      rs.Close
      Call pro_llena_listview1
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   var_swpassword = False
   var_modifica_registro_ubicacion_almacen = False
   Call activa_forma(var_activa_forma_rutas)
End Sub


Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         txt_articulo = lv_lista.selectedItem
         txt_nombre_articulo = lv_lista.selectedItem.SubItems(1)
      End If
      txt_articulo.SetFocus
      frm_lista.Visible = False
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub lv_articulos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_articulos, ColumnHeader)
End Sub

Private Sub lv_articulos_ItemClick(ByVal item As MSComctlLib.ListItem)
    Set lv_articulos.selectedItem = item
        pro_textos
        var_modifica_registro_ubicacion_almacen = True
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_articulos.SetFocus
      Call pro_avanzar(Me, lv_articulos, Button)
      lv_articulos.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_articulos.ListItems(1).Selected = True
      lv_articulos.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 4 Then
      numero_items_rutas = lv_articulos.ListItems.Count
      lv_articulos.ListItems(numero_items_rutas).Selected = True
      pro_textos
      lv_articulos.selectedItem.EnsureVisible
   End If
err0:
End Sub


Sub pro_guardar_ubicaciones_almacen()
   Dim ok As Boolean
   If txt_almacen <> "" And txt_nombre_almacen <> "" And txt_articulo <> "" And txt_ubicacion_1 <> "" Then
      If var_hubo_cambios Then
         If var_modifica_registro_ubicacion_almacen = False Then
            rs.Open "insert into tb_ubicaciones_almacen (vcha_alm_almacen_id, vcha_art_articulo_id, vcha_ubi_ubicacion_1, vcha_ubi_ubicacion_2, vcha_ubi_ubicacion_3) values ('" + txt_almacen + "','" + txt_articulo + "', '" + txt_ubicacion_1 + "', '" + txt_ubicacion_2 + "', '" + txt_ubicacion_3 + "')", cnn, adOpenDynamic, adLockOptimistic
         Else
            rs.Open "update tb_ubicaciones_almacen set vcha_ubi_ubicacion_1 = '" + txt_ubicacion_1 + "', vcha_ubi_ubicacion_2 = '" + txt_ubicacion_2 + "', vcha_ubi_ubicacion_3 = '" + txt_ubicacion_3 + "' where vcha_alm_almacen_id = '" + txt_almacen + "' and vcha_art_articulo_id = '" + txt_articulo + "'", cnn, adOpenDynamic, adLockOptimistic
         End If
         pro_actualiza_ListView
         txt_almacen.Enabled = False
         MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
         txt_registros = lv_articulos.ListItems.Count
         var_modifica_registro_ubicacion_almacen = True
      End If
   End If
   Set TB_RUTAS = Nothing: var_hubo_cambios = False
End Sub

Sub pro_elimina_rutas()
   Dim var_llave_usuarios As String
   ok = True
   On Error GoTo salir:
   If txt_almacen <> "" And txt_nombre_almacen <> "" And var_modifica_registro_ubicacion_almacen = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         rs.Open "DELETE FROM TB_UBICACIONES_ALMACEN WHERE VCHA_ALM_ALMACEN_ID ='" + txt_almacen + "' AND VCHA_ART_ARTICULO_ID = '" + txt_articulo + "'", cnn, adOpenDynamic, adLockOptimistic
      Else
         GoTo salir:
      End If
      MsgBox "Se Elimino Correctamente el Registro", vbInformation
      lv_articulos.ListItems.Remove (lv_articulos.selectedItem.Index)
      txt_registros = lv_articulos.ListItems.Count
      lv_articulos.selectedItem.Selected = True
      pro_textos
   End If
salir:
End Sub


Sub pro_llena_listview1()
   Dim list_item As ListItem
   numero_items_rutas = 0
   rs.Open "select * from vw_ubicaciones_almacen where vcha_alm_almacen_id = '" + txt_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set list_item = lv_articulos.ListItems.Add(, , rs!VCHA_ART_ARTICULO_ID)
         list_item.SubItems(1) = IIf(IsNull(rs!vcha_Art_nombre_español), "", rs!vcha_Art_nombre_español)
         list_item.SubItems(2) = IIf(IsNull(rs!vcha_ubi_ubicacion_1), "", rs!vcha_ubi_ubicacion_1)
         list_item.SubItems(3) = IIf(IsNull(rs!vcha_ubi_ubicacion_2), "", rs!vcha_ubi_ubicacion_2)
         list_item.SubItems(4) = IIf(IsNull(rs!vcha_ubi_ubicacion_3), "", rs!vcha_ubi_ubicacion_3)
         rs.MoveNext:
         numero_items_rutas = numero_items_rutas + 1
   Wend
   rs.Close
End Sub


Sub pro_textos()
'On Error GoTo err0:
Dim var_n As Double
   var_n = lv_articulos.ListItems.Count
   Me.txt_almacen.Enabled = False
   If var_n > 0 Then
      txt_articulo = lv_articulos.selectedItem
      txt_nombre_articulo = lv_articulos.selectedItem.SubItems(1)
      txt_ubicacion_1 = lv_articulos.selectedItem.SubItems(2)
      txt_ubicacion_2 = lv_articulos.selectedItem.SubItems(3)
      txt_ubicacion_3 = lv_articulos.selectedItem.SubItems(4)
      rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_articulo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_nombre = IIf(IsNull(rs!VCHA_ART_NOMBRE_ESPAÑOL), "", rs!VCHA_ART_NOMBRE_ESPAÑOL)
      Else
         txt_nombre_nombre = ""
      End If
      rs.Close
   End If
   var_numero_renglones = lv_articulos.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_articulos.ColumnHeaders(2).Width = 6000
   Else
      lv_articulos.ColumnHeaders(2).Width = 6200
   End If
   var_hubo_cambios = False
   var_modifica_registro_ubicacion_almacen = True
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro_ubicacion_almacen = False Then
        Set list_item = lv_articulos.ListItems.Add(, , txt_articulo)
        list_item.SubItems(1) = txt_nombre_articulo
        list_item.SubItems(2) = txt_ubicacion_1
        list_item.SubItems(3) = txt_ubicacion_2
        list_item.SubItems(4) = txt_ubicacion_3
        list_item.EnsureVisible
        list_item.Selected = True
        numero_items_rutas = numero_items_rutas + 1
    Else
        lv_articulos.ListItems.item(lv_articulos.selectedItem.Index).Checked = False
        lv_articulos.ListItems.item(lv_articulos.selectedItem.Index) = txt_articulo
        lv_articulos.ListItems.item(lv_articulos.selectedItem.Index).ListSubItems(1) = txt_nombre_articulo
        lv_articulos.ListItems.item(lv_articulos.selectedItem.Index).ListSubItems(2) = txt_ubicacion_1
        lv_articulos.ListItems.item(lv_articulos.selectedItem.Index).ListSubItems(3) = txt_ubicacion_2
        lv_articulos.ListItems.item(lv_articulos.selectedItem.Index).ListSubItems(4) = txt_ubicacion_3
        lv_articulos.ListItems.item(lv_articulos.selectedItem.Index).Selected = True
    End If
    lv_articulos.SetFocus
End Sub








Private Sub txt_articulo_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_articulo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select VCHA_ART_ARTICULO_ID, vcha_art_nombre_español from tb_articulos order by vcha_art_nombre_español", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ART_ARTICULO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_Art_nombre_español), "", rs!vcha_Art_nombre_español)
            rs.MoveNext
      Wend
      rs.Close
      frm_lista.Visible = True
      var_origen_codigo = 1
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_articulo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_articulo_LostFocus()
   Dim var_posible As Boolean
   If Trim(txt_articulo) <> "" Then
      var_posible = False
      rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_articulo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_posible = True
         txt_nombre_articulo = rs!VCHA_ART_NOMBRE_ESPAÑOL
            If Not rs.EOF Then
               var_posible = True
               txt_articulo = rs!VCHA_ART_ARTICULO_ID
               txt_nombre_articulo = rs!vcha_Art_nombre_español
               rsaux4.Open "select * from vw_ubicaciones_almacen where vcha_alm_almacen_id = '" + txt_almacen + "' and vcha_art_articulo_id = '" + txt_articulo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux4.EOF Then
                 txt_ubicacion_1 = IIf(IsNull(rsaux4!vcha_ubi_ubicacion_1), "", rsaux4!vcha_ubi_ubicacion_1)
                 txt_ubicacion_2 = IIf(IsNull(rsaux4!vcha_ubi_ubicacion_2), "", rsaux4!vcha_ubi_ubicacion_2)
                 txt_ubicacion_3 = IIf(IsNull(rsaux4!vcha_ubi_ubicacion_3), "", rsaux4!vcha_ubi_ubicacion_3)
                 var_hubo_cambios = False
                 var_modifica_registro_ubicacion_almacen = True
              Else
                  txt_ubicacion_1 = ""
                  txt_ubicacion_2 = ""
                  txt_ubicacion_3 = ""
               End If
               rsaux4.Close
               rs.Close
            Else
               var_posible = False
               rs.Close
            End If
      Else
         rs.Close
         rs.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + txt_articulo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rs!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               var_posible = True
               txt_articulo = rs!VCHA_ART_ARTICULO_ID
               txt_nombre_articulo = rsaux!vcha_Art_nombre_español
               rsaux4.Open "select * from vw_ubicaciones_almacen where vcha_alm_almacen_id = '" + txt_almacen + "' and vcha_art_articulo_id = '" + txt_articulo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux4.EOF Then
                 txt_ubicacion_1 = IIf(IsNull(rsaux4!vcha_ubi_ubicacion_1), "", rsaux4!vcha_ubi_ubicacion_1)
                 txt_ubicacion_2 = IIf(IsNull(rsaux4!vcha_ubi_ubicacion_2), "", rsaux4!vcha_ubi_ubicacion_2)
                 txt_ubicacion_3 = IIf(IsNull(rsaux4!vcha_ubi_ubicacion_3), "", rsaux4!vcha_ubi_ubicacion_3)
                 var_hubo_cambios = False
                 var_modifica_registro_ubicacion_almacen = True
              Else
                  txt_ubicacion_1 = ""
                  txt_ubicacion_2 = ""
                  txt_ubicacion_3 = ""
               End If
               rsaux4.Close
               rsaux.Close
               rs.Close
            Else
               var_posible = False
               rsaux.Close
               rs.Close
            End If
         Else
            rs.Close
         End If
         If var_posible = False Then
            MsgBox "Clave de artículo incorrecta", vbOKOnly, "ATENCION"
            txt_articulo = ""
            txt_nombre_articulo = ""
            txt_ubicacion_1 = ""
            txt_ubicacion_2 = ""
            txt_ubicacion_3 = ""
         End If
      End If
   Else
      txt_articulo = ""
      txt_nombre_articulo = ""
      txt_ubicacion_1 = ""
      txt_ubicacion_2 = ""
      txt_ubicacion_3 = ""
   End If
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim var_posible As Boolean
      If Trim(txt_buscar) <> "" Then
         var_posible = False
         txt_buscar = UCase(Me.txt_buscar)
         rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_buscar + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_posible = True
            rs.Close
            Call pro_busca_registro(Me.lv_articulos, txt_buscar, False)
            txt_buscar = ""
            pro_textos
         Else
            rs.Close
            rs.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + txt_buscar + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + rs!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  var_posible = True
                  txt_buscar = rs!VCHA_ART_ARTICULO_ID
                  rsaux.Close
                  rs.Close
               Else
                  var_posible = False
                  rsaux.Close
                  rs.Close
               End If
            Else
               rs.Close
            End If
            If var_posible = True Then
               Call pro_busca_registro(Me.lv_articulos, txt_buscar, False)
               txt_buscar = ""
               pro_textos
            End If
         End If
      End If
   End If
End Sub



Private Sub txt_nombre_almacen_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_almacen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub


Private Sub txt_almacen_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_almacen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub


Private Sub txt_nombre_articulo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select VCHA_ART_ARTICULO_ID, vcha_art_nombre_español from tb_articulos order by vcha_art_nombre_español", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ART_ARTICULO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_Art_nombre_español), "", rs!vcha_Art_nombre_español)
            rs.MoveNext
      Wend
      rs.Close
      frm_lista.Visible = True
      var_origen_codigo = 1
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_articulo_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_ubicacion_1_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_ubicacion_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_ubicacion_2_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_ubicacion_2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_ubicacion_3_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_ubicacion_3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub
