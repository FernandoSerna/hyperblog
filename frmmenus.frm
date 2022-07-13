VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmenus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menus"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   Icon            =   "frmmenus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11670
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11295
      Picture         =   "frmmenus.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Salir"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1410
      Picture         =   "frmmenus.frx":0F04
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1080
      Picture         =   "frmmenus.frx":1006
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   750
      Picture         =   "frmmenus.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   420
      Picture         =   "frmmenus.frx":11DA
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Guardar Alt + G"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmmenus.frx":12DC
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   45
      Width           =   330
   End
   Begin VB.Frame frm_submenus 
      Height          =   3090
      Left            =   3615
      TabIndex        =   11
      Top             =   1770
      Width           =   4755
      Begin VB.CommandButton cmd_cancelar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   375
         Picture         =   "frmmenus.frx":13DE
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Cancelar Alt + C"
         Top             =   285
         Width           =   330
      End
      Begin VB.CommandButton cmd_aceptar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   45
         Picture         =   "frmmenus.frx":1528
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Aceptar Alt + A"
         Top             =   285
         Width           =   330
      End
      Begin VB.TextBox txt_numero 
         Height          =   315
         Left            =   900
         TabIndex        =   29
         Top             =   2670
         Width           =   930
      End
      Begin VB.CheckBox chk_permiso4 
         Caption         =   "Permiso mancomunado 2"
         Height          =   210
         Left            =   2310
         TabIndex        =   27
         Top             =   2430
         Width           =   2265
      End
      Begin VB.CheckBox chk_permiso3 
         Caption         =   "Permiso 2"
         Height          =   210
         Left            =   900
         TabIndex        =   26
         Top             =   2445
         Width           =   1020
      End
      Begin VB.CheckBox chk_permiso2 
         Caption         =   "Permiso mancomunado 1"
         Height          =   210
         Left            =   2310
         TabIndex        =   25
         Top             =   2145
         Width           =   2265
      End
      Begin VB.CheckBox chk_permiso1 
         Caption         =   "Permiso 1"
         Height          =   210
         Left            =   915
         TabIndex        =   24
         Top             =   2160
         Width           =   1020
      End
      Begin VB.ComboBox cmb_formas 
         Height          =   315
         Left            =   915
         TabIndex        =   23
         Top             =   1785
         Width           =   3690
      End
      Begin VB.TextBox txt_accion 
         Height          =   285
         Left            =   915
         TabIndex        =   17
         Top             =   1800
         Width           =   795
      End
      Begin VB.TextBox txt_nivel 
         Enabled         =   0   'False
         Height          =   315
         Left            =   915
         TabIndex        =   16
         Top             =   1440
         Width           =   780
      End
      Begin VB.TextBox txt_titulo 
         Height          =   300
         Left            =   915
         TabIndex        =   15
         Top             =   1110
         Width           =   3660
      End
      Begin VB.TextBox txt_clave 
         Enabled         =   0   'False
         Height          =   300
         Left            =   915
         TabIndex        =   14
         Top             =   780
         Width           =   1320
      End
      Begin VB.Frame Frame2 
         Height          =   120
         Left            =   30
         TabIndex        =   22
         Top             =   540
         Width           =   4695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   28
         Top             =   2715
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Acción:"
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   21
         Top             =   1860
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nivel:"
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   20
         Top             =   1500
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Titulo:"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   19
         Top             =   1185
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   18
         Top             =   855
         Width           =   450
      End
      Begin VB.Label lbl_submenu 
         BackColor       =   &H8000000D&
         Caption         =   " Submenu"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   0
         TabIndex        =   12
         Top             =   15
         Width           =   4740
      End
   End
   Begin VB.Frame Frame5 
      Height          =   6840
      Left            =   5715
      TabIndex        =   9
      Top             =   420
      Width           =   5925
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   6360
         Left            =   45
         TabIndex        =   10
         Top             =   420
         Width           =   5790
         _ExtentX        =   10213
         _ExtentY        =   11218
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   3
         HotTracking     =   -1  'True
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Estructura del menu"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   30
         TabIndex        =   13
         Top             =   120
         Width           =   5850
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Menus"
      Height          =   975
      Left            =   30
      TabIndex        =   1
      Top             =   420
      Width           =   5655
      Begin VB.TextBox txt_menus 
         Height          =   315
         Index           =   0
         Left            =   915
         MaxLength       =   2
         TabIndex        =   3
         Top             =   240
         Width           =   1245
      End
      Begin VB.TextBox txt_menus 
         Height          =   315
         Index           =   1
         Left            =   915
         TabIndex        =   2
         Top             =   585
         Width           =   4620
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   300
         Width           =   450
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   4
         Top             =   645
         Width           =   600
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1245
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   5700
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   45
      Top             =   5430
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
            Picture         =   "frmmenus.frx":1672
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmenus.frx":1F4C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   690
      Top             =   5415
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmenus.frx":2826
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmenus.frx":3100
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmenus.frx":39DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmenus.frx":3F76
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmenus.frx":4852
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmenus.frx":512C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmenus.frx":5A06
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmenus.frx":5B18
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmenus.frx":5C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmenus.frx":5D3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmenus.frx":5E4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmenus.frx":5F60
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmenus.frx":64A2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   5895
      Left            =   30
      TabIndex        =   7
      Top             =   1365
      Width           =   5670
      Begin MSComctlLib.ListView lv_menus 
         Height          =   5700
         Left            =   45
         TabIndex        =   8
         Top             =   135
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   10054
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList"
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
            Text            =   "pais"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "estado"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "telefono"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   0
      TabIndex        =   6
      Top             =   285
      Width           =   11655
   End
   Begin VB.Menu mnu_opciones2 
      Caption         =   "opciones 2"
      Visible         =   0   'False
      Begin VB.Menu mnu_insertar_menu 
         Caption         =   "Insertar Menu"
      End
   End
   Begin VB.Menu mnu_opciones 
      Caption         =   "opciones"
      Visible         =   0   'False
      Begin VB.Menu mnu_insertar 
         Caption         =   "Insertar Opción"
      End
      Begin VB.Menu mnu_modificar 
         Caption         =   "Modificar Opción"
      End
      Begin VB.Menu mnu_eliminar 
         Caption         =   "Eliminar Opción"
      End
   End
End
Attribute VB_Name = "frmmenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim numero_items_menus As Integer
Dim bitacora As Boolean
Dim var_nivel As Integer
Dim var_opcion As Integer
Dim var_clave_nivel1 As Integer
Dim var_clave_nivel1_s As String
Dim var_clave_nivel2 As Integer
Dim var_clave_nivel2_s As String
Dim var_clave_nivel3 As Integer
Dim var_clave_nivel3_s As String
Dim var_clave_nivel4 As Integer
Dim var_clave_nivel4_s As String
Dim var_clave_nivel5 As Integer
Dim var_clave_nivel5_s As String
Dim var_nombre_submenu As String
Dim var_accion_submenu As String
Dim var_permiso1 As Integer
Dim var_permiso2 As Integer
Dim var_permiso3 As Integer
Dim var_permoso4 As Integer
Dim var_numero_submenu As Integer





Private Sub cmb_formas_Click()
   'txt_accion = Obtener_llave(cnn, rs, "TB_FORMAS", "VCHA_FOR_NOMBRE", cmb_formas, 0, "T")
    rs.Open "select * from tb_formas where vcha_for_nombre = '" + cmb_formas + "'", cnn, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
          txt_accion = IIf(IsNull(rs(0).Value), "", rs(0).Value)
    End If
    rs.Close

End Sub

Private Sub cmd_aceptar_Click()
   Set TB_SUBMENUS_INSERTA = New TB_SUBMENUS_INSERTA
   Set TB_SUBMENUS_MODIFICA = New TB_SUBMENUS_MODIFICA
   Set TB_SUBMENUS_ELIMINA = New TB_SUBMENUS_ELIMINA
   If Not IsNumeric(Me.txt_numero) Then
      txt_numero = 0
   End If
   If var_opcion = 1 Then
      rs.Open "select * from tb_submenus where char_sme_submenu_id = '" + txt_clave + "'", cnn, adOpenDynamic, adLockOptimistic
      If rs.EOF Then
         var_inserta = TB_SUBMENUS_INSERTA.Anadir(txt_menus(0), txt_clave, var_clave_nivel1, var_clave_nivel2, var_clave_nivel3, var_clave_nivel4, var_clave_nivel5, txt_titulo, 1, txt_accion, chk_permiso1, chk_permiso2, chk_permiso3, chk_permiso4, Val(txt_numero))
      Else
         MsgBox "Ya existe el submenu", vbOKOnly, "ATENCION"
      End If
      rs.Close
      pro_textos
   End If
   If var_opcion = 2 Then
      rs.Open "select * from tb_submenus where char_sme_submenu_id = '" + txt_clave + "'", cnn, adOpenDynamic, adLockOptimistic
      If rs.EOF Then
         var_inserta = TB_SUBMENUS_INSERTA.Anadir(txt_menus(0), txt_clave, var_clave_nivel1, var_clave_nivel2, var_clave_nivel3, var_clave_nivel4, var_clave_nivel5, txt_titulo, var_nivel + 1, txt_accion, chk_permiso1, chk_permiso2, chk_permiso3, chk_permiso4, Val(txt_numero))
      Else
         MsgBox "Ya existe el submenu", vbOKOnly, "ATENCION"
      End If
      rs.Close
      pro_textos
   End If
   If var_opcion = 3 Then
      rs.Open "select * from tb_submenus where char_sme_submenu_id = '" + txt_clave + "'", cnn, adOpenDynamic, adLockOptimistic
      If rs.EOF Then
         MsgBox "No existe el submenu", vbOKOnly, "ATENCION"
      Else
         var_inserta = TB_SUBMENUS_MODIFICA.Anadir(txt_menus(0), txt_clave, var_clave_nivel1, var_clave_nivel2, var_clave_nivel3, var_clave_nivel4, var_clave_nivel5, txt_titulo, var_nivel, txt_accion, chk_permiso1, chk_permiso2, chk_permiso3, chk_permiso4, Val(txt_numero))
      End If
      rs.Close
      pro_textos
   End If
   If var_opcion = 4 Then
      si = MsgBox("¿Se desea eliminar el submenu y todas sus dependencias?", vbOKCancel, "ATENCION")
      If si = 1 Then
         If var_nivel = 1 Then
            rs.Open "select * from tb_submenus where VCHA_MEN_MENU_ID = '" + txt_menus(0) + "' and INTE_SME_NIVEL1 = " + Str(var_clave_nivel1), cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               While Not rs.EOF
                  var_inserta = TB_SUBMENUS_ELIMINA.Anadir(rs(1).Value)
                  rs.MoveNext
               Wend
            End If
            rs.Close
         End If
         If var_nivel = 2 Then
            rs.Open "select * from tb_submenus where VCHA_MEN_MENU_ID = '" + txt_menus(0) + "' and INTE_SME_NIVEL1 = " + Str(var_clave_nivel1) + " and INTE_SME_NIVEL2 = " + Str(var_clave_nivel2), cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               While Not rs.EOF
                  var_inserta = TB_SUBMENUS_ELIMINA.Anadir(rs(1).Value)
                  rs.MoveNext
               Wend
            End If
            rs.Close
         End If
         If var_nivel = 3 Then
            rs.Open "select * from tb_submenus where VCHA_MEN_MENU_ID = '" + txt_menus(0) + "' and INTE_SME_NIVEL1 = " + Str(var_clave_nivel1) + " and INTE_SME_NIVEL2 = " + Str(var_clave_nivel2) + " and INTE_SME_NIVEL3 = " + Str(var_clave_nivel3), cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               While Not rs.EOF
                  var_inserta = TB_SUBMENUS_ELIMINA.Anadir(rs(1).Value)
                  rs.MoveNext
               Wend
            End If
            rs.Close
         End If
         If var_nivel = 4 Then
            var_inserta = TB_SUBMENUS_ELIMINA.Anadir(txt_clave)
         End If
      End If
      pro_textos
   End If
   frm_submenus.Visible = False

End Sub

Private Sub cmd_cancelar_Click()
   Set TB_SUBMENUS_INSERTA = New TB_SUBMENUS_INSERTA
   Set TB_SUBMENUS_MODIFICA = New TB_SUBMENUS_MODIFICA
   Set TB_SUBMENUS_ELIMINA = New TB_SUBMENUS_ELIMINA
      frm_submenus.Visible = False
      txt_clave = ""
      txt_titulo = ""
      txt_nivel = ""
      txt_accion = ""
      chk_permiso1 = 0
      chk_permiso2 = 0
      chk_permiso3 = 0
      chk_permiso4 = 0
   frm_submenus.Visible = False

End Sub

Private Sub cmd_deshacer_Click()
   frm_submenus.Visible = False
   Call pro_textos
End Sub

Private Sub cmd_eliminar_Click()
   frm_submenus.Visible = False
   Call pro_elimina_menus
   rs.Open "select * from tb_menus", cnn, adOpenDynamic, adLockOptimistic
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

Private Sub cmd_guardar_Click()
   frm_submenus.Visible = False
   Call pro_guardar_menus
   If rs.State = 1 Then
      rs.Close
   End If
   rs.Open "select * from tb_menus", cnn, adOpenDynamic, adLockOptimistic
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

Private Sub cmd_imprimir_Click()
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call gPrintListView(lv_menus, "LISTADO DE menus")
        End If

End Sub

Private Sub cmd_nuevo_Click()
   frm_submenus.Visible = False
   Call pro_limpiatextos(Me)
   txt_menus(0).Enabled = True
   txt_menus(0).SetFocus: var_modifica_registro_menu = False
   cmd_guardar.Enabled = True
   cmd_deshacer.Enabled = True
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

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 0
   Left = 0
    frm_submenus.Visible = False
    var_modifica_registro_menu = True
    lv_menus.SmallIcons = ImageList
    
    Call pro_encabezadosView(Me, lv_menus, False)
    Call pro_llena_listview1
    pro_textos

    rs.Open "select * from tb_menus", cnn, adOpenDynamic, adLockOptimistic
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
   var_swpassword = False
   var_modifica_registro_menu = False
   Call activa_forma(var_activa_forma_menus)
End Sub

Private Sub lv_menus_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_menus, ColumnHeader)
End Sub

Private Sub lv_menus_GotFocus()
   frm_submenus.Visible = False
End Sub

Private Sub lv_menus_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_menus.selectedItem = Item
        pro_textos
        var_modifica_registro_menu = True
        txt_menus(0).Enabled = False

End Sub


Private Sub lv_menus_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Button = 2 Then
      PopupMenu mnu_opciones2
   End If
End Sub

Sub pro_guardar_menus()
Dim ok As Boolean
   Set TB_MENUS = New TB_MENUS
   Set TB_BITACORA_MENUS = New TB_BITACORA_MONEDA
   ok = True
   If txt_menus(0) <> "" And txt_menus(1) <> "" Then
      If var_hubo_cambios Then
         rs.Open "select * from tb_menus where vcha_men_menu_id = '" + txt_menus(0) + "'", cnn, adOpenDynamic, adLockOptimistic
         ok = TB_MENUS.Anadir(txt_menus(0), txt_menus(1))
         If ok Then
            bitacora = True
            If var_modifica_registro_menu = False Then
               var_operacion_bitacora = "I"
               bitacora = TB_BITACORA_MENUS.Anadir(txt_menus(0), "VCHA_MEN_DESCRIPCION", var_operacion_bitacora, "", txt_menus(1), var_clave_usuario_global, fun_NombrePc, Date)
            Else
               var_operacion_bitacora = "M"
               If rs(0) <> txt_menus(0) Then
                  bitacora = TB_BITACORA_MENUS.Anadir(txt_menus(0), "VCHA_MEN_MENU_ID", var_operacion_bitacora, rs(0), txt_menus(0), var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(1) <> txt_menus(1) Then
                  bitacora = TB_BITACORA_MENUS.Anadir(txt_menus(0), "VCHA_MEN_DESCRIPCION", var_operacion_bitacora, rs(1), txt_menus(1), var_clave_usuario_global, fun_NombrePc, Date)
               End If
            End If
            rs.Close
            pro_actualiza_ListView
            txt_menus(0).Enabled = False
            MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
            txt_registros = lv_menus.ListItems.Count
            var_modifica_registro_menu = True
         Else
            MsgBox "No se puede grabar registro: " + TB_MENUS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
         End If
      End If
   End If
   Set TB_MENUS = Nothing: var_hubo_cambios = False
End Sub

Sub pro_elimina_menus()
   Dim var_llave_usuarios As String
   Set TB_MENUS = New TB_MENUS
   Set TB_BITACORA_MENUS = New TB_BITACORA_MONEDA
   On Error GoTo salir:
   ok = True
   If txt_menus(0) <> "" And txt_menus(1) <> "" And var_modifica_registro_menu = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_MENUS.Eliminar(txt_menus(0))
      Else
         GoTo salir:
      End If
      If ok Then
         var_operacion_bitacora = "E"
         bitacora = TB_BITACORA_MENUS.Anadir(txt_menus(0), "VCHA_MEN_DESCRIPCION", var_operacion_bitacora, txt_menus(1), "", var_clave_usuario_global, fun_NombrePc, Date)
         numero_items_menus = numero_items_menus - 1
         MsgBox "Se Elimino Correctamente el Registro", vbInformation
         lv_menus.ListItems.Remove (lv_menus.selectedItem.Index)
         Call pro_limpiatextos(Me)
         txt_registros = lv_menus.ListItems.Count
         lv_menus.selectedItem.Selected = True
         pro_textos
      Else
         MsgBox "No se puede grabar registro: " + TB_MENUS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
   Set TB_MENUS = Nothing
End Sub


Sub pro_llena_listview1()

Dim list_item As ListItem

   rs.Open "select * from TB_menus", cnn, adOpenDynamic, adLockOptimistic
   numero_items_menus = 0
   While Not rs.EOF
        Set list_item = lv_menus.ListItems.Add(, , rs(0).Value)
        list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
    rs.MoveNext:
    numero_items_menus = numero_items_menus + 1
    Wend
    rs.Close

End Sub


Sub pro_textos()
Dim nodX As Node
   If lv_menus.ListItems.Count > 0 Then
        txt_menus(0) = lv_menus.selectedItem
        txt_menus(1) = lv_menus.selectedItem.SubItems(1)
        rsaux.Open "select * from tb_submenus where vcha_men_menu_id = '" + txt_menus(0) + "' order by vcha_men_menu_id,inte_sme_nivel,inte_sme_numero,char_sme_submenu_id", cnn, adOpenDynamic, adLockOptimistic
        TreeView1.Nodes.Clear
        If Not rsaux.EOF Then
           While Not rsaux.EOF
              var_n = rsaux(8).Value
              If var_n = 1 Then
                 var_c = Trim(Mid(rsaux(1).Value, 1, 4))
                 Set nodX = TreeView1.Nodes.Add(, , """" + var_c + """", "" + rsaux(7).Value + "")
              End If
              If var_n = 2 Then
                 var_c2 = Trim(Mid(rsaux(1).Value, 1, 4))
                 var_c3 = Trim(Mid(rsaux(1).Value, 1, 6))
                 Set nodX = TreeView1.Nodes.Add("""" + var_c2 + """", tvwChild, """" + var_c3 + """", "" + rsaux(7).Value + "")
              End If
              If var_n = 3 Then
                 var_c3 = Trim(Mid(rsaux(1).Value, 1, 6))
                 var_c4 = Trim(Mid(rsaux(1).Value, 1, 8))
                 Set nodX = TreeView1.Nodes.Add("""" + var_c3 + """", tvwChild, """" + var_c4 + """", "" + rsaux(7).Value + "")
              End If
              If var_n = 4 Then
                 var_c4 = Trim(Mid(rsaux(1).Value, 1, 8))
                 var_c5 = Trim(Mid(rsaux(1).Value, 1, 10))
                 Set nodX = TreeView1.Nodes.Add("""" + var_c4 + """", tvwChild, """" + var_c5 + """", "" + rsaux(7).Value + "")
              End If
              rsaux.MoveNext:
           Wend
        End If
        rsaux.Close
        TreeView1.Style = 7
    End If
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro_menu = False Then
        Set list_item = lv_menus.ListItems.Add(, , txt_menus(0))
        list_item.SubItems(1) = txt_menus(1)
        list_item.EnsureVisible
        list_item.Selected = True
        numero_items_menus = numero_items_menus + 1
    Else
        lv_menus.ListItems.Item(lv_menus.selectedItem.Index).Checked = False
        lv_menus.ListItems.Item(lv_menus.selectedItem.Index) = txt_menus(0)
        lv_menus.ListItems.Item(lv_menus.selectedItem.Index).ListSubItems(1) = txt_menus(1)
        lv_menus.ListItems.Item(lv_menus.selectedItem.Index).Selected = True
    End If
    lv_menus.SetFocus
End Sub


Private Sub TreeView1_GotFocus()
   Me.frm_submenus.Visible = False
End Sub

Private Sub txt_accion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      frm_submenus.Visible = False
   End If
End Sub

Private Sub txt_clave_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 27 Then
      frm_submenus.Visible = False
   End If
End Sub

Private Sub txt_menus_Change(Index As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub

Private Sub txt_menus_GotFocus(Index As Integer)
   frm_submenus.Visible = False
End Sub

Private Sub txt_menus_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Index < 1 Then
          txt_menus(Index + 1).SetFocus
       Else
          txt_menus(0).Enabled = True
          txt_menus(0).SetFocus
       End If
    Else
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
       var_hubo_cambios = True
    End If
    
End Sub


Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim var_c As String
   If Button = 2 Then
      PopupMenu mnu_opciones
   End If
End Sub

Private Sub txt_menus_LostFocus(Index As Integer)
   If Len(Trim(txt_menus(0))) = 1 Then
      txt_menus(0) = "0" + Trim(txt_menus(0))
   End If
End Sub
Private Sub mnu_insertar_click()
   var_numero_submenu = 0
   var_opcion = 2
   var_clave_nivel1 = 0
   var_clave_nivel2 = 0
   var_clave_nivel3 = 0
   var_clave_nivel4 = 0
   var_clave_nivel5 = 0
   var_permiso1 = 0
   var_permiso2 = 0
   var_permiso3 = 0
   var_permiso4 = 0
   var_c = TreeView1.selectedItem.Key
   var_longitud = Len(Trim(var_c))
   If var_longitud = 6 Then
      var_c = Trim(Mid(var_c, 2, 4))
   End If
   If var_longitud = 8 Then
      var_c = Trim(Mid(var_c, 2, 6))
   End If
   If var_longitud = 10 Then
      var_c = Trim(Mid(var_c, 2, 8))
   End If
   If var_longitud = 12 Then
      var_c = Trim(Mid(var_c, 2, 10))
   End If
   var_longitud = Len(Trim(var_c))
   If var_longitud = 4 Then
      var_c2 = var_c + "000000"
      rs.Open "select * from tb_submenus where char_sme_submenu_id = '" + var_c2 + "'", cnn, adOpenDynamic, adLockOptimistic
      var_nivel = rs(8).Value
      var_nivel1 = rs(2).Value
      var_clave_nivel1 = rs(2).Value
      var_nombre_submenu = rs(7).Value
      var_permiso1 = rs(10).Value
      var_permiso2 = rs(11).Value
      var_permiso3 = rs(12).Value
      var_permiso4 = rs(13).Value
      var_numero_submenu = rs(14).Value
      txt_numero = rs(14).Value
      rs.Close
      rs.Open "select max(inte_sme_nivel2) from tb_submenus where vcha_men_menu_id = '" + txt_menus(0) + "' and inte_sme_nivel1 = " + Str(var_nivel1), cnn, adOpenDynamic, adLockOptimistic
      If IsNull(rs(0).Value) Then
         txt_clave = var_c + "010000"
         var_clave_nivel2 = 1
      Else
         If rs(0).Value + 1 < 10 Then
            txt_clave = var_c + "0" + Trim(Str(rs(0).Value) + 1) + "0000"
            var_clave_nivel2 = rs(0).Value + 1
         Else
            txt_clave = var_c + Trim(Str(rs(0).Value) + 1) + "0000"
            var_clave_nivel2 = rs(0).Value + 1
         End If
      End If
      rs.Close
   End If
   If var_longitud = 6 Then
      var_c2 = var_c + "0000"
      rs.Open "select * from tb_submenus where char_sme_submenu_id = '" + var_c2 + "'", cnn, adOpenDynamic, adLockOptimistic
      var_nivel = rs(8).Value
      var_nivel1 = rs(2).Value
      var_nivel2 = rs(3).Value
      var_clave_nivel1 = rs(2).Value
      var_clave_nivel2 = rs(3).Value
      var_nombre_submenu = rs(7).Value
      var_permiso1 = rs(10).Value
      var_permiso2 = rs(11).Value
      var_permiso3 = rs(12).Value
      var_permiso4 = rs(13).Value
      var_numero_submenu = rs(14).Value
      txt_numero = rs(14).Value
      rs.Close
      rs.Open "select max(inte_sme_nivel3) from tb_submenus where vcha_men_menu_id = '" + txt_menus(0) + "' and inte_sme_nivel1 = " + Str(var_nivel1) + " and inte_sme_nivel2 = " + Str(var_nivel2), cnn, adOpenDynamic, adLockOptimistic
      If IsNull(rs(0).Value) Then
         txt_clave = var_c + "0100"
         var_clave_nivel3 = 1
      Else
         If rs(0).Value + 1 < 10 Then
            txt_clave = var_c + "0" + Trim(Str(rs(0).Value) + 1) + "00"
            var_clave_nivel3 = rs(0).Value + 1
         Else
            txt_clave = var_c + Trim(Str(rs(0).Value) + 1) + "00"
            var_clave_nivel3 = rs(0).Value + 1
         End If
      End If
      rs.Close
   End If
   If var_longitud = 8 Then
      var_c2 = var_c + "00"
      rs.Open "select * from tb_submenus where char_sme_submenu_id = '" + var_c2 + "'", cnn, adOpenDynamic, adLockOptimistic
      var_nivel = rs(8).Value
      var_nivel1 = rs(2).Value
      var_nivel2 = rs(3).Value
      var_nivel3 = rs(4).Value
      var_nombre_submenu = rs(7).Value
      var_permiso1 = rs(10).Value
      var_permiso2 = rs(11).Value
      var_permiso3 = rs(12).Value
      var_permiso4 = rs(13).Value
      var_numero_submenu = rs(14).Value
      txt_numero = rs(14).Value
      rs.Close
      rs.Open "select max(inte_sme_nivel4) from tb_submenus where vcha_men_menu_id = '" + txt_menus(0) + "' and inte_sme_nivel1 = " + Str(var_nivel1) + " and inte_sme_nivel2 = " + Str(var_nivel2) + " and inte_sme_nivel3 =" + Str(var_nivel3), cnn, adOpenDynamic, adLockOptimistic
      If IsNull(rs(0).Value) Then
         txt_clave = var_c + "01"
         var_clave_nivel5 = 1
      Else
         If rs(0).Value + 1 < 10 Then
            txt_clave = var_c + "0" + Trim(Str(rs(0).Value) + 1)
            var_clave_nivel5 = rs(0).Value + 1
         Else
            txt_clave = var_c + Trim(Str(rs(0).Value) + 1)
            var_clave_nivel5 = rs(0).Value + 1
         End If
      End If
      rs.Close
   End If
   txt_numero = "0"
   txt_titulo = ""
   txt_nivel = var_nivel + 1
   txt_accion = ""
   chk_permiso1 = var_permiso1
   chk_permiso2 = var_permiso2
   chk_permiso3 = var_permiso3
   chk_permiso4 = var_permiso4
   rs.Open "select * from TB_FORMAS order by vcha_for_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
   Call RecsetToCombo(cmb_formas.hwnd, rs, 1)
   rs.Close
   lbl_submenu.Caption = "Insertar dentro de " + Trim(var_nombre_submenu)
   frm_submenus.Visible = True
   txt_titulo.SetFocus
End Sub

Private Sub mnu_modificar_click()
   var_opcion = 3
   var_clave_nivel1 = 0
   var_clave_nivel2 = 0
   var_clave_nivel3 = 0
   var_clave_nivel4 = 0
   var_clave_nivel5 = 0
   var_c = TreeView1.selectedItem.Key
   var_longitud = Len(Trim(var_c))
   If var_longitud = 6 Then
      var_c = Trim(Mid(var_c, 2, 4))
   End If
   If var_longitud = 8 Then
      var_c = Trim(Mid(var_c, 2, 6))
   End If
   If var_longitud = 10 Then
      var_c = Trim(Mid(var_c, 2, 8))
   End If
   If var_longitud = 12 Then
      var_c = Trim(Mid(var_c, 2, 10))
   End If
   var_longitud = Len(Trim(var_c))
   If var_longitud = 4 Then
      var_c2 = var_c + "000000"
      rs.Open "select * from tb_submenus where char_sme_submenu_id = '" + var_c2 + "'", cnn, adOpenDynamic, adLockOptimistic
      var_nivel = rs(8).Value
      var_nivel1 = rs(2).Value
      var_clave_nivel1 = rs(2).Value
      var_nombre_submenu = rs(7).Value
      var_accion_submenu = rs(9).Value
      var_permiso1 = rs(10).Value
      var_permiso2 = rs(11).Value
      var_permiso3 = rs(12).Value
      var_permiso4 = rs(13).Value
      var_numero_submenu = rs(14).Value
      txt_numero = var_numero_submenu
      rs.Close
   End If
   If var_longitud = 6 Then
      var_c2 = var_c + "0000"
      rs.Open "select * from tb_submenus where char_sme_submenu_id = '" + var_c2 + "'", cnn, adOpenDynamic, adLockOptimistic
      var_nivel = rs(8).Value
      var_nivel1 = rs(2).Value
      var_nivel2 = rs(3).Value
      var_clave_nivel1 = rs(2).Value
      var_clave_nivel2 = rs(3).Value
      var_accion_submenu = rs(9).Value
      var_nombre_submenu = rs(7).Value
      var_permiso1 = rs(10).Value
      var_permiso2 = rs(11).Value
      var_permiso3 = rs(12).Value
      var_permiso4 = rs(13).Value
      var_numero_submenu = rs(14).Value
      txt_numero = var_numero_submenu
      rs.Close
   End If
   If var_longitud = 8 Then
      var_c2 = var_c + "00"
      rs.Open "select * from tb_submenus where char_sme_submenu_id = '" + var_c2 + "'", cnn, adOpenDynamic, adLockOptimistic
      var_nivel = rs(8).Value
      var_nivel1 = rs(2).Value
      var_nivel2 = rs(3).Value
      var_nivel3 = rs(4).Value
      var_accion_submenu = rs(9).Value
      var_nombre_submenu = rs(7).Value
      var_permiso1 = rs(10).Value
      var_permiso2 = rs(11).Value
      var_permiso3 = rs(12).Value
      var_permiso4 = rs(13).Value
      var_numero_submenu = rs(14).Value
      txt_numero = var_numero_submenu
      rs.Close
   End If
   If var_longitud = 10 Then
      var_c2 = var_c
      rs.Open "select * from tb_submenus where char_sme_submenu_id = '" + var_c2 + "'", cnn, adOpenDynamic, adLockOptimistic
      var_nivel = rs(8).Value
      var_nivel1 = rs(2).Value
      var_nivel2 = rs(3).Value
      var_nivel3 = rs(4).Value
      var_accion_submenu = rs(9).Value
      var_nombre_submenu = rs(7).Value
      var_permiso1 = rs(10).Value
      var_permiso2 = rs(11).Value
      var_permiso3 = rs(12).Value
      var_permiso4 = rs(13).Value
      var_numero_submenu = rs(14).Value
      txt_numero = var_numero_submenu
      rs.Close
   End If
   
   txt_clave = var_c2
   txt_titulo = var_nombre_submenu
   txt_nivel = var_nivel
   txt_accion = var_accion_submenu
   chk_permiso1 = var_permiso1
   chk_permiso2 = var_permiso2
   chk_permiso3 = var_permiso3
   chk_permiso4 = var_permiso4
   rs.Open "select * from TB_FORMAS order by vcha_for_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
   Call RecsetToCombo(cmb_formas.hwnd, rs, 1)
   rs.Close
   'cmb_formas = Obtener_llave(cnn, rs, "TB_FORMAS", "VCHA_FOR_FORMA_ID", txt_accion, 1, "T")
   rs.Open "select * from tb_formas where vcha_for_forma_id = '" + Me.txt_accion + "'", cnn, adOpenDynamic, adLockOptimistic
   cmb_formas = rs!vcha_for_nombre
   rs.Close
   lbl_submenu.Caption = "Modificar " + Trim(var_nombre_submenu)
   frm_submenus.Visible = True
   txt_titulo.SetFocus
End Sub

Private Sub mnu_eliminar_click()
   var_opcion = 4
   var_clave_nivel1 = 0
   var_clave_nivel2 = 0
   var_clave_nivel3 = 0
   var_clave_nivel4 = 0
   var_clave_nivel5 = 0
   var_c = TreeView1.selectedItem.Key
   var_longitud = Len(Trim(var_c))
   If var_longitud = 6 Then
      var_c = Trim(Mid(var_c, 2, 4))
   End If
   If var_longitud = 8 Then
      var_c = Trim(Mid(var_c, 2, 6))
   End If
   If var_longitud = 10 Then
      var_c = Trim(Mid(var_c, 2, 8))
   End If
   If var_longitud = 12 Then
      var_c = Trim(Mid(var_c, 2, 10))
   End If
   var_longitud = Len(Trim(var_c))
   If var_longitud = 4 Then
      var_c2 = var_c + "000000"
      rs.Open "select * from tb_submenus where char_sme_submenu_id = '" + var_c2 + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
      var_nivel = rs(8).Value
      var_nivel1 = rs(2).Value
      var_clave_nivel1 = rs(2).Value
      var_nombre_submenu = rs(7).Value
      var_accion_submenu = rs(9).Value
      var_permiso1 = rs(10).Value
      var_permiso2 = rs(11).Value
      var_permiso3 = rs(12).Value
      var_permiso4 = rs(13).Value
      var_numero_submenu = rs(14).Value
      txt_numero = VAR_NUMERO
      End If
      rs.Close
   End If
   If var_longitud = 6 Then
      var_c2 = var_c + "0000"
      rs.Open "select * from tb_submenus where char_sme_submenu_id = '" + var_c2 + "'", cnn, adOpenDynamic, adLockOptimistic
      var_nivel = rs(8).Value
      var_nivel1 = rs(2).Value
      var_nivel2 = rs(3).Value
      var_clave_nivel1 = rs(2).Value
      var_clave_nivel2 = rs(3).Value
      var_accion_submenu = rs(9).Value
      var_nombre_submenu = rs(7).Value
      var_permiso1 = rs(10).Value
      var_permiso2 = rs(11).Value
      var_permiso3 = rs(12).Value
      var_permiso4 = rs(13).Value
      var_numero_submenu = rs(14).Value
      txt_numero = VAR_NUMERO
      rs.Close
   End If
   If var_longitud = 8 Then
      var_c2 = var_c + "00"
      rs.Open "select * from tb_submenus where char_sme_submenu_id = '" + var_c2 + "'", cnn, adOpenDynamic, adLockOptimistic
      var_nivel = rs(8).Value
      var_nivel1 = rs(2).Value
      var_nivel2 = rs(3).Value
      var_nivel3 = rs(4).Value
      var_clave_nivel1 = rs(2).Value
      var_clave_nivel2 = rs(3).Value
      var_clave_nivel3 = rs(4).Value
      var_accion_submenu = rs(9).Value
      var_nombre_submenu = rs(7).Value
      var_permiso1 = rs(10).Value
      var_permiso2 = rs(11).Value
      var_permiso3 = rs(12).Value
      var_permiso4 = rs(13).Value
      var_numero_submenu = rs(14).Value
      txt_numero = VAR_NUMERO
      rs.Close
   End If
   If var_longitud = 10 Then
      var_c2 = var_c
      rs.Open "select * from tb_submenus where char_sme_submenu_id = '" + var_c2 + "'", cnn, adOpenDynamic, adLockOptimistic
      var_nivel = rs(8).Value
      var_nivel1 = rs(2).Value
      var_nivel2 = rs(3).Value
      var_nivel3 = rs(4).Value
      var_accion_submenu = rs(9).Value
      var_nombre_submenu = rs(7).Value
      var_permiso1 = rs(10).Value
      var_permiso2 = rs(11).Value
      var_permiso3 = rs(12).Value
      var_permiso4 = rs(13).Value
      var_numero_submenu = rs(14).Value
      txt_numero = VAR_NUMERO
      rs.Close
   End If
   
   txt_clave = var_c2
   txt_titulo = var_nombre_submenu
   txt_nivel = var_nivel
   txt_accion = var_accion_submenu
   chk_permiso1 = var_permiso1
   chk_permiso2 = var_permiso2
   chk_permiso3 = var_permiso3
   chk_permiso4 = var_permiso4
   rs.Open "select * from TB_FORMAS order by vcha_for_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
   Call RecsetToCombo(cmb_formas.hwnd, rs, 1)
   rs.Close
   rs.Open "select * from tb_formas where vcha_for_forma_id = '" + Me.txt_accion + "'", cnn, adOpenDynamic, adLockOptimistic
   cmb_formas = rs!vcha_for_nombre
   rs.Close
   'cmb_formas = Obtener_llave(cnn, rs, "TB_FORMAS", "VCHA_FOR_FORMA_ID", txt_accion, 1, "T")
   lbl_submenu.Caption = "Se elimina " + Trim(var_nombre_submenu)
   frm_submenus.Visible = True
   txt_titulo.SetFocus
End Sub

Private Sub mnu_insertar_menu_click()
   rs.Open "select max(inte_sme_nivel1) from tb_submenus where vcha_men_menu_id = '" + txt_menus(0) + "'", cnn, adOpenDynamic, adLockOptimistic
   If IsNull(rs(0).Value) Then
      var_clave_nivel1 = 1
   Else
      var_clave_nivel1 = rs(0).Value + 1
   End If
   If var_clave_nivel1 + 1 <= 10 Then
      var_clave_nivel1_s = "0" + Trim(Str(var_clave_nivel1))
   Else
      var_clave_nivel1_s = Trim(Str(var_clave_nivel1))
   End If
   txt_clave = txt_menus(0) + var_clave_nivel1_s + "000000"
   rs.Close
   var_numero_submenu = 0
   var_opcion = 1
   var_nivel = 1
   txt_titulo = ""
   txt_nivel = 1
   txt_accion = ""
   txt_titulo = ""
   txt_accion = ""
   txt_numero = "0"
   chk_permiso1 = var_permiso1
   chk_permiso2 = var_permiso2
   chk_permiso3 = var_permiso3
   chk_permiso4 = var_permiso4
   rs.Open "select * from TB_FORMAS order by vcha_for_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
   Call RecsetToCombo(cmb_formas.hwnd, rs, 1)
   rs.Close
   frm_submenus.Visible = True
   txt_titulo.SetFocus
End Sub

Private Sub txt_nivel_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 27 Then
      frm_submenus.Visible = False
   End If
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
End Sub

Private Sub txt_titulo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 27 Then
      frm_submenus.Visible = False
   End If
End Sub
