VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmvehiculos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo de vehículos"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   Icon            =   "frmvehiculos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   5955
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5460
      Picture         =   "frmvehiculos.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Salir"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1470
      Picture         =   "frmvehiculos.frx":0F04
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      Picture         =   "frmvehiculos.frx":1006
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   810
      Picture         =   "frmvehiculos.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      Picture         =   "frmvehiculos.frx":11DA
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Guardar Alt + G"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmvehiculos.frx":12DC
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   45
      Width           =   330
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4275
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   21
      Top             =   60
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   " Vehiculos "
      Height          =   3345
      Left            =   150
      TabIndex        =   0
      Top             =   465
      Width           =   5655
      Begin VB.TextBox txt_vehiculos 
         Height          =   315
         Index           =   10
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   31
         Top             =   285
         Width           =   900
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2205
         TabIndex        =   30
         Top             =   285
         Width           =   3285
      End
      Begin VB.TextBox txt_vehiculos 
         Height          =   315
         Index           =   9
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   26
         Top             =   2940
         Width           =   1095
      End
      Begin VB.TextBox txt_vehiculos 
         Height          =   315
         Index           =   8
         Left            =   4020
         TabIndex        =   23
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txt_vehiculos 
         Height          =   315
         Index           =   7
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   22
         Top             =   2610
         Width           =   1095
      End
      Begin VB.TextBox txt_vehiculos 
         Height          =   315
         Index           =   5
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   17
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txt_vehiculos 
         Height          =   315
         Index           =   6
         Left            =   4035
         TabIndex        =   16
         Top             =   2280
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2205
         TabIndex        =   15
         Top             =   1620
         Width           =   3285
      End
      Begin VB.TextBox txt_vehiculos 
         Height          =   315
         Index           =   4
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   12
         Top             =   1950
         Width           =   2010
      End
      Begin VB.TextBox txt_vehiculos 
         Height          =   315
         Index           =   3
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1620
         Width           =   900
      End
      Begin VB.TextBox txt_vehiculos 
         Height          =   315
         Index           =   2
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1290
         Width           =   900
      End
      Begin VB.TextBox txt_vehiculos 
         Height          =   315
         Index           =   1
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   2
         Top             =   960
         Width           =   4170
      End
      Begin VB.TextBox txt_vehiculos 
         Height          =   315
         Index           =   0
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   1
         Top             =   630
         Width           =   1155
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Aseguradora:"
         Height          =   195
         Index           =   10
         Left            =   195
         TabIndex        =   32
         Top             =   345
         Width           =   945
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Alto:"
         Height          =   195
         Index           =   9
         Left            =   195
         TabIndex        =   27
         Top             =   3000
         Width           =   315
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Largo:"
         Height          =   195
         Index           =   8
         Left            =   2925
         TabIndex        =   25
         Top             =   2700
         Width           =   450
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Ancho:"
         Height          =   195
         Index           =   3
         Left            =   195
         TabIndex        =   24
         Top             =   2670
         Width           =   510
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Vigencia inicio:"
         Height          =   195
         Index           =   7
         Left            =   195
         TabIndex        =   19
         Top             =   2340
         Width           =   1065
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Vigencia final:"
         Height          =   195
         Index           =   6
         Left            =   2940
         TabIndex        =   18
         Top             =   2340
         Width           =   990
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Poliza:"
         Height          =   195
         Index           =   5
         Left            =   195
         TabIndex        =   14
         Top             =   2010
         Width           =   465
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Aseguradora:"
         Height          =   195
         Index           =   4
         Left            =   195
         TabIndex        =   13
         Top             =   1680
         Width           =   945
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
         Height          =   195
         Index           =   2
         Left            =   195
         TabIndex        =   6
         Top             =   1350
         Width           =   330
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Modelo:"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   5
         Top             =   1020
         Width           =   570
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Placas:"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   4
         Top             =   690
         Width           =   525
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   150
      TabIndex        =   7
      Top             =   3810
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   1905
         TabIndex        =   28
         Top             =   150
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3765
         TabIndex        =   29
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
               Object.ToolTipText     =   "Nuevo Registro"
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
         Caption         =   "Busqueda de Vehículo:"
         Height          =   195
         Left            =   195
         TabIndex        =   8
         Top             =   195
         Width           =   1680
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3045
      Left            =   135
      TabIndex        =   9
      Top             =   4410
      Width           =   5670
      Begin MSComctlLib.ListView lv_vehiculos 
         Height          =   2865
         Left            =   45
         TabIndex        =   20
         Top             =   135
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   5054
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Placas"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Modelo"
            Object.Width           =   6879
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "anio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "aseguradora"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Poliza"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Vigencia inicio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Vigencia final"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Ancho"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Largo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Alto"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   10
      Top             =   285
      Width           =   5655
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   3780
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
            Picture         =   "frmvehiculos.frx":13DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmvehiculos.frx":1CB8
            Key             =   ""
         EndProperty
      EndProperty
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
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmvehiculos.frx":2592
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmvehiculos.frx":2E6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmvehiculos.frx":3746
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmvehiculos.frx":3CE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmvehiculos.frx":45BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmvehiculos.frx":4E98
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmvehiculos.frx":5772
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmvehiculos.frx":5884
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmvehiculos.frx":5996
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmvehiculos.frx":5AA8
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmvehiculos"
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
      Call pro_elimina_vehiculos
      rs.Open "select * from tb_vehiculos", cnn, adOpenDynamic, adLockOptimistic
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
   Dim var_posible As Boolean
   var_posible = True
   If var_modifica_registro_vehiculo = False Then
      rs.Open "select * from tb_vehiculos where VCHA_VEH_PLACAS_ID = '" + Me.txt_vehiculos(0) + "'", cnn, adOpenDynamic, adLockOptimistic
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
         Call pro_guardar_vehiculos
         rs.Open "select * from tb_vehiculos", cnn, adOpenDynamic, adLockOptimistic
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
      MsgBox "Placas de vehiculo ya existen", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call gPrintListView(lv_vehiculos, "LISTADO DE vehiculos")
        End If

End Sub

Private Sub cmd_nuevo_Click()
        Call pro_limpiatextos(Me)
        txt_vehiculos(0).Enabled = True
        txt_vehiculos(0).SetFocus: var_modifica_registro_vehiculo = False
        cmd_guardar.Enabled = True
        cmd_deshacer.Enabled = True
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_vehiculo = False Then
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

Private Sub Combo1_Click()
   txt_vehiculos(0) = Obtener_llave(cnn, rs, "TB_EMPRESAS", "VCHA_EMP_NOMBRE", Combo1, 0, "T")
End Sub

Private Sub Combo2_Click()
   txt_vehiculos(3) = Obtener_llave(cnn, rs, "TB_ASEGURADORAS", "VCHA_ASE_NOMBRE", Combo2, 0, "T")
End Sub


Private Sub Combo2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_vehiculos(4).SetFocus
   Else
      KeyAscii = 0
   End If
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
   rs.Open "select * from tb_aseguradoras", cnn, adOpenDynamic, adLockBatchOptimistic
   Call RecsetToCombo(Combo2.hwnd, rs, 1)
   rs.Close
   var_modifica_registro_vehiculo = True
   lv_vehiculos.SmallIcons = ImageList1
   Call pro_encabezadosView(Me, lv_vehiculos, False)
   Call pro_llena_listview1
   pro_textos
   Call pro_AsignarAViewColor(lv_vehiculos, Picture1, vbWhite, vbGray)
   rs.Open "select * from tb_vehiculos", cnn, adOpenDynamic, adLockOptimistic
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
   var_modifica_registro_vehiculo = False
   Call activa_forma(var_activa_forma_vehiculos)
End Sub

Private Sub lv_vehiculos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_vehiculos, ColumnHeader)
End Sub

Private Sub lv_vehiculos_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_vehiculos.selectedItem = Item
        pro_textos
        var_modifica_registro_vehiculo = True
        txt_vehiculos(0).Enabled = True
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      Me.lv_vehiculos.SetFocus
      Call pro_avanzar(Me, lv_vehiculos, Button)
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_vehiculos.ListItems(1).Selected = True
      lv_vehiculos.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 4 Then
      numero_items_usos = lv_vehiculos.ListItems.Count
      lv_vehiculos.ListItems(numero_items_usos).Selected = True
      lv_vehiculos.selectedItem.EnsureVisible
      pro_textos
   End If
err0:
End Sub


Sub pro_guardar_vehiculos()

Dim ok As Boolean

Set TB_VEHICULOS = New TB_VEHICULOS
    
    
    If txt_vehiculos(0) <> "" And txt_vehiculos(1) <> "" Then
        If var_hubo_cambios Then
            ok = TB_VEHICULOS.Anadir(txt_vehiculos(0), txt_vehiculos(1), txt_vehiculos(2), txt_vehiculos(3), txt_vehiculos(4), txt_vehiculos(5), txt_vehiculos(6), Date, var_clave_usuario_global, fun_NombrePc, var_numero_planta)
            If ok Then
                pro_actualiza_ListView
                txt_vehiculos(0).Enabled = False
                MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
                txt_registros = lv_vehiculos.ListItems.Count
                var_modifica_registro_vehiculo = True
            Else
                MsgBox "No se puede grabar registro: " + TB_VEHICULOS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
        End If
    End If
    
Set TB_VEHICULOS = Nothing: var_hubo_cambios = False

End Sub

Sub pro_elimina_vehiculos()
Dim var_llave_usuarios As String

Set TB_VEHICULOS = New TB_VEHICULOS

    On Error GoTo salir:
    ok = True
    rs.Open "select * from TB_ARTICULOS,TB_DETALLE where TB_ARTICULOS.VCHA_ART_ARTICULO_ID = TB_DETALLE.VCHA_ART_ARTICULO_ID AND TB_ARTICULOS.VCHA_ART_LINEA = '" & txt_vehiculos(1) & "'", cnn, adOpenForwardOnly, adLockOptimistic
    If rs.RecordCount = 0 Then
        rs.Close
        If txt_vehiculos(0) <> "" And txt_vehiculos(1) <> "" And var_modifica_registro_vehiculo = True Then
            If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
                ok = TB_VEHICULOS.Eliminar(txt_vehiculos(0))
            Else
                GoTo salir:
            End If
            If ok Then
                MsgBox "Se Elimino Correctamente el Registro", vbInformation
                lv_vehiculos.ListItems.Remove (lv_vehiculos.selectedItem.Index)
                Call pro_limpiatextos(Me)
                txt_registros = lv_vehiculos.ListItems.Count
                lv_vehiculos.selectedItem.Selected = True
                pro_textos
            Else
                MsgBox "No se puede grabar registro: " + TB_VEHICULOS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
        End If
    Else
        SetTimer hwnd, NV_CLOSEMSGBOX, 1800, AddressOf TimerProc
        MsgBox "No se Puede Borrar Este Registro, Existen Dependencias", , "TRANSACCIONES [ AVISO ]"
        rs.Close
    End If

salir:
Set TB_VEHICULOS = Nothing

End Sub


Sub pro_llena_listview1()

Dim list_item As ListItem

    rs.Open "select * from TB_vehiculos", cnn, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
        Set list_item = lv_vehiculos.ListItems.Add(, , rs(0).Value): list_item.SmallIcon = 9
        list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
        list_item.SubItems(2) = IIf(IsNull(rs(2).Value), "", rs(2).Value)
        list_item.SubItems(3) = IIf(IsNull(rs(3).Value), "", rs(3).Value)
        list_item.SubItems(4) = IIf(IsNull(rs(4).Value), "", rs(4).Value)
        list_item.SubItems(5) = IIf(IsNull(rs(5).Value), "", rs(5).Value)
        list_item.SubItems(6) = IIf(IsNull(rs(6).Value), "", rs(6).Value)
    rs.MoveNext:
    Wend
    rs.Close

End Sub


Sub pro_textos()
On Error GoTo err0:
Dim var_n As Double
   var_n = lv_vehiculos.ListItems.Count
   If var_n > 0 Then
      txt_vehiculos(0) = lv_vehiculos.selectedItem
      txt_vehiculos(1) = lv_vehiculos.selectedItem.SubItems(1)
      txt_vehiculos(2) = lv_vehiculos.selectedItem.SubItems(2)
      txt_vehiculos(3) = lv_vehiculos.selectedItem.SubItems(3)
      txt_vehiculos(4) = lv_vehiculos.selectedItem.SubItems(4)
      txt_vehiculos(5) = lv_vehiculos.selectedItem.SubItems(5)
      txt_vehiculos(6) = lv_vehiculos.selectedItem.SubItems(6)
      Combo2 = Obtener_llave(cnn, rs, "TB_aseguradoras", "Vcha_ase_aseguradora_id", txt_vehiculos(3), 1, "T")
   End If
   var_numero_renglones = lv_vehiculos.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_vehiculos.ColumnHeaders(2).Width = 3850
   Else
      lv_vehiculos.ColumnHeaders(2).Width = 4099.71
   End If
   var_modifica_registro_vehiculo = True
   var_hubo_cambios = False
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro_vehiculo = False Then
        Set list_item = lv_vehiculos.ListItems.Add(, , txt_vehiculos(0)): list_item.SmallIcon = 9
        list_item.SubItems(1) = txt_vehiculos(1)
        list_item.SubItems(2) = txt_vehiculos(2)
        list_item.SubItems(3) = txt_vehiculos(3)
        list_item.SubItems(4) = txt_vehiculos(4)
        list_item.SubItems(5) = txt_vehiculos(5)
        list_item.SubItems(6) = txt_vehiculos(6)
        list_item.EnsureVisible
        list_item.Selected = True
    Else
        lv_vehiculos.ListItems.Item(lv_vehiculos.selectedItem.Index).Checked = False
        lv_vehiculos.ListItems.Item(lv_vehiculos.selectedItem.Index) = txt_vehiculos(0)
        lv_vehiculos.ListItems.Item(lv_vehiculos.selectedItem.Index).ListSubItems(1) = txt_vehiculos(1)
        lv_vehiculos.ListItems.Item(lv_vehiculos.selectedItem.Index).ListSubItems(2) = txt_vehiculos(2)
        lv_vehiculos.ListItems.Item(lv_vehiculos.selectedItem.Index).ListSubItems(3) = txt_vehiculos(3)
        lv_vehiculos.ListItems.Item(lv_vehiculos.selectedItem.Index).ListSubItems(4) = txt_vehiculos(4)
        lv_vehiculos.ListItems.Item(lv_vehiculos.selectedItem.Index).ListSubItems(5) = txt_vehiculos(5)
        lv_vehiculos.ListItems.Item(lv_vehiculos.selectedItem.Index).ListSubItems(6) = txt_vehiculos(6)
        lv_vehiculos.ListItems.Item(lv_vehiculos.selectedItem.Index).Selected = True
    End If
    lv_vehiculos.SetFocus
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_vehiculos, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub

Private Sub txt_vehiculos_Change(Index As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub

Private Sub txt_vehiculos_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
      var_hubo_cambios = True
   End If
End Sub
