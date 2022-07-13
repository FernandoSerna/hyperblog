VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmgruposactuales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grupos Actuales"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   Icon            =   "frmgruposactuales.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5955
   Begin VB.Frame frm_lista 
      Height          =   2415
      Left            =   180
      TabIndex        =   23
      Top             =   3045
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   24
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
         NumItems        =   3
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
            Text            =   "Grupo"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   25
         Top             =   135
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmgruposactuales.frx":08CA
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
      Picture         =   "frmgruposactuales.frx":09CC
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
      Picture         =   "frmgruposactuales.frx":0ACE
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
      Picture         =   "frmgruposactuales.frx":0BA0
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
      Picture         =   "frmgruposactuales.frx":0CA2
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
      Picture         =   "frmgruposactuales.frx":0DA4
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
      TabIndex        =   20
      Top             =   75
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   " Grupo Actual "
      Height          =   2070
      Left            =   150
      TabIndex        =   0
      Top             =   465
      Width           =   5655
      Begin VB.TextBox txt_descuento_3 
         Height          =   315
         Left            =   1245
         TabIndex        =   27
         Top             =   1650
         Width           =   900
      End
      Begin VB.CheckBox chk_cliente_telas 
         Alignment       =   1  'Right Justify
         Caption         =   "Cliente de telas de la textilera"
         Height          =   270
         Left            =   3045
         TabIndex        =   26
         Top             =   1305
         Width           =   2385
      End
      Begin VB.TextBox txt_descuento_2 
         Height          =   315
         Left            =   1245
         TabIndex        =   10
         Top             =   1305
         Width           =   900
      End
      Begin VB.TextBox txt_descuento_1 
         Height          =   315
         Left            =   1245
         TabIndex        =   9
         Top             =   945
         Width           =   900
      End
      Begin VB.TextBox txt_nombre_grupo_actual 
         Height          =   315
         Left            =   1245
         TabIndex        =   8
         Top             =   600
         Width           =   4245
      End
      Begin VB.TextBox txt_grupo_actual 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   255
         Width           =   900
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Garantia:"
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   28
         Top             =   1680
         Width           =   645
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Descuento 2:"
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   18
         Top             =   1335
         Width           =   960
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Descuento 1:"
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   13
         Top             =   990
         Width           =   960
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   12
         Top             =   645
         Width           =   555
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   11
         Top             =   315
         Width           =   450
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   150
      TabIndex        =   14
      Top             =   2520
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   2355
         TabIndex        =   22
         Top             =   165
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3930
         TabIndex        =   21
         Top             =   150
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
         Caption         =   "Busqueda de Grupo Actual:"
         Height          =   195
         Left            =   195
         TabIndex        =   15
         Top             =   225
         Width           =   1965
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4215
      Left            =   150
      TabIndex        =   16
      Top             =   3045
      Width           =   5655
      Begin MSComctlLib.ImageList icono_encabezado 
         Left            =   0
         Top             =   15
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
               Picture         =   "frmgruposactuales.frx":13DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmgruposactuales.frx":1CB8
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lv_gruposactuales 
         Height          =   3990
         Left            =   45
         TabIndex        =   19
         Top             =   150
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   7038
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
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
            Text            =   "descuento1"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "descuento2"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "tela"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "promedio"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   17
      Top             =   285
      Width           =   5655
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
            Picture         =   "frmgruposactuales.frx":2592
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgruposactuales.frx":2E6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgruposactuales.frx":3746
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgruposactuales.frx":3CE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgruposactuales.frx":45BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgruposactuales.frx":4E98
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgruposactuales.frx":5772
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgruposactuales.frx":5884
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgruposactuales.frx":5996
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgruposactuales.frx":5AA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgruposactuales.frx":5BBA
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmgruposactuales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim var_guardar_cambios As Boolean
Dim var_hubo_cambios As Boolean
Dim numero_items_gruposactales As Integer




Private Sub chk_cliente_telas_Click()
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
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         rsaux5.Open "select vcha_emp_conexion from tb_empresas where len(vcha_emp_conexion) > 0", cnn_distribucion, adOpenDynamic, adLockOptimistic
         While Not rsaux5.EOF
               var_conexion_importacion = IIf(IsNull(rsaux5(0).Value), "", rsaux5(0).Value)
               If Trim(var_conexion_importacion) <> "" Then
                  If cnn_importacion.State = 1 Then
                     cnn_importacion.Close
                  End If
                  cnn_importacion.Open var_conexion_importacion
                  Call pro_elimina_gruposactuales
               End If
               rsaux5.MoveNext
         Wend
         rsaux5.Close
         numero_items_gruposactales = numero_items_gruposactales - 1
         MsgBox "Se Elimino Correctamente el Registro", vbInformation
         lv_gruposactuales.ListItems.Remove (lv_gruposactuales.selectedItem.Index)
         Call pro_limpiatextos(Me)
         txt_registros = lv_gruposactuales.ListItems.Count
         lv_gruposactuales.selectedItem.Selected = True
         pro_textos
     End If
   
   
      rs.Open "select * from tb_gruposactuales", cnn_distribucion, adOpenDynamic, adLockOptimistic
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
   Dim var_descuento_anterior As Double
   Dim var_descuento_actual As Double
   If Trim(txt_descuento_1) = "" Then
      txt_descuento_1 = 0
   End If
   If Trim(txt_descuento_2) = "" Then
      txt_descuento_2 = 0
   End If
   If Not IsNumeric(txt_descuento_1) Then
      txt_descuento_1 = 0
   End If
   If Not IsNumeric(txt_descuento_2) Then
      txt_descuento_2 = 0
   End If
   If txt_nombre_grupo_actual = "" Or txt_descuento_1 = "" Or txt_descuento_2 = "" Then
      MsgBox "Información incompleta", vbOKOnly, "ATENCION"
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
         If var_empresa = "18" Then
            rs.Open "SELECT FLOA_GAC_DESCUENTO_1 FROM TB_GRUPOSACTUALES WHERE VCHA_GAC_GRUPO_aCTUAL_ID = '" + Me.txt_grupo_actual + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_descuento_anterior = IIf(IsNull(rs(0).Value), 0, rs(0))
            Else
               var_descuento_anterior = 0
            End If
            rs.Close
         End If
         rsaux5.Open "select vcha_emp_conexion from tb_empresas where len(vcha_emp_conexion) > 0  ORDER BY INTE_EMP_ORDEN_CONEXION", cnn_distribucion, adOpenDynamic, adLockOptimistic
         While Not rsaux5.EOF
               var_conexion_importacion = IIf(IsNull(rsaux5(0).Value), "", rsaux5(0).Value)
               If Trim(var_conexion_importacion) <> "" Then
                  If cnn_importacion.State = 1 Then
                     cnn_importacion.Close
                  End If
                  cnn_importacion.Open var_conexion_importacion
                  Call pro_guardar_gruposactuales
               End If
               rsaux5.MoveNext
         Wend
         rsaux5.Close
         If var_empresa = "18" Then
            rs.Open "SELECT FLOA_GAC_DESCUENTO_1 FROM TB_GRUPOSACTUALES WHERE VCHA_GAC_GRUPO_aCTUAL_ID = '" + Me.txt_grupo_actual + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_descuento_actual = IIf(IsNull(rs(0).Value), 0, rs(0))
            Else
               var_descuento_actual = 0
            End If
            rs.Close
            If var_descuento_anterior <> var_descuento_actual Then
               rs.Open "insert into TB_BITACORA_DESCUENTOS_GRUPOS_ACTUALES (VCHA_BIT_USUARIO, DTIM_BIT_FECHA, VCHA_GAC_GRUPO_ACTUAL_ID, FLOA_BIT_DESCUENTO_ANTERIOR, FLOA_BIT_DESCUENTO_ACTUAL) values ('" + var_clave_usuario_global + "',getdate(), '" + Me.txt_grupo_actual + "'," + CStr(var_descuento_anterior) + "," + CStr(var_descuento_actual) + ")"
            End If
         End If
         pro_actualiza_ListView
         txt_grupo_actual.Enabled = False
         MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
         txt_registros = lv_gruposactuales.ListItems.Count
         var_modifica_registro_ga = True
         var_hubo_cambios = False
      
      
         rs.Open "select * from tb_gruposactuales", cnn_distribucion, adOpenDynamic, adLockOptimistic
         If rs.BOF Then
            var_guardar_cambios = False
            cmd_guardar.Enabled = False
            cmd_deshacer.Enabled = False
            cmd_eliminar.Enabled = False
         Else
            cmd_guardar.Enabled = True
            cmd_deshacer.Enabled = True
            cmd_eliminar.Enabled = True
            var_guardar_cambios = False
         End If
         rs.Close
      End If
   End If
End Sub

Private Sub cmd_imprimir_Click()
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call gPrintListView(lv_gruposactuales, "LISTADO DE gruposactuales")
        End If

End Sub

Private Sub cmd_nuevo_Click()
   Call pro_limpiatextos(Me)
   txt_nombre_grupo_actual.Enabled = True
   txt_nombre_grupo_actual.SetFocus: var_modifica_registro_ga = False
   cmd_guardar.Enabled = True
   cmd_deshacer.Enabled = True
   var_guardar_cambios = True
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_ga = False Then
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
   var_cadena_seguridad = ""
   Top = 0
   Left = 2900
   var_modifica_registro_ga = True
   lv_gruposactuales.SmallIcons = ImageList
   Call pro_encabezadosView(Me, lv_gruposactuales, False)
   Call pro_llena_listview1
   pro_textos
   rs.Open "select * from tb_gruposactuales", cnn_distribucion, adOpenDynamic, adLockOptimistic
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
   var_guardar_cambios = False
   frm_lista.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    var_swpassword = False
    var_modifica_registro_ga = False
    Call activa_forma(var_activa_forma_gruposactuales)
End Sub

Private Sub lv_gruposactuales_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_gruposactuales, ColumnHeader)
End Sub

Private Sub lv_gruposactuales_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_gruposactuales.selectedItem = Item
        pro_textos
        var_modifica_registro_ga = True
        txt_grupo_actual.Enabled = False
End Sub

Private Sub lv_gruposactuales_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select vcha_cli_clave_id, vcha_cli_nombre, vcha_gac_grupo_actual_id from vw_clientes where vcha_emp_empresa_id = '" + var_empresa + "' order by vcha_cli_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            list_item.SubItems(2) = IIf(IsNull(rs!VCHA_GAC_GRUPO_aCTUAL_ID), "", rs!VCHA_GAC_GRUPO_aCTUAL_ID)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Clientes"
      var_tipo_lista = 2
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 7 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         Call pro_busca_registro(Me.lv_gruposactuales, lv_lista.selectedItem.SubItems(2), True)
      End If
      Me.lv_gruposactuales.SetFocus
   End If
   If KeyAscii = 27 Then
      Me.lv_gruposactuales.SetFocus
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      Me.lv_gruposactuales.SetFocus
      Call pro_avanzar(Me, lv_gruposactuales, Button)
      Me.lv_gruposactuales.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_gruposactuales.ListItems(1).Selected = True
      Me.lv_gruposactuales.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 4 Then
      numero_items_gruposactales = Me.lv_gruposactuales.ListItems.Count
      lv_gruposactuales.ListItems(numero_items_gruposactales).Selected = True
      Me.lv_gruposactuales.selectedItem.EnsureVisible
      pro_textos
   End If
err0:
End Sub


Sub pro_guardar_gruposactuales()

Dim ok As Boolean

Set TB_GRUPOSACTUALES = New TB_GRUPOSACTUALES
Set TB_BITACORA_GRUPOS_ACTUALES = New TB_BITACORA_GRUPOS_ACTUALES
    
    If txt_nombre_grupo_actual <> "" Then
        If var_hubo_cambios Then
           rs.Open "select * from tb_gruposactuales where vcha_gac_grupo_actual_id = '" + txt_grupo_actual + "'", cnn_importacion, adOpenDynamic, adLockOptimistic
           var_grupo_actual_regreso = txt_grupo_actual
           ok = TB_GRUPOSACTUALES.Anadir(txt_grupo_actual, txt_nombre_grupo_actual, txt_descuento_1, txt_descuento_2)
           If Trim(var_grupo_actual_regreso) <> "" Then
              txt_grupo_actual = var_grupo_actual_regreso
           End If
           If ok Then
              rsaux.Open "update tb_gruposactuales set floa_gac_descuento_3 = " + Me.txt_descuento_3 + ", vcha_emp_empresa_id = '" + var_empresa + "', INTE_GAC_TELA = " + CStr(Me.chk_cliente_telas) + " where vcha_gac_grupo_actual_id = '" + Me.txt_grupo_actual + "'", cnn_importacion, adOpenDynamic, adLockOptimistic
              bitacora = True
              If var_modifica_registro_ag = False Then
                 var_operacion_bitacora = "I"
                 bitacora = TB_BITACORA_GRUPOS_ACTUALES.Anadir(txt_grupo_actual, "VCHA_GAC_NOMBRE", var_operacion_bitacora, "", txt_nombre_grupo_actual, var_clave_usuario_global, fun_NombrePc, Date)
              Else
                 var_operacion_bitacora = "M"
                 If rs!VCHA_GAC_GRUPO_aCTUAL_ID <> txt_grupo_actual Then
                    bitacora = TB_BITACORA_GRUPOS_ACTUALES.Anadir(txt_grupo_actual, "VCHA_GAC_GRUPO_ACTUAL_ID", var_operacion_bitacora, rs!VCHA_GAC_GRUPO_aCTUAL_ID, txt_grupo_actual, var_clave_usuario_global, fun_NombrePc, Date)
                 End If
                 If rs!vcha_gac_nombre <> txt_nombre_grupo_actual Then
                    bitacora = TB_BITACORA_GRUPOS_ACTUALES.Anadir(txt_grupo_actual, "VCHA_GAC_NOMBRE", var_operacion_bitacora, rs!vcha_gac_nombre, txt_nombre_grupo_actual, var_clave_usuario_global, fun_NombrePc, Date)
                 End If
                 If rs!floa_gac_Descuento_1 <> txt_descuento_1 Then
                    bitacora = TB_BITACORA_GRUPOS_ACTUALES.Anadir(txt_grupo_actual, "VCHA_GAC_DESCUENTO1", var_operacion_bitacora, rs!floa_gac_Descuento_1, txt_descuento_1, var_clave_usuario_global, fun_NombrePc, Date)
                 End If
                 If rs!FLOA_GAC_DESCUENTO_2 <> txt_descuento_2 Then
                    bitacora = tb_bitacora_GRUPOSA_CTUALES.Anadir(txt_grupo_actual, "VCHA_GAC_DESCUENTO2", var_operacion_bitacora, rs!FLOA_GAC_DESCUENTO_2, txt_descuento_2, var_clave_usuario_global, fun_NombrePc, Date)
                 End If
              End If
              rs.Close
            Else
                MsgBox "No se puede grabar registro: " + TB_GRUPOSACTUALES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
            End If
        End If
    End If
    
Set TB_GRUPOSACTUALES = Nothing

End Sub

Sub pro_elimina_gruposactuales()
   Dim var_llave_usuarios As String
   Set TB_GRUPOSACTUALES = New TB_GRUPOSACTUALES
   Set TB_BITACORA_GRUPOS_ACTUALES = New TB_BITACORA_GRUPOS_ACTUALES
   On Error GoTo salir:
   ok = True
   If txt_grupo_actual <> "" And var_modifica_registro_ga = True Then
      'If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_GRUPOSACTUALES.Eliminar(txt_grupo_actual)
      'Else
      '   GoTo salir:
      'End If
      If ok Then
         var_operacion_bitacora = "E"
         bitacora = TB_BITACORA_GRUPOS_ACTUALES.Anadir(txt_grupo_actual, "VCHA_GAC_NOMBRE", var_operacion_bitacora, txt_nombre_grupo_actual, "", var_clave_usuario_global, fun_NombrePc, Date)
      Else
         MsgBox "No se puede grabar registro: " + TB_GRUPOSACTUALES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
   Set TB_GRUPOSACTUALES = Nothing
End Sub


Sub pro_llena_listview1()

Dim list_item As ListItem

    rs.Open "select distinct vcha_gac_grupo_actual_id, vcha_gac_nombre, floa_gac_descuento_1, floa_gac_descuento_2, INTE_GAC_TELA, floa_gac_descuento_3 from vw_clientes  where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_gac_grupo_actual_id is not null", cnn_distribucion, adOpenDynamic, adLockOptimistic
    numero_items_gruposactales = 0
    While Not rs.EOF
        Set list_item = lv_gruposactuales.ListItems.Add(, , Trim(rs!VCHA_GAC_GRUPO_aCTUAL_ID))
        list_item.SubItems(1) = IIf(IsNull(rs!vcha_gac_nombre), "", rs!vcha_gac_nombre)
        list_item.SubItems(2) = IIf(IsNull(rs!floa_gac_Descuento_1), 0, rs!floa_gac_Descuento_1)
        list_item.SubItems(3) = IIf(IsNull(rs!FLOA_GAC_DESCUENTO_2), "", rs!FLOA_GAC_DESCUENTO_2)
        list_item.SubItems(4) = IIf(IsNull(rs!INTE_GAC_TELA), 0, rs!INTE_GAC_TELA)
        list_item.SubItems(5) = IIf(IsNull(rs!floa_gac_descuento_3), 0, rs!floa_gac_descuento_3)
        rs.MoveNext:
        numero_items_gruposactales = numero_items_gruposactales + 1
 
    Wend
    rs.Close
End Sub


Sub pro_textos()
On Error GoTo err0:
Dim var_n As Double
   var_n = lv_gruposactuales.ListItems.Count
   If var_n > 0 Then
      txt_grupo_actual = lv_gruposactuales.selectedItem
      txt_nombre_grupo_actual = lv_gruposactuales.selectedItem.SubItems(1)
      txt_descuento_1 = lv_gruposactuales.selectedItem.SubItems(2)
      txt_descuento_2 = lv_gruposactuales.selectedItem.SubItems(3)
      Me.chk_cliente_telas = lv_gruposactuales.selectedItem.SubItems(4)
      Me.txt_descuento_3 = lv_gruposactuales.selectedItem.SubItems(5)
      txt_grupo_actual.Enabled = False
   End If
   var_numero_renglones = lv_gruposactuales.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_gruposactuales.ColumnHeaders(2).Width = 3850
   Else
      lv_gruposactuales.ColumnHeaders(2).Width = 4099.71
   End If
   var_modifica_registro_ga = True
   var_hubo_cambios = False
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro_ga = False Then
        Set list_item = lv_gruposactuales.ListItems.Add(, , txt_grupo_actual)
        list_item.SubItems(1) = txt_nombre_grupo_actual
        list_item.SubItems(2) = txt_descuento_1
        list_item.SubItems(3) = txt_descuento_2
        list_item.SubItems(4) = Me.chk_cliente_telas
        list_item.SubItems(5) = Me.txt_descuento_3
        list_item.EnsureVisible
        list_item.Selected = True
        numero_items_gruposactales = numero_items_gruposactales + 1
    Else
        lv_gruposactuales.ListItems.Item(lv_gruposactuales.selectedItem.Index).Checked = False
        lv_gruposactuales.ListItems.Item(lv_gruposactuales.selectedItem.Index) = txt_grupo_actual
        lv_gruposactuales.ListItems.Item(lv_gruposactuales.selectedItem.Index).ListSubItems(1) = txt_nombre_grupo_actual
        lv_gruposactuales.ListItems.Item(lv_gruposactuales.selectedItem.Index).ListSubItems(2) = txt_descuento_1
        lv_gruposactuales.ListItems.Item(lv_gruposactuales.selectedItem.Index).ListSubItems(3) = txt_descuento_2
        lv_gruposactuales.ListItems.Item(lv_gruposactuales.selectedItem.Index).ListSubItems(4) = Me.chk_cliente_telas
        lv_gruposactuales.ListItems.Item(lv_gruposactuales.selectedItem.Index).ListSubItems(5) = Me.txt_descuento_3
        lv_gruposactuales.ListItems.Item(lv_gruposactuales.selectedItem.Index).Selected = True
    End If
    lv_gruposactuales.SetFocus
End Sub



Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_gruposactuales, txt_buscar, True)
      txt_buscar = ""
      pro_textos
   End If
End Sub

Private Sub txt_descuento_1_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_descuento_1_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_descuento_1_LostFocus()
    If Not IsNumeric(txt_descuento_1) Then
       MsgBox "Descuento incorrecto", vbOKOnly, "ATENCION"
       txt_descuento_1 = 0
    Else
       If CDbl(txt_descuento_1) > 100 Then
          MsgBox "Descuento incorrecto", vbOKOnly, "ATENCION"
          txt_descuento_2 = 0
       End If
    End If
End Sub

Private Sub txt_descuento_2_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_descuento_2_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Me.txt_descuento_3.SetFocus
   End If
End Sub

Private Sub txt_descuento_2_LostFocus()
    If Not IsNumeric(txt_descuento_2) Then
       MsgBox "Descuento incorrecto", vbOKOnly, "ATENCION"
       txt_descuento_2 = 0
    Else
       If CDbl(txt_descuento_2) > 100 Then
          MsgBox "Descuento incorrecto", vbOKOnly, "ATENCION"
          txt_descuento_2 = 0
       End If
    End If
End Sub

Private Sub txt_descuento_3_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_descuento_3_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Me.cmd_guardar.Enabled = True Then
         Me.cmd_guardar.SetFocus
      End If
   End If
End Sub

Private Sub txt_grupo_actual_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_grupo_actual_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_grupo_actual_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_grupo_actual_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub
