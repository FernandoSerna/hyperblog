VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmgruposreales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grupos Reales"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   Icon            =   "frmgruposreales.frx":0000
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
      Left            =   165
      TabIndex        =   31
      Top             =   570
      Width           =   5670
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   32
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
         TabIndex        =   33
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmgruposreales.frx":08CA
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
      Picture         =   "frmgruposreales.frx":09CC
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
      Picture         =   "frmgruposreales.frx":0ACE
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
      Picture         =   "frmgruposreales.frx":0BA0
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
      Picture         =   "frmgruposreales.frx":0CA2
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
      Picture         =   "frmgruposreales.frx":0DA4
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
      Left            =   2685
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   25
      Top             =   90
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   24
      Top             =   285
      Width           =   5655
   End
   Begin VB.Frame Frame1 
      Caption         =   " Grupo Real "
      Height          =   2640
      Left            =   150
      TabIndex        =   0
      Top             =   465
      Width           =   5655
      Begin VB.TextBox txt_nombre_grupo_actual 
         Height          =   315
         Left            =   2475
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   14
         Top             =   2220
         Width           =   2760
      End
      Begin VB.TextBox txt_descuento_3 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1425
         TabIndex        =   11
         Top             =   1545
         Width           =   915
      End
      Begin VB.TextBox txt_prioridad 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1425
         MaxLength       =   1
         TabIndex        =   12
         Top             =   1875
         Width           =   390
      End
      Begin VB.TextBox txt_descuento_2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1425
         TabIndex        =   10
         Top             =   1215
         Width           =   915
      End
      Begin VB.TextBox txt_descuento_1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1425
         TabIndex        =   9
         Top             =   885
         Width           =   915
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   5235
         TabIndex        =   15
         ToolTipText     =   "Crear grupo actual automaticamente"
         Top             =   2205
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   11
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txt_nombre_grupo_real 
         Height          =   315
         Left            =   1425
         MaxLength       =   50
         TabIndex        =   8
         Top             =   555
         Width           =   4110
      End
      Begin VB.TextBox txt_grupo_real 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1425
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   225
         Width           =   945
      End
      Begin VB.TextBox txt_grupo_actual 
         Height          =   315
         Left            =   1410
         TabIndex        =   13
         Top             =   2220
         Width           =   1065
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Descuento 3:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   30
         Top             =   1605
         Width           =   960
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Prioridad:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   29
         Top             =   1935
         Width           =   660
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Descuento 2:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   28
         Top             =   1275
         Width           =   960
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Descuento 1:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   945
         Width           =   960
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   615
         Width           =   600
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   270
         Width           =   450
      End
      Begin VB.Label lab_estados 
         AutoSize        =   -1  'True
         Caption         =   "Grupo Actual (F5):"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   2265
         Width           =   1290
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   150
      TabIndex        =   21
      Top             =   3090
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   2085
         TabIndex        =   16
         Top             =   150
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3900
         TabIndex        =   26
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
         Caption         =   "Busqueda de Grupo Real:"
         Height          =   195
         Left            =   195
         TabIndex        =   22
         Top             =   195
         Width           =   1845
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3585
      Left            =   150
      TabIndex        =   23
      Top             =   3630
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
               Picture         =   "frmgruposreales.frx":13DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmgruposreales.frx":1CB8
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lv_gruposreales 
         Height          =   3405
         Left            =   45
         TabIndex        =   17
         Top             =   135
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   6006
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
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
            Text            =   "grupo actual"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "descuento1"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "descuento2"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "descuento3"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "prioridad"
            Object.Width           =   0
         EndProperty
      End
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
            Picture         =   "frmgruposreales.frx":2592
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgruposreales.frx":2E6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgruposreales.frx":3746
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgruposreales.frx":3CE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgruposreales.frx":45BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgruposreales.frx":4E98
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgruposreales.frx":5772
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgruposreales.frx":5884
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgruposreales.frx":5996
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgruposreales.frx":5AA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmgruposreales.frx":5BBA
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmgruposreales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim var_hubo_cambios As Boolean
Dim var_guardar_cambios As Boolean
Dim numero_items_gruposreales As Integer
Dim bitacora As Boolean
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
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         rsaux5.Open "select vcha_emp_conexion from tb_empresas where len(vcha_emp_conexion) > 0", cnn_distribucion, adOpenDynamic, adLockOptimistic
         While Not rsaux5.EOF
               var_conexion_importacion = IIf(IsNull(rsaux5(0).Value), "", rsaux5(0).Value)
               If Trim(var_conexion_importacion) <> "" Then
                  If cnn_importacion.State = 1 Then
                     cnn_importacion.Close
                  End If
                  cnn_importacion.Open var_conexion_importacion
                  Call pro_elimina_gruposreales
               End If
               rsaux5.MoveNext
         Wend
         rsaux5.Close
         numero_items_gruposreales = numero_items_gruposreales - 1
         MsgBox "Se Elimino Correctamente el Registro", vbInformation
         lv_gruposreales.ListItems.Remove (lv_gruposreales.selectedItem.Index)
         Call pro_limpiatextos(Me)
         txt_registros = lv_gruposreales.ListItems.Count
          lv_gruposreales.selectedItem.Selected = True
         If numero_items_gruposreales > 11 Then
            lv_gruposreales.ColumnHeaders(2).Width = 4200
         Else
            lv_gruposreales.ColumnHeaders(2).Width = 4499.71
         End If
         pro_textos
      
         rs.Open "select * from tb_gruposreales", cnn_distribucion, adOpenDynamic, adLockOptimistic
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
End Sub

Private Sub cmd_guardar_Click()
   If Not IsNumeric(txt_descuento_1) Then
      txt_descuento_1 = 0
   End If
   If Not IsNumeric(txt_descuento_2) Then
      txt_descuento_2 = 0
   End If
   If Not IsNumeric(txt_descuento_3) Then
      txt_descuento_3 = 0
   End If
   If txt_grupo_actual = "" Or txt_nombre_grupo_real = "" Then
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
         
         rsaux5.Open "select vcha_emp_conexion from tb_empresas where len(vcha_emp_conexion) > 0 and vcha_emp_empresa_id = '02' ORDER BY INTE_EMP_ORDEN_CONEXION", cnn_distribucion, adOpenDynamic, adLockOptimistic
         While Not rsaux5.EOF
               var_conexion_importacion = IIf(IsNull(rsaux5(0).Value), "", rsaux5(0).Value)
               If Trim(var_conexion_importacion) <> "" Then
                  If cnn_importacion.State = 1 Then
                     cnn_importacion.Close
                  End If
                  cnn_importacion.Open var_conexion_importacion
                  Call pro_guardar_gruposreales
               End If
               rsaux5.MoveNext
         Wend
         rsaux5.Close
         pro_actualiza_ListView
         MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
         txt_registros = lv_gruposreales.ListItems.Count
         var_modifica_registro_gr = True
         var_hubo_cambios = False
         
         
         rs.Open "select * from tb_gruposreales", cnn_distribucion, adOpenDynamic, adLockOptimistic
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
           Call gPrintListView(lv_gruposreales, "LISTADO DE gruposreales")
        End If

End Sub

Private Sub cmd_nuevo_Click()
        Call pro_limpiatextos(Me)
        txt_grupo_real.Enabled = False
        txt_nombre_grupo_real.SetFocus: var_modifica_registro_gr = False
        txt_nombre_grupo_real.Enabled = True
        cmd_guardar.Enabled = True
        cmd_deshacer.Enabled = True
        txt_grupo_actual.Enabled = True
        txt_nombre_grupo_actual.Enabled = True
        txt_nombre_grupo_real.SetFocus
        var_guardar_cambios = True
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_gr = False Then
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
   frm_lista.Visible = False
   txt_grupo_real.Enabled = False
   var_modifica_registro_gr = True
   lv_gruposreales.SmallIcons = ImageList
   Call pro_encabezadosView(Me, lv_gruposreales, False)
   Call pro_llena_listview1
   pro_textos
   var_guardar_cambios = False
   rs.Open "select * from tb_gruposreales", cnn_distribucion, adOpenDynamic, adLockOptimistic
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
  Dim var_guardar As Integer
    var_swpassword = False
    var_modifica_registro_gr = False
    Call activa_forma(var_activa_forma_gruposreales)
End Sub

Private Sub lv_gruposreales_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_gruposreales, ColumnHeader)
End Sub

Private Sub lv_gruposreales_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_gruposreales.selectedItem = Item
        pro_textos
        var_modifica_registro_gr = True
        txt_grupo_real.Enabled = False
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         txt_grupo_actual = lv_lista.selectedItem
         txt_nombre_grupo_actual = lv_lista.selectedItem.SubItems(1)
      Else
         txt_grupo_actual = ""
         txt_nombre_grupo_actual = ""
      End If
      txt_grupo_actual.SetFocus
      frm_lista.Visible = False
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_gruposreales.SetFocus
      Call pro_avanzar(Me, lv_gruposreales, Button)
      lv_gruposreales.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_gruposreales.ListItems(1).Selected = True
      lv_gruposreales.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 4 Then
      numero_items_gruposreales = lv_gruposreales.ListItems.Count
      lv_gruposreales.ListItems(numero_items_gruposreales).Selected = True
      lv_gruposreales.selectedItem.EnsureVisible
      pro_textos
   End If
err0:
End Sub


Sub pro_guardar_gruposreales()
Dim ok As Boolean
   Set TB_GRUPOSREALES = New TB_GRUPOSREALES
   Set TB_BITACORA_GRUPOS_REALES = New TB_BITACORA_GRUPOS_REALES
   ok = True
   If txt_grupo_actual <> "" And txt_nombre_grupo_real <> "" Then
      If var_hubo_cambios Then
         var_grupo_real_regreso = txt_grupo_real
         rs.Open "Select * from tb_gruposreales where vcha_gre_grupo_real_id = '" + txt_grupo_real + "'", cnn_importacion, adOpenDynamic, adLockOptimistic
         ok = TB_GRUPOSREALES.Anadir(txt_grupo_actual, txt_grupo_real, txt_nombre_grupo_real, txt_descuento_1, txt_descuento_2, txt_descuento_3, txt_prioridad)
         If Trim(var_grupo_real_regreso) <> "" Then
            txt_grupo_real = var_grupo_real_regreso
         End If
         If ok Then
            rsaux4.Open "update tb_gruposreales set vcha_emp_empresa_id = '" + var_empresa + "' where vcha_gre_grupo_real_id = '" + Me.txt_grupo_real + "'", cnn_importacion, adOpenDynamic, adLockOptimistic
            bitacora = True
            If var_modifica_registro_gr = False Then
               var_operacion_bitacora = "I"
               bitacora = TB_BITACORA_GRUPOS_REALES.Anadir(txt_grupo_real, "VCHA_GRE_NOMBRE", var_operacion_bitacora, "", txt_grupo_real, var_clave_usuario_global, fun_NombrePc, Date)
            Else
               var_operacion_bitacora = "M"
               If rs(0) <> txt_grupo_actual Then
                  bitacora = TB_BITACORA_GRUPOS_REALES.Anadir(txt_grupo_real, "VCHA_GAC_GRUPO_ACTUAL_ID", var_operacion_bitacora, rs(0), txt_grupo_actual, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(1) <> txt_grupo_real Then
                  bitacora = TB_BITACORA_GRUPOS_REALES.Anadir(txt_grupo_real, "VCHA_GRE_GRUPO_REAL_ID", var_operacion_bitacora, rs(1), txt_grupo_real, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(2) <> txt_nombre_grupo_real Then
                  bitacora = TB_BITACORA_GRUPOS_REALES.Anadir(txt_grupo_real, "VCHA_GRE_NOMBRE", var_operacion_bitacora, rs(2), txt_nombre_grupo_real, var_clave_usuario_global, fun_NombrePc, Date)
               End If
            End If
            rs.Close
         Else
            MsgBox "No se puede grabar registro: " + TB_GRUPOSREALES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
         End If
       End If
   End If
   Set TB_GRUPOSREALES = Nothing
End Sub

Sub pro_elimina_gruposreales()
   Dim var_llave_usuarios As String
   Set TB_GRUPOSREALES = New TB_GRUPOSREALES
   Set TB_BITACORA_GRUPOS_REALES = New TB_BITACORA_GRUPOS_REALES
   ok = True
   On Error GoTo salir:
   If txt_grupo_actual <> "" And txt_grupo_real <> "" Then
      'If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_GRUPOSREALES.Eliminar(txt_grupo_real)
      'Else
      '   GoTo salir:
      'End If
      If ok Then
         var_operacion_bitacora = "E"
         bitacora = TB_BITACORA_GRUPOS_REALES.Anadir(txt_grupo_real, "VCHA_GRE_NOMBRE", var_operacion_bitacora, txt_grupo_real, "", var_clave_usuario_global, fun_NombrePc, Date)
       Else
         MsgBox "No se puede eliminar registro: " + TB_GRUPOSREALES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
       End If
    End If
salir:
   Set TB_GRUPOSREALES = Nothing
End Sub


Sub pro_llena_listview1()
   Dim list_item As ListItem
   rs.Open "select * from TB_gruposreales  where vcha_emp_empresa_id = '" + var_empresa + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
   numero_items_gruposreales = 0
   While Not rs.EOF
      Set list_item = lv_gruposreales.ListItems.Add(, , rs!vcha_gre_grupo_real_id)
      list_item.SubItems(1) = IIf(IsNull(rs!VCHA_GRE_NOMBRE), "", rs!VCHA_GRE_NOMBRE)
      list_item.SubItems(2) = IIf(IsNull(rs!VCHA_GAC_GRUPO_aCTUAL_ID), "", rs!VCHA_GAC_GRUPO_aCTUAL_ID)
      list_item.SubItems(3) = IIf(IsNull(rs!floa_gre_descuento_1), 0, rs!floa_gre_descuento_1)
      list_item.SubItems(4) = IIf(IsNull(rs!floa_gre_descuento_2), 0, rs!floa_gre_descuento_2)
      list_item.SubItems(5) = IIf(IsNull(rs!floa_gre_descuento_3), 0, rs!floa_gre_descuento_3)
      list_item.SubItems(6) = IIf(IsNull(rs!CHAR_PRI_PRIORIDAD_ID), "", rs!CHAR_PRI_PRIORIDAD_ID)
      rs.MoveNext:
      numero_items_gruposreales = numero_items_gruposreales + 1
    Wend
    rs.Close
End Sub


Sub pro_textos()
On Error GoTo err0:
Dim var_n As Double
   var_n = lv_gruposreales.ListItems.Count
   If var_n > 0 Then
      txt_grupo_actual = lv_gruposreales.selectedItem.SubItems(2)
      txt_grupo_real = lv_gruposreales.selectedItem
      txt_nombre_grupo_real = lv_gruposreales.selectedItem.SubItems(1)
      txt_descuento_1 = lv_gruposreales.selectedItem.SubItems(3)
      txt_descuento_2 = lv_gruposreales.selectedItem.SubItems(4)
      txt_descuento_3 = lv_gruposreales.selectedItem.SubItems(5)
      txt_prioridad = lv_gruposreales.selectedItem.SubItems(6)
      rs.Open "select * from tb_gruposactuales where vcha_gac_grupo_actual_id = '" + txt_grupo_actual + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_grupo_actual = IIf(IsNull(rs1vcha_gac_nombre), "", rs!vcha_gac_nombre)
      Else
         txt_nombre_grupo_actual = ""
      End If
      rs.Close
      txt_grupo_real.Enabled = False
   End If
   var_numero_renglones = lv_gruposreales.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_gruposreales.ColumnHeaders(2).Width = 3850
   Else
      lv_gruposreales.ColumnHeaders(2).Width = 4099.71
   End If
   var_modifica_registro_gr = True
   var_hubo_cambios = False
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro_gr = False Then
        Set list_item = lv_gruposreales.ListItems.Add(, , txt_grupo_real)
        list_item.SubItems(1) = txt_nombre_grupo_real
        list_item.SubItems(2) = txt_grupo_actual
        list_item.SubItems(3) = txt_descuento_1
        list_item.SubItems(4) = txt_descuento_2
        list_item.SubItems(5) = txt_descuento_3
        list_item.SubItems(6) = txt_prioridad
        list_item.EnsureVisible
        list_item.Selected = True
        numero_items_gruposreales = numero_items_gruposreales + 1
    Else
        lv_gruposreales.ListItems.Item(lv_gruposreales.selectedItem.Index).Checked = False
        lv_gruposreales.ListItems.Item(lv_gruposreales.selectedItem.Index) = txt_grupo_real
        lv_gruposreales.ListItems.Item(lv_gruposreales.selectedItem.Index).ListSubItems(1) = txt_nombre_grupo_real
        lv_gruposreales.ListItems.Item(lv_gruposreales.selectedItem.Index).ListSubItems(2) = txt_grupo_actual
        lv_gruposreales.ListItems.Item(lv_gruposreales.selectedItem.Index).ListSubItems(3) = txt_descuento_1
        lv_gruposreales.ListItems.Item(lv_gruposreales.selectedItem.Index).ListSubItems(4) = txt_descuento_2
        lv_gruposreales.ListItems.Item(lv_gruposreales.selectedItem.Index).ListSubItems(5) = txt_descuento_3
        lv_gruposreales.ListItems.Item(lv_gruposreales.selectedItem.Index).ListSubItems(6) = txt_prioridad
        lv_gruposreales.ListItems.Item(lv_gruposreales.selectedItem.Index).Selected = True
    End If
'    lv_gruposreales.SetFocus
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
   Dim cmd As New Command
   Dim var_si As Integer
   Dim var_existe As Boolean
   var_si = MsgBox("¿Deseas crear el grupo actual automaticamente?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      rsaux.Open "SELECT * FROM TB_GRUPOSACTUALES WHERE VCHA_GAC_NOMBRE = '" + txt_nombre_grupo_real + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      var_existe = False
      If Not rsaux.EOF Then
         var_existe = True
      End If
      rsaux.Close
      If var_existe = False Then
         rsaux5.Open "select vcha_emp_conexion from tb_empresas where len(vcha_emp_conexion) > 0", cnn_distribucion, adOpenDynamic, adLockOptimistic
         While Not rsaux5.EOF
               var_conexion_importacion = IIf(IsNull(rsaux5(0).Value), "", rsaux5(0).Value)
               If Trim(var_conexion_importacion) <> "" Then
                  If cnn_importacion.State = 1 Then
                     cnn_importacion.Close
                  End If
                  cnn_importacion.Open var_conexion_importacion
                  Set cmd.ActiveConnection = cnn_importacion
                  cmd.CommandType = adCmdStoredProc
                  cmd.CommandText = "SP_CREACION_GRUPO_ACTUAL"
                  cmd("@GRUPO_REAL") = txt_grupo_real
                  cmd("@NOMBRE_GRUPO_REAL") = txt_nombre_grupo_real
                  cmd("@CLAVE_STRING") = ""
                  cmd.execute
                  txt_grupo_actual = cmd("@CLAVE_STRING")
                  txt_nombre_grupo_actual = txt_nombre_grupo_real
                  Set cmd = Nothing
                  rsaux4.Open "update tb_gruposactuales set vcha_emp_empresa_id = '" + var_empresa + "' where vcha_gac_grupo_actual_id = '" + txt_grupo_actual + "'", cnn_importacion, adOpenDynamic, adLockOptimistic
               End If
               rsaux5.MoveNext
         Wend
         rsaux5.Close
      Else
         var_si = MsgBox("Ya existe un grupo con el nombre " + txt_nombre_grupo_real + " ¿Deseas darlo de alta?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            
            rsaux5.Open "select vcha_emp_conexion from tb_empresas where len(vcha_emp_conexion) > 0", cnn_distribucion, adOpenDynamic, adLockOptimistic
            While Not rsaux5.EOF
                  var_conexion_importacion = IIf(IsNull(rsaux5(0).Value), "", rsaux5(0).Value)
                  If Trim(var_conexion_importacion) <> "" Then
                     If cnn_importacion.State = 1 Then
                        cnn_importacion.Close
                     End If
                     cnn_importacion.Open var_conexion_importacion
                     Set cmd.ActiveConnection = cnn_importacion
                     cmd.CommandType = adCmdStoredProc
                     cmd.CommandText = "SP_CREACION_GRUPO_ACTUAL"
                     cmd("@GRUPO_REAL") = txt_grupo_real
                     cmd("@NOMBRE_GRUPO_REAL") = txt_nombre_grupo_real
                     cmd("@CLAVE_STRING") = ""
                     cmd.execute
                     txt_grupo_actual = cmd("@CLAVE_STRING")
                     txt_nombre_grupo_actual = txt_nombre_grupo_real
                     Set cmd = Nothing
                  End If
                  rsaux5.MoveNext
            Wend
            rsaux5.Close
            
         End If
      End If
   Else
      MsgBox "Se a cancelado la creación del grupo actual", vbOKOnly, "ATENCION"
   End If
End Sub


Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_gruposreales, txt_buscar, True)
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
      txt_descuento_1 = "0"
   Else
      If CDbl(txt_descuento_1) > 100 Then
         MsgBox "Descuento incorrecto", vbOKOnly, "ATENCION"
         txt_descuento_1 = 0
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
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_descuento_2_LostFocus()
   If Not IsNumeric(txt_descuento_2) Then
      MsgBox "Descuento incorrecto", vbOKOnly, "ATENCION"
      txt_descuento_2 = "0"
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
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_descuento_3_LostFocus()
   If Not IsNumeric(txt_descuento_3) Then
      MsgBox "Descuento incorrecto", vbOKOnly, "ATENCION"
      txt_descuento_3 = "0"
   Else
      If CDbl(txt_descuento_3) > 100 Then
         MsgBox "Descuento incorrecto", vbOKOnly, "ATENCION"
         txt_descuento_3 = 0
      End If
   End If
End Sub

Private Sub txt_grupo_actual_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_grupo_actual_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para accesar al catálogo"
End Sub

Private Sub txt_grupo_actual_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_gruposactuales where vcha_emp_empresa_id = '" + var_empresa + "' order by vcha_gac_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_GAC_GRUPO_aCTUAL_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_gac_nombre), "", rs!vcha_gac_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "GRUPOS ACTUALES"
      var_tipo_lista = 1
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
      var_activa_forma_gruposactuales = Me.Name
      frmgruposactuales.Show
   End If
End Sub

Private Sub txt_grupo_actual_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_grupo_actual_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_grupo_actual) <> "" Then
      rs.Open "select * from TB_GRUPOSACTUALES where vcha_gac_grupo_actual_id = '" + txt_grupo_actual + "' AND VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn_distribucion, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_grupo_actual = IIf(IsNull(rs!vcha_gac_nombre), "", rs!vcha_gac_nombre)
      Else
         txt_grupo_actual = ""
         txt_nombre_grupo_actual = ""
      End If
      rs.Close
   Else
      txt_nombre_grupo_actual = ""
   End If
End Sub

Private Sub txt_grupo_real_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_grupo_real_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_grupo_actual_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_grupo_actual_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para accesar al catálogo"
End Sub

Private Sub txt_nombre_grupo_actual_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_gruposactuales order by vcha_gac_nombre", cnn_distribucion, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_GAC_GRUPO_aCTUAL_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_gac_nombre), "", rs!vcha_gac_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "GRUPOS ACTUALES"
      var_tipo_lista = 1
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
      var_activa_forma_gruposactuales = Me.Name
      frmgruposactuales.Show
   End If
End Sub

Private Sub txt_nombre_grupo_actual_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      If Me.cmd_guardar.Enabled = True Then
         Me.cmd_guardar.SetFocus
      End If
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_nombre_grupo_actual_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_grupo_real_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_grupo_real_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_prioridad_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_prioridad_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub
