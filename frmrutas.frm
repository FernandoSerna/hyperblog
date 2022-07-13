VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmrutas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo de rutas"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   Icon            =   "frmrutas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5955
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   150
      TabIndex        =   32
      Top             =   525
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   33
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
         TabIndex        =   34
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5445
      Picture         =   "frmrutas.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1470
      Picture         =   "frmrutas.frx":0F04
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      Picture         =   "frmrutas.frx":1006
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   810
      Picture         =   "frmrutas.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      Picture         =   "frmrutas.frx":11DA
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Guardar Alt + G"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmrutas.frx":12DC
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2715
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   27
      Top             =   75
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   " Rutas "
      Height          =   2805
      Left            =   150
      TabIndex        =   0
      Top             =   465
      Width           =   5655
      Begin VB.TextBox txt_nombre_agrupacion 
         Height          =   315
         Left            =   2220
         MaxLength       =   50
         TabIndex        =   18
         Top             =   2280
         Width           =   3285
      End
      Begin VB.TextBox txt_agrupacion 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   17
         Top             =   2280
         Width           =   915
      End
      Begin VB.TextBox txt_meta 
         Height          =   315
         Left            =   2985
         MaxLength       =   50
         TabIndex        =   16
         Top             =   1935
         Width           =   1350
      End
      Begin VB.TextBox txt_clasificacion 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   15
         Top             =   1935
         Width           =   900
      End
      Begin VB.TextBox txt_anterior 
         Height          =   330
         Left            =   2985
         TabIndex        =   14
         Top             =   1575
         Width           =   1350
      End
      Begin VB.TextBox txt_tolerancia 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   13
         Top             =   1583
         Width           =   900
      End
      Begin VB.TextBox txt_nombre_agente 
         Height          =   315
         Left            =   2205
         TabIndex        =   10
         Top             =   915
         Width           =   3285
      End
      Begin VB.TextBox txt_nombre_zona 
         Height          =   315
         Left            =   2205
         TabIndex        =   12
         Top             =   1245
         Width           =   3285
      End
      Begin VB.TextBox txt_zona 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1245
         Width           =   900
      End
      Begin VB.TextBox txt_agente 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   9
         Top             =   915
         Width           =   900
      End
      Begin VB.TextBox txt_nombre_ruta 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   8
         Top             =   585
         Width           =   4200
      End
      Begin VB.TextBox txt_ruta 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   7
         Top             =   255
         Width           =   900
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Agrupación:"
         Height          =   195
         Index           =   6
         Left            =   330
         TabIndex        =   38
         Top             =   2340
         Width           =   855
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Meta:"
         Height          =   195
         Index           =   4
         Left            =   2325
         TabIndex        =   37
         Top             =   1995
         Width           =   405
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clasificación:"
         Height          =   195
         Index           =   3
         Left            =   330
         TabIndex        =   36
         Top             =   1995
         Width           =   930
      End
      Begin VB.Label Anterior 
         AutoSize        =   -1  'True
         Caption         =   "Anterior:"
         Height          =   195
         Left            =   2355
         TabIndex        =   35
         Top             =   1643
         Width           =   585
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Tolerancia:"
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   31
         Top             =   1643
         Width           =   795
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Zona:"
         Height          =   195
         Index           =   7
         Left            =   330
         TabIndex        =   26
         Top             =   1305
         Width           =   420
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Index           =   5
         Left            =   330
         TabIndex        =   25
         Top             =   975
         Width           =   555
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Index           =   2
         Left            =   330
         TabIndex        =   20
         Top             =   645
         Width           =   885
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   1
         Left            =   330
         TabIndex        =   19
         Top             =   315
         Width           =   450
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   150
      TabIndex        =   21
      Top             =   3255
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   1575
         TabIndex        =   30
         Top             =   165
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3705
         TabIndex        =   29
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
         Caption         =   "Busqueda de ruta:"
         Height          =   195
         Left            =   195
         TabIndex        =   22
         Top             =   195
         Width           =   1305
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3465
      Left            =   150
      TabIndex        =   23
      Top             =   3795
      Width           =   5655
      Begin MSComctlLib.ListView lv_rutas 
         Height          =   3285
         Left            =   45
         TabIndex        =   28
         Top             =   120
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   5794
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   6879
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
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Anterior"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "clasificacion"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "meta"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "agrupacion"
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
               Picture         =   "frmrutas.frx":13DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmrutas.frx":1CB8
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   24
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
            Picture         =   "frmrutas.frx":2592
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmrutas.frx":2E6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmrutas.frx":3746
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmrutas.frx":3CE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmrutas.frx":45BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmrutas.frx":4E98
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmrutas.frx":5772
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmrutas.frx":5884
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmrutas.frx":5996
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmrutas.frx":5AA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmrutas.frx":5BBA
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmrutas"
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
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         rsaux5.Open "select vcha_emp_conexion from tb_empresas where len(vcha_emp_conexion) > 0", cnn, adOpenDynamic, adLockOptimistic
         While Not rsaux5.EOF
               var_conexion_importacion = IIf(IsNull(rsaux5(0).Value), "", rsaux5(0).Value)
               If Trim(var_conexion_importacion) <> "" Then
                  If cnn_importacion.State = 1 Then
                     cnn_importacion.Close
                  End If
                  cnn_importacion.Open var_conexion_importacion
                  Call pro_elimina_rutas
               End If
               rsaux5.MoveNext
         Wend
         rsaux5.Close
      
         numero_items_rutas = numero_items_rutas - 1
         MsgBox "Se Elimino Correctamente el Registro", vbInformation
         lv_rutas.ListItems.Remove (lv_rutas.selectedItem.Index)
         Call pro_limpiatextos(Me)
         txt_registros = lv_rutas.ListItems.Count
         lv_rutas.selectedItem.Selected = True
         pro_textos
      
      
         rs.Open "select * from tb_rutas", cnn, adOpenDynamic, adLockOptimistic
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
Dim var_posible As Boolean
   var_posible = True
   If var_modifica_registro_ruta = False Then
      rs.Open "select * from tb_rutas where vcha_rut_ruta_id = '" + Me.txt_ruta + "'", cnn, adOpenDynamic, adLockOptimistic
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
         
         rsaux5.Open "select vcha_emp_conexion from tb_empresas where len(vcha_emp_conexion) > 0", cnn, adOpenDynamic, adLockOptimistic
         While Not rsaux5.EOF
               var_conexion_importacion = IIf(IsNull(rsaux5(0).Value), "", rsaux5(0).Value)
               If Trim(var_conexion_importacion) <> "" Then
                  If cnn_importacion.State = 1 Then
                     cnn_importacion.Close
                  End If
                  cnn_importacion.Open var_conexion_importacion
                  Call pro_guardar_rutas
               End If
               rsaux5.MoveNext
         Wend
         rsaux5.Close
         
         pro_actualiza_ListView
         txt_ruta.Enabled = False
         MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
         txt_registros = lv_rutas.ListItems.Count
         var_modifica_registro_ruta = True
         var_hubo_cambios = False
         
         rs.Open "select * from tb_rutas", cnn, adOpenDynamic, adLockOptimistic
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
           Call gPrintListView(lv_rutas, "LISTADO DE rutas")
        End If

End Sub

Private Sub cmd_nuevo_Click()
        Call pro_limpiatextos(Me)
        txt_ruta.Enabled = True
        txt_ruta.SetFocus: var_modifica_registro_ruta = False
        cmd_guardar.Enabled = True
        cmd_deshacer.Enabled = True
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_ruta = False Then
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
   var_modifica_registro_ruta = True
   lv_rutas.SmallIcons = ImageList
   Call pro_encabezadosView(Me, lv_rutas, False)
   Call pro_llena_listview1
   pro_textos
   rs.Open "select * from tb_rutas", cnn, adOpenDynamic, adLockOptimistic
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
   var_modifica_registro_ruta = False
   Call activa_forma(var_activa_forma_rutas)
End Sub


Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_tipo_lista = 1 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_agente = lv_lista.selectedItem
            txt_nombre_agente = lv_lista.selectedItem.SubItems(1)
         Else
            txt_agente = ""
            txt_nombre_agente = ""
         End If
         txt_agente.SetFocus
      End If
      If var_tipo_lista = 2 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_zona = lv_lista.selectedItem
            txt_nombre_zona = lv_lista.selectedItem.SubItems(1)
         Else
            txt_zona = ""
            txt_nombre_zona = ""
         End If
         txt_zona.SetFocus
      End If
      If var_tipo_lista = 3 Then
         If lv_lista.ListItems.Count > 0 Then
            Me.txt_agrupacion = lv_lista.selectedItem
            Me.txt_nombre_agrupacion = lv_lista.selectedItem.SubItems(1)
         Else
            Me.txt_agrupacion = ""
            Me.txt_nombre_agrupacion = ""
         End If
         Me.txt_agrupacion.SetFocus
      End If
      frm_lista.Visible = False
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub lv_rutas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_rutas, ColumnHeader)
End Sub

Private Sub lv_rutas_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_rutas.selectedItem = Item
        pro_textos
        var_modifica_registro_ruta = True
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_rutas order by vcha_rut_NOMBRE", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_rut_ruta_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_rut_nombre), "", rs!vcha_rut_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "RUTAS"
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
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_rutas.SetFocus
      Call pro_avanzar(Me, lv_rutas, Button)
      lv_rutas.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_rutas.ListItems(1).Selected = True
      lv_rutas.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 4 Then
      numero_items_rutas = lv_rutas.ListItems.Count
      lv_rutas.ListItems(numero_items_rutas).Selected = True
      pro_textos
      lv_rutas.selectedItem.EnsureVisible
   End If
err0:
End Sub


Sub pro_guardar_rutas()
   Dim ok As Boolean
   Set TB_RUTAS = New TB_RUTAS
   Set TB_BITACORA_RUTAS = New TB_BITACORA_RUTAS
   If txt_ruta <> "" And txt_nombre_ruta <> "" Then
      If var_hubo_cambios Then
         rs.Open "SELECT * FROM TB_RUTAS WHERE VCHA_RUT_RUTA_ID = '" + txt_ruta + "'", cnn_importacion, adOpenDynamic, adLockOptimistic
         If Trim(txt_tolerancia) = "" Then
            txt_tolerancia = 0
         End If
         If Not IsNumeric(Me.txt_meta) Then
            Me.txt_meta = 0
         End If
         If Me.txt_agrupacion = "" Then
            Me.txt_agrupacion = Me.txt_ruta
         End If
         rsaux10.Open "select * from tb_rutas where vcha_rut_ruta_id = '" + Me.txt_agrupacion + "'", cnn, adOpenDynamic, adLockOptimistic
         If rsaux10.EOF Then
            Me.txt_agrupacion = Me.txt_ruta
         End If
         rsaux10.Close
         ok = TB_RUTAS.Anadir(txt_ruta, txt_nombre_ruta, txt_agente, txt_zona, txt_tolerancia)
         If ok Then
            rsaux4.Open "UPDATE TB_RUTAS SET VCHA_RUT_RUTA_ANTERIOR_ID = '" + txt_anterior + "', VCHA_RUT_CLASIFICACION_COMISION = '" + Me.txt_clasificacion + "', floa_rut_meta = " + Me.txt_meta + ", VCHA_RUT_AGRUPADOR_REPORTE = '" + Me.txt_agrupacion + "' WHERE VCHA_RUT_RUTA_ID = '" + Me.txt_ruta + "'", cnn_importacion, adOpenDynamic, adLockOptimistic
            bitacora = True
            If var_modifica_registro_ruta = False Then
               var_operacion_bitacora = "I"
               bitacora = TB_BITACORA_RUTAS.Anadir(txt_ruta, "VCHA_RUT_NOMBRE", var_operacion_bitacora, "", txt_nombre_ruta, var_clave_usuario_global, fun_NombrePc, Date)
            Else
               var_operacion_bitacora = "M"
               If rs(0).Value <> txt_ruta Then
                  bitacora = TB_BITACORA_RUTAS.Anadir(txt_ruta, "VCHA_RUT_PRODUCTO_ID", var_operacion_bitacora, rs(0), txt_ruta, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(1).Value <> txt_nombre_ruta Then
                  bitacora = TB_BITACORA_RUTAS.Anadir(txt_ruta, "VCHA_RUT_NOMBRE", var_operacion_bitacora, rs(1), txt_nombre_ruta, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(2).Value <> txt_agente Then
                  bitacora = TB_BITACORA_RUTAS.Anadir(txt_ruta, "VCHA_AGE_AGENTE_ID", var_operacion_bitacora, rs(2), txt_agente, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(3).Value <> txt_zona Then
                  bitacora = TB_BITACORA_RUTAS.Anadir(txt_ruta, "VCHA_ZON_ZONA_NOMBRE", var_operacion_bitacora, rs(3), txt_zona, var_clave_usuario_global, fun_NombrePc, Date)
               End If
            End If
            rs.Close
         Else
            MsgBox "No se puede grabar registro: " + TB_RUTAS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
         End If
      End If
   End If
   Set TB_RUTAS = Nothing
End Sub

Sub pro_elimina_rutas()
   Dim var_llave_usuarios As String
   Set TB_RUTAS = New TB_RUTAS
   Set TB_BITACORA_RUTAS = New TB_BITACORA_RUTAS
   ok = True
   On Error GoTo salir:
   If txt_ruta <> "" And txt_nombre_ruta <> "" And var_modifica_registro_ruta = True Then
      'If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_RUTAS.Eliminar(txt_ruta)
      'Else
         GoTo salir:
      'End If
      If ok Then
         var_operacion_bitacora = "E"
         bitacora = TB_BITACORA_RUTAS.Anadir(txt_ruta, "VCHA_RUT_NOMBRE", var_operacion_bitacora, txt_nombre_ruta, "", var_clave_usuario_global, fun_NombrePc, Date)
      Else
         MsgBox "No se puede grabar registro: " + TB_RUTAS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
   Set TB_RUTAS = Nothing
End Sub


Sub pro_llena_listview1()
   Dim list_item As ListItem
   numero_items_rutas = 0
   rs.Open "select * from TB_rutas", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set list_item = lv_rutas.ListItems.Add(, , rs!vcha_rut_ruta_id)
         list_item.SubItems(1) = IIf(IsNull(rs!vcha_rut_nombre), "", rs!vcha_rut_nombre)
         list_item.SubItems(2) = IIf(IsNull(rs!vcha_age_agente_id), "", rs!vcha_age_agente_id)
         list_item.SubItems(3) = IIf(IsNull(rs!VCHA_ZON_ZONA_ID), "", rs!VCHA_ZON_ZONA_ID)
         list_item.SubItems(4) = IIf(IsNull(rs!inte_rut_tolerancia), 0, rs!inte_rut_tolerancia)
         list_item.SubItems(5) = IIf(IsNull(rs!VCHA_rut_ruta_ANTERIOR_ID), 0, rs!VCHA_rut_ruta_ANTERIOR_ID)
         list_item.SubItems(6) = IIf(IsNull(rs!vcha_rut_clasificacion_comision), "", rs!vcha_rut_clasificacion_comision)
         list_item.SubItems(7) = IIf(IsNull(rs!floa_rut_meta), 0, rs!floa_rut_meta)
         list_item.SubItems(8) = IIf(IsNull(rs!vcha_rut_agrupador_Reporte), "", rs!vcha_rut_agrupador_Reporte)
         
         rs.MoveNext:
         numero_items_rutas = numero_items_rutas + 1
   Wend
   rs.Close
End Sub


Sub pro_textos()
'On Error GoTo err0:
Dim var_n As Double
   var_n = lv_rutas.ListItems.Count
   Me.txt_ruta.Enabled = False
   If var_n > 0 Then
      txt_ruta = lv_rutas.selectedItem
      txt_nombre_ruta = lv_rutas.selectedItem.SubItems(1)
      txt_agente = lv_rutas.selectedItem.SubItems(2)
      txt_zona = lv_rutas.selectedItem.SubItems(3)
      txt_tolerancia = lv_rutas.selectedItem.SubItems(4)
      txt_anterior = lv_rutas.selectedItem.SubItems(5)
      Me.txt_clasificacion = lv_rutas.selectedItem.SubItems(6)
      Me.txt_meta = lv_rutas.selectedItem.SubItems(7)
      Me.txt_agrupacion = lv_rutas.selectedItem.SubItems(8)
      rs.Open "select * from tb_rutas where vcha_rut_ruta_id = '" + Me.txt_agrupacion + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_agrupacion = IIf(IsNull(rs!vcha_rut_nombre), "", rs!vcha_rut_nombre)
      Else
         Me.txt_nombre_agrupacion = ""
      End If
      rs.Close
      rs.Open "select * from tb_agentes where vcha_age_agente_id = '" + txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_agente = IIf(IsNull(rs!vcha_age_nombre), "", rs!vcha_age_nombre)
      Else
         txt_nombre_agente = ""
      End If
      rs.Close
      rs.Open "select * from tb_zonas where vcha_zon_zona_id = '" + txt_zona + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_zona = IIf(IsNull(rs!VCHA_ZON_DESCRIPCION), "", rs!VCHA_ZON_DESCRIPCION)
      Else
         txt_nombre_zona = ""
      End If
      rs.Close
   End If
   var_numero_renglones = lv_rutas.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_rutas.ColumnHeaders(2).Width = 3850
   Else
      lv_rutas.ColumnHeaders(2).Width = 4099.71
   End If
   var_hubo_cambios = False
   var_modifica_registro_ruta = True
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro_ruta = False Then
        Set list_item = lv_rutas.ListItems.Add(, , txt_ruta)
        list_item.SubItems(1) = txt_nombre_ruta
        list_item.SubItems(2) = txt_agente
        list_item.SubItems(3) = txt_zona
        list_item.SubItems(4) = txt_tolerancia
        list_item.SubItems(5) = txt_anterior
        list_item.SubItems(6) = Me.txt_clasificacion
        list_item.SubItems(7) = Me.txt_meta
        list_item.SubItems(8) = Me.txt_agrupacion
        
        list_item.EnsureVisible
        list_item.Selected = True
        numero_items_rutas = numero_items_rutas + 1
    Else
        lv_rutas.ListItems.Item(lv_rutas.selectedItem.Index).Checked = False
        lv_rutas.ListItems.Item(lv_rutas.selectedItem.Index) = txt_ruta
        lv_rutas.ListItems.Item(lv_rutas.selectedItem.Index).ListSubItems(1) = txt_nombre_ruta
        lv_rutas.ListItems.Item(lv_rutas.selectedItem.Index).ListSubItems(2) = txt_agente
        lv_rutas.ListItems.Item(lv_rutas.selectedItem.Index).ListSubItems(3) = txt_zona
        lv_rutas.ListItems.Item(lv_rutas.selectedItem.Index).ListSubItems(4) = txt_tolerancia
        lv_rutas.ListItems.Item(lv_rutas.selectedItem.Index).ListSubItems(5) = txt_anterior
        lv_rutas.ListItems.Item(lv_rutas.selectedItem.Index).ListSubItems(6) = Me.txt_clasificacion
        lv_rutas.ListItems.Item(lv_rutas.selectedItem.Index).ListSubItems(7) = Me.txt_meta
        lv_rutas.ListItems.Item(lv_rutas.selectedItem.Index).ListSubItems(8) = Me.txt_agrupacion
        lv_rutas.ListItems.Item(lv_rutas.selectedItem.Index).Selected = True
    End If
    lv_rutas.SetFocus
End Sub





Private Sub txt_agente_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_agente_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para accesar al catálogo"
End Sub

Private Sub txt_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_agentes order by vcha_age_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_age_agente_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_age_nombre), "", rs!vcha_age_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
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
      var_activa_forma_agentes = Me.Name
      Me.Enabled = False
      frmagentes.Show
   End If
End Sub

Private Sub txt_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_agente_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_agente) <> "" Then
      rs.Open "select * from tb_agentes where vcha_age_agente_id = '" + txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_agente = IIf(IsNull(rs!vcha_age_nombre), "", rs!vcha_age_nombre)
      Else
         txt_nombre_agente = ""
         txt_agente = ""
         MsgBox "Clave de agente incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      txt_nombre_agente = ""
   End If
End Sub

Private Sub txt_agrupacion_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_agrupacion_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_rutas order by vcha_rut_NOMBRE", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_rut_ruta_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_rut_nombre), "", rs!vcha_rut_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "RUTAS"
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
End Sub

Private Sub txt_agrupacion_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_agrupacion_LostFocus()
   If Trim(Me.txt_agrupacion) <> "" Then
      rs.Open "select * from tb_rutas where vcha_rut_ruta_id = '" + Me.txt_agrupacion + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_agrupacion = IIf(IsNull(rs!vcha_rut_nombre), "", rs!vcha_rut_nombre)
      Else
          MsgBox "La ruta no existe", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      Me.txt_nombre_agrupacion = ""
   End If
End Sub

Private Sub txt_anterior_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_anterior_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_rutas, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub

Private Sub txt_clasificacion_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_clasificacion_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_meta_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_meta_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_agente_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para accesar al catálogo"
End Sub

Private Sub txt_nombre_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_agentes order by vcha_age_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_age_agente_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_age_nombre), "", rs!vcha_age_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
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
      var_activa_forma_agentes = Me.Name
      Me.Enabled = False
      frmagentes.Show
   End If
End Sub

Private Sub txt_nombre_agente_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 13
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_agente_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_agrupacion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub

Private Sub txt_nombre_ruta_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_ruta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_zona_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_zona_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para accesar al catálogo"
End Sub

Private Sub txt_nombre_zona_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_ZONAS order by vcha_zon_descripcion", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ZON_ZONA_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ZON_DESCRIPCION), "", rs!VCHA_ZON_DESCRIPCION)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ZONAS"
      var_tipo_lista = 2
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
      var_activa_forma_zonas = Me.Name
      Me.Enabled = False
      frmzonas.Show
   End If
End Sub

Private Sub txt_nombre_zona_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 13
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_zona_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_ruta_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_ruta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_tolerancia_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_tolerancia_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Me.txt_anterior.SetFocus
   End If
End Sub

Private Sub txt_zona_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_zona_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para accesar al catálogo"
End Sub

Private Sub txt_zona_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_ZONAS order by vcha_zon_descripcion", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ZON_ZONA_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ZON_DESCRIPCION), "", rs!VCHA_ZON_DESCRIPCION)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ZONAS"
      var_tipo_lista = 2
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
      var_activa_forma_zonas = Me.Name
      Me.Enabled = False
      frmzonas.Show
   End If
End Sub

Private Sub txt_zona_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_zona_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_zona) <> "" Then
      rs.Open "Select * from tb_zonas where vcha_zon_zona_id = '" + txt_zona + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_zona = IIf(IsNull(rs!VCHA_ZON_DESCRIPCION), "", rs!VCHA_ZON_DESCRIPCION)
      Else
         txt_nombre_zona = ""
         txt_zona = ""
         MsgBox "Clave de zona incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      txt_nombre_zona = ""
   End If
End Sub
