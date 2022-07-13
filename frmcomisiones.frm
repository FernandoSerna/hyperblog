VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcomisiones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comisiones"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   Icon            =   "frmcomisiones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5955
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   180
      TabIndex        =   24
      Top             =   2790
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
   Begin VB.Frame frm_agente_clonar 
      Height          =   1320
      Left            =   150
      TabIndex        =   28
      Top             =   495
      Width           =   5670
      Begin VB.TextBox txt_agente 
         Height          =   315
         Left            =   795
         TabIndex        =   12
         Top             =   840
         Width           =   900
      End
      Begin VB.TextBox txt_nombre_agente 
         Height          =   315
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   840
         Width           =   3900
      End
      Begin VB.CommandButton cmd_cancelar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   390
         Picture         =   "frmcomisiones.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Cancelar"
         Top             =   420
         Width           =   330
      End
      Begin VB.CommandButton cmd_aceptar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   60
         Picture         =   "frmcomisiones.frx":0A14
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Aceptar"
         Top             =   420
         Width           =   330
      End
      Begin VB.Frame Frame2 
         Height          =   120
         Left            =   30
         TabIndex        =   30
         Top             =   660
         Width           =   5580
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Index           =   3
         Left            =   165
         TabIndex        =   31
         Top             =   900
         Width           =   555
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   29
         Top             =   135
         Width           =   5595
      End
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1800
      Picture         =   "frmcomisiones.frx":0B5E
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Clonar Comisiones"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmcomisiones.frx":0C60
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
      Picture         =   "frmcomisiones.frx":0D62
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
      Picture         =   "frmcomisiones.frx":0E64
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
      Picture         =   "frmcomisiones.frx":0F36
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
      Picture         =   "frmcomisiones.frx":1038
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
      Picture         =   "frmcomisiones.frx":113A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   19
      Top             =   285
      Width           =   5655
   End
   Begin VB.Frame Frame1 
      Caption         =   " Comisión "
      Height          =   1680
      Left            =   150
      TabIndex        =   17
      Top             =   465
      Width           =   5655
      Begin VB.TextBox txt_nombre_linea 
         Height          =   315
         Left            =   2205
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         Top             =   255
         Width           =   3315
      End
      Begin VB.TextBox txt_linea 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   7
         Top             =   255
         Width           =   900
      End
      Begin VB.TextBox txt_limite_inferior 
         Height          =   315
         Left            =   1290
         TabIndex        =   9
         Top             =   585
         Width           =   1740
      End
      Begin VB.TextBox txt_limite_superior 
         Height          =   315
         Left            =   1290
         TabIndex        =   10
         Top             =   915
         Width           =   1740
      End
      Begin VB.TextBox txt_porcentaje 
         Height          =   315
         Left            =   1275
         TabIndex        =   11
         Top             =   1245
         Width           =   1740
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Porcentaje:"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   23
         Top             =   1305
         Width           =   810
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Limete Superior:"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   22
         Top             =   960
         Width           =   1140
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Limite Inferior:"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   21
         Top             =   645
         Width           =   975
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Linea:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   20
         Top             =   315
         Width           =   435
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2145
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   15
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   -15
      Top             =   945
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
            Picture         =   "frmcomisiones.frx":1774
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcomisiones.frx":204E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   960
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
            Picture         =   "frmcomisiones.frx":2928
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcomisiones.frx":3202
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcomisiones.frx":3ADC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcomisiones.frx":4078
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcomisiones.frx":4954
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcomisiones.frx":522E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcomisiones.frx":5B08
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcomisiones.frx":5C1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcomisiones.frx":5D2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcomisiones.frx":5E3E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   5055
      Left            =   150
      TabIndex        =   18
      Top             =   2145
      Width           =   5655
      Begin MSComctlLib.ListView lv_comisiones 
         Height          =   4860
         Left            =   30
         TabIndex        =   16
         Top             =   135
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   8573
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
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Porcentaje"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Limite inferior"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Limite superior"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "estatus"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "empresa"
            Object.Width           =   0
         EndProperty
      End
   End
End
Attribute VB_Name = "frmcomisiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim numero_items_comisiones As Integer
Dim var_tipo_lista As Integer

Private Sub cmd_aceptar_Click()
   Dim var_si As Integer
   If Trim(txt_nombre_agente) <> "" Then
   var_si = MsgBox("¿Desas clonar las comisiones del agente " + txt_nombre_agente, vbYesNo, "ATENCION")
   If var_si = 6 Then
      var_si = MsgBox("Confirmar la clonación del agente " + txt_nombre_agente, vbYesNo, "ATENCION")
      If var_si = 6 Then
         'MsgBox "exec SP_CLONAR_COMISIONES '" + txt_agente + "', '" + var_agente_seleccionado + "'"
         rs.Open "exec SP_CLONAR_COMISIONES '" + txt_agente + "', '" + var_agente_seleccionado + "'", cnn, adOpenDynamic, adLockOptimistic
         MsgBox "Se a terminado la clonación de las comisiones", vbOKOnly, "ATENCION"
         Call pro_llena_listview1
      End If
   End If
   Else
      MsgBox "Se debe de seleccionar un agente", vbOKOnly, "ATENCION"
   End If
   Me.frm_agente_clonar.Visible = False
End Sub

Private Sub cmd_cancelar_Click()
   Me.frm_agente_clonar.Visible = False
End Sub

Private Sub cmd_deshacer_Click()
   Call pro_textos
End Sub

Private Sub cmd_deshacer_GotFocus()
   Me.frm_agente_clonar.Visible = False
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
                  Call pro_elimina_comisiones
               End If
               rsaux5.MoveNext
         Wend
         rsaux5.Close
         
         MsgBox "Se Elimino Correctamente el Registro", vbInformation
         lv_comisiones.ListItems.Remove (lv_comisiones.selectedItem.Index)
         numero_items_comisiones = numero_items_comisiones - 1
         Call pro_limpiatextos(Me)
         txt_registros = lv_comisiones.ListItems.Count
         var_n = lv_comisiones.ListItems.Count
         If var_n > 0 Then
            lv_comisiones.selectedItem.Selected = True
         End If
         pro_textos
      
         rs.Open "select * from tb_comisiones", cnn, adOpenDynamic, adLockOptimistic
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

Private Sub cmd_eliminar_GotFocus()
   Me.frm_agente_clonar.Visible = False
End Sub

Private Sub cmd_guardar_Click()
   Dim var_comision_inferior As Double
   Dim var_comision_superior As Double
   
   If IsNumeric(Me.txt_limite_inferior) Then
      If IsNumeric(Me.txt_limite_superior) Then
         If IsNumeric(Me.txt_porcentaje) Then
            var_comision_inferior = CDbl(txt_limite_inferior)
            var_comision_superior = CDbl(txt_limite_superior)
            If var_comision_inferior <= var_comision_superior Then
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
                  If var_modifica_registro_comision = False Then
                     rsaux5.Open "select vcha_emp_conexion from tb_empresas where len(vcha_emp_conexion) > 0", cnn, adOpenDynamic, adLockOptimistic
                     While Not rsaux5.EOF
                           var_conexion_importacion = IIf(IsNull(rsaux5(0).Value), "", rsaux5(0).Value)
                           If Trim(var_conexion_importacion) <> "" Then
                              If cnn_importacion.State = 1 Then
                                 cnn_importacion.Close
                              End If
                              cnn_importacion.Open var_conexion_importacion
                              Call pro_guardar_comisiones
                           End If
                           rsaux5.MoveNext
                     Wend
                     rsaux5.Close
                     pro_actualiza_ListView
                     txt_linea.Enabled = False
                     txt_nombre_linea.Enabled = False
                     MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
                     txt_registros = lv_comisiones.ListItems.Count
                     var_modifica_registro_comision = True
                     var_hubo_cambios = False
                  Else
                     If var_hubo_cambios = True Then
                        MsgBox "Para actualizar la comision, debera de eliminarla con anterioridad", vbOKOnly, "ATENCION"
                     End If
                  End If
                  rs.Open "select * from tb_comisiones", cnn, adOpenDynamic, adLockOptimistic
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
                  var_n = lv_comisiones.ListItems.Count
                  If var_n > 0 Then
                     Me.lv_comisiones.SetFocus
                  Else
                     Me.cmd_nuevo.SetFocus
                  End If
               End If
            Else
               MsgBox "El limite inferior es mayor al superior", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Porcentaje incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Limite superior incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Limite inferior incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_guardar_GotFocus()
   Me.frm_agente_clonar.Visible = False
End Sub

Private Sub cmd_imprimir_Click()
        If vector_valida_passwords(var_indice_menu) = "*" Then
           frmpasswords.Show
        Else
           Call gPrintListView(lv_comisiones, "LISTADO DE comisiones")
        End If

End Sub

Private Sub cmd_imprimir_GotFocus()
   Me.frm_agente_clonar.Visible = False
End Sub

Private Sub cmd_nuevo_Click()
   Call pro_limpiatextos(Me)
   txt_linea.Enabled = True
   txt_linea.SetFocus: var_modifica_registro_comision = False
   txt_nombre_linea.Enabled = True
   cmd_guardar.Enabled = True
   cmd_deshacer.Enabled = True
End Sub

Private Sub cmd_nuevo_GotFocus()
   Me.frm_agente_clonar.Visible = False
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_comision = False Then
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







Private Sub cmd_salir_GotFocus()
   Me.frm_agente_clonar.Visible = False
End Sub

Private Sub Command1_Click()
   Me.frm_agente_clonar.Visible = True
   txt_agente.SetFocus
End Sub

Private Sub Command1_GotFocus()
   Me.frm_agente_clonar.Visible = False
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
   Left = 2900
   Me.frm_agente_clonar.Visible = False
   frm_lista.Visible = False
   numero_items_comisiones = 0
   var_modifica_registro_comision = True
   lv_comisiones.SmallIcons = ImageList1
   Call pro_encabezadosView(Me, lv_comisiones, False)
   Call pro_llena_listview1
   pro_textos
End Sub

Private Sub Form_Unload(Cancel As Integer)
    var_swpassword = False
    var_modifica_registro_comision = False
    Call activa_forma(var_activa_forma_comisiones)
End Sub

Private Sub lv_comisiones_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_comisiones, ColumnHeader)
End Sub

Private Sub lv_comisiones_GotFocus()
   Me.frm_agente_clonar.Visible = False
End Sub

Private Sub lv_comisiones_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_comisiones.selectedItem = Item
        pro_textos
End Sub



Sub pro_guardar_comisiones()
Dim ok As Boolean
Set TB_COMISIONES = New TB_COMISIONES
Set TB_BITACORA_COMISIONES = New TB_BITACORA_COMISIONES

   If txt_linea <> "" And txt_limite_inferior <> "" Then
      If var_hubo_cambios Then
         rs.Open "select * from tb_comisiones where vcha_age_agente_id = '" + var_agente_seleccionado + "' AND vcha_lin_linea_id  = '" + txt_linea + "'", cnn_importacion, adOpenDynamic, adLockOptimistic
         ok = TB_COMISIONES.Anadir(var_agente_seleccionado, txt_linea, txt_limite_inferior, txt_limite_superior, txt_porcentaje)
         If ok Then
            bitacora = True
            If var_modifica_registro_comision = False Then
               var_operacion_bitacora = "I"
               bitacora = TB_BITACORA_COMISIONES.Anadir(var_agente_seleccionado, txt_linea, "VCHA_LIN_NOMBRE", var_operacion_bitacora, "", txt_nombre_linea, var_clave_usuario_global, fun_NombrePc, Date)
            Else
               var_operacion_bitacora = "M"
               If rs(1).Value <> txt_linea Then
                  bitacora = TB_BITACORA_COMISIONES.Anadir(var_agente_seleccionado, txt_linea, "VCHA_LIN_LINEA_ID", var_operacion_bitacora, rs(0), txt_linea, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(2) <> txt_limite_inferior Then
                  bitacora = TB_BITACORA_COMISIONES.Anadir(var_agente_seleccionado, txt_linea, "VCHA_COM_LIMITE_INFERIOR", var_operacion_bitacora, rs(1), txt_limite_inferior, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(3) <> txt_limite_superior Then
                  bitacora = TB_BITACORA_COMISIONES.Anadir(var_agente_seleccionado, txt_linea, "VCHA_COM_LIMITE_SUPERIOR", var_operacion_bitacora, rs(2), txt_limite_superior, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(4) <> txt_porcentaje Then
                  bitacora = TB_BITACORA_COMISIONES.Anadir(var_agente_seleccionado, txt_linea, "VCHA_COM_PORCENTAJE", var_operacion_bitacora, rs(3), txt_porcentaje, var_clave_usuario_global, fun_NombrePc, Date)
               End If
            End If
            rs.Close
         Else
            MsgBox "No se puede grabar registro: " + TB_COMISIONES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
         End If
      End If
   End If
 Set TB_COMISIONES = Nothing

End Sub

Sub pro_elimina_comisiones()
   Dim var_llave_usuarios As String
   Set TB_COMISIONES = New TB_COMISIONES
   Set TB_BITACORA_COMISIONES = New TB_BITACORA_COMISIONES
   'On Error GoTo SALIR
   ok = True
   If txt_linea <> "" And txt_limite_inferior <> "" And var_modifica_registro_comision = True Then
      'If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         rsaux.Open "delete from tb_comisiones where vcha_age_agente_id = '" + var_agente_seleccionado + "' and vcha_lin_linea_id = '" + txt_linea + "' and FLOA_COM_LIMITE_INFERIOR = " + Me.txt_limite_inferior + " and FLOA_COM_LIMITE_SUPERIOR = " + Me.txt_limite_superior + " and FLOA_COM_PORCENTAJE = " + Me.txt_porcentaje, cnn, adOpenDynamic, adLockOptimistic
         ok = TB_COMISIONES.Eliminar(var_agente_seleccionado, txt_linea)
      'Else
      '   GoTo salir:
      'End If
      If ok Then
         var_operacion_bitacora = "E"
         bitacora = TB_BITACORA_COMISIONES.Anadir(var_agente_seleccionado, txt_limite_inferior, "VCHA_AGE_NOMBRE", var_operacion_bitacora, txt_nombre_linea, "", var_clave_usuario_global, fun_NombrePc, Date)
      Else
         MsgBox "No se puede grabar registro: " + TB_COMISIONES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
Set TB_COMISIONES = Nothing
End Sub


Sub pro_llena_listview1()
   numero_items_comisiones = 0
   Dim list_item As ListItem
   Dim var_n As Double
   rs.Open "select a.vcha_age_agente_id,a.vcha_lin_linea_id,b.vcha_lin_nombre,a.floa_com_limite_superior,a.floa_com_limite_inferior,a.floa_com_porcentaje from TB_comisiones a,tb_lineas b where a.vcha_age_agente_id = '" + var_agente_seleccionado + "' and a.vcha_lin_linea_id =  b.vcha_lin_linea_id order by a.vcha_lin_linea_id, a.floa_com_limite_inferior,a.floa_com_limite_superior,a.floa_com_porcentaje", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
      Set list_item = lv_comisiones.ListItems.Add(, , rs!vcha_lin_linea_id)
      list_item.SubItems(1) = IIf(IsNull(rs!VCHA_lin_NOMBRE), "", rs!VCHA_lin_NOMBRE)
      list_item.SubItems(2) = IIf(IsNull(rs!FLOA_COM_PORCENTAJE), "", rs!FLOA_COM_PORCENTAJE)
      list_item.SubItems(3) = IIf(IsNull(rs!FLOA_COM_LIMITE_INFERIOR), "", rs!FLOA_COM_LIMITE_INFERIOR)
      list_item.SubItems(4) = IIf(IsNull(rs!FLOA_COM_LIMITE_superior), "", rs!FLOA_COM_LIMITE_superior)
      rs.MoveNext:
     numero_items_comisiones = numero_items_comisiones + 1
   Wend
   rs.Close
   var_n = lv_comisiones.ListItems.Count
   If var_n = 0 Then
      cmd_guardar.Enabled = False
      cmd_deshacer.Enabled = False
      cmd_eliminar.Enabled = False
   Else
      cmd_guardar.Enabled = True
      cmd_deshacer.Enabled = True
      cmd_eliminar.Enabled = True
   End If
End Sub


Sub pro_textos()
On Error GoTo err0:
   Dim var_n As Double
   var_n = Me.lv_comisiones.ListItems.Count
   If var_n > 0 Then
      txt_linea = lv_comisiones.selectedItem
      txt_limite_inferior = lv_comisiones.selectedItem.SubItems(3)
      txt_limite_superior = lv_comisiones.selectedItem.SubItems(4)
      txt_porcentaje = lv_comisiones.selectedItem.SubItems(2)
      rs.Open "select * from tb_lineas where vcha_lin_linea_id = '" + txt_linea + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_linea = IIf(IsNull(rs!VCHA_lin_NOMBRE), "", rs!VCHA_lin_NOMBRE)
      Else
         txt_nombre_linea = ""
      End If
      rs.Close
      txt_linea.Enabled = False
      txt_nombre_linea.Enabled = False
   End If
   var_numero_renglones = lv_comisiones.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_comisiones.ColumnHeaders(2).Width = 4200
   Else
      lv_comisiones.ColumnHeaders(2).Width = 4500
   End If
   var_modifica_registro_comision = True
   var_hubo_cambios = False
err0:
End Sub

Private Sub pro_actualiza_ListView()
   lv_comisiones.ListItems.Clear
   pro_llena_listview1
End Sub




Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_tipo_lista = 1 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_linea = lv_lista.selectedItem
            txt_nombre_linea = lv_lista.selectedItem.SubItems(1)
         Else
           txt_linea = ""
           txt_nombre_linea = ""
         End If
         txt_linea.SetFocus
      End If
      If var_tipo_lista = 2 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_agente = lv_lista.selectedItem
            txt_nombre_agente = lv_lista.selectedItem.SubItems(1)
         Else
            txt_agente = ""
            txt_nombre_agente = ""
         End If
         txt_agente.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub txt_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_agentes order by vcha_age_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
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
End Sub

Private Sub txt_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_agente_LostFocus()
   If Trim(txt_agente) <> "" Then
      rs.Open "SELECT * FROM TB_AGENTES WHERE VCHA_AGE_AGENTE_ID = '" + txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_agente = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
      Else
         txt_agente = ""
         txt_nombre_agente = ""
         MsgBox "Clave de agente incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      txt_nombre_agente = ""
   End If
End Sub

Private Sub txt_limite_inferior_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_limite_inferior_GotFocus()
   Me.frm_agente_clonar.Visible = False
End Sub

Private Sub txt_limite_inferior_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_limite_inferior_LostFocus()
   If Not IsNumeric(txt_limite_inferior) Then
      MsgBox "Limite inferior incorrecto", vbOKOnly, "ATENCION"
      txt_limite_inferior = 0
   End If
End Sub

Private Sub txt_limite_superior_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_limite_superior_GotFocus()
   Me.frm_agente_clonar.Visible = False
End Sub

Private Sub txt_limite_superior_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_limite_superior_LostFocus()
   If Not IsNumeric(txt_limite_superior) Then
      MsgBox "Limite superior incorrecto", vbOKOnly, "ATENCION"
      txt_limite_superior = 0
   End If
End Sub

Private Sub txt_linea_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_linea_GotFocus()
   Me.frm_agente_clonar.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_linea_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_LINEAS order by vcha_lin_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_lin_linea_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_lin_NOMBRE), "", rs!VCHA_lin_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "LINEAS"
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
      var_llamado_comisiones = True
      var_catalogo_articulos = False
      frmlineas.Show
   End If
End Sub

Private Sub txt_linea_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_linea_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_linea) <> "" Then
      rs.Open "select * from tb_lineas where vcha_lin_linea_id = '" + txt_linea + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_linea = IIf(IsNull(rs1vcha_lin_nombre), "", rs!VCHA_lin_NOMBRE)
      Else
         MsgBox "Clave de Linea incorrecta", vbOKOnly, "ATENCION"
         txt_nombre_linea = ""
         txt_linea = ""
      End If
      rs.Close
   Else
      txt_nombre_linea = ""
   End If
End Sub

Private Sub txt_nombre_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_agentes order by vcha_age_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
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
End Sub

Private Sub txt_nombre_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      cmd_aceptar.SetFocus
   End If

End Sub

Private Sub txt_nombre_linea_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_linea_GotFocus()
   Me.frm_agente_clonar.Visible = False
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_linea_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_llamado_comisiones = True
      var_catalogo_articulos = False
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_LINEAS order by vcha_lin_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_lin_linea_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_lin_NOMBRE), "", rs!VCHA_lin_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "LINEAS"
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
      var_llamado_comisiones = True
      var_catalogo_articulos = False
      frmlineas.Show
   End If
End Sub

Private Sub txt_nombre_linea_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 13
   Case Else
       KeyAscii = 0
   End Select
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_linea_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_porcentaje_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_porcentaje_GotFocus()
   Me.frm_agente_clonar.Visible = False
End Sub

Private Sub txt_porcentaje_KeyPress(KeyAscii As Integer)
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

Private Sub txt_porcentaje_LostFocus()
   If Not IsNumeric(txt_porcentaje) Then
      MsgBox "Porcentaje incorrecto", vbOKOnly, "ATENCION"
      txt_porcentaje = 0
   End If
End Sub
