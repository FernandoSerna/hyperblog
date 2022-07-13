VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmov_almacenes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   Icon            =   "frmmov_almacenes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5955
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   135
      TabIndex        =   13
      Top             =   435
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1875
         Left            =   30
         TabIndex        =   14
         Top             =   435
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   3307
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
         TabIndex        =   15
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      Picture         =   "frmmov_almacenes.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   450
      Picture         =   "frmmov_almacenes.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Guardar Alt + G"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   780
      Picture         =   "frmmov_almacenes.frx":0ACE
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1110
      Picture         =   "frmmov_almacenes.frx":0BA0
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      Picture         =   "frmmov_almacenes.frx":0CA2
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
      Picture         =   "frmmov_almacenes.frx":0DA4
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
      Left            =   2670
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   12
      Top             =   90
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   " Almacen "
      Height          =   765
      Left            =   165
      TabIndex        =   11
      Top             =   405
      Width           =   5655
      Begin VB.TextBox txt_nombre_almacen 
         Height          =   315
         Left            =   1275
         TabIndex        =   8
         Top             =   300
         Width           =   4245
      End
      Begin VB.TextBox txt_almacen 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   1140
      End
   End
   Begin VB.Frame Frame3 
      Height          =   6045
      Left            =   150
      TabIndex        =   9
      Top             =   1155
      Width           =   5655
      Begin MSComctlLib.ImageList icono_encabezado 
         Left            =   15
         Top             =   825
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
               Picture         =   "frmmov_almacenes.frx":13DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmmov_almacenes.frx":1CB8
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lv_almacenes 
         Height          =   5850
         Left            =   30
         TabIndex        =   10
         Top             =   135
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   10319
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
            Text            =   "Almacen"
            Object.Width           =   9701
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Clave almacen"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   600
         Top             =   705
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
               Picture         =   "frmmov_almacenes.frx":2592
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmmov_almacenes.frx":2E6C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmmov_almacenes.frx":3746
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmmov_almacenes.frx":3CE2
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmmov_almacenes.frx":45BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmmov_almacenes.frx":4E96
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmmov_almacenes.frx":5770
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmmov_almacenes.frx":5A8A
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmmov_almacenes.frx":5DA4
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmmov_almacenes.frx":6340
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList3 
         Left            =   60
         Top             =   735
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
               Picture         =   "frmmov_almacenes.frx":665A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmmov_almacenes.frx":6F34
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   0
      Top             =   285
      Width           =   5655
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3120
      Top             =   -60
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
            Picture         =   "frmmov_almacenes.frx":780E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmov_almacenes.frx":80E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmov_almacenes.frx":89C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmov_almacenes.frx":8F5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmov_almacenes.frx":983A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmov_almacenes.frx":A114
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmov_almacenes.frx":A9EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmov_almacenes.frx":AB00
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmov_almacenes.frx":AC12
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmov_almacenes.frx":AD24
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmov_almacenes.frx":AE36
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmmov_almacenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmb_almacenes_Click()
    txt_almacen = Obtener_llave(cnn, rs, "TB_ALMACENES", "VCHA_ALM_NOMBRE", cmb_almacenes, 2, "T")
End Sub

Private Sub cmd_deshacer_Click()
   If lv_almacenes.ListItems.Count > 0 Then
      lv_almacenes.SetFocus
      txt_almacen = lv_almacenes.selectedItem.SubItems(1)
      txt_nombre_almacen = lv_almacenes.selectedItem
   Else
      txt_almacen = ""
      txt_nombre_almacen = ""
   End If
   If lv_almacenes.ListItems.Count <= 0 Then
      cmd_guardar.Enabled = False
      cmd_deshacer.Enabled = False
      cmd_eliminar.Enabled = False
   Else
      cmd_guardar.Enabled = True
      cmd_deshacer.Enabled = True
      cmd_eliminar.Enabled = True
   End If
End Sub

Private Sub cmd_eliminar_Click()
   Dim list_item As ListItem
   Set TB_MOVIMIENTOS_ALMACENES = New TB_MOVIMIENTOS_ALMACENES
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
      si = MsgBox("¿Deseas Eliminar el Registro?", vbYesNo, "ATENCION")
      If si = 6 Then
         ok = TB_MOVIMIENTOS_ALMACENES.Eliminar(var_movimiento_almacen, txt_almacen)
         If ok Then
            lv_almacenes.ListItems.Remove (lv_almacenes.selectedItem.Index)
            Call pro_limpiatextos(Me)
            MsgBox "El registro se elimino correctamente", vbOKOnly, "ATENCION"
            If lv_almacenes.ListItems.Count <= 0 Then
               cmd_guardar.Enabled = False
               cmd_deshacer.Enabled = False
               cmd_eliminar.Enabled = False
            Else
               cmd_guardar.Enabled = True
               cmd_deshacer.Enabled = True
               cmd_eliminar.Enabled = True
            End If
         End If
      End If
   End If
End Sub

Private Sub cmd_guardar_Click()
   Dim list_item As ListItem
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
      Set TB_MOVIMIENTOS_ALMACENES = New TB_MOVIMIENTOS_ALMACENES
      ok = TB_MOVIMIENTOS_ALMACENES.Anadir(var_movimiento_almacen, txt_almacen)
      If ok Then
         MsgBox "Información Guardada Corectamente", vbOKOnly, "ATENCION"
         Set list_item = lv_almacenes.ListItems.Add(, , txt_nombre_almacen)
         list_item.SubItems(1) = txt_almacen
      End If
   End If
   If lv_almacenes.ListItems.Count <= 0 Then
      cmd_guardar.Enabled = False
      cmd_deshacer.Enabled = False
      cmd_eliminar.Enabled = False
   Else
      cmd_guardar.Enabled = True
      cmd_deshacer.Enabled = True
      cmd_eliminar.Enabled = True
   End If
End Sub

Private Sub cmd_imprimir_Click()
   x = 1 + 1
End Sub

Private Sub cmd_nuevo_Click()
   cmd_guardar.Enabled = True
   cmd_deshacer.Enabled = False
   txt_almacen = ""
   txt_nombre_almacen = ""
   txt_almacen.SetFocus
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
Dim list_item As ListItem
   frm_lista.Visible = False
   var_cadena_seguridad = ""
   rs.Open "select * from vw_movimientos_almacenes where vcha_mov_movimiento_id = '" + var_movimiento_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
   If rs.BOF Then
      cmd_guardar.Enabled = False
      cmd_deshacer.Enabled = False
      cmd_eliminar.Enabled = False
   Else
      cmd_guardar.Enabled = True
      cmd_deshacer.Enabled = True
      cmd_eliminar.Enabled = True
   End If
   While Not rs.EOF
      Set list_item = lv_almacenes.ListItems.Add(, , rs!vcha_alm_nombre)
      list_item.SubItems(1) = IIf(IsNull(rs!vcha_alm_almacen_id), "", rs!vcha_alm_almacen_id)
      rs.MoveNext:
      txt_almacen = lv_almacenes.selectedItem.SubItems(1)
      txt_nombre_almacen = lv_almacenes.selectedItem
   Wend
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_mov_almacenes)
End Sub

Private Sub lv_almacenes_ItemClick(ByVal Item As MSComctlLib.ListItem)
   If lv_almacenes.ListItems.Count > 0 Then
      txt_almacen = lv_almacenes.selectedItem.SubItems(1)
      txt_nombre_almacen = lv_almacenes.selectedItem
   End If
End Sub


Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_almacen = lv_lista.selectedItem
      txt_nombre_almacen = lv_lista.selectedItem.SubItems(1)
      txt_almacen.SetFocus
   End If
   If KeyAscii = 27 Then
      txt_almacen.SetFocus
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub txt_almacen_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_almacen_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_almacenes order by vcha_alm_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_alm_almacen_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_alm_nombre), "", rs!vcha_alm_nombre)
            rs.MoveNext
      Wend
      rs.Close
      var_tipo_lista = 4
      lbl_lista = "Almacenes"
      var_n = lv_lista.ListItems.Count
      If var_n > 7 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
   If KeyCode = 117 Then
      Me.Enabled = False
      var_activa_forma_almacenes = Me.Name
      frmalmacenes.Show
   End If
End Sub

Private Sub txt_almacen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_almacen_LostFocus()
   If Trim(txt_almacen) <> "" Then
      rs.Open "select * from tb_almacenes where vcha_alm_almacen_id = '" + txt_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_almacen = IIf(IsNull(rs!vcha_alm_nombre), "", rs!vcha_alm_nombre)
      Else
         MsgBox "Clave de almacén incorrecto", vbOKOnly, "ATENCION"
         txt_nombre_almacen = ""
         txt_almacen = ""
      End If
      rs.Close
   Else
      txt_nombre_almacen = ""
   End If
End Sub

Private Sub txt_nombre_almacen_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_almacen_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_almacenes order by vcha_alm_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_alm_almacen_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_alm_nombre), "", rs!vcha_alm_nombre)
            rs.MoveNext
      Wend
      rs.Close
      var_tipo_lista = 4
      lbl_lista = "Almacenes"
      var_n = lv_lista.ListItems.Count
      If var_n > 7 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
   If KeyCode = 117 Then
      Me.Enabled = False
      var_activa_forma_almacenes = Me.Name
      frmalmacenes.Show
   End If
End Sub

Private Sub txt_nombre_almacen_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      If Me.cmd_guardar.Enabled = True Then
         Me.cmd_guardar.SetFocus
      End If
   End If
End Sub
