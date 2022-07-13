VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmdirecciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dirección"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5955
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   435
      Picture         =   "frmdirecciones.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Cancelar Alt + C"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   90
      TabIndex        =   21
      Top             =   375
      Width           =   5790
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   22
         Top             =   480
         Width           =   5700
         _ExtentX        =   10054
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
         TabIndex        =   23
         Top             =   120
         Width           =   5715
      End
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   45
      TabIndex        =   20
      Top             =   345
      Width           =   5850
   End
   Begin VB.CommandButton cmd_aceptar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frmdirecciones.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame5 
      Caption         =   " Dirección "
      Height          =   2250
      Left            =   90
      TabIndex        =   0
      Top             =   420
      Width           =   5775
      Begin VB.TextBox txt_codigo_postal 
         Height          =   315
         Left            =   885
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1905
         Width           =   1005
      End
      Begin VB.TextBox txt_pais 
         Height          =   315
         Left            =   885
         MaxLength       =   50
         TabIndex        =   1
         Top             =   180
         Width           =   1005
      End
      Begin VB.TextBox txt_estado 
         Height          =   315
         Left            =   885
         MaxLength       =   50
         TabIndex        =   3
         Top             =   525
         Width           =   1005
      End
      Begin VB.TextBox txt_ciudad 
         Height          =   315
         Left            =   885
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1215
         Width           =   1005
      End
      Begin VB.TextBox txt_colonia 
         Height          =   315
         Left            =   885
         MaxLength       =   50
         TabIndex        =   9
         Top             =   1560
         Width           =   1005
      End
      Begin VB.TextBox txt_nombre_pais 
         Height          =   315
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   2
         Top             =   180
         Width           =   3780
      End
      Begin VB.TextBox txt_nombre_estado 
         Height          =   315
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         Top             =   525
         Width           =   3780
      End
      Begin VB.TextBox txt_nombre_ciudad 
         Height          =   315
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1215
         Width           =   3780
      End
      Begin VB.TextBox txt_nombre_colonia 
         Height          =   315
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1560
         Width           =   3780
      End
      Begin VB.TextBox txt_municipio 
         Height          =   315
         Left            =   885
         MaxLength       =   50
         TabIndex        =   5
         Top             =   870
         Width           =   1005
      End
      Begin VB.TextBox txt_nombre_municipio 
         Height          =   315
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   6
         Top             =   870
         Width           =   3780
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "CP:"
         Height          =   195
         Index           =   22
         Left            =   165
         TabIndex        =   19
         Top             =   1965
         Width           =   255
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Pais:"
         Height          =   195
         Index           =   17
         Left            =   150
         TabIndex        =   18
         Top             =   240
         Width           =   345
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Index           =   18
         Left            =   150
         TabIndex        =   17
         Top             =   585
         Width           =   540
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad:"
         Height          =   195
         Index           =   19
         Left            =   150
         TabIndex        =   16
         Top             =   1275
         Width           =   540
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Colonia:"
         Height          =   195
         Index           =   20
         Left            =   150
         TabIndex        =   15
         Top             =   1620
         Width           =   570
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Municipio:"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   14
         Top             =   930
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmdirecciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tipo_lista As Integer
Private Sub cmd_cancelar_Click()
   var_aceptar_direccion = False
   Unload Me
End Sub

Private Sub cmd_aceptar_Click()
   var_aceptar_direccion = True
   var_dir_pais = txt_pais
   var_dir_nombre_pais = txt_nombre_pais
   var_dir_estado = txt_estado
   var_dir_nombre_estado = txt_nombre_estado
   var_dir_municipio = txt_municipio
   var_dir_nombre_municipio = txt_nombre_municipio
   var_dir_ciudad = txt_ciudad
   var_dir_nombre_ciudad = txt_nombre_ciudad
   var_dir_colonia = txt_colonia
   var_dir_nombre_colonia = txt_nombre_colonia
   var_dir_codigo_postal = txt_codigo_postal
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 2000
   Left = 2900
   var_cadena_seguridad = ""
   frm_lista.Visible = False
   var_aceptar_direccion = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_aceptar_direccion = True Then
      Select Case var_activa_forma_direcciones
          Case "frmtitulares"
            frmtitulares.txt_pais = var_dir_pais
            frmtitulares.txt_nombre_pais = var_dir_nombre_pais
            frmtitulares.txt_estado = var_dir_estado
            frmtitulares.txt_nombre_estado = var_dir_nombre_estado
            frmtitulares.txt_municipio = var_dir_municipio
            frmtitulares.txt_nombre_municipio = var_dir_nombre_municipio
            frmtitulares.txt_ciudad = var_dir_ciudad
            frmtitulares.txt_nombre_ciudad = var_dir_nombre_ciudad
            frmtitulares.txt_colonia = var_dir_colonia
            frmtitulares.txt_nombre_colonia = var_dir_nombre_colonia
            frmtitulares.txt_codigo_postal = var_dir_codigo_postal
          Case "frmclientes"
            frmclientes.txt_pais = var_dir_pais
            frmclientes.txt_nombre_pais = var_dir_nombre_pais
            frmclientes.txt_estado = var_dir_estado
            frmclientes.txt_nombre_estado = var_dir_nombre_estado
            frmclientes.txt_municipio = var_dir_municipio
            frmclientes.txt_nombre_municipio = var_dir_nombre_municipio
            frmclientes.txt_ciudad = var_dir_ciudad
            frmclientes.txt_nombre_ciudad = var_dir_nombre_ciudad
            frmclientes.txt_colonia = var_dir_colonia
            frmclientes.txt_nombre_colonia = var_dir_nombre_colonia
            frmclientes.txt_codigo_postal = var_dir_codigo_postal
          Case "frmestablecimientos"
            frmestablecimientos.txt_pais = var_dir_pais
            frmestablecimientos.txt_nombre_pais = var_dir_nombre_pais
            frmestablecimientos.txt_estado = var_dir_estado
            frmestablecimientos.txt_nombre_estado = var_dir_nombre_estado
            frmestablecimientos.txt_municipio = var_dir_municipio
            frmestablecimientos.txt_nombre_municipio = var_dir_nombre_municipio
            frmestablecimientos.txt_ciudad = var_dir_ciudad
            frmestablecimientos.txt_nombre_ciudad = var_dir_nombre_ciudad
            frmestablecimientos.txt_colonia = var_dir_colonia
            frmestablecimientos.txt_nombre_colonia = var_dir_nombre_colonia
            frmestablecimientos.txt_codigo_postal = var_dir_codigo_postal
      End Select
   End If
   Call activa_forma(var_activa_forma_direcciones)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_tipo_lista = 10 Then
         If lv_lista.ListItems.Count = 0 Then
            txt_ciudad = ""
            txt_nombre_ciudad = ""
         Else
            txt_ciudad = lv_lista.selectedItem
            txt_nombre_ciudad = lv_lista.selectedItem.SubItems(1)
         End If
         txt_ciudad.SetFocus
      End If
      If var_tipo_lista = 11 Then
         If lv_lista.ListItems.Count = 0 Then
            txt_municipio = ""
            txt_nombre_municipio = ""
         Else
            txt_municipio = lv_lista.selectedItem
            txt_nombre_municipio = lv_lista.selectedItem.SubItems(1)
         End If
         txt_municipio.SetFocus
      End If
      If var_tipo_lista = 12 Then
         If lv_lista.ListItems.Count = 0 Then
            txt_estado = ""
            txt_nombre_estado = ""
         Else
            txt_estado = lv_lista.selectedItem
            txt_nombre_estado = lv_lista.selectedItem.SubItems(1)
         End If
         txt_estado.SetFocus
      End If
      If var_tipo_lista = 13 Then
         If lv_lista.ListItems.Count = 0 Then
            txt_pais = ""
            txt_nombre_pais = ""
         Else
            txt_pais = lv_lista.selectedItem
            txt_nombre_pais = lv_lista.selectedItem.SubItems(1)
         End If
         txt_pais.SetFocus
      End If
      If var_tipo_lista = 14 Then
         If lv_lista.ListItems.Count = 0 Then
            txt_colonia = ""
            txt_nombre_colonia = ""
         Else
            txt_colonia = lv_lista.selectedItem
            txt_nombre_colonia = lv_lista.selectedItem.SubItems(1)
         End If
         txt_colonia.SetFocus
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

Private Sub txt_ciudad_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_ciudad_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_ciudades where vcha_pai_pais_id = '" + txt_pais + "' and vcha_est_estado_id = '" + txt_estado + "' order by vcha_ciu_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_ciu_ciudad_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CIUDADES DE " + txt_nombre_estado
      var_tipo_lista = 10
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
      frmciudades.Show
   End If
End Sub

Private Sub txt_ciudad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_ciudad_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_ciudad) <> "" Then
      rs.Open "select * from TB_CIUDADES where VCHA_CIU_CIUDAD_ID = '" + txt_ciudad + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_ciudad = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
      Else
         MsgBox "Clave de ciudad incorrecta", vbOKOnly, "ATENCION"
         txt_ciudad = ""
         txt_nombre_ciudad = ""
      End If
      rs.Close
   Else
      txt_nombre_ciudad = ""
   End If
End Sub

Private Sub txt_codigo_postal_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Me.cmd_aceptar.SetFocus
End Sub

Private Sub txt_colonia_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_colonia_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_colonias where vcha_pai_pais_id = '" + txt_pais + "' and vcha_est_estado_id = '" + txt_estado + "' and vcha_mun_municipio_id = '" + txt_municipio + "' order by vcha_col_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_COL_COLONIA_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "COLONIAS DE " + txt_nombre_estado
      var_tipo_lista = 14
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
      frmciudades.Show
   End If
End Sub

Private Sub txt_colonia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_colonia_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_colonia) <> "" Then
      rs.Open "select * from TB_COLONIAS where VCHA_COL_COLONIA_ID = '" + txt_colonia + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_colonia = IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
         txt_codigo_postal = IIf(IsNull(rs!vcha_col_cp), "", rs!vcha_col_cp)
      Else
         MsgBox "Clave de colonias incorrecta", vbOKOnly, "ATENCION"
         txt_colonia = ""
         txt_nombre_colonia = ""
      End If
      rs.Close
   Else
      txt_nombre_colonia = ""
   End If
End Sub

Private Sub txt_estado_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_estado_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_estados where vcha_pai_pais_id = '" + txt_pais + "' order by vcha_est_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_est_estado_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ESTADOS DE " + txt_nombre_pais
      var_tipo_lista = 12
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
      frmestados.Show
   End If
End Sub

Private Sub txt_estado_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_estado_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_estado) <> "" Then
      rs.Open "select * from TB_ESTADOS where VCHA_EST_ESTADO_ID = '" + txt_estado + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_estado = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
      Else
         MsgBox "Clave de estado incorrecta", vbOKOnly, "ATENCION"
         txt_estado = ""
         txt_nombre_estado = ""
      End If
      rs.Close
   Else
      txt_nombre_estado = ""
   End If
End Sub

Private Sub txt_municipio_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_municipio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_municipios where vcha_pai_pais_id = '" + txt_pais + "' and vcha_est_estado_id = '" + txt_estado + "' order by vcha_mun_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_mun_municipio_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_MUN_NOMBRE), "", rs!VCHA_MUN_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "MUNICIPIOS DE" + txt_nombre_estado
      var_tipo_lista = 11
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
      frmmunicipios.Show
   End If
End Sub

Private Sub txt_municipio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_municipio_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_ciudad_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_ciudad_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_ciudades where vcha_pai_pais_id = '" + txt_pais + "' and vcha_est_estado_id = '" + txt_estado + "' order by vcha_ciu_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_ciu_ciudad_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_ciu_nombre), "", rs!vcha_ciu_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CIUDADES DE " + txt_nombre_estado
      var_tipo_lista = 10
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
      frmciudades.Show
   End If
End Sub

Private Sub txt_nombre_ciudad_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_ciudad_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""

End Sub

Private Sub txt_nombre_colonia_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_colonia_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_colonias where vcha_pai_pais_id = '" + txt_pais + "' and vcha_est_estado_id = '" + txt_estado + "' order by vcha_col_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_COL_COLONIA_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_col_nombre), "", rs!vcha_col_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "COLONIAS DE " + txt_nombre_estado
      var_tipo_lista = 14
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
      frmciudades.Show
   End If
End Sub

Private Sub txt_nombre_colonia_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_colonia_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_estado_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_estado_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_estados where vcha_pai_pais_id = '" + txt_pais + "' order by vcha_est_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_est_estado_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ESTADOS DE " + txt_nombre_pais
      var_tipo_lista = 12
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
      frmestados.Show
   End If
End Sub

Private Sub txt_nombre_estado_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_estado_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_municipio_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_municipio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_municipios where vcha_pai_pais_id = '" + txt_pais + "' and vcha_est_estado_id = '" + txt_estado + "' and vcha_mun_municipio_id = '" + txt_municipio + "' order by vcha_mun_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_mun_municipio_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_MUN_NOMBRE), "", rs!VCHA_MUN_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "MUNICIPIOS DE" + txt_nombre_estado
      var_tipo_lista = 11
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
      frmmunicipios.Show
   End If
End Sub

Private Sub txt_nombre_municipio_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_municipio_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_pais_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_nombre_pais_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_paises order by vcha_pai_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_pai_pais_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_PAI_NOMBRE), "", rs!VCHA_PAI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "PAISES"
      var_tipo_lista = 13
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
      var_catalogo_articulos = True
      frmpaises.Show
   End If
End Sub

Private Sub txt_nombre_pais_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_nombre_pais_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_pais_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para entrar al catálogo"
End Sub

Private Sub txt_pais_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_paises order by vcha_pai_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_pai_pais_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_PAI_NOMBRE), "", rs!VCHA_PAI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "PAISES"
      var_tipo_lista = 13
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
      var_catalogo_articulos = True
      frmpaises.Show
   End If
End Sub

Private Sub txt_pais_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_pais_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_pais) <> "" Then
      rs.Open "select * from TB_PAISES where VCHA_PAI_PAIS_ID = '" + txt_pais + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_pais = IIf(IsNull(rs!VCHA_PAI_NOMBRE), "", rs!VCHA_PAI_NOMBRE)
      Else
         MsgBox "Clave de pais incorrecta", vbOKOnly, "ATENCION"
         txt_pais = ""
         txt_nombre_pais = ""
      End If
      rs.Close
   Else
      txt_nombre_pais = ""
   End If
End Sub
