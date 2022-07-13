VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmubicaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ubicaciones"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_disponibles 
      Height          =   3210
      Left            =   300
      TabIndex        =   5
      Top             =   45
      Width           =   7110
      Begin VB.TextBox txt_nombre_articulo 
         Height          =   345
         Left            =   90
         TabIndex        =   6
         Top             =   420
         Width           =   6915
      End
      Begin MSComctlLib.ListView lv_disponibles 
         Height          =   2325
         Left            =   75
         TabIndex        =   7
         Top             =   795
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   4101
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
            Text            =   "Código"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre del Artículo"
            Object.Width           =   9701
         EndProperty
      End
      Begin VB.Label Label18 
         BackColor       =   &H8000000D&
         Caption         =   " Artículos Disponibles"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   8
         Top             =   120
         Width           =   7035
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Ubicaciones "
      Height          =   2310
      Left            =   135
      TabIndex        =   3
      Top             =   945
      Width           =   7455
      Begin MSComctlLib.ListView lv_ubicaciones 
         Height          =   2040
         Left            =   45
         TabIndex        =   4
         Top             =   180
         Width           =   7350
         _ExtentX        =   12965
         _ExtentY        =   3598
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Ubicación"
            Object.Width           =   8114
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Cantidad"
            Object.Width           =   2893
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Activa"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Código "
      Height          =   765
      Left            =   135
      TabIndex        =   0
      Top             =   165
      Width           =   7455
      Begin VB.TextBox txt_descripcion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1545
         TabIndex        =   2
         Top             =   255
         Width           =   5835
      End
      Begin VB.TextBox txt_codigo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   255
         Width           =   1410
      End
   End
End
Attribute VB_Name = "frmubicaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Top = 1500
   Left = 2000
   Me.frm_disponibles.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub lv_disponibles_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.lv_disponibles.ListItems.Count > 0 Then
         Me.txt_codigo = Me.lv_disponibles.selectedItem
         Me.txt_descripcion = Me.lv_disponibles.selectedItem.SubItems(1)
         Me.txt_codigo.SetFocus
         Me.frm_disponibles.Visible = False
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_disponibles.Visible = False
   End If
End Sub

Private Sub lv_ubicaciones_GotFocus()
   Me.frm_disponibles.Visible = False
End Sub

Private Sub lv_ubicaciones_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 115 Then
      If Me.lv_ubicaciones.ListItems.Count > 0 Then
         If Me.lv_ubicaciones.selectedItem.SubItems(2) = "" Then
            var_si = MsgBox("¿Desea desactivar la ubicación?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               rs.Open "UPDATE TB_EXISTENCIAS_UBICACIONES SET INTE_UBI_ACTIVA = 1 WHERE VCHA_UBI_UBICACION = '" + Trim(Me.lv_ubicaciones.selectedItem) + "' AND VCHA_aRT_ARTICULO_ID = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               Me.lv_ubicaciones.selectedItem.SubItems(2) = "1"
            End If
         Else
            var_si = MsgBox("¿Desea activar la ubicación?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               rs.Open "UPDATE TB_EXISTENCIAS_UBICACIONES SET INTE_UBI_ACTIVA = 0 WHERE VCHA_UBI_UBICACION = '" + Trim(Me.lv_ubicaciones.selectedItem) + "' AND VCHA_aRT_ARTICULO_ID = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               Me.lv_ubicaciones.selectedItem.SubItems(2) = ""
            End If
         End If
      End If
   End If
   If KeyCode = 116 Then
      If Me.lv_ubicaciones.ListItems.Count > 0 Then
         If CDbl(Me.lv_ubicaciones.selectedItem.SubItems(1)) > 0 Then
            frmubicaciones_reorganziar.txt_ubicacion_origen = Me.lv_ubicaciones.selectedItem
            frmubicaciones_reorganziar.txt_codigo = Me.txt_codigo
            frmubicaciones_reorganziar.txt_descripcion = Me.txt_descripcion
            frmubicaciones_reorganziar.txt_cantidad_origen = Me.lv_ubicaciones.selectedItem.SubItems(1)
            frmubicaciones_reorganziar.Show
            frmubicaciones_reorganziar.txt_ubicacion_destino.SetFocus
         Else
            MsgBox "La cantidad debe de ser mayor de 0", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub

Private Sub lv_ubicaciones_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.lv_ubicaciones.ListItems.Count > 0 Then
         Me.txt_codigo.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub txt_codigo_Change()
   Me.txt_descripcion = ""
   Me.lv_ubicaciones.ListItems.Clear
End Sub

Private Sub txt_codigo_GotFocus()
   Me.frm_disponibles.Visible = False
End Sub

Private Sub txt_codigo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Me.txt_nombre_articulo = ""
      Me.lv_disponibles.ListItems.Clear
      Me.frm_disponibles.Visible = True
      Me.txt_nombre_articulo.SetFocus
   End If
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_descripcion.SetFocus
   End If
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub txt_codigo_LostFocus()
   If Trim(Me.txt_codigo) <> "" Then
      rsaux.Open "select * from tb_articulos where vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux.EOF Then
         Me.txt_descripcion = IIf(IsNull(rsaux!vcha_art_nombre_Español), "", rsaux!vcha_art_nombre_Español)
         Me.lv_ubicaciones.ListItems.Clear
      Else
         rsaux1.Open "select * from tb_Equivalencias where vcha_equ_codigo_equivalente = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux1.EOF Then
            Me.txt_codigo = IIf(IsNull(rsaux1!vcha_Art_articulo_id), "", rsaux1!vcha_Art_articulo_id)
            rsaux2.Open "select * from tb_Articulos where vcha_Art_Articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux2.EOF Then
               Me.txt_descripcion = IIf(IsNull(rsaux2!vcha_art_nombre_Español), "", rsaux2!vcha_art_nombre_Español)
               Me.lv_ubicaciones.ListItems.Clear
            Else
               MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
               Me.txt_codigo = ""
               Me.txt_descripcion = ""
               Me.lv_ubicaciones.ListItems.Clear
            End If
            rsaux2.Close
         Else
            MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
            Me.txt_codigo = ""
            Me.txt_descripcion = ""
            Me.lv_ubicaciones.ListItems.Clear
         End If
         rsaux1.Close
      End If
      rsaux.Close
      If Me.txt_codigo <> "" Then
         rs.Open "select * from tb_existencias_ubicaciones where vcha_art_Articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            While Not rs.EOF
                  Set list_item = lv_ubicaciones.ListItems.Add(, , rs!VCHA_UBI_UBICACION)
                  list_item.SubItems(1) = IIf(IsNull(rs!FLOA_UBI_CANTIDAD), "", rs!FLOA_UBI_CANTIDAD)
                  list_item.SubItems(2) = IIf(IsNull(rs!inte_ubi_activa), "", rs!inte_ubi_activa)
                  rs.MoveNext
            Wend
         End If
         rs.Close
      End If
   Else
      Me.lv_ubicaciones.ListItems.Clear
      Me.txt_descripcion = ""
   End If
End Sub

Private Sub txt_descripcion_GotFocus()
   Me.frm_disponibles.Visible = False
End Sub

Private Sub txt_descripcion_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Me.txt_nombre_articulo = ""
      Me.lv_disponibles.ListItems.Clear
      Me.frm_disponibles.Visible = True
      Me.txt_nombre_articulo.SetFocus
   End If
End Sub

Private Sub txt_descripcion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.lv_ubicaciones.SetFocus
   Else
      If KeyAscii = 27 Then
         Unload Me
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txt_nombre_articulo_Change()
   Me.lv_disponibles.ListItems.Clear
End Sub

Private Sub txt_nombre_articulo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 27 Then
      Me.txt_codigo.SetFocus
      Me.frm_disponibles.Visible = False
   End If
   If KeyAscii = 39 Or KeyAscii = 61 Or KeyAscii = 44 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Trim(Me.txt_nombre_articulo) <> "" Then
         var_cadena = ""
         var_like_1 = ""
         var_like_2 = ""
         var_like_3 = ""
         var_like_4 = ""
         var_like_5 = ""
         var_like_6 = ""
         var_like_7 = ""
         var_j = 1
         For var_i = 1 To Len(Me.txt_nombre_articulo)
             If Mid(Me.txt_nombre_articulo, var_i, 1) <> " " Then
                If var_j = 1 Then
                   var_like_1 = var_like_1 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
                If var_j = 2 Then
                   var_like_2 = var_like_2 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
                If var_j = 3 Then
                   var_like_3 = var_like_3 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
                If var_j = 4 Then
                   var_like_4 = var_like_4 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
                If var_j = 5 Then
                   var_like_5 = var_like_5 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
                If var_j = 6 Then
                   var_like_6 = var_like_6 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
                If var_j >= 7 Then
                   var_like_7 = var_like_7 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
             Else
                var_j = var_j + 1
             End If
         Next var_i
      End If
      If Trim(var_like_1) <> "" Then
         var_cadena = " vcha_art_nombre_Español like '%" + var_like_1 + "%'"
      End If
      If Trim(var_like_2) <> "" Then
         var_cadena = var_cadena + " and  vcha_art_nombre_Español like '%" + var_like_2 + "%'"
      End If
      If Trim(var_like_3) <> "" Then
         var_cadena = var_cadena + " and  vcha_art_nombre_Español like '%" + var_like_3 + "%'"
      End If
      If Trim(var_like_4) <> "" Then
         var_cadena = var_cadena + " and  vcha_art_nombre_Español like '%" + var_like_4 + "%'"
      End If
      If Trim(var_like_5) <> "" Then
         var_cadena = var_cadena + " and  vcha_art_nombre_Español like '%" + var_like_5 + "%'"
      End If
      If Trim(var_like_6) <> "" Then
         var_cadena = var_cadena + " and  vcha_art_nombre_Español like '%" + var_like_6 + "%'"
      End If
      If Trim(var_like_7) <> "" Then
         var_cadena = var_cadena + " and  vcha_art_nombre_Español like '%" + var_like_7 + "%'"
      End If
      Me.lv_disponibles.ListItems.Clear
      If Trim(var_cadena) <> "" Then
         var_cadena = "SELECT * FROM tb_articulos WHERE " + var_cadena
         rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
            Set list_item = lv_disponibles.ListItems.Add(, , rs!vcha_Art_articulo_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_art_nombre_Español), "", rs!vcha_art_nombre_Español)
            rs.MoveNext
         Wend
         rs.Close
         If Me.lv_disponibles.ListItems.Count > 0 Then
            Me.lv_disponibles.SetFocus
         End If
         If lv_disponibles.ListItems.Count > 11 Then
            lv_disponibles.ColumnHeaders(2).Width = 5300
         Else
            lv_disponibles.ColumnHeaders(2).Width = 5500
         End If
      End If
   End If
End Sub
