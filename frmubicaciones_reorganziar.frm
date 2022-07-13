VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmubicaciones_reorganziar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reorganizar ubicaciones"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   1335
      TabIndex        =   17
      Top             =   210
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   18
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
            Text            =   "CODIGO"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DESCRIPCION"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "CANTIDAD"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         Caption         =   " Artículos en ubicación"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   19
         Top             =   135
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7380
      Picture         =   "frmubicaciones_reorganziar.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frmubicaciones_reorganziar.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmubicaciones_reorganziar.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   30
      Left            =   0
      TabIndex        =   16
      Top             =   360
      Width           =   7860
   End
   Begin VB.Frame Frame2 
      Caption         =   " Ubicacion destino "
      Height          =   1140
      Left            =   90
      TabIndex        =   13
      Top             =   2130
      Width           =   7770
      Begin VB.TextBox txt_cantidad_destino 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   1035
         TabIndex        =   8
         Top             =   630
         Width           =   1920
      End
      Begin VB.TextBox txt_ubicacion_destino 
         Height          =   360
         Left            =   1035
         TabIndex        =   7
         Top             =   240
         Width           =   3915
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   165
         TabIndex        =   15
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ubicación:"
         Height          =   195
         Left            =   165
         TabIndex        =   14
         Top             =   330
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Ubicación original "
      Height          =   1560
      Left            =   90
      TabIndex        =   9
      Top             =   525
      Width           =   7755
      Begin VB.TextBox txt_cantidad_origen 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   1035
         TabIndex        =   6
         Top             =   1095
         Width           =   1920
      End
      Begin VB.TextBox txt_descripcion 
         Height          =   360
         Left            =   2970
         TabIndex        =   5
         Top             =   705
         Width           =   4665
      End
      Begin VB.TextBox txt_codigo 
         Height          =   360
         Left            =   1035
         TabIndex        =   4
         Top             =   705
         Width           =   1920
      End
      Begin VB.TextBox txt_ubicacion_origen 
         Height          =   360
         Left            =   1035
         TabIndex        =   3
         Top             =   315
         Width           =   1920
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   165
         TabIndex        =   12
         Top             =   1185
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   165
         TabIndex        =   11
         Top             =   795
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ubicación:"
         Height          =   195
         Left            =   165
         TabIndex        =   10
         Top             =   398
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmubicaciones_reorganziar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_guardar_Click()
   If Me.txt_ubicacion_origen <> "" Then
      If Me.txt_codigo <> "" Then
         If Me.txt_ubicacion_destino <> "" Then
            If IsNumeric(Me.txt_cantidad_destino) Then
               If CDbl(Me.txt_cantidad_origen) >= CDbl(Me.txt_cantidad_destino) Then
                  rs.Open "insert into tb_ubicaciones_Salidas (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_mov_movimiento_id, inte_sal_numero, vcha_art_articulo_id, vcha_ubi_ubicacion, floa_sal_Cantidad) VALUES ('" + var_empresa + "','" + var_unidad_organizacional + "','REUBICACION',0,'" + Me.txt_codigo + "','" + Me.txt_ubicacion_origen + "'," + CStr(CDbl(Me.txt_cantidad_destino)) + ")", cnn, adOpenDynamic, adLockOptimistic
                  rs.Open "insert into tb_ubicaciones_ENTRADAS (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_mov_movimiento_id, inte_ENT_numero, vcha_art_articulo_id, vcha_ubi_ubicacion, floa_ENT_Cantidad) VALUES ('" + var_empresa + "','" + var_unidad_organizacional + "','REUBICACION',0,'" + Me.txt_codigo + "','" + Me.txt_ubicacion_destino + "'," + CStr(CDbl(Me.txt_cantidad_destino)) + ")", cnn, adOpenDynamic, adLockOptimistic
                  MsgBox "Se a terminado la reubicacion de los códigos", vbOKOnly, "ATENCION"
                  Me.txt_cantidad_origen = Format(CDbl(Me.txt_cantidad_origen) - CDbl(Me.txt_cantidad_destino), "###,###,##0.00")
               Else
                  MsgBox "La cantidad destino no debe ser superior a la cantidad origen", vbOKOnly, "ATENCION"
                  Me.txt_cantidad_destino.SetFocus
               End If
            Else
               MsgBox "Cantidad destino incorrecta", vbOKOnly, "ATENCION"
               Me.txt_cantidad_destino.SetFocus
            End If
         Else
            MsgBox "No se a indicado la ubicación destino", vbOKOnly, "ATENCION"
            Me.txt_ubicacion_destino.SetFocus
         End If
      Else
         MsgBox "No se a indicado un artículo", vbOKOnly, "ATENCION"
         Me.txt_codigo.SetFocus
      End If
   Else
      MsgBox "No se a indicado la ubicación origen", vbOKOnly, "ATENCION"
      Me.txt_ubicacion_origen.SetFocus
   End If
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_ubicacion_origen.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 1500
   Left = 2000
   Me.frm_lista.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.lv_lista.ListItems.Count > 0 Then
         Me.txt_codigo = Me.lv_lista.selectedItem
         Me.txt_descripcion = Me.lv_lista.selectedItem.SubItems(1)
         Me.txt_cantidad_origen = Me.lv_lista.selectedItem.SubItems(2)
         Me.txt_ubicacion_destino.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub txt_cantidad_destino_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub

Private Sub txt_cantidad_origen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_ubicacion_destino.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_codigo_Change()
   Me.txt_descripcion = ""
   Me.txt_cantidad_origen = ""
   Me.txt_ubicacion_destino = ""
   Me.txt_cantidad_destino = ""
End Sub

Private Sub txt_codigo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If Me.txt_ubicacion_origen <> "" Then
         Me.lv_lista.ListItems.Clear
         rs.Open "SELECT dbo.TB_EXISTENCIAS_UBICACIONES.VCHA_UBI_UBICACION, dbo.TB_EXISTENCIAS_UBICACIONES.VCHA_ART_ARTICULO_ID, dbo.TB_EXISTENCIAS_UBICACIONES.FLOA_UBI_CANTIDAD , dbo.TB_Articulos.VCHA_ART_NOMBRE_ESPAÑOL FROM dbo.TB_EXISTENCIAS_UBICACIONES INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_EXISTENCIAS_UBICACIONES.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID WHERE (dbo.TB_EXISTENCIAS_UBICACIONES.VCHA_UBI_UBICACION = '" + Me.txt_ubicacion_origen + "')"
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_aRT_ARTICULO_ID)
               list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ART_NOMBRE_ESPAÑOL), "", rs!VCHA_ART_NOMBRE_ESPAÑOL)
               list_item.SubItems(2) = IIf(IsNull(rs!floa_ubi_cantidad), 0, rs!floa_ubi_cantidad)
               rs.MoveNext
         Wend
         rs.Close
         Me.frm_lista.Visible = True
         Me.lv_lista.SetFocus
      Else
         MsgBox "No se a seleccionado una ubicación", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.txt_ubicacion_origen <> "" Then
         If Me.txt_codigo <> "" Then
            rs.Open "select * from tb_articulos where vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_codigo = Me.txt_codigo
            Else
               rsaux.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  rsaux1.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + IIf(IsNull(rsaux!VCHA_aRT_ARTICULO_ID), "", rsaux!VCHA_aRT_ARTICULO_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     var_codigo = rsaux1!VCHA_aRT_ARTICULO_ID
                  Else
                     var_codigo = ""
                  End If
                  rsaux1.Close
               Else
                  var_codigo = ""
               End If
               rsaux.Close
            End If
            rs.Close
            If var_codigo <> "" Then
               Me.txt_codigo = var_codigo
               rs.Open "SELECT dbo.TB_EXISTENCIAS_UBICACIONES.VCHA_UBI_UBICACION, dbo.TB_EXISTENCIAS_UBICACIONES.VCHA_ART_ARTICULO_ID, dbo.TB_EXISTENCIAS_UBICACIONES.FLOA_UBI_CANTIDAD , dbo.TB_Articulos.VCHA_ART_NOMBRE_ESPAÑOL FROM dbo.TB_EXISTENCIAS_UBICACIONES INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_EXISTENCIAS_UBICACIONES.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID WHERE (dbo.TB_EXISTENCIAS_UBICACIONES.VCHA_UBI_UBICACION = '" + Me.txt_ubicacion_origen + "') AND (dbo.TB_EXISTENCIAS_UBICACIONES.VCHA_ART_ARTICULO_ID = '" + var_codigo + "')", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  Me.txt_descripcion = IIf(IsNull(rs!VCHA_ART_NOMBRE_ESPAÑOL), "", rs!VCHA_ART_NOMBRE_ESPAÑOL)
                  Me.txt_cantidad_origen = Format(IIf(IsNull(rs!floa_ubi_cantidad), 0, rs!floa_ubi_cantidad), "###,###,##0.00")
               Else
                  MsgBox "La ubicación " + Me.txt_ubicacion_origen + " no contiene el artículo indicado", vbOKOnly, "ATENCION"
               End If
               rs.Close
               Me.txt_descripcion.SetFocus
            Else
               MsgBox "El código no existe", vbOKOnly, "ATENCION"
            End If
         Else
            Me.txt_descripcion = ""
            Me.txt_cantidad_destino = ""
            Me.txt_cantidad_origen = ""
            Me.txt_ubicacion_destino = ""
         End If
      Else
         MsgBox "Se debe de indicar primero la ubicación origen", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_descripcion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_cantidad_origen.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_ubicacion_destino_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_cantidad_destino.SetFocus
   End If
End Sub

Private Sub txt_ubicacion_origen_Change()
   Me.txt_codigo = ""
   Me.txt_descripcion = ""
   Me.txt_cantidad_origen = ""
   Me.txt_ubicacion_destino = ""
   Me.txt_cantidad_destino = ""
End Sub

Private Sub txt_ubicacion_origen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_codigo.SetFocus
   End If
End Sub
