VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmarticulos_transito 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   3690
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   7695
      Begin VB.Frame frm_codigo_entrada 
         Height          =   1020
         Left            =   180
         TabIndex        =   2
         Top             =   1320
         Width           =   7335
         Begin VB.TextBox txt_codigo 
            Height          =   330
            Left            =   90
            TabIndex        =   5
            Top             =   465
            Width           =   1560
         End
         Begin VB.TextBox txt_descripcion 
            Height          =   330
            Left            =   1665
            TabIndex        =   4
            Top             =   465
            Width           =   5220
         End
         Begin VB.CommandButton cmd_aceptar_pedidos 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6915
            Picture         =   "frmarticulos_transito.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Aceptar Alt + A"
            Top             =   465
            Width           =   330
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000D&
            Caption         =   " Código entrada"
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   0
            TabIndex        =   6
            Top             =   15
            Width           =   7320
         End
      End
      Begin MSComctlLib.ListView lv_notas 
         Height          =   3495
         Left            =   60
         TabIndex        =   1
         Top             =   135
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   6165
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Entrada"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Consecutivo"
            Object.Width           =   0
         EndProperty
      End
   End
End
Attribute VB_Name = "frmarticulos_transito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_pedidos_Click()
   If Trim(Me.txt_codigo) <> "" Then
      var_si = MsgBox("Desa actualizar el registro", vbYesNo, "ATENCION")
      If var_si = 6 Then
         rsaux9.Open "UPDATE TB_TRANSiTO SET VCHA_aRT_ARTICULO_RECIVO = '" + Me.txt_codigo + "' WHERE INTE_TRA_CONSECUTIVO = " + CStr(CDbl(Me.lv_notas.selectedItem.SubItems(3))), cnn_admcdindustrial, adOpenDynamic, adLockOptimistic
         rsaux10.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + Me.lv_notas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
         If rsaux10.EOF Then
            rsaux9.Open "INSERT INTO TB_EQUIVALENCIAS (VCHA_aRT_ARTICULO_ID, VCHA_EQU_CODIGO_EQUIVALENTE) VALUES ('" + Me.txt_codigo + "', '" + Me.lv_notas.selectedItem + "')", cnn, adOpenDynamic, adLockOptimistic
         End If
         rsaux10.Close
         Me.lv_notas.selectedItem.SubItems(2) = Me.txt_codigo
         Me.lv_notas.SetFocus
      End If
   Else
      MsgBox "No se a seleccionado un código", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_aceptar_pedidos_KeyPress(KeyAscii As Integer)
   If KeyPress = 27 Then
      Me.frm_codigo_entrada.Visible = False
   End If
End Sub

Private Sub Form_Load()
   rsaux10.Open "select * from tb_plantas where vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn_admcdindustrial, adOpenDynamic, adLockOptimistic
   var_clave_unidad_planta = ""
   If Not rsaux10.EOF Then
      var_clave_unidad_planta = IIf(IsNull(rsaux10!vcha_pla_planta_id), "", rsaux10!vcha_pla_planta_id)
   End If
   rsaux10.Close

   'rsaux10.Open "select * from tb_transito where vcha_Tra_nota_Envio = '" + var_nota_traspasos_transito + "' and vcha_tra_planta_destino = '" + var_clave_unidad_planta + "'", cnn_admcdindustrial, adOpenDynamic, adLockOptimistic
   rsaux10.Open "select * from tb_transito where vcha_Tra_nota_Envio = '" + var_nota_traspasos_transito + "'", cnn_admcdindustrial, adOpenDynamic, adLockOptimistic
   While Not rsaux10.EOF
         Set list_item = lv_notas.ListItems.Add(, , rsaux10!VCHA_ART_ARTICULO_ID)
         list_item.SubItems(1) = IIf(IsNull(rsaux10!vcha_art_descripcion), "", rsaux10!vcha_art_descripcion)
         list_item.SubItems(2) = IIf(IsNull(rsaux10!vcha_art_articulo_recivo), "", rsaux10!vcha_art_articulo_recivo)
         list_item.SubItems(3) = rsaux10!inte_tra_consecutivo
         rsaux10.MoveNext
   Wend
   rsaux10.Close
   Me.frm_codigo_entrada.Visible = False
End Sub

Private Sub lv_notas_GotFocus()
   Me.frm_codigo_entrada.Visible = False
End Sub

Private Sub lv_notas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Me.txt_codigo = ""
      Me.txt_descripcion = ""
      Me.frm_codigo_entrada.Visible = True
      Me.txt_codigo.SetFocus
   End If
End Sub

Private Sub txt_codigo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frmbusqueda_articulo.Show 1
      Me.txt_codigo = var_codigo_seleccionado
   End If
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_descripcion.SetFocus
   End If
   If KeyAscii = 27 Then
      Me.frm_codigo_entrada.Visible = False
      Me.lv_notas.SetFocus
   End If
End Sub

Private Sub txt_codigo_LostFocus()
   If Me.txt_codigo <> "" Then
      rsaux10.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux10.EOF Then
         Me.txt_descripcion = IIf(IsNull(rsaux10!vcha_Art_nombre_español), "", rsaux10!vcha_Art_nombre_español)
      Else
         rsaux11.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux11.EOF Then
            rsaux9.Open "select * from tb_articulos where vcha_Art_Articulo_id = '" + IIf(IsNull(rsaux11!VCHA_ART_ARTICULO_ID), "", rsaux11!VCHA_ART_ARTICULO_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux9.EOF Then
               Me.txt_descripcion = IIf(IsNull(rsaux9!vcha_Art_nombre_español), "", rsaux9!vcha_Art_nombre_español)
            Else
               MsgBox "El código no existe", vbOKOnly, "ATENCION"
            End If
            rsaux9.Close
         Else
            MsgBox "El código no existe", vbOKOnly, "ATENCION"
         End If
         rsaux11.Close
      End If
      rsaux10.Close
   Else
      Me.txt_descripcion = ""
   End If
End Sub

Private Sub txt_descripcion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_aceptar_pedidos.SetFocus
   Else
      If KeyAscii = 27 Then
         Me.lv_notas.SetFocus
      Else
         KeyAscii = 0
      End If
   End If
End Sub
