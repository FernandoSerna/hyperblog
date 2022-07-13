VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmseries 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Series"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_unidad 
      Height          =   300
      Left            =   7710
      TabIndex        =   13
      Top             =   930
      Width           =   705
   End
   Begin VB.TextBox txt_empresa 
      Height          =   300
      Left            =   7710
      TabIndex        =   12
      Top             =   600
      Width           =   705
   End
   Begin VB.Frame Frame3 
      Height          =   1695
      Left            =   105
      TabIndex        =   9
      Top             =   1620
      Width           =   5565
      Begin MSComctlLib.ListView lv_series 
         Height          =   1440
         Left            =   45
         TabIndex        =   10
         Top             =   150
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   2540
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Activa"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Factura"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nota de Crédito"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Nota Cargo"
            Object.Width           =   2469
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "  Series "
      Height          =   1215
      Left            =   90
      TabIndex        =   5
      Top             =   405
      Width           =   5565
      Begin VB.TextBox txt_nota_cargo 
         Height          =   315
         Left            =   4155
         TabIndex        =   18
         Top             =   840
         Width           =   1260
      End
      Begin VB.TextBox txt_nota_credito 
         Height          =   315
         Left            =   4155
         TabIndex        =   16
         Top             =   510
         Width           =   1260
      End
      Begin VB.TextBox txt_factura 
         Height          =   315
         Left            =   4155
         TabIndex        =   14
         Top             =   180
         Width           =   1260
      End
      Begin VB.CheckBox chk_activa 
         Caption         =   "Serie Activa"
         Height          =   210
         Left            =   870
         TabIndex        =   11
         Top             =   555
         Width           =   1410
      End
      Begin VB.TextBox txt_serie 
         Height          =   315
         Left            =   870
         MaxLength       =   3
         TabIndex        =   6
         Top             =   180
         Width           =   900
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Nota de Cargo:"
         Height          =   195
         Index           =   3
         Left            =   2880
         TabIndex        =   19
         Top             =   900
         Width           =   1080
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Nota de Crédito:"
         Height          =   195
         Index           =   2
         Left            =   2880
         TabIndex        =   17
         Top             =   570
         Width           =   1155
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Factura:"
         Height          =   195
         Index           =   1
         Left            =   2880
         TabIndex        =   15
         Top             =   240
         Width           =   585
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   7
         Top             =   240
         Width           =   450
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5250
      Picture         =   "frmseries.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1095
      Picture         =   "frmseries.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   765
      Picture         =   "frmseries.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   435
      Picture         =   "frmseries.frx":080E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmseries.frx":0910
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   90
      TabIndex        =   8
      Top             =   270
      Width           =   5580
   End
End
Attribute VB_Name = "frmseries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_modificar As Boolean
Private Sub llena_lista()
   Dim list_item As ListItem
   lv_series.ListItems.Clear
   rs.Open "select * from tb_Series where vcha_uor_unidad_id = '" + txt_unidad + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
         Set list_item = lv_series.ListItems.Add(, , rs(2).Value)
         list_item.SubItems(1) = IIf(IsNull(rs(3).Value), 0, rs(3).Value)
         list_item.SubItems(2) = IIf(IsNull(rs(4).Value), 0, rs(4).Value)
         list_item.SubItems(3) = IIf(IsNull(rs(5).Value), 0, rs(5).Value)
         list_item.SubItems(4) = IIf(IsNull(rs(6).Value), 0, rs(6).Value)
         rs.MoveNext
      Wend
   End If
   rs.Close
End Sub

Private Sub cmd_deshacer_Click()
   var_modificar = True
   txt_serie.Enabled = False
   chk_activa.Enabled = False
End Sub

Private Sub cmd_eliminar_Click()
   Dim si As Integer
   si = MsgBox("¿Deseas eliminar la serie?", vbYesNo, "ATENCION")
   If si = 6 Then
      rs.Open "delete from tb_series where vcha_emp_empresa_id = '" + txt_empresa + "' and vcha_uor_unidad_id = '" + txt_unidad + "' and vcha_Ser_serie_id = '" + txt_serie + "'", cnn, adOpenDynamic, adLockOptimistic
      Call llena_lista
   End If
End Sub

Private Sub cmd_guardar_Click()
   If Trim(txt_serie) <> "" Then
      If Trim(txt_factura) <> "" And Trim(txt_nota_credito) <> "" And Trim(txt_nota_cargo) <> "" Then
         If var_modificar = False Then
            rs.Open "select * from tb_Series where vcha_ser_serie_id = '" + txt_serie + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               rs.Close
               MsgBox "Ya existe una serie " + txt_serie, vbOKOnly, "ATENCION"
            Else
               rs.Close
               If chk_activa = 1 Then
                  rsaux2.Open "update tb_series set inte_Ser_activa = 0 where vcha_emp_empresa_id = '" + txt_empresa + "' and vcha_uor_unidad_id = '" + txt_unidad + "'", cnn, adOpenDynamic, adLockOptimistic
               End If
               rsaux2.Open "insert into TB_SERIES (vcha_emp_empresa_id,vcha_uor_unidad_id,vcha_ser_serie_id,inte_ser_activa,inte_ser_factura, inte_ser_nota_credito, inte_ser_nota_cargo) VALUES ('" + txt_empresa + "', '" + txt_unidad + "', '" + txt_serie + "', " + Str(chk_activa) + ", " + txt_factura + ", " + txt_nota_credito + ", " + txt_nota_cargo + ")", cnn, adOpenDynamic, adLockOptimistic
               Call llena_lista
            End If
         Else
         End If
         var_modificar = True
         txt_serie.Enabled = False
         chk_activa.Enabled = False
      Else
         MsgBox "Existen inconsistencias en los folios", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "se debe de indicar una serie", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   var_modificar = False
   txt_serie.Enabled = True
   chk_activa.Enabled = True
   txt_factura.Enabled = True
   txt_nota_credito.Enabled = True
   txt_nota_cargo.Enabled = True
   txt_serie = ""
   chk_activa = 0
   txt_serie.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   txt_empresa = frmunidadesorganizacionales.txt_empresa
   txt_unidad = frmunidadesorganizacionales.txt_unidad
   var_modificar = True
   txt_serie.Enabled = False
   chk_activa.Enabled = False
   txt_factura.Enabled = False
   txt_nota_credito.Enabled = False
   txt_nota_cargo.Enabled = False
   Call llena_lista
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_series)
End Sub

Private Sub lv_series_ItemClick(ByVal Item As MSComctlLib.ListItem)
   txt_serie = lv_series.selectedItem
   chk_activa = lv_series.selectedItem.SubItems(1)
   txt_factura = lv_series.selectedItem.SubItems(2)
   txt_nota_credito = lv_series.selectedItem.SubItems(3)
   txt_nota_cargo = lv_series.selectedItem.SubItems(4)
   txt_factura.Enabled = False
   txt_nota_credito.Enabled = False
   txt_nota_cargo.Enabled = False
   txt_serie.Enabled = False
   chk_activa.Enabled = False
End Sub

Private Sub txt_factura_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      txt_nota_credito.SetFocus
   End If
End Sub

Private Sub txt_factura_LostFocus()
   If Not IsNumeric(txt_factura) Then
      MsgBox "Factura Incorrecta", vbOKOnly, "ATENCION"
      txt_factura = ""
   End If
End Sub

Private Sub txt_nota_cargo_LostFocus()
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If Not IsNumeric(txt_nota_cargo) Then
      MsgBox "Nota de cargo incorrecta", vbOKOnly, "ATENCION"
      txt_nota_cargo = ""
   End If
End Sub

Private Sub txt_nota_credito_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46, 27
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      txt_nota_cargo.SetFocus
   End If
End Sub

Private Sub txt_nota_credito_LostFocus()
   If Not IsNumeric(txt_nota_credito) Then
      MsgBox "Nota de crédito incorrecta", vbOKOnly, "ATENCION"
      txt_nota_credito = ""
   End If
End Sub

Private Sub txt_serie_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      txt_factura.SetFocus
   End If
End Sub

