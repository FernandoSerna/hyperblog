VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmpaqueteria_almacen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Paqueteria del almacen"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5415
      Picture         =   "frmpaqueteria_almacen.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   435
      Picture         =   "frmpaqueteria_almacen.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frmpaqueteria_almacen.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   0
      Width           =   330
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2565
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   21
      Top             =   90
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   105
      TabIndex        =   17
      Top             =   1800
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   1785
         TabIndex        =   18
         Top             =   165
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3735
         TabIndex        =   19
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
         Caption         =   "Busqueda de linea:"
         Height          =   195
         Left            =   195
         TabIndex        =   20
         Top             =   195
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Paqueterias  "
      Height          =   1410
      Left            =   105
      TabIndex        =   10
      Top             =   375
      Width           =   5655
      Begin VB.TextBox txt_referencia 
         Height          =   315
         Left            =   2970
         TabIndex        =   6
         Top             =   915
         Width           =   615
      End
      Begin VB.TextBox txt_seguro 
         Height          =   315
         Left            =   1290
         TabIndex        =   5
         Top             =   915
         Width           =   615
      End
      Begin VB.TextBox txt_nombre 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   4
         Top             =   570
         Width           =   4155
      End
      Begin VB.TextBox txt_paqueteria 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   3
         Top             =   225
         Width           =   630
      End
      Begin VB.TextBox txt_costo_seguro 
         Height          =   315
         Left            =   4815
         TabIndex        =   7
         Top             =   915
         Width           =   615
      End
      Begin VB.CommandButton cmd_precios_paqueteria 
         Caption         =   "$"
         Height          =   315
         Left            =   5115
         TabIndex        =   11
         Top             =   165
         Width           =   315
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Referencia:"
         Height          =   195
         Index           =   4
         Left            =   2100
         TabIndex        =   16
         Top             =   975
         Width           =   825
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Seguro:"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   15
         Top             =   975
         Width           =   555
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   14
         Top             =   630
         Width           =   885
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   13
         Top             =   285
         Width           =   450
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Costo seguro:"
         Height          =   195
         Index           =   3
         Left            =   3750
         TabIndex        =   12
         Top             =   975
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4815
      Left            =   105
      TabIndex        =   8
      Top             =   2355
      Width           =   5655
      Begin MSComctlLib.ListView lv_paqueterias 
         Height          =   4635
         Left            =   45
         TabIndex        =   9
         Top             =   135
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   8176
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
         NumItems        =   5
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
            Text            =   "SEGURO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "REFERENCIA"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "COSTO SEGURO"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   3030
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
            Picture         =   "frmpaqueteria_almacen.frx":083E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpaqueteria_almacen.frx":1118
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   2580
      Top             =   -315
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
            Picture         =   "frmpaqueteria_almacen.frx":19F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpaqueteria_almacen.frx":22CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpaqueteria_almacen.frx":2BA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpaqueteria_almacen.frx":3142
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpaqueteria_almacen.frx":3A1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpaqueteria_almacen.frx":42F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpaqueteria_almacen.frx":4BD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpaqueteria_almacen.frx":4CE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpaqueteria_almacen.frx":4DF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpaqueteria_almacen.frx":4F08
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpaqueteria_almacen.frx":501A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   105
      TabIndex        =   22
      Top             =   255
      Visible         =   0   'False
      Width           =   5655
   End
End
Attribute VB_Name = "frmpaqueteria_almacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim numero_items_lineas As Integer
Dim bitacora As Boolean




Private Sub cmd_guardar_Click()
Dim var_posible As Boolean
   var_posible = True
   If var_posible = True Then
      Call pro_guardar_lineas
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open "select * from tb_paqueteria", cnn, adOpenDynamic, adLockOptimistic
      If rs.BOF Then
         cmd_guardar.Enabled = False
      Else
         cmd_guardar.Enabled = True
      End If
      rs.Close
   Else
      MsgBox "Clave de paqueteria ya existe", vbOKOnly, "ATENCION"
   End If
End Sub


Private Sub cmd_nuevo_Click()
   Call pro_limpiatextos(Me)
   rs.Open "select max(cast(vcha_PAQ_CLAVE_ID as integer)) from TB_PAQUETERIA", cnn, adOpenDynamic, adLockOptimistic
   If rs.EOF Then
      Me.txt_paqueteria = "001"
   Else
      var_clave = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
      If Len(CStr(var_clave)) = 1 Then
         Me.txt_paqueteria = "00" + Trim(CStr(var_clave))
      Else
         If Len(CStr(var_clave)) = 2 Then
            Me.txt_paqueteria = "0" + Trim(CStr(var_clave))
         Else
            Me.txt_paqueteria = Trim(CStr(var_clave))
         End If
      End If
   End If
   rs.Close
   txt_paqueteria.Enabled = False
   
   Me.txt_nombre.SetFocus: var_modifica_registro_linea = False
   cmd_guardar.Enabled = True
End Sub

Private Sub cmd_precios_paqueteria_Click()
   If Me.txt_paqueteria <> "" Then
      var_paqueteria = Me.txt_paqueteria
      frmprecios_paqueteria.txt_paqueteria = Me.txt_paqueteria
      frmprecios_paqueteria.txt_nombre = Me.txt_nombre
      frmprecios_paqueteria.Show
   Else
      MsgBox "Debe de seleccionar una paqueteria", vbOKOnly, "ATENCION"
   End If
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
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 0
   Left = 2900
   var_modifica_registro_linea = True
   lv_paqueterias.SmallIcons = ImageList
   Call pro_llena_listview1
   pro_textos
   rs.Open "select * from TB_PAQUETERIA WHERE VCHA_PAQ_NOMBRE LIKE '%ALMACEN%'", cnn, adOpenDynamic, adLockOptimistic
   If rs.BOF Then
      cmd_guardar.Enabled = False
   Else
      cmd_guardar.Enabled = True
   End If
   rs.Close
   txt_paqueteria.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub lv_paqueterias_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_paqueterias, ColumnHeader)
End Sub

Private Sub lv_paqueterias_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Set lv_paqueterias.selectedItem = Item
   pro_textos
   var_modifica_registro_linea = True
   txt_paqueteria.Enabled = False
End Sub


Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_paqueterias.SetFocus
      Call pro_avanzar(Me, lv_paqueterias, Button)
      lv_paqueterias.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_paqueterias.ListItems(1).Selected = True
      lv_paqueterias.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 4 Then
      numero_items_lineas = lv_paqueterias.ListItems.Count
      lv_paqueterias.ListItems(numero_items_lineas).Selected = True
      lv_paqueterias.selectedItem.EnsureVisible
      pro_textos
   End If
err0:
End Sub


Sub pro_guardar_lineas()
Dim list_item As ListItem
   Dim ok As Boolean
   ok = True
   If txt_paqueteria <> "" And txt_nombre <> "" Then
      If IsNumeric(Me.txt_costo_seguro) Then
         If IsNumeric(Me.txt_referencia) Then
            If IsNumeric(Me.txt_seguro) Then
            
               rs.Open "SELECT * FROM TB_PAQUETERIA WHERE VCHA_PAQ_CLAVE_ID = '" + Me.txt_paqueteria + "'", cnn, adOpenDynamic, adLockOptimistic
               If rs.EOF Then
                  rsaux.Open "INSERT INTO TB_PAQUETERIA (VCHA_PAQ_CLAVE_ID,VCHA_PAQ_NOMBRE, FLOA_PAQ_SEGURO, FLOA_PAQ_COSTO_REFERENCIA, FLOA_PAQ_SEGURO_COSTO) VALUES ('" + Me.txt_paqueteria + "', '" + Me.txt_nombre + "'," + Me.txt_seguro + "," + Me.txt_referencia + "," + Me.txt_costo_seguro + ")", cnn, adOpenDynamic, adLockOptimistic
                  Set list_item = lv_paqueterias.ListItems.Add(, , txt_paqueteria)
                  list_item.SubItems(1) = Me.txt_nombre
                  list_item.SubItems(2) = Me.txt_seguro
                  list_item.SubItems(3) = Me.txt_referencia
                  list_item.SubItems(4) = Me.txt_costo_seguro
                  list_item.EnsureVisible
                  list_item.Selected = True
                  MsgBox "Se a inseertado el registro", vbOKOnly, "ATENCION"
               Else
                  var_si = MsgBox("Desea ejecutar los cambios", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     rsaux.Open "UPDATE TB_PAQUETERIA SET VCHA_PAQ_NOMBRE = '" + Me.txt_nombre + "', FLOA_PAQ_SEGURO = " + Me.txt_seguro + ", FLOA_PAQ_COSTO_REFERENCIA = " + Me.txt_referencia + ", FLOA_PAQ_SEGURO_COSTO = " + Me.txt_costo_seguro + "  WHERE VCHA_PAQ_CLAVE_ID = '" + Me.txt_paqueteria + "'", cnn, adOpenDynamic, adLockOptimistic
                     lv_paqueterias.ListItems.Item(lv_paqueterias.selectedItem.Index).Checked = False
                     lv_paqueterias.ListItems.Item(lv_paqueterias.selectedItem.Index) = Me.txt_paqueteria
                     lv_paqueterias.ListItems.Item(lv_paqueterias.selectedItem.Index).ListSubItems(1) = Me.txt_nombre
                     lv_paqueterias.ListItems.Item(lv_paqueterias.selectedItem.Index).ListSubItems(2) = Me.txt_seguro
                     lv_paqueterias.ListItems.Item(lv_paqueterias.selectedItem.Index).ListSubItems(3) = Me.txt_referencia
                     lv_paqueterias.ListItems.Item(lv_paqueterias.selectedItem.Index).ListSubItems(4) = Me.txt_costo_seguro
                     lv_paqueterias.ListItems.Item(lv_paqueterias.selectedItem.Index).Selected = True
                     MsgBox "Se han ejecutado los cambios correctamente", vbOKOnly, "ATENCION"
                  End If
               End If
               rs.Close
            Else
               MsgBox "Importe de seguro incorrecto", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Importe de referencia incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Costo de seguro incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Sub pro_elimina_lineas()
   
End Sub


Sub pro_llena_listview1()
   Dim list_item As ListItem
   rs.Open "select * from TB_PAQUETERIA WHERE VCHA_PAQ_NOMBRE LIKE '%ALMACEN%'", cnn, adOpenDynamic, adLockOptimistic
   numero_items_lineas = 0
   While Not rs.EOF
      Set list_item = lv_paqueterias.ListItems.Add(, , rs!vcha_paq_clave_id)
      list_item.SubItems(1) = IIf(IsNull(rs!vcha_paq_nombre), "", rs!vcha_paq_nombre)
      list_item.SubItems(2) = IIf(IsNull(rs!floa_paq_seguro), "", rs!floa_paq_seguro)
      list_item.SubItems(3) = IIf(IsNull(rs!floa_paq_costo_referencia), "", rs!floa_paq_costo_referencia)
      list_item.SubItems(4) = IIf(IsNull(rs!floa_paq_seguro), "", rs!floa_paq_seguro)
      rs.MoveNext:
    Wend
    rs.Close
End Sub


Sub pro_textos()
On Error GoTo err0:
Dim var_n As Double
   var_n = lv_paqueterias.ListItems.Count
   If var_n > 0 Then
      txt_paqueteria = lv_paqueterias.selectedItem
      Me.txt_nombre = lv_paqueterias.selectedItem.SubItems(1)
      Me.txt_seguro = lv_paqueterias.selectedItem.SubItems(2)
      Me.txt_referencia = lv_paqueterias.selectedItem.SubItems(3)
      Me.txt_costo_seguro = lv_paqueterias.selectedItem.SubItems(4)
   End If
   txt_paqueteria.Enabled = False
err0:
End Sub


Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_paqueterias, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub





Private Sub txt_costo_seguro_Change()
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub

Private Sub txt_costo_seguro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub

Private Sub txt_nombre_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_paqueteria_Change()
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub

Private Sub txt_paqueteria_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_referencia_Change()
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub

Private Sub txt_referencia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_seguro_Change()
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   var_hubo_cambios = True
End Sub

Private Sub txt_seguro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

