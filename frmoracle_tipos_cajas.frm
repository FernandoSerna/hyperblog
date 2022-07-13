VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_tipos_cajas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Catálogo de cajas"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   5340
      Left            =   75
      TabIndex        =   13
      Top             =   1860
      Width           =   5655
      Begin MSComctlLib.ListView lv_cajas 
         Height          =   5130
         Left            =   45
         TabIndex        =   5
         Top             =   135
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   9049
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
            Text            =   "Nombre"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Peso"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Caja "
      Height          =   915
      Left            =   75
      TabIndex        =   10
      Top             =   405
      Width           =   5655
      Begin VB.TextBox txt_descripcion 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   3
         Top             =   180
         Width           =   4155
      End
      Begin VB.TextBox txt_peso 
         Height          =   315
         Left            =   1290
         MaxLength       =   3
         TabIndex        =   4
         Top             =   525
         Width           =   975
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   12
         Top             =   240
         Width           =   885
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Peso:"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   11
         Top             =   585
         Width           =   405
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   75
      TabIndex        =   6
      Top             =   1320
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   1785
         TabIndex        =   7
         Top             =   165
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3735
         TabIndex        =   8
         Top             =   150
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
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
         Caption         =   "Busqueda de caja:"
         Height          =   195
         Left            =   195
         TabIndex        =   9
         Top             =   195
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmoracle_tipos_cajas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   390
      Picture         =   "frmoracle_tipos_cajas.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Guardar Alt + G"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5385
      Picture         =   "frmoracle_tipos_cajas.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   75
      TabIndex        =   14
      Top             =   285
      Visible         =   0   'False
      Width           =   5655
   End
End
Attribute VB_Name = "frmoracle_tipos_cajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim numero_items_lineas As Integer
Dim bitacora As Boolean




Private Sub cmd_guardar_Click()
        Dim var_posible As Boolean
        If Me.txt_descripcion <> "" Then
           If IsNumeric(Me.txt_peso) Then
              rs.Open "select * from tb_oracle_empaques where empaque = '" + Me.txt_descripcion + "'", cnn, adOpenDynamic, adLockOptimistic
              If Not rs.EOF Then
                 var_si = MsgBox("Se va a modificar el registro", vbYesNo, "ATENCION")
                 If var_si = 6 Then
                    rsaux.Open "UPDATE tb_oracle_empaques SET EMPAQUE = '" + Me.txt_descripcion + "', PESO = " + CStr(Me.txt_peso) + " WHERE EMPAQUE = '" + Trim(lv_cajas.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                    var_modifica_registro_linea = False
                    Call pro_actualiza_ListView
                 End If
              Else
                 rsaux.Open "INSERT INTO tb_oracle_empaques (EMPAQUE, PESO) VALUES ('" + Me.txt_descripcion + "'," + Me.txt_peso + ")", cnn, adOpenDynamic, adLockOptimistic
                 var_modifica_registro_linea = True
                 Call pro_actualiza_ListView
              End If
              rs.Close
           Else
              MsgBox "Peso incorrecto", vbOKOnly, "ATENCION"
           End If
        Else
           MsgBox "Descripcion del incorrecta", vbOKOnly, "ATENCION"
        End If
End Sub


Private Sub cmd_nuevo_Click()
   Me.txt_descripcion = ""
   Me.txt_peso = ""
   Me.txt_descripcion.SetFocus
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
   End If
   If Shift = 4 And KeyCode = 69 Then
   End If
   If Shift = 4 And KeyCode = 73 Then
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 0
   Left = 2900
   var_modifica_registro_linea = True
   Call pro_llena_listview1
   pro_textos
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_lineas)
End Sub

Private Sub lv_cajas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_cajas, ColumnHeader)
End Sub

Private Sub lv_cajas_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Set lv_cajas.selectedItem = Item
   pro_textos
   var_modifica_registro_linea = True
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_cajas.SetFocus
      Call pro_avanzar(Me, lv_cajas, Button)
      lv_cajas.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_cajas.ListItems(1).Selected = True
      lv_cajas.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 4 Then
      numero_items_lineas = lv_cajas.ListItems.Count
      lv_cajas.ListItems(numero_items_lineas).Selected = True
      lv_cajas.selectedItem.EnsureVisible
      pro_textos
   End If
err0:
End Sub

Sub pro_llena_listview1()
   Dim list_item As ListItem
   rs.Open "select * from tb_oracle_empaques", cnn, adOpenDynamic, adLockOptimistic
   numero_items_lineas = 0
   While Not rs.EOF
      Set list_item = lv_cajas.ListItems.Add(, , rs(0).Value)
      list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
      rs.MoveNext:
      numero_items_lineas = numero_items_lineas + 1
    Wend
    rs.Close
End Sub


Sub pro_textos()
On Error GoTo err0:
Dim var_n As Double
   var_n = lv_cajas.ListItems.Count
   If var_n > 0 Then
      txt_descripcion = lv_cajas.selectedItem
      txt_peso = lv_cajas.selectedItem.SubItems(1)
   End If
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro_linea = True Then
       Set list_item = lv_cajas.ListItems.Add(, , txt_descripcion)
       list_item.SubItems(1) = txt_peso
       list_item.EnsureVisible
       list_item.Selected = True
       numero_items_lineas = numero_items_lineas + 1
    Else
       lv_cajas.ListItems.Item(lv_cajas.selectedItem.Index).Checked = False
       lv_cajas.ListItems.Item(lv_cajas.selectedItem.Index) = txt_descripcion
       lv_cajas.ListItems.Item(lv_cajas.selectedItem.Index).ListSubItems(1) = txt_peso
       lv_cajas.ListItems.Item(lv_cajas.selectedItem.Index).Selected = True
    End If
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_cajas, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub






Private Sub txt_descripcion_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_peso_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub
