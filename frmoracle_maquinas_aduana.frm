VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_maquinas_aduana 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cátalogo de máquinas para aduana"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   10215
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   675
      Picture         =   "frmoracle_maquinas_aduana.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   5745
      Left            =   30
      TabIndex        =   9
      Top             =   1425
      Width           =   10125
      Begin MSComctlLib.ListView lv_maquinas 
         Height          =   5505
         Left            =   30
         TabIndex        =   5
         Top             =   150
         Width           =   10050
         _ExtentX        =   17727
         _ExtentY        =   9710
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Máquina"
            Object.Width           =   4233
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Tipo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Aduana"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Estacion"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "DVR"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Puerto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Com"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Fraccionado"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Máquina "
      Height          =   1035
      Left            =   30
      TabIndex        =   6
      Top             =   390
      Width           =   10140
      Begin VB.ComboBox cmb_metodo_fraccionado 
         Height          =   315
         ItemData        =   "frmoracle_maquinas_aduana.frx":0102
         Left            =   9060
         List            =   "frmoracle_maquinas_aduana.frx":0104
         TabIndex        =   22
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox cmb_com_bascula 
         Height          =   315
         ItemData        =   "frmoracle_maquinas_aduana.frx":0106
         Left            =   7470
         List            =   "frmoracle_maquinas_aduana.frx":0122
         Sorted          =   -1  'True
         TabIndex        =   21
         Top             =   150
         Width           =   975
      End
      Begin VB.ComboBox cmb_puerto 
         Height          =   315
         ItemData        =   "frmoracle_maquinas_aduana.frx":013E
         Left            =   6390
         List            =   "frmoracle_maquinas_aduana.frx":01A5
         TabIndex        =   18
         Top             =   555
         Width           =   975
      End
      Begin VB.ComboBox cmb_dvr 
         Height          =   315
         ItemData        =   "frmoracle_maquinas_aduana.frx":027C
         Left            =   4935
         List            =   "frmoracle_maquinas_aduana.frx":028F
         TabIndex        =   16
         Top             =   555
         Width           =   825
      End
      Begin VB.ComboBox cmb_estacion 
         Height          =   315
         ItemData        =   "frmoracle_maquinas_aduana.frx":02A2
         Left            =   3180
         List            =   "frmoracle_maquinas_aduana.frx":02CA
         TabIndex        =   15
         Top             =   540
         Width           =   1155
      End
      Begin VB.ComboBox cmb_metodo_aduana 
         Height          =   315
         ItemData        =   "frmoracle_maquinas_aduana.frx":02F4
         Left            =   1530
         List            =   "frmoracle_maquinas_aduana.frx":02FE
         TabIndex        =   12
         Top             =   540
         Width           =   825
      End
      Begin VB.ComboBox cmb_tipo 
         Height          =   315
         ItemData        =   "frmoracle_maquinas_aduana.frx":030A
         Left            =   4935
         List            =   "frmoracle_maquinas_aduana.frx":0314
         TabIndex        =   4
         Top             =   165
         Width           =   1455
      End
      Begin VB.TextBox txt_nombre 
         Height          =   315
         Left            =   1530
         MaxLength       =   50
         TabIndex        =   3
         Top             =   180
         Width           =   2790
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Método fraccionado:"
         Height          =   195
         Index           =   5
         Left            =   7560
         TabIndex        =   23
         Top             =   630
         Width           =   1470
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "COM Bascula:"
         Height          =   195
         Left            =   6420
         TabIndex        =   20
         Top             =   210
         Width           =   1020
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Puerto:"
         Height          =   195
         Index           =   4
         Left            =   5850
         TabIndex        =   19
         Top             =   585
         Width           =   510
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "DVR:"
         Height          =   195
         Index           =   3
         Left            =   4440
         TabIndex        =   17
         Top             =   585
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Estación:"
         Height          =   195
         Left            =   2475
         TabIndex        =   14
         Top             =   615
         Width           =   660
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Metodo aduana:"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   13
         Top             =   600
         Width           =   1170
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   2
         Left            =   270
         TabIndex        =   8
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Index           =   0
         Left            =   4470
         TabIndex        =   7
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   30
      Picture         =   "frmoracle_maquinas_aduana.frx":032A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   345
      Picture         =   "frmoracle_maquinas_aduana.frx":042C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Guardar Alt + G"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9810
      Picture         =   "frmoracle_maquinas_aduana.frx":052E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   30
      TabIndex        =   10
      Top             =   270
      Visible         =   0   'False
      Width           =   10275
   End
End
Attribute VB_Name = "frmoracle_maquinas_aduana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim numero_items_lineas As Integer
Dim bitacora As Boolean




Private Sub cmb_tipo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub

Private Sub cmd_eliminar_Click()
   var_si = MsgBox("¿Desea eliminar la máquina?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      rs.Open "DELETE FROM TB_ORACLE_MAQUINAS WHERE MAQUINA = '" + Me.txt_nombre + "'", cnn, adOpenDynamic, adLockOptimistic
      Me.lv_maquinas.ListItems.Remove (Me.lv_maquinas.selectedItem.Index)
   End If
End Sub

Private Sub cmd_guardar_Click()
        Dim var_posible As Boolean
        Dim var_com_bascula As Integer
        If Me.txt_nombre <> "" Then
           If Me.cmb_tipo <> "" Then
              If Me.cmb_dvr = "" Then
                 Me.cmb_dvr = "0"
              End If
              If Me.cmb_puerto = "" Then
                 Me.cmb_puerto = "0"
              End If
              If Me.cmb_metodo_aduana = "" Then
                 Me.cmb_metodo_aduana = "NO"
              End If
              If Me.cmb_tipo = "ESTACION" Then
                 var_tipo = "E"
              End If
              If Me.cmb_tipo = "ADUANA" Then
                 var_tipo = "S"
              End If
              If Me.cmb_metodo_aduana = "SI" Then
                 var_metodo_aduana = 1
              End If
              If Trim(Me.cmb_metodo_aduana) = "NO" Then
                 var_metodo_aduana = 0
              End If
              If Trim(Me.cmb_com_bascula) = "" Then
                 var_com_bascula = 0
              Else
                 var_com_bascula = Me.cmb_com_bascula
              End If
              rs.Open "select * from tb_oracle_maquinas where MAQUINA = '" + Me.txt_nombre + "'", cnn, adOpenDynamic, adLockOptimistic
              If Not rs.EOF Then
                 var_si = MsgBox("Se va a modificar el registro", vbYesNo, "ATENCION")
                 If Me.cmb_estacion.Text = "" Then
                    Me.cmb_estacion.Text = "0"
                 End If
                 If Me.cmb_metodo_fraccionado = "" Then
                    Me.cmb_metodo_fraccionado = 0
                 End If
                 If var_si = 6 Then
                    If rsaux.State = 1 Then
                       rsaux.Close
                    End If
                    rsaux.Open "UPDATE tb_oracle_maquinas SET uso = '" + var_tipo + "', metodo_aduana = " + CStr(var_metodo_aduana) + ", estacion = " + Me.cmb_estacion.Text + ", DVR = " + Me.cmb_dvr + ", PUERTO = " + Me.cmb_puerto + ", com_bascula = " + CStr(var_com_bascula) + ", METODO_FRACCIONADO = " + CStr(Me.cmb_metodo_fraccionado) + "  WHERE MAQUINA = '" + Me.txt_nombre + "'", cnn, adOpenDynamic, adLockOptimistic
                    var_modifica_registro_linea = False
                    Call pro_actualiza_ListView
                 End If
              Else
                 If Me.cmb_metodo_fraccionado = "" Then
                    Me.cmb_metodo_fraccionado = 0
                 End If
                 rsaux.Open "INSERT INTO tb_oracle_MAQUINAS (MAQUINA, uso, metodo_aduana, estacion, DVR, PUERTO, COM_BASCULA, METODO_FRACCIONADO) VALUES ('" + Me.txt_nombre + "','" + var_tipo + "'," + CStr(var_metodo_aduana) + "," + Me.cmb_estacion.Text + "," + Me.cmb_dvr + "," + Me.cmb_puerto + "," + CStr(var_com_bascula) + "," + CStr(Me.cmb_metodo_fraccionado) + ")", cnn, adOpenDynamic, adLockOptimistic
                 var_modifica_registro_linea = True
                 Call pro_actualiza_ListView
              End If
              rs.Close
              
           Else
              MsgBox "No se a seleccionado el tipo de máquina", vbOKOnly, "ATENCION"
           End If
        Else
           MsgBox "No se indico el nombre de la máquina", vbOKOnly, "ATENCION"
        End If
End Sub


Private Sub cmd_nuevo_Click()
   Me.txt_nombre = ""
   Me.txt_nombre.SetFocus
   Me.cmb_tipo = "ESTACION"
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
   Left = 700
   var_modifica_registro_linea = True
   Call pro_llena_listview1
   pro_textos
   If var_clave_usuario_global = "U0000000358" Or var_clave_usuario_global = "8" Or var_clave_usuario_global = "U0000000397" Or var_clave_usuario_global = "U0000000519" Or var_clave_usuario_global = "U0000000628" Then
      Me.cmb_dvr.Enabled = True
      Me.cmb_puerto.Enabled = True
   Else
      Me.cmb_dvr.Enabled = False
      Me.cmb_puerto.Enabled = False
   End If
   Me.cmb_metodo_fraccionado.AddItem ("0")
   Me.cmb_metodo_fraccionado.AddItem ("1")
   
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_maquinas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_maquinas, ColumnHeader)
End Sub

Private Sub lv_maquinas_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Set lv_maquinas.selectedItem = Item
   pro_textos
   var_modifica_registro_linea = True
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_maquinas.SetFocus
      Call pro_avanzar(Me, lv_maquinas, Button)
      lv_maquinas.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_maquinas.ListItems(1).Selected = True
      lv_maquinas.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 4 Then
      numero_items_lineas = lv_maquinas.ListItems.Count
      lv_maquinas.ListItems(numero_items_lineas).Selected = True
      lv_maquinas.selectedItem.EnsureVisible
      pro_textos
   End If
err0:
End Sub

Sub pro_llena_listview1()
   Dim list_item As ListItem
   rs.Open "select * from tb_oracle_maquinas", cnn, adOpenDynamic, adLockOptimistic
   numero_items_lineas = 0
   While Not rs.EOF
      Set list_item = lv_maquinas.ListItems.Add(, , rs!maquina)
      var_tipo = IIf(IsNull(rs!USO), "", rs!USO)
      list_item.SubItems(1) = var_tipo
      list_item.SubItems(2) = IIf(IsNull(rs!METODO_ADUANA), 0, rs!METODO_ADUANA)
      list_item.SubItems(3) = IIf(IsNull(rs!estacion), 0, rs!estacion)
      list_item.SubItems(4) = IIf(IsNull(rs!DVR), 0, rs!DVR)
      list_item.SubItems(5) = IIf(IsNull(rs!puerto), 0, rs!puerto)
      list_item.SubItems(6) = IIf(IsNull(rs!COM_BASCULA), 0, rs!COM_BASCULA)
      list_item.SubItems(7) = IIf(IsNull(rs!metodo_fraccionado), 0, rs!metodo_fraccionado)
      
      rs.MoveNext:
      numero_items_lineas = numero_items_lineas + 1
    Wend
    rs.Close
End Sub


Sub pro_textos()
On Error GoTo err0:
Dim var_n As Double
   var_n = lv_maquinas.ListItems.Count
   If var_n > 0 Then
      txt_nombre = lv_maquinas.selectedItem
      If lv_maquinas.selectedItem.SubItems(1) = "S" Then
         Me.cmb_tipo = "ADUANA"
      End If
      If lv_maquinas.selectedItem.SubItems(1) = "E" Then
         Me.cmb_tipo = "ESTACION"
      End If
      If lv_maquinas.selectedItem.SubItems(2) = "0" Then
         Me.cmb_metodo_aduana = "NO"
      End If
      If Me.lv_maquinas.selectedItem.SubItems(2) = "1" Then
         Me.cmb_metodo_aduana = "SI"
      End If
      Me.cmb_estacion = Me.lv_maquinas.selectedItem.SubItems(3)
      Me.cmb_dvr = Me.lv_maquinas.selectedItem.SubItems(4)
      Me.cmb_puerto = Me.lv_maquinas.selectedItem.SubItems(5)
      Me.cmb_com_bascula = Me.lv_maquinas.selectedItem.SubItems(6)
      Me.cmb_metodo_fraccionado = Me.lv_maquinas.selectedItem.SubItems(7)
      
      
   End If
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro_linea = True Then
       Set list_item = lv_maquinas.ListItems.Add(, , txt_nombre)
       If Me.cmb_tipo = "ADUANA" Then
          list_item.SubItems(1) = "S"
       End If
       If Me.cmb_tipo = "ESTACION" Then
          list_item.SubItems(1) = "E"
       End If
       If Me.cmb_metodo_aduana = "SI" Then
          list_item.SubItems(2) = 1
       End If
       If Me.cmb_metodo_aduana = "NO" Then
          list_item.SubItems(2) = 0
       End If
       list_item.SubItems(3) = Me.cmb_estacion.Text
       If Me.cmb_dvr = "" Then
          list_item.SubItems(4) = "0"
       Else
          list_item.SubItems(4) = cmb_dvr
       End If
       If Me.cmb_puerto = "" Then
          list_item.SubItems(5) = "0"
       Else
          list_item.SubItems(5) = cmb_puerto
       End If
       If Me.cmb_com_bascula = "" Then
          list_item.SubItems(6) = "0"
       Else
          list_item.SubItems(6) = Me.cmb_com_bascula
       End If
       If Me.cmb_metodo_fraccionado = "" Then
          list_item.SubItems(7) = "0"
       Else
          list_item.SubItems(7) = Me.cmb_metodo_fraccionado
       End If
       
       list_item.EnsureVisible
       list_item.Selected = True
       numero_items_lineas = numero_items_lineas + 1
    Else
       If Me.cmb_dvr = "" Then
          VAR_DVR = "0"
       Else
          VAR_DVR = cmb_dvr
       End If
       If Me.cmb_puerto = "" Then
          var_puerto = "0"
       Else
          var_puerto = cmb_puerto
       End If
       lv_maquinas.ListItems.Item(lv_maquinas.selectedItem.Index).Checked = False
       lv_maquinas.ListItems.Item(lv_maquinas.selectedItem.Index) = Me.txt_nombre
       If Me.cmb_tipo = "ADUANA" Then
          lv_maquinas.ListItems.Item(lv_maquinas.selectedItem.Index).ListSubItems(1) = "S"
       End If
       If Me.cmb_tipo = "ESTACION" Then
          lv_maquinas.ListItems.Item(lv_maquinas.selectedItem.Index).ListSubItems(1) = "E"
       End If
       If Me.cmb_metodo_aduana = "SI" Then
          lv_maquinas.ListItems.Item(lv_maquinas.selectedItem.Index).ListSubItems(2) = 1
       End If
       If Me.cmb_metodo_aduana = "NO" Then
          lv_maquinas.ListItems.Item(lv_maquinas.selectedItem.Index).ListSubItems(2) = 0
       End If
       Me.lv_maquinas.ListItems.Item(lv_maquinas.selectedItem.Index).ListSubItems(3) = Me.cmb_estacion
       Me.lv_maquinas.ListItems.Item(lv_maquinas.selectedItem.Index).ListSubItems(4) = VAR_DVR
       Me.lv_maquinas.ListItems.Item(lv_maquinas.selectedItem.Index).ListSubItems(5) = var_puerto
       Me.lv_maquinas.ListItems.Item(lv_maquinas.selectedItem.Index).ListSubItems(6) = Me.cmb_com_bascula
       Me.lv_maquinas.ListItems.Item(lv_maquinas.selectedItem.Index).ListSubItems(7) = Me.cmb_metodo_fraccionado
       
       lv_maquinas.ListItems.Item(lv_maquinas.selectedItem.Index).Selected = True
    End If
End Sub








Private Sub txt_nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmb_tipo.SetFocus
   End If
End Sub


Private Sub txt_nombre_LostFocus()
   If Trim(Me.txt_nombre) <> "" Then
      rs.Open "select * from tb_oracle_maquinas where maquina = '" + Me.txt_nombre + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         If rs!USO = "E" Then
            Me.cmb_tipo = "ESTACION"
         End If
         If rs!USO = "S" Then
            Me.cmb_tipo = "ADUANA"
         End If
      End If
      rs.Close
   End If
End Sub
