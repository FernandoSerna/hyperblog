VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_trajinantes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Trajinantes"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   30
      Picture         =   "frmoracle_trajinantes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   360
      Picture         =   "frmoracle_trajinantes.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Guardar Alt + G"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5340
      Picture         =   "frmoracle_trajinantes.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2775
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   7
      Top             =   90
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   "  Trajinantes "
      Height          =   1020
      Left            =   30
      TabIndex        =   2
      Top             =   465
      Width           =   5655
      Begin VB.TextBox txt_nombre 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   4
         Top             =   585
         Width           =   4275
      End
      Begin VB.TextBox txt_clave 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   3
         Top             =   255
         Width           =   900
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   6
         Top             =   645
         Width           =   600
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   315
         Width           =   450
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5415
      Left            =   30
      TabIndex        =   0
      Top             =   1440
      Width           =   5655
      Begin MSComctlLib.ImageList icono_encabezado 
         Left            =   0
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
               Picture         =   "frmoracle_trajinantes.frx":083E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmoracle_trajinantes.frx":1118
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lv_trajinantes 
         Height          =   5205
         Left            =   45
         TabIndex        =   1
         Top             =   150
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   9181
         View            =   3
         MultiSelect     =   -1  'True
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
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7056
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmoracle_trajinantes.frx":19F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_trajinantes.frx":22CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_trajinantes.frx":2BA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_trajinantes.frx":3142
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_trajinantes.frx":3A1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_trajinantes.frx":42F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_trajinantes.frx":4BD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_trajinantes.frx":4CE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_trajinantes.frx":4DF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_trajinantes.frx":4F08
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   30
      TabIndex        =   11
      Top             =   285
      Width           =   5655
   End
End
Attribute VB_Name = "frmoracle_trajinantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim comandoORA As New adodb.Command
Dim parametro As adodb.Parameter

Sub pro_textos()
On Error GoTo err0:
Dim var_n As Double
   var_n = lv_trajinantes.ListItems.Count
   If var_n > 0 Then
      txt_clave = lv_trajinantes.selectedItem
      txt_nombre = lv_trajinantes.selectedItem.SubItems(1)
   End If
err0:
End Sub


Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro_linea = True Then
       Set list_item = lv_trajinantes.ListItems.Add(, , txt_descripcion)
       list_item.SubItems(1) = txt_peso
       list_item.EnsureVisible
       list_item.Selected = True
       numero_items_lineas = numero_items_lineas + 1
    Else
       lv_trajinantes.ListItems.Item(lv_trajinantes.selectedItem.Index).Checked = False
       lv_trajinantes.ListItems.Item(lv_trajinantes.selectedItem.Index) = txt_descripcion
       lv_trajinantes.ListItems.Item(lv_trajinantes.selectedItem.Index).ListSubItems(1) = txt_peso
       lv_trajinantes.ListItems.Item(lv_trajinantes.selectedItem.Index).Selected = True
    End If
End Sub

Private Sub cmd_guardar_Click()
   If Trim(Me.txt_clave) = "" Then
      rs.Open "select max(clave)as clave from xxvia_tb_trajinantes", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_consecutivo = IIf(IsNull(rs!clave), 0, rs!clave) + 1
      Else
         var_consecutivo = 1
      End If
      rs.Close
      Me.txt_clave = var_consecutivo
      strconsulta = "insert into xxvia_Tb_trajinantes (clave, nombre) values (?,?)"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(var_consecutivo))
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_nombre)
           .Parameters.Append parametro
      End With
      Set rsaux4 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      Set list_item = Me.lv_trajinantes.ListItems.Add(, , Me.txt_clave)
      list_item.SubItems(1) = Me.txt_nombre
   Else
      strconsulta = "update xxvia_Tb_trajinantes set nombre = ? where clave = ?"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_nombre)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_clave))
           .Parameters.Append parametro
      End With
      Set rsaux4 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      Me.lv_trajinantes.selectedItem.SubItems(1) = Me.txt_nombre
   End If
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_clave = ""
   Me.txt_nombre = ""
   Me.txt_nombre.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Sub pro_llena_listview1()
   Dim list_item As ListItem
   Me.lv_trajinantes.ListItems.Clear
   rs.Open "select * from xxvia_tb_trajinantes", cnnoracle_4, adOpenDynamic, adLockOptimistic
   numero_items_transportes = 0
   While Not rs.EOF
      Set list_item = lv_trajinantes.ListItems.Add(, , rs(0).Value)
      list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
      rs.MoveNext:
      numero_items_transportes = numero_items_transportes + 1
    Wend
    rs.Close
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
   var_modifica_registro_transporte = True
   'lv_trajinantes.SmallIcons = ImageList1
   'Call pro_encabezadosView(Me, lv_trajinantes, False)
   Call pro_llena_listview1
   pro_textos

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   var_cadena_seguridad = ""
   Top = 0
   Left = 2900
   var_modifica_registro_linea = True
   Call pro_llena_listview1
   pro_textos
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_trajinantes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_trajinantes, ColumnHeader)
End Sub

Private Sub lv_trajinantes_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Call pro_textos
End Sub

Private Sub txt_clave_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_nombre.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub
