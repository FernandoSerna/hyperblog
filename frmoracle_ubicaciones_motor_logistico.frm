VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_ubicaciones_motor_logistico 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ubicaciones "
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   675
      TabIndex        =   23
      Top             =   105
      Width           =   7305
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1980
         Left            =   45
         TabIndex        =   24
         Top             =   375
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   3493
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
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   8819
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H000000C0&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   25
         Top             =   120
         Width           =   7230
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " Artículo "
      Height          =   840
      Left            =   135
      TabIndex        =   22
      Top             =   1395
      Width           =   8160
      Begin VB.TextBox txt_codigo 
         Height          =   450
         Left            =   225
         TabIndex        =   2
         Top             =   270
         Width           =   1695
      End
      Begin VB.TextBox txt_descripcion 
         Height          =   450
         Left            =   1950
         TabIndex        =   3
         Top             =   270
         Width           =   6075
      End
   End
   Begin VB.CommandButton com_guardar 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   465
      Picture         =   "frmoracle_ubicaciones_motor_logistico.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   150
      Picture         =   "frmoracle_ubicaciones_motor_logistico.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7935
      Picture         =   "frmoracle_ubicaciones_motor_logistico.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   120
      Left            =   75
      TabIndex        =   18
      Top             =   300
      Width           =   8205
   End
   Begin VB.Frame Frame2 
      Caption         =   " Ubicaciones "
      Height          =   1950
      Left            =   135
      TabIndex        =   11
      Top             =   2295
      Width           =   8160
      Begin VB.TextBox txt_ubicacion_6 
         Height          =   465
         Left            =   5085
         TabIndex        =   9
         Top             =   1245
         Width           =   2925
      End
      Begin VB.TextBox txt_ubicacion_5 
         Height          =   465
         Left            =   5100
         TabIndex        =   8
         Top             =   735
         Width           =   2925
      End
      Begin VB.TextBox txt_ubicacion_4 
         Height          =   465
         Left            =   5100
         TabIndex        =   7
         Top             =   225
         Width           =   2925
      End
      Begin VB.TextBox txt_ubicacion_3 
         Height          =   465
         Left            =   960
         TabIndex        =   6
         Top             =   1320
         Width           =   2925
      End
      Begin VB.TextBox txt_ubicacion_2 
         Height          =   465
         Left            =   975
         TabIndex        =   5
         Top             =   810
         Width           =   2925
      End
      Begin VB.TextBox txt_ubicacion_1 
         Height          =   465
         Left            =   975
         TabIndex        =   4
         Top             =   300
         Width           =   2925
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Sabado:"
         Height          =   195
         Left            =   4425
         TabIndex        =   17
         Top             =   1335
         Width           =   600
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Viernes:"
         Height          =   195
         Left            =   4440
         TabIndex        =   16
         Top             =   825
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Jueves:"
         Height          =   195
         Left            =   4440
         TabIndex        =   15
         Top             =   315
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Miercoles:"
         Height          =   195
         Left            =   195
         TabIndex        =   14
         Top             =   1410
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Martes:"
         Height          =   195
         Left            =   210
         TabIndex        =   13
         Top             =   900
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lunes:"
         Height          =   195
         Left            =   210
         TabIndex        =   12
         Top             =   390
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Centro de negocios "
      Height          =   840
      Left            =   120
      TabIndex        =   10
      Top             =   525
      Width           =   8160
      Begin VB.TextBox txt_nombre_subinventario 
         Height          =   450
         Left            =   1950
         TabIndex        =   1
         Top             =   270
         Width           =   6075
      End
      Begin VB.TextBox txt_subinventario 
         Height          =   450
         Left            =   225
         TabIndex        =   0
         Top             =   270
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmoracle_ubicaciones_motor_logistico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objConn As New adodb.Connection
Dim objCmd As New adodb.Command
Dim objParm As adodb.Parameter
Dim comandoORA As New adodb.Command
Dim parametro As adodb.Parameter
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_tipo_lista As Integer
Private Sub cmd_nuevo_Click()
   Me.txt_subinventario = ""
   Me.txt_nombre_subinventario = ""
   Me.txt_ubicacion_1 = ""
   Me.txt_ubicacion_2 = ""
   Me.txt_ubicacion_3 = ""
   Me.txt_ubicacion_4 = ""
   Me.txt_ubicacion_5 = ""
   Me.txt_ubicacion_6 = ""
   Me.txt_subinventario.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub com_guardar_Click()
   If Me.txt_subinventario <> "" Then
      If Me.txt_codigo <> "" Then
         rs.Open "select * from tb_oracle_ubicaciones_motor_logistico where clave = '" + Me.txt_subinventario + "' and codigo = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            rsaux.Open "update tb_oracle_ubicaciones_motor_logistico set ubicacion_1 = '" + Me.txt_ubicacion_1 + "', ubicacion_2 = '" + Me.txt_ubicacion_2 + "', ubicacion_3 = '" + Me.txt_ubicacion_3 + "', ubicacion_4 = '" + Me.txt_ubicacion_4 + "', ubicacion_5 = '" + Me.txt_ubicacion_5 + "', ubicacion_6 = '" + Me.txt_ubicacion_6 + "' where clave = '" + Me.txt_subinventario + "' and codigo = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            MsgBox "Se a actualizado el registro", vbOKOnly, "ATENCION"
         Else
            rsaux.Open "insert into tb_oracle_ubicaciones_motor_logistico (clave, codigo, ubicacion_1, ubicacion_2, ubicacion_3, ubicacion_4, ubicacion_5, ubicacion_6) values ('" + Me.txt_subinventario + "', '" + Me.txt_codigo + "','" + Me.txt_ubicacion_1 + "','" + Me.txt_ubicacion_2 + "','" + Me.txt_ubicacion_3 + "','" + Me.txt_ubicacion_4 + "','" + Me.txt_ubicacion_5 + "','" + Me.txt_ubicacion_6 + "')", cnn, adOpenDynamic, adLockOptimistic
            MsgBox "Se a insertado el registro", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
      End If
   Else
   End If
End Sub

Private Sub Form_Load()
   Top = 1300
   Left = 1700
   Me.frm_lista.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.lv_lista.ListItems.Count > 0 Then
         Me.txt_subinventario = Me.lv_lista.selectedItem
         Me.txt_nombre_subinventario = Me.lv_lista.selectedItem.SubItems(1)
      End If
      Me.txt_subinventario.SetFocus
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub txt_codigo_Change()
   Me.txt_ubicacion_1 = ""
   Me.txt_ubicacion_2 = ""
   Me.txt_ubicacion_3 = ""
   Me.txt_ubicacion_4 = ""
   Me.txt_ubicacion_5 = ""
   Me.txt_ubicacion_6 = ""
   Me.txt_descripcion = ""
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_codigo_LostFocus()
   If Me.txt_codigo <> "" Then
      strconsulta = "select * from xxvia_system_items_b where segment1 = ? and organization_id = 93"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
           .Parameters.Append parametro
      End With
      Set rsaux8 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux8.EOF Then
         Me.txt_descripcion = IIf(IsNull(rsaux8!Description), "", rsaux8!Description)
         rs.Open "select * from tb_oracle_ubicaciones_motor_logistico where clave = '" + Me.txt_subinventario + "' and codigo = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_ubicacion_1 = IIf(IsNull(rs!ubicacion_1), "", rs!ubicacion_1)
            Me.txt_ubicacion_2 = IIf(IsNull(rs!ubicacion_2), "", rs!ubicacion_2)
            Me.txt_ubicacion_3 = IIf(IsNull(rs!ubicacion_3), "", rs!ubicacion_3)
            Me.txt_ubicacion_4 = IIf(IsNull(rs!ubicacion_4), "", rs!ubicacion_4)
            Me.txt_ubicacion_5 = IIf(IsNull(rs!ubicacion_5), "", rs!ubicacion_5)
            Me.txt_ubicacion_6 = IIf(IsNull(rs!ubicacion_6), "", rs!ubicacion_6)
         Else
            Me.txt_ubicacion_1 = ""
            Me.txt_ubicacion_2 = ""
            Me.txt_ubicacion_3 = ""
            Me.txt_ubicacion_4 = ""
            Me.txt_ubicacion_5 = ""
            Me.txt_ubicacion_6 = ""
         End If
         rs.Close
      Else
         MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
         Me.txt_codigo = ""
         Me.txt_descripcion = ""
      End If
      rsaux8.Close
   Else
      Me.txt_descripcion = ""
   End If
End Sub

Private Sub txt_nombre_subinventario_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)

End Sub

Private Sub txt_subinventario_Change()
   Me.txt_nombre_subinventario = ""
   Me.txt_ubicacion_1 = ""
   Me.txt_ubicacion_2 = ""
   Me.txt_ubicacion_3 = ""
   Me.txt_ubicacion_4 = ""
   Me.txt_ubicacion_5 = ""
   Me.txt_ubicacion_6 = ""
   
End Sub

Private Sub txt_subinventario_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Me.lv_lista.ListItems.Clear
      var_tipo_lista = 1
      lbl_lista = "Centros de negocios"
      rs.Open "select secondary_inventory_name , description from mtl_secondary_inventories where attribute3 = 'PTO_VTA'AND DESCRIPTION NOT LIKE '%COSTALES%' ORDER BY DESCRIPTION", cnnoracle_4, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!secondary_inventory_name)
            list_item.SubItems(1) = IIf(IsNull(rs!Description), "", rs!Description)
            rs.MoveNext
      Wend
      rs.Close
      Me.frm_lista.Visible = True
      Me.lv_lista.SetFocus
   End If
End Sub

Private Sub txt_subinventario_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_subinventario_LostFocus()
   If Me.txt_subinventario <> "" Then
      strconsulta = "select secondary_inventory_name, description from mtl_secondary_inventories where attribute3 = 'PTO_VTA'AND DESCRIPTION NOT LIKE '%COSTALES%' and secondary_inventory_name = ? ORDER BY DESCRIPTION"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_subinventario)
           .Parameters.Append parametro
      End With
      Set rsaux8 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux8.EOF Then
         Me.txt_nombre_subinventario = IIf(IsNull(rsaux8!Description), "", rsaux8!Description)
      Else
         MsgBox "El subinventario no existe", vbOKOnly, "ATENCION"
         Me.txt_subinventario = ""
         Me.txt_nombre_subinventario = ""
      End If
      rsaux8.Close
   Else
      Me.txt_nombre_subinventario = ""
   End If
End Sub

Private Sub txt_ubicacion_1_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_ubicacion_2_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_ubicacion_3_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_ubicacion_4_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_ubicacion_5_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_ubicacion_6_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.com_guardar.SetFocus
   End If
End Sub
