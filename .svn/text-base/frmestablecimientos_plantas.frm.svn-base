VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmestablecimientos_plantas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relación establecimiento - planta"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   720
      TabIndex        =   15
      Top             =   -90
      Width           =   7470
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   30
         TabIndex        =   16
         Top             =   480
         Width           =   7365
         _ExtentX        =   12991
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
            Text            =   "Clave"
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7057
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Empresa"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   17
         Top             =   120
         Width           =   7395
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      Picture         =   "frmestablecimientos_plantas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   465
      Picture         =   "frmestablecimientos_plantas.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Guardar Alt + G"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   795
      Picture         =   "frmestablecimientos_plantas.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7845
      Picture         =   "frmestablecimientos_plantas.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   105
      TabIndex        =   10
      Top             =   345
      Width           =   8235
   End
   Begin VB.Frame frmestablecimientos_plantas 
      Caption         =   " Datos "
      Height          =   1605
      Left            =   165
      TabIndex        =   0
      Top             =   480
      Width           =   7995
      Begin VB.TextBox txt_nombre_planta 
         Height          =   390
         Left            =   2880
         TabIndex        =   9
         Top             =   1110
         Width           =   4995
      End
      Begin VB.TextBox txt_planta 
         Height          =   390
         Left            =   1365
         TabIndex        =   8
         Top             =   1110
         Width           =   1485
      End
      Begin VB.TextBox txt_nombre_establecimiento 
         Height          =   390
         Left            =   2880
         TabIndex        =   6
         Top             =   675
         Width           =   4995
      End
      Begin VB.TextBox txt_establecimiento 
         Height          =   390
         Left            =   1365
         TabIndex        =   5
         Top             =   675
         Width           =   1485
      End
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   390
         Left            =   2880
         TabIndex        =   3
         Top             =   240
         Width           =   4995
      End
      Begin VB.TextBox txt_cliente 
         Height          =   390
         Left            =   1365
         TabIndex        =   2
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Planta:"
         Height          =   195
         Left            =   165
         TabIndex        =   7
         Top             =   1215
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Establecimiento:"
         Height          =   195
         Left            =   165
         TabIndex        =   4
         Top             =   780
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   165
         TabIndex        =   1
         Top             =   345
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmestablecimientos_plantas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tipo_lista As Integer
Private Sub cmd_guardar_Click()
   If Trim(Me.txt_cliente) <> "" Then
      If Trim(Me.txt_establecimiento) <> "" Then
         If Trim(Me.txt_planta) <> "" Then
            rs.Open "SELECT * FROM TB_ESTABLECIMIENTOS_PLANTAS WHERE VCHA_eSB_ESTABLECIMIENTO_ID = '" + Me.txt_establecimiento + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               rsaux.Open "UPDATE TB_ESTABLECIMIENTOS_PLANTAS SET VCHA_UOR_UNIDAD_ID = '" + Me.txt_planta + "' WHERE VCHA_eSB_ESTABLECIMIENTO_ID = '" + Me.txt_establecimiento + "'", cnn, adOpenDynamic, adLockOptimistic
            Else
               rsaux.Open "INSERT INTO TB_ESTABLECIMIENTOS_plantas (VCHA_eSB_eSTABLECIMIENTO_ID, VCHA_UOR_UNIDAD_ID) VALUES ('" + Me.txt_establecimiento + "','" + Me.txt_planta + "')", cnn, adOpenDynamic, adLockOptimistic
            End If
            rs.Close
         Else
            MsgBox "No se a seleccionado una planta", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "No se a seleccionado un establecimiento", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a seleccionado un cliente", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_cliente = ""
   Me.txt_establecimiento = ""
   Me.txt_nombre_cliente = ""
   Me.txt_nombre_establecimiento = ""
   Me.txt_nombre_planta = ""
   Me.txt_planta = ""
   Me.txt_cliente.SetFocus
End Sub


Private Sub cmd_salir_Click()
   Unload Me
End Sub
  
Private Sub Form_Load()
    Top = 2500
    Left = 1500
    Me.frm_lista.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_tipo_lista = 1 Then
         Me.txt_cliente = lv_lista.selectedItem
         Me.txt_nombre_cliente = lv_lista.selectedItem.SubItems(1)
         Me.txt_cliente.SetFocus
      End If
      If var_tipo_lista = 2 Then
         Me.txt_establecimiento = lv_lista.selectedItem
         Me.txt_nombre_establecimiento = Me.lv_lista.selectedItem.SubItems(1)
         Me.txt_establecimiento.SetFocus
      End If
      If var_tipo_lista = 3 Then
         Me.txt_planta = lv_lista.selectedItem
         Me.txt_nombre_planta = Me.lv_lista.selectedItem.SubItems(1)
         Me.txt_planta.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      If var_tipo_lista = 1 Then
         Me.txt_cliente.SetFocus
      End If
      If var_tipo_lista = 2 Then
         Me.txt_establecimiento.SetFocus
      End If
      If var_tipo_lista = 3 Then
         Me.txt_planta.SetFocus
      End If
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub txt_cliente_Change()
   Me.txt_nombre_cliente = ""
   Me.txt_establecimiento = ""
   Me.txt_nombre_establecimiento = ""
   Me.txt_planta = ""
   Me.txt_nombre_planta = ""
End Sub

Private Sub txt_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "SELECT     TOP 100 PERCENT dbo.TB_CLIENTES.VCHA_CLI_CLAVE_ID, dbo.TB_CLIENTES.VCHA_CLI_NOMBRE, dbo.TB_EMPRESAS.VCHA_EMP_EMPRESA_ID, dbo.TB_EMPRESAS.VCHA_EMP_NOMBRE FROM dbo.TB_CLIENTES INNER JOIN dbo.TB_EMPRESAS ON dbo.TB_CLIENTES.VCHA_EMP_EMPRESA_ID = dbo.TB_EMPRESAS.VCHA_EMP_EMPRESA_ID WHERE (dbo.TB_CLIENTES.VCHA_AGE_AGENTE_ID = '00100') ORDER BY dbo.TB_CLIENTES.VCHA_CLI_NOMBRE", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            list_item.SubItems(2) = IIf(IsNull(rs!VCHA_EMP_NOMBRE), "", rs!VCHA_EMP_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Clientes"
      var_tipo_lista = 1
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
End Sub

Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_cliente_LostFocus()
   If Trim(Me.txt_cliente) <> "" Then
      rs.Open "select * from tb_clientes where vcha_cli_clave_id = '" + Me.txt_cliente + "' and vcha_age_agente_id = '00100'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_cliente = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
         Me.txt_establecimiento = ""
         Me.txt_nombre_establecimiento = ""
         Me.txt_planta = ""
         Me.txt_nombre_planta = ""
      Else
         MsgBox "El cliente no existe", vbOKOnly, "ATENCION"
         Me.txt_cliente = ""
         Me.txt_nombre_cliente = ""
         Me.txt_establecimiento = ""
         Me.txt_nombre_establecimiento = ""
         Me.txt_planta = ""
         Me.txt_nombre_planta = ""
      End If
      rs.Close
   Else
      Me.txt_nombre_cliente = ""
      Me.txt_establecimiento = ""
      Me.txt_nombre_establecimiento = ""
      Me.txt_planta = ""
      Me.txt_nombre_planta = ""
   End If
End Sub

Private Sub txt_establecimiento_Change()
   Me.txt_planta = ""
   Me.txt_nombre_planta = ""
End Sub

Private Sub txt_establecimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If Trim(Me.txt_cliente) <> "" Then
         lv_lista.ListItems.Clear
         rs.Open "select * from VW_ESTABLECIMIENTOS where vcha_cli_clave_id = '" + Me.txt_cliente + "' order by vcha_ESB_nombre", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!vcha_ESB_ESTABLECIMIENTO_id)
               list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ESB_NOMBRE), "", rs!VCHA_ESB_NOMBRE)
               rs.MoveNext
         Wend
         rs.Close
         lbl_lista = "Establecimientos"
         var_tipo_lista = 2
         Dim var_n As Integer
         var_n = lv_lista.ListItems.Count
         If var_n > 6 Then
            lv_lista.ColumnHeaders(2).Width = 4270.71
         Else
            lv_lista.ColumnHeaders(2).Width = 4499.71
         End If
         frm_lista.Visible = True
         lv_lista.SetFocus
      Else
         MsgBox "No se a seleccionado un cliente", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_establecimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_nombre_establecimiento.SetFocus
   End If
End Sub

Private Sub txt_establecimiento_LostFocus()
   If Trim(Me.txt_establecimiento) <> "" Then
      rs.Open "select * from vw_establecimientos where vcha_cli_clave_id = '" + Me.txt_cliente + "' and vcha_esb_establecimiento_id = '" + Me.txt_establecimiento + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_establecimiento = rs!VCHA_ESB_NOMBRE
         rsaux.Open "SELECT * FROM TB_ESTABLECIMIENTOS_PLANTAS WHERE VCHA_ESB_ESTABLECIMIENTO_ID = '" + Me.txt_establecimiento + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            rsaux1.Open "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE VCHA_UOR_UNIDAD_ID = '" + IIf(IsNull(rsaux!vcha_uor_unidad_id), "", rsaux!vcha_uor_unidad_id) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux1.EOF Then
               Me.txt_planta = rsaux!vcha_uor_unidad_id
               Me.txt_nombre_planta = rsaux1!VCHA_UOR_NOMBRE
            End If
            rsaux1.Close
         End If
         rsaux.Close
      Else
         MsgBox "El establecimiento no existe o no corresponde al cliente seleccionado", vbOKOnly, "ATENCION"
         Me.txt_establecimiento = ""
         Me.txt_nombre_establecimiento = ""
         Me.txt_nombre_planta = ""
         Me.txt_planta = ""
      End If
      rs.Close
   Else
      Me.txt_nombre_establecimiento = ""
      Me.txt_planta = ""
      Me.txt_nombre_planta = ""
   End If
End Sub

Private Sub txt_nombre_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_establecimiento.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_establecimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_planta.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_planta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_planta_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If Trim(Me.txt_establecimiento) <> "" Then
         lv_lista.ListItems.Clear
         rs.Open "select a.vcha_uor_unidad_id, a.vcha_uor_nombre, b.vcha_pla_planta_id from tb_unidadesorganizacionales a, admcdindustrial.sid.dbo.tb_plantas b where a.vcha_uor_unidad_id = b.vcha_uor_unidad_id order by a.vcha_uor_nombre", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!vcha_uor_unidad_id)
               list_item.SubItems(1) = IIf(IsNull(rs!VCHA_UOR_NOMBRE), "", rs!VCHA_UOR_NOMBRE)
               list_item.SubItems(2) = IIf(IsNull(rs!vcha_pla_planta_id), "", rs!vcha_pla_planta_id)
               rs.MoveNext
         Wend
         rs.Close
         lbl_lista = "Plantas"
         var_tipo_lista = 3
         Dim var_n As Integer
         var_n = lv_lista.ListItems.Count
         If var_n > 6 Then
            lv_lista.ColumnHeaders(2).Width = 4270.71
         Else
            lv_lista.ColumnHeaders(2).Width = 4499.71
         End If
         frm_lista.Visible = True
         lv_lista.SetFocus
      Else
         MsgBox "No se a seleccionado un establecimiento", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_planta_KeyPress(KeyAscii As Integer)
   Me.txt_nombre_planta.SetFocus
End Sub

Private Sub txt_planta_LostFocus()
   If Trim(Me.txt_planta) <> "" Then
      rs.Open "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE VCHA_UOR_UNIDAD_ID = '" + Me.txt_planta + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_planta = rs!VCHA_UOR_NOMBRE
      Else
         MsgBox "Clave de planta incorrecta", vbOKOnly, "ATENCION"
         Me.txt_planta = ""
         Me.txt_nombre_planta = ""
      End If
      rs.Close
   Else
      Me.txt_nombre_planta = ""
   End If
End Sub
