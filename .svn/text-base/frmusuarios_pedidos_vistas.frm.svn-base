VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmusuarios_pedidos_vistas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Usuarios para pedidos a vistas"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   1590
      TabIndex        =   21
      Top             =   405
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   22
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7937
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   23
         Top             =   135
         Width           =   5610
      End
   End
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   165
      TabIndex        =   20
      Top             =   375
      Width           =   7395
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7110
      Picture         =   "frmusuarios_pedidos_vistas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   810
      Picture         =   "frmusuarios_pedidos_vistas.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      Picture         =   "frmusuarios_pedidos_vistas.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Guardar Alt + G"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmusuarios_pedidos_vistas.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos para pedido a vistas "
      Height          =   2265
      Left            =   135
      TabIndex        =   0
      Top             =   480
      Width           =   7380
      Begin VB.TextBox txt_nombre_establecimiento 
         Height          =   315
         Left            =   2775
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1785
         Width           =   4470
      End
      Begin VB.TextBox txt_establecimiento 
         Height          =   315
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   9
         Top             =   1785
         Width           =   1440
      End
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   315
         Left            =   2775
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1425
         Width           =   4470
      End
      Begin VB.TextBox txt_cliente 
         Height          =   315
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1425
         Width           =   1440
      End
      Begin VB.TextBox txt_nombre_titular 
         Height          =   315
         Left            =   2775
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1065
         Width           =   4470
      End
      Begin VB.TextBox txt_titular 
         Height          =   315
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1065
         Width           =   1440
      End
      Begin VB.TextBox txt_nombre_agente 
         Height          =   315
         Left            =   2775
         MaxLength       =   50
         TabIndex        =   4
         Top             =   705
         Width           =   4470
      End
      Begin VB.TextBox txt_agente 
         Height          =   315
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   3
         Top             =   705
         Width           =   1440
      End
      Begin VB.TextBox txt_nombre_usuario 
         Height          =   315
         Left            =   2775
         MaxLength       =   50
         TabIndex        =   2
         Top             =   345
         Width           =   4470
      End
      Begin VB.TextBox txt_usuario 
         Height          =   315
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   1
         Top             =   345
         Width           =   1440
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Estabelcimiento:"
         Height          =   195
         Index           =   4
         Left            =   75
         TabIndex        =   15
         Top             =   1800
         Width           =   1155
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Index           =   3
         Left            =   75
         TabIndex        =   14
         Top             =   1440
         Width           =   525
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Titular:"
         Height          =   195
         Index           =   2
         Left            =   75
         TabIndex        =   13
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Index           =   1
         Left            =   75
         TabIndex        =   12
         Top             =   720
         Width           =   555
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   11
         Top             =   360
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmusuarios_pedidos_vistas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim var_tipo_lista As Integer

Private Sub cmd_guardar_Click()
   If Trim(Me.txt_usuario) <> "" Then
      If Trim(Me.txt_agente) <> "" Then
         If Trim(Me.txt_titular) <> "" Then
            If Trim(Me.txt_cliente) <> "" Then
               If Trim(Me.txt_establecimiento) <> "" Then
                  If rsaux2.State = 1 Then
                     rsaux2.Close
                  End If
                  rsaux2.Open "SELECT * FROM VW_CLIENTES WHERE VCHA_AGE_AGENTE_ID = '" + Me.txt_agente + "' AND VCHA_TIT_TITULAR_ID = '" + Me.txt_titular + "' AND VCHA_CLI_CLAVE_ID = '" + Me.txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux2.EOF Then
                     If rsaux4.State = 1 Then
                        rsaux4.Close
                     End If
                     rsaux4.Open "SELECT * FROM TB_DETALLE_ESTABLECIMIENTOS WHERE VCHA_ESB_ESTABLECIMIENTO_ID = '" + Me.txt_establecimiento + "' AND VCHA_CLI_CLAVE_ID = '" + Me.txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux4.EOF Then
                        rs.Open "select * from TB_USUARIOS_PEDIDOS_VISTAS where vcha_usu_usuario_id = '" + Me.txt_usuario + "'", cnn, adOpenDynamic, adLockBatchOptimistic
                        If Not rs.EOF Then
                           var_si = MsgBox("El usuario ya existe, ¿Desea aplicar los cambios?", vbYesNo, "ATENCION")
                           If var_si = 6 Then
                              rsaux.Open "UPDATE TB_USUARIOS_PEDIDOS_VISTAS SET VCHA_AGE_AGENTE_ID ='" + Me.txt_agente + "', VCHA_TIT_TITULAR_ID = '" + Me.txt_titular + "', VCHA_CLI_CLAVE_ID = '" + Me.txt_cliente + "', VCHA_ESB_ESTABLECIMIENTO_ID = '" + Me.txt_establecimiento + "' WHERE VCHA_USU_USUARIO_ID = '" + Me.txt_usuario + "'", cnn, adOpenDynamic, adLockOptimistic
                           End If
                        Else
                           rsaux.Open "INSERT INTO TB_USUARIOS_PEDIDOS_VISTAS (VCHA_USU_USUARIO_ID, VCHA_AGE_AGENTE_ID, VCHA_TIT_TITULAR_ID, VCHA_CLI_CLAVE_ID, VCHA_ESB_ESTABLECIMIENTO_ID) VALUES ('" + Me.txt_usuario + "','" + Me.txt_agente + "','" + Me.txt_titular + "','" + Me.txt_cliente + "','" + Me.txt_establecimiento + "')", cnn, adOpenDynamic, adLockOptimistic
                        End If
                        rs.Close
                     Else
                        MsgBox "El establecimiento no corresponde al cliente", vbOKOnly, "ATENCION"
                     End If
                     rsaux4.Close
                  Else
                     MsgBox "Inconsitencia en la información", vbOKOnly, "ATENCION"
                  End If
                  rsaux2.Close
               Else
                  MsgBox "No se a seleccionado el establecimiento", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "No se a seleccionado un cliente", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "No se a seleccionado un titular", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "No se a seleccionado un agente", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a seleccionado un usuario", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_agente = ""
   Me.txt_cliente = ""
   Me.txt_establecimiento = ""
   Me.txt_titular = ""
   Me.txt_usuario = ""
   Me.txt_nombre_agente = ""
   Me.txt_nombre_cliente = ""
   Me.txt_nombre_establecimiento = ""
   Me.txt_nombre_titular = ""
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 2300
   Left = 2100
   Me.frm_lista.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      If var_tipo_lista = 1 Then
         Me.txt_usuario.SetFocus
      End If
      If var_tipo_lista = 2 Then
         Me.txt_agente.SetFocus
      End If
      If var_tipo_lista = 3 Then
         Me.txt_titular.SetFocus
      End If
      If var_tipo_lista = 4 Then
         Me.txt_cliente.SetFocus
      End If
      If var_tipo_lista = 5 Then
         Me.txt_establecimiento.SetFocus
      End If
   End If
   If KeyAscii = 13 Then
      If Me.lv_lista.ListItems.Count > 0 Then
         If var_tipo_lista = 1 Then
            Me.txt_usuario = lv_lista.selectedItem
            Me.txt_nombre_usuario = lv_lista.selectedItem.SubItems(1)
            Me.txt_usuario.SetFocus
         End If
         If var_tipo_lista = 2 Then
            Me.txt_agente = lv_lista.selectedItem
            Me.txt_nombre_agente = lv_lista.selectedItem.SubItems(1)
            Me.txt_agente.SetFocus
         End If
         If var_tipo_lista = 3 Then
            Me.txt_titular = lv_lista.selectedItem
            Me.txt_nombre_titular = lv_lista.selectedItem.SubItems(1)
            Me.txt_titular.SetFocus
         End If
         If var_tipo_lista = 4 Then
            Me.txt_cliente = lv_lista.selectedItem
            Me.txt_nombre_cliente = lv_lista.selectedItem.SubItems(1)
            Me.txt_cliente.SetFocus
         End If
         If var_tipo_lista = 5 Then
            Me.txt_establecimiento = lv_lista.selectedItem
            Me.txt_nombre_establecimiento = lv_lista.selectedItem.SubItems(1)
            Me.txt_establecimiento.SetFocus
         End If
      
      End If
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub txt_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from vw_pedidos_2 where vcha_emp_empresa_id = '" + var_empresa + "' and char_tpe_tipo_pedido_id = 'V' order by vcha_age_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_age_agente_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_age_nombre), "", rs!vcha_age_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
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
   End If
End Sub

Private Sub txt_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Or KeyAscii = 44 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_agente_LostFocus()
   If Trim(Me.txt_agente) <> "" Then
      rs.Open "select * from vw_pedidos_2 where vcha_age_Agente_id = '" + Me.txt_agente + "'  and char_tpe_tipo_pedido_id = 'V' and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_agente = IIf(IsNull(rs!vcha_age_nombre), "", rs!vcha_age_nombre)
      Else
         MsgBox "Clave de agente no existe", vbOKOnly, "ATENCION"
         Me.txt_agente = ""
         Me.txt_nombre_agente = ""
      End If
      rs.Close
   Else
      Me.txt_nombre_agente = ""
   End If
End Sub

Private Sub txt_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from VW_PEDIDOS_2 where VCHA_TIT_TITULAR_ID = '" + Me.txt_titular + "'  and char_tpe_tipo_pedido_id = 'V' order by vcha_CLI_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_cli_nombre), "", rs!vcha_cli_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CLIENTES"
      var_tipo_lista = 4
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
   If KeyAscii = 39 Or KeyAscii = 61 Or KeyAscii = 44 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_cliente_LostFocus()
   If Trim(Me.txt_cliente) <> "" Then
      rs.Open "SELECT * FROM VW_PEDIDOS_2 WHERE VCHA_AGE_AGENTE_ID = '" + Me.txt_agente + "'  and char_tpe_tipo_pedido_id = 'V' AND VCHA_CLI_CLAVE_ID = '" + Me.txt_cliente + "' AND VCHA_TIT_TITULAR_ID = '" + Me.txt_titular + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_cliente = IIf(IsNull(rs!vcha_cli_nombre), "", rs!vcha_cli_nombre)
      Else
         MsgBox "Clave de cliente incorrecta", vbOKOnly, "ATENCION"
         Me.txt_cliente = ""
         Me.txt_nombre_cliente = ""
      End If
      rs.Close
   Else
      Me.txt_nombre_cliente = ""
   End If
End Sub

Private Sub txt_establecimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from VW_PEDIDOS_2 where VCHA_CLI_CLAVE_ID = '" + Me.txt_cliente + "'  and char_tpe_tipo_pedido_id = 'V' order by vcha_ESB_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_Esb_establecimiento_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_esb_nombre), "", rs!vcha_esb_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ESTABLECIMIENTOS"
      var_tipo_lista = 5
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

Private Sub txt_establecimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Or KeyAscii = 44 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_establecimiento_LostFocus()
   If Trim(Me.txt_establecimiento) <> "" Then
      rs.Open "SELECT * FROM VW_PEDIDOS_2 WHERE VCHA_AGE_aGENTE_ID = '" + Me.txt_agente + "'  and char_tpe_tipo_pedido_id = 'V' AND VCHA_TIT_TITULAR_ID = '" + Me.txt_titular + "' AND VCHA_CLI_CLAVE_ID = '" + Me.txt_cliente + "' AND VCHA_ESB_ESTABLECIMIENTO_ID = '" + Me.txt_establecimiento + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_establecimiento = IIf(IsNull(rs!vcha_esb_nombre), "", rs!vcha_esb_nombre)
      Else
         MsgBox "Clave de establecimiento incorrecto", vbOKOnly, "ATENCION"
         Me.txt_establecimiento = ""
         Me.txt_nombre_establecimiento = ""
      End If
      rs.Close
   Else
      Me.txt_nombre_establecimiento = ""
   End If
End Sub

Private Sub txt_nombre_agente_Change()
   Me.txt_titular = ""
   Me.txt_nombre_titular = ""
   Me.txt_cliente = ""
   Me.txt_nombre_cliente = ""
   Me.txt_establecimiento = ""
   Me.txt_nombre_establecimiento = ""
End Sub

Private Sub txt_nombre_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_agentes where vcha_emp_empresa_id = '" + var_empresa + "'  and char_tpe_tipo_pedido_id = 'V' order by vcha_age_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_age_agente_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_age_nombre), "", rs!vcha_age_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
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
   End If
End Sub

Private Sub txt_nombre_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_cliente_Change()
   Me.txt_establecimiento = ""
   Me.txt_nombre_establecimiento = ""
End Sub

Private Sub txt_nombre_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from VW_PEDIDOS_2 where VCHA_TIT_TITULAR_ID = '" + Me.txt_titular + "'  and char_tpe_tipo_pedido_id = 'V' order by vcha_CLI_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_cli_nombre), "", rs!vcha_cli_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CLIENTES"
      var_tipo_lista = 4
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

Private Sub txt_nombre_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_establecimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from VW_PEDIDOS_2 where VCHA_CLI_CLAVE_ID = '" + Me.txt_cliente + "'  and char_tpe_tipo_pedido_id = 'V' order by vcha_ESB_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_Esb_establecimiento_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_esb_nombre), "", rs!vcha_esb_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ESTABLECIMIENTOS"
      var_tipo_lista = 5
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

Private Sub txt_nombre_establecimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Me.cmd_guardar.SetFocus
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
  
End Sub

Private Sub txt_nombre_titular_Change()
   Me.txt_cliente = ""
   Me.txt_nombre_cliente = ""
   Me.txt_establecimiento = ""
   Me.txt_nombre_establecimiento = ""
End Sub

Private Sub txt_nombre_titular_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select vcha_tit_titular_id, vcha_tit_nombre from VW_PEDIDOS_2 where VCHA_AGE_AGENTE_ID = '" + Me.txt_agente + "'  and char_tpe_tipo_pedido_id = 'V' order by vcha_TIT_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_tit_titular_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_tit_NOMBRE), "", rs!VCHA_tit_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TITULARES"
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
   End If
End Sub

Private Sub txt_nombre_titular_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_usuario_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select vcha_usu_usuario_id, isnull(vcha_usu_nombre,'')+ ' '+ isnull(vcha_usu_apellidos,'') as nombre  from tb_usuarios order by vcha_usu_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_usu_usuario_id)
            list_item.SubItems(1) = IIf(IsNull(rs!nombre), "", rs!nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "USUARIOS"
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

Private Sub txt_nombre_usuario_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_titular_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select distinct vcha_tit_titular_id, vcha_tit_nombre from VW_PEDIDOS_2 where VCHA_AGE_AGENTE_ID = '" + Me.txt_agente + "'  and char_tpe_tipo_pedido_id = 'V' order by vcha_TIT_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_tit_titular_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_tit_NOMBRE), "", rs!VCHA_tit_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TITULARES"
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
   End If
End Sub

Private Sub txt_titular_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Or KeyAscii = 44 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_titular_LostFocus()
   If Trim(Me.txt_titular) <> "" Then
      rs.Open "SELECT * FROM VW_PEDIDOS_2 WHERE VCHA_AGE_AGENTE_ID = '" + Me.txt_agente + "'  and char_tpe_tipo_pedido_id = 'V' and vcha_tit_titular_id = '" + Me.txt_titular + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_titular = IIf(IsNull(rs!VCHA_tit_NOMBRE), "", rs!VCHA_tit_NOMBRE)
      Else
         MsgBox "Clave de titular incorrecta", vbOKOnly, "ATENCION"
         Me.txt_titular = ""
         Me.txt_nombre_titular = ""
      End If
      rs.Close
   Else
      Me.txt_nombre_titular = ""
   End If
End Sub

Private Sub txt_usuario_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select vcha_usu_usuario_id, isnull(vcha_usu_nombre,'')+ ' '+ isnull(vcha_usu_apellidos,'') as nombre  from tb_usuarios order by vcha_usu_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_usu_usuario_id)
            list_item.SubItems(1) = IIf(IsNull(rs!nombre), "", rs!nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "USUARIOS"
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

Private Sub txt_usuario_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Or KeyAscii = 44 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_usuario_LostFocus()
   If Trim(Me.txt_usuario) <> "" Then
      rs.Open "SELECT * FROM TB_USUARIOS WHERE VCHA_USU_USUARIO_ID = '" + Me.txt_usuario + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_usuario = IIf(IsNull(rs!vcha_usu_nombre), "", rs!vcha_usu_nombre) + " " + IIf(IsNull(rs!vcha_usu_apellidos), "", rs!vcha_usu_apellidos)
      Else
         Me.txt_usuario = ""
         Me.txt_nombre_usuario = ""
         MsgBox "Clave de usuario incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      Me.txt_nombre_usuario = ""
   End If
End Sub
