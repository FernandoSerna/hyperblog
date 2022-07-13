VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_asignar_clientes_rutas_distribucion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignar clientes a rutas"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   17325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_migrar_oracle 
      Caption         =   "ORACLE"
      Height          =   330
      Left            =   360
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Frame Frame1 
      Height          =   2355
      Left            =   0
      TabIndex        =   11
      Top             =   405
      Width           =   17235
      Begin VB.TextBox txt_clave_establecimiento 
         Height          =   390
         Left            =   1350
         TabIndex        =   20
         Top             =   1020
         Width           =   1815
      End
      Begin VB.TextBox txt_establecimiento 
         Height          =   390
         Left            =   3180
         TabIndex        =   5
         Top             =   1020
         Width           =   1815
      End
      Begin VB.TextBox txt_nombre_establecimiento 
         Height          =   420
         Left            =   1350
         TabIndex        =   6
         Top             =   1425
         Width           =   10110
      End
      Begin VB.TextBox txt_prioridad 
         Height          =   390
         Left            =   1350
         TabIndex        =   7
         Top             =   1860
         Width           =   855
      End
      Begin VB.TextBox txt_nombre 
         Height          =   420
         Left            =   1350
         TabIndex        =   4
         Top             =   585
         Width           =   10110
      End
      Begin VB.TextBox txt_clave 
         Height          =   390
         Left            =   1350
         TabIndex        =   3
         Top             =   165
         Width           =   1815
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Establecimiento:"
         Height          =   195
         Left            =   105
         TabIndex        =   16
         Top             =   1140
         Width           =   1155
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   105
         TabIndex        =   15
         Top             =   1575
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Prioridad:"
         Height          =   195
         Left            =   90
         TabIndex        =   14
         Top             =   2055
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   705
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Titular:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   270
         Width           =   480
      End
   End
   Begin VB.CommandButton com_guardar 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   30
      Picture         =   "frmoracle_asignar_clientes_rutas_distribucion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   4815
      Left            =   0
      TabIndex        =   9
      Top             =   2700
      Width           =   17235
      Begin VB.Frame frm_orden 
         BorderStyle     =   0  'None
         Height          =   825
         Left            =   7050
         TabIndex        =   17
         Top             =   1875
         Width           =   2190
         Begin VB.Label lbl_orden 
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   795
            Left            =   30
            TabIndex        =   18
            Top             =   15
            Width           =   2100
         End
      End
      Begin MSComctlLib.ListView lv_clientes 
         Height          =   4320
         Left            =   60
         TabIndex        =   0
         Top             =   450
         Width           =   17115
         _ExtentX        =   30189
         _ExtentY        =   7620
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave Titular"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Titular"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Clave"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Establecimiento"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Dirección"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Prioridad"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Clave Esb."
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Municipio"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Estado"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000FF&
         Caption         =   " Clientes de la ruta"
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   45
         TabIndex        =   10
         Top             =   135
         Width           =   17130
      End
   End
   Begin VB.Frame Frame3 
      Height          =   75
      Left            =   0
      TabIndex        =   8
      Top             =   330
      Width           =   17205
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   16875
      Picture         =   "frmoracle_asignar_clientes_rutas_distribucion.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "frmoracle_asignar_clientes_rutas_distribucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter
Dim objConn As New ADODB.Connection
Dim objCmd As New ADODB.Command
Dim objParm As ADODB.Parameter
Dim var_orden As String
Dim var_posicion As Double
Private Sub cmd_nuevo_Click()
   Me.txt_clave = ""
   Me.txt_nombre = ""
   Me.txt_nombre = ""
End Sub

Private Sub cmd_migrar_oracle_Click()
    rs.Open "delete from XXVIA_TB_CLIENTES_RUTAS_DISTR", cnnoracle_4, adOpenDynamic, adLockOptimistic
    rs.Open "SELECT * FROM TB_ORACLE_CLIENTES_RUTAS_DISTRIBUCION", cnn, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
          rsaux.Open "insert into XXVIA_TB_CLIENTES_RUTAS_DISTR (ruta, titular, nombre_titular, establecimiento, nombre_establecimiento, direccion, prioridad, cn_textilera) values ('" + rs!ruta + "', '" + rs!TITULAR + "', '" + rs!nombre_titular + "', '" + rs!ESTABLECIMIENTO + "', '" + rs!nombre_Establecimiento + "', '" + rs!direccion + "', " + CStr(rs!prioridad) + ", " + CStr(IIf(IsNull(rs!cn_textilera), 0, rs!cn_textilera)) + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
          rsaux.Open "INSERT INTO TB_ORACLE_BITACORA_RUTAS_CLIENTES (ACCION,RUTA,TITULAR,NOMBRE_TITULAR,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,FECHA_ACCION,USUARIO,MAQUINA) VALUES ('INSERTAR'," + rs!ruta + "', '" + rs!TITULAR + "', '" + rs!nombre_titular + "', '" + rs!ESTABLECIMIENTO + "', '" + rs!nombre_Establecimiento + "',GETDATE(),'" + var_clave_usuario_global + "','" + fun_NombrePc + "')", cnn, adOpenDynamic, adLockOptimistic
          rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub com_guardar_Click()
   If Me.lv_clientes.ListItems.Count > 0 Then
      If IsNumeric(Me.txt_prioridad) Then
         rs.Open "update XXVIA_TB_CLIENTES_RUTAS_DISTR set prioridad = " + Me.txt_prioridad + " where establecimiento = '" + Me.txt_establecimiento + "' and ruta = '" + var_ruta_distribucion + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         
         Me.lv_clientes.selectedItem.SubItems(5) = Me.txt_prioridad
         MsgBox "Se a asignado la prioridad", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub Form_Activate()
    Me.lv_clientes.ListItems.Clear
    rs.Open "SELECT * FROM XXVIA_VW_CLIENTES_RUTAS_DISTR where ruta = '" + var_ruta_distribucion + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
          Set list_item = Me.lv_clientes.ListItems.Add(, , rs!TITULAR)
          list_item.SubItems(1) = Format(rs!nombre_titular)
          list_item.SubItems(2) = IIf(IsNull(rs!ESTABLECIMIENTO), "", rs!ESTABLECIMIENTO)
          list_item.SubItems(3) = Format(rs!nombre_Establecimiento)
          list_item.SubItems(4) = IIf(IsNull(rs!direccion), "", rs!direccion)
          list_item.SubItems(5) = IIf(IsNull(rs!prioridad), "", rs!prioridad)
          list_item.SubItems(6) = IIf(IsNull(rs!clave_establecimiento), "", rs!clave_establecimiento)
          
          If rs!TITULAR <> "2040" Then
             strconsulta = "select * from xxvia_vw_clientes_bcp where site_use_id = ?"
             With comandoORA
                  .ActiveConnection = cnnoracle_4
                  .CommandType = adCmdText
                  .CommandText = strconsulta
                  Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, rs!ESTABLECIMIENTO)
                  .Parameters.Append parametro
             End With
             Set rsaux9 = comandoORA.execute
             Set comandoORA = Nothing
             Set parametro = Nothing
             list_item.SubItems(7) = IIf(IsNull(rsaux9!municipio), "", rsaux9!municipio)
             list_item.SubItems(8) = IIf(IsNull(rsaux9!estado), "", rsaux9!estado)
          
             rsaux9.Close
          End If
          rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub Form_Load()
   Top = 200
   Left = 2000
   Me.Caption = "Asignar clientes a ruta " + var_ruta_distribucion + " " + var_nombre_ruta_distribucion
   Me.frm_orden.Visible = False
End Sub

Private Sub lv_clientes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_clientes, ColumnHeader)
End Sub

Private Sub lv_clientes_GotFocus()
   If Me.lv_clientes.ListItems.Count > 0 Then
      Me.txt_clave = Me.lv_clientes.selectedItem
      Me.txt_nombre = Me.lv_clientes.selectedItem.SubItems(1)
      Me.txt_establecimiento = Me.lv_clientes.selectedItem.SubItems(2)
      Me.txt_nombre_establecimiento = Me.lv_clientes.selectedItem.SubItems(3)
      Me.txt_prioridad = Me.lv_clientes.selectedItem.SubItems(5)
      Me.txt_clave_establecimiento = Me.lv_clientes.selectedItem.SubItems(6)
   End If
End Sub

Private Sub lv_clientes_ItemClick(ByVal Item As MSComctlLib.ListItem)
   If Me.lv_clientes.ListItems.Count > 0 Then
      Me.txt_clave = Me.lv_clientes.selectedItem
      Me.txt_nombre = Me.lv_clientes.selectedItem.SubItems(1)
      Me.txt_establecimiento = Me.lv_clientes.selectedItem.SubItems(2)
      Me.txt_nombre_establecimiento = Me.lv_clientes.selectedItem.SubItems(3)
      Me.txt_prioridad = Me.lv_clientes.selectedItem.SubItems(5)
      Me.txt_clave_establecimiento = Me.lv_clientes.selectedItem.SubItems(6)
   End If
End Sub

Private Sub lv_clientes_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      If Me.lv_clientes.ListItems.Count > 0 Then
         var_si = MsgBox("¿Desea eliminar el registro?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            frmautoriza_cambios_distribucion.Show 1
            var_contraseña_cambios_distribucion = "x"
            If var_contraseña_cambios_distribucion <> "" Then
               rs.Open "DELETE FROM XXVIA_TB_CLIENTES_RUTAS_DISTR WHERE ESTABLECIMIENTO = '" + Me.lv_clientes.selectedItem.SubItems(2) + "' and ruta  = '" + var_ruta_distribucion + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               rsaux.Open "INSERT INTO TB_ORACLE_BITACORA_RUTAS_CLIENTES (ACCION,RUTA,TITULAR,NOMBRE_TITULAR,ESTABLECIMIENTO,NOMBRE_ESTABLECIMIENTO,FECHA_ACCION,USUARIO,MAQUINA) VALUES ('ELIMINAR','" + var_ruta_distribucion + "', '" + Me.txt_clave + "', '" + Me.txt_nombre + "', '" + Me.txt_establecimiento + "', '" + Me.txt_nombre_establecimiento + "',GETDATE(),'" + var_clave_usuario_global + "','" + fun_NombrePc + "')", cnn, adOpenDynamic, adLockOptimistic
               var_asunto = "Se notifica que el establecimiento " + Me.txt_establecimiento + " " + Me.txt_nombre_establecimiento + " del titular " + Me.txt_clave + " " + Me.txt_nombre + " a sido eliminado de la ruta " + var_ruta_distribucion + " por el usuario " + var_nombre_usuario_global
               var_cadena = "call xxvia_pk_correo.sp_enviar_email('','fserna@vianney.com.mx','','','Cambio de ruta del establecimiento " + Me.txt_establecimiento + " " + Me.txt_nombre_establecimiento + "','" + var_asunto + "','')"
               rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               lv_clientes.ListItems.Remove (lv_clientes.selectedItem.Index)
               MsgBox "Se a eliminado el registro", vbOKOnly, "ATENCION"
            End If
         End If
      End If
   End If
   If KeyCode = 115 Then
      If Me.txt_prioridad = "0" Then
         Me.txt_prioridad = ""
      End If
      Me.txt_prioridad.SetFocus
   End If
   If KeyCode = 116 Then
      frmoracle_busqueda_clientes_rutas_distribucion.Show 1
   End If
End Sub

Private Sub lv_clientes_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 13
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii > 0 Then
      
      If KeyAscii = 13 Then
         If var_orden <> "" Then
            var_j = Len(Trim(var_orden))
            If var_j <= 3 Then
               Me.lv_clientes.ListItems.Item(var_posicion).Selected = True
               rs.Open "update XXVIA_TB_CLIENTES_RUTAS_DISTR set prioridad = " + Me.lbl_orden + " where establecimiento = '" + Me.lv_clientes.selectedItem.SubItems(2) + "' and ruta = '" + var_ruta_distribucion + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               Me.lv_clientes.selectedItem.SubItems(5) = var_orden
               var_orden = ""
               Me.frm_orden.Visible = False
               Me.lbl_orden = ""
               var_posicion = 0
             Else
                MsgBox "Número incorrecto " + var_orden, vbOKOnly
                Me.lbl_orden = ""
                var_orden = ""
                Me.frm_orden.Visible = False
                var_posicion = 0
             End If
         End If
      Else
         Me.frm_orden.Visible = True
         var_orden = var_orden + Chr(KeyAscii)
         Me.lbl_orden.Caption = var_orden
         If var_posicion = 0 Then
            var_posicion = Me.lv_clientes.selectedItem.Index
            'MsgBox var_posicion
         End If
      End If
   End If

End Sub

Private Sub lv_clientes_LostFocus()
   Me.frm_orden.Visible = False
End Sub

Private Sub txt_clave_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frmoracle_busqueda_clientes_rutas_distribucion.Show 1
   End If
End Sub

Private Sub txt_clave_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_nombre.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_establecimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frmoracle_busqueda_clientes_rutas_distribucion.Show 1
   End If
End Sub

Private Sub txt_establecimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_nombre_establecimiento.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_establecimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frmoracle_busqueda_clientes_rutas_distribucion.Show 1
   End If
End Sub

Private Sub txt_nombre_establecimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_prioridad.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frmoracle_busqueda_clientes_rutas_distribucion.Show 1
   End If
End Sub

Private Sub txt_nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_establecimiento.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_prioridad_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frmoracle_busqueda_clientes_rutas_distribucion.Show 1
   End If
End Sub

Private Sub txt_prioridad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.com_guardar.SetFocus
   End If
End Sub
