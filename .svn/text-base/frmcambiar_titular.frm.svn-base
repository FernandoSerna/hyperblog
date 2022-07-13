VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcambiar_titular 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_tipo 
      Height          =   315
      Left            =   4590
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3000
      Width           =   915
   End
   Begin VB.TextBox txt_clave_establecimiento 
      Height          =   315
      Left            =   2520
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3015
      Width           =   1980
   End
   Begin VB.TextBox txt_clave_cliente 
      Height          =   315
      Left            =   75
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3015
      Width           =   2295
   End
   Begin VB.CommandButton cmd_cancelar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   375
      Picture         =   "frmcambiar_titular.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Cancelar Esc"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   45
      Picture         =   "frmcambiar_titular.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   0
      TabIndex        =   2
      Top             =   255
      Width           =   6240
   End
   Begin MSComctlLib.ListView lv_titulares 
      Height          =   2355
      Left            =   60
      TabIndex        =   3
      Top             =   480
      Width           =   6150
      _ExtentX        =   10848
      _ExtentY        =   4154
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
      NumItems        =   12
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
         Text            =   "pais"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "estado"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ciudad"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "colonia"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "domicilio"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "telefono"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "grupo real"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Limite"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "municipio"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "cp"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmcambiar_titular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_guardar_Click()
   
End Sub

Private Sub cmd_aceptar_Click()
Dim var_si As Integer
   If lv_titulares.ListItems.Count > 0 Then
      If txt_tipo = 1 Then
         var_si = MsgBox("¿Desea cambiarle el titular al cliente?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_si = MsgBox("Confirmar el cambio del titular al cliente", vbYesNo, "ATENCION")
            If var_si = 6 Then
               
               rsaux5.Open "select vcha_emp_conexion from tb_empresas where len(vcha_emp_conexion) > 0", cnn, adOpenDynamic, adLockOptimistic
               While Not rsaux5.EOF
                     var_conexion_importacion = IIf(IsNull(rsaux5(0).Value), "", rsaux5(0).Value)
                     If Trim(var_conexion_importacion) <> "" Then
                        If cnn_importacion.State = 1 Then
                           cnn_importacion.Close
                        End If
                        cnn_importacion.Open var_conexion_importacion
                        rs.Open "UPDATE TB_CLIENTES SET VCHA_TIT_TITULAR_ID = '" + Trim(lv_titulares.selectedItem) + "' WHERE VCHA_CLI_CLAVE_ID = '" + txt_clave_cliente + "'", cnn_importacion, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux5.MoveNext
               Wend
               rsaux5.Close
               
               
            Else
               MsgBox "Se a cancelado el cambio de titular", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Se a cancelado el cambio de titular", vbOKOnly, "ATENCION"
         End If
      End If
      If txt_tipo = 2 Then
         var_si = MsgBox("¿Desea cambiarle el titular al establecimiento?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_si = MsgBox("Confirmar el cambio del titular al establecimiento", vbYesNo, "ATENCION")
            If var_si = 6 Then
               rsaux5.Open "select vcha_emp_conexion from tb_empresas where len(vcha_emp_conexion) > 0", cnn, adOpenDynamic, adLockOptimistic
               While Not rsaux5.EOF
                     var_conexion_importacion = IIf(IsNull(rsaux5(0).Value), "", rsaux5(0).Value)
                     If Trim(var_conexion_importacion) <> "" Then
                        If cnn_importacion.State = 1 Then
                           cnn_importacion.Close
                        End If
                        cnn_importacion.Open var_conexion_importacion
                        rs.Open "UPDATE TB_ESTABLECIMIENTOS SET VCHA_TIT_TITULAR_ID = '" + Trim(lv_titulares.selectedItem) + "' WHERE VCHA_ESB_ESTABLECIMIENTO_ID = '" + txt_clave_establecimiento + "'", cnn_importacion, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux5.MoveNext
               Wend
               rsaux5.Close
               
               
            Else
               MsgBox "Se a cancelado el cambio de titular", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Se a cancelado el cambio de titular", vbOKOnly, "ATENCION"
         End If
      End If
   Else
      MsgBox "No se a seleccionado un titular", vbOKOnly, "ATENCION"
   End If
   Unload Me
End Sub

Private Sub cmd_cancelar_Click()
   Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 65 Then
      cmd_aceptar_Click
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   Dim list_item As ListItem
   rs.Open "select * from tb_titulares", cnn, adOpenDynamic, adLockOptimistic
   numero_items_titulares = 0
   While Not rs.EOF
      Set list_item = lv_titulares.ListItems.Add(, , rs!vcha_tit_titular_id)
      list_item.SubItems(1) = IIf(IsNull(rs!VCHA_tit_NOMBRE), "", rs!VCHA_tit_NOMBRE)
      rs.MoveNext:
      numero_items_titulares = numero_items_titulares + 1
   Wend
   rs.Close
   var_n = lv_titulares.ListItems.Count
   var_numero_renglones = lv_titulares.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_titulares.ColumnHeaders(2).Width = 4400
   Else
      lv_titulares.ColumnHeaders(2).Width = 4600
   End If
End Sub

Private Sub lv_titulares_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_titulares, ColumnHeader)
End Sub
