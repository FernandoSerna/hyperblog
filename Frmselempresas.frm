VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frmempresas 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empresas"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   Icon            =   "Frmselempresas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5730
   Begin MSComctlLib.ListView lv_empresas 
      Height          =   3285
      Left            =   90
      TabIndex        =   0
      Top             =   15
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   5794
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Clave"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Empresas"
         Object.Width           =   9613
      EndProperty
   End
   Begin MSComctlLib.ImageList imgSmall 
      Left            =   1440
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmselempresas.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmselempresas.frx":0E64
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   840
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
            Picture         =   "Frmselempresas.frx":0FBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmselempresas.frx":1898
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Frmempresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
   var_cadena_seguridad = ""
   var_top = 3772.5 - (Frmempresas.Height / 2)
   var_left = 5002.5 - (Frmempresas.Height / 2)
   Frmempresas.Top = var_top
   Frmempresas.Left = var_left
   If var_clave_usuario_global = "1" Then
      rs.Open "select * from Tb_empresas", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
         Set list_item = lv_empresas.ListItems.Add(, , rs(0).Value)
         list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
         rs.MoveNext:
         numero_items_usuarios = numero_items_usuarios + 1
      Wend
      rs.Close
   Else
      rs.Open "select * from vw_relacion_usuarioS_empresas with (nolock) where vcha_usu_usuario_id = '" + var_clave_usuario_global + "'", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
         Set list_item = lv_empresas.ListItems.Add(, , rs(0).Value)
         list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
         rs.MoveNext:
         numero_items_usuarios = numero_items_usuarios + 1
      Wend
      rs.Close
   End If
End Sub



Private Sub lv_empresas_DblClick()
   If lv_empresas.ListItems.Count > 0 Then
      If IsNull(lv_empresas.selectedItem) Then
         MsgBox "No se a seleccionado ninguna empresa", vbOKOnly, "ATENCION"
      Else
         var_empresa = lv_empresas.selectedItem
         
         var_nombre_empresa = lv_empresas.selectedItem.SubItems(1)
         var_x = ""
         For var_j = 1 To Len(var_nombre_empresa)
            If Mid(var_nombre_empresa, var_j, 1) <> "." Then
               If Mid(var_nombre_empresa, var_j, 1) <> "'" Then
                  If Mid(var_nombre_empresa, var_j, 1) <> "," Then
                     If Mid(var_nombre_empresa, var_j, 1) = " " Then
                        var_x = var_x + "_"
                     Else
                        var_x = var_x + Mid(var_nombre_empresa, var_j, 1)
                     End If
                  End If
               End If
            End If
         Next var_j
         var_nombre_empresa = var_x
         frmsistema_integral.Caption = frmsistema_integral.Caption + "  " + Me.lv_empresas.selectedItem.SubItems(1)
         Unload Me
         frmlistaunidades.Show
      End If
   End If
End Sub

Private Sub lv_empresas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_empresas.ListItems.Count > 0 Then
         If IsNull(lv_empresas.selectedItem) Then
            MsgBox "No se a seleccionado ninguna empresa", vbOKOnly, "ATENCION"
         Else
            var_empresa = lv_empresas.selectedItem
            var_nombre_empresa = lv_empresas.selectedItem.SubItems(1)
            For var_j = 1 To Len(var_nombre_empresa)
               If Mid(var_nombre_empresa, var_j, 1) <> "." Then
                  If Mid(var_nombre_empresa, var_j, 1) <> "'" Then
                     If Mid(var_nombre_empresa, var_j, 1) <> "," Then
                        If Mid(var_nombre_empresa, var_j, 1) = " " Then
                           var_x = var_x + "_"
                        Else
                           var_x = var_x + Mid(var_nombre_empresa, var_j, 1)
                        End If
                     End If
                  End If
               End If
            Next var_j
            var_nombre_empresa = var_x
            frmsistema_integral.Caption = frmsistema_integral.Caption + "  " + Me.lv_empresas.selectedItem.SubItems(1)
            Unload Me
            frmlistaunidades.Show
         End If
      End If
   End If
End Sub
