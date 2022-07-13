VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmlistaunidades 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plantas"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   Icon            =   "frmlistaunidades.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Clave"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Plantas"
         Object.Width           =   9613
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "EMPRESA"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmlistaunidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   var_cadena_seguridad = ""
   var_top = 3772.5 - (frmlistaunidades.Height / 2)
   var_left = 5002.5 - (frmlistaunidades.Height / 2)
   frmlistaunidades.Top = var_top
   frmlistaunidades.Left = var_left
   
   rs.Open "select * from VW_RELACIONES_USUARIOS_UNIDADES where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
      Set list_item = lv_empresas.ListItems.Add(, , rs(1).Value)
      list_item.SubItems(1) = IIf(IsNull(rs(2).Value), "", rs(2).Value)
      list_item.SubItems(2) = IIf(IsNull(rs(3).Value), "", rs(3).Value)
      rs.MoveNext:
      numero_items_usuarios = numero_items_usuarios + 1
   Wend
   rs.Close
End Sub

Private Sub lv_empresas_DblClick()
         If lv_empresas.ListItems.Count > 0 Then
            var_unidad_organizacional = lv_empresas.selectedItem
            var_empresa_global = lv_empresas.selectedItem
            rs.Open "select * from tb_unidadesorganizacionales where vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_responsabilidad_facturacion = IIf(IsNull(rs!RESPONSABILIDAD_FACTURACION), "", rs!RESPONSABILIDAD_FACTURACION)
            End If
            rs.Close
            Unload Me
            var_top = 3772.5 - (frmmenu1.Height / 2)
            var_left = 5002.5 - (frmmenu1.Height / 2)
            frmmenu1.Top = var_top
            frmmenu1.Left = var_left
            'Call ExplodeForm(frmmenu1, 800)
            frmmenu1.Label1.Caption = "BIENVENIDO AL SID " + Trim(var_nombre_usuario_global)
            'frmmenu1.Show
            If var_ejecutar_programa = 1 Then
On Error GoTo sigue
               rs.Open "select * from tb_oracle_maquinas where maquina = '" + fun_NombrePc + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  FileCopy "\\fscdindustrial\fscdind\update_sip\MovimientosInventario\RespaldoCatalogos.exe", App.Path + "\RespaldoCatalogos.exe"
                  x = Shell(App.Path + "\RespaldoCatalogos.exe " + Str(var_unidad_organizacional) + "|pvia")
               End If
               rs.Close
sigue:
            End If
         Else
            MsgBox "No existen plantas asignadas a este usuario", vbOKOnly, "ATENCION"
         End If
End Sub

Private Sub lv_empresas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNull(lv_empresas.selectedItem) Then
         MsgBox "No se a seleccionado ninguna empresa", vbOKOnly, "ATENCION"
      Else
         If lv_empresas.ListItems.Count > 0 Then
            var_unidad_organizacional = lv_empresas.selectedItem
            var_empresa_global = lv_empresas.selectedItem
            
            'rs.Open "SELECT * FROM TB_EMPRESAS WHERE VCHA_eMP_EMPRESA_ID = '" + Me.lv_empresas.selectedItem.SubItems(2) + "'", cnn, adOpenDynamic, adLockOptimistic
            'If Not rs.EOF Then
            '   VAR_SERVIDOR = IIf(IsNull(rs!VCHA_eMP_SERVIDOR), "", rs!VCHA_eMP_SERVIDOR)
            '   VAR_BASE_DATOS = IIf(IsNull(rs!VCHA_EMP_BASE_DE_dATOS), "", rs!VCHA_EMP_BASE_DE_dATOS)
            '   If VAR_SERVIDOR <> "" And VAR_BASE_DATOS <> "" Then
            '      parametros(0) = VAR_SERVIDOR
            '      parametros(1) = VAR_BASE_DATOS
            '      var_conexion_string = "Provider=SQLOLEDB.1;Password=" & parametros(3) & ";Persist Security Info=True;User ID=sa;Initial Catalog=" & parametros(1) & ";Data Source=" & parametros(0)

            '      cnn.Close
            '      cnn.Open var_conexion_string
            '   Else
            '      rs.Close
            '   End If
            'Else
            '   rs.Close
            'End If
            
            Unload Me
            var_top = 3772.5 - (frmmenu1.Height / 2)
            var_left = 5002.5 - (frmmenu1.Height / 2)
            frmmenu1.Top = var_top
            frmmenu1.Left = var_left
            'Call ExplodeForm(frmmenu1, 800)
            frmmenu1.Label1.Caption = "BIENVENIDO AL SID " + Trim(var_nombre_usuario_global)
            'frmmenu1.Show
            If var_ejecutar_programa = 1 Then
On Error GoTo sigue
               rs.Open "select * from tb_oracle_maquinas where maquina = '" + fun_NombrePc + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  FileCopy "\\fscdindustrial\fscdind\update_sip\MovimientosInventario\RespaldoCatalogos.exe", App.Path + "\RespaldoCatalogos.exe"
                  x = Shell(App.Path + "\RespaldoCatalogos.exe " + Str(var_unidad_organizacional) + "|pvia")
               End If
               rs.Close
sigue:
            End If
         Else
            MsgBox "No existen plantas asignadas a este usuario", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub
