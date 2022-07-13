VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_subinventarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Almacenes"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lv_almacenes 
      Height          =   2340
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   4128
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
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "marca"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmoracle_subinventarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   var_cadena_almacenes = ""
   If var_clave_usuario_global = "U0000000680" Then
      rsaux.Open "select secondary_inventory_name, description from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional + " and secondary_inventory_name = 'PRIVALIA'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      While Not rsaux.EOF
            Set list_item = Me.lv_almacenes.ListItems.Add(, , rsaux(0).Value)
            list_item.SubItems(1) = IIf(IsNull(rsaux(1).Value), "", rsaux(1).Value)
            rsaux.MoveNext
      Wend
      rsaux.Close
   Else
      rs.Open "select * from tb_permisos_movimientos where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            If var_cadena_almacenes = "" Then
               var_cadena_almacenes = "'" + IIf(IsNull(rs!vcha_per_almacen_1), "", rs!vcha_per_almacen_1) + "'"
            Else
               var_cadena_almacenes = var_cadena_almacenes + ",'" + IIf(IsNull(rs!vcha_per_almacen_1), "", rs!vcha_per_almacen_1) + "'"
            End If
            rs.MoveNext
      Wend
      rs.Close
      If var_clave_movimiento <> "VENDOR" Then
         'rsaux.Open "select secondary_inventory_name, description from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional + " and secondary_inventory_name in (" + var_cadena_almacenes + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
         'var_unidad_organizacional = 85
         rsaux.Open "select secondary_inventory_name, description from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
      Else
         rsaux.Open "select secondary_inventory_name, description from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
         'rsaux.Open "select secondary_inventory_name, description from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional + " and secondary_inventory_name = 'CDI_ALMPT'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      End If
      If Not rsaux.EOF Then
         While Not rsaux.EOF
               Set list_item = Me.lv_almacenes.ListItems.Add(, , rsaux(0).Value)
               list_item.SubItems(1) = IIf(IsNull(rsaux(1).Value), "", rsaux(1).Value)
               rsaux.MoveNext
         Wend
      Else
         MsgBox "No existen traspasos para esta organización", vbOKOnly, "ATENCION"
      End If
      rsaux.Close
   End If
End Sub

Private Sub lv_almacenes_DblClick()
   If Me.lv_almacenes.ListItems.Count > 0 Then
      var_almacen_global = Me.lv_almacenes.selectedItem
      var_nombre_almacen_global = Me.lv_almacenes.selectedItem.SubItems(1)
      Unload Me
   End If
End Sub

Private Sub lv_almacenes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call lv_almacenes_DblClick
   End If
End Sub
