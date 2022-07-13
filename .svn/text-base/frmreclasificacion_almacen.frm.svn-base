VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmreclasificacion_almacen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Almacen para reclasificación"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lv_lista 
      Height          =   2535
      Left            =   60
      TabIndex        =   0
      Top             =   15
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4471
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
End
Attribute VB_Name = "frmreclasificacion_almacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Top = 2500
   Left = 2500
   Dim list_item As ListItem
   rs.Open "select * from tb_almacenes where vcha_Emp_empresa_id = '" + var_empresa + "' AND INTE_ALM_RECLASIFICACION  =  1 order by vcha_alm_nombre", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set list_item = lv_lista.ListItems.Add(, , rs!vcha_alm_almacen_id)
         list_item.SubItems(1) = IIf(IsNull(rs!vcha_alm_nombre), "", rs!vcha_alm_nombre)
         rs.MoveNext:
         numero_items_MUNICIPIOS = numero_items_MUNICIPIOS + 1
   Wend
   rs.Close
   If lv_lista.ListItems.Count > 11 Then
      lv_lista.ColumnHeaders(2).Width = 4550
   Else
      lv_lista.ColumnHeaders(2).Width = 4790
   End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub lv_lista_DblClick()
   If lv_lista.ListItems.Count > 0 Then
      frmreclasificacion.txt_almacen = lv_lista.selectedItem
      frmreclasificacion.Show
      var_activa_forma_reporte_valuacion_devoluciones = "frmreclasificacion_almacen"
      Me.Enabled = False
   End If
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call lv_lista_DblClick
   End If
   If KeyAscii = 27 Then
     Unload Me
   End If
End Sub
