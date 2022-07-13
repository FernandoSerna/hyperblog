VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_subinventario_traspasos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Almacenes para traspasos"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lv_lista 
      Height          =   3525
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   6218
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Clave"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   10760
      EndProperty
   End
End
Attribute VB_Name = "frmoracle_subinventario_traspasos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
   'rs.Close
   rs.Open "ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'"
   rs.Open "seLECT DISTINCT ALMACENORIGENID FROM XXVIA_VW_TRANSITO_ENC_SUB WHERE ORGANIZACION = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            rsaux.Open "select description from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional + " AND secondary_inventory_name = '" + rs(0).Value + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            Set list_item = lv_lista.ListItems.Add(, , rs(0).Value)
            list_item.SubItems(1) = IIf(IsNull(rsaux(0).Value), "", rsaux(0).Value)
            rsaux.Close
            rs.MoveNext
      Wend
   Else
      MsgBox "No existen traspasos para esta organización", vbOKOnly, "ATENCION"
   End If
   rs.Close
End Sub

Private Sub lv_lista_DblClick()
   If Me.lv_lista.ListItems.Count > 0 Then
      var_almacen_destino_traspaso = Me.lv_lista.selectedItem
      frmoracle_notas_traspasos.Show
   End If
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      var_almacen_destino_traspaso = Me.lv_lista.selectedItem
      frmoracle_notas_traspasos.Show
   End If
End Sub
