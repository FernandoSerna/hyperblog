VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_notas_traspasos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Notas de traspasos"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lv_lista 
      Height          =   3525
      Left            =   30
      TabIndex        =   0
      Top             =   60
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Clave"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Retraso"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "almacen origen"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "almacen destino"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmoracle_notas_traspasos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   rs.Open "ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'"
   rs.Open "seLECT * FROM XXVIA_VW_TRANSITO_ENC_SUB WHERE ORGANIZACION = " + var_unidad_organizacional + " and ALMACENORIGENID = '" + var_almacen_destino_traspaso + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!folioenvio)
            list_item.SubItems(1) = IIf(IsNull(rs!almacenorigen), "", rs!almacenorigen)
            list_item.SubItems(2) = IIf(IsNull(rs!fechaenvio), "", rs!fechaenvio)
            list_item.SubItems(3) = IIf(IsNull(rs!diasretraso), "", rs!diasretraso)
            list_item.SubItems(4) = IIf(IsNull(rs!almacenorigenid), "", rs!almacenorigenid)
            list_item.SubItems(5) = IIf(IsNull(rs!almacendestino), "", rs!almacendestino)
            rs.MoveNext
      Wend
   Else
      MsgBox "No existen traspasos para esta organización", vbOKOnly, "ATENCION"
   End If
   rs.Close
End Sub

Private Sub lv_lista_DblClick()
      var_almacen_origen_traspaso = Me.lv_lista.selectedItem.SubItems(4)
      var_almacen_destino_traspaso = Me.lv_lista.selectedItem.SubItems(5)
      var_numero_nota_traspaso_n = Me.lv_lista.selectedItem
      frmoracle_traspasos.Show
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call lv_lista_DblClick
   End If
End Sub
