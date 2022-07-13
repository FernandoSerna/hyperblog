VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmplantas_transito 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Plantas"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lv_lista 
      Height          =   2115
      Left            =   15
      TabIndex        =   0
      Top             =   60
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   3731
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
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   5645
      EndProperty
   End
End
Attribute VB_Name = "frmplantas_transito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Dim list_item As ListItem
   rsaux1.Open "select distinct VCHA_PLA_PLANTA_ID, vcha_pla_descripc from tb_plantas where vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn_admcdindustrial, adOpenDynamic, adLockOptimistic
   numero_items_lineas = 0
   While Not rsaux1.EOF
      Set list_item = Me.lv_lista.ListItems.Add(, , rsaux1(0).Value)
      list_item.SubItems(1) = IIf(IsNull(rsaux1(1).Value), "", rsaux1(1).Value)
      rsaux1.MoveNext:
    Wend
    rsaux1.Close
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      var_planta_transito_global = Me.lv_lista.selectedItem
      frmnotas_traspasos_plantas.var_str_encabezado_forma = "Entradas Por traspaso Entre Plantas"
      frmnotas_traspasos_plantas.Show 1
      Unload Me
   End If
End Sub
