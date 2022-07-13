VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmbloquear_articulos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bloquear articulos"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Desbloquear articulos"
      Height          =   930
      Left            =   5910
      TabIndex        =   2
      Top             =   90
      Width           =   5580
   End
   Begin VB.CommandButton cmd_bloquear 
      Caption         =   "Bloquear articulos"
      Height          =   930
      Left            =   105
      TabIndex        =   0
      Top             =   90
      Width           =   5580
   End
   Begin MSComctlLib.ListView lv_articulos 
      Height          =   3975
      Left            =   105
      TabIndex        =   1
      Top             =   1140
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   7011
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   6350
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Detenido"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.ListView lv_desbloqueados 
      Height          =   3975
      Left            =   5910
      TabIndex        =   3
      Top             =   1140
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   7011
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   6350
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Detenido"
         Object.Width           =   1764
      EndProperty
   End
End
Attribute VB_Name = "frmbloquear_articulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_bloquear_Click()
   var_si = MsgBox("¿Bloquear los articulos?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      var_si = MsgBox("Confirmar el bloqueo de los articulos", vbYesNo, "ATENCION")
      If var_si = 6 Then
         rs.Open "select * from pumas", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               rsaux2.Open "update tb_Articulos set INTE_ART_DETENIDO = 1 where vcha_Art_Articulo_id = '" + rs!codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               rs.MoveNext
         Wend
         rs.Close
         lv_articulos.ListItems.Clear
         rs.Open "SELECT dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL, dbo.TB_ARTICULOS.INTE_ART_DETENIDO FROM  dbo.pumas INNER JOIN  dbo.TB_ARTICULOS ON dbo.pumas.codigo = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID"
         While Not rs.EOF
               Set list_item = Me.lv_articulos.ListItems.Add(, , rs!vcha_Art_articulo_id)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
               list_item.SubItems(2) = IIf(IsNull(rs!INTE_ART_detenido), 0, rs!INTE_ART_detenido)
               rs.MoveNext
         Wend
         rs.Close
      End If
   End If
End Sub

Private Sub Command1_Click()
   var_si = MsgBox("¿Desbloquear los articulos?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      var_si = MsgBox("Confirmar el desbloqueo de los articulos", vbYesNo, "ATENCION")
      If var_si = 6 Then
         rs.Open "select * from codigos_shrek", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               rsaux2.Open "update tb_Articulos set INTE_ART_DETENIDO = 0 where vcha_Art_Articulo_id = '" + rs!codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               rs.MoveNext
         Wend
         rs.Close
         lv_desbloqueados.ListItems.Clear
         rs.Open "SELECT dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL, dbo.TB_ARTICULOS.INTE_ART_DETENIDO FROM         dbo.TB_ARTICULOS INNER JOIN dbo.codigos_shrek ON dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID = dbo.codigos_shrek.codigo ", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = Me.lv_desbloqueados.ListItems.Add(, , rs!vcha_Art_articulo_id)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
               list_item.SubItems(2) = IIf(IsNull(rs!INTE_ART_detenido), 0, rs!INTE_ART_detenido)
               rs.MoveNext
         Wend
         rs.Close
      End If
   End If
End Sub

Private Sub Form_Load()
   Top = 1500
   Left = 0
   rs.Open "SELECT dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL, dbo.TB_ARTICULOS.INTE_ART_DETENIDO FROM  dbo.pumas INNER JOIN  dbo.TB_ARTICULOS ON dbo.pumas.codigo = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID"
   While Not rs.EOF
      Set list_item = Me.lv_articulos.ListItems.Add(, , rs!vcha_Art_articulo_id)
      list_item.SubItems(1) = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
      list_item.SubItems(2) = IIf(IsNull(rs!INTE_ART_detenido), 0, rs!INTE_ART_detenido)
      rs.MoveNext
    Wend
    rs.Close
    rs.Open "SELECT     dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL, dbo.TB_ARTICULOS.INTE_ART_DETENIDO FROM         dbo.TB_ARTICULOS INNER JOIN dbo.codigos_shrek ON dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID = dbo.codigos_shrek.codigo ", cnn, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
       Set list_item = Me.lv_desbloqueados.ListItems.Add(, , rs!vcha_Art_articulo_id)
       list_item.SubItems(1) = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
       list_item.SubItems(2) = IIf(IsNull(rs!INTE_ART_detenido), 0, rs!INTE_ART_detenido)
       rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub
