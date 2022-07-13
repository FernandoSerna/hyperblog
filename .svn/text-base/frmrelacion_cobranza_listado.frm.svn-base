VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmrelacion_cobranza_listado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relaciones de Cobranza No aplicadas"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   7275
   Begin VB.Frame frm_cargar_relacion 
      Height          =   915
      Left            =   2265
      TabIndex        =   1
      Top             =   1080
      Width           =   2610
      Begin VB.TextBox txt_folio 
         Height          =   315
         Left            =   105
         TabIndex        =   3
         Top             =   435
         Width           =   2355
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Folio de la Relación de Cobranza"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   2
         Top             =   15
         Width           =   2595
      End
   End
   Begin MSComctlLib.ListView lv_relaciones_cobranza 
      Height          =   3735
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   6588
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
         Text            =   "Folio"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Agente"
         Object.Width           =   7937
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha"
         Object.Width           =   2117
      EndProperty
   End
End
Attribute VB_Name = "frmrelacion_cobranza_listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frm_cargar_relacion.Visible = True
      txt_folio.SetFocus
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Dim var_fecha_cheques As String
   var_dia = CStr(Day(Date))
   var_mes = CStr(Month(Date))
   var_año = CStr(Year(Date))
   If Len(Trim(var_dia)) = 1 Then
      var_dia = "0" + var_dia
   End If
   If Len(Trim(var_mes)) = 1 Then
      var_mes = "0" + var_mes
   End If
   var_fecha_cheques = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"

Top = 2000
Left = 2100
frm_cargar_relacion.Visible = False
Dim list_item As ListItem
   'rs.Open "select distinct vcha_rco_folio, vcha_age_nombre, substring(cast(dtim_rco_fecha_relacion as varchar(50)),1,10) as dtim_rco_fecha_relacion from vw_relaciones_cobranza_no_aplicadas where vcha_emp_empresa_id = '" + var_empresa + "' and  dtim_rco_fecha_cheque <= " + var_fecha_cheques + " +1 -.000001 and vcha_Rco_estatus = 'I'", cnn, adOpenDynamic, adLockOptimistic
   rs.Open "select distinct vcha_rco_folio, vcha_age_nombre, dtim_rco_fecha_relacion from vw_relaciones_cobranza_no_aplicadas where vcha_emp_empresa_id = '" + var_empresa + "' and  dtim_rco_fecha_cheque <= " + var_fecha_cheques + " +1 -.000001 and vcha_Rco_estatus = 'I'", cnn, adOpenDynamic, adLockOptimistic
   numero_items_listadeprecios = 0
   If Not rs.EOF Then
      txt_folio = rs!vcha_Rco_folio
      While Not rs.EOF
          Set list_item = lv_relaciones_cobranza.ListItems.Add(, , rs!vcha_Rco_folio)
          list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
          'list_item.SubItems(2) = Format(IIf(IsNull(rs!dtim_rco_fecha_relacion), "", rs!dtim_rco_fecha_relacion), "Short Date")
          'var_x = CDate(rs!dtim_rco_fecha_relacion)
          list_item.SubItems(2) = Format(CDate(IIf(IsNull(rs!dtim_rco_fecha_relacion), "", rs!dtim_rco_fecha_relacion)), "Short Date")
          rs.MoveNext:
          numero_items_listadeprecios = numero_items_listadeprecios + 1
      Wend
   var_numero_renglones = lv_relaciones_cobranza.Height / 312.5
   var_n = Me.lv_relaciones_cobranza.ListItems.Count
   If var_n > var_numero_renglones Then
      lv_relaciones_cobranza.ColumnHeaders(2).Width = 4250.71
   Else
      lv_relaciones_cobranza.ColumnHeaders(2).Width = 4499.71
   End If
   End If
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_activa_menu = True Then
      Frmmenu2.Enabled = True
   End If
End Sub

Private Sub lv_relaciones_cobranza_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_relaciones_cobranza, ColumnHeader)
End Sub

Private Sub lv_relaciones_cobranza_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para buscar relaciones anteriores"
End Sub

Private Sub lv_relaciones_cobranza_ItemClick(ByVal Item As MSComctlLib.ListItem)
   txt_folio = lv_relaciones_cobranza.selectedItem
End Sub

Private Sub lv_relaciones_cobranza_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      var_activa_forma_relacion_cobranza = Me.Name
      frmrelacion_cobranza.Show 1
   End If
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub lv_relaciones_cobranza_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_folio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Trim(txt_folio) <> "" Then
         If rs.State = 1 Then
            rs.Close
         End If
         rs.Open "select * from tb_relacion_cobranza with (nolock) where vcha_rco_folio = '" + txt_folio + "' and vcha_emp_empresa_id = '" + var_empresa + "' and vcha_rco_estatus = 'I'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            rs.Close
            frm_cargar_relacion.Visible = False
            var_activa_forma_relacion_cobranza = Me.Name
            frmrelacion_cobranza.Show 1
         Else
            rs.Close
            MsgBox "Folio de relación de cobranza incorrecto", vbOKOnly, "ATENCION"
            frm_cargar_relacion.Visible = False
         End If
      End If
   End If
   If KeyAscii = 27 Then
      frm_cargar_relacion.Visible = False
   End If
End Sub
