VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmreporte_aplicacion_depositos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aplicación de depositos"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7320
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmreporte_aplicacion_depositos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Movimiento Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11235
      Picture         =   "frmreporte_aplicacion_depositos.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   90
      Left            =   0
      TabIndex        =   8
      Top             =   315
      Width           =   11685
   End
   Begin VB.Frame Frame2 
      Caption         =   " Aplicaciones"
      Height          =   2970
      Left            =   75
      TabIndex        =   6
      Top             =   4185
      Width           =   11490
      Begin MSComctlLib.ListView lv_aplicado 
         Height          =   2265
         Left            =   105
         TabIndex        =   4
         Top             =   180
         Width           =   11265
         _ExtentX        =   19870
         _ExtentY        =   3995
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Folio"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha relación"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Agente"
            Object.Width           =   6527
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "cliente"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Serie"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Documento"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Número"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Importe"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label lbl_total 
         Alignment       =   1  'Right Justify
         Caption         =   "999,000.000.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   9165
         TabIndex        =   10
         Top             =   2490
         Width           =   2130
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Aplicado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7320
         TabIndex        =   9
         Top             =   2542
         Width           =   1740
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Depositos "
      Height          =   3690
      Left            =   75
      TabIndex        =   5
      Top             =   480
      Width           =   11490
      Begin VB.TextBox txt_importe 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3195
         TabIndex        =   2
         Top             =   180
         Width           =   2475
      End
      Begin MSComctlLib.ListView lv_depositos 
         Height          =   2925
         Left            =   90
         TabIndex        =   3
         Top             =   675
         Width           =   11280
         _ExtentX        =   19897
         _ExtentY        =   5159
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha autorización"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Folio autorización"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Referencia"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Ficha"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "F. operacion"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Canal"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Tienda"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Descripcion"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Importe"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Cuenta"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Origen"
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Importe del deposito:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   150
         TabIndex        =   7
         Top             =   210
         Width           =   2970
      End
   End
End
Attribute VB_Name = "frmreporte_aplicacion_depositos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_nuevo_Click()
   Me.lv_aplicado.ListItems.Clear
   Me.lv_depositos.ListItems.Clear
   Me.txt_importe = ""
   Me.txt_importe.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 0
   lbl_total = "0.00"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   activa_forma (var_activa_forma_packing_list)
End Sub

Private Sub lv_depositos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_depositos, ColumnHeader)
End Sub

Private Sub lv_depositos_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Me.lv_aplicado.ListItems.Clear
End Sub

Private Sub lv_depositos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.lv_aplicado.ListItems.Clear
      If Me.lv_depositos.ListItems.Count > 0 Then
         var_cadena = "select rco.vcha_rco_folio as RC_Folio, rco.dtim_rco_fecha_relacion as Fecha_relaciOn, age.vcha_age_nombre as Agente, cli.vcha_cli_nombre as Cliente, rco.vcha_ser_serie_id as Serie, rco.vcha_car_documento as Documento, rco.inte_car_numero as Numero, rco.floa_RCO_IMPORTE as Importe from tb_relacion_cobranza rco, tb_agentes age, tb_clientes cli where rco.vcha_age_agente_id = age.vcha_age_agente_id and rco.vcha_cli_clave_id = cli.vcha_cli_clave_id and rco.inte_rco_numero_deposito = " + Me.lv_depositos.selectedItem.SubItems(1)
         Me.lbl_total.Caption = "0.00"
         rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            While Not rs.EOF
                  Set list_item = Me.lv_aplicado.ListItems.Add(, , rs(0).Value)
                  list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
                  list_item.SubItems(2) = IIf(IsNull(rs(2).Value), "", rs(2).Value)
                  list_item.SubItems(3) = IIf(IsNull(rs(3).Value), "", rs(3).Value)
                  list_item.SubItems(4) = IIf(IsNull(rs(4).Value), "", rs(4).Value)
                  list_item.SubItems(5) = IIf(IsNull(rs(5).Value), "", rs(5).Value)
                  list_item.SubItems(6) = IIf(IsNull(rs(6).Value), "", rs(6).Value)
                  list_item.SubItems(7) = Format(IIf(IsNull(rs(7).Value), 0, rs(7).Value), "###,###,##0.00")
                  Me.lbl_total = Me.lbl_total + IIf(IsNull(rs(7).Value), 0, rs(7).Value)
                  rs.MoveNext:
                  numero_items_lineas = numero_items_lineas + 1
            Wend
         Else
            MsgBox "El deposito no esta asignado", vbOKOnly, "ATENCION"
         End If
         rs.Close
         Me.lbl_total = Format(Me.lbl_total, "###,###,##0.00")
      End If
   End If
End Sub

Private Sub txt_importe_Change()
   Me.lv_aplicado.ListItems.Clear
   Me.lv_depositos.ListItems.Clear
   lbl_total = "0.00"
End Sub

Private Sub txt_importe_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_importe) Then
         Me.txt_importe = CDbl(Me.txt_importe)
         rs.Open "SELECT FECHA_AUTORIZACION, FOLIO_AUTORIZACION, REFERENCIA, FICHA, FECHA_OPERACION, CANAL, TIENDA, DESCRIPCION, IMPORTE, CUENTA, ORIGEN FROM VW_MOVIMIENTOS_CONSULTA WHERE IMPORTE = " + Me.txt_importe, cnnoracle_2, adOpenDynamic, adLockOptimistic
         Me.lv_depositos.ListItems.Clear
         While Not rs.EOF
               Set list_item = Me.lv_depositos.ListItems.Add(, , rs(0).Value)
               list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
               list_item.SubItems(2) = IIf(IsNull(rs(2).Value), "", rs(2).Value)
               list_item.SubItems(3) = IIf(IsNull(rs(3).Value), "", rs(3).Value)
               list_item.SubItems(4) = IIf(IsNull(rs(4).Value), "", rs(4).Value)
               list_item.SubItems(5) = IIf(IsNull(rs(5).Value), "", rs(5).Value)
               list_item.SubItems(6) = IIf(IsNull(rs(6).Value), "", rs(6).Value)
               list_item.SubItems(7) = IIf(IsNull(rs(7).Value), "", rs(7).Value)
               list_item.SubItems(8) = IIf(IsNull(rs(8).Value), "", rs(8).Value)
               list_item.SubItems(9) = IIf(IsNull(rs(9).Value), "", rs(9).Value)
               list_item.SubItems(10) = IIf(IsNull(rs(10).Value), "", rs(10).Value)
               rs.MoveNext:
               numero_items_lineas = numero_items_lineas + 1
         Wend

         rs.Close
      Else
         MsgBox "Importe incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
End Sub
