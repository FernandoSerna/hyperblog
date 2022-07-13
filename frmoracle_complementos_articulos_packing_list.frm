VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_complementos_articulos_packing_list 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Complementos de artículos para packing list de exportaciones"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_exportar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1980
      Picture         =   "frmoracle_complementos_articulos_packing_list.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Exportar a excel"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_complementos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   990
      Picture         =   "frmoracle_complementos_articulos_packing_list.frx":0312
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Complementos"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame frm_origenes 
      Height          =   2820
      Left            =   2625
      TabIndex        =   25
      Top             =   3465
      Width           =   4485
      Begin MSComctlLib.ListView lv_origenes 
         Height          =   2370
         Left            =   30
         TabIndex        =   26
         Top             =   405
         Width           =   4410
         _ExtentX        =   7779
         _ExtentY        =   4180
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   " Origenes"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   7
         Left            =   30
         TabIndex        =   27
         Top             =   135
         Width           =   4410
      End
   End
   Begin VB.CommandButton cmd_carga_mmasiva 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1650
      Picture         =   "frmoracle_complementos_articulos_packing_list.frx":0414
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Carga masiva"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   660
      Picture         =   "frmoracle_complementos_articulos_packing_list.frx":0516
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   4215
      Left            =   60
      TabIndex        =   23
      Top             =   3450
      Width           =   9570
      Begin MSComctlLib.ListView lv_complementos 
         Height          =   3975
         Left            =   90
         TabIndex        =   16
         Top             =   165
         Width           =   9360
         _ExtentX        =   16510
         _ExtentY        =   7011
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
         NumItems        =   15
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   12965
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fracción A."
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Composición"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Contenido"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Origen"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Aplica USA"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Aplica CA"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Hecho_en"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Fraccion Americana"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Criterio Origen USA"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Criterio Origen CA"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "APLICA AP"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "FOLIO_COLOMBIA"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "COMPLEMENTO"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3030
      Left            =   75
      TabIndex        =   17
      Top             =   405
      Width           =   9540
      Begin VB.TextBox txt_folio 
         Height          =   315
         Left            =   4680
         MaxLength       =   50
         TabIndex        =   42
         Top             =   2130
         Width           =   1695
      End
      Begin VB.TextBox txt_complemento 
         Height          =   315
         Left            =   7440
         MaxLength       =   50
         TabIndex        =   40
         Top             =   2130
         Width           =   1815
      End
      Begin VB.ComboBox cmb_aplica_AP 
         Height          =   315
         ItemData        =   "frmoracle_complementos_articulos_packing_list.frx":05E8
         Left            =   7740
         List            =   "frmoracle_complementos_articulos_packing_list.frx":05F5
         TabIndex        =   37
         Top             =   1680
         Width           =   960
      End
      Begin VB.TextBox txt_criterio_origen_CA 
         Height          =   315
         Left            =   6315
         TabIndex        =   36
         Top             =   2595
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.TextBox txt_criterio_origen_USA 
         Height          =   315
         Left            =   1545
         TabIndex        =   34
         Top             =   2595
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.TextBox txt_fraccion_americana 
         Height          =   315
         Left            =   1695
         MaxLength       =   50
         TabIndex        =   31
         Top             =   2130
         Width           =   1815
      End
      Begin VB.TextBox txt_hecho_en 
         Height          =   315
         Left            =   5505
         MaxLength       =   50
         TabIndex        =   13
         Top             =   1290
         Width           =   3900
      End
      Begin VB.ComboBox cmb_aplica_ca 
         Height          =   315
         ItemData        =   "frmoracle_complementos_articulos_packing_list.frx":0604
         Left            =   4620
         List            =   "frmoracle_complementos_articulos_packing_list.frx":0611
         TabIndex        =   15
         Top             =   1680
         Width           =   960
      End
      Begin VB.ComboBox cmb_aplica_usa 
         Height          =   315
         ItemData        =   "frmoracle_complementos_articulos_packing_list.frx":0620
         Left            =   1290
         List            =   "frmoracle_complementos_articulos_packing_list.frx":062D
         TabIndex        =   14
         Top             =   1680
         Width           =   960
      End
      Begin VB.TextBox txt_codigo 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   7
         Top             =   255
         Width           =   1170
      End
      Begin VB.TextBox txt_descripcion 
         Height          =   315
         Left            =   2490
         MaxLength       =   50
         TabIndex        =   8
         Top             =   255
         Width           =   6930
      End
      Begin VB.TextBox txt_fraccion 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   9
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txt_composicion 
         Height          =   315
         Left            =   4560
         MaxLength       =   50
         TabIndex        =   10
         Top             =   600
         Width           =   4860
      End
      Begin VB.TextBox txt_contenido 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   11
         Top             =   960
         Width           =   8130
      End
      Begin VB.TextBox txt_origen 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   12
         Top             =   1320
         Width           =   3330
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Complemento:"
         Height          =   195
         Index           =   9
         Left            =   6360
         TabIndex        =   43
         Top             =   2190
         Width           =   1005
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Folio Colombia:"
         Height          =   195
         Index           =   8
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   1065
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Folio Colombia:"
         Height          =   195
         Index           =   7
         Left            =   3600
         TabIndex        =   39
         Top             =   2190
         Width           =   1065
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Aplica Alianza del Pacífico:"
         Height          =   195
         Left            =   5760
         TabIndex        =   38
         Top             =   1740
         Width           =   1935
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Criterio origen CA:"
         Height          =   195
         Left            =   4905
         TabIndex        =   35
         Top             =   2625
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Criterio origen USA:"
         Height          =   195
         Left            =   135
         TabIndex        =   33
         Top             =   2595
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Fraccion  Americana:"
         Height          =   195
         Index           =   5
         Left            =   165
         TabIndex        =   32
         Top             =   2190
         Width           =   1500
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Hecho en:"
         Height          =   195
         Index           =   3
         Left            =   4710
         TabIndex        =   30
         Top             =   1365
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Aplica Centro America:"
         Height          =   195
         Left            =   2760
         TabIndex        =   29
         Top             =   1740
         Width           =   1605
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Aplica USA:"
         Height          =   195
         Left            =   180
         TabIndex        =   28
         Top             =   1740
         Width           =   855
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Artículo:"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   22
         Top             =   315
         Width           =   600
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Fraccion  A.:"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   21
         Top             =   660
         Width           =   900
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Composición:"
         Height          =   195
         Index           =   2
         Left            =   3540
         TabIndex        =   20
         Top             =   660
         Width           =   945
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Cóntenido:"
         Height          =   195
         Index           =   4
         Left            =   165
         TabIndex        =   19
         Top             =   1020
         Width           =   765
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Origen:"
         Height          =   195
         Index           =   6
         Left            =   165
         TabIndex        =   18
         Top             =   1380
         Width           =   510
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9240
      Picture         =   "frmoracle_complementos_articulos_packing_list.frx":063C
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1320
      Picture         =   "frmoracle_complementos_articulos_packing_list.frx":0C76
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   330
      Picture         =   "frmoracle_complementos_articulos_packing_list.frx":0D78
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   0
      Picture         =   "frmoracle_complementos_articulos_packing_list.frx":0E7A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   0
      Width           =   330
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3465
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_complementos_articulos_packing_list.frx":0F7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_complementos_articulos_packing_list.frx":1856
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_complementos_articulos_packing_list.frx":2130
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_complementos_articulos_packing_list.frx":26CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_complementos_articulos_packing_list.frx":2FA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_complementos_articulos_packing_list.frx":3882
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_complementos_articulos_packing_list.frx":415C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_complementos_articulos_packing_list.frx":426E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_complementos_articulos_packing_list.frx":4380
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_complementos_articulos_packing_list.frx":4492
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_complementos_articulos_packing_list.frx":45A4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame5 
      Height          =   120
      Left            =   0
      TabIndex        =   24
      Top             =   270
      Width           =   9630
   End
End
Attribute VB_Name = "frmoracle_complementos_articulos_packing_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter

Private Sub llena_lv()
    Me.lv_complementos.ListItems.Clear
    rs.Open "select a.*, b.description from xxvia_tb_complementos_pk_list a, xxvia_system_items_b b where a.codigo = b.segment1  and b.organization_id = 93", cnnoracle_4, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
          Set list_item = lv_complementos.ListItems.Add(, , IIf(IsNull(rs!codigo), "", rs!codigo))
          list_item.SubItems(1) = IIf(IsNull(rs!Description), "", rs!Description)
          list_item.SubItems(2) = IIf(IsNull(rs!fraccion_arancelaria), 0, rs!fraccion_arancelaria)
          list_item.SubItems(3) = IIf(IsNull(rs!composicion), "", rs!composicion)
          list_item.SubItems(4) = IIf(IsNull(rs!contenido), 0, rs!contenido)
          list_item.SubItems(5) = IIf(IsNull(rs!originario), "", rs!originario)
          list_item.SubItems(6) = IIf(IsNull(rs!aplica_usa), "", rs!aplica_usa)
          list_item.SubItems(7) = IIf(IsNull(rs!aplica_ca), "", rs!aplica_ca)
          list_item.SubItems(8) = IIf(IsNull(rs!hecho_en), "", rs!hecho_en)
          list_item.SubItems(9) = IIf(IsNull(rs!fraccion_americana), "", rs!fraccion_americana)
          list_item.SubItems(10) = IIf(IsNull(rs!criterio_usa), "", rs!criterio_usa)
          list_item.SubItems(11) = IIf(IsNull(rs!criterio_ca), "", rs!criterio_ca)
          list_item.SubItems(12) = IIf(IsNull(rs!aplica_ap), "", rs!aplica_ap)
          list_item.SubItems(13) = IIf(IsNull(rs!folio_colombia), "", rs!folio_colombia)
          list_item.SubItems(14) = IIf(IsNull(rs!complemento), "", rs!complemento)
          rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub cmb_aplica_ca_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub cmb_aplica_usa_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmb_aplica_ca.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub cmd_carga_mmasiva_Click()
On Error GoTo SALIR:
   strConnectionString = "DRIVER=Microsoft Excel Driver (*.xls); DBQ=C:\REPORTESSID\COMPLEMENTOS.XLS"
   If rsaux8.State = 1 Then
      rsaux8.Close
   End If
   rsaux8.Open "select * FROM [COMPLEMENTOS$]", strConnectionString
   While Not rsaux8.EOF
         strconsulta = "select * from xxvia_tb_complementos_pk_list where codigo = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rsaux8!codigo), "", rsaux8!codigo))
              .Parameters.Append parametro
         End With
         Set rsaux9 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         If rsaux9.EOF Then
            strconsulta = "insert into xxvia_tb_complementos_pk_list (codigo, fraccion_arancelaria, contenido, composicion, originario, hecho_en, aplica_usa, aplica_ca, fraccion_americana, criterio_usa, criterio_ca, aplica_ap, folio_colombia, complemento) values  (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rsaux8!codigo), "", rsaux8!codigo))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, IIf(IsNull(rsaux8!fraccion), 0, rsaux8!fraccion))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rsaux8!contenido), "", rsaux8!contenido))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rsaux8!composicion), "", rsaux8!composicion))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rsaux8!ORIGEN), "", rsaux8!ORIGEN))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rsaux8!hecho_en), "", rsaux8!hecho_en))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rsaux8!aplica_usa), "", rsaux8!aplica_usa))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rsaux8!aplica_ca), "", rsaux8!aplica_ca))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, IIf(IsNull(rsaux8!fraccion_americana), 0, rsaux8!fraccion_americana))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rsaux8!criterio_usa), 0, rsaux8!criterio_usa))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rsaux8!criterio_ca), 0, rsaux8!criterio_ca))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rsaux8!aplica_ap), 0, rsaux8!aplica_ap))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rsaux8!folio_colombia), 0, rsaux8!folio_colombia))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rsaux8!complemento), 0, rsaux8!complemento))
                 .Parameters.Append parametro
            End With
            Set rsaux10 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
         Else
            strconsulta = "update xxvia_tb_complementos_pk_list set fraccion_arancelaria = ?, contenido = ?, composicion = ?, originario = ?, hecho_en = ?, aplica_usa = ?, aplica_ca = ?, fraccion_americana = ?, criterio_usa = ?, criterio_ca = ?, aplica_ap = ?, folio_colombia = ?, complemento = ? where codigo = ?"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, IIf(IsNull(rsaux8!fraccion), 0, rsaux8!fraccion))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rsaux8!contenido), "", rsaux8!contenido))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rsaux8!composicion), "", rsaux8!composicion))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rsaux8!ORIGEN), "", rsaux8!ORIGEN))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rsaux8!hecho_en), "", rsaux8!hecho_en))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rsaux8!aplica_usa), "", rsaux8!aplica_usa))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rsaux8!aplica_ca), "", rsaux8!aplica_ca))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, IIf(IsNull(rsaux8!fraccion_americana), 0, rsaux8!fraccion_americana))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rsaux8!criterio_usa), "", rsaux8!criterio_usa))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rsaux8!criterio_ca), "", rsaux8!criterio_ca))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rsaux8!aplica_ap), 0, rsaux8!aplica_ap))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rsaux8!folio_colombia), 0, rsaux8!folio_colombia))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rsaux8!complemento), 0, rsaux8!complemento))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, IIf(IsNull(rsaux8!codigo), "", rsaux8!codigo))
                 .Parameters.Append parametro
            End With
            Set rsaux10 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
         End If
         rsaux9.Close
         rsaux8.MoveNext
   Wend
   rsaux8.Close
   Call llena_lv
   MsgBox "Se a terminado el proceso de carga masiva", vbOKOnly, "ATENCION"
   Exit Sub
SALIR:
   MsgBox "Error al cargar el archivo, verifique que el archivo se llame complementos, la hoja se llame complementos, que los nombres de las columnas sean codigo, fraccion, contenido, composicion, origen, hecho_en y que el archivo se guarde en c:\reportessid\", vbOKOnly, "ATENCION"
End Sub

Private Sub cmd_complementos_Click()
   If Me.txt_codigo <> "" Then
      var_codigo_complemento = Me.txt_codigo
      var_descripcion_complemento = Me.txt_descripcion
      frmoracle_complementos.Show 1
   Else
      MsgBox "No se a seleccionado un artículo", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_deshacer_Click()
   If Me.lv_complementos.ListItems.Count > 0 Then
      Me.lv_complementos.SetFocus
   End If
End Sub

Private Sub cmd_eliminar_Click()
   If Me.txt_codigo <> "" Then
      var_si = MsgBox("¿Desea eliminar el registro?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         var_si = MsgBox("Confirmar la eliminación del registro", vbYesNo, "ATENCION")
         If var_si = 6 Then
            valor = txt_codigo
            var_n = Me.lv_complementos.ListItems.Count
            var_encontro = 0
            var_i = 1
            While (var_i <= var_n)
                  Me.lv_complementos.ListItems.Item(var_i).Selected = True
                  valor = Trim(Me.lv_complementos.selectedItem)
                  If txt_codigo = valor Then
                     var_encontro = 1
                     var_i = var_n + 1
                  End If
                  var_i = var_i + 1
            Wend

            strconsulta = "DELETE FROM xxvia_tb_complementos_pk_list where codigo = ?"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
                 .Parameters.Append parametro
            End With
            Set rsaux10 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            Me.lv_complementos.ListItems.Remove (Me.lv_complementos.selectedItem.Index)
            Me.txt_codigo = ""
            Me.txt_descripcion = ""
            Me.txt_fraccion = ""
            Me.txt_composicion = ""
            Me.txt_contenido = ""
            Me.txt_origen = ""
            Me.txt_criterio_origen_CA = ""
            Me.txt_criterio_origen_USA = ""
            MsgBox "Se elimino el registro", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub

Private Sub cmd_guardar_Click()
   If Me.txt_descripcion <> "" Then
      If IsNumeric(Me.txt_fraccion) Then
         strconsulta = "select * from xxvia_tb_complementos_pk_list where codigo = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
              .Parameters.Append parametro
         End With
         Set rsaux9 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         If rsaux9.EOF Then
            strconsulta = "insert into xxvia_tb_complementos_pk_list (codigo, fraccion_arancelaria, contenido, composicion, originario, aplica_usa, aplica_ca, hecho_en, fraccion_americana, criterio_usa, criterio_ca, APLICA_AP, FOLIO_COLOMBIA, COMPLEMENTO) values  (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_fraccion))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_contenido)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_composicion)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_origen)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.cmb_aplica_usa)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.cmb_aplica_ca)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_hecho_en)
                 .Parameters.Append parametro
                 If Not IsNumeric(Me.txt_fraccion_americana) Then
                    Me.txt_fraccion_americana = 0
                 End If
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_fraccion_americana))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_criterio_origen_USA)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_criterio_origen_CA)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.cmb_aplica_AP)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_folio)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento)
                 .Parameters.Append parametro
            End With
            Set rsaux10 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            MsgBox "Se a insertado el registro", vbOKOnly, "ATENCION"
             
         Else
            var_si = MsgBox("¿Desea actualizar los datos?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               strconsulta = "update xxvia_tb_complementos_pk_list set fraccion_arancelaria = ?, contenido = ?, composicion = ?, originario = ?, aplica_usa = ?, aplica_ca = ?, hecho_en = ?, fraccion_americana = ?, criterio_usa = ?, criterio_ca = ?, APLICA_AP = ?, FOLIO_COLOMBIA = ?, COMPLEMENTO = ? where codigo = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_fraccion))
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_contenido)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_composicion)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_origen)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.cmb_aplica_usa)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.cmb_aplica_ca)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_hecho_en)
                    .Parameters.Append parametro
                    If Not IsNumeric(Me.txt_fraccion_americana) Then
                       Me.txt_fraccion_americana = 0
                    End If
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_fraccion_americana))
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_criterio_origen_USA)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_criterio_origen_CA)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.cmb_aplica_AP)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_folio)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
                    .Parameters.Append parametro
               End With
               Set rsaux10 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               MsgBox "Se a actualizado el registro", vbOKOnly, "ATENCION"
            Else
               MsgBox "Se a cancelado la actualización", vbOKOnly, "ATENCION"
            End If
         End If
         rsaux9.Close
         Me.lv_complementos.ListItems.Clear
         Call llena_lv
      Else
         MsgBox "La fracción arancelaria es incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a seleccionado un artículo", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_codigo = ""
   Me.txt_descripcion = ""
   Me.txt_fraccion = ""
   Me.txt_composicion = ""
   Me.txt_contenido = ""
   Me.txt_contenido = ""
   Me.txt_origen = ""
   Me.txt_hecho_en = ""
   Me.cmb_aplica_usa = ""
   Me.cmb_aplica_ca = ""
   Me.txt_fraccion_americana = ""
   Me.txt_criterio_origen_USA = ""
   Me.txt_criterio_origen_CA = ""
   Me.txt_codigo.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 1000
   Call llena_lv
   Me.frm_origenes.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_complementos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_complementos, ColumnHeader)
End Sub

Private Sub lv_complementos_ItemClick(ByVal Item As MSComctlLib.ListItem)
   If Me.lv_complementos.ListItems.Count > 0 Then
      Me.txt_codigo = Me.lv_complementos.selectedItem
      Me.txt_descripcion = Me.lv_complementos.selectedItem.SubItems(1)
      Me.txt_fraccion = Me.lv_complementos.selectedItem.SubItems(2)
      Me.txt_composicion = Me.lv_complementos.selectedItem.SubItems(3)
      Me.txt_contenido = Me.lv_complementos.selectedItem.SubItems(4)
      Me.txt_origen = Me.lv_complementos.selectedItem.SubItems(5)
      Me.cmb_aplica_usa = Me.lv_complementos.selectedItem.SubItems(6)
      Me.cmb_aplica_ca = Me.lv_complementos.selectedItem.SubItems(7)
      Me.txt_hecho_en = Me.lv_complementos.selectedItem.SubItems(8)
      Me.txt_fraccion_americana = Me.lv_complementos.selectedItem.SubItems(9)
      Me.txt_criterio_origen_USA = Me.lv_complementos.selectedItem.SubItems(10)
      Me.txt_criterio_origen_CA = Me.lv_complementos.selectedItem.SubItems(11)
      Me.cmb_aplica_AP = Me.lv_complementos.selectedItem.SubItems(12)
      Me.txt_folio = Me.lv_complementos.selectedItem.SubItems(13)
      Me.txt_complemento = Me.lv_complementos.selectedItem.SubItems(14)
   End If
End Sub

Private Sub lv_origenes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_origen = Me.lv_origenes.selectedItem
      Me.txt_origen.SetFocus
   End If
   If KeyAscii = 27 Then
      Me.txt_origen.SetFocus
   End If
End Sub

Private Sub lv_origenes_LostFocus()
   Me.frm_origenes.Visible = False
End Sub

Private Sub txt_codigo_Change()
   Me.txt_descripcion = ""
   Me.txt_contenido = ""
   Me.txt_composicion = ""
   Me.txt_fraccion = ""
   Me.txt_origen = ""
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_descripcion.SetFocus
   End If
End Sub

Private Sub txt_codigo_LostFocus()
   If Me.txt_codigo <> "" Then
      strconsulta = "select * from xxvia_tb_complementos_pk_list where codigo = ?"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If rsaux9.EOF Then
         strconsulta = "select * from XXVIA_SYSTEM_ITEMS_B where SEGMENT1 = ? AND ORGANIZATION_ID = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(var_unidad_organizacional))
              .Parameters.Append parametro
         End With
         Set rsaux10 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         If Not rsaux10.EOF Then
            Me.txt_descripcion = IIf(IsNull(rsaux10!Description), "", rsaux10!Description)
            'Me.txt_fraccion.SetFocus
         Else
            Me.txt_descripcion = ""
            Me.txt_fraccion = ""
            Me.txt_composicion = ""
            Me.txt_contenido = ""
            Me.txt_origen = ""
            MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
         End If
         rsaux10.Close
      Else
         valor = txt_codigo
         var_n = Me.lv_complementos.ListItems.Count
         var_encontro = 0
         var_i = 1
         While (var_i <= var_n)
               Me.lv_complementos.ListItems.Item(var_i).Selected = True
               valor = Trim(Me.lv_complementos.selectedItem)
               If txt_codigo = valor Then
                  var_encontro = 1
                  var_i = var_n + 1
               End If
               var_i = var_i + 1
         Wend
         If var_encontro = 1 Then
             Me.lv_complementos.SetFocus
             Me.txt_descripcion = Me.lv_complementos.selectedItem.SubItems(1)
             Me.txt_fraccion = Me.lv_complementos.selectedItem.SubItems(2)
             Me.txt_composicion = Me.lv_complementos.selectedItem.SubItems(3)
             Me.txt_contenido = Me.lv_complementos.selectedItem.SubItems(4)
             Me.txt_origen = Me.lv_complementos.selectedItem.SubItems(5)
             Me.cmb_aplica_usa = Me.lv_complementos.selectedItem.SubItems(6)
             Me.cmb_aplica_ca = Me.lv_complementos.selectedItem.SubItems(7)
             Me.txt_hecho_en = Me.lv_complementos.selectedItem.SubItems(8)
             Me.txt_fraccion_americana = Me.lv_complementos.selectedItem.SubItems(9)
             Me.txt_criterio_origen_USA = Me.lv_complementos.selectedItem.SubItems(10)
             Me.txt_criterio_origen_CA = Me.lv_complementos.selectedItem.SubItems(11)
             
             Set itmfound = Me.lv_complementos.findItem(Trim(txt_codigo), lvwText, , lvwPartial)
             If Not itmfound Is Nothing Then
                itmfound.EnsureVisible
                itmfound.Selected = True
             End If
         Else
             Me.txt_descripcion = ""
             Me.txt_fraccion = ""
             Me.txt_composicion = ""
             Me.txt_contenido = ""
             Me.txt_origen = ""
         End If
      End If
   Else
      Me.txt_descripcion = ""
      Me.txt_fraccion = ""
      Me.txt_composicion = ""
      Me.txt_contenido = ""
      Me.txt_origen = ""
   End If
End Sub

Private Sub txt_complemento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_criterio_origen_USA.SetFocus
   End If
End Sub

Private Sub txt_composicion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_contenido.SetFocus
   End If
End Sub

Private Sub txt_contenido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_origen.SetFocus
   End If
End Sub

Private Sub txt_criterio_origen_CA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Me.cmd_guardar.SetFocus
    End If

End Sub

Private Sub txt_descripcion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_fraccion.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_folio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_complemento.SetFocus
   End If
End Sub

Private Sub txt_fraccion_americana_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Me.txt_folio.SetFocus
    End If
End Sub

Private Sub txt_fraccion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_composicion.SetFocus
   End If
End Sub

Private Sub txt_hecho_en_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmb_aplica_usa.SetFocus
   End If
End Sub

Private Sub txt_origen_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      rs.Open "select distinct originario from xxvia_tb_complementos_pk_list", cnnoracle_4, adOpenDynamic, adLockOptimistic
      Me.lv_origenes.ListItems.Clear
      While Not rs.EOF
            Set list_item = lv_origenes.ListItems.Add(, , IIf(IsNull(rs(0).Value), "", rs(0).Value))
            rs.MoveNext
      Wend
      rs.Close
      Me.frm_origenes.Visible = True
      Me.lv_origenes.SetFocus
   End If
End Sub

Private Sub txt_origen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_hecho_en.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub
