VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmequipos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Equipos"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   Icon            =   "frmequipos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11670
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmequipos.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11190
      Picture         =   "frmequipos.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame7 
      Height          =   120
      Left            =   105
      TabIndex        =   41
      Top             =   255
      Width           =   11460
   End
   Begin VB.Frame Frame2 
      Height          =   6885
      Left            =   5685
      TabIndex        =   27
      Top             =   360
      Width           =   5880
      Begin VB.CommandButton com_nuevo_orden 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   45
         Picture         =   "frmequipos.frx":1006
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   375
         Width           =   330
      End
      Begin VB.CommandButton com_guardar_orden 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   375
         Picture         =   "frmequipos.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   375
         Width           =   330
      End
      Begin VB.CommandButton com_deshacer_orden 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   705
         Picture         =   "frmequipos.frx":120A
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   375
         Width           =   330
      End
      Begin VB.CommandButton com_eliminar_orden 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1035
         Picture         =   "frmequipos.frx":12DC
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   375
         Width           =   330
      End
      Begin VB.TextBox txt_cantidad_total 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4425
         TabIndex        =   24
         Top             =   6480
         Width           =   1395
      End
      Begin VB.TextBox txt_orden_surtido 
         Height          =   315
         Left            =   1005
         TabIndex        =   22
         Top             =   795
         Width           =   900
      End
      Begin MSComctlLib.ListView lv_orden_surtido 
         Height          =   5310
         Left            =   60
         TabIndex        =   23
         Top             =   1125
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   9366
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Pedido"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Ord. S."
            Object.Width           =   1570
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Agente"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Piezas"
            Object.Width           =   1587
         EndProperty
      End
      Begin VB.Frame Frame6 
         Height          =   75
         Left            =   15
         TabIndex        =   38
         Top             =   675
         Width           =   5835
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         Caption         =   " Ordenes de Surtido "
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   30
         TabIndex        =   37
         Top             =   120
         Width           =   5805
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Total Piezas:"
         Height          =   195
         Left            =   3465
         TabIndex        =   29
         Top             =   6525
         Width           =   915
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "N?mero:"
         Height          =   195
         Left            =   195
         TabIndex        =   28
         Top             =   855
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3225
      Left            =   105
      TabIndex        =   0
      Top             =   360
      Width           =   5505
      Begin MSComCtl2.MonthView mes 
         Height          =   2370
         Left            =   2100
         TabIndex        =   42
         Top             =   1185
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   58261505
         CurrentDate     =   37581
      End
      Begin VB.CommandButton cmdfecha 
         Height          =   285
         Index           =   0
         Left            =   2070
         Picture         =   "frmequipos.frx":13DE
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Seleccione la fecha"
         Top             =   840
         Width           =   315
      End
      Begin VB.TextBox txt_fecha 
         Height          =   315
         Left            =   990
         TabIndex        =   7
         Top             =   825
         Width           =   1050
      End
      Begin VB.CommandButton com_eliminar_equipo 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1035
         Picture         =   "frmequipos.frx":14E0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton com_deshacer_equipo 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   705
         Picture         =   "frmequipos.frx":15E2
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton com_guardar_equipo 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   375
         Picture         =   "frmequipos.frx":16B4
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   330
      End
      Begin VB.CommandButton com_nuevo_equipo 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   45
         Picture         =   "frmequipos.frx":17B6
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   330
      End
      Begin VB.TextBox txt_jaula 
         Height          =   315
         Left            =   990
         TabIndex        =   10
         Top             =   1155
         Width           =   1050
      End
      Begin VB.TextBox txt_cantidad_meta 
         Height          =   315
         Left            =   3630
         TabIndex        =   9
         Top             =   825
         Width           =   1755
      End
      Begin VB.Frame Frame4 
         Height          =   75
         Left            =   15
         TabIndex        =   35
         Top             =   660
         Width           =   5460
      End
      Begin MSComctlLib.ListView lv_jaulas 
         Height          =   1605
         Left            =   30
         TabIndex        =   39
         Top             =   1515
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   2831
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
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7585
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Meta"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   180
         TabIndex        =   40
         Top             =   885
         Width           =   495
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         Caption         =   "  Jaulas"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   30
         TabIndex        =   33
         Top             =   120
         Width           =   5430
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad Meta:"
         Height          =   195
         Left            =   2475
         TabIndex        =   26
         Top             =   885
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Jaula:"
         Height          =   195
         Left            =   195
         TabIndex        =   25
         Top             =   1215
         Width           =   420
      End
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   4680
      Top             =   465
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   12
      ImageHeight     =   12
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequipos.frx":18B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequipos.frx":2192
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   135
      Top             =   390
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
            Picture         =   "frmequipos.frx":2A6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequipos.frx":3346
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequipos.frx":3C20
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequipos.frx":41BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequipos.frx":4A98
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequipos.frx":5372
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequipos.frx":5C4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequipos.frx":5D5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequipos.frx":5E70
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequipos.frx":5F82
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequipos.frx":6094
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   3570
      Top             =   165
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequipos.frx":61A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequipos.frx":6A80
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequipos.frx":735A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequipos.frx":78F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequipos.frx":81D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequipos.frx":8AAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequipos.frx":9386
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequipos.frx":9498
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequipos.frx":95AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequipos.frx":96BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequipos.frx":97CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmequipos.frx":98E0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   3720
      Left            =   105
      TabIndex        =   30
      Top             =   3525
      Width           =   5505
      Begin VB.ComboBox cmb_puestos 
         Height          =   315
         ItemData        =   "frmequipos.frx":99F2
         Left            =   975
         List            =   "frmequipos.frx":99FC
         TabIndex        =   17
         Top             =   1155
         Width           =   2415
      End
      Begin VB.CommandButton com_nuevo_integrante 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   45
         Picture         =   "frmequipos.frx":9A15
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   375
         Width           =   330
      End
      Begin VB.CommandButton com_guardar_integrante 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   375
         Picture         =   "frmequipos.frx":9B17
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   375
         Width           =   330
      End
      Begin VB.CommandButton com_deshacer_integrante 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   705
         Picture         =   "frmequipos.frx":9C19
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   375
         Width           =   330
      End
      Begin VB.CommandButton com_eliminar_integrante 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1035
         Picture         =   "frmequipos.frx":9CEB
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   375
         Width           =   330
      End
      Begin VB.ComboBox cmb_integrantes 
         Height          =   315
         Left            =   2070
         TabIndex        =   16
         Top             =   825
         Width           =   3345
      End
      Begin VB.TextBox txt_integrante 
         Height          =   315
         Left            =   975
         TabIndex        =   15
         Top             =   825
         Width           =   1080
      End
      Begin MSComctlLib.ListView lv_integrantes 
         Height          =   2055
         Left            =   45
         TabIndex        =   31
         Top             =   1545
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   3625
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
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   5468
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Puesto"
            Object.Width           =   2293
         EndProperty
      End
      Begin VB.Frame Frame5 
         Height          =   75
         Left            =   15
         TabIndex        =   36
         Top             =   675
         Width           =   5460
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Puesto:"
         Height          =   195
         Left            =   165
         TabIndex        =   43
         Top             =   1185
         Width           =   540
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         Caption         =   " Integrante"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   30
         TabIndex        =   34
         Top             =   120
         Width           =   5430
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Integrante:"
         Height          =   195
         Left            =   165
         TabIndex        =   32
         Top             =   885
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmequipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_nombre_jaula As String
Dim numero_items_jaulas As Integer
Dim numero_items_integrantes As Integer
Dim numero_items_ordenes As Integer
Dim cantidad_total As Double
Dim var_guardar_jaula As Boolean
Dim var_guardar_integrante As Boolean
Dim var_guardar_orden As Boolean
Dim var_cantidad_surtir As Double



Private Sub cmb_integrantes_Click()
   txt_integrante = Obtener_llave(cnn, rs, "tb_personal", "VCHA_per_NOMBRE", cmb_integrantes, 0, "T")
End Sub

Private Sub cmb_integrantes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmb_puestos.SetFocus
   End If
End Sub

Private Sub cmb_puestos_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If com_guardar_integrante.Enabled = True Then
         com_guardar_integrante.SetFocus
      End If
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmdfecha_Click(Index As Integer)
   mes.Value = Date
   mes.Visible = True
   mes.SetFocus
End Sub

Private Sub com_deshacer_equipo_Click()
   com_deshacer_equipo.Enabled = False
   com_guardar_equipo.Enabled = False
   com_nuevo_equipo.Enabled = True
End Sub

Private Sub com_eliminar_equipo_Click()
   If Trim(txt_jaula) <> "" Then
      rs.Open "delete from tb_relacion_equipos where inte_jau_jaula_id = " + txt_jaula + " and dtim_equ_fecha = '" + txt_fecha + "'", cnn, adOpenDynamic, adLockOptimistic
      rs.Open "delete from tb_equipos where inte_jau_jaula_id = " + txt_jaula + " and dtim_equ_fecha = '" + txt_fecha + "'", cnn, adOpenDynamic, adLockOptimistic
      lv_jaulas.ListItems.Remove (lv_jaulas.selectedItem.Index)
      lv_integrantes.ListItems.Clear
      lv_orden_surtido.ListItems.Clear
      numero_items_jaulas = numero_items_jaulas - 1
      If numero_items_jaulas > 0 Then
         com_eliminar_equipo.Enabled = True
         txt_jaula = lv_jaulas.selectedItem
         txt_cantidad_meta = lv_jaulas.selectedItem.SubItems(2)
         integrantes
      Else
         com_eliminar_equipo.Enabled = False
         txt_jaula = ""
         txt_cantidad_meta = ""
         txt_integrante = ""
         cmb_integrantes.Text = ""
         txt_orden_surtido = ""
      End If
      com_guardar_equipo.Enabled = False
      com_nuevo_equipo.Enabled = True
      com_deshacer_equipo.Enabled = False
   End If
End Sub

Private Sub com_eliminar_integrante_Click()
   si = MsgBox("?Deseas eliminar al integrante del equipo?", vbYesNo, "ATENCION")
   If si = 6 Then
      rs.Open "delete from tb_equipos where inte_jau_jaula_id = " + txt_jaula + " and vcha_per_personal_id = '" + txt_integrante + "' and dtim_equ_fecha = '" + txt_fecha + "'"
      lv_integrantes.ListItems.Remove (lv_integrantes.selectedItem.Index)
      numero_items_integrantes = numero_items_integrantes - 1
      If numero_items_integrantes > 0 Then
         com_eliminar_integrante.Enabled = True
         txt_integrante = lv_integrantes.selectedItem
         cmb_integrantes.Text = lv_integrantes.selectedItem.SubItems(1)
      Else
         com_eliminar_integrante.Enabled = False
         txt_integrante = ""
         cmb_integrantes.Text = ""
      End If
      com_nuevo_integrante.Enabled = True
      com_guardar_integrante.Enabled = False
      com_deshacer_integrante.Enabled = False
   End If
End Sub

Private Sub com_eliminar_orden_Click()
   si = MsgBox("?Deseas eliminar la orden de surtido del equipo?", vbYesNo, "ATENCION")
   If si = 6 Then
      rs.Open "delete from tb_relacion_equipos where inte_ors_orden_surtido = " + txt_orden_surtido + " and dtim_equ_fecha = '" + txt_fecha + "'", cnn, adOpenDynamic, adLockOptimistic
      var_cantidad_surtir = var_cantidad_surtir - lv_orden_surtido.selectedItem.SubItems(3)
      lv_orden_surtido.ListItems.Remove (lv_orden_surtido.selectedItem.Index)
      txt_cantidad_total = Format(var_cantidad_surtir, "###,###,##0.00")
      numero_items_ordenes = numero_items_ordenes - 1
      If numero_items_ordenes > 0 Then
         com_eliminar_orden.Enabled = True
         txt_orden_surtido = lv_orden_surtido.selectedItem.SubItems(1)
      Else
         com_eliminar_orden.Enabled = False
         txt_orden_surtido = ""
      End If
      com_nuevo_orden.Enabled = True
      com_guardar_orden.Enabled = False
      com_deshacer_orden.Enabled = False
   End If
End Sub

Private Sub com_guardar_equipo_Click()
   Dim var_posible As Boolean
   If txt_cantidad_meta = "" Then
      txt_cantidad_meta = 0
   End If
   var_posible = False
   com_deshacer_equipo.Enabled = False
   com_guardar_equipo.Enabled = False
   com_nuevo_equipo.Enabled = True
   If Trim(txt_jaula) <> "" And var_guardar_jaula = True Then
      rs.Open "select * from tb_equipos where inte_jau_jaula_id = " + txt_jaula + " and DTIM_EQU_fecha = '" + txt_fecha + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_posible = False
      Else
         var_posible = True
      End If
      rs.Close
      com_nuevo_integrante.Enabled = True
      com_nuevo_integrante.SetFocus
      txt_jaula.Enabled = False
      txt_cantidad_meta.Enabled = False
      If var_posible = True Then
         Dim list_item As ListItem
         Set list_item = lv_jaulas.ListItems.Add(, , txt_jaula)
         list_item.SubItems(1) = var_nombre_jaula
         list_item.SubItems(2) = txt_cantidad_meta
         numero_items_jaulas = numero_items_jaulas + 1
         com_eliminar_equipo.Enabled = True
         com_nuevo_integrante.SetFocus
      Else
         lv_jaulas.SetFocus
         integrantes
         MsgBox "La jaula ya esta asignada a un equipo", vbOKOnly, "ATENCION"
      End If
   Else
      txt_jaula = ""
      txt_cantidad_meta = ""
      MsgBox "La clave de la jaula no existe", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub com_guardar_integrante_Click()
   Dim var_posible As Boolean
   var_posible = False
   rs.Open "select * from tb_Equipos where vcha_per_personal_id = '" + txt_integrante + "' and dtim_equ_fecha = '" + txt_fecha + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_posible = False
   Else
      var_posible = True
   End If
   rs.Close
   If var_posible = True Then
      rs.Open "insert into tb_equipos ([dtim_equ_fecha], [inte_jau_jaula_id], [vcha_per_personal_id], [floa_equ_cantidad_meta], [vcha_equ_puesto]) values ('" + txt_fecha + "', " + txt_jaula + ", '" + txt_integrante + "', " + txt_cantidad_meta + ", '" + cmb_puestos + "')", cnn, adOpenDynamic, adLockOptimistic
      Dim list_item As ListItem
      Set list_item = lv_integrantes.ListItems.Add(, , txt_integrante)
      list_item.SubItems(1) = cmb_integrantes.Text
      list_item.SubItems(2) = Me.cmb_puestos.Text
      numero_items_integrantes = numero_items_integrantes + 1
  Else
     MsgBox "No puede asignar esta persona al equipo ya que fue asignado a este o a otro equipo con anterioridad", vbOKOnly, "ATENCION"
  End If
  com_nuevo_integrante.Enabled = True
  com_deshacer_integrante.Enabled = False
  com_guardar_integrante.Enabled = False
  com_nuevo_integrante.SetFocus
  If numero_items_integrantes > 0 Then
     com_nuevo_orden.Enabled = True
     com_deshacer_orden.Enabled = False
     com_guardar_orden.Enabled = False
     com_eliminar_integrante.Enabled = True
  Else
     com_nuevo_orden.Enabled = False
     com_deshacer_orden.Enabled = False
     com_guardar_orden.Enabled = False
     com_eliminar_integrante.Enabled = True
  End If
End Sub

Private Sub com_guardar_orden_Click()
   Dim var_posible As Boolean
   If Trim(txt_integrante) <> "" Then
      If Trim(txt_orden_surtido) <> "" Then
         rs.Open "select * from tb_relacion_equipos where inte_jau_jaula_id = " + txt_jaula + " and dtim_equ_fecha = '" + txt_fecha + "' and inte_ors_orden_surtido = " + txt_orden_surtido, cnn, adOpenDynamic, adLockOptimistic
         var_posible = False
         If Not rs.EOF Then
            var_posible = False
            MsgBox "La orden de surtido ya fue asignada", vbOKOnly, "ATENCION"
         Else
            var_posible = True
         End If
         rs.Close
         If var_posible = True Then
            rs.Open "select * from vw_suma_cantidad_surtir where inte_ors_orden_surtido = " + txt_orden_surtido, cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               Dim list_item As ListItem
               Set list_item = lv_orden_surtido.ListItems.Add(, , rs!inte_ped_numero)
               list_item.SubItems(1) = IIf(IsNull(rs!inte_ors_orden_surtido), 0, rs!inte_ors_orden_surtido)
               list_item.SubItems(2) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
               list_item.SubItems(3) = Format(IIf(IsNull(rs!Cantidad), 0, rs!Cantidad), "###,###,##0.00")
               var_cantidad_surtir = var_cantidad_surtir + IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)
               numero_items_ordenes = numero_items_ordenes + 1
               rsaux.Open "insert into tb_relacion_equipos ([dtim_equ_fecha],[inte_jau_jaula_id], [inte_ors_orden_surtido], [FLOA_ORS_CANTIDAD]) values ('" + txt_fecha + "', " + txt_jaula + ", " + txt_orden_surtido + ", " + Str(var_cantidad_surtir) + ")", cnn, adOpenDynamic, adLockOptimistic
               txt_cantidad_total = Format(var_cantidad_surtir, "###,###,##0.00")
            Else
               MsgBox "La orden de surtido no existe", vbOKOnly, "ATENCION"
            End If
            rs.Close
         End If
      End If
   Else
      MsgBox "Las ordenes de surtido se podran dar de alta hasta que el equipo tenga integrantes", vbOKOnly, "ATENCION"
   End If
   com_nuevo_orden.Enabled = True
   com_deshacer_orden.Enabled = False
   com_guardar_orden.Enabled = False
   If numero_items_ordenes > 0 Then
      com_eliminar_orden.Enabled = True
   Else
      com_eliminar_orden.Enabled = False
   End If
End Sub

Private Sub com_nuevo_equipo_Click()
   var_guardar_jaula = False
   com_nuevo_equipo.Enabled = False
   com_guardar_equipo.Enabled = True
   com_deshacer_equipo.Enabled = True
   com_nuevo_integrante.Enabled = False
   com_deshacer_integrante.Enabled = False
   com_guardar_integrante.Enabled = False
   com_eliminar_integrante.Enabled = False
   com_nuevo_orden.Enabled = False
   com_guardar_orden.Enabled = False
   com_deshacer_orden.Enabled = False
   com_eliminar_orden.Enabled = False
   txt_jaula.Enabled = True
   txt_jaula = ""
   txt_integrante = ""
   cmb_integrantes.Text = ""
   lv_integrantes.ListItems.Clear
   txt_orden_surtido = ""
   lv_orden_surtido.ListItems.Clear
   txt_jaula.SetFocus
   txt_cantidad_meta = ""
   txt_cantidad_total = ""
   var_cantidad_surtir = 0
End Sub

Private Sub com_nuevo_integrante_Click()
   txt_integrante = ""
   cmb_integrantes.Text = ""
   com_nuevo_integrante.Enabled = False
   com_deshacer_integrante.Enabled = True
   com_guardar_integrante.Enabled = True
   cmb_puestos = "EMBARQUE"
   cmb_integrantes.Enabled = True
   txt_integrante.Enabled = True
   txt_integrante.SetFocus
End Sub

Private Sub com_nuevo_orden_Click()
   txt_orden_surtido = ""
   txt_orden_surtido.Enabled = True
   com_deshacer_orden.Enabled = True
   com_guardar_orden.Enabled = True
   com_nuevo_orden.Enabled = False
   txt_orden_surtido.SetFocus
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 0
   Left = 0
   mes.Visible = False
   txt_fecha = Date
   txt_fecha.Enabled = False
   txt_cantidad_meta.Enabled = False
   txt_integrante.Enabled = False
   cmb_integrantes.Enabled = False
   txt_jaula.Enabled = False
   txt_orden_surtido.Enabled = False
   txt_cantidad_total.Enabled = False
   com_guardar_equipo.Enabled = False
   com_deshacer_equipo.Enabled = False
   com_eliminar_equipo.Enabled = False
   com_nuevo_integrante.Enabled = False
   com_deshacer_integrante.Enabled = False
   com_guardar_integrante.Enabled = False
   com_eliminar_integrante.Enabled = False
   com_nuevo_orden.Enabled = False
   com_guardar_orden.Enabled = False
   com_deshacer_orden.Enabled = False
   com_eliminar_orden.Enabled = False
   numero_items_jaulas = 0
   numero_items_integrantes = 0
   numero_items_ordenes = 0
   rs.Open "select vcha_per_nombre from tb_personal order by vcha_per_nombre", cnn, adOpenDynamic, adLockOptimistic
   Call RecsetToCombo(cmb_integrantes.hwnd, rs, 0)
   rs.Close
   rs.Open "select distinct a.inte_jau_jaula_id,b.vcha_jau_nombre,a.floa_equ_cantidad_meta from tb_equipos a, tb_jaulas b where a.dtim_equ_fecha = '" + txt_fecha + "' and a.inte_jau_jaula_id = b.inte_jau_jaula_id", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      Dim list_item As ListItem
      txt_jaula = rs!inte_jau_jaula_id
      txt_cantidad_meta = IIf(IsNull(rs!floa_equ_cantidad_meta), 0, rs!floa_equ_cantidad_meta)
      While Not rs.EOF
         Set list_item = lv_jaulas.ListItems.Add(, , rs!inte_jau_jaula_id)
         list_item.SubItems(1) = IIf(IsNull(rs!vcha_jau_nombre), "", rs!vcha_jau_nombre)
         list_item.SubItems(2) = IIf(IsNull(rs!floa_equ_cantidad_meta), 0, rs!floa_equ_cantidad_meta)
         numero_items_jaulas = numero_items_jaulas + 1
         rs.MoveNext
      Wend
      rs.Close
      integrantes
   Else
      rs.Close
   End If
   ordenes_surtido
   If Trim(txt_jaula) <> "" Then
      com_eliminar_equipo.Enabled = True
      com_nuevo_integrante.Enabled = True
      com_eliminar_integrante.Enabled = True
      com_nuevo_orden.Enabled = True
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_equipos)
End Sub

Private Sub lv_integrantes_ItemClick(ByVal Item As MSComctlLib.ListItem)
   txt_integrante = lv_integrantes.selectedItem
   cmb_integrantes.Text = lv_integrantes.selectedItem.SubItems(1)
   cmb_puestos = lv_integrantes.selectedItem.SubItems(2)
End Sub

Private Sub lv_jaulas_ItemClick(ByVal Item As MSComctlLib.ListItem)
  
   txt_cantidad_meta = lv_jaulas.selectedItem.SubItems(2)
   txt_jaula = lv_jaulas.selectedItem
   If Trim(txt_jaula) <> "" Then
      integrantes
      ordenes_surtido
   End If
End Sub

Private Sub lv_orden_surtido_ItemClick(ByVal Item As MSComctlLib.ListItem)
   txt_orden_surtido = lv_orden_surtido.selectedItem.SubItems(1)
End Sub

Private Sub mes_DblClick()
   txt_fecha = mes.Value
   lv_jaulas.ListItems.Clear
   lv_integrantes.ListItems.Clear
   lv_orden_surtido.ListItems.Clear
   txt_jaula = ""
   txt_cantidad_meta = ""
   txt_integrante = ""
   cmb_integrantes.Text = ""
   txt_cantidad_total = ""
   txt_orden_surtido = ""
   mes.Visible = False
   txt_fecha.Enabled = False
   txt_cantidad_meta.Enabled = False
   txt_integrante.Enabled = False
   cmb_integrantes.Enabled = False
   txt_jaula.Enabled = False
   txt_orden_surtido.Enabled = False
   txt_cantidad_total.Enabled = False
   com_guardar_equipo.Enabled = False
   com_deshacer_equipo.Enabled = False
   com_eliminar_equipo.Enabled = False
   com_nuevo_integrante.Enabled = False
   com_deshacer_integrante.Enabled = False
   com_guardar_integrante.Enabled = False
   com_eliminar_integrante.Enabled = False
   com_nuevo_orden.Enabled = False
   com_guardar_orden.Enabled = False
   com_deshacer_orden.Enabled = False
   com_eliminar_orden.Enabled = False
   numero_items_jaulas = 0
   numero_items_integrantes = 0
   numero_items_ordenes = 0
   rs.Open "select vcha_per_nombre from tb_personal order by vcha_per_nombre", cnn, adOpenDynamic, adLockOptimistic
   Call RecsetToCombo(cmb_integrantes.hwnd, rs, 0)
   rs.Close
   rs.Open "select distinct a.inte_jau_jaula_id,b.vcha_jau_nombre,a.floa_equ_cantidad_meta from tb_equipos a, tb_jaulas b where a.dtim_equ_fecha = '" + txt_fecha + "' and a.inte_jau_jaula_id = b.inte_jau_jaula_id", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      Dim list_item As ListItem
      txt_jaula = rs!inte_jau_jaula_id
      txt_cantidad_meta = IIf(IsNull(rs!floa_equ_cantidad_meta), 0, rs!floa_equ_cantidad_meta)
      While Not rs.EOF
         Set list_item = lv_jaulas.ListItems.Add(, , rs!inte_jau_jaula_id)
         list_item.SubItems(1) = IIf(IsNull(rs!vcha_jau_nombre), "", rs!vcha_jau_nombre)
         list_item.SubItems(2) = IIf(IsNull(rs!floa_equ_cantidad_meta), 0, rs!floa_equ_cantidad_meta)
         numero_items_jaulas = numero_items_jaulas + 1
         rs.MoveNext
      Wend
      rs.Close
      integrantes
   Else
      rs.Close
   End If
   ordenes_surtido
   If Trim(txt_jaula) <> "" Then
      com_eliminar_equipo.Enabled = True
      com_nuevo_integrante.Enabled = True
      com_eliminar_integrante.Enabled = True
      com_nuevo_orden.Enabled = True
   End If
End Sub

Private Sub mes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_fecha = mes.Value
      lv_jaulas.ListItems.Clear
      lv_integrantes.ListItems.Clear
      lv_orden_surtido.ListItems.Clear
      txt_jaula = ""
      txt_cantidad_meta = ""
      txt_integrante = ""
      cmb_integrantes.Text = ""
      txt_cantidad_total = ""
      txt_orden_surtido = ""
      mes.Visible = False
      mes.Visible = False
      txt_fecha.Enabled = False
      txt_cantidad_meta.Enabled = False
      txt_integrante.Enabled = False
      cmb_integrantes.Enabled = False
      txt_jaula.Enabled = False
      txt_orden_surtido.Enabled = False
      txt_cantidad_total.Enabled = False
      com_guardar_equipo.Enabled = False
      com_deshacer_equipo.Enabled = False
      com_eliminar_equipo.Enabled = False
      com_nuevo_integrante.Enabled = False
      com_deshacer_integrante.Enabled = False
      com_guardar_integrante.Enabled = False
      com_eliminar_integrante.Enabled = False
      com_nuevo_orden.Enabled = False
      com_guardar_orden.Enabled = False
      com_deshacer_orden.Enabled = False
      com_eliminar_orden.Enabled = False
      numero_items_jaulas = 0
      numero_items_integrantes = 0
      numero_items_ordenes = 0
      rs.Open "select vcha_per_nombre from tb_personal order by vcha_per_nombre", cnn, adOpenDynamic, adLockOptimistic
      Call RecsetToCombo(cmb_integrantes.hwnd, rs, 0)
      rs.Close
      rs.Open "select distinct a.inte_jau_jaula_id,b.vcha_jau_nombre,a.floa_equ_cantidad_meta from tb_equipos a, tb_jaulas b where a.dtim_equ_fecha = '" + txt_fecha + "' and a.inte_jau_jaula_id = b.inte_jau_jaula_id", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Dim list_item As ListItem
         txt_jaula = rs!inte_jau_jaula_id
         txt_cantidad_meta = IIf(IsNull(rs!floa_equ_cantidad_meta), 0, rs!floa_equ_cantidad_meta)
         While Not rs.EOF
            Set list_item = lv_jaulas.ListItems.Add(, , rs!inte_jau_jaula_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_jau_nombre), "", rs!vcha_jau_nombre)
            list_item.SubItems(2) = IIf(IsNull(rs!floa_equ_cantidad_meta), 0, rs!floa_equ_cantidad_meta)
            numero_items_jaulas = numero_items_jaulas + 1
            rs.MoveNext
         Wend
         rs.Close
         integrantes
      Else
         rs.Close
      End If
      ordenes_surtido
      If Trim(txt_jaula) <> "" Then
         com_eliminar_equipo.Enabled = True
         com_nuevo_integrante.Enabled = True
         com_eliminar_integrante.Enabled = True
         com_nuevo_orden.Enabled = True
      End If
   End If
   If KeyAscii = 27 Then
      mes.Visible = False
   End If
End Sub

Private Sub mes_LostFocus()
   mes.Visible = False
End Sub


Private Sub txt_cantidad_meta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(txt_cantidad_meta) <> "" Then
         com_guardar_equipo.SetFocus
      Else
         MsgBox "Cantidad Incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_fecha_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
End Sub

Private Sub txt_integrante_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Trim(txt_integrante) <> "" Then
         rs.Open "select * from tb_personal where vcha_per_personal_id = '" + txt_integrante + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            cmb_integrantes = rs!vcha_per_nombre
            rs.Close
            txt_integrante.Enabled = False
            cmb_integrantes.Enabled = False
            com_guardar_integrante.SetFocus
         Else
            MsgBox "Clave de Integrante Incorrecta", vbOKOnly, "ATENCION"
            rs.Close
            cmb_integrantes.SetFocus
         End If
      End If
   End If
End Sub

Private Sub txt_jaula_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If Trim(txt_jaula) <> "" Then
         var_nombre_jaula = ""
         rs.Open "select * from tb_jaulas where inte_jau_jaula_id = " + txt_jaula, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_guardar_jaula = True
            var_nombre_jaula = rs!vcha_jau_nombre
            rs.Close
            txt_cantidad_meta.Enabled = True
            txt_cantidad_meta.SetFocus
         Else
            var_guardar_jaula = False
            var_nombre_jaula = ""
            rs.Close
            MsgBox "N?mero de Jaula Incorrecta", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub
Private Sub integrantes()
   If Trim(txt_jaula) <> "" Then
      lv_integrantes.ListItems.Clear
      rs.Open "select a.vcha_per_personal_id,a.vcha_per_nombre, b.vcha_equ_puesto from tb_personal a, tb_equipos b where a.vcha_per_personal_id = b.vcha_per_personal_id and b.dtim_equ_fecha = '" + txt_fecha + "' and b.INTE_JAU_JAULA_ID = " + txt_jaula, cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_integrante = rs!vcha_per_personal_id
         cmb_integrantes.Text = rs!vcha_per_nombre
         numero_items_integrantes = 0
         While Not rs.EOF
            Set list_item = lv_integrantes.ListItems.Add(, , rs!vcha_per_personal_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_per_nombre), "", rs!vcha_per_nombre)
            list_item.SubItems(2) = IIf(IsNull(rs!vcha_equ_puesto), "", rs!vcha_equ_puesto)
            numero_items_integrantes = numero_items_integrantes + 1
            rs.MoveNext
         Wend
      End If
      rs.Close
   End If
End Sub

Private Sub txt_orden_surtido_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
     If Trim(txt_orden_surtido) <> "" Then
        txt_orden_surtido.Enabled = False
        com_guardar_orden.SetFocus
     End If
   End If
End Sub

Private Sub ordenes_surtido()
   If Trim(txt_jaula) <> "" Then
      rs.Open "select a.inte_ped_numero,a.inte_ors_orden_surtido,a.vcha_age_nombre,a.cantidad from vw_suma_cantidad_surtir a, tb_relacion_equipos b where a.inte_ors_orden_surtido = b.inte_ors_orden_surtido and b.inte_jau_jaula_id = " + txt_jaula + " and dtim_equ_fecha = '" + txt_fecha + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         lv_orden_surtido.ListItems.Clear
         var_cantidad_surtir = 0
         numero_items_ordenes = 0
         txt_orden_surtido = rs!inte_ors_orden_surtido
         While Not rs.EOF
            Dim list_item As ListItem
            Set list_item = lv_orden_surtido.ListItems.Add(, , rs!inte_ped_numero)
            list_item.SubItems(1) = IIf(IsNull(rs!inte_ors_orden_surtido), 0, rs!inte_ors_orden_surtido)
            list_item.SubItems(2) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            list_item.SubItems(3) = Format(IIf(IsNull(rs!Cantidad), 0, rs!Cantidad), "###,###,##0.00")
            var_cantidad_surtir = var_cantidad_surtir + IIf(IsNull(rs!Cantidad), 0, rs!Cantidad)
            numero_items_ordenes = numero_items_ordenes + 1
            rs.MoveNext
         Wend
      Else
         lv_orden_surtido.ListItems.Clear
         var_cantidad_surtir = 0
         numero_items_ordenes = 0
         txt_orden_surtido = ""
      End If
      rs.Close
      txt_cantidad_total = Format(var_cantidad_surtir, "###,###,##0.00")
      If numero_items_ordenes > 0 Then
         com_eliminar_orden.Enabled = True
      Else
         com_eliminar_orden.Enabled = False
      End If
   End If
End Sub

