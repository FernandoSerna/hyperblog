VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcostos_predeterminados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Costos Predeterminados"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   6975
   Begin VB.Frame frm_lista 
      Height          =   2625
      Left            =   750
      TabIndex        =   22
      Top             =   390
      Width           =   5820
      Begin MSComctlLib.ListView lv_lista 
         Height          =   2145
         Left            =   30
         TabIndex        =   24
         Top             =   420
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   3784
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
            Text            =   "Nombre"
            Object.Width           =   10107
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Clave"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   30
         TabIndex        =   23
         Top             =   120
         Width           =   5745
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   75
      TabIndex        =   18
      Top             =   1860
      Width           =   6780
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   1815
         TabIndex        =   19
         Top             =   150
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   4470
         TabIndex        =   20
         Top             =   165
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ir al primero"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Un Registro Atras"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Un registro adelante"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ir al ultimo"
               ImageIndex      =   10
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Busqueda de artículo:"
         Height          =   195
         Left            =   195
         TabIndex        =   21
         Top             =   195
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      Picture         =   "frmcostos_predeterminados.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   390
      Picture         =   "frmcostos_predeterminados.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   720
      Picture         =   "frmcostos_predeterminados.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1050
      Picture         =   "frmcostos_predeterminados.frx":02D6
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1380
      Picture         =   "frmcostos_predeterminados.frx":03D8
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6540
      Picture         =   "frmcostos_predeterminados.frx":04DA
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Caption         =   " Costo Predeterminado "
      Height          =   1440
      Left            =   75
      TabIndex        =   13
      Top             =   390
      Width           =   6825
      Begin VB.TextBox txt_costo 
         Height          =   315
         Left            =   1140
         MaxLength       =   50
         TabIndex        =   11
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txt_nombre_proveedor 
         Height          =   315
         Left            =   2610
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         Top             =   300
         Width           =   4095
      End
      Begin VB.TextBox txt_nombre_articulo 
         Height          =   315
         Left            =   2610
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   10
         Top             =   630
         Width           =   4095
      End
      Begin VB.TextBox txt_clave_articulo 
         Height          =   315
         Left            =   1140
         MaxLength       =   50
         TabIndex        =   9
         Top             =   630
         Width           =   1455
      End
      Begin VB.TextBox txt_clave_proveedor 
         Height          =   315
         Left            =   1140
         MaxLength       =   50
         TabIndex        =   7
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Costo:"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   16
         Top             =   1020
         Width           =   450
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Artículo:"
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   15
         Top             =   690
         Width           =   600
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor:"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   14
         Top             =   360
         Width           =   780
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3630
      Left            =   90
      TabIndex        =   0
      Top             =   2415
      Width           =   6765
      Begin MSComctlLib.ListView lv_articulos 
         Height          =   3450
         Left            =   45
         TabIndex        =   12
         Top             =   135
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   6085
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
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
            Text            =   "Código"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "costo"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1770
      Top             =   -210
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
            Picture         =   "frmcostos_predeterminados.frx":0B14
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcostos_predeterminados.frx":13EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcostos_predeterminados.frx":1CC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcostos_predeterminados.frx":2264
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcostos_predeterminados.frx":2B40
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcostos_predeterminados.frx":341A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcostos_predeterminados.frx":3CF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcostos_predeterminados.frx":3E06
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcostos_predeterminados.frx":3F18
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcostos_predeterminados.frx":402A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcostos_predeterminados.frx":413C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   60
      TabIndex        =   17
      Top             =   255
      Width           =   6825
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   120
      Top             =   4920
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
            Picture         =   "frmcostos_predeterminados.frx":424E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcostos_predeterminados.frx":4B28
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   -45
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcostos_predeterminados.frx":5402
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcostos_predeterminados.frx":5CDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcostos_predeterminados.frx":65B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcostos_predeterminados.frx":6B52
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcostos_predeterminados.frx":742E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcostos_predeterminados.frx":7D08
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcostos_predeterminados.frx":85E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcostos_predeterminados.frx":86F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcostos_predeterminados.frx":8806
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcostos_predeterminados.frx":8918
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmcostos_predeterminados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tipo_lista As Integer
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report


Private Sub cmd_deshacer_Click()
   Me.txt_clave_articulo = ""
   Me.txt_nombre_articulo = ""
   Me.txt_clave_proveedor = ""
   Me.txt_nombre_proveedor = ""
End Sub

Private Sub cmd_eliminar_Click()
   Dim si As Integer
   Dim var_n
   If Trim(txt_clave_proveedor) <> "" Then
      If Trim(txt_clave_articulo) <> "" Then
         si = MsgBox("¿Desea eliminar el registro?", vbYesNo, "ATENCION")
         If si = 6 Then
            rs.Open "delete from tb_costos_predeterminados where vcha_pro_proveedor_id = '" + txt_clave_proveedor + "' and vcha_art_articulo_id = '" + txt_clave_articulo + "'", cnn, adOpenDynamic, adLockOptimistic
            lv_articulos.ListItems.Remove (lv_articulos.selectedItem.Index)
            var_n = lv_articulos.ListItems.Count
            If var_n > 0 Then
               lv_articulos.SetFocus
            Else
               txt_clave_articulo = ""
               txt_nombre_articulo = ""
               txt_costo = ""
            End If
         Else
            MsgBox "Se a cancelado la eliminación del registro", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Clave de artículo incorrecta", vbOKOnly, "ATENION"
      End If
   Else
      MsgBox "Clave de proveedor Incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_guardar_Click()
Dim si As Integer
   If Trim(txt_clave_proveedor) <> "" Then
      If Trim(txt_clave_articulo) <> "" Then
         If IsNumeric(txt_costo) Then
            rs.Open "select * from tb_costos_predeterminados where vcha_pro_proveedor_id = '" + txt_clave_proveedor + "' and vcha_art_articulo_id = '" + txt_clave_articulo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               si = MsgBox("Desea aplicar los cambios correspondientes", vbYesNo, "ATENCION")
               If si = 6 Then
                  rsaux.Open "update tb_costos_predeterminados set FLOA_CPR_COSTO_PREDETERMINADO  = " + txt_costo + " where vcha_pro_proveedor_id = '" + txt_clave_proveedor + "' and vcha_art_articulo_id = '" + txt_clave_articulo + "'", cnn, adOpenDynamic, adLockOptimistic
                  valor = Trim(txt_clave_articulo)
                  Set itmfound = lv_articulos.findItem(valor, lvwText, , lvwPartial)
                  itmfound.EnsureVisible
                  itmfound.Selected = True
                  lv_articulos.selectedItem.SubItems(2) = Format(txt_costo, "###,###,##0.00")
               Else
                  MsgBox "Se a cancelado la actualización", vbOKOnly, "ATENCION"
               End If
            Else
               si = MsgBox("Desea Insertar el registro correspondiente", vbYesNo, "ATENCION")
               If si = 6 Then
                  rsaux.Open "insert into tb_costos_predeterminados (vcha_pro_proveedor_id, vcha_art_articulo_id, floa_cpr_costo_predeterminado) values ('" + txt_clave_proveedor + "', '" + txt_clave_articulo + "', " + txt_costo + ")", cnn, adOpenDynamic, adLockOptimistic
                  Set list_item = lv_articulos.ListItems.Add(, , txt_clave_articulo)
                  list_item.SubItems(1) = txt_nombre_articulo
                  list_item.SubItems(2) = Format(txt_costo, "###,###,##0.00")
               Else
                  MsgBox "Se a cancelado la inserción", vbOKOnly, "ATENCION"
               End If
            End If
            rs.Close
         Else
            MsgBox "Costo incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Clave de artículo incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Clave de proveedor incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
   Set reporte = appl.OpenReport(App.Path + "\rep_costos_predeterminados.rpt")
   frmvistasprevias.cr.ReportSource = reporte
   For ntablas = 1 To reporte.Database.Tables.Count
       reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
   Next ntablas
   frmvistasprevias.cr.ViewReport
   frmvistasprevias.Caption = "Reporte de Costos Predeterminados"
   frmvistasprevias.Show 1
   Set reporte = Nothing
End Sub

Private Sub cmd_nuevo_Click()
   Call pro_limpiatextos(Me)
   Me.txt_clave_proveedor.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
   If Shift = 4 And KeyCode = 78 Then
      cmd_nuevo_Click
   End If
   If Shift = 4 And KeyCode = 71 Then
      cmd_guardar_Click
   End If
   If Shift = 4 And KeyCode = 68 Then
      cmd_deshacer_Click
   End If
   If Shift = 4 And KeyCode = 69 Then
      cmd_eliminar_Click
   End If
   If Shift = 4 And KeyCode = 73 Then
      cmd_imprimir_Click
   End If
End Sub

Private Sub Form_Load()
   If Trim(frmordenescompra.txt_proveedor) <> "" Then
      Me.txt_clave_proveedor = frmordenescompra.txt_proveedor
   End If
   If Trim(frmordenescompra.txt_codigo) <> "" Then
      Me.txt_clave_articulo = frmordenescompra.txt_codigo
   End If
   var_cadena_seguridad = ""
   Top = 800
   Left = 2200
   frm_lista.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    var_swpassword = False
    var_modifica_registro = False
    Call activa_forma(var_activa_forma_costos_predeterminados)
End Sub

Private Sub lv_articulos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_articulos, ColumnHeader)
End Sub

Private Sub lv_articulos_GotFocus()
   Dim var_n As Integer
   var_n = lv_articulos.ListItems.Count
   If var_n > 0 Then
      txt_clave_articulo = lv_articulos.selectedItem
      txt_nombre_articulo = lv_articulos.selectedItem.SubItems(1)
      txt_costo = lv_articulos.selectedItem.SubItems(2)
   End If
End Sub

Private Sub lv_articulos_ItemClick(ByVal Item As MSComctlLib.ListItem)
   txt_clave_articulo = lv_articulos.selectedItem
   txt_nombre_articulo = lv_articulos.selectedItem.SubItems(1)
   txt_costo = lv_articulos.selectedItem.SubItems(2)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_tipo_lista = 1 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_clave_proveedor = lv_lista.selectedItem.SubItems(1)
         Else
            txt_clave_proveedor = ""
         End If
         frm_lista.Visible = False
         txt_clave_proveedor.SetFocus
      End If
      If var_tipo_lista = 2 Then
         If lv_lista.ListItems.Count > 0 Then
            txt_clave_articulo = lv_lista.selectedItem.SubItems(1)
         Else
            txt_clave_articulo = ""
         End If
         frm_lista.Visible = False
         txt_clave_articulo.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      If var_tipo_lista = 1 Then
         frm_lista.Visible = False
         txt_clave_proveedor.SetFocus
      End If
      If var_tipo_lista = 2 Then
         frm_lista.Visible = False
         txt_clave_articulo.SetFocus
      End If
   End If
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_articulos.SetFocus
      Call pro_avanzar(Me, lv_articulos, Button)
      lv_articulos.selectedItem.EnsureVisible
      txt_clave_articulo = lv_articulos.selectedItem
      txt_nombre_articulo = lv_articulos.selectedItem.SubItems(1)
      txt_costo = lv_articulos.selectedItem.SubItems(2)
   End If
   If Button.Index = 1 Then
      lv_articulos.ListItems(1).Selected = True
      txt_clave_articulo = lv_articulos.selectedItem
      txt_nombre_articulo = lv_articulos.selectedItem.SubItems(1)
      txt_costo = lv_articulos.selectedItem.SubItems(2)
      lv_articulos.selectedItem.EnsureVisible
   End If
   If Button.Index = 4 Then
      numero_items_colores = lv_articulos.ListItems.Count
      lv_articulos.ListItems(numero_items_colores).Selected = True
      lv_articulos.selectedItem.EnsureVisible
      txt_clave_articulo = lv_articulos.selectedItem
      txt_nombre_articulo = lv_articulos.selectedItem.SubItems(1)
      txt_costo = lv_articulos.selectedItem.SubItems(2)
   End If
err0:
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(txt_clave_proveedor) <> "" Then
         If Trim(txt_buscar) <> "" Then
            rs.Open "seelct * from tb_costos_predeterminados where vcha_pro_proveedor_id = '" + txt_clave_porveedor + "' and vcha_art_articulo_id = '" + txt_buscar + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               valor = Trim(txt_buscar)
               Set itmfound = lv_articulos.findItem(valor, lvwText, , lvwPartial)
               itmfound.EnsureVisible
               itmfound.Selected = True
            End If
            rs.Close
         Else
            MsgBox "No se a indicado algun artículo", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "No se a seleccionado un proveedor", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_clave_articulo_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_clave_articulo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      var_tipo_lista = 2
      Dim list_item As ListItem
      rs.Open "select vcha_art_nombre_Español, vcha_art_Articulo_id from TB_articulos order by vcha_art_nombre_español", cnn, adOpenDynamic, adLockOptimistic
      numero_items = 0
      While Not rs.EOF
         Set list_item = lv_lista.ListItems.Add(, , rs(0).Value)
         list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
         rs.MoveNext:
         numero_items = numero_items + 1
      Wend
      rs.Close
      If numero_items > 8 Then
         lv_lista.ColumnHeaders(1).Width = 5430
      Else
         lv_lista.ColumnHeaders(1).Width = 5630
      End If
      lbl_lista = "Lista de Artículos"
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_clave_articulo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_costo.SetFocus
   End If
End Sub

Private Sub txt_clave_articulo_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_clave_articulo) <> "" Then
      rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_clave_articulo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_articulo = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
      Else
         txt_nombre_articulo = ""
         MsgBox "Clave de artículo incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      txt_nombre_articulo = ""
   End If
End Sub

Private Sub txt_clave_proveedor_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_clave_proveedor_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      var_tipo_lista = 1
      Dim list_item As ListItem
      rs.Open "select vcha_pro_nombre, vcha_pro_proveedor_id from TB_proveedores order by vcha_pro_nombre", cnn, adOpenDynamic, adLockOptimistic
      numero_items = 0
      While Not rs.EOF
         Set list_item = lv_lista.ListItems.Add(, , rs(0).Value)
         list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
         rs.MoveNext:
         numero_items = numero_items + 1
      Wend
      rs.Close
      If numero_items > 8 Then
         lv_lista.ColumnHeaders(1).Width = 5430
      Else
         lv_lista.ColumnHeaders(1).Width = 5630
      End If
      lbl_lista = "Lista de Proveedores"
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_clave_proveedor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_clave_proveedor_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_clave_proveedor) <> "" Then
      rs.Open "select * from tb_proveedores where vcha_pro_proveedor_id = '" + txt_clave_proveedor + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_proveedor = IIf(IsNull(rs!VCHA_PRO_NOMBRE), "", rs!VCHA_PRO_NOMBRE)
         lv_articulos.ListItems.Clear
         rsaux.Open "Select * from vw_costos_predeterminados where vcha_pro_proveedor_id = '" + txt_clave_proveedor + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            While Not rsaux.EOF
                  Set list_item = lv_articulos.ListItems.Add(, , rsaux!vcha_Art_articulo_id)
                  list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_art_nombre_español), "", rsaux!vcha_art_nombre_español)
                  list_item.SubItems(2) = Format(IIf(IsNull(rsaux!floa_cpr_costo_predeterminado), 0, rsaux!floa_cpr_costo_predeterminado), "###,###,##0.00")
                  rsaux.MoveNext
            Wend
         End If
         rsaux.Close
      Else
         txt_nombre_proveedor = ""
         MsgBox "Clave de proveedor incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      txt_nombre_proveedor = ""
   End If
End Sub

Private Sub txt_costo_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If Me.cmd_guardar.Enabled = True Then
         Me.cmd_guardar.SetFocus
      End If
   End If
End Sub

Private Sub txt_costo_LostFocus()
   If Not IsNumeric(txt_costo) Then
      MsgBox "Costo Incorrecto", vbOKOnly, "ATENCION"
      txt_costo = ""
   End If
End Sub

Private Sub txt_nombre_articulo_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_nombre_articulo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      var_tipo_lista = 2
      Dim list_item As ListItem
      rs.Open "select vcha_art_nombre_Español, vcha_art_Articulo_id from TB_articulos order by vcha_art_nombre_español", cnn, adOpenDynamic, adLockOptimistic
      numero_items = 0
      While Not rs.EOF
         Set list_item = lv_lista.ListItems.Add(, , rs(0).Value)
         list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
         rs.MoveNext:
         numero_items = numero_items + 1
      Wend
      rs.Close
      If numero_items > 8 Then
         lv_lista.ColumnHeaders(1).Width = 5430
      Else
         lv_lista.ColumnHeaders(1).Width = 5630
      End If
      lbl_lista = "Lista de Artículos"
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_articulo_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_articulo_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub

Private Sub txt_nombre_proveedor_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible"
End Sub

Private Sub txt_nombre_proveedor_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      var_tipo_lista = 1
      Dim list_item As ListItem
      rs.Open "select vcha_pro_nombre, vcha_pro_proveedor_id from TB_proveedores order by vcha_pro_nombre", cnn, adOpenDynamic, adLockOptimistic
      numero_items = 0
      While Not rs.EOF
         Set list_item = lv_lista.ListItems.Add(, , rs(0).Value)
         list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
         rs.MoveNext:
         numero_items = numero_items + 1
      Wend
      rs.Close
      If numero_items > 8 Then
         lv_lista.ColumnHeaders(1).Width = 5430
      Else
         lv_lista.ColumnHeaders(1).Width = 5630
      End If
      lbl_lista = "Lista de Proveedores"
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_proveedor_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_proveedor_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub
