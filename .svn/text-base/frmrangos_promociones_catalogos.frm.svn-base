VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmrangos_promociones_catalogos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rangos para promociones por vencimiento de catalogos"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   5955
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmrangos_promociones_catalogos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      Picture         =   "frmrangos_promociones_catalogos.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Guardar Alt + G"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   810
      Picture         =   "frmrangos_promociones_catalogos.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      Picture         =   "frmrangos_promociones_catalogos.frx":02D6
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1470
      Picture         =   "frmrangos_promociones_catalogos.frx":03D8
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5460
      Picture         =   "frmrangos_promociones_catalogos.frx":04DA
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2865
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   16
      Top             =   30
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   "  Canal de Venta "
      Height          =   615
      Left            =   150
      TabIndex        =   15
      Top             =   420
      Width           =   5655
      Begin VB.TextBox txt_nombre_canal_venta 
         Height          =   315
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   210
         Width           =   4305
      End
      Begin VB.TextBox txt_canal_venta 
         Height          =   315
         Left            =   225
         TabIndex        =   7
         Top             =   210
         Width           =   1005
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2490
      Left            =   150
      TabIndex        =   14
      Top             =   1920
      Width           =   5655
      Begin MSComctlLib.ListView lv_rangos 
         Height          =   2235
         Left            =   45
         TabIndex        =   11
         Top             =   165
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   3942
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
            Text            =   "L. Inferior"
            Object.Width           =   3228
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "L. Superior"
            Object.Width           =   3228
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descuento"
            Object.Width           =   3228
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   45
         Top             =   2505
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
               Picture         =   "frmrangos_promociones_catalogos.frx":0B14
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmrangos_promociones_catalogos.frx":13EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmrangos_promociones_catalogos.frx":1CC8
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmrangos_promociones_catalogos.frx":2264
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmrangos_promociones_catalogos.frx":2B40
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmrangos_promociones_catalogos.frx":341A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmrangos_promociones_catalogos.frx":3CF4
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmrangos_promociones_catalogos.frx":3E06
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmrangos_promociones_catalogos.frx":3F18
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmrangos_promociones_catalogos.frx":402A
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmrangos_promociones_catalogos.frx":413C
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   " Rangos "
      Height          =   825
      Left            =   150
      TabIndex        =   0
      Top             =   1080
      Width           =   5655
      Begin VB.TextBox txt_descuento 
         Height          =   315
         Left            =   4650
         TabIndex        =   10
         Top             =   300
         Width           =   705
      End
      Begin VB.TextBox txt_limite_inferior 
         Height          =   315
         Left            =   825
         TabIndex        =   8
         Top             =   300
         Width           =   705
      End
      Begin VB.TextBox txt_limite_superior 
         Height          =   315
         Left            =   2505
         TabIndex        =   9
         Top             =   300
         Width           =   705
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Index           =   3
         Left            =   5400
         TabIndex        =   19
         Top             =   360
         Width           =   120
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Descuento:"
         Height          =   195
         Index           =   1
         Left            =   3765
         TabIndex        =   18
         Top             =   360
         Width           =   825
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Inferior:"
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   13
         Top             =   360
         Width           =   525
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Superior:"
         Height          =   195
         Index           =   0
         Left            =   1770
         TabIndex        =   12
         Top             =   360
         Width           =   630
      End
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   17
      Top             =   285
      Width           =   5655
   End
End
Attribute VB_Name = "frmrangos_promociones_catalogos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_deshacer_Click()
   i = 0
End Sub

Private Sub cmd_eliminar_Click()
   Dim si As Integer
   si = MsgBox("¿Deseas eliminar el registro?", vbYesNo, "ATENCION")
   If si = 6 Then
      rs.Open "delete from tb_rangos_promociones_catalogos where vcha_can_canal_venta_id = '" + txt_canal_venta + "'  and inte_rpr_limite_inferior =  " + txt_limite_inferior + " and inte_rpr_limite_superior = " + txt_limite_superior + " and floa_rpr_descuento = " + txt_descuento, cnn, adOpenDynamic, adLockOptimistic
      lv_rangos.ListItems.Remove (lv_rangos.selectedItem.Index)
      If lv_rangos.ListItems.Count > 0 Then
         lv_rangos.selectedItem.Selected = True
      Else
         txt_limite_inferior = ""
         txt_limite_superior = ""
         txt_descuento = ""
      End If
   End If
End Sub

Private Sub cmd_guardar_Click()
   If Trim(txt_canal_venta) <> "" Then
      If IsNumeric(txt_limite_inferior) Then
         If IsNumeric(txt_limite_superior) Then
            If IsNumeric(txt_descuento) Then
               rs.Open "select * from TB_RANGOS_PROMOCIONES_CATALOGOS where vcha_Can_Canal_venta_id = '" + txt_canal_venta + "' and INTE_RPR_LIMITE_INFERIOR = " + txt_limite_inferior + " and INTE_RPR_LIMITE_SUPERIOR = " + txt_limite_superior + " and FLOA_RPR_DESCUENTO = " + txt_descuento, cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  rsaux2.Open "updtae tb_rangos_promociones_catalogos set floa_rpr_descuento = " + txt_descuento + " where vcha_can_canal_Venta_id = '" + txt_canal_venta + "' and inte_rpr_limite_inferior = " + txt_limite_inferior + " and inte_rpr_limite_superio = " + txt_limite_superior, cnn, adOpenDynamic, adLockOptimistic
                  lv_rangos.selectedItem = txt_limite_inferior
                  lv_rangos.selectedItem.SubItems(1) = txt_limite_superior
                  lv_rangos.selectedItem.SubItems(2) = txt_descuento
               Else
                  rsaux2.Open "insert into tb_rangos_promociones_catalogos (vcha_can_canal_venta_id, inte_rpr_limite_inferior, inte_rpr_limite_superior, floa_rpr_descuento) values ('" + txt_canal_venta + "', " + txt_limite_inferior + ", " + txt_limite_superior + ", " + txt_descuento + ")", cnn, adOpenDynamic, adLockOptimistic
                  Set list_item = lv_rangos.ListItems.Add(, , txt_limite_inferior)
                  list_item.SubItems(1) = txt_limite_superior
                  list_item.SubItems(2) = txt_descuento
               End If
               rs.Close
            Else
               MsgBox "Descuento Incorrecto", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Limite Superior Incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Limite Inferior Incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a seleccionado un canal de venta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
  i = 0
End Sub

Private Sub cmd_nuevo_Click()
   txt_canal_venta.Enabled = True
   txt_nombre_canal_venta.Enabled = True
   txt_canal_venta.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
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
   var_cadena_seguridad = ""
   Top = 1000
   Left = 2900
   txt_canal_venta.Enabled = False
   txt_nombre_canal_venta.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   var_swpassword = False
   var_modifica_registro = False
   Call activa_forma(var_activa_forma_rangos_promociones_catalogos)
End Sub

Private Sub txt_canal_venta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      txt_nombre_canal_venta.SetFocus
   End If
End Sub

Private Sub txt_canal_venta_LostFocus()
Dim list_item As ListItem
Dim numero_items_rangos As Integer
   If Trim(txt_canal_venta) <> "" Then
      rs.Open "select * from tb_canalesventas where vcha_can_canal_venta_id = '" + txt_canal_venta + "'", cnn, adOpenDynamic, adLockOptimistic
      lv_rangos.ListItems.Clear
      If Not rs.EOF Then
         txt_nombre_canal_venta = rs!vcha_can_nombre
         rsaux2.Open "select * from tb_rangos_promociones_catalogos where vcha_Can_canal_Venta_id = '" + txt_canal_venta + "'", cnn, adOpenDynamic, adLockOptimistic
         numero_items_rangos = 0
         If Not rsaux2.EOF Then
            txt_limite_inferior = IIf(IsNull(rsaux2!INTE_RPR_LIMITE_INFERIOR), 0, rsaux2!INTE_RPR_LIMITE_INFERIOR)
            txt_limite_superior = IIf(IsNull(rsaux2!INTE_RPR_LIMITE_SUPERIOR), 0, rsaux2!INTE_RPR_LIMITE_SUPERIOR)
            txt_descuento = IIf(IsNull(rsaux2!FLOA_RPR_DESCUENTO), 0, rsaux2!FLOA_RPR_DESCUENTO)
            While Not rsaux2.EOF
                  Set list_item = lv_rangos.ListItems.Add(, , rsaux2!INTE_RPR_LIMITE_INFERIOR)
                  list_item.SubItems(1) = IIf(IsNull(rsaux2!INTE_RPR_LIMITE_SUPERIOR), 0, rsaux2!INTE_RPR_LIMITE_SUPERIOR)
                  list_item.SubItems(2) = IIf(IsNull(rsaux2!FLOA_RPR_DESCUENTO), 0, rsaux2!FLOA_RPR_DESCUENTO)
                  rsaux2.MoveNext:
                  numero_items_rangos = numero_items_rangos + 1
            Wend
         End If
         rsaux2.Close
      Else
         MsgBox "Clave de canal de venta incorrecta", vbOKOnly, "ATENCION"
         txt_canal_venta = ""
         txt_nombre_canal_venta = ""
      End If
      rs.Close
   End If
End Sub

Private Sub txt_limite_inferior_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_limite_superior.SetFocus
   End If
End Sub

Private Sub txt_limite_superior_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_descuento.SetFocus
   End If
End Sub

Private Sub txt_nombre_canal_venta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_limite_inferior.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub
