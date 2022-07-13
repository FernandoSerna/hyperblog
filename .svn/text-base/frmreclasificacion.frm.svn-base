VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmreclasificacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reclasificacion de artículos"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_almacen 
      Height          =   285
      Left            =   6885
      TabIndex        =   21
      Top             =   1920
      Width           =   630
   End
   Begin VB.TextBox txt_buscar 
      Height          =   315
      Left            =   2055
      TabIndex        =   16
      Top             =   2505
      Width           =   1350
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2835
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   14
      Top             =   60
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   " Artículo "
      Height          =   1905
      Left            =   195
      TabIndex        =   10
      Top             =   405
      Width           =   5655
      Begin VB.TextBox txt_descripcion_general 
         Height          =   315
         Left            =   1305
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1440
         Width           =   4155
      End
      Begin VB.TextBox txt_descripcion 
         Height          =   315
         Left            =   1305
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         Top             =   570
         Width           =   4155
      End
      Begin VB.TextBox txt_codigo 
         Height          =   315
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   4
         Top             =   240
         Width           =   2070
      End
      Begin VB.TextBox txt_codigo_general 
         Height          =   315
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1110
         Width           =   2070
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   1500
         Width           =   885
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   630
         Width           =   885
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Código General:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   300
         Width           =   1140
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   1170
         Width           =   540
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4320
      Left            =   195
      TabIndex        =   8
      Top             =   2925
      Width           =   5655
      Begin MSComctlLib.ListView lv_reclasificacion 
         Height          =   4095
         Left            =   30
         TabIndex        =   9
         Top             =   150
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   7223
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
            Text            =   "Código"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "interno"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5505
      Picture         =   "frmreclasificacion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   840
      Picture         =   "frmreclasificacion.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   510
      Picture         =   "frmreclasificacion.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   180
      Picture         =   "frmreclasificacion.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   0
      Width           =   330
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   5115
      Top             =   1110
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
            Picture         =   "frmreclasificacion.frx":0940
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmreclasificacion.frx":121A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   1065
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
            Picture         =   "frmreclasificacion.frx":1AF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmreclasificacion.frx":23CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmreclasificacion.frx":2CA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmreclasificacion.frx":3244
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmreclasificacion.frx":3B20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmreclasificacion.frx":43FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmreclasificacion.frx":4CD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmreclasificacion.frx":4DE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmreclasificacion.frx":4EF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmreclasificacion.frx":500A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmreclasificacion.frx":511C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   600
      Left            =   195
      TabIndex        =   17
      Top             =   2310
      Width           =   5655
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3600
         TabIndex        =   18
         Top             =   195
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
         Caption         =   "Busqueda del Código:"
         Height          =   195
         Left            =   165
         TabIndex        =   19
         Top             =   255
         Width           =   1560
      End
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Index           =   0
      Left            =   180
      TabIndex        =   15
      Top             =   270
      Width           =   5655
   End
End
Attribute VB_Name = "frmreclasificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_eliminar_Click()
   If Trim(Me.txt_codigo) <> "" Then
      var_si = MsgBox("¿Desea eliminar el registro?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         var_si = MsgBox("Confirmar la eliminación del registro", vbYesNo, "ATENCION")
         If var_si = 6 Then
            rs.Open "DELETE FROM TB_RECLASIFICACION WHERE VCHA_ALM_ALMACEN_ID = '" + Me.txt_almacen + "' AND VCHA_ART_ARTICULO_ID = '" + Me.txt_codigo + "' AND VCHA_REC_CODIGO_GENERAL = '" + Me.txt_codigo_general + "'", cnn, adOpenDynamic, adLockOptimistic
            Me.lv_reclasificacion.ListItems.Remove (Me.lv_reclasificacion.selectedItem.Index)
         End If
      End If
   Else
      MsgBox "No se a seleccionado un registro", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_guardar_Click()
   Dim verificador As Integer
   rsaux.Open "select * from tb_articulos where vcha_art_articulo_id = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rsaux.EOF Then
      rsaux2.Open "select * from tb_articulos where vcha_art_articulo_id = '" + Me.txt_codigo_general + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux2.EOF Then
         var_codigo_1 = Mid(Me.txt_codigo_general, 1, 10)
         For var_i = 0 To 9
             var_codigo = var_codigo_1 + Trim(CStr(var_i))
             sum1 = 0
             sum2 = 0
             mcodigo = var_codigo
             longitud = Len(mcodigo)
             For icont = 1 To longitud
                 If ((icont / 2) - Int((icont / 2))) = 0 Then
                    sum2 = sum2 + Val(Mid(mcodigo, icont, 1))
                 Else
                    sum1 = sum1 + Val(Mid(mcodigo, icont, 1))
                 End If
             Next icont
             msuma = sum1 * 13 + sum2
             verificador = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
             If verificador = 10 Then
                verificador = 0
             End If
             var_codigo = var_codigo + Trim(CStr(verificador))
             
             rs.Open "SELECT * from tb_reclasificacion where vcha_alm_almacen_id = '" + txt_almacen + "' and vcha_art_articulo_id = '" + var_codigo + "' and vcha_rec_codigo_general = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
             If rs.EOF Then
                rsaux3.Open "INSERT INTO TB_RECLASIFICACION (VCHA_ALM_ALMACEN_ID, vcha_Art_articulo_id, vcha_rec_codigo_general) values ('" + txt_almacen + "','" + var_codigo + "','" + Me.txt_codigo + "')", cnn, adOpenDynamic, adLockOptimistic
                Set list_item = Me.lv_reclasificacion.ListItems.Add(, , var_codigo)
                list_item.SubItems(1) = Me.txt_descripcion_general
             End If
             rs.Close
         Next var_i
      Else
         MsgBox "El código del artículo no existe", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "El código general no existe", vbOKOnly, "ATENCION"
   End If
   rsaux.Close
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_codigo = ""
   Me.txt_codigo_general = ""
   Me.txt_descripcion = ""
   Me.txt_descripcion_general = ""
   Me.txt_codigo.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 2500
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call activa_forma(var_activa_forma_reporte_valuacion_devoluciones)
End Sub

Private Sub lv_reclasificacion_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_reclasificacion, ColumnHeader)
End Sub

Private Sub lv_reclasificacion_ItemClick(ByVal Item As MSComctlLib.ListItem)
   If lv_reclasificacion.ListItems.Count > 0 Then
      Me.txt_codigo_general = Me.lv_reclasificacion.selectedItem
      Me.txt_descripcion_general = Me.lv_reclasificacion.selectedItem.SubItems(1)
   End If
End Sub

Private Sub txt_codigo_general_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_codigo_general_LostFocus()
   If Trim(Me.txt_codigo_general) <> "" Then
      var_descuento = Mid(Me.txt_codigo_general, 11, 1)
      If var_descuento = "0" Then
         rs.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + Me.txt_codigo_general + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_descripcion_general = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
         Else
            MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         MsgBox "El artículo debe de ser de primera", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_codigo_LostFocus()
   If Trim(txt_codigo) = "" Then
      Me.txt_descripcion = ""
      Me.lv_reclasificacion.ListItems.Clear
   Else
      lv_reclasificacion.ListItems.Clear
      var_descuento = Mid(Me.txt_codigo, 11, 1)
      If var_descuento = "0" Then
         rs.Open "select * from tb_articulos where vcha_art_Articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_descripcion = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
            rsaux.Open "select distinct * from vw_Reclasificacion where vcha_rec_codigo_general = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               numero_items_ALMACENES = 0
               While Not rsaux.EOF
                     Set list_item = Me.lv_reclasificacion.ListItems.Add(, , rsaux!vcha_art_articulo_id)
                     list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_art_nombre_español), "", rsaux!vcha_art_nombre_español)
                     rsaux.MoveNext:
               Wend
               If lv_reclasificacion.ListItems.Count > 18 Then
                  lv_reclasificacion.ColumnHeaders(2).Width = 3750
               Else
                  lv_reclasificacion.ColumnHeaders(2).Width = 4100
               End If
            End If
            rsaux.Close
         Else
            MsgBox "Código de artículo incorrecto", vbOKOnly, "ATENCION"
            Me.txt_codigo = ""
            Me.txt_descripcion = ""
            Me.lv_reclasificacion.ListItems.Clear
         End If
         rs.Close
      End If
   End If
End Sub

Private Sub txt_descripcion_general_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub

Private Sub txt_descripcion_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub
