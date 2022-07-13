VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmclases_cartera 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clases de Movimientos en Cartera "
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5955
   Begin VB.Frame Frame3 
      Height          =   4785
      Left            =   150
      TabIndex        =   18
      Top             =   2430
      Width           =   5655
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   1635
         Top             =   135
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
               Picture         =   "frmclases_cartera.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmclases_cartera.frx":08DA
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmclases_cartera.frx":11B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmclases_cartera.frx":1750
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmclases_cartera.frx":202A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmclases_cartera.frx":2904
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmclases_cartera.frx":31DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmclases_cartera.frx":34F8
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmclases_cartera.frx":3812
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmclases_cartera.frx":3DAE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList icono_encabezado 
         Left            =   960
         Top             =   105
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
               Picture         =   "frmclases_cartera.frx":40C8
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmclases_cartera.frx":49A2
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lv_clases_cartera 
         Height          =   4590
         Left            =   45
         TabIndex        =   19
         Top             =   135
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   8096
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   6879
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ancho"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "alto"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "largo"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList3 
         Left            =   0
         Top             =   -15
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
               Picture         =   "frmclases_cartera.frx":527C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmclases_cartera.frx":5B56
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Clases de Movimientos en Cartera "
      Height          =   1440
      Left            =   150
      TabIndex        =   14
      Top             =   405
      Width           =   5655
      Begin VB.ComboBox cmb_documentos 
         Height          =   315
         ItemData        =   "frmclases_cartera.frx":6430
         Left            =   1845
         List            =   "frmclases_cartera.frx":6449
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   300
         Width           =   3600
      End
      Begin VB.TextBox txt_documento 
         Height          =   315
         Left            =   1140
         MaxLength       =   50
         TabIndex        =   6
         Top             =   300
         Width           =   690
      End
      Begin VB.TextBox txt_clase 
         Height          =   315
         Left            =   1140
         MaxLength       =   50
         TabIndex        =   8
         Top             =   630
         Width           =   690
      End
      Begin VB.TextBox txt_descripcion 
         Height          =   315
         Left            =   1140
         MaxLength       =   50
         TabIndex        =   9
         Top             =   960
         Width           =   4290
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Movimiento:"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   17
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clase:"
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   16
         Top             =   690
         Width           =   435
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   15
         Top             =   1020
         Width           =   885
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   150
      TabIndex        =   10
      Top             =   1875
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   1620
         TabIndex        =   11
         Top             =   150
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3600
         TabIndex        =   12
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
         Caption         =   "Busqueda de clase:"
         Height          =   195
         Left            =   195
         TabIndex        =   13
         Top             =   195
         Width           =   1410
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5460
      Picture         =   "frmclases_cartera.frx":64C4
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      Picture         =   "frmclases_cartera.frx":6AFE
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1110
      Picture         =   "frmclases_cartera.frx":6C00
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   795
      Picture         =   "frmclases_cartera.frx":6D02
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      Picture         =   "frmclases_cartera.frx":6DD4
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Guardar Alt + G"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmclases_cartera.frx":6ED6
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2595
      Top             =   975
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
            Picture         =   "frmclases_cartera.frx":6FD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclases_cartera.frx":78B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclases_cartera.frx":818C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclases_cartera.frx":8728
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclases_cartera.frx":9004
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclases_cartera.frx":98DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclases_cartera.frx":A1B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclases_cartera.frx":A2CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclases_cartera.frx":A3DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclases_cartera.frx":A4EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclases_cartera.frx":A600
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   20
      Top             =   285
      Width           =   5655
   End
End
Attribute VB_Name = "frmclases_cartera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_inserta As Boolean
Private Sub textos()
   txt_clase = lv_clases_cartera.selectedItem
   txt_descripcion = lv_clases_cartera.selectedItem.SubItems(1)
   var_inserta = False
End Sub
Private Sub cmb_documentos_Click()
   If cmb_documentos = "FACTURACION" Then
      txt_documento = "FA"
   Else
      If cmb_documentos = "NOTA DE CARGO" Then
         txt_documento = "NC"
      Else
         If cmb_documentos = "BONIFICACION" Then
            txt_documento = "BO"
         Else
            If cmb_documentos = "BONIFICACION FINANCIERA" Then
               txt_documento = "BF"
            Else
               If cmb_documentos = "DESCUENTO FINANCIERO" Then
                  txt_documento = "DF"
               Else
                  If cmb_documentos = "PAGO" Then
                     txt_documento = "PA"
                  Else
                     If cmb_documentos = "DEVOLUCION SOBRE VENTA" Then
                        txt_documento = "DV"
                     Else
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
End Sub

Private Sub cmb_documentos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub cmb_documentos_LostFocus()
Dim n As Integer
   If Trim(txt_documento) <> "" Then
      lv_clases_cartera.ListItems.Clear
      If Trim(txt_documento) = "FA" Then
         rs.Open "select * from tb_clases_cartera where vcha_Car_documento = 'FA'", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
            Set list_item = lv_clases_cartera.ListItems.Add(, , rs!vcha_car_clase_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_car_nombre), "", rs!vcha_car_nombre)
            rs.MoveNext
         Wend
         rs.Close
      Else
         If Trim(txt_documento) = "NC" Then
            rs.Open "select * from tb_clases_cartera where vcha_Car_documento = 'NC'", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
               Set list_item = lv_clases_cartera.ListItems.Add(, , rs!vcha_car_clase_id)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_car_nombre), "", rs!vcha_car_nombre)
               rs.MoveNext
            Wend
            rs.Close
         Else
            If Trim(txt_documento) = "BO" Then
               rs.Open "select * from tb_clases_cartera where vcha_Car_documento = 'BO'", cnn, adOpenDynamic, adLockOptimistic
               While Not rs.EOF
                  Set list_item = lv_clases_cartera.ListItems.Add(, , rs!vcha_car_clase_id)
                  list_item.SubItems(1) = IIf(IsNull(rs!vcha_car_nombre), "", rs!vcha_car_nombre)
                  rs.MoveNext
               Wend
               rs.Close
            Else
               If Trim(txt_documento) = "BF" Then
                  rs.Open "select * from tb_clases_cartera where vcha_Car_documento = 'BF'", cnn, adOpenDynamic, adLockOptimistic
                  While Not rs.EOF
                     Set list_item = lv_clases_cartera.ListItems.Add(, , rs!vcha_car_clase_id)
                     list_item.SubItems(1) = IIf(IsNull(rs!vcha_car_nombre), "", rs!vcha_car_nombre)
                     rs.MoveNext
                  Wend
                  rs.Close
               Else
                  If Trim(txt_documento) = "PA" Then
                     rs.Open "select * from tb_clases_cartera where vcha_Car_documento = 'PA'", cnn, adOpenDynamic, adLockOptimistic
                     While Not rs.EOF
                        Set list_item = lv_clases_cartera.ListItems.Add(, , rs!vcha_car_clase_id)
                        list_item.SubItems(1) = IIf(IsNull(rs!vcha_car_nombre), "", rs!vcha_car_nombre)
                        rs.MoveNext
                     Wend
                     rs.Close
                  Else
                     If Trim(txt_documento) = "DV" Then
                        rs.Open "select * from tb_clases_cartera where vcha_Car_documento = 'DV'", cnn, adOpenDynamic, adLockOptimistic
                        While Not rs.EOF
                           Set list_item = lv_clases_cartera.ListItems.Add(, , rs!vcha_car_clase_id)
                           list_item.SubItems(1) = IIf(IsNull(rs!vcha_car_nombre), "", rs!vcha_car_nombre)
                           rs.MoveNext
                        Wend
                        rs.Close
                     Else
                        If Trim(txt_documento) = "DF" Then
                           rs.Open "select * from tb_clases_cartera where vcha_Car_documento = 'DF'", cnn, adOpenDynamic, adLockOptimistic
                           While Not rs.EOF
                              Set list_item = lv_clases_cartera.ListItems.Add(, , rs!vcha_car_clase_id)
                              list_item.SubItems(1) = IIf(IsNull(rs!vcha_car_nombre), "", rs!vcha_car_nombre)
                              rs.MoveNext
                           Wend
                           rs.Close
                        Else
                           MsgBox "Clave de documento incorrecta", vbOKOnly, "ATENCION"
                        End If
                     End If
                  End If
               End If
            End If
         End If
      End If
      n = lv_clases_cartera.ListItems.Count
      If n > 0 And var_inserta = False Then
         Call textos
      Else
         txt_clase = ""
         txt_descripcion = ""
      End If
   End If
End Sub

Private Sub cmd_deshacer_Click()
Dim n As Integer
   var_inserta = False
   n = lv_clases_cartera.ListItems.Count
   If n > 0 And var_inserta = False Then
      Call textos
   Else
      txt_clase = ""
      txt_descripcion = ""
   End If
End Sub

Private Sub cmd_eliminar_Click()
Dim si As Integer
si = MsgBox("¿Deseas eliminar el registro?", vbYesNo, "ATENCION")
   If si = 6 Then
      rs.Open "delete from tb_clases_cartera where vcha_car_documento = '" + txt_documento + "' and vcha_car_clase_id ='" + txt_clase + "'", cnn, adOpenDynamic, adLockOptimistic
      If Trim(txt_documento) <> "" Then
         lv_clases_cartera.ListItems.Clear
         If Trim(txt_documento) = "FA" Then
            rs.Open "select * from tb_clases_cartera where vcha_Car_documento = 'FA'", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
               Set list_item = lv_clases_cartera.ListItems.Add(, , rs!vcha_car_clase_id)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_car_nombre), "", rs!vcha_car_nombre)
               rs.MoveNext
            Wend
            rs.Close
         Else
            If Trim(txt_documento) = "NC" Then
               rs.Open "select * from tb_clases_cartera where vcha_Car_documento = 'NC'", cnn, adOpenDynamic, adLockOptimistic
               While Not rs.EOF
                  Set list_item = lv_clases_cartera.ListItems.Add(, , rs!vcha_car_clase_id)
                  list_item.SubItems(1) = IIf(IsNull(rs!vcha_car_nombre), "", rs!vcha_car_nombre)
                  rs.MoveNext
               Wend
               rs.Close
            Else
               If Trim(txt_documento) = "BO" Then
                  rs.Open "select * from tb_clases_cartera where vcha_Car_documento = 'BO'", cnn, adOpenDynamic, adLockOptimistic
                  While Not rs.EOF
                     Set list_item = lv_clases_cartera.ListItems.Add(, , rs!vcha_car_clase_id)
                     list_item.SubItems(1) = IIf(IsNull(rs!vcha_car_nombre), "", rs!vcha_car_nombre)
                     rs.MoveNext
                  Wend
                  rs.Close
               Else
                  If Trim(txt_documento) = "BF" Then
                     rs.Open "select * from tb_clases_cartera where vcha_Car_documento = 'BF'", cnn, adOpenDynamic, adLockOptimistic
                     While Not rs.EOF
                        Set list_item = lv_clases_cartera.ListItems.Add(, , rs!vcha_car_clase_id)
                        list_item.SubItems(1) = IIf(IsNull(rs!vcha_car_nombre), "", rs!vcha_car_nombre)
                        rs.MoveNext
                     Wend
                     rs.Close
                 Else
                     If Trim(txt_documento) = "PA" Then
                        rs.Open "select * from tb_clases_cartera where vcha_Car_documento = 'PA'", cnn, adOpenDynamic, adLockOptimistic
                        While Not rs.EOF
                           Set list_item = lv_clases_cartera.ListItems.Add(, , rs!vcha_car_clase_id)
                           list_item.SubItems(1) = IIf(IsNull(rs!vcha_car_nombre), "", rs!vcha_car_nombre)
                           rs.MoveNext
                        Wend
                        rs.Close
                     Else
                        MsgBox "Clave de documento incorrecta", vbOKOnly, "ATENCION"
                     End If
                  End If
               End If
            End If
         End If
         n = lv_clases_cartera.ListItems.Count
         If n > 0 And var_inserta = False Then
            Call textos
         Else
            txt_clase = ""
            txt_descripcion = ""
         End If
      End If
   End If
End Sub

Private Sub cmd_guardar_Click()
   If var_inserta = True Then
      rs.Open "select * from tb_clases_cartera  where vcha_Car_documento = '" + txt_documento + "' and vcha_car_clase_id = '" + txt_clase + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         rs.Close
         MsgBox "La clave de clase " + Trim(txt_clase) + " ya existe", vbOKOnly, "ATENCION"
      Else
         rs.Close
         rsaux2.Open "insert into tb_clases_cartera (vcha_Car_documento, vcha_car_clase_id, vcha_car_nombre) values ('" + Trim(txt_documento) + "', '" + Trim(txt_clase) + "', '" + Trim(txt_descripcion) + "')", cnn, adOpenDynamic, adLockOptimistic
         If Trim(txt_documento) <> "" Then
            lv_clases_cartera.ListItems.Clear
            If Trim(txt_documento) = "FA" Then
               rs.Open "select * from tb_clases_cartera where vcha_Car_documento = 'FA'", cnn, adOpenDynamic, adLockOptimistic
               While Not rs.EOF
                  Set list_item = lv_clases_cartera.ListItems.Add(, , rs!vcha_car_clase_id)
                  list_item.SubItems(1) = IIf(IsNull(rs!vcha_car_nombre), "", rs!vcha_car_nombre)
                  rs.MoveNext
               Wend
               rs.Close
            Else
               If Trim(txt_documento) = "NC" Then
                  rs.Open "select * from tb_clases_cartera where vcha_Car_documento = 'NC'", cnn, adOpenDynamic, adLockOptimistic
                  While Not rs.EOF
                     Set list_item = lv_clases_cartera.ListItems.Add(, , rs!vcha_car_clase_id)
                     list_item.SubItems(1) = IIf(IsNull(rs!vcha_car_nombre), "", rs!vcha_car_nombre)
                     rs.MoveNext
                  Wend
                  rs.Close
               Else
                  If Trim(txt_documento) = "BO" Then
                     rs.Open "select * from tb_clases_cartera where vcha_Car_documento = 'BO'", cnn, adOpenDynamic, adLockOptimistic
                     While Not rs.EOF
                        Set list_item = lv_clases_cartera.ListItems.Add(, , rs!vcha_car_clase_id)
                        list_item.SubItems(1) = IIf(IsNull(rs!vcha_car_nombre), "", rs!vcha_car_nombre)
                        rs.MoveNext
                     Wend
                     rs.Close
                  Else
                     If Trim(txt_documento) = "BF" Then
                        rs.Open "select * from tb_clases_cartera where vcha_Car_documento = 'BF'", cnn, adOpenDynamic, adLockOptimistic
                        While Not rs.EOF
                           Set list_item = lv_clases_cartera.ListItems.Add(, , rs!vcha_car_clase_id)
                           list_item.SubItems(1) = IIf(IsNull(rs!vcha_car_nombre), "", rs!vcha_car_nombre)
                           rs.MoveNext
                        Wend
                        rs.Close
                    Else
                        If Trim(txt_documento) = "PA" Then
                           rs.Open "select * from tb_clases_cartera where vcha_Car_documento = 'PA'", cnn, adOpenDynamic, adLockOptimistic
                           While Not rs.EOF
                              Set list_item = lv_clases_cartera.ListItems.Add(, , rs!vcha_car_clase_id)
                              list_item.SubItems(1) = IIf(IsNull(rs!vcha_car_nombre), "", rs!vcha_car_nombre)
                              rs.MoveNext
                           Wend
                           rs.Close
                        Else
                           If Trim(txt_documento) = "DV" Then
                              rs.Open "select * from tb_clases_cartera where vcha_Car_documento = 'DV'", cnn, adOpenDynamic, adLockOptimistic
                              While Not rs.EOF
                                 Set list_item = lv_clases_cartera.ListItems.Add(, , rs!vcha_car_clase_id)
                                 list_item.SubItems(1) = IIf(IsNull(rs!vcha_car_nombre), "", rs!vcha_car_nombre)
                                 rs.MoveNext
                              Wend
                              rs.Close
                           Else
                              If Trim(txt_documento) = "DF" Then
                                 rs.Open "select * from tb_clases_cartera where vcha_Car_documento = 'DF'", cnn, adOpenDynamic, adLockOptimistic
                                 While Not rs.EOF
                                    Set list_item = lv_clases_cartera.ListItems.Add(, , rs!vcha_car_clase_id)
                                    list_item.SubItems(1) = IIf(IsNull(rs!vcha_car_nombre), "", rs!vcha_car_nombre)
                                    rs.MoveNext
                                 Wend
                                 rs.Close
                              Else
                                 MsgBox "Clave de documento incorrecta", vbOKOnly, "ATENCION"
                              End If
                           End If
                        End If
                     End If
                  End If
               End If
            End If
            n = lv_clases_cartera.ListItems.Count
            If n > 0 And var_inserta = False Then
               Call textos
            Else
               txt_clase = ""
               txt_descripcion = ""
            End If
         End If
        
      End If
   End If
   var_inserta = False
End Sub

Private Sub cmd_nuevo_Click()
   txt_documento = ""
   cmb_documentos = ""
   txt_clase = ""
   txt_descripcion = ""
   txt_documento.SetFocus
   var_inserta = True
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_inserta = True Then
      var_si = MsgBox("No se han guardado los cambios, ¿Desea salir?", vbYesNo, "ATENCION")
      If var_si <> 6 Then
         GoTo salir:
      End If
   Else
      If var_hubo_cambios = True Then
         var_si = MsgBox("No se han guardado los cambios, ¿Desea salir?", vbYesNo, "ATENCION")
         If var_si <> 6 Then
            GoTo salir:
         End If
      End If
   End If
   Unload Me
   Exit Sub
salir:
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
      'cmd_imprimir_Click
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 0
   Left = 2900
End Sub

Private Sub Form_Unload(Cancel As Integer)
    var_swpassword = False
    var_modifica_registro = False
    Call activa_forma(var_activa_forma_clases_cartera)
End Sub

Private Sub lv_clases_cartera_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_clases_cartera, ColumnHeader)
End Sub

Private Sub lv_clases_cartera_ItemClick(ByVal Item As MSComctlLib.ListItem)
   var_inserta = False
   Call textos
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_clases_cartera.SetFocus
      Call pro_avanzar(Me, lv_clases_cartera, Button)
      lv_clases.selectedItem.EnsureVisible
      textos
   End If
   If Button.Index = 1 Then
      lv_clases_cartera.ListItems(1).Selected = True
      lv_clases.selectedItem.EnsureVisible
      textos
   End If
   If Button.Index = 4 Then
      numero_items_clases = lv_clases.ListItems.Count
      lv_clases_cartera.ListItems(numero_items_clases).Selected = True
      lv_clases_cartera.selectedItem.EnsureVisible
      textos
   End If
err0:

End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_clases_cartera, txt_buscar, False)
      txt_buscar = ""
      Call textos
   End If
End Sub

Private Sub txt_clase_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      txt_descripcion.SetFocus
   End If
End Sub

Private Sub txt_clase_LostFocus()
   If Trim(txt_clase) <> "" Then
      
   Else
      MsgBox "Clave de clase incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub txt_descripcion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Me.cmd_guardar.Enabled = True Then
         Me.cmd_guardar.SetFocus
      End If
   End If
End Sub

Private Sub txt_documento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      cmb_documentos.SetFocus
   End If
End Sub

Private Sub txt_documento_LostFocus()
Dim n As Integer
   lv_clases_cartera.ListItems.Clear
   If Trim(txt_documento) <> "" Then
      If Trim(txt_documento) = "FA" Then
         cmb_documentos = "FACTURACION"
         rs.Open "select * from tb_clases_cartera where vcha_Car_documento = 'FA'", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
            Set list_item = lv_clases_cartera.ListItems.Add(, , rs!vcha_car_clase_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_car_nombre), "", rs!vcha_car_nombre)
            rs.MoveNext
         Wend
         rs.Close
      Else
         If Trim(txt_documento) = "NC" Then
            cmb_documentos = "NOTA DE CARGO"
            rs.Open "select * from tb_clases_cartera where vcha_Car_documento = 'NC'", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
               Set list_item = lv_clases_cartera.ListItems.Add(, , rs!vcha_car_clase_id)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_car_nombre), "", rs!vcha_car_nombre)
               rs.MoveNext
            Wend
            rs.Close
         Else
            If Trim(txt_documento) = "BO" Then
               cmb_documentos = "BONIFICACION"
               rs.Open "select * from tb_clases_cartera where vcha_Car_documento = 'BO'", cnn, adOpenDynamic, adLockOptimistic
               While Not rs.EOF
                  Set list_item = lv_clases_cartera.ListItems.Add(, , rs!vcha_car_clase_id)
                  list_item.SubItems(1) = IIf(IsNull(rs!vcha_car_nombre), "", rs!vcha_car_nombre)
                  rs.MoveNext
               Wend
               rs.Close
            Else
               If Trim(txt_documento) = "BF" Then
                  cmb_documentos = "BONIFICACION FINANCIERA"
                  rs.Open "select * from tb_clases_cartera where vcha_Car_documento = 'BF'", cnn, adOpenDynamic, adLockOptimistic
                  While Not rs.EOF
                     Set list_item = lv_clases_cartera.ListItems.Add(, , rs!vcha_car_clase_id)
                     list_item.SubItems(1) = IIf(IsNull(rs!vcha_car_nombre), "", rs!vcha_car_nombre)
                     rs.MoveNext
                  Wend
                  rs.Close
               Else
                  If Trim(txt_documento) = "PA" Then
                     cmb_documentos = "PAGO"
                     rs.Open "select * from tb_clases_cartera where vcha_Car_documento = 'PA'", cnn, adOpenDynamic, adLockOptimistic
                     While Not rs.EOF
                        Set list_item = lv_clases_cartera.ListItems.Add(, , rs!vcha_car_clase_id)
                        list_item.SubItems(1) = IIf(IsNull(rs!vcha_car_nombre), "", rs!vcha_car_nombre)
                        rs.MoveNext
                     Wend
                     rs.Close
                  Else
                     If Trim(txt_documento) = "DV" Then
                        cmb_documentos = "DEVOLUCION SOBRE VENTA"
                        rs.Open "select * from tb_clases_cartera where vcha_Car_documento = 'DV'", cnn, adOpenDynamic, adLockOptimistic
                        While Not rs.EOF
                           Set list_item = lv_clases_cartera.ListItems.Add(, , rs!vcha_car_clase_id)
                           list_item.SubItems(1) = IIf(IsNull(rs!vcha_car_nombre), "", rs!vcha_car_nombre)
                           rs.MoveNext
                        Wend
                        rs.Close
                     Else
                        If Trim(txt_documento) = "DF" Then
                           cmb_documentos = "DESCUENTO FINANCIERO"
                           rs.Open "select * from tb_clases_cartera where vcha_Car_documento = 'DF'", cnn, adOpenDynamic, adLockOptimistic
                           While Not rs.EOF
                              Set list_item = lv_clases_cartera.ListItems.Add(, , rs!vcha_car_clase_id)
                              list_item.SubItems(1) = IIf(IsNull(rs!vcha_car_nombre), "", rs!vcha_car_nombre)
                              rs.MoveNext
                           Wend
                           rs.Close
                        Else
                           MsgBox "Clave de documento incorrecta", vbOKOnly, "ATENCION"
                        End If
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
   n = lv_clases_cartera.ListItems.Count
   If n > 0 And var_inserta = False Then
      Call textos
   Else
      txt_clase = ""
      txt_descripcion = ""
   End If
End Sub
