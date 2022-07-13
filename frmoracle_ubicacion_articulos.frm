VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_ubicacion_articulos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ubicación de artículos"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   2910
      Left            =   60
      TabIndex        =   2
      Top             =   1065
      Width           =   6855
      Begin MSComctlLib.ListView lv_articulos 
         Height          =   2670
         Left            =   60
         TabIndex        =   3
         Top             =   165
         Width           =   6705
         _ExtentX        =   11827
         _ExtentY        =   4710
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
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   7762
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Tipo"
            Object.Width           =   882
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   60
      TabIndex        =   0
      Top             =   -45
      Width           =   6855
      Begin VB.TextBox txt_ubicacion 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   1260
         TabIndex        =   4
         Top             =   480
         Width           =   3960
      End
      Begin VB.Label lbl_ubicacion 
         BackColor       =   &H000000FF&
         Caption         =   " Ubicación"
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   30
         TabIndex        =   1
         Top             =   135
         Width           =   6780
      End
   End
End
Attribute VB_Name = "frmoracle_ubicacion_articulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_ubicacion_Change()
   Me.lv_articulos.ListItems.Clear
End Sub

Private Sub txt_ubicacion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(Me.txt_ubicacion) <> "" Then
         rs.Open "select segment1 as segment1, description From xxvia_system_items_b where atTribute2 = '" + Trim(Me.txt_ubicacion) + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = Me.lv_articulos.ListItems.Add(, , rs!segment1)
               list_item.SubItems(1) = Format(rs!Description)
               list_item.SubItems(2) = 1
               rs.MoveNext
         Wend
         rs.Close
         rs.Open "select segment1 as segment1, description From mtl_system_items where atTribute3 = '" + Trim(Me.txt_ubicacion) + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = Me.lv_articulos.ListItems.Add(, , rs!segment1)
               list_item.SubItems(1) = Format(rs!Description)
               list_item.SubItems(2) = 2
               rs.MoveNext
         Wend
         rs.Close
         rs.Open "select segment1 as segment1, description From mtl_system_items where atTribute4 = '" + Trim(Me.txt_ubicacion) + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = Me.lv_articulos.ListItems.Add(, , rs!segment1)
               list_item.SubItems(1) = Format(rs!Description)
               list_item.SubItems(2) = 3
               rs.MoveNext
         Wend
         rs.Close
         rs.Open "select segment1 as segment1, description From mtl_system_items where atTribute5 = '" + Trim(Me.txt_ubicacion) + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = Me.lv_articulos.ListItems.Add(, , rs!segment1)
               list_item.SubItems(1) = Format(rs!Description)
               list_item.SubItems(2) = 4
               rs.MoveNext
         Wend
         rs.Close
         rs.Open "select segment1 as segment1, description From mtl_system_items where atTribute6 = '" + Trim(Me.txt_ubicacion) + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = Me.lv_articulos.ListItems.Add(, , rs!segment1)
               list_item.SubItems(1) = Format(rs!Description)
               list_item.SubItems(2) = 5
               rs.MoveNext
         Wend
         rs.Close
      
      Else
         MsgBox "Ubicación incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
End Sub
