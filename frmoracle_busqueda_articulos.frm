VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_busqueda_articulos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Busqueda de artículos"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frm_lista 
      Height          =   3480
      Left            =   45
      TabIndex        =   2
      Top             =   -75
      Width           =   7305
      Begin VB.TextBox txt_descripcion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   45
         TabIndex        =   0
         Top             =   180
         Width           =   7185
      End
      Begin MSComctlLib.ListView lv_lista 
         Height          =   2610
         Left            =   30
         TabIndex        =   1
         Top             =   795
         Width           =   7230
         _ExtentX        =   12753
         _ExtentY        =   4604
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   10583
         EndProperty
      End
   End
End
Attribute VB_Name = "frmoracle_busqueda_articulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.lv_lista.ListItems.Count > 0 Then
         var_codigo_busqueda = Me.lv_lista.selectedItem
         var_descripcion_busqueda = Me.lv_lista.selectedItem.SubItems(1)
         Unload Me
      End If
   End If
End Sub

Private Sub txt_descripcion_Change()
   Me.lv_lista.ListItems.Clear
End Sub

Private Sub txt_descripcion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(Me.txt_descripcion) <> "" Then
         var_cadena_1 = ""
         var_cadena_2 = ""
         var_cadena_3 = ""
         var_cadena_4 = ""
         var_cadena_5 = ""
         var_cadena_6 = ""
         var_cadena_7 = ""
         var_cadena_8 = ""
         var_cadena_9 = ""
         var_cadena_10 = ""
         var_j = 1
         var_n = 1
         While var_j <= Len(Trim(Me.txt_descripcion))

               If var_n = 10 Then
                  If Mid(Trim(Me.txt_descripcion), var_j, 1) <> " " Then
                     var_cadena_10 = var_cadena_10 + Mid(Me.txt_descripcion, var_j, 1)
                  Else
                     var_n = 11
                  End If
               End If
               If var_n = 9 Then
                  If Mid(Trim(Me.txt_descripcion), var_j, 1) <> " " Then
                     var_cadena_9 = var_cadena_9 + Mid(Me.txt_descripcion, var_j, 1)
                  Else
                     var_n = 10
                  End If
               End If
               If var_n = 8 Then
                  If Mid(Trim(Me.txt_descripcion), var_j, 1) <> " " Then
                     var_cadena_8 = var_cadena_8 + Mid(Me.txt_descripcion, var_j, 1)
                  Else
                     var_n = 9
                  End If
               End If
               If var_n = 7 Then
                  If Mid(Trim(Me.txt_descripcion), var_j, 1) <> " " Then
                     var_cadena_7 = var_cadena_7 + Mid(Me.txt_descripcion, var_j, 1)
                  Else
                     var_n = 8
                  End If
               End If
               If var_n = 6 Then
                  If Mid(Trim(Me.txt_descripcion), var_j, 1) <> " " Then
                     var_cadena_6 = var_cadena_6 + Mid(Me.txt_descripcion, var_j, 1)
                  Else
                     var_n = 7
                  End If
               End If
               If var_n = 5 Then
                  If Mid(Trim(Me.txt_descripcion), var_j, 1) <> " " Then
                     var_cadena_5 = var_cadena_5 + Mid(Me.txt_descripcion, var_j, 1)
                  Else
                     var_n = 6
                  End If
               End If
               If var_n = 4 Then
                  If Mid(Trim(Me.txt_descripcion), var_j, 1) <> " " Then
                     var_cadena_4 = var_cadena_4 + Mid(Me.txt_descripcion, var_j, 1)
                  Else
                     var_n = 5
                  End If
               End If
               If var_n = 3 Then
                  If Mid(Trim(Me.txt_descripcion), var_j, 1) <> " " Then
                     var_cadena_3 = var_cadena_3 + Mid(Me.txt_descripcion, var_j, 1)
                  Else
                     var_n = 4
                  End If
               End If
               If var_n = 2 Then
                  If Mid(Trim(Me.txt_descripcion), var_j, 1) <> " " Then
                     var_cadena_2 = var_cadena_2 + Mid(Me.txt_descripcion, var_j, 1)
                  Else
                     var_n = 3
                  End If
               End If
               If var_n = 1 Then
                  If Mid(Trim(Me.txt_descripcion), var_j, 1) <> " " Then
                     var_cadena_1 = var_cadena_1 + Mid(Me.txt_descripcion, var_j, 1)
                  Else
                     var_n = 2
                  End If
               End If
               var_j = var_j + 1
         Wend
         var_cadena = "select segment1, description from xxvia_system_items_b where "
         If var_cadena_1 <> "" Then
            var_cadena = var_cadena + " upper(description) like '%" + UCase(var_cadena_1) + "%'"
         End If
         If var_cadena_2 <> "" Then
            var_cadena = var_cadena + " and upper(description) like '%" + UCase(var_cadena_2) + "%'"
         End If
         If var_cadena_3 <> "" Then
            var_cadena = var_cadena + " and upper(description) like '%" + UCase(var_cadena_3) + "%'"
         End If
         If var_cadena_4 <> "" Then
            var_cadena = var_cadena + " and upper(description) like '%" + UCase(var_cadena_4) + "%'"
         End If
         If var_cadena_5 <> "" Then
            var_cadena = var_cadena + " and upper(description) like '%" + UCase(var_cadena_5) + "%'"
         End If
         If var_cadena_6 <> "" Then
            var_cadena = var_cadena + " and upper(description) like '%" + UCase(var_cadena_6) + "%'"
         End If
         If var_cadena_7 <> "" Then
            var_cadena = var_cadena + " and upper(description) like '%" + UCase(var_cadena_7) + "%'"
         End If
         If var_cadena_8 <> "" Then
            var_cadena = var_cadena + " and upper(description) like '%" + UCase(var_cadena_8) + "%'"
         End If
         If var_cadena_9 <> "" Then
            var_cadena = var_cadena + " and upper(description) like '%" + UCase(var_cadena_9) + "%'"
         End If
         If var_cadena_10 <> "" Then
            var_cadena = var_cadena + " and upper(description) like '%" + UCase(var_cadena_10) + "%'"
         End If
         rs.Open var_cadena + " and organization_id = " + var_unidad_organizacional + " order by description", cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!SEGMENT1)
            list_item.SubItems(1) = IIf(IsNull(rs!Description), "", rs!Description)
            rs.MoveNext
         Wend
         rs.Close
         If lv_lista.ListItems.Count > 0 Then
            Me.lv_lista.SetFocus
         End If
      End If
   End If
End Sub
