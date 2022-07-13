VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmimpresion_etiquetas_textilera 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresión de Etiquetas"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1590
      Left            =   45
      TabIndex        =   3
      Top             =   435
      Width           =   7350
      Begin VB.TextBox txt_descripcion 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1170
         TabIndex        =   14
         Top             =   1005
         Width           =   6120
      End
      Begin VB.TextBox txt_cantidad 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5580
         TabIndex        =   13
         Top             =   465
         Width           =   1710
      End
      Begin VB.TextBox txt_codigo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   555
         Left            =   1185
         TabIndex        =   5
         Top             =   405
         Width           =   2895
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1110
         Width           =   885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   555
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   4830
         TabIndex        =   11
         Top             =   555
         Width           =   675
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000C0&
         Caption         =   "Artículo"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   4
         Top             =   120
         Width           =   7275
      End
   End
   Begin VB.CommandButton cmd_salir 
      Height          =   375
      Left            =   7035
      Picture         =   "frmimpresion_etiquetas_textilera.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmd_imprimir 
      Height          =   375
      Left            =   45
      Picture         =   "frmimpresion_etiquetas_textilera.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   15
      TabIndex        =   0
      Top             =   375
      Width           =   7410
   End
   Begin VB.Frame Frame3 
      Height          =   3300
      Left            =   45
      TabIndex        =   6
      Top             =   2100
      Width           =   7350
      Begin VB.TextBox txt_nombre_articulo 
         Height          =   330
         Left            =   75
         TabIndex        =   10
         Top             =   480
         Width           =   7185
      End
      Begin VB.Frame Frame4 
         Height          =   120
         Left            =   30
         TabIndex        =   9
         Top             =   840
         Width           =   7275
      End
      Begin MSComctlLib.ListView lv_lista 
         Height          =   2220
         Left            =   45
         TabIndex        =   8
         Top             =   1005
         Width           =   7230
         _ExtentX        =   12753
         _ExtentY        =   3916
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
            Text            =   "Clave"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   10142
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   " Busqueda de artículos"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   7
         Top             =   120
         Width           =   7275
      End
   End
End
Attribute VB_Name = "frmimpresion_etiquetas_textilera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VAR_TIPO_LISTA  As Integer
Private Sub cmd_imprimir_Click()
   Dim var_codigo As String
   Dim VERIFICADOR As Integer
   If Trim(Me.txt_descripcion) <> "" Then
      var_codigo = Me.txt_codigo
      rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      rs.Open "SELECT * FROM XXVIA_VW_EQUIVALENCIAS_TEXT WHERE cross_reference = '" + Trim(Me.txt_codigo) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         VAR_CODIGO_alta = Me.txt_codigo
         If CInt(Me.txt_cantidad) > 0 Then
            t11 = ""
            t22 = Trim(Me.txt_descripcion)
            T33 = CStr(CDbl(Mid(rs!cross_reference, 11, 1)))
            s = Trim(var_codigo)
            If CInt(T33) > 0 Then
               Open (App.Path & "\etiqueta.bat") For Output As #2
               'Print #2, "copy " + App.Path + "\etiqueta.txt lpt1"
               Print #2, "copy " + App.Path + "\etiqueta.txt \\" + fun_NombrePc + "\zebra"
               Open (App.Path & "\etiqueta.txt") For Output As #1
               Close #2
               For i = 1 To CInt(Me.txt_cantidad)
                   Print #1, "US"
                   Print #1, "N"
                   Print #1, "Q256,24"
                   Print #1, "q512"
                   Print #1, "A80,20,0,3,1,1,N,""" + t11 + """"
                   Print #1, "A80,50,0,3,1,1,N,""" + t22 + """"
                   Print #1, "A80,80,0,3,1,1,N,""" + "Descuento:" + """"
                   Print #1, "A300,63,0,5,1,1,N,""" + T33 + "0%"""
                   Print #1, "B80,120,0,3,2,4,80,B,""" + s + """"
                   Print #1, "P1"
               Next i
               Close #1
               x = Shell(App.Path & "\etiqueta.bat", vbHide)
            Else
               Open (App.Path & "\etiqueta.bat") For Output As #2
               'Print #2, "copy " + App.Path + "\etiqueta.txt lpt1"
               Print #2, "copy " + App.Path + "\etiqueta.txt \\" + fun_NombrePc + "\zebra"
               
               Open (App.Path & "\etiqueta.txt") For Output As #1
               Close #2
               For i = 1 To CInt(Me.txt_cantidad)
                   Print #1, "US"
                   Print #1, "N"
                   Print #1, "Q256,24"
                   Print #1, "q512"
                   Print #1, "A80,20,0,3,1,1,N,""" + t11 + """"
                   Print #1, "A80,50,0,3,1,1,N,""" + t22 + """"
                   Print #1, "B80,120,0,3,2,4,80,B,""" + s + """"
                   Print #1, "P1"
               Next i
               Close #1
               'z = "copy " + App.Path & "\etiqueta.txt lpt1"
               z = "copy " + App.Path + "\etiqueta.txt \\" + fun_NombrePc + "\zebra"

               x = Shell(App.Path & "\etiqueta.bat", vbHide)
            End If
         End If
      Else
         MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      MsgBox "No se a seleccionado un artículo", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 1000
   Left = 2000
   Me.txt_cantidad = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_lista_Click()
   If Me.lv_lista.ListItems.Count > 0 Then
      Me.txt_codigo = Me.lv_lista.selectedItem
      Me.txt_descripcion = Me.lv_lista.selectedItem.SubItems(1)
   End If
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub


Private Sub lv_lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
   If Me.lv_lista.ListItems.Count > 0 Then
      Me.txt_codigo = Me.lv_lista.selectedItem
      Me.txt_descripcion = Me.lv_lista.selectedItem.SubItems(1)
   End If
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If Me.lv_lista.ListItems.Count > 0 Then
      Me.txt_codigo = Me.lv_lista.selectedItem
      Me.txt_descripcion = Me.lv_lista.selectedItem.SubItems(1)
   End If
   If KeyAscii = 13 Then
      Me.txt_cantidad.SetFocus
   End If
End Sub


Private Sub txt_cantidad_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   End If
End Sub















Private Sub txt_codigo_Change()
   Me.txt_descripcion = ""
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_cantidad.SetFocus
   End If
End Sub

Private Sub txt_codigo_LostFocus()
   If Trim(Me.txt_codigo) <> "" Then
      var_cadena = "SELECT * FROM XXVIA_VW_EQUIVALENCIAS_TEXT WHERE cross_reference = '" + Me.txt_codigo + "'"
      rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_descripcion = IIf(IsNull(rs!Description), "", rs!Description)
      Else
         MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
         Me.txt_descripcion = ""
      End If
      rs.Close
   Else
      Me.txt_descripcion = ""
   End If
End Sub

Private Sub txt_nombre_articulo_Change()
   Me.lv_lista.ListItems.Clear
   Me.txt_codigo = ""
End Sub

Private Sub txt_nombre_articulo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Or KeyAscii = 44 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Trim(Me.txt_nombre_articulo) <> "" Then
         var_cadena = ""
         var_like_1 = ""
         var_like_2 = ""
         var_like_3 = ""
         var_like_4 = ""
         var_like_5 = ""
         var_like_6 = ""
         var_like_7 = ""
         var_j = 1
         For var_i = 1 To Len(Me.txt_nombre_articulo)
             If Mid(Me.txt_nombre_articulo, var_i, 1) <> " " Then
                If var_j = 1 Then
                   var_like_1 = var_like_1 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
                If var_j = 2 Then
                   var_like_2 = var_like_2 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
                If var_j = 3 Then
                   var_like_3 = var_like_3 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
                If var_j = 4 Then
                   var_like_4 = var_like_4 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
                If var_j = 5 Then
                   var_like_5 = var_like_5 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
                If var_j = 6 Then
                   var_like_6 = var_like_6 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
                If var_j >= 7 Then
                   var_like_7 = var_like_7 + Mid(Me.txt_nombre_articulo, var_i, 1)
                End If
             Else
                var_j = var_j + 1
             End If
         Next var_i
      End If
      If Trim(var_like_1) <> "" Then
         var_cadena = " DESCRIPTION like '%" + var_like_1 + "%'"
      End If
      If Trim(var_like_2) <> "" Then
         var_cadena = var_cadena + " and  DESCRIPTION like '%" + var_like_2 + "%'"
      End If
      If Trim(var_like_3) <> "" Then
         var_cadena = var_cadena + " and  DESCRIPTION like '%" + var_like_3 + "%'"
      End If
      If Trim(var_like_4) <> "" Then
         var_cadena = var_cadena + " and  DESCRIPTION like '%" + var_like_4 + "%'"
      End If
      If Trim(var_like_5) <> "" Then
         var_cadena = var_cadena + " and  DESCRIPTION like '%" + var_like_5 + "%'"
      End If
      If Trim(var_like_6) <> "" Then
         var_cadena = var_cadena + " and  DESCRIPTION like '%" + var_like_6 + "%'"
      End If
      If Trim(var_like_7) <> "" Then
         var_cadena = var_cadena + " and  DESCRIPTION like '%" + var_like_7 + "%'"
      End If
      Me.lv_lista.ListItems.Clear
      If Trim(var_cadena) <> "" Then
         var_cadena = "SELECT * FROM XXVIA_VW_EQUIVALENCIAS_TEXT WHERE " + var_cadena
         rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!cross_reference)
            list_item.SubItems(1) = IIf(IsNull(rs!Description), "", rs!Description)
            rs.MoveNext
         Wend
         rs.Close
      End If
   End If
End Sub
