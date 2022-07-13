VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_multiplo_articulos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Multiplo de artículos"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_cargar_archivo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1050
      Picture         =   "frmoracle_multiplo_articulos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Cargar archivo"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame frmbusqueda_pedido 
      Height          =   3900
      Left            =   450
      TabIndex        =   17
      Top             =   240
      Width           =   6525
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   105
         TabIndex        =   22
         Top             =   510
         Width           =   3135
      End
      Begin VB.FileListBox File1 
         Height          =   2235
         Left            =   3330
         Pattern         =   "*.xls"
         TabIndex        =   21
         Top             =   510
         Width           =   3075
      End
      Begin VB.DirListBox Dir1 
         Height          =   2340
         Left            =   90
         TabIndex        =   20
         Top             =   930
         Width           =   3150
      End
      Begin VB.CommandButton cmd_buscar_pedido 
         Caption         =   "Cargar archivo"
         Height          =   465
         Left            =   3330
         TabIndex        =   19
         Top             =   2790
         Width           =   3060
      End
      Begin VB.TextBox txt_ruta 
         Height          =   330
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   3390
         Width           =   6315
      End
      Begin VB.Label Label19 
         BackColor       =   &H8000000D&
         Caption         =   " Busqueda de pedido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   30
         TabIndex        =   23
         Top             =   120
         Width           =   6465
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4875
      Left            =   45
      TabIndex        =   14
      Top             =   2370
      Width           =   7575
      Begin MSComctlLib.ListView lv_multiplos 
         Height          =   4665
         Left            =   45
         TabIndex        =   15
         Top             =   135
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   8229
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
            Text            =   "Clave"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   8114
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Multiplo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Exportaciones"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   45
      TabIndex        =   11
      Top             =   1815
      Width           =   7575
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   1905
         TabIndex        =   12
         Top             =   150
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Busqueda de codigo:"
         Height          =   195
         Left            =   195
         TabIndex        =   13
         Top             =   195
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Multiplos de artículos"
      Height          =   1380
      Left            =   45
      TabIndex        =   4
      Top             =   435
      Width           =   7575
      Begin VB.CheckBox chk_exportaciones 
         Caption         =   "Check1"
         Height          =   225
         Left            =   2415
         TabIndex        =   25
         Top             =   960
         Width           =   1635
      End
      Begin VB.TextBox txt_codigo 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   7
         Top             =   255
         Width           =   1515
      End
      Begin VB.TextBox txt_descripcion 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   6
         Top             =   585
         Width           =   6195
      End
      Begin VB.TextBox txt_multiplo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   5
         Top             =   915
         Width           =   885
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "SKU:"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   10
         Top             =   255
         Width           =   375
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   9
         Top             =   615
         Width           =   885
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Multiplo:"
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   8
         Top             =   960
         Width           =   585
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7275
      Picture         =   "frmoracle_multiplo_articulos.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   720
      Picture         =   "frmoracle_multiplo_articulos.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   375
      Picture         =   "frmoracle_multiplo_articulos.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   45
      Picture         =   "frmoracle_multiplo_articulos.frx":0940
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   0
      Width           =   330
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3555
      Top             =   1680
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
            Picture         =   "frmoracle_multiplo_articulos.frx":0A42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_multiplo_articulos.frx":131C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_multiplo_articulos.frx":1BF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_multiplo_articulos.frx":2192
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_multiplo_articulos.frx":2A6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_multiplo_articulos.frx":3348
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_multiplo_articulos.frx":3C22
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_multiplo_articulos.frx":3D34
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_multiplo_articulos.frx":3E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_multiplo_articulos.frx":3F58
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_multiplo_articulos.frx":406A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame5 
      Height          =   120
      Left            =   45
      TabIndex        =   16
      Top             =   255
      Width           =   7575
   End
End
Attribute VB_Name = "frmoracle_multiplo_articulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter

Sub pro_llena_listview()

Dim list_item As ListItem

    rs.Open "select * from tb_oracle_multiplos", cnn, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
          strconsulta = "select * from xxvia_system_items_b where segment1 = ?"
          With comandoORA
               .ActiveConnection = cnnoracle_4
               .CommandType = adCmdText
               .CommandText = strconsulta
               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, rs!SEGMENT1)
               .Parameters.Append parametro
          End With
          Set rsaux6 = comandoORA.execute
          Set comandoORA = Nothing
          Set parametro = Nothing
    
         Set list_item = lv_multiplos.ListItems.Add(, , rs!SEGMENT1)
         list_item.SubItems(1) = rsaux6!Description
         list_item.SubItems(2) = rs!MULTIPLO
         list_item.SubItems(2) = IIf(IsNull(rs!MULTIPLO), 0, rs!MULTIPLO)
         rsaux6.Close
         rs.MoveNext
    Wend
    rs.Close

End Sub


Sub pro_textos()
'   On Error GoTo err0:
      Me.txt_codigo = lv_multiplos.selectedItem
      Me.txt_descripcion = lv_multiplos.selectedItem.SubItems(1)
      Me.txt_multiplo = lv_multiplos.selectedItem.SubItems(2)
err0:
End Sub


Private Sub cmd_buscar_pedido_Click()
On Error GoTo SALIR:
   
   var_archivo = Replace(Me.File1.FileName, ".xls", "")
   strConnectionString = "DRIVER=Microsoft Excel Driver (*.xls);" & "DBQ=" & Me.txt_ruta
   rsaux2.Open "SELECT * FROM [multiplos$] ", strConnectionString
   var_cadena_no_existen = ""
   If Not rsaux2.EOF Then
      While Not rsaux2.EOF
            var_codigo = CStr(IIf(IsNull(rsaux2!codigo), "", rsaux2!codigo))
            If Len(var_codigo) = 4 Then
               var_codigo = "0000" + CStr(var_codigo)
            End If
            If Len(var_codigo) = 5 Then
               var_codigo = "000" + CStr(var_codigo)
            End If
            If Len(var_codigo) = 6 Then
               var_codigo = "00" + CStr(var_codigo)
            End If
            If Len(var_codigo) = 7 Then
               var_codigo = "0" + CStr(var_codigo)
            End If
            strconsulta = "select * from xxvia_system_items_b a where a.segment1 = ? and organization_id = ?"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_codigo)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, var_unidad_organizacional)
                 .Parameters.Append parametro
            End With
            Set rsaux6 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            If Not rsaux6.EOF Then
               rsaux7.Open "DELETE FROM TB_ORACLE_MULTIPLOS WHERE SEGMENT1 = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               rsaux7.Open "INSERT INTO TB_ORACLE_MULTIPLOS (SEGMENT1, MULTIPLO, EXPORTACIONES) VALUES ('" + var_codigo + "'," + CStr(rsaux2!MULTIPLO) + "," + CStr(IIf(IsNull(rsaux2!EXPORTACIONES), 0, rsaux2!EXPORTACIONES)) + ")", cnn, adOpenDynamic, adLockOptimistic
            Else
               If var_cadena_no_existen = "" Then
                  var_cadena_no_existen = var_codigo
               Else
                  var_cadena_no_existen = var_cadena_no_existen + "," + var_codigo
               End If
            End If
            rsaux6.Close
            rsaux2.MoveNext
      Wend
      If var_cadena_no_existen <> "" Then
         MsgBox "Los siguientes códigos no estan dados de alta: " + var_cadena_no_existen, vbOKOnly, "ATENCION"
      End If
      Me.lv_multiplos.ListItems.Clear
      Call pro_llena_listview
      
      MsgBox "Se a terminado de cargar los multiplos", vbOKOnly, "ATENCION"
      Me.frmbusqueda_pedido.Visible = False
   End If
   rsaux2.Close
   Exit Sub
SALIR:
   MsgBox "Surgio un error al cargar el archivo, debe de tener las columnas CODIGO, MULTIPLO y EXPORTACIONES y la hoja debe de llamarse MULTIPLOS", vbOKOnly, "ATENCION"
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If rsaux6.State = 1 Then
      rsaux6.Close
   End If
   If rsaux7.State = 1 Then
      rsaux7.Close
   End If

End Sub

Private Sub cmd_buscar_pedido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frmbusqueda_pedido.Visible = False
   End If
End Sub

Private Sub cmd_cargar_archivo_Click()
   Me.frmbusqueda_pedido.Visible = True
   Me.Dir1.SetFocus
End Sub

Private Sub cmd_eliminar_Click()
   If Me.txt_codigo <> "" Then
   
      rs.Open "select * from tb_oracle_multiplos where segment1 = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_si = MsgBox("¿Desea eliminar el registro?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            Call pro_busca_registro(Me.lv_multiplos, Me.txt_codigo, False)
            lv_multiplos.ListItems.Remove (lv_multiplos.selectedItem.Index)
            rsaux.Open "DELETE FROM TB_ORACLE_MULTIPLOS WHERE SEGMENT1 = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            Me.txt_codigo = ""
         End If
      Else
         MsgBox "El código no se encuentra en la tabla de multiplos", vbOKOnly, "ATENCION"
      End If
      rs.Close
   End If
End Sub

Private Sub cmd_guardar_Click()
   If IsNumeric(Me.txt_multiplo) Then
      If CDbl(Me.txt_multiplo) > 1 Then
         If Me.txt_codigo <> "" Then
            rs.Open "select * from tb_oracle_multiplos where segment1 = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_si = MsgBox("¿Desea actualizar el registro?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  VAR_EXPORTACIONES = 0
                  If Me.chk_exportaciones.Value = 1 Then
                     VAR_EXPORTACIONES = 1
                  End If
                  rsaux.Open "update tb_oracle_multiplos set multiplo = " + Me.txt_multiplo + ", EXPORTACIONES = " + CStr(VAR_EXPORTACIONES) + " where segment1 = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                  Call pro_busca_registro(Me.lv_multiplos, txt_codigo, False)
                  Me.lv_multiplos.selectedItem.SubItems(2) = Me.txt_multiplo
                  Me.lv_multiplos.selectedItem.SubItems(3) = VAR_EXPORTACIONES
               End If
            Else
               strconsulta = "select * from xxvia_system_items_b where segment1 = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, Me.txt_codigo)
                    .Parameters.Append parametro
               End With
               Set rsaux6 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               If Not rsaux6.EOF Then
                  var_si = MsgBox("¿Desea insertar el registro?", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     VAR_EXPORTACIONES = 0
                     If Me.chk_exportaciones.Value = 1 Then
                        VAR_EXPORTACIONES = 1
                     End If
                     
                     rsaux.Open "insert into tb_oracle_multiplos (segment1, multiplo, EXPORTACIONES) values ('" + Me.txt_codigo + "', " + Me.txt_multiplo + "," + CStr(VAR_EXPORTACIONES) + ")", cnn, adOpenDynamic, adLockOptimistic
                     Set list_item = lv_multiplos.ListItems.Add(, , Me.txt_codigo)
                     list_item.SubItems(1) = Me.txt_descripcion
                     list_item.SubItems(2) = Me.txt_multiplo
                     Call pro_busca_registro(Me.lv_multiplos, txt_codigo, False)
                  End If
               Else
                  MsgBox "El articulo no existe", vbOKOnly, "ATENCION"
               End If
               rsaux6.Close
            End If
            rs.Close
         Else
            MsgBox "No se a seleccionado un artículo", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El multiplo debe de ser mayor a 1", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Número de multiplo incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_codigo = ""
   Me.txt_descripcion = ""
   Me.txt_multiplo = ""
   Me.txt_codigo.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Dir1_Change()
   Me.File1.Path = Me.Dir1.Path
End Sub

Private Sub Dir1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frmbusqueda_pedido.Visible = False
   End If
End Sub

Private Sub Drive1_Change()
   On Error GoTo SALIR:
   Me.Dir1.Path = Me.Drive1.Drive
   Me.Dir1.Refresh
   Exit Sub
SALIR:
   MsgBox "Unidad incorrecta"
   Me.Drive1.Drive = "c:"
End Sub

Private Sub Drive1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frmbusqueda_pedido.Visible = False
   End If
End Sub

Private Sub File1_Click()
   If CStr(Me.Dir1.Path) = "C:\" Or CStr(Me.Dir1.Path) = "c:\" Then
      Me.txt_ruta = CStr(Me.Dir1.Path) + Me.File1.FileName
   Else
      Me.txt_ruta = CStr(Me.Dir1.Path) + "\" + Me.File1.FileName
   End If
End Sub

Private Sub File1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frmbusqueda_pedido.Visible = False
   End If
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 2000
   Call pro_llena_listview
   Me.frmbusqueda_pedido.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_multiplos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_multiplos, ColumnHeader)
End Sub

Private Sub lv_multiplos_GotFocus()
   frmbusqueda_pedido.Visible = False
End Sub

Private Sub lv_multiplos_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Call pro_textos
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_multiplos, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub

Private Sub txt_codigo_Change()
   Me.txt_descripcion = ""
   Me.txt_multiplo = ""
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_descripcion.SetFocus
   End If
End Sub

Private Sub txt_codigo_LostFocus()
   If Trim(Me.txt_codigo) <> "" Then
      If Len(Me.txt_codigo) = 4 Then
         Me.txt_codigo = "0000" + Me.txt_codigo
      End If
      If Len(Me.txt_codigo) = 5 Then
         Me.txt_codigo = "000" + Me.txt_codigo
      End If
      If Len(Me.txt_codigo) = 6 Then
         Me.txt_codigo = "00" + Me.txt_codigo
      End If
      If Len(Me.txt_codigo) = 7 Then
         Me.txt_codigo = "0" + Me.txt_codigo
      End If
      rs.Open "select * from tb_oracle_multiplos where segment1 = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Call pro_busca_registro(Me.lv_multiplos, Me.txt_codigo, False)
         txt_buscar = ""
         pro_textos
      Else
          strconsulta = "select * from xxvia_system_items_b where segment1 = ?"
          With comandoORA
               .ActiveConnection = cnnoracle_4
               .CommandType = adCmdText
               .CommandText = strconsulta
               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, Me.txt_codigo)
               .Parameters.Append parametro
          End With
          Set rsaux6 = comandoORA.execute
          Set comandoORA = Nothing
          Set parametro = Nothing
          If Not rsaux6.EOF Then
             Me.txt_descripcion = rsaux6!Description
             Me.txt_multiplo = ""
             Me.txt_multiplo.SetFocus
          Else
             MsgBox "El articulo no existe", vbOKOnly, "ATENCION"
          End If
      End If
      rs.Close
   End If
End Sub

Private Sub txt_descripcion_GotFocus()
   frmbusqueda_pedido.Visible = False
End Sub

Private Sub txt_multiplo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub

Private Sub txt_ruta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frmbusqueda_pedido.Visible = False
   End If
End Sub
