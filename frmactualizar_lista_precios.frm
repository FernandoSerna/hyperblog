VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmactualizar_lista_precios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualizar listas de precios"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "  Listas de precios  "
      Height          =   2880
      Left            =   60
      TabIndex        =   7
      Top             =   45
      Width           =   6525
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         Picture         =   "frmactualizar_lista_precios.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   780
         Picture         =   "frmactualizar_lista_precios.frx":0216
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Marcar (Enter)"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1110
         Picture         =   "frmactualizar_lista_precios.frx":0460
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   120
         Picture         =   "frmactualizar_lista_precios.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   450
         Picture         =   "frmactualizar_lista_precios.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   225
         Width           =   330
      End
      Begin VB.Frame Frame6 
         Height          =   120
         Left            =   30
         TabIndex        =   8
         Top             =   540
         Width           =   6465
      End
      Begin MSComctlLib.ListView lv_listas 
         Height          =   2085
         Left            =   45
         TabIndex        =   14
         Top             =   720
         Width           =   6435
         _ExtentX        =   11351
         _ExtentY        =   3678
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
            Object.Width           =   8996
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame frmbusqueda_pedido 
      Height          =   3900
      Left            =   90
      TabIndex        =   0
      Top             =   2910
      Width           =   6525
      Begin VB.TextBox txt_ruta 
         Height          =   330
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   3390
         Width           =   6315
      End
      Begin VB.CommandButton cmd_buscar_pedido 
         Caption         =   "Actualizar precios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3330
         TabIndex        =   4
         Top             =   2805
         Width           =   3060
      End
      Begin VB.DirListBox Dir1 
         Height          =   2340
         Left            =   90
         TabIndex        =   3
         Top             =   930
         Width           =   3150
      End
      Begin VB.FileListBox File1 
         Height          =   2235
         Left            =   3330
         Pattern         =   "*.xls"
         TabIndex        =   2
         Top             =   510
         Width           =   3075
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   105
         TabIndex        =   1
         Top             =   510
         Width           =   3135
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
         TabIndex        =   6
         Top             =   120
         Width           =   6465
      End
   End
End
Attribute VB_Name = "frmactualizar_lista_precios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_mes As Integer

Private Sub cmd_buscar_pedido_Click()
   'On Error GoTo salir:
   var_contador = 0
   If Trim(Me.txt_ruta) <> "" Then
      For var_j = 1 To lv_listas.ListItems.Count
          lv_listas.ListItems.Item(var_j).Selected = True
          If Trim(Me.lv_listas.selectedItem.SubItems(2)) = "*" Then
             var_contador = var_contador + 1
          End If
      Next var_j
      If var_contador > 0 Then
         var_si = MsgBox("¿Desea actualizar los precios?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_si = MsgBox("Confirmar la actualización de los precios", vbYesNo, "ATENCION")
            If var_si = 6 Then
               strConnectionString = "DRIVER=Microsoft Excel Driver (*.xls); DBQ=" & Me.txt_ruta
               rsaux2.Open "SELECT * FROM [precios$]", strConnectionString
               cnn.BeginTrans
               'MsgBox cnn.ConnectionString
               rs.Open "select max(inte_tem_consecutivo) from TB_TEMP_CARGAR_LISTA_DE_PRECIOS", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
               Else
                  var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
               End If
               rs.Close
               var_consecutivo = var_consecutivo + 1
               rs.Open "insert into TB_TEMP_CARGAR_LISTA_DE_PRECIOS (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
               cnn.CommitTrans
               While Not rsaux2.EOF
                     var_codigo = Trim(IIf(IsNull(rsaux2!codigo), "", rsaux2!codigo))
                     rs.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_aRT_ARTICULO_ID = '" + CStr(var_codigo) + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        rsaux3.Open "INSERT INTO TB_TEMP_CARGAR_LISTA_DE_PRECIOS (INTE_TEM_CONSECUTIVO, VCHA_aRT_ARTICULO_ID, INTE_TEM_EXISTE, floa_dli_precio) VALUES (" + CStr(var_consecutivo) + ",'" + CStr(var_codigo) + "',0," + CStr(rsaux2!Precio) + ")", cnn, adOpenDynamic, adLockOptimistic
                     Else
                        rsaux3.Open "SELECT * FROM TB_EQUIVALENCIAS WHERE VCHA_EQU_CODIGO_EQUIVALENTE = '" + CStr(var_codigo) + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux3.EOF Then
                           rsaux4.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_ART_ARTICULO_ID = '" + IIf(IsNull(rsaux3!VCHA_ART_ARTICULO_ID), "", rsaux3!VCHA_ART_ARTICULO_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux4.EOF Then
                              var_codigo = IIf(IsNull(rsaux4!VCHA_ART_ARTICULO_ID), "", rsaux4!VCHA_ART_ARTICULO_ID)
                              rsaux5.Open "INSERT INTO TB_TEMP_CARGAR_LISTA_DE_PRECIOS (INTE_TEM_CONSECUTIVO, VCHA_aRT_ARTICULO_ID, INTE_TEM_EXISTE, floa_dli_precio) VALUES (" + CStr(var_consecutivo) + ",'" + var_codigo + "',0," + CStr(rsaux2!Precio) + ")", cnn, adOpenDynamic, adLockOptimistic
                           Else
                              rsaux5.Open "INSERT INTO TB_TEMP_CARGAR_LISTA_DE_PRECIOS (INTE_TEM_CONSECUTIVO, VCHA_aRT_ARTICULO_ID, INTE_TEM_EXISTE, floa_Dli_precio) VALUES (" + CStr(var_consecutivo) + ",'" + var_codigo + "',1," + CStr(rsaux2!Precio) + ")", cnn, adOpenDynamic, adLockOptimistic
                           End If
                           rsaux4.Close
                        Else
                           If Trim(var_codigo) <> "" Then
                              rsaux4.Open "INSERT INTO TB_TEMP_CARGAR_LISTA_DE_PRECIOS (INTE_TEM_CONSECUTIVO, VCHA_aRT_ARTICULO_ID, INTE_TEM_EXISTE, floa_dli_precio) VALUES (" + CStr(var_consecutivo) + ",'" + CStr(var_codigo) + "',1," + CStr(rsaux2!Precio) + ")", cnn, adOpenDynamic, adLockOptimistic
                           End If
                        End If
                        rsaux3.Close
                     End If
                     rs.Close
                     rsaux2.MoveNext
               Wend
               rsaux2.Close
               rs.Open "SELECT * FROM TB_TEMP_CARGAR_LISTA_DE_PRECIOS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " and inte_tem_existe = 1", cnn, adOpenDynamic, adLockOptimistic
               VAR_ACEPTAR = 0
               If rs.EOF Then
                  VAR_ACEPTAR = 0
               Else
                  var_cadena_CODIGOS = ""
                  While Not rs.EOF
                        If var_cadena_CODIGOS = "" Then
                           var_cadena_CODIGOS = IIf(IsNull(rs!VCHA_ART_ARTICULO_ID), "", rs!VCHA_ART_ARTICULO_ID)
                        Else
                           var_cadena_CODIGOS = var_cadena_CODIGOS + ", " + IIf(IsNull(rs!VCHA_ART_ARTICULO_ID), "", rs!VCHA_ART_ARTICULO_ID)
                        End If
                        rs.MoveNext
                  Wend
                  VAR_ACEPTAR = 1
                  var_si = MsgBox("Los siguientes códigos no estan dados de alta " + var_cadena_CODIGOS + " ¿Desea actualizar la lista de precios?", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     VAR_ACEPTAR = 0
                  End If
               End If
               rs.Close
               If VAR_ACEPTAR = 0 Then
                  For var_j = 1 To lv_listas.ListItems.Count
                      lv_listas.ListItems.Item(var_j).Selected = True
                      If Trim(Me.lv_listas.selectedItem.SubItems(2)) = "*" Then
                         rsaux1.Open "select * from TB_TEMP_CARGAR_LISTA_DE_PRECIOS where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_Art_Articulo_id is not null", cnn, adOpenDynamic, adLockOptimistic
                         While Not rsaux1.EOF
                               If Trim(Me.lv_listas.selectedItem) = "00" Then
                                  rsaux4.Open "SELECT * FROM TB_aRTICULOS WHERE VCHA_aRT_ARTICULO_ID = '" + rsaux1!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                                  If Not rsaux4.EOF Then
                                     var_precio = IIf(IsNull(rsaux4!mone_Art_precio_base), 0, rsaux4!mone_Art_precio_base)
                                     rsaux3.Open " INSERT INTO TB_RESPALDO_LISTA_PRECIOS (VCHA_LIS_LISTA_PRECIOS_ID, VCHA_ART_ARTICULO_ID, FLOA_RES_PRECIO_ANTERIOR, FLOA_RES_PRECIO_ACTUAL, DTIM_RES_FECHA_RESPALDO, VCHA_RES_MAQUINA, VCHA_RES_USUARIO, VCHA_RES_AFECTACION) VAlues  ( '" + Me.lv_listas.selectedItem + "', '" + rsaux1!VCHA_ART_ARTICULO_ID + "', " + CStr(var_precio) + ", " + CStr(IIf(IsNull(rsaux1!floa_dli_precio), 0, rsaux1!floa_dli_precio)) + ",GETDATE(), '" + fun_NombrePc + "', '" + var_clave_usuario_global + "', 'M')"
                                     rsaux3.Open "UPDATE TB_aRTICULOS SET MONE_ART_PRECIO_BASE = " + CStr(rsaux1!floa_dli_precio) + " WHERE VCHA_aRT_aRTICULO_ID = '" + rsaux1!VCHA_ART_ARTICULO_ID + "'"
                                  End If
                                  rsaux4.Close
                               Else
                                  'MsgBox rsaux1!vcha_Art_Articulo_id
                                  rsaux2.Open "select * from tb_Detalle_lista_precios where vcha_lis_lista_precios_id = '" + Me.lv_listas.selectedItem + "' and vcha_Art_Articulo_id = '" + rsaux1!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                                  If Not rsaux2.EOF Then
                                     rsaux3.Open " INSERT INTO TB_RESPALDO_LISTA_PRECIOS (VCHA_LIS_LISTA_PRECIOS_ID, VCHA_ART_ARTICULO_ID, FLOA_RES_PRECIO_ANTERIOR, FLOA_RES_PRECIO_ACTUAL, DTIM_RES_FECHA_RESPALDO, VCHA_RES_MAQUINA, VCHA_RES_USUARIO, VCHA_RES_AFECTACION) VAlues  ( '" + Me.lv_listas.selectedItem + "', '" + rsaux1!VCHA_ART_ARTICULO_ID + "', " + CStr(IIf(IsNull(rsaux2!floa_dli_precio), 0, rsaux2!floa_dli_precio)) + ", " + CStr(IIf(IsNull(rsaux1!floa_dli_precio), 0, rsaux1!floa_dli_precio)) + ",GETDATE(), '" + fun_NombrePc + "', '" + var_clave_usuario_global + "', 'M')", cnn, adOpenDynamic, adLockOptimistic
                                     'MsgBox "update tb_detalle_lista_precios set floa_dli_precio = " + CStr(rsaux1!floa_dli_precio) + " where vcha_lis_lista_precios_id = '" + Me.lv_listas.selectedItem + "' and vcha_Art_Articulo_id = '" + rsaux1!vcha_Art_Articulo_id + "'"
                                     rsaux3.Open "update tb_detalle_lista_precios set floa_dli_precio = " + CStr(rsaux1!floa_dli_precio) + " where vcha_lis_lista_precios_id = '" + Me.lv_listas.selectedItem + "' and vcha_Art_Articulo_id = '" + rsaux1!VCHA_ART_ARTICULO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
                                  Else
                                     rsaux3.Open " INSERT INTO TB_RESPALDO_LISTA_PRECIOS (VCHA_LIS_LISTA_PRECIOS_ID, VCHA_ART_ARTICULO_ID, FLOA_RES_PRECIO_ANTERIOR, FLOA_RES_PRECIO_ACTUAL, DTIM_RES_FECHA_RESPALDO, VCHA_RES_MAQUINA, VCHA_RES_USUARIO, VCHA_RES_AFECTACION) VAlues  ( '" + Me.lv_listas.selectedItem + "', '" + rsaux1!VCHA_ART_ARTICULO_ID + "', 0, " + CStr(IIf(IsNull(rsaux1!floa_dli_precio), 0, rsaux1!floa_dli_precio)) + ",GETDATE(), '" + fun_NombrePc + "', '" + var_clave_usuario_global + "', 'I')", cnn, adOpenDynamic, adLockOptimistic
                                     rsaux3.Open "insert into tb_Detalle_lista_precios (vcha_lis_lista_precios_id, vcha_Art_Articulo_id, floa_dli_precio) values ('" + Me.lv_listas.selectedItem + "','" + rsaux1!VCHA_ART_ARTICULO_ID + "'," + CStr(IIf(IsNull(rsaux1!floa_dli_precio), 0, rsaux1!floa_dli_precio)) + ")", cnn, adOpenDynamic, adLockOptimistic
                                  End If
                               rsaux2.Close
                               End If
                               rsaux1.MoveNext
                         Wend
                         rsaux1.Close
                      End If
                  Next var_j
                  MsgBox "Se a terminado de actualizar la lista de precios", vbOKOnly, "ATENCION"
               End If
            End If
         End If
      Else
         MsgBox "No se a seleccionado alguna lista de precios", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a seleccionado algún archivo", vbOKOnly, "ATENCION"
   End If
   Exit Sub
salir:
   MsgBox "Surgio un error al cargar el archivo", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If rsaux3.State = 1 Then
      rsaux3.Close
   End If
   If rsaux4.State = 1 Then
      rsaux4.Close
   End If
   If rsaux5.State = 1 Then
      rsaux5.Close
   End If
   If rsaux6.State = 1 Then
      rsaux6.Close
   End If
End Sub

Private Sub cmd_invertir_Click()
   n = lv_listas.ListItems.Count
   For i = 1 To n
      lv_listas.ListItems.Item(i).Selected = True
      If lv_listas.selectedItem.SubItems(2) = "*" Then
         lv_listas.selectedItem.SubItems(2) = ""
         lv_listas.ListItems.Item(i).Bold = False
         lv_listas.ListItems.Item(i).ForeColor = &H80000012
         lv_listas.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_listas.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_listas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_listas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_listas.selectedItem.SubItems(2) = "*"
         lv_listas.ListItems.Item(i).Bold = True
         lv_listas.ListItems.Item(i).ForeColor = &HFF0000
         lv_listas.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_listas.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_listas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_listas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      End If
   Next i

End Sub

Private Sub cmd_marcar_Click()
   i = lv_listas.selectedItem.Index
   If lv_listas.selectedItem.SubItems(2) = "*" Then
      lv_listas.selectedItem.SubItems(2) = ""
      lv_listas.ListItems.Item(i).Bold = False
      lv_listas.ListItems.Item(i).ForeColor = &H80000012
      lv_listas.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_listas.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_listas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_listas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_listas.Refresh
   Else
      lv_listas.selectedItem.SubItems(2) = "*"
      lv_listas.ListItems.Item(i).Bold = True
      lv_listas.ListItems.Item(i).ForeColor = &HFF0000
      lv_listas.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_listas.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_listas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_listas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_listas.Refresh
   End If
End Sub

Private Sub cmd_ninguno_Click()
   n = lv_listas.ListItems.Count
   For i = 1 To n
      lv_listas.ListItems.Item(i).Selected = True
      lv_listas.selectedItem.SubItems(2) = ""
      lv_listas.ListItems.Item(i).Bold = False
      lv_listas.ListItems.Item(i).ForeColor = &H80000012
      lv_listas.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_listas.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_listas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_listas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   lv_listas.Refresh
End Sub


Private Sub cmd_seleccion_Click()
   n = lv_listas.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_listas.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_listas.selectedItem.SubItems(2) = "" And var_rellena = True Then
         lv_listas.selectedItem.SubItems(2) = "*"
         lv_listas.ListItems.Item(i).Bold = True
         lv_listas.ListItems.Item(i).ForeColor = &HFF0000
         lv_listas.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_listas.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_listas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_listas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_listas.selectedItem.SubItems(2) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_listas.selectedItem.SubItems(2) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i

End Sub

Private Sub cmd_todos_Click()
   n = lv_listas.ListItems.Count
   For i = 1 To n
      lv_listas.ListItems.Item(i).Selected = True
      lv_listas.selectedItem.SubItems(2) = "*"
      lv_listas.ListItems.Item(i).Bold = True
      lv_listas.ListItems.Item(i).ForeColor = &HFF0000
      lv_listas.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_listas.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_listas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_listas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
   Next i
   lv_listas.Refresh
End Sub


Private Sub Dir1_Change()
   Me.File1.Path = Me.Dir1.Path
End Sub

Private Sub Drive1_Change()
   On Error GoTo salir:
   Me.Dir1.Path = Me.Drive1.Drive
   Me.Dir1.Refresh
   Exit Sub
salir:
   MsgBox "Unidad incorrecta"
   Me.Drive1.Drive = "c:"
End Sub

Private Sub File1_Click()
   If CStr(Me.Dir1.Path) = "C:\" Or CStr(Me.Dir1.Path) = "c:\" Then
      Me.txt_ruta = CStr(Me.Dir1.Path) + Me.File1.FileName
   Else
      Me.txt_ruta = CStr(Me.Dir1.Path) + "\" + Me.File1.FileName
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 200
   Left = 2400
   txt_inicio = Date
   txt_fin = Date
   rs.Open "select * from TB_LISTADEPRECIOS order by vcha_LIS_nombre ", cnn, adOpenDynamic, adLockOptimistic
   numero_items_ALMACENES = 0
   While Not rs.EOF
      Set list_item = lv_listas.ListItems.Add(, , rs!vcha_LIS_LISTA_iD)
      list_item.SubItems(1) = IIf(IsNull(rs!VCHA_lIS_NOMBRE), "", rs!VCHA_lIS_NOMBRE)
      list_item.SubItems(2) = ""
      rs.MoveNext:
   Wend
   rs.Close
   'If lv_listas.ListItems.Count > 7 Then
   '   lv_listas.ColumnHeaders(2).Width = 4220
   'Else
   '   lv_listas.ColumnHeaders(2).Width = 4499.71
   'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_despliega_menu = True Then
      var_swpassword = False
      var_modifica_registro = False
   End If
   Call activa_forma(var_activa_forma_articulos2)
End Sub


Private Sub lv_listas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_listas, ColumnHeader)
End Sub

Private Sub lv_listas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      i = lv_listas.selectedItem.Index
      If lv_listas.selectedItem.SubItems(2) = "*" Then
         lv_listas.selectedItem.SubItems(2) = ""
         lv_listas.ListItems.Item(i).Bold = False
         lv_listas.ListItems.Item(i).ForeColor = &H80000012
         lv_listas.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_listas.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_listas.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_listas.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_listas.Refresh
      Else
         lv_listas.selectedItem.SubItems(2) = "*"
         lv_listas.ListItems.Item(i).Bold = True
         lv_listas.ListItems.Item(i).ForeColor = &HFF0000
         lv_listas.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_listas.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_listas.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_listas.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_listas.Refresh
      End If
   End If
End Sub








