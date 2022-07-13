VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmconceptos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conceptos"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame6 
      Height          =   3150
      Left            =   90
      TabIndex        =   17
      Top             =   4020
      Width           =   5655
      Begin MSComctlLib.ListView lv_division 
         Height          =   2925
         Left            =   45
         TabIndex        =   7
         Top             =   165
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   5159
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
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7232
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Tipo Producto"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "tipoagente"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "zona"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "estatus"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "empresa"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "  Division "
      Height          =   1020
      Left            =   90
      TabIndex        =   16
      Top             =   3000
      Width           =   5655
      Begin VB.CommandButton com_nuevo_division 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   90
         Picture         =   "frmconceptos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton com_guardar_division 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   420
         Picture         =   "frmconceptos.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton com_eliminar_division 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   750
         Picture         =   "frmconceptos.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   210
         Width           =   330
      End
      Begin VB.TextBox txt_nombre_division 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   990
         MaxLength       =   50
         TabIndex        =   6
         Top             =   600
         Width           =   4560
      End
      Begin VB.TextBox txt_division 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   75
         MaxLength       =   2
         TabIndex        =   5
         Top             =   600
         Width           =   900
      End
   End
   Begin VB.Frame Frame4 
      Height          =   5985
      Left            =   5850
      TabIndex        =   15
      Top             =   1185
      Width           =   5655
      Begin MSComctlLib.ListView lv_subdivision 
         Height          =   5775
         Left            =   45
         TabIndex        =   13
         Top             =   150
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   10186
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
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7232
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Tipo Producto"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Division"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "zona"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "estatus"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "empresa"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Subdivision "
      Height          =   1110
      Left            =   5835
      TabIndex        =   14
      Top             =   60
      Width           =   5655
      Begin VB.CommandButton cmd_imprimir 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1065
         Picture         =   "frmconceptos.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   285
         Width           =   330
      End
      Begin VB.CommandButton com_nuevo_subdivision 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   75
         Picture         =   "frmconceptos.frx":0408
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   285
         Width           =   330
      End
      Begin VB.CommandButton com_guardar_subdivision 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   405
         Picture         =   "frmconceptos.frx":050A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   285
         Width           =   330
      End
      Begin VB.CommandButton com_eliminar_subdivision 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   735
         Picture         =   "frmconceptos.frx":060C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   285
         Width           =   330
      End
      Begin VB.TextBox txt_nombre_subdivision 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   990
         MaxLength       =   50
         TabIndex        =   12
         Top             =   660
         Width           =   4575
      End
      Begin VB.TextBox txt_subdivision 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   75
         MaxLength       =   50
         TabIndex        =   11
         Top             =   660
         Width           =   900
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2820
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   5655
      Begin MSComctlLib.ListView lv_tipoproductos 
         Height          =   2580
         Left            =   45
         TabIndex        =   1
         Top             =   150
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   4551
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
            Name            =   "Arial"
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
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7232
         EndProperty
      End
   End
End
Attribute VB_Name = "frmconceptos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Dim var_si As Integer
Private Sub com_nuevo_tipo_producto_Click()
   
End Sub

Private Sub cmd_imprimir_Click()
      Set reporte = appl.OpenReport(App.Path + "\rep_conceptos.rpt")
      frmvistasprevias.cr.ReportSource = reporte
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      frmvistasprevias.cr.ViewReport
      frmvistasprevias.Caption = "Reporte de conceptos"
      frmvistasprevias.Show 1
      Set reporte = Nothing
      var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
      If var_si = 6 Then
            Set reporte = appl.OpenReport(App.Path + "\rep_conceptos.rpt")
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\conceptos_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
      End If
      MsgBox "Se a terminado de guardar el archivo "

End Sub

Private Sub com_eliminar_division_Click()
   If Trim(txt_division) = "" Then
      MsgBox "No se a seleccionado una división", vbOKOnly, "ATENCION"
   Else
      rs.Open "SELECT * FROM TB_SUBDIVISIONES WHERE VCHA_TPR_TIPO_PRODUCTO_ID = '" + Me.lv_tipoproductos.selectedItem + "' AND VCHA_DIV_DIVISION_ID = '" + Me.txt_division + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         MsgBox "No se podra eliminar la división hasta que haya eliminado todas sus subdivisiones", vbOKOnly, "ATENCION"
      Else
         var_si = MsgBox("Desea eliminar la división", vbYesNo, "ATENCION")
         If var_si = 6 Then
            rsaux2.Open "DELETE FROM TB_DIVISIONES WHERE VCHA_DIV_DIVISION_ID = '" + txt_division + "' AND VCHA_TPR_TIPO_PRODUCTO_ID = '" + Me.lv_tipoproductos.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
            Me.lv_subdivision.ListItems.Clear
            Me.lv_division.ListItems.Clear
            rsaux.Open "SELECT * FROM TB_DIVISIONES WHERE VCHA_TPR_TIPO_PRODUCTO_ID = '" + Me.lv_tipoproductos.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               While Not rsaux.EOF
                     Set list_item = Me.lv_division.ListItems.Add(, , rsaux!vcha_div_division_id)
                     list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_div_nombre), "", rsaux!vcha_div_nombre)
                     numero_items_jaulas = numero_items_jaulas + 1
                     rsaux.MoveNext
               Wend
               rsaux.MoveFirst
               Me.txt_division = rsaux!vcha_div_division_id
               Me.txt_nombre_division = rsaux!vcha_div_nombre
            Else
               Me.txt_division = ""
               Me.txt_nombre_division = ""
            End If
            rsaux.Close
         End If
      End If
      rs.Close
   End If
End Sub

Private Sub com_eliminar_subdivision_Click()
   If Trim(txt_subdivision) = "" Then
      MsgBox "No se a seleccionado una subdivisión", vbOKOnly, "ATENCION"
   Else
      var_si = MsgBox("Desea elimiar la subdivisión", vbYesNo, "ATENCION")
      If var_si = 6 Then
         rsaux.Open "DELETE FROM TB_SUBDIVISIONES WHERE VCHA_SUB_SUBDIVISION_ID = '" + txt_subdivision + "' AND VCHA_DIV_DIVISION_ID = '" + Me.txt_division + "' AND VCHA_TPR_TIPO_PRODUCTO_ID = '" + Me.lv_tipoproductos.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
         Me.lv_subdivision.ListItems.Clear
         rsaux.Open "SELECT * FROM TB_SUBDIVISIONES WHERE VCHA_TPR_TIPO_PRODUCTO_ID = '" + Me.lv_tipoproductos.selectedItem + "' AND VCHA_DIV_DIVISION_ID = '" + Me.lv_division.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            While Not rsaux.EOF
                  Set list_item = Me.lv_subdivision.ListItems.Add(, , rsaux!VCHA_SUB_SUBDIVISION_ID)
                  list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_sub_nombre), "", rsaux!vcha_sub_nombre)
                  numero_items_jaulas = numero_items_jaulas + 1
                  rsaux.MoveNext
            Wend
            rsaux.MoveFirst
            Me.txt_subdivision = IIf(IsNull(rsaux!VCHA_SUB_SUBDIVISION_ID), "", rsaux!VCHA_SUB_SUBDIVISION_ID)
            Me.txt_nombre_subdivision = IIf(IsNull(rsaux!vcha_sub_nombre), "", rsaux!vcha_sub_nombre)
         Else
            Me.txt_subdivision = ""
            Me.txt_nombre_subdivision = ""
         End If
         rsaux.Close
            
      End If
   End If
End Sub

Private Sub com_guardar_division_Click()
   rs.Open "select * from tb_divisiones where vcha_div_division_id = '" + txt_division + "' and vcha_tpr_tipo_producto_id = '" + Me.lv_tipoproductos.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_si = MsgBox("La división ya existe, ¿Desea ejecutar los cambios?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         rsaux.Open "update tb_divisiones set vcha_div_nombre = '" + Me.txt_nombre_division + "' where vcha_tpr_tipo_producto_id = '" + Me.lv_tipoproductos.selectedItem + "' and vcha_div_division_id = '" + Me.txt_division + "'", cnn, adOpenDynamic, adLockOptimistic
      End If
   Else
      rsaux.Open "INSERT INTO TB_DIVISIONES (VCHA_TPR_TIPO_PRODUCTO_ID, VCHA_DIV_DIVISION_ID, VCHA_DIV_NOMBRE) VALUES ('" + Me.lv_tipoproductos.selectedItem + "', '" + txt_division + "','" + txt_nombre_division + "')", cnn, adOpenDynamic, adLockOptimistic
   End If
   lv_division.ListItems.Clear
   rsaux.Open "SELECT * FROM TB_DIVISIONES WHERE VCHA_TPR_TIPO_PRODUCTO_ID = '" + Me.lv_tipoproductos.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
   While Not rsaux.EOF
         Set list_item = Me.lv_division.ListItems.Add(, , rsaux!vcha_div_division_id)
         list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_div_nombre), "", rsaux!vcha_div_nombre)
         numero_items_jaulas = numero_items_jaulas + 1
         rsaux.MoveNext
   Wend
   rsaux.Close
   rs.Close
   
End Sub

Private Sub com_guardar_subdivision_Click()
   If Trim(txt_division) = "" Then
      MsgBox "No se a seleccionado una división", vbOKOnly, "ATENCION"
   Else
      rs.Open "select * from tb_subdivisiones where vcha_tpr_tipo_producto_id = '" + Me.lv_tipoproductos.selectedItem + "' and vcha_div_division_id = '" + Me.txt_division + "' and vcha_sub_subdivision_id = '" + Me.txt_subdivision + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_si = MsgBox("La subdivisión ya existe, ¿Desea aplicar los cambios?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            rsaux.Open "update tb_subdivisiones set vcha_sub_nombre = '" + Me.txt_nombre_subdivision + "' where vcha_tpr_tipo_producto_id = '" + Me.lv_tipoproductos.selectedItem + "' and vcha_div_division_id = '" + Me.txt_division + "' and vcha_sub_subdivision_id = '" + Me.txt_subdivision + "'", cnn, adOpenDynamic, adLockOptimistic
         End If
      Else
         rsaux.Open "insert into tb_subdivisiones (vcha_tpr_tipo_producto_id, vcha_div_division_id, vcha_sub_subdivision_id, vcha_sub_nombre) values ('" + Me.lv_tipoproductos.selectedItem + "','" + Me.txt_division + "','" + Me.txt_subdivision + "','" + Me.txt_nombre_subdivision + "')", cnn, adOpenDynamic, adLockOptimistic
      End If
      rs.Close
   
      Me.lv_subdivision.ListItems.Clear
      rsaux.Open "SELECT * FROM TB_SUBDIVISIONES WHERE VCHA_TPR_TIPO_PRODUCTO_ID = '" + Me.lv_tipoproductos.selectedItem + "' AND VCHA_DIV_DIVISION_ID = '" + Me.lv_division.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
      While Not rsaux.EOF
            Set list_item = Me.lv_subdivision.ListItems.Add(, , rsaux!VCHA_SUB_SUBDIVISION_ID)
            list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_sub_nombre), "", rsaux!vcha_sub_nombre)
            numero_items_jaulas = numero_items_jaulas + 1
            rsaux.MoveNext
      Wend
      rsaux.Close
   End If
End Sub

Private Sub com_nuevo_division_Click()
   Me.txt_division = ""
   Me.txt_nombre_division = ""
   Me.txt_subdivision = ""
   Me.txt_nombre_subdivision = ""
   Me.lv_subdivision.ListItems.Clear
   Me.txt_division.SetFocus
End Sub

Private Sub com_nuevo_subdivision_Click()
   Me.txt_subdivision = ""
   Me.txt_nombre_subdivision = ""
   Me.txt_subdivision.SetFocus
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 0
   Dim list_item As ListItem
   rs.Open "SELECT * FROM TB_TIPOS_PRODUCTOS", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set list_item = Me.lv_tipoproductos.ListItems.Add(, , rs!VCHA_tpr_tipo_producto_ID)
         list_item.SubItems(1) = IIf(IsNull(rs!vcha_tpr_nombre), "", rs!vcha_tpr_nombre)
         numero_items_jaulas = numero_items_jaulas + 1
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_reporte_valuacion_devoluciones)
End Sub

Private Sub lv_division_GotFocus()
   txt_division = lv_division.selectedItem
   txt_nombre_division = lv_division.selectedItem.SubItems(1)
   Me.lv_subdivision.ListItems.Clear
   rsaux2.Open "select * from tb_subdivisiones where vcha_tpr_tipo_producto_id = '" + Me.lv_tipoproductos.selectedItem + "' and vcha_div_division_id = '" + Me.txt_division + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rsaux2.EOF Then
      While Not rsaux2.EOF
            Set list_item = Me.lv_subdivision.ListItems.Add(, , rsaux2!VCHA_SUB_SUBDIVISION_ID)
            list_item.SubItems(1) = IIf(IsNull(rsaux2!vcha_sub_nombre), "", rsaux2!vcha_sub_nombre)
            rsaux2.MoveNext
      Wend
      rsaux2.MoveFirst
      Me.txt_subdivision = IIf(IsNull(rsaux2!VCHA_SUB_SUBDIVISION_ID), "", rsaux2!VCHA_SUB_SUBDIVISION_ID)
      Me.txt_nombre_subdivision = IIf(IsNull(rsaux2!vcha_sub_nombre), "", rsaux2!vcha_sub_nombre)
      If Me.lv_subdivision.ListItems.Count > 21 Then
         Me.lv_subdivision.ColumnHeaders(2).Width = 3850
      Else
         Me.lv_subdivision.ColumnHeaders(2).Width = 4100
      End If
   Else
      Me.txt_nombre_subdivision = ""
      Me.txt_subdivision = ""
   End If
   rsaux2.Close
End Sub

Private Sub lv_division_ItemClick(ByVal Item As MSComctlLib.ListItem)
   txt_division = lv_division.selectedItem
   txt_nombre_division = lv_division.selectedItem.SubItems(1)
   Me.lv_subdivision.ListItems.Clear
   rsaux2.Open "select * from tb_subdivisiones where vcha_tpr_tipo_producto_id = '" + Me.lv_tipoproductos.selectedItem + "' and vcha_div_division_id = '" + Me.txt_division + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rsaux2.EOF Then
      While Not rsaux2.EOF
            Set list_item = Me.lv_subdivision.ListItems.Add(, , rsaux2!VCHA_SUB_SUBDIVISION_ID)
            list_item.SubItems(1) = IIf(IsNull(rsaux2!vcha_sub_nombre), "", rsaux2!vcha_sub_nombre)
            rsaux2.MoveNext
      Wend
      rsaux2.MoveFirst
      Me.txt_subdivision = IIf(IsNull(rsaux2!VCHA_SUB_SUBDIVISION_ID), "", rsaux2!VCHA_SUB_SUBDIVISION_ID)
      Me.txt_nombre_subdivision = IIf(IsNull(rsaux2!vcha_sub_nombre), "", rsaux2!vcha_sub_nombre)
      If Me.lv_subdivision.ListItems.Count > 21 Then
         Me.lv_subdivision.ColumnHeaders(2).Width = 3850
      Else
         Me.lv_subdivision.ColumnHeaders(2).Width = 4100
      End If
   Else
      Me.txt_nombre_subdivision = ""
      Me.txt_subdivision = ""
   End If
   rsaux2.Close
End Sub

Private Sub lv_division_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.lv_subdivision.ListItems.Count > 0 Then
         Me.lv_subdivision.SetFocus
      End If
   End If
End Sub

Private Sub lv_tipoproductos_GotFocus()
      lv_division.ListItems.Clear
      Me.lv_subdivision.ListItems.Clear
      rsaux.Open "SELECT * FROM TB_DIVISIONES WHERE VCHA_TPR_TIPO_PRODUCTO_ID = '" + Me.lv_tipoproductos.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux.EOF Then
         While Not rsaux.EOF
               Set list_item = Me.lv_division.ListItems.Add(, , rsaux!vcha_div_division_id)
               list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_div_nombre), "", rsaux!vcha_div_nombre)
               numero_items_jaulas = numero_items_jaulas + 1
               rsaux.MoveNext
         Wend
         rsaux.MoveFirst
         Me.txt_division = rsaux!vcha_div_division_id
         Me.txt_nombre_division = rsaux!vcha_div_nombre
         If Me.lv_division.ListItems.Count > 10 Then
            Me.lv_division.ColumnHeaders(2).Width = 3850
         Else
            Me.lv_division.ColumnHeaders(2).Width = 4100.03
         End If

         Me.lv_subdivision.ListItems.Clear
         rsaux2.Open "select * from tb_subdivisiones where vcha_tpr_tipo_producto_id = '" + Me.lv_tipoproductos.selectedItem + "' and vcha_div_division_id = '" + Me.txt_division + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux2.EOF Then
            While Not rsaux2.EOF
                  Set list_item = Me.lv_subdivision.ListItems.Add(, , rsaux2!VCHA_SUB_SUBDIVISION_ID)
                  list_item.SubItems(1) = IIf(IsNull(rsaux2!vcha_sub_nombre), "", rsaux2!vcha_sub_nombre)
                  rsaux2.MoveNext
            Wend
            rsaux2.MoveFirst
            Me.txt_subdivision = IIf(IsNull(rsaux2!VCHA_SUB_SUBDIVISION_ID), "", rsaux2!VCHA_SUB_SUBDIVISION_ID)
            Me.txt_nombre_subdivision = IIf(IsNull(rsaux2!vcha_sub_nombre), "", rsaux2!vcha_sub_nombre)
         
            If Me.lv_subdivision.ListItems.Count > 21 Then
               Me.lv_subdivision.ColumnHeaders(2).Width = 3850
            Else
               Me.lv_subdivision.ColumnHeaders(2).Width = 4100
            End If
         
         Else
            Me.txt_nombre_subdivision = ""
            Me.txt_subdivision = ""
         End If
         rsaux2.Close
      Else
         Me.lv_division.ListItems.Clear
         Me.lv_subdivision.ListItems.Clear
         Me.txt_division = ""
         Me.txt_nombre_division = ""
         Me.txt_subdivision = ""
         Me.txt_nombre_subdivision = ""
      End If
      rsaux.Close
End Sub

Private Sub lv_tipoproductos_ItemClick(ByVal Item As MSComctlLib.ListItem)
      lv_division.ListItems.Clear
      Me.lv_subdivision.ListItems.Clear
      rsaux.Open "SELECT * FROM TB_DIVISIONES WHERE VCHA_TPR_TIPO_PRODUCTO_ID = '" + Me.lv_tipoproductos.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux.EOF Then
         While Not rsaux.EOF
               Set list_item = Me.lv_division.ListItems.Add(, , rsaux!vcha_div_division_id)
               list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_div_nombre), "", rsaux!vcha_div_nombre)
               numero_items_jaulas = numero_items_jaulas + 1
               rsaux.MoveNext
         Wend
         rsaux.MoveFirst
         If Me.lv_division.ListItems.Count > 10 Then
            Me.lv_division.ColumnHeaders(2).Width = 3850
         Else
            Me.lv_division.ColumnHeaders(2).Width = 4100.03
         End If
         Me.txt_division = rsaux!vcha_div_division_id
         Me.txt_nombre_division = rsaux!vcha_div_nombre
         Me.lv_subdivision.ListItems.Clear
         rsaux2.Open "select * from tb_subdivisiones where vcha_tpr_tipo_producto_id = '" + Me.lv_tipoproductos.selectedItem + "' and vcha_div_division_id = '" + Me.txt_division + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux2.EOF Then
            While Not rsaux2.EOF
                  Set list_item = Me.lv_subdivision.ListItems.Add(, , rsaux2!VCHA_SUB_SUBDIVISION_ID)
                  list_item.SubItems(1) = IIf(IsNull(rsaux2!vcha_sub_nombre), "", rsaux2!vcha_sub_nombre)
                  rsaux2.MoveNext
            Wend
            rsaux2.MoveFirst
            Me.txt_subdivision = IIf(IsNull(rsaux2!VCHA_SUB_SUBDIVISION_ID), "", rsaux2!VCHA_SUB_SUBDIVISION_ID)
            Me.txt_nombre_subdivision = IIf(IsNull(rsaux2!vcha_sub_nombre), "", rsaux2!vcha_sub_nombre)
            If Me.lv_subdivision.ListItems.Count > 21 Then
               Me.lv_subdivision.ColumnHeaders(2).Width = 3850
            Else
               Me.lv_subdivision.ColumnHeaders(2).Width = 4100
            End If
         Else
            Me.txt_nombre_subdivision = ""
            Me.txt_subdivision = ""
         End If
         rsaux2.Close
      Else
         Me.lv_division.ListItems.Clear
         Me.lv_subdivision.ListItems.Clear
         Me.txt_division = ""
         Me.txt_nombre_division = ""
         Me.txt_subdivision = ""
         Me.txt_nombre_subdivision = ""
      End If
      rsaux.Close
End Sub

Private Sub lv_tipoproductos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.lv_division.ListItems.Count > 0 Then
         lv_division.SetFocus
      End If
   End If
End Sub

Private Sub txt_division_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
End Sub


Private Sub txt_nombre_division_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.com_guardar_division.SetFocus
   End If
End Sub

Private Sub txt_nombre_subdivision_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.com_guardar_subdivision.SetFocus
   End If
End Sub

Private Sub txt_subdivision_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
End Sub
