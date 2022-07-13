VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmreporte_lista_precios_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Codigo de artículos"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_lista 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1440
      TabIndex        =   26
      Top             =   60
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Frame frm_estampados 
      Height          =   2835
      Left            =   135
      TabIndex        =   5
      Top             =   480
      Width           =   4365
      Begin VB.TextBox txt_clave_estampado 
         Height          =   360
         Left            =   75
         TabIndex        =   6
         Top             =   495
         Width           =   4215
      End
      Begin MSComctlLib.ListView lv_estampados 
         Height          =   1830
         Left            =   60
         TabIndex        =   7
         Top             =   915
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   3228
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
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7232
         EndProperty
      End
      Begin VB.Label lbl_estampado 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   8
         Top             =   105
         Width           =   4305
      End
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   105
      TabIndex        =   9
      Top             =   885
      Width           =   4365
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   10
         Top             =   480
         Width           =   4250
         _ExtentX        =   7488
         _ExtentY        =   3228
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
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7232
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   11
         Top             =   120
         Width           =   4305
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Código del Artículo "
      Height          =   810
      Left            =   105
      TabIndex        =   24
      Top             =   510
      Width           =   4425
      Begin VB.TextBox txt_tipo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   90
         MaxLength       =   1
         TabIndex        =   0
         Top             =   255
         Width           =   345
      End
      Begin VB.TextBox txt_division 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   450
         MaxLength       =   2
         TabIndex        =   1
         Top             =   255
         Width           =   585
      End
      Begin VB.TextBox txt_subdivision 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   1050
         MaxLength       =   2
         TabIndex        =   2
         Top             =   255
         Width           =   585
      End
      Begin VB.TextBox txt_estampado 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   1650
         MaxLength       =   5
         TabIndex        =   3
         Top             =   255
         Width           =   1230
      End
      Begin VB.TextBox txt_Descuento 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   3975
         MaxLength       =   1
         TabIndex        =   4
         Top             =   255
         Width           =   345
      End
      Begin VB.Label Label1 
         Caption         =   "descuento"
         Height          =   195
         Left            =   3015
         TabIndex        =   25
         Top             =   375
         Width           =   750
      End
   End
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   45
      TabIndex        =   23
      Top             =   375
      Width           =   4545
   End
   Begin VB.CommandButton cmd_imprimir 
      Height          =   375
      Left            =   90
      Picture         =   "frmreporte_lista_precios_01.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmd_salir 
      Height          =   375
      Left            =   4155
      Picture         =   "frmreporte_lista_precios_01.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   0
      Width           =   375
   End
   Begin VB.Frame Frame4 
      Caption         =   " Descripción del Artículo "
      Height          =   1725
      Left            =   120
      TabIndex        =   12
      Top             =   1380
      Width           =   4425
      Begin VB.TextBox txt_tipo_descripcion 
         Height          =   315
         Left            =   1050
         TabIndex        =   16
         Top             =   285
         Width           =   3285
      End
      Begin VB.TextBox txt_division_descripcion 
         Height          =   315
         Left            =   1050
         TabIndex        =   15
         Top             =   630
         Width           =   3285
      End
      Begin VB.TextBox txt_subdivision_descripcion 
         Height          =   315
         Left            =   1050
         TabIndex        =   14
         Top             =   975
         Width           =   3285
      End
      Begin VB.TextBox txt_estampado_descripcion 
         Height          =   315
         Left            =   1050
         TabIndex        =   13
         Top             =   1305
         Width           =   3285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   150
         TabIndex        =   20
         Top             =   345
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "División:"
         Height          =   195
         Left            =   150
         TabIndex        =   19
         Top             =   690
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Subdivisión:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1035
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Estampado:"
         Height          =   195
         Left            =   150
         TabIndex        =   17
         Top             =   1365
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmreporte_lista_precios_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_tipo_lista  As Integer

Private Sub cmd_imprimir_Click()
   var_filtro = ""
   var_filtro = " {VW_LISTA_PRECIOS_01.VCHA_LIS_LISTA_ID} = '" + Me.txt_lista + "' "
   If Me.txt_tipo <> "" Then
      var_filtro = var_filtro + " and Mid ({VW_LISTA_PRECIOS_01.VCHA_ART_ARTICULO_ID},1 ,1 ) = '" + Me.txt_tipo + "'"
   End If
   If Me.txt_division <> "" Then
      var_filtro = var_filtro + " and Mid ({VW_LISTA_PRECIOS_01.VCHA_ART_ARTICULO_ID},2 ,2 ) = '" + Me.txt_division + "'"
   End If
   If Me.txt_subdivision <> "" Then
      var_filtro = var_filtro + " and Mid ({VW_LISTA_PRECIOS_01.VCHA_ART_ARTICULO_ID},4 ,2 ) = '" + Me.txt_subdivision + "'"
   End If
   If Me.txt_estampado <> "" Then
      var_filtro = var_filtro + " and Mid ({VW_LISTA_PRECIOS_01.VCHA_ART_ARTICULO_ID},6 ,5 ) = '" + Me.txt_estampado + "'"
   End If
   
   Set reporte = appl.OpenReport(App.Path + "\rep_lista_precios_01.rpt")
   If var_filtro <> "" Then
      reporte.RecordSelectionFormula = var_filtro
   End If
   frmvistasprevias.cr.ReportSource = reporte
   For ntablas = 1 To reporte.Database.Tables.Count
       reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
   Next ntablas
   frmvistasprevias.cr.ViewReport
   frmvistasprevias.Caption = "Reporte de Comisiones por Linea"
   frmvistasprevias.Show 1
   Set reporte = Nothing
   var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      Set reporte = appl.OpenReport(App.Path + "\rep_lista_precios_01.rpt")
      If var_filtro <> "" Then
         reporte.RecordSelectionFormula = var_filtro
      End If
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      reporte.ExportOptions.FormatType = crEFTExcel80
      reporte.ExportOptions.DestinationType = crEDTDiskFile
      archivo = "c:\reportessid\Reporte_lista_precios" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
      reporte.ExportOptions.DiskFileName = archivo
      reporte.Export False
      Set reporte = Nothing
      MsgBox "Se a terminado de guardar el archivo " + archivo
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 1700
   Left = 3200
   Me.frm_lista.Visible = False
   Me.frm_estampados.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub lv_estampados_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.lv_estampados.ListItems.Count > 0 Then
         Me.txt_estampado = lv_estampados.selectedItem
         Me.txt_estampado.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_estampados.Visible = False
   End If
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      
      If lv_lista.ListItems.Count > 0 Then
         If var_tipo_lista = 1 Then
            Me.txt_tipo = lv_lista.selectedItem
            Me.txt_tipo.SetFocus
         End If
         If var_tipo_lista = 2 Then
            Me.txt_division = lv_lista.selectedItem
            Me.txt_division.SetFocus
         End If
         If var_tipo_lista = 3 Then
            Me.txt_subdivision = lv_lista.selectedItem
            Me.txt_subdivision.SetFocus
         End If
         If var_tipo_lista = 4 Then
            Me.txt_estampado = lv_lista.selectedItem
            Me.txt_estampado.SetFocus
         End If
      Else
         If var_tipo_lista = 1 Then
            Me.txt_tipo.SetFocus
         End If
         If var_tipo_lista = 2 Then
            Me.txt_division.SetFocus
         End If
         If var_tipo_lista = 3 Then
            Me.txt_subdivision.SetFocus
         End If
         If var_tipo_lista = 4 Then
            Me.txt_estampado.SetFocus
         End If
      End If
   End If
   If KeyAscii = 27 Then
      If var_tipo_lista = 1 Then
         Me.txt_tipo.SetFocus
      End If
      If var_tipo_lista = 2 Then
         Me.txt_division.SetFocus
      End If
      If var_tipo_lista = 3 Then
         Me.txt_subdivision.SetFocus
      End If
      If var_tipo_lista = 4 Then
         Me.txt_estampado.SetFocus
      End If
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub txt_cantidad_bulto_LostFocus()
End Sub

Private Sub txt_Cantidad_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 37 Then
      Me.txt_descuento.SetFocus
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

Private Sub txt_clave_estampado_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      lv_lista.ListItems.Clear
      rs.Open "select DISTINCT * from TB_ESTAMPADOS where vcha_est_nombre like '" + Trim(Me.txt_clave_estampado) + "%' ORDER BY VCHA_EST_nombre", cnn, adOpenDynamic, adLockOptimistic
      Me.lv_estampados.ListItems.Clear
      While Not rs.EOF
            Set list_item = lv_estampados.ListItems.Add(, , rs!VCHA_EST_ESTAMPADO_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_estampado = "ESTAMPADOS"
      var_tipo_lista = 4
      Dim var_n As Integer
      var_n = lv_estampados.ListItems.Count
      If var_n > 6 Then
         lv_estampados.ColumnHeaders(1).Width = 0
         lv_estampados.ColumnHeaders(2).Width = 3800
      Else
         lv_estampados.ColumnHeaders(1).Width = 0
         lv_estampados.ColumnHeaders(2).Width = 3800
      End If
      lv_estampados.SetFocus
   End If
   If KeyAscii = 27 Then
      Me.frm_estampados.Visible = False
   End If
End Sub

Private Sub txt_Descuento_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 37 Then
      Me.txt_estampado.SetFocus
   End If
End Sub

Private Sub txt_descuento_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_division_Change()
   If Len(Me.txt_division) = 2 Then
      Me.txt_subdivision.SetFocus
   End If
End Sub

Private Sub txt_division_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 37 Then
      Me.txt_tipo.SetFocus
   End If
   If KeyCode = 39 Then
      Me.txt_subdivision.SetFocus
   End If
   If KeyCode = 113 Then
      lv_lista.ListItems.Clear
      rs.Open "select DISTINCT * from TB_DIVISIONES WHERE VCHA_TPR_TIPO_PRODUCTO_ID = '" + Me.txt_tipo + "' order by vcha_DIV_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_div_division_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_div_nombre), "", rs!vcha_div_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "DIVISIONES"
      var_tipo_lista = 2
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3600
      Else
         lv_lista.ColumnHeaders(2).Width = 3800
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
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

Private Sub txt_division_LostFocus()
   If Trim(txt_division) <> "" Then
      If Trim(txt_tipo) <> "" Then
         If Len(Trim(txt_division)) = 1 Then
            txt_division = "0" + Trim(txt_division)
         End If
         rs.Open "SELECT * FROM TB_DIVISIONES WHERE VCHA_TPR_TIPO_PRODUCTO_ID = '" + Me.txt_tipo + "' AND VCHA_DIV_DIVISION_ID = '" + Me.txt_division + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_division_descripcion = IIf(IsNull(rs!vcha_div_nombre), "", rs!vcha_div_nombre)
            Me.txt_subdivision = ""
            Me.txt_subdivision_descripcion = ""
            Me.txt_estampado = ""
            Me.txt_estampado_descripcion = ""
            Me.txt_descuento = ""
         Else
            MsgBox "Clave de división no existe", vbOKOnly, "ATENCION"
            Me.txt_division = ""
            Me.txt_division_descripcion = ""
            Me.txt_subdivision = ""
            Me.txt_subdivision_descripcion = ""
            Me.txt_estampado = ""
            Me.txt_estampado_descripcion = ""
            Me.txt_descuento = ""
         End If
         rs.Close
      Else
         MsgBox "No se a indicado un tipo de producto", vbOKOnly, "ATENCION"
         Me.txt_division = ""
         Me.txt_division_descripcion = ""
         Me.txt_subdivision = ""
         Me.txt_subdivision_descripcion = ""
         Me.txt_estampado = ""
         Me.txt_estampado_descripcion = ""
         Me.txt_descuento = ""
      End If
   End If
End Sub

Private Sub txt_estampado_Change()
   If Len(Me.txt_estampado) = 5 Then
      Me.txt_descuento.SetFocus
   End If
End Sub

Private Sub txt_estampado_GotFocus()
    Me.frm_estampados.Visible = False
End Sub

Private Sub txt_estampado_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 37 Then
      Me.txt_subdivision.SetFocus
   End If
   If KeyCode = 39 Then
      Me.txt_descuento.SetFocus
   End If
   If KeyCode = 113 Then
   
      Me.lv_estampados.ListItems.Clear
      Me.txt_clave_estampado = ""
      Me.frm_estampados.Visible = True
      Me.txt_clave_estampado.SetFocus
   End If
End Sub

Private Sub txt_estampado_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_estampado_LostFocus()
   If Trim(txt_tipo) <> "" Then
      If Trim(txt_division) <> "" Then
         If Trim(txt_subdivision) <> "" Then
            If Trim(Me.txt_estampado) <> "" Then
               If Len(Trim(txt_estampado)) = 1 Then
                  txt_estampado = "0000" + Trim(txt_estampado)
               Else
                 If Len(Trim(txt_estampado)) = 2 Then
                    txt_estampado = "000" + Trim(txt_estampado)
                 Else
                    If Len(Trim(txt_estampado)) = 3 Then
                       txt_estampado = "00" + Trim(txt_estampado)
                    Else
                       If Len(Trim(txt_estampado)) = 4 Then
                          txt_estampado = "0" + Trim(txt_estampado)
                        End If
                     End If
                  End If
               End If
               rs.Open "SELECT * FROM TB_ESTAMPADOS WHERE VCHA_EST_eSTAMPADO_ID = '" + Me.txt_estampado + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  Me.txt_estampado_descripcion = IIf(IsNull(rs!vcha_est_nombre), "", rs!vcha_est_nombre)
               Else
                  MsgBox "Clave de estampado no existe", vbOKOnly, "ATENCION"
                  Me.txt_estampado = ""
                  Me.txt_estampado_descripcion = ""
                  Me.txt_descuento = ""
               End If
               rs.Close
            End If
         Else
            MsgBox "No se a seleccionado una subdivisión", vbOKOnly, "ATENCION"
            Me.txt_estampado = ""
            Me.txt_estampado_descripcion = ""
            Me.txt_descuento = ""
         End If
      Else
         MsgBox "No se a seleccionado una división", vbOKOnly, "ATENCION"
         Me.txt_subdivision = ""
         Me.txt_subdivision_descripcion = ""
         Me.txt_estampado = ""
         Me.txt_estampado_descripcion = ""
         Me.txt_descuento = ""
      End If
   Else
      MsgBox "No se a seleccionado un tipo de producto", vbOKOnly, "ATENCION"
      Me.txt_division = ""
      Me.txt_division_descripcion = ""
      Me.txt_subdivision = ""
      Me.txt_subdivision_descripcion = ""
      Me.txt_estampado = ""
      Me.txt_estampado_descripcion = ""
      Me.txt_descuento = ""
   End If
End Sub

Private Sub txt_subdivision_Change()
   If Len(Me.txt_subdivision) = 2 Then
      Me.txt_estampado.SetFocus
   End If
End Sub

Private Sub txt_subdivision_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 37 Then
      Me.txt_division.SetFocus
   End If
   If KeyCode = 39 Then
      Me.txt_estampado.SetFocus
   End If
   If KeyCode = 113 Then
      lv_lista.ListItems.Clear
      rs.Open "select DISTINCT  * from TB_SUBDIVISIONES WHERE VCHA_TPR_TIPO_PRODUCTO_ID = '" + Me.txt_tipo + "' AND VCHA_DIV_DIVISION_ID = '" + Me.txt_division + "' order by vcha_SUB_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_SUB_SUBDIVISION_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_sub_nombre), "", rs!vcha_sub_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "SUBDIVISIONES"
      var_tipo_lista = 3
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3600
      Else
         lv_lista.ColumnHeaders(2).Width = 3800
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
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

Private Sub txt_subdivision_LostFocus()
   If Trim(txt_subdivision) <> "" Then
      If Trim(txt_tipo) <> "" Then
        If Trim(txt_division) <> "" Then
            If Trim(txt_subdivision) <> "" Then
               If Len(Trim(Me.txt_subdivision)) = 1 Then
                  Me.txt_subdivision = "0" + Trim(txt_subdivision)
               End If
               rs.Open "SELECT * FROM TB_SUBDIVISIONES WHERE VCHA_SUB_SUBDIVISION_ID = '" + Me.txt_subdivision + "' AND VCHA_TPR_TIPO_PRODUCTO_ID = '" + Me.txt_tipo + "' AND VCHA_DIV_DIVISION_ID = '" + Me.txt_division + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  Me.txt_subdivision_descripcion = IIf(IsNull(rs!vcha_sub_nombre), "", rs!vcha_sub_nombre)
               Else
                  MsgBox "Clave de subdivisión incorrecta", vbOKOnly, "ATENCION"
                  Me.txt_subdivision = ""
                  Me.txt_subdivision_descripcion = ""
                  Me.txt_estampado = ""
                  Me.txt_estampado_descripcion = ""
                  Me.txt_descuento = ""
               End If
               rs.Close
            Else
               Me.txt_subdivision = ""
               Me.txt_subdivision_descripcion = ""
               Me.txt_estampado = ""
               Me.txt_estampado_descripcion = ""
               Me.txt_descuento = ""
            End If
         Else
            MsgBox "No se a indicado una división", vbOKOnly, "ATENCION"
            Me.txt_subdivision = ""
            Me.txt_subdivision_descripcion = ""
            Me.txt_estampado = ""
            Me.txt_estampado_descripcion = ""
            Me.txt_descuento = ""
         End If
      Else
         MsgBox "No se a indicado un tipo de producto", vbOKOnly, "ATENCION"
         Me.txt_subdivision = ""
         Me.txt_subdivision_descripcion = ""
         Me.txt_estampado = ""
         Me.txt_estampado_descripcion = ""
         Me.txt_descuento = ""
      End If
   End If
End Sub

Private Sub txt_tipo_Change()
   If Len(Me.txt_tipo) = 1 Then
      Me.txt_division.SetFocus
   End If
End Sub

Private Sub txt_tipo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 39 Then
      Me.txt_division.SetFocus
   End If
   If KeyCode = 113 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_TIPOS_PRODUCTOS order by vcha_tpr_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_tpr_tipo_producto_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_tpr_nombre), "", rs!vcha_tpr_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TIPOS PRODUCTOS"
      var_tipo_lista = 1
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 3600
      Else
         lv_lista.ColumnHeaders(2).Width = 3800
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_tipo_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
   
End Sub

Private Sub txt_tipo_LostFocus()
   If Trim(txt_tipo) = "" Then
      Me.txt_tipo_descripcion = ""
      Me.txt_division = ""
      Me.txt_division_descripcion = ""
      Me.txt_subdivision = ""
      Me.txt_subdivision_descripcion = ""
      Me.txt_estampado = ""
      Me.txt_estampado_descripcion = ""
      Me.txt_descuento = ""
   Else
      rs.Open "SELECT * FROM TB_TIPOS_PRODUCTOS WHERE VCHA_TPR_TIPO_PRODUCTO_ID = '" + Me.txt_tipo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_tipo_descripcion = IIf(IsNull(rs!vcha_tpr_nombre), "", rs!vcha_tpr_nombre)
         Me.txt_division = ""
         Me.txt_division_descripcion = ""
         Me.txt_subdivision = ""
         Me.txt_subdivision_descripcion = ""
         Me.txt_estampado = ""
         Me.txt_estampado_descripcion = ""
         Me.txt_descuento = ""
      Else
         MsgBox "Tipo de producto no existe", vbOKOnly, "ATENCION"
         Me.txt_tipo_descripcion = ""
         Me.txt_division = ""
         Me.txt_division_descripcion = ""
         Me.txt_subdivision = ""
         Me.txt_subdivision_descripcion = ""
         Me.txt_estampado = ""
         Me.txt_estampado_descripcion = ""
         Me.txt_descuento = ""
      End If
      rs.Close
   End If
End Sub

