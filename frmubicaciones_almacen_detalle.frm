VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmubicaciones_almacen_detalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ubicaciones"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmubicaciones_almacen_detalle.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   105
      TabIndex        =   19
      Top             =   435
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   20
         Top             =   480
         Width           =   5595
         _ExtentX        =   9869
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
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7937
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   21
         Top             =   135
         Width           =   5610
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5490
      Left            =   120
      TabIndex        =   17
      Top             =   1725
      Width           =   5655
      Begin MSComctlLib.ListView lv_ubicaciones 
         Height          =   5280
         Left            =   45
         TabIndex        =   11
         Top             =   120
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   9313
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Ubicación"
            Object.Width           =   9816
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Colores "
      Height          =   1320
      Left            =   120
      TabIndex        =   13
      Top             =   405
      Width           =   5655
      Begin VB.TextBox txt_ubicacion 
         Height          =   315
         Left            =   1125
         MaxLength       =   50
         TabIndex        =   10
         Top             =   885
         Width           =   2190
      End
      Begin VB.TextBox txt_articulo 
         Height          =   315
         Left            =   1125
         MaxLength       =   50
         TabIndex        =   8
         Top             =   555
         Width           =   1395
      End
      Begin VB.TextBox txt_nombre_articulo 
         Height          =   315
         Left            =   2535
         MaxLength       =   50
         TabIndex        =   9
         Top             =   555
         Width           =   3030
      End
      Begin VB.TextBox txt_almacen 
         Height          =   315
         Left            =   1125
         MaxLength       =   50
         TabIndex        =   6
         Top             =   225
         Width           =   1395
      End
      Begin VB.TextBox txt_nombre_almacen 
         Height          =   315
         Left            =   2535
         MaxLength       =   50
         TabIndex        =   7
         Top             =   225
         Width           =   3030
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Ubicación:"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   18
         Top             =   945
         Width           =   765
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Almacen:"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   15
         Top             =   285
         Width           =   660
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Artículo:"
         Height          =   195
         Index           =   2
         Left            =   195
         TabIndex        =   14
         Top             =   615
         Width           =   600
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2745
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   12
      Top             =   30
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5430
      Picture         =   "frmubicaciones_almacen_detalle.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1395
      Picture         =   "frmubicaciones_almacen_detalle.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1065
      Picture         =   "frmubicaciones_almacen_detalle.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   735
      Picture         =   "frmubicaciones_almacen_detalle.frx":0940
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Guardar Alt + G"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frmubicaciones_almacen_detalle.frx":0A42
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   135
      Top             =   4890
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
            Picture         =   "frmubicaciones_almacen_detalle.frx":0B44
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmubicaciones_almacen_detalle.frx":141E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList 
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
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmubicaciones_almacen_detalle.frx":1CF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmubicaciones_almacen_detalle.frx":25D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmubicaciones_almacen_detalle.frx":2EAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmubicaciones_almacen_detalle.frx":3448
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmubicaciones_almacen_detalle.frx":3D24
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmubicaciones_almacen_detalle.frx":45FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmubicaciones_almacen_detalle.frx":4ED8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmubicaciones_almacen_detalle.frx":4FEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmubicaciones_almacen_detalle.frx":50FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmubicaciones_almacen_detalle.frx":520E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   120
      TabIndex        =   16
      Top             =   285
      Width           =   5655
   End
End
Attribute VB_Name = "frmubicaciones_almacen_detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
   Dim list_item As ListItem
   Dim var_tipo_lista As Integer

Private Sub cmd_eliminar_Click()
   var_si = MsgBox("¿Desea eliminar el registro?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      Me.lv_ubicaciones.ListItems.Clear
      rs.Open "DELETE FROM TB_UBICACIONES_ALMACEN_DETALLE WHERE VCHA_ALM_ALMACEN_ID = '" + Me.txt_almacen + "' AND VCHA_aRT_ARTICULO_ID = '" + Me.txt_articulo + "' AND VCHA_UBI_UBICACION = '" + Me.txt_ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
      rsaux1.Open "SELECT dbo.TB_UBICACIONES_ALMACEN_DETALLE.VCHA_ALM_ALMACEN_ID, dbo.TB_UBICACIONES_ALMACEN_DETALLE.VCHA_ART_ARTICULO_ID, dbo.TB_UBICACIONES_ALMACEN_DETALLE.VCHA_UBI_UBICACION, dbo.TB_ARTICULOS.VCHA_ART_NOMBRE_ESPAÑOL,dbo.TB_ALMACENES.VCHA_ALM_NOMBRE FROM dbo.TB_UBICACIONES_ALMACEN_DETALLE INNER JOIN dbo.TB_ARTICULOS ON dbo.TB_UBICACIONES_ALMACEN_DETALLE.VCHA_ART_ARTICULO_ID = dbo.TB_ARTICULOS.VCHA_ART_ARTICULO_ID INNER JOIN dbo.TB_ALMACENES ON dbo.TB_UBICACIONES_ALMACEN_DETALLE.VCHA_ALM_ALMACEN_ID = dbo.TB_ALMACENES.VCHA_ALM_ALMACEN_ID WHERE  (dbo.TB_UBICACIONES_ALMACEN_DETALLE.VCHA_ALM_ALMACEN_ID = '" + Me.txt_almacen + "') AND   (dbo.TB_UBICACIONES_ALMACEN_DETALLE.VCHA_ART_ARTICULO_ID = '" + Me.txt_articulo + "')", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux1.EOF Then
         While Not rsaux1.EOF
               Set list_item = Me.lv_ubicaciones.ListItems.Add(, , rsaux1!VCHA_UBI_UBICACION)
               rsaux1.MoveNext
         Wend
         rsaux1.MoveFirst
         Me.txt_articulo = IIf(IsNull(rsaux1!vcha_Art_articulo_id), "", rsaux1!vcha_Art_articulo_id)
         Me.txt_nombre_articulo = IIf(IsNull(rsaux1!vcha_art_nombre_español), "", rsaux1!vcha_art_nombre_español)
         Me.txt_almacen = IIf(IsNull(rsaux1!VCHA_ALM_ALMACEN_ID), "", rsaux1!VCHA_ALM_ALMACEN_ID)
         Me.txt_nombre_almacen = IIf(IsNull(rsaux1!VCHA_ALM_NOMBRE), "", rsaux1!VCHA_ALM_NOMBRE)
         Me.txt_ubicacion = IIf(IsNull(rsaux1!VCHA_UBI_UBICACION), "", rsaux1!VCHA_UBI_UBICACION)
      Else
         Me.txt_almacen = ""
         Me.txt_nombre_almacen = ""
         Me.txt_articulo = ""
         Me.txt_nombre_articulo = ""
      End If
      rsaux1.Close
      
      
   End If
   
End Sub

Private Sub cmd_guardar_Click()
   If Me.txt_almacen <> "" Then
      If Me.txt_articulo <> "" Then
         If Me.txt_ubicacion <> "" Then
            rs.Open "select * from tb_ubicaciones_almacen_detalle where vcha_alm_almacen_id = '" + Me.txt_almacen + "' and vcha_Art_articulo_id = '" + Me.txt_articulo + "' AND VCHA_UBI_UBICACION = '" + Me.txt_ubicacion + "'", cnn, adOpenDynamic, adLockOptimistic
            If rs.EOF Then
               rsaux.Open "INSERT INTO TB_UBICACIONES_ALMACEN_dETALLE (VCHA_ALM_ALMACEN_ID, VCHA_aRT_ARTICULO_ID, VCHA_UBI_UBICACION) VALUES ('" + Me.txt_almacen + "','" + Me.txt_articulo + "','" + Me.txt_ubicacion + "')", cnn, adOpenDynamic, adLockOptimistic
               Set list_item = Me.lv_ubicaciones.ListItems.Add(, , Me.txt_ubicacion)
            Else
               MsgBox "La ubicación ya existe", vbOKOnly, "ATENCION"
            End If
            rs.Close
         End If
      Else
         MsgBox "No se a indicado un artículo", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a seleccionado un almacén", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
      Set reporte = appl.OpenReport(App.Path + "\rep_ubicaciones_almacen_detalle.rpt")
      frmvistasprevias.cr.ReportSource = reporte
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      frmvistasprevias.cr.ViewReport
      frmvistasprevias.Caption = "Reporte de ubicaciones"
      frmvistasprevias.Show 1
      Set reporte = Nothing
      var_si = MsgBox("Desea exportar el reporte a excel", vbYesNo, "ATENCION")
      If var_si = 6 Then
         Set reporte = appl.OpenReport(App.Path + "\rep_ubicaciones_almacen_detalle.rpt")
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         reporte.ExportOptions.FormatType = crEFTExcel80
         reporte.ExportOptions.DestinationType = crEDTDiskFile
         archivo = "c:\reportessid\reporte_ubicaciones_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
         reporte.ExportOptions.DiskFileName = archivo
         reporte.Export False
         Set reporte = Nothing
      End If
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_ubicacion = ""
   Me.txt_ubicacion.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   Me.txt_almacen = ""
   Me.txt_nombre_almacen = ""
   Me.txt_articulo = ""
   Me.txt_nombre_articulo = ""
   Me.lv_ubicaciones.ListItems.Clear
   Me.txt_almacen.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 78 Then
      cmd_nuevo_Click
   End If
   If Shift = 4 And KeyCode = 71 Then
      cmd_guardar_Click
   End If
   If Shift = 4 And KeyCode = 69 Then
      cmd_eliminar_Click
   End If
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 2900
   Me.frm_lista.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_tipo_lista = 20 Then
         Me.txt_almacen = Me.lv_lista.selectedItem
         Me.txt_nombre_almacen = Me.lv_lista.selectedItem.SubItems(1)
         Me.txt_almacen.SetFocus
      End If
      If var_tipo_lista = 21 Then
         Me.txt_articulo = Me.lv_lista.selectedItem
         Me.txt_nombre_articulo = Me.lv_lista.selectedItem.SubItems(1)
         Me.txt_nombre_almacen.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub lv_ubicaciones_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Me.txt_ubicacion = lv_ubicaciones.selectedItem
End Sub

Private Sub txt_almacen_Change()
   Me.lv_ubicaciones.ListItems.Clear
End Sub

Private Sub txt_almacen_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_almacenes order by vcha_alm_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ALM_ALMACEN_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ALMACENES"
      var_tipo_lista = 20
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_almacen_KeyPress(KeyAscii As Integer)
   pro_enfoque (KeyAscii)
End Sub

Private Sub txt_almacen_LostFocus()
   If Me.txt_almacen <> "" Then
      rs.Open "select * from tb_almacenes where vcha_alm_almacen_id = '" + Me.txt_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.lv_ubicaciones.ListItems.Clear
         Me.txt_nombre_almacen = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
         If Me.txt_articulo <> "" Then
            rsaux1.Open "SELECT VCHA_UBI_UBICACION FROM TB_UBICACIONES_ALMACEN_DETALLE WHERE VCHA_aLM_ALMACEN_ID = '" + Me.txt_almacen + "' AND VCHA_aRT_ARTICULO_ID = '" + Me.txt_articulo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux1.EOF Then
               While Not rsaux1.EOF
                     Set list_item = Me.lv_ubicaciones.ListItems.Add(, , rsaux1(0).Value)
                     rsaux1.MoveNext
               Wend
            End If
            rsaux1.Close
         End If
      Else
         MsgBox "El almacen no existe", vbOKOnly, "ATENCION"
         Me.txt_almacen = ""
         Me.txt_nombre_almacen = ""
         Me.lv_ubicaciones.ListItems.Clear
      End If
      rs.Close
   Else
      'MsgBox "Debe de seleccionar un almacen", vbOKOnly, "ATENCION"
      Me.txt_nombre_almacen = ""
      Me.lv_ubicaciones.ListItems.Clear
   End If
End Sub

Private Sub txt_articulo_Change()
   Me.lv_ubicaciones.ListItems.Clear
End Sub

Private Sub txt_articulo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_aRTICULOS order by VCHA_aRT_NOMBRE_ESPAÑOL", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_Art_articulo_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ARTICULOS"
      var_tipo_lista = 21
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_articulo_KeyPress(KeyAscii As Integer)
   pro_enfoque (KeyAscii)
End Sub

Private Sub txt_articulo_LostFocus()
   If Trim(Me.txt_articulo) <> "" Then
      rs.Open "select * from tb_Articulos where vcha_art_articulo_id = '" + Me.txt_articulo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.lv_ubicaciones.ListItems.Clear
         Me.txt_nombre_articulo = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
         Me.lv_ubicaciones.ListItems.Clear
         If Me.txt_almacen <> "" Then
            rsaux1.Open "SELECT VCHA_UBI_UBICACION FROM TB_UBICACIONES_ALMACEN_DETALLE WHERE VCHA_aLM_ALMACEN_ID = '" + Me.txt_almacen + "' AND VCHA_aRT_ARTICULO_ID = '" + Me.txt_articulo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux1.EOF Then
               While Not rsaux1.EOF
                     Set list_item = Me.lv_ubicaciones.ListItems.Add(, , rsaux1(0).Value)
                     rsaux1.MoveNext
               Wend
            End If
            rsaux1.Close
         Else
            Me.lv_ubicaciones.ListItems.Clear
         End If
      
      Else
         rsaux.Open "select * from VW_EQUIVALENCIAS where vcha_equ_codigo_equivalente = '" + Me.txt_articulo + "'", cnn, adOpenDynamic, adLockOptimistic
         Me.lv_ubicaciones.ListItems.Clear
         If Not rsaux.EOF Then
            Me.txt_nombre_articulo = IIf(IsNull(rsaux!vcha_art_nombre_español), "", rsaux!vcha_art_nombre_español)
            Me.txt_articulo = IIf(IsNull(rsaux!vcha_Art_articulo_id), "", rsaux!vcha_Art_articulo_id)
            Me.lv_ubicaciones.ListItems.Clear
            If Me.txt_almacen <> "" Then
               rsaux1.Open "SELECT VCHA_UBI_UBICACION FROM TB_UBICACIONES_ALMACEN_DETALLE WHERE VCHA_aLM_ALMACEN_ID = '" + Me.txt_almacen + "' AND VCHA_aRT_ARTICULO_ID = '" + Me.txt_articulo + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux1.EOF Then
                  While Not rsaux1.EOF
                        Set list_item = Me.lv_ubicaciones.ListItems.Add(, , rsaux1(0).Value)
                        rsaux1.MoveNext
                  Wend
               End If
               rsaux1.Close
            Else
               Me.lv_ubicaciones.ListItems.Clear
            End If
         
         Else
            MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
            Me.txt_articulo = ""
            Me.txt_nombre_articulo = ""
            Me.lv_ubicaciones.ListItems.Clear
         End If
         rsaux.Close
      End If
      rs.Close
   Else
      'MsgBox "Debe de indicar un artículo", vbOKOnly, "ATENCION"
      Me.txt_nombre_articulo = ""
      Me.lv_ubicaciones.ListItems.Clear
   End If
End Sub

Private Sub txt_nombre_almacen_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_almacenes order by vcha_alm_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ALM_ALMACEN_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ALMACENES"
      var_tipo_lista = 20
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_almacen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      pro_enfoque (KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_articulo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_aRTICULOS order by VCHA_aRT_NOMBRE_ESPAÑOL", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_Art_articulo_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_ART_nombre), "", rs!vcha_ART_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ARTICULOS"
      var_tipo_lista = 21
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_articulo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      pro_enfoque (KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_ubicacion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub
