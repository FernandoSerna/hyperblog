VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmreporte_catalogo_articulos_almacen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte del catálogo de artículos por almacén"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lv_almacenes 
      Height          =   2460
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   4339
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
End
Attribute VB_Name = "frmreporte_catalogo_articulos_almacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report


Private Sub Form_Load()
   Top = 2100
   Left = 3500
   lv_almacenes.ListItems.Clear
   rs.Open "select * from tb_almacenes where char_alm_tipo = 'A' order by vcha_alm_nombre ", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set list_item = lv_almacenes.ListItems.Add(, , rs!VCHA_ALM_ALMACEN_ID)
         list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
         rs.MoveNext
   Wend
   rs.Close
   Dim var_n As Integer
   var_n = lv_almacenes.ListItems.Count
   If var_n > 6 Then
      lv_almacenes.ColumnHeaders(2).Width = 4270.71
   Else
      lv_almacenes.ColumnHeaders(2).Width = 4499.71
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_reporte_catalogo_articulos_almacen)
End Sub

Private Sub lv_almacenes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_almacenes, ColumnHeader)
End Sub

Private Sub lv_almacenes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Set reporte = appl.OpenReport(App.Path + "\rep_catalogo_articulos_almacen.rpt")
      frmvistasprevias.cr.ReportSource = reporte
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      frmvistasprevias.cr.ViewReport
      frmvistasprevias.Caption = "Catálogo de artículos por almacén"
      frmvistasprevias.Show 1
      Set reporte = Nothing
      var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         Set reporte = appl.OpenReport(App.Path + "\rep_catalogo_articulos_almacen.rpt")
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         reporte.ExportOptions.FormatType = crEFTExcel80
         reporte.ExportOptions.DestinationType = crEDTDiskFile
         archivo = "c:\reportessid\catalogo_articulos" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
         reporte.ExportOptions.DiskFileName = archivo
         reporte.Export False
         Set reporte = Nothing
         MsgBox "Se a terminado de guardar el archivo " + archivo
     End If
   End If
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub
