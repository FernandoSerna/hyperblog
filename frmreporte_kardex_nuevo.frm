VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmreporte_kardex_nuevo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Kardex"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   240
      TabIndex        =   14
      Top             =   390
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   15
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
         TabIndex        =   16
         Top             =   135
         Width           =   5610
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Artículo  "
      Height          =   720
      Left            =   120
      TabIndex        =   13
      Top             =   1230
      Width           =   6015
      Begin VB.TextBox txt_codigo 
         Height          =   345
         Left            =   210
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txt_nombre_articulo 
         Height          =   345
         Left            =   1440
         TabIndex        =   5
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Almacén "
      Height          =   720
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   6015
      Begin VB.TextBox txt_nombre_almacen 
         Height          =   345
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox txt_almacen 
         Height          =   345
         Left            =   195
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   720
      Left            =   120
      TabIndex        =   9
      Top             =   1980
      Width           =   6030
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   1620
         TabIndex        =   6
         Top             =   255
         Width           =   1080
      End
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   3630
         TabIndex        =   7
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   1170
         TabIndex        =   11
         Top             =   315
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   3315
         TabIndex        =   10
         Top             =   315
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   8
      Top             =   345
      Width           =   6180
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmreporte_kardex_nuevo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5790
      Picture         =   "frmreporte_kardex_nuevo.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "frmreporte_kardex_nuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_tipo_lista As Integer

Private Sub cmd_imprimir_Click()
   Dim var_consecutivo As Double
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim año As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   
   'On Error GoTo salir:
   If Me.txt_almacen <> "" Then
      If Me.txt_codigo <> "" Then
         If IsDate(txt_inicio) Then
            If IsDate(txt_fin) Then
               If CDate(txt_inicio) <= CDate(txt_fin) Then
                  cnn.BeginTrans
                  rs.Open "select max(inte_TEM_consecutivo) from tb_temp_kardex", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
                  Else
                     var_consecutivo = 0
                  End If
                  var_consecutivo = var_consecutivo + 1
                  rs.Close
                  rs.Open "Insert into tb_temp_kardex (INTE_tem_CONSECUTIVO) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                  cnn.CommitTrans
                  var_fecha_fin_1 = CDate(txt_fin)
             
                  var_dia = CStr(Day(CDate(txt_inicio)))
                  var_mes = CStr(Month(CDate(txt_inicio)))
                  var_año = CStr(Year(CDate(txt_inicio)))
                  If Len(Trim(var_dia)) = 1 Then
                     var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
             
             
                  var_dia = CStr(Day(var_fecha_fin_1))
                  var_mes = CStr(Month(var_fecha_fin_1))
                  var_año = CStr(Year(var_fecha_fin_1))
                  If Len(Trim(var_dia)) = 1 Then
                     var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                  rs.Open "EXEC SP_KARDEX_NUEVO " + CStr(var_consecutivo) + ",'" + Me.txt_almacen + "','" + Me.txt_codigo + "'," + var_fecha_inicio + ", " + var_fecha_fin, cnn, adOpenDynamic, adLockOptimistic
                  rs.Open "delete from TB_TEMP_KARDEX where INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND VCHA_aRT_ARTICULO_ID IS NULL", cnn, adOpenDynamic, adLockOptimistic
                  Set reporte = appl.OpenReport(App.Path + "\REP_KARDEX_NUEVO.rpt")
                  reporte.RecordSelectionFormula = "{VW_TEMP_KARDEX.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Reporte de Kardex"
                  frmvistasprevias.Show 1
                  Set reporte = Nothing
                  var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     Set reporte = appl.OpenReport(App.Path + "\REP_KARDEX_NUEVO.rpt")
                     reporte.RecordSelectionFormula = "{VW_TEMP_KARDEX.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     reporte.ExportOptions.FormatType = crEFTExcel80
                     reporte.ExportOptions.DestinationType = crEDTDiskFile
                     archivo = "c:\reportessid\Reporte_kardex_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                     reporte.ExportOptions.DiskFileName = archivo
                     reporte.Export False
                     Set reporte = Nothing
                     MsgBox "Se a terminado de guardar el archivo " + archivo
                  End If
                  
                  
                  
                  rs.Open "delete from TB_TEMP_KARDEX where INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
               Else
                  MsgBox "La fecha de inicio debe de ser mayor", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "Fecha Final Incorrecta", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Fecha de Inicio Incorrecta", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Código incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Clave de almacén incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 2500
   Left = 3000
   txt_inicio = Date
   txt_fin = Date
   Me.frm_lista.Visible = False
   var_tipo_lista = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_kardex)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_tipo_lista = 1 Then
         If Me.lv_lista.ListItems.Count > 0 Then
            Me.txt_almacen = lv_lista.selectedItem
            Me.txt_nombre_almacen = Me.lv_lista.selectedItem.SubItems(1)
            var_tipo_lista = 0
            Me.txt_almacen.SetFocus
         End If
      End If
      If var_tipo_lista = 2 Then
         If Me.lv_lista.ListItems.Count > 0 Then
            Me.txt_codigo = lv_lista.selectedItem
            Me.txt_nombre_articulo = Me.lv_lista.selectedItem.SubItems(1)
            var_tipo_lista = 0
            Me.txt_codigo.SetFocus
         End If
      End If
   End If
   If KeyAscii = 27 Then
      If var_tipo_lista = 1 Then
         Me.txt_almacen.SetFocus
         var_tipo_lista = 0
      Else
         Me.txt_codigo.SetFocus
         var_tipo_lista = 0
      End If
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub txt_almacen_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_ALMACENES order by VCHA_ALM_ALMACEN_ID", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ALM_ALMACEN_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ALMACENES"
      var_tipo_lista = 1
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
   If KeyAscii = 13 Then
      Me.txt_nombre_almacen.SetFocus
   End If
End Sub

Private Sub txt_almacen_LostFocus()
   If Me.txt_almacen <> "" Then
      rs.Open "select * from tb_almacenes where vcha_alm_almacen_id = '" + Me.txt_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_almacen = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
      Else
         MsgBox "Clave de almacén incorrecta", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      Me.txt_nombre_almacen = ""
   End If
End Sub

Private Sub txt_codigo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_aRTICULOS order by VCHA_aRT_ARTICULO_ID", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_Art_articulo_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ARTICULOS"
      var_tipo_lista = 2
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

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_codigo_LostFocus()
   If Me.txt_codigo <> "" Then
      rs.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_ART_aRTICULO_ID = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_articulo = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
      Else
         rsaux.Open "SELECT * FROM TB_EQUIVALENCIAS WHERE VCHA_EQU_CODIGO_EQUIVALENTE = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            rsaux1.Open "SELECT * FROM TB_aRTICULOS WHERE VCHA_aRT_ARTICULO_ID = '" + rsaux!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux1.EOF Then
               Me.txt_codigo = IIf(IsNull(rsaux1!vcha_Art_articulo_id), "", rsaux1!vcha_Art_articulo_id)
               Me.txt_nombre_articulo = IIf(IsNull(rsaux1!vcha_art_nombre_español), "", rsaux1!vcha_art_nombre_español)
            Else
               Me.txt_nombre_articulo = ""
               Me.txt_codigo = ""
               MsgBox "Clave de artículo no existe", vbOKOnly, "ATENCION"
            End If
            rsaux1.Close
         Else
            Me.txt_nombre_articulo = ""
            Me.txt_codigo = ""
            MsgBox "Clave de artículo no existe", vbOKOnly, "ATENCION"
         End If
         rsaux.Close
      End If
      rs.Close
   Else
      Me.txt_nombre_articulo = ""
   End If

End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fin) Then
         frmcalendario.mes = CDate(Me.txt_fin)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      txt_fin = var_fecha_general
   End If
End Sub

Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_inicio) Then
         frmcalendario.mes = CDate(Me.txt_inicio)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      txt_inicio = var_fecha_general
   End If
End Sub

Private Sub txt_nombre_almacen_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_almacen_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_ALMACENES WHERE VCHA_eMP_EMPRESA_ID = '" + var_empresa + "' order by VCHA_ALM_ALMACEN_ID", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ALM_ALMACEN_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_ALM_NOMBRE), "", rs!VCHA_ALM_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "ALMACENES"
      var_tipo_lista = 1
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
    Call pro_enfoque(KeyAscii)
End Sub
