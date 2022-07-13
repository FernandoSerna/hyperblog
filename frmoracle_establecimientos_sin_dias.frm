VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_establecimientos_sin_dias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Establecimientos sin dias de despacho"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   420
      Picture         =   "frmoracle_establecimientos_sin_dias.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Establecimientos sin dias de despacho"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmoracle_establecimientos_sin_dias.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Establecimientos y dias de despacho"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5475
      Picture         =   "frmoracle_establecimientos_sin_dias.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   5655
      Left            =   90
      TabIndex        =   0
      Top             =   450
      Width           =   5715
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   375
         Picture         =   "frmoracle_establecimientos_sin_dias.frx":083E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   45
         Picture         =   "frmoracle_establecimientos_sin_dias.frx":0A54
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1035
         Picture         =   "frmoracle_establecimientos_sin_dias.frx":0B56
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   705
         Picture         =   "frmoracle_establecimientos_sin_dias.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Marcar (Enter)"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1365
         Picture         =   "frmoracle_establecimientos_sin_dias.frx":0E72
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   135
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_agentes 
         Height          =   5100
         Left            =   45
         TabIndex        =   6
         Top             =   480
         Width           =   5610
         _ExtentX        =   9895
         _ExtentY        =   8996
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
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "marca"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   60
      TabIndex        =   9
      Top             =   300
      Width           =   5790
   End
End
Attribute VB_Name = "frmoracle_establecimientos_sin_dias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_imprimir_Click()
   var_cadena_agentes = ""
   For var_j = 1 To Me.lv_agentes.ListItems.Count
       Me.lv_agentes.ListItems.Item(var_j).Selected = True
       If Me.lv_agentes.selectedItem.SubItems(2) = "*" Then
          If var_cadena_agentes = "" Then
             var_cadena_agentes = Me.lv_agentes.selectedItem
          Else
             var_cadena_agentes = var_cadena_agentes + ", " + Me.lv_agentes.selectedItem
          End If
       End If
   Next var_j
   If var_cadena_agentes <> "" Then
      var_cadena = "select distinct  arc.name as ANC, account_number TITULAR, account_full_name NOMBRE_TITULAR, party_site_number CLAVE, a.RAZON_SOCIAL_CLIENTE NOMBRE, a.calle||' '||a.num_calle||' '||a.colonia||' '||a.CIUDAD||' '||a.MUNICIPIO||' '||a.ESTADO||' '||a.CODIGO_POSTAL  as DIRECCION, a.ATTRIBUTE3 AS ACTIVO, B.RUTA, B.NOMBRE_RUTA from XXVIA_VW_CLIENTES_BCP a, hz_customer_profiles hcp, ar_collectors Arc, XXVIA_VW_DIAS_DESPACHO B Where A.site_use_id = hcp.site_use_id AND arc.collector_id  = hcp.collector_id AND hcp.collector_id  = (" + var_cadena_agentes + ") and a.SITE_USE_CODE = 'SHIP_TO' AND TO_CHAR(a.SITE_USE_ID) = B.SITE_USE_ID(+) ORDER BY ARC.NAME, A.RAZON_SOCIAL_CLIENTE"
      rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Set oExcel = CreateObject("Excel.Application")
         Set oWBook = oExcel.Workbooks.Add
         Set oSheet = oWBook.Worksheets(1)
         var_cadena = "PERIODO DEL " + Replace(var_fecha_inicio, "/", "_") + " AL " + Replace(var_fecha_fin, "/", "_")
         'MsgBox var_cadena
         oSheet.Name = "Reporte"
         Screen.MousePointer = vbHourglass
         iFila = 1
         iFila2 = 1
         iCol2 = 1
         iCol = 1
         'rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         For i = 0 To rs.Fields.Count - 1
             oSheet.Cells(iFila, i + 1) = rs.Fields(i).Name
         Next
         iFila = iFila + 1
         With oSheet
              ' carga los registros del recordset
              .Cells(iFila, iCol).CopyFromRecordset rs
              'oExcel.Columns(1).Select
              'oExcel.Selection.NumberFormat = "#,##0.00"
              'oExcel.Columns(1).Select
              'oExcel.Selection.Font.Color = vbRed
              .Columns.AutoFit ' ajusta el ancho de las columnas
         End With
         oWBook.SaveAs "c:\reportessid\reporte_establecimientos_sin_dias_despacho_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
         oExcel.Visible = True
         Set oExcel = Nothing
         Screen.MousePointer = vbDefault
      Else
         MsgBox "No existen establecimientos para los ANC seleccionados", vbOKOnly, "ATENCION"
      End If
            rs.Close
      
   Else
      MsgBox "No se selecciono ningun ANC", vbOKOnly, "ATENCION"
   End If

End Sub

Private Sub cmd_invertir_Click()
   n = lv_agentes.ListItems.Count
   For i = 1 To n
      lv_agentes.ListItems.Item(i).Selected = True
      If lv_agentes.selectedItem.SubItems(2) = "*" Then
         lv_agentes.selectedItem.SubItems(2) = ""
         lv_agentes.ListItems.Item(i).Bold = False
         lv_agentes.ListItems.Item(i).ForeColor = &H80000012
         lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_agentes.selectedItem.SubItems(2) = "*"
         lv_agentes.ListItems.Item(i).Bold = True
         lv_agentes.ListItems.Item(i).ForeColor = &HFF0000
         lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub cmd_marcar_Click()
   i = lv_agentes.selectedItem.Index
   If lv_agentes.selectedItem.SubItems(2) = "*" Then
      lv_agentes.selectedItem.SubItems(2) = ""
      lv_agentes.ListItems.Item(i).Bold = False
      lv_agentes.ListItems.Item(i).ForeColor = &H80000012
      lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_agentes.Refresh
   Else
      lv_agentes.selectedItem.SubItems(2) = "*"
      lv_agentes.ListItems.Item(i).Bold = True
      lv_agentes.ListItems.Item(i).ForeColor = &HFF0000
      lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_agentes.Refresh
   End If

End Sub

Private Sub cmd_ninguno_Click()
   n = lv_agentes.ListItems.Count
   For i = 1 To n
      lv_agentes.ListItems.Item(i).Selected = True
      lv_agentes.selectedItem.SubItems(2) = ""
      lv_agentes.ListItems.Item(i).Bold = False
      lv_agentes.ListItems.Item(i).ForeColor = &H80000012
      lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   lv_agentes.Refresh

End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmd_seleccion_Click()
   n = lv_agentes.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_agentes.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_agentes.selectedItem.SubItems(2) = "" And var_rellena = True Then
         lv_agentes.selectedItem.SubItems(2) = "*"
         lv_agentes.ListItems.Item(i).Bold = True
         lv_agentes.ListItems.Item(i).ForeColor = &HFF0000
         lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_agentes.selectedItem.SubItems(2) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_agentes.selectedItem.SubItems(2) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i

End Sub

Private Sub cmd_todos_Click()
   n = lv_agentes.ListItems.Count
   For i = 1 To n
      lv_agentes.ListItems.Item(i).Selected = True
      lv_agentes.selectedItem.SubItems(2) = "*"
      lv_agentes.ListItems.Item(i).Bold = True
      lv_agentes.ListItems.Item(i).ForeColor = &HFF0000
      lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
   Next i
   lv_agentes.Refresh

End Sub

Private Sub Command1_Click()
   var_cadena_agentes = ""
   For var_j = 1 To Me.lv_agentes.ListItems.Count
       Me.lv_agentes.ListItems.Item(var_j).Selected = True
       If Me.lv_agentes.selectedItem.SubItems(2) = "*" Then
          If var_cadena_agentes = "" Then
             var_cadena_agentes = Me.lv_agentes.selectedItem
          Else
             var_cadena_agentes = var_cadena_agentes + ", " + Me.lv_agentes.selectedItem
          End If
       End If
   Next var_j
   If var_cadena_agentes <> "" Then
      var_cadena = "select distinct  arc.name as ANC, account_number TITULAR, account_full_name NOMBRE_TITULAR, party_site_number CLAVE, a.RAZON_SOCIAL_CLIENTE NOMBRE, a.calle||' '||a.num_calle||' '||a.colonia||' '||a.CIUDAD||' '||a.MUNICIPIO||' '||a.ESTADO||' '||a.CODIGO_POSTAL  as DIRECCION, a.ATTRIBUTE3 AS ACTIVO, B.RUTA, B.NOMBRE_RUTA from XXVIA_VW_CLIENTES_BCP a, hz_customer_profiles hcp, ar_collectors Arc, XXVIA_VW_DIAS_DESPACHO B Where A.site_use_id = hcp.site_use_id AND arc.collector_id  = hcp.collector_id AND hcp.collector_id  = (" + var_cadena_agentes + ") and a.SITE_USE_CODE = 'SHIP_TO' AND TO_CHAR(a.SITE_USE_ID) = B.SITE_USE_ID(+) and nvl(b.ruta,' ') = ' '  ORDER BY ARC.NAME, A.RAZON_SOCIAL_CLIENTE"
      rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Set oExcel = CreateObject("Excel.Application")
         Set oWBook = oExcel.Workbooks.Add
         Set oSheet = oWBook.Worksheets(1)
         var_cadena = "PERIODO DEL " + Replace(var_fecha_inicio, "/", "_") + " AL " + Replace(var_fecha_fin, "/", "_")
         'MsgBox var_cadena
         oSheet.Name = "Reporte"
         Screen.MousePointer = vbHourglass
         iFila = 1
         iFila2 = 1
         iCol2 = 1
         iCol = 1
         'rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         For i = 0 To rs.Fields.Count - 1
             oSheet.Cells(iFila, i + 1) = rs.Fields(i).Name
         Next
         iFila = iFila + 1
         With oSheet
              ' carga los registros del recordset
              .Cells(iFila, iCol).CopyFromRecordset rs
              'oExcel.Columns(1).Select
              'oExcel.Selection.NumberFormat = "#,##0.00"
              'oExcel.Columns(1).Select
              'oExcel.Selection.Font.Color = vbRed
              .Columns.AutoFit ' ajusta el ancho de las columnas
         End With
         oWBook.SaveAs "c:\reportessid\reporte_establecimientos_sin_dias_despacho_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
         oExcel.Visible = True
         Set oExcel = Nothing
         Screen.MousePointer = vbDefault
      Else
         MsgBox "No existen establecimientos para los ANC seleccionados", vbOKOnly, "ATENCION"
      End If
            rs.Close
      
   Else
      MsgBox "No se selecciono ningun ANC", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Form_Load()
   Top = 800
   Left = 3000
   Dim list_item As ListItem
   rs.Open "select COLLECTOR_ID, NAME from ar_collectors", cnnoracle_4, adOpenDynamic, adLockOptimistic
   numero_items_permisos = 0
   While Not rs.EOF
      Set list_item = lv_agentes.ListItems.Add(, , rs!collector_id)
      list_item.SubItems(1) = IIf(IsNull(rs!Name), "", rs!Name)
      list_item.SubItems(2) = ""
      rs.MoveNext:
      numero_items_permisos = numero_items_permisos + 1
   Wend
   rs.Close

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_agentes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_agentes, ColumnHeader)
End Sub

Private Sub lv_agentes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_agentes.ListItems.Count > 0 Then
         i = lv_agentes.selectedItem.Index
         If lv_agentes.selectedItem.SubItems(2) = "*" Then
            lv_agentes.selectedItem.SubItems(2) = ""
            lv_agentes.ListItems.Item(i).Bold = False
            lv_agentes.ListItems.Item(i).ForeColor = &H80000012
            lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = False
            lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = False
            lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
            lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
            lv_agentes.Refresh
         Else
            lv_agentes.selectedItem.SubItems(2) = "*"
            lv_agentes.ListItems.Item(i).Bold = True
            lv_agentes.ListItems.Item(i).ForeColor = &HFF0000
            lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
            lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
            lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
            lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
            lv_agentes.Refresh
         End If
      End If
   End If
End Sub
