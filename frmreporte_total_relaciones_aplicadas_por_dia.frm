VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmreporte_total_relaciones_aplicadas_por_dia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Total relaciones aplicadas por dia"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   " Periodo"
      Height          =   660
      Left            =   75
      TabIndex        =   5
      Top             =   4650
      Width           =   7020
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   4200
         TabIndex        =   8
         Top             =   180
         Width           =   1095
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   3885
         TabIndex        =   9
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   1350
         TabIndex        =   7
         Top             =   240
         Width           =   420
      End
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmreporte_total_relaciones_aplicadas_por_dia.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6795
      Picture         =   "frmreporte_total_relaciones_aplicadas_por_dia.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   90
      Left            =   45
      TabIndex        =   2
      Top             =   300
      Width           =   7155
   End
   Begin VB.Frame frm_lista 
      Height          =   4245
      Left            =   75
      TabIndex        =   0
      Top             =   360
      Width           =   7020
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1380
         Picture         =   "frmreporte_total_relaciones_aplicadas_por_dia.frx":073C
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   720
         Picture         =   "frmreporte_total_relaciones_aplicadas_por_dia.frx":0952
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Marcar (Enter)"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1050
         Picture         =   "frmreporte_total_relaciones_aplicadas_por_dia.frx":0B9C
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   60
         Picture         =   "frmreporte_total_relaciones_aplicadas_por_dia.frx":0C6E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   390
         Picture         =   "frmreporte_total_relaciones_aplicadas_por_dia.frx":0D70
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   135
         Width           =   330
      End
      Begin VB.Frame Frame6 
         Height          =   120
         Left            =   15
         TabIndex        =   10
         Top             =   450
         Width           =   6975
      End
      Begin MSComctlLib.ListView lv_lista 
         Height          =   3480
         Left            =   45
         TabIndex        =   1
         Top             =   660
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   6138
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
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   10231
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
      End
   End
End
Attribute VB_Name = "frmreporte_total_relaciones_aplicadas_por_dia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
   If IsDate(Me.txt_inicio) Then
      If IsDate(Me.txt_fin) Then
         If CDate(Me.txt_inicio) <= CDate(Me.txt_fin) Then
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) from TB_REPORTE_TOTAL_RELACIONES_APLICADAS_POR_DIA"
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0))
            Else
               var_consecutivo = 0
            End If
            rs.Close
            var_consecutivo = var_consecutivo + 1
            rs.Open "insert into TB_REPORTE_TOTAL_RELACIONES_APLICADAS_POR_DIA (inte_tem_consecutivo ) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            
            cnn.CommitTrans
            dia = CStr(Day(CDate(Me.txt_inicio)))
            mes = CStr(Month(CDate(Me.txt_inicio)))
            año = CStr(Year(CDate(Me.txt_inicio)))
            If Len(dia) = 1 Then
               dia = "0" + dia
            End If
            If Len(mes) = 1 Then
               mes = "0" + mes
            End If
            If Len(año) = 1 Then
               año = "200" + año
            Else
               If Len(año) = 2 Then
                  año = "20" + año
               End If
            End If
            var_fecha_inicio = "{d '" + año + "-" + mes + "-" + dia + "'}"
            
            
            
            dia = CStr(Day(CDate(Me.txt_fin)))
            mes = CStr(Month(CDate(Me.txt_fin)))
            año = CStr(Year(CDate(Me.txt_fin)))
            If Len(dia) = 1 Then
               dia = "0" + dia
            End If
            If Len(mes) = 1 Then
               mes = "0" + mes
            End If
            If Len(año) = 1 Then
               año = "200" + año
            Else
               If Len(año) = 2 Then
                  año = "20" + año
               End If
            End If
            var_fecha_fin = "{d '" + año + "-" + mes + "-" + dia + "'}"
            
            var_filtro = ""
            For var_j = 1 To lv_lista.ListItems.Count
                lv_lista.ListItems.Item(var_j).Selected = True
                If lv_lista.selectedItem.SubItems(2) = "*" Then
                   If var_filtro = "" Then
                      var_filtro = " and ({TB_REPORTE_TOTAL_RELACIONES_APLICADAS_POR_DIA.vcha_age_agente_id} = '" + lv_lista.selectedItem + "'"
                   Else
                      var_filtro = var_filtro + " or {TB_REPORTE_TOTAL_RELACIONES_APLICADAS_POR_DIA.vcha_age_agente_id} = '" + lv_lista.selectedItem + "'"
                   End If
                End If
            Next var_j
            If var_filtro <> "" Then
               var_filtro = var_filtro + ")"
            End If
            
            var_cadena = "INSERT INTO TB_REPORTE_TOTAL_RELACIONES_APLICADAS_POR_DIA (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN,VCHA_AGE_AGENTE_ID, VCHA_AGE_NOMBRE, VCHA_RUT_NOMBRE, FLOA_TEM_IMPORTE, VCHA_eMP_EMPRESA_ID, VCHA_EMP_NOMBRE, VCHA_RCO_FOLIO, DTIM_RCO_FECHA)  "
            'var_cadena = var_cadena + " SELECT  " + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin + ",vcha_age_agente_id, VCHA_AGE_NOMBRE, VCHA_RUT_NOMBRE,sum(IMPORTE) FROM VW_RC_PORFECHA  WHERE  DTIM_RCO_FECHA_INSERCION >= " + var_fecha_inicio + " and DTIM_RCO_FECHA_INSERCION <= " + var_fecha_fin + "+1-.0000001 group by vcha_age_agente_id, VCHA_AGE_NOMBRE, VCHA_RUT_NOMBRE"
            var_cadena = var_cadena + " SELECT       " + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin + ", VCHA_AGE_AGENTE_ID, VCHA_AGE_NOMBRE, VCHA_RUT_NOMBRE, SUM(IMPORTE) AS Expr4, VCHA_EMP_EMPRESA_ID, VCHA_EMP_NOMBRE, FOLIO, FECHA From dbo.VW_RC_PORFECHA WHERE DTIM_RCO_FECHA_INSERCION >= " + var_fecha_inicio + " and DTIM_RCO_FECHA_INSERCION <= " + var_fecha_fin + "+1-.0000001 GROUP BY VCHA_AGE_AGENTE_ID, VCHA_AGE_NOMBRE, VCHA_RUT_NOMBRE, VCHA_EMP_EMPRESA_ID, VCHA_EMP_NOMBRE, FOLIO, FECHA "
            'MsgBox var_cadena
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            rs.Open "delete from TB_REPORTE_TOTAL_RELACIONES_APLICADAS_POR_DIA where INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " and vcha_age_agente_id is null", cnn, adOpenDynamic, adLockOptimistic
            Set reporte = appl.OpenReport(App.Path + "\rep_total_relaciones_aplicadas_por_dia.rpt")
            reporte.RecordSelectionFormula = "{TB_REPORTE_TOTAL_RELACIONES_APLICADAS_POR_DIA.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + var_filtro
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
               reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "RELACIONES APLICADAS"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            
            var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_total_relaciones_aplicadas_por_dia.rpt")
               reporte.RecordSelectionFormula = "{TB_REPORTE_TOTAL_RELACIONES_APLICADAS_POR_DIA.INTE_tem_CONSECUTIVO} = " + CStr(var_consecutivo) + var_filtro
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\relaciones_aplicadas" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
            
            
            rs.Open "delete from TB_REPORTE_TOTAL_RELACIONES_APLICADAS_POR_DIA where inte_tem_consecutivo = " + CStr(var_consecutivo)
         Else
            MsgBox "La fecha final debe de ser mayor a la fecha de inicio", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha inicial Incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_invertir_Click()
   n = lv_lista.ListItems.Count
   For i = 1 To n
      lv_lista.ListItems.Item(i).Selected = True
      If lv_lista.selectedItem.SubItems(2) = "*" Then
         lv_lista.selectedItem.SubItems(2) = ""
         lv_lista.ListItems.Item(i).Bold = False
         lv_lista.ListItems.Item(i).ForeColor = &H80000012
         lv_lista.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_lista.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_lista.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_lista.selectedItem.SubItems(2) = "*"
         lv_lista.ListItems.Item(i).Bold = True
         lv_lista.ListItems.Item(i).ForeColor = &HFF0000
         lv_lista.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_lista.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_lista.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub cmd_marcar_Click()
   i = lv_lista.selectedItem.Index
   If lv_lista.selectedItem.SubItems(2) = "*" Then
      lv_lista.selectedItem.SubItems(2) = ""
      lv_lista.ListItems.Item(i).Bold = False
      lv_lista.ListItems.Item(i).ForeColor = &H80000012
      lv_lista.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_lista.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_lista.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_lista.Refresh
   Else
      lv_lista.selectedItem.SubItems(2) = "*"
      lv_lista.ListItems.Item(i).Bold = True
      lv_lista.ListItems.Item(i).ForeColor = &HFF0000
      lv_lista.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_lista.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_lista.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_lista.Refresh
   End If
End Sub

Private Sub cmd_ninguno_Click()
   n = lv_lista.ListItems.Count
   For i = 1 To n
      lv_lista.ListItems.Item(i).Selected = True
      lv_lista.selectedItem.SubItems(2) = ""
      lv_lista.ListItems.Item(i).Bold = False
      lv_lista.ListItems.Item(i).ForeColor = &H80000012
      lv_lista.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_lista.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_lista.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   lv_lista.Refresh

End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmd_seleccion_Click()
   n = lv_lista.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_lista.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_lista.selectedItem.SubItems(2) = "" And var_rellena = True Then
         lv_lista.selectedItem.SubItems(2) = "*"
         lv_lista.ListItems.Item(i).Bold = True
         lv_lista.ListItems.Item(i).ForeColor = &HFF0000
         lv_lista.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_lista.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_lista.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_lista.selectedItem.SubItems(2) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_lista.selectedItem.SubItems(2) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub cmd_todos_Click()
   n = lv_lista.ListItems.Count
   For i = 1 To n
      lv_lista.ListItems.Item(i).Selected = True
      lv_lista.selectedItem.SubItems(2) = "*"
      lv_lista.ListItems.Item(i).Bold = True
      lv_lista.ListItems.Item(i).ForeColor = &HFF0000
      lv_lista.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_lista.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_lista.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
   Next i
   lv_lista.Refresh
End Sub

Private Sub Form_Load()
   Top = 1000
   Left = 2000
   Dim list_item As ListItem
   rs.Open "select vcha_age_agente_id, vcha_age_nombre from tb_agentes order by vcha_age_nombre", cnn, adOpenDynamic, adLockOptimistic
   numero_items_lineas = 0
   While Not rs.EOF
      Set list_item = lv_lista.ListItems.Add(, , rs(0).Value)
      list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
      rs.MoveNext:
      numero_items_lineas = numero_items_lineas + 1
    Wend
    rs.Close
    Me.txt_inicio = Date
    Me.txt_fin = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         i = lv_lista.selectedItem.Index
         If lv_lista.selectedItem.SubItems(2) = "*" Then
            lv_lista.selectedItem.SubItems(2) = ""
            lv_lista.ListItems.Item(i).Bold = False
            lv_lista.ListItems.Item(i).ForeColor = &H80000012
            lv_lista.ListItems.Item(i).ListSubItems(1).Bold = False
            lv_lista.ListItems.Item(i).ListSubItems(2).Bold = False
            lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
            lv_lista.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
            lv_lista.Refresh
         Else
            lv_lista.selectedItem.SubItems(2) = "*"
            lv_lista.ListItems.Item(i).Bold = True
            lv_lista.ListItems.Item(i).ForeColor = &HFF0000
            lv_lista.ListItems.Item(i).ListSubItems(1).Bold = True
            lv_lista.ListItems.Item(i).ListSubItems(2).Bold = True
            lv_lista.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
            lv_lista.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
            lv_lista.Refresh
         End If
      End If
   End If
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fin) Then
         frmcalendario.mes.Value = CDate(Me.txt_fin)
      Else
         frmcalendario.mes.Value = Date
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
