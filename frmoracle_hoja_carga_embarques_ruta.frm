VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_hoja_carga_embarques_ruta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hoja de carga para embarques en ruta"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   12450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_excel 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   720
      Picture         =   "frmoracle_hoja_carga_embarques_ruta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Generar archivo Beetrack"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmoracle_hoja_carga_embarques_ruta.frx":0312
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Guardar"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   60
      Left            =   75
      TabIndex        =   18
      Top             =   345
      Width           =   12345
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   375
      Picture         =   "frmoracle_hoja_carga_embarques_ruta.frx":0414
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   12015
      Picture         =   "frmoracle_hoja_carga_embarques_ruta.frx":0516
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   1950
      Left            =   90
      TabIndex        =   10
      Top             =   390
      Width           =   12300
      Begin VB.TextBox txt_color 
         Height          =   375
         Left            =   7395
         TabIndex        =   5
         Top             =   645
         Width           =   1485
      End
      Begin VB.TextBox txt_anden 
         Height          =   375
         Left            =   7395
         TabIndex        =   4
         Top             =   225
         Width           =   1470
      End
      Begin VB.TextBox txt_observaciones 
         Height          =   375
         Left            =   1335
         TabIndex        =   3
         Top             =   1470
         Width           =   10875
      End
      Begin VB.TextBox txt_unidad 
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   1065
         Width           =   4980
      End
      Begin VB.TextBox txt_ruta 
         Height          =   375
         Left            =   1335
         TabIndex        =   1
         Top             =   645
         Width           =   4980
      End
      Begin VB.TextBox txt_embarque 
         Height          =   375
         Left            =   1335
         TabIndex        =   0
         Top             =   225
         Width           =   1380
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones:"
         Height          =   195
         Left            =   90
         TabIndex        =   16
         Top             =   1560
         Width           =   1110
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Color:"
         Height          =   195
         Left            =   6660
         TabIndex        =   15
         Top             =   675
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Anden:"
         Height          =   195
         Left            =   6660
         TabIndex        =   14
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Unidad:"
         Height          =   195
         Left            =   90
         TabIndex        =   13
         Top             =   1155
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ruta:"
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   735
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Embarque:"
         Height          =   195
         Left            =   90
         TabIndex        =   11
         Top             =   315
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4905
      Left            =   90
      TabIndex        =   9
      Top             =   2295
      Width           =   12315
      Begin MSComctlLib.ListView lv_pedidos 
         CausesValidation=   0   'False
         Height          =   4650
         Left            =   60
         TabIndex        =   17
         Top             =   165
         Width           =   12165
         _ExtentX        =   21458
         _ExtentY        =   8202
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
            Text            =   "Pedido"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cliente"
            Object.Width           =   14464
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Orden"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Piezas"
            Object.Width           =   2187
         EndProperty
      End
   End
End
Attribute VB_Name = "frmoracle_hoja_carga_embarques_ruta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_guardar_Click()
   rs.Open "SELECT * FROM TB_ORACLE_ORDEN_CARGA_EMBARQUES_RUTA WHERE EMBARQUE = '" + Me.txt_embarque + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      rsaux.Open "UPDATE TB_ORACLE_ORDEN_CARGA_EMBARQUES_RUTA SET UNIDAD ='" + Me.txt_unidad + "', RUTA = '" + Me.txt_ruta + "', ANDEN = '" + Me.txt_anden + "', COLOR = '" + Me.txt_color + "', OBSERVACIONES = '" + Me.txt_observaciones + "' WHERE EMBARQUE = '" + Me.txt_embarque + "'", cnn, adOpenDynamic, adLockOptimistic
   Else
      rsaux.Open "INSERT INTO TB_ORACLE_ORDEN_CARGA_EMBARQUES_RUTA (EMBARQUE, UNIDAD, RUTA, ANDEN, COLOR, OBSERVACIONES) VALUES ('" + txt_embarque + "', '" + txt_unidad + "', '" + txt_ruta + "', '" + txt_anden + "', '" + txt_color + "', '" + txt_observaciones + "')", cnn, adOpenDynamic, adLockOptimistic
   End If
   MsgBox "Se a guardado la información", vbOKOnly, "ATENCION"
   rs.Close
End Sub

Private Sub cmd_imprimir_Click()
   If Me.lv_pedidos.ListItems.Count > 0 Then
      cnn.BeginTrans
      
      rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_ORACLE_REPORTE_HOJA_CARGA_EMBARQUES_RUTA", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux.EOF Then
         var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
      Else
         var_consecutivo = 1
      End If
      rsaux.Close
      rsaux1.Open "insert into TB_TEMP_ORACLE_REPORTE_HOJA_CARGA_EMBARQUES_RUTA (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
      cnn.CommitTrans
      VAR_TOTA_PEDIDOS = 0
      For var_j = 1 To Me.lv_pedidos.ListItems.Count
          Me.lv_pedidos.ListItems.Item(var_j).Selected = True
          rsaux1.Open "insert into TB_TEMP_ORACLE_REPORTE_HOJA_CARGA_EMBARQUES_RUTA (inte_tem_consecutivo, embarque, ruta, unidad, observaciones, anden, color, pedido, cliente, orden, piezas) values (" + CStr(var_consecutivo) + ",'" + Me.txt_embarque + "','" + Me.txt_ruta + "','" + Me.txt_unidad + "','" + Me.txt_observaciones + "','" + Me.txt_anden + "','" + Me.txt_color + "','" + Me.lv_pedidos.selectedItem + "','" + Me.lv_pedidos.selectedItem.SubItems(1) + "'," + Me.lv_pedidos.selectedItem.SubItems(2) + "," + Me.lv_pedidos.selectedItem.SubItems(3) + ")", cnn, adOpenDynamic, adLockOptimistic
          VAR_TOTAL_PEDIDOS = VAR_TOTAL_PEDIDOS + 1
      Next var_j
      rsaux.Open "delete from TB_TEMP_ORACLE_REPORTE_HOJA_CARGA_EMBARQUES_RUTA where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and embarque is null", cnn, adOpenDynamic, adLockOptimistic
      rsaux.Open "UPDATE TB_TEMP_ORACLE_REPORTE_HOJA_CARGA_EMBARQUES_RUTA SET TOTAL_PEDIDOS = " + CStr(VAR_TOTAL_PEDIDOS) + " WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            
      Set reporte = appl.OpenReport(App.Path + "\rep_oracle_hoja_orden_carga.rpt")
      reporte.RecordSelectionFormula = "{VW_ORACLE_REPORTE_HOJA_ORDEN.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
      frmvistasprevias.cr.ReportSource = reporte
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      frmvistasprevias.cr.ViewReport
      frmvistasprevias.Caption = "Hoja de carga para embarques en ruta"
      frmvistasprevias.Show 1
      Set reporte = Nothing
      
      
      
      var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         Set reporte = appl.OpenReport(App.Path + "\rep_oracle_hoja_orden_carga.rpt")
         reporte.RecordSelectionFormula = "{VW_ORACLE_REPORTE_HOJA_ORDEN.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         reporte.ExportOptions.FormatType = crEFTExcel80
         reporte.ExportOptions.DestinationType = crEDTDiskFile
         archivo = "c:\reportessid\rep_hoja_carga_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
         reporte.ExportOptions.DiskFileName = archivo
         reporte.Export False
         Set reporte = Nothing
         MsgBox "Se a terminado de guardar el archivo " + archivo
      End If
      
      
      
      
      
      
   Else
      MsgBox "No existen pedidos en el embarque seleccionado", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Me.txt_embarque = var_embarque_global
   rsaux.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen, cantidad_sin_catalogos from tb_oracle_pedidos_asignados_embarques where embarque = " + Me.txt_embarque + " order by orden_pedido, PEDIDO", cnn, adOpenDynamic, adLockOptimistic
   While Not rsaux.EOF
         Set list_item = lv_pedidos.ListItems.Add(, , rsaux!pedido)
         var_cliente = IIf(IsNull(rsaux!Cliente), "", rsaux!Cliente)
         If rsaux!Cliente = "VIANNEY TEXTIL HOGAR SA DE CV" Then
            var_cliente = IIf(IsNull(rsaux!nombre_Agente), "", rsaux!nombre_Agente)
         End If
         list_item.SubItems(1) = var_cliente
         list_item.SubItems(2) = rsaux!orden_pedido
         list_item.SubItems(3) = IIf(IsNull(rsaux!CANTIDAD_SIN_CATALOGOS), 0, rsaux!CANTIDAD_SIN_CATALOGOS)
         rsaux.MoveNext
   Wend
   rsaux.Close
   rs.Open "SELECT * FROM TB_ORACLE_ORDEN_CARGA_EMBARQUES_RUTA WHERE EMBARQUE = '" + Me.txt_embarque + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      Me.txt_anden = IIf(IsNull(rs!ANDEN), "", rs!ANDEN)
      Me.txt_color = IIf(IsNull(rs!Color), "", rs!Color)
      Me.txt_observaciones = IIf(IsNull(rs!observaciones), "", rs!observaciones)
      Me.txt_ruta = IIf(IsNull(rs!ruta), "", rs!ruta)
      Me.txt_unidad = IIf(IsNull(rs!unidad), "", rs!unidad)
   End If
   
   rs.Close
End Sub

Private Sub txt_anden_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_color_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_observaciones_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_ruta_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_unidad_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub
