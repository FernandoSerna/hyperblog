VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_embarques_a_surtir_lista 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Embarques a surtir  (F5 para imprimir reporte)"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   2880
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   70
      TabIndex        =   2
      Top             =   3120
      Width           =   5535
      Begin VB.CommandButton cmd_buscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   4080
         TabIndex        =   7
         Top             =   210
         Width           =   1335
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   255
         Width           =   1140
      End
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   2865
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   255
         Width           =   1140
      End
      Begin WMPLibCtl.WindowsMediaPlayer wmp1 
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   -960
         Width           =   375
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "full"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   661
         _cy             =   661
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   315
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2325
         TabIndex        =   5
         Top             =   315
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   70
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.CommandButton cmd_silencio 
         Caption         =   "Sonido"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2760
         Width           =   5295
      End
      Begin MSComctlLib.ListView lv_embarques 
         Height          =   2580
         Left            =   45
         TabIndex        =   1
         Top             =   120
         Width           =   5445
         _ExtentX        =   9604
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Embarque"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   4586
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Estatus"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Piezas"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
      End
      Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
         Height          =   375
         Left            =   600
         TabIndex        =   8
         Top             =   2160
         Width           =   495
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "full"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   873
         _cy             =   661
      End
   End
End
Attribute VB_Name = "frmoracle_embarques_a_surtir_lista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim objConn As New ADODB.Connection
Dim objCmd As New ADODB.Command
Dim objParm As ADODB.Parameter
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter

Private Sub Command1_Click()
   
End Sub

Private Sub cmd_buscar_Click()
    Me.Timer1.Enabled = True
    If IsDate(Me.txt_inicio) Then
       If IsDate(Me.txt_fin) Then
          If CDate(Me.txt_inicio) <= CDate(Me.txt_fin) Then
          
          Else
             MsgBox "La fecha de inicio debe de ser inferior a la fecha final", vbOKOnly, "ATENCION"
          End If
       Else
          MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
       End If
    Else
       MsgBox "Fecha inicio incorrecta", vbOKOnly, "ATENCION"
    End If
End Sub

Private Sub cmd_silencio_Click()
   If Me.cmd_silencio.Caption = "Sonido" Then
      Me.cmd_silencio.Caption = "Silencio"
      Me.wmp1.Controls.stop
   Else
      Me.cmd_silencio.Caption = "Sonido"
   End If
   
End Sub

Private Sub Form_Load()
   Top = 1200
   Left = 3000
    Me.lv_embarques.ListItems.Clear
    Me.txt_inicio = Date
    Me.txt_fin = Date
    
    var_dia = CStr(Day(CDate(txt_inicio)))
    var_mes = CStr(Month(CDate(txt_inicio)))
    var_año = CStr(Year(CDate(txt_inicio)))
    If Len(Trim(var_dia)) = 1 Then
       var_dia = "0" + var_dia
    End If
    If Len(Trim(var_mes)) = 1 Then
       var_mes = "0" + var_mes
    End If
            
    VAR_FECHA_INICIO_TABLA = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
    var_fecha_fin_1 = CDate(txt_fin) + 1
    var_dia = CStr(Day(var_fecha_fin_1))
    var_mes = CStr(Month(var_fecha_fin_1))
    var_año = CStr(Year(var_fecha_fin_1))
    If Len(Trim(var_dia)) = 1 Then
       var_dia = "0" + var_dia
    End If
    If Len(Trim(var_mes)) = 1 Then
       var_mes = "0" + var_mes
    End If
    VAR_FECHA_FIN_TABLA = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
    
    
    rs.Open "SELECT * FROM TB_ORACLE_EMBARQUES_SURTIR WHERE ISNULL(ESTATUS,'') = '' and fecha >= " + VAR_FECHA_INICIO_TABLA + " and fecha < " + VAR_FECHA_FIN_TABLA, cnn, adOpenDynamic, adLockOptimistic
    'rs.Open "SELECT * FROM TB_ORACLE_EMBARQUES_SURTIR WHERE  fecha >= " + var_Fecha_inicio_tabla + " and fecha < " + VAR_FECHA_FIN_TABLA, cnn, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
       While Not rs.EOF
             Set list_item = Me.lv_embarques.ListItems.Add(, , rs!Embarque)
             list_item.SubItems(1) = rs!Fecha
             list_item.SubItems(2) = IIf(IsNull(rs!ESTATUS), "", rs!ESTATUS)
             list_item.SubItems(3) = IIf(IsNull(rs!piezas_surtir), "", rs!piezas_surtir)
             
             rs.MoveNext
       Wend
       Me.wmp1.URL = App.Path + "\Mec_Alarm_10.wav"
       Me.wmp1.Controls.play
       Me.Timer1.Enabled = True
    Else
       Me.wmp1.URL = App.Path + "\Mec_Alarm_10.wav"
       Me.wmp1.Controls.stop
       Me.Timer1.Enabled = True
       
    End If
    rs.Close
   
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_embarques_GotFocus()
   Me.Timer1.Enabled = False
End Sub

Private Sub lv_embarques_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If Me.lv_embarques.ListItems.Count > 0 Then
         var_Cadena_embarques = ""
         var_piezas_totales = 0
         For var_j = 1 To Me.lv_embarques.ListItems.Count
             Me.lv_embarques.ListItems.Item(var_j).Selected = True
             If Me.lv_embarques.selectedItem.SubItems(4) = "*" Then
                var_piezas_totales = var_piezas_totales + CDbl(Me.lv_embarques.selectedItem.SubItems(3))
                If var_Cadena_embarques = "" Then
                   var_Cadena_embarques = "'" + Me.lv_embarques.selectedItem
                Else
                   var_Cadena_embarques = var_Cadena_embarques + "','" + Me.lv_embarques.selectedItem
                End If
             End If
             
         Next var_j
         If Len(var_Cadena_embarques) > 0 Then
            
            var_Cadena_embarques = var_Cadena_embarques + "'"
            rs.Open "SELECT PEDIDO FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES WHERE EMBARQUE in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
            
               'strconsulta = "select * from xxvia_Tb_encabezado_embarques where embarque = ?"
               'With comandoORA
               '     .ActiveConnection = cnnoracle_4
               '     .CommandType = adCmdText
               '     .CommandText = strconsulta
               '     Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.lv_embarques.selectedItem))
               '     .Parameters.Append parametro
               'End With
               'Set rsaux1 = comandoORA.execute
               'Set comandoORA = Nothing
               'Set parametro = Nothing
               'If Not rsaux1.EOF Then
               '   var_fecha_inicio = CStr(IIf(IsNull(rsaux1!FECHA_INICIO), "", rsaux1!FECHA_INICIO))
               '   var_fecha_fin = CStr(IIf(IsNull(rsaux1!FECHA_FIN), "", rsaux1!FECHA_FIN))
               '   VAR_ESTATUS = IIf(IsNull(rsaux1!char_Emb_estatus), "", rsaux1!char_Emb_estatus)
               'Else
               '   var_fecha_inicio = ""
               '   var_fecha_fin = ""
               '   VAR_ESTATUS = ""
               'End If
               'rsaux1.Close
            
               var_Cadena_pedidos = ""
               While Not rs.EOF
                     If var_Cadena_pedidos = "" Then
                        var_Cadena_pedidos = CStr(rs!pedido)
                     Else
                        var_Cadena_pedidos = var_Cadena_pedidos + "," + CStr(rs!pedido)
                     End If
                     rs.MoveNext
               Wend
            
               'strconsulta = "select A.segment1 as codigo, A.item_description as descripcion, B.ATTRIBUTE2 AS UBICACION, sum(src_requested_quantity) as cantidad from xxvia_tb_pedidos_divididos A, XXVIA_SYSTEM_ITEMS_B B where source_header_number in (?) AND A.ORGANIZATION_ID = B.ORGANIZATION_ID AND A.SEGMENT1 = B.SEGMENT1 group by A.segment1, A.item_description, B.ATTRIBUTE2 ORDER BY ATTRIBUTE2"
               'With comandoORA
               '     .ActiveConnection = cnnoracle_4
               '     .CommandType = adCmdText
               '     .CommandText = strconsulta
               '     Set parametro = .CreateParameter(, adNumeric, adParamInput, 1000, var_cadena_pedidos)
               '     .Parameters.Append parametro
               'End With
               'Set rsaux1 = comandoORA.execute
               'Set comandoORA = Nothing
             
         
               'var_Cadena_pedidos = Me.lv_embarques.selectedItem
               rsaux1.Open "select A.segment1 as codigo, A.item_description as descripcion, B.ATTRIBUTE2 AS UBICACION, B.ATTRIBUTE3 AS EXCEDENTE,sum(src_requested_quantity) as cantidad from xxvia_tb_pedidos_divididos A, XXVIA_SYSTEM_ITEMS_B B where source_header_number in (" + var_Cadena_pedidos + ") AND A.ORGANIZATION_ID = B.ORGANIZATION_ID AND A.SEGMENT1 = B.SEGMENT1 group by A.segment1, A.item_description, B.ATTRIBUTE2, B.ATTRIBUTE3  ORDER BY ATTRIBUTE2", cnnoracle_4, adOpenDynamic, adLockOptimistic
               cnn.BeginTrans
               rsaux2.Open "SELECT MAX(INTE_tEM_CONSECUTIVO) FROM TB_TEMP_ORACLE_embarque_concentrado", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux2.EOF Then
                  var_consecutivo = IIf(IsNull(rsaux2(0).Value), 0, rsaux2(0).Value) + 1
               Else
                  var_consecutivo = 1
               End If
               rsaux2.Close
               rsaux2.Open "INSERT INTO TB_TEMP_ORACLE_EMBARQUE_concentrado (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
               cnn.CommitTrans
            
               While Not rsaux1.EOF
                     'rsaux2.Open "INSERT INTO TB_TEMP_ORACLE_EMBARQUE_concentrado (INTE_tEM_CONSECUTIVO, EMBARQUE, FECHA_INICIO, FECHA_FIN, ESTATUS, CODIGO, DESCRIPCION, UBICACION, CANTIDAD, PASILLO, excedente, piezas_totales) VALUES (" + CStr(var_consecutivo) + ",'" + Me.lv_embarques.selectedItem + "','" + Me.lv_embarques.selectedItem.SubItems(1) + "',NULL,'" + VAR_ESTATUS + "','" + rsaux1!CODIGO + "','" + rsaux1!DESCRIPCION + "','" + IIf(IsNull(rsaux1!ubicacion), "", rsaux1!ubicacion) + "'," + CStr(rsaux1!Cantidad) + ",'" + Mid(IIf(IsNull(rsaux1!ubicacion), "", rsaux1!ubicacion), 1, 3) + "','" + IIf(IsNull(rsaux1!excedente), "", rsaux1!excedente) + "'," + Me.lv_embarques.selectedItem.SubItems(3) + ")", cnn, adOpenDynamic, adLockOptimistic
                     rsaux2.Open "INSERT INTO TB_TEMP_ORACLE_EMBARQUE_concentrado (INTE_tEM_CONSECUTIVO, EMBARQUE, FECHA_INICIO, FECHA_FIN, ESTATUS, CODIGO, DESCRIPCION, UBICACION, CANTIDAD, PASILLO, excedente, piezas_totales) VALUES (" + CStr(var_consecutivo) + ",'" + Replace(var_Cadena_embarques, "'", "") + "',NULL,NULL,'" + VAR_ESTATUS + "','" + rsaux1!CODIGO + "','" + rsaux1!DESCRIPCION + "','" + IIf(IsNull(rsaux1!ubicacion), "", rsaux1!ubicacion) + "'," + CStr(rsaux1!Cantidad) + ",'" + Mid(IIf(IsNull(rsaux1!ubicacion), "", rsaux1!ubicacion), 1, 3) + "','" + IIf(IsNull(rsaux1!excedente), "", rsaux1!excedente) + "'," + CStr(var_piezas_totales) + ")", cnn, adOpenDynamic, adLockOptimistic
                     rsaux1.MoveNext
               Wend
               rsaux1.Close
               
               rsaux1.Open "update TB_ORACLE_EMBARQUES_SURTIR set estatus = 'I', fecha_impresion = getdate() where embarque in (" + var_Cadena_embarques + ")", cnn, adOpenDynamic, adLockOptimistic
               For var_j = 1 To Me.lv_embarques.ListItems.Count
                   If Me.lv_embarques.selectedItem.SubItems(4) = "*" Then
                      Me.lv_embarques.selectedItem.SubItems(2) = "I"
                   End If
               Next var_j
               'rsaux1.Open "select distinct pasillo from TB_TEMP_ORACLE_EMBARQUE_concentrado where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " and codigo is not null", cnn, adOpenDynamic, adLockOptimistic
               'While Not rsaux1.EOF
                     Set reporte = appl.OpenReport(App.Path + "\rep_oracle_abastecimiento_ubicaciones_pasillo.rpt")
                     'reporte.RecordSelectionFormula = "{VW_ORACLE_REPORTE_EMBARQUES_CONSENTRADO_PASILLO.INTE_TEM_CONSECUTIVO}=" + CStr(var_consecutivo) + " and {VW_ORACLE_REPORTE_EMBARQUES_CONSENTRADO_PASILLO.PASILLO} = '" + IIf(IsNull(rsaux1!pasillo), "", rsaux1!pasillo) + "'"
                     reporte.RecordSelectionFormula = "{VW_ORACLE_REPORTE_EMBARQUES_CONSENTRADO_PASILLO.INTE_TEM_CONSECUTIVO}=" + CStr(var_consecutivo) + " and {VW_ORACLE_REPORTE_EMBARQUES_CONSENTRADO_PASILLO.codigo} <>''"
                     frmvistasprevias.cr.ReportSource = reporte
                     For ntablas = 1 To reporte.Database.Tables.Count
                         reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                     Next ntablas
                     frmvistasprevias.cr.ViewReport
                     frmvistasprevias.Caption = "Impresion de ubicaciones"
                     frmvistasprevias.Show 1
                     Set reporte = Nothing

               '      rsaux1.MoveNext
               'Wend
               'rsaux1.Close
               
            Else
               MsgBox "No se han asignado pedidos al embarque " + Me.lv_embarques.selectedItem, vbOKOnly, "ATENCION"
            End If
            rs.Close
         Else
            MsgBox "No se seleccionaron pedidos para imprimir, marque los pedidos posicionándose en el registro y oprimiendo la tecla ENTER", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub

Private Sub lv_embarques_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      i = lv_embarques.selectedItem.Index
      If lv_embarques.selectedItem.SubItems(4) = "*" Then
         lv_embarques.selectedItem.SubItems(4) = ""
         lv_embarques.ListItems.Item(i).Bold = False
         lv_embarques.ListItems.Item(i).ForeColor = &H80000012
         lv_embarques.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_embarques.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_embarques.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_embarques.ListItems.Item(i).ListSubItems(4).Bold = False
         lv_embarques.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_embarques.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_embarques.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
         lv_embarques.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
         lv_embarques.Refresh
      Else
         lv_embarques.selectedItem.SubItems(4) = "*"
         lv_embarques.ListItems.Item(i).Bold = True
         lv_embarques.ListItems.Item(i).ForeColor = &HFF0000
         lv_embarques.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_embarques.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_embarques.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_embarques.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_embarques.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_embarques.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_embarques.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_embarques.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
         lv_embarques.Refresh
      End If
      If Me.lv_embarques.ListItems.Count > 0 Then
         Me.lv_embarques.SetFocus
      End If
   End If

End Sub

Private Sub lv_embarques_LostFocus()
   'Me.Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Me.lv_embarques.ListItems.Clear
    If IsDate(Me.txt_inicio) Then
       If IsDate(Me.txt_fin) Then
          var_dia = CStr(Day(CDate(txt_inicio)))
          var_mes = CStr(Month(CDate(txt_inicio)))
          var_año = CStr(Year(CDate(txt_inicio)))
          If Len(Trim(var_dia)) = 1 Then
             var_dia = "0" + var_dia
          End If
          If Len(Trim(var_mes)) = 1 Then
             var_mes = "0" + var_mes
          End If
            
          VAR_FECHA_INICIO_TABLA = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
          var_fecha_fin_1 = CDate(txt_fin) + 1
          var_dia = CStr(Day(var_fecha_fin_1))
          var_mes = CStr(Month(var_fecha_fin_1))
          var_año = CStr(Year(var_fecha_fin_1))
          If Len(Trim(var_dia)) = 1 Then
             var_dia = "0" + var_dia
          End If
          If Len(Trim(var_mes)) = 1 Then
             var_mes = "0" + var_mes
          End If
          VAR_FECHA_FIN_TABLA = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
    
    
          'rs.Open "SELECT * FROM TB_ORACLE_EMBARQUES_SURTIR WHERE ISNULL(ESTATUS,'') = '' and fecha >= " + VAR_FECHA_INICIO_TABLA + " and fecha < " + VAR_FECHA_FIN_TABLA, cnn, adOpenDynamic, adLockOptimistic
          rs.Open "SELECT * FROM TB_ORACLE_EMBARQUES_SURTIR WHERE fecha >= " + VAR_FECHA_INICIO_TABLA + " and fecha < " + VAR_FECHA_FIN_TABLA, cnn, adOpenDynamic, adLockOptimistic
          var_bocina = "I"
          If Not rs.EOF Then
             While Not rs.EOF
                   Set list_item = Me.lv_embarques.ListItems.Add(, , rs!Embarque)
                   list_item.SubItems(1) = rs!Fecha
                   list_item.SubItems(2) = IIf(IsNull(rs!ESTATUS), "", rs!ESTATUS)
                   If IIf(IsNull(rs!ESTATUS), "", rs!ESTATUS) = "" Then
                      var_bocina = ""
                   End If
                   list_item.SubItems(3) = IIf(IsNull(rs!piezas_surtir), "", rs!piezas_surtir)
                   rs.MoveNext
              Wend
              If Me.cmd_silencio.Caption = "Sonido" Then
                 If var_bocina = "" Then
                    Me.wmp1.Controls.play
                 End If
              Else
                 Me.wmp1.Controls.stop
              End If
          Else
             Me.wmp1.Controls.stop
          End If
          rs.Close
       End If
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
