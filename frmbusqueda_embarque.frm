VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmbusqueda_embarque 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Busqueda de Embarques"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_carga_mmasiva 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   370
      Picture         =   "frmbusqueda_embarque.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Carga masiva"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   60
      TabIndex        =   8
      Top             =   360
      Width           =   11550
   End
   Begin VB.CommandButton cmd_aceptar 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   60
      Picture         =   "frmbusqueda_embarque.frx":0312
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Caption         =   " Embarques "
      Height          =   4800
      Left            =   75
      TabIndex        =   3
      Top             =   405
      Width           =   11535
      Begin MSComCtl2.MonthView mes 
         Height          =   2370
         Left            =   3885
         TabIndex        =   5
         Top             =   960
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   459800577
         CurrentDate     =   38397
      End
      Begin VB.TextBox txt_fin 
         Height          =   360
         Left            =   6585
         TabIndex        =   1
         Top             =   180
         Width           =   1350
      End
      Begin VB.TextBox txt_inicio 
         Height          =   360
         Left            =   3945
         TabIndex        =   0
         Top             =   180
         Width           =   1350
      End
      Begin MSComctlLib.ListView lv_embarques 
         Height          =   4140
         Left            =   105
         TabIndex        =   4
         Top             =   630
         Width           =   11340
         _ExtentX        =   20003
         _ExtentY        =   7303
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Embarque"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Ruta"
            Object.Width           =   4233
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Pedido"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cliente"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Máquina"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Fecha"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "estatus"
            Object.Width           =   1235
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Fin:"
         Height          =   195
         Left            =   5670
         TabIndex        =   7
         Top             =   270
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio:"
         Height          =   195
         Left            =   2985
         TabIndex        =   6
         Top             =   270
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmbusqueda_embarque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_mes As Integer

Private Sub cmd_aceptar_Click()
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         var_fecha_fin_1 = CDate(txt_fin) + 1
         var_dia = CStr(Day(CDate(txt_inicio)))
         var_mes2 = CStr(Month(CDate(txt_inicio)))
         var_año = CStr(Year(CDate(txt_inicio)))
         If Len(Trim(var_dia)) = 1 Then
            var_dia = "0" + var_dia
         End If
         If Len(Trim(var_mes2)) = 1 Then
            var_mes2 = "0" + var_mes2
         End If
         var_fecha_inicio = var_dia + "-" + var_mes2 + "-" + var_año
              
              
         var_dia = CStr(Day(var_fecha_fin_1))
         var_mes2 = CStr(Month(var_fecha_fin_1))
         var_año = CStr(Year(var_fecha_fin_1))
         If Len(Trim(var_dia)) = 1 Then
            var_dia = "0" + var_dia
         End If
         If Len(Trim(var_mes2)) = 1 Then
            var_mes2 = "0" + var_mes2
         End If
         var_fecha_fin = var_dia + "-" + var_mes2 + "-" + var_año
         
         var_cadena = "SELECT DISTINCT source_document_id, XE.JAULA, XE.EMBARQUE, xeo.source_header_number, OH.ORG_ID, J.NAME, OH.salesrep_id, HL.ADDRESS1 AS CUSTOMER_NAME, XE.char_emb_estatus, XE.FECHA_INICIO, XE.FECHA_FIN, xe.maquina"
         var_cadena = var_cadena + " FROM oe_order_headers_all oh, hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, HZ_CUST_SITE_USES_ALL HCSU, hz_cust_acct_sites_all f, jtf_rs_salesreps J, XXVIA_VW_EMBARQUES_ORDENES XEO, xxvia_tb_encabezado_embarques xe Where order_number = XEO.SOURCE_HEADER_NUMBER AND HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID = HL.LOCATION_ID AND HCSU.SITE_USE_ID = OH.INVOICE_TO_ORG_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND f.cust_acct_site_id = HCAS.CUST_ACCT_SITE_ID AND J.salesrep_id = OH.salesrep_id AND XE.embarque = xeo.inte_emb_embarque AND xe.fecha_inicio >= TO_dATE('" + var_fecha_inicio + "','DD-MM-YYYY') AND XE.fecha_inicio < TO_DATE('" + var_fecha_fin + "','DD-MM-YYYY') AND OH.ORG_ID = " + var_empresa + " ORDER BY xe.embarque, xeo.source_header_number"
         rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         'MsgBox var_cadena
         rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         lv_embarques.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_embarques.ListItems.Add(, , rs!Embarque)
               If rs!salesrep_id = -3 Then
                  VAR_AGENTE = "TIENDAS"
               Else
                  VAR_AGENTE = rs!Name
               End If
               VAR_NOMBRE_PROVEEDOR = ""
               If VAR_AGENTE = "TIENDAS" Then
                  rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(IIf(IsNull(rs!source_document_id), 0, rs!source_document_id)) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux2.EOF Then
                     VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                  End If
                  rsaux2.Close
               Else
                  VAR_NOMBRE_PROVEEDOR = IIf(IsNull(rs!customer_name), "", rs!customer_name)
               End If
               list_item.SubItems(1) = IIf(IsNull(VAR_AGENTE), "", VAR_AGENTE)
               list_item.SubItems(2) = rs!source_header_number
               list_item.SubItems(3) = VAR_NOMBRE_PROVEEDOR
               list_item.SubItems(4) = IIf(IsNull(rs!maquina), 0, rs!maquina)
               list_item.SubItems(5) = IIf(IsNull(rs!FECHA_INICIO), "", Trim(rs!FECHA_INICIO))
               list_item.SubItems(6) = IIf(IsNull(rs!char_Emb_estatus), "", Trim(rs!char_Emb_estatus))
               rs.MoveNext
         Wend
         rs.Close
         If lv_embarques.ListItems.Count > 17 Then
            lv_embarques.ColumnHeaders(2).Width = 2300.15
         Else
            lv_embarques.ColumnHeaders(2).Width = 2500.15
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de inicio incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      If var_mes > 0 Then
         If var_mes = 1 Then
            txt_inicio.SetFocus
         End If
         If var_mes = 2 Then
            txt_fin.SetFocus
         End If
      Else
         Unload Me
      End If
   End If
End Sub

Private Sub Form_Load()
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   mes.Visible = False
   Me.txt_fin = Date
   Me.txt_inicio = Date
   
   var_fecha_fin_1 = CDate(txt_fin) + 1
   var_dia = CStr(Day(CDate(txt_inicio)))
   var_mes2 = CStr(Month(CDate(txt_inicio)))
   var_año = CStr(Year(CDate(txt_inicio)))
   If Len(Trim(var_dia)) = 1 Then
      var_dia = "0" + var_dia
   End If
   If Len(Trim(var_mes2)) = 1 Then
      var_mes2 = "0" + var_mes2
   End If
   var_fecha_inicio = "{d '" + var_año + "-" + var_mes2 + "-" + var_dia + "'}"
              
              
   var_dia = CStr(Day(var_fecha_fin_1))
   var_mes2 = CStr(Month(var_fecha_fin_1))
   var_año = CStr(Year(var_fecha_fin_1))
   If Len(Trim(var_dia)) = 1 Then
      var_dia = "0" + var_dia
   End If
   If Len(Trim(var_mes2)) = 1 Then
      var_mes2 = "0" + var_mes2
   End If
   var_fecha_fin = "{d '" + var_año + "-" + var_mes2 + "-" + var_dia + "'}"
   Cadena = "SELECT dbo.TB_ENCABEZADO_EMBARQUES.VCHA_EMP_EMPRESA_ID, dbo.TB_ENCABEZADO_EMBARQUES.VCHA_UOR_UNIDAD_ID, dbo.TB_ENCABEZADO_EMBARQUES.INTE_EMB_EMBARQUE, dbo.TB_ENCABEZADO_EMBARQUES.INTE_JAU_JAULA_ID, dbo.TB_ENCABEZADO_EMBARQUES.VCHA_AGE_AGENTE_ID, dbo.TB_AGENTES.VCHA_AGE_NOMBRE, dbo.TB_ENCABEZADO_EMBARQUES.DTIM_EMB_FECHA_INICIO, dbo.TB_ENCABEZADO_EMBARQUES.CHAR_EMB_ESTATUS, dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID, dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO, dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO FROM dbo.TB_ENCABEZADO_MOVIMIENTOS INNER JOIN "
   Cadena = Cadena + " dbo.TB_DETALLE_EMBARQUES ON dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_EMP_EMPRESA_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_UOR_UNIDAD_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_ALM_ALMACEN_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_ALM_ALMACEN_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO = dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO INNER JOIN dbo.TB_ENC_ORDEN_SURTIDO ON dbo.TB_ENCABEZADO_MOVIMIENTOS.INTE_EMO_NUMERO_ORIGEN = dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO RIGHT OUTER JOIN"
   Cadena = Cadena + " dbo.TB_ENCABEZADO_EMBARQUES INNER JOIN dbo.TB_AGENTES ON dbo.TB_ENCABEZADO_EMBARQUES.VCHA_AGE_AGENTE_ID = dbo.TB_AGENTES.VCHA_AGE_AGENTE_ID ON dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_EMBARQUES.VCHA_EMP_EMPRESA_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_EMBARQUES.VCHA_UOR_UNIDAD_ID AND dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE = dbo.TB_ENCABEZADO_EMBARQUES.INTE_EMB_EMBARQUE   where DTIM_EMB_FECHA_INICIO >= " + var_fecha_inicio + " and DTIM_EMB_FECHA_INICIO <= " + var_fecha_fin
   Text1 = Cadena
   rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set list_item = lv_embarques.ListItems.Add(, , rs!inte_emb_embarque)
         list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", Trim(rs!VCHA_AGE_NOMBRE))
         list_item.SubItems(2) = IIf(IsNull(rs!inte_ors_orden_surtido), 0, Trim(rs!inte_ors_orden_surtido))
         list_item.SubItems(3) = IIf(IsNull(rs!VCHA_MOV_MOVIMIENTO_ID), "", Trim(rs!VCHA_MOV_MOVIMIENTO_ID))
         list_item.SubItems(4) = IIf(IsNull(rs!INTE_EMO_NUMERO), "", Trim(rs!INTE_EMO_NUMERO))
         list_item.SubItems(5) = IIf(IsNull(rs!inte_jau_jaula_id), "", Trim(rs!inte_jau_jaula_id))
         list_item.SubItems(6) = IIf(IsNull(rs!dtim_emb_fecha_inicio), "", Trim(rs!dtim_emb_fecha_inicio))
         list_item.SubItems(7) = IIf(IsNull(rs!char_Emb_estatus), "", Trim(rs!char_Emb_estatus))
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub lv_embarques_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_embarques, ColumnHeader)
End Sub

Private Sub lv_embarques_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_embarques.ListItems.Count > 0 Then
         frmnumero_embarque.txt_embarque = lv_embarques.selectedItem
         Unload Me
      End If
   End If
End Sub

Private Sub mes_DateDblClick(ByVal DateDblClicked As Date)
   If var_mes = 1 Then
      Me.txt_inicio = mes.Value
   End If
   If var_mes = 2 Then
      Me.txt_fin = mes.Value
   End If
   Me.mes.Visible = False
End Sub

Private Sub mes_LostFocus()
   mes.Visible = False
End Sub

Private Sub txt_fin_GotFocus()
   var_mes = 2
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_mes = 2
      If IsDate(txt_fin) Then
         mes.Value = CDate(txt_fin)
      Else
         mes.Value = Date
      End If
      mes.Visible = True
      mes.SetFocus
   End If
End Sub

Private Sub txt_fin_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Me.cmd_aceptar.SetFocus
End Sub

Private Sub txt_inicio_GotFocus()
   var_mes = 1
End Sub

Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_mes = 1
      If IsDate(txt_inicio) Then
         mes.Value = CDate(txt_inicio)
      Else
         mes.Value = Date
      End If
      mes.Visible = True
      mes.SetFocus
   End If
End Sub

Private Sub txt_inicio_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

