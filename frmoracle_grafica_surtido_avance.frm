VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmoracle_grafica_surtido_avance 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Avance de surtido"
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.MonthView mes 
      Height          =   2370
      Left            =   2040
      TabIndex        =   0
      Top             =   6405
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   59441153
      CurrentDate     =   38148
   End
   Begin MSComctlLib.ListView lv_grafica 
      Height          =   8340
      Left            =   30
      TabIndex        =   11
      Top             =   15
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   14711
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "O.S."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fecha"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Máquina"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Usuario"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Surtir"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Surtido"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Porcentaje"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Cantidad S/C"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Surtida S/C"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Porcentaje S/C"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   720
      Left            =   75
      TabIndex        =   1
      Top             =   8370
      Width           =   5670
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   3690
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   270
         Width           =   1095
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   1110
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   270
         Width           =   1080
      End
      Begin VB.CommandButton Command11 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2220
         Picture         =   "frmoracle_grafica_surtido_avance.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Fecha Inicial"
         Top             =   270
         Width           =   330
      End
      Begin VB.CommandButton Command12 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4830
         Picture         =   "frmoracle_grafica_surtido_avance.frx":1272
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Fecha Final"
         Top             =   270
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   3375
         TabIndex        =   7
         Top             =   330
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   660
         TabIndex        =   6
         Top             =   330
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Reporte "
      Height          =   720
      Left            =   5820
      TabIndex        =   8
      Top             =   8370
      Width           =   9360
      Begin VB.CommandButton Command1 
         Height          =   345
         Left            =   4140
         Picture         =   "frmoracle_grafica_surtido_avance.frx":24E4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Actualiza Grafica"
         Top             =   270
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Height          =   345
         Left            =   4515
         Picture         =   "frmoracle_grafica_surtido_avance.frx":25E6
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Exportar Gráfica"
         Top             =   270
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmoracle_grafica_surtido_avance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_fecha_inicio As String
Dim var_fecha_fin As String
Dim var_tipo_mes As Integer

Private Sub Command1_Click()
   Dim var_consecutivo As Integer
   rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_cadena_agentes = ""
   If IsDate(Me.txt_inicio) Then
      If IsDate(Me.txt_fin) Then
         'var_cadena = "SELECT DISTINCT a.SOURCE_HEADER_NUMBER AS PEDIDO from WSH_DLVB_DLVY_V B, hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, XXVIA_VW_AGENTES D Where a.delivery_id = B.delivery_id AND A.delivery_detail_id = B.delivery_detail_id AND HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID and TRUNC(B.creation_date) >= to_date('" + Me.txt_inicio + "','DD-MM-YYYY') AND TRUNC(B.creation_date) < TO_DATE('" + CStr(CDate(Me.txt_fin) + 1) + "','DD-MM-YYYY') AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID "
         'var_cadena = var_cadena + " AND HCAS.CUST_ACCOUNT_ID = D.CUST_ACCOUNT_ID AND OHA.SHIP_FROM_ORG_ID = " + var_unidad_organizacional + " ORDER BY a.SOURCE_HEADER_NUMBER"
         var_cadena = "select distinct source_header_number  as PEDIDO from xxvia_tb_encabezado_embarques, xxvia_tb_salidas_cajas where fecha_inicio >= to_date('" + Me.txt_inicio + "','DD-MM-YYYY') and fecha_inicio < to_date('" + CStr(CDate(Me.txt_fin) + 1) + "','DD-MM-YYYY') and inte_emb_embarque = embarque"
         rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         var_cadena_pedidos = ""
         While Not rsaux2.EOF
               If Trim(var_cadena_pedidos) = "" Then
                  var_cadena_pedidos = CStr(rsaux2!PEDIDO)
               Else
                  var_cadena_pedidos = var_cadena_pedidos + "," + CStr(rsaux2!PEDIDO)
               End If
               rsaux2.MoveNext
         Wend
         rsaux2.Close
         
         var_cadena = "select distinct source_header_number as PEDIDO from xxvia_tb_encabezado_embarques, xxvia_tb_salidas where fecha_inicio >= to_date('" + Me.txt_inicio + "','DD-MM-YYYY') and fecha_inicio < to_date('" + CStr(CDate(Me.txt_fin) + 1) + "','DD-MM-YYYY') and inte_emb_embarque = embarque"
         rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rsaux2.EOF
               If Trim(var_cadena_pedidos) = "" Then
                  var_cadena_pedidos = CStr(rsaux2!PEDIDO)
               Else
                  var_cadena_pedidos = var_cadena_pedidos + "," + CStr(rsaux2!PEDIDO)
               End If
               rsaux2.MoveNext
         Wend
         rsaux2.Close
         
         
         If var_cadena_pedidos <> "" Then
            cnn.BeginTrans
            rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM tb_Temp_oracle_grafica_surtido", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
            Else
               var_consecutivo = 1
            End If
            rsaux.Close
            rsaux1.Open "insert into tb_Temp_oracle_grafica_surtido (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            
            rsaux2.Open "SELECT pedido, fecha,MAQUINA, USUARIO, cantidad, cantidad_sin_catalogos, SUM(FLOA_SAL_CANTIDAD_LEIDA) as cantidad_surtida FROM XXVIA_TB_SALIDAS_cajas a , XXVIA_TB_ENCABEZADO_EMBARQUES b, XXVIA_tB_ORDENES_GRAFICA c WHERE pedido(+) = source_header_number and INTE_EMB_EMBARQUE = b.EMBARQUE AND pedido IN (" + var_cadena_pedidos + ") GROUP BY pedido, fecha,MAQUINA, USUARIO, cantidad, cantidad_sin_catalogos order by maquina", cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rsaux2.EOF
                  rsaux3.Open "select VCHA_USU_NOMBRE+' '+VCHA_USU_APELLIDOS from tb_usuarios where VCHA_USU_USUARIO_ID = '" + rsaux2!USUARIO + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux3.EOF Then
                     var_nombre_usuario = IIf(IsNull(rsaux3(0).Value), "", rsaux3(0).Value)
                  Else
                     var_nombre_usuario = ""
                  End If
                  rsaux3.Close
                  
                  var_año_s = CStr(Year(rsaux2!Fecha))
                  var_mes_s = CStr(Month(rsaux2!Fecha))
                  var_dia_s = CStr(Day(rsaux2!Fecha))
                  var_hora = CStr(Hour(rsaux2!Fecha))
                  var_minuto = CStr(Minute(rsaux2!Fecha))
                  var_segundo = CStr(Second(rsaux2!Fecha))
                  If Len(var_mes_s) = 1 Then
                     var_mes_s = "0" + var_mes_s
                  End If
                  If Len(var_dia_s) = 1 Then
                     var_dia_s = "0" + var_dia_s
                  End If
                  If Len(var_hora) = 1 Then
                     var_hora = "0" + var_hora
                  End If
                  If Len(var_minuto) = 1 Then
                     var_minuto = "0" + var_minuto
                  End If
                  If Len(var_segundo) = 1 Then
                     var_segundo = "0" + var_segundo
                  End If
                  var_fecha = "{ts '" + var_año_s + "-" + var_mes_s + "-" + var_dia_s + " " + var_hora + ":" + var_minuto + ":" + var_segundo + ".000'}"
                  rsaux3.Open "insert into tb_Temp_oracle_grafica_surtido (inte_tem_consecutivo, pedido, fecha, maquina, usuario, nombre_usuario, cantidad_surtir, cantidad_surtida, CANTIDAD_SURTIR_SIN_CATALOGO) values (" + CStr(var_consecutivo) + "," + CStr(rsaux2!PEDIDO) + "," + var_fecha + ",'" + rsaux2!maquina + "','" + rsaux2!USUARIO + "','" + var_nombre_usuario + "'," + CStr(rsaux2!Cantidad) + "," + CStr(rsaux2!CANTIDAD_SURTIDA) + "," + CStr(IIf(IsNull(rsaux2!CANTIDAD_SIN_CATALOGOS), 0, rsaux2!CANTIDAD_SIN_CATALOGOS)) + ")", cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.MoveNext
            Wend
            rsaux2.Close
            
            rsaux2.Open "SELECT pedido, fecha,MAQUINA, USUARIO, cantidad, cantidad_sin_catalogos, SUM(FLOA_SAL_CANTIDAD_LEIDA) as cantidad_surtida FROM XXVIA_TB_SALIDAS a , XXVIA_TB_ENCABEZADO_EMBARQUES b, XXVIA_tB_ORDENES_GRAFICA c WHERE pedido(+) = source_header_number and INTE_EMB_EMBARQUE = b.EMBARQUE AND pedido IN (" + var_cadena_pedidos + ") GROUP BY pedido, fecha,MAQUINA, USUARIO, cantidad, cantidad_sin_catalogos order by maquina", cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rsaux2.EOF
                  rsaux3.Open "select VCHA_USU_NOMBRE+' '+VCHA_USU_APELLIDOS from tb_usuarios where VCHA_USU_USUARIO_ID = '" + rsaux2!USUARIO + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux3.EOF Then
                     var_nombre_usuario = IIf(IsNull(rsaux3(0).Value), "", rsaux3(0).Value)
                  Else
                     var_nombre_usuario = ""
                  End If
                  rsaux3.Close
                  rsaux3.Open "insert into tb_Temp_oracle_grafica_surtido (inte_tem_consecutivo, pedido, fecha, maquina, usuario, nombre_usuario, cantidad_surtir, cantidad_surtida, CANTIDAD_SURTIR_SIN_CATALOGO) values (" + CStr(var_consecutivo) + "," + CStr(rsaux2!PEDIDO) + ",'" + rsaux2!Fecha + "','" + rsaux2!maquina + "','" + rsaux2!USUARIO + "','" + var_nombre_usuario + "'," + CStr(rsaux2!Cantidad) + "," + CStr(rsaux2!CANTIDAD_SURTIDA) + "," + CStr(IIf(IsNull(rsaux2!CANTIDAD_SIN_CATALOGOS), 0, rsaux2!CANTIDAD_SIN_CATALOGOS)) + ")", cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.MoveNext
            Wend
            rsaux2.Close
            
            var_cadena_pedidos = ""
            rsaux2.Open "SELECT DISTINCT PEDIDO FROM tb_Temp_oracle_grafica_surtido WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO IS NOT NULL", cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux2.EOF
                  If var_cadena_pedidos = "" Then
                     var_cadena_pedidos = CStr(rsaux2!PEDIDO)
                  Else
                     var_cadena_pedidos = var_cadena_pedidos + "," + CStr(rsaux2!PEDIDO)
                  End If
                  rsaux2.MoveNext
            Wend
            rsaux2.Close
            
            rsaux2.Open "SELECT source_header_number as pedido, SUM(FLOA_SAL_CANTIDAD_LEIDA) FROM XXVIA_TB_SALIDAS_CAJAS A, XXVIA_VW_ARTICULOS_CAT B,  XXVIA_TB_ENCABEZADO_EMBARQUES C WHERE ORGANIZACION = ORGANIZATION_ID AND a.inventory_item_id = B.item_id  AND SOURCE_HEADER_NUMBER IN (" + var_cadena_pedidos + ") AND INTE_EMB_EMBARQUE = EMBARQUE AND (LINEA <> 'CATALOGOS' OR LINEA IS NULL) GROUP BY source_header_number", cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rsaux2.EOF
                  rsaux3.Open "UPDATE tb_Temp_oracle_grafica_surtido SET CANTIDAD_SURTIDA_SIN_CATALOGO = " + CStr(IIf(IsNull(rsaux2(1).Value), 0, rsaux2(1).Value)) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO = " + CStr(rsaux2(0).Value), cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.MoveNext
            Wend
            rsaux2.Close
            
            rsaux2.Open "SELECT source_header_number as PEDIDO, SUM(FLOA_SAL_CANTIDAD_LEIDA) FROM XXVIA_TB_SALIDAS A, XXVIA_VW_ARTICULOS_CAT B,  XXVIA_TB_ENCABEZADO_EMBARQUES C WHERE ORGANIZACION = ORGANIZATION_ID AND a.inventory_item_id = B.item_id  AND SOURCE_HEADER_NUMBER IN (" + var_cadena_pedidos + ") AND INTE_EMB_EMBARQUE = EMBARQUE AND (LINEA <> 'CATALOGOS' OR LINEA IS NULL) GROUP BY source_header_number", cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rsaux2.EOF
                  rsaux3.Open "UPDATE tb_Temp_oracle_grafica_surtido SET CANTIDAD_SURTIDA_SIN_CATALOGO = " + CStr(IIf(IsNull(rsaux2(1).Value), 0, rsaux2(1).Value)) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO = " + CStr(rsaux2(0).Value), cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.MoveNext
            Wend
            rsaux2.Close
            
            
            
            Me.lv_grafica.ListItems.Clear
            rs.Open "select * FROM tb_Temp_oracle_grafica_surtido WHERE inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido is not null order by maquina, pedido", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  Set list_item = Me.lv_grafica.ListItems.Add(, , rs!PEDIDO)
                  list_item.SubItems(1) = IIf(IsNull(rs!Fecha), "", rs!Fecha)
                  list_item.SubItems(2) = IIf(IsNull(rs!maquina), "", rs!maquina)
                  list_item.SubItems(3) = IIf(IsNull(rs!NOMBRE_USUARIO), "", rs!NOMBRE_USUARIO)
                  list_item.SubItems(4) = IIf(IsNull(rs!CANTIDAD_SURTIR), "0", rs!CANTIDAD_SURTIR)
                  list_item.SubItems(5) = IIf(IsNull(rs!CANTIDAD_SURTIDA), "0", rs!CANTIDAD_SURTIDA)
                  If IIf(IsNull(rs!CANTIDAD_SURTIDA), 0, rs!CANTIDAD_SURTIDA) = 0 Then
                     var_porcentaje = 0
                  Else
                     var_porcentaje = (IIf(IsNull(rs!CANTIDAD_SURTIDA), 0, rs!CANTIDAD_SURTIDA) * 100) / IIf(IsNull(rs!CANTIDAD_SURTIR), 1, rs!CANTIDAD_SURTIR)
                  End If
                  list_item.SubItems(6) = Format(var_porcentaje, "##0.00")
                  list_item.SubItems(7) = IIf(IsNull(rs!CANTIDAD_SURTIr_SIN_CATALOGO), "0", rs!CANTIDAD_SURTIr_SIN_CATALOGO)
                  list_item.SubItems(8) = IIf(IsNull(rs!CANTIDAD_SURTIDA_SIN_CATALOGO), "0", rs!CANTIDAD_SURTIDA_SIN_CATALOGO)
                  If IIf(IsNull(rs!CANTIDAD_SURTIDA_SIN_CATALOGO), 0, rs!CANTIDAD_SURTIDA_SIN_CATALOGO) = 0 Then
                     If IIf(IsNull(rs!CANTIDAD_SURTIr_SIN_CATALOGO), 0, rs!CANTIDAD_SURTIr_SIN_CATALOGO) = 0 Then
                        var_porcentaje = 100
                     Else
                        var_porcentaje = 0
                     End If
                     
                  Else
                     If IIf(IsNull(rs!CANTIDAD_SURTIr_SIN_CATALOGO), 0, rs!CANTIDAD_SURTIr_SIN_CATALOGO) = 0 Then
                        var_porcentaje = 100
                     Else
                        var_porcentaje = (IIf(IsNull(rs!CANTIDAD_SURTIDA_SIN_CATALOGO), 0, rs!CANTIDAD_SURTIDA_SIN_CATALOGO) * 100) / IIf(IsNull(rs!CANTIDAD_SURTIr_SIN_CATALOGO), 1, rs!CANTIDAD_SURTIr_SIN_CATALOGO)
                     End If
                  End If
                  list_item.SubItems(9) = Format(var_porcentaje, "##0.00")
                  
                  
                  rs.MoveNext
            Wend
            rs.Close
         
         
            For var_i = 1 To lv_grafica.ListItems.Count
                lv_grafica.ListItems(var_i).Selected = True
                If (lv_grafica.selectedItem.SubItems(6) * 1) > 25 Then
                   lv_grafica.ListItems.Item(var_i).ForeColor = vbBlue
                   'lv_grafica.ListItems(var_i).ListSubItems(1).Bold = True
                   lv_grafica.ListItems(var_i).ListSubItems(1).ForeColor = vbBlue
                   lv_grafica.ListItems(var_i).ListSubItems(2).Bold = True
                   lv_grafica.ListItems(var_i).ListSubItems(2).ForeColor = vbBlue
                   lv_grafica.ListItems(var_i).ListSubItems(3).Bold = True
                   lv_grafica.ListItems(var_i).ListSubItems(3).ForeColor = vbBlue
                   lv_grafica.ListItems(var_i).ListSubItems(4).Bold = True
                   lv_grafica.ListItems(var_i).ListSubItems(4).ForeColor = vbBlue
                   lv_grafica.ListItems(var_i).ListSubItems(5).Bold = True
                   lv_grafica.ListItems(var_i).ListSubItems(5).ForeColor = vbBlue
                   lv_grafica.ListItems(var_i).ListSubItems(6).Bold = True
                   lv_grafica.ListItems(var_i).ListSubItems(6).ForeColor = vbBlue
                   lv_grafica.ListItems(var_i).ListSubItems(7).Bold = True
                   lv_grafica.ListItems(var_i).ListSubItems(7).ForeColor = vbBlue
                   lv_grafica.ListItems(var_i).ListSubItems(8).Bold = True
                   lv_grafica.ListItems(var_i).ListSubItems(8).ForeColor = vbBlue
                   lv_grafica.ListItems(var_i).ListSubItems(9).Bold = True
                   lv_grafica.ListItems(var_i).ListSubItems(9).ForeColor = vbBlue
                   lv_grafica.selectedItem.Bold = True
                End If
                If (lv_grafica.selectedItem.SubItems(6) * 1) > 50 Then
                   lv_grafica.ListItems.Item(var_i).ForeColor = &HC000C0
                   'lv_grafica.ListItems(var_i).ListSubItems(1).Bold = True
                   lv_grafica.ListItems(var_i).ListSubItems(1).ForeColor = &HC000C0
                   lv_grafica.ListItems(var_i).ListSubItems(2).Bold = True
                   lv_grafica.ListItems(var_i).ListSubItems(2).ForeColor = &HC000C0
                   lv_grafica.ListItems(var_i).ListSubItems(3).Bold = True
                   lv_grafica.ListItems(var_i).ListSubItems(3).ForeColor = &HC000C0
                   lv_grafica.ListItems(var_i).ListSubItems(4).Bold = True
                   lv_grafica.ListItems(var_i).ListSubItems(4).ForeColor = &HC000C0
                   lv_grafica.ListItems(var_i).ListSubItems(5).Bold = True
                   lv_grafica.ListItems(var_i).ListSubItems(5).ForeColor = &HC000C0
                   lv_grafica.ListItems(var_i).ListSubItems(6).Bold = True
                   lv_grafica.ListItems(var_i).ListSubItems(6).ForeColor = &HC000C0
                   lv_grafica.ListItems(var_i).ListSubItems(7).Bold = True
                   lv_grafica.ListItems(var_i).ListSubItems(7).ForeColor = &HC000C0
                   lv_grafica.ListItems(var_i).ListSubItems(8).Bold = True
                   lv_grafica.ListItems(var_i).ListSubItems(8).ForeColor = &HC000C0
                   lv_grafica.ListItems(var_i).ListSubItems(9).Bold = True
                   lv_grafica.ListItems(var_i).ListSubItems(9).ForeColor = &HC000C0
                   lv_grafica.selectedItem.Bold = True
                End If
                If (lv_grafica.selectedItem.SubItems(6) * 1) = 100 Then
                   lv_grafica.ListItems.Item(var_i).ForeColor = vbRed
                   'lv_grafica.ListItems(var_i).ListSubItems(1).Bold = True
                   lv_grafica.ListItems(var_i).ListSubItems(1).ForeColor = vbRed
                   lv_grafica.ListItems(var_i).ListSubItems(2).Bold = True
                   lv_grafica.ListItems(var_i).ListSubItems(2).ForeColor = vbRed
                   lv_grafica.ListItems(var_i).ListSubItems(3).Bold = True
                   lv_grafica.ListItems(var_i).ListSubItems(3).ForeColor = vbRed
                   lv_grafica.ListItems(var_i).ListSubItems(4).Bold = True
                   lv_grafica.ListItems(var_i).ListSubItems(4).ForeColor = vbRed
                   lv_grafica.ListItems(var_i).ListSubItems(5).Bold = True
                   lv_grafica.ListItems(var_i).ListSubItems(5).ForeColor = vbRed
                   lv_grafica.ListItems(var_i).ListSubItems(6).Bold = True
                   lv_grafica.ListItems(var_i).ListSubItems(6).ForeColor = vbRed
                   lv_grafica.ListItems(var_i).ListSubItems(7).Bold = True
                   lv_grafica.ListItems(var_i).ListSubItems(7).ForeColor = vbRed
                   lv_grafica.ListItems(var_i).ListSubItems(8).Bold = True
                   lv_grafica.ListItems(var_i).ListSubItems(8).ForeColor = vbRed
                   lv_grafica.ListItems(var_i).ListSubItems(9).Bold = True
                   lv_grafica.ListItems(var_i).ListSubItems(9).ForeColor = vbRed
                   lv_grafica.selectedItem.Bold = True
               End If
            Next var_i
            rsaux3.Open "delete from tb_Temp_oracle_grafica_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "No existen pedidos para la fecha seleccionada", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub

Private Sub Command11_Click()
   If IsDate(Me.txt_inicio) Then
      Me.mes.Value = CDate(Me.txt_inicio)
   Else
      mes.Value = Date
   End If
   var_tipo_mes = 1
   mes.Visible = True
   mes.SetFocus
End Sub

Private Sub Command12_Click()
   If IsDate(Me.txt_fin) Then
      mes.Value = CDate(Me.txt_fin)
   Else
      mes.Value = Date
   End If
   var_tipo_mes = 2
   mes.Visible = True
   mes.SetFocus
End Sub

Private Sub Command2_Click()
   Dim var_consecutivo As Integer
   rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   var_cadena_agentes = ""
   If IsDate(Me.txt_inicio) Then
      If IsDate(Me.txt_fin) Then
         'var_cadena = "SELECT DISTINCT a.SOURCE_HEADER_NUMBER AS PEDIDO from WSH_DLVB_DLVY_V B, hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, XXVIA_VW_AGENTES D Where a.delivery_id = B.delivery_id AND A.delivery_detail_id = B.delivery_detail_id AND HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID and TRUNC(B.creation_date) >= to_date('" + Me.txt_inicio + "','DD-MM-YYYY') AND TRUNC(B.creation_date) < TO_DATE('" + CStr(CDate(Me.txt_fin) + 1) + "','DD-MM-YYYY') AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.COLLECTOR_ID IN (" + var_cadena_agentes + ") "
         var_cadena = "SELECT DISTINCT a.SOURCE_HEADER_NUMBER AS PEDIDO from WSH_DLVB_DLVY_V B, hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, XXVIA_VW_AGENTES D Where a.delivery_id = B.delivery_id AND A.delivery_detail_id = B.delivery_detail_id AND HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID and TRUNC(B.creation_date) >= to_date('" + Me.txt_inicio + "','DD-MM-YYYY') AND TRUNC(B.creation_date) < TO_DATE('" + CStr(CDate(Me.txt_fin) + 1) + "','DD-MM-YYYY') AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID "
         var_cadena = var_cadena + " AND HCAS.CUST_ACCOUNT_ID = D.CUST_ACCOUNT_ID AND OHA.SHIP_FROM_ORG_ID = " + var_unidad_organizacional + " ORDER BY a.SOURCE_HEADER_NUMBER"
         Text1 = var_cadena
         rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         var_cadena_pedidos = ""
         While Not rsaux2.EOF
               If Trim(var_cadena_pedidos) = "" Then
                  var_cadena_pedidos = CStr(rsaux2!PEDIDO)
               Else
                  var_cadena_pedidos = var_cadena_pedidos + "," + CStr(rsaux2!PEDIDO)
               End If
               rsaux2.MoveNext
         Wend
         rsaux2.Close
         If var_cadena_pedidos <> "" Then
            cnn.BeginTrans
            rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM tb_Temp_oracle_grafica_surtido", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value) + 1
            Else
               var_consecutivo = 1
            End If
            rsaux.Close
            rsaux1.Open "insert into tb_Temp_oracle_grafica_surtido (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            
            rsaux2.Open "SELECT pedido, fecha,MAQUINA, USUARIO, cantidad, CANTIDAD_SIN_CATALOGOS, SUM(FLOA_SAL_CANTIDAD_LEIDA) as cantidad_surtida FROM XXVIA_TB_SALIDAS_cajas a , XXVIA_TB_ENCABEZADO_EMBARQUES b, XXVIA_tB_ORDENES_GRAFICA c WHERE pedido(+) = source_header_number and INTE_EMB_EMBARQUE = b.EMBARQUE AND pedido IN (" + var_cadena_pedidos + ") GROUP BY pedido, fecha,MAQUINA, USUARIO, cantidad, CANTIDAD_SIN_CATALOGOS order by maquina", cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rsaux2.EOF
                  rsaux3.Open "select VCHA_USU_NOMBRE+' '+VCHA_USU_APELLIDOS from tb_usuarios where VCHA_USU_USUARIO_ID = '" + rsaux2!USUARIO + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux3.EOF Then
                     var_nombre_usuario = IIf(IsNull(rsaux3(0).Value), "", rsaux3(0).Value)
                  Else
                     var_nombre_usuario = ""
                  End If
                  rsaux3.Close
                  rsaux3.Open "insert into tb_Temp_oracle_grafica_surtido (inte_tem_consecutivo, pedido, fecha, maquina, usuario, nombre_usuario, cantidad_surtir, cantidad_surtida, CANTIDAD_SURTIR_SIN_CATALOGO) values (" + CStr(var_consecutivo) + "," + CStr(rsaux2!PEDIDO) + ",'" + rsaux2!Fecha + "','" + rsaux2!maquina + "','" + rsaux2!USUARIO + "','" + var_nombre_usuario + "'," + CStr(rsaux2!Cantidad) + "," + CStr(rsaux2!CANTIDAD_SURTIDA) + "," + CStr(IIf(IsNull(rsaux2!CANTIDAD_SIN_CATALOGOS), 0, rsaux2!CANTIDAD_SIN_CATALOGOS)) + ")", cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.MoveNext
            Wend
            rsaux2.Close
            
            rsaux2.Open "SELECT pedido, fecha,MAQUINA, USUARIO, cantidad, SUM(FLOA_SAL_CANTIDAD_LEIDA) as cantidad_surtida FROM XXVIA_TB_SALIDAS a , XXVIA_TB_ENCABEZADO_EMBARQUES b, XXVIA_tB_ORDENES_GRAFICA c WHERE pedido(+) = source_header_number and INTE_EMB_EMBARQUE = b.EMBARQUE AND pedido IN (" + var_cadena_pedidos + ") GROUP BY pedido, fecha,MAQUINA, USUARIO, cantidad order by maquina", cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rsaux2.EOF
                  rsaux3.Open "select VCHA_USU_NOMBRE+' '+VCHA_USU_APELLIDOS from tb_usuarios where VCHA_USU_USUARIO_ID = '" + rsaux2!USUARIO + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux3.EOF Then
                     var_nombre_usuario = IIf(IsNull(rsaux3(0).Value), "", rsaux3(0).Value)
                  Else
                     var_nombre_usuario = ""
                  End If
                  rsaux3.Close
                  rsaux3.Open "insert into tb_Temp_oracle_grafica_surtido (inte_tem_consecutivo, pedido, fecha, maquina, usuario, nombre_usuario, cantidad_surtir, cantidad_surtida) values (" + CStr(var_consecutivo) + "," + CStr(rsaux2!PEDIDO) + ",'" + rsaux2!Fecha + "','" + rsaux2!maquina + "','" + rsaux2!USUARIO + "','" + var_nombre_usuario + "'," + CStr(rsaux2!Cantidad) + "," + CStr(rsaux2!CANTIDAD_SURTIDA) + ")", cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.MoveNext
            Wend
            rsaux2.Close
            rs.Open "delete FROM tb_Temp_oracle_grafica_surtido WHERE inte_tem_consecutivo = " + CStr(var_consecutivo) + " and pedido is null", cnn, adOpenDynamic, adLockOptimistic
            
            
            
            
            var_cadena_pedidos = ""
            rsaux2.Open "SELECT DISTINCT PEDIDO FROM tb_Temp_oracle_grafica_surtido WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO IS NOT NULL", cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux2.EOF
                  If var_cadena_pedidos = "" Then
                     var_cadena_pedidos = CStr(rsaux2!PEDIDO)
                  Else
                     var_cadena_pedidos = var_cadena_pedidos + "," + CStr(rsaux2!PEDIDO)
                  End If
                  rsaux2.MoveNext
            Wend
            rsaux2.Close
            
            rsaux2.Open "SELECT source_header_number as pedido, SUM(FLOA_SAL_CANTIDAD_LEIDA) FROM XXVIA_TB_SALIDAS_CAJAS A, XXVIA_VW_ARTICULOS_CAT B,  XXVIA_TB_ENCABEZADO_EMBARQUES C WHERE ORGANIZACION = ORGANIZATION_ID AND a.inventory_item_id = B.item_id  AND SOURCE_HEADER_NUMBER IN (" + var_cadena_pedidos + ") AND INTE_EMB_EMBARQUE = EMBARQUE AND (LINEA <> 'CATALOGOS' OR LINEA IS NULL) GROUP BY source_header_number", cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rsaux2.EOF
                  rsaux3.Open "UPDATE tb_Temp_oracle_grafica_surtido SET CANTIDAD_SURTIDA_SIN_CATALOGO = " + CStr(IIf(IsNull(rsaux2(1).Value), 0, rsaux2(1).Value)) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO = " + CStr(rsaux2(0).Value), cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.MoveNext
            Wend
            rsaux2.Close
            
            rsaux2.Open "SELECT source_header_number as PEDIDO, SUM(FLOA_SAL_CANTIDAD_LEIDA) FROM XXVIA_TB_SALIDAS A, XXVIA_VW_ARTICULOS_CAT B,  XXVIA_TB_ENCABEZADO_EMBARQUES C WHERE ORGANIZACION = ORGANIZATION_ID AND a.inventory_item_id = B.item_id  AND SOURCE_HEADER_NUMBER IN (" + var_cadena_pedidos + ") AND INTE_EMB_EMBARQUE = EMBARQUE AND (LINEA <> 'CATALOGOS' OR LINEA IS NULL) GROUP BY source_header_number", cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rsaux2.EOF
                  rsaux3.Open "UPDATE tb_Temp_oracle_grafica_surtido SET CANTIDAD_SURTIDA_SIN_CATALOGO = " + CStr(IIf(IsNull(rsaux2(1).Value), 0, rsaux2(1).Value)) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO = " + CStr(rsaux2(0).Value), cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.MoveNext
            Wend
            rsaux2.Close
            
            
            
            
            Set reporte = appl.OpenReport(App.Path + "\rep_oracle_grafica_surtido.rpt")
            reporte.RecordSelectionFormula = "{VW_ORACLE_GRAFICA_SURTIDO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\reporte_grafica_surtido_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
            Set reporte = Nothing
            MsgBox "Se a terminado de guardar el archivo " + archivo
            
                        
            
            rsaux3.Open "delete from tb_Temp_oracle_grafica_surtido where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "No existen pedidos para la fecha seleccionada", vbOKOnly, "ATENCION"
         End If
      End If
   End If
End Sub

Private Sub Form_Load()
   Me.mes.Visible = False
   Me.txt_fin = Date
   Me.txt_inicio = Date
   Top = 0
   Left = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_grafica_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_grafica, ColumnHeader)
End Sub

Private Sub mes_DateDblClick(ByVal DateDblClicked As Date)
   If var_tipo_mes = 1 Then
      txt_inicio = mes.Value
      
   End If
   If var_tipo_mes = 2 Then
      txt_fin = mes.Value
   End If
   mes.Visible = False
End Sub

Private Sub mes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      mes.Visible = False
   End If
End Sub

