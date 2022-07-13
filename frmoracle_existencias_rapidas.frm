VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_existencias_rapidas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Existencias rapidas"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_nombre_almacen 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2055
      TabIndex        =   12
      Top             =   345
      Width           =   9300
   End
   Begin VB.TextBox txt_cantidad_ordenes 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   390
      Left            =   9225
      TabIndex        =   11
      Top             =   6705
      Width           =   2265
   End
   Begin VB.CommandButton cmd_imprimir 
      Height          =   330
      Left            =   105
      Picture         =   "frmoracle_existencias_rapidas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6750
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Frame frm_ubicaciones 
      Height          =   4785
      Left            =   2205
      TabIndex        =   0
      Top             =   1695
      Width           =   6735
      Begin MSComctlLib.ListView lv_ubicaciones 
         Height          =   4125
         Left            =   60
         TabIndex        =   1
         Top             =   600
         Width           =   6570
         _ExtentX        =   11589
         _ExtentY        =   7276
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
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label lbl_ubicacion 
         BackColor       =   &H000000C0&
         ForeColor       =   &H8000000E&
         Height          =   315
         Left            =   30
         TabIndex        =   2
         Top             =   120
         Width           =   6660
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " Almacen "
      Height          =   690
      Left            =   45
      TabIndex        =   13
      Top             =   90
      Width           =   11475
      Begin VB.TextBox txt_clave_almacen 
         Height          =   315
         Left            =   105
         TabIndex        =   14
         Top             =   255
         Width           =   1890
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Artículo "
      Height          =   735
      Left            =   45
      TabIndex        =   24
      Top             =   825
      Width           =   11475
      Begin VB.TextBox txt_descripcion 
         Height          =   330
         Left            =   1995
         TabIndex        =   26
         Top             =   270
         Width           =   9300
      End
      Begin VB.TextBox txt_codigo 
         Height          =   330
         Left            =   105
         TabIndex        =   25
         Top             =   270
         Width           =   1875
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Existencias "
      Height          =   750
      Left            =   30
      TabIndex        =   17
      Top             =   1575
      Width           =   11505
      Begin VB.TextBox txt_fisico 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   330
         Left            =   1035
         TabIndex        =   20
         Top             =   270
         Width           =   1170
      End
      Begin VB.TextBox txt_apartado 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   330
         Left            =   4875
         TabIndex        =   19
         Top             =   270
         Width           =   1170
      End
      Begin VB.TextBox txt_disponible 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   330
         Left            =   8730
         TabIndex        =   18
         Top             =   285
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fisíco:"
         Height          =   195
         Left            =   540
         TabIndex        =   23
         Top             =   345
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Apartados:"
         Height          =   195
         Left            =   4065
         TabIndex        =   22
         Top             =   345
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Disponibles:"
         Height          =   195
         Left            =   7845
         TabIndex        =   21
         Top             =   345
         Width           =   855
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   " Ubicaciones "
      Height          =   1110
      Left            =   30
      TabIndex        =   4
      Top             =   2355
      Width           =   11505
      Begin VB.TextBox txt_ubicacion_6 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   345
         Left            =   8775
         TabIndex        =   30
         Top             =   690
         Width           =   1935
      End
      Begin VB.TextBox txt_ubicacion_5 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   345
         Left            =   4905
         TabIndex        =   29
         Top             =   675
         Width           =   1935
      End
      Begin VB.TextBox txt_ubicacion_4 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   345
         Left            =   1050
         TabIndex        =   28
         Top             =   675
         Width           =   1935
      End
      Begin VB.TextBox txt_ubicacion_3 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   345
         Left            =   8775
         TabIndex        =   7
         Top             =   285
         Width           =   1935
      End
      Begin VB.TextBox txt_ubicacion_2 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   345
         Left            =   4905
         TabIndex        =   6
         Top             =   270
         Width           =   1935
      End
      Begin VB.TextBox txt_ubicacion_1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   345
         Left            =   1050
         TabIndex        =   5
         Top             =   270
         Width           =   1935
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Ubicación 6:"
         Height          =   195
         Left            =   7815
         TabIndex        =   33
         Top             =   750
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Ubicación 5:"
         Height          =   195
         Left            =   3975
         TabIndex        =   32
         Top             =   750
         Width           =   900
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Ubiación 4:"
         Height          =   195
         Left            =   165
         TabIndex        =   31
         Top             =   750
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ubicación 3:"
         Height          =   195
         Left            =   7815
         TabIndex        =   10
         Top             =   345
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Ubicación 2:"
         Height          =   195
         Left            =   3975
         TabIndex        =   9
         Top             =   345
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Ubiación 1:"
         Height          =   195
         Left            =   165
         TabIndex        =   8
         Top             =   345
         Width           =   810
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Desgloce "
      Height          =   3165
      Left            =   45
      TabIndex        =   15
      Top             =   3480
      Width           =   11490
      Begin MSComctlLib.ListView lv_desgloce 
         Height          =   2880
         Left            =   60
         TabIndex        =   16
         Top             =   225
         Width           =   11370
         _ExtentX        =   20055
         _ExtentY        =   5080
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
            Text            =   "Pedido"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Agente"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cliente"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Cantidad"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Total en ordenes de surtido:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6705
      TabIndex        =   27
      Top             =   6795
      Width           =   2415
   End
End
Attribute VB_Name = "frmoracle_existencias_rapidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter


Private Sub Form_Load()
   Top = 0
   Left = 0
   Me.frm_ubicaciones.Visible = False
   If var_unidad_organizacional = 93 Then
      Me.txt_clave_almacen = "CDI_ALMPT"
      Me.txt_nombre_almacen = "PT. ALMACEN GENERAL"
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_desgloce_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_desgloce, ColumnHeader)
End Sub

Private Sub txt_clave_almacen_Change()
   Me.txt_nombre_almacen = ""
End Sub

Private Sub txt_clave_almacen_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Me.txt_codigo.SetFocus
   End If
End Sub

Private Sub txt_clave_almacen_LostFocus()
   If Trim(txt_clave_almacen) <> "" Then
      rsaux.Open "select secondary_inventory_name as vcha_alm_almacen_id, description as vcha_alm_nombre from mtl_secondary_inventories where organization_id = " + var_unidad_organizacional + " and secondary_inventory_name = '" + txt_clave_almacen + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rsaux.EOF Then
         txt_nombre_almacen = rsaux!vcha_alm_nombre
         var_almacen_Destino = txt_almacen
      Else
         Me.txt_clave_almacen = ""
         Me.txt_nombre_almacen = ""
         MsgBox "El almacén no existe", vbOKOnly, "ATENCION"
      End If
      rsaux.Close
   Else
      MsgBox "Clave de almacen incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub txt_codigo_Change()
      Me.lv_desgloce.ListItems.Clear
      Me.txt_descripcion = ""
      Me.txt_fisico = ""
      Me.txt_apartado = ""
      Me.txt_disponible = ""
      Me.txt_cantidad_ordenes = ""
      Me.txt_ubicacion_1 = ""
      Me.txt_ubicacion_2 = ""
      Me.txt_ubicacion_3 = ""
      Me.txt_ubicacion_4 = ""
      Me.txt_ubicacion_5 = ""
      Me.txt_ubicacion_6 = ""
End Sub

Private Sub txt_codigo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frmoracle_busqueda_articulos.Show 1
      Me.txt_codigo = var_codigo_busqueda
      Me.txt_descripcion = var_descripcion_busqueda
      Me.txt_codigo.SetFocus
   End If
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_descripcion.SetFocus
   End If
End Sub

Private Sub txt_codigo_LostFocus()
   If Trim(Me.txt_codigo) <> "" Then
      Me.lv_desgloce.ListItems.Clear
      Me.txt_descripcion = ""
      Me.txt_fisico = ""
      Me.txt_apartado = ""
      Me.txt_disponible = ""
      Me.txt_cantidad_ordenes = ""
      Me.txt_ubicacion_1 = ""
      Me.txt_ubicacion_2 = ""
      Me.txt_ubicacion_3 = ""
      Me.txt_ubicacion_4 = ""
      Me.txt_ubicacion_5 = ""
      Me.txt_ubicacion_6 = ""
      For var_j = Len(Me.txt_codigo) + 1 To 8
          Me.txt_codigo = "0" + Me.txt_codigo
      Next var_j
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      rsaux11.Open "SELECT * FROM xxvia_system_items_b WHERE SEGMENT1 = '" + Me.txt_codigo + "' AND ORGANIZATION_ID = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rsaux11.EOF Then
         Me.txt_descripcion = IIf(IsNull(rsaux11!Description), "", rsaux11!Description)
         strconsulta = "select * from Xxvia_vw_existencias_inv where organization_id = ? and subinventory_code = ? and segment1 = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(var_unidad_organizacional))
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_clave_almacen)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
             .Parameters.Append parametro
         End With
         Set rsaux9 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         
         'rsaux9.Open "SELECT * FROM Xxvia_vw_existencias_inv WHERE ORGANIZATION_ID = " + var_unidad_organizacional + " AND SEGMENT1 = '" + Me.txt_codigo + "'  AND SUBINVENTORY_CODE = '" + Me.txt_clave_almacen + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rsaux9.EOF Then
            Me.txt_fisico = IIf(IsNull(rsaux9!CANTMANO), 0, rsaux9!CANTMANO)
            Me.txt_apartado = IIf(IsNull(rsaux9!RESERVADA), 0, rsaux9!RESERVADA)
            Me.txt_disponible = IIf(IsNull(rsaux9!Disponible), 0, rsaux9!Disponible)
         Else
            Me.txt_fisico = 0
            Me.txt_apartado = 0
            Me.txt_disponible = 0
         End If
         rsaux9.Close
         'rsaux9.Open "select attribute2 ubicacion1, attribute3  ubicacion2, attribute4  ubicacion3, attribute5  ubicacion4, attribute6 ubicacion5, attribute7 ubicacion6, inventory_item_id item_id, segment1 item_number, description item_description from  xxvia_system_items_b where  organization_id  = " + var_unidad_organizacional + " and segment1 = '" + Me.txt_codigo + "' order by segment1", cnnoracle_4, adOpenDynamic, adLockOptimistic
         'Me.txt_ubicacion_1 = IIf(IsNull(rsaux9!UBICACION1), "", rsaux9!UBICACION1)
         'Me.txt_ubicacion_2 = IIf(IsNull(rsaux9!UBICACION2), "", rsaux9!UBICACION2)
         'Me.txt_ubicacion_3 = IIf(IsNull(rsaux9!UBICACION3), "", rsaux9!UBICACION3)
         'Me.txt_ubicacion_4 = IIf(IsNull(rsaux9!UBICACION4), "", rsaux9!UBICACION4)
         'Me.txt_ubicacion_5 = IIf(IsNull(rsaux9!UBICACION5), "", rsaux9!UBICACION5)
         'Me.txt_ubicacion_6 = IIf(IsNull(rsaux9!UBICACION6), "", rsaux9!UBICACION6)
         'rsaux9.Close
         strconsulta = "select * from xxvia_Tb_ubicaciones where codigo = ? and organizacion = ?  and almacen = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, var_unidad_organizacional)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_clave_almacen)
              .Parameters.Append parametro
         End With
         Set rsaux8 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         While Not rsaux8.EOF
               If rsaux8!NUMERO = 1 Then
                  Me.txt_ubicacion_1 = IIf(IsNull(rsaux8!ubicacion), "", rsaux8!ubicacion)
               End If
               If rsaux8!NUMERO = 2 Then
                  Me.txt_ubicacion_2 = IIf(IsNull(rsaux8!ubicacion), "", rsaux8!ubicacion)
               End If
               If rsaux8!NUMERO = 3 Then
                  Me.txt_ubicacion_3 = IIf(IsNull(rsaux8!ubicacion), "", rsaux8!ubicacion)
               End If
               If rsaux8!NUMERO = 4 Then
                  Me.txt_ubicacion_4 = IIf(IsNull(rsaux8!ubicacion), "", rsaux8!ubicacion)
               End If
               If rsaux8!NUMERO = 5 Then
                  Me.txt_ubicacion_5 = IIf(IsNull(rsaux8!ubicacion), "", rsaux8!ubicacion)
               End If
               If rsaux8!NUMERO = 6 Then
                  Me.txt_ubicacion_6 = IIf(IsNull(rsaux8!ubicacion), "", rsaux8!ubicacion)
               End If
               rsaux8.MoveNext
         Wend
         rsaux8.Close
         
         If Not rsaux11.EOF Then
            If var_unidad_organizacional = 93 And Me.txt_clave_almacen = "CDI_ALMPT" Then
               var_cadena = "SELECT mr.requirement_date as fecha, HCAS.CUST_ACCOUNT_ID, TL.NAME AS source_header_type_name, oha.source_document_id, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID, HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, MR.inventory_item_id, mr.requirement_date as date_requested, to_number(oha.order_number) as source_header_number, c.description as item_description, mr.reservation_quantity as  requested_quantity, c.segment1 FROM hz_cust_acct_sites_all HCAS,"
               var_cadena = var_cadena + " OE_ORDER_LINES_ALL OLA, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, MTL_RESERVATIONS MR, OE_TRANSACTION_TYPES_TL TL, mtl_parameters mpa Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID  = OHA.INVOICE_TO_ORG_ID AND c.segment1 = '" + Me.txt_codigo + "' AND OLA.HEADER_ID     = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND MR.inventory_item_id    = c.inventory_item_id AND OHA.ship_from_org_id      = C.ORGANIZATION_ID and mpa.organization_code = 'CDI' AND OLA.LINE_ID = MR.ORIG_DEMAND_SOURCE_LINE_ID and OHA.ORDER_TYPE_ID = TL.TRANSAcTION_TYPE_ID AND TL.language = 'ESA' and c.organization_id = mpa.organization_id and mpa.organization_code = 'CDI' and TL.NAME <> 'VIA_VXT_PUBLICIDAD'"
              
            Else
               var_cadena = "SELECT a.last_update_date as fecha, HCAS.CUST_ACCOUNT_ID, a.source_header_type_name, oha.source_document_id, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,A.item_description,A.source_line_number,A.requested_quantity,A.released_status, c.segment1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND c.segment1 = '" + Me.txt_codigo + "'"
               var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID  AND released_status = 'Y' AND A.organization_id = " + var_unidad_organizacional + " and subinventory = '" + Me.txt_clave_almacen + "'"
            End If
            rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               Me.lv_desgloce.ListItems.Clear
               var_cantidad_total = 0
               While Not rs.EOF
                     If rs!source_header_type_name = "VIA_PEDIDO_INTERNO" Or rs!source_header_type_name = "TEX_PEDIDO_INTERNO" Then
                        rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(rs!source_document_id) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux2.EOF Then
                           var_agente = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                        Else
                           rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rs!CUST_ACCOUNT_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                           var_agente = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                           rsaux4.Close
                        End If
                        rsaux2.Close
                     Else
                        rsaux4.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rs!CUST_ACCOUNT_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        var_agente = IIf(IsNull(rsaux4!Name), "", rsaux4!Name)
                        rsaux4.Close
                     End If
                     var_cliente = IIf(IsNull(rs!customer_name), "", rs!customer_name)
                     Set list_item = Me.lv_desgloce.ListItems.Add(, , rs!source_header_number)
                     list_item.SubItems(1) = Format(rs!Fecha, "Short date")
                     list_item.SubItems(2) = var_agente
                     list_item.SubItems(3) = var_cliente
                     list_item.SubItems(4) = Format(IIf(IsNull(rs!requested_quantity), 0, rs!requested_quantity), "###,###,##0.00")
                     var_cantidad_total = var_cantidad_total + IIf(IsNull(rs!requested_quantity), 0, rs!requested_quantity)
                     rs.MoveNext
               Wend
               Me.txt_cantidad_ordenes = Format(var_cantidad_total, "###,###,##0.00")
            Else
               Me.lv_desgloce.ListItems.Clear
            End If
            rs.Close
         End If
      Else
         MsgBox "El código del artículo no existe", vbOKOnly, "ATENCION"
         Me.txt_descripcion = ""
         Me.txt_codigo = ""
      End If
      rsaux11.Close
   Else
      Me.lv_desgloce.ListItems.Clear
      Me.txt_descripcion = ""
      Me.txt_fisico = ""
      Me.txt_apartado = ""
      Me.txt_disponible = ""
      Me.txt_cantidad_ordenes = ""
      Me.txt_ubicacion_1 = ""
      Me.txt_ubicacion_2 = ""
      Me.txt_ubicacion_3 = ""
      Me.txt_ubicacion_4 = ""
      Me.txt_ubicacion_5 = ""
      Me.txt_ubicacion_6 = ""
   End If
End Sub

Private Sub txt_descripcion_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frmoracle_busqueda_articulos.Show 1
      Me.txt_codigo = var_codigo_busqueda
      Me.txt_descripcion = var_descripcion_busqueda
      Me.txt_codigo.SetFocus
   End If
End Sub

Private Sub txt_descripcion_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      If KeyAscii <> 27 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txt_nombre_almacen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_codigo.SetFocus
   End If
End Sub

