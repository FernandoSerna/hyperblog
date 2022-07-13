VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmorden_compra_oracle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Información en la orden de compra"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11205
      Picture         =   "frmorden_compra_oracle.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   165
      Picture         =   "frmorden_compra_oracle.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Buscar Alt + B"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   135
      TabIndex        =   8
      Top             =   270
      Width           =   11400
   End
   Begin VB.Frame frm_archivo 
      Height          =   915
      Left            =   180
      TabIndex        =   0
      Top             =   285
      Width           =   2475
      Begin VB.TextBox txt_archivo 
         Height          =   390
         Left            =   90
         TabIndex        =   1
         Top             =   420
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         Caption         =   " Orden de compra"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   30
         TabIndex        =   2
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.Frame Frame3 
      Height          =   6810
      Left            =   150
      TabIndex        =   5
      Top             =   405
      Width           =   11385
      Begin MSComctlLib.ListView lv_archivo 
         Height          =   6555
         Left            =   45
         TabIndex        =   6
         Top             =   150
         Width           =   11280
         _ExtentX        =   19897
         _ExtentY        =   11562
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
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
            Text            =   "Orden"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Externo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Interno"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descripción"
            Object.Width           =   7585
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Cantidad"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Proveedor"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Costo"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   45
         Top             =   2505
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   11
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmorden_compra_oracle.frx":073C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmorden_compra_oracle.frx":1016
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmorden_compra_oracle.frx":18F0
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmorden_compra_oracle.frx":1E8C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmorden_compra_oracle.frx":2768
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmorden_compra_oracle.frx":3042
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmorden_compra_oracle.frx":391C
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmorden_compra_oracle.frx":3A2E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmorden_compra_oracle.frx":3B40
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmorden_compra_oracle.frx":3C52
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmorden_compra_oracle.frx":3D64
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lbl_cantidad 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5625
         TabIndex        =   7
         Top             =   5445
         Width           =   3510
      End
   End
End
Attribute VB_Name = "frmorden_compra_oracle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_buscar_Click()
   Me.frm_archivo.Visible = True
   Me.txt_archivo = ""
   Me.txt_archivo.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 0
   frm_archivo.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_archivo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_archivo, ColumnHeader)
End Sub

Private Sub txt_archivo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.lv_archivo.ListItems.Clear
      frmunidad_orden_compra.Show 1
      'txt_archivo = Trim(CStr(var_unidad_OC)) + Trim(txt_archivo)
      'txt_archivo = var_txt_archivo
      'Cadena = "SELECT sysdate as fecha, '' AS VAR_CAJA, 0 AS VAR_PESO, TO_NUMBER(PH.SEGMENT1) AS FOLIO, '' AS LOTE, '' as transporto, 'P' AS TIPO_PROV, '8' AS DESTINO, PO.LINE_NUM, IT.SEGMENT3 as codigo, PO.ITEM_DESCRIPTION as descripcion, PO.UNIT_MEAS_LOOKUP_CODE, (PO.QUANTITY - NVL (REC.NUMB_RC_CANTIDAD, 0)) AS CANTIDAD, PO.LIST_PRICE_PER_UNIT, PO.UNIT_PRICE as costo , PO.PO_HEADER_ID, PV.VENDOR_NAME, PO.PO_LINE_ID, PH.CURRENCY_CODE, PH.SEGMENT1, PV.VENDOR_ID as proveedor, d.destination_organization_id DEST"
      'Cadena = Cadena + "         FROM PO_LINES_ALL@PERPVIA.VIANNEY.COM.MX PO, PO_VENDORS@PERPVIA.VIANNEY.COM.MX PV,"
      'Cadena = Cadena + "         PO_HEADERS_ALL@PERPVIA.VIANNEY.COM.MX PH,"
      'Cadena = Cadena + "         MTL_SYSTEM_ITEMS_B@PERPVIA.VIANNEY.COM.MX IT, po_distributions_all@perpvia.vianney.com.mx d,"
      'Cadena = Cadena + "                        (SELECT   NUMB_RC_ORDEN_COMPRA_ID, NUMB_RC_NUMERO_LINEA,"
      'Cadena = Cadena + "                        NUMB_RC_LINEA_ID, SUM (NUMB_RC_CANTIDAD) NUMB_RC_CANTIDAD"
      'Cadena = Cadena + "                        From RC_TB_RECEPCIONES"
      'Cadena = Cadena + "                        Where NUMB_RC_ORDEN_COMPRA_ID = " + txt_archivo + " and numb_rc_org_id = " + var_unidad_OC
      'Cadena = Cadena + "                        GROUP BY NUMB_RC_ORDEN_COMPRA_ID, NUMB_RC_NUMERO_LINEA, NUMB_RC_LINEA_ID) REC"
      'Cadena = Cadena + "                        Where po.po_line_id = d.po_line_id AND PH.VENDOR_ID = PV.VENDOR_ID"
      'Cadena = Cadena + "                        AND PH.PO_HEADER_ID = PO.PO_HEADER_ID"
      'Cadena = Cadena + "                        AND IT.ORGANIZATION_ID = 83"
      'Cadena = Cadena + "                        AND IT.INVENTORY_ITEM_ID = PO.ITEM_ID"
      'Cadena = Cadena + "                        AND PH.APPROVED_FLAG = 'Y'"
      'Cadena = Cadena + "                        AND PO.PO_LINE_ID = REC.NUMB_RC_LINEA_ID(+)"
      'Cadena = Cadena + "                        AND (PO.QUANTITY - NVL (REC.NUMB_RC_CANTIDAD, 0)) > 0"
      'Cadena = Cadena + "                        AND (PO.CANCEL_FLAG = 'N' OR PO.CANCEL_FLAG IS NULL) "
      'Cadena = Cadena + "                        AND PO.PO_HEADER_ID = (SELECT PO_HEADER_ID"
      'Cadena = Cadena + "                        FROM PO_HEADERS_ALL@PERPVIA.VIANNEY.COM.MX"
      'Cadena = Cadena + "                        Where SEGMENT1 = " + txt_archivo
      'Cadena = Cadena + "                        AND ORG_ID = " + var_unidad_OC + ") ORDER BY PO.LINE_NUM"
      
      
      
      Cadena = "SELECT SYSDATE AS fecha, '' AS var_caja, 0 AS var_peso, TO_NUMBER (h.segment1) AS folio, '' AS lote, '' AS transporto, 'P' AS tipo_prov, '8' AS destino, l.line_num, i.segment3 AS codigo,"
      Cadena = Cadena + " l.item_description, l.unit_meas_lookup_code, (l.quantity - NVL (rec.numb_rc_cantidad, 0)) AS cantidad, l.list_price_per_unit, l.unit_price AS costo, l.po_header_id, v.vendor_name, l.po_line_id, h.currency_code, h.segment1, v.vendor_id AS proveedor, ll.ship_to_organization_id dest"
      Cadena = Cadena + " FROM mtl_system_items_b@perpvia.vianney.com.mx i, po_line_locations_all@perpvia.vianney.com.mx ll, po_lines_all@perpvia.vianney.com.mx l, po_headers_all@perpvia.vianney.com.mx h, po_vendors@perpvia.vianney.com.mx v, (SELECT   numb_rc_orden_compra_id, numb_rc_numero_linea, numb_rc_linea_id, SUM (numb_rc_cantidad) numb_rc_cantidad From rc_tb_recepciones "
      Cadena = Cadena + " Where numb_rc_orden_compra_id = " + txt_archivo + " And numb_rc_org_id = " + var_unidad_OC + " GROUP BY numb_rc_orden_compra_id, numb_rc_numero_linea, numb_rc_linea_id) rec Where ll.po_line_id = l.po_line_id AND ll.po_header_id = l.po_header_id AND ll.po_header_id = h.po_header_id AND l.po_header_id = h.po_header_id AND ll.org_id = l.org_id AND ll.org_id = h.org_id AND l.org_id = h.org_id "
      Cadena = Cadena + " AND i.inventory_item_id = l.item_id AND v.vendor_id = h.vendor_id AND l.po_line_id = rec.numb_rc_linea_id(+) AND (l.quantity - NVL (rec.numb_rc_cantidad, 0)) > 0 AND i.organization_id = 83 aND h.approved_flag = 'Y' AND (h.cancel_flag = 'N' OR h.cancel_flag IS NULL) AND h.po_header_id IN (SELECT po_header_id FROM po_headers_all@perpvia.vianney.com.mx "
      Cadena = Cadena + " WHERE segment1 = " + txt_archivo + " AND org_id = " + var_unidad_OC + ")"
      
      If rs.State = 1 Then
         rs.Close
      End If
      rs.Open Cadena, cnnoracle, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         While Not rs.EOF
               Set list_item = lv_archivo.ListItems.Add(, , txt_archivo)
               If var_empresa = "06" Or var_empresa = "31" Then
                  rsaux2.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_aRT_aRTICULO_ID = '" + Trim(rs!codigo) + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux2.EOF Then
                     list_item.SubItems(2) = IIf(IsNull(rsaux2!vcha_Art_articulo_id), "", rsaux2!vcha_Art_articulo_id)
                     list_item.SubItems(3) = IIf(IsNull(rsaux2!vcha_art_nombre_español), "", rsaux2!vcha_art_nombre_español)
                     rsaux3.Open "insert into orden (orden, vcha_Art_Articulo_id) values ('" + rs!codigo + "','" + list_item.SubItems(2) + "')", cnn, adOpenDynamic, adLockOptimistic
                  Else
                     rsaux4.Open "SELECT * FROM TB_EQUIVALENCIAS WHERE VCHA_EQU_CODIGO_EQUIVALENTE = '" + rs!codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux4.EOF Then
                        rsaux5.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_aRT_aRTICULO_ID = '" + Trim(rsaux4!vcha_Art_articulo_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux5.EOF Then
                           list_item.SubItems(2) = IIf(IsNull(rsaux5!vcha_Art_articulo_id), "", rsaux5!vcha_Art_articulo_id)
                           list_item.SubItems(3) = IIf(IsNull(rsaux5!vcha_art_nombre_español), "", rsaux5!vcha_art_nombre_español)
                           rsaux3.Open "insert into orden (orden, vcha_Art_Articulo_id) values ('" + rs!codigo + "','" + list_item.SubItems(2) + "')", cnn, adOpenDynamic, adLockOptimistic
                        End If
                        rsaux5.Close
                     End If
                     rsaux4.Close
                  End If
                  rsaux2.Close
               End If
               If var_empresa = "18" And Trim(rs!proveedor) = "2458" Then
                  list_item.SubItems(1) = Trim(rs!codigo) + "0"
               Else
                  list_item.SubItems(1) = rs!codigo
               End If
               If var_empresa = "18" And Trim(rs!proveedor) = "2458" Then
                  rsaux2.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + Trim(rs!codigo) + "0" + "'", cnn, adOpenDynamic, adLockOptimistic
               Else
                  rsaux2.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + rs!codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               End If
               If Not rsaux2.EOF Then
                  list_item.SubItems(2) = IIf(IsNull(rsaux2!vcha_Art_articulo_id), "", rsaux2!vcha_Art_articulo_id)
               End If
               rsaux2.Close
               'list_item.SubItems(3) = IIf(IsNull(rs!descripcion), "", rs!descripcion)
               list_item.SubItems(4) = Format(IIf(IsNull(rs!Cantidad), 0, rs!Cantidad), "###,###,##0.00")
               list_item.SubItems(5) = rs!proveedor
               list_item.SubItems(6) = Format(IIf(IsNull(rs!Costo), 0, rs!Costo), "###,###,##0.00")
               rs.MoveNext
         Wend
      Else
         MsgBox "No existe la orden de compra", vbOKOnly, "ATENCION"
      End If
      rs.Close
   End If
   If KeyAscii = 27 Then
      Me.frm_archivo.Visible = False
   End If
End Sub

Private Sub txt_archivo_LostFocus()
   Me.frm_archivo.Visible = False
End Sub
