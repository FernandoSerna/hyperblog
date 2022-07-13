VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_bitacora_lectura 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bitacora de lectura de artículos"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1200
      Left            =   60
      TabIndex        =   7
      Top             =   -15
      Width           =   11505
      Begin VB.TextBox txt_embarque 
         Height          =   405
         Left            =   4290
         TabIndex        =   1
         Top             =   210
         Width           =   1725
      End
      Begin VB.TextBox txt_cliente 
         Height          =   405
         Left            =   1110
         TabIndex        =   2
         Top             =   690
         Width           =   10320
      End
      Begin VB.TextBox txt_pedido 
         Height          =   405
         Left            =   1110
         TabIndex        =   0
         Top             =   210
         Width           =   1755
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Embarque:"
         Height          =   195
         Left            =   3450
         TabIndex        =   11
         Top             =   315
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   165
         TabIndex        =   9
         Top             =   765
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pedido:"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   315
         Width           =   540
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6090
      Left            =   60
      TabIndex        =   10
      Top             =   1125
      Width           =   11535
      Begin VB.TextBox txt_total 
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
         Height          =   510
         Left            =   9000
         TabIndex        =   6
         Top             =   5460
         Width           =   2475
      End
      Begin VB.TextBox txt_descripcion 
         Height          =   405
         Left            =   2565
         TabIndex        =   4
         Top             =   255
         Width           =   8865
      End
      Begin VB.TextBox txt_codigo 
         Height          =   405
         Left            =   1125
         TabIndex        =   3
         Top             =   255
         Width           =   1395
      End
      Begin VB.Frame Frame3 
         Height          =   90
         Left            =   15
         TabIndex        =   12
         Top             =   675
         Width           =   11475
      End
      Begin MSComctlLib.ListView lv_lista 
         Height          =   4590
         Left            =   60
         TabIndex        =   5
         Top             =   810
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   8096
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
            Text            =   "Caja"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cantidad"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha y Hora"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Maquina"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Usuario"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "DVR"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Puerto"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7875
         TabIndex        =   14
         Top             =   5520
         Width           =   915
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmoracle_bitacora_lectura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Top = 0
   Left = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_lineas)
End Sub

Private Sub txt_codigo_Change()
   Me.txt_descripcion = ""
   Me.txt_total = ""
   Me.lv_lista.ListItems.Clear
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_pedido) Then
         Me.lv_lista.ListItems.Clear
         If Trim(Me.txt_codigo) <> "" Then
            If Len(Me.txt_codigo) = 4 Then
               Me.txt_codigo = "0000" + Me.txt_codigo
            Else
               If Len(Me.txt_codigo) = 5 Then
                  Me.txt_codigo = "000" + Me.txt_codigo
               Else
                  If Len(Me.txt_codigo) = 6 Then
                     Me.txt_codigo = "00" + Me.txt_codigo
                  Else
                     If Len(Me.txt_codigo) = 7 Then
                        Me.txt_codigo = "0" + Me.txt_codigo
                     End If
                  End If
               End If
            End If
            rs.Open "SELECT * FROM xxvia_system_items_b WHERE SEGMENT1 = '" + Me.txt_codigo + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               Me.txt_descripcion = IIf(IsNull(rs!Description), "", rs!Description)
               rsaux.Open "SELECT caja, CANTIDAD, to_char(fecha_HORA,'dd/mm/yyyy hh24:mi:ss') AS FECHA, usuario, maquina, dvr, puerto FROM xxvia_tb_bitacora_lectura WHERE PEDIDO = " + Me.txt_pedido + " AND CODIGO = '" + Me.txt_codigo + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  rsaux2.Open "select sum(floa_Sal_cantidad_leida) as suma from xxvia_Tb_salidas_cajas where source_header_number = " + Me.txt_pedido + " and segment1 = '" + Me.txt_codigo + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux2.EOF Then
                     Me.txt_total = Format(IIf(IsNull(rsaux2!suma), 0, rsaux2!suma), "###,###,##0.00")
                  Else
                     Me.txt_total = Format(0, "###,###,##0.00")
                  End If
                  rsaux2.Close
                  While Not rsaux.EOF
                        Set list_item = Me.lv_lista.ListItems.Add(, , rsaux!Caja)
                        list_item.SubItems(1) = Format(rsaux!cantidad)
                        list_item.SubItems(2) = Format(rsaux!Fecha)
                        list_item.SubItems(3) = IIf(IsNull(rsaux!maquina), "", rsaux!maquina)
                        rsaux2.Open "select * from tb_usuarios where vcha_usu_usuario_id = '" + IIf(IsNull(rsaux!USUARIO), "", rsaux!USUARIO) + "'", cnn, adOpenDynamic, adLockOptimistic
                        list_item.SubItems(4) = IIf(IsNull(rsaux2!vcha_usu_nombre), "", rsaux2!vcha_usu_nombre) + " " + IIf(IsNull(rsaux2!vcha_usu_apellidos), "", rsaux2!vcha_usu_apellidos)
                        rsaux2.Close
                        list_item.SubItems(5) = IIf(IsNull(rsaux!DVR), "", rsaux!DVR)
                        list_item.SubItems(6) = IIf(IsNull(rsaux!puerto), "", rsaux!puerto)
                        rsaux.MoveNext
                  Wend
               Else
                  MsgBox "No hay bitacora para los articulos seleccionados"
               End If
               rsaux.Close
            Else
               MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
            End If
            rs.Close
         Else
            MsgBox "No se a indicado un código", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "No se a seleccionado un pedido", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_descripcion_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      If KeyAscii = 27 Then
         Unload Me
      Else
         KeyAscii = 0
      End If
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_pedido_Change()
   Me.txt_cliente = ""
   Me.txt_codigo = ""
   Me.txt_descripcion = ""
   Me.txt_embarque = ""
   Me.txt_total = ""
   Me.lv_lista.ListItems.Clear
End Sub

Private Sub txt_pedido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_pedido) Then
         rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         var_cadena = "SELECT A.DATE_REQUESTED, source_document_id, a.source_header_type_name, HCAS.CUST_ACCOUNT_ID, HCAS.CUST_ACCT_SITE_ID as customer_id, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, "
         var_cadena = var_cadena + " oha.attribute8, oha.attribute9 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND to_number(source_header_number) = " + Me.txt_pedido + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID  AND ROWNUM = 1"
         rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            If rsaux.State = 1 Then
               rsaux.Close
            End If
            rsaux.Open "SELECT INTE_EMB_EMBARQUE AS EMBARQUE FROM XXVIA_TB_sALIDAS_CAJAS WHERE SOURCE_HEADER_NUMBER = " + Me.txt_pedido, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               rsaux1.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + CStr(rsaux!Embarque), cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux1.EOF Then
                  If rs!source_header_type_name = "VIA_PEDIDO_INTERNO" Or rs!source_header_type_name = "TEX_PEDIDO_INTERNO" Then
                     rsaux2.Open "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(rs!source_document_id) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux2.EOF Then
                        Me.txt_cliente = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                     Else
                        Me.txt_cliente = IIf(IsNull(rs!customer_name), "", rs!customer_name)
                     End If
                     rsaux2.Close
                  Else
                     Me.txt_cliente = IIf(IsNull(rs!customer_name), "", rs!customer_name)
                  End If
                  Me.txt_embarque = rsaux!Embarque
                  'rsaux2.Open "SELECT * FROM TB_USUARIOS WHERE VCHA_USU_USUARIO_ID = '" + rsaux1!USUARIO + "'", cnn, adOpenDynamic, adLockOptimistic
                  'If Not rsaux2.EOF Then
                  '   Me.txt_usuario = IIf(IsNull(rsaux2!vcha_usu_nombre), "", rsaux2!vcha_usu_nombre) + " " + IIf(IsNull(rsaux2!vcha_usu_apellidos), "", rsaux2!vcha_usu_apellidos)
                  'End If
                  'rsaux2.Close
                  Me.txt_codigo.SetFocus
               Else
                  MsgBox "El pedido no a sido embarcado", vbOKOnly, "ATENCION"
               End If
               rsaux1.Close
            Else
               
            End If
            rsaux.Close
         Else
            MsgBox "El pedido no existe", vbOKOnly, "ATENCION"
            Me.txt_pedido = ""
         End If
         rs.Close
      Else
         MsgBox "Número de pedido incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
End Sub
