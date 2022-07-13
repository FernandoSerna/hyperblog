VERSION 5.00
Begin VB.Form frmoracle_regenerar_pedido 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generar archivo para pediddo"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3570
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frmoracle_regenerar_pedido.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprimir secuencia de pedidos"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3180
      Picture         =   "frmoracle_regenerar_pedido.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_archivo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   75
      Picture         =   "frmoracle_regenerar_pedido.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   150
      Left            =   45
      TabIndex        =   3
      Top             =   270
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   75
      TabIndex        =   0
      Top             =   420
      Width           =   3435
      Begin VB.TextBox txt_pedido 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1125
         TabIndex        =   2
         Top             =   225
         Width           =   1710
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pedido:"
         Height          =   300
         Left            =   150
         TabIndex        =   1
         Top             =   285
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmoracle_regenerar_pedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_archivo_Click()
   Dim var_pedido_surtido As Integer
   If IsNumeric(Me.txt_pedido) Then
      var_pedido_surtido = 0
      rs.Open "SELECT * FROM TB_TEMP_ORACLE_REGENERAR_PEDIDO WHERE PEDIDO_ORIGINAL = " + Me.txt_pedido, cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_pedido_surtido = 1
         VAR_PEDIDO_ACTUAL = IIf(IsNull(rs!PEDIDO_ACTUAL), 0, rs!PEDIDO_ACTUAL)
      End If
      rs.Close
      If var_pedido_surtido = 0 Then
         rs.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         rs.Open "SELECT C.ORDER_TYPE_ID, C.PRICE_LIST_ID, C.SOLD_TO_ORG_ID, C.SHIP_TO_ORG_ID, INVOICE_TO_ORG_ID, A.INVENTORY_ITEM_ID, SEGMENT1 AS CODIGO, REQUESTED_QUANTITY + NVL(CANCELLED_QUANTITY,0) AS CANTIDAD, RELEASED_STATUS FROM WSH_DELIVERABLES_v A, XXVIA_SYSTEM_ITEMS_B B, OE_ORDER_HEADERS_ALL C WHERE SOURCE_HEADER_NUMBER = '" + Me.txt_pedido + "' AND RELEASED_STATUS NOT IN ('Y','C') AND A.INVENTORY_ITEM_ID = B.INVENTORY_ITEM_ID AND A.ORGANIZATION_ID = b.ORGANIZATION_ID and (REQUESTED_QUANTITY + NVL(CANCELLED_QUANTITY,0)) > 0 AND A.SOURCE_HEADER_NUMBER = C.ORDER_NUMBER", cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            While Not rs.EOF
                  rsaux.Open "INSERT INTO TB_TEMP_ORACLE_REGENERAR_PEDIDO (PEDIDO_ORIGINAL, SEGMENT1, CANTIDAD, INVENTORY_ITEM_ID, SOLD_TO_ORG_ID, SHIP_TO_ORG_ID, INVOICE_TO_ORG_ID, PRICE_LIST_ID) VALUES (" + Me.txt_pedido + ",'" + rs!codigo + "'," + CStr(rs!Cantidad) + ", " + CStr(rs!INVENTORY_ITEM_ID) + ", " + CStr(rs!SOLD_TO_ORG_ID) + ", " + CStr(rs!SHIP_TO_ORG_ID) + ", " + CStr(rs!INVOICE_TO_ORG_ID) + "," + CStr(rs!PRICE_LIST_ID) + ")", cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
            Wend
            rs.Close
            rs.Open "select * from tb_temp_oracle_regenerar_pedido where pedido_original = " + Me.txt_pedido, cnn, adOpenDynamic, adLockOptimistic
            rsaux7.Open "select name from qp_secu_list_headers_v where list_header_id = " + CStr(rs!PRICE_LIST_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_lista_precios = rsaux7(0).Value
            rsaux7.Close
            
            var_cadena = "INSERT INTO oe_headers_iface_all (ORDER_SOURCE_ID, orig_sys_document_ref, creation_date, created_by, last_update_date, last_updated_by, operation_code , sold_to_org_id        , SHIP_TO_ORG_id                   ,INVOICE_TO_ORG_ID     , Order_type_ID, PRICE_LIST, org_id, ship_from_org_id)"
            var_cadena = var_cadena + "  VALUES (1001,'SID_RESURTIDO_" + Me.txt_pedido + "',SYSDATE,-1,SYSDATE, -1,'INSERT', " + CStr(rs!SOLD_TO_ORG_ID) + "," + CStr(rs!SHIP_TO_ORG_ID) + "," + CStr(rs!INVOICE_TO_ORG_ID) + ",1042,'" + CStr(var_lista_precios) + "'," + var_empresa + "," + var_unidad_organizacional + ")"
            If rsaux.State = 1 Then
               rsaux.Close
            End If
            rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_i = 0
            While Not rs.EOF
                  var_i = var_i + 1
                  rsaux10.Open "SELECT PRIMARY_UOM_CODE FROM xxvia_system_items_b WHERE INVENTORY_ITEM_ID = " + CStr(rs!INVENTORY_ITEM_ID) + " AND ORGANIZATION_ID = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux10.EOF Then
                     VAR_MEDIDA = rsaux10(0).Value
                  End If
                  rsaux10.Close
                                 
                  var_cadena = "INSERT INTO oe_lines_iface_all (ORDER_SOURCE_ID, orig_sys_document_ref,orig_sys_line_ref,inventory_item_id,ordered_quantity, operation_code, created_by, creation_date, last_updated_by, last_update_date, unit_selling_price, unit_list_price, calculate_price_flag, PRICING_QUANTITY, PRICING_QUANTITY_UOM, ATTRIBUTE1, subinventory, org_id, ship_from_org_id)"
                  var_cadena = var_cadena + " VALUES (1001,'SID_RESURTIDO_" + Trim(CStr(Me.txt_pedido)) + "','" + CStr(var_i) + "', " + CStr(rs!INVENTORY_ITEM_ID) + ", " + CStr(rs!Cantidad) + ",'INSERT', -1,SYSDATE, -1,SYSDATE,0,0,'Y', " + CStr(rs!Cantidad) + ", '" + VAR_MEDIDA + "','','CDI_ALMPT'," + var_empresa + "," + var_unidad_organizacional + ")"
                  'MsgBox var_cadena
                  rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
            Wend
            On Error GoTo SALIR
            rsaux.Open "INSERT INTO oe_actions_iface_all (order_source_ID, orig_sys_document_ref, operation_code) VALUES (1001, 'SID_RESURTIDO_" + Trim(Trim(CStr(Me.txt_pedido))) + "','BOOK_ORDER')", cnnoracle_4, adOpenDynamic, adLockOptimistic
            rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            rsaux.Open "CALL XXVIA_PK_INTERFACES_OM.importar_pedido('SID_RESURTIDO_" + Trim(Trim(CStr(Me.txt_pedido))) + "'," + var_empresa + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
            If rsaux.State = 1 Then
               rsaux.Close
            End If
            rsaux.Open "select order_number from oe_order_headers_all where orig_sys_document_ref = 'SID_RESURTIDO_" + Trim(CStr(Me.txt_pedido)) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rsaux.EOF
                  var_pedido = rsaux(0).Value
                  rsaux.MoveNext
            Wend
            rsaux.Close
            MsgBox var_pedido
            If var_pedido > 0 Then
               rsaux11.Open "update TB_TEMP_ORACLE_REGENERAR_PEDIDO set pedido_actual = " + CStr(var_pedido) + " WHERE PEDIDO_ORIGINAL = " + Me.txt_pedido, cnn, adOpenDynamic, adLockOptimistic
            End If
         Else
            MsgBox "El pedido no existe o no necesita resurtirse", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El pedido " + Me.txt_pedido + " ya se genero en el pedido " + CStr(VAR_PEDIDO_ACTUAL)
         
      End If
      If rs.State = 1 Then
         rs.Close
      End If
   Else
      MsgBox "Número de pedido incorrecto", vbOKOnly, "ATENCION"
   End If
   Exit Sub
SALIR:
   MsgBox Err.Description
   Resume
End Sub

Private Sub cmd_imprimir_Click()
   If IsNumeric(Me.txt_pedido) Then
      cnn.BeginTrans
      rs.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_ORACLE_SECUENCIA_PEDIDOS", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
      Else
         var_consecutivo = 1
      End If
      rs.Close
      rs.Open "INSERT INTO TB_TEMP_ORACLE_SECUENCIA_PEDIDOS (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
      cnn.CommitTrans
      
      var_pedido = CDbl(Me.txt_pedido)
      
      rs.Open "SELECT * FROM OE_ORDER_HEADERS_ALL WHERE ORIG_SYS_DOCUMENT_REF = 'SID_RESURTIDO_" + CStr(var_pedido) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         VAR_PEDIDO_ACTUAL = CDbl(rs!ORDER_NUMBER)
      Else
         VAR_PEDIDO_ACTUAL = 0
      End If
      rs.Close
      While VAR_PEDIDO_ACTUAL > 0
            rsaux.Open "INSERT INTO TB_TEMP_ORACLE_SECUENCIA_PEDIDOS (INTE_TEM_CONSECUTIVO, PEDIDO_ORIGINAL, PEDIDO_ACTUAL) VALUES (" + CStr(var_consecutivo) + "," + CStr(var_pedido) + "," + CStr(VAR_PEDIDO_ACTUAL) + ")", cnn, adOpenDynamic, adLockOptimistic
            var_pedido = VAR_PEDIDO_ACTUAL
            rs.Open "SELECT * FROM OE_ORDER_HEADERS_ALL WHERE ORIG_SYS_DOCUMENT_REF = 'SID_RESURTIDO_" + CStr(var_pedido) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               VAR_PEDIDO_ACTUAL = CDbl(rs!ORDER_NUMBER)
            Else
               VAR_PEDIDO_ACTUAL = 0
            End If
            rs.Close
      Wend
      Set reporte = appl.OpenReport(App.Path + "\rep_oracle_secuencia_pedidos_resurtidos.rpt")
      reporte.RecordSelectionFormula = "{VW_ORACLE_SECUENCIA_PEDIDOS_RESURTIDOS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
      frmvistasprevias.cr.ReportSource = reporte
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      frmvistasprevias.cr.ViewReport
      frmvistasprevias.Caption = "Secuencia pedidos resurtir"
      frmvistasprevias.Show 1
      Set reporte = Nothing
      rs.Open "DELETE FROM TB_TEMP_ORACLE_SECUENCIA_PEDIDOS WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
     
   Else
      MsgBox "No se a indicado un pedido", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 3200
   Left = 4200
End Sub

Private Sub txt_pedido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_archivo.SetFocus
   End If
End Sub
