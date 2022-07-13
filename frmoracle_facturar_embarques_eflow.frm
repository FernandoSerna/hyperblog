VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_facturar_embarques_eflow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generar facturas en Eflow por embarque"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4530
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_serie 
      Height          =   345
      Left            =   1035
      TabIndex        =   9
      Text            =   "FAEMX"
      Top             =   45
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6840
      Picture         =   "frmoracle_facturar_embarques_eflow.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_factura_nueva 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   360
      Picture         =   "frmoracle_facturar_embarques_eflow.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Crear facturas"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   3180
      Left            =   45
      TabIndex        =   5
      Top             =   1260
      Width           =   7110
      Begin MSComctlLib.ListView lv_facturas 
         Height          =   2970
         Left            =   60
         TabIndex        =   6
         Top             =   150
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   5239
         View            =   3
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
            Text            =   "Pedido"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cliente"
            Object.Width           =   8202
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Factura"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin VB.CommandButton cmd_aceptar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   30
      Picture         =   "frmoracle_facturar_embarques_eflow.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   -15
      TabIndex        =   1
      Top             =   255
      Width           =   7200
   End
   Begin VB.Frame Frame1 
      Height          =   900
      Left            =   30
      TabIndex        =   0
      Top             =   390
      Width           =   7140
      Begin VB.TextBox txt_embarque 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   2160
         TabIndex        =   4
         Top             =   210
         Width           =   2535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Embarque:"
         Height          =   195
         Left            =   1095
         TabIndex        =   3
         Top             =   398
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmoracle_facturar_embarques_eflow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objConn As New adodb.Connection
Dim objCmd As New adodb.Command
Dim objParm As adodb.Parameter

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const SW_SHOWNORMAL = 1
Dim var_ruta_facturas As String



Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
  "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal _
    szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long


Private Sub cmd_aceptar_pedidos_Click()
   For var_j = 1 To Me.lv_facturas.ListItems.Count
       Me.lv_facturas.ListItems.Item(var_j).Selected = True
       If IsNumeric(Me.lv_facturas.selectedItem.SubItems(2)) Then
          rsaux1.Open "select customer_trx_id, cadena as cadena, numero from xxvia_tb_control_doc_fiscales where serie = 'FAEMX' and numero = " + CStr(Me.lv_facturas.selectedItem.SubItems(2)), cnnoracle_4, adOpenDynamic, adLockOptimistic
          If rsaux1.EOF Then
             var_posible = 1
             MsgBox "No existe aun la factura " + Me.txt_serie + CStr(Me.lv_facturas.selectedItem.SubItems(2)) + ", ejecute nuevamente el concurrente", vbOKOnly, "ATENCION"
          Else
             var_cadena = rsaux1!Cadena
             var_cadena_rfc = Mid(var_cadena, 34, 12)
             var_cadena_str = ""
             Open ("C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(Str(Me.lv_facturas.selectedItem.SubItems(2))) + ".FAC") For Output As #1
             For var_i = 1 To Len(var_cadena)
                 If Asc(Mid(var_cadena, var_i, 1)) = 63 Then
                    If Trim(var_cadena_str) = "CONDICIONES_PAGO:- -" And Mid(Me.txt_serie, 1, 2) = "NC" Then
                        var_cadena_str = "CONDICIONES_PAGO:NO IDENTIFICADO"
                    End If
                    Print #1, var_cadena_str
                    var_cadena_str = ""
                 Else
                    var_cadena_str = var_cadena_str + Mid(var_cadena, var_i, 1)
                 End If
             Next var_i
             Print #1, "FIN:"
             Close #1
             var_archivo = "C:\SISTEMAS\sube_fact" + Trim(Str(Me.lv_facturas.selectedItem.SubItems(2))) + ".bat"
             x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(Str(Me.lv_facturas.selectedItem.SubItems(2))) + ".FAC" + "|https://facturas.vianney.mx/cgi-bin/cfds/timbrarGR|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
          End If
          rsaux1.Close
       End If
   Next var_j
End Sub

Private Sub cmd_factura_nueva_Click()
   For VAR_Z = 1 To Me.lv_facturas.ListItems.Count
       Me.lv_facturas.ListItems.Item(VAR_Z).Selected = True
       If IsNumeric(Me.lv_facturas.selectedItem.SubItems(2)) Then
          rsaux1.Open "select customer_trx_id from xxvia_Tb_control_doc_fiscales where serie = '" + Me.txt_serie + "' and numero = " + CStr(Me.lv_facturas.selectedItem.SubItems(2)), cnnoracle_4, adOpenDynamic, adLockOptimistic
          If rsaux1.EOF Then
             var_posible = 1
          Else
          End If
          rsaux1.Close
          URL = "https://facturas.vianney.mx/cgi-bin/cfds/cfdsORACLE?cmd=download_pdf&rfc_emisor=VTH981105F90&serie=" + Trim(Me.txt_serie) + "&folio=" + Trim(CStr(Me.lv_facturas.selectedItem.SubItems(2)))
          buf = Split(URL, ".")
          ext = buf(UBound(buf))
          strSavePath = "C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(CStr(Me.lv_facturas.selectedItem.SubItems(2))) + ".pdf"
          ret = URLDownloadToFile(0, URL, strSavePath, 0, 0)
          If ret = 0 Then
             Call ShellExecute(Me.hwnd, "print", "C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(CStr(Me.lv_facturas.selectedItem.SubItems(2))) + ".PDF", vbNullString, vbNullString, SW_SHOWNORMAL = 1)
          Else
             MsgBox "Error en la factura " + Me.txt_serie + Trim(CStr(Me.lv_facturas.selectedItem.SubItems(2)))
          End If
       End If
    Next VAR_Z
 End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 1500
   Left = 2300
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_embarque_Change()
   Me.lv_facturas.ListItems.Clear
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.lv_facturas.ListItems.Clear
      rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If IsNumeric(Me.txt_embarque) Then
         rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            If IIf(IsNull(rs!CHAR_EMB_ESTATUS), "", rs!CHAR_EMB_ESTATUS) = "I" Or IIf(IsNull(rs!CHAR_EMB_ESTATUS), "", rs!CHAR_EMB_ESTATUS) = "F" Then
               If Not rs.EOF Then
                  If rs!tipo_embarque = 1 Then
                     rsaux.Open "select distinct source_header_number from xxvia_tb_salidas where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  End If
                  If rs!tipo_embarque = 2 Then
                     rsaux.Open "select distinct source_header_number from xxvia_tb_SAlidas_cajas where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  End If
                  var_posible_embarque = 1
               End If
               var_cadena_pedidos = ""
               var_j = 0
               If var_posible_embarque = 1 Then
                  While Not rsaux.EOF
                        If var_cadena_pedidos = "" Then
                           var_cadena_pedidos = "'" + CStr(rsaux!source_header_number) + "'"
                        Else
                           var_cadena_pedidos = var_cadena_pedidos + ", '" + CStr(rsaux!source_header_number) + "'"
                        End If
                        var_j = var_j + 1
                        rsaux.MoveNext
                  Wend
                  rsaux.Close
                  If var_cadena_pedidos <> "" Then
                     rsaux1.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     var_cadena = " SELECT a.source_header_type_name as tipo_pedido, A.SOURCE_HEADER_NUMBER,  HL.ADDRESS1 AS CUSTOMER_NAME,  A.released_status,  sum(shipped_quantity) as cantidad from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID"
                     var_cadena = var_cadena + " AND A.SOURCE_HEADER_NUMBER in (" + var_cadena_pedidos + ") AND A.released_status = 'C' and A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID  and a.source_header_type_name not in ('VIA_PEDIDO_INTERNO','TEX_PEDIDO_INTERNO') group by  a.source_header_type_name , A.SOURCE_HEADER_NUMBER, HL.ADDRESS1,  A.released_status"
                     rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     var_i = 0
                     While Not rsaux.EOF
                           var_i = var_i + 1
                           var_tipo_pedido = rsaux!tipo_pedido
                           rsaux.MoveNext
                     Wend
                     Text1 = var_cadena
                     rsaux.Close
                     If var_j >= var_i Then
                        rsaux.Open "SELECT MIN(APS.TRX_NUMBER) as MINIMO, max(APS.TRX_NUMBER) as maximo From RA_CUSTOMER_TRX_ALL RCT, AR_PAYMENT_SCHEDULES_ALL APS Where INTERFACE_HEADER_ATTRIBUTE1 IN (" + CStr(var_cadena_pedidos) + ") AND INTERFACE_HEADER_ATTRIBUTE2 = '" + var_tipo_pedido + "' AND RCT.customer_trx_id = APS.customer_trx_id", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           txt_de_embarque = IIf(IsNull(rsaux(0).Value), "", rsaux(0).Value)
                           txt_a_embarque = IIf(IsNull(rsaux(1).Value), "", rsaux(1).Value)
                           If txt_de_embarque = "" Then
                              MsgBox "No se han generado todas las facturas, por favor espere un momento y vuelvalo a intentar", vbOKOnly, "ATENCION"
                           Else
                              If var_j >= var_i Then
                                 var_cadena = "SELECT HCAS.CUST_ACCOUNT_ID, HCSU.SITE_USE_ID, A.SOURCE_HEADER_NUMBER,  HL.ADDRESS1 AS CUSTOMER_NAME,  A.released_status, sum(shipped_quantity) as cantidad, HCSU.CUST_ACCT_SITE_ID from hz_cust_acct_sites_all HCAS,  HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID And HPS.LOCATION_ID = HL.LOCATION_ID And HCSU.SITE_USE_ID = OHA.INVOICE_TO_ORG_ID  AND A.SOURCE_HEADER_NUMBER in (" + var_cadena_pedidos + ") AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID  AND released_status = 'C' and  a.source_header_type_name not in ('VIA_PEDIDO_INTERNO','TEX_PEDIDO_INTERNO') group by  HCAS.CUST_ACCOUNT_ID, HCSU.SITE_USE_ID, A.SOURCE_HEADER_NUMBER, HL.ADDRESS1,  A.released_status,  HCSU.CUST_ACCT_SITE_ID "
                                 rsaux4.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 If Not rsaux4.EOF Then
                                    rsaux2.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux4!CUST_ACCOUNT_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                    'Me.txt_agente = rsaux2!collector_id
                                    'Me.txt_nombre_agente = rsaux2!Name
                                    rsaux2.Close
                                    While Not rsaux4.EOF
                                          rsaux5.Open "SELECT APS.TRX_NUMBER as factura From RA_CUSTOMER_TRX_ALL RCT, AR_PAYMENT_SCHEDULES_ALL APS Where INTERFACE_HEADER_ATTRIBUTE1 IN ('" + CStr(rsaux4!source_header_number) + "') AND INTERFACE_HEADER_ATTRIBUTE2 = '" + var_tipo_pedido + "' AND RCT.customer_trx_id = APS.customer_trx_id", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                          Set list_item = Me.lv_facturas.ListItems.Add(, , rsaux4!source_header_number)
                                          list_item.SubItems(1) = IIf(IsNull(rsaux4!customer_name), "", rsaux4!customer_name)
                                          If Not rsaux5.EOF Then
                                             list_item.SubItems(2) = IIf(IsNull(rsaux5!FACTURA), "", rsaux5!FACTURA)
                                          End If
                                          rsaux5.Close
                                          rsaux4.MoveNext
                                    Wend
                                 End If
                                 rsaux4.Close
                              Else
                                 MsgBox "No se han cerrado todos los pedidos del embarque", vbOKOnly, "ATENCION"
                              End If
                           End If
                        End If
                        rsaux.Close
                     Else
                        MsgBox "No se han generado todas las facturas", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "El embarque no tiene facturas", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "No existen facturas para el embarque seleccionado", vbOKOnly, "ATENCION"
               End If
            Else
               If IIf(IsNull(rs!CHAR_EMB_ESTATUS), "", rs!CHAR_EMB_ESTATUS) = "F" Then
                  MsgBox "El embarque ya fue facturado", vbOKOnly, "ATENCION"
               End If
               If IIf(IsNull(rs!CHAR_EMB_ESTATUS), "", rs!CHAR_EMB_ESTATUS) = "" Then
                  MsgBox "El embarque no a sido cerrado", vbOKOnly, "ATENCION"
               End If
               If IIf(IsNull(rs!CHAR_EMB_ESTATUS), "", rs!CHAR_EMB_ESTATUS) = "E" Then
                  MsgBox "El embarque no a sido cerrado", vbOKOnly, "ATENCION"
               End If
            End If
         Else
            MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         MsgBox "Embarque incorrecto", vbOKOnly, "ATENCION"
      End If
End If
End Sub
