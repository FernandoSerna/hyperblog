VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmoracle_facturas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturar embarques"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_embarque_pedido 
      Height          =   870
      Left            =   480
      TabIndex        =   34
      Top             =   360
      Width           =   1860
      Begin VB.TextBox txt_embarque_pedido 
         Height          =   375
         Left            =   60
         TabIndex        =   35
         Top             =   450
         Width           =   1725
      End
      Begin VB.Label Label6 
         BackColor       =   &H000000C0&
         Caption         =   " Embarque"
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   45
         TabIndex        =   36
         Top             =   120
         Width           =   1770
      End
   End
   Begin VB.CommandButton cmd_cgi 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2325
      Picture         =   "frmoracle_facturas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Interface con tiendas."
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton cmd_corregir_embarques 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2010
      Picture         =   "frmoracle_facturas.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Corrección de embarques"
      Top             =   0
      Width           =   315
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   4815
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   75
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmd_factura_nueva 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      Picture         =   "frmoracle_facturas.frx":024C
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Crear facturas"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   225
      Left            =   3975
      TabIndex        =   29
      Top             =   45
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   285
      Left            =   3315
      TabIndex        =   28
      Top             =   15
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame frm_correo 
      Height          =   810
      Left            =   2340
      TabIndex        =   25
      Top             =   375
      Width           =   2100
      Begin VB.TextBox txt_embarque_correo 
         Height          =   315
         Left            =   90
         TabIndex        =   26
         Top             =   420
         Width           =   1920
      End
      Begin VB.Label Label11 
         BackColor       =   &H000000C0&
         Caption         =   "Pedido"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   30
         TabIndex        =   27
         Top             =   120
         Width           =   2040
      End
   End
   Begin VB.CommandButton cmd_correo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1680
      Picture         =   "frmoracle_facturas.frx":034E
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Enviar Información"
      Top             =   0
      Width           =   315
   End
   Begin VB.Frame frm_embarque_nota_envio 
      Height          =   870
      Left            =   240
      TabIndex        =   20
      Top             =   360
      Width           =   1860
      Begin VB.TextBox txt_embraue_nota_envio 
         Height          =   375
         Left            =   60
         TabIndex        =   21
         Top             =   450
         Width           =   1725
      End
      Begin VB.Label Label5 
         BackColor       =   &H000000C0&
         Caption         =   " Pedido"
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   45
         TabIndex        =   22
         Top             =   120
         Width           =   1770
      End
   End
   Begin VB.CommandButton cmd_concurrente_facturacion 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1050
      Picture         =   "frmoracle_facturas.frx":0450
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Correr concurrente de facturación"
      Top             =   0
      Width           =   315
   End
   Begin VB.Frame frm_embarque_relacion 
      Height          =   870
      Left            =   720
      TabIndex        =   14
      Top             =   600
      Width           =   1860
      Begin VB.TextBox txt_embarque_relacion 
         Height          =   375
         Left            =   60
         TabIndex        =   16
         Top             =   435
         Width           =   1725
      End
      Begin VB.Label Label3 
         BackColor       =   &H000000C0&
         Caption         =   " Embarque"
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   45
         TabIndex        =   15
         Top             =   120
         Width           =   1770
      End
   End
   Begin VB.Frame Frame4 
      Height          =   600
      Left            =   30
      TabIndex        =   10
      Top             =   1155
      Width           =   7035
      Begin VB.TextBox txt_nombre_agente 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1785
         TabIndex        =   13
         Top             =   180
         Width           =   5175
      End
      Begin VB.TextBox txt_agente 
         Enabled         =   0   'False
         Height          =   330
         Left            =   780
         TabIndex        =   12
         Top             =   180
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   255
         Width           =   555
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6765
      Picture         =   "frmoracle_facturas.frx":0552
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2880
      Picture         =   "frmoracle_facturas.frx":0B8C
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Crear facturas"
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmd_relacion_facturas 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   735
      Picture         =   "frmoracle_facturas.frx":0C8E
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Relación de Facturas"
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton cmd_nota_envio 
      Height          =   315
      Left            =   390
      Picture         =   "frmoracle_facturas.frx":0D90
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Generar Nota de Envio y Correo"
      Top             =   0
      Width           =   345
   End
   Begin VB.Frame Frame3 
      Height          =   750
      Left            =   30
      TabIndex        =   2
      Top             =   405
      Width           =   7035
      Begin VB.TextBox txt_embarque 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   555
         Left            =   1995
         TabIndex        =   3
         Top             =   150
         Width           =   2925
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Embarque:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   330
         TabIndex        =   4
         Top             =   240
         Width           =   1560
      End
   End
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   15
      TabIndex        =   1
      Top             =   345
      Width           =   7080
   End
   Begin VB.Frame Frame1 
      Height          =   3165
      Left            =   15
      TabIndex        =   0
      Top             =   1785
      Width           =   7050
      Begin VB.Frame frm_mensaje 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   1410
         Left            =   315
         TabIndex        =   17
         Top             =   705
         Width           =   6435
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            Caption         =   "Procesando facturación, espere un momento por favor."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1050
            Left            =   150
            TabIndex        =   18
            Top             =   225
            Width           =   6045
         End
      End
      Begin MSComctlLib.ListView lv_facturas 
         Height          =   2970
         Left            =   45
         TabIndex        =   9
         Top             =   135
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
         NumItems        =   6
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
            Text            =   "Cantidad"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Pedido"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cliente"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Titular"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.CommandButton cmd_imprimir_factura 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1350
      Picture         =   "frmoracle_facturas.frx":0E92
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Imprimir facturas"
      Top             =   0
      Width           =   330
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   5805
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   810
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
End
Attribute VB_Name = "frmoracle_facturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objConn As New ADODB.Connection
Dim objCmd As New ADODB.Command
Dim objParm As ADODB.Parameter
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter
Dim var_embarque_pedido As Double
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Private Sub cmd_cgi_Click()
   rs.Open "select * from xxvia_tb_icg_tran_cedis_tienda", cnnicg, adOpenDynamic, adLockOptimistic
   rs.Close
End Sub

Private Sub cmd_concurrente_facturacion_Click()
   Dim clnt As New SoapClient30
   Dim var_con As String
   clnt.MSSoapInit var_webservice
   var_con = clnt.ejecutar_autoinvoice("OM_FACTURAS", 4001)
   Set clint = Nothing
   x = 0
   If x = 1 Then
                  If objConn.State = 1 Then
                     objConn.RollbackTrans
                     objConn.Close
                  End If
                  objConn.Open var_conexion_oracle
                  '… Establecer conexión a la base de datos con el objeto objConn.
                  With objCmd
                       objConn.BeginTrans
                       .ActiveConnection = objConn
                       'LISTO
                       .CommandText = "xxvia_pk_fact_pos_ar_VIANNEY.ejecuta_conc_fact_VIANNEY"
                       .CommandType = adCmdStoredProc
                                  
                       rsaux10.Open "select * from tb_unidadesorganizacionales where vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                       If Not rsaux10.EOF Then
                          var_responsabilidad_facturacion = IIf(IsNull(rsaux10!RESPONSABILIDAD_FACTURACION), "", rsaux10!RESPONSABILIDAD_FACTURACION)
                       End If
                       rsaux10.Close
                                
                                
                       Set objParm = .CreateParameter("p_responsabilidad", adVarChar, adParamInput, 100, var_responsabilidad_facturacion)
                       .Parameters.Append objParm
                                 
                                   
                       'Set objParm = .CreateParameter("p_customer_trx_id", adNumeric, adParamInput, 50, var_customer_trx_id)
                       Set objParm = .CreateParameter("p_customer_trx_id", adNumeric, adParamInput, 50, Null)
                       .Parameters.Append objParm
                               
                       Set objParm = .CreateParameter("p_esperar", adNumeric, adParamInput, 50, 1)
                       .Parameters.Append objParm
                              
                       Set objParm = .CreateParameter("p_fact_pagada", adVarChar, adParamInput, 50, "Y")
                       .Parameters.Append objParm
                                   
                       var_estatus_factura = ""
                       Set objParm = .CreateParameter("p_estatus", adVarChar, adParamOutput, 50, var_estatus_factura)
                       .Parameters.Append objParm
                       rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                       rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                       On Error GoTo SALIR:
                       .execute
                                    
                       var_estatus_factura = .Parameters("p_estatus").Value
                       objConn.CommitTrans
                  End With
                  Set objConn = Nothing
                  Set objCmd = Nothing
                  If objConn.State = 1 Then
                     objConn.RollbackTrans
                     objConn.Close
                  End If
                  objConn.Open var_conexion_oracle
                  '… Establecer conexión a la base de datos con el objeto objConn.
                  With objCmd
                       objConn.BeginTrans
                       .ActiveConnection = objConn
                       'LISTO
                       .CommandText = "xxvia_pk_fact_pos_ar_VIANNEY.ejecuta_conc_fact_VIANNEY"
                       .CommandType = adCmdStoredProc
                                  
                       rsaux10.Open "select * from tb_unidadesorganizacionales where vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                       If Not rsaux10.EOF Then
                          var_responsabilidad_facturacion = IIf(IsNull(rsaux10!RESPONSABILIDAD_FACTURACION), "", rsaux10!RESPONSABILIDAD_FACTURACION)
                       End If
                       rsaux10.Close
                                
                                
                       Set objParm = .CreateParameter("p_responsabilidad", adVarChar, adParamInput, 100, var_responsabilidad_facturacion)
                       .Parameters.Append objParm
                                 
                                   
                       'Set objParm = .CreateParameter("p_customer_trx_id", adNumeric, adParamInput, 50, var_customer_trx_id)
                       Set objParm = .CreateParameter("p_customer_trx_id", adNumeric, adParamInput, 50, Null)
                       .Parameters.Append objParm
                             
                       Set objParm = .CreateParameter("p_esperar", adNumeric, adParamInput, 50, 1)
                       .Parameters.Append objParm
                               
                       Set objParm = .CreateParameter("p_fact_pagada", adVarChar, adParamInput, 50, "N")
                       .Parameters.Append objParm
                                   
                       var_estatus_factura = ""
                       Set objParm = .CreateParameter("p_estatus", adVarChar, adParamOutput, 50, var_estatus_factura)
                       .Parameters.Append objParm
                       rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                       rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                       On Error GoTo SALIR:
                       .execute
                                    
                       var_estatus_factura = .Parameters("p_estatus").Value
                       objConn.CommitTrans
                  End With
                  Set objConn = Nothing
                  Set objCmd = Nothing
   Else
      MsgBox "Favor de correr el concurrente XXVIA - Facturacion VIANNEY (Eflow) directamente en Oracle", vbOKOnly, "ATENCION"
   End If
   Exit Sub
SALIR:
   'MsgBox Err.Description
   'MsgBox Err.Number
   If Err.Number = -2147467259 Or Err.Number = -2147217900 Then
      'MsgBox Err.Description
      rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic

      Resume
   End If
   MsgBox "el proceso de facturación termino con error"
   'MsgBox Err.Description
   'Resume
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   If rsaux1.State = 1 Then
      rsaux1.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If rsaux3.State = 1 Then
      rsaux3.Close
   End If
   If rsaux4.State = 1 Then
      rsaux4.Close
   End If
   Exit Sub
salir_factura:
   MsgBox "Surgio un error al generar los documentos electrónicos", vbOKOnly, "ATENCION"
   Me.Label4.Visible = False
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   If rsaux1.State = 1 Then
      rsaux1.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If rsaux3.State = 1 Then
      rsaux3.Close
   End If
   If rsaux4.State = 1 Then
      rsaux4.Close
   End If
   If objConn.State = 1 Then
      objConn.RollbackTrans
      objConn.Close
   End If
   
End Sub

Private Sub cmd_corregir_embarques_Click()
    rs.Open "select max(embarque)+100 from xxvia_tb_encabezado_embarques", cnnoracle_4, adOpenDynamic, adLockOptimistic
    var_embarque = rs(0).Value
    rs.Close
    rs.Open "insert into xxvia_tb_encabezado_embarques (embarque) values (" + CStr(var_embarque) + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
    MsgBox "Se a terminado el proceso", vbOKOnly, "ATENCION"
End Sub

Private Sub cmd_correo_Click()
   Me.frm_correo.Visible = True
   Me.txt_embarque_correo = ""
   Me.txt_embarque_correo.SetFocus
End Sub

Private Sub cmd_factura_nueva_Click()
   Dim clnt As New SoapClient30
   Dim var_con As String
   Dim var_customer_trx_id As Double
   Dim var_estatus_factura As String
   On Error GoTo SALIR:
   
   
   rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   If IsNumeric(Me.txt_embarque) Then
      rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         If IIf(IsNull(rs!char_emb_estatus), "", rs!char_emb_estatus) = "I" Then
            If Me.lv_facturas.ListItems.Count > 0 Then
               var_cadena_pedidos_global = ""
               For var_i = 1 To Me.lv_facturas.ListItems.Count
                   If var_cadena_pedidos_global = "" Then
                      var_cadena_pedidos_global = Me.lv_facturas.selectedItem
                   Else
                      var_cadena_pedidos_global = var_cadena_pedidos_global + "," + CStr(Me.lv_facturas.selectedItem)
                   End If
               Next var_i
               
               var_si = MsgBox("¿Desea correr el proceso de facturación?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  Me.frm_mensaje.Visible = True
                  clnt.MSSoapInit var_webservice
                  For var_j = 1 To 2
                      var_con = clnt.ejecutar_autoinvoice("OM_FACTURAS", 4002)
                  Next var_j
                  Set clint = Nothing
               
                                 
                                 
                  For var_j = 1 To Me.lv_facturas.ListItems.Count
                      Me.lv_facturas.ListItems.Item(var_j).Selected = True
                      rsaux.Open "SELECT oha.header_id, oha.ordered_date, oha.order_number,  HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, f.orig_system_reference from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU, hz_cust_acct_sites_all f Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND HCSU.site_use_code = 'BILL_TO' and f.cust_acct_site_id = HCAS.CUST_ACCT_SITE_ID and order_number  = '" + Me.lv_facturas.selectedItem.SubItems(3) + "'"
                      If Not rsaux.EOF Then
                         var_cadena = " SELECT a.source_header_type_name as tipo_pedido, A.SOURCE_HEADER_NUMBER,  HL.ADDRESS1 AS CUSTOMER_NAME,  A.released_status,  sum(shipped_quantity) as cantidad from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID"
                         var_cadena = var_cadena + " AND A.SOURCE_HEADER_NUMBER in (" + Me.lv_facturas.selectedItem.SubItems(3) + ") AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID  and a.source_header_type_name not in ('VIA_PEDIDO_INTERNO','TEX_PEDIDO_INTERNO')  group by   a.source_header_type_name, A.SOURCE_HEADER_NUMBER, HL.ADDRESS1,  A.released_status"
                         rsaux1.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                         If Not rsaux1.EOF Then
                            var_encontros = 0
                            VAR_Z = 0
                            While var_encontros = 0
                                  If VAR_Z = 25 Then
                                     Call Command2_Click
                                  End If
                                  var_tipo_pedido = rsaux1!tipo_pedido
                                  var_cadena = "SELECT RCT.CUSTOMER_TRX_ID, APS.TRX_NUMBER, APS.AMOUNT_DUE_ORIGINAL, APS.STATUS, APS.CLASS, HCAS.CUST_ACCT_SITE_ID, HL.ADDRESS1 AS CUSTOMER_NAME From hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, HZ_CUST_SITE_USES_ALL HCSU, RA_CUSTOMER_TRX_ALL RCT, AR_PAYMENT_SCHEDULES_ALL APS Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND RCT.BILL_TO_SITE_USE_ID = HCSU.SITE_USE_ID AND INTERFACE_HEADER_ATTRIBUTE1 IN ('" + Me.lv_facturas.selectedItem + "') AND INTERFACE_HEADER_ATTRIBUTE2 = '" + var_tipo_pedido + "' AND RCT.customer_trx_id = APS.customer_trx_id "
                                  rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                  If Not rsaux2.EOF Then
                                     var_encontros = 1
                                     var_customer_trx_id = rsaux2!customer_Trx_id
                                  End If
                                  rsaux2.Close
                                  VAR_Z = VAR_Z + 1
                            Wend
                         End If
                         rsaux1.Close
                      End If
                      rsaux.Close
                  Next var_j
                     
                  x = 0
                  If x = 1 Then
                  If objConn.State = 1 Then
                     objConn.RollbackTrans
                     objConn.Close
                  End If
                  objConn.Open var_conexion_oracle
                  '… Establecer conexión a la base de datos con el objeto objConn.
                  With objCmd
                       objConn.BeginTrans
                       'LISTO
                       .ActiveConnection = objConn
                       
                       .CommandText = "xxvia_pk_fact_pos_ar_VIANNEY.ejecuta_conc_fact_VIANNEY"
                       .CommandType = adCmdStoredProc
                                  
                       rsaux10.Open "select * from tb_unidadesorganizacionales where vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                       If Not rsaux10.EOF Then
                          var_responsabilidad_facturacion = IIf(IsNull(rsaux10!RESPONSABILIDAD_FACTURACION), "", rsaux10!RESPONSABILIDAD_FACTURACION)
                       End If
                       rsaux10.Close
                                
                                
                       Set objParm = .CreateParameter("p_responsabilidad", adVarChar, adParamInput, 100, var_responsabilidad_facturacion)
                       .Parameters.Append objParm
                                 
                                   
                       'Set objParm = .CreateParameter("p_customer_trx_id", adNumeric, adParamInput, 50, var_customer_trx_id)
                       Set objParm = .CreateParameter("p_customer_trx_id", adNumeric, adParamInput, 50, Null)
                       .Parameters.Append objParm
                               
                       Set objParm = .CreateParameter("p_esperar", adNumeric, adParamInput, 50, 1)
                       .Parameters.Append objParm
                              
                       Set objParm = .CreateParameter("p_fact_pagada", adVarChar, adParamInput, 50, "Y")
                       .Parameters.Append objParm
                                   
                       var_estatus_factura = ""
                       Set objParm = .CreateParameter("p_estatus", adVarChar, adParamOutput, 50, var_estatus_factura)
                       .Parameters.Append objParm
                       rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                       rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                       On Error GoTo SALIR:
                       .execute
                                    
                       var_estatus_factura = .Parameters("p_estatus").Value
                       objConn.CommitTrans
                  End With
                  Set objConn = Nothing
                  Set objCmd = Nothing
                  If objConn.State = 1 Then
                     objConn.RollbackTrans
                     objConn.Close
                  End If
                  objConn.Open var_conexion_oracle
                  '… Establecer conexión a la base de datos con el objeto objConn.
                  With objCmd
                       objConn.BeginTrans
                       .ActiveConnection = objConn
                       'LISTO
                       .CommandText = "xxvia_pk_fact_pos_ar_VIANNEY.ejecuta_conc_fact_VIANNEY"
                       .CommandType = adCmdStoredProc
                                  
                       rsaux10.Open "select * from tb_unidadesorganizacionales where vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                       If Not rsaux10.EOF Then
                          var_responsabilidad_facturacion = IIf(IsNull(rsaux10!RESPONSABILIDAD_FACTURACION), "", rsaux10!RESPONSABILIDAD_FACTURACION)
                       End If
                       rsaux10.Close
                                
                                
                       Set objParm = .CreateParameter("p_responsabilidad", adVarChar, adParamInput, 100, var_responsabilidad_facturacion)
                       .Parameters.Append objParm
                                 
                                   
                       'Set objParm = .CreateParameter("p_customer_trx_id", adNumeric, adParamInput, 50, var_customer_trx_id)
                       Set objParm = .CreateParameter("p_customer_trx_id", adNumeric, adParamInput, 50, Null)
                       .Parameters.Append objParm
                             
                       Set objParm = .CreateParameter("p_esperar", adNumeric, adParamInput, 50, 1)
                       .Parameters.Append objParm
                               
                       Set objParm = .CreateParameter("p_fact_pagada", adVarChar, adParamInput, 50, "N")
                       .Parameters.Append objParm
                                   
                       var_estatus_factura = ""
                       Set objParm = .CreateParameter("p_estatus", adVarChar, adParamOutput, 50, var_estatus_factura)
                       .Parameters.Append objParm
                       rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                       rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                       On Error GoTo SALIR:
                       .execute
                                    
                       var_estatus_factura = .Parameters("p_estatus").Value
                       objConn.CommitTrans
                  End With
                  Set objConn = Nothing
                  Set objCmd = Nothing
                  Else
                     MsgBox "Favor de correr el concurrente XXVIA - Facturacion VIANNEY (Eflow) directamente en Oracle", vbOKOnly, "ATENCION"
                  End If

                     
                     
                     
                  For var_j = 1 To Me.lv_facturas.ListItems.Count
                      Me.lv_facturas.ListItems.Item(var_j).Selected = True
                      rsaux.Open "SELECT oha.header_id, oha.ordered_date, oha.order_number,  HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  E.NAME, f.orig_system_reference from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU, hz_customer_profiles D, ar_collectors E, hz_cust_acct_sites_all f Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND D.collector_id = e.collector_id and oha.order_type_id in (1106) and HCSU.site_use_code = 'BILL_TO' and f.cust_acct_site_id = HCAS.CUST_ACCT_SITE_ID and order_number  = '" + Me.lv_facturas.selectedItem.SubItems(3) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                      If Not rsaux.EOF Then
                         'var_cadena = "SELECT e.collector_id, A.SOURCE_HEADER_NUMBER,  HL.ADDRESS1 AS CUSTOMER_NAME,  A.released_status,  E.NAME , sum(shipped_quantity) as cantidad from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, hz_customer_profiles D, ar_collectors e Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id "
                         'var_cadena = var_cadena + " AND A.SOURCE_HEADER_NUMBER in (" + Me.lv_facturas.selectedItem.SubItems(3) + ") AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND D.collector_id = e.collector_id  group by  e.collector_id, A.SOURCE_HEADER_NUMBER, HL.ADDRESS1,  A.released_status,  E.NAME"
                         var_cadena = "SELECT A.SOURCE_HEADER_NUMBER,  HL.ADDRESS1 AS CUSTOMER_NAME,  A.released_status,  sum(shipped_quantity) as cantidad from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID "
                         var_cadena = var_cadena + " AND A.SOURCE_HEADER_NUMBER in (" + Me.lv_facturas.selectedItem.SubItems(3) + ") AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID group by  A.SOURCE_HEADER_NUMBER, HL.ADDRESS1,  A.released_status"
                         rsaux1.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                         If Not rsaux1.EOF Then
                            var_encontros = 0
                            VAR_Z = 0
                            While var_encontros = 0
                                  If VAR_Z = 1000 Then
                                     VAR_Z = 0
                                  End If
                                  
                                  'var_cadena = "SELECT  a.source_header_type_name from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND to_number(source_header_number) IN (" + Me.lv_facturas.selectedItem.SubItems(3) + ")"
                                  'var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id(+) = e.collector_id  AND ROWNUM = 1"
                                  var_cadena = "SELECT  a.source_header_type_name from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND to_number(source_header_number) IN (" + Me.lv_facturas.selectedItem.SubItems(3) + ")"
                                  var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND  ROWNUM = 1"
                                   
                                  rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                  If Not rsaux2.EOF Then
                                     var_tipo_pedido = rsaux2!source_header_type_name
                                  End If
                                  rsaux2.Close
                                    
                                     
                                  var_cadena = "SELECT APS.TRX_NUMBER, APS.AMOUNT_DUE_ORIGINAL, APS.STATUS, APS.CLASS, HCAS.CUST_ACCT_SITE_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  E.COLLECTOR_ID, E.NAME From hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, HZ_CUST_SITE_USES_ALL HCSU, RA_CUSTOMER_TRX_ALL RCT, AR_PAYMENT_SCHEDULES_ALL APS, ar_collectors E, hz_customer_profiles D Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND RCT.BILL_TO_SITE_USE_ID = HCSU.SITE_USE_ID AND INTERFACE_HEADER_ATTRIBUTE1 IN ('" + Me.lv_facturas.selectedItem + "')  AND INTERFACE_HEADER_ATTRIBUTE2 = '" + var_tipo_pedido + "'  AND RCT.customer_trx_id = APS.customer_trx_id AND E.collector_id = D.COLLECTOR_ID AND D.site_use_id = HCSU.SITE_USE_ID "
                                  rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                  If Not rsaux2.EOF Then
                                     var_encontros = 1
                                  End If
                                  rsaux2.Close
                                  VAR_Z = VAR_Z + 1
                            Wend
                            
                            'var_cadena = "SELECT  a.source_header_type_name from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND to_number(source_header_number) IN (" + Me.lv_facturas.selectedItem.SubItems(3) + ")"
                            'var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id(+) = e.collector_id  AND ROWNUM = 1"
                             
                            var_cadena = "SELECT  a.source_header_type_name from hz_cust_acct_sites_all HCAS,  HZ_PARTY_SITES HPS,  HZ_LOCATIONS HL,  OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A,HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID "
                            var_cadena = var_cadena + " AND to_number(source_header_number) IN (" + Me.lv_facturas.selectedItem.SubItems(3) + ") AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID  AND A.inventory_item_id  = c.inventory_item_id  AND A.ORGANIZATION_ID = C.ORGANIZATION_ID and ROWNUM = 1"
                              
                            rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                            If Not rsaux2.EOF Then
                               var_tipo_pedido = rsaux2!source_header_type_name
                            End If
                            rsaux2.Close
                            var_cadena = "SELECT APS.TRX_NUMBER, APS.AMOUNT_DUE_ORIGINAL, APS.STATUS, APS.CLASS, HCAS.CUST_ACCT_SITE_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  E.COLLECTOR_ID, E.NAME From hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, HZ_CUST_SITE_USES_ALL HCSU, RA_CUSTOMER_TRX_ALL RCT, AR_PAYMENT_SCHEDULES_ALL APS, ar_collectors E, hz_customer_profiles D Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND RCT.BILL_TO_SITE_USE_ID = HCSU.SITE_USE_ID AND INTERFACE_HEADER_ATTRIBUTE1 IN ('" + Me.lv_facturas.selectedItem + "') AND INTERFACE_HEADER_ATTRIBUTE2 = '" + var_tipo_pedido + "' AND RCT.customer_trx_id = APS.customer_trx_id AND E.collector_id = D.COLLECTOR_ID AND D.site_use_id = HCSU.SITE_USE_ID "
                            rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                            If Not rsaux2.EOF Then
                               var_importe_factura = rsaux2!amount_due_original
                               If rsaux3.State = 1 Then
                                  rsaux3.Close
                               End If
                               rsaux3.Open "SELECT CUSTOMER_SITE_USE_ID, SUM(AMOUNT_DUE_REMAINING*-1) AS IMPORTE_TOTAL From AR_PAYMENT_SCHEDULES_ALL WHERE CLASS ='PMT' AND STATUS = 'OP' AND customer_site_use_id = " + Me.lv_facturas.selectedItem.SubItems(5) + " GROUP BY CUSTOMER_SITE_USE_ID", cnnoracle_4, adOpenDynamic, adLockOptimistic
                               If Not rsaux3.EOF Then
                                  VAR_IMPORTE_TOTAL = IIf(IsNull(rsaux3(0).Value), 0, rsaux3(0).Value)
                                  If VAR_IMPORTE_TOTAL >= var_importe_factura Then
                                     If rsaux4.State = 1 Then
                                        rsaux4.Close
                                     End If
                                     rsaux4.Open "SELECT CUSTOMER_ID AS TITULAR_ID, CUSTOMER_SITE_USE_ID AS CLIENTE_ID, AMOUNT_DUE_REMAINING* -1 as AMOUNT_DUE_REMAINING, CASH_RECEIPT_ID, TRX_NUMBER, TRX_DATE From AR_PAYMENT_SCHEDULES_ALL WHERE CLASS ='PMT' AND STATUS = 'OP' AND customer_site_use_id = " + Me.lv_facturas.selectedItem.SubItems(5) + " ORDER BY 6 DESC", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                       
                                     var_importe_aplicar = 0
                                     While Not rsaux4.EOF Or var_importe_factura > 0
                                           If var_importe_factura > 0 Then
                                              If rsaux4!amount_due_remaining >= var_importe_factura Then
                                                 var_importe_aplicar = var_importe_factura
                                                 var_importe_factura = 0
                                              Else
                                                 var_importe_aplicar = rsaux4!amount_due_remaining
                                                 var_importe_factura = var_importe_factura - rsaux4!amount_due_remaining
                                              End If
                                              var_numero_doposito = rsaux4!CASH_RECEIPT_ID
                                              var_numero_factura = rsaux2!trx_number
                                              Set clnt = Nothing
                                              clnt.MSSoapInit var_webservice
                                              On Error GoTo SALIR
                                              var_arreglo = clnt.aplicar_recibo(CStr(var_numero_doposito), CStr(var_numero_factura), CDbl(var_importe_aplicar), Date, CInt(var_empresa))
                                              Set clint = Nothing
                                           End If
                                           rsaux4.MoveNext
                                     Wend
                                     rsaux4.Close
                                  End If
                               Else
                                  MsgBox "La factura no puede ser liquidada ya que no tiene el suficiente importe el cliente", vbOKOnly, "ATENCION"
                               End If
                               rsaux3.Close
                            End If
                            rsaux2.Close
                         End If
                         rsaux1.Close
                      End If
                      rsaux.Close
                  Next var_j
                  If rsaux.State = 1 Then
                     rsaux.Close
                  End If
                  rsaux.Open "update xxvia_tb_encabezado_embarques SET CHAR_EMB_ESTATUS = 'F' where embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  MsgBox "Se a terminado el proceso de facturacion", vbOKOnly, "ATENCION"
                  Me.txt_embarque = ""
                  Me.lv_facturas.ListItems.Clear
                  Me.frm_mensaje.Visible = False
               Else
                  MsgBox "Se a cancelado el proceso de facturación", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "El embarque no contiene movimientos", vbOKOnly, "ATENCION"
            End If
         Else
            If IIf(IsNull(rs!char_emb_estatus), "", rs!char_emb_estatus) = "F" Then
               MsgBox "El embarque ya fue facturado", vbOKOnly, "ATENCION"
            End If
            If IIf(IsNull(rs!char_emb_estatus), "", rs!char_emb_estatus) = "" Then
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
   Exit Sub
SALIR:
   'MsgBox Err.Description
   'MsgBox Err.Number
   If Err.Number = -2147467259 Or Err.Number = -2147217900 Then
      'MsgBox Err.Description
      rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic

      Resume
   End If
   'MsgBox Err.Description
   MsgBox "el proceso de facturación termino con error"
   'MsgBox Err.Description
   'Resume
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   If rsaux1.State = 1 Then
      rsaux1.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If rsaux3.State = 1 Then
      rsaux3.Close
   End If
   If rsaux4.State = 1 Then
      rsaux4.Close
   End If
   Exit Sub
salir_factura:
   MsgBox "Surgio un error al generar los documentos electrónicos", vbOKOnly, "ATENCION"
   Me.Label4.Visible = False
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   If rsaux1.State = 1 Then
      rsaux1.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If rsaux3.State = 1 Then
      rsaux3.Close
   End If
   If rsaux4.State = 1 Then
      rsaux4.Close
   End If
   If objConn.State = 1 Then
      objConn.RollbackTrans
      objConn.Close
   End If

End Sub

Private Sub cmd_imprimir_Click()
   Dim clnt As New SoapClient30
   Dim var_con As String
   Dim var_customer_trx_id As Double
   Dim var_estatus_factura As String
   'On Error GoTo salir:
   
   
   rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   If IsNumeric(Me.txt_embarque) Then
      rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         If IIf(IsNull(rs!char_emb_estatus), "", rs!char_emb_estatus) = "I" Then
            If Me.lv_facturas.ListItems.Count > 0 Then
               var_cadena_pedidos_global = ""
               For var_i = 1 To Me.lv_facturas.ListItems.Count
                   If var_cadena_pedidos_global = "" Then
                      var_cadena_pedidos_global = Me.lv_facturas.selectedItem
                   Else
                      var_cadena_pedidos_global = var_cadena_pedidos_global + "," + CStr(Me.lv_facturas.selectedItem)
                   End If
               Next var_i
               
               x = 1
               If x = 0 Then
               var_cadena = "SELECT a.source_line_id, OHA.HEADER_ID from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND A.SOURCE_HEADER_NUMBER in (" + var_cadena_pedidos_global + ") "
               var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id = e.collector_id AND released_status = 'B' order by A.source_header_number"
               rsaux7.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux7.EOF Then
                  var_tipo_depurado = 1
                  frmoracle_depurar_pedidos.Show 1
               End If
               rsaux7.Close
               
               
               var_cadena = "SELECT a.source_line_id, OHA.HEADER_ID from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND A.SOURCE_HEADER_NUMBER in (" + var_cadena_pedidos_global + ") "
               var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id = e.collector_id AND released_status = 'B' order by A.source_header_number"
               rsaux7.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If rsaux7.EOF Then
                  var_i = 100
               Else
                  var_i = 0
               End If
               rsaux7.Close
               End If
               var_i = 100
               If var_i = 0 Then
                  MsgBox "No se depuraron las ordenes de surtido"
               Else
                  var_si = MsgBox("¿Desea correr el proceso de facturación?", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     Me.frm_mensaje.Visible = True
                     clnt.MSSoapInit var_webservice
                     For var_j = 1 To 2
                         var_con = clnt.ejecutar_autoinvoice("OM_FACTURAS", 4002)
                     Next var_j
                     Set clint = Nothing
                  
                                    
                                    
                     For var_j = 1 To Me.lv_facturas.ListItems.Count
                         Me.lv_facturas.ListItems.Item(var_j).Selected = True
                         rsaux.Open "SELECT oha.header_id, oha.ordered_date, oha.order_number,  HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, f.orig_system_reference from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU, hz_cust_acct_sites_all f Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND HCSU.site_use_code = 'BILL_TO' and f.cust_acct_site_id = HCAS.CUST_ACCT_SITE_ID and order_number  = '" + Me.lv_facturas.selectedItem.SubItems(3) + "'"
                         If Not rsaux.EOF Then
                            var_cadena = " SELECT A.SOURCE_HEADER_NUMBER,  HL.ADDRESS1 AS CUSTOMER_NAME,  A.released_status,  sum(shipped_quantity) as cantidad from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID"
                            var_cadena = var_cadena + " AND A.SOURCE_HEADER_NUMBER in (" + Me.lv_facturas.selectedItem.SubItems(3) + ") AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID group by   A.SOURCE_HEADER_NUMBER, HL.ADDRESS1,  A.released_status"
                            rsaux1.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                            If Not rsaux1.EOF Then
                               var_encontros = 0
                               VAR_Z = 0
                               While var_encontros = 0
                                     If VAR_Z = 25 Then
                                        Call Command2_Click
                                     End If
                                     var_cadena = "SELECT  a.source_header_type_name from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND to_number(source_header_number) IN (" + Me.lv_facturas.selectedItem.SubItems(3) + ")"
                                     var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND  ROWNUM = 1"
                                     rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                     If Not rsaux2.EOF Then
                                        var_tipo_pedido = rsaux2!source_header_type_name
                                     End If
                                     rsaux2.Close
                                     var_cadena = "SELECT RCT.CUSTOMER_TRX_ID, APS.TRX_NUMBER, APS.AMOUNT_DUE_ORIGINAL, APS.STATUS, APS.CLASS, HCAS.CUST_ACCT_SITE_ID, HL.ADDRESS1 AS CUSTOMER_NAME From hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, HZ_CUST_SITE_USES_ALL HCSU, RA_CUSTOMER_TRX_ALL RCT, AR_PAYMENT_SCHEDULES_ALL APS Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND RCT.BILL_TO_SITE_USE_ID = HCSU.SITE_USE_ID AND INTERFACE_HEADER_ATTRIBUTE1 IN ('" + Me.lv_facturas.selectedItem + "') AND INTERFACE_HEADER_ATTRIBUTE2 = '" + var_tipo_pedido + "' AND RCT.customer_trx_id = APS.customer_trx_id "
                                     rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                     If Not rsaux2.EOF Then
                                        var_encontros = 1
                                        var_customer_trx_id = rsaux2!customer_Trx_id
                                     End If
                                     rsaux2.Close
                                     VAR_Z = VAR_Z + 1
                               Wend
                            End If
                            rsaux1.Close
                         End If
                         rsaux.Close
                     Next var_j
                     
                     x = 0
                     If x = 1 Then
                     If objConn.State = 1 Then
                        objConn.RollbackTrans
                        objConn.Close
                     End If
                     objConn.Open var_conexion_oracle
                     '… Establecer conexión a la base de datos con el objeto objConn.
                     With objCmd
                          objConn.BeginTrans
                          .ActiveConnection = objConn
                          'LISTO
                          .CommandText = "xxvia_pk_fact_pos_ar_vianney.ejecuta_conc_fact"
                          .CommandType = adCmdStoredProc
                                     
                          rsaux10.Open "select * from tb_unidadesorganizacionales where vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                          If Not rsaux10.EOF Then
                             var_responsabilidad_facturacion = IIf(IsNull(rsaux10!RESPONSABILIDAD_FACTURACION), "", rsaux10!RESPONSABILIDAD_FACTURACION)
                          End If
                          rsaux10.Close
                                   
                                   
                          Set objParm = .CreateParameter("p_responsabilidad", adVarChar, adParamInput, 100, var_responsabilidad_facturacion)
                          .Parameters.Append objParm
                                    
                                      
                          'Set objParm = .CreateParameter("p_customer_trx_id", adNumeric, adParamInput, 50, var_customer_trx_id)
                          Set objParm = .CreateParameter("p_customer_trx_id", adNumeric, adParamInput, 50, Null)
                          .Parameters.Append objParm
                                
                          Set objParm = .CreateParameter("p_esperar", adNumeric, adParamInput, 50, 1)
                          .Parameters.Append objParm
                                
                          Set objParm = .CreateParameter("p_fact_pagada", adVarChar, adParamInput, 50, "Y")
                          .Parameters.Append objParm
                                   
                          var_estatus_factura = ""
                          Set objParm = .CreateParameter("p_estatus", adVarChar, adParamOutput, 50, var_estatus_factura)
                          .Parameters.Append objParm
                          rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                          rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                          On Error GoTo SALIR:
                          .execute
                                    
                          var_estatus_factura = .Parameters("p_estatus").Value
                          objConn.CommitTrans
                     End With
                     Set objConn = Nothing
                     Set objCmd = Nothing



                     If objConn.State = 1 Then
                        objConn.RollbackTrans
                        objConn.Close
                     End If
                     objConn.Open var_conexion_oracle
                     '… Establecer conexión a la base de datos con el objeto objConn.
                     With objCmd
                          objConn.BeginTrans
                          .ActiveConnection = objConn
                          'LISTO
                          .CommandText = "xxvia_pk_fact_pos_ar_vianney.ejecuta_conc_fact"
                          .CommandType = adCmdStoredProc
                                     
                          rsaux10.Open "select * from tb_unidadesorganizacionales where vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                          If Not rsaux10.EOF Then
                             var_responsabilidad_facturacion = IIf(IsNull(rsaux10!RESPONSABILIDAD_FACTURACION), "", rsaux10!RESPONSABILIDAD_FACTURACION)
                          End If
                          rsaux10.Close
                                   
                                   
                          Set objParm = .CreateParameter("p_responsabilidad", adVarChar, adParamInput, 100, var_responsabilidad_facturacion)
                          .Parameters.Append objParm
                                    
                                      
                          'Set objParm = .CreateParameter("p_customer_trx_id", adNumeric, adParamInput, 50, var_customer_trx_id)
                          Set objParm = .CreateParameter("p_customer_trx_id", adNumeric, adParamInput, 50, Null)
                          .Parameters.Append objParm
                                
                          Set objParm = .CreateParameter("p_esperar", adNumeric, adParamInput, 50, 1)
                          .Parameters.Append objParm
                                
                          Set objParm = .CreateParameter("p_fact_pagada", adVarChar, adParamInput, 50, "N")
                          .Parameters.Append objParm
                                   
                          var_estatus_factura = ""
                          Set objParm = .CreateParameter("p_estatus", adVarChar, adParamOutput, 50, var_estatus_factura)
                          .Parameters.Append objParm
                          rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                          rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                          On Error GoTo SALIR:
                          .execute
                                    
                          var_estatus_factura = .Parameters("p_estatus").Value
                          objConn.CommitTrans
                     End With
                     Set objConn = Nothing
                     Set objCmd = Nothing
                     Else
                        MsgBox "Favor de correr el concurrente XXVIA - Facturacion VIANNEY (Eflow) directamente en Oracle", vbOKOnly, "ATENCION"
                     End If

                     
                     
                     
                     For var_j = 1 To Me.lv_facturas.ListItems.Count
                         Me.lv_facturas.ListItems.Item(var_j).Selected = True
                         rsaux.Open "SELECT oha.header_id, oha.ordered_date, oha.order_number,  HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  E.NAME, f.orig_system_reference from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, HZ_CUST_SITE_USES_ALL HCSU, hz_customer_profiles D, ar_collectors E, hz_cust_acct_sites_all f Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND D.collector_id = e.collector_id and oha.order_type_id in (1106) and HCSU.site_use_code = 'BILL_TO' and f.cust_acct_site_id = HCAS.CUST_ACCT_SITE_ID and order_number  = '" + Me.lv_facturas.selectedItem.SubItems(3) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                         If Not rsaux.EOF Then
                            'var_cadena = "SELECT e.collector_id, A.SOURCE_HEADER_NUMBER,  HL.ADDRESS1 AS CUSTOMER_NAME,  A.released_status,  E.NAME , sum(shipped_quantity) as cantidad from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, hz_customer_profiles D, ar_collectors e Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id "
                            'var_cadena = var_cadena + " AND A.SOURCE_HEADER_NUMBER in (" + Me.lv_facturas.selectedItem.SubItems(3) + ") AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND D.collector_id = e.collector_id  group by  e.collector_id, A.SOURCE_HEADER_NUMBER, HL.ADDRESS1,  A.released_status,  E.NAME"
                            var_cadena = "SELECT A.SOURCE_HEADER_NUMBER,  HL.ADDRESS1 AS CUSTOMER_NAME,  A.released_status,  sum(shipped_quantity) as cantidad from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID "
                            var_cadena = var_cadena + " AND A.SOURCE_HEADER_NUMBER in (" + Me.lv_facturas.selectedItem.SubItems(3) + ") AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID group by  A.SOURCE_HEADER_NUMBER, HL.ADDRESS1,  A.released_status"
                            rsaux1.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                            If Not rsaux1.EOF Then
                               var_encontros = 0
                               VAR_Z = 0
                               While var_encontros = 0
                                     If VAR_Z = 1000 Then
                                        VAR_Z = 0
                                     End If
                                     
                                     'var_cadena = "SELECT  a.source_header_type_name from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND to_number(source_header_number) IN (" + Me.lv_facturas.selectedItem.SubItems(3) + ")"
                                     'var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id(+) = e.collector_id  AND ROWNUM = 1"
                                     var_cadena = "SELECT  a.source_header_type_name from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND to_number(source_header_number) IN (" + Me.lv_facturas.selectedItem.SubItems(3) + ")"
                                     var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND  ROWNUM = 1"
                                     
                                     rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                     If Not rsaux2.EOF Then
                                        var_tipo_pedido = rsaux2!source_header_type_name
                                     End If
                                     rsaux2.Close
                                     
                                     
                                     var_cadena = "SELECT APS.TRX_NUMBER, APS.AMOUNT_DUE_ORIGINAL, APS.STATUS, APS.CLASS, HCAS.CUST_ACCT_SITE_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  E.COLLECTOR_ID, E.NAME From hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, HZ_CUST_SITE_USES_ALL HCSU, RA_CUSTOMER_TRX_ALL RCT, AR_PAYMENT_SCHEDULES_ALL APS, ar_collectors E, hz_customer_profiles D Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND RCT.BILL_TO_SITE_USE_ID = HCSU.SITE_USE_ID AND INTERFACE_HEADER_ATTRIBUTE1 IN ('" + Me.lv_facturas.selectedItem + "')  AND INTERFACE_HEADER_ATTRIBUTE2 = '" + var_tipo_pedido + "'  AND RCT.customer_trx_id = APS.customer_trx_id AND E.collector_id = D.COLLECTOR_ID AND D.site_use_id = HCSU.SITE_USE_ID "
                                     rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                     If Not rsaux2.EOF Then
                                        var_encontros = 1
                                     End If
                                     rsaux2.Close
                                     VAR_Z = VAR_Z + 1
                               Wend
                               
                               'var_cadena = "SELECT  a.source_header_type_name from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND to_number(source_header_number) IN (" + Me.lv_facturas.selectedItem.SubItems(3) + ")"
                               'var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id(+) = e.collector_id  AND ROWNUM = 1"
                               
                               var_cadena = "SELECT  a.source_header_type_name from hz_cust_acct_sites_all HCAS,  HZ_PARTY_SITES HPS,  HZ_LOCATIONS HL,  OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A,HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID "
                               var_cadena = var_cadena + " AND to_number(source_header_number) IN (" + Me.lv_facturas.selectedItem.SubItems(3) + ") AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID  AND A.inventory_item_id  = c.inventory_item_id  AND A.ORGANIZATION_ID = C.ORGANIZATION_ID and ROWNUM = 1"
                               
                               rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                               If Not rsaux2.EOF Then
                                  var_tipo_pedido = rsaux2!source_header_type_name
                               End If
                               rsaux2.Close
                               var_cadena = "SELECT APS.TRX_NUMBER, APS.AMOUNT_DUE_ORIGINAL, APS.STATUS, APS.CLASS, HCAS.CUST_ACCT_SITE_ID, HL.ADDRESS1 AS CUSTOMER_NAME,  E.COLLECTOR_ID, E.NAME From hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, HZ_CUST_SITE_USES_ALL HCSU, RA_CUSTOMER_TRX_ALL RCT, AR_PAYMENT_SCHEDULES_ALL APS, ar_collectors E, hz_customer_profiles D Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND RCT.BILL_TO_SITE_USE_ID = HCSU.SITE_USE_ID AND INTERFACE_HEADER_ATTRIBUTE1 IN ('" + Me.lv_facturas.selectedItem + "') AND INTERFACE_HEADER_ATTRIBUTE2 = '" + var_tipo_pedido + "' AND RCT.customer_trx_id = APS.customer_trx_id AND E.collector_id = D.COLLECTOR_ID AND D.site_use_id = HCSU.SITE_USE_ID "
                               rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                               If Not rsaux2.EOF Then
                                  var_importe_factura = rsaux2!amount_due_original
                                  If rsaux3.State = 1 Then
                                     rsaux3.Close
                                  End If
                                  rsaux3.Open "SELECT CUSTOMER_SITE_USE_ID, SUM(AMOUNT_DUE_REMAINING*-1) AS IMPORTE_TOTAL From AR_PAYMENT_SCHEDULES_ALL WHERE CLASS ='PMT' AND STATUS = 'OP' AND customer_site_use_id = " + Me.lv_facturas.selectedItem.SubItems(5) + " GROUP BY CUSTOMER_SITE_USE_ID", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                  If Not rsaux3.EOF Then
                                     VAR_IMPORTE_TOTAL = IIf(IsNull(rsaux3(0).Value), 0, rsaux3(0).Value)
                                     If VAR_IMPORTE_TOTAL >= var_importe_factura Then
                                        If rsaux4.State = 1 Then
                                           rsaux4.Close
                                        End If
                                        rsaux4.Open "SELECT CUSTOMER_ID AS TITULAR_ID, CUSTOMER_SITE_USE_ID AS CLIENTE_ID, AMOUNT_DUE_REMAINING* -1 as AMOUNT_DUE_REMAINING, CASH_RECEIPT_ID, TRX_NUMBER, TRX_DATE From AR_PAYMENT_SCHEDULES_ALL WHERE CLASS ='PMT' AND STATUS = 'OP' AND customer_site_use_id = " + Me.lv_facturas.selectedItem.SubItems(5) + " ORDER BY 6 DESC", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                        
                                        var_importe_aplicar = 0
                                        While Not rsaux4.EOF Or var_importe_factura > 0
                                              If var_importe_factura > 0 Then
                                                 If rsaux4!amount_due_remaining >= var_importe_factura Then
                                                    var_importe_aplicar = var_importe_factura
                                                    var_importe_factura = 0
                                                 Else
                                                    var_importe_aplicar = rsaux4!amount_due_remaining
                                                    var_importe_factura = var_importe_factura - rsaux4!amount_due_remaining
                                                 End If
                                                 var_numero_doposito = rsaux4!CASH_RECEIPT_ID
                                                 var_numero_factura = rsaux2!trx_number
                                                 Set clnt = Nothing
                                                 clnt.MSSoapInit var_webservice
                                                 On Error GoTo SALIR
                                                 var_arreglo = clnt.aplicar_recibo(CStr(var_numero_doposito), CStr(var_numero_factura), CDbl(var_importe_aplicar), Date, CInt(var_empresa))
                                                 Set clint = Nothing
                                              End If
                                              rsaux4.MoveNext
                                        Wend
                                        rsaux4.Close
                                     End If
                                  Else
                                     MsgBox "La factura no puede ser liquidada ya que no tiene el suficiente importe el cliente", vbOKOnly, "ATENCION"
                                  End If
                                  rsaux3.Close
                               End If
                               rsaux2.Close
                            End If
                            rsaux1.Close
                         End If
                         rsaux.Close
                     Next var_j
                     If rsaux.State = 1 Then
                        rsaux.Close
                     End If
                     rsaux.Open "update xxvia_tb_encabezado_embarques SET CHAR_EMB_ESTATUS = 'F' where embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     MsgBox "Se a terminado el proceso de facturacion", vbOKOnly, "ATENCION"
                     Me.txt_embarque = ""
                     Me.lv_facturas.ListItems.Clear
                     Me.frm_mensaje.Visible = False
                  Else
                     MsgBox "Se a cancelado el proceso de facturación", vbOKOnly, "ATENCION"
                  End If
               End If
            Else
               MsgBox "El embarque no contiene movimientos", vbOKOnly, "ATENCION"
            End If
         Else
            If IIf(IsNull(rs!char_emb_estatus), "", rs!char_emb_estatus) = "F" Then
               MsgBox "El embarque ya fue facturado", vbOKOnly, "ATENCION"
            End If
            If IIf(IsNull(rs!char_emb_estatus), "", rs!char_emb_estatus) = "" Then
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
   Exit Sub
SALIR:
   'MsgBox Err.Description
   'MsgBox Err.Number
   If Err.Number = -2147467259 Or Err.Number = -2147217900 Then
      'MsgBox Err.Description
      rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic

      Resume
   End If
   MsgBox "el proceso de facturación termino con error"
   'MsgBox Err.Description
   'Resume
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   If rsaux1.State = 1 Then
      rsaux1.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If rsaux3.State = 1 Then
      rsaux3.Close
   End If
   If rsaux4.State = 1 Then
      rsaux4.Close
   End If
   Exit Sub
salir_factura:
   MsgBox "Surgio un error al generar los documentos electrónicos", vbOKOnly, "ATENCION"
   Me.Label4.Visible = False
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   If rsaux1.State = 1 Then
      rsaux1.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If rsaux3.State = 1 Then
      rsaux3.Close
   End If
   If rsaux4.State = 1 Then
      rsaux4.Close
   End If
   If objConn.State = 1 Then
      objConn.RollbackTrans
      objConn.Close
   End If
End Sub

Private Sub cmd_imprimir_factura_Click()
   frmoracle_imprimir_facturas.Show 1
End Sub

Private Sub cmd_nota_envio_Click()
    Dim var_i As Integer
    var_i = 0
    If var_i = 0 Then
       Me.frm_embarque_pedido.Visible = True
       Me.txt_embarque_pedido.SetFocus
    Else
       frmoracle_asignar_chofer_unidad.Show
    End If
   'Me.frm_embarque_nota_envio.Visible = True
   'Me.txt_embraue_nota_envio = ""
   'Me.txt_embraue_nota_envio.SetFocus
End Sub

Private Sub cmd_relacion_facturas_Click()
   Me.frm_embarque_relacion.Visible = True
   Me.txt_embarque_relacion = ""
   Me.txt_embarque_relacion.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   rs.Open "Select CONSECUTIVO, CAST(SUBSTRING(DEVOLUCION,7,10) AS INTEGER) AS NUMERO  from valuacion_devoluciones_mal", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux.Open "SELECT * FROM XXVIA_tB_dEVOLUCIONES_CLIENTES WHERE MOVIMIENTO = 'DC' AND NUMERO = " + CStr(rs!numero), cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            rsaux2.Open "UPDATE valuacion_devoluciones_mal SET AGENTE = " + CStr(rsaux!Agente) + ", NOMBRE_aGENTE = '" + rsaux!NOMBRE_AGENTE + "' WHERE CONSECUTIVO = " + CStr(rs!CONSECUTIVO), cnn, adOpenDynamic, adLockOptimistic
         End If
         rsaux.Close
         rs.MoveNext
   Wend
   rs.Close
   rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   rs.Open "SELECT * FROM VALUACION_dEVOLUCIONES_MAL", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux.Open "SELECT  distinct hp.party_name as nombre_titular,  account_number as vcha_tit_titular_id, hcas.cust_account_id AS VCHA_CLI_CLAVE_ID FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors Arc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND hcp.collector_id = " + CStr(rs!Agente) + " and hcas.cust_account_id = " + CStr(rs!TITULAR), cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            rsaux1.Open "UPDATE valuacion_devoluciones_mal SET NOMBRE_TITULAR = '" + CStr(rsaux!nombre_titular) + "' WHERE CONSECUTIVO = " + CStr(rs!CONSECUTIVO), cnn, adOpenDynamic, adLockOptimistic
         End If
         rsaux.Close
         rs.MoveNext
   Wend
   rs.Close
   
   
   
   rs.Open "SELECT * FROM VALUACION_dEVOLUCIONES_MAL", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux.Open "SELECT * FROM RA_CUSTOMER_TRX_ALL WHERE customer_trx_id = " + CStr(rs!FACTURA_ORACLE), cnnoracle_4, adOpenDynamic, adLockOptimistic
         FACTURA = rsaux!trx_number
         rsaux.Close
         rsaux.Open "SELECT * FROM RA_CUSTOMER_TRX_ALL WHERE customer_trx_id = " + CStr(rs!NOTA_CREDITO_ORACLE), cnnoracle_4, adOpenDynamic, adLockOptimistic
         NOTA_CREDITO = rsaux!trx_number
         rsaux.Close
         rsaux.Open "UPDATE VALUACION_dEVOLUCIONES_MAL SET FACTURA = " + CStr(FACTURA) + ", NOTA_CREDITO = " + CStr(NOTA_CREDITO) + " WHERE CONSECUTIVO = " + CStr(rs!CONSECUTIVO), cnn, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
   rs.Open "SELECT * FROM VALUACION_dEVOLUCIONES_MAL", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux.Open "SELECT INVENTORY_ITEM_ID, DESCRIPTION FROM xxvia_system_items_b WHERE SEGMENT1 = '" + rs!codigo + "' AND ORGANIZATION_ID = 93", cnnoracle_4, adOpenDynamic, adLockOptimistic
         var_inventory_item_id = rsaux!inventory_item_id
         var_descripcion = rsaux!Description
         rsaux.Close
         rsaux.Open "SELECT L.CUSTOMER_TRX_ID, TRX_DATE, TRX_NUMBER, E.sold_to_customer_id,  L.inventory_item_id, SALES_ORDER_LINE, SUM(NVL(L.GROSS_unit_selling_price,l.unit_selling_price)) AS PRECIO FROM RA_CUSTOMER_TRX_LINES_ALL L, RA_CUSTOMER_TRX_ALL E, ra_cust_trx_types_all TYPES Where TYPES.TYPE = 'INV' AND TYPES.cust_trx_type_id = E.cust_trx_type_id AND TYPES.org_id = E.org_id AND l.customer_trx_id = E.customer_trx_id AND L.inventory_item_id = " + CStr(var_inventory_item_id) + " AND E.sold_to_customer_id = " + CStr(rs!TITULAR) + " AND L.CUSTOMER_TRX_ID = " + CStr(rs!FACTURA_ORACLE) + " GROUP BY L.CUSTOMER_TRX_ID, TRX_DATE, TRX_NUMBER, E.sold_to_customer_id,  L.inventory_item_id, SALES_ORDER_LINE ORDER BY TRX_NUMBER DESC", cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            var_precio = rsaux!Precio
         End If
         rsaux.Close
         rsaux5.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         var_cadena = "SELECT ARPA.APPLIED_CUSTOMER_TRX_ID AS FACTURA_ID, ARPA.CUSTOMER_TRX_ID AS NOTA_CREDITO_ID, ARPA.ACCTD_AMOUNT_APPLIED_TO AS MONTO_APLICADO, RCT.CUST_TRX_TYPE_ID, RCTL.ATTRIBUTE11, RCTL.ATTRIBUTE10, ARPA.AMOUNT_APPLIED, acr.amount FROM AR_RECEIVABLE_APPLICATIONS_ALL ARPA, RA_CUSTOMER_TRX_ALL RCT, RA_CUSTOMER_TRX_LINES_ALL RCTL, ar_cash_receipts_all acr WHERE ARPA.APPLICATION_TYPE = 'CM' AND ARPA.CUSTOMER_TRX_ID = RCT.CUSTOMER_TRX_ID AND RCT.CUST_TRX_TYPE_ID IN (SELECT ATTRIBUTE2 From RA_CUST_TRX_TYPES_ALL WHERE ATTRIBUTE2 IS NOT NULL) AND ARPA.CUSTOMER_TRX_ID  = RCTL.CUSTOMER_TRX_ID AND RCTL.ATTRIBUTE11 IS NOT NULL AND ARPA.APPLIED_CUSTOMER_TRX_ID = " + CStr(rs!FACTURA_ORACLE) + " and RCTL.ATTRIBUTE10 = acr.cash_receipt_id and ARPA.ACCTD_AMOUNT_APPLIED_TO > 0 order by arpa.last_update_date desc"
         rsaux5.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rsaux5.EOF Then
            var_cadena = "select amount_applied from ar_receivable_applications_all Where applied_customer_trx_id = " + CStr(rs!FACTURA_ORACLE) + " and display = 'Y' and application_type = 'CASH' and cash_receipt_id = " + rsaux5!attribute10
            rsaux6.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rsaux6.EOF Then
               VAR_IMPORTE_TOTAL = rsaux6!amount_applied + CDbl(rsaux5!attribute11)
               If VAR_IMPORTE_TOTAL = 0 Then
                  VAR_PORCENTAJE_FIN = 0
               Else
                  VAR_PORCENTAJE_FIN = 100 - (Round((rsaux6!amount_applied * 100) / VAR_IMPORTE_TOTAL, 2))
               End If
               VAR_NOTA_CREDITO_DF = rsaux5!NOTA_CREDITO_ID
               VAR_PRECIO_NETO = var_precio
               var_precio = var_precio * (1 - (VAR_PORCENTAJE_FIN / 100))
               rsaux.Open "SELECT * FROM RA_CUSTOMER_TRX_ALL WHERE customer_trx_id = " + CStr(rs!FACTURA_ORACLE), cnnoracle_4, adOpenDynamic, adLockOptimistic
               FACTURA = rsaux!trx_number
               rsaux.Close
               rsaux.Open "SELECT * FROM RA_CUSTOMER_TRX_ALL WHERE customer_trx_id = " + CStr(VAR_NOTA_CREDITO_DF), cnnoracle_4, adOpenDynamic, adLockOptimistic
               NOTA_CREDITO = rsaux!trx_number
               rsaux.Close
               rsaux.Open "UPDATE VALUACION_dEVOLUCIONES_MAL SET DESCUENTO_FINANCIERO_FINAL = " + CStr(VAR_PORCENTAJE_FIN) + ", FACTURA_FINAL = " + CStr(FACTURA) + ", NOTA_CREDITO_FINAL = " + CStr(NOTA_CREDITO) + ", PRECIO_FINAL = " + CStr(var_precio) + ", PRECIO_NETO = " + CStr(VAR_PRECIO_NETO) + ", INVENTORY_ITEM_ID = " + CStr(var_inventory_item_id) + ", DESCRIPCION = '" + var_descripcion + "'  WHERE CONSECUTIVO = " + CStr(rs!CONSECUTIVO), cnn, adOpenDynamic, adLockOptimistic
            End If
            rsaux6.Close
         Else
            var_precio = (rs!Precio * 100) / rs!DESCUENTO
            rsaux.Open "UPDATE VALUACION_dEVOLUCIONES_MAL SET PRECIO_FINAL = " + CStr(var_precio) + " WHERE CONSECUTIVO = " + CStr(rs!CONSECUTIVO), cnn, adOpenDynamic, adLockOptimistic
         End If
         rsaux5.Close
         rs.MoveNext
   Wend
   rs.Close
   rs.Open "SELECT * FROM VALUACION_dEVOLUCIONES_MAL", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux.Open "select TRX_NUMBER from RA_CUSTOMER_TRX_LINES_ALL A, RA_CUSTOMER_TRX_ALL B  where sales_order = " + CStr(rs!pedido) + " AND inventory_item_id = " + CStr(rs!inventory_item_id) + " AND A.CUSTOMER_TRX_ID = B.customer_trx_id", cnnoracle_4, adOpenDynamic, adLockOptimistic
         VAR_CADENA_NOTAS_CREDITO = ""
         While Not rsaux.EOF
               If VAR_CADENA_NOTAS_CREDITO = "" Then
                  VAR_CADENA_NOTAS_CREDITO = CStr(rsaux(0).Value)
               Else
                  VAR_CADENA_NOTAS_CREDITO = VAR_CADENA_NOTAS_CREDITO + ", " + CStr(rsaux(0).Value)
               End If
               rsaux.MoveNext
         Wend
         rsaux.Close
         rsaux.Open "UPDATE VALUACION_dEVOLUCIONES_MAL SET NOTA_CREDITO_CLIENTE = '" + CStr(VAR_CADENA_NOTAS_CREDITO) + "' WHERE CONSECUTIVO = " + CStr(rs!CONSECUTIVO), cnn, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub Command2_Click()
   Dim clnt As New SoapClient30
   Dim var_con As String
   clnt.MSSoapInit var_webservice
   'For var_j = 1 To 2
       var_con = clnt.ejecutar_autoinvoice("OM_FACTURAS", 4001)
   'Next var_j
   Set clint = Nothing
End Sub

Private Sub Form_Load()
   Top = 1500
   Left = 2300
   Me.frm_embarque_relacion.Visible = False
   Me.frm_mensaje.Visible = False
   Me.frm_embarque_nota_envio.Visible = False
   Me.frm_correo.Visible = False
   frm_embarque_pedido.Visible = False
   
   
   
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub Frame5_DragDrop(Source As Control, x As Single, Y As Single)

End Sub

Private Sub txt_embarque_Change()
   Me.lv_facturas.ListItems.Clear
   Me.txt_agente = ""
   Me.txt_nombre_agente = ""
End Sub

Private Sub txt_embarque_correo_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If rsaux3.State = 1 Then
         rsaux3.Close
      End If
      If rsaux.State = 1 Then
         rsaux.Close
      End If
      If IsNumeric(Me.txt_embarque_correo) Then
         rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         rs.Open "SELECT * FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = " + Me.txt_embarque_correo, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_cadena = "SELECT HPS.PARTY_ID, HPS.PARTY_SITE_NUMBER,HCA.price_list_id,hcsu.order_type_id,hca.cust_account_id,hcp.site_use_id AS VCHA_CLI_CLAVE_ID,hl.address1 VCHA_CLI_NOMBRE, HCA.* FROM hz_parties hp,hz_party_sites hps,hz_cust_accounts hca,hz_cust_acct_sites_all hcas,hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'BILL_TO'  AND hcp.site_use_id = " + CStr(rs!INVOICE_TO_ORG_ID)
            rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               var_cliente = rsaux!party_site_number
               rsaux4.Open "select hz_contact_points.email_address from hz_contact_points where  hz_contact_points.owner_table_name = 'HZ_PARTIES' and hz_contact_points.owner_table_id =  " + CStr(rsaux!PARTY_ID) + " and hz_contact_points.contact_point_type = 'EMAIL' and    nvl(hz_contact_points.primary_flag,'N') = 'Y' and nvl(hz_contact_points.status,'I') = 'A'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux4.EOF Then
                  VAR_CORREO_ELECTRONICO = rsaux4(0).Value
               Else
                  VAR_CORREO_ELECTRONICO = ""
               End If
               rsaux4.Close
               If VAR_CORREO_ELECTRONICO = "" Then
                  VAR_CORREO_ELECTRONICO = "vluna@vianney.com.mx"
               End If
               If Trim(VAR_CORREO_ELECTRONICO) <> "" Then
                  var_numero_folio_2 = CDbl(Me.txt_embarque_correo)
                  If var_numero_folio_2 > 0 Then
                     var_nombre_archivo = ""
                     If Len(Trim(Str(var_numero_folio_2))) = 1 Then
                         var_nombre_archivo = "00000" + Trim(Str(var_numero_folio_2))
                     End If
                     If Len(Trim(Str(var_numero_folio_2))) = 2 Then
                        var_nombre_archivo = "0000" + Trim(Str(var_numero_folio_2))
                     End If
                     If Len(Trim(Str(var_numero_folio_2))) = 3 Then
                        var_nombre_archivo = "000" + Trim(Str(var_numero_folio_2))
                     End If
                     If Len(Trim(Str(var_numero_folio_2))) = 4 Then
                        var_nombre_archivo = "00" + Trim(Str(var_numero_folio_2))
                     End If
                     If Len(Trim(Str(var_numero_folio_2))) = 5 Then
                        var_nombre_archivo = "0" + Trim(Str(var_numero_folio_2))
                     End If
                     If Len(Trim(Str(var_numero_folio_2))) = 6 Then
                        var_nombre_archivo = Trim(Str(var_numero_folio_2))
                     End If
                     If Dir("c:\notas_franquicias\nota_env.dbf") <> "" Then
                        Set var_tabla = CreateObject("ADODB.connection")
                        var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + "c:\notas_franquicias\" + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
                        rsaux2.Open "delete from nota_env", var_tabla, adOpenDynamic, adLockOptimistic
                        var_eliminar = DeleteFile("c:\notas_franquicias\temp_" + Trim(var_nombre_archivo) + ".dbf")
                        var_eliminar = DeleteFile("c:\notas_franquicias\" + Trim(var_nombre_archivo) + ".dbf")
                        var_copia = CopyFile("c:\notas_franquicias\nota_env.dbf", "c:\notas_franquicias\temp_" + Trim(var_nombre_archivo) + ".dbf", 1)
                        var_si = MsgBox("              ¿Enviar Correo?", vbYesNo, "ATENCION")
                        If var_si = 6 Then
                         
                           rsaux4.Open "SELECT SUBSTR(C.SEGMENT1,4,5) AS CODIGO, B.SHIPPED_QUANTITY, B.UNIT_SELLING_PRICE AS floa_Sal_precio FROM WSH_DELIVERABLES_V A, OE_ORDER_LINES_ALL B, xxvia_system_items_b C WHERE source_header_number = " + CStr(Me.txt_embarque_correo) + " AND A.SOURCE_HEADER_ID = B.HEADER_ID AND A.SOURCE_LINE_ID = B.line_id AND A.SHIPPED_QUANTITY > 0 AND A.ORGANIZATION_ID = C.ORGANIZATION_ID  AND A.INVENTORY_ITEM_ID = C.INVENTORY_ITEM_ID"
                           While Not rsaux4.EOF
                                 var_precio_articulo = IIf(IsNull(rsaux4!floa_Sal_precio), 0, rsaux4!floa_Sal_precio)
                                 Cadena = "insert into c:\notas_franquicias\temp_" + Trim(var_nombre_archivo) + ".dbf (cvenota, cvecliente, clapr, canp1, canp2, canp3, canp4, canp5, canp6, prepr, cvepedido, anocosto, tallas, talla1, talla2, talla3, talla4, talla5, talla6) values ('" + Trim(Str(Me.txt_embarque_correo)) + "', '" + var_clave_cliente + "', '" + Trim(rsaux4!codigo) + "', " + Trim(CStr(IIf(IsNull(rsaux4!SHIPPED_QUANTITY), 0, rsaux4!SHIPPED_QUANTITY))) + ", 0, 0, 0, 0, 0, " + Trim(CStr(Round(var_precio_articulo, 4))) + ", 0, '" + Trim(CStr(2005)) + "',0,0,0,0,0,0,0)"
                                 rsaux2.Open Cadena, var_tabla, adOpenDynamic, adLockOptimistic
                                 rsaux4.MoveNext
                           Wend
                           rsaux4.Close
                           var_tabla.Close
                           var_copia = CopyFile("c:\notas_franquicias\temp_" + Trim(var_nombre_archivo) + ".dbf", "c:\notas_franquicias\" + Trim(var_nombre_archivo) + ".dbf", 1)
                           x = 0
                           If x = 1 Then
                           If Trim(VAR_CORREO_ELECTRONICO) <> "" Then
                              If MAPISession1.SessionID = 0 Then
                                 MAPISession1.SignOn
                              End If
                              MAPIMessages1.SessionID = MAPISession1.SessionID
                              MAPIMessages1.Compose
                              MAPIMessages1.RecipDisplayName = VAR_CORREO_ELECTRONICO
                              MAPIMessages1.RecipAddress = VAR_CORREO_ELECTRONICO
                              MAPIMessages1.AddressResolveUI = True
                              MAPIMessages1.ResolveName
                              MAPIMessages1.MsgSubject = "Nota de envio " + Str(var_numero_folio_2)
                              MAPIMessages1.MsgNoteText = "Se adjunta nota de envio número " + Str(var_numero_folio_2)
                              MAPIMessages1.AttachmentPathName = App.Path + "\" + Trim(var_nombre_archivo) + ".dbf"
                              MAPIMessages1.send True
                              If MAPISession1.SessionID > 0 Then
                                 MAPISession1.SignOff
                              End If
                              End If
                           Else
                              'MsgBox "El cliente no cuenta con una cuenta de correo electronico", vbOKOnly, "ATENCION"
                              MsgBox "Se a generado la nota, favor de buscarla en la ruta c:\notas_franquicias\."
                           End If
                        End If
                     Else
                        MsgBox "No se encuentra el archivo " + App.Path + "\nota_env.dbf, consulte con el administrador del sistema", vbOKOnly, "ATENCION"
                     End If
                  End If
               Else
                  MsgBox "El cliente no cuenta con una cuenta de correo electronico", vbOKOnly, "ATENCION"
               End If
            End If
            rsaux.Close
         Else
            MsgBox "El pedido no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
         Me.frm_correo.Visible = False
      Else
         MsgBox "Número de embarque incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_embarque_correo_LostFocus()
   Me.frm_correo.Visible = False
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.lv_facturas.ListItems.Clear
      rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If IsNumeric(Me.txt_embarque) Then
         rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            If IIf(IsNull(rs!char_emb_estatus), "", rs!char_emb_estatus) = "I" Then
               If Not rs.EOF Then
                  If rs!tipo_embarque = 1 Then
                     rsaux.Open "select distinct source_header_number from xxvia_tb_salidas where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  End If
                  If rs!tipo_embarque = 2 Then
                     rsaux.Open "select distinct source_header_number from xxvia_tb_SAlidas_cajas where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  End If
                  var_posible_embarque = 1
               End If
               var_Cadena_pedidos = ""
               var_j = 0
               If var_posible_embarque = 1 Then
                  While Not rsaux.EOF
                        If var_Cadena_pedidos = "" Then
                           var_Cadena_pedidos = "'" + CStr(rsaux!source_header_number) + "'"
                        Else
                           var_Cadena_pedidos = var_Cadena_pedidos + ", '" + CStr(rsaux!source_header_number) + "'"
                        End If
                        var_j = var_j + 1
                        rsaux.MoveNext
                  Wend
                  rsaux.Close
                  If var_Cadena_pedidos <> "" Then
                     rsaux1.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     'var_cadena = "SELECT e.collector_id, A.SOURCE_HEADER_NUMBER,  HL.ADDRESS1 AS CUSTOMER_NAME,  A.released_status,  E.NAME , sum(shipped_quantity) as cantidad from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, hz_customer_profiles D, ar_collectors e Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id "
                     'var_cadena = var_cadena + " AND A.SOURCE_HEADER_NUMBER in (" + var_cadena_pedidos + ") AND A.released_status = 'C' and A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND D.collector_id = e.collector_id  group by  e.collector_id, A.SOURCE_HEADER_NUMBER, HL.ADDRESS1,  A.released_status,  E.NAME "
               
                     var_cadena = " SELECT a.source_header_type_name as tipo_pedido, A.SOURCE_HEADER_NUMBER,  HL.ADDRESS1 AS CUSTOMER_NAME,  A.released_status,  sum(shipped_quantity) as cantidad from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID"
                     var_cadena = var_cadena + " AND A.SOURCE_HEADER_NUMBER in (" + var_Cadena_pedidos + ") AND A.released_status = 'C' and A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID  and a.source_header_type_name not in ('VIA_PEDIDO_INTERNO','TEX_PEDIDO_INTERNO') group by  a.source_header_type_name , A.SOURCE_HEADER_NUMBER, HL.ADDRESS1,  A.released_status"
               
                     rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     var_i = 0
                     While Not rsaux.EOF
                           var_i = var_i + 1
                           var_tipo_pedido = rsaux!tipo_pedido
                           rsaux.MoveNext
                     Wend
                     Me.Text1 = var_cadena
                     
                     rsaux.Close
                     
                     If var_j >= var_i Then
                        'var_cadena = "SELECT  a.source_header_type_name from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND  to_number(source_header_number) IN (" + VAR_CADENA_PEDIDOS + ")"
                        'var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND  ROWNUM = 1"
                        'rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        'If Not rsaux2.EOF Then
                        '   var_tipo_pedido = rsaux2!source_header_type_name
                        'End If
                        'rsaux2.Close
                        rsaux.Open "SELECT MIN(APS.TRX_NUMBER) as MINIMO, max(APS.TRX_NUMBER) as maximo From RA_CUSTOMER_TRX_ALL RCT, AR_PAYMENT_SCHEDULES_ALL APS Where INTERFACE_HEADER_ATTRIBUTE1 IN (" + CStr(var_Cadena_pedidos) + ") AND INTERFACE_HEADER_ATTRIBUTE2 = '" + var_tipo_pedido + "' AND RCT.customer_trx_id = APS.customer_trx_id", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           txt_de_embarque = IIf(IsNull(rsaux(0).Value), "", rsaux(0).Value)
                           txt_a_embarque = IIf(IsNull(rsaux(1).Value), "", rsaux(1).Value)
                           If txt_de_embarque = "" Then
                              Dim clnt As New SoapClient30
                              Dim var_con As String
                              clnt.MSSoapInit var_webservice
                              var_con = clnt.ejecutar_autoinvoice("OM_FACTURAS", 4001)
                              Set clint = Nothing
                           
                              MsgBox "No se han generado todas las facturas, por favor espere un momento y vuelvalo a intentar", vbOKOnly, "ATENCION"
                           Else
                              If var_j >= var_i Then
                                 var_cadena = "SELECT HCAS.CUST_ACCOUNT_ID, HCSU.SITE_USE_ID, A.SOURCE_HEADER_NUMBER,  HL.ADDRESS1 AS CUSTOMER_NAME,  A.released_status, sum(shipped_quantity) as cantidad, HCSU.CUST_ACCT_SITE_ID from hz_cust_acct_sites_all HCAS,  HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID And HPS.LOCATION_ID = HL.LOCATION_ID And HCSU.SITE_USE_ID = OHA.INVOICE_TO_ORG_ID  AND A.SOURCE_HEADER_NUMBER in (" + var_Cadena_pedidos + ") AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID  AND released_status = 'C' and  a.source_header_type_name not in ('VIA_PEDIDO_INTERNO','TEX_PEDIDO_INTERNO') group by  HCAS.CUST_ACCOUNT_ID, HCSU.SITE_USE_ID, A.SOURCE_HEADER_NUMBER, HL.ADDRESS1,  A.released_status,  HCSU.CUST_ACCT_SITE_ID "
                                 rsaux4.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 If Not rsaux4.EOF Then
                                    rsaux2.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux4!CUST_ACCOUNT_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                    Me.txt_agente = rsaux2!collector_id
                                    Me.txt_nombre_agente = rsaux2!Name
                                    rsaux2.Close
                                    While Not rsaux4.EOF
                                          Set list_item = Me.lv_facturas.ListItems.Add(, , rsaux4!source_header_number)
                                          list_item.SubItems(1) = IIf(IsNull(rsaux4!customer_name), "", rsaux4!customer_name)
                                          list_item.SubItems(2) = IIf(IsNull(rsaux4!cantidad), "", rsaux4!cantidad)
                                          list_item.SubItems(3) = IIf(IsNull(rsaux4!source_header_number), "", rsaux4!source_header_number)
                                          list_item.SubItems(4) = IIf(IsNull(rsaux4!CUST_ACCT_SITE_ID), "", rsaux4!CUST_ACCT_SITE_ID)
                                          list_item.SubItems(5) = IIf(IsNull(rsaux4!site_use_id), "", rsaux4!site_use_id)
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
               If IIf(IsNull(rs!char_emb_estatus), "", rs!char_emb_estatus) = "F" Then
                  MsgBox "El embarque ya fue facturado", vbOKOnly, "ATENCION"
               End If
               If IIf(IsNull(rs!char_emb_estatus), "", rs!char_emb_estatus) = "" Then
                  MsgBox "El embarque no a sido cerrado", vbOKOnly, "ATENCION"
               End If
               If IIf(IsNull(rs!char_emb_estatus), "", rs!char_emb_estatus) = "E" Then
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

Private Sub txt_embarque_pedido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_embarque_pedido) Then
         If rs.State = 1 Then
            rs.Close
         End If
         rs.Open "select * from xxvia_Tb_encabezado_Embarques where embarque = " + Me.txt_embarque_pedido, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            VAR_ESTATUS = IIf(IsNull(rs!char_emb_estatus), "", rs!char_emb_estatus)
            VAR_ESTATUS = "I"
            If VAR_ESTATUS = "I" Then
               var_embarque_pedido = CDbl(Me.txt_embarque_pedido)
               Me.frm_embarque_nota_envio.Visible = True
               Me.txt_embraue_nota_envio = ""
               Me.txt_embraue_nota_envio.SetFocus
            Else
               MsgBox "El embarque no a sido cerrado", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         MsgBox "Número de embarque incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_embarque_pedido_LostFocus()
   Me.frm_embarque_pedido.Visible = False
End Sub

Private Sub txt_embarque_relacion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_embarque_relacion) Then
         cnn.BeginTrans
         rs.Open "select max(inte_tem_consecutivo) from tb_Temp_oracle_relacion_facturas", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
         Else
            var_consecutivo = 1
         End If
         rs.Close
         rs.Open "insert into tb_Temp_oracle_relacion_facturas (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
         cnn.CommitTrans
         var_posible_embarque = 0
         rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque_relacion, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            If rs!tipo_embarque = 1 Then
               If rsaux.State = 1 Then
                  rsaux.Close
               End If
               rsaux.Open "select distinct source_header_number from xxvia_tb_salidas where inte_emb_embarque = " + Me.txt_embarque_relacion, cnnoracle_4, adOpenDynamic, adLockOptimistic
            End If
            If rs!tipo_embarque = 2 Then
               If rsaux.State = 1 Then
                  rsaux.Close
               End If
               rsaux.Open "select distinct source_header_number from xxvia_tb_SAlidas_cajas where inte_emb_embarque = " + Me.txt_embarque_relacion, cnnoracle_4, adOpenDynamic, adLockOptimistic
            End If
            var_posible_embarque = 1
         End If
         rs.Close
         var_Cadena_pedidos = ""
         var_j = 0
         If var_posible_embarque = 1 Then
            While Not rsaux.EOF
                  If var_Cadena_pedidos = "" Then
                     var_Cadena_pedidos = "'" + CStr(rsaux!source_header_number) + "'"
                  Else
                     var_Cadena_pedidos = var_Cadena_pedidos + ", '" + CStr(rsaux!source_header_number) + "'"
                  End If
                  var_j = var_j + 1
                  rsaux.MoveNext
            Wend
            rsaux.Close
            If var_Cadena_pedidos <> "" Then
               rsaux2.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               'var_cadena = "SELECT  a.source_header_type_name from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND to_number(source_header_number) IN (" + var_cadena_pedidos + ")"
               'var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id(+) = e.collector_id  AND ROWNUM = 1"
               var_cadena = "SELECT  a.source_header_type_name from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND  to_number(source_header_number) IN (" + var_Cadena_pedidos + ")"
               var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND  ROWNUM = 1 and source_header_type_name not in ('VIA_PEDIDO_INTERNO','TEX_PEDIDO_INTERNO')"
               
               rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux2.EOF Then
                  var_tipo_pedido = rsaux2!source_header_type_name
               End If
               rsaux2.Close
            
            
               rsaux1.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               'var_cadena = "SELECT e.collector_id, A.SOURCE_HEADER_NUMBER,  HL.ADDRESS1 AS CUSTOMER_NAME,  A.released_status,  E.NAME , sum(shipped_quantity) as cantidad from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, hz_customer_profiles D, ar_collectors e Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id "
               'var_cadena = var_cadena + " AND A.SOURCE_HEADER_NUMBER in (" + var_cadena_pedidos + ") AND A.released_status = 'C' and A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND D.collector_id = e.collector_id  group by  e.collector_id, A.SOURCE_HEADER_NUMBER, HL.ADDRESS1,  A.released_status,  E.NAME "
               var_cadena = " SELECT A.SOURCE_HEADER_NUMBER,  HL.ADDRESS1 AS CUSTOMER_NAME,  A.released_status,  sum(shipped_quantity) as cantidad from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND A.SOURCE_HEADER_NUMBER in (" + var_Cadena_pedidos + ") AND A.released_status = 'C' and A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID  group by  A.SOURCE_HEADER_NUMBER, HL.ADDRESS1,  A.released_status"
               rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               var_i = 0
               While Not rsaux.EOF
                     var_i = var_i + 1
                     rsaux.MoveNext
               Wend
               rsaux.Close
               'MsgBox var_j
               'MsgBox var_i
               If var_j >= var_i Then
                  
                  'var_cadena = "SELECT INTERFACE_HEADER_ATTRIBUTE1, RCT.customer_trx_id, HCAS.CUST_ACCOUNT_ID, APS.TRX_NUMBER, APS.AMOUNT_DUE_ORIGINAL, APS.STATUS, APS.CLASS, HCAS.CUST_ACCT_SITE_ID, HL.ADDRESS1 AS CUSTOMER_NAME,HCSU.site_use_id,  sum(quantity_invoiced) as CANTIDAD  From hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, HZ_CUST_SITE_USES_ALL HCSU, RA_CUSTOMER_TRX_ALL RCT, AR_PAYMENT_SCHEDULES_ALL APS, ra_customer_trx_lines_all rctl Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND RCT.BILL_TO_SITE_USE_ID = HCSU.SITE_USE_ID AND INTERFACE_HEADER_ATTRIBUTE1 IN (" + VAR_CADENA_PEDIDOS + ") AND INTERFACE_HEADER_ATTRIBUTE2 = '" + var_tipo_pedido + "' AND RCT.customer_trx_id = APS.customer_trx_id and rctl.customer_trx_id = rct.customer_trx_id and extended_amount >0"
                  'var_cadena = var_cadena + " GROUP BY INTERFACE_HEADER_ATTRIBUTE1, RCT.customer_trx_id, HCAS.CUST_ACCOUNT_ID, APS.TRX_NUMBER, APS.AMOUNT_DUE_ORIGINAL, APS.STATUS, APS.CLASS, HCAS.CUST_ACCT_SITE_ID, HL.ADDRESS1, HCSU.site_use_id"
                  
                  var_cadena = "SELECT INTERFACE_HEADER_ATTRIBUTE1, RCT.customer_trx_id, HCAS.CUST_ACCOUNT_ID, APS.TRX_NUMBER, APS.AMOUNT_DUE_ORIGINAL, HCAS.CUST_ACCT_SITE_ID, HL.ADDRESS1 AS CUSTOMER_NAME,HCSU.site_use_id,  sum(quantity_invoiced) as CANTIDAD  From hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, HZ_CUST_SITE_USES_ALL HCSU, RA_CUSTOMER_TRX_ALL RCT,  ra_customer_trx_lines_all rctl, xxvia_importe_facturas APS Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND RCT.BILL_TO_SITE_USE_ID = HCSU.SITE_USE_ID AND INTERFACE_HEADER_ATTRIBUTE1 IN (" + var_Cadena_pedidos + ") AND INTERFACE_HEADER_ATTRIBUTE2 = '" + var_tipo_pedido + "' and rctl.customer_trx_id = rct.customer_trx_id and extended_amount >0 AND APS.CUSTOMER_TRX_ID = RCT.CUSTOMER_TRX_ID"
                  var_cadena = var_cadena + " GROUP BY INTERFACE_HEADER_ATTRIBUTE1, RCT.customer_trx_id, HCAS.CUST_ACCOUNT_ID, APS.TRX_NUMBER, APS.AMOUNT_DUE_ORIGINAL, HCAS.CUST_ACCT_SITE_ID, HL.ADDRESS1, HCSU.site_use_id"
                  Text1 = var_cadena
                  rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     While Not rsaux.EOF
                           rsaux1.Open "select * from xxvia_vw_agentes where CUST_ACCOUNT_ID = " + CStr(rsaux!CUST_ACCOUNT_ID) + " and site_use_id = " + CStr(IIf(IsNull(rsaux!site_use_id), 0, rsaux!site_use_id)), cnnoracle_4, adOpenDynamic, adLockOptimistic
                           var_agente = rsaux1!collector_id
                           var_nombre_agente = rsaux1!Name
                           rsaux1.Close
                           var_cadena = " INSERT INTO TB_TEMP_ORACLE_RELACION_FACTURAS (INTE_TEM_CONSECUTIVO, TRX_NUMBER, AMOUNT_DUE_ORIGINAL, CUST_ACCT_SITE_ID, CUSTOMER_NAME, COLLECTOR_ID, NAME, EMBARQUE, CANTIDAD) VALUES "
                           var_cadena = var_cadena + " (" + CStr(var_consecutivo) + ", " + CStr(rsaux!trx_number) + "," + CStr(rsaux!amount_due_original) + ", " + CStr(rsaux!CUST_ACCT_SITE_ID) + ", '" + rsaux!customer_name + "'," + CStr(var_agente) + ", '" + var_nombre_agente + "'," + Me.txt_embarque_relacion + "," + CStr(IIf(IsNull(rsaux!cantidad), 0, rsaux!cantidad)) + ")"
                           rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                           rsaux.MoveNext
                     Wend
                     rsaux1.Open "DELETE FROM TB_TEMP_ORACLE_RELACION_fACTURAS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND TRX_NUMBER IS NULL", cnn, adOpenDynamic, adLockOptimistic
                     rsaux1.Open "select distinct collector_id FROM TB_TEMP_ORACLE_RELACION_fACTURAS where INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                     While Not rsaux1.EOF
                           Set reporte = appl.OpenReport(App.Path + "\rep_oracle_relacion_facturas.rpt")
                           reporte.RecordSelectionFormula = "{VW_TEMP_ORACLE_RELACION_FACTURAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_TEMP_ORACLE_RELACION_FACTURAS.COLLECTOR_ID} = " + CStr(rsaux1!collector_id)
                           frmvistasprevias.cr.ReportSource = reporte
                           For ntablas = 1 To reporte.Database.Tables.Count
                               reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                           Next ntablas
                           frmvistasprevias.cr.ViewReport
                           frmvistasprevias.Caption = "Relación de facturas"
                           frmvistasprevias.Show 1
                           Set reporte = Nothing
                           rsaux1.MoveNext
                     Wend
                     rsaux1.Close
                     rsaux1.Open "delete from tb_Temp_oracle_relacion_facturas where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  Else
                     MsgBox "No existen facturas", vbOKOnly, "ATENCION"
                  End If
                  rsaux.Close
               Else
                  MsgBox "No se han generado todas las facturas", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "No se han generado todas las facturas", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "No existen facturas para el embarque seleccionado", vbOKOnly, "ATENCION"
         End If
         
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_embarque_relacion.Visible = False
   End If
End Sub

Private Sub txt_embarque_relacion_LostFocus()
   Me.frm_embarque_relacion.Visible = False
End Sub

Private Sub txt_embraue_nota_envio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim dl As Long                                 ' Valor devuelto por la función API
      Dim sAttributes As String                  ' Aributos
      Dim sDriver As String                       ' Nombre del controlador
      Dim sDescription As String                ' Descripción del DSN
      Dim sDsnName As String                  ' Nombre del DSN

      Const ODBC_ADD_SYS_DSN As Long = 4         ' Se creará un DSN de sistema
      Const vbAPINull As Long = 0&                         ' Puntero NULL

      ' se elimina
      Const ODBC_REMOVE_SYS_DSN As Long = 6    ' Se eliminará un DSN de sistema
      sDsnName = "DSN=sqlquezada2"
      sDriver = "SQL Server"
      dl = SQLConfigDataSource(0&, ODBC_REMOVE_SYS_DSN, sDriver, sDsnName)

      'se crea
      sDsnName = "sqlsistema"
      sDescription = "sqlsistema"
      sDriver = "SQL Server"
      sAttributes = "DSN=" & sDsnName & Chr(0)
      sAttributes = sAttributes & "Server=" + parametros(0) & Chr$(0)
      sAttributes = sAttributes & "Description=" & sDescription & Chr(0)
      sAttributes = sAttributes & "Database=" + var_bd_reportes & Chr(0)
      strAttributes = strAttributes & "UID=sa" & Chr$(0)
      strAttributes = strAttributes & "PWD=elia" & Chr$(0)
      dl = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYS_DSN, sDriver, sAttributes)
      
      Dim var_location_id As Double
      Dim VAR_CLAVE_USUARIO_MOV As String
      Dim var_fecha_inicio As String
      Dim var_fecha_fin As String
      Dim var_consignacion As String
      Dim var_carta_porte As Integer
      If IsNumeric(Me.txt_embraue_nota_envio) Then
         var_posible_embarque = 1
         var_Cadena_pedidos = Me.txt_embraue_nota_envio
         var_j = 0
         If var_posible_embarque = 1 Then
            If rsaux.State = 1 Then
            rsaux.Close
            End If
            rsaux.Open "alter session set nls_languAge = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            'var_cadena = "SELECT  A.ATTRIBUTE1, B.description as nombre_almacen, g.organization_id, oh.ordered_date, oh.source_document_id, oh.header_id, oh.order_number, oh.transactional_curr_code, NVL(ol.ordered_quantity,0) AS CANTIDAD_PEDIDA, NVL(ol.cancelled_quantity,0) AS CANTIDAD_NEGADA, NVL(ol.shipped_quantity,0)   AS CANTIDAD_surtida, ol.line_id, ol.ordered_item, g.description, ol.order_quantity_uom, ol.inventory_item_id, ol.price_list_id, ol.unit_selling_price, DECODE(ol.cancelled_flag,'Y','CANCELADA','SURTIDA') line_status, ol.flow_status_code, h.linea FROM oe_order_headers_all oh, oe_order_lines_all ol, OE_ORDER_LINES_HISTORY OLH, xxvia_system_items_b g, xxvia_vw_articulos_cat h, po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE order_number  IN (" + var_Cadena_pedidos + ") "
            'var_cadena = var_cadena + " AND oh.header_id = ol.header_id AND ol.ship_from_org_id = " + var_unidad_organizacional + " AND oL.header_id = oLh.header_id(+) AND OL.LINE_ID = OLH.LINE_ID(+) AND ol.inventory_item_id = g.inventory_item_id"
            'var_cadena = var_cadena + " AND g.organization_id = ol.ship_from_org_id AND h.item_id = g.inventory_item_id AND h.organization_id = g.organization_id AND requisition_header_id = OH.source_document_id AND B.secondary_inventory_name = A.ATTRIBUTE1 "
            var_cadena = "SELECT  oh.ordered_date, oh.source_document_id, oh.header_id, oh.order_number, oh.transactional_curr_code, NVL(ol.ordered_quantity,0) AS CANTIDAD_PEDIDA, NVL(ol.cancelled_quantity,0) AS CANTIDAD_NEGADA, NVL(ol.shipped_quantity,0)   AS CANTIDAD_surtida, ol.line_id, ol.ordered_item, ol.order_quantity_uom, ol.inventory_item_id, ol.price_list_id, ol.unit_selling_price, DECODE(ol.cancelled_flag,'Y','CANCELADA','SURTIDA') line_status, ol.flow_status_code"
            var_cadena = var_cadena + " FROM oe_order_headers_all oh, oe_order_lines_all ol, OE_ORDER_LINES_HISTORY OLH WHERE order_number  = " + var_Cadena_pedidos
            var_cadena = var_cadena + " AND oh.header_id = ol.header_id AND ol.ship_from_org_id = 93 AND oL.header_id = oLh.header_id(+) AND OL.LINE_ID = OLH.LINE_ID(+) and  NVL(ol.shipped_quantity,0) > 0"
            If rsaux.State = 1 Then
               rsaux.Close
            End If
            rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_posible_embarque = 0
            If Not rsaux.EOF Then
               var_posible_embarque = 1
            End If
            rsaux.Close
            rsaux.Open "select * from xxvia_tb_encabezado_embarques where embarque = " + CStr(var_embarque_pedido), cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_chofer = ""
            If Not rsaux.EOF Then
               var_chofer = IIf(IsNull(rsaux!CHOFER), "", rsaux!CHOFER)
            Else
               var_chofer = ""
            End If
            If var_chofer = "" Then
               var_posible_embarque = 2
            End If
            rsaux.Close
            rsaux.Open "select * from xxvia_tb_encabezado_embarques where embarque = " + CStr(var_embarque_pedido), cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_transporte = ""
            If Not rsaux.EOF Then
               var_transporte = IIf(IsNull(rsaux!transporte), "", rsaux!transporte)
            Else
               var_transporte = ""
            End If
            If var_transporte = "" Then
               var_posible_embarque = 3
            End If
            rsaux.Close
            var_posible_embarque = 1
            If var_posible_embarque = 1 Then
               var_j = 1
               var_i = 1
               If rsaux.State = 1 Then
                  rsaux.Close
               End If
               rsaux.Open "select distinct INTE_EMB_EMBARQUE from xxvia_tb_salidas where SOURCE_HEADER_NUMBER = " + Me.txt_embraue_nota_envio, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  var_numero_embarque = IIf(IsNull(rsaux!inte_emb_embarque), 0, rsaux!inte_emb_embarque)
               End If
               rsaux.Close
               rsaux.Open "select distinct INTE_EMB_EMBARQUE from xxvia_tb_SAlidas_cajas where SOURCE_HEADER_NUMBER = " + Me.txt_embraue_nota_envio + " AND INTE_EMB_EMBARQUE = " + CStr(var_embarque_pedido), cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  var_si = 1
               Else
                  var_si = 0
               End If
               rsaux.Close
               var_numero_embarque = var_embarque_pedido
               rsaux.Open "alter session set nls_languAge = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               If var_si = 1 Then
                  If var_j = var_i Then
                     var_cadena = "SELECT  A.ATTRIBUTE1, B.description as nombre_almacen, g.organization_id, oh.ordered_date, oh.source_document_id, oh.header_id, oh.order_number, oh.transactional_curr_code, NVL(ol.ordered_quantity,0) AS CANTIDAD_PEDIDA, NVL(ol.cancelled_quantity,0) AS CANTIDAD_NEGADA, NVL(ol.shipped_quantity,0)   AS CANTIDAD_surtida, ol.line_id, ol.ordered_item, g.description, ol.order_quantity_uom, ol.inventory_item_id, ol.price_list_id, ol.unit_selling_price, DECODE(ol.cancelled_flag,'Y','CANCELADA','SURTIDA') line_status, ol.flow_status_code, h.linea FROM oe_order_headers_all oh, oe_order_lines_all ol, OE_ORDER_LINES_HISTORY OLH, xxvia_system_items_b g, xxvia_vw_articulos_cat h, po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE order_number  IN (" + var_Cadena_pedidos + ") "
                     var_cadena = var_cadena + " AND oh.header_id = ol.header_id AND ol.ship_from_org_id = " + var_unidad_organizacional + " AND oL.header_id = oLh.header_id(+) AND OL.LINE_ID = OLH.LINE_ID(+) AND ol.inventory_item_id = g.inventory_item_id"
                     var_cadena = var_cadena + " AND g.organization_id = ol.ship_from_org_id AND h.item_id = g.inventory_item_id AND h.organization_id = g.organization_id AND requisition_header_id = OH.source_document_id AND B.secondary_inventory_name = A.ATTRIBUTE1 "
                     If rsaux.State = 1 Then
                        rsaux.Close
                     End If
                     rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rsaux.EOF Then
                        'MsgBox rsaux!ATTRIBUTE1
                        cnn.BeginTrans
                        If rs.State = 1 Then
                           rs.Close
                        End If
                        rs.Open "select max(inte_tem_consecutivo) from tb_Temp_oracle_NOTA_ENVIO", cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
                        Else
                           var_consecutivo = 1
                        End If
                        rs.Close
                        rs.Open "insert into tb_Temp_oracle_NOTA_ENVIO (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                        cnn.CommitTrans
                        If rsaux1.State = 1 Then
                           rsaux1.Close
                        End If
                        rsaux1.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        rsaux1.Open "SELECT LAST_UPDATE_dATE FROM WSH_DELIVERABLES_V WHERE SOURCE_HEADER_NUMBER = '" + CStr(rsaux!order_number) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        VAR_FECHA_MOVIMIENTO = CStr(Date)
                        VAR_FECHA_MOVIMIENTO = CStr(rsaux1!LAST_UPDATE_DATE)
                        If rsaux3.State = 1 Then
                           rsaux3.Close
                        End If
                        rsaux3.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        rsaux3.Open "SELECT * FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = " + CStr(rsaux!order_number), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        var_header = rsaux3!header_id
                        var_dia = CStr(Day(IIf(IsNull(rsaux3!pricing_date), Date, rsaux3!pricing_date)))
                        var_mes = CStr(Month(IIf(IsNull(rsaux3!pricing_date), Date, rsaux3!pricing_date)))
                        var_año = CStr(Year(IIf(IsNull(rsaux3!pricing_date), Date, rsaux3!pricing_date)))
                        If Len(var_dia) = 1 Then
                           var_dia = "0" + var_dia
                        End If
                        If Len(var_mes) = 1 Then
                           var_mes = "0" + var_mes
                        End If
                        If Len(var_año) = 2 Then
                           var_año = "20" + var_año
                        End If
                        VAR_FECHA_PRECIO = var_dia + "/" + var_mes + "/" + var_año
                        rsaux3.Close
                        'SELECT DISTINCT b.secondary_inventory_name AS CLAVE_ALMACEN, B.DESCRIPTION AS NOMBRE_ALMACEN  FROM WSH_DELIVERABLES_V A, mtl_subinventories_all_v B WHERE source_header_number = 74124 AND B.secondary_inventory_name = A.subinventory AND A.organization_id = b.organization_id AND A.SOURCE_HEADER_ID = 4355290
                        rsaux3.Open "SELECT DISTINCT b.secondary_inventory_name AS CLAVE_ALMACEN, B.DESCRIPTION AS NOMBRE_ALMACEN  FROM WSH_DELIVERABLES_V A, mtl_subinventories_all_v B WHERE source_header_number = " + CStr(rsaux!order_number) + " AND B.secondary_inventory_name = A.subinventory AND A.organization_id = b.organization_id AND A.SOURCE_HEADER_ID = " + CStr(var_header), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux3.EOF Then
                           var_almacen = rsaux3!CLAVE_ALMACEN
                           var_nombre_almacen = rsaux3!nombre_almacen
                        End If
                        rsaux3.Close
                        If var_almacen = "CDISTEX_PT" Then
                           var_almacen = "TEX_PT_QL"
                           var_nombre_almacen = "EL VERGEL PRODUCTO TERMINADO TEX"
                        End If
                     
                        rsaux3.Open "select * from mtl_secondary_inventories where secondary_inventory_name = '" + rsaux!attribute1 + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        var_consignacion = IIf(IsNull(rsaux3!attribute3), "", rsaux3!attribute3)
                        var_almacen_icg = IIf(IsNull(rsaux!attribute1), "", rsaux!attribute1)
                        If Not rsaux3.EOF Then
                           var_location_id = IIf(IsNull(rsaux3!LOCATION_ID), 0, rsaux3!LOCATION_ID)
                           If var_location_id > 0 Then
                              rsaux4.Open "select ADDRESS_LINE_1, ADDRESS_LINE_2, TOWN_OR_CITY, REGION_1, COUNTRY, POSTAL_CODE  from hr_locations_all where location_id = '" + CStr(CDbl(var_location_id)) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              VAR_DIRECCION = IIf(IsNull(rsaux4!ADDRESS_LINE_1), "", rsaux4!ADDRESS_LINE_1)
                              VAR_COLONIA = IIf(IsNull(rsaux4!ADDRESS_LINE_2), "", rsaux4!ADDRESS_LINE_2)
                              var_ciudad = IIf(IsNull(rsaux4!TOWN_OR_CITY), "", rsaux4!TOWN_OR_CITY)
                              var_estado = IIf(IsNull(rsaux4!REGION_1), "", rsaux4!REGION_1)
                              var_pais = IIf(IsNull(rsaux4!COUNTRY), "", rsaux4!COUNTRY)
                              VAR_CP = IIf(IsNull(rsaux4!POSTAL_code), "", rsaux4!POSTAL_code)
                              rsaux4.Close
                           Else
                              VAR_DIRECCION = ""
                              VAR_COLONIA = ""
                              var_ciudad = ""
                              var_estado = ""
                              var_pais = ""
                              VAR_CP = ""
                           End If
                        End If
                        rsaux3.Close
                        rsaux3.Open "SELECT * FROM TB_USUARIOS WHERE VCHA_USU_USUARIO_ID = '" + VAR_CLAVE_USUARIO_MOV + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux3.EOF Then
                           VAR_NOMBRE_USUARIO_ENTREGO = IIf(IsNull(rsaux3!vcha_usu_nombre), "", rsaux3!vcha_usu_nombre) + " " + IIf(IsNull(rsaux3!vcha_usu_apellidos), "", rsaux3!vcha_usu_apellidos)
                        Else
                        End If
                        rsaux3.Close
                        var_clave_Destino = rsaux!attribute1
                        var_nombre_destino = rsaux!nombre_almacen
                        rsaux.Close
                        rsaux.Open "select * from xxvia_Tb_Salidas_cajas a, xxvia_vw_categorias_item_b b where source_header_number = " + Me.txt_embraue_nota_envio + " and inte_emb_embarque = " + Me.txt_embarque_pedido + " and a.segment1 = codigo and organization_id = " + CStr(var_unidad_organizacional) + " and floa_Sal_Cantidad_leida > 0", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        While Not rsaux.EOF
                              If var_consignacion = "PTO_CONS2" Then
                                 If rsaux1.State = 1 Then
                                    rsaux1.Close
                                 End If
                                 'rsaux1.Open "SELECT OPERAND FROM qp_list_lines_v WHERE list_header_id =  9007 AND  PRODUCT_ATTR_VAL_DISP = '" + rsaux!ORDERED_ITEM + "' and start_date_active <= to_date('" + CStr(VAR_FECHA_PRECIO) + "','DD/MM/YYYY') and (end_date_active is null or end_date_active >= to_date('" + CStr(VAR_FECHA_PRECIO) + "','DD/MM/YYYY')) and Product_Attr_Value = " + CStr(rsaux!inventory_item_id), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 rsaux1.Open "SELECT OPERAND FROM qp_list_lines_v WHERE list_header_id =  9007 and start_date_active <= to_date('" + CStr(VAR_FECHA_PRECIO) + "','DD/MM/YYYY') and (end_date_active is null or end_date_active >= to_date('" + CStr(VAR_FECHA_PRECIO) + "','DD/MM/YYYY')) and Product_Attr_Value = " + CStr(rsaux!inventory_item_id), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 If Not rsaux1.EOF Then
                                    'var_precio = IIf(IsNull(rsaux1!OPERAND), rsaux!unit_selling_price, rsaux1!OPERAND) / 1.16
                                    var_precio = rsaux1!OPERAND / 1.16
                                 Else
                                    var_precio = 0
                                 End If
                                 rsaux1.Close
                              Else
                                 'var_precio = rsaux!unit_selling_price
                                 var_precio = 0
                              End If
                              var_cadena = "INSERT INTO tb_Temp_oracle_NOTA_ENVIO (INTE_TEM_CONSECUTIVO, PEDIDO,                  FECHA,                 ALMACEN,                       CLIENTE,            NOMBRE_CLIENTE,                                            EMBARQUE,                 LINEA, NOMBRE_LINEA,            CODIGO,                     NOMBRE_ARTICULO,              ENTREGO, INICIO, TERMINO, DIRECCION, COLONIA, CIUDAD, ESTADO, PAIS, CP, CANTIDAD, PRECIO, INTE_EMB_EMBARQUE) VALUES "
                              'var_cadena = var_cadena + " (" + CStr(var_consecutivo) + ", " + CStr(rsaux!order_number) + ",'" + VAR_FECHA_MOVIMIENTO + "', '" + var_nombre_almacen + "','" + CStr(rsaux!attribute1) + "', '" + rsaux!nombre_almacen + "'," + Me.txt_embraue_nota_envio + ",'','" + rsaux!Linea + "','" + rsaux!ORDERED_ITEM + "', '" + rsaux!Description + "','" + VAR_NOMBRE_USUARIO_ENTREGO + "','" + CStr(VAR_FECHA_PRECIO) + "','" + CStr(VAR_FECHA_MOVIMIENTO) + "','" + VAR_DIRECCION + "','" + VAR_COLONIA + "','" + var_ciudad + "','" + var_estado + "','" + var_pais + "','" + VAR_CP + "'," + CStr(rsaux!CANTIDAD_SURTIDA) + "," + CStr(var_precio) + "," + CStr(var_numero_embarque) + ")"
                              var_cadena = var_cadena + " (" + CStr(var_consecutivo) + ", " + CStr(rsaux!source_header_number) + ",'" + VAR_FECHA_MOVIMIENTO + "', '" + var_nombre_almacen + "','" + CStr(var_clave_Destino) + "', '" + var_nombre_destino + "'," + Me.txt_embraue_nota_envio + ",'','" + rsaux!Linea + "','" + rsaux!SEGMENT1 + "', '" + rsaux!Descripcion + "','" + VAR_NOMBRE_USUARIO_ENTREGO + "','" + CStr(VAR_FECHA_PRECIO) + "','" + CStr(VAR_FECHA_MOVIMIENTO) + "','" + VAR_DIRECCION + "','" + VAR_COLONIA + "','" + var_ciudad + "','" + var_estado + "','" + var_pais + "','" + VAR_CP + "'," + CStr(rsaux!FLOA_SAL_CANTIDAD_LEIDA) + "," + CStr(var_precio) + "," + CStr(var_numero_embarque) + ")"
                              'var_nombre_almacen_consignacioN = rsaux!nombre_almacen
                              var_nombre_almacen_consignacioN = ""
                              'MsgBox var_cadena
                              If rsaux1.State = 1 Then
                                 rsaux1.Close
                              End If
                              rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              rsaux.MoveNext
                        Wend
                        rsaux1.Open "DELETE FROM tb_Temp_oracle_NOTA_ENVIO WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO IS NULL", cnn, adOpenDynamic, adLockOptimistic
                        rsaux.MoveFirst
                        'If rsaux!order_number = 125410 Then
                        '   rsaux1.Open "select segment1, sum(floa_Sal_Cantidad_leida) as cantidad from xxvia_tb_salidas_cajas where source_header_number = 125410 group by segment1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        '   While Not rsaux1.EOF
                        '         rsaux9.Open "update tb_Temp_oracle_NOTA_ENVIO set cantidad = " + CStr(rsaux1!cantidad) + " where pedido = 125410 and inte_tem_consecutivo = " + CStr(var_consecutivo) + " and codigo = '" + rsaux1!segment1 + "'", cnn, adOpenDynamic, adLockOptimistic
                        '         rsaux1.MoveNext
                        '   Wend
                        '   rsaux1.Close
                        'End If
                        rsaux1.Open "select sum(cantidad) from tb_Temp_oracle_NOTA_ENVIO where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux1.EOF Then
                           var_cantidad_oracle = IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value)
                        Else
                          var_cantidad_oracle = 0
                        End If
                        rsaux1.Close
                        rsaux1.Open "select sum(floa_sal_Cantidad_leida) from xxvia_tb_salidas where source_header_number = " + CStr(Me.txt_embraue_nota_envio) + " and inte_emb_embarque = " + Me.txt_embarque_pedido, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux1.EOF Then
                           var_cantidad_leida = IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value)
                        Else
                           var_cantidad_leida = 0
                        End If
                        rsaux1.Close
                        rsaux1.Open "select sum(floa_sal_Cantidad_leida) from xxvia_tb_salidas_cajas where source_header_number = " + CStr(Me.txt_embraue_nota_envio) + " and inte_emb_embarque = " + Me.txt_embarque_pedido, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux1.EOF Then
                           var_cantidad_leida = var_cantidad_leida + IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value)
                        Else
                           var_cantidad_leida = var_cantidad_leida + 0
                        End If
                        rsaux1.Close
                        'If Round(var_cantidad_leida, 2) = Round(var_cantidad_leida, 2) Then
                        If Round(var_cantidad_leida, 2) <= Round(var_cantidad_oracle, 2) Then
                           strconsulta = "select nvl(attribute7,'N') activa_pos_icg from mtl_secondary_inventories where secondary_inventory_name=?"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_almacen_icg)
                                .Parameters.Append parametro
                           End With
                           Set rsaux9 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           If Not rsaux9.EOF Then
                              var_posible_cgi = IIf(IsNull(rsaux9(0).Value), "N", rsaux9(0))
                           Else
                              var_posible_cgi = "N"
                           End If
                           rsaux9.Close
                           var_posible_cgi = "N"
                           rsaux9.Open "select * from TB_ORACLE_NOTAS_IMPRESAS_ICG where pedido = " + Me.txt_embraue_nota_envio, cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux9.EOF Then
                              var_posible_cgi = "N"
                           End If
                           rsaux9.Close
                           var_posible_cgi = "N"
                           If var_posible_cgi = "Y" Then
                              If CDbl(Me.txt_embraue_nota_envio) >= 161037 Then
                                 If cnnicg_sql.State = 1 Then
                                    cnnicg_sql.Close
                                 End If
                                 cnnicg_sql.Open "Provider=SQLOLEDB.1;Password=icgfront2013;Persist Security Info=True;User ID=ICGAdmin;Initial Catalog=general;Data Source=sqlposprod"
                                 rsaux1.Open "SELECT source_header_number, inte_paq_caja, segment1, sum(floa_sal_Cantidad_leida) as FLOA_SAL_CANTIDAD_LEIDA FROM XXVIA_TB_SALIDAS_CAJAS WHERE SOURCE_HEADER_NUMBER = " + Me.txt_embraue_nota_envio + " AND FLOA_SAL_CANTIDAD_LEIDA >0 group by source_header_number, inte_paq_caja, segment1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 While Not rsaux1.EOF
                                       'VAR_EMBARQUE_ICG = rsaux1!inte_emb_embarque
                                       VAR_CAJA_ICG = rsaux1!INTE_PAQ_CAJA
                                       If Len(Trim(Str(VAR_CAJA_ICG))) = 1 Then
                                          var_referencia_caja = "00" + Trim(Str(VAR_CAJA_ICG))
                                       End If
                                       If Len(Trim(Str(VAR_CAJA_ICG))) = 2 Then
                                          var_referencia_caja = "0" + Trim(Str(VAR_CAJA_ICG))
                                       End If
                                       If Len(Trim(Str(VAR_CAJA_ICG))) = 3 Then
                                          var_referencia_caja = Trim(Str(VAR_CAJA_ICG))
                                       End If
                                       VAR_CAJA_S = var_referencia_caja
                                       var_dia_s = CStr(Day(Now))
                                       var_mes_s = CStr(Month(Now))
                                       var_año_s = CStr(Year(Now))
                                       If Len(var_dia_s) = 1 Then
                                          var_dia_s = "0" + var_dia_s
                                       End If
                                       If Len(var_mes_s) = 1 Then
                                          var_mes_s = "0" + var_mes_s
                                       End If
                                       If Len(var_año_s) = 2 Then
                                          var_año_s = "20" + var_año_s
                                       End If
                                       var_fecha = var_dia_s + "-" + var_mes_s + "-" + var_año_s
                                       strconsulta = "select * from XXVIA_TB_ICG_TRAN_CEDIS_TIENDA where NUMB_ORGANIZATION_ID = ? and  VCHA_SUBINVENTORY_CODE = ? and VCHA_TRANSFER_SUBINVENTORY = ? and VCHA_NOTA_ENVIO = ? and VCHA_NUMERO_CAJA = ? and VCHA_CODIGO = ? "
                                       With comandoORA
                                            .ActiveConnection = cnnicg
                                            .CommandType = adCmdText
                                            .CommandText = strconsulta
                                            Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(var_unidad_organizacional))
                                            .Parameters.Append parametro
                                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_almacen)
                                            .Parameters.Append parametro
                                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_almacen_icg)
                                            .Parameters.Append parametro
                                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, CStr(rsaux1!source_header_number))
                                            .Parameters.Append parametro
                                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, VAR_CAJA_S)
                                            .Parameters.Append parametro
                                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rsaux1!SEGMENT1)
                                            .Parameters.Append parametro
                                       End With
                                       Set rsaux9 = comandoORA.execute
                                       Set comandoORA = Nothing
                                       Set parametro = Nothing
                                       If rsaux9.EOF Then
                                          strconsulta = "INSERT INTO XXVIA_TB_ICG_TRAN_CEDIS_TIENDA (NUMB_ORGANIZATION_ID, VCHA_SUBINVENTORY_CODE, VCHA_TRANSFER_SUBINVENTORY, DATE_FECHA, VCHA_NOTA_ENVIO, VCHA_NUMERO_CAJA, VCHA_CODIGO, NUMB_CANTIDAD, NUMB_STATUS, NUMB_ORG_ORIGEN) VALUES (?, ?, ?, SYSDATE, ?, ?, ?, ?, 3, ?)"
                                          With comandoORA
                                               .ActiveConnection = cnnicg
                                               .CommandType = adCmdText
                                               .CommandText = strconsulta
                                               Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(var_unidad_organizacional))
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_almacen)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_almacen_icg)
                                               .Parameters.Append parametro
   '                                           Set parametro = .CreateParameter(, adDate, adParamInput, 200, CDate(var_fecha))
   '                                           .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, CStr(rsaux1!source_header_number))
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, VAR_CAJA_S)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rsaux1!SEGMENT1)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, rsaux1!FLOA_SAL_CANTIDAD_LEIDA)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(var_unidad_organizacional))
                                               .Parameters.Append parametro
                                          End With
                                          Set rsaux8 = comandoORA.execute
                                          Set comandoORA = Nothing
                                          Set parametro = Nothing
                                       End If
                                       rsaux9.Close
                                       rsaux1.MoveNext
                                 Wend
                                 rsaux1.MoveFirst
                                 strconsulta = "UPDATE XXVIA_TB_ICG_TRAN_CEDIS_TIENDA SET NUMB_STATUS = 0 where NUMB_ORGANIZATION_ID = ? and  VCHA_SUBINVENTORY_CODE = ? and VCHA_TRANSFER_SUBINVENTORY = ? and VCHA_NOTA_ENVIO = ? and NUMB_STATUS = 3"
                                 With comandoORA
                                      .ActiveConnection = cnnicg
                                      .CommandType = adCmdText
                                      .CommandText = strconsulta
                                      Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, CDbl(var_unidad_organizacional))
                                      .Parameters.Append parametro
                                      Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_almacen)
                                      .Parameters.Append parametro
                                      Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_almacen_icg)
                                      .Parameters.Append parametro
                                      Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, CStr(rsaux1!source_header_number))
                                      .Parameters.Append parametro
                                 End With
                                 Set rsaux9 = comandoORA.execute
                                 Set comandoORA = Nothing
                                 Set parametro = Nothing
                           
''''''''''''   '''''''''''' comienza traspaso de costales
                                 x = 1
                                 If x = 1 Then
                                    var_posible_pedido = 1
                                    var_pedido_tienda = Me.txt_embraue_nota_envio
                                    If rsaux8.State = 1 Then
                                       rsaux8.Close
                                    End If
                                    rsaux8.Open "SELECT * FROM TB_ORACLE_PEDIDOS_TIENDAS_COSTALES WHERE PEDIDO = " + CStr(var_pedido_tienda), cnn, adOpenDynamic, adLockOptimistic
                                    If Not rsaux8.EOF Then
                                       var_posible_pedido = 0
                                    End If
                                    rsaux8.Close
                                    If var_posible_pedido = 1 Then
                                       strconsulta = "SELECT A.ATTRIBUTE1, B.description FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B,  OE_ORDER_HEADERS_ALL OHA Where requisition_header_id = OHA.SOURCE_DOCUMENT_ID AND secondary_inventory_name = A.ATTRIBUTE1 AND OHA.ORDER_NUMBER = ?"
                                       With comandoORA
                                            .ActiveConnection = cnnoracle_4
                                            .CommandType = adCmdText
                                            .CommandText = strconsulta
                                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_pedido_tienda)
                                            .Parameters.Append parametro
                                       End With
                                       Set rsaux8 = comandoORA.execute
                                       Set comandoORA = Nothing
                                       Set parametro = Nothing
                                       var_almacen_tienda = IIf(IsNull(rsaux8!attribute1), "", rsaux8!attribute1)
                                       p_almacendestinofinal = var_almacen_tienda
                                       rsaux8.Close
                                       If var_almacen_tienda <> "" Then
                                          var_i = 0
                                          rsaux8.Open "SELECT XXVIA_SQ_LINEA_TM.nextval FROM dual", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                          If Not rsaux8.EOF Then
                                             p_origenencabezadoid = rsaux8(0).Value
                                          End If
                                          rsaux8.Close
                                          rsaux8.Open "select XXVIA_SQ_ENCABEZADO_MT_ID.nextval from dual", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                          If Not rsaux8.EOF Then
                                             P_ENCABEZADO_MT_ID = rsaux8(0).Value
                                          End If
                                          rsaux8.Close
                                          strconsulta = "SELECT TIPO_CAJA, COUNT(*) AS CANTIDAD FROM XXVIA_VW_CAJAS_POR_PEDIDO WHERE SOURCE_HEADER_NUMBER = ? AND (TIPO_CAJA LIKE '%COSTAL%' OR TIPO_CAJA LIKE '%CAJA BIASI%') and INTE_EMB_EMBARQUE = " + Me.txt_embarque_pedido + " GROUP BY TIPO_CAJA"
                                          With comandoORA
                                               .ActiveConnection = cnnoracle_4
                                               .CommandType = adCmdText
                                               .CommandText = strconsulta
                                               Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, Me.txt_embraue_nota_envio)
                                               .Parameters.Append parametro
                                          End With
                                          Set rsaux6 = comandoORA.execute
                                          Set comandoORA = Nothing
                                          Set parametro = Nothing
                                          While Not rsaux6.EOF
                                                var_i = var_i + 1
                                                rs.Open "select * from tb_oracle_empaques where empaque = '" + rsaux6!tipo_caja + "' and codigo is not null", cnn, adOpenDynamic, adLockOptimistic
                                                If Not rs.EOF Then
                                                   strconsulta = "select PRIMARY_UOM_CODE, INVENTORY_ITEM_ID from xxvia_system_items_b where SEGMENT1 = ? AND ORGANIZATION_ID = ?"
                                                   With comandoORA
                                                        .ActiveConnection = cnnoracle_4
                                                        .CommandType = adCmdText
                                                        .CommandText = strconsulta
                                                        Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, rs!codigo)
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, var_unidad_organizacional)
                                                        .Parameters.Append parametro
                                                   End With
                                                   Set rsaux8 = comandoORA.execute
                                                   Set comandoORA = Nothing
                                                   Set parametro = Nothing
                                                   var_inventory_item_id = rsaux8!inventory_item_id
                                                   p_um = rsaux8!PRIMARY_UOM_CODE
                                                   rsaux8.Close
                                                   p_organizacion_id = var_unidad_organizacional
                                                   p_organizacion_destino = var_unidad_organizacional
                                                   If var_empresa = 92 Then
                                                      p_subinventario = "CDI_ALMPT"
                                                   End If
                                                   If var_empresa = 83 Then
                                                      p_subinventario = "TEX_PT_QL"
                                                   End If
                                                   p_subinventario_destino = "TRANS"
                                                   p_codigoarticulo = rs!codigo
                                                   p_cantidadorigen = rsaux6!cantidad
                                                   p_Cantidadrecibida = 0
                                                   p_origentransaccion = "SID_COSTALES_" + CStr(var_pedido_tienda)
                                                   p_referencia_transaccion = var_pedido_tienda
                                                   p_mensajeerror = ""
                                                   strconsulta = "call xxvia_pk_inventarios.xxvia_sp_inventarios4 (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
                                                   With comandoORA
                                                        .ActiveConnection = cnnoracle_4
                                                        .CommandType = adCmdText
                                                        .CommandText = strconsulta
                                                        Set parametro = .CreateParameter(, adInteger, adParamInput, 100, p_organizacion_id)
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adInteger, adParamInput, 100, p_organizacion_destino)
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, p_subinventario)
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, p_subinventario_destino)
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adInteger, adParamInput, 100, 2)
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, p_codigoarticulo)
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adInteger, adParamInput, 100, p_cantidadorigen)
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adInteger, adParamInput, 100, p_cantidadorigen)
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, p_origentransaccion)
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adInteger, adParamInput, 100, p_origenencabezadoid)
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adDouble, adParamInput, 100, Null)
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adDouble, adParamInput, 100, Null)
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, p_referencia_transaccion)
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adDouble, adParamInput, 100, Null)
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, p_almacendestinofinal)
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adDate, adParamInput, 100, Date)
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, Null)
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, p_um)
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adDouble, adParamInput, 100, P_ENCABEZADO_MT_ID)
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_clave_usuario_global)
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, fun_NombrePc)
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, Null)
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, "")
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, "")
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adVarChar, adParamOutput, 100, p_mensajeerror)
                                                        .Parameters.Append parametro
                                                   End With
                                                   'MsgBox strconsulta
                                                   Set rsaux9 = comandoORA.execute
                                                   Set comandoORA = Nothing
                                                   Set parametro = Nothing
                                                End If
                                                rs.Close
                                                rsaux6.MoveNext
                                          Wend
                                          strconsulta = "call xxvia_pk_inventarios.xxvia_valida_interface (1,?,?)"
                                          With comandoORA
                                               .ActiveConnection = cnnoracle_4
                                               .CommandType = adCmdText
                                               .CommandText = strconsulta
                                               Set parametro = .CreateParameter(, adDouble, adParamInput, 100, p_origenencabezadoid)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adDouble, adParamInput, 200, 0)
                                               .Parameters.Append parametro
                                          End With
                                          rsaux9.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                          On Error GoTo salir2
                                          Set rsaux9 = comandoORA.execute
                                          Set comandoORA = Nothing
                                          Set parametro = Nothing
                                          rsaux8.Open "insert into tb_oracle_pedidos_tiendas_costales (pedido) values (" + CStr(var_pedido_tienda) + ")", cnn, adOpenDynamic, adLockOptimistic
                                       End If
                                    End If
                                 End If
   ''''''''''''''''''''''''  fin de traspaso de costales
                                 If rsaux6.State = 1 Then
                                    rsaux6.Close
                                 End If
                                 rsaux10.Open "SELECT * FROM TB_ORACLE_NOTAS_IMPRESAS_ICG WHERE PEDIDO = " + Me.txt_embraue_nota_envio, cnn, adOpenDynamic, adLockOptimistic
                                 If rsaux10.EOF Then
                                    'MsgBox cnnicg_sql.ConnectionString
                                    cnn.CommandTimeout = 360
                                    rsaux11.Open "INSERT INTO TB_ORACLE_NOTAS_IMPRESAS_ICG (PEDIDO, FECHA) VALUES (" + Me.txt_embraue_nota_envio + ",GETDATE())", cnn, adOpenDynamic, adLockOptimistic
                                    x = 1
                                    If x = 1 Then
                                       If cnnicg_sql.State = 1 Then
                                          cnnicg_sql.Close
                                       End If
                                       cnnicg_sql.Open "Provider=SQLOLEDB.1;Password=icgfront2013;Persist Security Info=True;User ID=ICGAdmin;Initial Catalog=general;Data Source=sqlposprod"
                                       rsaux9.Open "exec vyt_crea_pedido_cedis " + var_unidad_organizacional + ", '" + CStr(Me.txt_embraue_nota_envio) + "'", cnnicg_sql, adOpenDynamic, adLockOptimistic
                                       rsaux9.Open "call xxpos.xxvia_pk_motor_logistico.xxvia_sp_senales_eviandas_a_cn (" + Me.txt_embraue_nota_envio + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                    End If
                                 End If
                                 If rsaux10.State = 1 Then
                                    rsaux10.Close
                                 End If
                                 If rsaux1.State = 1 Then
                                    rsaux1.Close
                                 End If
                                 If cnnicg_sql.State = 1 Then
                                    cnnicg_sql.Close
                                 End If
                              End If
                           End If
                           rsaux11.Open "INSERT INTO TB_ORACLE_NOTAS_IMPRESAS_ICG (PEDIDO, FECHA) VALUES (" + Me.txt_embraue_nota_envio + ",GETDATE())", cnn, adOpenDynamic, adLockOptimistic
                           'error aqui
                           strconsulta = "select nvl(attribute7,'N') activa_pos_icg from mtl_secondary_inventories where secondary_inventory_name=?"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_almacen_icg)
                                .Parameters.Append parametro
                           End With
                           Set rsaux9 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           If Not rsaux9.EOF Then
                              var_posible_cgi = IIf(IsNull(rsaux9(0).Value), "N", rsaux9(0))
                           Else
                              var_posible_cgi = "N"
                           End If
                           rsaux9.Close
                           var_posible_cgi = "Y"
                           If var_posible_cgi = "Y" Then
                              x = 0
                              If x = 1 Then
                                 var_conexion_string_p = "Provider=SQLOLEDB.1;Password=icgfront2013;Persist Security Info=True;User ID=ICGAdmin;Initial Catalog=general;Data Source=sqlposprod"
                                 If cnn_icg_posprod.State = 1 Then
                                    cnn_icg_posprod.Close
                                 End If
                                 cnn_icg_posprod.Open var_conexion_string_p
                                 cnn_icg_posprod.CommandTimeout = 360
                                 'rsaux9.Open "exec [sqlposprod.vianney.com.mx].general.dbo.vyt_crea_pedido_cedis " + var_unidad_organizacional + ", '" + CStr(Me.txt_embraue_nota_envio) + "'", cnn, adOpenDynamic, adLockOptimistic
                                 'MsgBox cnn_icg_posprod.ConnectionString
                                 'rsaux9.Open "exec vyt_crea_pedido_cedis " + var_unidad_organizacional + ", '" + CStr(Me.txt_embraue_nota_envio) + "'", cnn_icg_posprod, adOpenDynamic, adLockOptimistic
                              End If
                           End If
                           rsaux1.Open "DELETE FROM TB_ORACLE_CAJAS_EMBARQUES WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                           rsaux1.Open "select  inte_paq_caja, tipo_Caja, MAX(sello)  as cantidad, MAX(NVL(TRANSPORTE,' ')) AS TRANSPORTE from xxvia_tb_salidas_Cajas where source_header_number = " + Me.txt_embraue_nota_envio + " and inte_emb_embarque = " + Me.txt_embarque_pedido + " and floa_sal_cantidad_leida > 0 GROUP BY inte_paq_caja, tipo_Caja order by inte_paq_caja ", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           While Not rsaux1.EOF
                                 If rsaux2.State = 1 Then
                                    rsaux2.Close
                                 End If
                                 rsaux2.Open "INSERT INTO TB_ORACLE_CAJAS_EMBARQUES (INTE_TEM_CONSECUTIVO, PEDIDO, CAJA, TIPO_CAJA, SELLO, TRANSPORTE) VALUES (" + CStr(var_consecutivo) + "," + Me.txt_embraue_nota_envio + "," + CStr(rsaux1!INTE_PAQ_CAJA) + ",'" + IIf(IsNull(rsaux1!tipo_caja), "", rsaux1!tipo_caja) + "','" + IIf(IsNull(rsaux1!cantidad), "", rsaux1!cantidad) + "','" + rsaux1!transporte + "')", cnn, adOpenDynamic, adLockOptimistic
                                 rsaux1.MoveNext
                           Wend
                           rsaux1.Close
                           x = 0
                           If x = 0 Then
                              'select a.description, ADDRESS_LINE_1, ADDRESS_LINE_2, TOWN_OR_CITY, REGION_1, POSTAL_CODE  into lv_nombre_subinventario, lv_direccion_1, lv_direccion_2 , lv_ciudad, lv_Estado, lv_cp  from mtl_secondary_inventories A, hr_locations_all B where A.ORGANIZATION_ID = LV_ORGANIZACION_DESTINO AND A.secondary_inventory_name = var_destino AND A.LOCATION_ID = B.LOCATION_ID;
                              If var_location_id > 0 Then
                              
                              
                              
                                 'strconsulta = "call XXVIA_SP_TIMBRAR_TRASPASOS(?,?,?)"
                                 strconsulta = "call XXVIA_SP_TIMBRAR_TRASPASOS_2(?,?,?,?)"
                                 'strconsulta = "call XXVIA_SP_TIMBRAR_TRASPASOS_3(?,?,?,?)"
                                 With comandoORA
                                      .ActiveConnection = cnnoracle_4
                                      .CommandType = adCmdText
                                      .CommandText = strconsulta
                                      Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, Me.txt_embraue_nota_envio)
                                      .Parameters.Append parametro
                                      Set parametro = .CreateParameter(, adNumeric, adParamInput, 50, 3)
                                      .Parameters.Append parametro
                                      Set parametro = .CreateParameter(, adNumeric, adParamInput, 50, var_unidad_organizacional)
                                      .Parameters.Append parametro
                                      Set parametro = .CreateParameter(, adNumeric, adParamInput, 50, var_numero_embarque)
                                      .Parameters.Append parametro
                                 End With
                                 Set rsaux2 = comandoORA.execute
                                 Set comandoORA = Nothing
                                 Set parametro = Nothing
                                 var_serie = "TRX" + Me.txt_embarque_pedido + "_"
                                 strconsulta = "select customer_trx_id, cadena as cadena, numero from xxvia_tb_control_doc_fiscales where serie = '" + var_serie + "' and numero = ? AND ORGANIZACION = ?"
                                 With comandoORA
                                      .ActiveConnection = cnnoracle_4
                                      .CommandType = adCmdText
                                      .CommandText = strconsulta
                                      Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, Me.txt_embraue_nota_envio)
                                      .Parameters.Append parametro
                                      Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, var_unidad_organizacional)
                                      .Parameters.Append parametro
                                 End With
                                 Set rsaux2 = comandoORA.execute
                                 Set comandoORA = Nothing
                                 Set parametro = Nothing
                                 If Not rsaux2.EOF Then
                                    var_cadena = Replace(rsaux2!Cadena, " ", "")
                                    var_cadena_rfc = Mid(var_cadena, 34, 12)
                                    VAR_CADENA_STR = ""
                                    Open ("C:\SISTEMAS\TRX" + Trim(Me.txt_embarque_pedido) + "_" + Trim(Me.txt_embraue_nota_envio) + ".FAC") For Output As #1
                                    For var_i = 1 To Len(var_cadena)
                                        If Asc(Mid(var_cadena, var_i, 1)) = 63 Then
                                           Print #1, VAR_CADENA_STR
                                           VAR_CADENA_STR = ""
                                        Else
                                           VAR_CADENA_STR = VAR_CADENA_STR + Mid(var_cadena, var_i, 1)
                                        End If
                                    Next var_i
                                    Print #1, "FIN:"
                                    Close #1
                                    var_archivo = "C:\SISTEMAS\sube_fact" + Trim(Str(Me.txt_embarque_pedido)) + "_" + Trim(Me.txt_embraue_nota_envio) + ".bat"
                                    x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\TRX" + Me.txt_embarque_pedido + "_" + Trim(Me.txt_embraue_nota_envio) + ".FAC" + "|https://facturas2.vianney.mx/cgi-bin/cfds/timbrarGR33|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
                                    rsaux2.Close
                                    rsaux1.Open "select *  from tb_oracle_tiempo_impresion_documentos where pedido = " + Me.txt_embraue_nota_envio, cnn, adOpenDynamic, adLockOptimistic
                                    If rsaux1.EOF Then
                                       strconsulta = "select oha.source_document_id, A.ATTRIBUTE1, B.description  from oe_order_headers_all oha, po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B  where order_number = ? and requisition_header_id = oha.source_document_id and secondary_inventory_name = A.ATTRIBUTE1"
                                       With comandoORA
                                            .ActiveConnection = cnnoracle_4
                                            .CommandType = adCmdText
                                            .CommandText = strconsulta
                                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, Me.txt_embraue_nota_envio)
                                            .Parameters.Append parametro
                                       End With
                                       Set rsaux2 = comandoORA.execute
                                       Set comandoORA = Nothing
                                       Set parametro = Nothing
                                       If Not rsaux2.EOF Then
                                          var_clave_almacen = rsaux2!attribute1
                                          var_nombre_almacen = rsaux2!Description
                                       Else
                                          var_clave_almacen = ""
                                          var_nombre_almacen = ""
                                       End If
                                       rsaux2.Close
                                       rsaux2.Open "insert into tb_oracle_tiempo_impresion_documentos (pedido,fecha, tienda, nombre) values (" + Me.txt_embraue_nota_envio + ",GETDATE(),'" + var_clave_almacen + "','" + var_nombre_almacen + "')", cnn, adOpenDynamic, adLockOptimistic
                                    End If
                                    rsaux1.Close
                                 Else
                                 End If
                                 'rsaux2.Close
                              Else
                                 MsgBox "El subinventario " + var_almacen_icg + "   " + var_nombre_almacen_consignacioN + ", no tiene una dirección asignada, favor de validarlo con el departamento de costos o contraloria", vbOKOnly, "ATENCION"
                              End If
                           'Else
                              rsaux1.Open "select *  from tb_oracle_tiempo_impresion_documentos where pedido = " + Me.txt_embraue_nota_envio, cnn, adOpenDynamic, adLockOptimistic
                              If rsaux1.EOF Then
                                 strconsulta = "select oha.source_document_id, A.ATTRIBUTE1, B.description  from oe_order_headers_all oha, po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B  where order_number = ? and requisition_header_id = oha.source_document_id and secondary_inventory_name = A.ATTRIBUTE1"
                                 With comandoORA
                                      .ActiveConnection = cnnoracle_4
                                      .CommandType = adCmdText
                                      .CommandText = strconsulta
                                      Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, Me.txt_embraue_nota_envio)
                                      .Parameters.Append parametro
                                 End With
                                 Set rsaux2 = comandoORA.execute
                                 Set comandoORA = Nothing
                                 Set parametro = Nothing
                                 If Not rsaux2.EOF Then
                                    var_clave_almacen = rsaux2!attribute1
                                    var_nombre_almacen = rsaux2!Description
                                 Else
                                    var_clave_almacen = ""
                                    var_nombre_almacen = ""
                                 End If
                                 rsaux2.Close
                                 rsaux2.Open "insert into tb_oracle_tiempo_impresion_documentos (pedido,fecha, tienda, nombre) values (" + Me.txt_embraue_nota_envio + ",GETDATE(),'" + var_clave_almacen + "','" + var_nombre_almacen + "')", cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux1.Close
                              rsaux1.Open "select distinct pedido from tb_Temp_oracle_NOTA_ENVIO where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                              While Not rsaux1.EOF
                                    Set reporte = appl.OpenReport(App.Path + "\rep_oracle_nota_envio_linea.rpt")
                                    reporte.RecordSelectionFormula = "{VW_ORACLE_NOTA_ENVIO_LINEA.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_ORACLE_NOTA_ENVIO_LINEA.cantidad} > 0 and {VW_ORACLE_NOTA_ENVIO_LINEA.pedido} = '" + CStr(rsaux1!pedido) + "'"
                                    frmvistasprevias.cr.ReportSource = reporte
                                    For ntablas = 1 To reporte.Database.Tables.Count
                                        reporte.Database.Tables(ntablas).SetLogOnInfo "admcdindustrial", var_bd_reportes, parametros(4), parametros(5)
                                    Next ntablas
                                    frmvistasprevias.cr.ViewReport
                                    frmvistasprevias.Caption = "Nota de envio a tiendas"
                                    frmvistasprevias.Show 1
                                    Set reporte = Nothing
                                    var_si = MsgBox("¿Desea el reporte a detalle?", vbYesNo, "ATENCION")
                                    If var_si = 6 Then
                                       If var_consignacion = "PTO_CONS2" Then
                                          Set reporte = appl.OpenReport(App.Path + "\rep_oracle_nota_envio_detalle_consig.rpt")
                                          reporte.RecordSelectionFormula = "{VW_ORACLE_NOTA_ENVIO_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_ORACLE_NOTA_ENVIO_DETALLE.cantidad} > 0 and {VW_ORACLE_NOTA_ENVIO_DETALLE.pedido} = '" + CStr(rsaux1!pedido) + "'"
                                          frmvistasprevias.cr.ReportSource = reporte
                                          For ntablas = 1 To reporte.Database.Tables.Count
                                              reporte.Database.Tables(ntablas).SetLogOnInfo "admcdindustrial", var_bd_reportes, parametros(4), parametros(5)
                                          Next ntablas
                                          frmvistasprevias.cr.ViewReport
                                          frmvistasprevias.Caption = "Nota de envio a tiendas"
                                          frmvistasprevias.Show 1
                                          Set reporte = Nothing
                                          Set reporte = appl.OpenReport(App.Path + "\rep_oracle_nota_envio_detalle_consig.rpt")
                                          reporte.RecordSelectionFormula = "{VW_ORACLE_NOTA_ENVIO_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_ORACLE_NOTA_ENVIO_DETALLE.cantidad} > 0 and {VW_ORACLE_NOTA_ENVIO_DETALLE.pedido} = '" + CStr(rsaux1!pedido) + "'"
                                          For ntablas = 1 To reporte.Database.Tables.Count
                                              reporte.Database.Tables(ntablas).SetLogOnInfo "admcdindustrial", var_bd_reportes, parametros(4), parametros(5)
                                          Next ntablas
                                          reporte.ExportOptions.FormatType = crEFTExcel80
                                          reporte.ExportOptions.DestinationType = crEDTDiskFile
                                          archivo = "c:\reportessid\NOTA_ENVIO_" + Me.txt_embraue_nota_envio + "_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + ".xls"
                                          reporte.ExportOptions.DiskFileName = archivo
                                          reporte.Export False
                                          Set reporte = Nothing
                                          var_si = MsgBox("¿Desea enviar la nota por correo?", vbYesNo, "ATENCION")
                                          If var_si = 6 Then
                                             VAR_CORREO_ELECTRONICO = "vluna@vianney.com.mx"
                                             If Trim(VAR_CORREO_ELECTRONICO) <> "" Then
                                                If MAPISession1.SessionID = 0 Then
                                                   MAPISession1.SignOn
                                                End If
                                                MAPIMessages1.SessionID = MAPISession1.SessionID
                                                MAPIMessages1.Compose
                                                MAPIMessages1.RecipDisplayName = VAR_CORREO_ELECTRONICO
                                                MAPIMessages1.RecipAddress = VAR_CORREO_ELECTRONICO
                                                MAPIMessages1.AddressResolveUI = True
                                                MAPIMessages1.ResolveName
                                                MAPIMessages1.MsgSubject = "Nota de envio " + Str(Me.txt_embraue_nota_envio)
                                                MAPIMessages1.MsgNoteText = "Se adjunta nota de envio número " + Str(Me.txt_embraue_nota_envio) + " del cliente " + var_nombre_almacen_consignacioN
                                                MAPIMessages1.AttachmentPathName = archivo
                                                MAPIMessages1.send True
                                                If MAPISession1.SessionID > 0 Then
                                                    MAPISession1.SignOff
                                                End If
                                             End If
                                          End If
                                          MsgBox "Se a terminado de guardar el archivo " + archivo
                                       Else
                                          Set reporte = appl.OpenReport(App.Path + "\rep_oracle_nota_envio_detalle.rpt")
                                          reporte.RecordSelectionFormula = "{VW_ORACLE_NOTA_ENVIO_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {VW_ORACLE_NOTA_ENVIO_DETALLE.cantidad} > 0 and {VW_ORACLE_NOTA_ENVIO_DETALLE.pedido} = '" + CStr(rsaux1!pedido) + "'"
                                          frmvistasprevias.cr.ReportSource = reporte
                                          For ntablas = 1 To reporte.Database.Tables.Count
                                              reporte.Database.Tables(ntablas).SetLogOnInfo "admcdindustrial", var_bd_reportes, parametros(4), parametros(5)
                                          Next ntablas
                                          frmvistasprevias.cr.ViewReport
                                          frmvistasprevias.Caption = "Nota de envio a tiendas"
                                          frmvistasprevias.Show 1
                                          Set reporte = Nothing
                                       End If
                                    End If
                                    rsaux1.MoveNext
                              Wend
                              rsaux1.Close
                           End If
                        Else
                           MsgBox "No se a terminado de procesar la información en oracle, vuelva a intentar", vbOKOnly, "ATENCION"
                        End If
                        rsaux1.Open "delete from tb_Temp_oracle_NOTA_ENVIO where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                        'rsaux.Close
                        Me.frm_embarque_nota_envio.Visible = False
                     Else
                        MsgBox "No existen pedidos para el embarque", vbOKOnly, "ATENCION"
                     End If
                     rsaux.Close
                  Else
                     MsgBox "No se han generado todos los pedidos", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "El pedido no existe o no corresponde al embarque seleccionado", vbOKOnly, "ATENCION"
               End If
            Else
               If var_posible_embarque = 0 Then
                  MsgBox "El pedido no ha sido cerrado", vbOKOnly, "ATENCION"
               End If
               If var_posible_embarque = 2 Then
                  MsgBox "No se ha asignado un chofer al embarque", vbOKOnly, "ATENCION"
               End If
               If var_posible_embarque = 3 Then
                  MsgBox "El embarque no tiene un transporte asignado.", vbOKOnly, "ATENCION"
               End If
            End If
         Else
            rsaux.Close
            MsgBox "No existen movimientos para el embarque seleccionado", vbOKOnly, "ATENCION"
         End If
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_embarque_nota_envio.Visible = False
   End If
   Exit Sub
salir2:
   If Err.Number = -2147217900 Then
      If rsaux10.State = 1 Then
         rsaux10.Close
         'MsgBox Err.Description
      End If
      rsaux10.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      rsaux10.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      'MsgBox Err.Description
      Resume
   End If
End Sub
