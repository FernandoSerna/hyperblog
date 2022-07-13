VERSION 5.00
Begin VB.Form frmoracle_imprimir_facturas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imprimir facturas"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      Height          =   870
      Left            =   45
      TabIndex        =   28
      Top             =   2130
      Width           =   7050
      Begin VB.CommandButton Command6 
         Height          =   315
         Left            =   6630
         Picture         =   "frmoracle_imprimir_facturas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   450
         Width           =   345
      End
      Begin VB.CommandButton cmd_imprimir_factura 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6645
         Picture         =   "frmoracle_imprimir_facturas.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Imprimir Movimiento"
         Top             =   450
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox txt_serie_factura 
         Height          =   375
         Left            =   705
         TabIndex        =   30
         Top             =   420
         Width           =   1290
      End
      Begin VB.TextBox txt_de_factura 
         Height          =   375
         Left            =   2805
         TabIndex        =   31
         Top             =   420
         Width           =   1290
      End
      Begin VB.TextBox txt_a_factura 
         Height          =   375
         Left            =   4380
         TabIndex        =   32
         Top             =   420
         Width           =   1290
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   270
         TabIndex        =   36
         Top             =   510
         Width           =   405
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "De:"
         Height          =   195
         Left            =   2370
         TabIndex        =   35
         Top             =   510
         Width           =   255
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "A:"
         Height          =   195
         Left            =   4170
         TabIndex        =   33
         Top             =   510
         Width           =   150
      End
      Begin VB.Label Label13 
         BackColor       =   &H000000C0&
         Caption         =   " Facturas"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   30
         TabIndex        =   29
         Top             =   150
         Width           =   6975
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2175
      Left            =   3600
      TabIndex        =   14
      Top             =   -60
      Width           =   3495
      Begin VB.CommandButton Command5 
         Height          =   315
         Left            =   60
         Picture         =   "frmoracle_imprimir_facturas.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   390
         Width           =   345
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command2"
         Height          =   300
         Left            =   1590
         TabIndex        =   38
         Top             =   420
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   300
         Left            =   1230
         TabIndex        =   37
         Top             =   420
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.TextBox txt_pedido 
         Height          =   375
         Left            =   525
         TabIndex        =   21
         Top             =   825
         Width           =   1665
      End
      Begin VB.TextBox txt_serie_pedido 
         Height          =   375
         Left            =   525
         TabIndex        =   20
         Top             =   1230
         Width           =   1290
      End
      Begin VB.TextBox txt_de_pedido 
         Height          =   375
         Left            =   525
         TabIndex        =   19
         Top             =   1635
         Width           =   1290
      End
      Begin VB.TextBox txt_a_pedido 
         Height          =   375
         Left            =   2100
         TabIndex        =   18
         Top             =   1635
         Width           =   1290
      End
      Begin VB.Frame Frame4 
         Height          =   45
         Left            =   30
         TabIndex        =   17
         Top             =   705
         Width           =   3435
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   45
         Picture         =   "frmoracle_imprimir_facturas.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Imprimir Movimiento"
         Top             =   390
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox txt_copias_pedido 
         Height          =   375
         Left            =   2460
         TabIndex        =   15
         Top             =   1230
         Width           =   930
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Ped:"
         Height          =   195
         Left            =   90
         TabIndex        =   27
         Top             =   930
         Width           =   330
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   90
         TabIndex        =   26
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "De:"
         Height          =   195
         Left            =   90
         TabIndex        =   25
         Top             =   1725
         Width           =   255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "A:"
         Height          =   195
         Left            =   1890
         TabIndex        =   24
         Top             =   1725
         Width           =   150
      End
      Begin VB.Label Label7 
         BackColor       =   &H000000C0&
         Caption         =   " Por pedido"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   45
         TabIndex        =   23
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Copias:"
         Height          =   195
         Left            =   1890
         TabIndex        =   22
         Top             =   1320
         Width           =   525
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   45
      TabIndex        =   5
      Top             =   -60
      Width           =   3495
      Begin VB.CommandButton eflow 
         Caption         =   "Eflow"
         Height          =   315
         Left            =   2310
         TabIndex        =   43
         Top             =   390
         Width           =   1125
      End
      Begin VB.CommandButton cmd_subir_facturas_eflow 
         Caption         =   "Subir facturas a Eflow"
         Height          =   315
         Left            =   390
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   390
         Width           =   1920
      End
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   45
         Picture         =   "frmoracle_imprimir_facturas.frx":0408
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   390
         Width           =   345
      End
      Begin VB.TextBox txt_copias_embarque 
         Height          =   375
         Left            =   2460
         TabIndex        =   13
         Top             =   1230
         Width           =   930
      End
      Begin VB.CommandButton cmd_imprimir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   45
         Picture         =   "frmoracle_imprimir_facturas.frx":050A
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Imprimir Movimiento"
         Top             =   390
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Frame Frame2 
         Height          =   45
         Left            =   30
         TabIndex        =   11
         Top             =   705
         Width           =   3450
      End
      Begin VB.TextBox txt_a_embarque 
         Height          =   375
         Left            =   2100
         TabIndex        =   3
         Top             =   1635
         Width           =   1290
      End
      Begin VB.TextBox txt_de_embarque 
         Height          =   375
         Left            =   525
         TabIndex        =   2
         Top             =   1635
         Width           =   1290
      End
      Begin VB.TextBox txt_serie_embarque 
         Height          =   375
         Left            =   525
         TabIndex        =   1
         Top             =   1230
         Width           =   1290
      End
      Begin VB.TextBox txt_embarque 
         Height          =   375
         Left            =   525
         TabIndex        =   0
         Top             =   825
         Width           =   1665
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Copias:"
         Height          =   195
         Left            =   1890
         TabIndex        =   12
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label Label5 
         BackColor       =   &H000000C0&
         Caption         =   " Por embarque"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   45
         TabIndex        =   10
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "A:"
         Height          =   195
         Left            =   1890
         TabIndex        =   9
         Top             =   1725
         Width           =   150
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "De:"
         Height          =   195
         Left            =   90
         TabIndex        =   8
         Top             =   1725
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   90
         TabIndex        =   7
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Emb:"
         Height          =   195
         Left            =   90
         TabIndex        =   6
         Top             =   930
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmoracle_imprimir_facturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_orden As String
Dim var_posicion As Integer
Dim var_consecutivo_general As Double
Dim var_imprime_pedidos As Integer
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter

'Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const SW_SHOWNORMAL = 1
Dim var_ruta_facturas As String



Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
  "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal _
    szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
  
Sub DownloadFilefromWeb()
End Sub







Private Sub cmd_imprimir_Click()
   If Not IsNumeric(Me.txt_copias_embarque) Then
      Me.txt_copias_embarque = "1"
   End If
   If var_ruta_facturas <> "" Then
      If IsNumeric(Me.txt_de_embarque) Then
         If IsNumeric(Me.txt_a_embarque) Then
            If CDbl(Me.txt_de_embarque) <= CDbl(Me.txt_a_embarque) Then
               var_posible = 0
               rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  If rs!tipo_embarque = 1 Then
                     rsaux.Open "select distinct source_header_number from xxvia_tb_salidas where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  End If
                  If rs!tipo_embarque = 2 Then
                     If rsaux.State = 1 Then
                        rsaux.Close
                     End If
                     rsaux.Open "select distinct source_header_number from xxvia_tb_SAlidas_cajas where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
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
                     rsaux1.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     'var_cadena = "SELECT e.collector_id, A.SOURCE_HEADER_NUMBER,  HL.ADDRESS1 AS CUSTOMER_NAME,  A.released_status,  E.NAME , sum(shipped_quantity) as cantidad from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, hz_customer_profiles D, ar_collectors e Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id "
                     'var_cadena = var_cadena + " AND A.SOURCE_HEADER_NUMBER in (" + var_cadena_pedidos + ") AND A.released_status = 'C' and A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND D.collector_id = e.collector_id  group by  e.collector_id, A.SOURCE_HEADER_NUMBER, HL.ADDRESS1,  A.released_status,  E.NAME "
                     
                     var_cadena = "SELECT A.SOURCE_HEADER_NUMBER,  HL.ADDRESS1 AS CUSTOMER_NAME,  A.released_status,  sum(shipped_quantity) as cantidad from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID"
                     var_cadena = var_cadena + " AND A.SOURCE_HEADER_NUMBER in (" + var_Cadena_pedidos + ") AND A.released_status = 'C' and A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID group by A.SOURCE_HEADER_NUMBER, HL.ADDRESS1,  A.released_status"
                     
                     rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     var_i = 0
                     While Not rsaux.EOF
                           var_i = var_i + 1
                           rsaux.MoveNext
                     Wend
                     rsaux.Close
                     If var_j = var_i Then
                        'var_cadena = "SELECT  a.source_header_type_name from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND to_number(source_header_number) IN (" + var_cadena_pedidos + ")"
                        'var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id(+) = e.collector_id  AND ROWNUM = 1"
                        
                        var_cadena = "SELECT  a.source_header_type_name from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND to_number(source_header_number) IN (" + var_Cadena_pedidos + ")"
                        var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND  ROWNUM = 1"
                        
                        
                        rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux2.EOF Then
                           var_tipo_pedido = rsaux2!source_header_type_name
                        End If
                        rsaux2.Close
               
                        rsaux.Open "SELECT distinct APS.TRX_NUMBER as maximo From RA_CUSTOMER_TRX_ALL RCT, AR_PAYMENT_SCHEDULES_ALL APS Where INTERFACE_HEADER_ATTRIBUTE1 IN (" + CStr(var_Cadena_pedidos) + ") AND INTERFACE_HEADER_ATTRIBUTE2 = '" + var_tipo_pedido + "' AND RCT.customer_trx_id = APS.customer_trx_id order by APS.TRX_NUMBER", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           While Not rsaux.EOF
                                 var_archivo = var_ruta_facturas + Trim(Me.txt_serie_embarque) + "\" + Trim(Me.txt_serie_embarque) + Trim(CStr(rsaux!maximo)) + ".pdf"
                                 Archivoabuscar = var_archivo
                                 Text1 = var_archivo
                                 'MsgBox var_Archivo
                                 Archivoabuscar = Dir(var_archivo)
                                 If Archivoabuscar = "" Then
                                     var_posible = 1
                                 End If
                                 rsaux.MoveNext
                           Wend
                           If var_posible = 0 Then
                              rsaux.MoveFirst
                              While Not rsaux.EOF
                                    If var_unidad_organizacional = "93" Then
                                       VAR_Z = CDbl(Me.txt_copias_embarque)
                                       For var_i = 1 To CDbl(VAR_Z)
                                           var_archivo = var_ruta_facturas + Trim(Me.txt_serie_embarque) + "\" + Trim(Me.txt_serie_embarque) + Trim(Str(rsaux!maximo)) + ".PDF"
                                           Call Shell("c:\archivos de programa\adobe\reader 8.0\reader\acrord32.exe  /p /h " + var_archivo, vbMaximizedFocus)
                                       Next var_i
                                    Else
                                       Open (App.Path & "\EJPDF" + Trim(Me.txt_serie_embarque) + Trim(CStr(var_j)) + ".bat") For Output As #2
                                       Print #2, "START " + var_ruta_facturas + Trim(Me.txt_serie_embarque) + "\" + Trim(Me.txt_serie_embarque) + Trim(Str(rsaux!maximo)) + ".PDF"
                                       Close #2
                                       var_archivo = App.Path & "\EJPDF" + Trim(Me.txt_serie_embarque) + Trim(CStr(var_j)) + ".bat"
                                       x = Shell(var_archivo, vbHide)
                                    End If
                                    rsaux.MoveNext
                              Wend
                              rsaux.Close
                           Else
                              MsgBox "No se han generado todas las facturas", vbOKOnly, "ATENCION"
                           End If
                        Else
                           MsgBox "No se han generado todas las facturas", vbOKOnly, "ATENCION"
                        End If
                     Else
                        MsgBox "No se han generado todas las facturas", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "El embarque no tiene pedidos", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "El embarque esta vacio", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "El número de factura inicio debe de ser menor al número de factura final", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Número de factura final incorrecto", vbOKOnly, "atencion"
         End If
      Else
         MsgBox "Número de factura inicio incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a indicado una ruta donde se deben de generar las facturas electrónica", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_factura_Click()
   If IsNumeric(Me.txt_de_factura) Then
      If IsNumeric(Me.txt_a_factura) Then
         If CDbl(Me.txt_de_factura) <= CDbl(Me.txt_a_factura) Then
            If Me.txt_serie_factura <> "" Then
               For var_j = CDbl(Me.txt_de_factura) To CDbl(Me.txt_a_factura)
                   var_archivo = var_ruta_facturas + Me.txt_serie_factura + "\" + Trim(Me.txt_serie_factura) + Trim(CStr(var_j)) + ".pdf"
                   Archivoabuscar = var_archivo
                   Text1 = var_archivo
                   Archivoabuscar = Dir(var_archivo)
                   If Archivoabuscar = "" Then
                      var_posible = 1
                   End If
                   If var_posible = 1 Then
                      Open (App.Path & "\EJPDF" + Trim(Me.txt_serie_embarque) + Trim(CStr(var_j)) + ".bat") For Output As #2
                      Print #2, "START " + var_ruta_facturas + Me.txt_serie_factura + "\" + Trim(Me.txt_serie_factura) + Trim(Str(var_j)) + ".PDF"
                      Close #2
                      var_archivo = App.Path & "\EJPDF" + Trim(Me.txt_serie_embarque) + Trim(CStr(var_j)) + ".bat"
                      x = Shell(var_archivo, vbHide)
                   End If
               Next var_j
            Else
               MsgBox "Se debe de indicar una serie", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El número final de factura debe de ser mayor al inicial"
         End If
      Else
         MsgBox "Número de factura final incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Número de factura inical incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub







Private Sub cmd_subir_facturas_eflow_Click()
   
   Dim strSavePath As String
   Dim URL As String, ext As String
   Dim buf, ret As Long
   If Not IsNumeric(Me.txt_copias_embarque) Then
      Me.txt_copias_embarque = "1"
   End If
   If var_ruta_facturas <> "" Then
      If IsNumeric(Me.txt_de_embarque) Then
         If IsNumeric(Me.txt_a_embarque) Then
            If CDbl(Me.txt_de_embarque) <= CDbl(Me.txt_a_embarque) Then
               var_posible = 0
               rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  If rs!tipo_embarque = 1 Then
                     rsaux.Open "select distinct source_header_number from xxvia_tb_salidas where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  End If
                  If rs!tipo_embarque = 2 Then
                     If rsaux.State = 1 Then
                        rsaux.Close
                     End If
                     rsaux.Open "select distinct source_header_number from xxvia_tb_SAlidas_cajas where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
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
                     rsaux1.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     'var_cadena = "SELECT e.collector_id, A.SOURCE_HEADER_NUMBER,  HL.ADDRESS1 AS CUSTOMER_NAME,  A.released_status,  E.NAME , sum(shipped_quantity) as cantidad from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, hz_customer_profiles D, ar_collectors e Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id "
                     'var_cadena = var_cadena + " AND A.SOURCE_HEADER_NUMBER in (" + var_cadena_pedidos + ") AND A.released_status = 'C' and A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND D.collector_id = e.collector_id  group by  e.collector_id, A.SOURCE_HEADER_NUMBER, HL.ADDRESS1,  A.released_status,  E.NAME "
                     
                     var_cadena = "SELECT A.SOURCE_HEADER_NUMBER,  HL.ADDRESS1 AS CUSTOMER_NAME,  A.released_status,  sum(shipped_quantity) as cantidad from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID"
                     var_cadena = var_cadena + " AND A.SOURCE_HEADER_NUMBER in (" + var_Cadena_pedidos + ") AND A.released_status = 'C' and A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND a.source_header_type_name not in ('VIA_PEDIDO_INTERNO','TEX_PEDIDO_INTERNO')group by A.SOURCE_HEADER_NUMBER, HL.ADDRESS1,  A.released_status"
                     
                     rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     var_i = 0
                     While Not rsaux.EOF
                           var_i = var_i + 1
                           rsaux.MoveNext
                     Wend
                     rsaux.Close
                     If var_j >= var_i Then
                        'var_cadena = "SELECT  a.source_header_type_name from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND to_number(source_header_number) IN (" + var_cadena_pedidos + ")"
                        'var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id(+) = e.collector_id  AND ROWNUM = 1"
                        
                        var_cadena = "SELECT  a.source_header_type_name from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND to_number(source_header_number) IN (" + var_Cadena_pedidos + ")"
                        var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND  ROWNUM = 1 AND a.source_header_type_name not in ('VIA_PEDIDO_INTERNO','TEX_PEDIDO_INTERNO')"
                        
                        
                        rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux2.EOF Then
                           var_tipo_pedido = rsaux2!source_header_type_name
                        End If
                        rsaux2.Close
               
                        rsaux.Open "SELECT distinct APS.TRX_NUMBER as maximo From RA_CUSTOMER_TRX_ALL RCT, AR_PAYMENT_SCHEDULES_ALL APS Where INTERFACE_HEADER_ATTRIBUTE1 IN (" + CStr(var_Cadena_pedidos) + ") AND INTERFACE_HEADER_ATTRIBUTE2 = '" + var_tipo_pedido + "' AND RCT.customer_trx_id = APS.customer_trx_id order by APS.TRX_NUMBER", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           While Not rsaux.EOF
                                 rsaux1.Open "select customer_trx_id from xxvia_Tb_control_doc_fiscales where serie = '" + Me.txt_serie_embarque + "' and numero = " + CStr(rsaux!maximo), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 If rsaux1.EOF Then
                                    var_posible = 1
                                 Else
                                    'rsaux10.Open "SELECT customer_trx_id FROM XXVIA_TB_CONTROL_DOC_FISCALES WHERE CADENA LIKE '%LUGAR_EXPEDICION:,%' AND SERIE = '" + Me.txt_serie_embarque + "' AND NUMERO = " + CStr(rsaux!maximo), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                    'If rsaux10.EOF Then
                                    '   rsaux2.Open "CALL XXVIA_SEND_POST('" + CStr(rsaux1!CUSTOMER_TRX_ID) + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                    'Else
                                    '   MsgBox "El documento " + Me.txt_serie_embarque + CStr(rsaux!maximo) + " No tiene lugar de expedición", vbOKOnly, "ATENCION"
                                    'End If
                                    'rsaux10.Close
                                 End If
                                 rsaux1.Close
                                 rsaux.MoveNext
                           Wend
                           If var_posible = 0 Then
                           
                           
                           
                           
                           
                           
                              rsaux.MoveFirst
                              While Not rsaux.EOF
                                    VAR_Z = CDbl(Me.txt_copias_embarque)
                                    For var_i = 1 To 1
                                        x = 1
                                        If x = 1 Then
                                           rsaux1.Open "select customer_trx_id, cadena as cadena, numero from xxvia_tb_control_doc_fiscales where serie = '" + Me.txt_serie_embarque + "' and numero = " + CStr(rsaux!maximo), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                           If rsaux1.EOF Then
                                              var_posible = 1
                                              MsgBox "Vuelva a correr el concurrente Eflow rapido " + Me.txt_serie_embarque + CStr(rsaux!maximo) + ", ejecute nuevamente el concurrente", vbOKOnly, "ATENCION"
                                           Else
                                              var_cadena = rsaux1!Cadena
                                              var_cadena_rfc = Mid(var_cadena, 34, 12)
                                              var_cadena_str = ""
                                              Open ("C:\SISTEMAS\" + Trim(Me.txt_serie_embarque) + Trim(Str(rsaux!maximo)) + ".FAC") For Output As #1
                                              For var_i_2 = 1 To Len(var_cadena)
                                                  If Asc(Mid(var_cadena, var_i_2, 1)) = 63 Then
                                                     If Trim(var_cadena_str) = "CONDICIONES_PAGO:- -" And Mid(Me.txt_serie_embarque, 1, 2) = "NC" Then
                                                        var_cadena_str = "CONDICIONES_PAGO:NO IDENTIFICADO"
                                                     End If
                                                     Print #1, var_cadena_str
                                                     var_cadena_str = ""
                                                  Else
                                                     var_cadena_str = var_cadena_str + Mid(var_cadena, var_i_2, 1)
                                                  End If
                                              Next var_i_2
                                              Print #1, "FIN:"
                                              Close #1
                         
                                              strconsulta = "SELECT SERIE FROM XXVIA_tB_cONTROL_DOC_FISCALES WHERE SERIE = ? and numero = ? AND NUMERO_TIENDA = '3.3'"
                                              With comandoORA
                                                   .ActiveConnection = cnnoracle_4
                                                   .CommandType = adCmdText
                                                   .CommandText = strconsulta
                                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Trim(Me.txt_serie_embarque))
                                                   .Parameters.Append parametro
                                                   Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, rsaux!maximo)
                                                   .Parameters.Append parametro
                                              End With
                                              Set rsaux9 = comandoORA.execute
                                              Set comandoORA = Nothing
                                              Set parametro = Nothing
                                              If Not rsaux9.EOF Then
                                                 VAR_33 = 1
                                              Else
                                                 VAR_33 = 0
                                              End If
                                              rsaux9.Close
                                              If VAR_33 = 1 Then
                                                 x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\" + Trim(Trim(Me.txt_serie_embarque)) + Trim(Str(rsaux!maximo)) + ".FAC" + "|https://facturas2.vianney.mx/cgi-bin/cfds/timbrarGR33|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
                                              Else
                                                 x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\" + Trim(Trim(Me.txt_serie_embarque)) + Trim(Str(rsaux!maximo)) + ".FAC" + "|https://facturas.vianney.mx/cgi-bin/cfds/timbrarGR|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
                                              End If
                                              
                                              
                                              
                                              
                                              'var_archivo = "C:\SISTEMAS\sube_fact" + Trim(Str(rsaux!maximo)) + ".bat"
                                              'x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\" + Trim(Me.txt_serie_embarque) + Trim(Str(rsaux!maximo)) + ".FAC" + "|https://facturas.vianney.mx/cgi-bin/cfds/timbrarGR|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
                                              var_archivo_listo = 0
                                              While var_archivo_listo = 0
                                                    var_archivo = UCase("C:\SISTEMAS\" + Me.txt_serie_embarque + CStr(rsaux!maximo) + " .TXT")
                                                    var_zzz = Dir(var_archivo, vbNormal)
                                                    If var_zzz <> "" Then
                                                       var_archivo_listo = 1
                                                    End If
                                              Wend
                                           End If
                                           rsaux1.Close
                                           
                                           
                                           
                                           
                                           
                                           
                                           
                                           'URL = "https://facturas.vianney.mx/cgi-bin/cfds/cfdsORACLE?cmd=download_pdf&rfc_emisor=VTH981105F90&serie=" + Trim(Me.txt_serie_embarque) + "&folio=" + Trim(CStr(rsaux!maximo))
                                           'buf = Split(URL, ".")
                                           'ext = buf(UBound(buf))
                                           'strSavePath = "C:\SISTEMAS\" + Trim(Me.txt_serie_embarque) + Trim(CStr(rsaux!maximo)) + ".pdf"
                                           'ret = URLDownloadToFile(0, URL, strSavePath, 0, 0)
                                           'If ret = 0 Then
                                           '   Call ShellExecute(Me.hwnd, "print", "C:\SISTEMAS\" + Trim(Me.txt_serie_embarque) + Trim(CStr(rsaux!maximo)) + ".PDF", vbNullString, vbNullString, SW_SHOWNORMAL = 1)
                                           'Else
                                           '   MsgBox "Error en la factura " + Me.txt_serie_embarque + Trim(CStr(rsaux!maximo))
                                           'End If
                                        Else
                                           'Call ShellExecute(Me.hwnd, "print", "C:\SISTEMAS\" + Trim(Me.txt_serie_embarque) + Trim(CStr(rsaux!maximo)) + ".PDF", vbNullString, vbNullString, SW_SHOWNORMAL = 1)
                                        End If
                                        
                                        
                                        
                                    Next var_i
                                    rsaux.MoveNext
                              Wend
                              rsaux.Close
                           Else
                              MsgBox "Vuelva a ejecutar el concurrente Eflow rapido", vbOKOnly, "ATENCION"
                              
                              
                              
                              
                           End If
                        Else
                           MsgBox "No se han generado todas las facturas", vbOKOnly, "ATENCION"
                        End If
                     Else
                        MsgBox "No se han generado todas las facturas", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "El embarque no tiene pedidos", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "El embarque esta vacio", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "El número de factura inicio debe de ser menor al número de factura final", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Número de factura final incorrecto", vbOKOnly, "atencion"
         End If
      Else
         MsgBox "Número de factura inicio incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a indicado una ruta donde se deben de generar las facturas electrónica", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Command1_Click()
   If Not IsNumeric(Me.txt_copias_pedido) Then
      Me.txt_copias_pedido = "1"
   End If
   If var_ruta_facturas <> "" Then
      If IsNumeric(Me.txt_de_pedido) Then
         If IsNumeric(Me.txt_a_pedido) Then
            If CDbl(Me.txt_de_pedido) <= CDbl(Me.txt_a_pedido) Then
               var_posible = 0
               For var_j = CDbl(Me.txt_de_pedido) To CDbl(Me.txt_a_pedido)
                   var_archivo = var_ruta_facturas + Me.txt_serie_pedido + "\" + Trim(Me.txt_serie_pedido) + Trim(CStr(var_j)) + ".pdf"
                   Archivoabuscar = var_archivo
                   
                   Text1 = var_archivo
                   Archivoabuscar = Dir(var_archivo)
                   If Archivoabuscar = "" Then
                      var_posible = 1
                   End If
               Next var_j
               If var_posible = 0 Then
                  For var_j = CDbl(Me.txt_de_pedido) To CDbl(Me.txt_a_pedido)
                      If var_unidad_organizacional = "90" Then
                         VAR_Z = CDbl(Me.txt_copias_pedido)
                         For var_i = 1 To VAR_Z
                             var_archivo = var_ruta_facturas + Me.txt_serie_pedido + "\" + Trim(Me.txt_serie_pedido) + Trim(Str(var_j)) + ".PDF"
                             Call Shell("c:\archivos de programa\adobe\reader 8.0\reader\acrord32.exe  /p /h " + var_archivo, vbMaximizedFocus)
                         Next var_i
                      Else
                         Open (App.Path & "\EJPDF" + Trim(Me.txt_serie_pedido) + Trim(CStr(var_j)) + ".bat") For Output As #2
                         Print #2, "START " + var_ruta_facturas + Me.txt_serie_pedido + "\"; Trim(Me.txt_serie_pedido) + Trim(Str(var_j)) + ".PDF"
                         Close #2
                         var_archivo = App.Path & "\EJPDF" + Trim(Me.txt_serie_pedido) + Trim(CStr(var_j)) + ".bat"
                         x = Shell(var_archivo, vbHide)
                      End If
                  Next var_j
               Else
                  MsgBox "No se han generado todas las facturas", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "El número de factura inicio debe de ser menor al número de factura final", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Número de factura final incorrecto", vbOKOnly, "atencion"
         End If
      Else
         MsgBox "Número de factura inicio incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a indicado una ruta donde se deben de generar las facturas electrónica", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Command2_Click()
   'Call ShellExecute(Me.hwnd, vbNullString, "https://facturas.vianney.mx/cgi-bin/cfds/cfdsORACLE?cmd=download_pdf&rfc_emisor=VTH981105F90&serie=FAEMX&folio=65156 ", vbNullString, vbNullString, SW_SHOWNORMAL)
   Call ShellExecute(Me.hwnd, "OPEN", "https://facturas.vianney.mx/cgi-bin/cfds/cfdsORACLE?cmd=download_pdf&rfc_emisor=VTH981105F90&serie=FAEMX&folio=65156 ", vbNullString, vbNullString, SW_SHOWNORMAL)
End Sub

Private Sub Command3_Click()
   'Call ShellExecute(Me.hwnd, vbNullString, "https://facturas.vianney.mx/cgi-bin/cfds/cfdsORACLE?cmd=download_pdf&rfc_emisor=VTH981105F90&serie=FAEMX&folio=65156 ", vbNullString, vbNullString, SW_SHOWNORMAL)
   'Call ShellExecute(vbNull, "PRINT", "https://facturas.vianney.mx/cgi-bin/cfds/cfdsORACLE?cmd=download_pdf&rfc_emisor=VTH981105F90&serie=FAEMX&folio=65156", vbNullString, vbNullString, vbHide)
   'Call ShellExecute(Me.hwnd, "print", "https://facturas.vianney.mx/cgi-bin/cfds/cfdsORACLE?cmd=download_pdf&rfc_emisor=VTH981105F90&serie=FAEMX&folio=65156 ", vbNullString, vbNullString, SW_SHOWNORMAL = 1)
   
    Dim strSavePath As String
    Dim URL As String, ext As String
    Dim buf, ret As Long
    URL = "https://facturas.vianney.mx/cgi-bin/cfds/cfdsORACLE?cmd=download_pdf&rfc_emisor=VTH981105F90&serie=FAEMX&folio=65608"
    buf = Split(URL, ".")
    ext = buf(UBound(buf))
    strSavePath = "C:\SISTEMAS\FAEMX65608"
    ret = URLDownloadToFile(0, URL, strSavePath, 0, 0)
    If ret = 0 Then
        Call ShellExecute(Me.hwnd, "print", "C:\SISTEMAS\ARCHIVO_G.PDF", vbNullString, vbNullString, SW_SHOWNORMAL = 1)
    Else
        MsgBox "Error"
    End If
   


End Sub

Private Sub Command4_Click()
   Dim strSavePath As String
   Dim URL As String, ext As String
   Dim buf, ret As Long
   If Not IsNumeric(Me.txt_copias_embarque) Then
      Me.txt_copias_embarque = "1"
   End If
   If var_ruta_facturas <> "" Then
      If IsNumeric(Me.txt_de_embarque) Then
         If IsNumeric(Me.txt_a_embarque) Then
            If CDbl(Me.txt_de_embarque) <= CDbl(Me.txt_a_embarque) Then
               var_posible = 0
               rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  If rs!tipo_embarque = 1 Then
                     rsaux.Open "select distinct source_header_number from xxvia_tb_salidas where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  End If
                  If rs!tipo_embarque = 2 Then
                     If rsaux.State = 1 Then
                        rsaux.Close
                     End If
                     rsaux.Open "select distinct source_header_number from xxvia_tb_SAlidas_cajas where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
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
                     rsaux1.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     'var_cadena = "SELECT e.collector_id, A.SOURCE_HEADER_NUMBER,  HL.ADDRESS1 AS CUSTOMER_NAME,  A.released_status,  E.NAME , sum(shipped_quantity) as cantidad from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, hz_customer_profiles D, ar_collectors e Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id "
                     'var_cadena = var_cadena + " AND A.SOURCE_HEADER_NUMBER in (" + var_cadena_pedidos + ") AND A.released_status = 'C' and A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND D.collector_id = e.collector_id  group by  e.collector_id, A.SOURCE_HEADER_NUMBER, HL.ADDRESS1,  A.released_status,  E.NAME "
                     
                     var_cadena = "SELECT A.SOURCE_HEADER_NUMBER,  HL.ADDRESS1 AS CUSTOMER_NAME,  A.released_status,  sum(shipped_quantity) as cantidad from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID"
                     var_cadena = var_cadena + " AND A.SOURCE_HEADER_NUMBER in (" + var_Cadena_pedidos + ") AND A.released_status = 'C' and A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND a.source_header_type_name not in ('VIA_PEDIDO_INTERNO','TEX_PEDIDO_INTERNO')group by A.SOURCE_HEADER_NUMBER, HL.ADDRESS1,  A.released_status"
                     
                     rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     var_i = 0
                     While Not rsaux.EOF
                           var_i = var_i + 1
                           rsaux.MoveNext
                     Wend
                     rsaux.Close
                     If var_j >= var_i Then
                        'var_cadena = "SELECT  a.source_header_type_name from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND to_number(source_header_number) IN (" + var_cadena_pedidos + ")"
                        'var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id(+) = e.collector_id  AND ROWNUM = 1"
                        
                        var_cadena = "SELECT  a.source_header_type_name from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND to_number(source_header_number) IN (" + var_Cadena_pedidos + ")"
                        var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND  ROWNUM = 1 AND a.source_header_type_name not in ('VIA_PEDIDO_INTERNO','TEX_PEDIDO_INTERNO')"
                        
                        
                        rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux2.EOF Then
                           var_tipo_pedido = rsaux2!source_header_type_name
                        End If
                        rsaux2.Close
               
                        rsaux.Open "SELECT distinct APS.TRX_NUMBER as maximo From RA_CUSTOMER_TRX_ALL RCT, AR_PAYMENT_SCHEDULES_ALL APS Where INTERFACE_HEADER_ATTRIBUTE1 IN (" + CStr(var_Cadena_pedidos) + ") AND INTERFACE_HEADER_ATTRIBUTE2 = '" + var_tipo_pedido + "' AND RCT.customer_trx_id = APS.customer_trx_id order by APS.TRX_NUMBER", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           While Not rsaux.EOF
                                 rsaux1.Open "select customer_trx_id from xxvia_Tb_control_doc_fiscales where serie = '" + Me.txt_serie_embarque + "' and numero = " + CStr(rsaux!maximo), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 If rsaux1.EOF Then
                                    var_posible = 1
                                 Else
                                    'rsaux10.Open "SELECT customer_trx_id FROM XXVIA_TB_CONTROL_DOC_FISCALES WHERE CADENA LIKE '%LUGAR_EXPEDICION:,%' AND SERIE = '" + Me.txt_serie_embarque + "' AND NUMERO = " + CStr(rsaux!maximo), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                    'If rsaux10.EOF Then
                                    '   rsaux2.Open "CALL XXVIA_SEND_POST('" + CStr(rsaux1!CUSTOMER_TRX_ID) + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                    'Else
                                    '   MsgBox "El documento " + Me.txt_serie_embarque + CStr(rsaux!maximo) + " No tiene lugar de expedición", vbOKOnly, "ATENCION"
                                    'End If
                                    'rsaux10.Close
                                 End If
                                 rsaux1.Close
                                 rsaux.MoveNext
                           Wend
                           If var_posible = 0 Then
                           
                           
                           
                           
                           
                           
                              rsaux.MoveFirst
                              While Not rsaux.EOF
                                    VAR_Z = CDbl(Me.txt_copias_embarque)
                                    x = 0
                                    If x = 1 Then
                                    rsaux1.Open "select customer_trx_id, cadena as cadena, numero from xxvia_tb_control_doc_fiscales where serie = '" + Me.txt_serie_embarque + "' and numero = " + CStr(rsaux!maximo), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                    var_j = rsaux1!numero
                                    var_cadena = rsaux1!Cadena
                                    var_cadena_rfc = Mid(var_cadena, 34, 12)
                                    var_cadena_str = ""
                                    Open ("C:\SISTEMAS\" + Trim(Me.txt_serie_embarque) + Trim(Str(var_j)) + ".FAC") For Output As #1
                                    For var_m = 1 To Len(var_cadena)
                                       'MsgBox Asc(Mid(var_cadena, var_i, 1))
                                       If Asc(Mid(var_cadena, var_m, 1)) = 63 Then
                                          If Trim(var_cadena_str) = "CONDICIONES_PAGO:- -" And Mid(Me.txt_serie_embarque, 1, 2) = "NC" Then
                                              var_cadena_str = "CONDICIONES_PAGO:NO IDENTIFICADO"
                                          End If
                                          Print #1, var_cadena_str
                                          var_cadena_str = ""
                                       Else
                                          var_cadena_str = var_cadena_str + Mid(var_cadena, var_m, 1)
                                       End If
                                    Next var_m
                                    Print #1, "FIN:"
                                    Close #1
                                    var_archivo = "C:\SISTEMAS\sube_fact" + Trim(Str(var_j)) + ".bat"
                                    x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\" + Trim(Me.txt_serie_embarque) + Trim(Str(var_j)) + ".FAC" + "|https://facturas.vianney.mx/cgi-bin/cfds/timbrarGR|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
                                    rsaux1.Close
                                    End If
                                    For var_i = 1 To VAR_Z
                                        x = 1
                                        If x = 1 Then
                                           URL = "https://facturas.vianney.mx/cgi-bin/cfds/cfdsORACLE?cmd=download_pdf&rfc_emisor=VTH981105F90&serie=" + Trim(Me.txt_serie_embarque) + "&folio=" + Trim(CStr(rsaux!maximo))
                                           buf = Split(URL, ".")
                                           ext = buf(UBound(buf))
                                           strSavePath = "C:\SISTEMAS\" + Trim(Me.txt_serie_embarque) + Trim(CStr(rsaux!maximo)) + ".pdf"
                                           ret = URLDownloadToFile(0, URL, strSavePath, 0, 0)
                                           If ret = 0 Then
                                              Call ShellExecute(Me.hwnd, "print", "C:\SISTEMAS\" + Trim(Me.txt_serie_embarque) + Trim(CStr(rsaux!maximo)) + ".PDF", vbNullString, vbNullString, SW_SHOWNORMAL = 1)
                                           Else
                                              MsgBox "Error en la factura " + Me.txt_serie_embarque + Trim(CStr(rsaux!maximo))
                                           End If
                                        Else
                                           Call ShellExecute(Me.hwnd, "print", "C:\SISTEMAS\" + Trim(Me.txt_serie_embarque) + Trim(CStr(rsaux!maximo)) + ".PDF", vbNullString, vbNullString, SW_SHOWNORMAL = 1)
                                        End If
                                        
                                        
                                        
                                    Next var_i
                                    rsaux.MoveNext
                              Wend
                              rsaux.Close
                           Else
                              MsgBox "No se han generado todas las facturas", vbOKOnly, "ATENCION"
                           End If
                        Else
                           MsgBox "No se han generado todas las facturas", vbOKOnly, "ATENCION"
                        End If
                     Else
                        MsgBox "No se han generado todas las facturas", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "El embarque no tiene pedidos", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "El embarque esta vacio", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "El número de factura inicio debe de ser menor al número de factura final", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Número de factura final incorrecto", vbOKOnly, "atencion"
         End If
      Else
         MsgBox "Número de factura inicio incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a indicado una ruta donde se deben de generar las facturas electrónica", vbOKOnly, "ATENCION"
   End If

End Sub

Private Sub Command5_Click()
   If Not IsNumeric(Me.txt_copias_pedido) Then
      Me.txt_copias_pedido = "1"
   End If
   If var_ruta_facturas <> "" Then
      If IsNumeric(Me.txt_de_pedido) Then
         If IsNumeric(Me.txt_a_pedido) Then
            If CDbl(Me.txt_de_pedido) <= CDbl(Me.txt_a_pedido) Then
               var_posible = 0
               For var_j = CDbl(Me.txt_de_pedido) To CDbl(Me.txt_a_pedido)
                   rsaux1.Open "select customer_trx_id from xxvia_Tb_control_doc_fiscales where serie = '" + Me.txt_serie_pedido + "' and numero = " + CStr(var_j), cnnoracle_4, adOpenDynamic, adLockOptimistic
                   If rsaux1.EOF Then
                      var_posible = 1
                   Else
                      'rsaux2.Open "CALL XXVIA_SEND_POST('" + CStr(rsaux1!CUSTOMER_TRX_ID) + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                   End If
                   rsaux1.Close
                   
               Next var_j
               If var_posible = 0 Then
                  For var_j = CDbl(Me.txt_de_pedido) To CDbl(Me.txt_a_pedido)
                      
                       URL = "https://facturas.vianney.mx/cgi-bin/cfds/cfdsORACLE?cmd=download_pdf&rfc_emisor=VTH981105F90&serie=" + Trim(Me.txt_serie_pedido) + "&folio=" + Trim(CStr(var_j))
                       buf = Split(URL, ".")
                       ext = buf(UBound(buf))
                       strSavePath = "C:\SISTEMAS\" + Trim(Me.txt_serie_pedido) + Trim(CStr(var_j)) + ".pdf"
                       ret = URLDownloadToFile(0, URL, strSavePath, 0, 0)
                       If ret = 0 Then
                          Call ShellExecute(Me.hwnd, "print", "C:\SISTEMAS\" + Trim(Me.txt_serie_pedido) + Trim(CStr(var_j)) + ".PDF", vbNullString, vbNullString, SW_SHOWNORMAL = 1)
                       Else
                          MsgBox "Error en la factura " + Me.txt_serie_pedido + Trim(CStr(var_j))
                       End If
                  Next var_j
               Else
                  MsgBox "No se han generado todas las facturas", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "El número de factura inicio debe de ser menor al número de factura final", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Número de factura final incorrecto", vbOKOnly, "atencion"
         End If
      Else
         MsgBox "Número de factura inicio incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a indicado una ruta donde se deben de generar las facturas electrónica", vbOKOnly, "ATENCION"
   End If

End Sub

Private Sub Command6_Click()
   If IsNumeric(Me.txt_de_factura) Then
      If IsNumeric(Me.txt_a_factura) Then
         If CDbl(Me.txt_de_factura) <= CDbl(Me.txt_a_factura) Then
            If Me.txt_serie_factura <> "" Then
               For var_j = CDbl(Me.txt_de_factura) To CDbl(Me.txt_a_factura)
                   rsaux1.Open "select customer_trx_id from xxvia_Tb_control_doc_fiscales where serie = '" + Me.txt_serie_factura + "' and numero = " + CStr(var_j), cnnoracle_4, adOpenDynamic, adLockOptimistic
                   If rsaux1.EOF Then
                      var_posible = 1
                   Else
                      'rsaux2.Open "CALL XXVIA_SEND_POST('" + CStr(rsaux1!CUSTOMER_TRX_ID) + "')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                   End If
                   rsaux1.Close
               Next var_j
               
               For var_j = CDbl(Me.txt_de_factura) To CDbl(Me.txt_a_factura)
                   URL = "https://facturas.vianney.mx/cgi-bin/cfds/cfdsORACLE?cmd=download_pdf&rfc_emisor=VTH981105F90&serie=" + Trim(Me.txt_serie_factura) + "&folio=" + Trim(CStr(var_j))
                   buf = Split(URL, ".")
                   ext = buf(UBound(buf))
                   strSavePath = "C:\SISTEMAS\" + Trim(Me.txt_serie_factura) + Trim(CStr(var_j)) + ".pdf"
                   ret = URLDownloadToFile(0, URL, strSavePath, 0, 0)
                   If ret = 0 Then
                      Call ShellExecute(Me.hwnd, "print", "C:\SISTEMAS\" + Trim(Me.txt_serie_factura) + Trim(CStr(var_j)) + ".PDF", vbNullString, vbNullString, SW_SHOWNORMAL = 1)
                   Else
                      MsgBox "Error en la factura " + Me.txt_serie_factura + Trim(CStr(var_j))
                   End If
               Next var_j
            Else
               MsgBox "Se debe de indicar una serie", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El número final de factura debe de ser mayor al inicial"
         End If
      Else
         MsgBox "Número de factura final incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Número de factura inical incorrecto", vbOKOnly, "ATENCION"
   End If

End Sub

Private Sub eflow_Click()
   If Me.txt_serie_embarque = "FAEVII" Then
        strconsulta = "select distinct source_header_number from xxvia_tb_SAlidas_cajas where inte_emb_embarque = ? "
        With comandoORA
             .ActiveConnection = cnnoracle_4
             .CommandType = adCmdText
             .CommandText = strconsulta
             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(Me.txt_embarque))
             .Parameters.Append parametro
        End With
        Set rs = comandoORA.execute
        Set comandoORA = Nothing
        Set parametro = Nothing
        While Not rs.EOF
              strconsulta = "select * from ra_customer_trx_all a, AR.AR_PAYMENT_SCHEDULES_ALL b where ct_reference = ? and a.customer_trx_id = b.customer_trx_id"
              With comandoORA
                   .ActiveConnection = cnnoracle_4
                   .CommandType = adCmdText
                   .CommandText = strconsulta
                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!source_header_number))
                   .Parameters.Append parametro
              End With
              Set rsaux = comandoORA.execute
              Set comandoORA = Nothing
              Set parametro = Nothing
              If Not rsaux.EOF Then
                 'select fnd_document_sequences.NAME from ra_cust_trx_types_all, fnd_doc_sequence_assignments, fnd_document_sequences  where ra_cust_trx_types_all.name = fnd_doc_sequence_assignments.Category_code and cust_trx_type_id = 1244 and fnd_doc_sequence_assignments.doc_sequence_id = fnd_document_sequences.doc_sequence_id AND fnd_document_sequences.END_DATE IS NULL
                 While Not rsaux.EOF
                       VAR_TIPO_CAMBIO = rsaux!EXCHANGE_RATE
                       var_customer_trx_id = rsaux!CUSTOMER_tRX_ID
                       VAR_IMPORTE_TOTAL = rsaux!AMOUNT_LINE_ITEMS_ORIGINAL
                       VAR_IMPORTE_IVA = IIf(IsNull(rsaux!TAX_ORIGINAL), 0, rsaux!TAX_ORIGINAL)
                       VAR_TAZA_IVA = Round(((VAR_IMPORTE_IVA * 100) / VAR_IMPORTE_TOTAL), 0)
                       VAR_FECHA_FACTURA = Format(rsaux!trx_date, "YYYY-MM-DD")
                       var_fecha_creacion = Format(rsaux!creation_Date, "hh:mm:ss")
                       var_fecha_str = VAR_FECHA_FACTURA + "T" + var_fecha_creacion
                       
                       strconsulta = "SELECT SITE_USE_ID, PARTY_SITE_NUMBER, RAZON_SOCIAL_CLIENTE, CALLE, NUM_CALLE, NUM_INTERIOR, MUNICIPIO, COLONIA, CIUDAD, CODIGO_POSTAL, ESTADO, PAIS, RFC, PAYMENT_TERM_DESCRIPTION, orig_system_reference, METODO_PAGO, NUMERO_CUENTA  FROM OE_ORDER_HEADERS_ALL A, XXVIA_VW_CLIENTES_BCP B WHERE A.INVOICE_TO_ORG_ID = B.SITE_USE_ID AND ORDER_NUMBER = ?"
                       With comandoORA
                            .ActiveConnection = cnnoracle_4
                            .CommandType = adCmdText
                            .CommandText = strconsulta
                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!source_header_number))
                            .Parameters.Append parametro
                       End With
                       Set rsaux1 = comandoORA.execute
                       Set comandoORA = Nothing
                       Set parametro = Nothing
                       var_clave_cliente = rsaux1!party_site_number
                       var_nombre_cliente = rsaux1!RAZON_SOCIAL_CLIENTE
                       var_calle_cliente = rsaux1!calle
                       VAR_NUMERO = IIf(IsNull(rsaux1!NUM_CALLE), "", rsaux1!NUM_CALLE)
                       var_numero_interior = IIf(IsNull(rsaux1!num_interior), "", rsaux1!num_interior)
                       VAR_municipio_CLIENTE = rsaux1!municipio
                       var_colonia_cliente = rsaux1!colonia
                       VAR_CIUDAD_CLIENTE = rsaux1!ciudad
                       var_cp_cliente = rsaux1!codigo_postal
                       var_estado_cliente = rsaux1!estado
                       VAR_PAIS_CLIENTE = rsaux1!pais
                       var_rfc = rsaux1!rfc
                       var_termino_pago = rsaux1!PAYMENT_TERM_DESCRIPTION
                       var_referencia = IIf(IsNull(rsaux1!orig_system_reference), "", rsaux1!orig_system_reference)
                       VAR_METODO_PAGO = IIf(IsNull(rsaux1!METODO_PAGO), "NO IDENTIFICADO", rsaux1!METODO_PAGO)
                       var_numero_cuenta = IIf(IsNull(rsaux1!NUMERO_CUENTA), "NO IDENTIFICADO", rsaux1!NUMERO_CUENTA)
                       strconsulta = "select * from XXVIA_TB_TAX_ID where cliente = ?"
                       With comandoORA
                            .ActiveConnection = cnnoracle_4
                            .CommandType = adCmdText
                            .CommandText = strconsulta
                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_clave_cliente)
                            .Parameters.Append parametro
                       End With
                       Set rsaux2 = comandoORA.execute
                       Set comandoORA = Nothing
                       Set parametro = Nothing
                       If rsaux2.EOF Then
                          VAR_TAX_ID = IIf(IsNull(rsaux2!TAX_ID), "", rsaux2!TAX_ID)
                       Else
                          VAR_TAX_ID = ""
                       End If
                       rsaux2.Close
                                              
                       strconsulta = "SELECT b.collector_id, c.Name FROM hz_cust_accounts_all a,   hz_customer_profiles b,   ar_collectors c Where A.cust_account_id = b.cust_account_id   AND b.collector_id      = c.collector_id AND site_use_id         = ?"
                       With comandoORA
                            .ActiveConnection = cnnoracle_4
                            .CommandType = adCmdText
                            .CommandText = strconsulta
                            Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CStr(rsaux1!site_use_id))
                            .Parameters.Append parametro
                       End With
                       Set rsaux2 = comandoORA.execute
                       Set comandoORA = Nothing
                       Set parametro = Nothing
                       VAR_AGENTE = rsaux2!collector_id
                       var_nombre_agente = rsaux2!Name
                       rsaux1.Close
                       rsaux2.Close
                       
                       strconsulta = "SELECT PARTY_SITE_NUMBER, RAZON_SOCIAL_CLIENTE, CALLE, NUM_CALLE, NUM_INTERIOR, MUNICIPIO, COLONIA, CIUDAD, CODIGO_POSTAL, ESTADO, PAIS, RFC, PAYMENT_TERM_DESCRIPTION  FROM OE_ORDER_HEADERS_ALL A, XXVIA_VW_CLIENTES_BCP B WHERE A.SHIP_TO_ORG_ID = B.SITE_USE_ID AND ORDER_NUMBER = ?"
                       With comandoORA
                            .ActiveConnection = cnnoracle_4
                            .CommandType = adCmdText
                            .CommandText = strconsulta
                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rs!source_header_number))
                            .Parameters.Append parametro
                       End With
                       Set rsaux1 = comandoORA.execute
                       Set comandoORA = Nothing
                       Set parametro = Nothing
                       var_clave_esb = rsaux1!party_site_number
                       var_nombre_Esb = rsaux1!RAZON_SOCIAL_CLIENTE
                       var_calle_esb = rsaux1!calle
                       var_numero_esb = IIf(IsNull(rsaux1!NUM_CALLE), "", rsaux1!NUM_CALLE)
                       var_numero_interior_esb = IIf(IsNull(rsaux1!num_interior), "", rsaux1!num_interior)
                       var_municipio_esb = rsaux1!municipio
                       var_colonia_esb = rsaux1!colonia
                       var_ciudad_Esb = rsaux1!ciudad
                       var_cp_esb = rsaux1!codigo_postal
                       var_estado_esb = rsaux1!estado
                       var_pais_esb = rsaux1!pais
                       rsaux1.Close
                       
                       strconsulta = "select B.SEGMENT1 AS CODIGO, A.DESCRIPTION AS DESCRIPCION, QUANTITY_INVOICED AS CANTIDAD, UNIT_SELLING_PRICE  AS PRECIO from ra_customer_trx_lines_all A, XXVIA_SYSTEM_ITEMS_B B WHERE CUSTOMER_tRX_ID = ? AND A.INVENTORY_ITEM_ID = B.INVENTORY_ITEM_ID AND WAREHOUSE_ID = B.ORGANIZATION_ID AND UNIT_SELLING_PRICE>0"
                       With comandoORA
                            .ActiveConnection = cnnoracle_4
                            .CommandType = adCmdText
                            .CommandText = strconsulta
                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rsaux!CUSTOMER_tRX_ID))
                            .Parameters.Append parametro
                       End With
                       Set rsaux1 = comandoORA.execute
                       Set comandoORA = Nothing
                       Set parametro = Nothing
                       While Not rsaux1.EOF
                             var_cadena = "INSERT INTO [TB_ORACLE_FACTURAS]"
                             var_cadena = var_cadena + " ([INTE_TEM_CONSECUTIVO]"
                             var_cadena = var_cadena + " ,[CUSTOMER_TRX_ID]"
                             var_cadena = var_cadena + " ,[SERIE]"
                             var_cadena = var_cadena + " ,[NUMERO]"
                             var_cadena = var_cadena + " ,[NUM_CTE]"
                             var_cadena = var_cadena + " ,[NOMBRE_CTE]"
                             var_cadena = var_cadena + " ,[RFC_CTE]"
                             var_cadena = var_cadena + " ,[CALLE_CTE]"
                             var_cadena = var_cadena + " ,[NUMERO_CTE]"
                             var_cadena = var_cadena + " ,[INTERIOR_CTE]"
                             var_cadena = var_cadena + " ,[COLONIA_CTE]"
                             var_cadena = var_cadena + " ,[CIUDAD_CTE]"
                             var_cadena = var_cadena + " ,[MUNICIPIO_CTE]"
                             var_cadena = var_cadena + " ,[ESTADO_CTE]"
                             var_cadena = var_cadena + " ,[PAIS_CTE]"
                             var_cadena = var_cadena + " ,[ENTREGA]"
                             var_cadena = var_cadena + " ,[FECHA]"
                             var_cadena = var_cadena + " ,[AGENTE]"
                             var_cadena = var_cadena + " ,[CONDICIONES_PAGO]"
                             var_cadena = var_cadena + " ,[FORMA_PAGO]"
                             var_cadena = var_cadena + " ,[METODO_PAGO]"
                             var_cadena = var_cadena + " ,[EMBARQUE]"
                             var_cadena = var_cadena + " ,[REFERENCIA_BANCARIA]"
                             var_cadena = var_cadena + " ,[PEDIDO]"
                             var_cadena = var_cadena + " ,[IMPORTE]"
                             var_cadena = var_cadena + " ,[DESCUENTO]"
                             var_cadena = var_cadena + " ,[IMPUESTO]"
                             var_cadena = var_cadena + " ,[IMPUESTO_TAZA]"
                             var_cadena = var_cadena + " ,[IMPUESTO_DESGLOSADO]"
                             var_cadena = var_cadena + " ,[RET_IVA]"
                             var_cadena = var_cadena + " ,[RET_IVA_TAZA]"
                             var_cadena = var_cadena + " ,[COMENT]"
                             var_cadena = var_cadena + " ,[COMENT_INICIO]"
                             var_cadena = var_cadena + " ,[COMENT_FIN]"
                             var_cadena = var_cadena + " ,[CANTIDAD]"
                             var_cadena = var_cadena + " ,[CODIGO]"
                             var_cadena = var_cadena + " ,[DESCRIPCION]"
                             var_cadena = var_cadena + " ,[PRECIO]"
                             var_cadena = var_cadena + " ,[DESCUENTO_PRECIO]"
                             var_cadena = var_cadena + " ,[FRACCION]"
                             var_cadena = var_cadena + " ,[ORIGEN]"
                             var_cadena = var_cadena + " ,[TAX_ID]"
                             var_cadena = var_cadena + " ,[TIPO_CAMBIO])"
                             var_cadena = var_cadena + " Values"
                             var_cadena = var_cadena + " (0"
                             var_cadena = var_cadena + " , " + CStr(var_customer_trx_id)
                             var_cadena = var_cadena + " , 'FAEVII'"
                             var_cadena = var_cadena + " ," + rsaux!TRX_NUMBER
                             var_cadena = var_cadena + " ,'" + var_clave_cliente + "'"
                             var_cadena = var_cadena + " ,'" + var_nombre_cliente + "'"
                             var_cadena = var_cadena + " ,'" + var_rfc + "'"
                             var_cadena = var_cadena + " ,'" + var_calle_cliente + "'"
                             var_cadena = var_cadena + " ,'" + VAR_NUMERO + "'"
                             var_cadena = var_cadena + " ,'" + var_numero_interior + "'"
                             var_cadena = var_cadena + " ,'" + IIf(IsNull(var_colonia_cliente), "", var_colonia_cliente) + "'"
                             var_cadena = var_cadena + " ,'" + IIf(IsNull(VAR_CIUDAD_CLIENTE), "", VAR_CIUDAD_CLIENTE) + "'"
                             var_cadena = var_cadena + " ,'" + IIf(IsNull(VAR_municipio_CLIENTE), "", VAR_municipio_CLIENTE) + "'"
                             var_cadena = var_cadena + " ,'" + var_estado_cliente + "'"
                             var_cadena = var_cadena + " ,'" + VAR_PAIS_CLIENTE + "'"
                             If VAR_PAIS_CLIENTE = "US" Then
                                strconsulta = "select * from xxvia_tb_complementos_pk_list where codigo = ?"
                                With comandoORA
                                     .ActiveConnection = cnnoracle_4
                                     .CommandType = adCmdText
                                     .CommandText = strconsulta
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rsaux1!CODIGO)
                                     .Parameters.Append parametro
                                End With
                                Set rsaux9 = comandoORA.execute
                                Set comandoORA = Nothing
                                Set parametro = Nothing
                                If Not rsaux9.EOF Then
                                   VAR_FRACCION = IIf(IsNull(rsaux9!fraccion_americana), "", rsaux9!fraccion_americana)
                                   VAR_ORIGEN = IIf(IsNull(rsaux9!originario), "", rsaux9!originario)
                                Else
                                   VAR_ORIGEN = ""
                                   VAR_FRACCION = ""
                                End If
                                rsaux9.Close
                             Else
                                strconsulta = "select * from xxvia_tb_complementos_pk_list where codigo = ?"
                                With comandoORA
                                     .ActiveConnection = cnnoracle_4
                                     .CommandType = adCmdText
                                     .CommandText = strconsulta
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rsaux1!CODIGO)
                                     .Parameters.Append parametro
                                End With
                                Set rsaux9 = comandoORA.execute
                                Set comandoORA = Nothing
                                Set parametro = Nothing
                                If Not rsaux9.EOF Then
                                   VAR_FRACCION = IIf(IsNull(rsaux9!fraccion_arancelaria), "", rsaux9!fraccion_arancelaria)
                                   VAR_ORIGEN = IIf(IsNull(rsaux9!originario), "", rsaux9!originario)
                                Else
                                   VAR_ORIGEN = ""
                                   VAR_FRACCION = ""
                                End If
                                rsaux9.Close
                             End If
                             
                             var_cadena = var_cadena + " ,'" + var_clave_esb + " " + var_nombre_Esb + " " + var_calle_esb + " " + var_numero_esb + " " + var_numero_interior_esb + " " + IIf(IsNull(var_colonia_esb), "", var_colonia_esb) + " " + IIf(IsNull(var_municipio_esb), "", var_municipio_esb) + " " + IIf(IsNull(var_ciudad_Esb), "", var_ciudad_Esb) + " " + var_pais_esb + "'"
                             var_cadena = var_cadena + " ,'" + var_fecha_str + "'"
                             var_cadena = var_cadena + " ,'" + CStr(VAR_AGENTE) + " " + var_nombre_agente + "'"
                             var_cadena = var_cadena + " ,'" + var_termino_pago + "'"
                             var_cadena = var_cadena + " ,'PAGO HECHO EN UNA SOLA EXHIBICION '"
                             var_cadena = var_cadena + " ,'" + VAR_METODO_PAGO + "'"
                             var_cadena = var_cadena + " ,'" + Me.txt_embarque + "'"
                             var_cadena = var_cadena + " ,'" + var_referencia + "'"
                             var_cadena = var_cadena + " ,'" + CStr((rs!source_header_number)) + "'"
                             var_cadena = var_cadena + " ," + CStr(VAR_IMPORTE_TOTAL)
                             var_cadena = var_cadena + " ,0"
                             var_cadena = var_cadena + " ," + CStr(VAR_IMPORTE_IVA)
                             var_cadena = var_cadena + " ," + CStr(VAR_TAZA_IVA)
                             var_cadena = var_cadena + " ,0"
                             var_cadena = var_cadena + " ,0"
                             var_cadena = var_cadena + " ,0"
                             var_cadena = var_cadena + " ,''"
                             var_cadena = var_cadena + " ,''"
                             var_cadena = var_cadena + " ,''"
                             var_cadena = var_cadena + " ," + CStr(rsaux1!Cantidad)
                             var_cadena = var_cadena + " ,'" + rsaux1!CODIGO + "'"
                             var_cadena = var_cadena + " ,'" + rsaux1!DESCRIPCION + "'"
                             var_cadena = var_cadena + " ," + CStr(rsaux1!Precio)
                             var_cadena = var_cadena + " ,0"
                             var_cadena = var_cadena + " ,'" + CStr(VAR_FRACCION) + "'"
                             var_cadena = var_cadena + " ,'" + VAR_ORIGEN + "'"
                             var_cadena = var_cadena + " ,'" + VAR_TAX_ID + "'"
                             var_cadena = var_cadena + " ," + CStr(IIf(IsNull(VAR_TIPO_CAMBIO), 1, VAR_TIPO_CAMBIO)) + ")"
                             'MsgBox var_cadena
                             rsaux3.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                             rsaux1.MoveNext
                       Wend
                       rsaux1.Close
                       strconsulta = "select A.CUSTOMER_TRX_ID, B.SEGMENT1 AS CODIGO, A.DESCRIPTION AS DESCRIPCION, QUANTITY_INVOICED AS CANTIDAD, UNIT_SELLING_PRICE  AS PRECIO from ra_customer_trx_lines_all A, XXVIA_SYSTEM_ITEMS_B B WHERE CUSTOMER_tRX_ID = ? AND A.INVENTORY_ITEM_ID = B.INVENTORY_ITEM_ID AND WAREHOUSE_ID = B.ORGANIZATION_ID AND UNIT_SELLING_PRICE<0"
                       With comandoORA
                            .ActiveConnection = cnnoracle_4
                            .CommandType = adCmdText
                            .CommandText = strconsulta
                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rsaux!CUSTOMER_tRX_ID))
                            .Parameters.Append parametro
                       End With
                       Set rsaux1 = comandoORA.execute
                       Set comandoORA = Nothing
                       Set parametro = Nothing
                       While Not rsaux1.EOF
                             rsaux3.Open "update tb_oracle_facturas set DESCUENTO_PRECIO = " + CStr(rsaux1!Precio) + " where CUSTOMER_TRX_ID = " + CStr(rsaux1!CUSTOMER_tRX_ID) + " AND CODIGO = '" + rsaux1!CODIGO + "'", cnn, adOpenDynamic, adLockOptimistic
                             rsaux1.MoveNext
                       Wend
                       rsaux1.Close
                       strconsulta = "select SUM(QUANTITY_INVOICED * UNIT_SELLING_PRICE)  AS PRECIO from ra_customer_trx_lines_all A, XXVIA_SYSTEM_ITEMS_B B WHERE CUSTOMER_tRX_ID = ? AND A.INVENTORY_ITEM_ID = B.INVENTORY_ITEM_ID AND WAREHOUSE_ID = B.ORGANIZATION_ID AND UNIT_SELLING_PRICE<0"
                       With comandoORA
                            .ActiveConnection = cnnoracle_4
                            .CommandType = adCmdText
                            .CommandText = strconsulta
                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rsaux!CUSTOMER_tRX_ID))
                            .Parameters.Append parametro
                       End With
                       Set rsaux1 = comandoORA.execute
                       Set comandoORA = Nothing
                       Set parametro = Nothing
                       If Not rsaux1.EOF Then
                          VAR_DESCUENTO = IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value)
                       End If
                       rsaux1.Close
                       rsaux1.Open "SELECT SUM(PRECIO*CANTIDAD) FROM TB_ORACLE_FACTURAS WHERE CUSTOMER_TRX_ID = " + CStr(var_customer_trx_id), cnn, adOpenDynamic, adLockOptimistic
                       rsaux3.Open "UPDATE TB_ORACLE_fACTURAS SET IMPORTE = " + CStr(IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value)) + ", DESCUENTO = " + CStr((VAR_DESCUENTO * -1)) + " WHERE CUSTOMER_TRX_ID = " + CStr(var_customer_trx_id), cnn, adOpenDynamic, adLockOptimistic
                       rsaux3.Open "UPDATE TB_ORACLE_fACTURAS SET  DESCUENTO = " + CStr((VAR_DESCUENTO * -1)) + " WHERE CUSTOMER_TRX_ID = " + CStr(var_customer_trx_id), cnn, adOpenDynamic, adLockOptimistic
                       rsaux1.Close
                       rsaux.MoveNext
                 Wend
                 rsaux.Close
              Else
                    
              End If
              rs.MoveNext
        Wend
        rs.Close
   Else
      MsgBox "Modulo unico para facturas de exportaciones", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Form_Load()
   rs.Open "SELECT * FROM TB_UNIDADESORGANIZACIONALES WHERE VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      Me.txt_serie_embarque = IIf(IsNull(rs!Serie), "", rs!Serie)
      Me.txt_serie_pedido = IIf(IsNull(rs!Serie), "", rs!Serie)
      var_ruta_facturas = IIf(IsNull(rs!ruta_facturas), "", rs!ruta_facturas)
   End If
   rs.Close
End Sub

Private Sub txt_a_embarque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   End If
End Sub

Private Sub txt_a_factura_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_imprimir_factura.SetFocus
   End If
End Sub

Private Sub txt_a_pedido_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
End Sub

Private Sub txt_copias_embarque_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
End Sub

Private Sub txt_de_embarque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_a_embarque.SetFocus
   End If
End Sub

Private Sub txt_de_factura_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_a_factura.SetFocus
   End If
End Sub

Private Sub txt_de_pedido_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_embarque) Then
         rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            If rs!tipo_embarque = 1 Then
               rsaux.Open "select distinct source_header_number from xxvia_tb_salidas where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
            End If
            If rs!tipo_embarque = 2 Then
               rsaux.Open "select distinct source_header_number from xxvia_tb_SAlidas_cajas where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
            End If
            var_posible_embarque = 1
         End If
         rs.Close
         var_Cadena_pedidos = ""
         var_j = 0
         VAR_PEDIDO_HEADER_ID = ""
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
               
               var_cadena = " SELECT A.SOURCE_HEADER_NUMBER,  HL.ADDRESS1 AS CUSTOMER_NAME,  A.released_status,  sum(shipped_quantity) as cantidad from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID"
               var_cadena = var_cadena + " AND A.SOURCE_HEADER_NUMBER in (" + var_Cadena_pedidos + ") AND A.released_status = 'C' and A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID  and  a.source_header_type_name not in ('VIA_PEDIDO_INTERNO','TEX_PEDIDO_INTERNO') group by  A.SOURCE_HEADER_NUMBER, HL.ADDRESS1,  A.released_status"
               
               rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               var_i = 0
               If Not rsaux.EOF Then
                  VAR_PEDIDO_HEADER_ID = rsaux!source_header_number
                  While Not rsaux.EOF
                        var_i = var_i + 1
                        rsaux.MoveNext
                  Wend
               End If
               rsaux.Close
               If var_j >= var_i Then
                  rsaux7.Open "SELECT HEADER_ID FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = " + VAR_PEDIDO_HEADER_ID + " and order_type_id not in (1002,1023)", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux7.EOF Then
                     VAR_HEADER_ID = rsaux7!header_id
                  Else
                     VAR_HEADER_ID = 0
                  End If
                  rsaux7.Close
               
                  'var_cadena = "SELECT  a.source_header_type_name from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND to_number(source_header_number) IN (" + var_cadena_pedidos + ")"
                  'var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id(+) = e.collector_id  AND ROWNUM = 1"
                  var_cadena = "SELECT  a.source_header_type_name from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND  to_number(source_header_number) IN (" + var_Cadena_pedidos + ")"
                  var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND  ROWNUM = 1 and A.SOURCE_HEADER_ID       = " + CStr(VAR_HEADER_ID) + " AND A.ORGANIZATION_ID = " + var_unidad_organizacional + " and  a.source_header_type_name not in ('VIA_PEDIDO_INTERNO','TEX_PEDIDO_INTERNO')"
                  rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux2.EOF Then
                     var_tipo_pedido = rsaux2!source_header_type_name
                  End If
                  rsaux2.Close
               
                  rsaux.Open "SELECT MIN(RCT.CUSTOMER_TRX_ID) AS CUSTOMER_TRX_ID, MIN(APS.TRX_NUMBER) as MINIMO, max(APS.TRX_NUMBER) as maximo From RA_CUSTOMER_TRX_ALL RCT, AR_PAYMENT_SCHEDULES_ALL APS Where INTERFACE_HEADER_ATTRIBUTE1 IN (" + CStr(var_Cadena_pedidos) + ") AND INTERFACE_HEADER_ATTRIBUTE2 = '" + var_tipo_pedido + "' AND RCT.customer_trx_id = APS.customer_trx_id", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     strconsulta = "select fnd_document_sequences.NAME from ra_cust_trx_types_all, fnd_doc_sequence_assignments, fnd_document_sequences, RA_CUSTOMER_TRX_ALL  where ra_cust_trx_types_all.name = fnd_doc_sequence_assignments.Category_code  and fnd_doc_sequence_assignments.doc_sequence_id = fnd_document_sequences.doc_sequence_id AND fnd_document_sequences.END_DATE IS NULL AND  ra_cust_trx_types_all.cust_trx_type_id = RA_CUSTOMER_TRX_ALL.cust_trx_type_id AND RA_CUSTOMER_TRX_ALL.CUSTOMER_TRX_ID = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rsaux!CUSTOMER_tRX_ID))
                          .Parameters.Append parametro
                     End With
                     Set rsaux11 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     If Not rsaux11.EOF Then
                        Me.txt_serie_embarque = rsaux11(0).Value
                     End If
                     rsaux11.Close
                     
                     Me.txt_de_embarque = IIf(IsNull(rsaux(1).Value), "", rsaux(1).Value)
                     Me.txt_a_embarque = IIf(IsNull(rsaux(2).Value), "", rsaux(2).Value)
                     If Me.txt_de_embarque = "" Then
                        MsgBox "No se han generado todas las facturas", vbOKOnly, "ATENCION"
                     End If
                  End If
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
         MsgBox "Número de embarque incorrecto", vbOKOnly, "ATENCION"
      End If
      If rsaux.State = 1 Then
         rsaux.Close
      End If
   End If
End Sub

Private Sub txt_pedido_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_pedido) Then
         var_Cadena_pedidos = "'" + Me.txt_pedido + "'"
         rsaux1.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         'var_cadena = "SELECT  a.source_header_type_name from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND to_number(source_header_number) IN (" + var_cadena_pedidos + ")"
         'var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id(+) = e.collector_id  AND ROWNUM = 1"
         var_cadena = "SELECT  a.source_header_type_name from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND to_number(source_header_number) IN (" + var_Cadena_pedidos + ")"
         var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID  AND ROWNUM = 1"
         rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rsaux2.EOF Then
            var_tipo_pedido = rsaux2!source_header_type_name
         End If
         rsaux2.Close
         
         rsaux.Open "SELECT MIN(APS.TRX_NUMBER) as MINIMO, max(APS.TRX_NUMBER) as maximo From RA_CUSTOMER_TRX_ALL RCT, AR_PAYMENT_SCHEDULES_ALL APS Where INTERFACE_HEADER_ATTRIBUTE1 IN (" + CStr(var_Cadena_pedidos) + ") and INTERFACE_HEADER_ATTRIBUTE2 = '" + var_tipo_pedido + "' AND RCT.customer_trx_id = APS.customer_trx_id", cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
               Me.txt_de_pedido = IIf(IsNull(rsaux(0).Value), "", rsaux(0).Value)
               Me.txt_a_pedido = IIf(IsNull(rsaux(1).Value), "", rsaux(1).Value)
               If Me.txt_de_pedido = "" Then
                  MsgBox "No se han generado facturas para el pedido seleccionado", vbOKOnly, "ATENCION"
               End If
         Else
            MsgBox "No se han generado facturas para el pedido seleccionado", vbOKOnly, "ATENCION"
         End If
         rsaux.Close
      Else
         MsgBox "Número de pedido incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_serie_embarque_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
End Sub

Private Sub txt_serie_factura_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_de_factura.SetFocus
   End If
End Sub
