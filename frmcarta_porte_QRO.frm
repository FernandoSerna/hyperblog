VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcarta_porte_QRO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carta Porte Queretaro"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton com_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   30
      Picture         =   "frmcarta_porte_QRO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   75
      Left            =   0
      TabIndex        =   5
      Top             =   330
      Width           =   10365
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9960
      Picture         =   "frmcarta_porte_QRO.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame frm_lista 
      Height          =   2880
      Left            =   1200
      TabIndex        =   1
      Top             =   960
      Width           =   8500
      Begin MSComctlLib.ListView lv_lista 
         Height          =   2415
         Left            =   45
         TabIndex        =   2
         Top             =   375
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   4260
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   10583
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H000000C0&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   3
         Top             =   120
         Width           =   8415
      End
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   360
      Picture         =   "frmcarta_porte_QRO.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   6195
      Left            =   0
      TabIndex        =   7
      Top             =   405
      Width           =   10395
      Begin VB.TextBox txt_chofer 
         Height          =   390
         Left            =   990
         TabIndex        =   29
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txt_nombre_chofer 
         Height          =   390
         Left            =   2820
         TabIndex        =   28
         Top             =   600
         Width           =   6975
      End
      Begin VB.TextBox txt_embarque 
         Height          =   390
         Left            =   990
         TabIndex        =   27
         Top             =   165
         Width           =   1815
      End
      Begin VB.TextBox txt_unidad 
         Height          =   390
         Left            =   1005
         TabIndex        =   26
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txt_nombre_unidad 
         Height          =   390
         Left            =   2880
         TabIndex        =   25
         Top             =   1680
         Width           =   6975
      End
      Begin VB.TextBox txt_RFC 
         Enabled         =   0   'False
         Height          =   390
         Left            =   1005
         TabIndex        =   24
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txt_licencia 
         Enabled         =   0   'False
         Height          =   390
         Left            =   3885
         TabIndex        =   23
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txt_permsct 
         Enabled         =   0   'False
         Height          =   390
         Left            =   1005
         TabIndex        =   22
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txt_numpermisosct 
         Enabled         =   0   'False
         Height          =   390
         Left            =   4725
         TabIndex        =   21
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txt_seguro 
         Enabled         =   0   'False
         Height          =   390
         Left            =   1005
         TabIndex        =   20
         Top             =   2640
         Width           =   3855
      End
      Begin VB.TextBox txt_poliza 
         Enabled         =   0   'False
         Height          =   390
         Left            =   6045
         TabIndex        =   19
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txt_configuracion_vehicular 
         Enabled         =   0   'False
         Height          =   390
         Left            =   1920
         TabIndex        =   18
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox txt_placaVM 
         Enabled         =   0   'False
         Height          =   390
         Left            =   4605
         TabIndex        =   17
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox txt_modelo_VM 
         Enabled         =   0   'False
         Height          =   390
         Left            =   7605
         TabIndex        =   16
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox txt_remolque 
         Enabled         =   0   'False
         Height          =   390
         Left            =   1920
         TabIndex        =   15
         Top             =   3600
         Width           =   1815
      End
      Begin VB.TextBox txt_placa_remolque 
         Enabled         =   0   'False
         Height          =   390
         Left            =   4605
         TabIndex        =   14
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Frame Frame2 
         Height          =   75
         Left            =   0
         TabIndex        =   13
         Top             =   1560
         Width           =   10365
      End
      Begin VB.Frame Frame4 
         Height          =   75
         Left            =   0
         TabIndex        =   12
         Top             =   4080
         Width           =   10365
      End
      Begin VB.TextBox txt_pedido 
         Height          =   390
         Left            =   900
         TabIndex        =   11
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox txt_cliente 
         Height          =   390
         Left            =   900
         TabIndex        =   10
         Top             =   5160
         Width           =   8895
      End
      Begin VB.TextBox txt_tipo 
         Height          =   390
         Left            =   900
         TabIndex        =   9
         Top             =   4680
         Width           =   3735
      End
      Begin VB.TextBox txt_uuid 
         Height          =   390
         Left            =   900
         TabIndex        =   8
         Top             =   5640
         Width           =   8895
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Chofer"
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   780
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Embarque:"
         Height          =   195
         Left            =   120
         TabIndex        =   47
         Top             =   270
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Unidad:"
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   1785
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "RFC:"
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   1260
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Licencia:"
         Height          =   195
         Left            =   3120
         TabIndex        =   44
         Top             =   1185
         Width           =   645
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Perm. SCT:"
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   2280
         Width           =   810
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Número permiso SCT:"
         Height          =   195
         Left            =   3120
         TabIndex        =   42
         Top             =   2160
         Width           =   1545
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Seguro:"
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   2745
         Width           =   555
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Poliza:"
         Height          =   195
         Left            =   5040
         TabIndex        =   40
         Top             =   2745
         Width           =   465
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Configuración vehicular:"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   3225
         Width           =   1710
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Placa VM:"
         Height          =   195
         Left            =   3840
         TabIndex        =   38
         Top             =   3225
         Width           =   735
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Modelo VM:"
         Height          =   195
         Left            =   6720
         TabIndex        =   37
         Top             =   3225
         Width           =   855
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Tipo remolque:"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   3705
         Width           =   1050
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Número permiso SCT:"
         Height          =   195
         Left            =   3000
         TabIndex        =   35
         Top             =   1680
         Width           =   1545
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Placa:"
         Height          =   195
         Left            =   3840
         TabIndex        =   34
         Top             =   3705
         Width           =   450
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Pedido:"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   4305
         Width           =   540
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   4800
         Width           =   360
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   5280
         Width           =   525
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "UUID:"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   5760
         Width           =   450
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Chofer"
      Height          =   195
      Left            =   120
      TabIndex        =   50
      Top             =   1560
      Width           =   465
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Unidad:"
      Height          =   195
      Left            =   3120
      TabIndex        =   49
      Top             =   2640
      Width           =   75
   End
End
Attribute VB_Name = "frmcarta_porte_QRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tipo As Integer
Dim objConn As New ADODB.Connection
Dim objCmd As New ADODB.Command
Dim objParm As ADODB.Parameter
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter
   Dim cn As New ADODB.Connection
   Dim DSN As String
   Dim cn2 As New ADODB.Connection

Private Sub Text4_Change()

End Sub

Private Sub cmd_imprimir_Click()
If IsNumeric(Me.txt_embarque) Then
      If IsNumeric(Me.txt_pedido) Then
         'If CDbl(Me.txt_de) <= CDbl(Me.txt_a) Then
            If Me.txt_embarque <> "" Then
               If rsaux1.State = 1 Then
                  rsaux1.Close
               End If
               var_posible_embarque = 1
               var_Cadena_pedidos = Me.txt_pedido
               var_j = 0
               rsaux.Open "alter session set nls_languAge = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               var_cadena = "SELECT  oh.ordered_date, oh.source_document_id, oh.header_id, oh.order_number, oh.transactional_curr_code, NVL(ol.ordered_quantity,0) AS CANTIDAD_PEDIDA, NVL(ol.cancelled_quantity,0) AS CANTIDAD_NEGADA, NVL(ol.shipped_quantity,0)   AS CANTIDAD_surtida, ol.line_id, ol.ordered_item, ol.order_quantity_uom, ol.inventory_item_id, ol.price_list_id, ol.unit_selling_price, DECODE(ol.cancelled_flag,'Y','CANCELADA','SURTIDA') line_status, ol.flow_status_code"
               var_cadena = var_cadena + " FROM oe_order_headers_all oh, oe_order_lines_all ol, OE_ORDER_LINES_HISTORY OLH WHERE order_number  = " + Me.txt_pedido
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
               rsaux.Open "select * from xxvia_tb_encabezado_embarques where embarque = " + CStr(Me.txt_embarque), cnnoracle_4, adOpenDynamic, adLockOptimistic
               var_chofer = ""
               If Not rsaux.EOF Then
                  var_chofer = IIf(IsNull(rsaux!CHOFER_qro), "", rsaux!CHOFER_qro)
               Else
                  var_chofer = ""
               End If
               If var_chofer = "" Then
                  var_posible_embarque = 2
               End If
               rsaux.Close
               rsaux.Open "select * from xxvia_tb_encabezado_embarques where embarque = " + CStr(Me.txt_embarque), cnnoracle_4, adOpenDynamic, adLockOptimistic
               var_transporte = ""
               If Not rsaux.EOF Then
                  var_transporte = IIf(IsNull(rsaux!transporte_qro), "", rsaux!transporte_qro)
               Else
                  var_transporte = ""
               End If
               If var_transporte = "" Then
                  var_posible_embarque = 3
               End If
               rsaux.Close
               If var_posible_embarque = 1 Then
                  If rsaux1.State = 1 Then
                     rsaux1.Close
                  End If
                  var_tipo = 3
                  var_cadena = "CALL XXVIA_SP_TIMBRAR_TRASPASOS_11(?,?,?,?)"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = var_cadena
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_pedido))
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(var_tipo))
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(var_unidad_organizacional))
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_embarque))
                       .Parameters.Append parametro
                  End With
                  Set rsaux11 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  
                  If rsaux1.State = 1 Then
                     rsaux1.Close
                  End If
                  'rsaux.Open "call XXVIA_SP_TIMBRAR_TRASPASOS_3 ()"
                  var_cadena = "select customer_trx_id, cadena  as cadena, numero from xxvia_tb_control_doc_fiscales where serie = 'CPQR0" + Me.txt_embarque + "' and numero = " + CStr(Me.txt_pedido)
                  rsaux1.Open "select customer_trx_id, cadena  as cadena, numero from xxvia_tb_control_doc_fiscales where serie = 'CPQRO" + Me.txt_embarque + "_' and numero = " + CStr(Me.txt_pedido), cnnoracle_4, adOpenDynamic, adLockOptimistic
                  var_cadena = Replace(Replace(rsaux1!Cadena, "T23:", "T00:"), "AUTORIZADO  ", "AUTORIZADO ")
                  var_cadena_rfc = Mid(var_cadena, 34, 12)
                  VAR_CADENA_STR = ""
                  Open ("C:\SISTEMAS\CPCDMX" + Trim(Me.txt_embarque) + "_" + Trim(Str(Me.txt_pedido)) + ".FAC") For Output As #1
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
                        
                  var_archivo = "C:\SISTEMAS\sube_fact_" + Trim("CPCDMX" + Me.txt_embarque) + "_" + Trim(Str(txt_pedido)) + ".bat"
                  'x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\" + Trim(Me.txt_serie) + Trim(Str(var_j)) + ".FAC" + "|https://facturas2.vianney.mx/cgi-bin/cfds/timbrarGR33|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
                  x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\CPCDMX" + Trim(Me.txt_embarque) + "_" + Me.txt_pedido + ".FAC" + "|https://facturas2.vianney.mx/cgi-bin/cfds/timbrarGR33|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
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
               MsgBox "No se indico el embarque.", vbOKOnly, "ATENCION"
            End If
         'Else
         '   MsgBox "El número final de factura debe de ser mayor al inicial"
         'End If
      Else
         MsgBox "Número de pedido incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Número de pedido incorrecto", vbOKOnly, "ATENCION"
   End If
   Exit Sub
salir1:
      If Err.Number = -2147217900 Then
         'MsgBox Err.Description
         rsaux14.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         rsaux14.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         Resume
      End If

End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub com_guardar_Click()
   If IsNumeric(Me.txt_embarque) Then
      rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         If Me.txt_chofer <> "" Then
            If Me.txt_unidad <> "" Then
               If rsaux1.State = 1 Then
                  rsaux1.Close
               End If
               rsaux1.Open "UPDATE XXVIA_TB_ENCABEZADO_EMBARQUES SET CHOFER_qro = '" + Me.txt_chofer + "', TRANSPORTE_qro = '" + Me.txt_unidad + "' WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
            Else
               MsgBox "Unidad invalida.", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Chofer invalido.", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El embarque " + Me.txt_embarque + " no existe.", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      MsgBox "Número de embarque invalido.", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Form_Load()
   DSN = "eflow"
   If cn.State = 1 Then
      cn.Close
   End If
   cn.Open ("DSN=" & DSN & ";")
   Me.frm_lista.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If var_tipo = 1 Then
          Me.txt_chofer = Me.lv_lista.selectedItem
          Me.txt_nombre_chofer = Me.lv_lista.selectedItem.SubItems(1)
          rs.Open "select * from xxvia_tb_choferes where id_chofer = '" + Me.lv_lista.selectedItem + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
          If Not rs.EOF Then
             Me.txt_RFC = IIf(IsNull(rs!rfc), "", rs!rfc)
             Me.txt_licencia = IIf(IsNull(rs!licencia), "", rs!licencia)
          End If
          Me.txt_chofer.SetFocus
          rs.Close
       End If
       If var_tipo = 2 Then
          Me.txt_unidad = Me.lv_lista.selectedItem
          Me.txt_nombre_unidad = Me.lv_lista.selectedItem.SubItems(1)
          rsaux.Open "SELECT * FROM XXVIA_tB_tRANSPORTES WHERE CLAVE = '" + txt_unidad + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
          If Not rsaux.EOF Then
             Me.txt_permsct = IIf(IsNull(rsaux!PERMSCT), "", rsaux!PERMSCT)
             Me.txt_seguro = IIf(IsNull(rsaux!NOMBREASEG), "", rsaux!NOMBREASEG)
             Me.txt_numpermisosct = IIf(IsNull(rsaux!NUMPERMIsoSCT), "", rsaux!NUMPERMIsoSCT)
             Me.txt_poliza = IIf(IsNull(rsaux!NUMPOLIZASEG), "", rsaux!NUMPOLIZASEG)
             Me.txt_configuracion_vehicular = IIf(IsNull(rsaux!configvehicular), "", rsaux!configvehicular)
             Me.txt_placaVM = IIf(IsNull(rsaux!placavm), "", rsaux!placavm)
             Me.txt_modelo_VM = IIf(IsNull(rsaux!aniomodelovm), "", rsaux!aniomodelovm)
             Me.txt_remolque = IIf(IsNull(rsaux!subtiporem), "", rsaux!subtiporem)
             Me.txt_placa_remolque = IIf(IsNull(rsaux!placas), "", rsaux!placas)
          End If
          rsaux.Close
          Me.txt_unidad.SetFocus
       End If
    End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub txt_chofer_Change()
   Me.txt_nombre_chofer = ""
   Me.txt_RFC = ""
   Me.txt_licencia = ""
End Sub

Private Sub txt_chofer_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_tipo = 1
      Me.frm_lista.Visible = True
      Me.lv_lista.ListItems.Clear
      rs.Open "select * from xxvia_tb_choferes where nvl(rfc,' ')<> ' ' order by nombre", cnnoracle_4, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!id_chofer)
            list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE), "", rs!NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      Me.lv_lista.SetFocus
   End If
End Sub

Private Sub txt_chofer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      rs.Open "select * from xxvia_tb_choferes where id_chofer = '" + Me.txt_chofer + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_chofer = IIf(IsNull(rs!NOMBRE), "", rs!NOMBRE)
         Me.txt_RFC = IIf(IsNull(rs!rfc), "", rs!rfc)
         Me.txt_licencia = IIf(IsNull(rs!licencia), "", rs!licencia)
      Else
        Me.txt_chofer = ""
        Me.txt_nombre_chofer = ""
        Me.txt_licencia = ""
        Me.txt_RFC = ""
      End If
      rs.Close
      Me.txt_unidad.SetFocus
   End If
End Sub

Private Sub txt_embarque_Change()
   Me.txt_chofer = ""
   Me.txt_configuracion_vehicular = ""
   Me.txt_licencia = ""
   Me.txt_modelo_VM = ""
   Me.txt_nombre_chofer = ""
   Me.txt_nombre_unidad = ""
   Me.txt_numpermisosct = ""
   Me.txt_permsct = ""
   Me.txt_placa_remolque = ""
   Me.txt_placaVM = ""
   Me.txt_poliza = ""
   Me.txt_remolque = ""
   Me.txt_RFC = ""
   Me.txt_seguro = ""
   Me.txt_unidad = ""
   Me.txt_pedido = ""
   Me.txt_tipo = ""
   Me.txt_cliente = ""
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_embarque) Then
         rs.Open "select * from xxvia_Tb_encabezado_embarques where embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_chofer = IIf(IsNull(rs!CHOFER_qro), "", rs!CHOFER_qro)
            If var_chofer <> "" Then
               If rsaux.State = 1 Then
                  rsaux.Close
               End If
               rsaux.Open "SELECT * FROM XXVIA_TB_CHOFERES WHERE ID_CHOFER = '" + var_chofer + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  Me.txt_chofer = IIf(IsNull(rsaux!id_chofer), "", rsaux!id_chofer)
                  Me.txt_nombre_chofer = IIf(IsNull(rsaux!NOMBRE), "", rsaux!NOMBRE)
                  Me.txt_licencia = IIf(IsNull(rsaux!licencia), "", rsaux!licencia)
                  Me.txt_RFC = IIf(IsNull(rsaux!rfc), "", rsaux!rfc)
               End If
               rsaux.Close
            Else
               MsgBox "El embarque no tiene chofer asignado.", vbOKOnly, "ATENCION"
            End If
            var_transporte = IIf(IsNull(rs!transporte_qro), "", rs!transporte_qro)
            If var_transporte <> "" Then
               rsaux.Open "SELECT * FROM XXVIA_tB_tRANSPORTES WHERE CLAVE = '" + var_transporte + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  Me.txt_unidad = IIf(IsNull(rsaux!clave), "", rsaux!clave)
                  Me.txt_nombre_unidad = IIf(IsNull(rsaux!NOMBRE), "", rsaux!NOMBRE)
                  Me.txt_permsct = IIf(IsNull(rsaux!PERMSCT), "", rsaux!PERMSCT)
                  Me.txt_seguro = IIf(IsNull(rsaux!NOMBREASEG), "", rsaux!NOMBREASEG)
                  Me.txt_numpermisosct = IIf(IsNull(rsaux!NUMPERMIsoSCT), "", rsaux!NUMPERMIsoSCT)
                  Me.txt_poliza = IIf(IsNull(rsaux!NUMPOLIZASEG), "", rsaux!NUMPOLIZASEG)
                  Me.txt_configuracion_vehicular = IIf(IsNull(rsaux!configvehicular), "", rsaux!configvehicular)
                  Me.txt_placaVM = IIf(IsNull(rsaux!placavm), "", rsaux!placavm)
                  Me.txt_modelo_VM = IIf(IsNull(rsaux!aniomodelovm), "", rsaux!aniomodelovm)
                  Me.txt_remolque = IIf(IsNull(rsaux!subtiporem), "", rsaux!subtiporem)
                  Me.txt_placa_remolque = IIf(IsNull(rsaux!placas), "", rsaux!placas)
                  
               End If
               rsaux.Close
            Else
               MsgBox "El embarque no tiene transporte asignado.", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "El embarque no existe.", vbOKOnly, "ATENCION"
         End If
         rs.Close
         Me.txt_chofer.SetFocus
      Else
         MsgBox "Número de embarque incorrecto.", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_pedido_Change()
   Me.txt_tipo = ""
   Me.txt_cliente = ""
End Sub

Private Sub txt_pedido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_pedido) Then
         If Me.txt_embarque <> "" Then
            strconsulta = "SELECT * FROM XXVIA_TB_sALIDAS_CAJAS WHERE SOURCE_HEADER_NUMBER = ? "
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_pedido))
                 .Parameters.Append parametro
            End With
            Set rs = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            If Not rs.EOF Then
               If rs!inte_emb_embarque = CDbl(Me.txt_embarque) Then
                  strconsulta = "SELECT * FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ? "
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_pedido))
                       .Parameters.Append parametro
                  End With
                  Set rsaux = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  If Not rsaux.EOF Then
                     Me.txt_tipo = CStr(rsaux!ORDER_TYPE_ID)
                     If Me.txt_tipo = 1002 Then
                        var_requisicion = IIf(IsNull(rsaux!source_document_id), "", rsaux!source_document_id)
                        rsaux2.Open "SELECT A.ATTRIBUTE1, B.description, B.SECONDARY_INVENTORY_NAME FROM po_requisition_headers_ALL A, MTL_SECONDARY_INVENTORIES B WHERE requisition_header_id IN (" + CStr(var_requisicion) + ") AND secondary_inventory_name = A.ATTRIBUTE1", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux2.EOF Then
                           Me.txt_cliente = IIf(IsNull(rsaux2!Description), "", rsaux2!Description)
                        End If
                        rsaux2.Close
                     Else
                        rsaux2.Open "SELECT * FROM XXVIA_VW_CLIENTES_BCP WHERE SITE_USE_ID = " + CStr(rsaux!SHIP_TO_ORG_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux2.EOF Then
                           Me.txt_cliente = rsaux2!razon_social_cliente
                        End If
                        rsaux2.Close
                                                rsaux11.Open "select * from ra_customer_Trx_all where ct_reference = '" + Me.txt_pedido + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        If Not rsaux11.EOF Then
                           var_cadena = "select serie, numero  from xxvia_tb_control_doc_fiscales where customer_Trx_id = " + CStr(rsaux11!customer_Trx_id)
                           rsaux12.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If Not rsaux12.EOF Then
                              Set rsaux10 = cn.execute("SELECT * FROM facturas where factura = '" + rsaux12!Serie + CStr(rsaux12!numero) + "'")
                              If Not rsaux10.EOF Then
                                 Me.txt_uuid = IIf(IsNull(rsaux10!sat_uuid), "", rsaux10!sat_uuid)
                                 rsaux2.Open "update xxvia_Tb_control_doc_fiscales set cadena_original = '" + Me.txt_uuid + "' where customer_Trx_id = " + CStr(rsaux11!customer_Trx_id), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 rsaux2.Open "SELECT * FROM XXVIA_VW_CLIENTES_BCP WHERE SITE_USE_ID = " + CStr(rsaux!SHIP_TO_ORG_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 If Not rsaux2.EOF Then
                                    Me.txt_cliente = rsaux2!razon_social_cliente
                                 End If
                                 rsaux2.Close
                              Else
                                 MsgBox "El pedido no a sido timbrado", vbOKOnly, "ATENCION"
                                 Me.txt_uuid = ""
                                 Me.txt_cliente = ""
                                 Me.txt_tipo = ""
                              End If
                           Else
                              MsgBox "El pedido no a sido timbrado", vbOKOnly, "ATENCION"
                              Me.txt_uuid = ""
                              Me.txt_cliente = ""
                              Me.txt_tipo = ""
                           End If
                        Else
                           MsgBox "El pedido no a sido facturado", vbOKOnly, "ATENCION"
                           Me.txt_uuid = ""
                           Me.txt_cliente = ""
                           Me.txt_tipo = ""
                        End If
                        
                        
                     End If
                  Else
                     MsgBox "El pedido no existe", vbOKOnly, "ATENCION"
                  End If
                  rsaux.Close
               Else
                  Me.txt_pedido = ""
                  MsgBox "El pedido corresponde al embarque " + CStr(rs!inte_emb_embarque), vbOKOnly, "ATENCION"
               End If
            Else
               Me.txt_pedido = ""
               MsgBox "El pedido no existe o no se ha leido aun", vbOKOnly, "ATENCION"
            End If
            rs.Close
         Else
            Me.txt_pedido = ""
            MsgBox "No se a seleccionado un embarque.", vbOKOnly, "ATENCION"
         End If
      Else
         Me.txt_pedido = ""
         MsgBox "Número de pedido incorrecto.", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_unidad_Change()
   Me.txt_configuracion_vehicular = ""
   Me.txt_modelo_VM = ""
   Me.txt_nombre_unidad = ""
   Me.txt_numpermisosct = ""
   Me.txt_permsct = ""
   Me.txt_placa_remolque = ""
   Me.txt_placaVM = ""
   Me.txt_poliza = ""
   Me.txt_remolque = ""
   Me.txt_seguro = ""
End Sub

Private Sub txt_unidad_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_tipo = 2
      Me.frm_lista.Visible = True
      Me.lv_lista.ListItems.Clear
      rs.Open "select * from xxvia_tb_transportes order by nombre", cnnoracle_4, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!clave)
            list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE), "", rs!NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      Me.lv_lista.SetFocus
      
   End If
End Sub

Private Sub txt_unidad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      rsaux.Open "SELECT * FROM XXVIA_tB_tRANSPORTES WHERE CLAVE = '" + txt_unidad + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rsaux.EOF Then
         Me.txt_nombre_unidad = IIf(IsNull(rsaux!NOMBRE), "", rsaux!NOMBRE)
         Me.txt_permsct = IIf(IsNull(rsaux!PERMSCT), "", rsaux!PERMSCT)
         Me.txt_seguro = IIf(IsNull(rsaux!NOMBREASEG), "", rsaux!NOMBREASEG)
         Me.txt_numpermisosct = IIf(IsNull(rsaux!NUMPERMIsoSCT), "", rsaux!NUMPERMIsoSCT)
         Me.txt_poliza = IIf(IsNull(rsaux!NUMPOLIZASEG), "", rsaux!NUMPOLIZASEG)
         Me.txt_configuracion_vehicular = IIf(IsNull(rsaux!configvehicular), "", rsaux!configvehicular)
         Me.txt_placaVM = IIf(IsNull(rsaux!placavm), "", rsaux!placavm)
         Me.txt_modelo_VM = IIf(IsNull(rsaux!aniomodelovm), "", rsaux!aniomodelovm)
         Me.txt_remolque = IIf(IsNull(rsaux!subtiporem), "", rsaux!subtiporem)
         Me.txt_placa_remolque = IIf(IsNull(rsaux!placas), "", rsaux!placas)
      End If
      rsaux.Close
   End If
End Sub



