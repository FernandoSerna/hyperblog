VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmembarques 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar Embarque"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   Icon            =   "frmembarques.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   6210
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   255
      TabIndex        =   24
      Top             =   555
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   25
         Top             =   480
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   3228
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7937
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H000000FF&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   26
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   165
      Picture         =   "frmembarques.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5685
      Picture         =   "frmembarques.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos para embarque "
      Height          =   2535
      Left            =   165
      TabIndex        =   12
      Top             =   435
      Width           =   5850
      Begin VB.TextBox txt_nombre_transporte 
         Height          =   315
         Left            =   2280
         TabIndex        =   8
         Top             =   1770
         Width           =   3405
      End
      Begin VB.TextBox txt_transporte 
         Height          =   315
         Left            =   1155
         TabIndex        =   7
         Top             =   1770
         Width           =   1095
      End
      Begin VB.TextBox txt_jaula 
         Height          =   315
         Left            =   1155
         TabIndex        =   2
         Top             =   765
         Width           =   1095
      End
      Begin VB.ComboBox cmb_choferes 
         Height          =   315
         Left            =   2280
         TabIndex        =   6
         Top             =   1410
         Width           =   3420
      End
      Begin VB.TextBox txt_cubicaje 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1005
         TabIndex        =   10
         Top             =   3225
         Width           =   1095
      End
      Begin VB.TextBox txt_chofer 
         Height          =   315
         Left            =   1155
         TabIndex        =   5
         Top             =   1425
         Width           =   1095
      End
      Begin VB.TextBox txt_fecha 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3090
         TabIndex        =   1
         Top             =   420
         Width           =   1095
      End
      Begin VB.TextBox txt_numero 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1155
         TabIndex        =   0
         Top             =   420
         Width           =   1095
      End
      Begin VB.TextBox txt_agente 
         Height          =   315
         Left            =   1155
         TabIndex        =   3
         Top             =   1095
         Width           =   1095
      End
      Begin VB.ComboBox cmb_agentes 
         Height          =   315
         Left            =   2265
         TabIndex        =   4
         Top             =   1080
         Width           =   3420
      End
      Begin VB.TextBox txt_vehiculo 
         Height          =   315
         Left            =   1155
         TabIndex        =   9
         Top             =   2115
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Transporte:"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   1845
         Width           =   810
      End
      Begin VB.Label lab_estados 
         AutoSize        =   -1  'True
         Caption         =   "Anden:"
         Height          =   195
         Index           =   10
         Left            =   240
         TabIndex        =   23
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lab_estados 
         AutoSize        =   -1  'True
         Caption         =   "3"
         Height          =   195
         Index           =   9
         Left            =   2385
         TabIndex        =   22
         Top             =   3255
         Width           =   90
      End
      Begin VB.Label lab_estados 
         AutoSize        =   -1  'True
         Caption         =   "Cm."
         Height          =   195
         Index           =   7
         Left            =   2145
         TabIndex        =   21
         Top             =   3315
         Width           =   270
      End
      Begin VB.Label lab_estados 
         AutoSize        =   -1  'True
         Caption         =   "Cubicaje:"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   20
         Top             =   3285
         Width           =   660
      End
      Begin VB.Label lab_estados 
         AutoSize        =   -1  'True
         Caption         =   "Chofer:"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   19
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lab_estados 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Index           =   4
         Left            =   2415
         TabIndex        =   18
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lab_estados 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   600
      End
      Begin VB.Label lab_estados 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   1140
         Width           =   555
      End
      Begin VB.Label lab_estados 
         AutoSize        =   -1  'True
         Caption         =   "Placas:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   2175
         Width           =   525
      End
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   5040
      Top             =   645
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
            Picture         =   "frmembarques.frx":1006
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmembarques.frx":18E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmembarques.frx":21BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmembarques.frx":2756
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmembarques.frx":3032
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmembarques.frx":390C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmembarques.frx":41E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmembarques.frx":42F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmembarques.frx":440A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmembarques.frx":451C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmembarques.frx":462E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Index           =   0
      Left            =   150
      TabIndex        =   16
      Top             =   285
      Width           =   5850
   End
End
Attribute VB_Name = "frmembarques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objConn As New adodb.Connection
Dim objCmd As New adodb.Command
Dim objParm As adodb.Parameter
Dim var_contador_porcentaje As Integer
Dim var_cubicaje As Double
Dim var_ventana As Integer



Private Sub cmb_agentes_Click()
   'txt_agente = Obtener_llave(cnn, rsaux, "TB_agentes", "VCHA_age_NOMBRE", cmb_agentes, 1, "T")
   rs.Open "select * from ar_collectors where name = '" + Me.cmb_agentes + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      Me.txt_agente = IIf(IsNull(rs!collector_id), "", rs!collector_id)
   End If
   rs.Close
   'lv_rutas.ListItems.Clear
   'Dim list_item As ListItem
   'numero_items_rutas = 0
   'rs.Open "select vcha_rut_ruta_id,vcha_rut_nombre from TB_rutas where vcha_age_agente_id = '" + txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
   'While Not rs.EOF
   '      Set list_item = lv_rutas.ListItems.Add(, , rs(0).Value)
   '      list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
   '      rs.MoveNext:
   '      numero_items_aseguradoras = numero_items_aseguradoras + 1
   'Wend
   'rs.Close
End Sub

Private Sub cmb_agentes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(txt_agente) <> "" Then
         'txt_agente.Enabled = False
         'cmb_agentes.Enabled = False
         Call pro_enfoque(KeyAscii)
      End If
   End If
End Sub

Private Sub cmb_almacenes_Click()
   txt_almacen = Obtener_llave(cnn, rsaux, "TB_almacenes", "VCHA_alm_NOMBRE", cmb_almacenes, 2, "T")
End Sub

Private Sub cmb_almacenes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(txt_almacen) <> "" Then
         txt_almacen.Enabled = False
         cmb_almacenes.Enabled = False
         txt_vehiculo.Enabled = True
         txt_vehiculo.SetFocus
      End If
   End If
End Sub


Private Sub cmb_choferes_Click()
   txt_chofer = Obtener_llave(cnn, rsaux, "TB_CHOFERES", "VCHA_CHO_NOMBRE", cmb_choferes, 0, "T")
End Sub

Private Sub cmb_choferes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_transporte.SetFocus
   End If
End Sub



Private Sub cmd_guardar_Click()
   Dim clnt As New SoapClient30
   Dim var_arreglo() As String
   Dim var_posible As Boolean
   Dim var_posible_chofer As Boolean
   var_posible = True
   var_posible_chofer = True
   If var_empresa = "02" Or var_empresa = "03" Then
      If Me.txt_vehiculo <> "" Then
         rs.Open "SELECT * FROM TB_TRANSPORTES WHERE VCHA_TRN_NOMBRE = '" + Me.txt_vehiculo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_posible = True
         Else
            MsgBox "Placas de vehiculo incorrectas", vbOKOnly, "ATENCION"
            var_posible = False
         End If
         rs.Close
      Else
         var_posible = False
      End If
      If Me.txt_chofer <> "" Then
         rs.Open "SELECT * FROM tb_choferes WHERE vcha_cho_chofer_id = '" + Me.txt_chofer + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_posible_chofer = True
         Else
            MsgBox "Chofer incorrecto", vbOKOnly, "ATENCION"
            var_posible_chofer = False
         End If
         rs.Close
      Else
         var_posible_chofer = False
      End If
   Else
      var_posible_chofer = True
      var_posible = True
   End If
   
   
   If var_posible = True Then
      If var_posible_chofer = True Then
         If Trim(txt_numero) <> "" And Trim(Me.txt_agente) <> "" And Trim(Me.txt_jaula) <> "" Then
            If Me.txt_transporte <> "" Then
               If Me.txt_transporte = "19" Or Me.txt_transporte = "20" Or Me.txt_transporte = "21" Or Me.txt_transporte = "22" Or Me.txt_transporte = "23" Or Me.txt_transporte = "24" Or Me.txt_transporte = "25" Then
                  If Me.txt_vehiculo <> "" Then
                     var_posible_placas = 1
                  Else
                     var_posible_placas = 0
                  End If
               Else
                  var_posible_placas = 1
               End If
               If var_posible_placas = 1 Then
                  cnnoracle_4.BeginTrans
                  rs.Open "select MAX(EMBARQUE) AS MAXIMO_EMBARQUE from XXVIA_TB_ENCABEZADO_EMBARQUES", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If rs.EOF Then
                     var_numero_embarque = 1
                  Else
                     var_numero_embarque = IIf(IsNull(rs!maximo_embarque), 0, rs!maximo_embarque) + 1
                  End If
                  rs.Close
                  Me.txt_numero = var_numero_embarque
                  rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  rs.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  'clnt.MSSoapInit var_webservice
                  'var_arreglo = clnt.crear_embarque(Me.txt_numero, Me.txt_numero, "000001_VTH_PROPIO_T_D2D")
                  'Set clnt = Nothing
            
                  objConn.Open var_conexion_oracle
                  '… Establecer conexión a la base de datos con el objeto objConn.
                  With objCmd
                       objConn.BeginTrans
                       .ActiveConnection = objConn
                       .CommandText = "XXVIA_PK_INTERFACES_OM.crear_embarque"
                       .CommandType = adCmdStoredProc
                                   
                       Set objParm = .CreateParameter("p_api_version_number", adNumeric, adParamInput, 50, 1)
                       .Parameters.Append objParm
                       
                       Set objParm = .CreateParameter("p_action_code", adVarChar, adParamInput, 100, "CREATE")
                      .Parameters.Append objParm
                     
                       Set objParm = .CreateParameter("p_name", adVarChar, adParamInput, 100, Me.txt_numero)
                       .Parameters.Append objParm
                  
                       Set objParm = .CreateParameter("p_trip_name", adVarChar, adParamInput, 100, Me.txt_numero)
                       .Parameters.Append objParm
                       
                       Set objParm = .CreateParameter("p_ship_method_code", adVarChar, adParamInput, 100, "000001_VTH_PROPIO_T_D2D")
                       .Parameters.Append objParm
                     
                       Set objParm = .CreateParameter("x_trip_id", adNumeric, adParamOutput, 50, 0)
                       .Parameters.Append objParm
                                   
                       Set objParm = .CreateParameter("x_trip_name", adVarChar, adParamOutput, 50, "")
                       .Parameters.Append objParm
                     
                       rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                       rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                       On Error GoTo SALIR:
                       .execute
                                       
                       VAR_X_TRIP_ID = .Parameters("x_trip_id").Value
                       var_x_trip_name = .Parameters("x_trip_name").Value
                       objConn.CommitTrans
                  End With
                  Set objConn = Nothing
                  Set objCmd = Nothing
            
                  If Not IsNumeric(VAR_X_TRIP_ID) Then
                     MsgBox var_arreglo(0)
                     MsgBox "A surgido un error en crear el embarque en ORACLE, intentelo nuevamente", vbOKOnly, "ATENCION"
                  Else
                     var_cadena = "insert into xxvia_tb_encabezado_embarques (EMBARQUE,               JAULA,              VEHICULO,               AGENTE,          FECHA_INICIO, CHAR_EMB_ESTATUS, CHOFER,  BLOQUEADO, BLOQUEADO_POR, TIPO_EMBARQUE, MAQUINA, USUARIO, ARREGLO_0, ARREGLO_1, ORGANIZACION, transporte) "
                     var_cadena = var_cadena + " values (" + CStr(var_numero_embarque) + ", " + txt_jaula + ", '" + txt_vehiculo + "', '" + txt_agente + "', sysdate, '',          '" + txt_chofer + "', 0, ''," + CStr(var_tipo_embarque) + ", '" + fun_NombrePc + "','" + var_clave_usuario_global + "','" + CStr(VAR_X_TRIP_ID) + "','" + var_x_trip_name + "'," + var_unidad_organizacional + ",'" + Me.txt_transporte + "')"
                     rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If var_bandera_asignacion = 1 Then
                        frmnumero_embarque.txt_embarque = var_numero_embarque
                     End If
                     var_agente_asignar = Me.txt_agente
                     var_embarque_asignar = var_numero_embarque
                     var_nombre_agente_asignar = Me.cmb_agentes.Text
                     Unload Me
                  End If
                  cnnoracle_4.CommitTrans
                  If var_bandera_asignacion = 0 Then
                  If IsNumeric(VAR_X_TRIP_ID) Then
                     frmoracle_asignar_pedidos_embarque.txt_embarque = var_embarque_asignar
                     frmoracle_asignar_pedidos_embarque.Show 1
                     Unload Me
                  End If
                  End If
               Else
                  MsgBox "Se debe de indicar las placas del vehiculo", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "Se debe de indicar el transporte", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Falta información", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Chofer incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Debe de indicar un vehiculo", vbOKOnly, "ATENCION"
   End If
   Exit Sub
SALIR:
   If Err.Number = -2147467259 Or Err.Number = -2147217900 Then
      
      'MsgBox Err.Description
      rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic

      Resume
   End If
   MsgBox "No se pudo crear el embarque", vbOKOnly, "ATENCION"
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 71 Then
      cmd_guardar_Click
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   'Me.txt_jaula = var_anden_global
   Me.txt_jaula = 1
   Me.frm_lista.Visible = False
   var_cadena_seguridad = ""
   Top = 2000
   Left = 2500
   var_cubicaje = 0
   rs.Open "select MAX(EMBARQUE) AS MAXIMO_EMBARQUE from XXVIA_TB_ENCABEZADO_EMBARQUES", cnnoracle_4, adOpenDynamic, adLockOptimistic
   If rs.EOF Then
      var_numero_embarque = 1
   Else
      var_numero_embarque = IIf(IsNull(rs!maximo_embarque), 0, rs!maximo_embarque) + 1
   End If
   rs.Close
   txt_numero = var_numero_embarque
   txt_fecha = Date
   rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   'rs.Open "select '', collector_id, name from ar_collectors order by name", cnnoracle_4, adOpenDynamic, adLockBatchOptimistic
   'rs.Open "SELECT  distinct '',  arc.collector_id , arc.name   FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, ar_collectors arc, ar_customer_profile_classes arcpc Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND arc.collector_id = hcp.collector_id AND arcpc.customer_profile_class_id = hcp.profile_class_id and site_use_code = 'BILL_TO' AND hcas.org_id = " + var_empresa + " order by arc.name", cnnoracle_4, adOpenDynamic, adLockOptimistic
   rs.Open " select distinct '', collector_id, name from xxvia_ar_collectors", cnnoracle_4, adOpenDynamic, adLockOptimistic
   Call RecsetToCombo(cmb_agentes.hwnd, rs, 2)
   rs.Close
   rs.Open "select * from tb_choferes order by vcha_cho_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
   Call RecsetToCombo(cmb_choferes.hwnd, rs, 1)
   rs.Close
   'txt_vehiculo.Enabled = False
   'txt_agente.Enabled = False
   'cmb_agentes.Enabled = False
   txt_jaula.Enabled = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_embarques)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_ventana = 1 Then
         Me.txt_vehiculo = lv_lista.selectedItem.SubItems(1)
         Me.txt_vehiculo.SetFocus
      End If
      If var_ventana = 2 Then
         Me.txt_transporte = lv_lista.selectedItem
         Me.txt_nombre_transporte = lv_lista.selectedItem.SubItems(1)
         Me.txt_transporte.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      Me.txt_vehiculo.SetFocus
   End If
   
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub txt_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_almacen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Trim(txt_almacen) <> "" Then
          rs.Open "select * from tb_almacenes where vcha_alm_almacen_id = '" + txt_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
          If Not rs.EOF Then
             cmb_almacenes.Text = rs!VCHA_ALM_NOMBRE
             cmb_almacenes.Enabled = False
             txt_almacen.Enabled = False
             rs.Close
             txt_vehiculo.Enabled = True
             txt_vehiculo.SetFocus
          Else
             txt_almacen = ""
             MsgBox "Almacen Incorrecto", vbOKOnly, "ATENCION"
             rs.Close
             cmb_almacenes.SetFocus
          End If
       End If
    End If
End Sub


Private Sub txt_agente_LostFocus()
      If IsNumeric(txt_agente) Then
         rs.Open "select * from ar_collectors where collector_id = " + txt_agente, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            cmb_agentes.Text = rs!Name
            'cmb_agentes.Enabled = False
            'txt_agente.Enabled = False
            rs.Close
            numero_items_rutas = 0
            txt_vehiculo.Enabled = True
            'txt_vehiculo.SetFocus
         Else
            txt_agente = ""
            MsgBox "Agente Incorrecto", vbOKOnly, "ATENCION"
            rs.Close
            'cmb_agentes.SetFocus
         End If
      End If
End Sub

Private Sub txt_chofer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      If Trim(txt_chofer) <> "" Then
         rs.Open "select * from tb_choferes where vcha_cho_chofer_id = '" + txt_chofer + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            cmb_choferes.Text = rs!vcha_cho_nombre
            'cmb_choferes.Enabled = False
            'txt_chofer.Enabled = False
            rs.Close
         Else
            txt_chofer = ""
            rs.Close
            MsgBox "Chofer Incorrecto", vbOKOnly, "ATENCION"
            'cmb_choferes.SetFocus
         End If
      End If
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_cubicaje_KeyPress(KeyAscii As Integer)
   'Me.cmd_guardar.SetFocus
End Sub

Private Sub txt_fecha_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_jaula_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If Trim(txt_jaula) <> "" Then
         rs.Open "select * from tb_jaulas where inte_jau_jaula_id = " + txt_jaula, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            rs.Close
            txt_agente.Enabled = True
            cmb_agentes.Enabled = True
         Else
            rs.Close
            txt_agente.Enabled = False
            cmb_agentes.Enabled = False
            MsgBox "La jaula no existe", vbOKOnly, "ATENCION"
         End If
      End If
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_transporte_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_ventana = 2
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_oracle_transportes", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!CLAVE)
            list_item.SubItems(1) = IIf(IsNull(rs!nombre), "", rs!nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TRANSPORTES"
      VAR_TIPO_LISTA = 100
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_nombre_transporte_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_transporte_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_ventana = 2
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_oracle_transportes", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!CLAVE)
            list_item.SubItems(1) = IIf(IsNull(rs!nombre), "", rs!nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TRANSPORTES"
      VAR_TIPO_LISTA = 100
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_transporte_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_transporte_LostFocus()
   If Trim(Me.txt_transporte) <> "" Then
      rs.Open "SELECT * FROM TB_ORACLE_TRANSPORTES WHERE CLAVE = '" + Me.txt_transporte + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_transporte = IIf(IsNull(rs!nombre), "", rs!nombre)
      Else
         MsgBox "Transporte incorrecto", vbOKOnly, "ATENCION"
         Me.txt_nombre_transporte = ""
         Me.txt_transporte = ""
      End If
      rs.Close
   End If
End Sub

Private Sub txt_vehiculo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_ventana = 1
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_transportes", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_TRN_TRANSPORTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_trn_nombre), "", rs!vcha_trn_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TRANSPORTES"
      VAR_TIPO_LISTA = 100
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_vehiculo_KeyPress(KeyAscii As Integer)
   Dim var_largo As Double
   Dim var_ancho As Double
   Dim var_alto As Double
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub

