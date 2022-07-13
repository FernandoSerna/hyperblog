VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_embarques 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Creación de embarques"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   1485
      TabIndex        =   17
      Top             =   750
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   18
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Volumen"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H000000FF&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   19
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8820
      Picture         =   "frmoracle_embarques.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frmoracle_embarques.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   4950
      Left            =   90
      TabIndex        =   8
      Top             =   390
      Width           =   9060
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   120
         Picture         =   "frmoracle_embarques.frx":073C
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   960
         Width           =   330
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   450
         Picture         =   "frmoracle_embarques.frx":083E
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   960
         Width           =   330
      End
      Begin VB.TextBox txt_volumen 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7665
         TabIndex        =   6
         Top             =   4215
         Width           =   1320
      End
      Begin VB.CommandButton cmd_asignar_maquinas 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8670
         Picture         =   "frmoracle_embarques.frx":0A54
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Asignar"
         Top             =   4575
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox txt_agente 
         Height          =   315
         Left            =   6000
         TabIndex        =   22
         Top             =   585
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txt_numero 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1065
         TabIndex        =   0
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txt_fecha 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3015
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txt_vehiculo 
         Height          =   315
         Left            =   1065
         TabIndex        =   7
         Top             =   4545
         Width           =   1320
      End
      Begin VB.TextBox txt_chofer 
         Height          =   315
         Left            =   1065
         TabIndex        =   2
         Top             =   3855
         Width           =   1320
      End
      Begin VB.ComboBox cmb_choferes 
         Height          =   315
         Left            =   2415
         TabIndex        =   3
         Top             =   3855
         Width           =   6600
      End
      Begin VB.TextBox txt_transporte 
         Height          =   315
         Left            =   1065
         TabIndex        =   4
         Top             =   4215
         Width           =   1320
      End
      Begin VB.TextBox txt_nombre_transporte 
         Height          =   315
         Left            =   2400
         TabIndex        =   5
         Top             =   4200
         Width           =   4440
      End
      Begin MSComctlLib.ListView lv_agentes 
         Height          =   2400
         Left            =   60
         TabIndex        =   11
         Top             =   1290
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   4233
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
            Text            =   "Clave"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   12347
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Volumen:"
         Height          =   195
         Left            =   6975
         TabIndex        =   25
         Top             =   4275
         Width           =   660
      End
      Begin VB.Label lab_estados 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   21
         Top             =   660
         Width           =   600
      End
      Begin VB.Label lab_estados 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Index           =   4
         Left            =   2340
         TabIndex        =   20
         Top             =   660
         Width           =   495
      End
      Begin VB.Label lbl_anden 
         AutoSize        =   -1  'True
         Caption         =   "8"
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
         Left            =   1425
         TabIndex        =   16
         Top             =   240
         Width           =   165
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Estación:"
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
         Left            =   150
         TabIndex        =   15
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label lab_estados 
         AutoSize        =   -1  'True
         Caption         =   "Placas:"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   14
         Top             =   4605
         Width           =   525
      End
      Begin VB.Label lab_estados 
         AutoSize        =   -1  'True
         Caption         =   "Chofer:"
         Height          =   195
         Index           =   5
         Left            =   150
         TabIndex        =   13
         Top             =   3915
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Transporte:"
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   4275
         Width           =   810
      End
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   90
      TabIndex        =   24
      Top             =   270
      Width           =   9060
   End
End
Attribute VB_Name = "frmoracle_embarques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objConn As New ADODB.Connection
Dim objCmd As New ADODB.Command
Dim objParm As ADODB.Parameter
Dim var_contador_porcentaje As Integer
Dim var_cubicaje As Double
Dim var_ventana As Integer
Private Sub cmb_choferes_Click()
   rsaux.Open "SELECT * FROM TB_CHOFERES WHERE VCHA_CHO_NOMBRE = '" + Me.cmb_choferes + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rsaux.EOF Then
      'txt_chofer = Obtener_llave(cnn, rsaux, "TB_CHOFERES", "VCHA_CHO_NOMBRE", cmb_choferes, 0, "T")
      txt_chofer = IIf(IsNull(rsaux!vcha_cho_chofer_id), "", rsaux!vcha_cho_chofer_id)
   End If
   rsaux.Close
End Sub

Private Sub cmb_choferes_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub cmd_asignar_maquinas_Click()
   frmoracle_asignar_maquinas.Show 1
End Sub

Private Sub cmd_guardar_Click()
   Dim clnt As New SoapClient30
   Dim var_arreglo() As String
   Dim var_posible As Boolean
   Dim var_posible_chofer As Boolean
   Dim var_primer_agente As String
   Dim var_cadena_agentes As String
   var_posible = True
   var_posible_chofer = True
   Me.txt_agente = ""
   var_cadena_agentes = ""
   For var_j = 1 To lv_agentes.ListItems.Count
       lv_agentes.ListItems.Item(var_j).Selected = True
       If lv_agentes.selectedItem.SubItems(2) = "*" Then
          If Me.txt_agente = "" Then
             var_primer_agente = lv_agentes.selectedItem
             Me.txt_agente = "'" + lv_agentes.selectedItem
             var_cadena_agentes = Me.lv_agentes.selectedItem
          Else
             Me.txt_agente = Me.txt_agente + "','" + lv_agentes.selectedItem
             var_cadena_agentes = var_cadena_agentes + "," + Me.lv_agentes.selectedItem
          End If
       End If
   Next var_j
   var_agente_asignar = Me.txt_agente + "'"
   If var_empresa = "02" Or var_empresa = "03" Then
      If Me.txt_vehiculo <> "" Then
         rs.Open "SELECT * FROM TB_TRANSPORTES WHERE VCHA_TRN_NOMBRE = '" + Me.txt_vehiculo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_posible = True
            var_volumen_transporte = IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN)
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
         If Trim(txt_numero) <> "" And Trim(Me.txt_agente) <> "" And Trim(Me.lbl_anden) <> "" Then
            If Me.txt_transporte <> "" Then
               rs.Open "SELECT * FROM TB_oracle_TRANSPORTES WHERE clave = '" + Me.txt_transporte + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_volumen_transporte = IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN)
               End If
               rs.Close
            
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
                     var_cadena = "insert into xxvia_tb_encabezado_embarques (EMBARQUE,               JAULA,              VEHICULO,               AGENTE,          FECHA_INICIO, CHAR_EMB_ESTATUS, CHOFER,  BLOQUEADO, BLOQUEADO_POR, TIPO_EMBARQUE, MAQUINA, USUARIO, ARREGLO_0, ARREGLO_1, ORGANIZACION, transporte, agentes) "
                     var_cadena = var_cadena + " values (" + CStr(var_numero_embarque) + ", " + Me.lbl_anden + ", '" + txt_vehiculo + "', '" + var_primer_agente + "', sysdate, '',          '" + txt_chofer + "', 0, ''," + CStr(var_tipo_embarque) + ", '" + fun_NombrePc + "','" + var_clave_usuario_global + "','" + CStr(VAR_X_TRIP_ID) + "','" + var_x_trip_name + "'," + var_unidad_organizacional + ",'" + Me.txt_transporte + "','" + var_cadena_agentes + "')"
                     rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If var_bandera_asignacion = 1 Then
                        frmnumero_embarque.txt_embarque = var_numero_embarque
                     End If
                     var_agente_asignar = Me.txt_agente
                     var_embarque_asignar = var_numero_embarque
                     frmoracle_asignar_maquinas.Show 1
                     'var_nombre_agente_asignar = Me.cmb_agentes.Text
                     Unload Me
                  End If
                  cnnoracle_4.CommitTrans
                  'If var_bandera_asignacion = 0 Then
                  If IsNumeric(VAR_X_TRIP_ID) Then
                     frmoracle_asignar_pedidos_embarque.txt_embarque = var_embarque_asignar
                     frmoracle_asignar_pedidos_embarque.Show 1
                     Unload Me
                  End If
                  'End If
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

Private Sub cmd_ninguno_Click()
    
      Dim list_item As ListItem
      
      For i = 1 To Me.lv_agentes.ListItems.Count
      Me.lv_agentes.ListItems.Item(i).Selected = True
      If lv_agentes.selectedItem.SubItems(2) = "*" Then
          lv_agentes.ListItems.Item(i).SubItems(2) = " "
          lv_agentes.ListItems.Item(i).Bold = False
          lv_agentes.ListItems.Item(i).ForeColor = &H80000012
          lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = False
          lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = False
          lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
          lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_agentes.ListItems.Item(i).SubItems(2) = "*"
         lv_agentes.ListItems.Item(i).Bold = True
         lv_agentes.ListItems.Item(i).ForeColor = &H8000&
         lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
         lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
      End If
      Next i
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmd_todos_Click()
      Dim list_item As ListItem
      
      For i = 1 To Me.lv_agentes.ListItems.Count
      Me.lv_agentes.ListItems.Item(i).Selected = True
      If lv_agentes.selectedItem.SubItems(2) = "*" Then
          lv_agentes.ListItems.Item(i).SubItems(2) = " "
          lv_agentes.ListItems.Item(i).Bold = False
          lv_agentes.ListItems.Item(i).ForeColor = &H80000012
          lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = False
          lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = False
          lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
          lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_agentes.ListItems.Item(i).SubItems(2) = "*"
         lv_agentes.ListItems.Item(i).Bold = True
         lv_agentes.ListItems.Item(i).ForeColor = &H8000&
         lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
         lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
      End If
      Next i

End Sub

Private Sub Form_Load()
   Me.lbl_anden = var_anden_global
   Me.frm_lista.Visible = False
   var_cadena_seguridad = ""
   'Top = 2000
   'Left = 2500
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
   If var_cambio_embarque = 1 Then
       
      rs.Open "SELECT * FROM tb_oracle_pedidos_asignados_embarques WHERE PEDIDO = '" + var_pedido_cambio_embarque + "'", cnn, adOpenDynamic, adLockOptimistic
      
   Else
      rs.Open "select distinct agente from tb_oracle_pedidos_asignados_embarques where agente is not null and embarque = 0", cnn, adOpenDynamic, adLockOptimistic
   End If
   var_cadena_agentes = ""
   While Not rs.EOF
         If var_cadena_agentes = "" Then
            var_cadena_agentes = CStr(rs!Agente)
         Else
            var_cadena_agentes = var_cadena_agentes + "," + CStr(rs!Agente)
         End If
         rs.MoveNext
   Wend
   rs.Close
   If var_cadena_agentes <> "" Then
      rs.Open " select distinct collector_id, name from xxvia_ar_collectors where collector_id in (" + var_cadena_agentes + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_agentes.ListItems.Add(, , rs!collector_id)
            list_item.SubItems(1) = IIf(IsNull(rs!Name), "", rs!Name)
            rs.MoveNext
      Wend
      rs.Close
   Else
      MsgBox "No hay pedidos por surtir", vbOKOnly, "ATENCION"
   End If
   rs.Open "select * from XXVIA_tb_choferes WHERE NVL(LICENCIA,' ') <> ' ' order by NOMBRE", cnnoracle_4, adOpenDynamic, adLockBatchOptimistic
   Call RecsetToCombo(cmb_choferes.hwnd, rs, 1)
   rs.Close
   
   
   
      Dim numero_lineas As Integer
      Dim numero_seleccionado1 As Integer
      Dim numero_seleccionado2 As Integer
      Dim primera_vez As Boolean
      Dim segunda_vez As Boolean
      Dim i As Integer
      Dim n As Integer
      For i = 1 To Me.lv_agentes.ListItems.Count
          Me.lv_agentes.ListItems.Item(i).Selected = True
          Me.lv_agentes.selectedItem.SubItems(2) = "*"
            lv_agentes.ListItems.Item(i).Bold = True
            lv_agentes.ListItems.Item(i).ForeColor = &H8000&
            lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
            lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
            lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
            lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
      Next i
      lv_agentes.Refresh
   
   
   
   
End Sub

Private Sub lv_agentes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_agentes, ColumnHeader)
End Sub

Private Sub lv_agentes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim numero_lineas As Integer
      Dim numero_seleccionado1 As Integer
      Dim numero_seleccionado2 As Integer
      Dim primera_vez As Boolean
      Dim segunda_vez As Boolean
      Dim i As Integer
      Dim n As Integer
      Dim list_item As ListItem
      i = lv_agentes.selectedItem.Index
      If lv_agentes.selectedItem.SubItems(2) = "*" Then
          lv_agentes.ListItems.Item(i).SubItems(2) = " "
          lv_agentes.ListItems.Item(i).Bold = False
          lv_agentes.ListItems.Item(i).ForeColor = &H80000012
          lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = False
          lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = False
          lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
          lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_agentes.ListItems.Item(i).SubItems(2) = "*"
         lv_agentes.ListItems.Item(i).Bold = True
         lv_agentes.ListItems.Item(i).ForeColor = &H8000&
         lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
         lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
      End If
      lv_agentes.Refresh
   End If
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
         Me.txt_volumen = lv_lista.selectedItem.SubItems(2)
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

Private Sub txt_nombre_transporte_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_ventana = 2
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_oracle_transportes", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!clave)
            list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE), "", rs!NOMBRE)
            list_item.SubItems(2) = IIf(IsNull(rs!VOLUMEN), "", rs!VOLUMEN)
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

Private Sub txt_transporte_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_ventana = 2
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_oracle_transportes", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!clave)
            list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE), "", rs!NOMBRE)
            list_item.SubItems(2) = IIf(IsNull(rs!VOLUMEN), "0", rs!VOLUMEN)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TRANSPORTES"
      VAR_TIPO_LISTA = 100
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 2800
         lv_lista.ColumnHeaders(3).Width = 1400
      Else
         lv_lista.ColumnHeaders(2).Width = 3000.18
         lv_lista.ColumnHeaders(3).Width = 1400
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
         Me.txt_nombre_transporte = IIf(IsNull(rs!NOMBRE), "", rs!NOMBRE)
         Me.txt_volumen = IIf(IsNull(rs!VOLUMEN), "", rs!VOLUMEN)
      Else
         MsgBox "Transporte incorrecto", vbOKOnly, "ATENCION"
         Me.txt_nombre_transporte = ""
         Me.txt_transporte = ""
         Me.txt_volumen = ""
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
         lv_lista.ColumnHeaders(3).Width = 0
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
         lv_lista.ColumnHeaders(3).Width = 0
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_vehiculo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub

