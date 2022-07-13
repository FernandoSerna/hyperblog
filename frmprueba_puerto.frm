VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmprueba_puerto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Prototipo de bascula"
   ClientHeight    =   6570
   ClientLeft      =   -15
   ClientTop       =   330
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_cerrar 
      Height          =   375
      Left            =   4200
      Picture         =   "frmprueba_puerto.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txt_cantidad 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3840
      TabIndex        =   2
      Top             =   600
      Width           =   1890
   End
   Begin VB.TextBox txt_caja 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3000
      TabIndex        =   1
      Top             =   97
      Width           =   1095
   End
   Begin VB.TextBox txt_pedido 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1080
      TabIndex        =   0
      Top             =   97
      Width           =   1095
   End
   Begin VB.TextBox txt_foco 
      Height          =   285
      Left            =   9000
      TabIndex        =   3
      Top             =   840
      Width           =   150
   End
   Begin VB.TextBox txt_codigo 
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   630
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "bascula 2"
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Top             =   6600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   360
      TabIndex        =   10
      Top             =   5640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   5760
      Top             =   960
   End
   Begin VB.CommandButton CMDENVIAR 
      Caption         =   "ENVIAR"
      Height          =   405
      Left            =   2700
      TabIndex        =   9
      Top             =   6570
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.CommandButton cmdcon 
      Caption         =   "CONECTAR"
      Height          =   465
      Left            =   720
      TabIndex        =   8
      Top             =   5880
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmprueba_puerto.frx":0102
      Left            =   2220
      List            =   "frmprueba_puerto.frx":011E
      TabIndex        =   7
      Top             =   6000
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtenviar 
      Height          =   915
      Left            =   360
      TabIndex        =   5
      Top             =   6000
      Visible         =   0   'False
      Width           =   1665
   End
   Begin MSComctlLib.ListView lv_entradas 
      Height          =   4530
      Left            =   120
      TabIndex        =   18
      Top             =   1920
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   7990
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
         Text            =   "Código"
         Object.Width           =   2478
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   9701
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Pedido"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Enviado"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Localizador"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Peso:"
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
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
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
      Left            =   120
      TabIndex        =   16
      Top             =   720
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Caja:"
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
      Left            =   2280
      TabIndex        =   14
      Top             =   157
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Pedido:"
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
      Left            =   120
      TabIndex        =   13
      Top             =   157
      Width           =   915
   End
   Begin VB.Label lbl_bascula 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   960
      TabIndex        =   12
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblmostrar 
      Height          =   435
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   2325
   End
End
Attribute VB_Name = "frmprueba_puerto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim textout, textin, texto As String
Dim objConn As New ADODB.Connection
Dim objCmd As New ADODB.Command
Dim objParm As ADODB.Parameter
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter
Dim var_cantidad_leida As Double
Dim var_descripcion_articulo As String

   
   


Private Sub cmd_cerrar_Click()
Dim var_pedido As Integer
var_pedido = CDbl(Me.txt_pedido)
                  rsaux13.Open "select * from TB_ORACLE_PESOS_aRTICULOS where pedido = " + CStr(var_pedido) + " and caja = " + Me.txt_caja + " and codigo = 'ULTIMO'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux13.EOF Then
                     rsaux14.Open "UPDATE TB_ORACLE_PESOS_aRTICULOS SET PESO = " + Me.lbl_bascula + " where pedido = " + CStr(var_pedido) + " and caja = " + Me.txt_caja + " and codigo = 'ULTIMO'", cnn, adOpenDynamic, adLockOptimistic
                  Else
                     rsaux14.Open "INSERT INTO TB_ORACLE_PESOS_aRTICULOS (PEDIDO, CAJA, CODIGO, PESO) VALUES (" + CStr(var_pedido) + ", " + Me.txt_caja + ",'ULTIMO'," + CStr(CDbl(Me.lbl_bascula)) + ")", cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux13.Close
                        rsaux14.Open "select * from tb_oracle_pesos_articulos  where pedido = " + CStr(var_pedido) + " and caja = " + Me.txt_caja + " order by consecutivo", cnn, adOpenDynamic, adLockOptimistic
                        var_peso = 0
                        var_anterior = 0
                        While Not rsaux14.EOF
                              var_peso = rsaux14!PESO
                              rsaux14.MoveNext
                              If Not rsaux14.EOF Then
                                 var_peso = rsaux14!PESO - var_peso
                                 If var_peso <= 0 Then
                                    VAR_POSIBLE_PESO_rEAL = 0
                                 End If
                                 rsaux14.MovePrevious
                                 rsaux15.Open "UPDATE TB_ORACLE_PESOS_ARTICULOS SET PESO_rEAL = " + CStr(var_peso) + " WHERE CONSECUTIVO = " + CStr(rsaux14!CONSECUTIVO), cnn, adOpenDynamic, adLockOptimistic
                                 rsaux14.MoveNext
                              End If
                        Wend
                        rsaux14.Close
                        'MsgBox " select * from TB_ORACLE_PESOS_ARTICULOS WHERE PEDIDO = " + CStr(var_pedido) + " AND CAJA = " + Me.txt_caja + " and codigo <> 'ULTIMO'"
                        rsaux14.Open " select * from TB_ORACLE_PESOS_ARTICULOS WHERE PEDIDO = " + CStr(var_pedido) + " AND CAJA = " + Me.txt_caja + " and codigo <> 'ULTIMO'", cnn, adOpenDynamic, adLockOptimistic
                        While Not rsaux14.EOF
                              strconsulta = "select * from XXVIA_SYSTEM_ITEMS_B where organization_id = ? and segment1 = ? "
                              With comandoORA
                                   .ActiveConnection = cnnoracle_4
                                   .CommandType = adCmdText
                                   .CommandText = strconsulta
                                   Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(var_unidad_organizacional))
                                   .Parameters.Append parametro
                                   Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rsaux14!codigo)
                                   .Parameters.Append parametro
                              End With
                              Set rsaux15 = comandoORA.execute
                              Set comandoORA = Nothing
                              Set parametro = Nothing
         
                              If Not rsaux15.EOF Then
                                 rsaux9.Open "UPDATE TB_ORACLE_PESOS_ARTICULOS SET PESO_SISTEMA = " + CStr(IIf(IsNull(rsaux15!UNIT_WEIGHT), 0, rsaux15!UNIT_WEIGHT)) + " WHERE CONSECUTIVO = " + CStr(rsaux14!CONSECUTIVO), cnn, adOpenDynamic, adLockOptimistic
                              End If
                              rsaux15.Close
                              rsaux14.MoveNext
                        Wend
                        rsaux14.Close
                        
                        VAR_POSIBLE_PESO_rEAL = 1
                                                
                        If VAR_POSIBLE_PESO_rEAL = 1 Then
                           'rsaux.Open "select peso form tb_"
                           rsaux.Open "SELECT SUM(cANTIDAD) AS CANTIDAD FROM tb_codigos_prueba_bascula WHERE PEDIDO = " + Me.txt_pedido + " AND CAJA = " + Me.txt_caja, cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux.EOF Then
                              var_cantidad_caja = IIf(IsNull(rsaux!cantidad), 0, rsaux!cantidad)
                           Else
                              var_cantidad_caja = 0
                           End If
                           rsaux.Close
                        
                           rsaux.Open "SELECT SUM(cANTIDAD) AS CANTIDAD, SUM(PESO_REAL) AS PESO_REAL, SUM(PESO_SISTEMA) AS PESO_SISTEMA FROM TB_ORACLE_PESOS_aRTICULOS WHERE PEDIDO = " + Me.txt_pedido + " AND CAJA = " + Me.txt_caja, cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux.EOF Then
                              VAR_CANTIDAD_PESO = IIf(IsNull(rsaux!cantidad), 0, rsaux!cantidad)
                              VAR_PESO_rEAL = IIf(IsNull(rsaux!PESO_REAL), 0, rsaux!PESO_REAL)
                              VAR_PESO_SISTEMA = IIf(IsNull(rsaux!PESO_SISTEMA), 0, rsaux!PESO_SISTEMA)
                           Else
                              VAR_CANTIDAD_PESO = 0
                              VAR_PESO_rEAL = 0
                              VAR_PESO_SISTEMA = 0
                           End If
                           rsaux.Close
                        
                           var_posible = 0
                           If var_cantidad_caja = VAR_CANTIDAD_PESO Then
                              If Round(CDbl(VAR_PESO_rEAL), 2) = Round(CDbl(Me.lbl_bascula), 2) Then
                                 var_posible = 1
                              Else
                                 MsgBox "Existe diferencia entre el peso leido y el peso en bascula", vbOKOnly, "ATENCION"
                              End If
                           Else
                              MsgBox "Hay diferencia de piezas entre las piezas leidas y las piezas en bascula", vbOKOnly, "ATENCION"
                           End If
                        Else
                           var_posible = 0
                           MsgBox "Existe variación en los pesos leidos y el de la bascula"
                        End If
                        
                        If var_posible = 1 Then
                           MsgBox "SE IMPRIME ETIQUETA", vbOKOnly, "ATENCION"
                        Else
                           frmmensaje.lbl_articulo = ""
                           frmmensaje.lbl_mensaje = "No existe movimiento de peso anterior"
                           frmmensaje.Show 1
                        End If
                        
                        
                        

End Sub

Private Sub cmdcon_Click()
'On Error GoTo SALIR
   If cmdcon.Caption = "CONECTAR" Then
   
   puerto.CommPort = Val(Me.Combo1.ListIndex + 1)
   puerto.PortOpen = True
   Me.CMDENVIAR.Value = True
   Me.Timer1.Enabled = True
   cmdcon.Caption = "DESCONECTAR"
Else
   cmdcon.Caption = "DESCONECTAR"
   Timer1.Enabled = False
   Me.CMDENVIAR.Visible = False
   puerto.PortOpen = False
   cmdcon.Caption = "CONECTAR"
End If
Exit Sub
SALIR:
MsgBox "ERROR"
End Sub

Private Sub Command1_Click()

         strconsulta = "select * from XXVIA_TB_PESOS_BASCULA where NAME_COMPUTER = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, fun_NombrePc)
              .Parameters.Append parametro
         End With
         Set rs = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         If Not rs.EOF Then
         Me.txtenviar = CStr(rs!Weight)
         End If

End Sub

Private Sub Command2_Click()
Me.Timer1.Enabled = True
End Sub

Private Sub Form_Load()
 Me.Timer1.Enabled = True
   rs.Open "select * from tb_oracle_maquinas where maquina = '" + fun_NombrePc + "'", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_puerto = IIf(IsNull(rs!COM_BASCULA), 0, rs!COM_BASCULA)
      If var_puerto > 0 Then
         x = Shell(App.Path + "/puerto.exe ")
         Me.Timer1.Enabled = True
      End If
   Else
      Me.Timer1.Enabled = False
   End If
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub Timer1_Timer()
On Error GoTo SALIR:
   
     VAR_ZZ = 0
     If var_z = 0 Then
         var_maquina_bascula = fun_NombrePc
         strconsulta = "select * from XXVIA_TB_PESOS_BASCULA where NAME_COMPUTER = '" + var_maquina_bascula + "'"
         rs_bascula.Open strconsulta, cnn, adOpenDynamic, adLockOptimistic
         If Not rs_bascula.EOF Then
            If IsNumeric(rs_bascula!Weight) Then
               Me.lbl_bascula = CStr(rs_bascula!Weight)
            Else
               Me.lbl_bascula = "0.00"
            End If
         Else
            Me.lbl_bascula = "ERROR"
         End If
         rs_bascula.Close
     
     
     Else
         If rs_bascula.State = 1 Then
            rs_bascula.Close
         End If
         strconsulta = "select * from XXVIA_TB_PESOS_BASCULA where NAME_COMPUTER = ?"
         var_maquina_bascula = fun_NombrePc
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_maquina_bascula)
              .Parameters.Append parametro
         End With
         Set rs_bascula = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         If Not rs_bascula.EOF Then
         Me.lbl_bascula = CStr(rs_bascula!Weight)
         End If
         rs_bascula.Close
   End If
   Exit Sub
   
SALIR:
   Me.lbl_bascula = "0.00"

End Sub


Private Sub txt_caja_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      rs.Open "SELECT * FROM tb_codigos_prueba_bascula WHERE PEDIDO = " + Me.txt_pedido + " AND CAJA = " + Me.txt_caja, cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         While Not rs.EOF
               Set list_item = lv_entradas.ListItems.Add(, , Trim(rs!codigo))
               list_item.SubItems(1) = rs!DESCRIPCION
               list_item.SubItems(2) = rs!cantidad
               rs.MoveNext
         Wend
      End If
      Me.txt_codigo.SetFocus
   End If
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Dim var_recontable As Integer
   Dim var_cantidad_caja As Integer
   Dim var_caja As String
   Dim var_estatus_caja As String
   Dim var_posible_caja As Integer
   Dim var_codigo_caja As String
   txt_codigo = Trim(txt_codigo)
   If KeyAscii = 13 Then
      var_localizador_subinventario = " "
      var_encontro = 0
      var_cantidad_leida = 1
      var_cantidad_leida_seg_nivel = 1
      var_posible_caja = 1
      var_cantidad_leida_caja = 0
      Dim var_tela As String
      var_tela = ""
      For var_j = 1 To Len(Me.txt_codigo)
          If Mid(Me.txt_codigo, var_j, 1) = "-" Then
             var_tela = var_tela + Mid(Me.txt_codigo, var_j, 1)
          End If
      Next var_j
      var_estatus_caja = ""
      var_codigo_caja = ""
      If Mid(Me.txt_codigo, 1, 2) = "CA" Or var_tela = "---" Then
         rs.Open "SELECT * FROM XXVIA_TB_CAJAS_PROD WHERE vcha_caj_caja_id = '" + UCase(Me.txt_codigo) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            If rsaux8.State = 1 Then
               rsaux8.Close
            End If
            rsaux8.Open "select * from xxvia_system_items_b where segment1 = '" + IIf(IsNull(rs!VCHA_ART_ARTICULO_ID), "", rs!VCHA_ART_ARTICULO_ID) + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rsaux8.EOF Then
               var_estatus_caja = IIf(IsNull(rs!vcha_caj_staus), "", rs!vcha_caj_staus)
               'var_estatus_caja = "A"
               
               If var_estatus_caja <> "S" Then
                  var_codigo_caja = Me.txt_codigo
                  Me.txt_codigo = IIf(IsNull(rs!VCHA_ART_ARTICULO_ID), "", rs!VCHA_ART_ARTICULO_ID)
                  var_salida_masiva = IIf(IsNull(rsaux8!attribute10), "N", rsaux8!attribute10)
                  var_cantidad_leida_caja = rs!numb_caj_cantidad
                  strconsulta = "update xxvia_Tb_cajas_prod set vcha_caj_staus = 'S' where vcha_caj_Caja_id = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_codigo_caja)
                       .Parameters.Append parametro
                  End With
                  Set rsaux9 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
               
               Else
                  var_posible_caja = 0
               End If
            End If
            rsaux8.Close
         End If
         rs.Close
      End If
      If var_posible_caja = 1 Then
      var_cadena = "SELECT A.INVENTORY_ITEM_ID, B.DESCRIPTION, cross_reference, b.segment1, NVL(a.description,'') AS localizador, B.UNIT_WEIGHT FROM (select INVENTORY_ITEM_ID, description, cross_reference from mtl_cross_references_b) A, (select inventory_item_id, DESCRIPTION, organization_id, segment1, UNIT_WEIGHT from xxvia_system_items_b) B Where a.inventory_item_id = B.inventory_item_id AND B.organization_id = " + var_unidad_organizacional + " AND CROSS_REFERENCE       = '" + Me.txt_codigo + "'"
      x = 0
      If x = 0 Then
         strconsulta = "SELECT A.INVENTORY_ITEM_ID, B.DESCRIPTION, cross_reference, b.segment1, NVL(a.description,'') AS localizador, B.UNIT_WEIGHT, a.cantidad FROM (select INVENTORY_ITEM_ID, description, cross_reference, nvl(attribute1,1) as cantidad from mtl_cross_references_b) A, (select inventory_item_id, DESCRIPTION, organization_id, segment1, UNIT_WEIGHT from xxvia_system_items_b) B Where a.inventory_item_id = B.inventory_item_id AND B.organization_id = ? AND CROSS_REFERENCE = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
              .Parameters.Append parametro
         End With
         Set rsaux8 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
      End If
      var_cantidad_leida = 1
      If Not rsaux8.EOF Then
         var_peso = IIf(IsNull(rsaux8!UNIT_WEIGHT), 0, rsaux8!UNIT_WEIGHT)
         var_cantidad_leida_seg_nivel = IIf(IsNull(rsaux8!cantidad), 1, rsaux8!cantidad)
         'If IIf(IsNull(rsaux8!localizador), "", rsaux8!localizador) <> "" Then
         '   var_localizador_subinventario = txt_almacen + IIf(IsNull(rsaux8!localizador), "", rsaux8!localizador)
         '   If var_localizador_subinventario <> "" Then
         '       Me.txt_codigo = rsaux8!SEGMENT1
         '   End If
         'Else
            Me.txt_codigo = rsaux8!SEGMENT1
         'End If
      End If
      rsaux8.Close
      
      
      
      
      
      
      
      
      'rsaux8.Open "SELECT A.INVENTORY_ITEM_ID, B.DESCRIPTION, cross_reference, b.segment1, nvl(a.description,'') as localizador FROM mtl_cross_references_v A, xxvia_system_items_b B WHERE A.inventory_item_id = B.inventory_item_id AND B.organization_id = " + var_unidad_organizacional + " AND CROSS_REFERENCE = '" + Me.txt_codigo + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      'If Not rsaux8.EOF Then
      '   Me.txt_codigo = rsaux8!SEGMENT1
      'Else
      '   rsaux9.Open "select * from xxvia_system_items_b where segment1 = '" + Me.txt_codigo + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
      '   If Not rsaux9.EOF Then
      '      Me.txt_codigo = rsaux9!SEGMENT1
      '   Else
      '      Me.txt_codigo = ""
      '   End If
      '   rsaux9.Close
      'End If
      'rsaux8.Close
      
      If Trim(Me.txt_codigo) <> "" Then
         rsaux8.Open "select * from xxvia_system_items_b where segment1 = '" + Me.txt_codigo + "' and organization_id = " + var_unidad_organizacional, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rsaux8.EOF Then
            
            
            
            var_unidad_medida = rsaux8!PRIMARY_UOM_CODE
            var_descripcion_articulo = rsaux8!Description
            var_inventory_item_id = rsaux8!inventory_item_id
            var_salida_masiva = IIf(IsNull(rsaux8!attribute10), "N", rsaux8!attribute10)
            var_clave_lista_precios = 9007
            'var_cadena = "select * from  qp_secu_list_headers_v a, qp_list_lines_v b Where a.list_header_id = b.list_header_id and  B.product_attr_value = " + CStr(var_inventory_item_id) + " AND a.list_header_id = " + CStr(var_clave_lista_precios) + " and   product_attr_val_disp = '" + Me.txt_codigo + "' AND NVL(OPERAND,0) > 0"
            var_cadena = "select * from  qp_secu_list_headers_v a, qp_list_lines_v b Where a.list_header_id = b.list_header_id and  B.product_attr_value = " + CStr(var_inventory_item_id) + " AND a.list_header_id = " + CStr(var_clave_lista_precios) + " and   product_attr_val_disp = '" + Me.txt_codigo + "'"
            rsaux10.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            'rsaux10.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            'If Not rsaux10.EOF Then
            If var_cadena <> "" Then
               'If IIf(IsNull(rsaux10!OPERAND), 0, rsaux10!OPERAND) > 0 Then
                If var_cadena <> "" Then
                  If var_unidad_organizacional = "900" Then
                     var_salida_masiva = "Y"
                  End If
                  'If var_unidad_organizacional = "85" Or var_unidad_organizacional = "94" Then
                  '   var_salida_masiva = "Y"
                  'End If
                  If var_clave_usuario_global = "U0000000635" Then
                     var_salida_masiva = "Y"
                  End If
                  If var_cantidad_leida_caja = 0 Then
                     If var_cantidad_leida_seg_nivel = 1 Then
                        If var_salida_masiva = "Y" Then
                           var_codigo_global = Me.txt_codigo
                           frmoracle_cantidad.Show 1
                           var_cantidad_leida = var_cantidad_global
                           Me.txt_codigo = var_codigo_global
                        Else
                           var_cantidad_leida = 1
                        End If
                        Me.txt_foco.Enabled = True
                        Me.txt_foco.SetFocus
                     Else
                        var_cantidad_leida = var_cantidad_leida_seg_nivel
                        Me.txt_foco.Enabled = True
                        Me.txt_foco.SetFocus
                     End If
                  Else
                     var_cantidad_leida = var_cantidad_leida_caja
                     Me.txt_foco.Enabled = True
                     Me.txt_foco.SetFocus
                  End If
               Else
                  txt_codigo = ""
                  frmmensaje.lbl_mensaje = "El artículo no tiene precio"
                  frmmensaje.Show
                  If var_codigo_caja <> "" Then
                     strconsulta = "update xxvia_Tb_cajas_prod set vcha_caj_staus = 'A' where vcha_caj_Caja_id = ?"
                     With comandoORA
                          .ActiveConnection = cnnoracle_4
                          .CommandType = adCmdText
                          .CommandText = strconsulta
                          Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_codigo_caja)
                          .Parameters.Append parametro
                     End With
                     Set rsaux9 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                  End If
               
               
               End If
            Else
               txt_codigo = ""
               frmmensaje.lbl_mensaje = "El artículo no se encuentra en la lista de precios del cliente"
               frmmensaje.Show
               If var_codigo_caja <> "" Then
                  strconsulta = "update xxvia_Tb_cajas_prod set vcha_caj_staus = 'A' where vcha_caj_Caja_id = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_codigo_caja)
                       .Parameters.Append parametro
                  End With
                  Set rsaux9 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
               End If
            
            End If
            'rsaux10.Close
         Else
            txt_codigo = ""
            frmmensaje.lbl_mensaje = "Error en código"
            frmmensaje.Show
         End If
         rsaux8.Close
      Else
         If var_localizador = 2 And Me.txt_codigo = "" Then
            txt_codigo = ""
            frmmensaje.lbl_mensaje = "El artículo necesita localizador"
            frmmensaje.Show
         Else
            txt_codigo = ""
            frmmensaje.lbl_mensaje = "El artículo no existe"
            frmmensaje.Show
         End If
      End If
      Else
         Me.txt_codigo = ""
         frmmensaje.lbl_mensaje = "La caja ya habia sido leida"
         frmmensaje.Show
      End If
   End If
   End If
End Sub

Private Sub txt_foco_GotFocus()
   If IsNumeric(Me.txt_pedido.Text) Then
      If IsNumeric(Me.txt_caja.Text) Then
         Sleep 3000
         If rs.State = 1 Then
            rs.Close
         End If
         rs.Open "select PESO from TB_ORACLE_PESOS_aRTICULOS where consecutivo =  (select max(consecutivo) from TB_ORACLE_PESOS_aRTICULOS where PEDIDO = " + Me.txt_pedido + " and caja  = " + Me.txt_caja + ")", cnn, adOpenDynamic, adLockOptimistic
         If rs.EOF Then
            var_peso = 0
         Else
            var_peso = rs!PESO
         End If
         rs.Close
         If var_peso < CDbl(Me.lbl_bascula) Or var_peso = 0 Then
            rs.Open "insert into TB_ORACLE_PESOS_aRTICULOS (pedido, caja, codigo, peso, peso_real, peso_sistema, cantidad) values (" + Me.txt_pedido + "," + Me.txt_caja + ",'" + Me.txt_codigo + "'," + Me.lbl_bascula + ",0,0," + CStr(var_cantidad_leida) + ")"
            
            rs.Open "select max(consecutivo) as consecutivo from TB_ORACLE_PESOS_aRTICULOS where pedido = " + Me.txt_pedido + " and caja = " + Me.txt_caja, cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               rsaux.Open "update TB_ORACLE_PESOS_aRTICULOS set peso = " + Me.lbl_bascula + " where consecutivo = " + CStr(rs!CONSECUTIVO), cnn, adOpenDynamic, adLockOptimistic
            End If
            rs.Close
            
            
            rs.Open "SELECT * FROM tb_codigos_prueba_bascula WHERE PEDIDO = " + Me.txt_pedido + " AND CAJA = " + Me.txt_caja + " AND CODIGO = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               rsaux.Open "UPDATE tb_codigos_prueba_bascula SET CANTIDAD =  CANTIDAD + " + CStr(var_cantidad_leida) + "WHERE PEDIDO = " + Me.txt_pedido + " AND CAJA = " + Me.txt_caja + " AND CODIGO = '" + Me.txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
               For var_j = 1 To Me.lv_entradas.ListItems.Count
                   Me.lv_entradas.ListItems.Item(var_j).Selected = True
                     
                   If Me.lv_entradas.selectedItem = Me.txt_codigo Then
                      Me.lv_entradas.selectedItem.SubItems(2) = CDbl(Me.lv_entradas.selectedItem.SubItems(2)) + var_cantidad_leida
                   End If
               Next var_j
            Else
               rsaux.Open "INSERT INTO tb_codigos_prueba_bascula (CODIGO, DESCRIPCION, CANTIDAD, PEDIDO, CAJA) VALUES ('" + Me.txt_codigo + "','" + var_descripcion_articulo + "'," + CStr(var_cantidad_leida) + "," + Me.txt_pedido + "," + Me.txt_caja + ")", cnn, adOpenDynamic, adLockOptimistic
               Set list_item = lv_entradas.ListItems.Add(, , Trim(txt_codigo))
               list_item.SubItems(1) = var_descripcion_articulo
               list_item.SubItems(2) = var_cantidad_leida
            End If
            rs.Close
            
            Me.txt_codigo.BackColor = &H80000005
            Me.txt_codigo.Enabled = True
            Me.txt_codigo.Text = ""
            Me.txt_codigo.SetFocus
         Else
            Me.txt_codigo.BackColor = &H80000005
            Me.txt_codigo.Enabled = True
            Me.txt_codigo.Text = ""
            Me.txt_codigo.SetFocus
            frmmensaje.lbl_articulo = ""
            frmmensaje.lbl_mensaje = "No existe movimiento de peso anterior"
            frmmensaje.Show 1
         End If
      Else
           MsgBox "Número de caja incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Número de pedido incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub txt_pedido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_caja.SetFocus
   End If
End Sub
