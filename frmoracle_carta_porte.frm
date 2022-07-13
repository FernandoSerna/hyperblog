VERSION 5.00
Begin VB.Form frmoracle_carta_porte 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carta porte"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1080
      TabIndex        =   18
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox txt_uuid 
      Height          =   375
      Left            =   1200
      TabIndex        =   17
      Top             =   3000
      Width           =   7935
   End
   Begin VB.TextBox txt_distancia 
      Height          =   375
      Left            =   1200
      TabIndex        =   15
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox txt_placa 
      Height          =   375
      Left            =   1200
      TabIndex        =   13
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmoracle_carta_porte.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9000
      Picture         =   "frmoracle_carta_porte.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   60
      Left            =   0
      TabIndex        =   10
      Top             =   360
      Width           =   9225
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      Picture         =   "frmoracle_carta_porte.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.TextBox txt_chofer 
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   1560
      Width           =   7935
   End
   Begin VB.TextBox txt_cliente 
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   1080
      Width           =   7935
   End
   Begin VB.TextBox txt_pedido 
      Height          =   405
      Left            =   3600
      TabIndex        =   3
      Top             =   585
      Width           =   1455
   End
   Begin VB.TextBox txt_embarque 
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "UUID:"
      Height          =   195
      Left            =   240
      TabIndex        =   16
      Top             =   3090
      Width           =   450
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Distancia:"
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   2610
      Width           =   705
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Placas:"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   2130
      Width           =   525
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Chofer:"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   1650
      Width           =   510
   End
   Begin VB.Label Label3 
      Caption         =   "Cliente:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1140
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Pedido:"
      Height          =   195
      Left            =   2880
      TabIndex        =   2
      Top             =   690
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "Embarque:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   660
      Width           =   855
   End
End
Attribute VB_Name = "frmoracle_carta_porte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
    rs.Open "SELECT * FROM direcciones_envio_121121", cnn, adOpenDynamic, adLockOptimistic
    var_i = 50
    While Not rs.EOF
          var_i = var_i + 1
          'rsaux.Open "INSERT INTO TB_CHOFERES (VCHA_CHO_CHOFER_ID, VCHA_CHO_NOMBRE) VALUES (" + CStr(var_i) + ",'" + rs!NOMBRE + "')", cnn, adOpenDynamic, adLockOptimistic
          
          var_cadena = "INSERT INTO XXVIA_TB_ANES_CARTA_PORTE (CLAVE,CALLE, NUMERO_EXTERIOR, LOCALIDAD, MUNICIPIO, ESTADO, PAIS, CODIGO_POSTAL,colonia)"
          'var_cadena = var_cadena + " VALUES('" + rs!CLAVE + "','" + IIf(IsNull(rs!CALLE), "", rs!CALLE) + "','" + IIf(IsNull(rs!NUM_EXT), "", rs!NUM_EXT) + "','" + CStr(IIf(IsNull(rs!LOCALIDAD), "", rs!LOCALIDAD)) + "','" + CStr(IIf(IsNull(rs!MUNICIPIO), "", rs!MUNICIPIO)) + "','" + CStr(IIf(IsNull(rs!ESTADO), "", rs!ESTADO)) + "','MX','" + CStr(IIf(IsNull(rs!CODIGO_POSTAL), "", rs!CODIGO_POSTAL)) + "','" + IIf(IsNull(rs!colonia), "", rs!colonia) + "')"
          var_cadena = var_cadena + " VALUES('" + CStr(rs!site_use_id) + "','" + IIf(IsNull(rs!CALLE), "", rs!CALLE) + "','','" + CStr(IIf(IsNull(rs!LOCALIDAD), "", rs!LOCALIDAD)) + "','" + CStr(IIf(IsNull(rs!MUNICIPIO), "", rs!MUNICIPIO)) + "','" + CStr(IIf(IsNull(rs!ESTADO), "", rs!ESTADO)) + "','MX','" + CStr(IIf(IsNull(rs!POSTAL_code), "", rs!POSTAL_code)) + "','" + CStr(IIf(IsNull(rs!colonia), "", rs!colonia)) + "')"
          'var_cadena = var_cadena + " VALUES('" + rs!CLAVE + "','','','" + CStr(IIf(IsNull(rs!LOCALIDAD), "", rs!LOCALIDAD)) + "','" + CStr(IIf(IsNull(rs!MUNICIPIO), "", rs!MUNICIPIO)) + "','" + CStr(IIf(IsNull(rs!ESTADO), "", rs!ESTADO)) + "','MX','" + CStr(IIf(IsNull(rs!CODIGO_POSTAL), "", rs!CODIGO_POSTAL)) + "','" + CStr(IIf(IsNull(rs!colonia), "", rs!colonia)) + "')"
          'var_cadena = "UPDATE XXVIA_TB_choferes SET rfc = '" + CStr(IIf(IsNull(rs!rfc), "", rs!rfc)) + "', licencia = '" + IIf(IsNull(rs!licencia), "", rs!licencia) + "' WHERE id_chofer = '" + CStr(rs!clave) + "'"
          'var_cadena = "INSERT INTO XXVIA_TB_CHOFERES (ID_CHOFER, NOMBRE,RFC, LICENCIA) VALUES (" + CStr(var_i) + ",'" + rs!NOMBRE + "','" + rs!RFC + "','" + rs!LICENCIA + "')"
          
          rsaux2.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
          rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub Form_Load()
   Top = 1400
   Left = 1200
   Me.txt_embarque = 343916
   Me.txt_pedido = 699465

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim clnt As New SoapClient30
          var_cadena = "SELECT * from xxvia_tb_encabezado_embarques where embarque = ?"
          strconsulta = var_cadena
          With comandoORA
               .ActiveConnection = cnnoracle_4
               .CommandType = adCmdText
               .CommandText = strconsulta
               Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, Me.txt_embarque)
               .Parameters.Append parametro
          End With
          Set rsaux12 = comandoORA.execute
          Set comandoORA = Nothing
          Set parametro = Nothing
          If Not rsaux12.EOF Then
             Me.txt_pedido.SetFocus
          Else
             MsgBox "El embarque no existe.", vbOKOnly, "ATENCION"
          End If
          rsaux12.Close
   End If
End Sub

Private Sub txt_pedido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      
      Dim var_location_id As Double
      Dim VAR_CLAVE_USUARIO_MOV As String
      Dim var_fecha_inicio As String
      Dim var_fecha_fin As String
      Dim var_consignacion As String
      If IsNumeric(Me.txt_pedido) Then
         var_posible_embarque = 1
         var_Cadena_pedidos = Me.txt_pedido
         var_numero_embarque = Me.txt_embarque
         strconsulta = "call XXVIA_SP_TIMBRAR_TRASPASOS_3(?,?,?,?)"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, Me.txt_pedido)
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
         VAR_SERIE = "TRX" + Me.txt_embarque + "_"
         strconsulta = "select customer_trx_id, cadena as cadena, numero from xxvia_tb_control_doc_fiscales where serie = '" + VAR_SERIE + "' and numero = ? AND ORGANIZACION = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, Me.txt_pedido)
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
            Open ("C:\SISTEMAS\TRX" + Trim(Me.txt_embarque) + "_" + Trim(Me.txt_pedido) + ".FAC") For Output As #1
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
            var_archivo = "C:\SISTEMAS\sube_fact" + Trim(Str(Me.txt_embarque)) + "_" + Trim(Me.txt_pedido) + ".bat"
            x = Shell("c:\sistemas\facturar " + """" + "facturar|C:\SISTEMAS\|C:\SISTEMAS\TRX" + Me.txt_embarque + "_" + Trim(Me.txt_pedido) + ".FAC" + "|https://facturas2.vianney.mx/cgi-bin/cfds/timbrarGR33|cfdsvianney|9y3jv^TI;4g#|1" + """", vbHide)
            rsaux2.Close
         Else
            MsgBox "No existen movimientos para el embarque seleccionado", vbOKOnly, "ATENCION"
         End If
      End If
   End If
   If KeyAscii = 27 Then
      'Me.frm_embarque_nota_envio.Visible = False
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
