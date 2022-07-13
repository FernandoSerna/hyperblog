VERSION 5.00
Begin VB.Form frmoracle_embarque_concentrado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Embarque"
   ClientHeight    =   1020
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   2940
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_embarque 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmoracle_embarque_concentrado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim objConn As New adodb.Connection
Dim objCmd As New adodb.Command
Dim objParm As adodb.Parameter
Dim comandoORA As New adodb.Command
Dim parametro As adodb.Parameter

Private Sub Form_Load()
   Top = 3300
   Left = 4500
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_embarque) Then
         rs.Open "SELECT PEDIDO FROM TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES WHERE EMBARQUE = '" + Me.txt_embarque + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            
            strconsulta = "select * from xxvia_Tb_encabezado_embarques where embarque = ?"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_embarque))
                 .Parameters.Append parametro
            End With
            Set rsaux1 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            If Not rsaux1.EOF Then
               var_fecha_inicio = CStr(IIf(IsNull(rsaux1!FECHA_INICIO), "", rsaux1!FECHA_INICIO))
               var_fecha_fin = CStr(IIf(IsNull(rsaux1!FECHA_FIN), "", rsaux1!FECHA_FIN))
               VAR_ESTATUS = IIf(IsNull(rsaux1!char_Emb_estatus), "", rsaux1!char_Emb_estatus)
            Else
               var_fecha_inicio = ""
               var_fecha_fin = ""
               VAR_ESTATUS = ""
            End If
            rsaux1.Close
            
            var_cadena_pedidos = ""
            While Not rs.EOF
                  If var_cadena_pedidos = "" Then
                     var_cadena_pedidos = CStr(rs!pedido)
                  Else
                     var_cadena_pedidos = var_cadena_pedidos + "," + CStr(rs!pedido)
                  End If
                  rs.MoveNext
            Wend
            'strconsulta = "select A.segment1 as codigo, A.item_description as descripcion, B.ATTRIBUTE2 AS UBICACION, sum(src_requested_quantity) as cantidad from xxvia_tb_pedidos_divididos A, XXVIA_SYSTEM_ITEMS_B B where source_header_number in (?) AND A.ORGANIZATION_ID = B.ORGANIZATION_ID AND A.SEGMENT1 = B.SEGMENT1 group by A.segment1, A.item_description, B.ATTRIBUTE2 ORDER BY ATTRIBUTE2"
            'With comandoORA
            '     .ActiveConnection = cnnoracle_4
            '     .CommandType = adCmdText
            '     .CommandText = strconsulta
            '     Set parametro = .CreateParameter(, adNumeric, adParamInput, 1000, var_cadena_pedidos)
            '     .Parameters.Append parametro
            'End With
            'Set rsaux1 = comandoORA.execute
            'Set comandoORA = Nothing
            'Set parametro = Nothing
            rsaux1.Open "select A.segment1 as codigo, A.item_description as descripcion, B.ATTRIBUTE2 AS UBICACION, sum(src_requested_quantity) as cantidad from xxvia_tb_pedidos_divididos A, XXVIA_SYSTEM_ITEMS_B B where source_header_number in (" + var_cadena_pedidos + ") AND A.ORGANIZATION_ID = B.ORGANIZATION_ID AND A.SEGMENT1 = B.SEGMENT1 group by A.segment1, A.item_description, B.ATTRIBUTE2 ORDER BY ATTRIBUTE2"
            cnn.BeginTrans
            rsaux2.Open "SELECT MAX(INTE_tEM_CONSECUTIVO) FROM TB_TEMP_ORACLE_embarque_concentrado", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux2.EOF Then
               var_consecutivo = IIf(IsNull(rsaux2(0).Value), 0, rsaux2(0).Value) + 1
            Else
               var_consecutivo = 1
            End If
            rsaux2.Close
            rsaux2.Open "INSERT INTO TB_TEMP_ORACLE_EMBARQUE_concentrado (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            
            While Not rsaux1.EOF
                  rsaux2.Open "INSERT INTO TB_TEMP_ORACLE_EMBARQUE_concentrado (INTE_tEM_CONSECUTIVO, EMBARQUE, FECHA_INICIO, FECHA_FIN, ESTATUS, CODIGO, DESCRIPCION, UBICACION, CANTIDAD) VALUES (" + CStr(var_consecutivo) + ",'" + Me.txt_embarque + "','" + var_fecha_inicio + "','" + var_fecha_fin + "','" + VAR_ESTATUS + "','" + rsaux1!CODIGO + "','" + rsaux1!DESCRIPCION + "','" + IIf(IsNull(rsaux1!ubicacion), "", rsaux1!ubicacion) + "'," + CStr(rsaux1!Cantidad) + ")", cnn, adOpenDynamic, adLockOptimistic
                  rsaux1.MoveNext
            Wend
            rsaux1.Close
            
            
         Else
            MsgBox "El embarque no tiene pedidos asociados", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         MsgBox "Número de embarque incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
End Sub
