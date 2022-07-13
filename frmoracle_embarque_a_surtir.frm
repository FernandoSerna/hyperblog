VERSION 5.00
Begin VB.Form frmoracle_embarque_a_surtir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Embarque a surtir"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   2775
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmoracle_embarque_a_surtir"
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
         rs.Open "Select * from tb_oracle_pedidos_asignados_embarques where embarque = '" + Me.txt_embarque + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            rsaux1.Open "Select SUM(CANTIDAD_SIN_CATALOGOS)  from tb_oracle_pedidos_asignados_embarques where embarque = '" + Me.txt_embarque + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux1.EOF Then
               var_cantidad_total = IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value)
            Else
               var_cantidad_total = 0
            End If
            rsaux1.Close
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
            If VAR_ESTATUS = "" Then
               rsaux2.Open "SELECT * FROM TB_ORACLE_EMBARQUES_SURTIR WHERE EMBARQUE = '" + Me.txt_embarque + "'", cnn, adOpenDynamic, adLockOptimistic
               If rsaux2.EOF Then
                  rsaux1.Open "INSERT INTO TB_ORACLE_EMBARQUES_SURTIR (EMBARQUE, PIEZAS_SURTIR) VALUES ('" + Me.txt_embarque + "'," + Str(var_cantidad_total) + ")", cnn, adOpenDynamic, adLockOptimistic
                  MsgBox "Se a insertado el embarque", vbOKOnly, "ATENCION"
               Else
                  MsgBox "El embarque ya habia sido cargado", vbOKOnly, "ATENCION"
               End If
               rsaux2.Close
            Else
               MsgBox "El embarque ya fue cerrado", vbOKOnly, "ATENCION"
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
