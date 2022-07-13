VERSION 5.00
Begin VB.Form frmdepura_ordenes 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   480
      Left            =   180
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2400
      Width           =   4260
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   15
      Top             =   1005
   End
   Begin VB.CommandButton cmd_depurar 
      Caption         =   "Depurar pedidos"
      Height          =   2055
      Left            =   105
      TabIndex        =   0
      Top             =   75
      Width           =   4470
   End
End
Attribute VB_Name = "frmdepura_ordenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_depurar_Click()
Dim cnn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim rs_2 As ADODB.Recordset
   Set cnn = CreateObject("ADODB.connection")
   Set rs = CreateObject("ADODB.recordset")
   Set rs_2 = CreateObject("ADODB.recordset")
   cnn.Open "Provider=SQLOLEDB.1;Password=elia;Persist Security Info=True;User ID=sa;Initial Catalog=vianney;Data Source=distribucion"
   cnn.CursorLocation = adUseClient
   Dim var_dias As Integer
   rs.Open "select * from tb_principal", cnn, adOpenDynamic, adLockOptimistic
   'var_dias = IIf(IsNull(rs!INTE_PRI_DIAS_DEPURACION), 3, rs!INTE_PRI_DIAS_DEPURACION)
   var_dias = 3
   rs.Close
   
   var_cadena = "SELECT dbo.TB_ENC_ORDEN_SURTIDO.INTE_PED_NUMERO, dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_ORDEN_SURTIDO, dbo.TB_ENC_ORDEN_SURTIDO.DTIM_ORS_FECHA_CARGA, CAST(GETDATE() - dbo.TB_ENC_ORDEN_SURTIDO.DTIM_ORS_FECHA_CARGA AS integer) AS diferencia, dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_LIBERADA FROM         dbo.TB_ENC_ORDEN_SURTIDO INNER JOIN dbo.TB_ENCABEZADO_PEDIDOS ON                       dbo.TB_ENC_ORDEN_SURTIDO.INTE_PED_NUMERO = dbo.TB_ENCABEZADO_PEDIDOS.INTE_PED_NUMERO WHERE     (dbo.TB_ENC_ORDEN_SURTIDO.CHAR_TPE_TIPO_PEDIDO_ID = 'FT') AND (dbo.TB_ENCABEZADO_PEDIDOS.CHAR_PED_ESTATUS <> 'E') AND (dbo.TB_ENCABEZADO_PEDIDOS.CHAR_PED_ESTATUS <> 'C') AND (CAST(GETDATE() - dbo.TB_ENC_ORDEN_SURTIDO.DTIM_ORS_FECHA_CARGA AS integer) > 3) AND (dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_LIBERADA <> 1) OR"
   var_cadena = var_cadena + " (dbo.TB_ENC_ORDEN_SURTIDO.CHAR_TPE_TIPO_PEDIDO_ID = 'FT') AND (dbo.TB_ENCABEZADO_PEDIDOS.CHAR_PED_ESTATUS <> 'E') AND (dbo.TB_ENCABEZADO_PEDIDOS.CHAR_PED_ESTATUS <> 'C') AND (CAST(GETDATE() - dbo.TB_ENC_ORDEN_SURTIDO.DTIM_ORS_FECHA_CARGA AS integer) > +" + CStr(var_dias) + ") AND (dbo.TB_ENC_ORDEN_SURTIDO.INTE_ORS_LIBERADA IS NULL)"
   Text1 = var_cadena
   rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rs_2.Open "exec SP_CANCELACION_PEDIDO_TIENDA " + CStr(rs!INTE_PED_NUMERO), cnn, adOpenDynamic, adLockOptimistic
         rs.MoveNext
   Wend
   rs.Close
   End
End Sub

Private Sub Timer1_Timer()
   Call cmd_depurar_Click
End Sub
