VERSION 5.00
Begin VB.Form frmoracle_negado_distribucion_hora_x_hora 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hora x hora"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   15
      TabIndex        =   4
      Top             =   330
      Width           =   2910
   End
   Begin VB.Frame Frame4 
      Caption         =   " Fecha "
      Height          =   765
      Left            =   45
      TabIndex        =   2
      Top             =   435
      Width           =   2865
      Begin VB.TextBox txt_inicio 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   780
         TabIndex        =   3
         Top             =   315
         Width           =   1350
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2565
      Picture         =   "frmoracle_negado_distribucion_hora_x_hora.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   0
      Picture         =   "frmoracle_negado_distribucion_hora_x_hora.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Negado de distribución"
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "frmoracle_negado_distribucion_hora_x_hora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Dim comandoORA As New adodb.Command
Dim parametro As adodb.Parameter
Dim objConn As New adodb.Connection
Dim objCmd As New adodb.Command
Dim objParm As adodb.Parameter
Dim var_contador_porcentaje As Integer
Dim var_cubicaje As Double
Dim var_ventana As Integer

Private Sub cmd_imprimir_Click()
   If IsDate(Me.txt_inicio) Then
      var_fecha_anterior = CDate(Me.txt_inicio) - 1
      var_dia_str = CStr(Day(CDate(var_fecha_anterior)))
      If Len(var_dia_str) = 1 Then
         var_dia_str = "0" + var_dia_str
      End If
      var_mes_str = CStr(Month(CDate(var_fecha_anterior)))
      If Len(var_mes_str) = 1 Then
         var_mes_str = "0" + var_mes_str
      End If
      var_año_str = CStr(Year(CDate(var_fecha_anterior)))
      If Len(var_año_str) = 2 Then
         var_año_str = "20" + var_año_str
      End If
      var_fecha_anterior_str = var_dia_str + "/" + var_mes_str + "/" + var_año_str
      
      VAR_FECHA_ACTUAL = CDate(Me.txt_inicio)
      var_dia_str = CStr(Day(CDate(VAR_FECHA_ACTUAL)))
      If Len(var_dia_str) = 1 Then
         var_dia_str = "0" + var_dia_str
      End If
      var_mes_str = CStr(Month(CDate(VAR_FECHA_ACTUAL)))
      If Len(var_mes_str) = 1 Then
         var_mes_str = "0" + var_mes_str
      End If
      var_año_str = CStr(Year(CDate(VAR_FECHA_ACTUAL)))
      If Len(var_año_str) = 2 Then
         var_año_str = "20" + var_año_str
      End If
      VAR_FECHA_ACTUAL_STR = var_dia_str + "/" + var_mes_str + "/" + var_año_str
      VAR_FECHA_SQL = "{d '" + var_año_str + "-" + var_mes_str + "-" + var_dia_str + "'}"
      var_fecha_inicio = var_fecha_anterior_str + " 23:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 00:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('NO LOCALIZADO','XXVIA1') AND  fecha_negado >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      cnn.BeginTrans
      rsaux9.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_ORACLE_NEGADO_DISTRIBUCION_HXH", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux9.EOF Then
         var_consecutivo = IIf(IsNull(rsaux9(0).Value), 0, rsaux9(0).Value)
      Else
         var_consecutivo = 0
      End If
      rsaux9.Close
      var_consecutivo = var_consecutivo + 1
      rsaux9.Open "INSERT INTO TB_ORACLE_NEGADO_DISTRIBUCION_HXH (INTE_tEM_CONSECUTIVO,TIPO, FECHA, HORA_23) VALUES (" + CStr(var_consecutivo) + ",'NO ENCONTRADO'," + VAR_FECHA_SQL + "," + CStr(var_cantidad) + ")", cnn, adOpenDynamic, adLockOptimistic
      cnn.CommitTrans
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 00:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 01:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '24', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in ('NO LOCALIZADO','XXVIA1')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_24 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 01:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 02:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '01', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('NO LOCALIZADO','XXVIA1')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_01 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 02:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 03:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('NO LOCALIZADO','XXVIA1')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_02 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 03:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 04:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '02', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('NO LOCALIZADO','XXVIA1')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_03 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 04:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 05:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('NO LOCALIZADO','XXVIA1')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_04 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 05:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 06:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('NO LOCALIZADO','XXVIA1')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_05 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 06:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 07:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('NO LOCALIZADO','XXVIA1')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_06 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 07:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 08:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('NO LOCALIZADO','XXVIA1')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_07 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 08:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 09:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('NO LOCALIZADO','XXVIA1')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_08 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 09:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 10:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('NO LOCALIZADO','XXVIA1')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_09 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 10:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 11:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('NO LOCALIZADO','XXVIA1')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_10 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 11:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 12:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('NO LOCALIZADO','XXVIA1')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_11 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 12:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 13:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('NO LOCALIZADO','XXVIA1')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_12 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 13:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 14:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('NO LOCALIZADO','XXVIA1')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_13 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 14:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 15:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('NO LOCALIZADO','XXVIA1')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_14 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 15:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 16:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('NO LOCALIZADO','XXVIA1')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_15 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 16:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 17:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('NO LOCALIZADO','XXVIA1')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_16 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 17:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 18:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('NO LOCALIZADO','XXVIA1')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_17 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 18:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 19:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('NO LOCALIZADO','XXVIA1')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_18 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 19:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 20:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('NO LOCALIZADO','XXVIA1')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_19 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 20:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 21:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('NO LOCALIZADO','XXVIA1')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_20 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 21:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 22:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('NO LOCALIZADO','XXVIA1')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_21 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 22:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 23:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('NO LOCALIZADO','XXVIA1')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_22 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      
      
      
''''ENVIADOS POR PAQUETERIA
      
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('MUEBLES ENVIADOS POR PAQUETERIA','XXVIA2') AND fecha_negado >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      cnn.BeginTrans

      rsaux9.Open "INSERT INTO TB_ORACLE_NEGADO_DISTRIBUCION_HXH (INTE_tEM_CONSECUTIVO,TIPO, FECHA, HORA_23) VALUES (" + CStr(var_consecutivo) + ",'ENVIADOS POR PAQUETERIA'," + VAR_FECHA_SQL + "," + CStr(var_cantidad) + ")", cnn, adOpenDynamic, adLockOptimistic
      cnn.CommitTrans
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 00:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 01:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      
      strconsulta = "select '24', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('MUEBLES ENVIADOS POR PAQUETERIA','XXVIA2') AND fecha_negado  >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_24 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND TIPO = 'ENVIADOS POR PAQUETERIA'", cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 01:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 02:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '01', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('MUEBLES ENVIADOS POR PAQUETERIA','XXVIA2')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_01 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND TIPO = 'ENVIADOS POR PAQUETERIA'", cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 02:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 03:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('MUEBLES ENVIADOS POR PAQUETERIA','XXVIA2')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_02 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND TIPO = 'ENVIADOS POR PAQUETERIA'", cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 03:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 04:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '02', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('MUEBLES ENVIADOS POR PAQUETERIA','XXVIA2')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_03 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND TIPO = 'ENVIADOS POR PAQUETERIA'", cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 04:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 05:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('MUEBLES ENVIADOS POR PAQUETERIA','XXVIA2')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_04 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND TIPO = 'ENVIADOS POR PAQUETERIA'", cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 05:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 06:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('MUEBLES ENVIADOS POR PAQUETERIA','XXVIA2')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_05 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND TIPO = 'ENVIADOS POR PAQUETERIA'", cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 06:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 07:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('MUEBLES ENVIADOS POR PAQUETERIA','XXVIA2')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_06 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND TIPO = 'ENVIADOS POR PAQUETERIA'", cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 07:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 08:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('MUEBLES ENVIADOS POR PAQUETERIA','XXVIA2')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_07 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND TIPO = 'ENVIADOS POR PAQUETERIA'", cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 08:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 09:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('MUEBLES ENVIADOS POR PAQUETERIA','XXVIA2')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_08 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND TIPO = 'ENVIADOS POR PAQUETERIA'", cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 09:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 10:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('MUEBLES ENVIADOS POR PAQUETERIA','XXVIA2')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_09 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND TIPO = 'ENVIADOS POR PAQUETERIA'", cnn, adOpenDynamic, adLockOptimistic
      
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 10:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 11:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('MUEBLES ENVIADOS POR PAQUETERIA','XXVIA2')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_10 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND TIPO = 'ENVIADOS POR PAQUETERIA'", cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 11:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 12:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('MUEBLES ENVIADOS POR PAQUETERIA','XXVIA2')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_11 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND TIPO = 'ENVIADOS POR PAQUETERIA'", cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 12:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 13:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('MUEBLES ENVIADOS POR PAQUETERIA','XXVIA2')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_12 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND TIPO = 'ENVIADOS POR PAQUETERIA'", cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 13:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 14:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('MUEBLES ENVIADOS POR PAQUETERIA','XXVIA2')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_13 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND TIPO = 'ENVIADOS POR PAQUETERIA'", cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 14:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 15:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('MUEBLES ENVIADOS POR PAQUETERIA','XXVIA2')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_14 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND TIPO = 'ENVIADOS POR PAQUETERIA'", cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 15:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 16:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('MUEBLES ENVIADOS POR PAQUETERIA','XXVIA2')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_15 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND TIPO = 'ENVIADOS POR PAQUETERIA'", cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 16:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 17:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('MUEBLES ENVIADOS POR PAQUETERIA','XXVIA2')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_16 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND TIPO = 'ENVIADOS POR PAQUETERIA'", cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 17:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 18:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('MUEBLES ENVIADOS POR PAQUETERIA','XXVIA2')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_17 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND TIPO = 'ENVIADOS POR PAQUETERIA'", cnn, adOpenDynamic, adLockOptimistic
      
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 18:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 19:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('MUEBLES ENVIADOS POR PAQUETERIA','XXVIA2')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_18 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND TIPO = 'ENVIADOS POR PAQUETERIA'", cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 19:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 20:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('MUEBLES ENVIADOS POR PAQUETERIA','XXVIA2')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_19 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND TIPO = 'ENVIADOS POR PAQUETERIA'", cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 20:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 21:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('MUEBLES ENVIADOS POR PAQUETERIA','XXVIA2')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_20 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND TIPO = 'ENVIADOS POR PAQUETERIA'", cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 21:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 22:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('MUEBLES ENVIADOS POR PAQUETERIA','XXVIA2')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_21 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND TIPO = 'ENVIADOS POR PAQUETERIA'", cnn, adOpenDynamic, adLockOptimistic
      
      var_fecha_inicio = VAR_FECHA_ACTUAL_STR + " 22:00:00"
      var_fecha_fin = VAR_FECHA_ACTUAL_STR + " 23:00:00"
      rsaux9.Open "alter session set nls_date_format='DD/MM/YYYY hh24:mi:ss'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      strconsulta = "select '23', sum(cantidad) as cantidad from xxvia_Tb_negado_distribucion where causa_negado in('MUEBLES ENVIADOS POR PAQUETERIA','XXVIA2')AND  FECHA_NEGADO >= to_date(?,'DD/MM/YYYY hh24:mi:ss') and fecha_negado < to_date(?,'DD/MM/YYYY hh24:mi:ss') and cantidad> 0"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_inicio)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_fecha_fin)
           .Parameters.Append parametro
      End With
      Set rsaux9 = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux9.EOF Then
         var_cantidad = IIf(IsNull(rsaux9!Cantidad), 0, rsaux9!Cantidad)
      Else
         var_cantidad = 0
      End If
      rsaux9.Close
      rsaux9.Open "UPDATE TB_ORACLE_NEGADO_DISTRIBUCION_HXH SET HORA_22 = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND TIPO = 'ENVIADOS POR PAQUETERIA'", cnn, adOpenDynamic, adLockOptimistic
      
      
      
      
''''FIN ENVIADOS POR PAQUETERIA
      
      
      
      
                  Set reporte = appl.OpenReport(App.Path + "\rep_negado_distribucion_hxh.rpt")
                  reporte.RecordSelectionFormula = "{TB_ORACLE_NEGADO_DISTRIBUCION_HXH.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
                  frmvistasprevias.cr.ReportSource = reporte
                  For ntablas = 1 To reporte.Database.Tables.Count
                      reporte.Database.Tables(ntablas).SetLogOnInfo "sqlsistema", var_bd_reportes, "sa", "elia"
                  Next ntablas
                  frmvistasprevias.cr.ViewReport
                  frmvistasprevias.Caption = "Devolución de clientes"
                  frmvistasprevias.Show 1
                  Set reporte = Nothing
      
      
      
      
   Else
      MsgBox "Fecha incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 2900
   Left = 3500
   Me.txt_inicio = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_inicio) Then
         frmcalendario.mes = CDate(Me.txt_inicio)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      txt_inicio = var_fecha_general
   End If
End Sub

