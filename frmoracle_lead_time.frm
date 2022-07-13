VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_lead_time 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lead time"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6540
      Picture         =   "frmoracle_lead_time.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmoracle_lead_time.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   45
      TabIndex        =   12
      Top             =   255
      Width           =   6825
   End
   Begin VB.Frame Frame3 
      Height          =   6840
      Left            =   60
      TabIndex        =   0
      Top             =   345
      Width           =   6825
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1410
         Picture         =   "frmoracle_lead_time.frx":073C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   750
         Picture         =   "frmoracle_lead_time.frx":0952
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Marcar (Enter)"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         Picture         =   "frmoracle_lead_time.frx":0B9C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   90
         Picture         =   "frmoracle_lead_time.frx":0C6E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   135
         Width           =   330
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   420
         Picture         =   "frmoracle_lead_time.frx":0D70
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   135
         Width           =   330
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   3990
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   165
         Width           =   1140
      End
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   5595
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   150
         Width           =   1140
      End
      Begin MSComctlLib.ListView lv_almacenes 
         Height          =   6255
         Left            =   60
         TabIndex        =   8
         Top             =   480
         Width           =   6690
         _ExtentX        =   11800
         _ExtentY        =   11033
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
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre CN"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "marca"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   3465
         TabIndex        =   10
         Top             =   225
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   5265
         TabIndex        =   9
         Top             =   210
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmoracle_lead_time"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter

Private Sub cmd_imprimir_Click()
   Dim var_entrega_x As String
   Dim var_fecha_creacion As String
   var_cadena_almacenes = ""
   For var_j = 1 To Me.lv_almacenes.ListItems.Count
       Me.lv_almacenes.ListItems(var_j).Selected = True
       If Me.lv_almacenes.selectedItem.SubItems(2) = "*" Then
          If var_cadena_almacenes = "" Then
             var_cadena_almacenes = "'" + Me.lv_almacenes.selectedItem + "'"
          Else
             var_cadena_almacenes = var_cadena_almacenes + ",'" + Me.lv_almacenes.selectedItem + "'"
          End If
       End If
   Next var_j
   If var_cadena_almacenes <> "" Then
      var_dia_inicio = CStr(Day(CDate(Me.txt_inicio)))
      var_mes_inicio = CStr(Month(CDate(Me.txt_inicio)))
      var_año_inicio = CStr(Year(CDate(Me.txt_inicio)))
      If Len(var_dia_inicio) = 1 Then
         var_dia_inicio = "0" + var_dia_inicio
      End If
      If Len(var_mes_inicio) = 1 Then
         var_mes_inicio = "0" + var_mes_inicio
      End If
      If Len(var_año_inicio) = 1 Then
         var_año_inicio = "20" + var_año_inicio
      End If
      var_fecha_inicio = "{d '" + var_año_inicio + "-" + var_mes_inicio + "-" + var_dia_inicio + "'}"
   
      var_dia_fin = CStr(Day(CDate(Me.txt_fin) + 1))
      var_mes_fin = CStr(Month(CDate(Me.txt_fin) + 1))
      var_año_fin = CStr(Year(CDate(Me.txt_fin) + 1))
      If Len(var_dia_fin) = 1 Then
         var_dia_fin = "0" + var_dia_fin
      End If
      If Len(var_mes_fin) = 1 Then
         var_mes_fin = "0" + var_mes_fin
      End If
      If Len(var_año_fin) = 1 Then
         var_año_fin = "20" + var_año_fin
      End If
      var_fecha_fin = "{d '" + var_año_fin + "-" + var_mes_fin + "-" + var_dia_fin + "'}"
      
      var_dia_fin = CStr(Day(CDate(Me.txt_fin)))
      var_mes_fin = CStr(Month(CDate(Me.txt_fin)))
      var_año_fin = CStr(Year(CDate(Me.txt_fin)))
      If Len(var_dia_fin) = 1 Then
         var_dia_fin = "0" + var_dia_fin
      End If
      If Len(var_mes_fin) = 1 Then
         var_mes_fin = "0" + var_mes_fin
      End If
      If Len(var_año_fin) = 1 Then
         var_año_fin = "20" + var_año_fin
      End If
      var_fecha_FIN_REPORTE = "{d '" + var_año_fin + "-" + var_mes_fin + "-" + var_dia_fin + "'}"
      If rsaux.State = 1 Then
         rsaux.Close
      End If
      rsaux.Open "select * from tb_oracle_tiempo_impresion_documentos where fecha >= " + var_fecha_inicio + " and fecha < " + var_fecha_fin + " AND TIENDA IN (" + var_cadena_almacenes + ")", cnn, adOpenDynamic, adLockOptimistic
      'rsaux.Open "select * from tb_oracle_tiempo_impresion_documentos where fecha >= " + var_fecha_inicio + " and fecha < " + var_fecha_fin + " AND TIENDA IN (" + var_cadena_almacenes + ")", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux.EOF Then
         cnn.BeginTrans
         rsaux1.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_temp_ORACLE_LEAD_TIME", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux1.EOF Then
            var_consecutivo = IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value)
         Else
            var_consecutivo = 0
         End If
         rsaux1.Close
         var_consecutivo = var_consecutivo + 1
         rsaux1.Open "INSERT INTO TB_temp_ORACLE_LEAD_TIME (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
         cnn.CommitTrans
         rsaux9.Open "alter session set nls_date_format = 'DD-MON-YYYY HH24:MI:SS'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
         
               strconsulta = "SELECT creation_Date from oe_order_headers_all where order_number = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, CDbl(rsaux!pedido))
                   .Parameters.Append parametro
               End With
               Set rsaux9 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
               If Not rsaux9.EOF Then
                  'var_fecha_creacion = IIf(IsNull(rsaux9!creation_Date), "", rsaux9!creation_Date)
                  var_dia_str = CStr(Day(rsaux9!creation_Date))
                  If Len(var_dia_str) = 1 Then
                     var_dia_str = "0" + var_dia_str
                  End If
                  var_mes_str = CStr(Month(rsaux9!creation_Date))
                  If Len(var_mes_str) = 1 Then
                     var_mes_str = "0" + var_mes_str
                  End If
                  var_año_str = CStr(Year(rsaux9!creation_Date))
                  If Len(var_año_str) = 1 Then
                     var_año_str = "0" + var_año_str
                  End If
                  var_hora_str = CStr(Hour(rsaux9!creation_Date))
                  If Len(var_hora_str) = 1 Then
                     var_hora_str = "0" + var_hora_str
                  End If
                  var_minute_str = CStr(Minute(rsaux9!creation_Date))
                  If Len(var_minute_str) = 1 Then
                     var_minute_str = "0" + var_minute_str
                  End If
                  VAR_SEGUNDO_STR = CStr(Second(rsaux9!creation_Date))
                  If Len(VAR_SEGUNDO_STR) = 1 Then
                     VAR_SEGUNDO_STR = "0" + VAR_SEGUNDO_STR
                  End If
                  var_fecha_creacion = "{d '" + var_año_str + "-" + var_mes_str + "-" + var_dia_str + "'}"
              
               End If
               rsaux9.Close
               x = 1
               If x = 0 Then
               cnn_lead_time.CommandTimeout = 300
               var_cadena = "select max(a.fechamodificado) as fecha_entrada FROM sqlposprod.dbicgprod_via.dbo.ALBVENTACAB a, sqlposprod.dbicgprod_via.dbo.ALBCOMPRACAB b WHERE substring(a.SUALBARAN,1,LEN(cast(" + CStr(rsaux!pedido) + " as varchar(50)))) = cast(" + CStr(rsaux!pedido) + " as varchar(50)) and '.'+a.numserie+'-'+cast(a.numalbaran as varchar(50)) = b.SUALBARAN"
               rs.Open var_cadena, cnn_lead_time, adOpenDynamic, adLockOptimistic
               End If
               'MsgBox rs!FECHA_ENTRADA
               'CAST('01/01/2000 14:30:20:999' AS datetime2)
               'var_entrega_x = IIf(IsNull(rs!FECHA_ENTRADA), "0", "1111")
               'MsgBox var_entrega_x
               'If Not rs.EOF Then
               If var_x = 1 Then
                  If var_entrega_x = "1111" Then
                     var_dia_str = CStr(Day(rsaux!Fecha))
                     If Len(var_dia_str) = 1 Then
                        var_dia_str = "0" + var_dia_str
                     End If
                     var_mes_str = CStr(Month(rsaux!Fecha))
                     If Len(var_mes_str) = 1 Then
                        var_mes_str = "0" + var_mes_str
                     End If
                     var_año_str = CStr(Year(rsaux!Fecha))
                     If Len(var_año_str) = 1 Then
                        var_año_str = "0" + var_año_str
                     End If
                     
                     var_hora_str = CStr(Hour(rsaux!Fecha))
                     If Len(var_hora_str) = 1 Then
                        var_hora_str = "0" + var_hora_str
                     End If
                     var_minute_str = CStr(Minute(rsaux!Fecha))
                     If Len(var_minute_str) = 1 Then
                        var_minute_str = "0" + var_minute_str
                     End If
                     VAR_SEGUNDO_STR = CStr(Second(rsaux!Fecha))
                     If Len(VAR_SEGUNDO_STR) = 1 Then
                        VAR_SEGUNDO_STR = "0" + VAR_SEGUNDO_STR
                     End If
                     var_x = Replace(UCase(Mid(CStr(rsaux!Fecha), 21, 4)), ".", "")
                     var_fecha_pedido = var_dia_str + "/" + var_mes_str + "/" + var_año_str + " " + var_hora_str + ":" + var_minute_str + ":" + VAR_SEGUNDO_STR + ":500 " + var_x
                  
                     var_dia_str = CStr(Day(rs!FECHA_ENTRADA))
                     If Len(var_dia_str) = 1 Then
                        var_dia_str = "0" + var_dia_str
                     End If
                     var_mes_str = CStr(Month(rs!FECHA_ENTRADA))
                     If Len(var_mes_str) = 1 Then
                        var_mes_str = "0" + var_mes_str
                     End If
                     var_año_str = CStr(Year(rs!FECHA_ENTRADA))
                     If Len(var_año_str) = 1 Then
                        var_año_str = "0" + var_año_str
                     End If
                     
                     var_hora_str = CStr(Hour(rs!FECHA_ENTRADA))
                     If Len(var_hora_str) = 1 Then
                        var_hora_str = "0" + var_hora_str
                     End If
                     var_minute_str = CStr(Minute(rs!FECHA_ENTRADA))
                     If Len(var_minute_str) = 1 Then
                        var_minute_str = "0" + var_minute_str
                     End If
                     VAR_SEGUNDO_STR = CStr(Second(rs!FECHA_ENTRADA))
                     If Len(VAR_SEGUNDO_STR) = 1 Then
                        VAR_SEGUNDO_STR = "0" + VAR_SEGUNDO_STR
                     End If
                     var_x = Replace(UCase(Mid(CStr(rs!FECHA_ENTRADA), 21, 4)), ".", "")
                     var_fecha_entrega = var_dia_str + "/" + var_mes_str + "/" + var_año_str + " " + var_hora_str + ":" + var_minute_str + ":" + VAR_SEGUNDO_STR + ":500 " + var_x
                     
                     rsaux2.Open "INSERT INTO TB_temp_ORACLE_LEAD_TIME (INTE_TEM_CONSECUTIVO,PEDIDO, TIENDA, NOMBRE_TIENDA,FECHA_INICIO,FECHA_FINAL, INICIO_REPORTE, FIN_REPORTE, FECHA_CREACION) VALUES (" + CStr(var_consecutivo) + "," + CStr(rsaux!pedido) + ",'" + rsaux!tienda + "','" + rsaux!NOMBRE + "',cast('" + Mid(var_fecha_pedido, 1, 10) + "' as datetime),cast('" + Mid(var_fecha_entrega, 1, 10) + "' as datetime)," + var_fecha_inicio + "," + var_fecha_FIN_REPORTE + "," + var_fecha_creacion + ")", cnn, adOpenDynamic, adLockOptimistic
                  Else
                     var_dia_str = CStr(Day(rsaux!Fecha))
                     If Len(var_dia_str) = 1 Then
                        var_dia_str = "0" + var_dia_str
                     End If
                     var_mes_str = CStr(Month(rsaux!Fecha))
                     If Len(var_mes_str) = 1 Then
                        var_mes_str = "0" + var_mes_str
                     End If
                     var_año_str = CStr(Year(rsaux!Fecha))
                     If Len(var_año_str) = 1 Then
                        var_año_str = "0" + var_año_str
                     End If
                     
                     var_hora_str = CStr(Hour(rsaux!Fecha))
                     If Len(var_hora_str) = 1 Then
                        var_hora_str = "0" + var_hora_str
                     End If
                     var_minute_str = CStr(Minute(rsaux!Fecha))
                     If Len(var_minute_str) = 1 Then
                        var_minute_str = "0" + var_minute_str
                     End If
                     VAR_SEGUNDO_STR = CStr(Second(rsaux!Fecha))
                     If Len(VAR_SEGUNDO_STR) = 1 Then
                        VAR_SEGUNDO_STR = "0" + VAR_SEGUNDO_STR
                     End If
                     var_x = Replace(UCase(Mid(CStr(rsaux!Fecha), 21, 4)), ".", "")
                     var_fecha_pedido = var_dia_str + "/" + var_mes_str + "/" + var_año_str + " " + var_hora_str + ":" + var_minute_str + ":" + VAR_SEGUNDO_STR + ":500 " + var_x
                  
                     
                     rsaux2.Open "INSERT INTO TB_temp_ORACLE_LEAD_TIME (INTE_TEM_CONSECUTIVO,PEDIDO, TIENDA, NOMBRE_TIENDA,FECHA_INICIO,INICIO_REPORTE, FIN_REPORTE, FECHA_cREACION) VALUES (" + CStr(var_consecutivo) + "," + CStr(rsaux!pedido) + ",'" + rsaux!tienda + "','" + rsaux!NOMBRE + "',cast('" + Mid(var_fecha_pedido, 1, 10) + "' as datetime)," + var_fecha_inicio + "," + var_fecha_FIN_REPORTE + "," + var_fecha_creacion + ")", cnn, adOpenDynamic, adLockOptimistic
                  End If
               Else
                  var_dia_str = CStr(Day(rsaux!Fecha))
                  If Len(var_dia_str) = 1 Then
                     var_dia_str = "0" + var_dia_str
                  End If
                  var_mes_str = CStr(Month(rsaux!Fecha))
                  If Len(var_mes_str) = 1 Then
                     var_mes_str = "0" + var_mes_str
                  End If
                  var_año_str = CStr(Year(rsaux!Fecha))
                  If Len(var_año_str) = 1 Then
                     var_año_str = "0" + var_año_str
                  End If
                  
                  var_hora_str = CStr(Hour(rsaux!Fecha))
                  If Len(var_hora_str) = 1 Then
                     var_hora_str = "0" + var_hora_str
                  End If
                  var_minute_str = CStr(Minute(rsaux!Fecha))
                  If Len(var_minute_str) = 1 Then
                     var_minute_str = "0" + var_minute_str
                  End If
                  VAR_SEGUNDO_STR = CStr(Second(rsaux!Fecha))
                  If Len(VAR_SEGUNDO_STR) = 1 Then
                     VAR_SEGUNDO_STR = "0" + VAR_SEGUNDO_STR
                  End If
                  var_x = Replace(UCase(Mid(CStr(rsaux!Fecha), 21, 4)), ".", "")
                  var_fecha_pedido = var_dia_str + "/" + var_mes_str + "/" + var_año_str + " " + var_hora_str + ":" + var_minute_str + ":" + VAR_SEGUNDO_STR + ":500 " + var_x
                  var_fecha_pedido = "{d '" + var_año_str + "-" + var_mes_str + "-" + var_dia_str + "'}"
                  'var_dia_str = CStr(Day(rs!FECHA_ENTRADA))
                  var_dia_str = CStr(Day(Date))
                  If Len(var_dia_str) = 1 Then
                     var_dia_str = "0" + var_dia_str
                  End If
                  'var_mes_str = CStr(Month(rs!FECHA_ENTRADA))
                  var_mes_str = CStr(Month(Date))
                  If Len(var_mes_str) = 1 Then
                     var_mes_str = "0" + var_mes_str
                  End If
                  'var_año_str = CStr(Year(rs!FECHA_ENTRADA))
                  var_año_str = CStr(Year(Date))
                  If Len(var_año_str) = 1 Then
                     var_año_str = "0" + var_año_str
                  End If
                  
                  'var_hora_str = CStr(Hour(rs!FECHA_ENTRADA))
                  var_hora_str = CStr(Hour(Date))
                  If Len(var_hora_str) = 1 Then
                     var_hora_str = "0" + var_hora_str
                  End If
                  'var_minute_str = CStr(Minute(rs!FECHA_ENTRADA))
                  var_minute_str = CStr(Minute(Date))
                  If Len(var_minute_str) = 1 Then
                     var_minute_str = "0" + var_minute_str
                  End If
                  'VAR_SEGUNDO_STR = CStr(Second(rs!FECHA_ENTRADA))
                  VAR_SEGUNDO_STR = CStr(Second(Date))
                  If Len(VAR_SEGUNDO_STR) = 1 Then
                     VAR_SEGUNDO_STR = "0" + VAR_SEGUNDO_STR
                  End If
                  'var_X = Replace(UCase(Mid(CStr(rs!FECHA_ENTRADA), 21, 4)), ".", "")
                  var_x = Replace(UCase(Mid(CStr(Date), 21, 4)), ".", "")
                  var_fecha_entrega = var_dia_str + "/" + var_mes_str + "/" + var_año_str + " " + var_hora_str + ":" + var_minute_str + ":" + VAR_SEGUNDO_STR + ":500 " + var_x
                  
                  rsaux2.Open "INSERT INTO TB_temp_ORACLE_LEAD_TIME (INTE_TEM_CONSECUTIVO,PEDIDO, TIENDA, NOMBRE_TIENDA,FECHA_INICIO, INICIO_REPORTE, FIN_REPORTE, FECHA_CREACION) VALUES (" + CStr(var_consecutivo) + "," + CStr(rsaux!pedido) + ",'" + rsaux!tienda + "','" + rsaux!NOMBRE + "',cast(" + var_fecha_pedido + " as datetime)," + var_fecha_inicio + "," + var_fecha_FIN_REPORTE + ", " + var_fecha_creacion + ")", cnn, adOpenDynamic, adLockOptimistic
                  'rsaux2.Open "INSERT INTO TB_temp_ORACLE_LEAD_TIME (INTE_TEM_CONSECUTIVO,PEDIDO, TIENDA, NOMBRE_TIENDA,FECHA_INICIO, INICIO_REPORTE, FIN_REPORTE, FECHA_CREACION) VALUES (" + CStr(var_consecutivo) + "," + CStr(rsaux!pedido) + ",'" + rsaux!tienda + "','" + rsaux!NOMBRE + "',getdate()," + var_fecha_inicio + "," + var_fecha_FIN_REPORTE + ", " + var_fecha_creacion + ")", cnn, adOpenDynamic, adLockOptimistic
               End If
               'rs.Close
               rsaux.MoveNext
         Wend
         rsaux.Close
         'rsaux.Open "select pedido, fecha_final - fecha_inicio as diferencia from tb_temp_oracle_lead_time  WHERE FECHA_FINAL IS NOT NULL AND INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         rsaux.Open "select pedido, fecha_final - fecha_CREACION as diferencia from tb_temp_oracle_lead_time  WHERE FECHA_FINAL IS NOT NULL AND INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               var_diferencia = CStr(rsaux!diferencia)
               var_mes = Mid(var_diferencia, 4, 2)
               VAR_DIAS = Mid(var_diferencia, 1, 2)
               VAR_HORAS = Mid(var_diferencia, 12, 8)
               VAR_CADENA_TIEMPO = ""
               If var_mes > "01" Then
                  'VAR_CADENA_TIEMPO = CStr(CInt(var_mes)) + " MES "
               End If
               
               'VAR_CADENA_TIEMPO = VAR_CADENA_TIEMPO + CStr(CInt(VAR_DIAS)) + " DIAS " + VAR_HORAS + " HORAS"
               VAR_CADENA_TIEMPO = VAR_CADENA_TIEMPO + CStr(CInt(VAR_DIAS)) + " DIAS"
               rsaux1.Open "update tb_temp_oracle_lead_time set diferencia = '" + VAR_CADENA_TIEMPO + "' where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux!pedido), cnn, adOpenDynamic, adLockOptimistic
               rsaux.MoveNext
         Wend
         rsaux.Close
         rsaux.Open "select pedido, fecha_final - fecha_INICIO as diferencia from tb_temp_oracle_lead_time  WHERE FECHA_FINAL IS NOT NULL AND INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               var_diferencia = CStr(rsaux!diferencia)
               var_mes = Mid(var_diferencia, 4, 2)
               VAR_DIAS = Mid(var_diferencia, 1, 2)
               VAR_HORAS = Mid(var_diferencia, 12, 8)
               VAR_CADENA_TIEMPO = ""
               If var_mes > "01" Then
                  'VAR_CADENA_TIEMPO = CStr(CInt(var_mes)) + " MES "
               End If
               
               'VAR_CADENA_TIEMPO = VAR_CADENA_TIEMPO + CStr(CInt(VAR_DIAS)) + " DIAS " + VAR_HORAS + " HORAS"
               VAR_CADENA_TIEMPO = VAR_CADENA_TIEMPO + CStr(CInt(VAR_DIAS)) + " DIAS"
               rsaux1.Open "update tb_temp_oracle_lead_time set diferencia_2 = '" + VAR_CADENA_TIEMPO + "' where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " and pedido = " + CStr(rsaux!pedido), cnn, adOpenDynamic, adLockOptimistic
               rsaux.MoveNext
         Wend
         rsaux.Close
         
         
         
         rsaux.Open "SELECT isnull(MAX(DIFERENCIA),0) as diferencia FROM TB_TEMP_ORACLE_LEAD_TIME WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            rsaux1.Open "UPDATE TB_TEMP_ORACLE_LEAD_TIME SET MAYOR = 1 WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND DIFERENCIA = '" + rsaux(0).Value + "'", cnn, adOpenDynamic, adLockOptimistic
         End If
         rsaux.Close

         rsaux.Open "SELECT isnull(MIN(DIFERENCIA),0) as diferencia FROM TB_TEMP_ORACLE_LEAD_TIME WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            rsaux1.Open "UPDATE TB_TEMP_ORACLE_LEAD_TIME SET MAYOR = 2 WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND DIFERENCIA = '" + rsaux(0).Value + "'", cnn, adOpenDynamic, adLockOptimistic
         End If
         rsaux.Close
         rsaux.Open "SELECT COUNT(*) FROM TB_TEMP_ORACLE_LEAD_TIME WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO IS NOT NULL", cnn, adOpenDynamic, adLockOptimistic
         VAR_COUNT = IIf(IsNull(rsaux(0).Value), 1, rsaux(0).Value)
         rsaux.Close
         var_z = 0
         If var_z = 1 Then
         rsaux.Open "select CAST(SUM(CAST(FECHA_FINAL AS FLOAT) - CAST(FECHA_INICIO AS FLOAT))/" + CStr(VAR_COUNT) + " AS DATETIME) from tb_temp_oracle_lead_time WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND FECHA_FINAL IS NOT NULL", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            var_diferencia = CStr(rsaux(0).Value)
            var_mes = Mid(var_diferencia, 4, 2)
            VAR_DIAS = Mid(var_diferencia, 1, 2)
            VAR_HORAS = Mid(var_diferencia, 12, 8)
            VAR_CADENA_TIEMPO = ""
            If var_mes > "01" Then
               'VAR_CADENA_TIEMPO = CStr(CInt(var_mes)) + " MES "
            End If
            VAR_CADENA_TIEMPO = VAR_CADENA_TIEMPO + CStr(CInt(VAR_DIAS)) + " DIAS " + VAR_HORAS + " HORAS"
            rsaux1.Open "update tb_temp_oracle_lead_time set PROMEDIO = '" + VAR_CADENA_TIEMPO + "' where inte_Tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         End If
         rsaux.Close
         End If
         rsaux.Open "delete from tb_temp_oracle_lead_time where inte_Tem_consecutivo = " + CStr(var_consecutivo) + " and pedido is null", cnn, adOpenDynamic, adLockOptimistic
         
         Set reporte = appl.OpenReport(App.Path + "\REP_ORACLE_LEAD_TIME.rpt")
         reporte.RecordSelectionFormula = "{VW_ORACLE_LEAD_TIME.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
         frmvistasprevias.cr.ReportSource = reporte
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Reporte de lead time"
         frmvistasprevias.Show 1
         Set reporte = Nothing
         var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            x = 0
            If x = 1 Then
               Set reporte = appl.OpenReport(App.Path + "\REP_ORACLE_LEAD_TIME_eXCEL.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_LEAD_TIME.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\reporte_lead_time_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            Else
               Set oexcel = CreateObject("Excel.Application")
               Set owbook = oexcel.Workbooks.Add
               Set osheet = owbook.Worksheets(1)
               osheet.Name = "LEAD TIME"
               Screen.MousePointer = vbHourglass
               iFila = 1
               ifila2 = 1
               icol2 = 1
               iCol = 1
               var_cadena = " SELECT INICIO_REPORTE, FIN_REPORTE, PEDIDO, TIENDA, NOMBRE_TIENDA,  FECHA_CREACION FECHA_CREACION_PEDIDO, FECHA_INICIO FECHA_FACTURACION, FECHA_FINAL FECHA_DE_ENTRADA, DIFERENCIA DIFERENCIA_CREACION, DIFERENCIA_2 DIFERENCIA_FACTURACION, PROMEDIO PROMEDIO_CREACION_ENTRADA FROM VW_ORACLE_LEAD_TIME WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo) + " ORDER BY TIENDA"
               rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               For i = 0 To rsaux10.Fields.Count - 1
                   osheet.Cells(iFila, i + 1) = rsaux10.Fields(i).Name
                   osheet.Cells(iFila, i + 1).Font.Bold = 1
               Next
               iFila = iFila + 1
               With osheet
                    ' carga los registros del recordset
                    .Cells(iFila, iCol).CopyFromRecordset rsaux10
                    'oExcel.Columns(1).Select
                    'oExcel.Selection.NumberFormat = "#,##0.00"
                    'oExcel.Columns(1).Select
                    'oExcel.Selection.Font.Color = vbRed
                    .Columns.AutoFit ' ajusta el ancho de las columnas
                End With
                owbook.SaveAs "c:\reportessid\LEAD_TIME_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                oexcel.Visible = True
                Set oexcel = Nothing
                Screen.MousePointer = vbDefault
                rsaux10.Close
            End If
         
         
         
         
         End If
         'rsaux.Open "delete from tb_temp_oracle_lead_time where inte_Tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      Else
         MsgBox "No existen pedidos para el periodo seleccionado", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a seleccionado ningún centro de negocios", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_invertir_Click()
   n = lv_almacenes.ListItems.Count
   For i = 1 To n
      lv_almacenes.ListItems.Item(i).Selected = True
      If lv_almacenes.selectedItem.SubItems(2) = "*" Then
         lv_almacenes.selectedItem.SubItems(2) = ""
         lv_almacenes.ListItems.Item(i).Bold = False
         lv_almacenes.ListItems.Item(i).ForeColor = &H80000012
         lv_almacenes.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_almacenes.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_almacenes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_almacenes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_almacenes.selectedItem.SubItems(2) = "*"
         lv_almacenes.ListItems.Item(i).Bold = True
         lv_almacenes.ListItems.Item(i).ForeColor = &HFF0000
         lv_almacenes.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_almacenes.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_almacenes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_almacenes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub cmd_marcar_Click()
   i = lv_almacenes.selectedItem.Index
   If lv_almacenes.selectedItem.SubItems(2) = "*" Then
      lv_almacenes.selectedItem.SubItems(2) = ""
      lv_almacenes.ListItems.Item(i).Bold = False
      lv_almacenes.ListItems.Item(i).ForeColor = &H80000012
      lv_almacenes.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_almacenes.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_almacenes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_almacenes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_almacenes.Refresh
   Else
      lv_almacenes.selectedItem.SubItems(2) = "*"
      lv_almacenes.ListItems.Item(i).Bold = True
      lv_almacenes.ListItems.Item(i).ForeColor = &HFF0000
      lv_almacenes.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_almacenes.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_almacenes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_almacenes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_almacenes.Refresh
   End If
End Sub

Private Sub cmd_ninguno_Click()
   n = lv_almacenes.ListItems.Count
   For i = 1 To n
      lv_almacenes.ListItems.Item(i).Selected = True
      lv_almacenes.selectedItem.SubItems(2) = ""
      lv_almacenes.ListItems.Item(i).Bold = False
      lv_almacenes.ListItems.Item(i).ForeColor = &H80000012
      lv_almacenes.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_almacenes.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_almacenes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_almacenes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   lv_almacenes.Refresh
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmd_seleccion_Click()
   n = lv_almacenes.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_almacenes.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_almacenes.selectedItem.SubItems(2) = "" And var_rellena = True Then
         lv_almacenes.selectedItem.SubItems(2) = "*"
         lv_almacenes.ListItems.Item(i).Bold = True
         lv_almacenes.ListItems.Item(i).ForeColor = &HFF0000
         lv_almacenes.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_almacenes.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_almacenes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_almacenes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_almacenes.selectedItem.SubItems(2) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_almacenes.selectedItem.SubItems(2) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub cmd_todos_Click()
   n = lv_almacenes.ListItems.Count
   For i = 1 To n
      lv_almacenes.ListItems.Item(i).Selected = True
      lv_almacenes.selectedItem.SubItems(2) = "*"
      lv_almacenes.ListItems.Item(i).Bold = True
      lv_almacenes.ListItems.Item(i).ForeColor = &HFF0000
      lv_almacenes.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_almacenes.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_almacenes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_almacenes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
   Next i
   lv_almacenes.Refresh
End Sub

Private Sub Form_Load()
   txt_fin = Date
   txt_inicio = Date
   Top = 0
   Left = 2200
   var_conexion_string = "Provider=SQLOLEDB.1;Password=icgfront2013;Persist Security Info=True;User ID=ICGAdmin;Initial Catalog=DBICGPROD_VIA;Data Source=SQLPOSPROD"
   If cnn_lead_time.State = 1 Then
      cnn_lead_time.Close
   End If
   cnn_lead_time.Open var_conexion_string
   
   rs.Open "select secondary_inventory_name as CLAVE, description AS NOMBRE from mtl_subinventories_all_v where attribute3 = 'PTO_VTA' and organization_id = 93 and secondary_inventory_name like '%TD%' ORDER BY DESCRIPTION", cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set list_item = lv_almacenes.ListItems.Add(, , rs!CLAVE)
         list_item.SubItems(1) = IIf(IsNull(rs!NOMBRE), "", rs!NOMBRE)
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
   If cnn_lead_time.State = 1 Then
      cnn_lead_time.Close
   End If
End Sub

Private Sub lv_almacenes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_almacenes, ColumnHeader)
End Sub

Private Sub lv_almacenes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_almacenes.ListItems.Count > 0 Then
         i = lv_almacenes.selectedItem.Index
         If lv_almacenes.selectedItem.SubItems(2) = "*" Then
            lv_almacenes.selectedItem.SubItems(2) = ""
            lv_almacenes.ListItems.Item(i).Bold = False
            lv_almacenes.ListItems.Item(i).ForeColor = &H80000012
            lv_almacenes.ListItems.Item(i).ListSubItems(1).Bold = False
            lv_almacenes.ListItems.Item(i).ListSubItems(2).Bold = False
            lv_almacenes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
            lv_almacenes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
            lv_almacenes.Refresh
         Else
            lv_almacenes.selectedItem.SubItems(2) = "*"
            lv_almacenes.ListItems.Item(i).Bold = True
            lv_almacenes.ListItems.Item(i).ForeColor = &HFF0000
            lv_almacenes.ListItems.Item(i).ListSubItems(1).Bold = True
            lv_almacenes.ListItems.Item(i).ListSubItems(2).Bold = True
            lv_almacenes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
            lv_almacenes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
            lv_almacenes.Refresh
         End If
      End If
   End If
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fin) Then
         frmcalendario.mes = CDate(Me.txt_fin)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      txt_fin = var_fecha_general
   End If
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
