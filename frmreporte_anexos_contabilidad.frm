VERSION 5.00
Begin VB.Form frmreporte_anexos_contabilidad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anexos"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "6"
      Height          =   315
      Left            =   420
      TabIndex        =   5
      ToolTipText     =   "Anexo 6"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   " Fecha "
      Height          =   840
      Left            =   75
      TabIndex        =   3
      Top             =   450
      Width           =   4335
      Begin VB.TextBox txt_inicio 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1035
         TabIndex        =   4
         Top             =   210
         Width           =   1980
      End
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   15
      TabIndex        =   2
      Top             =   360
      Width           =   4485
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Caption         =   "13"
      Height          =   315
      Left            =   90
      TabIndex        =   1
      ToolTipText     =   "Anexo 13"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4050
      Picture         =   "frmreporte_anexos_contabilidad.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
End
Attribute VB_Name = "frmreporte_anexos_contabilidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
   Dim var_consecutivo As Double
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim año As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   
   'On Error GoTo salir:
  If IsDate(txt_inicio) Then
       cnn.BeginTrans
       rs.Open "select max(inte_tem_consecutivo) from TB_TEMP_REPORTE_ANEXO_13", cnn, adOpenDynamic, adLockOptimistic
       If Not rs.EOF Then
          var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
       Else
          var_consecutivo = 0
       End If
       var_consecutivo = var_consecutivo + 1
       rs.Close
       rs.Open "Insert into TB_TEMP_REPORTE_ANEXO_13 (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
       cnn.CommitTrans
        
       var_dia = CStr(Day(CDate(txt_inicio)))
       var_mes = CStr(Month(CDate(txt_inicio)))
       var_año = CStr(Year(CDate(txt_inicio)))
       If Len(Trim(var_dia)) = 1 Then
          var_dia = "0" + var_dia
       End If
       If Len(Trim(var_mes)) = 1 Then
          var_mes = "0" + var_mes
       End If
       var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
       
       var_cadena = "SELECT VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, MAX(INTE_EMO_NUMERO) AS INTE_EMO_NUMERO From dbo.TB_ENCABEZADO_MOVIMIENTOS WHERE (DTIM_EMO_FECHA < " + var_fecha_inicio + ") GROUP BY VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID"
       rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
       While Not rs.EOF
             var_cadena = "INSERT INTO TB_TEMP_REPORTE_ANEXO_13 (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA, VCHA_UOR_UNIDAD_ID,            VCHA_ALM_ALMACEN_ID,              VCHA_MOV_MOVIMIENTO_ID,     INTE_TEM_NUMERO_ANTERIOR, INTE_TEM_NUMERO_POSTERIOR, DTIM_TEM_FECHA_ANTERIOR,  DTIM_TEM_FECHA_POSTERIOR, FLOA_TEM_CANTIDAD_ANTERIOR, FLOA_TEM_CANTIDAD_POSTERIOR, FLOA_TEM_PRECIO_ANTERIOR, FLOA_TEM_PRECIO_POSTERIOR) "
             var_cadena = var_cadena + " VALUES (" + CStr(var_consecutivo) + "," + var_fecha_inicio + ",'" + rs!vcha_uor_unidad_id + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_EMO_NUMERO) + ", 0,               null,NULL,                  0,                          0,                           0,                         0)"
             rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
             rs.MoveNext
       Wend
       rs.Close
       
       var_cadena = "SELECT VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, MIN(INTE_EMO_NUMERO) AS INTE_EMO_NUMERO From dbo.TB_ENCABEZADO_MOVIMIENTOS WHERE (DTIM_EMO_FECHA >= " + var_fecha_inicio + ") GROUP BY VCHA_EMP_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID"
       rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
       While Not rs.EOF
             rsaux1.Open "SELECT * FROM TB_TEMP_REPORTE_ANEXO_13 WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND VCHA_UOR_UNIDAD_ID = '" + rs!vcha_uor_unidad_id + "' AND VCHA_ALM_ALMACEN_ID = '" + rs!VCHA_ALM_ALMACEN_ID + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + rs!VCHA_MOV_MOVIMIENTO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
             If Not rsaux1.EOF Then
                rsaux2.Open "UPDATE TB_TEMP_REPORTE_ANEXO_13 SET INTE_TEM_NUMERO_POSTERIOR = " + CStr(rs!INTE_EMO_NUMERO) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND VCHA_UOR_UNIDAD_ID = '" + rs!vcha_uor_unidad_id + "' AND VCHA_ALM_ALMACEN_ID = '" + rs!VCHA_ALM_ALMACEN_ID + "' AND VCHA_MOV_MOVIMIENTO_ID = '" + rs!VCHA_MOV_MOVIMIENTO_ID + "'", cnn, adOpenDynamic, adLockOptimistic
             Else
                var_cadena = "INSERT INTO TB_TEMP_REPORTE_ANEXO_13 (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA, VCHA_UOR_UNIDAD_ID,            VCHA_ALM_ALMACEN_ID,              VCHA_MOV_MOVIMIENTO_ID,     INTE_TEM_NUMERO_ANTERIOR, INTE_TEM_NUMERO_POSTERIOR, DTIM_TEM_FECHA_ANTERIOR,  DTIM_TEM_FECHA_POSTERIOR, FLOA_TEM_CANTIDAD_ANTERIOR, FLOA_TEM_CANTIDAD_POSTERIOR, FLOA_TEM_PRECIO_ANTERIOR, FLOA_TEM_PRECIO_POSTERIOR) "
                var_cadena = var_cadena + " VALUES (" + CStr(var_consecutivo) + "," + var_fecha_inicio + ",'" + rs!vcha_uor_unidad_id + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', 0, " + CStr(rs!INTE_EMO_NUMERO) + ",               null, NULL,                  0,                          0,                           0,                         0)"
                rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
             End If
             rsaux1.Close
             rs.MoveNext
       Wend
       rs.Close
       
       rs.Open "select * from TB_TEMP_REPORTE_ANEXO_13 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_mov_movimiento_id is not null", cnn, adOpenDynamic, adLockOptimistic
       While Not rs.EOF
             'If IIf(IsNull(rs!INTE_TEM_NUMERO_ANTERIOR), 0, rs!INTE_TEM_NUMERO_ANTERIOR) > 0 Then
                var_cadena = "SELECT     MIN(DTIM_ENT_FECHA) AS FECHA, SUM(FLOA_ENT_CANTIDAD) AS CANTIDAD, SUM(FLOA_ENT_PRECIO * FLOA_ENT_CANTIDAD) AS PRECIO, VCHA_UOR_UNIDAD_ID ,  VCHA_MOV_MOVIMIENTO_ID, INTE_ENT_NUMERO From dbo.TB_ENTRADAS WHERE (VCHA_UOR_UNIDAD_ID = '" + rs!vcha_uor_unidad_id + "')  AND (VCHA_MOV_MOVIMIENTO_ID = '" + rs!VCHA_MOV_MOVIMIENTO_ID + "') AND (INTE_ENT_NUMERO = " + CStr(rs!INTE_TEM_NUMERO_ANTERIOR) + ") GROUP BY VCHA_UOR_UNIDAD_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_ENT_NUMERO"
                rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                If Not rsaux.EOF Then
                   var_dia = CStr(Day(CDate(rsaux!fecha)))
                   var_mes = CStr(Month(CDate(rsaux!fecha)))
                   var_año = CStr(Year(CDate(rsaux!fecha)))
                   If Len(Trim(var_dia)) = 1 Then
                      var_dia = "0" + var_dia
                   End If
                   If Len(Trim(var_mes)) = 1 Then
                      var_mes = "0" + var_mes
                   End If
                   var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                   
                   rsaux1.Open "UPDATE TB_TEMP_REPORTE_ANEXO_13 SET DTIM_TEM_FECHA_ANTERIOR = " + var_fecha_inicio + ", FLOA_TEM_cANTIDAD_ANTERIOR = " + CStr(rsaux!Cantidad) + ", FLOA_TEM_PRECIO_ANTERIOr = " + CStr(IIf(IsNull(rsaux!Precio), 0, rsaux!Precio)) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND (VCHA_UOR_UNIDAD_ID = '" + rs!vcha_uor_unidad_id + "') AND (VCHA_MOV_MOVIMIENTO_ID = '" + rs!VCHA_MOV_MOVIMIENTO_ID + "') AND (INTE_tem_NUMERO_anterior = " + CStr(rs!INTE_TEM_NUMERO_ANTERIOR) + ")"
                End If
                rsaux.Close
             'End If
             
                var_cadena = "select vcha_uor_unidad_id, vcha_mov_movimiento_id, inte_sal_numero, min(dtim_sal_Fecha) as fecha, sum(floa_sal_cantidad) as cantidad, sum(((floa_sal_precio * floa_sal_cantidad)*(1-(floa_sal_Descuento_1/100))) * (1 -(floa_sal_Descuento_2/100)))as precio from tb_salidas where vcha_uor_unidad_id = '" + rs!vcha_uor_unidad_id + "' and vcha_mov_movimiento_id = '" + rs!VCHA_MOV_MOVIMIENTO_ID + "' and inte_sal_numero  = " + CStr(rs!INTE_TEM_NUMERO_ANTERIOR) + " group by vcha_uor_unidad_id, vcha_mov_movimiento_id, inte_sal_numero"
                rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                If Not rsaux.EOF Then
                   var_dia = CStr(Day(CDate(rsaux!fecha)))
                   var_mes = CStr(Month(CDate(rsaux!fecha)))
                   var_año = CStr(Year(CDate(rsaux!fecha)))
                   If Len(Trim(var_dia)) = 1 Then
                      var_dia = "0" + var_dia
                   End If
                   If Len(Trim(var_mes)) = 1 Then
                      var_mes = "0" + var_mes
                   End If
                   var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                   
                   rsaux1.Open "UPDATE TB_TEMP_REPORTE_ANEXO_13 SET DTIM_TEM_FECHA_ANTERIOR = " + var_fecha_inicio + ", FLOA_TEM_cANTIDAD_ANTERIOR = " + CStr(rsaux!Cantidad) + ", FLOA_TEM_PRECIO_ANTERIOR = " + CStr(IIf(IsNull(rsaux!Precio), 0, rsaux!Precio)) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND (VCHA_UOR_UNIDAD_ID = '" + rs!vcha_uor_unidad_id + "') AND (VCHA_MOV_MOVIMIENTO_ID = '" + rs!VCHA_MOV_MOVIMIENTO_ID + "') AND (INTE_tem_NUMERO_ANTERIOR = " + CStr(rs!INTE_TEM_NUMERO_ANTERIOR) + ")"
                End If
                rsaux.Close
             
             
             
                var_cadena = "SELECT     MIN(DTIM_ENT_FECHA) AS FECHA, SUM(FLOA_ENT_CANTIDAD) AS CANTIDAD, SUM(FLOA_ENT_PRECIO * FLOA_ENT_CANTIDAD) AS PRECIO, VCHA_UOR_UNIDAD_ID ,  VCHA_MOV_MOVIMIENTO_ID, INTE_ENT_NUMERO From dbo.TB_ENTRADAS WHERE (VCHA_UOR_UNIDAD_ID = '" + rs!vcha_uor_unidad_id + "')  AND (VCHA_MOV_MOVIMIENTO_ID = '" + rs!VCHA_MOV_MOVIMIENTO_ID + "') AND (INTE_ENT_NUMERO = " + CStr(rs!INTE_TEM_NUMERO_posterior) + ") GROUP BY VCHA_UOR_UNIDAD_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_ENT_NUMERO"
                rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                If Not rsaux.EOF Then
                   var_dia = CStr(Day(CDate(rsaux!fecha)))
                   var_mes = CStr(Month(CDate(rsaux!fecha)))
                   var_año = CStr(Year(CDate(rsaux!fecha)))
                   If Len(Trim(var_dia)) = 1 Then
                      var_dia = "0" + var_dia
                   End If
                   If Len(Trim(var_mes)) = 1 Then
                      var_mes = "0" + var_mes
                   End If
                   var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                   
                   rsaux1.Open "UPDATE TB_TEMP_REPORTE_ANEXO_13 SET DTIM_TEM_FECHA_POSTERIOR = " + var_fecha_inicio + ", FLOA_TEM_cANTIDAD_POSTERIOR = " + CStr(rsaux!Cantidad) + ", FLOA_TEM_PRECIO_POSTERIOR = " + CStr(IIf(IsNull(rsaux!Precio), 0, rsaux!Precio)) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND (VCHA_UOR_UNIDAD_ID = '" + rs!vcha_uor_unidad_id + "') AND (VCHA_MOV_MOVIMIENTO_ID = '" + rs!VCHA_MOV_MOVIMIENTO_ID + "') AND (INTE_tem_NUMERO_POSTERIOR = " + CStr(rs!INTE_TEM_NUMERO_posterior) + ")"
                End If
                rsaux.Close
             
             
             
             
             'If IIf(IsNull(rs!INTE_TEM_NUMERO_POSTERIOR), 0, rs!INTE_TEM_NUMERO_POSTERIOR) > 0 Then
                var_cadena = "select vcha_uor_unidad_id, vcha_mov_movimiento_id, inte_sal_numero, min(dtim_sal_Fecha) as fecha, sum(floa_sal_cantidad) as cantidad, sum(((floa_sal_precio * floa_sal_cantidad)*(1-(floa_sal_Descuento_1/100))) * (1 -(floa_sal_Descuento_2/100)))as precio from tb_salidas where vcha_uor_unidad_id = '" + rs!vcha_uor_unidad_id + "' and vcha_mov_movimiento_id = '" + rs!VCHA_MOV_MOVIMIENTO_ID + "' and inte_sal_numero  = " + CStr(rs!INTE_TEM_NUMERO_posterior) + " group by vcha_uor_unidad_id, vcha_mov_movimiento_id, inte_sal_numero"
                rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                If Not rsaux.EOF Then
                   var_dia = CStr(Day(CDate(rsaux!fecha)))
                   var_mes = CStr(Month(CDate(rsaux!fecha)))
                   var_año = CStr(Year(CDate(rsaux!fecha)))
                   If Len(Trim(var_dia)) = 1 Then
                      var_dia = "0" + var_dia
                   End If
                   If Len(Trim(var_mes)) = 1 Then
                      var_mes = "0" + var_mes
                   End If
                   var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                   
                   rsaux1.Open "UPDATE TB_TEMP_REPORTE_ANEXO_13 SET DTIM_TEM_FECHA_POSTERIOR = " + var_fecha_inicio + ", FLOA_TEM_cANTIDAD_POSTERIOR = " + CStr(rsaux!Cantidad) + ", FLOA_TEM_PRECIO_POSTERIOr = " + CStr(IIf(IsNull(rsaux!Precio), 0, rsaux!Precio)) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND (VCHA_UOR_UNIDAD_ID = '" + rs!vcha_uor_unidad_id + "') AND (VCHA_MOV_MOVIMIENTO_ID = '" + rs!VCHA_MOV_MOVIMIENTO_ID + "') AND (INTE_tem_NUMERO_POSTERIOR = " + CStr(rs!INTE_TEM_NUMERO_posterior) + ")"
                End If
                rsaux.Close
             'End If
             
             
             
             rs.MoveNext
       Wend
       rs.Close
       
       
       
       
       
       
       
       
       
       var_cadena = "delete from TB_TEMP_REPORTE_ANEXO_13 where dtim_tem_fecha_anterior is null and dtim_tem_fecha_posterior is null"
       rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
      
       Set reporte = appl.OpenReport(App.Path + "\rep_anexo_13.rpt")
       reporte.RecordSelectionFormula = "{VW_TEMP_REPORTE_ANEXO_13.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
       For ntablas = 1 To reporte.Database.Tables.Count
           reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
       Next ntablas
       reporte.ExportOptions.FormatType = crEFTExcel80
       reporte.ExportOptions.DestinationType = crEDTDiskFile
       archivo = "c:\reportessid\Anexo_13_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
       reporte.ExportOptions.DiskFileName = archivo
       reporte.Export False
       Set reporte = Nothing
       MsgBox "Se a terminado de guardar el archivo " + archivo
      
      
      rs.Open "delete from TB_TEMP_REPORTE_ANEXO_13 where INTE_Tem_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
   Else
      MsgBox "Fecha de Inicio Incorrecta", vbOKOnly, "ATENCION"
   End If
   Exit Sub
salir:
   MsgBox "A surgido un error al generar el reporte", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   Dim var_consecutivo As Double
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim año As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   
   'On Error GoTo salir:
  If IsDate(txt_inicio) Then
       Dim var_fecha As Date
       
       cnn.BeginTrans
       If rs.State = 1 Then
          rs.Close
       End If
       rs.Open "select max(inte_tem_consecutivo) from TB_TEMP_REPORTE_ANEXO_6", cnn, adOpenDynamic, adLockOptimistic
       If Not rs.EOF Then
          var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
       Else
          var_consecutivo = 0
       End If
       var_consecutivo = var_consecutivo + 1
       rs.Close
       rs.Open "Insert into TB_TEMP_REPORTE_ANEXO_6 (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
       cnn.CommitTrans
        
       var_dia = CStr(Day(CDate(txt_inicio)))
       var_mes = CStr(Month(CDate(txt_inicio)))
       var_año = CStr(Year(CDate(txt_inicio)))
       If Len(Trim(var_dia)) = 1 Then
          var_dia = "0" + var_dia
       End If
       If Len(Trim(var_mes)) = 1 Then
          var_mes = "0" + var_mes
       End If
       var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
       
       
       var_cadena = "SELECT VCHA_EMP_EMPRESA_ID, VCHA_SER_SERIE_ID, VCHA_CAR_TIPO_DOCUMENTO, MAX(DTIM_CAR_FECHA) AS FECHA From dbo.TB_ENCABEZADO_CARTERA WHERE (DTIM_CAR_FECHA < " + var_fecha_inicio + ") AND VCHA_CAR_TIPO_DOCUMENTO IN ('FA','NG','NC') and (CHAR_CAR_ESTATUS <> 'C' OR CHAR_CAR_ESTATUS IS NULL) GROUP BY VCHA_EMP_EMPRESA_ID, VCHA_SER_SERIE_ID, VCHA_CAR_TIPO_DOCUMENTO"
       rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
       While Not rs.EOF
             var_dia = CStr(Day((rs!fecha)))
             var_mes = CStr(Month((rs!fecha)))
             var_año = CStr(Year((rs!fecha)))
             var_minuto = CStr(Minute((rs!fecha)))
             var_segundo = CStr(Second((rs!fecha)))
             var_hora = CStr(Hour(CDate(rs!fecha)))
             If Len(Trim(var_dia)) = 1 Then
                var_dia = "0" + var_dia
             End If
             If Len(Trim(var_mes)) = 1 Then
                var_mes = "0" + var_mes
             End If
             If Len(Trim(var_hora)) = 1 Then
                var_hora = "0" + var_hora
             End If
             If Len(Trim(var_minuto)) = 1 Then
                var_minuto = "0" + var_minuto
             End If
             If Len(Trim(var_segundo)) = 1 Then
                var_segundo = "0" + var_segundo
             End If
             var_año = Mid(var_año, 3, 2)
             var_fecha_cartera = x
             cFechainicio = Format(rs!fecha, "dd-mm-yy hh:mm:ss")
             cFechainicio = cFechainicio
             var_cadena = "INSERT INTO TB_TEMP_REPORTE_ANEXO_6 (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA, VCHA_EMP_EMPRESA_ID,            VCHA_SER_SERIE_ID,              VCHA_CAR_DOCUMENTO,     INTE_TEM_NUMERO_ANTERIOR, DTIM_TEM_FECHA_ANTERIOR, FLOA_TEM_IMPORTE_NETO_ANTERIOR, FLOA_TEM_TIPO_CAMBIO_ANTERIOR,  VCHA_TEM_MONEDA_ANTERIOR, INTE_TEM_NUMERO_POSTERIOR, DTIM_TEM_FECHA_POSTERIOR, FLOA_TEM_IMPORTE_NETO_POSTERIOR, FLOA_TEM_TIPO_CAMBIO_POSTERIOR, VCHA_TEM_MONEDA_POSTERIOR) "
             var_cadena = var_cadena + " VALUES (" + CStr(var_consecutivo) + "," + var_fecha_inicio + ",'" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!vcha_ser_Serie_id + "', '" + rs!vcha_Car_tipo_documento + "', 0, cast('" + cFechainicio + "' as datetime),               0,                  0,                              '',                           0,                    NULL,                     0,                          0,                               '')"
             'MsgBox var_cadena
             rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
             rs.MoveNext
        Wend
       rs.Close
       
       var_cadena = "SELECT VCHA_EMP_EMPRESA_ID, VCHA_SER_SERIE_ID, VCHA_CAR_TIPO_DOCUMENTO, MIN(DTIM_CAR_fecha) AS FECHA From dbo.TB_ENCABEZADO_CARTERA WHERE (DTIM_CAR_FECHA >= " + var_fecha_inicio + ") AND VCHA_CAR_TIPO_DOCUMENTO IN ('FA','NG','NC') and (CHAR_CAR_ESTATUS <> 'C' OR CHAR_CAR_ESTATUS IS NULL) GROUP BY VCHA_EMP_EMPRESA_ID, VCHA_SER_SERIE_ID, VCHA_CAR_TIPO_DOCUMENTO"
       rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
       While Not rs.EOF
             var_dia = CStr(Day((rs!fecha)))
             var_mes = CStr(Month((rs!fecha)))
             var_año = CStr(Year((rs!fecha)))
             var_minuto = CStr(Minute((rs!fecha)))
             var_segundo = CStr(Second((rs!fecha)))
             var_hora = CStr(Hour(CDate(rs!fecha)))
             If Len(Trim(var_dia)) = 1 Then
                var_dia = "0" + var_dia
             End If
             If Len(Trim(var_mes)) = 1 Then
                var_mes = "0" + var_mes
             End If
             If Len(Trim(var_hora)) = 1 Then
                var_hora = "0" + var_hora
             End If
             If Len(Trim(var_minuto)) = 1 Then
                var_minuto = "0" + var_minuto
             End If
             If Len(Trim(var_segundo)) = 1 Then
                var_segundo = "0" + var_segundo
             End If
             var_año = Mid(var_año, 3, 2)
             var_fecha_cartera = x
             cFechainicio = Format(rs!fecha, "dd-mm-yy hh:mm:ss")
             cFechainicio = cFechainicio
             
             rsaux1.Open "SELECT * FROM TB_TEMP_REPORTE_ANEXO_6 WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND VCHA_EMP_EMPRESA_ID = '" + rs!VCHA_EMP_EMPRESA_ID + "' AND VCHA_CAR_DOCUMENTO = '" + rs!vcha_Car_tipo_documento + "' AND VCHA_SER_SERIE_ID = '" + rs!vcha_ser_Serie_id + "'", cnn, adOpenDynamic, adLockOptimistic
             If Not rsaux1.EOF Then
                rsaux2.Open "UPDATE TB_TEMP_REPORTE_ANEXO_6 SET dtim_tem_Fecha_posterior = cast('" + cFechainicio + "' as smalldatetime) WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND VCHA_eMP_EMPRESA_ID = '" + rs!VCHA_EMP_EMPRESA_ID + "' AND VCHA_cAR_DOCUMENTO = '" + rs!vcha_Car_tipo_documento + "' AND VCHA_SER_SERIE_ID = '" + rs!vcha_ser_Serie_id + "'", cnn, adOpenDynamic, adLockOptimistic
             Else
                var_cadena = "INSERT INTO TB_TEMP_REPORTE_ANEXO_6 (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA, VCHA_EMP_EMPRESA_ID,            VCHA_SER_SERIE_ID,              VCHA_CAR_DOCUMENTO,     INTE_TEM_NUMERO_ANTERIOR, DTIM_TEM_FECHA_ANTERIOR, FLOA_TEM_IMPORTE_NETO_ANTERIOR, FLOA_TEM_TIPO_CAMBIO_ANTERIOR,  VCHA_TEM_MONEDA_ANTERIOR, INTE_TEM_NUMERO_POSTERIOR, DTIM_TEM_FECHA_POSTERIOR, FLOA_TEM_IMPORTE_NETO_POSTERIOR, FLOA_TEM_TIPO_CAMBIO_POSTERIOR, VCHA_TEM_MONEDA_POSTERIOR) "
                var_cadena = var_cadena + " VALUES (" + CStr(var_consecutivo) + "," + var_fecha_inicio + ",'" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!vcha_ser_Serie_id + "', '" + rs!vcha_Car_tipo_documento + "', 0,           NULL,                    0,                         0,                              '',                       0,                         cast('" + cFechainicio + "' as datetime),                     0,                          0,                               '')"
                rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
             End If
             rsaux1.Close
             rs.MoveNext
       Wend
       rs.Close
       
       rs.Open "select * from TB_TEMP_REPORTE_ANEXO_6 where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and vcha_Car_documento is not null", cnn, adOpenDynamic, adLockOptimistic
       While Not rs.EOF
       
       
       
             var_fecha_Valida = IIf(IsNull(rs!INTE_TEM_NUMERO_ANTERIOR), "", rs!INTE_TEM_NUMERO_ANTERIOR)
             If Len(var_fecha_Valida) > 0 Then
                cFechainicio = Format(rs!dtim_Tem_Fecha_anterior, "dd-mm-yy hh:mm:ss")
                cFechainicio = cFechainicio
                var_cadena = "SELECT  * From tb_Encabezado_cartera WHERE vcha_Emp_empresa_id = '" + rs!VCHA_EMP_EMPRESA_ID + "' and vcha_Car_TIPO_documento = '" + rs!vcha_Car_documento + "' and vcha_Ser_serie_id = '" + rs!vcha_ser_Serie_id + "' and cast(dtim_Car_fecha as smalldatetime) = cast('" + cFechainicio + "' as smalldatetime) "
                rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                If Not rsaux.EOF Then
                   rsaux1.Open "UPDATE TB_TEMP_REPORTE_ANEXO_6 SET inte_tem_numero_anterior = " + CStr(rsaux!inte_car_numero) + ", FLOA_TEM_IMPORTE_NETO_ANTERIOR = " + CStr(IIf(IsNull(rsaux!floa_Car_importe_neto), 0, rsaux!floa_Car_importe_neto)) + ", floa_tem_tipo_cambio_anterior = " + CStr(IIf(IsNull(rsaux!floa_car_tipo_cambio), 1, rsaux!floa_car_tipo_cambio)) + ", vcha_TEM_MONEDA_ANTERIOR = '" + IIf(IsNull(rsaux!vcha_mon_moneda_id), "", rsaux!vcha_mon_moneda_id) + "' WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND (vcha_emp_empresa_id = '" + rs!VCHA_EMP_EMPRESA_ID + "') AND (vcha_Car_documento = '" + rs!vcha_Car_documento + "') and vcha_ser_Serie_id = '" + rs!vcha_ser_Serie_id + "' AND cast(dtim_Tem_Fecha_Anterior as smalldatetime) = cast('" + cFechainicio + "' as smalldatetime)", cnn, adOpenDynamic, adLockOptimistic
                End If
                rsaux.Close
             End If
             
             var_fecha_Valida = IIf(IsNull(rs!INTE_TEM_NUMERO_posterior), "", rs!INTE_TEM_NUMERO_posterior)
             If Len(var_fecha_Valida) > 0 Then
                cFechainicio = Format(rs!dtim_Tem_Fecha_posterior, "dd-mm-yy hh:mm:ss")
                cFechainicio = cFechainicio
                var_cadena = "SELECT  * From tb_Encabezado_cartera WHERE vcha_Emp_empresa_id = '" + rs!VCHA_EMP_EMPRESA_ID + "' and vcha_Car_TIPO_documento = '" + rs!vcha_Car_documento + "' and vcha_Ser_serie_id = '" + rs!vcha_ser_Serie_id + "' and cast(dtim_Car_fecha as smalldatetime) =  cast('" + cFechainicio + "' as smalldatetime) "
                rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                If Not rsaux.EOF Then
                   'MsgBox "UPDATE TB_TEMP_REPORTE_ANEXO_6 SET inte_tem_numero_posterior = " + CStr(rsaux!inte_car_numero) + ", FLOA_TEM_IMPORTE_NETO_POSTERIOR = " + CStr(IIf(IsNull(rsaux!floa_Car_importe_neto), 0, rsaux!floa_Car_importe_neto)) + ", floa_tem_tipo_cambio_posterior = " + CStr(IIf(IsNull(rsaux!floa_car_tipo_cambio), 1, rsaux!floa_car_tipo_cambio)) + ", VCHA_TEM_MONEDA_POSTERIOR = '" + IIf(IsNull(rsaux!vcha_mon_moneda_id), "1", rsaux!vcha_mon_moneda_id) + "' WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND (vcha_emp_empresa_id = '" + rs!VCHA_EMP_EMPRESA_ID + "') AND (vcha_Car_documento = '" + rs!vcha_Car_documento + "') and vcha_ser_Serie_id = '" + rs!vcha_ser_Serie_id + "' AND cast(dtim_Tem_fecha_posterior as smalldatetime) =  cast('" + cFechainicio + "' as smalldatetime)+ "
                   rsaux1.Open "UPDATE TB_TEMP_REPORTE_ANEXO_6 SET inte_tem_numero_posterior = " + CStr(rsaux!inte_car_numero) + ", FLOA_TEM_IMPORTE_NETO_POSTERIOR = " + CStr(IIf(IsNull(rsaux!floa_Car_importe_neto), 0, rsaux!floa_Car_importe_neto)) + ", floa_tem_tipo_cambio_posterior = " + CStr(IIf(IsNull(rsaux!floa_car_tipo_cambio), 1, rsaux!floa_car_tipo_cambio)) + ", VCHA_TEM_MONEDA_POSTERIOR = '" + IIf(IsNull(rsaux!vcha_mon_moneda_id), "1", rsaux!vcha_mon_moneda_id) + "' WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND (vcha_emp_empresa_id = '" + rs!VCHA_EMP_EMPRESA_ID + "') AND (vcha_Car_documento = '" + rs!vcha_Car_documento + "') and vcha_ser_Serie_id = '" + rs!vcha_ser_Serie_id + "' AND cast(dtim_Tem_fecha_posterior as smalldatetime) =  cast('" + cFechainicio + "' as smalldatetime) ", cnn, adOpenDynamic, adLockOptimistic
                End If
                rsaux.Close
             End If
             
             
             
             
             
             
             rs.MoveNext
       Wend
       rs.Close
       
       
       
       
       
       
       
       
       
       var_cadena = "delete from TB_TEMP_REPORTE_ANEXO_6 where dtim_tem_fecha_anterior is null and dtim_tem_fecha_posterior is null"
       rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
      
       Set reporte = appl.OpenReport(App.Path + "\rep_anexo_6.rpt")
       reporte.RecordSelectionFormula = "{VW_REPORTE_ANEXO_6.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
       For ntablas = 1 To reporte.Database.Tables.Count
           reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
       Next ntablas
       reporte.ExportOptions.FormatType = crEFTExcel80
       reporte.ExportOptions.DestinationType = crEDTDiskFile
       archivo = "c:\reportessid\Anexo_6_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
       reporte.ExportOptions.DiskFileName = archivo
       reporte.Export False
       Set reporte = Nothing
       MsgBox "Se a terminado de guardar el archivo " + archivo
      
      
      rs.Open "delete from TB_TEMP_REPORTE_ANEXO_6 where INTE_Tem_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
   Else
      MsgBox "Fecha de Inicio Incorrecta", vbOKOnly, "ATENCION"
   End If
   Exit Sub
salir:
   MsgBox "A surgido un error al generar el reporte", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 3000
   Left = 3500
   txt_inicio = Date
   txt_fin = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
   var_swpassword = False
   var_modifica_registro = False
   Call activa_forma(var_activa_forma_articulos2)
End Sub


Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_inicio) Then
         frmcalendario.mes.Value = CDate(Me.txt_inicio)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_inicio = var_fecha_general
   End If
End Sub



