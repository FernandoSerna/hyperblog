VERSION 5.00
Begin VB.Form frmoracle_reporte_validacion_doc_fiscales_vs_eflow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte Oracle VS Eflow"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "FAEECA"
      Height          =   315
      Left            =   720
      TabIndex        =   9
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   360
      Picture         =   "frmoracle_reporte_validacion_doc_fiscales_vs_eflow.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   0
      Picture         =   "frmoracle_reporte_validacion_doc_fiscales_vs_eflow.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3930
      Picture         =   "frmoracle_reporte_validacion_doc_fiscales_vs_eflow.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   840
      Left            =   45
      TabIndex        =   0
      Top             =   435
      Width           =   4245
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   2865
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   300
         Width           =   1140
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2325
         TabIndex        =   4
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   375
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   15
      TabIndex        =   7
      Top             =   270
      Width           =   4275
   End
End
Attribute VB_Name = "frmoracle_reporte_validacion_doc_fiscales_vs_eflow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter

Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim año As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_consecutivo As Integer
   Dim var_afectacion_movimiento As String
   Dim var_vistas As String
   Dim cn As New ADODB.Connection
   Dim DSN As String
   Dim cn2 As New ADODB.Connection
   'Dim DSN As String
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            var_dia = CStr(Day(CDate(txt_inicio)))
            var_mes = CStr(Month(CDate(txt_inicio)))
            var_año = CStr(Year(CDate(txt_inicio)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            
            var_fecha_inicio = var_dia + "/" + var_mes + "/" + var_año
            VAR_FECHA_INICIO_TABLA = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            var_fecha_fin_1 = CDate(txt_fin)
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            
            VAR_FECHA_FIN_TABLA = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            
            var_fecha_fin_1 = CDate(txt_fin) + 1
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            
            
            var_fecha_fin = var_dia + "/" + var_mes + "/" + var_año
            rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_cadena = "select SERIE, SERIE||TRX_NUMBER AS FOLIO, TRX_DATE AS FECHA, BILL_CTE_LOC AS TITULAR, BILL_CUST_NAME AS CLIENTE, CUSTOMER_tRX_ID from xxvia_vw_documento_fiscales where printing_option IN ('PRI', 'REP') and TRX_DATE >= TO_DATE('" + var_fecha_inicio + "','DD/MM/YYYY') AND TRX_DATE < TO_DATE('" + var_fecha_fin + "','DD/MM/YYYY')"
            'var_cadena = "select SERIE||TRX_NUMBER AS FOLIO, TRX_DATE AS FECHA, BILL_CTE_LOC AS TITULAR, BILL_CUST_NAME AS CLIENTE, CUSTOMER_tRX_ID from xxvia_vw_documento_fiscales where printing_option IN ('PRI', 'REP') and  customer_Trx_id in "
            'var_cadena = var_cadena + "(44832664,44835934,44835935,44835936,44835937,44835939,44835941,44835943,44835945,44835952,44835954,44835956,44835957,44835960,44835962,44835964,44835965,44835966,44835968,44847937,44847939,44847940,44847941,44847942,44847943,44847944,44847946,44847948,44847949,44847950,44847951,44847953,44847955,44847956,44847957,44852008,44852010,44852011,44852012,44852013,44852014,44851984,44852015,44852016,44852017,44852018,44852019,44852020,44856437,44856438,44856439,44856440,44856441,44856442,44856444,44856445,44856446,44856448,44856450,44856451,44856452,44856454,44856455,44859477,44859478,44859480,44859481,44859483,44859484,44859485,44859487,44859488,44859490,44859492,44859493,44859495,44859497,44859498,44859499,44859501,44859503,44859505,44859738,44859740,44859741,44852053,44814758,44835230,44835231,44835232)"
            

           rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               cnn.BeginTrans
               rsaux.Open " SELECT MAX(INTE_tEM_CONSECUTIVO) FROM TB_TEMP_ORACLE_DOC_FISCALES_VS_EFLOW", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
               Else
                  var_consecutivo = 0
               End If
               rsaux.Close
               var_consecutivo = var_consecutivo + 1
               rsaux.Open "INSERT INTO TB_TEMP_ORACLE_DOC_FISCALES_VS_EFLOW (INTE_TEM_cONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
               cnn.CommitTrans
               DSN = "eflow"
               cn.Open ("DSN=" & DSN & ";")
               'DSN = "EFLOW2"
               'cn2.Open ("DSN=" + DSN + ";")
               While Not rs.EOF
                     var_dia = CStr(Day(CDate(rs!Fecha)))
                     var_mes = CStr(Month(CDate(rs!Fecha)))
                     var_año = CStr(Year(CDate(rs!Fecha)))
                     If Len(Trim(var_dia)) = 1 Then
                        var_dia = "0" + var_dia
                     End If
                     If Len(Trim(var_mes)) = 1 Then
                        var_mes = "0" + var_mes
                     End If
            
                     var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                     var_folio = rs!Folio
                     If CDate(rs!Fecha) < CDate("28/01/2021") Then
                        If rs!Serie = "FAEMXX" Then
                           var_folio = Replace(rs!Folio, "FAEMXX", "FAEMX")
                        End If
                        If rs!Serie = "FAEVBII" Then
                           var_folio = Replace(rs!Folio, "FAEVBII", "FAEVBI")
                        End If
                        If rs!Serie = "FAEVVXX" Then
                           var_folio = Replace(rs!Folio, "FAEVVXX", "FAEVXX")
                        End If
                     Else
                        var_folio = rs!Folio
                     End If
                     Set rsaux1 = cn.execute("SELECT * FROM facturas where factura = '" + var_folio + "'")
                     var_customer_trx_id = IIf(IsNull(rs!customer_Trx_id), "", rs!customer_Trx_id)
                     VAR_VERSION = ""
                     If Not rsaux1.EOF Then
                        var_documento_eflow = IIf(IsNull(rsaux1!FACTURA), "", rsaux1!FACTURA)
                        VAR_ESTATUS = IIf(IsNull(rsaux1!estatus), 0, rsaux1!estatus)
                        VAR_SAT_UUID = IIf(IsNull(rsaux1!sat_uuid), "", rsaux1!sat_uuid)
                        VAR_VERSION = "3.2"
                        'MsgBox CStr(rsaux1!Fecha)
                     Else
                        VAR_ESTATUS = 0
                        var_documento_eflow = ""
                        VAR_SAT_UUID = ""
                        VAR_VERSION = ""
                     End If
                     
                     If var_documento_eflow = "x" Then

                        Set rsaux1 = cn.execute("SELECT * FROM facturas where factura = '" + rs!Folio + "'")
                        If Not rsaux1.EOF Then
                           var_documento_eflow = IIf(IsNull(rsaux1!FACTURA), "", rsaux1!FACTURA)
                           VAR_ESTATUS = IIf(IsNull(rsaux1!estatus), 0, rsaux1!estatus)
                           VAR_SAT_UUID = IIf(IsNull(rsaux1!sat_uuid), "", rsaux1!sat_uuid)
                           'MsgBox CStr(rsaux1!Fecha)
                           VAR_VERSION = "3.3"
                           
                        Else
                           VAR_ESTATUS = 0
                           var_documento_eflow = ""
                           VAR_SAT_UUID = ""
                           VAR_VERSION = ""
                        End If
                     End If
                     
                     rsaux.Open "INSERT INTO TB_TEMP_ORACLE_DOC_FISCALES_VS_EFLOW (INTE_TEM_CONSECUTIVO, DOCUMENTO_ORACLE, FECHA, TITULAR, CLIENTE, DOCUMENTO_EFLOW, FECHA_INICIO, FECHA_FIN, estatus, sat_uuid, VERSION, CUSTOMER_TRX_ID) VALUES (" + CStr(var_consecutivo) + ",'" + IIf(IsNull(rs!Folio), "", rs!Folio) + "'," + var_fecha + ",'" + Replace(IIf(IsNull(rs!TITULAR), "", rs!TITULAR), "'", "") + "','" + Replace(IIf(IsNull(rs!Cliente), "", rs!Cliente), "'", "") + "','" + var_documento_eflow + "'," + VAR_FECHA_INICIO_TABLA + "," + VAR_FECHA_FIN_TABLA + "," + CStr(VAR_ESTATUS) + ", '" + Mid(CStr(VAR_SAT_UUID), 1, 500) + "','" + VAR_VERSION + "'," + CStr(var_customer_trx_id) + ")", cnn, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
                     'MsgBox rs(4).Value
               Wend
               rsaux2.Open "delete from TB_TEMP_ORACLE_DOC_FISCALES_VS_EFLOW WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " and fecha_inicio is null", cnn, adOpenDynamic, adLockOptimistic
               x = 0
               If x = 1 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_oracle_vs_eflow.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_ORACLE_VS_EFLOW.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\rep_comparacion_oracle_vs_eflow_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               Else
                  
                  Set oexcel = CreateObject("Excel.Application")
                  Set owbook = oexcel.Workbooks.Add
                  Set osheet = owbook.Worksheets(1)
                  osheet.Name = "ORACLE VS EFLOW"
                  Screen.MousePointer = vbHourglass
                  iFila = 1
                  ifila2 = 1
                  icol2 = 1
                  iCol = 1
                  var_cadena = "select distinct FECHA_INICIO, FECHA_FIN, FECHA, DOCUMENTO_ORACLE, DOCUMENTO_EFLOW, ESTATUS, SAT_UUID, TITULAR, CLIENTE, VERSION, CUSTOMER_TRX_ID  from TB_TEMP_ORACLE_DOC_FISCALES_VS_EFLOW WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND DOCUMENTO_ORACLE IS NOT NULL order by DOCUMENTO_ORACLE"
                  rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  For i = 0 To rsaux10.Fields.Count - 1
                      osheet.Cells(iFila, i + 1) = rsaux10.Fields(i).Name
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
                  archivo = "c:\reportessid\rep_comparacion_oracle_vs_eflow_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                  owbook.SaveAs archivo
                  oexcel.Visible = True
                  Set oexcel = Nothing
                  Screen.MousePointer = vbDefault
                  rsaux10.Close
                  
                  
                  
                  
                  
               End If
               MsgBox "Se a terminado de guardar el archivo " + archivo
               If rsaux1.State = 1 Then
                  rsaux1.Close
               End If
            Else
               MsgBox "No existen documentos para el periodo indicado", vbOKOnly, "ATENCION"
            End If
            rs.Close
            rs.Open "delete from TB_TEMP_ORACLE_DOC_FISCALES_VS_EFLOW WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "La fecha de inicio debe de ser menor o igual a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de Inicio incorrecta", vbOKOnly, "ATENCION"
   End If
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   If rsaux1.State = 1 Then
      rsaux1.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If rsaux3.State = 1 Then
      rsaux3.Close
   End If

End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim año As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_consecutivo As Integer
   Dim var_afectacion_movimiento As String
   Dim var_vistas As String
   Dim cn As New ADODB.Connection
   Dim DSN As String
   Dim cn2 As New ADODB.Connection
               DSN = "eflow"
               cn.Open ("DSN=" & DSN & ";")
   
   rs.Open "SELECT CUSTOMER_tRX_ID, SERIE||TO_CHAR(NUMERO) AS FOLIO FROM XXVIA_tB_cONTROL_DOC_FISCALES WHERE FECHA_CREACION >= TO_DATE('01/01/2018','DD/MM/YYYY') and nvl(SERIE||TO_CHAR(NUMERO),' ') <>' ' and nvl(to_char(cadena_original),' ') = ' ' AND SUBSTR(SERIE,1,2) = 'FA'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set rsaux1 = cn.execute("SELECT * FROM facturas where factura = '" + rs!Folio + "'")
         If Not rsaux1.EOF Then
            VAR_SAT_UUID = IIf(IsNull(rsaux1!sat_uuid), "", rsaux1!sat_uuid)
         Else
            VAR_VERSION = ""
         End If
         strconsulta = "UPDATE XXVIA_TB_cONTROL_DOC_FISCALES SET CADENA_ORIGINAL = ? WHERE CUSTOMER_TRX_ID = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, VAR_SAT_UUID)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, rs!customer_Trx_id)
              .Parameters.Append parametro
         End With
         Set rsaux9 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         
         rs.MoveNext
                     
   Wend
   rs.Close
   MsgBox "termino"
End Sub

Private Sub Command2_Click()
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim año As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_consecutivo As Integer
   Dim var_afectacion_movimiento As String
   Dim var_vistas As String
   Dim cn As New ADODB.Connection
   Dim DSN As String
   Dim cn2 As New ADODB.Connection
   'Dim DSN As String
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            var_dia = CStr(Day(CDate(txt_inicio)))
            var_mes = CStr(Month(CDate(txt_inicio)))
            var_año = CStr(Year(CDate(txt_inicio)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            
            var_fecha_inicio = var_dia + "/" + var_mes + "/" + var_año
            VAR_FECHA_INICIO_TABLA = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            var_fecha_fin_1 = CDate(txt_fin)
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            
            VAR_FECHA_FIN_TABLA = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            
            var_fecha_fin_1 = CDate(txt_fin) + 1
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            
            
            var_fecha_fin = var_dia + "/" + var_mes + "/" + var_año
            rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_cadena = "SELECT FOLIO, FECHA, TITULAR, CLIENTE, CUSTOMER_tRX_ID  FROM XXVIA_VW_DOC_FISCALES_FAEECA WHERE FECHA >= TO_DATE('" + var_fecha_inicio + "','DD/MM/YYYY') AND FECHA < TO_DATE('" + var_fecha_fin + "','DD/MM/YYYY')"
           rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               cnn.BeginTrans
               rsaux.Open " SELECT MAX(INTE_tEM_CONSECUTIVO) FROM TB_TEMP_ORACLE_DOC_FISCALES_VS_EFLOW", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux.EOF Then
                  var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
               Else
                  var_consecutivo = 0
               End If
               rsaux.Close
               var_consecutivo = var_consecutivo + 1
               rsaux.Open "INSERT INTO TB_TEMP_ORACLE_DOC_FISCALES_VS_EFLOW (INTE_TEM_cONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
               cnn.CommitTrans
               DSN = "eflow"
               cn.Open ("DSN=" & DSN & ";")
               'DSN = "EFLOW2"
               'cn2.Open ("DSN=" + DSN + ";")
               While Not rs.EOF
                     var_dia = CStr(Day(CDate(rs!Fecha)))
                     var_mes = CStr(Month(CDate(rs!Fecha)))
                     var_año = CStr(Year(CDate(rs!Fecha)))
                     If Len(Trim(var_dia)) = 1 Then
                        var_dia = "0" + var_dia
                     End If
                     If Len(Trim(var_mes)) = 1 Then
                        var_mes = "0" + var_mes
                     End If
            
                     var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                     Set rsaux1 = cn.execute("SELECT * FROM facturas where factura = '" + rs!Folio + "'")
                     var_customer_trx_id = IIf(IsNull(rs!customer_Trx_id), "", rs!customer_Trx_id)
                     VAR_VERSION = ""
                     If Not rsaux1.EOF Then
                        var_documento_eflow = IIf(IsNull(rsaux1!FACTURA), "", rsaux1!FACTURA)
                        VAR_ESTATUS = IIf(IsNull(rsaux1!estatus), 0, rsaux1!estatus)
                        VAR_SAT_UUID = IIf(IsNull(rsaux1!sat_uuid), "", rsaux1!sat_uuid)
                        VAR_VERSION = "3.2"
                        'MsgBox CStr(rsaux1!Fecha)
                     Else
                        VAR_ESTATUS = 0
                        var_documento_eflow = ""
                        VAR_SAT_UUID = ""
                        VAR_VERSION = ""
                     End If
                     
                     If var_documento_eflow = "x" Then

                        Set rsaux1 = cn.execute("SELECT * FROM facturas where factura = '" + rs!Folio + "'")
                        If Not rsaux1.EOF Then
                           var_documento_eflow = IIf(IsNull(rsaux1!FACTURA), "", rsaux1!FACTURA)
                           VAR_ESTATUS = IIf(IsNull(rsaux1!estatus), 0, rsaux1!estatus)
                           VAR_SAT_UUID = IIf(IsNull(rsaux1!sat_uuid), "", rsaux1!sat_uuid)
                           'MsgBox CStr(rsaux1!Fecha)
                           VAR_VERSION = "3.3"
                           
                        Else
                           VAR_ESTATUS = 0
                           var_documento_eflow = ""
                           VAR_SAT_UUID = ""
                           VAR_VERSION = ""
                        End If
                     End If
                     
                     rsaux.Open "INSERT INTO TB_TEMP_ORACLE_DOC_FISCALES_VS_EFLOW (INTE_TEM_CONSECUTIVO, DOCUMENTO_ORACLE, FECHA, TITULAR, CLIENTE, DOCUMENTO_EFLOW, FECHA_INICIO, FECHA_FIN, estatus, sat_uuid, VERSION, CUSTOMER_TRX_ID) VALUES (" + CStr(var_consecutivo) + ",'" + IIf(IsNull(rs!Folio), "", rs!Folio) + "'," + var_fecha + ",'" + Replace(IIf(IsNull(rs!TITULAR), "", rs!TITULAR), "'", "") + "','" + Replace(IIf(IsNull(rs!Cliente), "", rs!Cliente), "'", "") + "','" + var_documento_eflow + "'," + VAR_FECHA_INICIO_TABLA + "," + VAR_FECHA_FIN_TABLA + "," + CStr(VAR_ESTATUS) + ", '" + Mid(CStr(VAR_SAT_UUID), 1, 500) + "','" + VAR_VERSION + "'," + CStr(var_customer_trx_id) + ")", cnn, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
                     'MsgBox rs(4).Value
               Wend
               rsaux2.Open "delete from TB_TEMP_ORACLE_DOC_FISCALES_VS_EFLOW WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " and fecha_inicio is null", cnn, adOpenDynamic, adLockOptimistic
               x = 0
               If x = 1 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_oracle_vs_eflow.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_ORACLE_VS_EFLOW.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\rep_comparacion_oracle_vs_eflow_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               Else
                  
                  Set oexcel = CreateObject("Excel.Application")
                  Set owbook = oexcel.Workbooks.Add
                  Set osheet = owbook.Worksheets(1)
                  osheet.Name = "ORACLE VS EFLOW"
                  Screen.MousePointer = vbHourglass
                  iFila = 1
                  ifila2 = 1
                  icol2 = 1
                  iCol = 1
                  var_cadena = "select distinct FECHA_INICIO, FECHA_FIN, FECHA, DOCUMENTO_ORACLE, DOCUMENTO_EFLOW, ESTATUS, SAT_UUID, TITULAR, CLIENTE, VERSION, CUSTOMER_TRX_ID  from TB_TEMP_ORACLE_DOC_FISCALES_VS_EFLOW WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND DOCUMENTO_ORACLE IS NOT NULL order by DOCUMENTO_ORACLE"
                  rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  For i = 0 To rsaux10.Fields.Count - 1
                      osheet.Cells(iFila, i + 1) = rsaux10.Fields(i).Name
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
                  archivo = "c:\reportessid\rep_comparacion_oracle_vs_eflow_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                  owbook.SaveAs archivo
                  oexcel.Visible = True
                  Set oexcel = Nothing
                  Screen.MousePointer = vbDefault
                  rsaux10.Close
                  
                  
                  
                  
                  
               End If
               MsgBox "Se a terminado de guardar el archivo " + archivo
               If rsaux1.State = 1 Then
                  rsaux1.Close
               End If
            Else
               MsgBox "No existen documentos para el periodo indicado", vbOKOnly, "ATENCION"
            End If
            rs.Close
            rs.Open "delete from TB_TEMP_ORACLE_DOC_FISCALES_VS_EFLOW WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "La fecha de inicio debe de ser menor o igual a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de Inicio incorrecta", vbOKOnly, "ATENCION"
   End If
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   If rsaux1.State = 1 Then
      rsaux1.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If rsaux3.State = 1 Then
      rsaux3.Close
   End If


End Sub

Private Sub Command3_Click()
   Dim cn As New ADODB.Connection
   Dim DSN As String
   Dim cn2 As New ADODB.Connection
                  Set oexcel = CreateObject("Excel.Application")
                  Set owbook = oexcel.Workbooks.Add
                  Set osheet = owbook.Worksheets(1)
                  osheet.Name = "ORACLE VS EFLOW"
                  Screen.MousePointer = vbHourglass
                  iFila = 1
                  ifila2 = 1
                  icol2 = 1
                  iCol = 1
               DSN = "eflow"
               cn.Open ("DSN=" & DSN & ";")
                     Set rsaux1 = cn.execute("SELECT  factura, fecha, ESTATUS FROM facturas where serie in ('FAIN','FAEVII','MTUINT','MPUINT','NCMPU','FAEVIND','FAEMX','FAEVXX','FAEVBI','FAEVPO','FAEVPS','FAEVPC','NCEMX','GTM_AR','NCEVPE','NCEVXX','NCEVBI','NCEMXO','NCTEII','FAEE','FAETR','FAEVIIN','INVIAGS','INVIAGS_','FAEII','VTHIN','FAEMX','FAEVXX','FAEVPO','FAEVPS','FAEG','UTVIN','NDCORR','NCEVIIN__','NCE VIIN','NCETR','NCEG','FAEVBI','FAEECA','FAEVPO','NCEERE','NCEERE','FAEVIIN__','NCIN','AR-INVOICE','FAEI','NCEVPS','NCEVPC','FAEERE','NCEVII')")
                     MsgBox CStr(rsaux1.Fields.Count)
                  With osheet
                      ' carga los registros del recordset
                      .Cells(iFila, iCol).CopyFromRecordset rsaux1
                      'oExcel.Columns(1).Select
                      'oExcel.Selection.NumberFormat = "#,##0.00"
                      'oExcel.Columns(1).Select
                      'oExcel.Selection.Font.Color = vbRed
                      .Columns.AutoFit ' ajusta el ancho de las columnas
                  End With
                  archivo = "c:\reportessid\rep_comparacion_oracle_vs_eflow_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                  owbook.SaveAs archivo
                  oexcel.Visible = True
                  Set oexcel = Nothing
                  Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 3000
   Left = 3500
   txt_inicio = Date
   txt_fin = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
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






