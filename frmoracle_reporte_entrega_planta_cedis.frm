VERSION 5.00
Begin VB.Form frmoracle_reporte_entrega_planta_cedis 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de tiempo de entrega"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   0
      Picture         =   "frmoracle_reporte_entrega_planta_cedis.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3945
      Picture         =   "frmoracle_reporte_entrega_planta_cedis.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   840
      Left            =   60
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
      Left            =   30
      TabIndex        =   7
      Top             =   270
      Width           =   4275
   End
End
Attribute VB_Name = "frmoracle_reporte_entrega_planta_cedis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
   Dim iFila As Long, iCol As Integer, i As Integer
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            cnn.CommandTimeout = 720
            cnn.BeginTrans
            rs.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_ORACLE_REPORTE_TIEMPO_NOTAS_PLANTAS_CEDIS", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!numero), 0, rs!numero)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "insert into TB_TEMP_ORACLE_REPORTE_TIEMPO_NOTAS_PLANTAS_CEDIS (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
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
            
            var_fecha_inicio = var_dia + "/" + var_mes + "/" + var_año
            VAR_FECHA_INICIO_TABLA = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
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
             
            rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
              
            var_cadena = "select haou.NAME Oranización_Origen, sin.DESCRIPTION Almacén_Origen, mmtE.SHIPMENT_NUMBER nota, to_char(max(mmtE.creation_date),'DD/MM/YYYY HH24:MI:SS') fecha_envio, to_char(max(mmtR.creation_date),'DD/MM/YYYY HH24:MI:SS') fecha_Recepcion, trunc(max(nvl(mmtR.creation_date,sysdate)) - min(mmtE.CREATION_DATE))  dias, trunc((((max(nvl(mmtR.creation_date,sysdate)) - min(mmtE.CREATION_DATE)) - trunc(max(nvl(mmtR.creation_date,sysdate)) - min(mmtE.CREATION_DATE)))*1440) /60)horas , trunc((((((max(nvl(mmtR.creation_date,sysdate)) - min(mmtE.CREATION_DATE)) - trunc(max(nvl(mmtR.creation_date,sysdate)) - min(mmtE.CREATION_DATE)))*1440) /60) - trunc((((max(nvl(mmtR.creation_date,sysdate)) - min(mmtE.CREATION_DATE)) - trunc(max(nvl(mmtR.creation_date,sysdate)) - min(mmtE.CREATION_DATE)))*1440) /60)) *60) minutos, "
            var_cadena = var_cadena + " trunc((((((((max(nvl(mmtR.creation_date,sysdate)) - min(mmtE.CREATION_DATE)) - trunc(max(nvl(mmtR.creation_date,sysdate)) - min(mmtE.CREATION_DATE)))*1440) /60) - trunc((((max(nvl(mmtR.creation_date,sysdate)) - min(mmtE.CREATION_DATE)) - trunc(max(nvl(mmtR.creation_date,sysdate)) - min(mmtE.CREATION_DATE)))*1440) /60)) *60)  - trunc((((((max(nvl(mmtR.creation_date,sysdate)) - min(mmtE.CREATION_DATE)) - trunc(max(nvl(mmtR.creation_date,sysdate)) - min(mmtE.CREATION_DATE)))*1440) /60) - trunc((((max(nvl(mmtR.creation_date,sysdate)) - min(mmtE.CREATION_DATE)) - trunc(max(nvl(mmtR.creation_date,sysdate)) - min(mmtE.CREATION_DATE)))*1440) /60)) *60)) * 60)  segundos, mmtE.TRANSFER_ORGANIZATION_ID, mmtE.TRANSFER_SUBINVENTORY from INV.MTL_MATERIAL_TRANSACTIONS mmtE, INV.MTL_MATERIAL_TRANSACTIONS mmtR, HR.HR_ALL_ORGANIZATION_UNITS haou,"
            var_cadena = var_cadena + " INV.MTL_SECONDARY_INVENTORIES Sin Where mmtE.TRANSACTION_TYPE_ID = 21 and mmtE.CREATION_DATE >= TO_DATE('" + var_fecha_inicio + "','DD/MM/YYYY') AND mmtE.CREATION_DATE < TO_DATE('" + var_fecha_fin + "','DD/MM/YYYY') and mmtR.TRANSACTION_TYPE_ID (+)= 12 and mmtE.SHIPMENT_NUMBER =  mmtR.SHIPMENT_NUMBER (+) and mmtE.ORGANIZATION_ID =  haou.ORGANIZATION_ID and mmtE.ORGANIZATION_ID = sin.ORGANIZATION_ID and case when mmtE.SUBINVENTORY_CODE = 'ALMCAJAS' and mmtE.TRANSACTION_SOURCE_NAME is not null then mmtE.TRANSACTION_SOURCE_NAME else mmtE.SUBINVENTORY_CODE end = sin.SECONDARY_INVENTORY_NAME AND mmtE.TRANSFER_ORGANIZATION_ID = 93 and mmtE.TRANSFER_SUBINVENTORY = 'CDI_ALMPT' group by mmtE.SHIPMENT_NUMBER, haou.NAME, sin.DESCRIPTION, mmtE.TRANSFER_ORGANIZATION_ID, mmtE.TRANSFER_SUBINVENTORY ORDER BY 1,2,3"
            rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               While Not rs.EOF
                     var_cadena = "INSERT INTO TB_TEMP_ORACLE_REPORTE_TIEMPO_NOTAS_PLANTAS_CEDIS (INTE_TEM_CONSECUTIVO, ORGANIZACION, ALMACEN_ORIGEN, NOTA, FECHA_ENVIO, FECHA_RECEPCION, DIAS, HORAS, MINUTOS, ORGANIZACION_TRANSFERIR, ALMACEN_TRANSFERIR, FECHA_INICIO, FECHA_FIN)"
                     var_cadena = var_cadena + " VALUES (" + CStr(var_consecutivo) + ", '" + rs!Oranización_Origen + "','" + rs!ALMACéN_ORIGEN + "','" + rs!NOTA + "','" + CStr(rs!FECHA_ENVIO) + "','" + CStr(IIf(IsNull(rs!FECHA_RECEPCION), "", rs!FECHA_RECEPCION)) + "'," + CStr(rs!DIAS) + "," + CStr(rs!horas) + "," + CStr(rs!MINUTOS) + ",93,'CDI_ALMPT'," + VAR_FECHA_INICIO_TABLA + "," + VAR_FECHA_FIN_TABLA + ")"
                     rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               rsaux1.Open "DELETE FROM TB_TEMP_ORACLE_REPORTE_TIEMPO_NOTAS_PLANTAS_CEDIS WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND ORGANIZACION IS NULL", cnn, adOpenDynamic, adLockOptimistic
               
               'Set reporte = appl.OpenReport(App.Path + "\rep_oracle_valuacion_devoluciones_detalle.rpt")
               'reporte.RecordSelectionFormula = "{VW_ORACLE_VALUACION_DEVOLUCIONES_DETALLE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               'frmvistasprevias.cr.ReportSource = reporte
               'For ntablas = 1 To reporte.Database.Tables.Count
               '    reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               'Next ntablas
               'frmvistasprevias.cr.ViewReport
               'frmvistasprevias.Caption = "Valuación de devoluciones a detalle"
               'frmvistasprevias.Show 1
               'Set reporte = Nothing
    
               'var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
               var_si = 6
               If var_si = 6 Then
                  Set oexcel = CreateObject("Excel.Application")
                  Set oWBook = oexcel.Workbooks.Add
                  Set oSheet = oWBook.Worksheets(1)
                  oSheet.Name = "TIEMPO PLANTA - CEDIS"
                  Screen.MousePointer = vbHourglass
                  iFila = 1
                  iFila2 = 1
                  iCol2 = 1
                  iCol = 1
                  var_cadena = " SELECT FECHA_INICIO, FECHA_FIN, ORGANIZACION, ALMACEN_ORIGEN, NOTA, FECHA_ENVIO, FECHA_RECEPCION, DIAS, HORAS, MINUTOS, ORGANIZACION_TRANSFERIR, ALMACEN_TRANSFERIR From dbo.TB_TEMP_ORACLE_REPORTE_TIEMPO_NOTAS_PLANTAS_CEDIS WHERE (INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND (ORGANIZACION IS NOT NULL))"
                  rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  For i = 0 To rsaux10.Fields.Count - 1
                      oSheet.Cells(iFila, i + 1) = rsaux10.Fields(i).Name
                  Next
                  iFila = iFila + 1
                  With oSheet
                      ' carga los registros del recordset
                      .Cells(iFila, iCol).CopyFromRecordset rsaux10
                      'oExcel.Columns(1).Select
                      'oExcel.Selection.NumberFormat = "#,##0.00"
                      'oExcel.Columns(1).Select
                      'oExcel.Selection.Font.Color = vbRed
                      .Columns.AutoFit ' ajusta el ancho de las columnas
                  End With
                  archivo = "c:\reportessid\LEAD_TIME_PANTAS_CEDIS" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                  oWBook.SaveAs archivo
                  oexcel.Visible = True
                  Set oexcel = Nothing
                  Screen.MousePointer = vbDefault
                  rsaux10.Close
                  MsgBox "Se a terminado de guardar el archivo " + archivo
               End If
            Else
               MsgBox "No existen notas para el periodo indicado", vbOKOnly, "ATENCION"
            End If
            rs.Close
            rs.Open "delete from TB_TEMP_ORACLE_REPORTE_TIEMPO_NOTAS_PLANTAS_CEDIS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "La fecha de inicio debe de ser menor o igual a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de Inicio incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
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


