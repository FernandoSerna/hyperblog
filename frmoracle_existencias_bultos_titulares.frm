VERSION 5.00
Begin VB.Form frmoracle_existencias_bultos_titulares 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Existencias bultos titulares"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   2880
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.TextBox txt_fecha 
         Height          =   375
         Left            =   1050
         TabIndex        =   2
         Top             =   210
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   390
         TabIndex        =   1
         Top             =   270
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmoracle_existencias_bultos_titulares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim comandoORA As New adodb.Command
Dim parametro As adodb.Parameter

Private Sub Form_Load()
   Top = 3200
   Left = 4200
   Me.txt_fecha = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_fecha_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fecha) Then
         frmcalendario.mes = CDate(Me.txt_fecha)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      txt_fecha = var_fecha_general
   End If

End Sub

Private Sub txt_fecha_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsDate(Me.txt_fecha) Then
         var_dia_s = CStr(Day(CDate(Me.txt_fecha)))
         var_mes_s = CStr(Month(CDate(Me.txt_fecha)))
         var_año_s = CStr(Year(CDate(Me.txt_fecha)))
            
         If Len(var_dia_s) = 1 Then
            var_dia_s = "0" + var_dia_s
         End If
         If Len(var_mes_s) = 1 Then
            var_mes_s = "0" + var_mes_s
         End If
         If Len(var_año_s) = 2 Then
            var_año_s = "20" + var_dia_s
         End If
         var_fecha_s = var_dia_s + "/" + var_mes_s + "/" + var_año_s
         var_fecha = "{d '" + var_año_s + "-" + var_mes_s + "-" + var_dia_s + "'}"
         strconsulta = "select titular, nombre_titular, CODIGO, DESCRIPTION, SUM(CANTIDAD) AS CANTIDAD  from xxvia_vw_exis_bultos_tit WHERE CREATION_DATE < TO_DATE(?,'DD/MM/YYYY') GROUP BY  titular, nombre_titular, CODIGO, DESCRIPTION ORDER BY NOMBRE_TITULAR"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 50, var_fecha_s)
              .Parameters.Append parametro
         End With
         Set rsaux = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         If Not rsaux.EOF Then
            cnn.BeginTrans
            rs.Open "SELECT MAX(INTE_tEM_CONSECUTIVO) FROM TB_TEMP_ORACLE_EXISTENCIAS_TITULAR", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rs.Open "INSERT INTO TB_TEMP_ORACLE_EXISTENCIAS_TITULAR (INTE_tEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ") ", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            While Not rsaux.EOF
                  rs.Open "Insert Into TB_TEMP_ORACLE_EXISTENCIAS_TITULAR (INTE_TEM_CONSECUTIVO, FECHA, TITULAR, NOMBRE_TITULAR, CODIGO, DESCRIPCION, CANTIDAD) VALUES (" + CStr(var_consecutivo) + "," + var_fecha + ",'" + rsaux!TITULAR + "','" + rsaux!NOMBRE_TITULAR + "','" + rsaux!CODIGO + "','" + rsaux!Description + "'," + CStr(rsaux!Cantidad) + ")", cnn, adOpenDynamic, adLockOptimistic
                  rsaux.MoveNext
            Wend
            
            Set reporte = appl.OpenReport(App.Path + "\rep_oracle_existencias_bultos_titular.rpt")
            reporte.RecordSelectionFormula = "{VW_ORACLE_EXISTENCIAS_BULTOS_TITULAR.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Ordenes de surtido historica"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            
            var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_existencias_bultos_titular.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_EXISTENCIAS_BULTOS_TITULAR.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\Existencias_bultos_titular_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
            
            
            
         Else
         End If
         rsaux.Close
      End If
   End If
End Sub
