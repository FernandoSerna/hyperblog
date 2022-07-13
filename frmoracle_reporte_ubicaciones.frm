VERSION 5.00
Begin VB.Form frmoracle_reporte_ubicaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de ubicaciones"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      Picture         =   "frmoracle_reporte_ubicaciones.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4305
      Picture         =   "frmoracle_reporte_ubicaciones.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   90
      Left            =   30
      TabIndex        =   1
      Top             =   285
      Width           =   4635
   End
   Begin VB.Frame Frame1 
      Height          =   1110
      Left            =   75
      TabIndex        =   0
      Top             =   345
      Width           =   4560
      Begin VB.TextBox txt_ubicacion_1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   570
         Left            =   75
         TabIndex        =   2
         Top             =   420
         Width           =   4395
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000FF&
         Caption         =   " Ubicación"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   30
         TabIndex        =   5
         Top             =   135
         Width           =   4485
      End
   End
End
Attribute VB_Name = "frmoracle_reporte_ubicaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter


Private Sub cmd_imprimir_Click()
   Dim var_ubicacion As String
   
   var_ubicacion = Trim(Me.txt_ubicacion_1)
   If var_ubicacion <> "" Then
      cnn.BeginTrans
      rsaux1.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_ORACLE_REPORTE_UBICACIONES", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux1.EOF Then
         var_consecutivo = IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value) + 1
      Else
         var_consecutivo = 1
      End If
      rsaux1.Close
      rsaux1.Open "INSERT INTO TB_TEMP_ORACLE_REPORTE_UBICACIONES (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
      cnn.CommitTrans
      If var_unidad_organizacional = 93 Then
         var_almacen = "CDI_ALMPT"
      End If
      If var_unidad_organizacional = 90 Then
         var_almacen = "CDISTEX_PT"
      End If
      If var_unidad_organizacional <> 930000 Then
      For var_j = 1 To 6
          If var_j = 1 Then
             rs.Open "select nvl(attribute2,' ') ubicacion,inventory_item_id item_id, segment1, description from  xxvia_system_items_b where  organization_id  = " + var_unidad_organizacional + " and ATTRIBUTE2 LIKE '" + var_ubicacion + "%' and organization_id = " + CStr(var_unidad_organizacional) + " order by attribute2 ", cnnoracle_4, adOpenDynamic, adLockOptimistic
          End If
          If var_j = 2 Then
             rs.Open "select nvl(attribute3,' ') ubicacion,inventory_item_id item_id, segment1, description from  xxvia_system_items_b where  organization_id  = " + var_unidad_organizacional + " and ATTRIBUTE3 LIKE '" + var_ubicacion + "%' and organization_id = " + CStr(var_unidad_organizacional) + " order by attribute3", cnnoracle_4, adOpenDynamic, adLockOptimistic
          End If
          If var_j = 3 Then
             rs.Open "select nvl(attribute4,' ') ubicacion,inventory_item_id item_id, segment1, description from  xxvia_system_items_b where  organization_id  = " + var_unidad_organizacional + " and ATTRIBUTE4 LIKE '" + var_ubicacion + "%' and organization_id = " + CStr(var_unidad_organizacional) + " order by attribute4", cnnoracle_4, adOpenDynamic, adLockOptimistic
          End If
          If var_j = 4 Then
             rs.Open "select nvl(attribute5,' ') ubicacion,inventory_item_id item_id, segment1, description from  xxvia_system_items_b where  organization_id  = " + var_unidad_organizacional + " and ATTRIBUTE5 LIKE '" + var_ubicacion + "%' and organization_id = " + CStr(var_unidad_organizacional) + " order by attribute5", cnnoracle_4, adOpenDynamic, adLockOptimistic
          End If
          If var_j = 5 Then
             rs.Open "select nvl(attribute6,' ') ubicacion,inventory_item_id item_id, segment1, description from  xxvia_system_items_b where  organization_id  = " + var_unidad_organizacional + " and ATTRIBUTE6 LIKE '" + var_ubicacion + "%' and organization_id = " + CStr(var_unidad_organizacional) + " order by attribute6", cnnoracle_4, adOpenDynamic, adLockOptimistic
          End If
          If var_j = 6 Then
             rs.Open "select nvl(attribute7,' ') ubicacion,inventory_item_id item_id, segment1, description from  xxvia_system_items_b where  organization_id  = " + var_unidad_organizacional + " and ATTRIBUTE7 LIKE '" + var_ubicacion + "%' and organization_id = " + CStr(var_unidad_organizacional) + " order by attribute7", cnnoracle_4, adOpenDynamic, adLockOptimistic
          End If
          If Not rs.EOF Then
             While Not rs.EOF
                   rsaux1.Open "INSERT INTO TB_TEMP_ORACLE_REPORTE_UBICACIONES (INTE_TEM_CONSECUTIVO, CODIGO, DESCRIPCION, UBICACION, NUMERO_UBICACION, inventory_item_id) VALUES (" + CStr(var_consecutivo) + ",'" + IIf(IsNull(rs!SEGMENT1), "", rs!SEGMENT1) + "','" + IIf(IsNull(rs!Description), "", rs!Description) + "','" + rs!ubicacion + "'," + CStr(var_j) + "," + CStr(rs!ITEM_ID) + ")", cnn, adOpenDynamic, adLockOptimistic
                   rs.MoveNext
             Wend
          End If
          rs.Close
      Next var_j
      End If
      If var_unidad_organizacional <> "930000" Then
         rs.Open "select CODIGO AS SEGMENT1, min(numero_ubicacion) AS NUMERO  from tb_temp_oracle_Reporte_ubicaciones where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and inventory_item_id is not null group by CODIGO", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               strconsulta = "select * from Xxvia_vw_existencias_inv where organization_id = ? and subinventory_code = ? and segment1 = ?"
               With comandoORA
                    .ActiveConnection = cnnoracle_4
                    .CommandType = adCmdText
                    .CommandText = strconsulta
                    Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, CDbl(var_unidad_organizacional))
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_almacen)
                    .Parameters.Append parametro
                    Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, rs!SEGMENT1)
                    .Parameters.Append parametro
               End With
               Set rsaux9 = comandoORA.execute
               Set comandoORA = Nothing
               Set parametro = Nothing
         
         
               'rsaux9.Open "SELECT * FROM Xxvia_vw_existencias_inv WHERE ORGANIZATION_ID = " + var_unidad_organizacional + " AND SEGMENT1 = '" + rs!segment1 + "'  AND SUBINVENTORY_CODE = 'CDI_ALMPT'", cnnoracle_4, adOpenDynamic, adLockOptimistic
               If Not rsaux9.EOF Then
                  var_cantidad = IIf(IsNull(rsaux9!cantmano), 0, rsaux9!cantmano)
               Else
                  var_cantidad = 0
               End If
               rsaux9.Close
               rsaux9.Open "UPDATE tb_temp_oracle_Reporte_ubicaciones SET CANTIDAD = " + CStr(var_cantidad) + " WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND CODIGO = '" + rs!SEGMENT1 + "' AND NUMERO_UBICACION = " + CStr(rs!numero), cnn, adOpenDynamic, adLockOptimistic
               rs.MoveNext
         Wend
         rs.Close
      End If
      rs.Open "DELETE FROM TB_tEMP_ORACLE_REPORTE_UBICACIONES WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND CODIGO IS NULL", cnn, adOpenDynamic, adLockOptimistic
      Set reporte = appl.OpenReport(App.Path + "\rep_oracle_ubicaciones.rpt")
      var_cadena = "{VW_ORACLE_REPORTE_UBICACIONES.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
      reporte.RecordSelectionFormula = var_cadena
      frmvistasprevias.cr.ReportSource = reporte
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      frmvistasprevias.cr.ViewReport
      frmvistasprevias.Caption = "Reporte de ubicaciones"
      frmvistasprevias.Show 1
      Set reporte = Nothing
         
      var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         Set reporte = appl.OpenReport(App.Path + "\rep_oracle_ubicaciones.rpt")
         var_cadena = "{VW_ORACLE_REPORTE_UBICACIONES.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
         reporte.RecordSelectionFormula = var_cadena
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         reporte.ExportOptions.FormatType = crEFTExcel80
         reporte.ExportOptions.DestinationType = crEDTDiskFile
         archivo = "c:\reportessid\ubicaciones_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
         reporte.ExportOptions.DiskFileName = archivo
         reporte.Export False
         Set reporte = Nothing
      End If
      
      rs.Open "DELETE FROM TB_tEMP_ORACLE_REPORTE_UBICACIONES WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
      
   Else
      MsgBox "No se han seleccionado ubicaciones", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 3300
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_ubicacion_1_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   End If
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub
