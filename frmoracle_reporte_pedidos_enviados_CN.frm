VERSION 5.00
Begin VB.Form frmoracle_reporte_pedidos_enviados_CN 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Envíos a CN"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   0
      Picture         =   "frmoracle_reporte_pedidos_enviados_CN.frx":0000
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
      Picture         =   "frmoracle_reporte_pedidos_enviados_CN.frx":0102
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
Attribute VB_Name = "frmoracle_reporte_pedidos_enviados_CN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objConn As New ADODB.Connection
Dim objCmd As New ADODB.Command
Dim objParm As ADODB.Parameter
Dim clnt As New SoapClient30
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter

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
            var_fecha_inicio_reporte = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
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
            var_fecha_FIN_REPORTE = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"

            
            
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
         
            strconsulta = "select E.FECHA, order_number,C.ATTRIBUTE1 ALMACEN, D.DESCRIPTION NOMBRE_ALMACEN, B.SEGMENT1 CODIGO, B.ITEM_dESCRIPTION DESCRIPCION, SUM(FLOA_SAL_cANTIDAD_LEIDA)  CANTIDAD"
            strconsulta = strconsulta + " from oe_order_headers_all a, xxvia_tb_Salidas_cajas B, po_requisition_headers_ALL C, MTL_SECONDARY_INVENTORIES D, XXVIA_VW_FECHA_LECTURA_PEDIDOS E Where order_number = source_header_number and ORDER_TYPE_ID = 1002 AND e.fecha >= TO_dATE(?,'DD/MM/YYYY') AND E.FECHA < TO_DATE(?,'DD/MM/YYYY') AND requisition_header_id = source_document_id AND secondary_inventory_name = C.ATTRIBUTE1 AND FLOA_SAL_CANTIDAD_LEIDA > 0 AND A.HEADER_ID = E.HEADER_ID and char_paq_estatus = 'S'"
            strconsulta = strconsulta + " GROUP BY E.FECHA, order_number, C.ATTRIBUTE1, D.DESCRIPTION, B.SEGMENT1, B.ITEM_dESCRIPTION ORDER BY E.FECHA, C.ATTRIBUTE1"
             
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 20, CStr(var_fecha_inicio))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 20, CStr(var_fecha_fin))
                 .Parameters.Append parametro
            End With
            Set rs = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
                     
                     
                     
            If Not rs.EOF Then
               Set oexcel = CreateObject("Excel.Application")
               Set owbook = oexcel.Workbooks.Add
               Set osheet = owbook.Worksheets(1)
               var_cadena = "PERIODO DEL " + Replace(var_fecha_inicio, "/", "_") + " AL " + Replace(var_fecha_fin, "/", "_")
               'MsgBox var_cadena
               osheet.Name = "DEL " + Replace(var_fecha_inicio, "/", "_") + " AL " + Replace(CStr(CDate(var_fecha_fin) - 1), "/", "_")
               Screen.MousePointer = vbHourglass
               iFila = 1
               ifila2 = 1
               icol2 = 1
               iCol = 1
               'rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               For i = 0 To rs.Fields.Count - 1
                   osheet.Cells(iFila, i + 1) = rs.Fields(i).Name
               Next
               rs.MoveFirst
               iFila = iFila + 1
               With osheet
                  ' carga los registros del recordset
                  'MsgBox rs.RecordCount
                  .Cells(iFila, iCol).CopyFromRecordset rs
                  'oexcel.Columns(13).Select
                  'oexcel.Selection.NumberFormat = "###,###,##0.00"
                  
                  'oExcel.Columns(1).Select
                  'oExcel.Selection.Font.Color = vbRed
                  .Columns.AutoFit ' ajusta el ancho de las columnas
                  VAR_FILAS = iFila + rs.RecordCount
                  
                  '.Cells(VAR_FILAS, 10).Value = "TOTAL BULTOS:"
                  '.Cells(VAR_FILAS, 10).Font.Bold = True
                  '.Cells(VAR_FILAS, 10).horizontalAlignment = xlRight
                  
                  '.Cells(VAR_FILAS, 11).Value = rs.RecordCount
                  '.Cells(VAR_FILAS, 11).NumberFormat = "###,###,##0"
                  '.Cells(VAR_FILAS, 11).Font.Bold = True
                  
                  '.Cells(VAR_FILAS, 12).Value = "TOTAL PIEZAS:"
                  '.Cells(VAR_FILAS, 12).Font.Bold = True
                  '.Cells(VAR_FILAS, 12).horizontalAlignment = xlRight
                  
                  '.Cells(VAR_FILAS, 13).Value = var_cantidad
                  '.Cells(VAR_FILAS, 13).Font.Bold = True
                  '.Cells(VAR_FILAS, 13).NumberFormat = "###,###,##0.00"
                  .Columns.AutoFit
               End With
               
               
               With osheet
                  .Cells(1, 1).Font.Bold = True
                  .Cells(1, 2).Font.Bold = True
                  .Cells(1, 3).Font.Bold = True
                  .Cells(1, 4).Font.Bold = True
                  .Cells(1, 5).Font.Bold = True
                  .Cells(1, 6).Font.Bold = True
                  .Cells(1, 7).Font.Bold = True
                  .Columns.AutoFit
               End With
               
               
               
               owbook.SaveAs "c:\reportessid\reporte_envios_CN_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               oexcel.Visible = True
               Set oexcel = Nothing
               Screen.MousePointer = vbDefault
            Else
               MsgBox "No existen envios para el periodo seleccionado", vbOKOnly, "ATENCION"
            End If
         
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
   Me.txt_inicio = Date
   Me.txt_fin = Date
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
