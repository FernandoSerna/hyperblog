VERSION 5.00
Begin VB.Form frmoracle_reporte_bultos_por_embarque 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte bultos por embarque por periodo"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3960
      Picture         =   "frmoracle_reporte_bultos_por_embarque.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      Picture         =   "frmoracle_reporte_bultos_por_embarque.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   30
      TabIndex        =   5
      Top             =   270
      Width           =   4275
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   840
      Left            =   60
      TabIndex        =   0
      Top             =   420
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
End
Attribute VB_Name = "frmoracle_reporte_bultos_por_embarque"
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
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            cnn.CommandTimeout = 720
            x = 0
            If x = 1 Then
               cnn.BeginTrans
               rs.Open "select max(inte_tem_consecutivo) as numero from TB_TEMP_ORACLE_CONTROL_RELACION_MAYOREO", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
               Else
                  var_consecutivo = 0
               End If
               var_consecutivo = var_consecutivo + 1
               rs.Close
               rs.Open "insert into TB_TEMP_ORACLE_CONTROL_RELACION_MAYOREO (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
               cnn.CommitTrans
            End If
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
            var_cadena = "select embarque, source_header_number pedido, fecha_inicio, fecha_fin, char_emb_estatus estatus_embarque,  vehiculo, tipo_Caja, b.INTE_PAQ_CAJA numero_caja, sello, char_paq_estatus estatus_caja, sum(floa_Sal_Cantidad_leida) as cantidad  from xxvia_Tb_encabezado_embarques a, xxvia_Tb_Salidas_cajas b Where A.Embarque = b.inte_emb_Embarque and a.fecha_inicio >= to_DATE('" + var_fecha_inicio + "','DD/MM/YYYY') and fecha_INICIO < to_date('" + var_fecha_fin + "','DD/MM/YYYY') and floa_sal_Cantidad_leida>0 group by embarque, source_header_number, fecha_inicio, fecha_fin, char_emb_estatus,  vehiculo, tipo_Caja, b.INTE_PAQ_CAJA, sello, char_paq_estatus order by embarque, b.SOURCE_HEADER_NUMBER"
            Me.txt_inicio = var_cadena
            rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            
            
            
            
            
            
            
            If Not rs.EOF Then
               Set oexcel = CreateObject("Excel.Application")
               Set owbook = oexcel.Workbooks.Add
               Set osheet = owbook.Worksheets(1)
               var_cadena = "PERIODO DEL " + Replace(var_fecha_inicio, "/", "_") + " AL " + Replace(var_fecha_fin, "/", "_")
               'MsgBox var_cadena
               osheet.Name = "DEL " + Replace(var_fecha_inicio, "/", "_") + " AL " + Replace(var_fecha_fin, "/", "_")
               Screen.MousePointer = vbHourglass
               iFila = 1
               ifila2 = 1
               icol2 = 1
               iCol = 1
               'rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               For i = 0 To rs.Fields.Count - 1
                   osheet.Cells(iFila, i + 1) = rs.Fields(i).Name
               Next
               iFila = iFila + 1
               With osheet
                  ' carga los registros del recordset
                  .Cells(iFila, iCol).CopyFromRecordset rs
                  'oExcel.Columns(1).Select
                  'oExcel.Selection.NumberFormat = "#,##0.00"
                  'oExcel.Columns(1).Select
                  'oExcel.Selection.Font.Color = vbRed
                  .Columns.AutoFit ' ajusta el ancho de las columnas
               End With
               owbook.SaveAs "c:\reportessid\reporte_de_bultos_por_embarque_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               oexcel.Visible = True
               Set oexcel = Nothing
               Screen.MousePointer = vbDefault
            Else
               MsgBox "No existen embarques para el periodo indicado", vbOKOnly, "ATENCION"
            End If
            rs.Close
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


