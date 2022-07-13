VERSION 5.00
Begin VB.Form frmoracle_salidas_privalia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de salidas Privalia"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_entradas_privalia 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   360
      Picture         =   "frmoracle_salidas_privalia.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3945
      Picture         =   "frmoracle_salidas_privalia.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   0
      Picture         =   "frmoracle_salidas_privalia.frx":073C
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
      TabIndex        =   7
      Top             =   270
      Width           =   4275
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
End
Attribute VB_Name = "frmoracle_salidas_privalia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report



Private Sub cmd_entradas_privalia_Click()
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
             
            rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            rsaux.Open "ALTER SESSION SET NLS_DATE_FORMAT = 'DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_cadena = "select '" + Me.txt_inicio + "' as FECHA_INICIO, '" + Me.txt_fin + "' as FECHA_FIN, CODIGO, SUM(CANTIDAD) AS CANTIDAD FROM XXVIA_tB_DEVOLUCIONES_CLIENTES where to_date(SUBSTR(fecha_fin,1,10),'DD/MM/YYYY')  >= to_date('" + Me.txt_inicio + "','DD/MM/YYYY')  AND to_date(SUBSTR(fecha_fin,1,10),'DD/MM/YYYY') < TO_DATE('" + Me.txt_fin + "','DD/MM/YYYY') + 1 AND MOVIMIENTO = 'SP' AND CANTIDAD > 0 group by CODIGO"
              
            Set oexcel = CreateObject("Excel.Application")
            Set owbook = oexcel.Workbooks.Add
            Set osheet = owbook.Worksheets(1)
            osheet.Name = "SALIDAS PRIVALIA"
            Screen.MousePointer = vbHourglass
            iFila = 1
            ifila2 = 1
            icol2 = 1
            iCol = 1
            MsgBox var_cadena
            rsaux10.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
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
            archivo = "c:\reportessid\salidas_privalia_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            owbook.SaveAs archivo
            oexcel.Visible = True
            Set oexcel = Nothing
            Screen.MousePointer = vbDefault
            rsaux10.Close
            MsgBox "Se a terminado de guardar el archivo " + archivo
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
             
            rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            rsaux.Open "ALTER SESSION SET NLS_DATE_FORMAT = 'DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_cadena = "select '" + Me.txt_inicio + "' as FECHA_INICIO, '" + Me.txt_fin + "' as FECHA_FIN, CODIGO, DESCRIPCION, REFERENCIA, SUM(CANTIDAD) AS CANTIDAD FROM XXVIA_tB_DEVOLUCIONES_CLIENTES where to_date(SUBSTR(fecha_fin,1,10),'DD/MM/YYYY')  >= to_date('" + Me.txt_inicio + "','DD/MM/YYYY')  AND to_date(SUBSTR(fecha_fin,1,10),'DD/MM/YYYY') < TO_DATE('" + Me.txt_fin + "','DD/MM/YYYY') + 1 AND MOVIMIENTO = 'SP' AND CANTIDAD > 0 group by CODIGO, DESCRIPCION, REFERENCIA"
            Text1 = var_cadena
            Set oexcel = CreateObject("Excel.Application")
            Set owbook = oexcel.Workbooks.Add
            Set osheet = owbook.Worksheets(1)
            osheet.Name = "SALIDAS PRIVALIA"
            Screen.MousePointer = vbHourglass
            iFila = 1
            ifila2 = 1
            icol2 = 1
            iCol = 1
            rsaux10.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
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
            archivo = "c:\reportessid\salidas_privalia_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            owbook.SaveAs archivo
            oexcel.Visible = True
            Set oexcel = Nothing
            Screen.MousePointer = vbDefault
            rsaux10.Close
            MsgBox "Se a terminado de guardar el archivo " + archivo
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


