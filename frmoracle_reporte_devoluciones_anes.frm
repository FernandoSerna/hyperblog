VERSION 5.00
Begin VB.Form frmoracle_reporte_devoluciones_anes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Devoluciones ANEs."
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   0
      Picture         =   "frmoracle_reporte_devoluciones_anes.frx":0000
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
      Picture         =   "frmoracle_reporte_devoluciones_anes.frx":0102
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
Attribute VB_Name = "frmoracle_reporte_devoluciones_anes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter


Private Sub cmd_imprimir_Click()

                   var_dia = CStr(Day(CDate(Me.txt_inicio)))
                   var_mes = CStr(Month(CDate(Me.txt_inicio)))
                   var_año = CStr(Year(CDate(Me.txt_inicio)))
                   If Len(Trim(var_dia)) = 1 Then
                      var_dia = "0" + var_dia
                   End If
                   If Len(Trim(var_mes)) = 1 Then
                      var_mes = "0" + var_mes
                   End If
            
                   var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"

                   var_dia = CStr(Day(CDate(Me.txt_fin) + 1))
                   var_mes = CStr(Month(CDate(Me.txt_fin) + 1))
                   var_año = CStr(Year(CDate(Me.txt_fin) + 1))
                   If Len(Trim(var_dia)) = 1 Then
                      var_dia = "0" + var_dia
                   End If
                   If Len(Trim(var_mes)) = 1 Then
                      var_mes = "0" + var_mes
                   End If
            
                   var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"



                  Set oexcel = CreateObject("Excel.Application")
                  Set owbook = oexcel.Workbooks.Add
                  Set osheet = owbook.Worksheets(1)
                  osheet.Name = "DEVOLUCIONES ANES"
                  Screen.MousePointer = vbHourglass
                  iFila = 1
                  ifila2 = 1
                  icol2 = 1
                  iCol = 1
                  var_cadena = "select DISTINCT ESTABLECIMIENTO from TB_DEVOLUCIONES where FECHA_INICIO >= " + var_fecha_inicio + " and FECHA_INICIO < " + var_fecha_fin + " and ESTATUS = 'I'"
                  rsaux10.Open var_cadena, cnn_devolucion_anes, adOpenDynamic, adLockOptimistic
                  While Not rsaux10.EOF
                        var_establecimiento = rsaux10!establecimiento
                        If var_establecimiento = 7572 Then
                           var_establecimiento = 7573
                        End If
                        strconsulta = "SELECT CALLE||' '||NUM_CALLE||', '||COLONIA||', '||CIUDAD||', '||MUNICIPIO||', '|| ESTADO||', '||CODIGO_POSTAL AS DIRECCION,  ACCOUNT_NUMBER AS TITULAR, ACCOUNT_FULL_NAME AS NOMBRE_TITULAR FROM XXVIA_VW_CLIENTES_BCP WHERE SITE_USE_ID = ?"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_establecimiento)
                             .Parameters.Append parametro
                        End With
                        Set rsaux9 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        VAR_DIRECCION = ""
                        'MsgBox "UPDATE TB_DEVOLUCIONES SET DIRECCION = '" + rsaux9!DIRECCION + "' WHERE ESTABLECIMIENTO = " + CStr(rsaux10!ESTABLECIMIENTO)
                        'If Not rsaux9.EOF Then
                           rsaux8.Open "UPDATE TB_DEVOLUCIONES SET DIRECCION = '" + rsaux9!DIRECCION + "', CLAVE_TITULAR = '" + rsaux9!TITULAR + "', NOMBRE_TITULAR = '" + rsaux9!NOMBRE_TITULAR + "' WHERE ESTABLECIMIENTO = " + CStr(rsaux10!establecimiento), cnn_devolucion_anes, adOpenDynamic, adLockOptimistic
                        'End If
                        rsaux10.MoveNext
                  Wend
                  rsaux10.Close
                  
                  var_cadena = "select FECHA_INICIO AS FECHA, NUMERO, NOMBRE_AGENTE, CLAVE_TITULAR, NOMBRE_TITULAR, ESTABLECIMIENTO, NOMBRE_ESTABLECIMIENTO , CODIGO, CANTIDAD, DIRECCION, REFERENCIA, TIPO_DEVOLUCION_1, TIPO_DEVOLUCION_2 from TB_DEVOLUCIONES where FECHA_INICIO >= " + var_fecha_inicio + " and FECHA_INICIO < " + var_fecha_fin + " and ESTATUS = 'I' ORDER BY NUMERO"
                  rsaux10.Open var_cadena, cnn_devolucion_anes, adOpenDynamic, adLockOptimistic
                  
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
                  archivo = "c:\reportessid\rep_devoluciones_ANEs_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                  owbook.SaveAs archivo
                  oexcel.Visible = True
                  Set oexcel = Nothing
                  Screen.MousePointer = vbDefault
                  rsaux10.Close

End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   If cnn_devolucion_anes.State = 1 Then
      cnn_devolucion_anes.Close
   End If
   cnn_devolucion_anes.Open "Provider=SQLOLEDB.1;Password=" & parametros(3) & ";Persist Security Info=True;User ID=sa;Initial Catalog=devolucion_anes;Data Source=SQLQUEZADA2"
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
